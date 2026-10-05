"""
formatter.py
------------
Core, reusable TOC (Table of Contents) formatting logic for the TOC Formatter Tool.

DESIGN GOAL (per project brief):
    "Is tool ko hum reusable rakhenge, taaki future mein agar numbering ya
    formatting change ho to bas ek function update karna pade."
    -> If numbering style or output formatting rules change in the future,
       only ONE function/config in this file should need to change.

Everything that "knows about" numbering schemes lives in this file, behind
a small set of pure functions. app.py (the UI) never parses text itself -
it only calls into this module.

Public API used by app.py / utils.py:
    parse_toc(lines)            -> list[TocEntry]
    renumber(entries)           -> list[TocEntry]   (assigns clean 1. / 1.1. / ... numbers)
    render(entries, fmt)        -> str               (fmt: 'txt' | 'md' | 'json')
    process_text(raw_text)      -> (entries, rendered_txt)
"""

import re
import json
from dataclasses import dataclass, field, asdict

# ---------------------------------------------------------------------------
# CONFIG - the only section you should need to touch if numbering/formatting
# rules change in the future.
# ---------------------------------------------------------------------------

INDENT_UNIT = "    "        # 4 spaces per nesting level in txt/preview output
MD_BULLET_STYLE = "-"       # bullet used for markdown export (numbers still shown)
NUMBER_SUFFIX = ""          # appended after the number, e.g. "." for "1." - empty for "1"
NUMBER_TITLE_SEP = "\t"     # separator between the number and the title, e.g. "1\tIndustry Outlook"
BLANK_LINE_BETWEEN_ENTRIES = False  # compact list (no blank line after every entry) when False

# Section headings that mark the start of a FLAT, independently-numbered list
# (e.g. a "List of Key Figures and Tables" appendix). Real TOCs commonly end
# with a section like this where every following line is a standalone
# figure/table caption - siblings of each other, not nested under anything -
# and gets numbered 1, 2, 3, ... from scratch rather than continuing the
# main document's hierarchical numbering. Matched case-insensitively against
# the full entry title. Add more patterns here if other restart-style
# sections need the same treatment.
RESTART_LIST_TRIGGERS = [
    re.compile(r"^list of (key )?figures?(\s*(and|&)\s*tables?)?$", re.IGNORECASE),
    re.compile(r"^list of tables?(\s*(and|&)\s*figures?)?$", re.IGNORECASE),
]
# Restarted-list entries use their own number format - a classic flat
# numbered list ("1.\tTitle") - independent of NUMBER_SUFFIX/NUMBER_TITLE_SEP
# above, which apply to the main hierarchical TOC.
RESTART_LIST_NUMBER_SUFFIX = "."
RESTART_LIST_NUMBER_TITLE_SEP = "\t"

# Regex patterns used to *detect and parse* numbering already present in a line.
# Order matters: first match wins.
#
# IMPORTANT: the trailing "\." here is MANDATORY (not "\.?"). Real TOC titles
# routinely start with a bare number that ISN'T a section number at all -
# e.g. "3000 Series" / "5000 Series" (aluminium alloy grades), "24/7 Support",
# "5 Forces Analysis". An optional trailing dot used to let a line like
# "3000 Series" be misread as a new top-level entry numbered "3000.", which
# then hijacked the whole counter stack and exploded every number after it
# (e.g. jumping straight to "6001.", "6002.", ...). Requiring the dot means
# we only ever trust digits as a real section number when the source
# document actually punctuated it as one (e.g. "5.4." or "1.").
DECIMAL_PATTERN = re.compile(r"^\s*(\d+(?:\.\d+)*)\.\s+(.+?)\s*$")

# Legacy / old-style single-symbol markers, checked in this priority order when
# a line is NOT decimal. Each entry: (name, regex, symbol_kind)
LEGACY_PATTERNS = [
    ("upper_alpha", re.compile(r"^\s*([A-Z])\.\s+(.+?)\s*$"), "alpha_upper"),
    ("lower_roman", re.compile(r"^\s*([ivxlcdm]+)\.\s+(.+?)\s*$", re.IGNORECASE), "roman_lower"),
    ("lower_alpha", re.compile(r"^\s*([a-z])\.\s+(.+?)\s*$"), "alpha_lower"),
    ("number", re.compile(r"^\s*(\d+)\.\s+(.+?)\s*$"), "number"),
]

ROMAN_RE = re.compile(r"^(m{0,4}(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3}))$", re.IGNORECASE)


@dataclass
class TocEntry:
    """One line of the Table of Contents, normalized."""
    level: int                 # 1-based nesting depth (1 = top level)
    title: str                 # the text content, numbering stripped
    number: str = ""           # final decimal number, e.g. "5.4.1"
    source_number: str = ""    # decimal number as found in the ORIGINAL doc (empty if legacy/plain)
    raw: str = ""              # original raw line, for debugging/traceability
    restart_group: bool = False  # True for entries in a flat restarted list (see RESTART_LIST_TRIGGERS)
    is_trigger: bool = False     # True for the trigger heading itself (e.g. "List of Key Figures and Tables")


# ---------------------------------------------------------------------------
# PARSING
# ---------------------------------------------------------------------------

def _is_roman(token: str) -> bool:
    return bool(token) and bool(ROMAN_RE.match(token))


def _detect_line(line: str, indent_stack, level_hint=None, hint_reliable=True, context=None):
    """
    Detect the numbering marker (if any) on a single line and figure out its
    nesting level.

    Returns (level, title, source_number, marked) or None if the line is
    blank and should be skipped entirely. `source_number` is the ORIGINAL
    decimal number string (e.g. "5.4.1") when the line already used modern
    decimal numbering, or None for legacy/plain lines (those get a
    brand-new number assigned later by renumber()). `marked` is False only
    for lines whose level came purely from a fallback/soft-hint guess (see
    step 3) - used by parse_toc to decide how the NEXT line's fallback
    should behave.

    `level_hint`, when given, is a STRUCTURAL level read directly from the
    source .docx (see utils._structural_level). `hint_reliable` says which
    of two tiers it came from:
      - True:  a named style like "TOC 2"/"Heading 3" - a deliberate,
        trustworthy signal. Every line gets its own hint trusted as-is.
      - False: a bare outline/list level with no matching style. This is
        still useful as a STARTING point, but real documents sometimes give
        a run of plain, unnumbered bullet items (e.g. product/grade names
        with no section number of their own) a slightly different soft
        level on each line - ordinary Word list-continuation quirks, not
        real nesting. So a soft hint is only consulted for the FIRST
        unmarked line after a marked one; consecutive unmarked lines after
        that stay siblings of each other regardless of their own soft hint
        (see step 3 and parse_toc's context tracking).

    Detection strategy, checked in order:
      1. Decimal numbering e.g. "5.4.1.3.1. Packaged Food" -> level is the
         count of dot-separated segments (or a RELIABLE hint, if stronger),
         and the ORIGINAL number is kept as-is (Phase 1 only re-indents; it
         never renumbers numbers that are already correct decimal
         numbering). A trailing period is REQUIRED (see DECIMAL_PATTERN) so
         titles that merely start with a number, like "3000 Series" or
         "24/7 Support", are never mistaken for section numbers.
      2. Legacy single-symbol markers (A. / a. / i. / 1.) -> level is derived
         from a running "symbol-kind stack" (or a RELIABLE hint, if
         stronger): the first symbol kind seen is level 1, a new symbol kind
         encountered deeper is pushed as a new level, and returning to a
         previously-seen symbol kind pops back to that level (sibling
         item). These get a FRESH decimal number from renumber() (Phase 3:
         old format -> new format conversion).
      3. No marker in the text at all: a RELIABLE hint is trusted outright;
         otherwise real leading whitespace/tabs (plain pasted text only) or
         a SOFT hint sets the level for the first line of a run, and
         further fallback details are handled by parse_toc via `context`.
    """
    stripped = line.strip()
    if not stripped:
        return None

    # 1) Decimal numbering (already in, or close to, the target format).
    # A bare single integer ("1.", "2.") is ambiguous: it could be a
    # top-level modern number OR the innermost digit of a legacy
    # A./a./i./1. run. We only trust it as modern decimal when there's no
    # legacy symbol context open yet (indent_stack empty) or when it has
    # multiple dot-separated segments (unambiguously modern, e.g. "5.4.1"),
    # UNLESS a RELIABLE structural level_hint is available, in which case
    # ambiguity doesn't matter - the hint decides the level either way. A
    # soft (unreliable) hint is ignored here in favor of the text, which is
    # already unambiguous once we have a real number in hand.
    m = DECIMAL_PATTERN.match(line)
    if m:
        number_str, title = m.group(1), m.group(2)
        is_multi_segment = "." in number_str
        if level_hint is not None and hint_reliable:
            return level_hint, title, number_str, True
        if is_multi_segment or not indent_stack:
            level = number_str.count(".") + 1
            return level, title, number_str, True
        # else: bare integer inside an active legacy run -> fall through to
        # the legacy "number" pattern below so it nests under the run.

    # 2) Legacy single-symbol markers
    # NOTE: single lowercase letters that are also valid roman numerals
    # (i, v, x, l, c, d, m) are always treated as roman here (lower_roman is
    # checked before lower_alpha) - this matches the brief's canonical
    # A. / a. / i. / 1. example. A long run of plain a./b./c.../i. lettering
    # (rather than true roman numerals) is a known edge case this heuristic
    # will misread when there's no reliable level_hint to fall back on.
    for name, pattern, kind in LEGACY_PATTERNS:
        m = pattern.match(line)
        if not m:
            continue
        symbol, title = m.group(1), m.group(2)
        if level_hint is not None and hint_reliable:
            level = level_hint
        else:
            level = _level_from_symbol_stack(indent_stack, kind)
        return level, title.strip(), None, True

    # 3) No recognizable marker in the text at all (e.g. a plain Word heading
    # like "Executive Summary" with no typed-in number, or a bullet like
    # "3000 Series" with no number of its own).

    # 3a) A RELIABLE hint (named Heading/TOC style) is trusted unconditionally,
    # every single line - this is what correctly lets a real Word document's
    # Heading 1/2/3/4 structure produce proper nesting straight away.
    if level_hint is not None and hint_reliable:
        return level_hint, stripped, None, True

    # 3b) Otherwise (a SOFT hint, or no hint at all): if the PREVIOUS line was
    # also unmarked, this line is a continuation of that same flat run -
    # always its SIBLING, at the exact same level. This deliberately
    # overrides a soft hint/indentation for this case: real documents
    # sometimes carry inconsistent per-paragraph outline/list levels across
    # a run of plain bullet items (e.g. "3000 Series" / "5000 Series" /
    # "6000 Series" each ending up at a slightly different XML level due to
    # how they were typed/pasted in Word), which would otherwise make each
    # successive unmarked bullet nest one level deeper than the last
    # instead of staying siblings.
    prev_level, prev_marked = context if context else (None, True)
    if prev_level is not None and not prev_marked:
        return prev_level, stripped, None, False

    # 3c) First line of a new unmarked run: soft hint > real indentation >
    # one level deeper than the preceding marked line > level 1.
    if level_hint is not None:
        return level_hint, stripped, None, False

    leading_ws = len(line) - len(line.lstrip(" \t"))
    if leading_ws > 0:
        expanded = line[:leading_ws].expandtabs(4)
        level = (len(expanded) // 4) + 1
        return level, stripped, None, False

    if prev_level is None:
        level = 1
    elif prev_marked:
        level = prev_level + 1
    else:
        level = prev_level
    return level, stripped, None, False


def _level_from_symbol_stack(indent_stack, kind):
    """
    indent_stack is a list of (kind, level) pairs representing the currently
    "open" legacy levels, deepest last. Mutates indent_stack in place.
    """
    # If this kind is already open at some depth, we've returned to that level.
    for i, (k, lvl) in enumerate(indent_stack):
        if k == kind:
            del indent_stack[i + 1:]
            return lvl

    # New, deeper kind -> push a new level
    new_level = (indent_stack[-1][1] + 1) if indent_stack else 1
    indent_stack.append((kind, new_level))
    return new_level


def parse_toc(lines) -> list:
    """
    Parse raw TOC lines into TocEntry objects with a detected nesting
    `level`. `lines` accepts several shapes for flexibility:
      - a single newline-joined string
      - a list of plain text strings (level is inferred purely from the text)
      - a list of (text, structural_level_or_None) 2-tuples, treated as a
        RELIABLE hint (backwards-compatible shorthand)
      - a list of (text, structural_level_or_None, is_reliable) 3-tuples, as
        produced by utils.extract_lines_from_docx() - structural_level
        (from a Word Heading/TOC style, outline level, or list level) takes
        priority over text-based inference whenever it's present; see
        _detect_line's docstring for how is_reliable changes the handling.

    Numbers themselves are regenerated later by renumber() when there's no
    number already in the source text - here we only need level + title.
    """
    if isinstance(lines, str):
        lines = lines.splitlines()

    entries = []
    indent_stack = []  # tracks legacy symbol-kind levels across the whole doc
    prev_level = None
    prev_marked = True
    for item in lines:
        if isinstance(item, tuple):
            if len(item) == 3:
                raw_line, level_hint, hint_reliable = item
            else:
                raw_line, level_hint = item
                hint_reliable = True
        else:
            raw_line, level_hint, hint_reliable = item, None, True

        result = _detect_line(
            raw_line, indent_stack, level_hint=level_hint, hint_reliable=hint_reliable,
            context=(prev_level, prev_marked),
        )
        if result is None:
            continue
        level, title, source_number, marked = result
        title = re.sub(r"\s+", " ", title).strip()
        if not title:
            continue
        level = max(level, 1)
        entries.append(TocEntry(
            level=level,
            title=title,
            source_number=source_number or "",
            raw=raw_line,
        ))
        prev_level, prev_marked = level, marked
    return entries


# ---------------------------------------------------------------------------
# RENUMBERING - always produces clean decimal numbering: 1. / 1.1. / 5.4.1.3.1.
# ---------------------------------------------------------------------------

def renumber(entries: list) -> list:
    """
    Finalize the decimal number for every entry (mutates and returns the same
    list):

      - If the entry already had a decimal number in the source document
        (source_number set), that number is TRUSTED AS-IS (Phase 1: only
        indentation is normalized, real numbers are never rewritten) - and
        the running counter stack is synced to it so any later legacy/plain
        entries continue counting on from the right place. This also resets
        the level-normalization stack described below.
      - Otherwise (legacy symbols, structural-hint-derived, or plain/fallback
        lines) a fresh, contiguous decimal number is generated.

    For entries WITHOUT a source number, entry.level is an "absolute" value
    from parse_toc that isn't necessarily already relative to the running
    counters (e.g. two true siblings can carry the exact same absolute
    level while appearing at very different points in the counter stack).
    Comparing that absolute value against the ever-changing `len(counters)`
    one entry at a time - the old approach - breaks precisely on repeated
    identical levels: after the first sibling deepens the counters, the
    second sibling's IDENTICAL level no longer looks "too deep" relative to
    the now-longer counters, so it silently gets nested a level below its
    sibling instead of next to it.

    Instead we run a small level-stack normalizer (the same idea as the
    legacy A./a./i./1. symbol stack, generalized to plain numbers): each
    non-anchor entry's raw level is compared only to the PREVIOUS non-anchor
    entry's raw level - equal means "sibling, same depth", greater means
    "one level deeper", smaller means "pop back to the matching (or next
    shallower) open level". The resulting depth is then layered on top of
    whatever counters depth the most recent decimal anchor established.
    """
    counters = []   # final decimal counters actually used for numbering
    raw_stack = []  # stack of raw (pre-clamp) levels for non-anchor entries,
                     # reset every time a decimal-numbered anchor appears
    anchor_depth = 0  # counters depth at the most recent anchor
    restarted = False  # True once a RESTART_LIST_TRIGGERS heading is seen
    flat_counter = 0    # 1, 2, 3, ... used for entries after the trigger

    for entry in entries:
        if not restarted and any(p.match(entry.title) for p in RESTART_LIST_TRIGGERS):
            # This is the trigger heading itself (e.g. "List of Key Figures
            # and Tables") - it's a plain section title, not itself a list
            # item, so it gets no number. Everything after it restarts as a
            # flat, independently-numbered list.
            entry.is_trigger = True
            entry.number = ""
            restarted = True
            continue

        if restarted:
            flat_counter += 1
            entry.level = 1
            entry.number = str(flat_counter)
            entry.restart_group = True
            continue

        if entry.source_number:
            counters = [int(p) for p in entry.source_number.split(".")]
            entry.level = len(counters)
            entry.number = entry.source_number
            raw_stack = []
            anchor_depth = len(counters)
            continue

        raw_level = entry.level
        while raw_stack and raw_stack[-1] > raw_level:
            raw_stack.pop()
        if not (raw_stack and raw_stack[-1] == raw_level):
            raw_stack.append(raw_level)

        depth = anchor_depth + len(raw_stack)

        # Depth can still only increase by at most 1 beyond the current
        # counters (e.g. if the source skipped a level entirely); clamp and
        # keep raw_stack consistent with the clamped depth so later
        # comparisons stay correct.
        if depth > len(counters) + 1:
            depth = len(counters) + 1
            del raw_stack[depth - anchor_depth:]

        if depth > len(counters):
            counters.append(1)
        else:
            counters = counters[:depth]
            counters[-1] += 1

        entry.level = depth
        entry.number = ".".join(str(c) for c in counters)
    return entries


# ---------------------------------------------------------------------------
# RENDERING
# ---------------------------------------------------------------------------

def render(entries: list, fmt: str = "txt") -> str:
    """Render normalized entries to 'txt', 'md', or 'json'."""
    fmt = fmt.lower()
    if fmt == "json":
        return json.dumps([asdict(e) for e in entries], indent=2, ensure_ascii=False)

    if not entries:
        return ""

    # Indentation is RELATIVE to the shallowest level present in the MAIN
    # hierarchical part of the document. e.g. if a TOC excerpt starts at
    # "5.4." (an absolute level-2 number), that line still renders with zero
    # indent, and "5.4.1." gets one level of indent, matching the brief's
    # worked example. Entries in a restarted flat list (see
    # RESTART_LIST_TRIGGERS) and the trigger heading itself are excluded from
    # this calculation and always render at zero indent - they're a
    # standalone flat list, not part of the main hierarchy.
    main_levels = [e.level for e in entries if not e.restart_group and not e.is_trigger]
    base_level = min(main_levels) if main_levels else 1

    lines = []
    for e in entries:
        if e.is_trigger:
            # Plain section heading, no number of its own.
            lines.append(e.title)
        elif e.restart_group:
            number = f"{e.number}{RESTART_LIST_NUMBER_SUFFIX}"
            if fmt == "md":
                lines.append(f"{MD_BULLET_STYLE} {number}{RESTART_LIST_NUMBER_TITLE_SEP}{e.title}")
            else:
                lines.append(f"{number}{RESTART_LIST_NUMBER_TITLE_SEP}{e.title}")
        else:
            indent = INDENT_UNIT * (e.level - base_level)
            number = f"{e.number}{NUMBER_SUFFIX}"
            if fmt == "md":
                heading_marks = "#" * min(e.level + 1, 6)  # start at H2
                lines.append(f"{indent}{MD_BULLET_STYLE} {number}{NUMBER_TITLE_SEP}{e.title}")
            else:  # txt (default)
                lines.append(f"{indent}{number}{NUMBER_TITLE_SEP}{e.title}")

        if BLANK_LINE_BETWEEN_ENTRIES:
            lines.append("")
    text = "\n".join(lines).rstrip() + "\n"
    return text


# ---------------------------------------------------------------------------
# ONE-SHOT CONVENIENCE FUNCTION (used by app.py)
# ---------------------------------------------------------------------------

def process_text(raw_text: str, fmt: str = "txt"):
    """
    Full pipeline: raw TOC text -> parsed entries -> renumbered -> rendered.
    Returns (entries, rendered_string).
    """
    entries = parse_toc(raw_text)
    entries = renumber(entries)
    rendered = render(entries, fmt=fmt)
    return entries, rendered