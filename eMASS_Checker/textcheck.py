"""Text quality checks: offline spelling, built-in grammar/punctuation rules,
assessor language rules, and an optional local LanguageTool server.

Every check returns TextHit objects with character offsets into the text so the
GUI can highlight the span and build a proposed replacement. Nothing in this
module sends data off the computer. LanguageTool is reached only on 127.0.0.1.
"""
from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from functools import lru_cache

APP_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def user_dir() -> str:
    base = os.environ.get("APPDATA") or os.path.join(os.path.expanduser("~"), ".config")
    d = os.path.join(base, "eMASS_Checker")
    os.makedirs(d, exist_ok=True)
    return d


# --------------------------------------------------------------------------------------
# Result object
# --------------------------------------------------------------------------------------
@dataclass
class TextHit:
    rule: str            # rule ID, for example AG-OPEN-EXAMINE
    category: str        # Spelling, Grammar, Punctuation, Assessor Language, Terminology ...
    severity: str        # Error, Warning, Info
    start: int
    end: int
    message: str         # what is wrong
    why: str             # why it matters / source of the rule
    replacement: str | None = None   # replacement for text[start:end]; None = flag only
    safe: bool = False   # True = mechanical fix that "Apply Safe Fixes" may apply
    word: str | None = None  # for spelling hits (Add to dictionary)

    def apply(self, text: str) -> str:
        if self.replacement is None:
            return text
        return text[: self.start] + self.replacement + text[self.end:]


# --------------------------------------------------------------------------------------
# Rules config
# --------------------------------------------------------------------------------------
def bundled_rules_path() -> str:
    return os.path.join(APP_DIR, "config", "rules.json")


def user_rules_path() -> str:
    return os.path.join(user_dir(), "rules.json")


def load_rules() -> dict:
    path = user_rules_path() if os.path.exists(user_rules_path()) else bundled_rules_path()
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def user_dictionary_path() -> str:
    p = os.path.join(user_dir(), "user_dictionary.txt")
    if not os.path.exists(p):
        with open(p, "w", encoding="utf-8") as fh:
            fh.write("# One word per line. Words here are never flagged as misspelled.\n")
    return p


# --------------------------------------------------------------------------------------
# Spelling (offline, bundled SCOWL/Hunspell en_US word list)
# --------------------------------------------------------------------------------------
class Speller:
    LETTERS = "abcdefghijklmnopqrstuvwxyz"

    def __init__(self, rules: dict):
        with open(os.path.join(APP_DIR, "data", "en_US_words.txt"), encoding="utf-8") as fh:
            words = fh.read().split("\n")
        self.exact = set(words)
        self.lower = {w.lower() for w in words}
        self.extra = set()
        for w in rules.get("spelling_whitelist", []):
            self.extra.add(w.lower())
        for t in rules.get("canonical_terms", []):
            for part in re.split(r"[\s/]+", t):
                self.extra.add(part.lower())
        self.reload_user()

    def reload_user(self):
        self.user = set()
        try:
            with open(user_dictionary_path(), encoding="utf-8") as fh:
                for line in fh:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        self.user.add(line.lower())
        except OSError:
            pass

    def add_user_word(self, word: str):
        word = word.strip()
        if not word:
            return
        with open(user_dictionary_path(), "a", encoding="utf-8") as fh:
            fh.write(word + "\n")
        self.user.add(word.lower())
        self.known.cache_clear()

    @lru_cache(maxsize=50000)
    def known(self, w: str) -> bool:
        lw = w.lower()
        if w in self.exact or lw in self.lower or lw in self.extra or lw in self.user:
            return True
        for suf in ("'s", "’s", "s'", "’"):
            if lw.endswith(suf) and self.known(w[: -len(suf)]):
                return True
        if "-" in w:
            parts = [p for p in w.split("-") if p]
            return bool(parts) and all(self.known(p) or len(p) <= 2 for p in parts)
        return False

    def _edits1(self, w):
        splits = [(w[:i], w[i:]) for i in range(len(w) + 1)]
        deletes = [a + b[1:] for a, b in splits if b]
        transposes = [a + b[1] + b[0] + b[2:] for a, b in splits if len(b) > 1]
        replaces = [a + c + b[1:] for a, b in splits if b for c in self.LETTERS]
        inserts = [a + c + b for a, b in splits for c in self.LETTERS]
        return transposes, deletes, replaces, inserts

    @lru_cache(maxsize=5000)
    def suggest(self, word: str, n: int = 5) -> tuple:
        w = word.lower()
        pool = self.lower | self.extra | self.user
        found, seen = [], {}
        t, d, r, i = self._edits1(w)
        for rank, group in enumerate((t, r, d, i)):
            for c in group:
                if c in pool and c not in seen:
                    seen[c] = rank
                    found.append(c)
        if not found and len(w) <= 14:
            for e1 in set(t + d + r + i):
                t2, d2, r2, i2 = self._edits1(e1)
                for c in t2 + d2 + r2:
                    if c in pool and c not in seen:
                        seen[c] = 4
                        found.append(c)
                if len(found) > 25:
                    break
        # rank: transposition first, then same first letter, similar length
        found.sort(key=lambda c: (seen[c], c[0] != w[0], abs(len(c) - len(w)), c))
        out = []
        for c in found[:n]:
            if word[:1].isupper():
                c = c[:1].upper() + c[1:]
            out.append(c)
        return tuple(out)


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------
PROTECT_PATTERNS = [
    r"https?://\S+", r"\bwww\.\S+", r"\b[\w.+-]+@[\w-]+\.[\w.-]+\b",
    r"\b[A-Z]{2}\.L[12]-\d+\.\d+\.\d+(?:\[[a-z]\])*", r"\b\d+\.\d+\.\d+(?:\[[a-z]\])*",
    r"\b[\w\-()]+\.(?:pdf|docx?|xlsx?|xlsm|pptx?|png|jpe?g|gif|bmp|csv|txt|msg|eml|zip|vsdx?|json|xml|html?|exe|msi|ps1)\b",
    r"\b[A-Za-z]:\\[^\s;]+", r"\\\\[^\s;]+",
    r"\b(?:e\.g|i\.e|etc|vs|approx|No|Inc|Corp|Ltd|LLC|Co|Dept|Sec|Rev|Ver|v)\.",
    r"\b\d+(?:\.\d+)+\b",
]
ABBREV_BEFORE_PERIOD = re.compile(r"(?:\b(?:e\.g|i\.e|etc|vs|approx|No|Inc|Corp|Ltd|Co|Dept|Sec|Rev|Ver|Mr|Ms|Mrs|Dr|St|U\.S)|\b[A-Z])\.$")
LOWER_START_BRANDS = re.compile(r"^(?:eMASS|iOS|iPhone|iPad|macOS|eDiscovery|eBay|pfSense|vCenter|vSphere|iDRAC)\b")
WORD_RE = re.compile(r"[A-Za-z][A-Za-z'’\-]*[A-Za-z]|[A-Za-z]")


def protected_spans(text: str):
    spans = []
    for p in PROTECT_PATTERNS:
        for m in re.finditer(p, text):
            spans.append((m.start(), m.end()))
    return spans


def in_spans(pos: int, spans) -> bool:
    return any(a <= pos < b for a, b in spans)


def sentence_bounds(text: str):
    """Yield (start, end) for each sentence (end exclusive, includes terminal punctuation)."""
    bounds, start, i, n = [], 0, 0, len(text)
    while i < n:
        ch = text[i]
        if ch in ".!?" and (i + 1 == n or text[i + 1] in " \n\r\t\"')"):
            if ch == "." and ABBREV_BEFORE_PERIOD.search(text[max(0, i - 6): i + 1]):
                i += 1
                continue
            j = i + 1
            while j < n and text[j] in "\"')":
                j += 1
            bounds.append((start, j))
            while j < n and text[j] in " \n\r\t":
                j += 1
            start = j
            i = j
            continue
        if ch == "\n" and i + 1 < n and text[i + 1] == "\n":
            if text[start:i].strip():
                bounds.append((start, i))
            start = i + 1
        i += 1
    if text[start:].strip():
        bounds.append((start, n))
    return [(a, b) for a, b in bounds if text[a:b].strip()]


def sentence_containing(text: str, pos: int):
    for a, b in sentence_bounds(text):
        if a <= pos < b:
            return a, b
    return 0, len(text)


def removal_hit_for_sentence(text, a, b):
    """Return (start, end) that removes a sentence and the whitespace that follows it."""
    e = b
    while e < len(text) and text[e] in " \t":
        e += 1
    s = a
    if e >= len(text):  # last sentence: also remove the space before it
        while s > 0 and text[s - 1] in " \t":
            s -= 1
    return s, e


def match_case(src: str, repl: str) -> str:
    if src[:1].isupper() and repl[:1].islower():
        return repl[:1].upper() + repl[1:]
    return repl


def phrase_regex(phrase: str) -> str:
    """Turn a plain phrase into a word-bounded, whitespace-tolerant regex."""
    esc = r"\s+".join(re.escape(p) for p in phrase.split())
    left = r"\b" if re.match(r"\w", phrase) else ""
    right = r"\b" if re.search(r"\w$", phrase) else ""
    return left + esc + right


# --------------------------------------------------------------------------------------
# Built-in checker
# --------------------------------------------------------------------------------------
GUIDE = "Assessor reporting guideline"
EMASS = "eMASS template instruction"


class TextChecker:
    def __init__(self, rules: dict | None = None):
        self.rules = rules or load_rules()
        self.speller = Speller(self.rules)
        self.lt = None  # LanguageToolClient when enabled
        r = self.rules
        self._prohib_remove = [(p, re.compile(phrase_regex(p), re.I)) for p in r["prohibited_remove_sentence"]]
        self._prohib_flag = [re.compile(p, re.I) for p in r["prohibited_flag"]]
        self._status_cs = [(p, re.compile(phrase_regex(p))) for p in r["status_terms_flag_case_sensitive"]]
        self._status_ci = [(p, re.compile(p if "\\" in p else phrase_regex(p), re.I)) for p in r["status_terms_flag"]]
        self._qual = [(p, re.compile(phrase_regex(p), re.I)) for p in r["unsupported_qualifiers_flag"]]
        self._filler = [(p, re.compile(r"\b" + re.escape(p) + r"\b\s?", re.I)) for p in r["filler_qualifiers_remove"]]
        self._time = [re.compile(p, re.I) for p in r["time_reference_patterns"]]
        self._file = re.compile(r["filename_pattern"], re.I)
        self._evid = re.compile(r["evidence_id_pattern"]) if r.get("evidence_id_pattern") else None
        self._actor_replace = sorted(r["actor_variants_replace"], key=len, reverse=True)
        self._actor_flag = [re.compile(p) for p in r["actor_variants_flag"]]
        terms = sorted(r["canonical_terms"], key=len, reverse=True)
        self._terms = [(t, re.compile(r"(?<![\w&])" + r"\s+".join(re.escape(x) for x in t.split()) + r"(?![\w&])", re.I)) for t in terms]
        self._variants = [(k, v, re.compile(r"(?<![\w&])" + re.escape(k).replace(r"\ ", r"\s+") + r"(?![\w&])", re.I))
                          for k, v in sorted(r["term_variants"].items(), key=lambda kv: -len(kv[0]))]
        self._overcap = [(p, re.compile(r"(?<=[a-z,;:] )" + re.escape(p) + r"\b")) for p in r["over_capitalized_phrases"]]
        self._contr = [(k, v, re.compile(r"\b" + re.escape(k).replace("'", "['’]") + r"\b", re.I)) for k, v in r["contractions"].items()]

    # ------------------------------------------------------------------
    def check(self, text: str, field: str, opts: dict | None = None) -> list[TextHit]:
        """field: Examine, Test, Overall Comments, Findings, Executive Summary, ESP Name, Assessed By"""
        opts = opts or {}
        if text is None:
            return []
        text = str(text)
        hits: list[TextHit] = []
        if not text.strip():
            return hits
        narrative = field in ("Examine", "Test", "Overall Comments", "Findings", "Executive Summary")
        hits += self._whitespace_chars(text)
        if field in ("Examine", "Test", "Overall Comments"):
            hits += self._opener(text, field)
        if narrative:
            hits += self._actor(text)
            hits += self._prohibited(text)
            hits += self._time_refs(text)
            hits += self._dashes(text)
            hits += self._terminology(text)
            if field != "Executive Summary":
                hits += self._filenames(text)
            hits += self._punctuation(text, field)
            hits += self._grammar(text, field)
            if opts.get("spelling", True):
                hits += self._spelling(text)
            if opts.get("languagetool") and self.lt is not None:
                hits += self.lt.check(text, self.speller)
        elif field == "Assessed By":
            hits += self._actor(text, only_prohibited=True)
        return dedupe(hits)

    # ------------------------------------------------------------------
    def _whitespace_chars(self, t):
        h = []
        if t != t.strip():
            lead = len(t) - len(t.lstrip())
            trail = len(t.rstrip())
            if lead:
                h.append(TextHit("TX-WS-EDGE", "Punctuation", "Info", 0, lead, "Leading whitespace.", "Stray spaces or line breaks at the start of a cell carry into eMASS.", "", True))
            if trail < len(t):
                h.append(TextHit("TX-WS-EDGE", "Punctuation", "Info", trail, len(t), "Trailing whitespace.", "Stray spaces or line breaks at the end of a cell carry into eMASS.", "", True))
        for m in re.finditer(r"(?<=\S)[ \t]{2,}(?=\S)", t):
            h.append(TextHit("TX-WS-DOUBLE", "Punctuation", "Info", m.start(), m.end(), "Multiple consecutive spaces.", "Use a single space between words and after punctuation.", " ", True))
        for m in re.finditer(r"[   ​‌‍﻿]", t):
            rep = "" if m.group() in "​‌‍﻿" else " "
            h.append(TextHit("TX-CHAR-HIDDEN", "Punctuation", "Warning", m.start(), m.end(), "Hidden or non-breaking space character.", "Invisible characters pasted from Word or email can cause search and import problems.", rep, True))
        for m in re.finditer(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", t):
            h.append(TextHit("TX-CHAR-CTRL", "Punctuation", "Error", m.start(), m.end(), "Control character in text.", "Control characters are invalid in Excel XML and may break the eMASS import.", "", True))
        return h

    # ------------------------------------------------------------------
    ACTOR_START = re.compile(
        r"^\s*(?:the\s+)?(?:assigned\s+certified\s+assessors?|certified\s+assessors?|assessment\s+team|assessing\s+team|audit\s+team|"
        r"c3pao(?:\s+assessment)?(?:\s+team)?|lead\s+assessor|assessors?|ccas?|team|we)\s+(\w+)", re.I)
    LEADING_VERB = re.compile(r"^\s*(reviewed|examined|validated|confirmed|verified|tested|observed|inspected|evaluated|analyzed|assessed|noted|determined|found)\b\s*", re.I)

    def _opener(self, t, field):
        allowed = self.rules["openers"].get(field, [])
        if not allowed or any(t.lstrip().startswith(a) for a in allowed):
            return []
        lead = len(t) - len(t.lstrip())
        target_verbs = {"Examine": ("reviewed", "examined"), "Test": ("validated",), "Overall Comments": ("confirmed",)}[field]
        default_verb = target_verbs[0]
        why = f"{GUIDE}: {field} statements begin with " + " or ".join(f'"{a}"' for a in allowed) + ". Use the same actor phrase throughout."
        msg = f'{field} does not begin with the required opener ({" / ".join(allowed)}).'
        # case-only difference, e.g. "the assessment team reviewed"
        for a in allowed:
            if t.lstrip().lower().startswith(a.lower()):
                return [TextHit(f"AG-OPEN-{field.upper().replace(' ', '')}", "Assessor Language", "Warning", lead, lead + len(a), msg + " (capitalization)", why, a, True)]
        m = self.ACTOR_START.match(t)
        if m:
            verb = m.group(1).lower()
            if field == "Test" and verb in ("reviewed", "examined"):
                return [TextHit("AG-TEST-NOTTEST", "Assessor Language", "Warning", lead, m.end(), "Test narrative describes a document review, not a test.",
                                f"{GUIDE}: Test statements begin with \"The assessment team validated\" only where supported by an actual test. Move review content to Examine or confirm a live test occurred.", None)]
            new_verb = verb if verb in target_verbs else default_verb
            rep = f"The assessment team {new_verb}"
            note = "" if verb in target_verbs else f' (changes the verb "{verb}" to "{new_verb}", confirm this matches what occurred)'
            return [TextHit(f"AG-OPEN-{field.upper().replace(' ', '')}", "Assessor Language", "Warning", lead, m.end(), msg + note, why, rep, False)]
        m = self.LEADING_VERB.match(t)
        if m:
            verb = m.group(1).lower()
            new_verb = verb if verb in target_verbs else default_verb
            return [TextHit(f"AG-OPEN-{field.upper().replace(' ', '')}", "Assessor Language", "Warning", lead, m.end(), msg, why, f"The assessment team {new_verb} ", False)]
        return [TextHit(f"AG-OPEN-{field.upper().replace(' ', '')}", "Assessor Language", "Warning", lead, min(len(t), lead + 40), msg + " No automatic rewrite offered because the sentence structure must change; edit the proposed text manually.", why, None)]

    # ------------------------------------------------------------------
    def _actor(self, t, only_prohibited=False):
        h = []
        actor = self.rules.get("actor_phrase", "the assessment team")
        taken = []
        for v in self._actor_replace:
            if only_prohibited and "assigned certified" not in v:
                continue
            for m in re.finditer(phrase_regex(v), t, re.I):
                if any(a <= m.start() < b for a, b in taken):
                    continue
                taken.append((m.start(), m.end()))
                rep = match_case(m.group(), actor)
                is_proh = "assigned certified" in v.lower()
                h.append(TextHit("AG-ACTOR-PROHIBITED" if is_proh else "AG-ACTOR-CONSISTENT", "Assessor Language", "Error" if is_proh else "Warning",
                                 m.start(), m.end(), f'"{m.group()}" used as the actor phrase.',
                                 f'{GUIDE}: use "{actor}" consistently.' + (' "The assigned certified assessor" is prohibited in eMASS deliverables.' if is_proh else ""),
                                 rep, is_proh))
        if only_prohibited:
            return h
        for rx in self._actor_flag:
            for m in rx.finditer(t):
                if m.group() == "us" and t[max(0, m.start() - 1):m.start()] == ".":
                    continue
                h.append(TextHit("AG-ACTOR-PRONOUN", "Assessor Language", "Warning", m.start(), m.end(), f'First-person or informal actor "{m.group()}".',
                                 f'{GUIDE}: use "{actor}" as the single actor phrase throughout. Rewrite the sentence manually.', None))
        return h

    # ------------------------------------------------------------------
    def _prohibited(self, t):
        h = []
        for p, rx in self._prohib_remove:
            for m in rx.finditer(t):
                a, b = sentence_containing(t, m.start())
                why = f"{GUIDE}: remove advisory, recommendation, awareness, best-practice, and continuous-improvement language from eMASS narratives."
                comma = max(t.rfind(sep, a, m.start()) for sep in (",", ";", "\u2014", " \u2013 ", " -- "))
                sent_end = b
                while sent_end > a and t[sent_end - 1] in " \t\n\"')":
                    sent_end -= 1
                term = t[sent_end - 1] if sent_end > a and t[sent_end - 1] in ".!?" else ""
                while comma > a and t[comma - 1] in " \t":
                    comma -= 1
                if comma > a and len(t[a:comma].split()) >= 4:
                    # remove only the trailing clause: ", which is a best practice." -> "."
                    h.append(TextHit("AG-ADVISORY", "Assessor Language", "Error", comma, sent_end, f'Advisory or recommendation language: "{m.group()}". Proposed fix removes the clause.',
                                     why, term or ".", False))
                    continue
                s_, e_ = removal_hit_for_sentence(t, a, b)
                if not (t[:s_] + t[e_:]).strip():
                    h.append(TextHit("AG-ADVISORY", "Assessor Language", "Error", m.start(), m.end(), f'Advisory or recommendation language: "{m.group()}". The whole narrative is advisory; rewrite manually.', why, None))
                    continue
                h.append(TextHit("AG-ADVISORY", "Assessor Language", "Error", s_, e_, f'Advisory or recommendation language: "{m.group()}". Proposed fix removes the sentence.', why, "", False))
        for rx in self._prohib_flag:
            for m in rx.finditer(t):
                h.append(TextHit("AG-ADVISORY-REVIEW", "Assessor Language", "Warning", m.start(), m.end(), f'Possible advisory language: "{m.group()}".',
                                 f"{GUIDE}: narratives state facts only. Confirm this is quoted OSC policy language or reword (prohibited terms must be reworded even inside quoted policy text).", None))
        for p, rx in self._status_cs:
            for m in rx.finditer(t):
                h.append(TextHit("AG-STATUS-TERM", "Assessor Language", "Error", m.start(), m.end(), f'Non-permitted status term "{m.group()}".',
                                 f"{GUIDE}: PARTIAL and PENDING are prohibited. Formal results are MET, NOT MET, or N/A only. Assessor determination required.", None))
        for p, rx in self._status_ci:
            for m in rx.finditer(t):
                h.append(TextHit("AG-TDEE", "Assessor Language", "Error", m.start(), m.end(), f'"{m.group()}" used as a result category.',
                                 f"{GUIDE}: do not create a separate formal category for Temporary Deficiency or Enduring Exception, and do not write \"MET via TD/EE\". Use the formal MET / NOT MET / N/A result with factual rationale.", None))
        for p, rx in self._qual:
            for m in rx.finditer(t):
                h.append(TextHit("AG-QUALIFIER", "Assessor Language", "Warning", m.start(), m.end(), f'Unsupported qualifier "{m.group()}".',
                                 f"{GUIDE}: remove unsupported phrases such as \"appears,\" \"seems,\" or \"likely.\" State what the evidence shows; rewrite manually so the claim stays accurate.", None))
        for p, rx in self._filler:
            for m in rx.finditer(t):
                h.append(TextHit("AG-FILLER", "Assessor Language", "Info", m.start(), m.end(), f'Unsupported intensifier "{m.group().strip()}".',
                                 f"{GUIDE}: remove unsupported qualifiers. Removing the word does not change the factual claim.", "", False))
        return h

    # ------------------------------------------------------------------
    def _time_refs(self, t):
        h, taken = [], []
        rep_default = self.rules.get("time_reference_replacement", "during the assessment session")
        for i, rx in enumerate(self._time):
            for m in rx.finditer(t):
                if any(a <= m.start() < b for a, b in taken):
                    continue
                taken.append((m.start(), m.end()))
                rep = match_case(m.group(), rep_default) if i == 0 else None
                h.append(TextHit("AG-TIMEREF", "Assessor Language", "Warning", m.start(), m.end(), f'Session day, date, or time reference: "{m.group()}".',
                                 f'{GUIDE}: do not reference specific session days or times; use "the assessment session" or "the walkthrough" when timing context is necessary.', rep, False))
        return h

    def _dashes(self, t):
        h = []
        for m in re.finditer(r"\s*—\s*|\s+–\s+|\s+--\s+", t):
            h.append(TextHit("AG-EMDASH", "Punctuation", "Warning", m.start(), m.end(), "Em dash (or dash used as one).",
                             f"{GUIDE}: no em dashes in deliverables. Proposed fix uses a comma; adjust if a period reads better.", ", ", False))
        return h

    def _filenames(self, t):
        h = []
        for m in self._file.finditer(t):
            h.append(TextHit("AG-FILENAME", "Assessor Language", "Warning", m.start(), m.end(), f'Specific filename cited: "{m.group()}".',
                             f'{GUIDE}: narrative fields reference "the associated evidence" generically; traceability lives in the Artifacts column.', None))
        if self._evid:
            for m in self._evid.finditer(t):
                h.append(TextHit("AG-EVIDENCE-ID", "Assessor Language", "Warning", m.start(), m.end(), f'Evidence request ID cited: "{m.group()}".',
                                 f'{GUIDE}: narrative fields should not cite evidence request IDs.', None))
        return h

    # ------------------------------------------------------------------
    def _terminology(self, t):
        h, taken = [], []
        prot = [(a, b) for a, b in protected_spans(t)]
        for k, v, rx in self._variants:
            for m in rx.finditer(t):
                if m.group() == v or in_spans(m.start(), prot) or any(a <= m.start() < b for a, b in taken):
                    continue
                taken.append((m.start(), m.end()))
                h.append(TextHit("AG-TERM", "Terminology", "Warning", m.start(), m.end(), f'Non-standard term "{m.group()}"; use "{v}".',
                                 f"{GUIDE}: use exact product, program, and template terminology.", v, True))
        for term, rx in self._terms:
            for m in rx.finditer(t):
                if m.group() == term or in_spans(m.start(), prot) or any(a <= m.start() < b for a, b in taken):
                    continue
                taken.append((m.start(), m.end()))
                if term.lower() in ("it", "duo", "san", "ad", "gcc") and m.group().islower():
                    continue  # ordinary English word, not the product
                h.append(TextHit("AG-TERM-CASE", "Terminology", "Warning", m.start(), m.end(), f'Inconsistent capitalization "{m.group()}"; use "{term}".',
                                 f"{GUIDE}: remove inconsistent capitalization and use exact product, program, and template terminology.", term, True))
        for p, rx in self._overcap:
            for m in rx.finditer(t):
                rep = m.group().lower()
                h.append(TextHit("AG-CAPS", "Terminology", "Info", m.start(), m.end(), f'"{m.group()}" capitalized mid-sentence.',
                                 f"{GUIDE}: remove inconsistent capitalization; use sentence case.", rep, True))
        acr = {x.upper() for x in self.rules.get("spelling_whitelist", [])} | {x for x in self.rules.get("canonical_terms", []) if x.isupper()}
        acr |= {"MET", "NOT", "N/A", "OSC", "SSP", "CUI", "NA"}
        for m in re.finditer(r"\b(?:[A-Z]{4,}\b[ ,]*){1,}", t):
            words = [w for w in re.findall(r"[A-Z]{4,}", m.group()) if w not in acr]
            if len(words) >= 2 or (words and len(words[0]) >= 7):
                h.append(TextHit("AG-ALLCAPS", "Terminology", "Info", m.start(), m.end(), f'All-caps text "{m.group().strip()}".',
                                 f"{GUIDE}: proper sentence case capitalization throughout; no all-caps.", None))
        return h

    # ------------------------------------------------------------------
    def _punctuation(self, t, field):
        h = []
        prot = protected_spans(t)
        for m in re.finditer(r"[ \t]+([,.;:!?])(?![\w.])", t):
            if in_spans(m.start(1), prot):
                continue
            h.append(TextHit("TX-SPACE-BEFORE", "Punctuation", "Info", m.start(), m.end(), "Space before punctuation.", "Punctuation attaches to the preceding word.", m.group(1), True))
        for m in re.finditer(r"([,;:])(?=[A-Za-z])", t):
            if in_spans(m.start(), prot) or re.match(r"\d", t[m.start() + 1:m.start() + 2] or "") or t[max(0, m.start() - 4):m.start()].lower().endswith("http"):
                continue
            h.append(TextHit("TX-SPACE-AFTER", "Punctuation", "Info", m.start(), m.end(), f'Missing space after "{m.group(1)}".', "Put one space after commas, semicolons, and colons.", m.group(1) + " ", True))
        for m in re.finditer(r"(?<=[a-z]{2})\.(?=[A-Z][a-z])", t):
            if in_spans(m.start(), prot):
                continue
            h.append(TextHit("TX-SPACE-AFTER", "Punctuation", "Info", m.start(), m.end(), "Missing space after period.", "Sentences are separated by a period and a space.", ". ", True))
        for m in re.finditer(r"([,;:!?])\1+|(?<!\.)\.\.(?!\.)|[,;:]\.|\.,|,;|;,", t):
            if in_spans(m.start(), prot):
                continue
            rep = m.group()[-1] if m.group()[-1] == "." else m.group()[0]
            h.append(TextHit("TX-DOUBLE-PUNCT", "Punctuation", "Warning", m.start(), m.end(), f'Doubled or conflicting punctuation "{m.group()}".', "Use a single punctuation mark.", rep, True))
        opens, closes = t.count("("), t.count(")")
        if opens != closes:
            pos = t.find("(") if opens > closes else t.rfind(")")
            h.append(TextHit("TX-PAREN", "Punctuation", "Warning", max(pos, 0), max(pos, 0) + 1, "Unbalanced parentheses.", "Every opening parenthesis needs a closing one.", None))
        if t.count('"') % 2:
            pos = t.find('"')
            h.append(TextHit("TX-QUOTE", "Punctuation", "Warning", pos, pos + 1, "Unbalanced quotation marks.", "Quoted text needs opening and closing marks.", None))
        if t.count("“") != t.count("”"):
            pos = max(t.find("“"), t.find("”"))
            h.append(TextHit("TX-QUOTE", "Punctuation", "Warning", pos, pos + 1, "Unbalanced curly quotation marks.", "Quoted text needs opening and closing marks.", None))
        st = t.rstrip()
        if st and st[-1] not in ".!?\"')”" and len(st.split()) > 3:
            h.append(TextHit("TX-END-PUNCT", "Punctuation", "Info", len(st), len(st), "Narrative does not end with terminal punctuation.", "Complete sentences end with a period.", ".", True))
        return h

    # ------------------------------------------------------------------
    VOWEL_SOUND_EXCEPT_CONS = re.compile(r"^(?:uni|use|usa|usu|uti|ura|ure|uro|eu|ewe|one|once|ubiq|uk\b|u\b)", re.I)
    CONS_SOUND_EXCEPT_VOWEL = re.compile(r"^(?:hour|honest|honor|honour|heir|herb\b)", re.I)

    def _article_ok(self, art, word):
        w = word.strip("\"'(")
        if not w:
            return True
        if w.isupper() and len(w) > 1 and not re.match(r"^(?:NIST|NASA|FIPS|SIEM|CUI)$", w):  # spoken as letters
            vowel = w[0] in "AEFHILMNORSX"
            if w.startswith(("CUI",)):
                vowel = False
        elif re.match(r"^\d", w):
            vowel = bool(re.match(r"^(?:8|11|18|80+\b)", w))
        else:
            vowel = w[0].lower() in "aeiou"
            if vowel and self.VOWEL_SOUND_EXCEPT_CONS.match(w):
                vowel = False
            if not vowel and self.CONS_SOUND_EXCEPT_VOWEL.match(w):
                vowel = True
        return (art.lower() == "an") == vowel

    def _grammar(self, t, field):
        h = []
        prot = protected_spans(t)
        for m in re.finditer(r"\b(\w+)(\s+)(\1)\b", t, re.I):
            w = m.group(1).lower()
            if w in ("that", "had", "is") or w.isdigit():
                continue
            h.append(TextHit("TX-REPEAT", "Grammar", "Warning", m.start(2), m.end(3), f'Repeated word "{m.group(1)} {m.group(3)}".', "A word appears twice in a row.", "", True))
        for m in re.finditer(r"\b(a|an|A|An)\s+([\w\"'(][\w-]*)", t):
            if in_spans(m.start(2), prot) or m.group(2).lower() in ("n/a",):
                continue
            if not self._article_ok(m.group(1), m.group(2)):
                rep = ("an" if m.group(1).lower() == "a" else "a")
                rep = match_case(m.group(1), rep)
                h.append(TextHit("TX-ARTICLE", "Grammar", "Warning", m.start(1), m.end(1), f'Check article: "{m.group(1)} {m.group(2)}".', 'Use "an" before a vowel sound and "a" before a consonant sound.', rep, False))
        for m in re.finditer(r"(?<![\w.])i(?=\s)", t):
            h.append(TextHit("TX-LOWER-I", "Grammar", "Warning", m.start(), m.end(), 'Lowercase "i" as a pronoun.', 'The pronoun "I" is capitalized (and first person is not used in eMASS narratives).', "I", True))
        # sentence capitalization
        for a, b in sentence_bounds(t):
            seg = t[a:b]
            lead = len(seg) - len(seg.lstrip(" \t\n\"'(“"))
            p = a + lead
            if p < len(t) and t[p].islower() and not in_spans(p, prot) and not LOWER_START_BRANDS.match(t[p:]):
                if field == "Findings" and p == 0:
                    pass
                h.append(TextHit("TX-SENT-CAP", "Grammar", "Warning", p, p + 1, "Sentence starts with a lowercase letter.", "Sentences begin with a capital letter.", t[p].upper(), True))
            nwords = len(seg.split())
            if nwords > self.rules.get("max_sentence_words", 60):
                h.append(TextHit("TX-LONG-SENT", "Grammar", "Info", a, b, f"Long sentence ({nwords} words).", "Very long sentences are hard to read and often hide run-ons; consider splitting.", None))
        for k, v, rx in self._contr:
            for m in rx.finditer(t):
                h.append(TextHit("TX-CONTRACTION", "Grammar", "Info", m.start(), m.end(), f'Contraction "{m.group()}".', "Formal assessment narratives avoid contractions.", match_case(m.group(), v), True))
        return h

    # ------------------------------------------------------------------
    def _spelling(self, t):
        h = []
        prot = protected_spans(t)
        sent_starts = {a + (len(t[a:b]) - len(t[a:b].lstrip(" \t\n\"'(“"))) for a, b in sentence_bounds(t)}
        for m in WORD_RE.finditer(t):
            w = m.group()
            if in_spans(m.start(), prot) or len(w) < 3:
                continue
            core = w.strip("'’-")
            if not core or core.isupper() or any(c.isdigit() for c in core):
                continue
            if re.search(r"[a-z][A-Z]", core):  # camelCase product names
                continue
            if self.speller.known(core):
                continue
            capital = core[0].isupper() and m.start() not in sent_starts
            sugg = self.speller.suggest(core)
            if capital and not sugg:
                continue  # unknown proper noun (person or company name) with no close match
            sev = "Info" if capital else "Warning"
            rep = sugg[0] if sugg else None
            s = m.start() + w.find(core)
            h.append(TextHit("TX-SPELL", "Spelling", sev, s, s + len(core), f'Possible misspelling "{core}".' + (f" Suggestions: {', '.join(sugg)}." if sugg else " No suggestion found."),
                             "Word not found in the offline dictionary or your custom dictionary. Use Add to Dictionary for valid names and terms.", rep, False, word=core))
        return h


def dedupe(hits):
    """Drop exact duplicates and, within one rule, hits nested inside a longer hit."""
    out = []
    for x in sorted(hits, key=lambda z: (z.rule, z.start, -(z.end - z.start))):
        if any(y.rule == x.rule and y.start <= x.start and x.end <= y.end and (y.end > y.start or x.start == y.start) for y in out):
            continue
        out.append(x)
    return sorted(out, key=lambda z: (z.start, z.rule))


# --------------------------------------------------------------------------------------
# LanguageTool (local server only)
# --------------------------------------------------------------------------------------
LT_DOWNLOAD_URL = "https://languagetool.org/download/LanguageTool-stable.zip"


def find_lt_jar(folder: str) -> str | None:
    if not folder or not os.path.isdir(folder):
        return None
    for root, _dirs, files in os.walk(folder):
        if "languagetool-server.jar" in files:
            return os.path.join(root, "languagetool-server.jar")
    return None


class LanguageToolClient:
    """Starts (or connects to) a LanguageTool HTTP server on 127.0.0.1 and checks text.
    Requests never leave the computer: the URL is forced to localhost and system proxies are bypassed."""

    def __init__(self, folder: str, port: int = 8081, java: str = "java", disabled_rules=None, log=print):
        self.folder, self.port, self.java = folder, int(port), java or "java"
        self.disabled = list(disabled_rules or [])
        self.proc = None
        self.log = log
        self.base = f"http://127.0.0.1:{self.port}/v2"
        self.opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))

    def _alive(self) -> bool:
        try:
            with self.opener.open(self.base + "/languages", timeout=2) as r:
                return r.status == 200
        except Exception:
            return False

    def start(self, wait: int = 90) -> None:
        if self._alive():
            return
        jar = find_lt_jar(self.folder)
        if not jar:
            raise RuntimeError("languagetool-server.jar was not found. Set the LanguageTool folder in Settings or use Download LanguageTool.")
        flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
        self.proc = subprocess.Popen([self.java, "-Xmx1g", "-cp", jar, "org.languagetool.server.HTTPServer", "--port", str(self.port)],
                                     cwd=os.path.dirname(jar), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=flags)
        t0 = time.time()
        while time.time() - t0 < wait:
            if self.proc.poll() is not None:
                raise RuntimeError("LanguageTool server exited. Confirm Java 17 or newer is installed (run 'java -version').")
            if self._alive():
                return
            time.sleep(1)
        raise RuntimeError("LanguageTool server did not start in time.")

    def stop(self):
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()
        self.proc = None

    def check(self, text: str, speller: Speller | None = None) -> list[TextHit]:
        data = urllib.parse.urlencode({"language": "en-US", "text": text, "disabledRules": ",".join(self.disabled)}).encode()
        try:
            with self.opener.open(urllib.request.Request(self.base + "/check", data=data), timeout=60) as r:
                res = json.loads(r.read().decode("utf-8"))
        except Exception as e:  # noqa: BLE001
            return [TextHit("LT-UNAVAILABLE", "Grammar", "Info", 0, 0, f"LanguageTool check failed: {e}", "Local grammar engine unavailable for this cell.", None)]
        out = []
        for m in res.get("matches", []):
            off, ln = m.get("offset", 0), m.get("length", 0)
            frag = text[off:off + ln]
            if speller and m.get("rule", {}).get("issueType") == "misspelling" and speller.known(frag):
                continue
            reps = [r["value"] for r in m.get("replacements", [])[:5]]
            cat = m.get("rule", {}).get("category", {}).get("name", "Grammar")
            category = "Punctuation" if "punct" in cat.lower() else ("Spelling" if "typo" in cat.lower() or "spell" in cat.lower() else "Grammar")
            msg = m.get("message", "Grammar issue.")
            if reps:
                msg += f" Suggestions: {', '.join(reps)}."
            out.append(TextHit("LT-" + m.get("rule", {}).get("id", "RULE"), category, "Warning" if category != "Grammar" or m.get("rule", {}).get("issueType") in ("grammar", "misspelling") else "Info",
                               off, off + ln, msg, f"LanguageTool (local): {m.get('rule', {}).get('description', '')}", reps[0] if reps else None, False))
        return out


def download_languagetool(dest_folder: str, progress=None) -> str:
    """Download and extract LanguageTool once. Only the LanguageTool software is downloaded; no assessment data is sent."""
    import zipfile
    os.makedirs(dest_folder, exist_ok=True)
    zpath = os.path.join(dest_folder, "LanguageTool-stable.zip")
    with urllib.request.urlopen(LT_DOWNLOAD_URL, timeout=60) as r, open(zpath, "wb") as fh:
        total = int(r.headers.get("Content-Length") or 0)
        got = 0
        while True:
            chunk = r.read(1 << 20)
            if not chunk:
                break
            fh.write(chunk)
            got += len(chunk)
            if progress:
                progress(got, total)
    with zipfile.ZipFile(zpath) as z:
        z.extractall(dest_folder)
    os.remove(zpath)
    jar = find_lt_jar(dest_folder)
    if not jar:
        raise RuntimeError("Download finished but languagetool-server.jar was not found.")
    return os.path.dirname(jar)
