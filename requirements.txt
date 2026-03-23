"""
SpareX Assist - AI Chat Assistant for Spare Finder
===================================================
A smart assistant that helps you find spare parts from your Excel data
using text chat.

Features:
  - Automatically loads spare data from the designated Excel file
  - Intelligent Q&A with natural language understanding
  - Modern dark-themed chat GUI
  - Search by any column (part name, number, location, etc.)

Requirements (install once):
  pip install openpyxl pandas

Author: SpareX Team
"""

import os
import sys
import re
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

# ---------------------------------------------------------------------------
# Dependency check & imports
# ---------------------------------------------------------------------------
# Only run dependency check if NOT frozen (i.e., running as script)
if not getattr(sys, 'frozen', False):
    MISSING_LIBS = []

    try:
        import pandas as pd
    except ImportError:
        MISSING_LIBS.append("pandas")

    try:
        import openpyxl
    except ImportError:
        MISSING_LIBS.append("openpyxl")

    if MISSING_LIBS:
        print(f"Installing missing libraries: {', '.join(MISSING_LIBS)} ...")
        import subprocess
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install"] + MISSING_LIBS,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        # Re-import after install
        import pandas as pd
        import openpyxl
else:
    # In frozen mode, just import them directly.
    # If they are missing, it will crash, but we can't pip install anyway.
    import pandas as pd
    import openpyxl

# ---------------------------------------------------------------------------
# Color palette & styling constants
# ---------------------------------------------------------------------------
BG_DARK       = "#1a1a2e"
BG_PANEL      = "#16213e"
BG_CARD       = "#0f3460"
BG_INPUT      = "#1a1a2e"
FG_TEXT        = "#e0e0e0"
FG_ACCENT     = "#00d2ff"
FG_USER_MSG   = "#ffffff"
FG_BOT_MSG    = "#c8f7c5"
FG_ERROR      = "#ff6b6b"
FG_HIGHLIGHT  = "#ffd93d"
BG_USER_BUBBLE = "#0f3460"
BG_BOT_BUBBLE  = "#1b4332"
BTN_BG        = "#00bcd4"
BTN_HOVER     = "#00e5ff"
FONT_TITLE    = ("Segoe UI", 16, "bold")
FONT_HEADER   = ("Segoe UI", 12, "bold")
FONT_BODY     = ("Segoe UI", 11)
FONT_SMALL    = ("Segoe UI", 9)
FONT_INPUT    = ("Segoe UI", 12)
FONT_BTN      = ("Segoe UI", 11, "bold")


# ---------------------------------------------------------------------------
# Keyword aliases – maps common short names to terms found in DESCRIPTION
# ---------------------------------------------------------------------------
KEYWORD_ALIASES = {
    "motor": ["ret motor"],
    "ret": ["ret motor"],
    "adapter": ["adapter", "power adapter", "psu adapter", "sma adapter",
                "pa adapter", "pwr adapter", "interface"],
    "pin": ["pin", "pin block", "probe"],
    "probe": ["probe", "rf probe", "probe hss", "probe ascf"],
    "connector": ["connector", "rf connector", "power connector",
                  "trx rf connector"],
    "fuse": ["fuse", "smd fuse"],
    "cable": ["cable", "fibre cable", "rf cable"],
    "sfp": ["sfp", "sfp module"],
    "card": ["card", "fpga card"],
    "fan": ["fan", "cooling fan"],
    "relay": ["relay"],
    "cover": ["cover"],
    "plug": ["plug"],
    "screw": ["screw"],
    "washer": ["washer"],
    "tape": ["tape"],
    "switch": ["switch"],
    "coupler": ["coupler"],
    "splitter": ["splitter", "power splitter"],
    "ic": ["ic"],
    "receptacle": ["recepticle", "receptacle", "trx receptacle"],
    "attenuator": ["attenuator", "sma attenuator"],
    "bushing": ["bushing"],
    "usb": ["usb hub", "usb"],
    "controller": ["controller"],
    "shunt": ["shunt"],
    "spring": ["gas spring", "spring"],
    "pad": ["pad", "pads"],
    "unit": ["unit"],
    "block": ["dc block", "pin block"],
    "power": ["power adapter", "power splitter", "power connector",
              "pwr adapter", "power supply"],
    "rf": ["rf probe", "rf connector", "rf cable", "rf adapter"],
    "sma": ["sma adapter", "sma connector", "sma attenuator"],
    "bearing": ["bearing"],
}


# ---------------------------------------------------------------------------
# Intent patterns for natural language Q&A
# ---------------------------------------------------------------------------
# Each tuple: (compiled regex, intent_name)
INTENT_PATTERNS = [
    # --- LOCATION questions ---
    (re.compile(r"^(?:where\s+is|where\s+are|where\s+can\s+i\s+find|"
                r"location\s+of|find\s+location|kahan\s+hai|kaha\s+hai|"
                r"locate)\s+(.+)$", re.I),
     "location"),

    # --- STOCK questions ---
    (re.compile(r"^(?:how\s+many|stock\s+of|quantity\s+of|count\s+of|"
                r"kitne|kitna|available\s+stock|closing\s+stock\s+of|"
                r"stock\s+check|check\s+stock)\s+(.+)$", re.I),
     "stock"),

    # --- VENDOR questions ---
    (re.compile(r"^(?:vendor\s+of|supplier\s+of|who\s+supplies|"
                r"who\s+is\s+the\s+vendor|manufacturer\s+of|"
                r"vendor\s+for|supplier\s+for)\s+(.+)$", re.I),
     "vendor"),

    # --- PROJECT questions ---
    (re.compile(r"^(?:which\s+project\s+(?:has|uses|for)|project\s+of|"
                r"project\s+for|kis\s+project)\s+(.+)$", re.I),
     "project"),

    # --- TEST BENCH questions ---
    (re.compile(r"^(?:test\s+bench\s+(?:for|of)|which\s+bench|"
                r"bench\s+for|used\s+(?:on|in|at)\s+which\s+bench)\s+(.+)$", re.I),
     "bench"),

    # --- ALMIRA / cabinet listing ---
    (re.compile(r"^(?:what\s+is\s+in|show\s+(?:all\s+in|items\s+in)|"
                r"list\s+(?:all\s+in|items\s+in)|parts\s+in)\s+(m\d+)$", re.I),
     "almira_list"),

    # --- ZONE listing ---
    (re.compile(r"^(?:what\s+is\s+in|show\s+(?:all\s+in|items\s+in)|"
                r"list\s+(?:all\s+in|items\s+in)|parts\s+in)\s+"
                r"(?:zone\s+)?(00[a-dA-D])$", re.I),
     "zone_list"),

    # --- TYPE filter ---
    (re.compile(r"^(?:show\s+all|list\s+all|all)\s+(.+?)(?:\s+type)?$", re.I),
     "type_filter"),

    # --- PART CODE lookup ---
    (re.compile(r"^(?:part\s+code\s+(?:of|for)|code\s+(?:of|for)|"
                r"what\s+is\s+the\s+(?:part\s+)?code\s+(?:of|for))\s+(.+)$", re.I),
     "partcode"),

    # --- FULL INFO ---
    (re.compile(r"^(?:tell\s+me\s+about|details\s+of|info\s+(?:of|about|on)|"
                r"full\s+info|show\s+details)\s+(.+)$", re.I),
     "full_info"),
]


# ---------------------------------------------------------------------------
# SpareEngine – loads Excel / CSV and performs intelligent Q&A search
# ---------------------------------------------------------------------------
class SpareEngine:
    """Handles loading spare data and intelligent Q&A searching."""

    def __init__(self):
        self.df = None
        self.file_path = None
        self.columns = []

    def load_file(self, path: str) -> str:
        """Load an Excel or CSV file. Returns a status message."""
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in (".xlsx", ".xls"):
                self.df = pd.read_excel(path, engine="openpyxl")
            elif ext == ".csv":
                self.df = pd.read_csv(path)
            else:
                return f"❌ Unsupported file type: {ext}"

            # Clean up column names
            self.df.columns = [str(c).strip() for c in self.df.columns]
            self.columns = list(self.df.columns)
            self.file_path = path
            rows = len(self.df)
            cols_str = ", ".join(self.columns)
            return (
                f"✅ Loaded **{os.path.basename(path)}** successfully!\n"
                f"   📊 {rows} spare records found.\n"
                f"   📋 Columns: {cols_str}"
            )
        except Exception as e:
            return f"❌ Error loading file: {e}"

    # ---- Helper utilities ----

    def _normalize(self, text: str) -> str:
        """Lowercase, strip, collapse whitespace."""
        return re.sub(r"\s+", " ", str(text).lower().strip())

    def _fuzzy_match(self, query: str, text: str) -> bool:
        """Check if all query words appear in text (order-independent)."""
        q_words = self._normalize(query).split()
        t = self._normalize(text)
        return all(w in t for w in q_words)

    def _safe_str(self, val) -> str:
        """Convert a value to string, returning '—' for NaN/None."""
        if val is None:
            return "—"
        try:
            if pd.isna(val):
                return "—"
        except (TypeError, ValueError):
            pass
        s = str(val).strip()
        return s if s else "—"

    def _has_column(self, name: str) -> bool:
        """Check if a column exists (case-insensitive)."""
        return any(self._normalize(c) == self._normalize(name) for c in self.columns)

    def _get_column(self, name: str) -> str:
        """Get the actual column name (case-insensitive match)."""
        for c in self.columns:
            if self._normalize(c) == self._normalize(name):
                return c
        return name

    # ---- Core part finding ----

    def _find_parts(self, search_term: str, max_results: int = 20):
        """
        Find spare parts matching the search_term.
        Searches DESCRIPTION and PART CODE columns first, then falls back
        to searching all columns. Also expands keyword aliases.
        Returns a list of matching row dicts.
        """
        if self.df is None:
            return []

        search_lower = self._normalize(search_term)
        results = []
        seen_indices = set()

        # --- Step 1: expand aliases ---
        expanded_terms = [search_term]
        for alias_key, alias_values in KEYWORD_ALIASES.items():
            if search_lower == alias_key or search_lower in [self._normalize(v) for v in alias_values]:
                for av in alias_values:
                    if self._normalize(av) != search_lower:
                        expanded_terms.append(av)
                # Also add the alias key itself if it isn't the original
                if alias_key != search_lower:
                    expanded_terms.append(alias_key)

        # --- Step 2: search DESCRIPTION and PART CODE with all terms ---
        desc_col = self._get_column("DESCRIPTION")
        code_col = self._get_column("PART CODE")

        for idx, row in self.df.iterrows():
            if idx in seen_indices:
                continue
            desc_val = self._normalize(str(row.get(desc_col, "")))
            code_val = self._normalize(str(row.get(code_col, "")))
            combined_target = desc_val + " " + code_val

            for term in expanded_terms:
                if self._fuzzy_match(term, combined_target):
                    results.append(row.to_dict())
                    seen_indices.add(idx)
                    break

            if len(results) >= max_results:
                break

        # --- Step 3: if nothing found, search ALL columns ---
        if not results:
            for idx, row in self.df.iterrows():
                if idx in seen_indices:
                    continue
                combined = " ".join(str(v) for v in row.values)
                for term in expanded_terms:
                    if self._fuzzy_match(term, combined):
                        results.append(row.to_dict())
                        seen_indices.add(idx)
                        break
                if len(results) >= max_results:
                    break

        return results

    # ---- Intent detection ----

    def _detect_intent(self, query: str):
        """
        Detect the user's intent from their query.
        Returns (intent_name, extracted_search_term) or (None, query).
        """
        q = query.strip()
        for pattern, intent in INTENT_PATTERNS:
            m = pattern.match(q)
            if m:
                return intent, m.group(1).strip()
        return None, q

    # ---- Smart Q&A entry point ----

    def smart_query(self, query: str):
        """
        Main Q&A method. Detects intent, finds parts, and returns a
        formatted conversational answer string plus optional result dicts.

        Returns: (answer_text: str, result_cards: list[dict], speech_text: str)
        """
        if self.df is None:
            return ("⚠️ Data file not loaded!", [], "Data file not loaded.")

        intent, search_term = self._detect_intent(query)

        # ----- ALMIRA LISTING -----
        if intent == "almira_list":
            return self._answer_almira_list(search_term)

        # ----- ZONE LISTING -----
        if intent == "zone_list":
            return self._answer_zone_list(search_term)

        # ----- TYPE FILTER -----
        if intent == "type_filter":
            return self._answer_type_filter(search_term)

        # ----- Find matching parts for all other intents -----
        parts = self._find_parts(search_term)
        if not parts:
            msg = f"😕 No results found for \"{search_term}\".\nTry different keywords or check spelling."
            return (msg, [], f"No results found for {search_term}.")

        # ----- LOCATION -----
        if intent == "location":
            return self._answer_location(search_term, parts)

        # ----- STOCK -----
        if intent == "stock":
            return self._answer_stock(search_term, parts)

        # ----- VENDOR -----
        if intent == "vendor":
            return self._answer_vendor(search_term, parts)

        # ----- PROJECT -----
        if intent == "project":
            return self._answer_project(search_term, parts)

        # ----- TEST BENCH -----
        if intent == "bench":
            return self._answer_bench(search_term, parts)

        # ----- PART CODE -----
        if intent == "partcode":
            return self._answer_partcode(search_term, parts)

        # ----- FULL INFO -----
        if intent == "full_info":
            return self._answer_full_info(search_term, parts)

        # ----- DEFAULT: general search with smart summary -----
        return self._answer_general(search_term, parts)

    # ================================================================
    # Answer formatters for each intent
    # ================================================================

    def _answer_location(self, term, parts):
        """Format location answer."""
        lines = []
        speech_parts = []
        for p in parts[:5]:
            desc = self._safe_str(p.get(self._get_column("DESCRIPTION")))
            almira = self._safe_str(p.get(self._get_column("ALMIRA NO")))
            zone = self._safe_str(p.get(self._get_column("ZONE")))
            bin_val = self._safe_str(p.get(self._get_column("BIN")))
            location = self._safe_str(p.get(self._get_column("LOCATION")))

            lines.append(
                f"📍 **{desc}** is located at:\n"
                f"   🗄️ Almira: **{almira}** | Zone: **{zone}** | Bin: **{bin_val}**\n"
                f"   📦 Location Code: **{location}**"
            )
            speech_parts.append(
                f"{desc} is in Almira {almira}, Zone {zone}, {bin_val}"
            )

        answer = f"🔍 Location for \"{term}\":\n\n" + "\n\n".join(lines)
        if len(parts) > 5:
            answer += f"\n\n... and {len(parts) - 5} more result(s)."
        speech = ". ".join(speech_parts[:2])
        return (answer, [], speech)

    def _answer_stock(self, term, parts):
        """Format stock answer."""
        lines = []
        speech_parts = []
        for p in parts[:5]:
            desc = self._safe_str(p.get(self._get_column("DESCRIPTION")))
            opening = self._safe_str(p.get(self._get_column("OPENING STOCK")))
            issued = self._safe_str(p.get(self._get_column("ISSUED")))
            recd = self._safe_str(p.get(self._get_column("RECD.")))
            closing = self._safe_str(p.get(self._get_column("CLOSING STOCK")))

            lines.append(
                f"📊 **{desc}**:\n"
                f"   Opening: {opening} | Issued: {issued} | "
                f"Received: {recd} | Closing: **{closing}**"
            )
            speech_parts.append(f"{desc} has closing stock {closing}")

        answer = f"📦 Stock info for \"{term}\":\n\n" + "\n\n".join(lines)
        if len(parts) > 5:
            answer += f"\n\n... and {len(parts) - 5} more result(s)."
        speech = ". ".join(speech_parts[:2])
        return (answer, [], speech)

    def _answer_vendor(self, term, parts):
        """Format vendor answer."""
        lines = []
        speech_parts = []
        for p in parts[:5]:
            desc = self._safe_str(p.get(self._get_column("DESCRIPTION")))
            vendor = self._safe_str(p.get(self._get_column("VENDOR")))
            address = self._safe_str(p.get(self._get_column("VENDOR ADDRESS")))

            lines.append(
                f"🏭 **{desc}**:\n"
                f"   Vendor: **{vendor}** ({address})"
            )
            speech_parts.append(f"{desc} vendor is {vendor}")

        answer = f"🏢 Vendor info for \"{term}\":\n\n" + "\n\n".join(lines)
        if len(parts) > 5:
            answer += f"\n\n... and {len(parts) - 5} more result(s)."
        speech = ". ".join(speech_parts[:2])
        return (answer, [], speech)

    def _answer_project(self, term, parts):
        """Format project answer."""
        projects = set()
        desc_name = "—"
        for p in parts:
            proj = self._safe_str(p.get(self._get_column("PROJECT")))
            if proj != "—":
                projects.add(proj)
            d = self._safe_str(p.get(self._get_column("DESCRIPTION")))
            if d != "—":
                desc_name = d

        if projects:
            proj_list = ", ".join(sorted(projects))
            answer = (
                f"📌 **{desc_name}** is found in the following project(s):\n\n"
                f"   🏗️ {proj_list}"
            )
            speech = f"{desc_name} is used in projects {proj_list}"
        else:
            answer = f"😕 No project information found for \"{term}\"."
            speech = f"No project information found for {term}"

        return (answer, [], speech)

    def _answer_bench(self, term, parts):
        """Format test bench answer."""
        benches = set()
        desc_name = "—"
        for p in parts:
            bench = self._safe_str(p.get(self._get_column("TEST BENCH")))
            if bench != "—":
                benches.add(bench)
            d = self._safe_str(p.get(self._get_column("DESCRIPTION")))
            if d != "—":
                desc_name = d

        if benches:
            bench_list = ", ".join(sorted(benches))
            answer = (
                f"🔬 **{desc_name}** is used on:\n\n"
                f"   🧪 {bench_list}"
            )
            speech = f"{desc_name} is used on {bench_list}"
        else:
            answer = f"😕 No test bench information found for \"{term}\"."
            speech = f"No test bench information found for {term}"

        return (answer, [], speech)

    def _answer_partcode(self, term, parts):
        """Format part code answer."""
        lines = []
        speech_parts = []
        for p in parts[:5]:
            desc = self._safe_str(p.get(self._get_column("DESCRIPTION")))
            code = self._safe_str(p.get(self._get_column("PART CODE")))

            lines.append(f"🔖 **{desc}** → Part Code: **{code}**")
            speech_parts.append(f"{desc} part code is {code}")

        answer = f"🏷️ Part code for \"{term}\":\n\n" + "\n".join(lines)
        speech = ". ".join(speech_parts[:2])
        return (answer, [], speech)

    def _answer_full_info(self, term, parts):
        """Return full info cards."""
        count = min(len(parts), 5)
        answer = f"🔍 Found **{len(parts)}** result(s) for \"{term}\":\n"
        speech = f"Found {len(parts)} results for {term}."

        # Return cards for display
        cards = parts[:count]
        return (answer, cards, speech)

    def _answer_almira_list(self, almira_code):
        """List all parts in a specific almira/cabinet."""
        if self.df is None:
            return ("⚠️ Data file not loaded!", [], "Data file not loaded.")

        almira_col = self._get_column("ALMIRA NO")
        code_upper = almira_code.upper()
        matches = self.df[
            self.df[almira_col].astype(str).str.strip().str.upper() == code_upper
        ]

        if matches.empty:
            return (
                f"😕 No items found in Almira **{code_upper}**.",
                [],
                f"No items found in Almira {code_upper}."
            )

        desc_col = self._get_column("DESCRIPTION")
        zone_col = self._get_column("ZONE")
        bin_col = self._get_column("BIN")

        lines = []
        for _, row in matches.head(15).iterrows():
            desc = self._safe_str(row.get(desc_col))
            zone = self._safe_str(row.get(zone_col))
            bin_val = self._safe_str(row.get(bin_col))
            lines.append(f"  • {desc}  (Zone: {zone}, {bin_val})")

        answer = (
            f"🗄️ Items in Almira **{code_upper}** ({len(matches)} total):\n\n"
            + "\n".join(lines)
        )
        if len(matches) > 15:
            answer += f"\n\n... and {len(matches) - 15} more."

        speech = f"Almira {code_upper} has {len(matches)} items."
        return (answer, [], speech)

    def _answer_zone_list(self, zone_code):
        """List all parts in a specific zone."""
        if self.df is None:
            return ("⚠️ Data file not loaded!", [], "Data file not loaded.")

        zone_col = self._get_column("ZONE")
        code_upper = zone_code.upper()
        matches = self.df[
            self.df[zone_col].astype(str).str.strip().str.upper() == code_upper
        ]

        if matches.empty:
            return (
                f"😕 No items found in Zone **{code_upper}**.",
                [],
                f"No items found in Zone {code_upper}."
            )

        desc_col = self._get_column("DESCRIPTION")
        almira_col = self._get_column("ALMIRA NO")
        bin_col = self._get_column("BIN")

        lines = []
        for _, row in matches.head(15).iterrows():
            desc = self._safe_str(row.get(desc_col))
            almira = self._safe_str(row.get(almira_col))
            bin_val = self._safe_str(row.get(bin_col))
            lines.append(f"  • {desc}  (Almira: {almira}, {bin_val})")

        answer = (
            f"🏷️ Items in Zone **{code_upper}** ({len(matches)} total):\n\n"
            + "\n".join(lines)
        )
        if len(matches) > 15:
            answer += f"\n\n... and {len(matches) - 15} more."

        speech = f"Zone {code_upper} has {len(matches)} items."
        return (answer, [], speech)

    def _answer_type_filter(self, type_term):
        """List all parts of a specific TYPE."""
        if self.df is None:
            return ("⚠️ Data file not loaded!", [], "Data file not loaded.")

        type_col = self._get_column("TYPE")
        term_lower = self._normalize(type_term)

        matches = self.df[
            self.df[type_col].astype(str).apply(self._normalize) == term_lower
        ]

        # If exact match fails, try fuzzy
        if matches.empty:
            mask = self.df[type_col].astype(str).apply(
                lambda x: self._fuzzy_match(type_term, x)
            )
            matches = self.df[mask]

        if matches.empty:
            return (
                f"😕 No items found with type \"{type_term}\".",
                [],
                f"No items found with type {type_term}."
            )

        desc_col = self._get_column("DESCRIPTION")
        loc_col = self._get_column("LOCATION")

        lines = []
        for _, row in matches.head(15).iterrows():
            desc = self._safe_str(row.get(desc_col))
            loc = self._safe_str(row.get(loc_col))
            lines.append(f"  • {desc}  (Location: {loc})")

        answer = (
            f"📋 All **{type_term.upper()}** type items ({len(matches)} total):\n\n"
            + "\n".join(lines)
        )
        if len(matches) > 15:
            answer += f"\n\n... and {len(matches) - 15} more."

        speech = f"Found {len(matches)} items of type {type_term}."
        return (answer, [], speech)

    def _answer_general(self, term, parts):
        """
        Default answer: show a smart summary for the first result,
        then return remaining results as cards.
        """
        first = parts[0]
        desc = self._safe_str(first.get(self._get_column("DESCRIPTION")))
        almira = self._safe_str(first.get(self._get_column("ALMIRA NO")))
        zone = self._safe_str(first.get(self._get_column("ZONE")))
        bin_val = self._safe_str(first.get(self._get_column("BIN")))
        location = self._safe_str(first.get(self._get_column("LOCATION")))
        closing = self._safe_str(first.get(self._get_column("CLOSING STOCK")))
        project = self._safe_str(first.get(self._get_column("PROJECT")))
        vendor = self._safe_str(first.get(self._get_column("VENDOR")))
        part_code = self._safe_str(first.get(self._get_column("PART CODE")))
        bench = self._safe_str(first.get(self._get_column("TEST BENCH")))

        answer = (
            f"🔍 Found **{len(parts)}** result(s) for \"{term}\":\n\n"
            f"📌 **{desc}**\n"
            f"   📍 Location: Almira **{almira}** | Zone **{zone}** | **{bin_val}**\n"
            f"   📦 Location Code: **{location}**\n"
            f"   📊 Closing Stock: **{closing}**\n"
            f"   🏗️ Project: {project}\n"
            f"   🏷️ Part Code: {part_code}\n"
            f"   🏭 Vendor: {vendor}"
        )
        if bench != "—":
            answer += f"\n   🧪 Test Bench: {bench}"

        speech = (
            f"Found {len(parts)} results for {term}. "
            f"{desc} is in Almira {almira}, Zone {zone}, {bin_val}. "
            f"Closing stock is {closing}."
        )

        # Return remaining items as cards (skip first since we showed it above)
        cards = parts[1:6] if len(parts) > 1 else []
        if len(parts) > 1:
            answer += f"\n\n--- More results ---"

        return (answer, cards, speech)

    # ---- Legacy search (kept as fallback) ----

    def search(self, query: str, max_results: int = 10):
        """
        Search across ALL columns for rows matching the query.
        Returns a list of dicts (matching rows) and a summary string.
        """
        if self.df is None:
            return [], "⚠️ No file loaded yet. Please load an Excel file first."

        query = query.strip()
        if not query:
            return [], "⚠️ Please enter a search query."

        # --- Detect column-specific search like "column: value" ---
        col_match = re.match(r"^(.+?):\s*(.+)$", query)
        target_col = None
        search_term = query

        if col_match:
            potential_col = col_match.group(1).strip()
            # Check if this matches a column name (case-insensitive)
            for col in self.columns:
                if self._normalize(potential_col) == self._normalize(col):
                    target_col = col
                    search_term = col_match.group(2).strip()
                    break

        results = []
        for _, row in self.df.iterrows():
            if target_col:
                # Search only in the specified column
                if self._fuzzy_match(search_term, str(row[target_col])):
                    results.append(row.to_dict())
            else:
                # Search across all columns
                combined = " ".join(str(v) for v in row.values)
                if self._fuzzy_match(search_term, combined):
                    results.append(row.to_dict())

            if len(results) >= max_results:
                break

        if results:
            summary = f"🔍 Found **{len(results)}** result(s) for \"{search_term}\""
            if target_col:
                summary += f" in column **{target_col}**"
            summary += ":\n"
        else:
            summary = f"😕 No results found for \"{search_term}\". Try different keywords."

        return results, summary

    def get_stats(self) -> str:
        """Get stats about the loaded data."""
        if self.df is None:
            return "No file loaded."
        return (
            f"📁 File: {os.path.basename(self.file_path)}\n"
            f"📊 Total Records: {len(self.df)}\n"
            f"📋 Columns ({len(self.columns)}): {', '.join(self.columns)}"
        )

    def get_all_data_preview(self, n: int = 5) -> str:
        """Return first n rows as a preview."""
        if self.df is None:
            return "No file loaded."
        preview = self.df.head(n).to_string(index=False)
        return f"📄 First {n} rows preview:\n{preview}"



# ---------------------------------------------------------------------------
# SpareXApp – Main GUI Application
# ---------------------------------------------------------------------------
class SpareXApp:
    """Modern dark-themed chat GUI for SpareX Assist."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.engine = SpareEngine()

        self._setup_window()
        self._build_ui()
        self._show_welcome()
        self._auto_load_file()

    # ---------- Window setup ----------
    def _setup_window(self):
        self.root.title("SpareX Assist – AI Spare Finder")
        self.root.configure(bg=BG_DARK)
        self.root.minsize(750, 600)

        # Center the window
        w, h = 900, 700
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        # Style
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Dark.TFrame", background=BG_DARK)
        style.configure("Panel.TFrame", background=BG_PANEL)
        style.configure(
            "Accent.TButton",
            background=BTN_BG,
            foreground="white",
            font=FONT_BTN,
            padding=(12, 6),
        )
        style.map(
            "Accent.TButton",
            background=[("active", BTN_HOVER)],
        )


    # ---------- Build UI Components ----------
    def _build_ui(self):
        # ---- Top bar ----
        top_frame = tk.Frame(self.root, bg=BG_PANEL, pady=8, padx=12)
        top_frame.pack(fill=tk.X)

        title_lbl = tk.Label(
            top_frame,
            text="🔧 SpareX Assist",
            font=FONT_TITLE,
            bg=BG_PANEL,
            fg=FG_ACCENT,
        )
        title_lbl.pack(side=tk.LEFT)

        subtitle_lbl = tk.Label(
            top_frame,
            text="AI Chat Assistant for Spare Finder",
            font=FONT_SMALL,
            bg=BG_PANEL,
            fg="#888",
        )
        subtitle_lbl.pack(side=tk.LEFT, padx=(10, 0), pady=(5, 0))

        # ---- Chat area ----
        chat_frame = tk.Frame(self.root, bg=BG_DARK)
        chat_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 0))

        self.chat_canvas = tk.Canvas(chat_frame, bg=BG_DARK, highlightthickness=0)
        scrollbar = ttk.Scrollbar(chat_frame, orient=tk.VERTICAL, command=self.chat_canvas.yview)
        self.chat_inner = tk.Frame(self.chat_canvas, bg=BG_DARK)

        self.chat_inner.bind(
            "<Configure>",
            lambda e: self.chat_canvas.configure(scrollregion=self.chat_canvas.bbox("all")),
        )
        self.chat_canvas.create_window((0, 0), window=self.chat_inner, anchor="nw")
        self.chat_canvas.configure(yscrollcommand=scrollbar.set)

        self.chat_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Mouse wheel scroll
        self.chat_canvas.bind_all(
            "<MouseWheel>",
            lambda e: self.chat_canvas.yview_scroll(-1 * (e.delta // 120), "units"),
        )

        # ---- Input bar ----
        input_frame = tk.Frame(self.root, bg=BG_PANEL, pady=10, padx=10)
        input_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self.entry_var = tk.StringVar()
        self.entry = tk.Entry(
            input_frame,
            textvariable=self.entry_var,
            font=FONT_INPUT,
            bg=BG_INPUT,
            fg=FG_TEXT,
            insertbackground=FG_ACCENT,
            bd=0,
            relief=tk.FLAT,
        )
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=8, padx=(0, 8))
        self.entry.bind("<Return>", lambda e: self._on_send())

        # Send button
        send_btn = tk.Button(
            input_frame,
            text="💬 Send",
            font=FONT_BTN,
            bg=BTN_BG,
            fg="white",
            activebackground=BTN_HOVER,
            activeforeground="white",
            bd=0,
            padx=16,
            pady=6,
            cursor="hand2",
            command=self._on_send,
        )
        send_btn.pack(side=tk.LEFT, padx=(0, 5))



        # Status bar
        self.status_var = tk.StringVar(value="Loading spare data...")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=FONT_SMALL,
            bg=BG_DARK,
            fg="#666",
            anchor="w",
            padx=12,
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    # ---------- Chat bubble helpers ----------
    def _add_message(self, text: str, sender: str = "bot"):
        """Add a chat message bubble to the chat area."""
        is_user = sender == "user"

        bubble_frame = tk.Frame(self.chat_inner, bg=BG_DARK, pady=4)
        bubble_frame.pack(fill=tk.X, padx=10)

        # Alignment
        anchor = tk.E if is_user else tk.W

        # Label
        name = "You" if is_user else "🤖 SpareX"
        name_lbl = tk.Label(
            bubble_frame,
            text=name,
            font=("Segoe UI", 9, "bold"),
            bg=BG_DARK,
            fg=FG_ACCENT if not is_user else "#aaa",
            anchor=anchor,
        )
        name_lbl.pack(anchor=anchor)

        # Message bubble
        msg_bg = BG_USER_BUBBLE if is_user else BG_BOT_BUBBLE
        msg_fg = FG_USER_MSG if is_user else FG_BOT_MSG

        msg_lbl = tk.Label(
            bubble_frame,
            text=text,
            font=FONT_BODY,
            bg=msg_bg,
            fg=msg_fg,
            wraplength=550,
            justify=tk.LEFT,
            padx=14,
            pady=10,
            anchor="w",
        )
        msg_lbl.pack(anchor=anchor, pady=(2, 0))

        # Scroll to bottom
        self.root.update_idletasks()
        self.chat_canvas.yview_moveto(1.0)

    def _add_result_card(self, data: dict):
        """Add a result card for a spare part."""
        card_frame = tk.Frame(self.chat_inner, bg=BG_CARD, padx=14, pady=10, bd=0)
        card_frame.pack(fill=tk.X, padx=30, pady=3, anchor=tk.W)

        for i, (key, val) in enumerate(data.items()):
            if pd.isna(val):
                val = "—"
            # Skip unnamed columns
            if str(key).startswith("Unnamed"):
                continue
            row_frame = tk.Frame(card_frame, bg=BG_CARD)
            row_frame.pack(fill=tk.X, pady=1)

            key_lbl = tk.Label(
                row_frame,
                text=f"{key}:",
                font=("Segoe UI", 10, "bold"),
                bg=BG_CARD,
                fg=FG_HIGHLIGHT,
                anchor="w",
                width=20,
            )
            key_lbl.pack(side=tk.LEFT)

            val_lbl = tk.Label(
                row_frame,
                text=str(val),
                font=("Segoe UI", 10),
                bg=BG_CARD,
                fg=FG_TEXT,
                anchor="w",
                wraplength=400,
            )
            val_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Scroll to bottom
        self.root.update_idletasks()
        self.chat_canvas.yview_moveto(1.0)

    # ---------- Welcome message ----------
    def _show_welcome(self):
        welcome = (
            "👋 Welcome to SpareX Assist!\n\n"
            "I'm your AI assistant for finding spare parts.\n"
            "Just ask me anything! Examples:\n\n"
            "📍 \"where is RET MOTOR\"\n"
            "📊 \"how many RF PROBE\"\n"
            "🏭 \"vendor of PROBE HSS\"\n"
            "🏗️ \"which project has PIN\"\n"
            "🧪 \"test bench for RF PROBE\"\n"
            "🗄️ \"what is in M03\"\n"
            "📋 \"show all ADAPTER\"\n"
            "🔍 Or just type any keyword like \"motor\"\n\n"
            "💡 Type \"help\" for all commands"
        )
        self._add_message(welcome, "bot")

    # ---------- Auto-load Excel file ----------
    def _auto_load_file(self):
        """Automatically load the designated Excel file on startup."""
        # Determine the directory where this script/exe is located
        if getattr(sys, 'frozen', False):
            # If bundled, the file will be in the temporary folder (sys._MEIPASS)
            # Try bundling location first
            if hasattr(sys, '_MEIPASS'):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        filename = "Tester Spares Master File 2025-2026 1.xlsx"
        path = os.path.join(base_path, filename)

        # Fallback: if not found in bundled location, check next to exe
        if not os.path.exists(path) and getattr(sys, 'frozen', False):
            fallback_path = os.path.join(os.path.dirname(sys.executable), filename)
            if os.path.exists(fallback_path):
                path = fallback_path

        if not os.path.exists(path):
            self._add_message(
                f"❌ Data file not found!\n"
                f"Expected file: {filename}\n"
                f"Please place it in: {script_dir}",
                "bot",
            )
            self.status_var.set("Error – Data file not found")
            return

        self.status_var.set(f"Loading {filename} ...")
        result = self.engine.load_file(path)
        self._add_message(result, "bot")

        if self.engine.df is not None:
            self.status_var.set(f"Loaded: {filename} – {len(self.engine.df)} records")
        else:
            self.status_var.set("Error – Failed to load data file")

    # ---------- Send text query ----------
    def _on_send(self):
        query = self.entry_var.get().strip()
        if not query:
            return

        self.entry_var.set("")
        self._add_message(query, "user")
        self._process_query(query)


    # ---------- Process queries ----------
    def _process_query(self, query: str):
        """Process a user query (text or voice) using the smart Q&A engine."""
        q_lower = query.lower().strip()

        # --- Built-in commands ---
        if q_lower in ("help", "commands", "?"):
            self._show_help()
            return

        if q_lower in ("stats", "info", "status"):
            self._add_message(self.engine.get_stats(), "bot")
            return

        if q_lower in ("preview", "show data", "sample"):
            self._add_message(self.engine.get_all_data_preview(), "bot")
            return

        if q_lower in ("clear", "cls"):
            for widget in self.chat_inner.winfo_children():
                widget.destroy()
            self._show_welcome()
            return

        if q_lower == "reload":
            self._auto_load_file()
            return

        # --- Smart Q&A ---
        if self.engine.df is None:
            self._add_message(
                "⚠️ Data file not loaded!\nPlease ensure the Excel file is in the correct folder and restart the application.",
                "bot",
            )
            return

        self.status_var.set(f"Thinking about \"{query}\" ...")

        # Use the smart Q&A engine
        answer, cards, speech = self.engine.smart_query(query)

        # Show the answer
        self._add_message(answer, "bot")

        # Show result cards if any
        for item in cards:
            self._add_result_card(item)

        self.status_var.set("Ready")

    def _show_help(self):
        help_text = (
            "📖 **SpareX Assist – Q&A Commands:**\n\n"
            "📍 **Location:**\n"
            "   \"where is RET MOTOR\"\n"
            "   \"find location motor\"\n\n"
            "📊 **Stock Check:**\n"
            "   \"how many RF PROBE\"\n"
            "   \"stock of PIN\"\n\n"
            "🏭 **Vendor Info:**\n"
            "   \"vendor of PROBE HSS\"\n"
            "   \"who supplies RET MOTOR\"\n\n"
            "🏗️ **Project Info:**\n"
            "   \"which project has PIN\"\n\n"
            "🧪 **Test Bench:**\n"
            "   \"test bench for RF PROBE\"\n\n"
            "🗄️ **Almira Listing:**\n"
            "   \"what is in M03\"\n\n"
            "📋 **Type Filter:**\n"
            "   \"show all ADAPTER\"\n"
            "   \"all PIN\"\n\n"
            "🔍 **General Search:**\n"
            "   Just type any keyword: motor, bearing, 12345\n\n"
            "⚙️ **Other Commands:**\n"
            "   stats    → Show file info & columns\n"
            "   preview  → Show first 5 rows of data\n"
            "   reload   → Reload the Excel data file\n"
            "   clear    → Clear chat history\n"
            "   help     → Show this help message"
        )
        self._add_message(help_text, "bot")


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def main():
    root = tk.Tk()

    # Set icon if available
    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    app = SpareXApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # If the app crashes, show a message box (helpful for frozen apps without console)
        # We need to import messagebox here in case it wasn't imported successfully earlier
        # or if the crash happened before main().
        try:
            from tkinter import messagebox
            import traceback
            tb = traceback.format_exc()
            root_err = tk.Tk()
            root_err.withdraw()
            messagebox.showerror("SpareX Assist Error", f"An unexpected error occurred:\n\n{e}\n\nDetails:\n{tb}")
        except:
            # If even that fails, just print to stderr (will be logged if run from cmd)
            print(f"Critical error: {e}", file=sys.stderr)
