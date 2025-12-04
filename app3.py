# final_streamlit_app_complete.py
import streamlit as st
import pandas as pd
import json
import re
from rapidfuzz import fuzz
import pdfplumber
from io import BytesIO
import os

# ---------------------------
# Config / filenames
# ---------------------------
DIN_FILENAME = "DIN_standards.json"
MATERIAL_GRADES_FILENAME = "material_grades.json"
SYNONYMS_FILENAME = "synonyms.json"

# ---------------------------
# Helpers: normalization / parsing
# ---------------------------
def normalize(text):
    if text is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())

def normalize_preserve_space(text):
    if text is None:
        return ""
    return re.sub(r"[^A-Z0-9\s]", " ", str(text).upper()).strip()

def normalized_token_in_text(token, text):
    if not token or not text:
        return False
    tnorm = normalize(str(token))
    if not tnorm:
        return False
    txt_norm = normalize(str(text))
    if tnorm in txt_norm:
        return True
    token_ws = normalize_preserve_space(str(token))
    text_ws = normalize_preserve_space(str(text))
    try:
        pattern = r"\b" + re.escape(token_ws) + r"\b"
        if re.search(pattern, text_ws):
            return True
    except re.error:
        pass
    return False

def parse_grades_field(g):
    """Ensure grades separated by commas or semicolons become separate values."""
    if g is None:
        return []
    out = []
    if isinstance(g, list):
        for el in g:
            if el is None:
                continue
            if isinstance(el, (int, float)):
                out.append(str(el))
            elif isinstance(el, str):
                parts = re.split(r'[;,]', el)
                out.extend([p.strip() for p in parts if p.strip()])
            else:
                out.append(str(el).strip())
        return out
    if isinstance(g, (int, float)):
        return [str(g)]
    if isinstance(g, str):
        parts = re.split(r'[;,]', g)
        return [p.strip() for p in parts if p.strip()]
    return []

# ---------------------------
# Load JSON files (cached)
# ---------------------------
@st.cache_data
def load_json_file(fname):
    if not os.path.exists(fname):
        return None
    with open(fname, "r", encoding="utf-8") as f:
        return json.load(f)

din_raw = load_json_file(DIN_FILENAME) or []
material_data = load_json_file(MATERIAL_GRADES_FILENAME) or []
synonyms_data = load_json_file(SYNONYMS_FILENAME) or {}

# ✅ FIX: refined synonym detection logic for Allen Cap vs Hex Bolt
def find_all_synonym_matches(desc, synonyms_dict):
    """
    Return a list of (synonym, main_term) pairs found in desc.
    Special-case CAPSCREW HEX HD → allow both ALLEN CAP SCREW and HEX BOLT mapping.
    Otherwise: if CAPSCREW or ALLEN CAP appears, prefer ALLEN CAP SCREW mapping only.
    """
    if not desc:
        return []
    desc_up = desc.upper()
    out = []

    # If the phrase includes CAPSCREW and HEX HD (or HEXHD) -> allow both Allen Cap Screw and Hex Bolt matches
    if "CAPSCREW" in desc_up and ("HEX HD" in desc_up or "HEXHD" in desc_up or "HEX HEAD" in desc_up):
        for main, syns in synonyms_dict.items():
            syn_list = syns if isinstance(syns, list) else [syns]
            for s in syn_list:
                if s and s.upper() in desc_up:
                    if main.upper() in ["ALLEN CAP SCREW", "HEX BOLT"]:
                        out.append((s, main))
        return out

    # Regular matching:
    for main, syns in synonyms_dict.items():
        syn_list = syns if isinstance(syns, list) else [syns]
        for s in syn_list:
            if s and s.upper() in desc_up:
                # If token is capscrew/allencap, only map it to ALLEN CAP SCREW (avoid mapping to HEX BOLT)
                tok = s.upper()
                if ("CAPSCREW" in tok or "ALLEN CAP" in tok or "SOCKET CAP" in tok or "ALLEN" in tok) and main.upper() != "ALLEN CAP SCREW":
                    # skip mapping to other mains when token is clearly Allen/capscsrew-related
                    continue
                out.append((s, main))
    return out

# ---------------------------
# Smart cleaner — split comma-separated values automatically
# ---------------------------
def split_comma_values_list(lst):
    """Ensure each value inside JSON lists is split correctly if joined by commas or stored as list-strings."""
    out = []
    for v in lst:
        if isinstance(v, str):
            # Handle literal stringified lists like "['DIN 125A','DIN 125B']"
            if v.strip().startswith("[") and v.strip().endswith("]"):
                try:
                    inner = json.loads(v.replace("'", '"'))
                    if isinstance(inner, list):
                        out.extend([str(x).strip() for x in inner if x])
                        continue
                except Exception:
                    pass
            # Handle comma or semicolon separated strings
            parts = [p.strip() for p in re.split(r'[;,]', v) if p.strip()]
            out.extend(parts)
        elif isinstance(v, list):
            # Recursively flatten nested lists
            out.extend(split_comma_values_list(v))
        else:
            out.append(v)
    return out

def clean_din_json(data):
    """Recursively normalize DIN JSON so all fields are consistent lists or strings."""
    if isinstance(data, dict):
        new_data = {}
        for k, v in data.items():
            if isinstance(v, list):
                new_data[k] = split_comma_values_list(v)
            elif isinstance(v, dict):
                new_data[k] = clean_din_json(v)
            elif isinstance(v, str):
                # Split comma-separated strings even if not inside lists
                if "," in v or ";":
                    parts = [p.strip() for p in re.split(r'[;,]', v) if p.strip()]
                    new_data[k] = parts if len(parts) > 1 else parts[0]
                else:
                    new_data[k] = v.strip()
            else:
                new_data[k] = v
        return new_data
    elif isinstance(data, list):
        return [clean_din_json(x) if isinstance(x, (dict, list)) else x for x in data]
    return data

# Clean and normalize the loaded DIN data
din_raw = clean_din_json(din_raw)

# ---------------------------
# Build DIN index & collect finishes
# ---------------------------
def build_din_index(din_data):
    categories = {}
    global_finishes = []

    def add_item(cat_key, item):
        categories.setdefault(cat_key, []).append(item)

    if isinstance(din_data, dict):
        for topk, val in din_data.items():
            topk_lower = str(topk).strip().lower()
            if "finish" in topk_lower and isinstance(val, list):
                for v in val:
                    if isinstance(v, str):
                        global_finishes.append(v.strip().upper())
                    elif isinstance(v, list):
                        for x in v:
                            global_finishes.append(str(x).strip().upper())
                continue
            if isinstance(val, list):
                category_key = topk_lower.rstrip('s')
                for entry in val:
                    if not isinstance(entry, dict):
                        continue
                    type_key = None
                    for k in entry.keys():
                        if k.strip().lower().endswith(" type"):
                            type_key = k
                            break
                    type_name = entry.get(type_key) if type_key else entry.get("Type") or entry.get("Bolt Type") or entry.get("Nut Type") or entry.get("name") or ""
                    standard = entry.get("Standard") or entry.get("standard") or ""
                    inches = entry.get("Inches") or entry.get("inches") or ""
                    metrics = entry.get("Metrics") or entry.get("metrics") or ""
                    grades = parse_grades_field(entry.get("Grades") or entry.get("grades"))
                    entry_finishes = []
                    for k, v in entry.items():
                        if "finish" in k.lower():
                            if isinstance(v, list):
                                for el in v:
                                    if el:
                                        entry_finishes.append(str(el).strip())
                            elif isinstance(v, str):
                                entry_finishes.extend([x.strip() for x in re.split(r'[;,]', v) if x.strip()])
                    add_item(category_key, {
                        "type_name": str(type_name).strip(),
                        "standard": str(standard).strip(),
                        "inches": inches,
                        "metrics": metrics,
                        "grades": grades,
                        "finishes": [f for f in entry_finishes],
                        "raw": entry
                    })
                    for f in entry_finishes:
                        if f:
                            global_finishes.append(str(f).upper())
    elif isinstance(din_data, list):
        for entry in din_data:
            if not isinstance(entry, dict):
                continue
            type_key = None
            for k in entry.keys():
                if k.strip().lower().endswith(" type"):
                    type_key = k
                    break
            if type_key:
                category_key = type_key.strip().split()[0].lower()
            else:
                category_key = "unknown"
                for guess in ["washer", "bolt", "nut", "stud", "screw"]:
                    if any(guess in k.lower() for k in entry.keys()):
                        category_key = guess
                        break
            type_name = entry.get(type_key) if type_key else entry.get("Type") or entry.get("Bolt Type") or entry.get("Nut Type") or entry.get("name") or ""
            standard = entry.get("Standard") or entry.get("standard") or ""
            inches = entry.get("Inches") or entry.get("inches") or ""

            # ✅ Split Metrics properly — make sure each value inside list or string becomes its own entry
            raw_metrics = entry.get("Metrics") or entry.get("metrics") or []
            if isinstance(raw_metrics, list):
               metrics = []
               for m in raw_metrics:
                    if isinstance(m, str):
                       parts = [p.strip() for p in re.split(r'[;,]', m) if p.strip()]
                       metrics.extend(parts)
                    else:
                       metrics.append(str(m).strip())
            else:
                metrics = [p.strip() for p in re.split(r'[;,]', str(raw_metrics)) if p.strip()]

            grades = parse_grades_field(entry.get("Grades") or entry.get("grades"))

            entry_finishes = []
            for k, v in entry.items():
                if "finish" in k.lower():
                    if isinstance(v, list):
                        for el in v:
                            if el:
                                entry_finishes.append(str(el).strip())
                    elif isinstance(v, str):
                        entry_finishes.extend([x.strip() for x in re.split(r'[;,]', v) if x.strip()])
            add_item(category_key, {
                "type_name": str(type_name).strip(),
                "standard": str(standard).strip(),
                "inches": inches,
                "metrics": metrics,
                "grades": grades,
                "finishes": [f for f in entry_finishes],
                "raw": entry
            })
            for f in entry_finishes:
                if f:
                    global_finishes.append(str(f).upper())

    seen = set()
    dedup_fin = []
    for f in global_finishes:
        fu = str(f).upper().strip()
        if fu and fu not in seen:
            seen.add(fu)
            dedup_fin.append(fu)
    return categories, dedup_fin

# ✅ Build DIN index safely (ensures din_index and din_global_finishes are defined)
din_index, din_global_finishes = build_din_index(din_raw)

# ✅ Type lookup for later use (unchanged logic)
type_to_entries = {}
for cat, items in din_index.items():
    for it in items:
        tname = (it.get("type_name") or "").strip()
        if not tname:
            continue
        type_to_entries.setdefault(tname, []).append(it)

category_to_types = {cat: sorted(list({it['type_name'] for it in items if it.get('type_name')})) for cat, items in din_index.items()}

# ----------------------------------------------------
# Helper: flatten list values for dropdown display
# ----------------------------------------------------
def flatten_dropdown_values(values):
    """Flatten and clean dropdown option lists for selectboxes."""
    out = []
    for v in values:
        if isinstance(v, list):
            out.extend([str(x).strip() for x in v if x])
        elif isinstance(v, str):
            out.append(v.strip())
    # remove duplicates while preserving order
    seen = set()
    final = []
    for x in out:
        if x not in seen:
            seen.add(x)
            final.append(x)
    return final

# ---------------------------
# Material grade extraction helpers
# ---------------------------
def extract_candidate_terms(desc):
    patterns = [
        r"SA\d{3,4}\s*GR\s*[A-Z0-9]+",
        r"A\d{3,4}\s*[A-Z0-9]+",
        r"\b[A-Z]{1,3}\d{2,4}[A-Z0-9\-]*\b",
        r"HASTELLOY\s*[A-Z0-9]*",
        r"INCONEL\s*[A-Z0-9]*",
        r"ALLOY\s*[0-9A-Z\-]*",
        r"MONEL\s*[0-9A-Z]*",
        r"TI[A-Z0-9.\-]+",
        r"F\d{2,3}",
        r"\bCL\s*\d+(\.\d+)?\b",
        r"\b\d+\.\d+\b",
        r"ASTM\s*[A-Z]?\d{1,4}[A-Z0-9\-]*"
    ]
    found = []
    desc_upper = desc.upper()
    for pat in patterns:
        for match in re.findall(pat, desc_upper):
            if match and match not in found:
                found.append(match.strip())
    return found

def extract_all_grades(obj, collected=None, title=None):
    if collected is None:
        collected = []
    if isinstance(obj, dict):
        local_title = title or obj.get("title", "")
        for k, v in obj.items():
            if isinstance(v, (dict, list)):
                extract_all_grades(v, collected, local_title)
            elif isinstance(v, (str, int, float)):
                v_clean = str(v).strip()
                if v_clean and len(v_clean) > 0:
                    collected.append((v_clean, local_title, k))
    elif isinstance(obj, list):
        for item in obj:
            extract_all_grades(item, collected, title)
    return collected

all_grades = extract_all_grades(material_data)

def get_grades_from_desc(desc, threshold=85):
    desc_upper = desc.upper()
    candidates = extract_candidate_terms(desc)
    matches = []
    exact_match = None
    for cand in candidates:
        norm_cand = normalize(cand)
        for grade_val, title, key in all_grades:
            norm_grade = normalize(str(grade_val))
            if not norm_grade:
                continue
            if norm_grade in normalize(desc_upper):
                exact_match = grade_val
            if norm_grade in norm_cand or norm_cand in norm_grade:
                matches.append((grade_val, title, key))
            elif fuzz.partial_ratio(norm_grade, norm_cand) >= threshold:
                matches.append((grade_val, title, key))
    if not matches:
        for grade_val, title, key in all_grades:
            if fuzz.partial_ratio(normalize(str(grade_val)), desc_upper) >= threshold:
                matches.append((grade_val, title, key))
    if exact_match:
        matches = sorted(matches, key=lambda x: 0 if x[0] == exact_match else 1)
    seen = set()
    out = []
    for g, t, k in matches:
        if g not in seen:
            seen.add(g)
            out.append(g)
    return out

def get_finish_from_desc(desc, threshold=85):
    desc_upper = " " + re.sub(r"[^A-Z0-9\s]", " ", (desc or "").upper()) + " "
    matched_finishes = []
    for finish in din_global_finishes:
        fin_norm = finish.upper().strip()
        if not fin_norm:
            continue
        try:
            if re.search(r"\b" + re.escape(fin_norm) + r"\b", desc_upper):
                matched_finishes.append(fin_norm)
            elif fuzz.partial_ratio(fin_norm, desc_upper) >= threshold:
                matched_finishes.append(fin_norm)
        except re.error:
            if fin_norm in desc_upper:
                matched_finishes.append(fin_norm)
    for tok in ["ZINC", "HDG", "BLACK", "PLATED", "GALVANIZED", "PASSIVATED"]:
        if tok in desc_upper and tok not in matched_finishes:
            matched_finishes.append(tok)
    return matched_finishes

# ---------------------------
# Category detection + type mapping heuristics
# ---------------------------
CATEGORY_KEYWORDS = {
    "stud": ["STUD", "STUDS"],
    "nut": ["NUT", "NUTS"],
    "bolt": ["BOLT", "BOLTS", "CAPSCREW", "HEXHD", "HEX HEAD", "SOCKET", "GRUB"],
    "washer": ["WASHER", "WASHERS", "FLAT WASHER", "SPRING WASHER", "TOOTH"],
    "screw": ["SCREW", "SCREWS"]
}
def detect_category(desc):
    desc_norm = normalize_preserve_space(desc)
    detected = []
    for cat, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if kw in desc_norm:
                detected.append(cat)
                break
    for cat in din_index.keys():
        if cat and cat in desc_norm and cat not in detected:
            detected.append(cat)
    # dedupe while preserving order
    seen = set()
    out = []
    for d in detected:
        if d not in seen:
            out.append(d)
            seen.add(d)
    return out

# stronger type preferences ensuring HEX NUT / HEX BOLT prioritized (avoid heavy mapping first)
TYPE_PREFERENCE = {
    "NUT": ["HEX NUT", "HEXAGON NUT", "HEXAGON NUT", "HEAVY HEX NUT", "HEAVY NUT"],
    "BOLT": ["HEX BOLT", "HEXAGON BOLT", "HEX HEAD BOLT", "HEX HEAD", "HEAVY HEX BOLT"]
}

# ---------------------------
# Utility: extract family token
# ---------------------------
def standard_family_of(standard_str: str):
    if not standard_str:
        return ""
    s = str(standard_str).upper().strip()
    families = ["DIN", "ISO", "ASTM", "ASME", "ANSI", "BS"]
    for f in families:
        if s.startswith(f):
            return f
        if re.search(r"\b" + re.escape(f) + r"\b", s):
            return f
    m = re.match(r"^[A-Z]+", s)
    if m:
        return m.group(0)
    return "OTHER"

# ---------------------------
# Dimension detection helper (metric vs inch)
# ---------------------------
def detect_dimension_unit(desc):
    """Detect metric or inch from description text."""
    if not desc:
        return ""
    d = desc.upper().replace("×", "X")  # normalize x symbol

    # ---- Metric detection ----
    if re.search(r"\bM\s*\d+(\.\d+)?\b", d):          # M10, M6, M2.5
        return "metric"
    if re.search(r"\b\d+\s*MM\b", d):                 # 10MM
        return "metric"
    if re.search(r"\bM\d+\s*X", d):                   # M10X1.5
        return "metric"
    if any(tok in d for tok in [" DIN", " ISO", " METRIC"]):
        return "metric"

    # ---- Inch detection ----
    if re.search(r"\d+\s*/\s*\d+\s*(?:\"|''|INCH|\bIN\b)", d):   # 3/8", 3/8'', 3/8IN
        return "inch"
    if re.search(r"\d+\s*-\s*\d+\s*/\s*\d+\s*(?:\"|''|INCH|\bIN\b)", d):  # 1-1/2"
        return "inch"
    if re.search(r"\d+\s*(?:\"|''|\bINCH|\bIN\b)\b", d):         # 1", 2'', 2IN
        return "inch"
    if any(tok in d for tok in ["UNC", "UNF", "#", " INCH THREAD"]):
        return "inch"

    return ""

# ---------------------------
# Streamlit UI & Styling
# ---------------------------
st.set_page_config(page_title="Material Grade Mapping — compact", layout="wide")
st.markdown("<h1 style='text-align:center;font-family:system-ui;'>🔩 Material Grade + DIN Standards Mapper</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center;color:gray;'>Compact editor — editable dropdowns with small selected labels above each dropdown.</p>", unsafe_allow_html=True)

st.markdown(
    """
    <style>
    /* Make selectboxes roomy and wrap selected labels */
    div[role="listbox"], div[role="combobox"], div.stSelectbox > div[role="button"] {
        min-width: 320px !important;
        max-width: 100% !important;
        white-space: normal !important;
        overflow: visible !important;
    }
    .stSelectbox [role="listbox"] { min-width: 420px !important; white-space: normal !important; }
    .row-card { border: 1px solid rgba(0,0,0,0.06); padding: 8px; border-radius: 8px; margin-bottom: 6px; background: #fff; }
    .row-header { background: rgba(0,0,0,0.03); padding: 6px; border-radius: 6px; margin-bottom: 8px; }
    .small-selected { font-size:12px; color:#666; margin-bottom:4px; min-height:14px; }
    </style>
    """,
    unsafe_allow_html=True,
)

uploaded_file = st.file_uploader("Upload your Excel, CSV or PDF file", type=["xlsx", "xls", "csv", "pdf"])

def read_uploaded_file(upl):
    if upl is None:
        return pd.DataFrame()
    fname = getattr(upl, "name", "")
    try:
        if fname.lower().endswith((".xlsx", ".xls")):
            try:
                return pd.read_excel(upl)
            except Exception:
                upl.seek(0)
                return pd.read_excel(BytesIO(upl.read()))
        elif fname.lower().endswith(".csv"):
            try:
                return pd.read_csv(upl)
            except Exception:
                upl.seek(0)
                return pd.read_csv(BytesIO(upl.read()))
        elif fname.lower().endswith(".pdf"):
            try:
                lines = []
                with pdfplumber.open(upl) as pdf:
                    for p in pdf.pages:
                        txt = p.extract_text()
                        if txt:
                            for ln in txt.splitlines():
                                if ln.strip():
                                    lines.append(ln.strip())
                if not lines:
                    return pd.DataFrame({"Extracted_Text": []})
                return pd.DataFrame({"Extracted_Text": lines})
            except Exception:
                upl.seek(0)
                try:
                    with pdfplumber.open(BytesIO(upl.read())) as pdf:
                        lines = []
                        for p in pdf.pages:
                            txt = p.extract_text()
                            if txt:
                                for ln in txt.splitlines():
                                    if ln.strip():
                                        lines.append(ln.strip())
                        if not lines:
                            return pd.DataFrame({"Extracted_Text": []})
                        return pd.DataFrame({"Extracted_Text": lines})
                except Exception as e:
                    st.error(f"PDF read error: {e}")
                    return pd.DataFrame({"Extracted_Text": []})
        else:
            return pd.DataFrame()
    except Exception as e:
        st.error(f"File read error: {e}")
        return pd.DataFrame()

# Helper: safe options + normalized index selection
def options_and_normalized_index(options_list, desired_value):
    opts = [""] + list(options_list)
    if desired_value is None:
        return opts, 0
    try:
        if desired_value in opts:
            return opts, opts.index(desired_value)
    except Exception:
        pass
    dnorm = normalize(str(desired_value))
    for idx, o in enumerate(opts):
        if not o:
            continue
        if normalize(str(o)) == dnorm:
            return opts, idx
    for idx, o in enumerate(opts):
        if not o:
            continue
        if dnorm and dnorm in normalize(str(o)):
            return opts, idx
    return opts, 0

# Helper to get session value (fallback provided)
def current_select_value(key, fallback=""):
    try:
        val = st.session_state.get(key, fallback)
        if val is None:
            return fallback
        return val
    except Exception:
        return fallback

if uploaded_file:
    df = read_uploaded_file(uploaded_file)
    if df is None or df.empty:
        st.stop()

    # header heuristics quietly
    first_row = df.iloc[0].astype(str).tolist() if len(df) > 0 else []
    try:
        header_likely_invalid = sum(bool(re.search(r"\d", str(c))) for c in df.columns) > len(df.columns) / 2 if len(df.columns) > 0 else False
    except Exception:
        header_likely_invalid = False
    data_looks_like_header = sum(bool(re.search(r"[A-Za-z]", x)) for x in first_row) > len(first_row) / 2 if first_row else False
    if header_likely_invalid or data_looks_like_header:
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
        try:
            df = pd.read_excel(uploaded_file, header=None)
        except Exception:
            try:
                uploaded_file.seek(0)
            except Exception:
                pass
            try:
                df = pd.read_csv(uploaded_file, header=None)
            except Exception:
                pass
        df.columns = [f"Column_{i}" for i in range(df.shape[1])]

    # detect description column flexibly
    desc_col = None
    for col in df.columns:
        col_lower = str(col).lower()
        if any(word in col_lower for word in ["desc", "description", "material", "details", "part", "combined", "extract"]):
            desc_col = col
            break
    if not desc_col:
        for col in df.columns:
            sample = " ".join(df[col].astype(str).head(5).tolist()).lower()
            if any(word in sample for word in ["bolt", "nut", "screw", "washer", "steel", "stud", "thread"]):
                desc_col = col
                break
    if not desc_col:
        desc_col = df.columns[0]

    # ensure required columns exist
    for c in ["Grade", "Finish", "DIN Category", "Type", "Standard", "Inches", "Metrics", "DIN Details", "Detected Unit", "Standard Family"]:
        if c not in df.columns:
            df[c] = ""

    # Precompute suggestions (these use original desc so pre-mapping remains unchanged)
    grade_results = []
    finish_suggestions = []
    category_suggestions = []
    unit_suggestions = []
    for _, row in df.iterrows():
        desc = str(row[desc_col])
        grade_results.append(get_grades_from_desc(desc))
        finish_suggestions.append(get_finish_from_desc(desc))
        category_suggestions.append(detect_category(desc))
        unit_suggestions.append(detect_dimension_unit(desc))

    # -------------------------
    # Header
    # -------------------------
    st.markdown("<div class='row-header'><b>Mapped rows — editable</b></div>", unsafe_allow_html=True)

    # column widths for the stacked block view (we'll use them consistently for each type-block)
    block_widths = [0.3, 3.2, 1.6, 1.0, 0.9, 2.0, 1.4, 1.6]

    # Top-level small header columns (keeps the same look but won't try to reserve for all possible dynamic columns)
    st_columns = st.columns([0.3, 3.2])
    st_columns[0].markdown("**#**")
    st_columns[1].markdown("**Description (mapped types below)**")

    # Render rows
    for i, row in df.iterrows():
        orig_desc = str(row[desc_col])
        # compute synonym matches (ALL) & safe replacement for local matching ONLY
        matched_syn_pairs = find_all_synonym_matches(orig_desc, synonyms_data)  # list of (synonym, main)
        matched_syn = matched_syn_pairs[0][0] if matched_syn_pairs else None
        matched_mains = [m for s, m in matched_syn_pairs]

        # Build preferred types map from synonyms to prefer exact main-term matches in type lists
        preferred_types_from_syn = {'nut': [], 'bolt': [], 'screw': [], 'washer': [], 'stud': []}
        if matched_mains:
            for main in matched_mains:
                m_up = str(main).upper()
                # guess which category this main term belongs to and store it
                if "NUT" in m_up:
                    preferred_types_from_syn['nut'].append(str(main).strip())
                elif "BOLT" in m_up:
                    preferred_types_from_syn['bolt'].append(str(main).strip())
                elif "SCREW" in m_up:
                    preferred_types_from_syn['screw'].append(str(main).strip())
                elif "WASHER" in m_up:
                    preferred_types_from_syn['washer'].append(str(main).strip())
                elif "STUD" in m_up:
                    preferred_types_from_syn['stud'].append(str(main).strip())
                else:
                    # if no category token but exact name in any type list, try to place it
                    for cat, opts in category_to_types.items():
                        for opt in opts:
                            if normalize(opt) == normalize(main):
                                if cat in preferred_types_from_syn:
                                    preferred_types_from_syn[cat].append(str(main).strip())

        # If synonyms found, replace synonyms in a canonical desc_for_match for better matching
        if matched_syn_pairs:
            desc = orig_desc
            desc_for_match = orig_desc
            for s, m in matched_syn_pairs:
                try:
                    desc_for_match = re.sub(re.escape(s), str(m), str(desc_for_match), flags=re.IGNORECASE)
                except Exception:
                    desc_for_match = desc_for_match.replace(s, m)
            # Add main category tokens inferred from synonyms into category_suggestions
            for main_category in matched_mains:
                m_up = str(main_category).upper()
                if "NUT" in m_up and "nut" not in category_suggestions[i]:
                    category_suggestions[i].append("nut")
                if "BOLT" in m_up and "bolt" not in category_suggestions[i]:
                    category_suggestions[i].append("bolt")
                if "SCREW" in m_up and "screw" not in category_suggestions[i]:
                    category_suggestions[i].append("screw")
        else:
            desc = orig_desc
            desc_for_match = orig_desc

        # detected categories for this row (may be 0..n)
        # ✅ FIX: ensure we combine explicit detected categories + those inferred from matched mains and dedupe
        inferred_cats = []
        for m in matched_mains:
            m_up = str(m).upper()
            if "NUT" in m_up and "nut" not in inferred_cats:
                inferred_cats.append("nut")
            if "BOLT" in m_up and "bolt" not in inferred_cats:
                inferred_cats.append("bolt")
            if "SCREW" in m_up and "screw" not in inferred_cats:
                inferred_cats.append("screw")
            if "WASHER" in m_up and "washer" not in inferred_cats:
                inferred_cats.append("washer")
            if "STUD" in m_up and "stud" not in inferred_cats:
                inferred_cats.append("stud")

        detected_cats = list(dict.fromkeys((category_suggestions[i] or []) + inferred_cats))

        # ensure at least one "unknown" category if none detected to still show a single editable block
        if not detected_cats:
            detected_cats = [""]

        # We'll store per-type results in df columns like Type_1, Standard_1, Grade_1, Finish_1, Unit_1
        # Create them if they don't exist yet (safe)
        for idx in range(len(detected_cats)):
            tcol = f"Type_{idx+1}"
            scol = f"Standard_{idx+1}"
            gcol = f"Grade_{idx+1}"
            fcol = f"Finish_{idx+1}"
            ucol = f"Unit_{idx+1}"
            famcol = f"Standard_Family_{idx+1}"
            if tcol not in df.columns:
                df[tcol] = ""
            if scol not in df.columns:
                df[scol] = ""
            if gcol not in df.columns:
                df[gcol] = ""
            if fcol not in df.columns:
                df[fcol] = ""
            if ucol not in df.columns:
                df[ucol] = ""
            if famcol not in df.columns:
                df[famcol] = ""

        # Render the top number + description row (first type block will show number & description)
        first_block = True
        for t_idx, cat in enumerate(detected_cats):
            # For each type-block create the same 8 columns layout (number, desc, type, family, unit, dim, grade, finish)
            cnum, cdesc, ctype_col, cfamily_col, cunit_col, cdim_col, cgrade_col, cfinish_col = st.columns(block_widths)

            # number and description displayed only for first block; subsequent blocks show blanks in those spots for neat alignment
            if first_block:
                cnum.markdown(f"<div style='padding:6px;border:1px solid rgba(0,0,0,0.06);border-radius:6px;text-align:center'>{i+1}</div>", unsafe_allow_html=True)
                cdesc.markdown(f"<div style='padding:6px;border:1px solid rgba(0,0,0,0.06);border-radius:6px'>{(desc if len(desc) < 300 else desc[:297] + '...')}</div>", unsafe_allow_html=True)
            else:
                cnum.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)
                cdesc.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)

            # show synonym note only on first block (keeps UI small)
            if first_block and matched_syn_pairs:
                note_parts = []
                seen_notes = set()
                for s, m in matched_syn_pairs:
                    label = f"({s} is synonym of {m})"
                    if label not in seen_notes:
                        note_parts.append(label)
                        seen_notes.add(label)
                if note_parts:
                    note_html = "<div style='font-size:12px;color:gray;margin-top:2px;'>" + " ".join(note_parts) + "</div>"
                    cdesc.markdown(note_html, unsafe_allow_html=True)

            # Now per-type mapping logic (adapted from your previous single/multi-case logic but applied to the specific category)
            # Keys are per-row-per-type so they don't clash
            key_t = f"row_{i}_type_{t_idx+1}"
            key_fam = f"row_{i}_family_{t_idx+1}"
            key_unit = f"row_{i}_unit_{t_idx+1}"
            key_dim = f"row_{i}_dim_{t_idx+1}"
            key_grade = f"row_{i}_grade_{t_idx+1}"
            key_finish = f"row_{i}_finish_{t_idx+1}"

            desc_up = desc_for_match.upper()

            # prepare options for this category (if category empty -> all types)
            t_opts = []
            if cat:
                # category_to_types keys come from din_index keys (lowercase): use cat directly
                t_opts = category_to_types.get(cat, []) or []
            if not t_opts:
                # fallback all types (deduped)
                all_types = []
                for k, v in category_to_types.items():
                    all_types.extend(v)
                seent = set()
                t_opts = [x for x in all_types if x and (x not in seent and not seent.add(x))]

            # ✅ FIX: If we have preferred types inferred from synonyms, ensure they appear in the options
            pref_for_cat = None
            if cat and preferred_types_from_syn.get(cat):
                for candidate in preferred_types_from_syn.get(cat):
                    # normalize both candidate and option compare to avoid case mismatch
                    match_opt = next((o for o in t_opts if normalize(o) == normalize(candidate)), None)
                    if match_opt:
                        pref_for_cat = match_opt
                        break
                # If not present in t_opts, insert preferred at front so user can see it and select
                for candidate in preferred_types_from_syn.get(cat):
                    if candidate and all(normalize(candidate) != normalize(o) for o in t_opts):
                        # insert at beginning (display), but keep original list intact after
                        t_opts.insert(0, candidate)

            # Determine a reasonable default
            existing_type_val = ""
            col_type_name = f"Type_{t_idx+1}"
            if df.at[i, col_type_name]:
                existing_type_val = df.at[i, col_type_name]

            pre_type = existing_type_val or pref_for_cat or ""
            # prefer HEX BOLT/HEX NUT where applicable
            if not pre_type and cat:
                if cat == "bolt":
                    pre_type = next((opt for opt in t_opts if opt.upper() == "HEX BOLT"), "")
                elif cat == "nut":
                    pre_type = next((opt for opt in t_opts if opt.upper() == "HEX NUT"), "")
            if not pre_type and t_opts:
                pre_type = t_opts[0]

            # display small-selected above
            display_type = current_select_value(key_t, pre_type)
            ctype_col.markdown(f"<div class='small-selected'>{display_type}</div>", unsafe_allow_html=True)
            try:
                t_options_final = [""] + t_opts
                t_index = t_options_final.index(display_type) if display_type in t_options_final else (t_options_final.index(pre_type) if pre_type in t_options_final else 0)
            except Exception:
                t_index = 0
            sel_type = ctype_col.selectbox("", options=t_options_final, index=t_index, key=key_t, help=f"Type (category: {cat or 'any'})")

            chosen_type = st.session_state.get(key_t, sel_type) or ""
            df.at[i, col_type_name] = chosen_type

            # -------------------- FAMILY --------------------
            # build families from din_index
            fams_all = set()
            for items in din_index.values():
                for it in items:
                    sv = str(it.get("standard") or "").strip()
                    fam = standard_family_of(sv)
                    if fam:
                        fams_all.add(fam)
            # derive families for selected type (current selection)
            sel_types_for_family = [chosen_type] if chosen_type else []
            fams_for_selected_types = set()
            for tsel in sel_types_for_family:
                for items in din_index.values():
                    for it in items:
                        if it.get("type_name") == tsel:
                            sv = str(it.get("standard") or "").strip()
                            fam = standard_family_of(sv) or standard_family_of(str(it.get("metrics") or ""))
                            if fam:
                                fams_for_selected_types.add(fam)
            if len(sel_types_for_family) == 1 and fams_for_selected_types:
                family_options = [""] + sorted(list(fams_for_selected_types))
            else:
                family_options = [""] + sorted(list(fams_all))

            prev_family = df.at[i, "Standard"] and standard_family_of(df.at[i, "Standard"])
            fam_idx = 0
            if prev_family and prev_family in family_options:
                fam_idx = family_options.index(prev_family)
            else:
                for f in family_options:
                    if f and f in desc_for_match.upper():
                        fam_idx = family_options.index(f)
                        break
            display_family = current_select_value(key_fam, family_options[fam_idx] if fam_idx < len(family_options) else "")
            cfamily_col.markdown(f"<div class='small-selected'>{display_family}</div>", unsafe_allow_html=True)
            family_sel = cfamily_col.selectbox("", options=family_options, index=fam_idx, key=key_fam, help="Standard family (DIN/ISO/ASME etc.)")
            df.at[i, f"Standard_Family_{t_idx+1}"] = st.session_state.get(key_fam, family_sel) or ""

            # -------------------- UNIT --------------------
            unit_options = ["metric", "inch"]
            auto_unit = detect_dimension_unit(desc_for_match)
            default_unit = st.session_state.get(key_unit, df.at[i, f"Unit_{t_idx+1}"] if (ucol := f"Unit_{t_idx+1}") in df.columns else "") or (auto_unit if auto_unit in ("metric","inch") else "")
            if not default_unit:
                default_unit = "metric" if "M" in desc_for_match.upper() or "MM" in desc_for_match.upper() else ("inch" if "#" in desc_for_match or "INCH" in desc_for_match.upper() else "metric")
            display_unit = current_select_value(key_unit, default_unit)
            cunit_col.markdown(f"<div class='small-selected'>{display_unit}</div>", unsafe_allow_html=True)
            try:
                unit_index = unit_options.index(default_unit) if default_unit in unit_options else 0
            except Exception:
                unit_index = 0
            unit_sel = cunit_col.selectbox("", options=unit_options, index=unit_index, key=key_unit, help="Unit (metric/inch)")
            df.at[i, f"Unit_{t_idx+1}"] = st.session_state.get(key_unit, unit_sel) or unit_sel

            # -------------------- DIM / STANDARD (filtered) --------------------
            # Build dimensional options filtered by selected Type, Family and Unit (reuse your logic)
            dim_opts = []
            current_sel = chosen_type or ""
            current_family = st.session_state.get(key_fam, df.at[i, f"Standard_Family_{t_idx+1}"] or "") or ""
            current_unit = st.session_state.get(key_unit, df.at[i, f"Unit_{t_idx+1}"] or "") or ""

            for catk, items in din_index.items():
                for it in items:
                    try:
                        if not it.get("type_name"):
                            continue
                        if current_sel and it.get("type_name").strip() != current_sel.strip():
                            continue
                        std_raw = str(it.get("standard") or "").strip()
                        metrics_raw = it.get("metrics") or ""
                        if isinstance(metrics_raw, list):
                            metrics_raw = ", ".join([str(m).strip() for m in metrics_raw if m])
                        inches_raw = str(it.get("inches") or "").strip()
                        # family filter
                        if current_family:
                            fam_it = standard_family_of(std_raw) or standard_family_of(metrics_raw)
                            if fam_it != current_family:
                                if std_raw and re.match(r"^\d", std_raw) and not metrics_raw:
                                    cand = f"{current_family} {std_raw}"
                                    dim_opts.append(cand)
                                else:
                                    continue
                        # unit filter
                        if current_unit == "metric":
                            if metrics_raw:
                                parts = [p.strip() for p in re.split(r'[;,]', str(metrics_raw)) if p.strip()]
                                for p in parts:
                                    if p not in dim_opts:
                                        dim_opts.append(p)
                            else:
                                if std_raw:
                                    if re.search(r"\bDIN\b|\bISO\b|\d", std_raw):
                                        if std_raw not in dim_opts:
                                            dim_opts.append(std_raw)
                        elif current_unit == "inch":
                            if inches_raw:
                                if inches_raw not in dim_opts:
                                    dim_opts.append(inches_raw)
                            else:
                                if std_raw:
                                    if re.search(r"INCH", std_raw.upper()) or re.search(r"['\"]", std_raw):
                                        if std_raw not in dim_opts:
                                            dim_opts.append(std_raw)
                        else:
                            if metrics_raw:
                                parts = [p.strip() for p in re.split(r'[;,]', str(metrics_raw)) if p.strip()]
                                for p in parts:
                                    if p not in dim_opts:
                                        dim_opts.append(p)
                            elif std_raw:
                                if std_raw not in dim_opts:
                                    dim_opts.append(std_raw)
                    except Exception:
                        continue

            # fallback: if none found, but family selected
            if not dim_opts and current_family:
                for catk, items in din_index.items():
                    for it in items:
                        std_raw = str(it.get("standard") or "").strip()
                        metrics_raw = it.get("metrics") or ""
                        if isinstance(metrics_raw, list):
                            metrics_raw = ", ".join([str(m).strip() for m in metrics_raw if m])
                        inches_raw = str(it.get("inches") or "").strip()
                        fam_it = standard_family_of(std_raw) or standard_family_of(metrics_raw)
                        if fam_it == current_family:
                            if current_unit == "metric":
                                if metrics_raw and metrics_raw not in dim_opts:
                                    parts = [p.strip() for p in re.split(r'[;,]', str(metrics_raw)) if p.strip()]
                                    for p in parts:
                                        if p not in dim_opts:
                                            dim_opts.append(p)
                            elif current_unit == "inch":
                                if inches_raw and inches_raw not in dim_opts:
                                    dim_opts.append(inches_raw)
                            else:
                                cand = metrics_raw or std_raw or inches_raw
                                if cand and cand not in dim_opts:
                                    dim_opts.append(cand)

            # final fallback: all unique standards/metrics
            if not dim_opts:
                seen_dim = set()
                for catk, items in din_index.items():
                    for it in items:
                        metrics_raw = it.get("metrics") or ""
                        if isinstance(metrics_raw, list):
                            metrics_raw = ", ".join([str(m).strip() for m in metrics_raw if m])
                        std_raw = str(it.get("standard") or "").strip()
                        inches_raw = str(it.get("inches") or "").strip()
                        if current_unit == "metric" and metrics_raw:
                            parts = [p.strip() for p in re.split(r'[;,]', str(metrics_raw)) if p.strip()]
                            for p in parts:
                                if p and p not in seen_dim:
                                    seen_dim.add(p)
                                    dim_opts.append(p)
                        elif current_unit == "inch" and inches_raw:
                            if inches_raw and inches_raw not in seen_dim:
                                seen_dim.add(inches_raw)
                                dim_opts.append(inches_raw)
                        else:
                            cand_list = [metrics_raw, std_raw, inches_raw]
                            for cand in cand_list:
                                if cand and cand not in seen_dim:
                                    seen_dim.add(cand)
                                    dim_opts.append(cand)

            # dedupe & keep order
            dim_list = []
            seen_dim2 = set()
            for d in dim_opts:
                if not d:
                    continue
                dclean = str(d).strip()
                if dclean not in seen_dim2:
                    seen_dim2.add(dclean)
                    dim_list.append(dclean)

            dim_key = key_dim
            default_dim = df.at[i, f"Standard_{t_idx+1}"] or (dim_list[0] if dim_list else "")

            # If synonym main_term present and it matches one option, prefer it (auto-map)
            if matched_syn_pairs:
                for cand in dim_list:
                    try:
                        if normalize(str(cand)) == normalize(str(matched_mains[0])):
                            default_dim = cand
                            break
                    except Exception:
                        continue

            display_dim = current_select_value(dim_key, default_dim)
            cdim_col.markdown(f"<div class='small-selected'>{display_dim}</div>", unsafe_allow_html=True)
            dim_options_final, dim_idx = options_and_normalized_index(dim_list, default_dim)
            dim_sel = cdim_col.selectbox("", options=dim_options_final, index=dim_idx, key=dim_key, help="Dim Standard (filtered by Type + Family + Unit)")
            df.at[i, f"Standard_{t_idx+1}"] = st.session_state.get(dim_key, dim_sel) or ""

            # -------------------- Grade --------------------
            grade_set = []
            current_dim = st.session_state.get(dim_key, df.at[i, f"Standard_{t_idx+1}"]) or ""
            for catk, items in din_index.items():
                for it in items:
                    if it.get("type_name") != chosen_type:
                        continue
                    if current_dim:
                        if current_dim not in str(it.get("metrics","")) and current_dim not in str(it.get("standard","")):
                            if not (normalize(current_dim).replace(" ", "") in normalize(str(it.get("metrics","")) + normalize(str(it.get("standard",""))))):
                                continue
                    for g in it.get("grades", []):
                        if g:
                            gs = str(g).strip()
                            if gs not in grade_set:
                                grade_set.append(gs)
            # prepend detected grades (computed earlier from original desc)
            for g in grade_results[i]:
                if g and g not in grade_set:
                    grade_set.insert(0, g)
            # fallback to global material_data grades
            if not grade_set and all_grades:
                for gval, _, _ in all_grades:
                    if gval not in grade_set:
                        grade_set.append(gval)

            # automatic exact grade mapping
            default_grade = df.at[i, f"Grade_{t_idx+1}"] or ""
            exact_found = ""
            try:
                for g in grade_set:
                    if normalized_token_in_text(g, desc_for_match):
                        exact_found = g
                        break
                if not exact_found:
                    for gval, _, _ in all_grades:
                        if normalized_token_in_text(gval, desc_for_match):
                            exact_found = gval
                            break
            except Exception:
                exact_found = ""
            # 🩵 Only apply auto-mapping if user has NOT manually changed the grade
            manual_grade = st.session_state.get(key_grade, "")
            if exact_found and not manual_grade:
               df.at[i, f"Grade_{t_idx+1}"] = exact_found
               if key_grade not in st.session_state:
                  st.session_state[key_grade] = exact_found
                  default_grade = exact_found
            else:
              default_grade = manual_grade or df.at[i, f"Grade_{t_idx+1}"] or (grade_set[0] if grade_set else "")

            display_grade = current_select_value(key_grade, default_grade)
            cgrade_col.markdown(f"<div class='small-selected'>{display_grade}</div>", unsafe_allow_html=True)
            grade_options = [""] + grade_set
            try:
                target = st.session_state.get(key_grade, default_grade) or ""
                matched_index = 0
                for idx_opt, opt in enumerate([""] + grade_set):
                    if not opt:
                        continue
                    if normalize(opt) == normalize(str(target)):
                        matched_index = idx_opt
                        break
                grade_sel = cgrade_col.selectbox("", options=grade_options, index=matched_index, key=key_grade, help="Grade (editable)")
            except Exception:
                grade_sel = cgrade_col.selectbox("", options=grade_options, index=0, key=key_grade, help="Grade (editable)")
            df.at[i, f"Grade_{t_idx+1}"] = st.session_state.get(key_grade, grade_sel) or ""

            # -------------------- Finish --------------------
            fin_set = []
            for catk, items in din_index.items():
                for it in items:
                    if it.get("type_name") != chosen_type:
                        continue
                    if current_dim:
                        if current_dim not in str(it.get("metrics","")) and current_dim not in str(it.get("standard","")):
                            if not (normalize(current_dim).replace(" ", "") in normalize(str(it.get("metrics","")) + normalize(str(it.get("standard",""))))):
                                continue
                    for f in it.get("finishes", []):
                        if f:
                            fup = str(f).strip().upper()
                            if fup not in fin_set:
                                fin_set.append(fup)
            # add suggestions and global finishes
            for f in finish_suggestions[i]:
                fu = str(f).strip().upper()
                if fu and fu not in fin_set:
                    fin_set.append(fu)
            for f in din_global_finishes:
                if f and f not in fin_set:
                    fin_set.append(f)
            default_fin = df.at[i, f"Finish_{t_idx+1}"] or ""
            try:
                for fin in fin_set:
                    if normalized_token_in_text(fin, desc_for_match):
                        default_fin = fin
                        df.at[i, f"Finish_{t_idx+1}"] = fin
                        st.session_state[key_finish] = fin
                        break
            except Exception:
                pass
            display_fin = current_select_value(key_finish, default_fin)
            cfinish_col.markdown(f"<div class='small-selected'>{display_fin}</div>", unsafe_allow_html=True)
            fin_options_full = [""] + fin_set
            try:
                fin_sel = cfinish_col.selectbox("", options=fin_options_full, index=fin_options_full.index(default_fin) if default_fin in fin_options_full else 0, key=key_finish, help="Finish (editable)")
            except Exception:
                fin_sel = cfinish_col.selectbox("", options=fin_options_full, index=0, key=key_finish, help="Finish (editable)")
            df.at[i, f"Finish_{t_idx+1}"] = st.session_state.get(key_finish, fin_sel) or ""

            # DIN details quiet (store primary types joined)
            # We'll not overwrite your original "Type" column — but we set it to the comma-joined chosen types for compatibility
            # We'll update it after finishing all type blocks for the row (below)

            first_block = False

        # After rendering all type-blocks for this row, persist combined Type field from Type_1..Type_n
        chosen_types_for_row = []
        for idx in range(len(detected_cats)):
            tval = df.at[i, f"Type_{idx+1}"] if f"Type_{idx+1}" in df.columns else ""
            if tval:
                chosen_types_for_row.append(tval)
        df.at[i, "Type"] = ",".join(chosen_types_for_row)
        # Compose DIN Details field similar to previous format (categories/types/standards)
        cat_str = ",".join(detected_cats) if detected_cats else ""
        # For Standard put comma-joined Standard_1..n
        stds = []
        for idx in range(len(detected_cats)):
            sval = df.at[i, f"Standard_{idx+1}"] if f"Standard_{idx+1}" in df.columns else ""
            if sval:
                stds.append(sval)
        df.at[i, "DIN Details"] = f"{cat_str}/{df.at[i,'Type']}/{','.join(stds)}"

    # -------------------------
    # Final table generation area (button at bottom)
    # -------------------------
    st.markdown("---")
    st.markdown("### Finalize and generate merged table")
    if st.button("Generate Final Table"):
        # Build final table with 5 columns: Sr No, Description, Dimension Standard, Grade, Finish
        final_rows = []
        for i, row in df.iterrows():
            desc_text = str(row[desc_col])
            # gather all Standard_X fields in order if present
            std_parts = []
            grade_parts = []
            finish_parts = []
            # collect any fields that match pattern Type_1..N
            type_cols = [c for c in df.columns if re.match(r"Type_\d+$", c)]
            type_cols_sorted = sorted(type_cols, key=lambda x: int(x.split("_")[1]))  # Type_1, Type_2 ...
            for idx_col, tcol in enumerate(type_cols_sorted, start=1):
                sval = row.get(f"Standard_{idx_col}", "")
                if sval and sval not in std_parts:
                    std_parts.append(str(sval))
                gval = row.get(f"Grade_{idx_col}", "")
                if gval and gval not in grade_parts:
                    grade_parts.append(str(gval))
                fval = row.get(f"Finish_{idx_col}", "")
                if fval and fval not in finish_parts:
                    finish_parts.append(str(fval))

            # fallback: if no Standard_X collected, try the main "Standard" column
            if not std_parts and row.get("Standard"):
                std_parts.append(row.get("Standard"))
            # fallback for grade / finish: main columns
            if not grade_parts and row.get("Grade"):
                grade_parts.append(row.get("Grade"))
            if not finish_parts and row.get("Finish"):
                finish_parts.append(row.get("Finish"))

            std_merged = " / ".join(std_parts) if std_parts else ""
            grade_merged = " / ".join(grade_parts) if grade_parts else ""
            finish_merged = " / ".join(finish_parts) if finish_parts else ""

            final_rows.append({
                "Sr No": i+1,
                "Description": desc_text,
                "Dimension Standard": std_merged,
                "Grade": grade_merged,
                "Finish": finish_merged
            })

        final_df = pd.DataFrame(final_rows, columns=["Sr No", "Description", "Dimension Standard", "Grade", "Finish"])
        st.markdown("### Final Table")
        st.dataframe(final_df, use_container_width=True)

        # Also provide CSV download
        csv = final_df.to_csv(index=False).encode("utf-8")
        st.download_button("Download Final CSV", data=csv, file_name="final_table.csv", mime="text/csv")

    else:
        st.info("When you're ready, click 'Generate Final Table' to merge per-type columns into the final 5-column table.")
