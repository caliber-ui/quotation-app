import streamlit as st
import json
import re
from rapidfuzz import process

# -------------------------
# Load & Normalize JSON
# -------------------------
@st.cache_data
def load_catalogue(file):
    raw = json.load(file)
    normalized = []

    for entry in raw:
        # Case 1: Already structured
        if "screw_type" in entry:
            if "stud price" in entry.get("screw_type", "").lower():
                entry["screw_type"] = "Stud Screw"
                # Clean metric
                if "dimensions_in_metric" in entry:
                    cleaned = []
                    for dim in entry["dimensions_in_metric"]:
                        new_dia = {k.replace(" ", ""): v for k, v in dim.get("diameter", {}).items()}
                        cleaned.append({"length_mm": dim.get("length_mm"), "diameter": new_dia})
                    entry["dimensions_in_metric"] = cleaned
                # Clean inch
                if "dimensions_in_inches" in entry:
                    cleaned = []
                    for dim in entry["dimensions_in_inches"]:
                        new_dia = {k.strip(): v for k, v in dim.get("diameter", {}).items()}
                        cleaned.append({"length": dim.get("length"), "diameter": new_dia})
                    entry["dimensions_in_inches"] = cleaned
            normalized.append(entry)

        # Case 2: Flat (Rivet style)
        elif "title" in entry and ("dimensions_unit" in entry or "data" in entry):
            screw_type = entry["title"]
            dim_unit = entry.get("dimensions_unit", "").lower()
            unit = entry.get("approx_count_unit", "N/A")
            if "stud" in screw_type.lower():
                screw_type = "Stud Screw"
            new_entry = {"screw_type": screw_type, "standard": "N/A", "unit": unit,
                         "dimensions_in_metric": [], "dimensions_in_inches": []}
            if dim_unit.startswith("inch"):
                for d in entry.get("data", []):
                    new_entry["dimensions_in_inches"].append({
                        "length": d.get("length"),
                        "diameter": {k.replace("diameter_", "").strip(): v for k, v in d.items() if k.startswith("diameter_")}
                    })
            elif dim_unit.startswith("metric"):
                for d in entry.get("data", []):
                    diam_clean = {k.replace("diameter_", "").replace(" ", ""): v for k, v in d.items() if k.startswith("diameter_")}
                    new_entry["dimensions_in_metric"].append({"length_mm": d.get("length_mm"), "diameter": diam_clean})
            normalized.append(new_entry)

        # Case 3: Nut structures
        elif "approx_count_per_50_kgs" in entry or "hex_locknuts_bsw_bsf_approx_weight_per_100_pcs" in entry:
            screw_type = entry.get("title", "Nut")

            def looks_like_inches(s):
                s = str(s)
                return '"' in s or '/' in s or "ba" in s.lower()

            if "approx_count_per_50_kgs" in entry:
                for nut_type, data_list in entry["approx_count_per_50_kgs"].items():
                    ne = {"screw_type": f"{screw_type} - {nut_type}", "standard": "N/A",
                          "unit": "Approx. Count per 50 kgs.", "dimensions_in_metric": [], "dimensions_in_inches": []}
                    for d in data_list:
                        size, val = d.get("size"), d.get("specification") or d.get("height_d") or d.get("count")
                        target = "dimensions_in_inches" if looks_like_inches(size) else "dimensions_in_metric"
                        ne[target].append({"length_mm": None, "diameter": {size: val}})
                    normalized.append(ne)

            if "hex_locknuts_bsw_bsf_approx_weight_per_100_pcs" in entry:
                for nut_type, data_list in entry["hex_locknuts_bsw_bsf_approx_weight_per_100_pcs"].items():
                    ne = {"screw_type": f"{screw_type} - {nut_type}", "standard": "N/A",
                          "unit": "Approx. Weight per 100 pcs.", "dimensions_in_metric": [], "dimensions_in_inches": []}
                    for d in data_list:
                        size, val = d.get("size"), d.get("weight") or d.get("count")
                        target = "dimensions_in_inches" if looks_like_inches(size) else "dimensions_in_metric"
                        ne[target].append({"length_mm": None, "diameter": {size: val}})
                    normalized.append(ne)

        # Case 4: Washer structure
        elif "title" in entry and "unit" in entry:
            for ftype, data_list in entry.items():
                if ftype in ["title", "unit"]:
                    continue
                ne = {"screw_type": ftype, "standard": "N/A", "unit": entry.get("unit", "N/A"),
                      "dimensions_in_metric": [], "dimensions_in_inches": []}
                for d in data_list:
                    size, wt = d.get("size"), d.get("weight")
                    ne["dimensions_in_metric"].append({"length_mm": None, "diameter": {size: wt}})
                normalized.append(ne)
    return normalized


# -------------------------
# Utility Functions
# -------------------------
def find_rate(entries, length_val, diameter, dim_type):
    dims_key = "dimensions_in_metric" if dim_type == "metric" else "dimensions_in_inches"
    for e in entries:
        for dim in e.get(dims_key, []):
            if dim.get("length_mm") == length_val or dim.get("length") == length_val or dim.get("length_mm") is None:
                val = dim.get("diameter", {}).get(diameter)
                if val:
                    return val
    return None


def sort_diameters(diam_list, dim_type):
    def metric_key(x):
        try:
            return int(''.join(ch for ch in x if ch.isdigit()) or 9999)
        except:
            return 9999
    def inch_key(x):
        try:
            val = x.replace('"', "")
            if "/" in val:
                n, d = val.split("/")
                return float(n) / float(d)
            return float(val)
        except:
            return 9999
    return sorted(list(dict.fromkeys(diam_list)), key=metric_key if dim_type == "metric" else inch_key)


def normalize_name(name):
    return name.split(" - ")[0].strip().lower()


def parse_combo_input(inp):
    words = re.findall(r"[A-Za-z]+|\d+", inp.lower())
    combo, qty = [], 1
    for w in words:
        if w.isdigit():
            qty = int(w)
        elif w in ["stud", "nut", "washer", "bolt", "screw"]:
            combo.append({"type": w, "count": qty})
            qty = 1
    return combo


# -------------------------
# Streamlit App
# -------------------------
st.title("🔩 Screw Weight & Price Calculator from Catalogue")

uploaded_file = st.file_uploader("Upload catalogue.json", type="json")

if uploaded_file:
    catalogue = load_catalogue(uploaded_file)
    screw_input = st.text_input("Enter Screw/Washer/Nut Type or Combo (e.g. STUD WITH 2 NUT AND WASHER)")

    if screw_input:
        combo_items = parse_combo_input(screw_input)

        # -------------------------
        # COMBO MODE (Corrected)
        # -------------------------
        if len(combo_items) > 1:
            st.header("⚙️ Combo Mode")
            total_weight_pc, total_price_all = 0.0, 0.0

            for comp in combo_items:
                ctype, ccount = comp["type"], comp["count"]
                st.subheader(f"🔹 {ctype.title()} (x{ccount})")

                # ----------- FIXED: Handle subtype for all components -------------
                entries_filtered = [e for e in catalogue if ctype in e.get("screw_type", "").lower()]
                subtype_labels = sorted(set([e["screw_type"] for e in entries_filtered]))

                if len(subtype_labels) > 1:
                    chosen_type = st.selectbox(f"Select {ctype.title()} Type", subtype_labels, key=f"type_{ctype}")
                    entries = [e for e in entries_filtered if e["screw_type"] == chosen_type]
                else:
                    entries = entries_filtered
                # -------------------------------------------------------------------

                if not entries:
                    st.warning(f"No data found for {ctype}")
                    continue

                has_metric = any(e.get("dimensions_in_metric") for e in entries)
                has_inches = any(e.get("dimensions_in_inches") for e in entries)
                dim_type = "metric"
                if has_metric and has_inches:
                    dim_type = st.radio(f"{ctype} Dimension Type", ["metric", "inches"], key=f"dim_{ctype}")
                elif has_inches:
                    dim_type = "inches"

                # Diameters
                diam_list = []
                for e in entries:
                    dims = e.get("dimensions_in_metric", []) if dim_type == "metric" else e.get("dimensions_in_inches", [])
                    for d in dims:
                        diam_list.extend(list(d.get("diameter", {}).keys()))
                diam_list = sort_diameters(diam_list, dim_type)
                diameter = st.selectbox(f"{ctype.title()} Diameter", diam_list, key=f"dia_{ctype}")

                # Length (Fixed for all types)
                lengths = []
                for e in entries:
                    dims = e.get("dimensions_in_metric", []) if dim_type == "metric" else e.get("dimensions_in_inches", [])
                    for d in dims:
                        if diameter in d.get("diameter", {}):
                            length_val = d.get("length_mm") or d.get("length")
                            if length_val is not None:
                                lengths.append(length_val)

                lengths = sorted(set(lengths), key=lambda x: int(x) if isinstance(x, (int, str)) and str(x).isdigit() else str(x))
                if lengths:
                    if len(lengths) == 1:
                        length_val = lengths[0]
                        st.info(f"Only one length available → Auto-selected: {length_val}")
                    else:
                        length_choice = st.selectbox(f"{ctype.title()} Length", ["--manual--"] + [str(l) for l in lengths], key=f"len_{ctype}")
                        if length_choice == "--manual--":
                            if dim_type == "metric":
                                length_val = st.number_input(f"Enter Length for {ctype} (mm)", min_value=1, step=1, key=f"manual_len_{ctype}")
                            else:
                                length_val = st.text_input(f"Enter Length for {ctype} (inches)", key=f"manual_len_{ctype}")
                        else:
                            length_val = int(length_choice) if dim_type == "metric" and str(length_choice).isdigit() else length_choice
                else:
                    length_val = None
                    st.info(f"No length available for {ctype} → Proceeding without length.")

                # Rate and price
                rate_val = find_rate(entries, length_val, diameter, dim_type)
                if not rate_val:
                    st.warning(f"No rate found for {ctype}")
                    continue
                unit = entries[0].get("unit", "").lower()
                if "price per piece" in unit:
                    per_pc_wt = float(rate_val)
                    st.success(f"Per Piece Weight (direct): {per_pc_wt:.4f} kg")
                elif "approx. weight per 100" in unit or "100 nos" in unit:
                    per_pc_wt = float(rate_val) / 100
                    st.success(f"Weight per 100 pcs: {float(rate_val)} kg")
                elif "approx. count per 50" in unit:
                    pcs_per_50kg = float(rate_val)
                    per_pc_wt = 50.0 / pcs_per_50kg
                    st.success(f"Pieces per 50kg: {int(pcs_per_50kg)}")
                else:
                    per_pc_wt = float(rate_val)
                    st.info(f"Raw value used: {per_pc_wt}")

                rate_per_kg = st.number_input(f"Rate per kg for {ctype}", min_value=0.0, step=0.1, key=f"rate_{ctype}")
                qty = st.number_input(f"Quantity for {ctype}", min_value=1, value=1, key=f"qty_{ctype}")

                comp_weight = per_pc_wt * ccount * qty
                comp_price = comp_weight * rate_per_kg
                total_weight_pc += per_pc_wt * ccount
                total_price_all += comp_price

                st.success(f"{ctype.title()} weight per piece: {per_pc_wt:.4f} kg")
                st.success(f"{ctype.title()} total weight: {comp_weight:.4f} kg")
                st.success(f"{ctype.title()} total price: ₹{comp_price:.2f}")

            st.markdown("---")
            st.success(f"⚖️ Combined per-piece weight: {total_weight_pc:.4f} kg")
            st.success(f"💰 Combined total price: ₹{total_price_all:.2f}")
        # -------------------------
        # SINGLE ITEM MODE (your full original logic)
        # -------------------------
        else:
            st.header("🔧 Single Item Mode")

            screw_types = [item.get("screw_type", "") for item in catalogue if "screw_type" in item]
            matches = process.extract(screw_input, screw_types, limit=5)
            matches = [(m, sc, idx) for m, sc, idx in matches if sc > 70]

            if not matches:
                st.error("❌ No good match found. Try typing more clearly.")
                st.stop()

            match_labels = [f"{m} ({sc:.1f}%)" for m, sc, _ in matches]
            chosen_match = st.selectbox("Select Closest Type", match_labels)
            best_match_name = matches[match_labels.index(chosen_match)][0]
            best_match_base = normalize_name(best_match_name)

            matched_entries = [item for item in catalogue if normalize_name(item.get("screw_type", "")) == best_match_base]

            entry_labels = [f"{entry.get('screw_type', 'Unknown')} | Standard: {entry.get('standard', 'N/A')} | Unit: {entry.get('unit', 'N/A')}" for entry in matched_entries]
            chosen_entry_label = st.selectbox("Select Entry (Standard)", entry_labels)
            chosen_entry_idx = entry_labels.index(chosen_entry_label)
            relevant_entries = [matched_entries[chosen_entry_idx]]

            has_metric = any(entry.get("dimensions_in_metric") and len(entry.get("dimensions_in_metric")) > 0 for entry in relevant_entries)
            has_inches = any(entry.get("dimensions_in_inches") and len(entry.get("dimensions_in_inches")) > 0 for entry in relevant_entries)

            if has_metric and has_inches:
                dim_type = st.radio("Select Dimension Type", ["metric", "inches"])
            elif has_metric:
                dim_type = "metric"
                st.info("Only metric dimensions available for this type.")
            elif has_inches:
                dim_type = "inches"
                st.info("Only inch dimensions available for this type.")
            else:
                st.error("No dimension data available for this type.")
                st.stop()

            diameters = []
            if dim_type == "metric":
                for entry in relevant_entries:
                    for dim in entry.get("dimensions_in_metric", []):
                        diameters.extend(list(dim.get("diameter", {}).keys()))
            else:
                for entry in relevant_entries:
                    for dim in entry.get("dimensions_in_inches", []):
                        diameters.extend(list(dim.get("diameter", {}).keys()))
            diameters = sort_diameters(diameters, dim_type)

            if len(diameters) == 1:
                diameter = diameters[0]
                st.info(f"Only one diameter available → Auto-selected: {diameter}")
            else:
                diameter_choice = st.selectbox("Select Diameter", ["--manual--"] + diameters)
                if diameter_choice == "--manual--":
                    diameter = st.text_input("Enter Diameter (e.g., M6 or 1/4\")")
                else:
                    diameter = diameter_choice

            lengths = []
            has_length_values = False
            if dim_type == "metric":
                for entry in relevant_entries:
                    for dim in entry.get("dimensions_in_metric", []):
                        if diameter in dim.get("diameter", {}):
                            if dim.get("length_mm") is not None:
                                lengths.append(dim.get("length_mm"))
                                has_length_values = True
            else:
                for entry in relevant_entries:
                    for dim in entry.get("dimensions_in_inches", []):
                        if diameter in dim.get("diameter", {}):
                            if dim.get("length") is not None:
                                lengths.append(dim.get("length"))
                                has_length_values = True

            lengths = sorted(set(lengths), key=lambda x: int(x) if isinstance(x, (int, str)) and str(x).isdigit() else str(x))
            if not has_length_values:
                length_val = None
                st.info("No length dimension for this item. Proceeding without length.")
            else:
                if len(lengths) == 1:
                    length_val = lengths[0]
                    st.info(f"Only one length available → Auto-selected: {length_val}")
                else:
                    length_choice = st.selectbox("Select Length", ["--manual--"] + [str(l) for l in lengths])
                    if length_choice == "--manual--":
                        if dim_type == "metric":
                            length_val = st.number_input("Enter Length (mm)", min_value=1, step=1)
                        else:
                            length_val = st.text_input("Enter Length (inches, e.g. 1\")")
                    else:
                        length_val = int(length_choice) if dim_type == "metric" and str(length_choice).isdigit() else length_choice

            quantity = st.number_input("Enter Quantity", min_value=1, value=1)
            rate_per_kg = st.number_input("Enter Rate Price (per kg)", min_value=0.0, step=0.1)

            if st.button("Calculate"):
                rate_value = find_rate(relevant_entries, length_val, diameter, dim_type)
                if rate_value:
                    chosen_unit = relevant_entries[0].get("unit", "").strip().lower()
                    if "price per piece" in chosen_unit:
                        weight_per_pc = float(rate_value)
                        st.success(f"Per Piece Weight (direct): {weight_per_pc:.4f} kg")
                    elif "approx. count per 50" in chosen_unit:
                        pcs_per_50kg = float(rate_value)
                        weight_per_pc = 50.0 / pcs_per_50kg
                        st.success(f"Pieces per 50kg: {int(pcs_per_50kg)}")
                    elif "approx. weight per 100" in chosen_unit or "100 nos" in chosen_unit:
                        weight_per_100 = float(rate_value)
                        weight_per_pc = weight_per_100 / 100.0
                        st.success(f"Weight per 100 pcs: {weight_per_100} kg")
                    else:
                        pcs_per_50kg = float(rate_value)
                        weight_per_pc = 50.0 / pcs_per_50kg
                        st.success(f"Pieces per 50kg (inferred): {int(pcs_per_50kg)}")

                    total_weight = weight_per_pc * quantity
                    per_piece_price = weight_per_pc * rate_per_kg
                    total_price = per_piece_price * quantity

                    st.success(f"Weight per piece: {weight_per_pc:.4f} kg")
                    st.success(f"💰 Price per piece: {per_piece_price:.2f}")
                    st.success(f"Total weight for {quantity} pcs: {total_weight:.4f} kg")
                    st.success(f"💰 Final Price for {quantity} pcs: {total_price:.2f}")
                else:
                    st.error("❌ No matching rate found for this diameter/length.")
