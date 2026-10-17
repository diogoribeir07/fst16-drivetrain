import os
import sys
import re
import json
import pandas as pd

base_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, base_dir)

app_path = os.path.join(base_dir, "app.py")

with open(app_path, "r", encoding="utf-8") as f:
    src = f.read()

# 1. Check forbidden bureaucratic terms in app.py
print("1. Checking forbidden bureaucratic terms in app.py...")
forbidden = [
    "Component Status",
    "DEFAULT_STATUSES",
    "DL Review",
    "DL Status",
    "DL Decision",
    "DL Recommendation",
    "Component Trade-Off Study",
    "sign-off",
    "Approved Components",
    "Under Trade-Off Study",
    "Ready for DL Review",
    "Ready for Review",
    "EXCEL_TASKS",
    "load_design_tasks"
]

for term in forbidden:
    occurrences = [i for i in range(len(src)) if src.startswith(term, i)]
    if occurrences:
        for idx in occurrences:
            line = src[src.rfind('\n', 0, idx)+1 : src.find('\n', idx)]
            if "drop(columns=" in line or "in df.columns" in line:
                continue
            raise AssertionError(f"Found forbidden term '{term}' in app.py:\n  {line.strip()}")

print("   PASSED: Zero forbidden bureaucratic / DL terms in app.py.")

# 2. Check forbidden bureaucratic terms across all Excel files
print("2. Checking forbidden terms across all Excel files...")
excel_files = ["FST16_DT_Parts_Export.xlsx", "FST16_Trash_Internal.xlsx", "FST16_Assemblies_Export_DL.xlsx", "FST16_Requirements.xlsx"]
excel_forbidden = ["Component Status", "DL", "sign-off", "signoff", "Approved Components", "Under Trade-Off Study", "Ready for DL Review"]

for ef in excel_files:
    ef_path = os.path.join(base_dir, ef)
    if os.path.exists(ef_path):
        df_check = pd.read_excel(ef_path)
        assert "Component Status" not in df_check.columns, f"'Component Status' column found in {ef}!"
        for col in df_check.columns:
            for r_idx, cell in enumerate(df_check[col]):
                s_cell = str(cell)
                for term in excel_forbidden:
                    if term == "DL":
                        if re.search(r'\bDL\b', s_cell):
                            raise AssertionError(f"Forbidden term 'DL' found in {ef} [{col}] row {r_idx}:\n  {s_cell}")
                    elif term.lower() in s_cell.lower():
                        raise AssertionError(f"Forbidden term '{term}' found in {ef} [{col}] row {r_idx}:\n  {s_cell}")

print("   PASSED: Zero forbidden bureaucratic terms in all Excel files.")

# 3. Check emojis in app.py and Excel files
print("3. Checking emojis in app.py and Excel files...")
emojis = [c for c in src if ord(c) > 127 and (ord(c) >= 126976 or (ord(c) >= 9728 and ord(c) <= 10175))]
if emojis:
    raise AssertionError(f"Found emojis in app.py: {emojis}")

for ef in excel_files:
    ef_path = os.path.join(base_dir, ef)
    if os.path.exists(ef_path):
        df_check = pd.read_excel(ef_path)
        for col in df_check.columns:
            for r_idx, cell in enumerate(df_check[col]):
                s_cell = str(cell)
                e_matches = [c for c in s_cell if ord(c) > 127 and (ord(c) >= 126976 or (ord(c) >= 9728 and ord(c) <= 10175))]
                if e_matches:
                    raise AssertionError(f"Found emojis in {ef} [{col}] row {r_idx}: {e_matches}")

print("   PASSED: Zero emojis found.")

# 4. Check UI titles, headers, tabs, buttons, inputs, metrics, expanders length (strictly 1 to 3 words)
print("4. Checking UI element lengths (strictly 1 to 3 words)...")
patterns = [
    (r'st\.title\(f?[\'"]([^\'"]+)[\'"]\)', 'st.title'),
    (r'st\.subheader\(f?[\'"]([^\'"]+)[\'"]\)', 'st.subheader'),
    (r'st\.sidebar\.markdown\(f?[\'"]\#+\s*([^\'"]+)[\'"]\)', 'st.sidebar.markdown header'),
    (r'st\.markdown\(f?[\'"]\#+\s*([^\'"]+)[\'"]\)', 'st.markdown header'),
    (r'st\.tabs\(\[([^\]]+)\]\)', 'st.tabs'),
    (r'st\.button\(f?[\'"]([^\'"]+)[\'"]', 'st.button'),
    (r'st\.sidebar\.button\(f?[\'"]([^\'"]+)[\'"]', 'st.sidebar.button'),
    (r'st\.form_submit_button\(f?[\'"]([^\'"]+)[\'"]', 'st.form_submit_button'),
    (r'st\.expander\(f?[\'"]([^\'"]+)[\'"]', 'st.expander'),
    (r'(?:\w+\.)?metric\(f?[\'"]([^\'"]+)[\'"]', 'metric'),
    (r'st\.selectbox\(f?[\'"]([^\'"]+)[\'"]', 'st.selectbox'),
    (r'st\.text_input\(f?[\'"]([^\'"]+)[\'"]', 'st.text_input'),
    (r'st\.number_input\(f?[\'"]([^\'"]+)[\'"]', 'st.number_input'),
    (r'st\.text_area\(f?[\'"]([^\'"]+)[\'"]', 'st.text_area'),
    (r'st\.radio\(f?[\'"]([^\'"]+)[\'"]', 'st.radio'),
    (r'render_image_input\(f?[\'"]([^\'"]+)[\'"]', 'render_image_input label'),
    (r'render_image_input\([^)]*label=f?[\'"]([^\'"]+)[\'"]', 'render_image_input kwarg'),
]

for pat, kind in patterns:
    matches = re.finditer(pat, src)
    for m in matches:
        text = m.group(1).strip()
        if kind == 'st.tabs':
            if ' for ' in text:
                continue
            items = [x.strip(' \n\'"') for x in text.split(',')]
            for item in items:
                words = item.split()
                if len(words) > 3:
                    raise AssertionError(f"Tab '{item}' has {len(words)} words (>3 words)")
        else:
            clean_text = re.sub(r'\{[^}]+\}', 'X', text).strip()
            words = clean_text.split()
            if len(words) > 3:
                raise AssertionError(f"{kind} '{text}' has {len(words)} words (>3 words)")

# Check all lines with markdown headers
for i, line in enumerate(src.splitlines()):
    l = line.strip()
    m_h = re.search(r'^\s*st\.(?:sidebar\.)?(?:markdown|write)\(f?[\'"]\s*(#{1,6})\s+([^"\'\n]+)', l)
    if m_h:
        h_text = m_h.group(2).strip()
        h_text_clean = re.sub(r'<[^>]+>.*', '', h_text).strip()
        clean_words = re.sub(r'\{[^}]+\}', 'X', h_text_clean).split()
        if len(clean_words) > 3:
            raise AssertionError(f"Markdown header on line {i+1} '{h_text_clean}' has {len(clean_words)} words (>3 words)")

print("   PASSED: All UI titles, headers, tabs, buttons, inputs, metrics, and expanders are strictly 1 to 3 words.")

# 5. Check backend functions and data integrity
print("5. Testing backend functions and data integrity...")
exec_code = src.split("st.set_page_config")[0]
env = {"__file__": app_path}
exec(exec_code, env)

# Measurements
p_m = env["parse_measurements"]("Feature: Bore\nValue: 80 mm")
assert len(p_m) == 2, "parse_measurements failed"
f_m = env["format_measurements"](p_m)
assert "Bore" in f_m, "format_measurements failed"

# Tolerances
test_tol = [{"Interface / Feature": "Bore", "Nominal (mm)": "100", "Fit / Tolerance": "H7", "Roughness / Notes": "Ra 1.6"}]
f_t = env["format_tolerancing_data"](test_tol)
p_t = env["parse_tolerancing_data"](f_t)
assert len(p_t) == 1 and p_t[0]["Nominal (mm)"] == "100", "tolerances parsing failed"

# Load parts
df_parts = env["load_parts"]()
print(f"   Loaded df_parts: {len(df_parts)} rows, {len(df_parts.columns)} columns")
assert "Component Status" not in df_parts.columns, "Component Status MUST NOT be in df_parts.columns!"
assert "Justification" in df_parts.columns, "Justification MUST be in df_parts.columns"
assert "Part Name" in df_parts.columns, "Part Name MUST be in df_parts.columns"

# Check defaults exist
assert len(env["DEFAULT_CUSTOM_TOLERANCES"]) == 0, "DEFAULT_CUSTOM_TOLERANCES should be empty"
assert len(env["DEFAULT_OEM_TOLERANCES"]) == 0, "DEFAULT_OEM_TOLERANCES should be empty"
active_tols = [str(x).strip() for x in df_parts["Tolerancing Data"].dropna() if str(x).strip() not in ["", "nan", "[]"]]
assert len(active_tols) == 0, f"Expected 0 active tolerance entries, got {len(active_tols)}"
assert "UP_03" in env["DEFAULT_JUSTIFICATIONS"], "UP_03 missing in DEFAULT_JUSTIFICATIONS"
assert "UP_04" in env["DEFAULT_JUSTIFICATIONS"], "UP_04 missing in DEFAULT_JUSTIFICATIONS"
assert "UP_01" not in env["DEFAULT_JUSTIFICATIONS"], "UP_01 should NOT be in DEFAULT_JUSTIFICATIONS"

# Check clean justifications for UP_03 and UP_04, and empty for UP_01
up01_just = [str(x).strip() for x in df_parts[df_parts["Part Name"] == "UP_01"]["Justification"].dropna() if str(x).strip() not in ["", "nan"]]
assert len(up01_just) == 0, f"UP_01 should NOT have any justification, found: {up01_just}"

up03_just = df_parts[df_parts["Part Name"] == "UP_03"]["Justification"].iloc[0]
up04_just = df_parts[df_parts["Part Name"] == "UP_04"]["Justification"].iloc[0]
assert "Under Trade-Off Study" not in up03_just and "DL" not in up03_just, "UP_03 justification has bureaucratic text"
assert "Under Trade-Off Study" not in up04_just and "DL" not in up04_just, "UP_04 justification has bureaucratic text"
assert len(str(up03_just).strip()) > 10, "UP_03 justification missing"
assert len(str(up04_just).strip()) > 10, "UP_04 justification missing"

# Load assemblies
df_asm = env["load_assemblies"]()
print(f"   Loaded df_assemblies: {len(df_asm)} rows")
assert not df_asm.empty, "df_assemblies should not be empty"

# Load trash
df_trash = env["load_trash"]()
print(f"   Loaded df_trash: {len(df_trash)} rows")
assert "Component Status" not in df_trash.columns, "Component Status MUST NOT be in df_trash.columns!"

# Load requirements
df_req = env["load_requirements"]()
print(f"   Loaded df_requirements: {len(df_req)} rows, {len(df_req.columns)} columns")
assert len(df_req) >= 45, f"Expected >= 45 requirements, got {len(df_req)}"
assert "ID" in df_req.columns and "Requirement" in df_req.columns, "Requirements missing required columns"
assert "Component Status" not in df_req.columns, "Component Status MUST NOT be in df_requirements.columns!"

# Test editing and saving parts without LossySetitemError, TypeError or NameError
for test_p_name in ["UP_01", "HB_01", "61816", "UP_03", "85X100X9 HMSA10 V"]:
    matching_indices = df_parts[df_parts["Part Name"] == test_p_name].index.tolist()
    if not matching_indices:
        continue
    idx_to_test = matching_indices[0]
    row_to_edit = df_parts.loc[idx_to_test]
    
    # Simulate variables as in app.py tab_editar
    user_edit_stiff_vals = {}
    updated_lc_images_pasted = {}
    curr_lc_images = env["parse_lc_images"](row_to_edit.get("LC Images", ""))
    curr_struct = env["parse_structural_data"](row_to_edit.get("Stiffness Data", ""))
    
    # Simulate save logic
    text_cols_save = ["Description", "Specifications", "Tolerancing Data", "Stiffness Data", "LC Images", "Justification", "Image Path"]
    for col_name in text_cols_save:
        if col_name in df_parts.columns and df_parts[col_name].dtype != object:
            df_parts[col_name] = df_parts[col_name].astype(object)
            
    df_parts.at[idx_to_test, "Tolerancing Data"] = env["format_tolerancing_data"]([])
    if test_p_name == "UP_01":
        df_parts.at[idx_to_test, 'Stiffness Data'] = env["format_stiffness_data"](user_edit_stiff_vals)
    else:
        if user_edit_stiff_vals:
            df_parts.at[idx_to_test, 'Stiffness Data'] = json.dumps(user_edit_stiff_vals)
        elif curr_struct:
            df_parts.at[idx_to_test, 'Stiffness Data'] = json.dumps(curr_struct)
            
    df_parts.at[idx_to_test, "Tolerancing Data"] = ""

print("ALL TEST SUITES PASSED WITH FLYING COLORS!")
