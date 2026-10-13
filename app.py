import streamlit as st
import pandas as pd
import os
import json
from datetime import datetime
from PIL import Image
try:
    from PIL import ImageGrab
except ImportError:
    ImageGrab = None

# Configuration
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PARTS = os.path.join(BASE_DIR, "FST16_DT_Parts_Export.xlsx")
EXCEL_ASSEMBLIES = os.path.join(BASE_DIR, "FST16_Assemblies_Export_DL.xlsx")
EXCEL_TRASH = os.path.join(BASE_DIR, "FST16_Trash_Internal.xlsx")
IMAGE_DIR = os.path.join(BASE_DIR, "images")

if not os.path.exists(IMAGE_DIR):
    os.makedirs(IMAGE_DIR)

def resolve_image_path(path_str):
    if not path_str or pd.isna(path_str):
        return None
    clean_path = str(path_str).replace("\\", "/").strip()
    if os.path.isabs(clean_path) and os.path.exists(clean_path):
        return clean_path
    full_path = os.path.join(BASE_DIR, clean_path)
    if os.path.exists(full_path):
        return full_path
    if os.path.exists(clean_path):
        return clean_path
    return None

def grab_image_from_clipboard():
    if ImageGrab is None:
        return None
    try:
        img = ImageGrab.grabclipboard()
        if isinstance(img, Image.Image):
            return img
        elif isinstance(img, list) and len(img) > 0:
            if os.path.exists(img[0]):
                return Image.open(img[0])
    except Exception:
        pass
    return None

def render_image_input(label="Attach Image", key="img_input"):
    clipboard_key = f"clipboard_{key}"
    col_up, col_btn = st.columns([3, 1])
    with col_up:
        uploaded_file = st.file_uploader(
            f"{label} (Arrastar & largar, ou clicar e premir Ctrl+V)",
            type=["jpg", "png", "jpeg"],
            key=f"uploader_{key}"
        )
    with col_btn:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if st.button("Colar do Clipboard", key=f"btn_paste_{key}", use_container_width=True):
            pasted = grab_image_from_clipboard()
            if pasted is not None:
                st.session_state[clipboard_key] = pasted
                st.toast("Imagem colada do clipboard com sucesso.")
            else:
                st.warning("Nenhuma imagem detetada no clipboard do sistema. No browser, clica na caixa de upload e prime Ctrl+V.")

    pasted_img = st.session_state.get(clipboard_key, None)
    if pasted_img is not None and uploaded_file is None:
        c_prev, c_clear = st.columns([3, 1])
        with c_prev:
            st.image(pasted_img, caption="Imagem colada do clipboard", width=260)
        with c_clear:
            if st.button("Limpar Imagem Colada", key=f"btn_clear_{key}"):
                st.session_state.pop(clipboard_key, None)
                st.rerun()

    return uploaded_file, pasted_img

def save_uploaded_or_pasted_image(uploaded_file, pasted_image, part_name, version_str):
    version_clean = str(version_str).lstrip('vV').replace('.', '_')
    if uploaded_file is not None:
        filename = f"{part_name}_v{version_clean}_{uploaded_file.name}"
        filepath = os.path.join(IMAGE_DIR, filename)
        with open(filepath, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return f"images/{filename}"
    elif pasted_image is not None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{part_name}_v{version_clean}_clipboard_{timestamp}.png"
        filepath = os.path.join(IMAGE_DIR, filename)
        pasted_image.save(filepath, format="PNG")
        return f"images/{filename}"
    return ""

from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

def save_formatted_excel(df, filename, table_name="DataTable"):
    try:
        # Auto-organize Parts by Subassembly and Part Name
        if "Subassembly" in df.columns and "Part Name" in df.columns:
            df = df.sort_values(by=["Subassembly", "Part Name"]).reset_index(drop=True)
            
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')
            worksheet = writer.sheets['Data']
            
            for idx, col in enumerate(df.columns, 1):
                col_letter = get_column_letter(idx)
                
                # Widths and Wrap Text
                if col in ["Description", "Constraints", "Part Configuration", "Image Path", "Stiffness Data"]:
                    worksheet.column_dimensions[col_letter].width = 45
                    for row in range(2, len(df) + 2):
                        worksheet.cell(row=row, column=idx).alignment = Alignment(wrap_text=True, vertical='top')
                else:
                    max_len = max(df[col].astype(str).map(len).max() if not df.empty else 0, len(str(col)))
                    worksheet.column_dimensions[col_letter].width = max(min(max_len + 5, 25), 14)
                    for row in range(2, len(df) + 2):
                        worksheet.cell(row=row, column=idx).alignment = Alignment(horizontal="center", vertical='top')
                        
            # Apply a beautiful Excel Table style with alternating colors and filters
            if not df.empty:
                max_row = len(df) + 1
                max_col = len(df.columns)
                tab = Table(displayName=table_name, ref=f"A1:{get_column_letter(max_col)}{max_row}")
                style = TableStyleInfo(
                    name="TableStyleMedium2", showFirstColumn=False,
                    showLastColumn=False, showRowStripes=True, showColumnStripes=False
                )
                tab.tableStyleInfo = style
                worksheet.add_table(tab)
                
    except PermissionError:
        st.error(f"PERMISSION DENIED: Please close '{filename}' in Microsoft Excel before interacting with the app!")
        st.stop()

# ==================== FEA STIFFNESS & DISPLACEMENTS (UP_01) ====================
LC_KEYS = ["LC1", "LC2", "LC3", "LC4", "LC5", "LC6", "LC7", "LC8", "LC9"]

LC_INFO = {
    "LC1": {"desc": "Pure Braking", "ref_fst15": 0.2955, "default_fst16": 0.3211},
    "LC2": {"desc": "Pure Acceleration", "ref_fst15": 0.4246, "default_fst16": 0.4234},
    "LC3": {"desc": "Pure Cornering", "ref_fst15": 0.3381, "default_fst16": 0.5343},
    "LC4": {"desc": "Braking + Cornering 1", "ref_fst15": 0.4214, "default_fst16": 0.4075},
    "LC5": {"desc": "Braking + Cornering 2", "ref_fst15": 0.4229, "default_fst16": 0.3582},
    "LC6": {"desc": "Acceleration + Cornering 1", "ref_fst15": 0.3262, "default_fst16": 0.4746},
    "LC7": {"desc": "Acceleration + Cornering 2", "ref_fst15": 0.3574, "default_fst16": 0.5312},
    "LC8": {"desc": "1kN Fy", "ref_fst15": 0.0, "default_fst16": 0.0},
    "LC9": {"desc": "-1kN Fy", "ref_fst15": 0.0, "default_fst16": 0.0},
}

DEFAULT_FST15_DISPLACEMENTS = {k: LC_INFO[k]["ref_fst15"] for k in LC_KEYS}
DEFAULT_FST16_DISPLACEMENTS = {k: LC_INFO[k]["default_fst16"] for k in LC_KEYS}

def parse_stiffness_data(val):
    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "nan":
        return {}
    val_str = str(val).strip()
    try:
        raw = json.loads(val_str)
        res = {}
        for lc in LC_KEYS:
            if lc in raw:
                item = raw[lc]
                if isinstance(item, dict):
                    res[lc] = float(item.get("disp", 0.0))
                else:
                    res[lc] = float(item)
        return res
    except Exception:
        pass
    return {}

def format_stiffness_data(disp_dict):
    cleaned = {}
    for lc in LC_KEYS:
        if lc in disp_dict and disp_dict[lc] is not None:
            try:
                if isinstance(disp_dict[lc], dict):
                    cleaned[lc] = round(float(disp_dict[lc].get("disp", 0.0)), 4)
                else:
                    cleaned[lc] = round(float(disp_dict[lc]), 4)
            except (ValueError, TypeError):
                pass
    return json.dumps(cleaned)

def format_v_label(v):
    v_str = str(v)
    return v_str if v_str.startswith("v") or v_str.startswith("FST") else f"v{v_str}"

def build_comparison_dataframe(curr_stiff, ref_stiff, ref_label="FST15", curr_label=None):
    rows = []
    curr_col = f"Displacement ({curr_label}) (mm)" if curr_label else "Displacement (mm)"
    ref_col = f"Ref. ({ref_label}) (mm)"
    for lc in LC_KEYS:
        desc = LC_INFO[lc]["desc"]
        
        c_item = curr_stiff.get(lc, 0.0)
        c_disp = c_item.get("disp", 0.0) if isinstance(c_item, dict) else float(c_item)
        
        r_item = ref_stiff.get(lc, 0.0)
        r_disp = r_item.get("disp", 0.0) if isinstance(r_item, dict) else float(r_item)
        
        diff = c_disp - r_disp
        if abs(diff) < 1e-5:
            diff = 0.0
            pct = 0.0
        else:
            pct = (diff / r_disp * 100.0) if r_disp != 0 else 0.0
            
        rows.append({
            "Load Case": lc,
            "Description": desc,
            curr_col: c_disp,
            ref_col: r_disp,
            "Δ (mm)": diff,
            "%": pct
        })
    return pd.DataFrame(rows)

def style_comparison_dataframe(df_comp):
    def highlight_delta(val):
        try:
            v = float(val)
            if v < -0.00001:
                return 'color: #33cc33; font-weight: 600;'
            elif v > 0.00001:
                return 'color: #ff4b4b; font-weight: 600;'
        except Exception:
            pass
        return ''

    delta_col = "Δ (mm)"
    pct_col = "%"

    format_dict = {}
    for c in df_comp.columns:
        if c == delta_col:
            format_dict[c] = "{:+.4f}"
        elif c == pct_col:
            format_dict[c] = "{:+.1f}%"
        elif "(mm)" in c:
            format_dict[c] = "{:.4f}"

    styled = df_comp.style.map(
        highlight_delta,
        subset=[c for c in df_comp.columns if c in [delta_col, pct_col]]
    ).format(format_dict)
    return styled

def build_mass_comparison_dataframe(eval_weight, ref_weight, curr_label="Current", ref_label="FST15"):
    try:
        e_w = float(eval_weight)
    except:
        e_w = 0.0
    try:
        r_w = float(ref_weight)
    except:
        r_w = 0.0

    diff = e_w - r_w
    if abs(diff) < 1e-4:
        diff = 0.0
        pct = 0.0
    else:
        pct = (diff / r_w * 100.0) if r_w != 0 else 0.0

    curr_col = f"Weight ({curr_label}) (g)" if curr_label else "Weight (g)"
    ref_col = f"Ref. ({ref_label}) (g)"

    return pd.DataFrame([{
        "Metric": "Total Weight",
        curr_col: e_w,
        ref_col: r_w,
        "Δ (g)": diff,
        "%": pct
    }])

def build_all_iterations_mass_dataframe(part_data, ref_weight, ref_label="FST15"):
    try:
        r_w = float(ref_weight)
    except:
        r_w = 0.0

    ref_col = f"Ref. ({ref_label}) (g)"
    rows = []
    for _, r in part_data.sort_values("Date").iterrows():
        try:
            w = float(r["Weight (g)"])
        except:
            w = 0.0
        diff = w - r_w
        if abs(diff) < 1e-4:
            diff = 0.0
            pct = 0.0
        else:
            pct = (diff / r_w * 100.0) if r_w != 0 else 0.0

        rows.append({
            "Version": format_v_label(r["Version"]),
            "Date": str(r["Date"]),
            "Weight (g)": w,
            ref_col: r_w,
            "Δ (g)": diff,
            "%": pct
        })
    return pd.DataFrame(rows)

def style_mass_dataframe(df_mass):
    def highlight_delta(val):
        try:
            v = float(val)
            if v < -0.01:
                return 'color: #33cc33; font-weight: 600;'
            elif v > 0.01:
                return 'color: #ff4b4b; font-weight: 600;'
        except Exception:
            pass
        return ''

    delta_col = "Δ (g)"
    pct_col = "%"

    format_dict = {}
    for c in df_mass.columns:
        if c == delta_col:
            format_dict[c] = "{:+.2f}"
        elif c == pct_col:
            format_dict[c] = "{:+.1f}%"
        elif "(g)" in c:
            format_dict[c] = "{:.2f}"

    styled = df_mass.style.map(
        highlight_delta,
        subset=[c for c in df_mass.columns if c in [delta_col, pct_col]]
    ).format(format_dict)
    return styled

def load_parts():
    if os.path.exists(EXCEL_PARTS):
        try:
            df = pd.read_excel(EXCEL_PARTS)
            if "Stiffness Data" not in df.columns:
                df["Stiffness Data"] = ""
            
            # Auto-seed UP_01 stiffness data if empty
            needs_save = False
            up01_mask = df["Part Name"].astype(str) == "UP_01"
            if up01_mask.any():
                for idx in df[up01_mask].index.tolist():
                    val = df.at[idx, "Stiffness Data"]
                    if pd.isna(val) or str(val).strip() == "" or str(val).strip() == "nan":
                        ver = str(df.at[idx, "Version"])
                        if "FST15" in ver.upper():
                            df.at[idx, "Stiffness Data"] = format_stiffness_data(DEFAULT_FST15_DISPLACEMENTS)
                            needs_save = True
                        else:
                            df.at[idx, "Stiffness Data"] = format_stiffness_data(DEFAULT_FST16_DISPLACEMENTS)
                            needs_save = True
                if needs_save:
                    try:
                        save_parts(df)
                    except Exception:
                        pass
            return df
        except PermissionError:
            st.error(f"PERMISSION DENIED: Please close '{EXCEL_PARTS}' in Microsoft Excel to continue.")
            st.stop()
    else:
        return pd.DataFrame(columns=[
            "Subassembly", "Part Name", "Version", "Date", 
            "Weight (g)", "Description", "Constraints", "Image Path", "Stiffness Data"
        ])

def save_parts(df):
    save_formatted_excel(df, EXCEL_PARTS, table_name="PartsTable")

def load_assemblies():
    if os.path.exists(EXCEL_ASSEMBLIES):
        try:
            return pd.read_excel(EXCEL_ASSEMBLIES)
        except PermissionError:
            st.error(f"PERMISSION DENIED: Please close '{EXCEL_ASSEMBLIES}' in Microsoft Excel to continue.")
            st.stop()
    else:
        return pd.DataFrame(columns=[
            "Assembly Version", "Date", "Total Weight (g)", 
            "Description", "Part Configuration"
        ])

def save_assemblies(df):
    save_formatted_excel(df, EXCEL_ASSEMBLIES, table_name="AssembliesTable")

def load_trash():
    if os.path.exists(EXCEL_TRASH):
        try:
            df = pd.read_excel(EXCEL_TRASH)
            if "Stiffness Data" not in df.columns:
                df["Stiffness Data"] = ""
            return df
        except PermissionError:
            st.error(f"PERMISSION DENIED: Please close '{EXCEL_TRASH}' in Microsoft Excel to continue.")
            st.stop()
    else:
        return pd.DataFrame(columns=[
            "Subassembly", "Part Name", "Version", "Date", 
            "Weight (g)", "Description", "Constraints", "Image Path", "Stiffness Data"
        ])

def save_trash(df):
    save_formatted_excel(df, EXCEL_TRASH, table_name="TrashTable")

def parse_measurements(val):
    if pd.isna(val) or str(val).strip() == "":
        return []
    lines = str(val).split("\n")
    res = []
    for l in lines:
        if ":" in l:
            k, v = l.split(":", 1)
            res.append({"Property": k.strip(), "Value": v.strip()})
        else:
            if l.strip():
                res.append({"Property": "Note", "Value": l.strip()})
    return res

def format_measurements(records):
    lines = [f"{str(r.get('Property', '')).strip()}: {str(r.get('Value', '')).strip()}" 
             for r in records if str(r.get('Property', '')).strip() != ""]
    return "\n".join(lines)

def parse_config(val):
    if pd.isna(val) or str(val).strip() == "":
        return {}
    res = {}
    for l in str(val).split("\n"):
        if ":" in l:
            k, v = l.split(":", 1)
            res[k.strip()] = v.strip().replace("v", "")
    return res

def format_config(config_dict):
    return "\n".join([f"{k}: v{v}" for k, v in config_dict.items()])


st.set_page_config(page_title="FST 16 Drivetrain Tracker", layout="wide")

# Initialize session state for navigation
if 'action' not in st.session_state:
    st.session_state.action = "create"
if 'selected_part' not in st.session_state:
    st.session_state.selected_part = None

df_parts = load_parts()
df_assemblies = load_assemblies()
df_trash = load_trash()

if 'current_page' not in st.session_state:
    st.session_state.current_page = "Part Management"

# ==================== SIDEBAR ====================
st.sidebar.markdown("""
<style>
div[data-testid="stSidebar"] button:has(p:contains('FST16 - Wheel Assembly')) {
    background-color: transparent !important;
    border: none !important;
    font-size: 28px !important;
    font-weight: 700 !important;
    padding: 0 !important;
    box-shadow: none !important;
    justify-content: flex-start !important;
    margin-bottom: 20px;
}
div[data-testid="stSidebar"] button:has(p:contains('FST16 - Wheel Assembly')):hover {
    color: #ff4b4b !important;
}
</style>
""", unsafe_allow_html=True)

if st.sidebar.button("FST16 - Wheel Assembly", use_container_width=True):
    st.session_state.current_page = "Assembly Tree"
    st.rerun()

st.sidebar.divider()

if st.sidebar.button("+ Create Part", use_container_width=True):
    st.session_state.current_page = "Part Management"
    st.session_state.action = "create"
    st.session_state.selected_part = None
    st.rerun()
    
st.sidebar.markdown("### Parts by Subassembly")
subassemblies = [s for s in df_parts["Subassembly"].unique() if s != "Other"] if not df_parts.empty else []

for sub in subassemblies:
    with st.sidebar.expander(sub, expanded=True):
        parts_in_sub = df_parts[df_parts["Subassembly"] == sub]["Part Name"].unique()
        for p in parts_in_sub:
            if st.button(p, key=f"btn_{p}", use_container_width=True):
                st.session_state.current_page = "Part Management"
                st.session_state.selected_part = p
                st.session_state.action = "view"
                st.rerun()

if not df_parts.empty and "Other" in df_parts["Subassembly"].unique():
    st.sidebar.markdown("### OEM Catalog")
    with st.sidebar.expander("Seals & Bearings", expanded=False):
        parts_in_sub = df_parts[df_parts["Subassembly"] == "Other"]["Part Name"].unique()
        for p in parts_in_sub:
            if st.button(str(p), key=f"btn_oem_{p}", use_container_width=True):
                st.session_state.current_page = "Part Management"
                st.session_state.selected_part = p
                st.session_state.action = "view"
                st.rerun()

# TRASH SECTION
st.sidebar.divider()
with st.sidebar.expander("Recycle Bin", expanded=False):
    if not df_trash.empty:
        trash_parts = df_trash["Part Name"].unique()
        for p in trash_parts:
            col1, col2 = st.columns([3, 2])
            col1.write(p)
            if col2.button("Restore", key=f"restore_{p}"):
                restored_rows = df_trash[df_trash["Part Name"] == p]
                df_parts = pd.concat([df_parts, restored_rows], ignore_index=True)
                save_parts(df_parts)
                df_trash = df_trash[df_trash["Part Name"] != p]
                save_trash(df_trash)
                st.rerun()
                
        if st.button("Empty Trash", use_container_width=True, type="primary"):
            df_trash = df_trash.iloc[0:0]
            save_trash(df_trash)
            st.rerun()
    else:
        st.caption("Trash is empty.")

# ==================== ASSEMBLY TREE ====================
if st.session_state.current_page == "Assembly Tree":
    st.title("Wheel Assembly - Drivetrain")
    st.markdown("Track full assembly versions. Each version locks specific part iterations.")
    
    selected_parts_df = pd.DataFrame()
    
    ref_parts_df = pd.DataFrame()
    
    if not df_assemblies.empty:
        col1, col2 = st.columns(2)
        with col1:
            asm_versions_ref = df_assemblies["Assembly Version"].tolist()
            default_ref_idx = asm_versions_ref.index("FST15") if "FST15" in asm_versions_ref else 0
            ref_asm_v = st.selectbox("Baseline Assembly (Reference):", asm_versions_ref, index=default_ref_idx)
        with col2:
            asm_versions_curr = df_assemblies["Assembly Version"].tolist()
            default_curr_idx = 0
            if len(asm_versions_curr) > 1:
                default_curr_idx = len(asm_versions_curr) - 1 # default to the newest one
            selected_asm_v = st.selectbox("Evaluate Assembly (Current):", asm_versions_curr, index=default_curr_idx)
            
        ref_row = df_assemblies[df_assemblies["Assembly Version"] == ref_asm_v].iloc[0]
        ref_config = parse_config(ref_row["Part Configuration"])
        for p_name, p_ver in ref_config.items():
            if p_name in ["Inner Bearing", "Outer Bearing", "Seal"]:
                actual_name = p_ver
                actual_ver = "OEM"
            else:
                actual_name = p_name
                actual_ver = p_ver
                
            row = df_parts[(df_parts["Part Name"].astype(str) == str(actual_name)) & (df_parts["Version"].astype(str) == str(actual_ver))]
            if not row.empty:
                row_copy = row.copy()
                row_copy["Part Name"] = p_name
                row_copy["Version"] = actual_name
                ref_parts_df = pd.concat([ref_parts_df, row_copy.iloc[[-1]]])
        
        asm_row = df_assemblies[df_assemblies["Assembly Version"] == selected_asm_v].iloc[0]
        st.caption(f"Created at: {asm_row['Date']}")
        if str(asm_row['Description']).strip() and str(asm_row['Description']) != "nan":
            st.info(asm_row['Description'])
            
        config = parse_config(asm_row["Part Configuration"])
        for p_name, p_ver in config.items():
            if p_name in ["Inner Bearing", "Outer Bearing", "Seal"]:
                actual_name = p_ver
                actual_ver = "OEM"
            else:
                actual_name = p_name
                actual_ver = p_ver
                
            row = df_parts[(df_parts["Part Name"].astype(str) == str(actual_name)) & (df_parts["Version"].astype(str) == str(actual_ver))]
            if not row.empty:
                row_copy = row.copy()
                row_copy["Part Name"] = p_name
                row_copy["Version"] = actual_name
                selected_parts_df = pd.concat([selected_parts_df, row_copy.iloc[[-1]]])
    else:
        st.info("No Assembly Versions saved yet.")
        if not df_parts.empty:
            for p in df_parts["Part Name"].unique():
                part_latest = df_parts[df_parts["Part Name"] == p].sort_values("Date").iloc[[-1]]
                selected_parts_df = pd.concat([selected_parts_df, part_latest])
                
        # If no assemblies, reference is just the first iteration of every part
        for p in df_parts["Part Name"].unique():
            part_first = df_parts[df_parts["Part Name"] == p].sort_values("Date").iloc[[0]]
            ref_parts_df = pd.concat([ref_parts_df, part_first])
            ref_asm_v = "First Iteration"

    if not df_parts.empty:
        dashboard_data = []
        
        all_displayed_parts = set()
        if not ref_parts_df.empty: all_displayed_parts.update(ref_parts_df["Part Name"].astype(str).tolist())
        if not selected_parts_df.empty: all_displayed_parts.update(selected_parts_df["Part Name"].astype(str).tolist())
        
        for p_name in sorted(list(all_displayed_parts)):
            if not ref_parts_df.empty and str(p_name) in ref_parts_df["Part Name"].astype(str).values:
                orig_weight = ref_parts_df[ref_parts_df["Part Name"].astype(str) == str(p_name)].iloc[0]["Weight (g)"]
            else:
                orig_weight = 0.0
                
            if not selected_parts_df.empty and str(p_name) in selected_parts_df["Part Name"].astype(str).values:
                curr_row = selected_parts_df[selected_parts_df["Part Name"].astype(str) == str(p_name)].iloc[0]
                curr_weight = curr_row["Weight (g)"]
                version = curr_row["Version"]
                subassembly = curr_row["Subassembly"]
            else:
                curr_weight = 0.0
                version = "-- Not Used --"
                # To get subassembly, try to find it in ref_parts if it exists
                if not ref_parts_df.empty and str(p_name) in ref_parts_df["Part Name"].astype(str).values:
                    subassembly = ref_parts_df[ref_parts_df["Part Name"].astype(str) == str(p_name)].iloc[0]["Subassembly"]
                else:
                    subassembly = "Other"
                
            if orig_weight == 0.0 and curr_weight == 0.0:
                continue
                
            diff = curr_weight - orig_weight
            
            in_ref = "X" if (not ref_parts_df.empty and str(p_name) in ref_parts_df["Part Name"].astype(str).values) else ""
            in_curr = "X" if (not selected_parts_df.empty and str(p_name) in selected_parts_df["Part Name"].astype(str).values) else ""
            
            dashboard_data.append({
                "Subassembly": subassembly,
                "Part Name": p_name,
                f"In {ref_asm_v}": in_ref,
                f"In {selected_asm_v}": in_curr,
                "Ref. Weight (g)": orig_weight,
                "Current Weight (g)": curr_weight,
                "Diff (g)": diff
            })
            
        dash_df = pd.DataFrame(dashboard_data, columns=[
            "Subassembly", "Part Name", 
            f"In {ref_asm_v}", f"In {selected_asm_v}", 
            "Ref. Weight (g)", "Current Weight (g)", "Diff (g)"
        ])
        
        total_curr = dash_df["Current Weight (g)"].sum() if not dash_df.empty else 0.0
        total_orig = dash_df["Ref. Weight (g)"].sum() if not dash_df.empty else 0.0
        total_diff = total_curr - total_orig
        
        col1, col2 = st.columns(2)
        col1.metric(f"Reference Weight ({ref_asm_v})", f"{total_orig:.2f} g")
        col2.metric(f"Current Weight ({selected_asm_v})", f"{total_curr:.2f} g", delta=f"{total_diff:+.2f} g", delta_color="inverse")
        
        st.divider()
        st.markdown(f"### Part Comparison vs {ref_asm_v}")
        
        # Style the dataframe for better visualization
        def highlight_diff(val):
            try:
                v = float(val)
                if v < 0:
                    return 'color: #33cc33' # Green for weight saved
                elif v > 0:
                    return 'color: #ff4b4b' # Red for weight added
            except:
                pass
            return ''
            
        styled_df = dash_df.style.map(highlight_diff, subset=['Diff (g)']).format({
            "Current Weight (g)": "{:.2f}",
            "Ref. Weight (g)": "{:.2f}",
            "Diff (g)": "{:+.2f}"
        })
        
        st.dataframe(styled_df, use_container_width=True, hide_index=True)

    st.divider()
    with st.expander("Manage / Create Assembly Versions"):
        mode = st.radio("Action", ["Create New Version", "Edit Existing Assembly"], horizontal=True)
        is_edit = mode == "Edit Existing Assembly"
        
        target_asm = None
        if is_edit:
            if df_assemblies.empty:
                st.warning("No assemblies available to edit.")
            else:
                target_asm = st.selectbox("Select Assembly to Edit:", df_assemblies["Assembly Version"].tolist())
                
        if not is_edit or (is_edit and not df_assemblies.empty and target_asm):
            with st.form("new_asm_form"):
                if is_edit:
                    target_row = df_assemblies[df_assemblies["Assembly Version"] == target_asm].iloc[0]
                    new_asm_v = st.text_input("Version ID", value=target_asm, disabled=True)
                    new_asm_desc = st.text_area("Global description of changes", value=target_row["Description"])
                    base_config = parse_config(target_row["Part Configuration"])
                else:
                    new_asm_v = st.text_input("Version ID (e.g., 1.1, 2.0)", value="1.0" if df_assemblies.empty else "")
                    new_asm_desc = st.text_area("Global description of changes")
                    base_config = {}
                    if not df_assemblies.empty:
                        base_config = parse_config(df_assemblies.iloc[-1]["Part Configuration"])
                
                st.markdown("#### Part Configuration")
                st.write("Configure the exact part versions for this assembly.")
                
                new_config = {}
                unique_parts = df_parts["Part Name"].unique() if not df_parts.empty else []
                
                bearings_list = []
                seals_list = []
                custom_parts = []
                
                if not df_parts.empty:
                    bearings_list = df_parts[df_parts["Description"].astype(str).str.contains("Schaeffler", na=False)]["Part Name"].unique().tolist()
                    seals_list = df_parts[df_parts["Description"].astype(str).str.contains("SKF", na=False)]["Part Name"].unique().tolist()
                    
                    for p in unique_parts:
                        if p not in bearings_list and p not in seals_list:
                            custom_parts.append(p)
                
                if len(custom_parts) > 0:
                    st.markdown("##### Custom Components")
                    for p in custom_parts:
                        raw_versions = df_parts[df_parts["Part Name"].astype(str) == str(p)]["Version"].tolist()
                        p_versions = ["-- Not Used --"] + raw_versions
                        
                        default_v = base_config.get(p, "-- Not Used --")
                        if default_v not in p_versions:
                            default_v = raw_versions[-1]
                            
                        idx = p_versions.index(default_v)
                        sel_v = st.selectbox(f"{p}", p_versions, index=idx)
                        if sel_v != "-- Not Used --":
                            new_config[p] = sel_v
                else:
                    st.warning("No custom parts exist yet.")
                    
                st.markdown("##### OEM Components")
                b_opts = ["-- Not Used --"] + [str(b) for b in bearings_list]
                s_opts = ["-- Not Used --"] + [str(s) for s in seals_list]
                
                existing_inner = base_config.get("Inner Bearing", "-- Not Used --")
                existing_outer = base_config.get("Outer Bearing", "-- Not Used --")
                existing_seal = base_config.get("Seal", "-- Not Used --")
                
                # Fallback for old un-migrated configs
                legacy_bearings = [k for k in base_config.keys() if k in bearings_list]
                legacy_seals = [k for k in base_config.keys() if k in seals_list]
                if existing_inner == "-- Not Used --" and len(legacy_bearings) > 0: existing_inner = legacy_bearings[0]
                if existing_outer == "-- Not Used --" and len(legacy_bearings) > 1: existing_outer = legacy_bearings[1]
                if existing_seal == "-- Not Used --" and len(legacy_seals) > 0: existing_seal = legacy_seals[0]
                
                inner = st.selectbox("Inner Bearing", b_opts, index=b_opts.index(existing_inner) if existing_inner in b_opts else 0)
                outer = st.selectbox("Outer Bearing", b_opts, index=b_opts.index(existing_outer) if existing_outer in b_opts else 0)
                seal = st.selectbox("Seal", s_opts, index=s_opts.index(existing_seal) if existing_seal in s_opts else 0)
                
                if inner != "-- Not Used --": new_config["Inner Bearing"] = inner
                if outer != "-- Not Used --": new_config["Outer Bearing"] = outer
                if seal != "-- Not Used --": new_config["Seal"] = seal
                    
                submit_label = "Save Changes to Assembly" if is_edit else "Create New Assembly"
                submitted = st.form_submit_button(submit_label)
                if submitted and new_asm_v:
                    asm_weight = 0
                    for p, v in new_config.items():
                        r = df_parts[(df_parts["Part Name"].astype(str) == str(p)) & (df_parts["Version"].astype(str) == str(v))]
                        if not r.empty:
                            asm_weight += r.iloc[-1]["Weight (g)"]
                    
                    new_row = {
                        "Assembly Version": new_asm_v,
                        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Total Weight (g)": asm_weight,
                        "Description": new_asm_desc,
                        "Part Configuration": format_config(new_config)
                    }
                    
                    if is_edit:
                        idx = df_assemblies[df_assemblies["Assembly Version"] == target_asm].index[0]
                        for k, v in new_row.items():
                            df_assemblies.at[idx, k] = v
                        st.success(f"Assembly '{target_asm}' updated!")
                    else:
                        df_assemblies = pd.concat([df_assemblies, pd.DataFrame([new_row])], ignore_index=True)
                        st.success("Assembly version created!")
                        
                    save_assemblies(df_assemblies)
                    st.rerun()

# ==================== PART MANAGEMENT ====================
elif st.session_state.current_page == "Part Management":
    
    if st.session_state.action == "create":
        st.title("Create New Part")
        
        with st.container():
            col1, col2 = st.columns(2)
            with col1:
                part_number_input = st.text_input("Part ID (e.g., type '2' to create 'UP_02')")
                subassembly = st.selectbox("Subassembly", ["UP", "HB", "GB", "Other"])
                version = st.text_input("Initial Version", value="1.0")
            with col2:
                weight = st.number_input("Initial Weight (g)", min_value=0.0, format="%.2f")
            
            st.markdown("**Fixed Constraints (Key-Value Table)**")
            st.write("Click empty cells below to add rows.")
            empty_df = pd.DataFrame([{"Property": "", "Value": ""}, {"Property": "", "Value": ""}])
            edited_meas = st.data_editor(empty_df, num_rows="dynamic", use_container_width=True)
            
            desc = st.text_area("Initial Description")
            uploaded_img, pasted_img = render_image_input("Attach Image", key="new_part_img")
            
            if st.button("Save Part", type="primary", use_container_width=True, key="btn_save_new_part"):
                version = str(version).lstrip('vV')
                if part_number_input:
                    clean_input = str(part_number_input).strip()
                    if clean_input.isdigit():
                        final_part_name = f"{subassembly}_{int(clean_input):02d}"
                    else:
                        final_part_name = clean_input

                    img_path = save_uploaded_or_pasted_image(uploaded_img, pasted_img, final_part_name, version)
                    if img_path:
                        st.session_state.pop("clipboard_new_part_img", None)
                    
                    new_stiffness = format_stiffness_data(DEFAULT_FST16_DISPLACEMENTS) if final_part_name == "UP_01" else ""
                    new_row = {
                        "Subassembly": subassembly, "Part Name": final_part_name, "Version": version,
                        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Weight (g)": weight, "Description": desc, 
                        "Constraints": format_measurements(edited_meas.to_dict('records')),
                        "Image Path": img_path,
                        "Stiffness Data": new_stiffness
                    }
                    df_parts = pd.concat([df_parts, pd.DataFrame([new_row])], ignore_index=True)
                    save_parts(df_parts)
                    st.session_state.action = "view"
                    st.session_state.selected_part = final_part_name
                    st.success("Part created!")
                    st.rerun()
                    
    elif st.session_state.action == "view" and st.session_state.selected_part:
        selected_part = st.session_state.selected_part
        
        # Check if the part actually exists in the loaded dataframe
        if df_parts.empty or selected_part not in df_parts["Part Name"].values:
            st.warning("Please wait or refresh (F5). The part might have been deleted.")
            if st.button("Back to Menu"):
                st.session_state.action = "create"
                st.session_state.selected_part = None
                st.rerun()
        else:
            part_data = df_parts[df_parts["Part Name"] == selected_part].sort_values("Date")
            latest_iter = part_data.iloc[-1]
            
            st.title(f"{selected_part} (Subassembly: {latest_iter['Subassembly']})")
            
            tab_geral, tab_nova, tab_editar = st.tabs(["Overview & History", "+ New Iteration", "Edit Iteration"])
            
            # TAB 1: OVERVIEW
            with tab_geral:
                col1, col2, col3 = st.columns(3)
                col1.metric("Current Version", f"v{latest_iter['Version']}")
                col2.metric("Current Weight", f"{latest_iter['Weight (g)']} g")
                
                weight_diff = latest_iter['Weight (g)'] - part_data.iloc[0]['Weight (g)']
                col3.metric("Net Diff (vs Original)", f"{weight_diff:+.2f} g")
                
                meas_list = parse_measurements(latest_iter['Constraints'])
                if meas_list:
                    st.markdown("**Fixed Constraints:**")
                    st.dataframe(pd.DataFrame(meas_list), hide_index=True)
                
                if len(part_data) > 1:
                    st.divider()
                    st.markdown("### Version Comparison")

                    comp_versions = []
                    for v in part_data["Version"].astype(str).tolist():
                        if v not in comp_versions:
                            comp_versions.append(v)

                    col1, col2 = st.columns(2)
                    with col1:
                        default_ref_idx = comp_versions.index("FST15") if "FST15" in comp_versions else 0
                        ref_v = st.selectbox(
                            "Baseline Iteration (Reference):",
                            comp_versions,
                            index=default_ref_idx,
                            format_func=format_v_label,
                            key=f"comp_ref_{selected_part}"
                        )
                    with col2:
                        default_eval_idx = len(comp_versions) - 1 if len(comp_versions) > 1 else 0
                        eval_v = st.selectbox(
                            "Evaluate Iteration (Current):",
                            comp_versions,
                            index=default_eval_idx,
                            format_func=format_v_label,
                            key=f"comp_eval_{selected_part}"
                        )

                    ref_rows = part_data[part_data["Version"].astype(str) == str(ref_v)]
                    ref_row = ref_rows.iloc[-1] if not ref_rows.empty else latest_iter
                    eval_rows = part_data[part_data["Version"].astype(str) == str(eval_v)]
                    eval_row = eval_rows.iloc[-1] if not eval_rows.empty else latest_iter

                    ref_label = format_v_label(ref_v)
                    eval_label = format_v_label(eval_v)

                    st.markdown("#### Mass Difference")
                    df_mass = build_mass_comparison_dataframe(
                        eval_weight=eval_row['Weight (g)'],
                        ref_weight=ref_row['Weight (g)'],
                        curr_label=eval_label,
                        ref_label=ref_label
                    )
                    st.dataframe(style_mass_dataframe(df_mass), use_container_width=True, hide_index=True)

                    if len(comp_versions) > 2:
                        with st.expander(f"All Iterations vs Baseline ({ref_label})", expanded=False):
                            df_all_mass = build_all_iterations_mass_dataframe(part_data, ref_row['Weight (g)'], ref_label)
                            st.dataframe(style_mass_dataframe(df_all_mass), use_container_width=True, hide_index=True)

                    if selected_part == "UP_01":
                        st.markdown("#### FEA Displacements")
                        ref_stiff = parse_stiffness_data(ref_row.get("Stiffness Data", ""))
                        if not ref_stiff and str(ref_v) == "FST15":
                            ref_stiff = DEFAULT_FST15_DISPLACEMENTS

                        eval_stiff = parse_stiffness_data(eval_row.get("Stiffness Data", ""))
                        if not eval_stiff:
                            if str(eval_v) == "FST15":
                                eval_stiff = DEFAULT_FST15_DISPLACEMENTS
                            else:
                                eval_stiff = DEFAULT_FST16_DISPLACEMENTS

                        df_view = build_comparison_dataframe(eval_stiff, ref_stiff, ref_label=ref_label, curr_label=eval_label)
                        st.dataframe(style_comparison_dataframe(df_view), use_container_width=True, hide_index=True)

                st.divider()
                st.subheader("History")
                for i, row in part_data.sort_values("Date", ascending=False).iterrows():
                    with st.container():
                        st.markdown(f"### v{row['Version']} <span style='font-size:16px; color:gray;'>({row['Date']})</span>", unsafe_allow_html=True)
                        c1, c2 = st.columns([2, 1])
                        with c1:
                            st.write(f"**Weight:** {row['Weight (g)']} g")
                            st.write(f"**Notes:** {row['Description']}")
                            hist_meas = parse_measurements(row['Constraints'])
                            if hist_meas:
                                st.write("**Constraints:**")
                                st.dataframe(pd.DataFrame(hist_meas), hide_index=True)
                        with c2:
                            img_file = resolve_image_path(row.get('Image Path'))
                            if img_file:
                                st.image(img_file, use_container_width=True)

                        if selected_part == "UP_01":
                            h_stiff = parse_stiffness_data(row.get("Stiffness Data", ""))
                            if h_stiff:
                                ver_str = str(row['Version'])
                                with st.expander(f"FEA Data (v{ver_str})", expanded=False):
                                    earlier_rows = part_data[part_data["Date"] < row["Date"]].sort_values("Date")
                                    if not earlier_rows.empty:
                                        p_row = earlier_rows.iloc[-1]
                                        p_stiff = parse_stiffness_data(p_row.get("Stiffness Data", ""))
                                        if not p_stiff: p_stiff = DEFAULT_FST15_DISPLACEMENTS
                                        h_ref_lbl = f"v{p_row['Version']}"
                                    else:
                                        p_stiff = DEFAULT_FST15_DISPLACEMENTS
                                        h_ref_lbl = "FST15"
                                    h_df = build_comparison_dataframe(h_stiff, p_stiff, ref_label=h_ref_lbl, curr_label=format_v_label(ver_str))
                                    st.dataframe(style_comparison_dataframe(h_df), use_container_width=True, hide_index=True)

                        st.markdown("---")

            # TAB 2: NEW ITERATION
            with tab_nova:
                if selected_part == "UP_01":
                    st.subheader(f"Create New Iteration — {selected_part}")
                    current_v = str(latest_iter['Version'])
                    try:
                        if '.' in current_v:
                            major, minor = map(int, current_v.split('.'))
                            next_minor = f"{major}.{minor + 1}"
                        else:
                            major = int(current_v)
                            next_minor = f"{major}.1"
                    except:
                        next_minor = f"{current_v}_new"

                    col1, col2 = st.columns(2)
                    with col1:
                        new_version = st.text_input("New Version ID", value=next_minor, key="up_new_ver")
                        weight = st.number_input("New Total Weight (g)", value=float(latest_iter['Weight (g)']), format="%.2f", key="up_new_wt")
                    with col2:
                        desc = st.text_area("What changed in this version?", key="up_new_desc")
                        uploaded_img, pasted_img = render_image_input("Attach Image", key="up_new_img")

                    st.markdown("**Constraints (Inherited from previous version)**")
                    meas_df_nova = pd.DataFrame(parse_measurements(latest_iter['Constraints']))
                    if meas_df_nova.empty or "Property" not in meas_df_nova.columns:
                        meas_df_nova = pd.DataFrame([{"Property": "", "Value": ""}])
                    edited_meas_nova = st.data_editor(meas_df_nova, num_rows="dynamic", use_container_width=True, key="up_new_meas")

                    st.divider()
                    st.markdown("#### FEA Displacements")

                    latest_stiff = parse_stiffness_data(latest_iter.get("Stiffness Data", ""))
                    if not latest_stiff:
                        latest_stiff = DEFAULT_FST16_DISPLACEMENTS

                    editor_rows = []
                    for lc in LC_KEYS:
                        val = latest_stiff.get(lc, LC_INFO[lc]["default_fst16"])
                        lbl = f"{lc} ({LC_INFO[lc]['desc']})" if lc in ["LC8", "LC9"] else lc
                        editor_rows.append({
                            "Load Case": lbl,
                            "Displacement (mm)": val,
                        })
                    df_stiff_input = pd.DataFrame(editor_rows)

                    edited_stiff_df = st.data_editor(
                        df_stiff_input,
                        disabled=["Load Case"],
                        column_config={
                            "Displacement (mm)": st.column_config.NumberColumn(
                                "Displacement (mm)",
                                min_value=0.0,
                                max_value=10.0,
                                step=0.0001,
                                format="%.4f"
                            )
                        },
                        use_container_width=True,
                        hide_index=True,
                        key=f"up01_stiff_editor_nova_{latest_iter['Version']}"
                    )

                    user_stiff_vals = {}
                    for idx, r in edited_stiff_df.iterrows():
                        lc = LC_KEYS[idx]
                        try:
                            d_val = float(r["Displacement (mm)"])
                        except:
                            d_val = 0.0
                        user_stiff_vals[lc] = d_val

                    if st.button("Save New Iteration", type="primary", use_container_width=True, key="save_up_new_btn"):
                        new_version = str(new_version).lstrip('vV')
                        img_path = save_uploaded_or_pasted_image(uploaded_img, pasted_img, selected_part, new_version)
                        if img_path:
                            st.session_state.pop("clipboard_up_new_img", None)

                        final_meas = [r for r in edited_meas_nova.to_dict('records') if str(r.get('Property', '')).strip() != ""]
                        new_row = {
                            "Subassembly": latest_iter['Subassembly'],
                            "Part Name": selected_part,
                            "Version": new_version,
                            "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Weight (g)": weight,
                            "Description": desc,
                            "Constraints": format_measurements(final_meas),
                            "Image Path": img_path,
                            "Stiffness Data": format_stiffness_data(user_stiff_vals)
                        }
                        df_parts = pd.concat([df_parts, pd.DataFrame([new_row])], ignore_index=True)
                        save_parts(df_parts)
                        st.session_state.pop(f"up01_stiff_editor_nova_{latest_iter['Version']}", None)
                        st.success(f"Nova iteração v{new_version} criada com sucesso!")
                        st.rerun()

                else:
                    st.subheader("Create New Iteration")
                    
                    with st.container():
                        current_v = str(latest_iter['Version'])
                        try:
                            if '.' in current_v:
                                major, minor = map(int, current_v.split('.'))
                                next_minor = f"{major}.{minor + 1}"
                            else:
                                major = int(current_v)
                                next_minor = f"{major}.1"
                        except:
                            next_minor = f"{current_v}_new"
                        
                        new_version = st.text_input("New Version ID", value=next_minor, key="non_up_new_ver")
                        weight = st.number_input("New Total Weight (g)", value=float(latest_iter['Weight (g)']), format="%.2f", key="non_up_new_wt")
                        
                        st.markdown("**Constraints (Inherited from previous version)**")
                        meas_df_nova = pd.DataFrame(parse_measurements(latest_iter['Constraints']))
                        if meas_df_nova.empty or "Property" not in meas_df_nova.columns:
                            meas_df_nova = pd.DataFrame([{"Property": "", "Value": ""}])
                        edited_meas_nova = st.data_editor(meas_df_nova, num_rows="dynamic", use_container_width=True, key="non_up_new_meas")
                        
                        desc = st.text_area("What changed in this version?", key="non_up_new_desc")
                        uploaded_img, pasted_img = render_image_input("Attach Image", key="non_up_new_img")
                        
                        if st.button("Save New Iteration", type="primary", use_container_width=True, key="save_non_up_new_btn"):
                            new_version = str(new_version).lstrip('vV')
                            img_path = save_uploaded_or_pasted_image(uploaded_img, pasted_img, selected_part, new_version)
                            if img_path:
                                st.session_state.pop("clipboard_non_up_new_img", None)
                                    
                            final_meas = [r for r in edited_meas_nova.to_dict('records') if str(r.get('Property', '')).strip() != ""]
                            
                            new_row = {
                                "Subassembly": latest_iter['Subassembly'], "Part Name": selected_part, "Version": new_version,
                                "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "Weight (g)": weight, "Description": desc, 
                                "Constraints": format_measurements(edited_meas_nova.to_dict('records')),
                                "Image Path": img_path,
                                "Stiffness Data": ""
                            }
                            df_parts = pd.concat([df_parts, pd.DataFrame([new_row])], ignore_index=True)
                            save_parts(df_parts)
                            st.success("New iteration added!")
                            st.rerun()

            # TAB 3: EDIT ITERATION
            with tab_editar:
                if selected_part == "UP_01":
                    st.subheader("Edit Iteration Data — UP_01")
                    version_indices = part_data.index.tolist()
                    def format_version_option(idx):
                        r = part_data.loc[idx]
                        return f"v{r['Version']} ({r['Date']})"

                    selected_idx = st.selectbox("Select Version to Edit:", version_indices, format_func=format_version_option, index=len(version_indices)-1, key="sel_edit_up")
                    row_to_edit = part_data.loc[selected_idx]
                    version_to_edit = row_to_edit["Version"]

                    col1, col2 = st.columns(2)
                    with col1:
                        edit_weight = st.number_input("Update Weight (g)", value=float(row_to_edit['Weight (g)']), format="%.2f", key="edit_up_wt")
                        edit_desc = st.text_area("Update Description", value=row_to_edit['Description'], key="edit_up_desc")
                    with col2:
                        uploaded_edit_img, pasted_edit_img = render_image_input("Update Image", key="edit_up_img")

                    st.markdown("**Update Constraints**")
                    meas_df_edit = pd.DataFrame(parse_measurements(row_to_edit['Constraints']))
                    if meas_df_edit.empty or "Property" not in meas_df_edit.columns:
                        meas_df_edit = pd.DataFrame([{"Property": "", "Value": ""}])
                    edited_meas_edit = st.data_editor(meas_df_edit, num_rows="dynamic", use_container_width=True, key="de_edit_up")

                    st.divider()
                    st.markdown("#### FEA Displacements")

                    curr_edit_stiff = parse_stiffness_data(row_to_edit.get("Stiffness Data", ""))
                    if not curr_edit_stiff:
                        curr_edit_stiff = DEFAULT_FST16_DISPLACEMENTS

                    edit_rows = []
                    for lc in LC_KEYS:
                        val = curr_edit_stiff.get(lc, LC_INFO[lc]["default_fst16"])
                        lbl = f"{lc} ({LC_INFO[lc]['desc']})" if lc in ["LC8", "LC9"] else lc
                        edit_rows.append({
                            "Load Case": lbl,
                            "Displacement (mm)": val,
                        })
                    df_edit_stiff_input = pd.DataFrame(edit_rows)

                    edited_stiff_edit_df = st.data_editor(
                        df_edit_stiff_input,
                        disabled=["Load Case"],
                        column_config={
                            "Displacement (mm)": st.column_config.NumberColumn(
                                "Displacement (mm)",
                                min_value=0.0,
                                max_value=10.0,
                                step=0.0001,
                                format="%.4f"
                            )
                        },
                        use_container_width=True,
                        hide_index=True,
                        key=f"stiff_editor_edit_{selected_idx}"
                    )

                    user_edit_vals = {}
                    for idx, r in edited_stiff_edit_df.iterrows():
                        lc = LC_KEYS[idx]
                        try:
                            d_val = float(r["Displacement (mm)"])
                        except:
                            d_val = 0.0
                        user_edit_vals[lc] = d_val

                    c_save, c_del_iter, c_del = st.columns([2, 2, 2])
                    with c_save:
                        save_clicked = st.button("Save Changes", type="primary", use_container_width=True, key="save_up_edit_btn")
                    with c_del_iter:
                        del_iter_clicked = st.button("Delete Iteration", use_container_width=True, key="del_up_iter_btn")
                    with c_del:
                        del_clicked = st.button("Delete Entire Part", use_container_width=True, key="del_up_part_btn")

                    if save_clicked:
                        new_img_path = save_uploaded_or_pasted_image(uploaded_edit_img, pasted_edit_img, selected_part, version_to_edit)
                        if new_img_path:
                            df_parts.at[selected_idx, 'Image Path'] = new_img_path
                            st.session_state.pop("clipboard_edit_up_img", None)

                        final_meas = [r for r in edited_meas_edit.to_dict('records') if str(r.get('Property', '')).strip() != ""]
                        df_parts.at[selected_idx, 'Weight (g)'] = edit_weight
                        df_parts.at[selected_idx, 'Description'] = edit_desc
                        df_parts.at[selected_idx, 'Constraints'] = format_measurements(final_meas)
                        df_parts.at[selected_idx, 'Stiffness Data'] = format_stiffness_data(user_edit_vals)

                        save_parts(df_parts)
                        st.success(f"Data for v{version_to_edit} updated successfully!")
                        st.rerun()

                    if del_iter_clicked:
                        if len(part_data) <= 1:
                            st.error("Cannot delete the only iteration. Use 'Delete Entire Part' instead.")
                        else:
                            df_parts = df_parts.drop(selected_idx)
                            save_parts(df_parts)
                            st.success(f"Iteration v{version_to_edit} deleted!")
                            st.rerun()

                    if del_clicked:
                        part_to_trash = df_parts[df_parts["Part Name"] == selected_part]
                        df_trash = pd.concat([df_trash, part_to_trash], ignore_index=True)
                        save_trash(df_trash)
                        df_parts = df_parts[df_parts["Part Name"] != selected_part]
                        save_parts(df_parts)
                        st.session_state.action = "create"
                        st.session_state.selected_part = None
                        st.rerun()

                else:
                    st.subheader("Edit Iteration Data")
                    
                    version_indices = part_data.index.tolist()
                    def format_version_option(idx):
                        r = part_data.loc[idx]
                        return f"v{r['Version']} ({r['Date']})"
                        
                    selected_idx = st.selectbox("Select Version to Edit:", version_indices, format_func=format_version_option, index=len(version_indices)-1)
                    row_to_edit = part_data.loc[selected_idx]
                    version_to_edit = row_to_edit["Version"]
                    
                    with st.container():
                        st.caption(f"Editing data for v{version_to_edit}. Changes will overwrite this specific iteration.")
                        edit_weight = st.number_input("Update Weight (g)", value=float(row_to_edit['Weight (g)']), format="%.2f", key="edit_non_up_wt")
                        
                        st.markdown("**Update Constraints**")
                        meas_df_edit = pd.DataFrame(parse_measurements(row_to_edit['Constraints']))
                        if meas_df_edit.empty or "Property" not in meas_df_edit.columns:
                            meas_df_edit = pd.DataFrame([{"Property": "", "Value": ""}])
                        edited_meas_edit = st.data_editor(meas_df_edit, num_rows="dynamic", use_container_width=True, key="de_edit_non_up")
                        
                        edit_desc = st.text_area("Update Description", value=row_to_edit['Description'], key="edit_non_up_desc")
                        uploaded_edit_img, pasted_edit_img = render_image_input("Update Image", key="edit_non_up_img")
                        
                        st.markdown("""
                        <style>
                        button:has(p:contains('Delete Entire Part')), button:has(p:contains('Delete Iteration')) {
                            background-color: rgba(255, 50, 50, 0.1) !important;
                            color: #ff4b4b !important;
                            border: 1px solid #ff4b4b !important;
                        }
                        button:has(p:contains('Delete Entire Part')):hover, button:has(p:contains('Delete Iteration')):hover {
                            background-color: rgba(255, 50, 50, 0.2) !important;
                            color: #ff3333 !important;
                            border: 1px solid #ff3333 !important;
                        }
                        </style>
                        """, unsafe_allow_html=True)
                        
                        c_save, c_del_iter, c_del = st.columns([2, 2, 2])
                        with c_save:
                            save_clicked = st.button("Save Changes", type="primary", use_container_width=True, key="save_non_up_edit_btn")
                        with c_del_iter:
                            del_iter_clicked = st.button("Delete Iteration", use_container_width=True, key="del_non_up_iter_btn")
                        with c_del:
                            del_clicked = st.button("Delete Entire Part", use_container_width=True, key="del_non_up_part_btn")
                        
                        if save_clicked:
                            new_img_path = save_uploaded_or_pasted_image(uploaded_edit_img, pasted_edit_img, selected_part, version_to_edit)
                            if new_img_path:
                                df_parts.at[selected_idx, 'Image Path'] = new_img_path
                                st.session_state.pop("clipboard_edit_non_up_img", None)
                                
                            df_parts.at[selected_idx, 'Weight (g)'] = edit_weight
                            df_parts.at[selected_idx, 'Description'] = edit_desc
                            df_parts.at[selected_idx, 'Constraints'] = format_measurements(edited_meas_edit.to_dict('records'))
                            
                            save_parts(df_parts)
                            st.success(f"Data for v{version_to_edit} updated successfully!")
                            st.rerun()

                        if del_iter_clicked:
                            if len(part_data) <= 1:
                                st.error("Cannot delete the only iteration. Use 'Delete Entire Part' instead.")
                            else:
                                df_parts = df_parts.drop(selected_idx)
                                save_parts(df_parts)
                                st.success(f"Iteration v{version_to_edit} deleted!")
                                st.rerun()

                        if del_clicked:
                            # Move to trash
                            part_to_trash = df_parts[df_parts["Part Name"] == selected_part]
                            df_trash = pd.concat([df_trash, part_to_trash], ignore_index=True)
                            save_trash(df_trash)
                            
                            # Remove from DB
                            df_parts = df_parts[df_parts["Part Name"] != selected_part]
                            save_parts(df_parts)
                            
                            st.session_state.action = "create"
                            st.session_state.selected_part = None
                            st.rerun()
