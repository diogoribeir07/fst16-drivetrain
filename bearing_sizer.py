"""
Bearing Sizer & Selector Module for FST 16 Wheel Assembly
Author: Drivetrain / Wheel Assembly Team (Diogo Ribeiro)
Calculates bearing kinematics, load decoupling, static safety (ISO 76),
fatigue rating life (ISO 281), camber stiffness, Motosh preload torque,
thermal fit expansion, and EHL lubrication ratio.
"""

import os
import math
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_CALC_PATH = os.path.join(BASE_DIR, "FST16_Wheel_Bearing_Calculator.xlsx")

# Master Bearing Catalog
BEARING_DATA = [{'Part_Number': 'Schaeffler 61816-Y', 'Type': 'Thin DGBB (Brass Cage)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.8, 'Static_C0r_kN': 13.8, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 6.0, 'Ball_Count_Z': 28, 'Limiting_Speed_Grease_rpm': 8000, 'Limiting_Speed_Oil_rpm': 12300, 'Radial_Stiffness_Kr_kN_mm': 500.0, 'Axial_Stiffness_Ka_kN_mm': 7.353, 'Mass_g': 155.0, 'Unit_Cost_EUR': 45.0, 'Lead_Time_Weeks': 1.0, 'Status': 'Selected Baseline'}, {'Part_Number': 'Schaeffler 61816-2RSR', 'Type': 'Thin DGBB (Rubber Sealed)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.8, 'Static_C0r_kN': 13.8, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 6.0, 'Ball_Count_Z': 28, 'Limiting_Speed_Grease_rpm': 3100, 'Limiting_Speed_Oil_rpm': 3100, 'Radial_Stiffness_Kr_kN_mm': 500.0, 'Axial_Stiffness_Ka_kN_mm': 7.353, 'Mass_g': 160.0, 'Unit_Cost_EUR': 50.0, 'Lead_Time_Weeks': 1.0, 'Status': 'Rejected Seal Drag'}, {'Part_Number': 'FAG B71816-C-TPA-P4S-UL', 'Type': 'Thin ACBB (15-deg Spindle)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 14.6, 'Static_C0r_kN': 16.6, 'Contact_Angle_deg': 15.0, 'Ball_Diam_Dw_mm': 6.747, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 12000, 'Limiting_Speed_Oil_rpm': 19000, 'Radial_Stiffness_Kr_kN_mm': 580.0, 'Axial_Stiffness_Ka_kN_mm': 75.0, 'Mass_g': 145.0, 'Unit_Cost_EUR': 370.0, 'Lead_Time_Weeks': 18.0, 'Status': 'Candidate 15-deg'}, {'Part_Number': 'FAG B71816-E-TPA-P4S-UL', 'Type': 'Thin ACBB (25-deg Light)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.7, 'Static_C0r_kN': 15.6, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 9500, 'Limiting_Speed_Oil_rpm': 15000, 'Radial_Stiffness_Kr_kN_mm': 450.0, 'Axial_Stiffness_Ka_kN_mm': 121.4, 'Mass_g': 145.0, 'Unit_Cost_EUR': 380.0, 'Lead_Time_Weeks': 18.0, 'Status': 'Target Upgrade'}, {'Part_Number': 'FAG B71816-E-TPA-P4S-UM', 'Type': 'Thin ACBB (25-deg Medium)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.7, 'Static_C0r_kN': 15.6, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 9000, 'Limiting_Speed_Oil_rpm': 14000, 'Radial_Stiffness_Kr_kN_mm': 520.0, 'Axial_Stiffness_Ka_kN_mm': 195.3, 'Mass_g': 145.0, 'Unit_Cost_EUR': 410.0, 'Lead_Time_Weeks': 18.0, 'Status': 'Candidate Medium'}, {'Part_Number': 'FAG B71816-E-TPA-P4S-UH', 'Type': 'Thin ACBB (25-deg Heavy)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.7, 'Static_C0r_kN': 15.6, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 8200, 'Limiting_Speed_Oil_rpm': 13000, 'Radial_Stiffness_Kr_kN_mm': 590.0, 'Axial_Stiffness_Ka_kN_mm': 275.3, 'Mass_g': 145.0, 'Unit_Cost_EUR': 440.0, 'Lead_Time_Weeks': 20.0, 'Status': 'Candidate Heavy'}, {'Part_Number': 'FAG HCB71816-E-TPA-P4S-UL', 'Type': 'Hybrid Ceramic ACBB (Si3N4)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 9.5, 'Static_C0r_kN': 10.8, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 12000, 'Limiting_Speed_Oil_rpm': 18000, 'Radial_Stiffness_Kr_kN_mm': 490.0, 'Axial_Stiffness_Ka_kN_mm': 135.0, 'Mass_g': 125.0, 'Unit_Cost_EUR': 620.0, 'Lead_Time_Weeks': 24.0, 'Status': 'Hybrid Ceramic'}, {'Part_Number': 'FAG XCB71816-E-TPA-P4S-UL', 'Type': 'Cronidur 30 ACBB (F1 Grade)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 12.5, 'Static_C0r_kN': 12.0, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 14000, 'Limiting_Speed_Oil_rpm': 22000, 'Radial_Stiffness_Kr_kN_mm': 550.0, 'Axial_Stiffness_Ka_kN_mm': 145.0, 'Mass_g': 125.0, 'Unit_Cost_EUR': 1250.0, 'Lead_Time_Weeks': 26.0, 'Status': 'Cronidur 30 F1'}, {'Part_Number': 'Schaeffler 61815-Y', 'Type': 'Thin DGBB (FST15 Inner)', 'Bore_d_mm': 75.0, 'OD_D_mm': 95.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 85.0, 'Dynamic_Cr_kN': 13.7, 'Static_C0r_kN': 13.4, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 6.0, 'Ball_Count_Z': 26, 'Limiting_Speed_Grease_rpm': 8500, 'Limiting_Speed_Oil_rpm': 13000, 'Radial_Stiffness_Kr_kN_mm': 480.0, 'Axial_Stiffness_Ka_kN_mm': 7.1, 'Mass_g': 140.0, 'Unit_Cost_EUR': 42.0, 'Lead_Time_Weeks': 1.0, 'Status': 'FST15 Baseline'}, {'Part_Number': 'FAG B71815-E-TPA-P4S-UL', 'Type': 'Thin ACBB (FST15 Stepped)', 'Bore_d_mm': 75.0, 'OD_D_mm': 95.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 85.0, 'Dynamic_Cr_kN': 13.2, 'Static_C0r_kN': 14.8, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 33, 'Limiting_Speed_Grease_rpm': 10000, 'Limiting_Speed_Oil_rpm': 16000, 'Radial_Stiffness_Kr_kN_mm': 430.0, 'Axial_Stiffness_Ka_kN_mm': 115.0, 'Mass_g': 135.0, 'Unit_Cost_EUR': 360.0, 'Lead_Time_Weeks': 18.0, 'Status': 'FST15 Stepped ACBB'}, {'Part_Number': 'Schaeffler 61916', 'Type': 'Medium Thin DGBB (ISO 19)', 'Bore_d_mm': 80.0, 'OD_D_mm': 110.0, 'Width_B_mm': 16.0, 'Pitch_dm_mm': 95.0, 'Dynamic_Cr_kN': 26.5, 'Static_C0r_kN': 24.0, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 10.0, 'Ball_Count_Z': 18, 'Limiting_Speed_Grease_rpm': 6700, 'Limiting_Speed_Oil_rpm': 8000, 'Radial_Stiffness_Kr_kN_mm': 620.0, 'Axial_Stiffness_Ka_kN_mm': 12.0, 'Mass_g': 380.0, 'Unit_Cost_EUR': 75.0, 'Lead_Time_Weeks': 2.0, 'Status': 'Rejected Packaging'}, {'Part_Number': 'FAG B71916-E-TPA-P4S-UL', 'Type': 'Angular Contact (ISO 19)', 'Bore_d_mm': 80.0, 'OD_D_mm': 110.0, 'Width_B_mm': 16.0, 'Pitch_dm_mm': 95.0, 'Dynamic_Cr_kN': 33.5, 'Static_C0r_kN': 24.0, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 10.0, 'Ball_Count_Z': 21, 'Limiting_Speed_Grease_rpm': 8500, 'Limiting_Speed_Oil_rpm': 13000, 'Radial_Stiffness_Kr_kN_mm': 680.0, 'Axial_Stiffness_Ka_kN_mm': 185.0, 'Mass_g': 380.0, 'Unit_Cost_EUR': 450.0, 'Lead_Time_Weeks': 20.0, 'Status': 'Rejected Packaging'}, {'Part_Number': 'Schaeffler 6016', 'Type': 'Standard DGBB (ISO 02)', 'Bore_d_mm': 80.0, 'OD_D_mm': 125.0, 'Width_B_mm': 22.0, 'Pitch_dm_mm': 102.5, 'Dynamic_Cr_kN': 51.0, 'Static_C0r_kN': 40.0, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 14.0, 'Ball_Count_Z': 12, 'Limiting_Speed_Grease_rpm': 5600, 'Limiting_Speed_Oil_rpm': 6700, 'Radial_Stiffness_Kr_kN_mm': 780.0, 'Axial_Stiffness_Ka_kN_mm': 18.0, 'Mass_g': 850.0, 'Unit_Cost_EUR': 85.0, 'Lead_Time_Weeks': 2.0, 'Status': 'Rejected Mass'}, {'Part_Number': 'FAG 7216-B-XL-TVP', 'Type': 'Standard ACBB (ISO 02)', 'Bore_d_mm': 80.0, 'OD_D_mm': 140.0, 'Width_B_mm': 26.0, 'Pitch_dm_mm': 110.0, 'Dynamic_Cr_kN': 85.0, 'Static_C0r_kN': 76.5, 'Contact_Angle_deg': 40.0, 'Ball_Diam_Dw_mm': 18.0, 'Ball_Count_Z': 14, 'Limiting_Speed_Grease_rpm': 4800, 'Limiting_Speed_Oil_rpm': 6300, 'Radial_Stiffness_Kr_kN_mm': 1100.0, 'Axial_Stiffness_Ka_kN_mm': 320.0, 'Mass_g': 1450.0, 'Unit_Cost_EUR': 140.0, 'Lead_Time_Weeks': 3.0, 'Status': 'Rejected Mass'}, {'Part_Number': 'FAG 32016-X-XL', 'Type': 'Tapered Roller Bearing', 'Bore_d_mm': 80.0, 'OD_D_mm': 125.0, 'Width_B_mm': 29.0, 'Pitch_dm_mm': 102.5, 'Dynamic_Cr_kN': 162.0, 'Static_C0r_kN': 212.0, 'Contact_Angle_deg': 14.0, 'Ball_Diam_Dw_mm': 12.0, 'Ball_Count_Z': 24, 'Limiting_Speed_Grease_rpm': 3000, 'Limiting_Speed_Oil_rpm': 4300, 'Radial_Stiffness_Kr_kN_mm': 2200.0, 'Axial_Stiffness_Ka_kN_mm': 850.0, 'Mass_g': 1250.0, 'Unit_Cost_EUR': 120.0, 'Lead_Time_Weeks': 3.0, 'Status': 'Rejected Mass Friction'}, {'Part_Number': 'FAG 3216-B-TVH', 'Type': 'Double Row Angular Contact', 'Bore_d_mm': 80.0, 'OD_D_mm': 140.0, 'Width_B_mm': 44.4, 'Pitch_dm_mm': 110.0, 'Dynamic_Cr_kN': 99.0, 'Static_C0r_kN': 102.0, 'Contact_Angle_deg': 32.0, 'Ball_Diam_Dw_mm': 16.0, 'Ball_Count_Z': 24, 'Limiting_Speed_Grease_rpm': 2800, 'Limiting_Speed_Oil_rpm': 3800, 'Radial_Stiffness_Kr_kN_mm': 1400.0, 'Axial_Stiffness_Ka_kN_mm': 420.0, 'Mass_g': 2450.0, 'Unit_Cost_EUR': 210.0, 'Lead_Time_Weeks': 4.0, 'Status': 'Rejected Road Car Unit'}, {'Part_Number': 'FAG QJ216-XL-MPA', 'Type': 'Four-Point Contact Ball', 'Bore_d_mm': 80.0, 'OD_D_mm': 140.0, 'Width_B_mm': 26.0, 'Pitch_dm_mm': 110.0, 'Dynamic_Cr_kN': 136.0, 'Static_C0r_kN': 137.0, 'Contact_Angle_deg': 35.0, 'Ball_Diam_Dw_mm': 18.0, 'Ball_Count_Z': 14, 'Limiting_Speed_Grease_rpm': 3400, 'Limiting_Speed_Oil_rpm': 5300, 'Radial_Stiffness_Kr_kN_mm': 1300.0, 'Axial_Stiffness_Ka_kN_mm': 350.0, 'Mass_g': 1700.0, 'Unit_Cost_EUR': 260.0, 'Lead_Time_Weeks': 6.0, 'Status': 'Rejected Split Ring'}]

def read_bytes_safely(file_path):
    """Safely read binary file bytes, handling Windows Excel process file sharing locks."""
    if not os.path.exists(file_path):
        return None
    try:
        with open(file_path, "rb") as f:
            return f.read()
    except (PermissionError, IOError):
        try:
            import ctypes
            from ctypes import wintypes
            GENERIC_READ = 0x80000000
            FILE_SHARE_READ = 0x00000001
            FILE_SHARE_WRITE = 0x00000002
            FILE_SHARE_DELETE = 0x00000004
            OPEN_EXISTING = 3
            FILE_ATTRIBUTE_NORMAL = 0x80
            CreateFileW = ctypes.windll.kernel32.CreateFileW
            CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
            CreateFileW.restype = wintypes.HANDLE
            ReadFile = ctypes.windll.kernel32.ReadFile
            CloseHandle = ctypes.windll.kernel32.CloseHandle

            h = CreateFileW(file_path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
            if h and h != -1 and h != wintypes.HANDLE(-1).value:
                size = os.path.getsize(file_path)
                buf = ctypes.create_string_buffer(size)
                read_bytes = wintypes.DWORD()
                ReadFile(h, buf, size, ctypes.byref(read_bytes), None)
                CloseHandle(h)
                return bytes(buf.raw[:read_bytes.value])
        except Exception:
            pass
    return None

def get_bearing_catalog_df():
    """Returns pandas DataFrame of master catalog."""
    excel_bytes = read_bytes_safely(EXCEL_CALC_PATH)
    if excel_bytes:
        import io
        try:
            df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Bearing_Database", header=2)
            if "Part_Number" in df.columns:
                return df
        except Exception:
            pass
        try:
            df_raw = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Bearing_Database", header=None)
            for r_idx in range(min(10, len(df_raw))):
                row_vals = df_raw.iloc[r_idx].astype(str).tolist()
                if "Part_Number" in row_vals:
                    df = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Bearing_Database", header=r_idx)
                    if "Part_Number" in df.columns:
                        return df
        except Exception:
            pass
    return pd.DataFrame(BEARING_DATA)

def get_bearing_dict(part_number):
    """Retrieves single bearing spec dictionary."""
    df_cat = get_bearing_catalog_df()
    if "Part_Number" in df_cat.columns:
        match = df_cat[df_cat["Part_Number"] == part_number]
        if not match.empty:
            return match.iloc[0].to_dict()
    for b in BEARING_DATA:
        if b["Part_Number"] == part_number:
            return b
    return BEARING_DATA[0]

def compute_kinematics_and_loads(outer_pname, inner_pname, fy, fz, reff, lb, v_kmh, temp_c):
    """Comprehensive analytical calculation engine with correct units and guards."""
    b_out = get_bearing_dict(outer_pname)
    b_in = get_bearing_dict(inner_pname)

    # 1. Statics & Geometry
    mx = abs(fy) * reff
    
    alpha_out_rad = math.radians(b_out["Contact_Angle_deg"])
    alpha_in_rad = math.radians(b_in["Contact_Angle_deg"])
    
    a_out = (b_out["Pitch_dm_mm"] / 2.0) * math.tan(alpha_out_rad) if b_out["Contact_Angle_deg"] > 0 else 0.0
    a_in = (b_in["Pitch_dm_mm"] / 2.0) * math.tan(alpha_in_rad) if b_in["Contact_Angle_deg"] > 0 else 0.0
    
    effective_spread_h = lb + a_out + a_in
    f_couple = mx / (effective_spread_h / 1000.0) if effective_spread_h > 0 else 0.0
    
    # Vertical wheel load distributed between outer and inner bearings
    fr_out = f_couple + (fz / 2.0)
    fr_in = abs(-f_couple + (fz / 2.0))
    fa = abs(fy)

    # 2. ISO 76 Static Equivalent Load and Safety
    if b_out["Contact_Angle_deg"] > 0:
        p0_out = max(fr_out, 0.5 * fr_out + 0.38 * fa)
    else:
        p0_out = max(fr_out, 0.6 * fr_out + 0.5 * fa)
        
    if b_in["Contact_Angle_deg"] > 0:
        p0_in = max(fr_in, 0.5 * fr_in + 0.38 * fa)
    else:
        p0_in = max(fr_in, 0.6 * fr_in + 0.5 * fa)
        
    s0_out = (b_out["Static_C0r_kN"] * 1000.0) / p0_out if p0_out > 0 else 99.0
    s0_in = (b_in["Static_C0r_kN"] * 1000.0) / p0_in if p0_in > 0 else 99.0

    # 3. ISO 281 Rating Fatigue Life across Duty Spectrum
    # Wheel speed with v=0 protection
    if reff > 0 and v_kmh > 0:
        n_wheel = ((v_kmh * 1000.0 / 3600.0) / (2.0 * math.pi * reff)) * 60.0
    else:
        n_wheel = 0.0
    
    # Spectrum: 60% cruise (1250 N), 30% 1.0g cornering (0.55*Fr), 8% peak cornering (Fr), 2% curb (2600 N)
    p_eq_out = (0.60 * (1250.0**3) + 0.30 * ((0.55 * fr_out)**3) + 0.08 * (fr_out**3) + 0.02 * (2600.0**3))**(1.0 / 3.0)
    p_eq_in = (0.60 * (1250.0**3) + 0.30 * ((0.55 * fr_in)**3) + 0.08 * (fr_in**3) + 0.02 * (2600.0**3))**(1.0 / 3.0)
    
    if n_wheel > 0 and p_eq_out > 0:
        l10h_out = (1e6 / (60.0 * n_wheel)) * ((b_out["Dynamic_Cr_kN"] * 1000.0) / p_eq_out)**3
    else:
        l10h_out = 99999.0
        
    if n_wheel > 0 and p_eq_in > 0:
        l10h_in = (1e6 / (60.0 * n_wheel)) * ((b_in["Dynamic_Cr_kN"] * 1000.0) / p_eq_in)**3
    else:
        l10h_in = 99999.0
    
    l10km_out = l10h_out * v_kmh if l10h_out < 99999.0 else 999999.0
    l10km_in = l10h_in * v_kmh if l10h_in < 99999.0 else 999999.0

    # 4. Elastokinematic Camber Stiffness & Deflection
    # Kr in kN/mm -> 1 kN/mm = 1e6 N/m
    kr_out_npm = b_out["Radial_Stiffness_Kr_kN_mm"] * 1e6  # N/m
    c_tilt_rad = 0.5 * kr_out_npm * ((effective_spread_h / 1000.0)**2)  # N*m/rad
    c_tilt_deg = c_tilt_rad * (math.pi / 180.0)  # N*m/deg
    gamma_deg = (mx / c_tilt_deg) if c_tilt_deg > 0 else 0.0
    delta_tire_top_mm = reff * math.radians(gamma_deg) * 1000.0
    
    stiffness_ratio_vs_physical = (effective_spread_h / lb)**2 if lb > 0 else 1.0

    # 5. Preload & Motosh Torque for M78x1.5
    # Manufacturer catalog preload classes: UL (84N), UM (302N), UH (712N)
    if "-UH" in b_out["Part_Number"]:
        f_preload = 712.0
    elif "-UM" in b_out["Part_Number"]:
        f_preload = 302.0
    elif b_out["Contact_Angle_deg"] > 0:
        f_preload = 84.0  # Light preload UL
    else:
        f_preload = (b_out["Axial_Stiffness_Ka_kN_mm"] * 1000.0) * 0.100  # 100 um clearance uptake = 735.3 N
        
    d2 = 77.026
    pitch = 1.50
    psi_rad = math.atan(pitch / (math.pi * d2))
    alpha_n_rad = math.radians(30.0)
    mu = 0.150
    dc = 81.5
    
    motosh_bracket = (d2 / 2000.0) * ((math.tan(psi_rad) + mu / math.cos(alpha_n_rad)) / (1.0 - mu * math.tan(psi_rad) / math.cos(alpha_n_rad))) + mu * (dc / 2000.0)
    torque_nm = f_preload * motosh_bracket

    # 6. Thermal Fit & Expansion (Al 7075-T6 vs 100Cr6)
    delta_alpha = 11.9e-6
    delta_t = temp_c - 20.0
    bore_exp_mm = 100.0 * delta_alpha * delta_t
    shaft_exp_mm = 80.0 * delta_alpha * delta_t
    net_fit_at_temp_mm = 0.030 - bore_exp_mm  # P7 interference (+0.030 mm) minus hub expansion
    
    bore_exp_140_mm = 100.0 * delta_alpha * (140.0 - 20.0)
    net_fit_140_mm = 0.030 - bore_exp_140_mm

    # 7. Kinematics (dm, ndm, nc, nb, Fc, Mg, anti-skid threshold)
    ndm = n_wheel * b_out["Pitch_dm_mm"]
    gamma_kin = (b_out["Ball_Diam_Dw_mm"] / b_out["Pitch_dm_mm"]) * math.cos(alpha_out_rad)
    nc = (n_wheel / 2.0) * (1.0 - gamma_kin)
    nb = (n_wheel / 2.0) * (b_out["Pitch_dm_mm"] / b_out["Ball_Diam_Dw_mm"]) * (1.0 - (gamma_kin**2)) if b_out["Ball_Diam_Dw_mm"] > 0 else 0.0
    
    # Ball mass
    rho_ball = 3200.0 if "HCB" in b_out["Part_Number"] else 7800.0
    ball_vol_m3 = (4.0 / 3.0) * math.pi * ((b_out["Ball_Diam_Dw_mm"] / 2000.0)**3)
    ball_mass_kg = ball_vol_m3 * rho_ball
    omega_c = 2.0 * math.pi * nc / 60.0
    omega_b = 2.0 * math.pi * nb / 60.0
    
    # Centrifugal force Fc = m * r * omega^2
    fc_ball = ball_mass_kg * (b_out["Pitch_dm_mm"] / 2000.0) * (omega_c**2)
    
    # Ball moment of inertia Jb = 0.1 * m * Dw^2
    j_ball = 0.1 * ball_mass_kg * ((b_out["Ball_Diam_Dw_mm"] / 1000.0)**2)
    # Gyroscopic moment Mg = Jb * omega_b * omega_c * sin(alpha) in N*m
    mg_ball_nm = j_ball * omega_b * omega_c * math.sin(alpha_out_rad)
    mg_ball_nmm = mg_ball_nm * 1000.0
    
    # Anti-skid minimum preload threshold:
    # Under high speed, gyroscopic slip is resisted by traction: Fa_skid = Z * (Mg / (r_ball * mu)) * sin(alpha)
    if b_out["Contact_Angle_deg"] > 0:
        r_ball_m = (b_out["Ball_Diam_Dw_mm"] / 2000.0)
        mu_traction = 0.06
        fa_skid_n = b_out["Ball_Count_Z"] * (mg_ball_nm / (r_ball_m * mu_traction)) * math.sin(alpha_out_rad)
    else:
        # For DGBB, ISO minimum radial load to prevent skidding = 0.01 * C0r
        fa_skid_n = 0.01 * (b_out["Static_C0r_kN"] * 1000.0)

    # 8. EHL Lubrication (Motul Gear 300 75W90)
    # Walther ASTM D341 kinematic viscosity at temp_c:
    nu_operating = 152.0 * math.exp(-0.038 * (temp_c - 40.0))
    nu_operating = max(nu_operating, 10.0)
    
    if n_wheel > 0:
        nu1_rated = 45000.0 * (n_wheel**(-0.83)) * (b_out["Pitch_dm_mm"]**(-0.5))
        kappa = nu_operating / nu1_rated if nu1_rated > 0 else 1.0
    else:
        nu1_rated = 1.0
        kappa = 99.0

    return {
        "b_out": b_out,
        "b_in": b_in,
        "mx": mx,
        "a_out": a_out,
        "a_in": a_in,
        "effective_spread_h": effective_spread_h,
        "f_couple": f_couple,
        "fr_out": fr_out,
        "fr_in": fr_in,
        "fa": fa,
        "p0_out": p0_out,
        "p0_in": p0_in,
        "s0_out": s0_out,
        "s0_in": s0_in,
        "n_wheel": n_wheel,
        "p_eq_out": p_eq_out,
        "p_eq_in": p_eq_in,
        "l10h_out": l10h_out,
        "l10h_in": l10h_in,
        "l10km_out": l10km_out,
        "l10km_in": l10km_in,
        "c_tilt_deg": c_tilt_deg,
        "c_tilt_rad": c_tilt_rad,
        "gamma_deg": gamma_deg,
        "delta_tire_top_mm": delta_tire_top_mm,
        "stiffness_ratio_vs_physical": stiffness_ratio_vs_physical,
        "f_preload": f_preload,
        "torque_nm": torque_nm,
        "bore_exp_mm": bore_exp_mm,
        "shaft_exp_mm": shaft_exp_mm,
        "net_fit_at_temp_mm": net_fit_at_temp_mm,
        "bore_exp_140_mm": bore_exp_140_mm,
        "net_fit_140_mm": net_fit_140_mm,
        "ndm": ndm,
        "nc": nc,
        "nb": nb,
        "fc_ball": fc_ball,
        "mg_ball_nmm": mg_ball_nmm,
        "fa_skid_n": fa_skid_n,
        "nu_operating": nu_operating,
        "nu1_rated": nu1_rated,
        "kappa": kappa
    }

def evaluate_all_load_cases(outer_pname, effective_spread_h, reff):
    """Evaluates 8 vehicle dynamic load cases."""
    b_out = get_bearing_dict(outer_pname)
    cases = [
        {"Case": "LC1 Pure Braking", "Fy": 0.0, "Fz": 3207.37, "Desc": "Maximum longitudinal deceleration"},
        {"Case": "LC2 Pure Acceleration", "Fy": 0.0, "Fz": 2409.45, "Desc": "Full throttle traction launch"},
        {"Case": "LC3 Pure Cornering", "Fy": -4436.26, "Fz": 3574.61, "Desc": "Peak 1.8g steady state cornering"},
        {"Case": "LC4 Combined BC1", "Fy": -3089.33, "Fz": 3125.00, "Desc": "Heavy braking with trail cornering"},
        {"Case": "LC5 Combined BC2", "Fy": -2468.43, "Fz": 2823.56, "Desc": "Mid braking turn entry"},
        {"Case": "LC6 Combined AC1", "Fy": -2228.85, "Fz": 2130.33, "Desc": "Corner exit acceleration phase"},
        {"Case": "LC7 Combined AC2", "Fy": -2898.75, "Fz": 2879.34, "Desc": "Power oversteer lateral cornering"},
        {"Case": "LC Bump 3.5g", "Fy": 0.0, "Fz": 5200.00, "Desc": "Dynamic curb strike event"}
    ]

    results = []
    for c in cases:
        fy = abs(c["Fy"])
        fz = c["Fz"]
        mx = fy * reff
        f_couple = mx / (effective_spread_h / 1000.0) if effective_spread_h > 0 else 0.0
        fr = f_couple + (fz / 2.0)
        
        if b_out["Contact_Angle_deg"] > 0:
            p0 = max(fr, 0.5 * fr + 0.38 * fy)
        else:
            p0 = max(fr, 0.6 * fr + 0.5 * fy)
            
        s0 = (b_out["Static_C0r_kN"] * 1000.0) / p0 if p0 > 0 else 99.0
        status = "Safe" if s0 >= 1.5 else ("Marginal" if s0 >= 1.0 else "Exceeded")
        
        results.append({
            "Load Case": c["Case"],
            "Fy (N)": fy,
            "Fz (N)": fz,
            "Mx (N*m)": round(mx, 1),
            "Fr Outer (N)": round(fr, 1),
            "P0 (N)": round(p0, 1),
            "Static S0": round(s0, 2),
            "Compliance": status
        })
    return pd.DataFrame(results)

def render_bearing_sizer_page():
    """Main Streamlit renderer for Bearing Sizer & Selector."""
    st.title("Bearing Sizer")
    st.caption("FST16 Bearing Sizer & Selector")
    
    df_catalog = get_bearing_catalog_df()
    bearing_names = df_catalog["Part_Number"].tolist()
    
    # Defaults: Outer 61816-Y, Inner 61816-Y (Dual 61816 Architecture)
    idx_out_default = bearing_names.index("Schaeffler 61816-Y") if "Schaeffler 61816-Y" in bearing_names else 0
    idx_in_default = bearing_names.index("Schaeffler 61816-Y") if "Schaeffler 61816-Y" in bearing_names else 0

    tab1, tab2, tab3, tab4 = st.tabs(["Calculator", "Load Cases", "Database", "User Manual"])

    with tab1:
        # Quick Architecture Presets
        st.subheader("Architecture Presets")
        col_p1, col_p2, col_p3, col_p4 = st.columns(4)
        if col_p1.button("Dual 61816 Baseline", use_container_width=True):
            st.session_state["bs_sel_outer"] = "Schaeffler 61816-Y"
            st.session_state["bs_sel_inner"] = "Schaeffler 61816-Y"
            st.rerun()
        if col_p2.button("Dual 71816 Target", use_container_width=True):
            st.session_state["bs_sel_outer"] = "FAG B71816-E-TPA-P4S-UL"
            st.session_state["bs_sel_inner"] = "FAG B71816-E-TPA-P4S-UL"
            st.rerun()
        if col_p3.button("Hybrid Ceramic Si3N4", use_container_width=True):
            st.session_state["bs_sel_outer"] = "FAG HCB71816-E-TPA-P4S-UL"
            st.session_state["bs_sel_inner"] = "FAG HCB71816-E-TPA-P4S-UL"
            st.rerun()
        if col_p4.button("FST15 Stepped", use_container_width=True):
            st.session_state["bs_sel_outer"] = "Schaeffler 61816-Y"
            st.session_state["bs_sel_inner"] = "Schaeffler 61815-Y"
            st.rerun()

        st.divider()

        # Architecture & Dynamic Inputs
        col_arch, col_cond = st.columns([1, 1])
        
        with col_arch:
            st.subheader("Candidate Architecture")
            sel_outer = st.selectbox("Outer Bearing", bearing_names, index=idx_out_default, key="bs_sel_outer")
            sel_inner = st.selectbox("Inner Bearing", bearing_names, index=idx_in_default, key="bs_sel_inner")
            inp_lb = st.number_input("Center Span", min_value=15.0, max_value=80.0, value=32.0, step=0.5, key="bs_inp_lb")
            
        with col_cond:
            st.subheader("Dynamic Inputs")
            col_in1, col_in2 = st.columns(2)
            with col_in1:
                # Lateral g slider linked to lateral force
                inp_lat_g = st.slider("Lateral Load", min_value=0.0, max_value=2.5, value=1.8, step=0.1, key="bs_inp_lat_g")
                # Default lateral force scaled by cornering load (1.8g -> 4436.26 N)
                default_fy = round(inp_lat_g * (4436.26 / 1.8), 1)
                inp_fy = st.number_input("Lateral Force", min_value=0.0, max_value=8000.0, value=default_fy, step=50.0, key="bs_inp_fy")
                inp_fz = st.number_input("Vertical Load", min_value=0.0, max_value=6000.0, value=1787.31, step=50.0, key="bs_inp_fz")
                inp_reff = st.number_input("Tire Radius", min_value=0.180, max_value=0.350, value=0.250, step=0.005, format="%.3f", key="bs_inp_reff")
            with col_in2:
                inp_v = st.slider("Track Speed", min_value=0.0, max_value=160.0, value=100.0, step=5.0, key="bs_inp_v")
                inp_temp = st.slider("Track Temperature", min_value=20.0, max_value=150.0, value=80.0, step=5.0, key="bs_inp_temp")

        # Analytical Calculation
        res = compute_kinematics_and_loads(
            outer_pname=sel_outer,
            inner_pname=sel_inner,
            fy=inp_fy,
            fz=inp_fz,
            reff=inp_reff,
            lb=inp_lb,
            v_kmh=inp_v,
            temp_c=inp_temp
        )

        st.divider()

        # Primary Performance Metrics Cards (Exact 8 Cards Required)
        st.subheader("Performance Summary")
        m1, m2, m3, m4 = st.columns(4)
        
        # 1. Effective Spread H
        m1.metric("Effective Spread", f"{res['effective_spread_h']:.1f} mm", delta=f"{res['stiffness_ratio_vs_physical']:.2f}x base", help="Virtual distance between load reaction centers. 25 deg contact angle projects load lines outward to 74 mm.")
        
        # 2. Moment Couple
        m2.metric("Moment Couple", f"{res['f_couple']/1000.0:.2f} kN", delta=f"Mx: {res['mx']:.1f} N*m", delta_color="inverse", help="Radial couple force = Mx / H. Expanding spread from 32 mm to 74 mm reduces this force by 58%.")
        
        # 3. Static Safety S0
        s0_val = res['s0_out']
        s0_badge = "Safe" if s0_val >= 1.5 else ("Marginal" if s0_val >= 1.0 else "Exceeded")
        m3.metric("Static Safety", f"{s0_val:.2f}", delta=s0_badge, delta_color="normal" if s0_val >= 1.5 else ("off" if s0_val >= 1.0 else "inverse"), help="ISO 76 static safety against raceway denting. Single row = 0.98; matched preloaded DB pair = 1.90 (safe).")
        
        # 4. L10h Life
        if res['l10h_out'] < 99999.0:
            m4.metric("Fatigue Life", f"{res['l10h_out']:.1f} h", delta=f"{res['l10km_out']:.0f} km", help="ISO 281 fatigue life across track duty spectrum. Full season is 500 km; survives over 7,700 km.")
        else:
            m4.metric("Fatigue Life", "Inf", delta="Static Pit", help="Vehicle stationary.")

        m5, m6, m7, m8 = st.columns(4)
        
        # 5. Camber Rigidity
        m5.metric("Camber Rigidity", f"{res['c_tilt_deg']:,.0f} N*m/deg", delta=f"gamma: {res['gamma_deg']:.3f} deg", delta_color="inverse", help="Wheel hub angular stiffness = 0.5 * Kr * H^2. Holds tire contact patch 100% flat for maximum grip.")
        
        # 6. Preload Nut Torque
        m6.metric("Nut Torque", f"{res['torque_nm']:.1f} N*m", delta=f"{res['f_preload']:.0f} N preload", help="Motosh tightening torque for M78x1.5 castle nut. 71816 needs only 1.1 N*m; 61816 needs 11.0 N*m.")
        
        # 7. Thermal Fit Change
        fit_80_um = res['net_fit_at_temp_mm'] * 1000.0
        fit_label = "Interference" if fit_80_um > 0 else "Loose Fit"
        m7.metric("Thermal Fit", f"{fit_80_um:+.1f} um", delta=fit_label, delta_color="normal" if fit_80_um > 0 else "inverse", help="At 80 C, hub bore expands 71 um, creating loose fit. Axial nut clamping prevents outer ring slip.")
        
        # 8. EHL Lubrication Kappa
        kappa_val = res['kappa']
        if kappa_val >= 2.0:
            k_badge = "Full EHL"
            k_color = "normal"
        elif kappa_val >= 1.0:
            k_badge = "Adequate"
            k_color = "off"
        else:
            k_badge = "Boundary"
            k_color = "inverse"
        m8.metric("Lubrication Ratio", f"{kappa_val:.2f}", delta=k_badge, delta_color=k_color, help="ISO 281 viscosity ratio kappa = nu / nu1. Values > 1.5 guarantee full elastohydrodynamic film separation.")

        # Operational Warnings
        if s0_val < 1.5:
            st.warning(f"Static safety factor S0 = {s0_val:.2f} is below standard threshold 1.5 under peak cornering!")
        if fit_80_um < 0:
            st.info(f"Hub bore expands by {res['bore_exp_mm']*1000.0:.1f} um at {inp_temp:.0f} C. Axial clamping by Preload Nut prevents outer ring slip.")
        if kappa_val < 1.0:
            st.warning(f"Viscosity ratio kappa = {kappa_val:.2f} indicates mixed boundary lubrication. EP additives required.")

        st.divider()

        # Interactive Comparative Visuals
        st.subheader("Comparative Visuals")
        col_c1, col_c2 = st.columns(2)

        with col_c1:
            # Camber Stiffness Comparison Bar Chart
            res_dual_61816 = compute_kinematics_and_loads("Schaeffler 61816-Y", "Schaeffler 61816-Y", inp_fy, inp_fz, inp_reff, inp_lb, inp_v, inp_temp)
            res_stepped = compute_kinematics_and_loads("Schaeffler 61816-Y", "Schaeffler 61815-Y", inp_fy, inp_fz, inp_reff, inp_lb, inp_v, inp_temp)
            res_dual_71816 = compute_kinematics_and_loads("FAG B71816-E-TPA-P4S-UL", "FAG B71816-E-TPA-P4S-UL", inp_fy, inp_fz, inp_reff, inp_lb, inp_v, inp_temp)
            
            fig_stiff = go.Figure()
            cfgs = ["Dual 61816-Y", "FST15 Stepped", "Dual 71816-E", "Current Setup"]
            stiff_vals = [res_dual_61816['c_tilt_deg'], res_stepped['c_tilt_deg'], res_dual_71816['c_tilt_deg'], res['c_tilt_deg']]
            colors = ["#4A90E2", "#7ED321", "#F5A623", "#BD10E0"]
            
            fig_stiff.add_trace(go.Bar(
                x=cfgs,
                y=stiff_vals,
                text=[f"{v:,.0f} N*m/deg" for v in stiff_vals],
                textposition="outside",
                marker_color=colors
            ))
            fig_stiff.update_layout(
                title=dict(text="Camber Rigidity Comparison (N*m/deg)", font=dict(size=14)),
                height=320,
                margin=dict(l=30, r=20, t=40, b=30),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(showgrid=True, gridcolor="#E0E0E0")
            )
            st.plotly_chart(fig_stiff, use_container_width=True)

        with col_c2:
            # Hub Bore Thermal Clearance Loss Curve (20°C to 150°C)
            temps = list(range(20, 155, 5))
            fits = [(0.030 - (100.0 * 11.9e-6 * (t - 20.0))) * 1000.0 for t in temps]
            
            fig_therm = go.Figure()
            fig_therm.add_trace(go.Scatter(
                x=temps,
                y=fits,
                mode="lines+markers",
                name="Effective Fit",
                line=dict(color="#D0021B", width=2.5)
            ))
            fig_therm.add_hline(y=0.0, line_dash="dash", line_color="#333333", annotation_text="Transition to Loose Fit")
            fig_therm.add_vline(x=inp_temp, line_dash="dot", line_color="#4A90E2", annotation_text=f"T = {inp_temp:.0f} C")
            fig_therm.update_layout(
                title=dict(text="Hub Fit vs Temperature (um)", font=dict(size=14)),
                xaxis_title="Operating Temperature (C)",
                yaxis_title="Effective Fit (um)",
                height=320,
                margin=dict(l=30, r=20, t=40, b=30),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(showgrid=True, gridcolor="#E0E0E0")
            )
            st.plotly_chart(fig_therm, use_container_width=True)

        # Deep Technical Expanders
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            with st.expander("Kinematic Details"):
                st.write(f"**Rotational Speed:** {res['n_wheel']:.1f} rpm")
                st.write(f"**Speed Parameter (n * dm):** {res['ndm']:,.0f} mm*rpm (Catalog limit: 1,000,000)")
                st.write(f"**Cage Orbital Speed (nc):** {res['nc']:.1f} rpm")
                st.write(f"**Ball Spin Speed (nb):** {res['nb']:.1f} rpm")
                st.write(f"**Ball Centrifugal Force (Fc):** {res['fc_ball']:.4f} N per ball")
                st.write(f"**Ball Gyroscopic Moment (Mg):** {res['mg_ball_nmm']:.4f} N*mm")
                st.write(f"**Anti-Skid Preload Threshold:** {res['fa_skid_n']:.1f} N axial clamping")
                st.write(f"**EHL Viscosity Ratio (kappa):** {res['kappa']:.2f} (nu: {res['nu_operating']:.1f} / nu1: {res['nu1_rated']:.1f} cSt)")

        with col_e2:
            with st.expander("Thermal Details"):
                st.write(f"**Hub Bore Expansion (Delta D @ {inp_temp:.0f}C):** {res['bore_exp_mm']*1000.0:.1f} um")
                st.write(f"**Shaft Journal Expansion (Delta d @ {inp_temp:.0f}C):** {res['shaft_exp_mm']*1000.0:.1f} um")
                st.write(f"**Net Bore Fit @ {inp_temp:.0f}C:** {res['net_fit_at_temp_mm']*1000.0:+.1f} um (Initial P7: +30.0 um)")
                st.write(f"**Net Bore Fit @ 140C (Extreme):** {res['net_fit_140_mm']*1000.0:+.1f} um")
                st.write(f"**Tire Top Deflection (Peak):** {res['delta_tire_top_mm']:.2f} mm lateral")
                st.write(f"**Required Axial Clamping:** Preload Nut with M78x1.5 locked by cotter pin")

        st.divider()

        # Excel Download Section
        excel_bytes = read_bytes_safely(EXCEL_CALC_PATH)
        if excel_bytes:
            st.download_button(
                label="Download Excel",
                data=excel_bytes,
                file_name="FST16_Wheel_Bearing_Calculator.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("Workbook not found.")

    with tab2:
        st.subheader("Load Spectrum")
        df_lc = evaluate_all_load_cases(
            outer_pname=sel_outer,
            effective_spread_h=res['effective_spread_h'],
            reff=inp_reff
        )
        
        st.dataframe(df_lc, use_container_width=True, hide_index=True)
        
        # Load Case Safety Factor Chart
        fig_lc = go.Figure()
        colors_lc = ["#2ECC71" if s >= 1.5 else ("#F39C12" if s >= 1.0 else "#E74C3C") for s in df_lc["Static S0"]]
        
        fig_lc.add_trace(go.Bar(
            x=df_lc["Load Case"],
            y=df_lc["Static S0"],
            marker_color=colors_lc,
            text=df_lc["Static S0"],
            textposition="outside"
        ))
        fig_lc.add_hline(y=1.5, line_dash="dash", line_color="#2ECC71", annotation_text="Standard Safety Threshold S0 = 1.5")
        fig_lc.add_hline(y=1.0, line_dash="dot", line_color="#E74C3C", annotation_text="Static Capacity Limit S0 = 1.0")
        fig_lc.update_layout(
            title=dict(text="Static Safety Factor across Load Cases", font=dict(size=14)),
            yaxis_title="Static Safety Factor (S0)",
            height=360,
            margin=dict(l=30, r=20, t=40, b=50),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(showgrid=True, gridcolor="#E0E0E0")
        )
        st.plotly_chart(fig_lc, use_container_width=True)

    with tab3:
        st.subheader("Database Catalog")
        
        cat_filter = st.selectbox("Filter Status", ["All", "Baseline", "Upgrade", "Option", "Rejected"], key="bs_cat_filter")
        
        if cat_filter == "All":
            df_display = df_catalog
        elif cat_filter == "Baseline":
            df_display = df_catalog[df_catalog["Status"].str.contains("Baseline", case=False, na=False)]
        elif cat_filter == "Upgrade":
            df_display = df_catalog[df_catalog["Status"].str.contains("Upgrade", case=False, na=False)]
        elif cat_filter == "Option":
            df_display = df_catalog[df_catalog["Status"].str.contains("Option", case=False, na=False)]
        else:
            df_display = df_catalog[df_catalog["Status"].str.contains("Rejected", case=False, na=False)]
            
        st.dataframe(df_display, use_container_width=True, hide_index=True)
        


    with tab4:
        st.subheader("System Overview")
        st.markdown(
            "This calculator bridges vehicle dynamics cornering loads with rolling bearing contact mechanics, "
            "evaluating the complete survivability and stiffness of the FST 16 Wheel Assembly."
        )

        with st.expander("Metric Guide", expanded=True):
            st.markdown(
                "### Understanding the 8 Live Summary Cards\n\n"
                "1. **Effective Spread (H)**:\n"
                "   - *Physical Meaning*: The virtual distance between bearing load reaction points along the wheel axis.\n"
                "   - *Background Calculation*: Deep groove ball bearings react loads vertically ($H = L_b = 32\\text{ mm}$). "
                "Angular contact bearings react loads along contact lines angled at $25^\\circ$, projecting the load center outward "
                "by apex distance $a = (d_m/2)\\tan\\alpha = 21\\text{ mm}$. The virtual stance expands to $H = 32 + 2(21) = 74\\text{ mm}$.\n"
                "   - *Why it matters*: Doubles overturning resistance and cuts radial bearing loads by $58\\%$.\n\n"
                "2. **Moment Couple (F_couple)**:\n"
                "   - *Physical Meaning*: The push-pull radial force reacting the tire overturning moment ($M_x = F_y \\cdot R_{\\text{eff}}$).\n"
                "   - *Background Calculation*: $F_{\\text{couple}} = M_x / H$. At $1.8\\text{g}$ ($M_x = 1,109\\text{ N}\\cdot\\text{m}$), "
                "61816 experiences $34.7\\text{ kN}$, while 71816 experiences only $15.0\\text{ kN}$.\n\n"
                "3. **Static Safety Factor (S0)**:\n"
                "   - *Physical Meaning*: Margin against permanent plastic indentation (brinelling) of raceways during hard curb strikes.\n"
                "   - *Background Calculation*: $S_0 = C_{0r} / P_0$ per ISO 76. Single 71816 row shows $S_0 = 0.98$ (at yield boundary), "
                "but as a matched preloaded DB pair ($C_{0r} = 31.2\\text{ kN}$), the system achieves $S_0 = 1.90$ (fully safe, $>1.5$).\n\n"
                "4. **Fatigue Rating Life (L10h / L10km)**:\n"
                "   - *Physical Meaning*: Operating hours/kilometers survived by $90\\%$ of bearings before material spalling flaking.\n"
                "   - *Background Calculation*: ISO 281 cubic equivalent load across the 4 track regimes (cruising, moderate cornering, peak 1.8g, curb bumps). "
                "Season target is $500\\text{ km}$; 71816 survives over $7,700\\text{ km}$ ($>15\\times$ safety buffer).\n\n"
                "5. **Camber Rigidity (c_tilt)**:\n"
                "   - *Physical Meaning*: Wheel hub angular resistance against tilting relative to the upright body.\n"
                "   - *Background Calculation*: $c_{\\text{tilt}} = 0.5 \\cdot K_r \\cdot H^2$. Scales with the SQUARE of effective spread ($H^2$). "
                "71816 provides $21,485\\text{ N}\\cdot\\text{m/deg}$ ($+464\\%$ stiffer than 61816), keeping the Hoosier tire contact patch 100% flat.\n\n"
                "6. **Preload Nut Torque**:\n"
                "   - *Physical Meaning*: Tightening torque required on the $M78 \\times 1.5$ Upright Lid castle nut.\n"
                "   - *Background Calculation*: Motosh thread friction equation. 71816 has factory-ground ring offsets, requiring only $1.1\\text{ N}\\cdot\\text{m}$ to clamp. "
                "61816 has $100\\text{ }\\mu\\text{m}$ internal play, requiring $11.0\\text{ N}\\cdot\\text{m}$ to compress rolling elements with $735\\text{ N}$.\n\n"
                "7. **Thermal Fit Change**:\n"
                "   - *Physical Meaning*: Differential thermal expansion between aluminum hub ($23.4 \\times 10^{-6}$) and steel bearing ($11.5 \\times 10^{-6}$).\n"
                "   - *Background Calculation*: At $80^\\circ\\text{C}$ on track, hub bore expands $71\\text{ }\\mu\\text{m}$, turning the initial $30\\text{ }\\mu\\text{m}$ P7 press fit "
                "into a $-41\\text{ }\\mu\\text{m}$ loose fit. Axial clamping by the Preload Nut sandwiches the bearing to prevent outer ring spinning.\n\n"
                "8. **EHL Lubrication Ratio (kappa)**:\n"
                "   - *Physical Meaning*: Lubricant film thickness separating rolling balls from steel raceways.\n"
                "   - *Background Calculation*: $\\kappa = \\nu / \\nu_1$. Motul Gear 300 75W90 operating viscosity at $80^\\circ\\text{C}$ is $24.8\\text{ cSt}$; "
                "rated requirement is $15.4\\text{ cSt} \\implies \\kappa = 1.69$. Values between $1.5 - 4.0$ guarantee full fluid film separation."
            )

        with st.expander("Preload Mechanics"):
            st.markdown(
                "### Why Preload Dictates Performance\n\n"
                "- **Play Elimination**: Standard bearings have $10 - 25\\text{ }\\mu\\text{m}$ radial play and $100 - 150\\text{ }\\mu\\text{m}$ axial play. "
                "Without preload, wheels wobble inside the clearance and track-temperature loose fit.\n"
                "- **Non-Linear Contact Stiffening**: Squeezing balls into raceways increases contact patch size, boosting radial stiffness by $+53\\%$.\n"
                "- **Anti-Skid Gyroscopic Protection**: At $100\\text{ km/h}$, each ball orbits at $643\\text{ rpm}$ and spins at $8,300\\text{ rpm}$, generating a "
                "gyroscopic moment ($M_g = 0.189\\text{ mN}\\cdot\\text{m}$) trying to skid the ball transversely. Preload clamping ($84\\text{ N}$ on 71816, $735\\text{ N}$ on 61816) "
                "keeps friction force far above the $35\\text{ N}$ anti-skid threshold, preventing ball scuffing."
            )

        with st.expander("Thermal Expansion"):
            st.markdown(
                "### Aluminum Hub vs Steel Bearing Fit Physics\n\n"
                "```text\n"
                "  20 deg C (Room Temp):     P7 Radial Interference fit (+30 um) [TIGHT PRESS FIT]\n"
                "  43 deg C (Warmup):        Transition line (Fit = 0 um)       [ZERO PLAY]\n"
                "  80 deg C (Track Soak):    Effective Fit = -41 um             [LOOSE CLEARANCE!]\n"
                "```\n"
                "- **The Hazard**: Without axial clamping, the outer ring spins inside the aluminum bore under torque reversals, causing fretting corrosion and bore ovalization.\n"
                "- **The FST 16 Solution**: Decouples thermal retention from radial fit by applying axial preload clamping with the $M78 \\times 1.5$ nut, locked with a cotter pin."
            )

        with st.expander("Judges Interview Tips"):
            st.markdown(
                "### Top 4 Questions Formula Student Judges Ask\n\n"
                "1. **Why didn't you use tapered roller bearings?**\n"
                "   - *Answer*: Tapered roller bearings (32016-X) weigh $1.25\\text{ kg}$ each ($2.5\\text{ kg}$ per hub), adding $+8.8\\text{ kg}$ of unsprung mass across the car. "
                "Their $\\varnothing 125\\text{ mm}$ outer diameter leaves no packaging space for our brake calipers inside 10-inch wheel rims.\n\n"
                "2. **Why do you use dual 61816 instead of 61816 outer and 61815 inner?**\n"
                "   - *Answer*: FST15 used stepped $80\\text{ mm} / 75\\text{ mm}$ bearings. Standardizing on dual $\\varnothing 80\\text{ mm}$ in FST 16 unifies the wheel assembly into a "
                "single bearing SKU, enables single-mandrel thermal installation tooling, and provides identical radial and axial stiffness across both rows.\n\n"
                "3. **How do you guarantee positive locking on the wheel bearing nut?**\n"
                "   - *Answer*: Formula Student Rule T.10 strictly prohibits threadlocker fluid alone or Nyloc nuts in hot brake zones. Our custom $M78 \\times 1.5$ nut has 6 radial castellated slots, "
                "positively locked by a $\\varnothing 1.6\\text{ mm}$ stainless steel cotter pin through the Upright Lid journal.\n\n"
                "4. **What is your bearing selection decision strategy?**\n"
                "   - *Answer*: 'If possible YES (FAG 71816-E super-precision), if not Schaeffler 61816-Y'. Both share the identical $\\varnothing 80 \\times \\varnothing 100 \\times 10\\text{ mm}$ envelope. "
                "If technical sponsorship is granted, we install 71816-E for maximum camber stiffness. If not, 61816-Y is off-the-shelf in 48 hours for under €1,000 total inventory."
            )
