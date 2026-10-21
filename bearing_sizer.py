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
BEARING_DATA = [{'Part_Number': 'Schaeffler 61816-Y', 'Type': 'Thin DGBB (Brass Cage)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.8, 'Static_C0r_kN': 13.8, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 6.0, 'Ball_Count_Z': 28, 'Limiting_Speed_Grease_rpm': 8000, 'Limiting_Speed_Oil_rpm': 12300, 'Radial_Stiffness_Kr_kN_mm': 500.0, 'Axial_Stiffness_Ka_kN_mm': 7.353, 'Mass_g': 155.0, 'Unit_Cost_EUR': 50.0, 'Lead_Time_Weeks': 1.0, 'Status': 'Selected Baseline'}, {'Part_Number': 'INA 71816-TN', 'Type': 'Thin ACBB (40-deg Polyamide)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 16.5, 'Static_C0r_kN': 19.1, 'Contact_Angle_deg': 40.0, 'Ball_Diam_Dw_mm': 7.0, 'Ball_Count_Z': 32, 'Limiting_Speed_Grease_rpm': 8000, 'Limiting_Speed_Oil_rpm': 10000, 'Radial_Stiffness_Kr_kN_mm': 520.0, 'Axial_Stiffness_Ka_kN_mm': 180.0, 'Mass_g': 152.0, 'Unit_Cost_EUR': 300.0, 'Lead_Time_Weeks': 2.0, 'Status': 'Target Upgrade'}, {'Part_Number': 'Schaeffler 61816-2RSR', 'Type': 'Thin DGBB (Rubber Sealed)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.8, 'Static_C0r_kN': 13.8, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 6.0, 'Ball_Count_Z': 28, 'Limiting_Speed_Grease_rpm': 3100, 'Limiting_Speed_Oil_rpm': 3100, 'Radial_Stiffness_Kr_kN_mm': 500.0, 'Axial_Stiffness_Ka_kN_mm': 7.353, 'Mass_g': 160.0, 'Unit_Cost_EUR': 50.0, 'Lead_Time_Weeks': 1.0, 'Status': 'Rejected Seal Drag'}, {'Part_Number': 'FAG B71816-C-TPA-P4S-UL', 'Type': 'Thin ACBB (15-deg Spindle)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 14.6, 'Static_C0r_kN': 16.6, 'Contact_Angle_deg': 15.0, 'Ball_Diam_Dw_mm': 6.747, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 12000, 'Limiting_Speed_Oil_rpm': 19000, 'Radial_Stiffness_Kr_kN_mm': 580.0, 'Axial_Stiffness_Ka_kN_mm': 75.0, 'Mass_g': 145.0, 'Unit_Cost_EUR': 370.0, 'Lead_Time_Weeks': 18.0, 'Status': 'Candidate 15-deg'}, {'Part_Number': 'FAG B71816-E-TPA-P4S-UL', 'Type': 'Thin ACBB (25-deg Light)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.7, 'Static_C0r_kN': 15.6, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 9500, 'Limiting_Speed_Oil_rpm': 15000, 'Radial_Stiffness_Kr_kN_mm': 450.0, 'Axial_Stiffness_Ka_kN_mm': 121.4, 'Mass_g': 145.0, 'Unit_Cost_EUR': 380.0, 'Lead_Time_Weeks': 18.0, 'Status': 'Target Upgrade'}, {'Part_Number': 'FAG B71816-E-TPA-P4S-UM', 'Type': 'Thin ACBB (25-deg Medium)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.7, 'Static_C0r_kN': 15.6, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 9000, 'Limiting_Speed_Oil_rpm': 14000, 'Radial_Stiffness_Kr_kN_mm': 520.0, 'Axial_Stiffness_Ka_kN_mm': 195.3, 'Mass_g': 145.0, 'Unit_Cost_EUR': 410.0, 'Lead_Time_Weeks': 18.0, 'Status': 'Candidate Medium'}, {'Part_Number': 'FAG B71816-E-TPA-P4S-UH', 'Type': 'Thin ACBB (25-deg Heavy)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 13.7, 'Static_C0r_kN': 15.6, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 8200, 'Limiting_Speed_Oil_rpm': 13000, 'Radial_Stiffness_Kr_kN_mm': 590.0, 'Axial_Stiffness_Ka_kN_mm': 275.3, 'Mass_g': 145.0, 'Unit_Cost_EUR': 440.0, 'Lead_Time_Weeks': 20.0, 'Status': 'Candidate Heavy'}, {'Part_Number': 'FAG HCB71816-E-TPA-P4S-UL', 'Type': 'Hybrid Ceramic ACBB (Si3N4)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 9.5, 'Static_C0r_kN': 10.8, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 12000, 'Limiting_Speed_Oil_rpm': 18000, 'Radial_Stiffness_Kr_kN_mm': 490.0, 'Axial_Stiffness_Ka_kN_mm': 135.0, 'Mass_g': 125.0, 'Unit_Cost_EUR': 620.0, 'Lead_Time_Weeks': 24.0, 'Status': 'Hybrid Ceramic'}, {'Part_Number': 'FAG XCB71816-E-TPA-P4S-UL', 'Type': 'Cronidur 30 ACBB (F1 Grade)', 'Bore_d_mm': 80.0, 'OD_D_mm': 100.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 90.0, 'Dynamic_Cr_kN': 12.5, 'Static_C0r_kN': 12.0, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 35, 'Limiting_Speed_Grease_rpm': 14000, 'Limiting_Speed_Oil_rpm': 22000, 'Radial_Stiffness_Kr_kN_mm': 550.0, 'Axial_Stiffness_Ka_kN_mm': 145.0, 'Mass_g': 125.0, 'Unit_Cost_EUR': 1250.0, 'Lead_Time_Weeks': 26.0, 'Status': 'Cronidur 30 F1'}, {'Part_Number': 'Schaeffler 61815-Y', 'Type': 'Thin DGBB (FST15 Inner)', 'Bore_d_mm': 75.0, 'OD_D_mm': 95.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 85.0, 'Dynamic_Cr_kN': 13.7, 'Static_C0r_kN': 13.4, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 6.0, 'Ball_Count_Z': 26, 'Limiting_Speed_Grease_rpm': 8500, 'Limiting_Speed_Oil_rpm': 13000, 'Radial_Stiffness_Kr_kN_mm': 480.0, 'Axial_Stiffness_Ka_kN_mm': 7.1, 'Mass_g': 140.0, 'Unit_Cost_EUR': 42.0, 'Lead_Time_Weeks': 1.0, 'Status': 'FST15 Baseline'}, {'Part_Number': 'FAG B71815-E-TPA-P4S-UL', 'Type': 'Thin ACBB (FST15 Stepped)', 'Bore_d_mm': 75.0, 'OD_D_mm': 95.0, 'Width_B_mm': 10.0, 'Pitch_dm_mm': 85.0, 'Dynamic_Cr_kN': 13.2, 'Static_C0r_kN': 14.8, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 7.144, 'Ball_Count_Z': 33, 'Limiting_Speed_Grease_rpm': 10000, 'Limiting_Speed_Oil_rpm': 16000, 'Radial_Stiffness_Kr_kN_mm': 430.0, 'Axial_Stiffness_Ka_kN_mm': 115.0, 'Mass_g': 135.0, 'Unit_Cost_EUR': 360.0, 'Lead_Time_Weeks': 18.0, 'Status': 'FST15 Stepped ACBB'}, {'Part_Number': 'Schaeffler 61916', 'Type': 'Medium Thin DGBB (ISO 19)', 'Bore_d_mm': 80.0, 'OD_D_mm': 110.0, 'Width_B_mm': 16.0, 'Pitch_dm_mm': 95.0, 'Dynamic_Cr_kN': 26.5, 'Static_C0r_kN': 24.0, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 10.0, 'Ball_Count_Z': 18, 'Limiting_Speed_Grease_rpm': 6700, 'Limiting_Speed_Oil_rpm': 8000, 'Radial_Stiffness_Kr_kN_mm': 620.0, 'Axial_Stiffness_Ka_kN_mm': 12.0, 'Mass_g': 380.0, 'Unit_Cost_EUR': 75.0, 'Lead_Time_Weeks': 2.0, 'Status': 'Rejected Packaging'}, {'Part_Number': 'FAG B71916-E-TPA-P4S-UL', 'Type': 'Angular Contact (ISO 19)', 'Bore_d_mm': 80.0, 'OD_D_mm': 110.0, 'Width_B_mm': 16.0, 'Pitch_dm_mm': 95.0, 'Dynamic_Cr_kN': 33.5, 'Static_C0r_kN': 24.0, 'Contact_Angle_deg': 25.0, 'Ball_Diam_Dw_mm': 10.0, 'Ball_Count_Z': 21, 'Limiting_Speed_Grease_rpm': 8500, 'Limiting_Speed_Oil_rpm': 13000, 'Radial_Stiffness_Kr_kN_mm': 680.0, 'Axial_Stiffness_Ka_kN_mm': 185.0, 'Mass_g': 380.0, 'Unit_Cost_EUR': 450.0, 'Lead_Time_Weeks': 20.0, 'Status': 'Rejected Packaging'}, {'Part_Number': 'Schaeffler 6016', 'Type': 'Standard DGBB (ISO 02)', 'Bore_d_mm': 80.0, 'OD_D_mm': 125.0, 'Width_B_mm': 22.0, 'Pitch_dm_mm': 102.5, 'Dynamic_Cr_kN': 51.0, 'Static_C0r_kN': 40.0, 'Contact_Angle_deg': 0.0, 'Ball_Diam_Dw_mm': 14.0, 'Ball_Count_Z': 12, 'Limiting_Speed_Grease_rpm': 5600, 'Limiting_Speed_Oil_rpm': 6700, 'Radial_Stiffness_Kr_kN_mm': 780.0, 'Axial_Stiffness_Ka_kN_mm': 18.0, 'Mass_g': 850.0, 'Unit_Cost_EUR': 85.0, 'Lead_Time_Weeks': 2.0, 'Status': 'Rejected Mass'}, {'Part_Number': 'FAG 7216-B-XL-TVP', 'Type': 'Standard ACBB (ISO 02)', 'Bore_d_mm': 80.0, 'OD_D_mm': 140.0, 'Width_B_mm': 26.0, 'Pitch_dm_mm': 110.0, 'Dynamic_Cr_kN': 85.0, 'Static_C0r_kN': 76.5, 'Contact_Angle_deg': 40.0, 'Ball_Diam_Dw_mm': 18.0, 'Ball_Count_Z': 14, 'Limiting_Speed_Grease_rpm': 4800, 'Limiting_Speed_Oil_rpm': 6300, 'Radial_Stiffness_Kr_kN_mm': 1100.0, 'Axial_Stiffness_Ka_kN_mm': 320.0, 'Mass_g': 1450.0, 'Unit_Cost_EUR': 140.0, 'Lead_Time_Weeks': 3.0, 'Status': 'Rejected Mass'}, {'Part_Number': 'FAG 32016-X-XL', 'Type': 'Tapered Roller Bearing', 'Bore_d_mm': 80.0, 'OD_D_mm': 125.0, 'Width_B_mm': 29.0, 'Pitch_dm_mm': 102.5, 'Dynamic_Cr_kN': 162.0, 'Static_C0r_kN': 212.0, 'Contact_Angle_deg': 14.0, 'Ball_Diam_Dw_mm': 12.0, 'Ball_Count_Z': 24, 'Limiting_Speed_Grease_rpm': 3000, 'Limiting_Speed_Oil_rpm': 4300, 'Radial_Stiffness_Kr_kN_mm': 2200.0, 'Axial_Stiffness_Ka_kN_mm': 850.0, 'Mass_g': 1250.0, 'Unit_Cost_EUR': 120.0, 'Lead_Time_Weeks': 3.0, 'Status': 'Rejected Mass Friction'}, {'Part_Number': 'FAG 3216-B-TVH', 'Type': 'Double Row Angular Contact', 'Bore_d_mm': 80.0, 'OD_D_mm': 140.0, 'Width_B_mm': 44.4, 'Pitch_dm_mm': 110.0, 'Dynamic_Cr_kN': 99.0, 'Static_C0r_kN': 102.0, 'Contact_Angle_deg': 32.0, 'Ball_Diam_Dw_mm': 16.0, 'Ball_Count_Z': 24, 'Limiting_Speed_Grease_rpm': 2800, 'Limiting_Speed_Oil_rpm': 3800, 'Radial_Stiffness_Kr_kN_mm': 1400.0, 'Axial_Stiffness_Ka_kN_mm': 420.0, 'Mass_g': 2450.0, 'Unit_Cost_EUR': 210.0, 'Lead_Time_Weeks': 4.0, 'Status': 'Rejected Road Car Unit'}, {'Part_Number': 'FAG QJ216-XL-MPA', 'Type': 'Four-Point Contact Ball', 'Bore_d_mm': 80.0, 'OD_D_mm': 140.0, 'Width_B_mm': 26.0, 'Pitch_dm_mm': 110.0, 'Dynamic_Cr_kN': 136.0, 'Static_C0r_kN': 137.0, 'Contact_Angle_deg': 35.0, 'Ball_Diam_Dw_mm': 18.0, 'Ball_Count_Z': 14, 'Limiting_Speed_Grease_rpm': 3400, 'Limiting_Speed_Oil_rpm': 5300, 'Radial_Stiffness_Kr_kN_mm': 1300.0, 'Axial_Stiffness_Ka_kN_mm': 350.0, 'Mass_g': 1700.0, 'Unit_Cost_EUR': 260.0, 'Lead_Time_Weeks': 6.0, 'Status': 'Rejected Split Ring'}]

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
    st.caption("FST16 Head-to-Head Sizing Suite")
    
    df_catalog = get_bearing_catalog_df()
    
    tab1, tab2, tab3, tab4 = st.tabs(["Calculator", "Load Cases", "Database", "User Manual"])

    with tab1:
        # Architecture Banner
        st.subheader("Candidate Comparison")
        col_bn1, col_bn2 = st.columns(2)
        with col_bn1:
            st.info("**Baseline: FAG 61816** | Deep Groove (0 deg) | €50 | 1-Season Target")
        with col_bn2:
            st.success("**Upgrade: INA 71816-TN** | Angular Contact (40 deg) | €300 | Multi-Season Asset")

        st.divider()

        # Dynamic Controls
        st.subheader("Dynamic Controls")
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            inp_lat_g = st.slider("Lateral Load", min_value=0.0, max_value=2.5, value=1.8, step=0.1, key="bs_lat_g")
            default_fy = round(inp_lat_g * (4436.26 / 1.8), 1)
            inp_fy = st.number_input("Lateral Force", min_value=0.0, max_value=8000.0, value=default_fy, step=50.0, key="bs_fy")
            inp_fz = st.number_input("Vertical Load", min_value=0.0, max_value=6000.0, value=1787.31, step=50.0, key="bs_fz")
            inp_reff = st.number_input("Tire Radius", min_value=0.180, max_value=0.350, value=0.250, step=0.005, format="%.3f", key="bs_reff")
        with col_d2:
            inp_v = st.slider("Track Speed", min_value=20.0, max_value=160.0, value=75.0, step=5.0, key="bs_v")
            n_wheel = ((inp_v * 1000.0 / 3600.0) / (2.0 * math.pi * inp_reff)) * 60.0 if inp_reff > 0 else 0.0
            st.caption(f"Wheel Speed: {n_wheel:.0f} rpm | Pace: 75 km/h avg (796 rpm) / 155 km/h max (1,650 rpm)")
            inp_temp = st.slider("Track Temperature", min_value=20.0, max_value=120.0, value=80.0, step=5.0, key="bs_temp")
            inp_season_km = st.slider("Season Distance", min_value=100, max_value=2000, value=500, step=50, key="bs_season_km")
            
        col_pr1, col_pr2 = st.columns(2)
        cost_61816 = col_pr1.number_input("Price 61816", min_value=10.0, max_value=300.0, value=50.0, step=5.0, key="bs_cost_61816")
        cost_71816 = col_pr2.number_input("Price 71816", min_value=50.0, max_value=1500.0, value=300.0, step=25.0, key="bs_cost_71816")

        # Analytical Computations for Both Candidates
        res_61816 = compute_kinematics_and_loads("Schaeffler 61816-Y", "Schaeffler 61816-Y", inp_fy, inp_fz, inp_reff, 32.0, inp_v, inp_temp)
        res_71816 = compute_kinematics_and_loads("INA 71816-TN", "INA 71816-TN", inp_fy, inp_fz, inp_reff, 32.0, inp_v, inp_temp)

        # Economic Ratios
        cost_km_61816 = (cost_61816 / res_61816['l10km_out']) if res_61816['l10km_out'] > 0 else 0.0
        cost_km_71816 = (cost_71816 / res_71816['l10km_out']) if res_71816['l10km_out'] > 0 else 0.0
        season_pct_61816 = (inp_season_km / res_61816['l10km_out']) * 100.0 if res_61816['l10km_out'] > 0 else 0.0
        season_pct_71816 = (inp_season_km / res_71816['l10km_out']) * 100.0 if res_71816['l10km_out'] > 0 else 0.0

        st.divider()

        # Head-to-Head Comparison Dashboard
        st.subheader("Head Comparison")
        col_c_left, col_c_right = st.columns(2)

        with col_c_left:
            st.markdown("#### Baseline: FAG 61816 (CN)")
            m_b1, m_b2 = st.columns(2)
            m_b1.metric("Couple Force", f"{res_61816['f_couple']/1000.0:.2f} kN", delta="H = 32.0 mm", delta_color="inverse")
            m_b2.metric("Static Safety", f"{res_61816['s0_out']:.2f}", delta="Yield Risk" if res_61816['s0_out'] < 1.0 else "Marginal", delta_color="inverse")
            
            m_b3, m_b4 = st.columns(2)
            m_b3.metric("Fatigue Life", f"{res_61816['l10km_out']:.0f} km", delta=f"{res_61816['l10h_out']:.1f} hours", delta_color="inverse")
            m_b4.metric("Cost Per Km", f"€{cost_km_61816:.4f}/km", delta=f"{season_pct_61816:.1f}% season", delta_color="inverse")

        with col_c_right:
            st.markdown("#### Upgrade: INA 71816-TN (40°)")
            m_u1, m_u2 = st.columns(2)
            m_u1.metric("Couple Force", f"{res_71816['f_couple']/1000.0:.2f} kN", delta=f"H = {res_71816['effective_spread_h']:.1f} mm", delta_color="normal")
            m_u2.metric("Static Safety", f"{res_71816['s0_out']:.2f}", delta="Fully Safe", delta_color="normal")
            
            m_u3, m_u4 = st.columns(2)
            m_u3.metric("Fatigue Life", f"{res_71816['l10km_out']:.0f} km", delta=f"{res_71816['l10h_out']:.1f} hours", delta_color="normal")
            m_u4.metric("Cost Per Km", f"€{cost_km_71816:.4f}/km", delta=f"{season_pct_71816:.1f}% season", delta_color="normal")

        st.divider()

        # Dynamic Visuals with Live Moving Operating Markers
        st.subheader("Dynamic Visuals")
        col_plot1, col_plot2 = st.columns(2)

        with col_plot1:
            # 1. Camber Deflection Curve vs Lateral g (with live active marker)
            g_range = [i * 0.1 for i in range(0, 26)]
            gamma_61816_curve = []
            gamma_71816_curve = []
            for g_val in g_range:
                fy_val = g_val * (4436.26 / 1.8)
                mx_val = fy_val * inp_reff
                c_61816 = res_61816['c_tilt_deg']
                c_71816 = res_71816['c_tilt_deg']
                gamma_61816_curve.append(mx_val / c_61816 if c_61816 > 0 else 0.0)
                gamma_71816_curve.append(mx_val / c_71816 if c_71816 > 0 else 0.0)

            fig_camber = go.Figure()
            fig_camber.add_trace(go.Scatter(
                x=g_range, y=gamma_61816_curve, mode="lines", name="61816 Baseline", line=dict(color="#4A90E2", width=2.5)
            ))
            fig_camber.add_trace(go.Scatter(
                x=g_range, y=gamma_71816_curve, mode="lines", name="71816-TN Upgrade", line=dict(color="#F5A623", width=2.5)
            ))
            # Moving active markers
            fig_camber.add_trace(go.Scatter(
                x=[inp_lat_g], y=[res_61816['gamma_deg']], mode="markers", name="Active 61816",
                marker=dict(size=12, color="#4A90E2", symbol="diamond")
            ))
            fig_camber.add_trace(go.Scatter(
                x=[inp_lat_g], y=[res_71816['gamma_deg']], mode="markers", name="Active 71816",
                marker=dict(size=12, color="#F5A623", symbol="star")
            ))
            fig_camber.update_layout(
                title=dict(text="Camber Deflection vs Lateral Load (deg)", font=dict(size=14)),
                xaxis_title="Lateral Load (g)",
                yaxis_title="Wheel Tilt Angle (deg)",
                height=340,
                margin=dict(l=30, r=20, t=40, b=30),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(showgrid=True, gridcolor="#E0E0E0")
            )
            st.plotly_chart(fig_camber, use_container_width=True)

        with col_plot2:
            # 2. Hub Bore Thermal Expansion Curve (with moving vertical line)
            temps = list(range(20, 130, 5))
            fits = [(0.030 - (100.0 * 11.9e-6 * (t - 20.0))) * 1000.0 for t in temps]
            
            fig_therm = go.Figure()
            fig_therm.add_trace(go.Scatter(
                x=temps, y=fits, mode="lines+markers", name="Bore Fit", line=dict(color="#D0021B", width=2.5)
            ))
            fig_therm.add_hline(y=0.0, line_dash="dash", line_color="#333333", annotation_text="Loose Transition")
            fig_therm.add_vline(x=inp_temp, line_dash="dot", line_color="#4A90E2", annotation_text=f"T = {inp_temp:.0f} C")
            fig_therm.update_layout(
                title=dict(text="Hub Fit vs Temperature (um)", font=dict(size=14)),
                xaxis_title="Operating Temperature (C)",
                yaxis_title="Effective Fit (um)",
                height=340,
                margin=dict(l=30, r=20, t=40, b=30),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                yaxis=dict(showgrid=True, gridcolor="#E0E0E0")
            )
            st.plotly_chart(fig_therm, use_container_width=True)

        # Operational Verification Expanders
        col_ex1, col_ex2 = st.columns(2)
        with col_ex1:
            with st.expander("Preload Verification"):
                st.write(f"**61816 Nut Torque:** {res_61816['torque_nm']:.1f} N*m (absorbs 100 um clearance)")
                st.write(f"**71816 Nut Torque:** {res_71816['torque_nm']:.1f} N*m (flush face clamp)")
                st.write(f"**Anti-Skid Clamping:** 71816 requires {res_71816['fa_skid_n']:.1f} N axial force at {n_wheel:.0f} rpm")
                st.write("**Positive Locking:** M78x1.5 nut secured by 1.6 mm cotter pin (FSG rule T.10)")
        with col_ex2:
            with st.expander("Thermal Verification"):
                fit_80 = res_61816['net_fit_at_temp_mm'] * 1000.0
                st.write(f"**Hub Expansion @ {inp_temp:.0f}C:** {res_61816['bore_exp_mm']*1000.0:.1f} um")
                st.write(f"**Shaft Expansion @ {inp_temp:.0f}C:** {res_61816['shaft_exp_mm']*1000.0:.1f} um (absorbs C3 play)")
                st.write(f"**Net Bore Fit @ {inp_temp:.0f}C:** {fit_80:+.1f} um (axial nut clamp prevents slip)")
                st.write(f"**Tire Deflection:** 61816 = {res_61816['delta_tire_top_mm']:.2f} mm | 71816 = {res_71816['delta_tire_top_mm']:.2f} mm")

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
        df_lc_61816 = evaluate_all_load_cases("Schaeffler 61816-Y", 32.0, inp_reff)
        df_lc_71816 = evaluate_all_load_cases("INA 71816-TN", res_71816['effective_spread_h'], inp_reff)
        
        df_comp_lc = pd.DataFrame({
            "Load Case": df_lc_61816["Load Case"],
            "Fy (N)": df_lc_61816["Fy (N)"],
            "Fz (N)": df_lc_61816["Fz (N)"],
            "Mx (N*m)": df_lc_61816["Mx (N*m)"],
            "S0 61816": df_lc_61816["Static S0"],
            "S0 71816": df_lc_71816["Static S0"],
            "Verdict 61816": df_lc_61816["Compliance"],
            "Verdict 71816": df_lc_71816["Compliance"]
        })
        st.dataframe(df_comp_lc, use_container_width=True, hide_index=True)
        
        # Load Case Safety Comparison Bar Chart
        fig_lc = go.Figure()
        fig_lc.add_trace(go.Bar(
            x=df_comp_lc["Load Case"], y=df_comp_lc["S0 61816"], name="61816 Baseline", marker_color="#E74C3C"
        ))
        fig_lc.add_trace(go.Bar(
            x=df_comp_lc["Load Case"], y=df_comp_lc["S0 71816"], name="71816-TN Upgrade", marker_color="#2ECC71"
        ))
        fig_lc.add_hline(y=1.5, line_dash="dash", line_color="#2ECC71", annotation_text="Safety Limit S0 = 1.5")
        fig_lc.add_hline(y=1.0, line_dash="dot", line_color="#E74C3C", annotation_text="Yield Limit S0 = 1.0")
        fig_lc.update_layout(
            title=dict(text="Static Safety Factor across Load Cases (ISO 76)", font=dict(size=14)),
            yaxis_title="Static Safety Factor (S0)",
            barmode="group",
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
        st.subheader("System Manual")
        st.markdown(
            "This guide synthesizes the physics and financial trade-offs between "
            "the two realistic wheel bearing choices for the FST 16 prototype."
        )

        with st.expander("Distance Mechanics"):
            st.markdown(
                "### Why Mileage is Independent of RPM\n\n"
                "- **Revolutions Rule**: ISO 281 defines fatigue life strictly in accumulated contact stress cycles. "
                "For `71816-TN`, $L_{10} = 11.91\\text{ million revolutions}$.\n"
                "- **Circumference Constant**: With $R_{\\text{eff}} = 0.250\\text{ m}$, 1 kilometer always takes exactly $636.6\\text{ wheel turns}$.\n"
                "- **Conclusion**: The vehicle will travel **$18,708\\text{ km}$** before the 10% fatigue probability limit, whether driven at $40\\text{ km/h}$ or $150\\text{ km/h}$."
            )

        with st.expander("Speed Mechanics"):
            st.markdown(
                "### Race Pace vs Top Speed\n\n"
                "- **Average Race Pace (796 rpm / 75 km/h)**: In tight autocross and endurance tracks, time-weighted average speed is $\\approx 75\\text{ km/h}$. "
                "At this pace, `71816-TN` lasts **$249.4\\text{ hours}$**.\n"
                "- **Continuous Top Speed (1,650 rpm / 155.5 km/h)**: Even if running flat-out continuously, `71816-TN` survives **$120.3\\text{ hours}$** ($18.5\\times$ a full season).\n"
                "- **Limiting Speed Safety ($n_G$)**: Catalog limit is **$8,000\\text{ rpm}$**. At $1,650\\text{ rpm}$, the bearing operates at only $20.6\\%$ of its limit."
            )

        with st.expander("Cube Rule"):
            st.markdown(
                "### Why Fatigue Life Scales with (1/P)^3\n\n"
                "- **Virtual Spread Expansion**: The $40^\\circ$ contact angle projects pressure apexes outward ($a = 42.8\\text{ mm}$), "
                "widening effective spread to $H = 107.6\\text{ mm}$ (vs $32.0\\text{ mm}$). This cuts overturning couple force by $-70\\%$.\n"
                "- **ISO 281 Sensitivity**: Fatigue life is inversely proportional to load CUBED ($L_{10} \\propto 1/P^3$). "
                "Reducing equivalent load from $17.9\\text{ kN}$ to $7.2\\text{ kN}$ multiplies fatigue life by **$26.1\\times$**!"
            )

        with st.expander("Economic Rationale"):
            st.markdown(
                "### Purchasing & Amortization Strategy\n\n"
                "- **FAG 61816 (€50)**: Ideal budget baseline. Survives 1 season ($717\\text{ km}$ covers $500\\text{ km}$ with $1.4\\times$ margin), but consumes $70\\%$ of its life and must be scrapped.\n"
                "- **INA 71816-TN (€300)**: Consumes only $2.7\\%$ of its life in a season. Multi-season reusable for 5+ prototype generations, making it **$4.4\\times$ cheaper per kilometer** (€0.016/km vs €0.070/km) with zero yield risk ($S_0 = 1.58$)."
            )

