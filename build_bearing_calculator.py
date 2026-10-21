"""
Script to generate FST16_Wheel_Bearing_Calculator.xlsx
Workbook includes:
- Sheet 1: Bearing_Database
- Sheet 2: Bearing_Calculator
Fully compliant with openpyxl, ISO 76, ISO 281, and Drivetrain_Tracker rules.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
import os

def create_wheel_bearing_calculator(output_path):
    wb = openpyxl.Workbook()
    
    # ----------------------------------------------------
    # STYLES & PALETTE (Clean FST Engineering Theme)
    # ----------------------------------------------------
    font_family = "Segoe UI"
    
    title_font = Font(name=font_family, size=14, bold=True, color="1F4E79")
    section_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    header_font = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    bold_font = Font(name=font_family, size=10, bold=True)
    normal_font = Font(name=font_family, size=10)
    note_font = Font(name=font_family, size=9, italic=True, color="595959")
    
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    section_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    accent_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    input_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Soft yellow for user inputs
    result_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Soft green for primary outputs
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    double_bottom_side = Side(border_style="double", color="1F4E79")
    thick_bottom_side = Side(border_style="medium", color="1F4E79")
    
    border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    border_result = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=double_bottom_side)
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    # ====================================================
    # SHEET 1: Bearing_Database
    # ====================================================
    ws_db = wb.active
    ws_db.title = "Bearing_Database"
    ws_db.views.sheetView[0].showGridLines = True
    
    ws_db.cell(row=1, column=1, value="FST 16 Wheel Bearing Catalog & Technical Database").font = title_font
    ws_db.row_dimensions[1].height = 28
    
    db_headers = [
        "Part_Number", "Type", "Bore_d_mm", "OD_D_mm", "Width_B_mm", "Pitch_dm_mm",
        "Dynamic_Cr_kN", "Static_C0r_kN", "Contact_Angle_deg", "Ball_Diam_Dw_mm", "Ball_Count_Z",
        "Limiting_Speed_Grease_rpm", "Limiting_Speed_Oil_rpm", "Radial_Stiffness_Kr_kN_mm",
        "Axial_Stiffness_Ka_kN_mm", "Mass_g", "Unit_Cost_EUR", "Lead_Time_Weeks", "Status"
    ]
    
    ws_db.row_dimensions[3].height = 24
    for col_idx, h in enumerate(db_headers, 1):
        cell = ws_db.cell(row=3, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = border_all
        
    db_rows = [['Schaeffler 61816-Y', 'Thin DGBB (Brass Cage)', 80.0, 100.0, 10.0, 90.0, 13.8, 13.8, 0.0, 6.0, 28, 8000, 12300, 500.0, 7.353, 155.0, 50.0, 1.0, 'Selected Baseline'], ['INA 71816-TN', 'Thin ACBB (40-deg Polyamide)', 80.0, 100.0, 10.0, 90.0, 16.5, 19.1, 40.0, 7.0, 32, 8000, 10000, 520.0, 180.0, 152.0, 300.0, 2.0, 'Target Upgrade'], ['Schaeffler 61816-2RSR', 'Thin DGBB (Rubber Sealed)', 80.0, 100.0, 10.0, 90.0, 13.8, 13.8, 0.0, 6.0, 28, 3100, 3100, 500.0, 7.353, 160.0, 50.0, 1.0, 'Rejected Seal Drag'], ['FAG B71816-C-TPA-P4S-UL', 'Thin ACBB (15-deg Spindle)', 80.0, 100.0, 10.0, 90.0, 14.6, 16.6, 15.0, 6.747, 35, 12000, 19000, 580.0, 75.0, 145.0, 370.0, 18.0, 'Candidate 15-deg'], ['FAG B71816-E-TPA-P4S-UL', 'Thin ACBB (25-deg Light)', 80.0, 100.0, 10.0, 90.0, 13.7, 15.6, 25.0, 7.144, 35, 9500, 15000, 450.0, 121.4, 145.0, 380.0, 18.0, 'Target Upgrade'], ['FAG B71816-E-TPA-P4S-UM', 'Thin ACBB (25-deg Medium)', 80.0, 100.0, 10.0, 90.0, 13.7, 15.6, 25.0, 7.144, 35, 9000, 14000, 520.0, 195.3, 145.0, 410.0, 18.0, 'Candidate Medium'], ['FAG B71816-E-TPA-P4S-UH', 'Thin ACBB (25-deg Heavy)', 80.0, 100.0, 10.0, 90.0, 13.7, 15.6, 25.0, 7.144, 35, 8200, 13000, 590.0, 275.3, 145.0, 440.0, 20.0, 'Candidate Heavy'], ['FAG HCB71816-E-TPA-P4S-UL', 'Hybrid Ceramic ACBB (Si3N4)', 80.0, 100.0, 10.0, 90.0, 9.5, 10.8, 25.0, 7.144, 35, 12000, 18000, 490.0, 135.0, 125.0, 620.0, 24.0, 'Hybrid Ceramic'], ['FAG XCB71816-E-TPA-P4S-UL', 'Cronidur 30 ACBB (F1 Grade)', 80.0, 100.0, 10.0, 90.0, 12.5, 12.0, 25.0, 7.144, 35, 14000, 22000, 550.0, 145.0, 125.0, 1250.0, 26.0, 'Cronidur 30 F1'], ['Schaeffler 61815-Y', 'Thin DGBB (FST15 Inner)', 75.0, 95.0, 10.0, 85.0, 13.7, 13.4, 0.0, 6.0, 26, 8500, 13000, 480.0, 7.1, 140.0, 42.0, 1.0, 'FST15 Baseline'], ['FAG B71815-E-TPA-P4S-UL', 'Thin ACBB (FST15 Stepped)', 75.0, 95.0, 10.0, 85.0, 13.2, 14.8, 25.0, 7.144, 33, 10000, 16000, 430.0, 115.0, 135.0, 360.0, 18.0, 'FST15 Stepped ACBB'], ['Schaeffler 61916', 'Medium Thin DGBB (ISO 19)', 80.0, 110.0, 16.0, 95.0, 26.5, 24.0, 0.0, 10.0, 18, 6700, 8000, 620.0, 12.0, 380.0, 75.0, 2.0, 'Rejected Packaging'], ['FAG B71916-E-TPA-P4S-UL', 'Angular Contact (ISO 19)', 80.0, 110.0, 16.0, 95.0, 33.5, 24.0, 25.0, 10.0, 21, 8500, 13000, 680.0, 185.0, 380.0, 450.0, 20.0, 'Rejected Packaging'], ['Schaeffler 6016', 'Standard DGBB (ISO 02)', 80.0, 125.0, 22.0, 102.5, 51.0, 40.0, 0.0, 14.0, 12, 5600, 6700, 780.0, 18.0, 850.0, 85.0, 2.0, 'Rejected Mass'], ['FAG 7216-B-XL-TVP', 'Standard ACBB (ISO 02)', 80.0, 140.0, 26.0, 110.0, 85.0, 76.5, 40.0, 18.0, 14, 4800, 6300, 1100.0, 320.0, 1450.0, 140.0, 3.0, 'Rejected Mass'], ['FAG 32016-X-XL', 'Tapered Roller Bearing', 80.0, 125.0, 29.0, 102.5, 162.0, 212.0, 14.0, 12.0, 24, 3000, 4300, 2200.0, 850.0, 1250.0, 120.0, 3.0, 'Rejected Mass Friction'], ['FAG 3216-B-TVH', 'Double Row Angular Contact', 80.0, 140.0, 44.4, 110.0, 99.0, 102.0, 32.0, 16.0, 24, 2800, 3800, 1400.0, 420.0, 2450.0, 210.0, 4.0, 'Rejected Road Car Unit'], ['FAG QJ216-XL-MPA', 'Four-Point Contact Ball', 80.0, 140.0, 26.0, 110.0, 136.0, 137.0, 35.0, 18.0, 14, 3400, 5300, 1300.0, 350.0, 1700.0, 260.0, 6.0, 'Rejected Split Ring']]
    
    for r_idx, row_data in enumerate(db_rows, 4):
        ws_db.row_dimensions[r_idx].height = 20
        for c_idx, val in enumerate(row_data, 1):
            cell = ws_db.cell(row=r_idx, column=c_idx, value=val)
            cell.font = normal_font
            cell.border = border_all
            if c_idx in [1, 2, 19]:
                cell.alignment = align_left
            elif c_idx in [3, 4, 5, 11, 12, 13]:
                cell.alignment = align_center
                if c_idx in [11, 12, 13]:
                    cell.number_format = "#,##0"
                else:
                    cell.number_format = "0.0"
            else:
                cell.alignment = align_right
                if c_idx in [7, 8, 14, 15]:
                    cell.number_format = "0.00"
                elif c_idx in [16, 17]:
                    cell.number_format = "#,##0.0"
                else:
                    cell.number_format = "0.0"
                    
    for col in ws_db.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws_db.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    # ====================================================
    # SHEET 2: Bearing_Calculator
    # ====================================================
    ws_calc = wb.create_sheet(title="Bearing_Calculator")
    ws_calc.views.sheetView[0].showGridLines = True
    
    ws_calc.column_dimensions["A"].width = 5
    ws_calc.column_dimensions["B"].width = 38
    ws_calc.column_dimensions["C"].width = 20
    ws_calc.column_dimensions["D"].width = 20
    ws_calc.column_dimensions["E"].width = 15
    ws_calc.column_dimensions["F"].width = 42
    
    # Title Block
    ws_calc.cell(row=1, column=2, value="FST 16 Wheel Bearing Sizing & Verification Calculator").font = title_font
    ws_calc.cell(row=2, column=2, value="Formula Student Lisboa - Hub & Upright Wheel Assembly").font = note_font
    ws_calc.row_dimensions[1].height = 25
    ws_calc.row_dimensions[2].height = 18
    
    current_row = 4
    
    def add_section_header(title):
        nonlocal current_row
        ws_calc.row_dimensions[current_row].height = 22
        cell = ws_calc.cell(row=current_row, column=2, value=title)
        cell.font = section_font
        cell.fill = section_fill
        cell.alignment = align_left
        for col in range(3, 7):
            c = ws_calc.cell(row=current_row, column=col)
            c.fill = section_fill
        current_row += 1

    def add_table_header(col_labels):
        nonlocal current_row
        ws_calc.row_dimensions[current_row].height = 20
        for idx, lbl in enumerate(col_labels):
            cell = ws_calc.cell(row=current_row, column=idx + 2, value=lbl)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = align_center
            cell.border = border_all
        current_row += 1

    # ----------------------------------------------------
    # SECTION 1: CANDIDATE BEARING SELECTION
    # ----------------------------------------------------
    add_section_header("1. BEARING SELECTION & CANDIDATE ARCHITECTURE")
    add_table_header(["Interface / Position", "Bearing Designation", "Series / Angle", "Reference Link"])
    
    # Outer Bearing (Row current_row, e.g. 6)
    r_outer = current_row
    ws_calc.cell(row=r_outer, column=2, value="Outer Bearing (Rim Side)").font = bold_font
    ws_calc.cell(row=r_outer, column=2).border = border_all
    cell_outer = ws_calc.cell(row=r_outer, column=3, value="Schaeffler 61816-Y")
    cell_outer.font = bold_font
    cell_outer.fill = input_fill
    cell_outer.alignment = align_center
    cell_outer.border = border_all
    
    ws_calc.cell(row=r_outer, column=4, value='=VLOOKUP(C' + str(r_outer) + ', Bearing_Database!A4:S25, 2, FALSE)').font = normal_font
    ws_calc.cell(row=r_outer, column=4).border = border_all
    ws_calc.cell(row=r_outer, column=5, value='="Bore: " & VLOOKUP(C' + str(r_outer) + ', Bearing_Database!A4:S25, 3, FALSE) & " mm | Angle: " & VLOOKUP(C' + str(r_outer) + ', Bearing_Database!A4:S25, 9, FALSE) & " deg"').font = note_font
    ws_calc.cell(row=r_outer, column=5).border = border_all
    current_row += 1
    
    # Inner Bearing (Row current_row, e.g. 7)
    r_inner = current_row
    ws_calc.cell(row=r_inner, column=2, value="Inner Bearing (Motor Side)").font = bold_font
    ws_calc.cell(row=r_inner, column=2).border = border_all
    cell_inner = ws_calc.cell(row=r_inner, column=3, value="Schaeffler 61816-Y")
    cell_inner.font = bold_font
    cell_inner.fill = input_fill
    cell_inner.alignment = align_center
    cell_inner.border = border_all
    
    ws_calc.cell(row=r_inner, column=4, value='=VLOOKUP(C' + str(r_inner) + ', Bearing_Database!A4:S25, 2, FALSE)').font = normal_font
    ws_calc.cell(row=r_inner, column=4).border = border_all
    ws_calc.cell(row=r_inner, column=5, value='="Bore: " & VLOOKUP(C' + str(r_inner) + ', Bearing_Database!A4:S25, 3, FALSE) & " mm | Angle: " & VLOOKUP(C' + str(r_inner) + ', Bearing_Database!A4:S25, 9, FALSE) & " deg"').font = note_font
    ws_calc.cell(row=r_inner, column=5).border = border_all
    current_row += 2

    # Data validation dropdown for bearing selection cells
    dv_bearings = DataValidation(type="list", formula1="Bearing_Database!$A$4:$A$20", allow_blank=True)
    dv_bearings.error = "Select a bearing from candidate catalog"
    dv_bearings.errorTitle = "Invalid Bearing"
    dv_bearings.prompt = "Choose candidate bearing from catalog"
    dv_bearings.promptTitle = "Bearing Catalog"
    ws_calc.add_data_validation(dv_bearings)
    dv_bearings.add(f"C{r_outer}")
    dv_bearings.add(f"C{r_inner}")
    
    # ----------------------------------------------------
    # SECTION 2: OPERATING CONDITIONS & VEHICLE DYNAMICS
    # ----------------------------------------------------
    add_section_header("2. OPERATING CONDITIONS & VEHICLE DYNAMICS (PEAK CORNERING)")
    add_table_header(["Operating Parameter", "Value", "Unit", "Physical Engineering Note"])
    
    inputs_data = [
        ("Tire Lateral Force (Fy)", 4436.26, "N", "Peak 1.8g steady-state cornering (LC3)", "0.0"),
        ("Tire Vertical Wheel Load (Fz)", 1787.31, "N", "Vertical wheel load under weight transfer", "0.0"),
        ("Loaded Tire Rolling Radius (Reff)", 0.250, "m", "Hoosier 18.0x6.0-10 under racing load", "0.000"),
        ("Bearing Physical Center Span (Lb)", 32.0, "mm", "Center-to-center raceway distance", "0.0"),
        ("Vehicle Track Speed (v)", 100.0, "km/h", "Operating cornering entry speed", "0.0"),
        ("Track Operating Temperature (T)", 80.0, "deg C", "Equilibrium temperature under braking & motor", "0.0"),
        ("Thread Pitch (M78x1.5)", 1.50, "mm", "Thread pitch for Preload Nut clamping", "0.00")
    ]
    
    input_row_map = {}
    for param, val, unit, note, num_fmt in inputs_data:
        r = current_row
        input_row_map[param] = r
        ws_calc.cell(row=r, column=2, value=param).font = bold_font
        ws_calc.cell(row=r, column=2).border = border_all
        c_val = ws_calc.cell(row=r, column=3, value=val)
        c_val.font = bold_font
        c_val.fill = input_fill
        c_val.alignment = align_right
        c_val.number_format = num_fmt
        c_val.border = border_all
        
        c_u = ws_calc.cell(row=r, column=4, value=unit)
        c_u.font = normal_font
        c_u.alignment = align_center
        c_u.border = border_all
        
        c_n = ws_calc.cell(row=r, column=5, value=note)
        c_n.font = note_font
        c_n.border = border_all
        current_row += 1
        
    current_row += 1
    
    r_fy = input_row_map["Tire Lateral Force (Fy)"]
    r_fz = input_row_map["Tire Vertical Wheel Load (Fz)"]
    r_reff = input_row_map["Loaded Tire Rolling Radius (Reff)"]
    r_lb = input_row_map["Bearing Physical Center Span (Lb)"]
    r_v = input_row_map["Vehicle Track Speed (v)"]
    r_temp = input_row_map["Track Operating Temperature (T)"]
    r_pitch = input_row_map["Thread Pitch (M78x1.5)"]
    
    # ----------------------------------------------------
    # SECTION 3: LOAD DISTRIBUTION & OVERTURNING COUPLE
    # ----------------------------------------------------
    add_section_header("3. LOAD DISTRIBUTION & CONTACT MECHANICS")
    add_table_header(["Calculated Parameter", "Formula / Value", "Unit", "Engineering Verification"])
    
    r_mx = current_row
    r_aout = current_row + 1
    r_ain = current_row + 2
    r_h = current_row + 3
    r_couple = current_row + 4
    r_frout = current_row + 5
    r_frin = current_row + 6
    r_fa = current_row + 7
    
    ws_calc.cell(row=r_mx, column=2, value="Overturning Moment (Mx)").font = bold_font
    ws_calc.cell(row=r_mx, column=2).border = border_all
    c = ws_calc.cell(row=r_mx, column=3, value=f"=C{r_fy}*C{r_reff}")
    c.font = bold_font; c.alignment = align_right; c.number_format = "0.00"; c.border = border_all
    ws_calc.cell(row=r_mx, column=4, value="N*m").font = normal_font; ws_calc.cell(row=r_mx, column=4).border = border_all; ws_calc.cell(row=r_mx, column=4).alignment = align_center
    ws_calc.cell(row=r_mx, column=5, value="Mx = Fy * Reff").font = note_font; ws_calc.cell(row=r_mx, column=5).border = border_all
    
    ws_calc.cell(row=r_aout, column=2, value="Outer Pressure Apex Distance (a_out)").font = normal_font
    ws_calc.cell(row=r_aout, column=2).border = border_all
    c = ws_calc.cell(row=r_aout, column=3, value=f'=IF(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE)>0, (VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 6, FALSE)/2)*TAN(RADIANS(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE))), 0)')
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.00"; c.border = border_all
    ws_calc.cell(row=r_aout, column=4, value="mm").font = normal_font; ws_calc.cell(row=r_aout, column=4).border = border_all; ws_calc.cell(row=r_aout, column=4).alignment = align_center
    ws_calc.cell(row=r_aout, column=5, value="a_out = (dm/2)*tan(alpha)").font = note_font; ws_calc.cell(row=r_aout, column=5).border = border_all

    ws_calc.cell(row=r_ain, column=2, value="Inner Pressure Apex Distance (a_in)").font = normal_font
    ws_calc.cell(row=r_ain, column=2).border = border_all
    c = ws_calc.cell(row=r_ain, column=3, value=f'=IF(VLOOKUP(C{r_inner}, Bearing_Database!A4:S25, 9, FALSE)>0, (VLOOKUP(C{r_inner}, Bearing_Database!A4:S25, 6, FALSE)/2)*TAN(RADIANS(VLOOKUP(C{r_inner}, Bearing_Database!A4:S25, 9, FALSE))), 0)')
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.00"; c.border = border_all
    ws_calc.cell(row=r_ain, column=4, value="mm").font = normal_font; ws_calc.cell(row=r_ain, column=4).border = border_all; ws_calc.cell(row=r_ain, column=4).alignment = align_center
    ws_calc.cell(row=r_ain, column=5, value="a_in = (dm/2)*tan(alpha)").font = note_font; ws_calc.cell(row=r_ain, column=5).border = border_all

    ws_calc.cell(row=r_h, column=2, value="Effective Bearing Spread (H)").font = bold_font
    ws_calc.cell(row=r_h, column=2).border = border_all
    c = ws_calc.cell(row=r_h, column=3, value=f"=C{r_lb}+C{r_aout}+C{r_ain}")
    c.font = bold_font; c.fill = result_fill; c.alignment = align_right; c.number_format = "0.00"; c.border = border_result
    ws_calc.cell(row=r_h, column=4, value="mm").font = bold_font; ws_calc.cell(row=r_h, column=4).border = border_all; ws_calc.cell(row=r_h, column=4).alignment = align_center
    ws_calc.cell(row=r_h, column=5, value="H = Lb + a_out + a_in (O-arrangement spread)").font = note_font; ws_calc.cell(row=r_h, column=5).border = border_all

    ws_calc.cell(row=r_couple, column=2, value="Decoupled Overturning Couple (F_couple)").font = bold_font
    ws_calc.cell(row=r_couple, column=2).border = border_all
    c = ws_calc.cell(row=r_couple, column=3, value=f"=C{r_mx}/(C{r_h}/1000)")
    c.font = bold_font; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_all
    ws_calc.cell(row=r_couple, column=4, value="N").font = normal_font; ws_calc.cell(row=r_couple, column=4).border = border_all; ws_calc.cell(row=r_couple, column=4).alignment = align_center
    ws_calc.cell(row=r_couple, column=5, value="F_couple = Mx / H").font = note_font; ws_calc.cell(row=r_couple, column=5).border = border_all

    ws_calc.cell(row=r_frout, column=2, value="Peak Outer Radial Force (Fr_out)").font = bold_font
    ws_calc.cell(row=r_frout, column=2).border = border_all
    c = ws_calc.cell(row=r_frout, column=3, value=f"=C{r_couple}+(C{r_fz}/2)")
    c.font = bold_font; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_all
    ws_calc.cell(row=r_frout, column=4, value="N").font = normal_font; ws_calc.cell(row=r_frout, column=4).border = border_all; ws_calc.cell(row=r_frout, column=4).alignment = align_center
    ws_calc.cell(row=r_frout, column=5, value="Fr_out = F_couple + (Fz / 2)").font = note_font; ws_calc.cell(row=r_frout, column=5).border = border_all

    ws_calc.cell(row=r_frin, column=2, value="Peak Inner Radial Force (Fr_in)").font = bold_font
    ws_calc.cell(row=r_frin, column=2).border = border_all
    c = ws_calc.cell(row=r_frin, column=3, value=f"=ABS(-C{r_couple}+(C{r_fz}/2))")
    c.font = bold_font; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_all
    ws_calc.cell(row=r_frin, column=4, value="N").font = normal_font; ws_calc.cell(row=r_frin, column=4).border = border_all; ws_calc.cell(row=r_frin, column=4).alignment = align_center
    ws_calc.cell(row=r_frin, column=5, value="Fr_in = |-F_couple + (Fz / 2)|").font = note_font; ws_calc.cell(row=r_frin, column=5).border = border_all

    ws_calc.cell(row=r_fa, column=2, value="Axial Thrust Force (Fa)").font = bold_font
    ws_calc.cell(row=r_fa, column=2).border = border_all
    c = ws_calc.cell(row=r_fa, column=3, value=f"=C{r_fy}")
    c.font = bold_font; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_all
    ws_calc.cell(row=r_fa, column=4, value="N").font = normal_font; ws_calc.cell(row=r_fa, column=4).border = border_all; ws_calc.cell(row=r_fa, column=4).alignment = align_center
    ws_calc.cell(row=r_fa, column=5, value="Fa = Fy (cornering thrust)").font = note_font; ws_calc.cell(row=r_fa, column=5).border = border_all

    current_row = r_fa + 2

    # ----------------------------------------------------
    # SECTION 4: STATIC SAFETY (ISO 76) & FATIGUE LIFE (ISO 281)
    # ----------------------------------------------------
    add_section_header("4. STATIC SAFETY FACTOR (ISO 76) & RATING LIFE (ISO 281)")
    add_table_header(["Life / Safety Parameter", "Outer Bearing", "Inner Bearing", "Unit", "Physical Engineering Formulation"])
    
    # Static equivalent load P0
    r_p0 = current_row
    ws_calc.cell(row=r_p0, column=2, value="Equivalent Static Load (P0)").font = bold_font
    ws_calc.cell(row=r_p0, column=2).border = border_all
    c_p0_out = ws_calc.cell(row=r_p0, column=3, value=f'=IF(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE)>0, MAX(C{r_frout}, 0.5*C{r_frout}+0.38*C{r_fa}), MAX(C{r_frout}, 0.6*C{r_frout}+0.5*C{r_fa}))')
    c_p0_out.font = bold_font; c_p0_out.alignment = align_right; c_p0_out.number_format = "#,##0.0"; c_p0_out.border = border_all
    c_p0_in = ws_calc.cell(row=r_p0, column=4, value=f'=IF(VLOOKUP(C{r_inner}, Bearing_Database!A4:S25, 9, FALSE)>0, MAX(C{r_frin}, 0.5*C{r_frin}+0.38*C{r_fa}), MAX(C{r_frin}, 0.6*C{r_frin}+0.5*C{r_fa}))')
    c_p0_in.font = bold_font; c_p0_in.alignment = align_right; c_p0_in.number_format = "#,##0.0"; c_p0_in.border = border_all
    ws_calc.cell(row=r_p0, column=5, value="N").font = normal_font; ws_calc.cell(row=r_p0, column=5).alignment = align_center; ws_calc.cell(row=r_p0, column=5).border = border_all
    ws_calc.cell(row=r_p0, column=6, value="ISO 76 static equivalent load").font = note_font; ws_calc.cell(row=r_p0, column=6).border = border_all
    current_row += 1
    
    # Static Safety Factor S0
    r_s0 = current_row
    ws_calc.cell(row=r_s0, column=2, value="Static Safety Factor (S0 = C0r / P0)").font = bold_font
    ws_calc.cell(row=r_s0, column=2).border = border_all
    c_s0_out = ws_calc.cell(row=r_s0, column=3, value=f'=(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 8, FALSE)*1000)/C{r_p0}')
    c_s0_out.font = bold_font; c_s0_out.fill = result_fill; c_s0_out.alignment = align_right; c_s0_out.number_format = "0.00"; c_s0_out.border = border_result
    c_s0_in = ws_calc.cell(row=r_s0, column=4, value=f'=(VLOOKUP(C{r_inner}, Bearing_Database!A4:S25, 8, FALSE)*1000)/D{r_p0}') # Fixed: references D{r_p0}
    c_s0_in.font = bold_font; c_s0_in.fill = result_fill; c_s0_in.alignment = align_right; c_s0_in.number_format = "0.00"; c_s0_in.border = border_result
    ws_calc.cell(row=r_s0, column=5, value="ratio").font = normal_font; ws_calc.cell(row=r_s0, column=5).alignment = align_center; ws_calc.cell(row=r_s0, column=5).border = border_all
    ws_calc.cell(row=r_s0, column=6, value="Target S0 >= 1.5 (Green >= 1.5, Yellow 1.0-1.5, Red < 1.0)").font = note_font; ws_calc.cell(row=r_s0, column=6).border = border_all
    current_row += 1
    
    # Wheel rotational speed (n_wheel)
    r_nwheel = current_row
    ws_calc.cell(row=r_nwheel, column=2, value="Wheel Rotational Speed (n_wheel)").font = normal_font
    ws_calc.cell(row=r_nwheel, column=2).border = border_all
    c_nw_out = ws_calc.cell(row=r_nwheel, column=3, value=f'=((C{r_v}*1000/3600)/(2*PI()*C{r_reff}))*60')
    c_nw_out.font = normal_font; c_nw_out.alignment = align_right; c_nw_out.number_format = "#,##0.0"; c_nw_out.border = border_all
    c_nw_in = ws_calc.cell(row=r_nwheel, column=4, value=f'=C{r_nwheel}')
    c_nw_in.font = normal_font; c_nw_in.alignment = align_right; c_nw_in.number_format = "#,##0.0"; c_nw_in.border = border_all
    ws_calc.cell(row=r_nwheel, column=5, value="rpm").font = normal_font; ws_calc.cell(row=r_nwheel, column=5).alignment = align_center; ws_calc.cell(row=r_nwheel, column=5).border = border_all
    ws_calc.cell(row=r_nwheel, column=6, value="n = (v / (2 * pi * Reff)) * 60").font = note_font; ws_calc.cell(row=r_nwheel, column=6).border = border_all
    current_row += 1
    
    # Dynamic equivalent load across duty spectrum (Peq)
    # Duty spectrum: 60% cruise (1250 N), 30% cornering 1.0g (0.55*Fr), 8% peak 1.8g (Fr), 2% curb (2600 N)
    r_peq = current_row
    ws_calc.cell(row=r_peq, column=2, value="Dynamic Equivalent Spectrum Load (Peq)").font = bold_font
    ws_calc.cell(row=r_peq, column=2).border = border_all
    c_peq_out = ws_calc.cell(row=r_peq, column=3, value=f'=(0.60*(1250^3) + 0.30*((0.55*C{r_frout})^3) + 0.08*(C{r_frout}^3) + 0.02*(2600^3))^(1/3)')
    c_peq_out.font = bold_font; c_peq_out.alignment = align_right; c_peq_out.number_format = "#,##0.0"; c_peq_out.border = border_all
    c_peq_in = ws_calc.cell(row=r_peq, column=4, value=f'=(0.60*(1250^3) + 0.30*((0.55*C{r_frin})^3) + 0.08*(C{r_frin}^3) + 0.02*(2600^3))^(1/3)')
    c_peq_in.font = bold_font; c_peq_in.alignment = align_right; c_peq_in.number_format = "#,##0.0"; c_peq_in.border = border_all
    ws_calc.cell(row=r_peq, column=5, value="N").font = normal_font; ws_calc.cell(row=r_peq, column=5).alignment = align_center; ws_calc.cell(row=r_peq, column=5).border = border_all
    ws_calc.cell(row=r_peq, column=6, value="ISO 281 cubic duty spectrum mean load").font = note_font; ws_calc.cell(row=r_peq, column=6).border = border_all
    current_row += 1
    
    # ISO 281 Fatigue Life L10h
    r_l10h = current_row
    ws_calc.cell(row=r_l10h, column=2, value="Basic Rating Fatigue Life (L10h)").font = bold_font
    ws_calc.cell(row=r_l10h, column=2).border = border_all
    c_l10h_out = ws_calc.cell(row=r_l10h, column=3, value=f'=(10^6/(60*C{r_nwheel})) * ((VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 7, FALSE)*1000)/C{r_peq})^3')
    c_l10h_out.font = bold_font; c_l10h_out.fill = result_fill; c_l10h_out.alignment = align_right; c_l10h_out.number_format = "#,##0.0"; c_l10h_out.border = border_result
    c_l10h_in = ws_calc.cell(row=r_l10h, column=4, value=f'=(10^6/(60*C{r_nwheel})) * ((VLOOKUP(C{r_inner}, Bearing_Database!A4:S25, 7, FALSE)*1000)/D{r_peq})^3') # Fixed: references D{r_peq}
    c_l10h_in.font = bold_font; c_l10h_in.fill = result_fill; c_l10h_in.alignment = align_right; c_l10h_in.number_format = "#,##0.0"; c_l10h_in.border = border_result
    ws_calc.cell(row=r_l10h, column=5, value="Hours").font = normal_font; ws_calc.cell(row=r_l10h, column=5).alignment = align_center; ws_calc.cell(row=r_l10h, column=5).border = border_all
    ws_calc.cell(row=r_l10h, column=6, value="L10h = (10^6 / (60*n)) * (Cr / Peq)^3").font = note_font; ws_calc.cell(row=r_l10h, column=6).border = border_all
    current_row += 1
    
    # Distance Life (km)
    r_l10km = current_row
    ws_calc.cell(row=r_l10km, column=2, value="Equivalent Distance Life (L10km)").font = normal_font
    ws_calc.cell(row=r_l10km, column=2).border = border_all
    c_km_out = ws_calc.cell(row=r_l10km, column=3, value=f'=C{r_l10h}*C{r_v}')
    c_km_out.font = normal_font; c_km_out.alignment = align_right; c_km_out.number_format = "#,##0"; c_km_out.border = border_all
    c_km_in = ws_calc.cell(row=r_l10km, column=4, value=f'=D{r_l10h}*C{r_v}') # Fixed: references D{r_l10h}
    c_km_in.font = normal_font; c_km_in.alignment = align_right; c_km_in.number_format = "#,##0"; c_km_in.border = border_all
    ws_calc.cell(row=r_l10km, column=5, value="km").font = normal_font; ws_calc.cell(row=r_l10km, column=5).alignment = align_center; ws_calc.cell(row=r_l10km, column=5).border = border_all
    ws_calc.cell(row=r_l10km, column=6, value="L10km = L10h * v (Full FS season ~ 500 km)").font = note_font; ws_calc.cell(row=r_l10km, column=6).border = border_all
    current_row += 2
    
    # ----------------------------------------------------
    # SECTION 5: ELASTOKINEMATIC CAMBER STIFFNESS
    # ----------------------------------------------------
    add_section_header("5. ELASTOKINEMATIC CAMBER STIFFNESS & ANGULAR COMPLIANCE")
    add_table_header(["Rigidity Parameter", "Value", "Unit", "Physical Engineering Note"])
    
    # Tilting stiffness c_tilt = 0.5 * Kr * H^2
    # Kr in kN/mm -> 1 kN/mm = 1e6 N/m
    r_ctilt = current_row
    ws_calc.cell(row=r_ctilt, column=2, value="Angular Camber Stiffness (c_tilt)").font = bold_font
    ws_calc.cell(row=r_ctilt, column=2).border = border_all
    c = ws_calc.cell(row=r_ctilt, column=3, value=f'=0.5 * (VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 14, FALSE)*1000000) * ((C{r_h}/1000)^2) * (PI()/180)')
    c.font = bold_font; c.fill = result_fill; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_result
    ws_calc.cell(row=r_ctilt, column=4, value="N*m / deg").font = bold_font; ws_calc.cell(row=r_ctilt, column=4).border = border_all; ws_calc.cell(row=r_ctilt, column=4).alignment = align_center
    ws_calc.cell(row=r_ctilt, column=5, value="c_tilt = 0.5 * Kr * H^2 (N*m/rad converted to N*m/deg)").font = note_font; ws_calc.cell(row=r_ctilt, column=5).border = border_all
    current_row += 1
    
    # Camber angular deflection gamma = Mx / c_tilt
    r_gamma = current_row
    ws_calc.cell(row=r_gamma, column=2, value="Angular Camber Deflection (gamma)").font = bold_font
    ws_calc.cell(row=r_gamma, column=2).border = border_all
    c = ws_calc.cell(row=r_gamma, column=3, value=f'=C{r_mx}/C{r_ctilt}')
    c.font = bold_font; c.fill = result_fill; c.alignment = align_right; c.number_format = "0.000"; c.border = border_result
    ws_calc.cell(row=r_gamma, column=4, value="deg").font = bold_font; ws_calc.cell(row=r_gamma, column=4).border = border_all; ws_calc.cell(row=r_gamma, column=4).alignment = align_center
    ws_calc.cell(row=r_gamma, column=5, value="Elastokinematic hub deflection under peak cornering").font = note_font; ws_calc.cell(row=r_gamma, column=5).border = border_all
    current_row += 1

    # Tire Top Lateral Deflection
    r_topdef = current_row
    ws_calc.cell(row=r_topdef, column=2, value="Tire Top Lateral Deflection").font = normal_font
    ws_calc.cell(row=r_topdef, column=2).border = border_all
    c = ws_calc.cell(row=r_topdef, column=3, value=f'=C{r_reff}*RADIANS(C{r_gamma})*1000')
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.00"; c.border = border_all
    ws_calc.cell(row=r_topdef, column=4, value="mm").font = normal_font; ws_calc.cell(row=r_topdef, column=4).border = border_all; ws_calc.cell(row=r_topdef, column=4).alignment = align_center
    ws_calc.cell(row=r_topdef, column=5, value="Reff * radians(gamma) * 1000 lateral displacement").font = note_font; ws_calc.cell(row=r_topdef, column=5).border = border_all
    current_row += 1
    
    # Relative Camber Stiffness Gain vs Baseline
    r_gain = current_row
    ws_calc.cell(row=r_gain, column=2, value="Rigidity Ratio vs Physical Span").font = normal_font
    ws_calc.cell(row=r_gain, column=2).border = border_all
    c = ws_calc.cell(row=r_gain, column=3, value=f'=(C{r_h}/C{r_lb})^2')
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.00"; c.border = border_all
    ws_calc.cell(row=r_gain, column=4, value="ratio (x)").font = normal_font; ws_calc.cell(row=r_gain, column=4).border = border_all; ws_calc.cell(row=r_gain, column=4).alignment = align_center
    ws_calc.cell(row=r_gain, column=5, value="Scaling factor (H / Lb)^2").font = note_font; ws_calc.cell(row=r_gain, column=5).border = border_all
    current_row += 2
    
    # ----------------------------------------------------
    # SECTION 6: PRELOAD & PRELOAD NUT TIGHTENING TORQUE (MOTOSH)
    # ----------------------------------------------------
    add_section_header("6. PRELOAD MECHANICS & M78x1.5 PRELOAD NUT TORQUE")
    add_table_header(["Fastener / Preload Parameter", "Value", "Unit", "Mathematical Formulation"])
    
    # Required axial preload force (recognizes UL, UM, UH, and DGBB)
    r_fpreload = current_row
    ws_calc.cell(row=r_fpreload, column=2, value="Required Axial Preload (F_preload)").font = bold_font
    ws_calc.cell(row=r_fpreload, column=2).border = border_all
    preload_formula = (
        f'=IF(ISNUMBER(SEARCH("-UH", C{r_outer})), 712, '
        f'IF(ISNUMBER(SEARCH("-UM", C{r_outer})), 302, '
        f'IF(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE)>0, 84, '
        f'(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 15, FALSE)*1000)*0.100)))'
    )
    c = ws_calc.cell(row=r_fpreload, column=3, value=preload_formula)
    c.font = bold_font; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_all
    ws_calc.cell(row=r_fpreload, column=4, value="N").font = normal_font; ws_calc.cell(row=r_fpreload, column=4).border = border_all; ws_calc.cell(row=r_fpreload, column=4).alignment = align_center
    ws_calc.cell(row=r_fpreload, column=5, value="Catalog class (84N UL, 302N UM, 712N UH, or Ka*0.1mm play)").font = note_font; ws_calc.cell(row=r_fpreload, column=5).border = border_all
    current_row += 1
    
    # Thread pitch diameter d2
    r_d2 = current_row
    ws_calc.cell(row=r_d2, column=2, value="Thread Pitch Diameter (d2)").font = normal_font
    ws_calc.cell(row=r_d2, column=2).border = border_all
    c = ws_calc.cell(row=r_d2, column=3, value=77.026)
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.000"; c.border = border_all
    ws_calc.cell(row=r_d2, column=4, value="mm").font = normal_font; ws_calc.cell(row=r_d2, column=4).border = border_all; ws_calc.cell(row=r_d2, column=4).alignment = align_center
    ws_calc.cell(row=r_d2, column=5, value="M78x1.5 ISO metric fine thread").font = note_font; ws_calc.cell(row=r_d2, column=5).border = border_all
    current_row += 1
    
    # Thread lead angle psi
    r_psi = current_row
    ws_calc.cell(row=r_psi, column=2, value="Thread Lead Angle (psi)").font = normal_font
    ws_calc.cell(row=r_psi, column=2).border = border_all
    c = ws_calc.cell(row=r_psi, column=3, value=f'=DEGREES(ATAN(C{r_pitch}/(PI()*C{r_d2})))')
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.000"; c.border = border_all
    ws_calc.cell(row=r_psi, column=4, value="deg").font = normal_font; ws_calc.cell(row=r_psi, column=4).border = border_all; ws_calc.cell(row=r_psi, column=4).alignment = align_center
    ws_calc.cell(row=r_psi, column=5, value="psi = atan(pitch / (pi * d2))").font = note_font; ws_calc.cell(row=r_psi, column=5).border = border_all
    current_row += 1
    
    # Friction coefficient mu
    r_mu = current_row
    ws_calc.cell(row=r_mu, column=2, value="Friction Coefficient (mu_thread = mu_collar)").font = normal_font
    ws_calc.cell(row=r_mu, column=2).border = border_all
    c = ws_calc.cell(row=r_mu, column=3, value=0.150)
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.000"; c.border = border_all
    ws_calc.cell(row=r_mu, column=4, value="-").font = normal_font; ws_calc.cell(row=r_mu, column=4).border = border_all; ws_calc.cell(row=r_mu, column=4).alignment = align_center
    ws_calc.cell(row=r_mu, column=5, value="Lightly lubricated aluminum-to-steel interface").font = note_font; ws_calc.cell(row=r_mu, column=5).border = border_all
    current_row += 1
    
    # Mean collar bearing diameter dc
    r_dc = current_row
    ws_calc.cell(row=r_dc, column=2, value="Collar Contact Mean Diameter (dc)").font = normal_font
    ws_calc.cell(row=r_dc, column=2).border = border_all
    c = ws_calc.cell(row=r_dc, column=3, value=81.5)
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.0"; c.border = border_all
    ws_calc.cell(row=r_dc, column=4, value="mm").font = normal_font; ws_calc.cell(row=r_dc, column=4).border = border_all; ws_calc.cell(row=r_dc, column=4).alignment = align_center
    ws_calc.cell(row=r_dc, column=5, value="Nut face contact against inner ring face").font = note_font; ws_calc.cell(row=r_dc, column=5).border = border_all
    current_row += 1
    
    # Motosh Preload Nut Torque
    r_torque = current_row
    ws_calc.cell(row=r_torque, column=2, value="Preload Nut Tightening Torque (T)").font = bold_font
    ws_calc.cell(row=r_torque, column=2).border = border_all
    motosh_formula = f'=C{r_fpreload} * ( (C{r_d2}/2000)*((TAN(RADIANS(C{r_psi})) + C{r_mu}/COS(RADIANS(30))) / (1 - C{r_mu}*TAN(RADIANS(C{r_psi}))/COS(RADIANS(30)))) + C{r_mu}*(C{r_dc}/2000) )'
    c = ws_calc.cell(row=r_torque, column=3, value=motosh_formula)
    c.font = bold_font; c.fill = result_fill; c.alignment = align_right; c.number_format = "0.00"; c.border = border_result
    ws_calc.cell(row=r_torque, column=4, value="N*m").font = bold_font; ws_calc.cell(row=r_torque, column=4).border = border_all; ws_calc.cell(row=r_torque, column=4).alignment = align_center
    ws_calc.cell(row=r_torque, column=5, value="Motosh tightening torque (FSG cotter pin locked)").font = note_font; ws_calc.cell(row=r_torque, column=5).border = border_all
    current_row += 2
    
    # ----------------------------------------------------
    # SECTION 7: THERMAL FIT EXPANSION & CLEARANCE LOSS
    # ----------------------------------------------------
    add_section_header("7. THERMAL FIT EXPANSION & CLEARANCE LOSS (AL 7075 vs 100Cr6)")
    add_table_header(["Thermal Fit Parameter", "Track Operating Temp", "140 deg C (Extreme Brake)", "Unit / Physical Meaning"])
    
    # Thermal expansion coefficients
    r_alpha = current_row
    ws_calc.cell(row=r_alpha, column=2, value="Differential Expansion (Delta_alpha)").font = normal_font
    ws_calc.cell(row=r_alpha, column=2).border = border_all
    ws_calc.cell(row=r_alpha, column=3, value=11.9e-6).font = normal_font; ws_calc.cell(row=r_alpha, column=3).alignment = align_right; ws_calc.cell(row=r_alpha, column=3).number_format = "0.00E+00"; ws_calc.cell(row=r_alpha, column=3).border = border_all
    ws_calc.cell(row=r_alpha, column=4, value=11.9e-6).font = normal_font; ws_calc.cell(row=r_alpha, column=4).alignment = align_right; ws_calc.cell(row=r_alpha, column=4).number_format = "0.00E+00"; ws_calc.cell(row=r_alpha, column=4).border = border_all
    ws_calc.cell(row=r_alpha, column=5, value="1/K (alpha_Al: 23.4e-6 - alpha_Steel: 11.5e-6)").font = note_font; ws_calc.cell(row=r_alpha, column=5).border = border_all
    current_row += 1
    
    # Temperature delta
    r_dt = current_row
    ws_calc.cell(row=r_dt, column=2, value="Temperature Delta (Delta_T = T - 20 deg C)").font = normal_font
    ws_calc.cell(row=r_dt, column=2).border = border_all
    ws_calc.cell(row=r_dt, column=3, value=f'=C{r_temp}-20').font = normal_font; ws_calc.cell(row=r_dt, column=3).alignment = align_right; ws_calc.cell(row=r_dt, column=3).number_format = "0.0"; ws_calc.cell(row=r_dt, column=3).border = border_all
    ws_calc.cell(row=r_dt, column=4, value=120.0).font = normal_font; ws_calc.cell(row=r_dt, column=4).alignment = align_right; ws_calc.cell(row=r_dt, column=4).number_format = "0.0"; ws_calc.cell(row=r_dt, column=4).border = border_all
    ws_calc.cell(row=r_dt, column=5, value="K (thermal gradient from ambient 20 deg C)").font = note_font; ws_calc.cell(row=r_dt, column=5).border = border_all
    current_row += 1
    
    # Hub Bore Thermal Expansion
    r_exp_bore = current_row
    ws_calc.cell(row=r_exp_bore, column=2, value="Hub Bore Expansion (Delta_D_hub)").font = bold_font
    ws_calc.cell(row=r_exp_bore, column=2).border = border_all
    c_b80 = ws_calc.cell(row=r_exp_bore, column=3, value=f'=100.0 * C{r_alpha} * C{r_dt}')
    c_b80.font = bold_font; c_b80.alignment = align_right; c_b80.number_format = "0.000"; c_b80.border = border_all
    c_b140 = ws_calc.cell(row=r_exp_bore, column=4, value=f'=100.0 * D{r_alpha} * D{r_dt}')
    c_b140.font = bold_font; c_b140.alignment = align_right; c_b140.number_format = "0.000"; c_b140.border = border_all
    ws_calc.cell(row=r_exp_bore, column=5, value="mm (Hub bore grows faster than steel outer ring)").font = note_font; ws_calc.cell(row=r_exp_bore, column=5).border = border_all
    current_row += 1
    
    # Journal Expansion
    r_exp_shaft = current_row
    ws_calc.cell(row=r_exp_shaft, column=2, value="Upright Journal Expansion (Delta_d_shaft)").font = bold_font
    ws_calc.cell(row=r_exp_shaft, column=2).border = border_all
    c_s80 = ws_calc.cell(row=r_exp_shaft, column=3, value=f'=80.0 * C{r_alpha} * C{r_dt}')
    c_s80.font = bold_font; c_s80.alignment = align_right; c_s80.number_format = "0.000"; c_s80.border = border_all
    c_s140 = ws_calc.cell(row=r_exp_shaft, column=4, value=f'=80.0 * D{r_alpha} * D{r_dt}')
    c_s140.font = bold_font; c_s140.alignment = align_right; c_s140.number_format = "0.000"; c_s140.border = border_all
    ws_calc.cell(row=r_exp_shaft, column=5, value="mm (Aluminum shaft expands, locking inner ring)").font = note_font; ws_calc.cell(row=r_exp_shaft, column=5).border = border_all
    current_row += 1
    
    # Net Outer Ring Fit Status
    r_net_fit = current_row
    ws_calc.cell(row=r_net_fit, column=2, value="Effective Fit at Hub Bore (P7 = +0.030 mm)").font = bold_font
    ws_calc.cell(row=r_net_fit, column=2).border = border_all
    c_f80 = ws_calc.cell(row=r_net_fit, column=3, value=f'=0.030 - C{r_exp_bore}')
    c_f80.font = bold_font; c_f80.fill = result_fill; c_f80.alignment = align_right; c_f80.number_format = "+0.000;-0.000;0.000"; c_f80.border = border_result
    c_f140 = ws_calc.cell(row=r_net_fit, column=4, value=f'=0.030 - D{r_exp_bore}')
    c_f140.font = bold_font; c_f140.fill = result_fill; c_f140.alignment = align_right; c_f140.number_format = "+0.000;-0.000;0.000"; c_f140.border = border_result
    ws_calc.cell(row=r_net_fit, column=5, value="mm (Negative = Loose clearance, requires axial clamp!)").font = note_font; ws_calc.cell(row=r_net_fit, column=5).border = border_all
    current_row += 2
    
    # ----------------------------------------------------
    # SECTION 8: KINEMATICS & HIGH-SPEED DYNAMICS
    # ----------------------------------------------------
    add_section_header("8. BEARING KINEMATICS & BALL DYNAMICS")
    add_table_header(["Kinematic Parameter", "Value", "Unit", "Kinematic Relationship"])
    
    r_ndm = current_row
    ws_calc.cell(row=r_ndm, column=2, value="Speed Parameter (n * dm)").font = bold_font
    ws_calc.cell(row=r_ndm, column=2).border = border_all
    c = ws_calc.cell(row=r_ndm, column=3, value=f'=C{r_nwheel} * VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 6, FALSE)')
    c.font = bold_font; c.alignment = align_right; c.number_format = "#,##0"; c.border = border_all
    ws_calc.cell(row=r_ndm, column=4, value="mm*rpm").font = normal_font; ws_calc.cell(row=r_ndm, column=4).border = border_all; ws_calc.cell(row=r_ndm, column=4).alignment = align_center
    ws_calc.cell(row=r_ndm, column=5, value="n * dm (Catalog limit: 1,000,000)").font = note_font; ws_calc.cell(row=r_ndm, column=5).border = border_all
    current_row += 1
    
    r_nc = current_row
    ws_calc.cell(row=r_nc, column=2, value="Cage Orbital Speed (nc)").font = normal_font
    ws_calc.cell(row=r_nc, column=2).border = border_all
    c = ws_calc.cell(row=r_nc, column=3, value=f'=(C{r_nwheel}/2) * (1 - (VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 10, FALSE)/VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 6, FALSE))*COS(RADIANS(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE))))')
    c.font = normal_font; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_all
    ws_calc.cell(row=r_nc, column=4, value="rpm").font = normal_font; ws_calc.cell(row=r_nc, column=4).border = border_all; ws_calc.cell(row=r_nc, column=4).alignment = align_center
    ws_calc.cell(row=r_nc, column=5, value="nc = (n/2) * (1 - (Dw/dm)*cos(alpha))").font = note_font; ws_calc.cell(row=r_nc, column=5).border = border_all
    current_row += 1
    
    r_nb = current_row
    ws_calc.cell(row=r_nb, column=2, value="Ball Spin Speed (nb)").font = normal_font
    ws_calc.cell(row=r_nb, column=2).border = border_all
    c = ws_calc.cell(row=r_nb, column=3, value=f'=(C{r_nwheel}/2) * (VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 6, FALSE)/VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 10, FALSE)) * (1 - ((VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 10, FALSE)/VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 6, FALSE))*COS(RADIANS(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE))))^2)')
    c.font = normal_font; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_all
    ws_calc.cell(row=r_nb, column=4, value="rpm").font = normal_font; ws_calc.cell(row=r_nb, column=4).border = border_all; ws_calc.cell(row=r_nb, column=4).alignment = align_center
    ws_calc.cell(row=r_nb, column=5, value="Ball rotational speed around its own axis").font = note_font; ws_calc.cell(row=r_nb, column=5).border = border_all
    current_row += 1
    
    r_fc = current_row
    ws_calc.cell(row=r_fc, column=2, value="Ball Centrifugal Force (Fc)").font = normal_font
    ws_calc.cell(row=r_fc, column=2).border = border_all
    # Correct formula: Fc = m_ball * r_orbital * omega_c^2 (no extraneous 0.5)
    c = ws_calc.cell(row=r_fc, column=3, value=f'=((4/3)*PI()*((VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 10, FALSE)/2000)^3)*7800) * (VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 6, FALSE)/2000) * ((2*PI()*C{r_nc}/60)^2)')
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.000"; c.border = border_all
    ws_calc.cell(row=r_fc, column=4, value="N").font = normal_font; ws_calc.cell(row=r_fc, column=4).border = border_all; ws_calc.cell(row=r_fc, column=4).alignment = align_center
    ws_calc.cell(row=r_fc, column=5, value="Centrifugal force per rolling element Fc = m * r * omega^2").font = note_font; ws_calc.cell(row=r_fc, column=5).border = border_all
    current_row += 1

    r_mg = current_row
    ws_calc.cell(row=r_mg, column=2, value="Ball Gyroscopic Moment (Mg)").font = normal_font
    ws_calc.cell(row=r_mg, column=2).border = border_all
    # Mg = Jb * omega_b * omega_c * sin(alpha) in N*mm
    # Jb = 0.1 * m * Dw^2
    mg_formula = f'=(0.1 * ((4/3)*PI()*((VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 10, FALSE)/2000)^3)*7800) * ((VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 10, FALSE)/1000)^2)) * (2*PI()*C{r_nb}/60) * (2*PI()*C{r_nc}/60) * SIN(RADIANS(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE))) * 1000'
    c = ws_calc.cell(row=r_mg, column=3, value=mg_formula)
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.0000"; c.border = border_all
    ws_calc.cell(row=r_mg, column=4, value="N*mm").font = normal_font; ws_calc.cell(row=r_mg, column=4).border = border_all; ws_calc.cell(row=r_mg, column=4).alignment = align_center
    ws_calc.cell(row=r_mg, column=5, value="Mg = Jb * omega_b * omega_c * sin(alpha)").font = note_font; ws_calc.cell(row=r_mg, column=5).border = border_all
    current_row += 1

    r_antiskid = current_row
    ws_calc.cell(row=r_antiskid, column=2, value="Anti-Skid Preload Threshold (Fa_skid)").font = bold_font
    ws_calc.cell(row=r_antiskid, column=2).border = border_all
    # For ACBB: Z * (Mg_Nm / (r_ball * mu)) * sin(alpha). For DGBB: 0.01 * C0r
    antiskid_formula = f'=IF(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE)>0, VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 11, FALSE) * ((C{r_mg}/1000) / ((VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 10, FALSE)/2000) * 0.06)) * SIN(RADIANS(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE))), 0.01 * (VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 8, FALSE)*1000))'
    c = ws_calc.cell(row=r_antiskid, column=3, value=antiskid_formula)
    c.font = bold_font; c.alignment = align_right; c.number_format = "#,##0.0"; c.border = border_all
    ws_calc.cell(row=r_antiskid, column=4, value="N").font = normal_font; ws_calc.cell(row=r_antiskid, column=4).border = border_all; ws_calc.cell(row=r_antiskid, column=4).alignment = align_center
    ws_calc.cell(row=r_antiskid, column=5, value="Minimum axial preload to prevent ball skidding under Mg").font = note_font; ws_calc.cell(row=r_antiskid, column=5).border = border_all
    current_row += 2
    
    # ----------------------------------------------------
    # SECTION 9: EHL LUBRICATION (MOTUL GEAR 300 75W90)
    # ----------------------------------------------------
    add_section_header("9. EHL LUBRICATION RATIO (MOTUL GEAR 300 75W90)")
    add_table_header(["Lubrication Parameter", "Value", "Unit", "Physical Meaning & Threshold"])
    
    # Operating viscosity nu at operating temperature (dynamic Walther approximation)
    r_nu = current_row
    ws_calc.cell(row=r_nu, column=2, value="Kinematic Viscosity at Operating Temp (nu)").font = bold_font
    ws_calc.cell(row=r_nu, column=2).border = border_all
    c = ws_calc.cell(row=r_nu, column=3, value=f'=MAX(10.0, 152.0 * EXP(-0.038 * (C{r_temp} - 40)))')
    c.font = bold_font; c.alignment = align_right; c.number_format = "0.0"; c.border = border_all
    ws_calc.cell(row=r_nu, column=4, value="mm^2/s").font = normal_font; ws_calc.cell(row=r_nu, column=4).border = border_all; ws_calc.cell(row=r_nu, column=4).alignment = align_center
    ws_calc.cell(row=r_nu, column=5, value="Motul Gear 300 75W90 ester synthetic at track temp").font = note_font; ws_calc.cell(row=r_nu, column=5).border = border_all
    current_row += 1
    
    # Rated viscosity nu1 (ISO 281)
    r_nu1 = current_row
    ws_calc.cell(row=r_nu1, column=2, value="Reference Viscosity for Full Film (nu1)").font = normal_font
    ws_calc.cell(row=r_nu1, column=2).border = border_all
    c = ws_calc.cell(row=r_nu1, column=3, value=f'=45000 * (C{r_nwheel}^(-0.83)) * ((VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 6, FALSE))^(-0.5))')
    c.font = normal_font; c.alignment = align_right; c.number_format = "0.0"; c.border = border_all
    ws_calc.cell(row=r_nu1, column=4, value="mm^2/s").font = normal_font; ws_calc.cell(row=r_nu1, column=4).border = border_all; ws_calc.cell(row=r_nu1, column=4).alignment = align_center
    ws_calc.cell(row=r_nu1, column=5, value="nu1 = 45000 * n^(-0.83) * dm^(-0.5) (ISO 281)").font = note_font; ws_calc.cell(row=r_nu1, column=5).border = border_all
    current_row += 1
    
    # Viscosity ratio kappa = nu / nu1
    r_kappa = current_row
    ws_calc.cell(row=r_kappa, column=2, value="Viscosity Ratio (kappa = nu / nu1)").font = bold_font
    ws_calc.cell(row=r_kappa, column=2).border = border_all
    c = ws_calc.cell(row=r_kappa, column=3, value=f'=C{r_nu}/C{r_nu1}')
    c.font = bold_font; c.fill = result_fill; c.alignment = align_right; c.number_format = "0.00"; c.border = border_result
    ws_calc.cell(row=r_kappa, column=4, value="ratio").font = bold_font; ws_calc.cell(row=r_kappa, column=4).border = border_all; ws_calc.cell(row=r_kappa, column=4).alignment = align_center
    ws_calc.cell(row=r_kappa, column=5, value="kappa >= 2.0: Full elastohydrodynamic separation").font = note_font; ws_calc.cell(row=r_kappa, column=5).border = border_all
    current_row += 2
    
    # ----------------------------------------------------
    # SECTION 10: MULTI-LOAD CASE STATIC SAFETY SUMMARY
    # ----------------------------------------------------
    add_section_header("10. MULTI-LOAD CASE SAFETY FACTOR SUMMARY (ALL 8 FS CASES)")
    add_table_header(["Load Case", "Fy (N)", "Mx (N*m)", "Fr_outer (N)", "P0_outer (N)", "S0_outer", "Evaluation"])
    
    cases = [
        ("LC1 - Pure Braking", 0.0, 3207.37),
        ("LC2 - Pure Acceleration", 0.0, 2409.45),
        ("LC3 - Pure Cornering (1.8g)", -4436.26, 3574.61),
        ("LC4 - Combined Braking + Cornering 1", -3089.33, 3125.00),
        ("LC5 - Combined Braking + Cornering 2", -2468.43, 2823.56),
        ("LC6 - Combined Accel + Cornering 1", -2228.85, 2130.33),
        ("LC7 - Combined Accel + Cornering 2", -2898.75, 2879.34),
        ("LC Bump - Dynamic Curb Strike (3.5g)", 0.0, 5200.00)
    ]
    
    lc_start_row = current_row
    for lc_name, lc_fy, lc_fz in cases:
        r = current_row
        ws_calc.row_dimensions[r].height = 20
        
        # Name
        ws_calc.cell(row=r, column=2, value=lc_name).font = bold_font
        ws_calc.cell(row=r, column=2).border = border_all
        
        # Fy
        c_fy = ws_calc.cell(row=r, column=3, value=lc_fy)
        c_fy.font = normal_font; c_fy.alignment = align_right; c_fy.number_format = "#,##0.0"; c_fy.border = border_all
        
        # Mx = |Fy| * Reff
        c_mx = ws_calc.cell(row=r, column=4, value=f'=ABS(C{r})*C{r_reff}')
        c_mx.font = normal_font; c_mx.alignment = align_right; c_mx.number_format = "#,##0.0"; c_mx.border = border_all
        
        # Fr_outer = (Mx / H) + (Fz / 2)
        c_fro = ws_calc.cell(row=r, column=5, value=f'=(D{r}/(C{r_h}/1000)) + ({lc_fz}/2)')
        c_fro.font = normal_font; c_fro.alignment = align_right; c_fro.number_format = "#,##0.0"; c_fro.border = border_all
        
        # S0_outer = C0r / P0
        p0_expr = f'IF(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 9, FALSE)>0, MAX(E{r}, 0.5*E{r}+0.38*ABS(C{r})), MAX(E{r}, 0.6*E{r}+0.5*ABS(C{r})))'
        c_s0 = ws_calc.cell(row=r, column=6, value=f'=(VLOOKUP(C{r_outer}, Bearing_Database!A4:S25, 8, FALSE)*1000)/({p0_expr})')
        c_s0.font = bold_font; c_s0.alignment = align_right; c_s0.number_format = "0.00"; c_s0.border = border_all
        
        # Evaluation in column 7
        c_eval = ws_calc.cell(row=r, column=7, value=f'=IF(F{r}>=1.5, "Safe", IF(F{r}>=1.0, "Marginal", "Exceeded"))')
        c_eval.font = bold_font; c_eval.alignment = align_center; c_eval.border = border_all
        
        current_row += 1
        
    lc_end_row = current_row - 1
    
    # Format column 7 header for LC summary
    ws_calc.cell(row=lc_start_row - 1, column=7, value="Compliance").font = header_font
    ws_calc.cell(row=lc_start_row - 1, column=7).fill = header_fill
    ws_calc.cell(row=lc_start_row - 1, column=7).alignment = align_center
    ws_calc.cell(row=lc_start_row - 1, column=7).border = border_all
    ws_calc.column_dimensions["G"].width = 16
    
    # ----------------------------------------------------
    # CONDITIONAL FORMATTING RULES (Green >= 1.5, Yellow 1.0-1.5, Red < 1.0)
    # ----------------------------------------------------
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(name=font_family, size=10, bold=True, color="006100")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    yellow_font = Font(name=font_family, size=10, bold=True, color="9C6500")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(name=font_family, size=10, bold=True, color="9C0006")
    
    rule_green = CellIsRule(operator='greaterThanOrEqual', formula=['1.5'], stopIfTrue=True, fill=green_fill, font=green_font)
    rule_yellow = CellIsRule(operator='between', formula=['1.0', '1.4999'], stopIfTrue=True, fill=yellow_fill, font=yellow_font)
    rule_red = CellIsRule(operator='lessThan', formula=['1.0'], stopIfTrue=True, fill=red_fill, font=red_font)
    
    # Apply to S0 row (C{r_s0}:D{r_s0})
    ws_calc.conditional_formatting.add(f"C{r_s0}:D{r_s0}", rule_green)
    ws_calc.conditional_formatting.add(f"C{r_s0}:D{r_s0}", rule_yellow)
    ws_calc.conditional_formatting.add(f"C{r_s0}:D{r_s0}", rule_red)
    
    # Apply to LC summary table S0 column (F{lc_start_row}:F{lc_end_row})
    ws_calc.conditional_formatting.add(f"F{lc_start_row}:F{lc_end_row}", rule_green)
    ws_calc.conditional_formatting.add(f"F{lc_start_row}:F{lc_end_row}", rule_yellow)
    ws_calc.conditional_formatting.add(f"F{lc_start_row}:F{lc_end_row}", rule_red)
    
    wb.save(output_path)
    print(f"Workbook successfully saved to: {output_path}")

if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "FST16_Wheel_Bearing_Calculator.xlsx")
    create_wheel_bearing_calculator(out)
