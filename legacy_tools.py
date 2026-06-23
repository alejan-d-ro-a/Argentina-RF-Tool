import os
import time
import traceback
import threading
import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import numpy as np
import sys
import subprocess
import unicodedata
import re
import csv
import chardet
import gc
import calendar
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Border, Side
import win32com.client
import shutil

if sys.platform == "win32":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        myappid = 'arg.batch.ept.generator.v1'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except: pass

def set_appwindow(root):
    if sys.platform == "win32":
        try:
            import ctypes
            hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
            if not hwnd: hwnd = root.winfo_id()
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            WS_EX_TOOLWINDOW = 0x00000080
            style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
            style = style & ~WS_EX_TOOLWINDOW
            style = style | WS_EX_APPWINDOW
            ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
            icon_path = str(resource_path("argentina.ico"))
            if os.path.exists(icon_path):
                WM_SETICON = 0x0080
                IMAGE_ICON = 1
                LR_LOADFROMFILE = 0x0010
                hicon = ctypes.windll.user32.LoadImageW(0, icon_path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE)
                if hicon:
                    ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 1, hicon)
                    ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, 0, hicon)
        except: pass

def apply_rounded_corners(window, radius):
    if sys.platform == "win32":
        try:
            import ctypes
            window.update_idletasks()
            if isinstance(window, tk.Toplevel):
                hwnd = window.winfo_id()
            else:
                hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
            if not hwnd: hwnd = window.winfo_id()
            w = window.winfo_width()
            h = window.winfo_height()
            rgn = ctypes.windll.gdi32.CreateRoundRectRgn(0, 0, w, h, radius, radius)
            ctypes.windll.user32.SetWindowRgn(hwnd, rgn, True)
        except: pass

def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

COLOR_BG = "#F0F8FF"
COLOR_FRAME = "#E6F2FF"
COLOR_TEXT = "#003366"
COLOR_BTN = "#00509E"
COLOR_BTN_DARK = "#002244"
COLOR_BTN_HOVER = "#0073E6"
COLOR_BTN_CANCEL = "#A00000"
COLOR_BTN_CANCEL_HOVER = "#D00000"
COLOR_DISABLED_BG = "#D3D3D3"
COLOR_DISABLED_FG = "#808080"

FONT_NORMAL = ("Comic Sans MS", 9)
FONT_TITLE = ("Comic Sans MS", 10, "bold")
FONT_LARGE = ("Comic Sans MS", 11, "bold")
FONT_HUGE = ("Comic Sans MS", 12, "bold")
FONT_DYNAMIC = ("Comic Sans MS", 16, "bold")

class ProcessCancelledException(Exception): pass

def clean_series(series): return series.astype(str).str.replace('\xa0', ' ', regex=False).str.strip().str.upper().str.replace(r'\.0$', '', regex=True)

def build_dict(df, key_col, val_col, sanitize_val=False):
    if not key_col or not val_col or key_col not in df.columns or val_col not in df.columns: return {}
    temp = df.dropna(subset=[key_col, val_col]).copy()
    temp[key_col] = clean_series(temp[key_col])
    invalid_strs = ['', 'NAN', 'NONE', 'NULL', 'NAT']
    temp = temp[~temp[val_col].astype(str).str.upper().str.strip().isin(invalid_strs)]
    if sanitize_val: temp[val_col] = clean_series(temp[val_col])
    else: temp[val_col] = temp[val_col].astype(str).str.replace('\xa0', ' ', regex=False).str.strip().str.upper()
    return temp.drop_duplicates(subset=[key_col]).set_index(key_col)[val_col].to_dict()

def buscar_columna_inteligente(columnas_reales, palabras_clave, reject_words=None):
    if reject_words is None: reject_words = []
    for col in columnas_reales:
        col_norm = ''.join(c for c in unicodedata.normalize('NFD', str(col)) if unicodedata.category(c) != 'Mn')
        col_clean = col_norm.upper().replace(" ", "").replace("_", "").replace("-", "").replace("(", "").replace(")", "")
        if all(kw.upper() in col_clean for kw in palabras_clave):
            if not any(rk.upper() in col_clean for rk in reject_words):
                return col
    return None

find_column_smart = buscar_columna_inteligente

def read_safe_file(filepath):
    if not filepath or not isinstance(filepath, str) or not os.path.exists(filepath): return pd.DataFrame()
    ext = os.path.splitext(filepath)[1].lower()
    if ext in ['.xlsx', '.xls']:
        try: return pd.read_excel(filepath, dtype=str)
        except: return pd.DataFrame()
    else:
        encodings = ['latin1', 'cp1252', 'utf-8', 'utf-8-sig', 'iso-8859-1']
        try:
            with open(filepath, 'rb') as f: raw = f.read(10000)
            detected = chardet.detect(raw)['encoding']
            if detected and detected not in encodings: encodings.insert(0, detected)
        except: pass
        for enc in encodings:
            try: return pd.read_csv(filepath, encoding=enc, low_memory=False, dtype=str)
            except: continue
        return pd.DataFrame()

def read_safe_csv(path):
    return read_safe_file(path)

def flip_vendor(series): return np.where(series == 'HW_MVS', 'HW_TA', np.where(series == 'HW_TA', 'HW_MVS', series))

def normalize_column_name(col):
    if not isinstance(col, str): col = str(col)
    col = unicodedata.normalize('NFKD', col).encode('ASCII', 'ignore').decode('ASCII')
    col = col.strip().upper()
    col = re.sub(r'[^\w]', '', col)
    return col

# =========================================================================
# DATA VALIDATION LOGIC
# =========================================================================
def generate_data_validation(paths_dict, path_batch, out_dir, fmt_csv, fmt_xlsx, update_callback):
    update_callback(5, "Reading Batchfile...")
    df_batch = read_safe_file(path_batch)
    if df_batch.empty:
        raise ValueError("Batchfile is empty or could not be read.")
    if df_batch.shape[1] < 8:
        raise ValueError("The Batchfile must have at least 8 columns to extract UNICO (col 1) and Cluster (col 8).")

    unicos_series = clean_series(df_batch.iloc[:, 0])
    clusters_series = clean_series(df_batch.iloc[:, 7])

    df_map = pd.DataFrame({'UNICO': unicos_series, 'CLUSTER': clusters_series})
    df_map = df_map[(df_map['UNICO'] != '') & (df_map['UNICO'] != 'NAN')]
    dict_unico_cluster = df_map.drop_duplicates(subset=['UNICO']).set_index('UNICO')['CLUSTER'].to_dict()
    
    del df_batch, df_map
    gc.collect()

    all_results = []
    total_techs = sum(1 for p in paths_dict.values() if p and os.path.exists(p))
    if total_techs == 0: raise ValueError("No valid data files provided for validation.")

    progress_step = 80 / total_techs
    current_progress = 10

    def extract_unico_from_cell(cell_str):
        s = str(cell_str).strip().upper()
        l = len(s)
        if l in [7, 8]: return s[1:6]
        if l in [9, 14]: return s[:6]
        return s 

    for tech, path in paths_dict.items():
        if not path or not os.path.exists(path): continue
        update_callback(current_progress, f"Processing {tech} Data...")
        df_data = read_safe_file(path)
        if df_data.empty:
            current_progress += progress_step
            continue

        c_cell = buscar_columna_inteligente(df_data.columns, ['CELL']) or 'cell'
        c_date = buscar_columna_inteligente(df_data.columns, ['FECHA', 'DATE', 'TIME']) or 'fecha'

        if c_cell not in df_data.columns or c_date not in df_data.columns:
            current_progress += progress_step
            continue

        def parse_dt(v):
            if pd.isna(v): return pd.NaT
            s = str(v).strip()
            try:
                if ' ' in s: s = s.split()[0]
                for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"):
                    try: return datetime.strptime(s, fmt).date()
                    except: pass
                return pd.to_datetime(s, format='mixed', dayfirst=False).date()
            except: return pd.NaT

        df_data['ParsedDate'] = df_data[c_date].apply(parse_dt)
        df_data = df_data.dropna(subset=['ParsedDate'])
        if df_data.empty:
            current_progress += progress_step
            continue

        df_data['CELL_CLEAN'] = df_data[c_cell].astype(str).str.strip().str.upper()
        df_data['UNICO'] = df_data['CELL_CLEAN'].apply(extract_unico_from_cell)
        df_data['Cluster'] = df_data['UNICO'].map(dict_unico_cluster).fillna("SIN CLUSTER")
        df_data['Cluster'] = df_data['Cluster'].astype(str).str.upper()
        df_data['UNICO'] = df_data['UNICO'].astype(str).str.upper()

        agg_df = df_data.groupby(['Cluster', 'UNICO'])['ParsedDate'].agg(Fecha_Minima='min', Fecha_Maxima='max').reset_index()
        agg_df['Tecnología'] = tech
        all_results.append(agg_df)

        current_progress += progress_step
        del df_data
        gc.collect()

    if not all_results: raise ValueError("No valid data could be extracted from the provided files.")
    
    update_callback(95, "Exporting Consolidated Data...")
    final_df = pd.concat(all_results, ignore_index=True)
    final_df.sort_values(by=['Cluster', 'UNICO', 'Tecnología'], inplace=True)
    
    out_file_base = os.path.join(out_dir, "DATA_VALIDATION_REPORT")
    try:
        if fmt_csv: final_df.to_csv(out_file_base + ".csv", index=False, encoding='utf-8-sig', sep=',')
        if fmt_xlsx: final_df.to_excel(out_file_base + ".xlsx", index=False, sheet_name="DATA_VALIDATION")
    except PermissionError as e: 
        raise PermissionError("Permission denied. Ensure the file is not open in another program.\n\n" + str(e))
        
    update_callback(100, "Validation completed.")
    return out_file_base + (".csv" if fmt_csv else ".xlsx")

# =========================================================================
# NEW LOGIC: IT FINAL REPORT (USING WIN32COM PARA PRESERVAR IMÁGENES Y FORMATO)
# =========================================================================
def generate_it_final_report(p_rnd, p_template_unused, p_ctrl, p_cambios, date_elec, date_mech, out_dir, fmt_csv, fmt_xlsx, update_callback):
    import shutil
    import win32com.client
    
    p_template = resource_path("Cluster Final Report_PRE_vs_POST TEMPLATE.xlsx")
    if not os.path.exists(p_template):
        raise FileNotFoundError(f"Template internal file not found: {p_template}")

    def clean_val(v):
        if pd.isna(v) or str(v).strip().upper() == 'NAN': 
            return ""
        return v

    update_callback(10, "Extracting Cluster Name from RND...")
    df_rnd = pd.concat([pd.read_excel(p_rnd, sheet_name=s) for s in pd.ExcelFile(p_rnd).sheet_names if 'MVS' in str(s).upper() or 'TP' in str(s).upper()])
    c_cluster_rnd = buscar_columna_inteligente(df_rnd.columns, ['CLUSTER'])
    if not c_cluster_rnd: raise ValueError("Cluster column not found in RND.")
    cluster_name = str(df_rnd[c_cluster_rnd].dropna().iloc[0]).strip().upper()

    update_callback(20, "Preparing Template environment...")
    final_out = os.path.join(out_dir, f"Final_Report_{cluster_name}.xlsx")
    final_out_abs = os.path.abspath(final_out)
    
    shutil.copy(p_template, final_out_abs)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    
    try:
        wb = excel.Workbooks.Open(final_out_abs)

        update_callback(25, f"Updating Template for {cluster_name}...")
        for ws in wb.Sheets:
            ws.Cells.Replace(What="XXXX_XXX", Replacement=cluster_name, LookAt=2, MatchCase=False)

        update_callback(40, "Processing Site List (Cluster Control)...")
        df_ctrl = read_safe_file(p_ctrl)
        c_clust_id = buscar_columna_inteligente(df_ctrl.columns, ['CLUST', 'ID']) or 'Clust_Id'
        c_sitio = buscar_columna_inteligente(df_ctrl.columns, ['SITIO'])
        c_sitio_mvs = buscar_columna_inteligente(df_ctrl.columns, ['SITIO', 'MVS'])
        c_site_name = buscar_columna_inteligente(df_ctrl.columns, ['SITE', 'NAME'])
        c_operador = buscar_columna_inteligente(df_ctrl.columns, ['OPERADOR'])
        c_rsh = buscar_columna_inteligente(df_ctrl.columns, ['CLUSTER', 'ACTIVATION']) or buscar_columna_inteligente(df_ctrl.columns, ['ACTIVATION'])
        c_bbu = buscar_columna_inteligente(df_ctrl.columns, ['TARGET', 'BBU'])

        df_ctrl_clust = df_ctrl[df_ctrl[c_clust_id].astype(str).str.strip().str.upper() == cluster_name].copy()
        if c_rsh:
            df_ctrl_clust = df_ctrl_clust[~df_ctrl_clust[c_rsh].astype(str).str.upper().str.contains('BLOCKED', na=False)]

        site_list_data = []
        for _, row in df_ctrl_clust.iterrows():
            site_list_data.append(tuple(clean_val(x) for x in [
                row.get(c_clust_id, cluster_name), 
                row.get(c_sitio_mvs, ''), 
                row.get(c_sitio, ''), 
                row.get(c_bbu, ''),
                row.get(c_site_name, ''), 
                row.get(c_operador, ''), 
                row.get(c_rsh, '')
            ]))

        ws_info = None
        for sheet in wb.Sheets:
            if sheet.Name == "Cluster Information":
                ws_info = sheet
                break
        if not ws_info: ws_info = wb.Sheets(1)

        found_site = ws_info.Cells.Find(What="Cluster Name", LookAt=2, MatchCase=False)
        if not found_site: 
            raise ValueError("Marker 'Cluster Name' not found in template.")
        
        if found_site and site_list_data:
            start_row = found_site.Row + 1
            col_start = found_site.Column
            num_rows = len(site_list_data)
            
            ws_info.Rows(f"{start_row}:{start_row + num_rows - 1}").Insert()
            
            end_row = start_row + num_rows - 1
            end_col = col_start + len(site_list_data[0]) - 1
            rng = ws_info.Range(ws_info.Cells(start_row, col_start), ws_info.Cells(end_row, end_col))
            rng.Value = site_list_data
            
            rng.Font.Bold = False
            rng.Font.Name = "Calibri"
            rng.Font.Size = 11
            rng.HorizontalAlignment = -4108 
            rng.VerticalAlignment = -4108   
            rng.Interior.ColorIndex = -4142 
            
            for b_id in range(7, 13):
                rng.Borders(b_id).LineStyle = 1
                rng.Borders(b_id).Weight = 2
                
            rng.Columns(7).NumberFormat = "dd/mm/yyyy"

        update_callback(55, "Processing Cluster Information (RND)...")
        c_rat = buscar_columna_inteligente(df_rnd.columns, ['RAT']) or buscar_columna_inteligente(df_rnd.columns, ['TECH'])
        c_type = buscar_columna_inteligente(df_rnd.columns, ['CELL', 'TYPE'], ['LOCAL']) or buscar_columna_inteligente(df_rnd.columns, ['DUPLEX'])
        
        if c_rat: df_rnd = df_rnd[~df_rnd[c_rat].astype(str).str.upper().str.contains('5G|NR', na=False)]
        if c_type: df_rnd = df_rnd[~df_rnd[c_type].astype(str).str.upper().str.contains('NB|IOT', na=False)]

        c_site_id = df_rnd.columns[0] if len(df_rnd.columns) > 0 else None
        c_site_name_rnd = df_rnd.columns[1] if len(df_rnd.columns) > 1 else None
        c_cell = buscar_columna_inteligente(df_rnd.columns, ['CELL', 'NAME'], ['LOCAL', 'PHYSICAL']) or buscar_columna_inteligente(df_rnd.columns, ['CELDA']) or buscar_columna_inteligente(df_rnd.columns, ['CELLNAME'])
        c_sec = buscar_columna_inteligente(df_rnd.columns, ['SECTOR'])
        c_cid = buscar_columna_inteligente(df_rnd.columns, ['CELLID'], ['LOCAL']) or buscar_columna_inteligente(df_rnd.columns, ['CELL', 'ID'], ['LOCAL', 'INDEX'])
        c_local_cid = buscar_columna_inteligente(df_rnd.columns, ['RAT', 'NE', 'ID'])
        c_arfcn = buscar_columna_inteligente(df_rnd.columns, ['ARFCN', 'DL']) or buscar_columna_inteligente(df_rnd.columns, ['EARFCN']) or buscar_columna_inteligente(df_rnd.columns, ['UARFCN'])
        c_pci = buscar_columna_inteligente(df_rnd.columns, ['PCI']) or buscar_columna_inteligente(df_rnd.columns, ['PSC'])
        c_azi = buscar_columna_inteligente(df_rnd.columns, ['AZIMUTH']) or buscar_columna_inteligente(df_rnd.columns, ['AZIMUT']) or buscar_columna_inteligente(df_rnd.columns, ['AZI'])
        c_etilt = buscar_columna_inteligente(df_rnd.columns, ['ELECTRICAL']) or buscar_columna_inteligente(df_rnd.columns, ['RET', 'TILT']) or buscar_columna_inteligente(df_rnd.columns, ['ETILT'])
        c_mtilt = buscar_columna_inteligente(df_rnd.columns, ['MECHANICAL']) or buscar_columna_inteligente(df_rnd.columns, ['MTILT'])
        c_height = buscar_columna_inteligente(df_rnd.columns, ['HEIGHT']) or buscar_columna_inteligente(df_rnd.columns, ['ALTURA'])
        c_ant = buscar_columna_inteligente(df_rnd.columns, ['ANTENNA']) or buscar_columna_inteligente(df_rnd.columns, ['ANTENA'])

        cluster_info_data = []
        for _, row in df_rnd.iterrows():
            c_val = str(row.get(c_cell, '')).strip()
            sec_val = str(row.get(c_sec, '')).strip()
            if not sec_val and c_val: sec_val = c_val[-1]
            cluster_info_data.append(tuple(clean_val(x) for x in [
                row.get(c_rat, ''), 
                row.get(c_site_id, ''), 
                row.get(c_site_name_rnd, ''),
                c_val, 
                sec_val, 
                row.get(c_cid, ''), 
                row.get(c_local_cid, ''),
                row.get(c_arfcn, ''), 
                row.get(c_pci, ''), 
                row.get(c_azi, ''),
                row.get(c_etilt, ''), 
                row.get(c_mtilt, ''), 
                row.get(c_height, ''), 
                row.get(c_ant, '')
            ]))

        found_tech = ws_info.Cells.Find(What="Tech", LookAt=2, MatchCase=False)
        if not found_tech: 
            raise ValueError("Marker 'Tech' not found in template.")
                
        if found_tech and cluster_info_data:
            start_row = found_tech.Row + 1
            col_start = found_tech.Column
            num_rows = len(cluster_info_data)
            
            ws_info.Rows(f"{start_row}:{start_row + num_rows - 1}").Insert()
                
            end_row = start_row + num_rows - 1
            end_col = col_start + len(cluster_info_data[0]) - 1
            rng = ws_info.Range(ws_info.Cells(start_row, col_start), ws_info.Cells(end_row, end_col))
            rng.Value = cluster_info_data
            
            rng.Font.Bold = False
            rng.Font.Name = "Calibri"
            rng.Font.Size = 11
            rng.HorizontalAlignment = -4108
            rng.VerticalAlignment = -4108
            rng.Interior.ColorIndex = -4142
            
            for b_id in range(7, 13):
                rng.Borders(b_id).LineStyle = 1
                rng.Borders(b_id).Weight = 2

        ws_info.Columns.AutoFit()

        update_callback(75, "Processing Optimization Changes...")
        xl_cambios = pd.ExcelFile(p_cambios)
        sheet_cambios = None
        for s in xl_cambios.sheet_names:
            if 'CAMBIOS' in s.upper() or 'PROPOSAL' in s.upper():
                sheet_cambios = s
                break
        if not sheet_cambios:
            sheet_cambios = xl_cambios.sheet_names[1]
            
        df_raw = pd.read_excel(p_cambios, sheet_name=sheet_cambios, header=None)
        header_idx = 0
        for i, row in df_raw.head(15).iterrows():
            row_str = " ".join(row.dropna().astype(str)).upper()
            if 'CELL' in row_str and 'NAME' in row_str:
                header_idx = i
                break
                
        df_cambios = pd.read_excel(p_cambios, sheet_name=sheet_cambios, header=header_idx)

        c_cambio_codigo = buscar_columna_inteligente(df_cambios.columns, ['CODIGO', 'UNICO'])
        c_cambio_site = buscar_columna_inteligente(df_cambios.columns, ['SITE', 'NAME'])
        c_cambio_cell = buscar_columna_inteligente(df_cambios.columns, ['CELL', 'NAME'])
        c_cambio_param = buscar_columna_inteligente(df_cambios.columns, ['TUNING', 'PARAMETER']) or buscar_columna_inteligente(df_cambios.columns, ['PARAMETER'])
        c_cambio_target = buscar_columna_inteligente(df_cambios.columns, ['OPTIMIZATION', 'TARGET'])
        
        c_pre_azi = buscar_columna_inteligente(df_cambios.columns, ['PRE', 'AZIMUTH']) or buscar_columna_inteligente(df_cambios.columns, ['PRE', 'AZI'])
        c_prop_azi = buscar_columna_inteligente(df_cambios.columns, ['PROPOSED', 'AZIMUTH']) or buscar_columna_inteligente(df_cambios.columns, ['PROPOSED', 'AZI'])
        
        c_pre_etilt = buscar_columna_inteligente(df_cambios.columns, ['PRE', 'E-TILT']) or buscar_columna_inteligente(df_cambios.columns, ['PRE', 'TILT'], reject_words=['M-TILT', 'MTILT', 'MECHANICAL'])
        c_prop_etilt = buscar_columna_inteligente(df_cambios.columns, ['PROPOSED', 'E-TILT']) or buscar_columna_inteligente(df_cambios.columns, ['PROPOSED', 'TILT'], reject_words=['M-TILT', 'MTILT', 'MECHANICAL'])
        
        c_pre_mtilt = buscar_columna_inteligente(df_cambios.columns, ['PRE', 'M-TILT']) or buscar_columna_inteligente(df_cambios.columns, ['PRE', 'TILT'], reject_words=['E-TILT', 'ETILT', 'ELECTRICAL'])
        c_prop_mtilt = buscar_columna_inteligente(df_cambios.columns, ['PROPOSED', 'M-TILT']) or buscar_columna_inteligente(df_cambios.columns, ['PROPOSED', 'TILT'], reject_words=['E-TILT', 'ETILT', 'ELECTRICAL'])

        cambios_data = []
        for _, row in df_cambios.iterrows():
            param_str = str(row.get(c_cambio_param, '')).strip().upper()
            if param_str == 'NAN' or param_str == '': 
                continue

            val_unico = str(row.get(c_cambio_codigo, '')).strip() if c_cambio_codigo else ""
            val_site = str(row.get(c_cambio_site, '')).strip() if c_cambio_site else ""
            val_cell = str(row.get(c_cambio_cell, '')).strip() if c_cambio_cell else ""
            
            if val_unico.lower() == 'nan': val_unico = ""
            if val_site.lower() == 'nan': val_site = ""
            if val_cell.lower() == 'nan': val_cell = ""

            site_prefix = val_unico if val_unico else val_site
            site_cell = f"{site_prefix} / {val_cell}" if site_prefix and val_cell else (site_prefix or val_cell)

            target = row.get(c_cambio_target, '')
            if str(target).strip().upper() == 'NAN': 
                target = ''

            if 'AZIMUT' in param_str:
                date_to_use = date_mech if date_mech else date_elec
                cambios_data.append(tuple(clean_val(x) for x in [site_cell, "Azimuth", row.get(c_pre_azi, ''), row.get(c_prop_azi, ''), "Done", date_to_use, target]))
            elif 'E-TILT' in param_str or 'ETILT' in param_str or 'ELECT' in param_str:
                date_to_use = date_elec if date_elec else date_mech
                cambios_data.append(tuple(clean_val(x) for x in [site_cell, "E-Tilt", row.get(c_pre_etilt, ''), row.get(c_prop_etilt, ''), "Done", date_to_use, target]))
            elif 'M-TILT' in param_str or 'MTILT' in param_str or 'MECH' in param_str:
                date_to_use = date_mech if date_mech else date_elec
                cambios_data.append(tuple(clean_val(x) for x in [site_cell, "M-Tilt", row.get(c_pre_mtilt, ''), row.get(c_prop_mtilt, ''), "Done", date_to_use, target]))

        ws_opt = None
        for sheet in wb.Sheets:
            if sheet.Name == "Optimization Changes":
                ws_opt = sheet
                break

        if ws_opt and cambios_data:
            found_opt = ws_opt.Cells.Find(What="Site / Cellname", LookAt=2, MatchCase=False)
            if not found_opt: 
                raise ValueError("Marker 'Site / Cellname' not found in template.")
            
            if found_opt:
                start_row = found_opt.Row + 1
                col_start = found_opt.Column
                num_rows = len(cambios_data)
                
                ws_opt.Rows(f"{start_row}:{start_row + num_rows - 1}").Insert()
                    
                end_row = start_row + num_rows - 1
                end_col = col_start + len(cambios_data[0]) - 1
                rng = ws_opt.Range(ws_opt.Cells(start_row, col_start), ws_opt.Cells(end_row, end_col))
                rng.Value = cambios_data
                
                rng.Font.Bold = False
                rng.Font.Name = "Calibri"
                rng.Font.Size = 11
                rng.HorizontalAlignment = -4108
                rng.VerticalAlignment = -4108
                rng.Interior.ColorIndex = -4142
                
                for b_id in range(7, 13):
                    rng.Borders(b_id).LineStyle = 1
                    rng.Borders(b_id).Weight = 2
                    
                rng.Columns(6).NumberFormat = "dd/mm/yyyy"

            ws_opt.Columns.AutoFit()

        update_callback(95, "Saving Final Report...")
        wb.Save()
        
    finally:
        if wb:
            wb.Close(False)
        excel.Quit()
        del excel

    update_callback(100, "Generation completed.")
    return final_out

# =========================================================================
# LÓGICA DE MASTER EPT
# =========================================================================
def generate_ept(path_mvs, path_tp, path_nics, out_dir, fmt_csv, fmt_xlsx, update_callback):
    update_callback(5.12, "Waiting for instructions...")
    df_mvs, df_tp = read_safe_csv(path_mvs), read_safe_csv(path_tp)
    if df_mvs.empty or df_tp.empty: raise ValueError("MVS, TP, and NICS files are mandatory for MASTER EPT.")
    df_input = pd.concat([df_mvs, df_tp], ignore_index=True)
    del df_mvs
    del df_tp
    gc.collect()
    
    update_callback(18.45, "Applying RF rules...")
    c_vendor   = buscar_columna_inteligente(df_input.columns, ['VENDOR']) or 'Vendor'
    c_grilla   = buscar_columna_inteligente(df_input.columns, ['GRILLA']) or buscar_columna_inteligente(df_input.columns, ['GU']) or 'Grilla Única'
    c_outdoor  = buscar_columna_inteligente(df_input.columns, ['OUTDOOR']) or buscar_columna_inteligente(df_input.columns, ['ISOUT']) or 'isOutdoor'
    c_cellname = buscar_columna_inteligente(df_input.columns, ['CELL', 'NAME'], reject_words=['LOCAL']) or buscar_columna_inteligente(df_input.columns, ['CELDA']) or 'Cell Name'
    c_operator = buscar_columna_inteligente(df_input.columns, ['OPERATOR']) or buscar_columna_inteligente(df_input.columns, ['OPERADOR']) or 'Operator'
    c_band     = buscar_columna_inteligente(df_input.columns, ['BAND']) or buscar_columna_inteligente(df_input.columns, ['FRECUENCIA']) or 'Frequency Band'
    c_pci_in   = buscar_columna_inteligente(df_input.columns, ['PCI']) or buscar_columna_inteligente(df_input.columns, ['PHYSICAL', 'CELL']) or 'PCI'
    c_local_id = buscar_columna_inteligente(df_input.columns, ['LOCAL', 'CELL']) or buscar_columna_inteligente(df_input.columns, ['LOCAL', 'ID']) or 'Local Cell ID'
    c_enb_name = buscar_columna_inteligente(df_input.columns, ['ENODEB', 'NAME'], reject_words=['ID']) or 'eNodeB Name'
    c_tac      = buscar_columna_inteligente(df_input.columns, ['TAC']) or 'TAC'
    c_enb_id   = buscar_columna_inteligente(df_input.columns, ['ENODEBID']) or buscar_columna_inteligente(df_input.columns, ['NODEBID']) or 'eNodeB ID'
    
    optional_cols = ['Site OSS Name', 'CellID', 'SwapDate', 'Status']
    for col_name in optional_cols:
        if col_name not in df_input.columns:
            df_input[col_name] = ""
            
    req_cols = [c_vendor, c_grilla, c_cellname, c_band, c_pci_in]
    missing = [c for c in req_cols if c not in df_input.columns]
    if missing: raise KeyError(f"Missing required columns in the provided files:\n{', '.join(missing)}")
    
    c_cell_id  = buscar_columna_inteligente(df_input.columns, ['CELLID'], reject_words=['LOCAL']) or 'Cell ID'
    c_long     = buscar_columna_inteligente(df_input.columns, ['LONGITUDE']) or buscar_columna_inteligente(df_input.columns, ['LONGITUD']) or 'Longitude'
    c_lat      = buscar_columna_inteligente(df_input.columns, ['LATITUDE']) or buscar_columna_inteligente(df_input.columns, ['LATITUD']) or 'Latitude'
    c_azimuth  = buscar_columna_inteligente(df_input.columns, ['AZIMUTH']) or buscar_columna_inteligente(df_input.columns, ['AZIMUT']) or 'Azimuth'
    c_height   = buscar_columna_inteligente(df_input.columns, ['HEIGHT']) or buscar_columna_inteligente(df_input.columns, ['ALTURA']) or 'Height'
    c_mech     = buscar_columna_inteligente(df_input.columns, ['MECHANICAL']) or buscar_columna_inteligente(df_input.columns, ['MTILT']) or 'Mechanical Downtilt'
    c_ant_mod  = buscar_columna_inteligente(df_input.columns, ['ANTENNA']) or buscar_columna_inteligente(df_input.columns, ['ANTENA']) or 'Antenna Model'
    c_earfcn   = buscar_columna_inteligente(df_input.columns, ['EARFCN']) or 'DlEarfcn'
    c_bwidth   = buscar_columna_inteligente(df_input.columns, ['BANDWIDTH']) or buscar_columna_inteligente(df_input.columns, ['ANCHO']) or 'DlBandwidth'
    c_cluster_in = buscar_columna_inteligente(df_input.columns, ['CLUSTER']) or buscar_columna_inteligente(df_input.columns, ['CLUST', 'ID'])
    c_etilt_in = buscar_columna_inteligente(df_input.columns, ['ETILT']) or buscar_columna_inteligente(df_input.columns, ['ELECTRICAL']) or 'Electrical Downtilt'
    
    v_vendor = df_input[c_vendor].fillna('').astype(str).str.upper().str.strip()
    v_gu     = df_input[c_grilla].fillna('').astype(str).str.upper().str.strip()
    v_out    = df_input[c_outdoor].fillna('').astype(str).str.lower().str.strip()
    v_cell   = df_input[c_cellname].fillna('').astype(str).str.strip()
    cond = ((v_vendor == 'HUAWEI') & (v_gu == 'GU_OK') & ((v_out == 'outdoor') | (v_out == '')) & (v_cell != '') & (v_cell.str.len().isin([7, 8, 9, 14])))
    df_base = df_input[cond].copy()
    if df_base.empty: raise ValueError("No sectors passed the initial EPT filter.")
    del df_input
    gc.collect()
    
    update_callback(32.80, "Applying RF rules...")
    df_base['Operator_Clean'] = df_base[c_operator].fillna('').astype(str).str.strip()
    df_base['Frequency_Band_Clean'] = pd.to_numeric(df_base[c_band], errors='coerce').fillna(0)
    df_base['PCI_INPUT_Clean'] = pd.to_numeric(df_base[c_pci_in], errors='coerce').fillna(0)
    df_base['Local_Cell_ID_Clean'] = pd.to_numeric(df_base[c_local_id], errors='coerce').fillna(0).astype(int).astype(str)
    df_base['eNodeB_Name_Raw_Clean'] = df_base.get(c_enb_name, '').fillna('').astype(str).str.strip().str.upper()
    df_base['COMPOSITE_KEY'] = df_base['eNodeB_Name_Raw_Clean'] + "_" + df_base['Local_Cell_ID_Clean']

    cell_len = df_base[c_cellname].str.len()
    cond_7_8 = cell_len.isin([7, 8])
    cond_9_14 = cell_len.isin([9, 14])
    df_base['eNodeB Name'] = np.select(
        [cond_7_8, cond_9_14],
        [df_base[c_cellname].str[1:6], df_base[c_cellname].str[0:6]],
        default="REVIEW_LENGTH"
    )

    if c_cluster_in is not None:
        df_base['CLUSTER'] = df_base[c_cluster_in].replace('', np.nan).fillna("NO_CLUSTER_FOUND")
    else:
        df_base['CLUSTER'] = "NO_CLUSTER_FOUND"
    
    update_callback(55.25, "Applying RF rules...")
    if path_nics and os.path.exists(path_nics):
        f_lst = os.path.join(path_nics, "LST_CELL.csv")
        f_pb = os.path.join(path_nics, "LST_CELL_PB.csv")
        f_pa = os.path.join(path_nics, "LST_CELL_PA.csv")
        df_lst = read_safe_csv(f_lst)
        if not df_lst.empty:
            c_lst_cell = buscar_columna_inteligente(df_lst.columns, ['CELL', 'NAME'], reject_words=['LOCAL']) or 'Cell Name'
            c_lst_pci  = buscar_columna_inteligente(df_lst.columns, ['PHYSICAL', 'CELL']) or buscar_columna_inteligente(df_lst.columns, ['PCI']) or 'Physical cell ID'
            c_lst_mode = buscar_columna_inteligente(df_lst.columns, ['TRANSMISSION', 'RECEPTION']) or buscar_columna_inteligente(df_lst.columns, ['MODE']) or 'Cell transmission and reception mode'
            df_base = df_base.merge(df_lst[[c_lst_cell, c_lst_pci, c_lst_mode]].drop_duplicates(subset=[c_lst_cell]), how='left', left_on=c_cellname, right_on=c_lst_cell)
            df_base['PCI'] = df_base[c_lst_pci].fillna(df_base['PCI_INPUT_Clean'])
            
            mode_str = df_base.get(c_lst_mode, '').astype(str).str.upper().str.strip()
            df_base['Number of Transmission Antenna Ports'] = mode_str.str.extract(r'(\d+)T')[0].fillna(4).astype(int).clip(upper=4)
            df_base['Number of Transmission Antennas'] = mode_str.str.extract(r'T(\d+)R')[0].fillna(4).astype(int)
        else:
            df_base['PCI'], df_base['Number of Transmission Antenna Ports'], df_base['Number of Transmission Antennas'] = df_base['PCI_INPUT_Clean'], 4, 4
        del df_lst
        gc.collect()

        update_callback(72.90, "Applying RF rules...")
        df_pb = read_safe_csv(f_pb)
        if not df_pb.empty:
            c_pb_ne, c_pb_loc = buscar_columna_inteligente(df_pb.columns, ['NENAME']) or 'NENAME', buscar_columna_inteligente(df_pb.columns, ['LOCAL', 'CELL']) or 'Local Cell ID'
            c_pb_rsp, c_pb_val = buscar_columna_inteligente(df_pb.columns, ['REFERENCE', 'POWER']) or 'Reference signal power(0.1dBm)', buscar_columna_inteligente(df_pb.columns, ['PB']) or 'PB'
            df_pb['COMPOSITE_KEY'] = df_pb[c_pb_ne].fillna('').astype(str).str.strip().str.upper() + "_" + pd.to_numeric(df_pb[c_pb_loc], errors='coerce').fillna(0).astype(int).astype(str)
            df_base = df_base.merge(df_pb.drop_duplicates(subset=['COMPOSITE_KEY'])[['COMPOSITE_KEY', c_pb_rsp, c_pb_val]], how='left', on='COMPOSITE_KEY')
            
            rs_numeric = pd.to_numeric(df_base[c_pb_rsp], errors='coerce')
            mode_upper = df_base.get('Cell transmission and reception mode', '').astype(str).str.upper()
            
            cond_32 = mode_upper.str.contains('32T32R')
            cond_band = df_base['Frequency_Band_Clean'].isin([28, 5])
            
            df_base['RS Power'] = np.select(
                [rs_numeric.notna() & (rs_numeric != 0), cond_32, cond_band],
                [rs_numeric, 113, 175],
                default=170
            )
            df_base['PB'] = pd.to_numeric(df_base.get(c_pb_val), errors='coerce').fillna(1)
        else:
            df_base['RS Power'], df_base['PB'] = 170, 1
        del df_pb

        df_pa = read_safe_csv(f_pa)
        if not df_pa.empty:
            c_pa_ne, c_pa_loc = buscar_columna_inteligente(df_pa.columns, ['NENAME']) or 'NENAME', buscar_columna_inteligente(df_pa.columns, ['LOCAL', 'CELL']) or 'Local Cell ID'
            c_pa_val = buscar_columna_inteligente(df_pa.columns, ['PA', 'DISTRIBUTION']) or 'PA for even power distribution(dB)'
            df_pa['COMPOSITE_KEY'] = df_pa[c_pa_ne].fillna('').astype(str).str.strip().str.upper() + "_" + pd.to_numeric(df_pa[c_pa_loc], errors='coerce').fillna(0).astype(int).astype(str)
            df_base = df_base.merge(df_pa.drop_duplicates(subset=['COMPOSITE_KEY'])[['COMPOSITE_KEY', c_pa_val]], how='left', on='COMPOSITE_KEY')
            df_base['PA'] = pd.to_numeric(df_base[c_pa_val].astype(str).str.replace('dB', '', case=False).str.strip(), errors='coerce').fillna(-3)
        else:
            df_base['PA'] = -3
        del df_pa
        gc.collect()
    else:
        df_base['PCI'], df_base['Number of Transmission Antenna Ports'], df_base['Number of Transmission Antennas'] = df_base['PCI_INPUT_Clean'], 4, 4
        df_base['RS Power'], df_base['PB'], df_base['PA'] = 170, 1, -3
        
    update_callback(88.35, "Exporting EPT...")
    df_base['Electrical Downtilt'] = pd.to_numeric(df_base.get(c_etilt_in, 0), errors='coerce').fillna(0.0)
    df_base['TAC'], df_base['eNodeB ID'] = df_base.get(c_tac, ""), df_base.get(c_enb_id, "")
    df_base['Cell ID'], df_base['Cell Name'], df_base['Active'] = df_base.get(c_cell_id, ""), df_base.get(c_cellname, ""), 'Y'
    df_base['Longitude'] = pd.to_numeric(df_base.get(c_long, np.nan), errors='coerce').replace(0, np.nan)
    df_base['Latitude'] = pd.to_numeric(df_base.get(c_lat, np.nan), errors='coerce').replace(0, np.nan)
    site_clean = df_base['eNodeB Name'].astype(str).str.strip().str.upper()
    df_base['Longitude'] = df_base.groupby(site_clean, dropna=False)['Longitude'].ffill().bfill().round(6)
    df_base['Latitude'] = df_base.groupby(site_clean, dropna=False)['Latitude'].ffill().bfill().round(6)
    df_base['Azimuth'], df_base['Height'] = df_base.get(c_azimuth, ""), df_base.get(c_height, "")
    df_base['Mechanical Downtilt'] = pd.to_numeric(df_base.get(c_mech, 0), errors='coerce').fillna(0)
    df_base['Antenna Model'], df_base['DlEarfcn'] = df_base.get(c_ant_mod, ""), df_base.get(c_earfcn, "")
    df_base['DlBandwidth'] = pd.to_numeric(df_base.get(c_bwidth, np.nan), errors='coerce')
    df_base['DlBandwidth'] = np.where(df_base['DlBandwidth'] > 100, df_base['DlBandwidth'] / 1000, df_base['DlBandwidth'])
    df_base = df_base[df_base['DlBandwidth'].round(3) != 0.2].reset_index(drop=True)
    df_base['PDSCH Actual Load(DL)'], df_base['isOutdoor'] = 0.5, 1
    raw_op = df_base['Operator_Clean'].astype(str).str.replace(r'\.0$', '', regex=True).str.replace(r'[^0-9]', '', regex=True)
    df_base['MCC'] = raw_op.str[:3]
    df_base['MNC'] = raw_op.str[3:].str.zfill(2)
    final_cols = ['TAC', 'eNodeB ID', 'Cell ID', 'eNodeB Name', 'CLUSTER', 'Cell Name', 'Active', 'Longitude', 'Latitude', 'Azimuth', 'Height', 'Mechanical Downtilt', 'Electrical Downtilt', 'Antenna Model', 'PCI', 'DlEarfcn', 'DlBandwidth', 'RS Power', 'PA', 'PB', 'Number of Transmission Antenna Ports', 'Number of Transmission Antennas', 'PDSCH Actual Load(DL)', 'isOutdoor', 'MCC', 'MNC']
    for col in final_cols:
        if col not in df_base.columns: df_base[col] = ""
    df_final = df_base[final_cols].drop_duplicates()
    
    update_callback(96.72, "Exporting EPT...")
    out_file_base = os.path.join(out_dir, "MASTER_EPT")
    try:
        if fmt_csv: df_final.to_csv(out_file_base + ".csv", index=False, encoding='latin1', sep=',')
        if fmt_xlsx: df_final.to_excel(out_file_base + ".xlsx", index=False)
    except PermissionError as e:
        raise PermissionError("Permission denied. The file might be open in another program. Please close it and try again.\n\n" + f"({str(e)})")
    return out_file_base + (".csv" if fmt_csv else ".xlsx")

# =========================================================================
# CLUSTER FRIENDLY Y ITSC
# =========================================================================
def generate_cluster_friendly(path_ctrl, path_clus, out_dir, fmt_csv, fmt_xlsx, update_callback):
    update_callback(5, "Waiting for instructions...")
    if not path_ctrl or not os.path.exists(path_ctrl):
        raise ValueError("Cluster Control file not found.")
    ext_ctrl = os.path.splitext(path_ctrl)[1].lower()
    if ext_ctrl in ['.xlsx', '.xls']:
        try:
            xl = pd.ExcelFile(path_ctrl)
            sheet_name = 'Site List' if 'Site List' in xl.sheet_names else xl.sheet_names[0]
            df_ctrl = pd.read_excel(path_ctrl, sheet_name=sheet_name, header=1, dtype=str)
        except Exception as e:
            raise ValueError(f"Error reading Cluster Control Excel: {e}")
    else:
        df_raw = read_safe_csv(path_ctrl)
        if df_raw.empty:
            raise ValueError("Cluster Control file is empty or could not be read.")
        new_headers = df_raw.iloc[1].astype(str).values
        df_ctrl = df_raw.iloc[2:].copy()
        df_ctrl.columns = new_headers
        df_ctrl.reset_index(drop=True, inplace=True)
    if df_ctrl.empty:
        raise ValueError("Cluster Control file has no data after skipping header row.")
    update_callback(20, "Reading Cluster Control...")
    if not path_clus or not os.path.exists(path_clus):
        raise ValueError("Cluster mapping file not found.")
    ext_clus = os.path.splitext(path_clus)[1].lower()
    if ext_clus in ['.xlsx', '.xls']:
        try:
            xl_clus = pd.ExcelFile(path_clus)
            if 'Sitios GU X Clúster' in xl_clus.sheet_names:
                df_clus = pd.read_excel(path_clus, sheet_name='Sitios GU X Clúster', dtype=str)
            else:
                df_clus = pd.read_excel(path_clus, sheet_name=xl_clus.sheet_names[0], dtype=str)
        except Exception as e:
            raise ValueError(f"Error reading Cluster mapping Excel: {e}")
    else:
        df_clus = read_safe_csv(path_clus)
    if df_clus.empty:
        raise ValueError("Cluster mapping file is empty or could not be read.")
    update_callback(40, "Reading Cluster mapping...")
    col_sitio = None
    for col in df_ctrl.columns:
        if col is None: continue
        norm = normalize_column_name(str(col))
        if norm in ['SITIO', 'SITE']:
            col_sitio = col
            break
    if col_sitio is None:
        raise KeyError("Column not found: 'Sitio' in Cluster Control.")
    col_site_id = None
    col_sitio_mvs = None
    for col in df_clus.columns:
        if col is None: continue
        norm = normalize_column_name(str(col))
        if norm in ['SITEID', 'SITE_ID', 'SITE']: col_site_id = col
        if norm in ['SITIOMVS', 'SITIO_MVS', 'SITIO MVS']: col_sitio_mvs = col
    if col_site_id is None: raise KeyError("Column not found: 'Site ID' in Cluster mapping.")
    if col_sitio_mvs is None: raise KeyError("Column not found: 'Sitio MVS' in Cluster mapping.")
    
    df_clus['_site_id_norm'] = df_clus[col_site_id].fillna('').astype(str).str.strip().str.upper()
    df_clus['_site_id_norm'] = df_clus['_site_id_norm'].apply(lambda x: re.sub(r'\s+', ' ', x) if isinstance(x, str) else x)
    mapping = df_clus.drop_duplicates(subset=['_site_id_norm']).set_index('_site_id_norm')[col_sitio_mvs].to_dict()
    df_out = df_ctrl.copy()
    sitio_idx = df_out.columns.get_loc(col_sitio)
    cols = list(df_out.columns)
    cols.insert(sitio_idx + 1, 'Sitio MVS')
    cols.insert(sitio_idx + 2, 'Flag')
    df_out = df_out.reindex(columns=cols)
    sitio_vals = df_out[col_sitio].fillna('').astype(str).str.strip().str.upper()
    sitio_vals = sitio_vals.apply(lambda x: re.sub(r'\s+', ' ', x) if isinstance(x, str) else x)
    df_out['Sitio MVS'] = sitio_vals.map(mapping).fillna(sitio_vals)
    df_out['Flag'] = ''
    update_callback(80, "Generating output...")
    out_file_base = os.path.join(out_dir, "CLUSTER_CONTROL_FRIENDLY")
    try:
        if fmt_csv: df_out.to_csv(out_file_base + ".csv", index=False, encoding='utf-8-sig', sep=',')
        if fmt_xlsx: df_out.to_excel(out_file_base + ".xlsx", index=False)
    except PermissionError as e:
        raise PermissionError("Permission denied. The file might be open in another program. Please close it and try again.\n\n" + f"({str(e)})")
    update_callback(100, "Generation completed.")
    return out_file_base + (".csv" if fmt_csv else ".xlsx")

def generate_itsc_template(clusters_str, cluster_control_path, out_dir, fmt_csv, fmt_xlsx, update_callback):
    update_callback(5, "Waiting for instructions...")
    clusters = [c.strip() for c in clusters_str.split(',') if c.strip()]
    if not clusters: raise ValueError("You must specify at least one cluster.")
    if not cluster_control_path or not os.path.exists(cluster_control_path):
        raise ValueError("Cluster Control file not found.")
    ext = os.path.splitext(cluster_control_path)[1].lower()
    if ext in ['.xlsx', '.xls']:
        try: df_ctrl = pd.read_excel(cluster_control_path, dtype=str)
        except Exception as e: raise ValueError(f"Error reading Cluster Control Excel: {e}")
    else: df_ctrl = read_safe_csv(cluster_control_path)
    if df_ctrl.empty: raise ValueError("Cluster Control file is empty or could not be read.")
    update_callback(20, "Identifying columns...")
    col_clust_id = buscar_columna_inteligente(df_ctrl.columns, ['CLUST', 'ID']) or buscar_columna_inteligente(df_ctrl.columns, ['CLUST_ID']) or 'Clust_Id'
    col_activation = buscar_columna_inteligente(df_ctrl.columns, ['CLUSTER', 'ACTIVATION']) or buscar_columna_inteligente(df_ctrl.columns, ['ACTIVATION']) or 'cluster activation'
    col_sitio = buscar_columna_inteligente(df_ctrl.columns, ['SITIO']) or 'Sitio'
    col_sitio_mvs = buscar_columna_inteligente(df_ctrl.columns, ['SITIO', 'MVS']) or 'Sitio MVS'
    col_2g = buscar_columna_inteligente(df_ctrl.columns, ['2G']) or '2G'
    col_3g = buscar_columna_inteligente(df_ctrl.columns, ['3G']) or '3G'
    col_4g = buscar_columna_inteligente(df_ctrl.columns, ['4G']) or '4G'
    required_cols = [col_clust_id, col_activation, col_sitio, col_sitio_mvs, col_2g, col_3g, col_4g]
    missing = [c for c in required_cols if c not in df_ctrl.columns]
    if missing: raise KeyError(f"Missing required columns: {missing}")
    update_callback(30, "Filtering by clusters...")
    total_inicial = len(df_ctrl)
    df_ctrl = df_ctrl[df_ctrl[col_clust_id].astype(str).str.upper().isin([c.upper() for c in clusters])].copy()
    update_callback(35, "Filtering by clusters...")
    def keep_activation(val):
        if pd.isna(val): return False
        v = str(val).strip().lower()
        date_pattern = r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}'
        has_date = bool(re.search(date_pattern, v))
        v_norm = re.sub(r'[^\w\s]', '', v)
        has_customer_activate = 'customer activate' in v_norm
        has_customer_activated = 'customer activated' in v_norm
        if has_date: return True
        if has_customer_activate or has_customer_activated: return True
        return False
    update_callback(40, "Applying activation filter...")
    antes_activ = len(df_ctrl)
    df_ctrl = df_ctrl[df_ctrl[col_activation].apply(keep_activation)].copy()
    if df_ctrl.empty: raise ValueError(f"No records remain after the activation filter. Discarded {antes_activ} rows. Check column '{col_activation}'.")
    update_callback(55, "Processing sites and technologies...")
    def extract_5char(value):
        if pd.isna(value): return None
        s = str(value).strip()
        if '.' in s: s = s.split('.')[0]
        s = re.sub(r'[^A-Za-z0-9]', '', s)
        if len(s) == 5: return s
        return None
    tech_map = {'2G': 'G', '3G': 'U', '4G': 'L'}
    keywords = ['MORAN', 'MOCN', 'MORAN (MVS)', 'KEEP TP', 'KEEP MVS']
    def normalize_text(txt):
        if not isinstance(txt, str): txt = str(txt)
        return re.sub(r'[^\w]', '', txt.upper())
    rows = []
    sitios_sin_5car = 0
    primer_sitio_invalido = None
    for idx, row in df_ctrl.iterrows():
        site_code = extract_5char(row[col_sitio])
        if site_code is None: site_code = extract_5char(row[col_sitio_mvs])
        if site_code is None:
            sitios_sin_5car += 1
            if primer_sitio_invalido is None: primer_sitio_invalido = f"Sitio='{row[col_sitio]}', Sitio MVS='{row[col_sitio_mvs]}'"
            continue
        cluster_id = str(row[col_clust_id]).strip()
        for tech_col, tech_letter in tech_map.items():
            tech_value_raw = row[tech_col]
            if pd.isna(tech_value_raw): continue
            tech_value_norm = normalize_text(tech_value_raw)
            found = False
            for kw in keywords:
                if normalize_text(kw) in tech_value_norm:
                    found = True
                    break
            if found:
                ne_id = f"{cluster_id}_{site_code}_{tech_letter}"
                rows.append({
                    'NE_ID': ne_id, 'SITE_ID': site_code, 'NE_NAME': site_code, 'TYPE': '',
                    'SOURCE_SYSTEM': '', 'SOURCE_ID': '', 'IS_VALID': '', 'LAST_SYNC_TIME': '',
                    'ISDP_PROJECT_CODE': '', 'cluster_id': cluster_id
                })
    if not rows: raise ValueError(f"No records generated. Sites without 5 characters: {sitios_sin_5car}. Check columns '{col_sitio}'. Invalid example: {primer_sitio_invalido}")
    df_out = pd.DataFrame(rows)
    column_order = ['NE_ID', 'SITE_ID', 'NE_NAME', 'TYPE', 'SOURCE_SYSTEM', 'SOURCE_ID',
                    'IS_VALID', 'LAST_SYNC_TIME', 'ISDP_PROJECT_CODE', 'cluster_id']
    df_out = df_out[column_order]
    update_callback(80, "Exporting file...")
    out_file_base = os.path.join(out_dir, "ITSC_TEMPLATE")
    try:
        if fmt_csv: df_out.to_csv(out_file_base + ".csv", index=False, encoding='utf-8-sig', sep=',')
        if fmt_xlsx: df_out.to_excel(out_file_base + ".xlsx", index=False)
    except PermissionError as e: raise PermissionError("Permission denied. The file might be open in another program. Please close it and try again.\n\n" + f"({str(e)})")
    update_callback(100, "Generation completed.")
    return out_file_base + (".csv" if fmt_csv else ".xlsx")

def generate_itsc_swap(swap_table_text, out_dir, fmt_csv, fmt_xlsx, update_callback):
    update_callback(5, "Processing sites and technologies...")
    lines = [line.strip() for line in swap_table_text.splitlines() if line.strip()]
    if not lines: raise ValueError("The SWAP data table is empty. Please enter at least one site.")
    
    tech_map = {'2G': 'G', '3G': 'U', '4G': 'L', '5G': 'N'}
    rows = []
    for line in lines:
        parts = line.split(None, 1)
        if len(parts) < 2: continue
        site, tech_str = parts[0], parts[1]
        tech_list = [x.strip().upper() for x in tech_str.split('_') if x.strip()]
        for tech in tech_list:
            suffix = tech_map.get(tech, '')
            if suffix:
                ne_id = f"{site}_{suffix}"
                rows.append({
                    'NE_ID': ne_id, 'SITE_ID': site, 'NE_NAME': site, 'TYPE': '',
                    'SOURCE_SYSTEM': '', 'SOURCE_ID': '', 'IS_VALID': '', 'LAST_SYNC_TIME': '',
                    'ISDP_PROJECT_CODE': '', 'cluster_id': ''
                })
    if not rows: raise ValueError("Objects could not be generated. Check format: Site Tech1_Tech2 (e.g. XX001 2G_3G)")
    df = pd.DataFrame(rows)
    column_order = ['NE_ID', 'SITE_ID', 'NE_NAME', 'TYPE', 'SOURCE_SYSTEM', 'SOURCE_ID',
                    'IS_VALID', 'LAST_SYNC_TIME', 'ISDP_PROJECT_CODE', 'cluster_id']
    df = df[column_order]
    out_file_base = os.path.join(out_dir, "ITSC_SWAP_TEMPLATE")
    try:
        if fmt_csv: df.to_csv(out_file_base + ".csv", index=False, encoding='utf-8-sig', sep=',')
        if fmt_xlsx: df.to_excel(out_file_base + ".xlsx", index=False)
    except PermissionError as e: raise PermissionError("Permission denied. The file might be open in another program. Please close it and try again.\n\n" + f"({str(e)})")
    update_callback(100, "Generation completed.")
    return out_file_base + (".csv" if fmt_csv else ".xlsx")

# =========================================================================
# GENERATE BATCH Y AUDIT
# =========================================================================
def generate_batch_audit(path_old, path_new, out_dir, update_callback):
    update_callback(5, "Waiting for instructions...")
    
    update_callback(10, "Processing AUDIT... (OLD)")
    df_old = read_safe_file(path_old)
    if df_old.empty: raise ValueError("MASTERBATCH OLD is empty or could not be read.")
    
    update_callback(30, "Processing AUDIT... (NEW)")
    df_new = read_safe_file(path_new)
    if df_new.empty: raise ValueError("MASTERBATCH NEW is empty or could not be read.")

    c_clus_old = buscar_columna_inteligente(df_old.columns, ['CLUSTER']) or 'Cluster'
    c_unico_old = buscar_columna_inteligente(df_old.columns, ['UNICO']) or 'UNICO'
    c_cell_old = buscar_columna_inteligente(df_old.columns, ['CELL', 'NAME']) or 'CellName'

    c_clus_new = buscar_columna_inteligente(df_new.columns, ['CLUSTER']) or 'Cluster'
    c_unico_new = buscar_columna_inteligente(df_new.columns, ['UNICO']) or 'UNICO'
    c_cell_new = buscar_columna_inteligente(df_new.columns, ['CELL', 'NAME']) or 'CellName'

    for col in [c_clus_old, c_unico_old, c_cell_old]:
        if col not in df_old.columns:
            raise KeyError(f"Required column '{col}' not found in MASTERBATCH OLD.")
            
    for col in [c_clus_new, c_unico_new, c_cell_new]:
        if col not in df_new.columns:
            raise KeyError(f"Required column '{col}' not found in MASTERBATCH NEW.")

    update_callback(50, "Processing AUDIT... (Analyzing Deltas)")
    
    df_old[c_clus_old] = df_old[c_clus_old].fillna('SIN CLUSTER').astype(str).str.strip().str.upper()
    df_new[c_clus_new] = df_new[c_clus_new].fillna('SIN CLUSTER').astype(str).str.strip().str.upper()

    agg_old = df_old.groupby(c_clus_old).agg(
        OLD_Sites=(c_unico_old, lambda x: x.dropna().nunique()),
        OLD_Cells=(c_cell_old, lambda x: x.dropna().nunique())
    ).reset_index().rename(columns={c_clus_old: 'Cluster'})
    
    agg_new = df_new.groupby(c_clus_new).agg(
        NEW_Sites=(c_unico_new, lambda x: x.dropna().nunique()),
        NEW_Cells=(c_cell_new, lambda x: x.dropna().nunique())
    ).reset_index().rename(columns={c_clus_new: 'Cluster'})

    df_audit = pd.merge(agg_old, agg_new, on='Cluster', how='outer').fillna(0)
    df_audit['OLD_Sites'] = df_audit['OLD_Sites'].astype(int)
    df_audit['NEW_Sites'] = df_audit['NEW_Sites'].astype(int)
    df_audit['OLD_Cells'] = df_audit['OLD_Cells'].astype(int)
    df_audit['NEW_Cells'] = df_audit['NEW_Cells'].astype(int)

    df_audit['Delta Sitios'] = df_audit['NEW_Sites'] - df_audit['OLD_Sites']
    df_audit['Delta Celdas'] = df_audit['NEW_Cells'] - df_audit['OLD_Cells']
    
    df_audit = df_audit[['Cluster', 'OLD_Sites', 'NEW_Sites', 'Delta Sitios', 'OLD_Cells', 'NEW_Cells', 'Delta Celdas']]
    df_audit.sort_values(by='Cluster', inplace=True)
    
    # Agregar Fila de Totales
    totals = df_audit[['OLD_Sites', 'NEW_Sites', 'Delta Sitios', 'OLD_Cells', 'NEW_Cells', 'Delta Celdas']].sum()
    total_row = pd.DataFrame({
        'Cluster': ['TOTAL'],
        'OLD_Sites': [totals['OLD_Sites']],
        'NEW_Sites': [totals['NEW_Sites']],
        'Delta Sitios': [totals['Delta Sitios']],
        'OLD_Cells': [totals['OLD_Cells']],
        'NEW_Cells': [totals['NEW_Cells']],
        'Delta Celdas': [totals['Delta Celdas']]
    })
    df_audit = pd.concat([df_audit, total_row], ignore_index=True)

    update_callback(70, "Processing AUDIT... (Exporting to Excel)")
    out_file = os.path.join(out_dir, "MASTERBATCH_AUDIT.xlsx")
    
    try:
        with pd.ExcelWriter(out_file, engine='openpyxl') as writer:
            df_audit.to_excel(writer, sheet_name='AUDIT', index=False)
            df_old.to_excel(writer, sheet_name='MASTERBATCH OLD', index=False)
            del df_old
            gc.collect()
            df_new.to_excel(writer, sheet_name='MASTERBATCH NEW', index=False)
            del df_new
            gc.collect()
    except PermissionError as e:
        raise PermissionError("Permission denied. The file might be open in another program. Please close it and try again.\n\n" + f"({str(e)})")

    update_callback(100, "Generation completed.")
    return out_file

def generate_batch(p_master_ept, p_ran, p_obj, p_rnd, p_dat, p_ctrl, p_prog, p_clus, out_dir, fmt_csv, fmt_xlsx, update_callback):
    update_callback(5, "Waiting for instructions...")

    update_callback(8, "Loading reference files...")
    df_clus = read_safe_file(p_clus) if p_clus and os.path.exists(p_clus) else pd.DataFrame()
    df_ctrl = read_safe_file(p_ctrl) if p_ctrl and os.path.exists(p_ctrl) else pd.DataFrame()
    df_datos = read_safe_file(p_dat) if p_dat and os.path.exists(p_dat) else pd.DataFrame()
    df_prog = read_safe_file(p_prog) if p_prog and os.path.exists(p_prog) else pd.DataFrame()
    df_master_ept = read_safe_file(p_master_ept) if p_master_ept and os.path.exists(p_master_ept) else pd.DataFrame()

    col_site_id = buscar_columna_inteligente(df_clus.columns, ['SITE', 'ID']) or buscar_columna_inteligente(df_clus.columns, ['SITEID']) or 'Site ID'
    col_sitio_mvs = buscar_columna_inteligente(df_clus.columns, ['SITIO', 'MVS']) or 'Sitio MVS'
    col_cluster_id = buscar_columna_inteligente(df_clus.columns, ['CLUSTER']) or buscar_columna_inteligente(df_clus.columns, ['ID_CLUSTER']) or 'ID_Clúster'
    dict_clus_site = build_dict(df_clus, col_site_id, col_cluster_id) if col_site_id in df_clus.columns else {}
    dict_clus_mvs  = build_dict(df_clus, col_sitio_mvs, col_cluster_id) if col_sitio_mvs in df_clus.columns else {}

    col_emg = buscar_columna_inteligente(df_datos.columns, ['EMG']) or 'EMG'
    col_codigo = buscar_columna_inteligente(df_datos.columns, ['CÓDIGO']) or 'Código'
    dict_datos_emg = build_dict(df_datos, col_emg, col_codigo, sanitize_val=True) if col_emg in df_datos.columns else {}
    dict_datos_cod = build_dict(df_datos, col_codigo, col_emg, sanitize_val=True) if col_codigo in df_datos.columns else {}

    col_ctrl_sitio = buscar_columna_inteligente(df_ctrl.columns, ['SITIO']) or 'Sitio'
    col_ctrl_sitio_mvs = buscar_columna_inteligente(df_ctrl.columns, ['SITIO', 'MVS']) or 'SITIO MVS'
    col_ctrl_emg = buscar_columna_inteligente(df_ctrl.columns, ['EMG']) or 'EMG'
    col_ctrl_cluster = buscar_columna_inteligente(df_ctrl.columns, ['CLUST', 'ID']) or buscar_columna_inteligente(df_ctrl.columns, ['CLUST_ID']) or 'Clust_Id'
    dict_ctrl_sitio = build_dict(df_ctrl, col_ctrl_sitio, col_ctrl_cluster) if col_ctrl_sitio in df_ctrl.columns else {}
    dict_ctrl_mvs   = build_dict(df_ctrl, col_ctrl_sitio_mvs, col_ctrl_cluster) if col_ctrl_sitio_mvs in df_ctrl.columns else {}
    dict_ctrl_emg   = build_dict(df_ctrl, col_ctrl_emg, col_ctrl_cluster) if col_ctrl_emg in df_ctrl.columns else {}

    col_ept_cell = buscar_columna_inteligente(df_master_ept.columns, ['CELL', 'NAME']) or 'Cell Name'
    col_ept_cluster = buscar_columna_inteligente(df_master_ept.columns, ['CLUSTER']) or 'CLUSTER'
    dict_ept_cluster = build_dict(df_master_ept, col_ept_cell, col_ept_cluster) if col_ept_cell in df_master_ept.columns else {}

    col_ept_enodeb = buscar_columna_inteligente(df_master_ept.columns, ['ENODEB', 'NAME']) or buscar_columna_inteligente(df_master_ept.columns, ['ENODEBNAME']) or 'eNodeB Name'
    
    col_datos_lon = buscar_columna_inteligente(df_datos.columns, ['LONG']) or buscar_columna_inteligente(df_datos.columns, ['LON']) if not df_datos.empty else None
    col_datos_lat = buscar_columna_inteligente(df_datos.columns, ['LAT']) if not df_datos.empty else None
    dict_datos_lon = build_dict(df_datos, col_emg, col_datos_lon) if col_datos_lon else {}
    dict_datos_lat = build_dict(df_datos, col_emg, col_datos_lat) if col_datos_lat else {}

    col_clus_lon = buscar_columna_inteligente(df_clus.columns, ['LONG']) or buscar_columna_inteligente(df_clus.columns, ['LON']) if not df_clus.empty else None
    col_clus_lat = buscar_columna_inteligente(df_clus.columns, ['LAT']) if not df_clus.empty else None
    dict_clus_lon = build_dict(df_clus, col_site_id, col_clus_lon) if col_clus_lon else {}
    dict_clus_lat = build_dict(df_clus, col_site_id, col_clus_lat) if col_clus_lat else {}

    dict_ept_lon, dict_ept_lat, dict_ept_azi = {}, {}, {}
    if not df_master_ept.empty:
        col_ept_lon = buscar_columna_inteligente(df_master_ept.columns, ['LONG']) or buscar_columna_inteligente(df_master_ept.columns, ['LON'])
        col_ept_lat = buscar_columna_inteligente(df_master_ept.columns, ['LAT'])
        col_ept_azi = (buscar_columna_inteligente(df_master_ept.columns, ['AZIMUTH']) or 
                       buscar_columna_inteligente(df_master_ept.columns, ['AZIMUT']) or 
                       buscar_columna_inteligente(df_master_ept.columns, ['AZI']))
        if col_ept_lon: dict_ept_lon = build_dict(df_master_ept, col_ept_cell, col_ept_lon)
        if col_ept_lat: dict_ept_lat = build_dict(df_master_ept, col_ept_cell, col_ept_lat)
        if col_ept_azi: dict_ept_azi = build_dict(df_master_ept, col_ept_cell, col_ept_azi)

    del df_clus
    del df_datos
    gc.collect()

    update_callback(15, "Processing RAN REPORTS...")
    all_dfs = []
    if os.path.exists(p_ran):
        for root, dirs, files in os.walk(p_ran):
            for file in files:
                if not file.upper().endswith('.CSV'): continue
                filepath = os.path.join(root, file)
                name_parts = file.upper().replace('.CSV', '').split('_')
                if len(name_parts) != 2: continue
                tech_val, region_val = name_parts[0], name_parts[1]
                try:
                    try: df = pd.read_csv(filepath, dtype=str, encoding='utf-8-sig')
                    except UnicodeDecodeError: df = pd.read_csv(filepath, dtype=str, encoding='latin1')
                    df.columns = df.columns.astype(str).str.replace('\xa0', ' ', regex=False).str.strip().str.lower()
                    cell_col = [col for col in df.columns if 'cell' in col and 'name' in col]
                    if not cell_col: continue
                    df = df.rename(columns={cell_col[0]: 'cellname_raw'})
                    df = df.dropna(subset=['cellname_raw'])
                    df['CellName'] = df['cellname_raw'].astype(str).str.replace('\xa0', ' ', regex=False).str.strip().str.upper()
                    if tech_val == '4G' and 'rat' in df.columns:
                        df = df[df['rat'].astype(str).str.strip().str.upper() != 'NBIOT']
                    df['length'] = df['CellName'].str.len()
                    df = df[df['length'].isin([7, 8, 9, 14])]
                    if df.empty: continue
                    df['Tech'] = tech_val
                    df['Region'] = region_val
                    df['UNICO'] = np.where(df['length'].isin([7, 8]), df['CellName'].str[1:6], df['CellName'].str[:6])
                    df['Vendor'] = np.where(df['length'].isin([7, 8]), 'HW_MVS', 'HW_TA')
                    df['Sector'] = df['CellName'].str[-1]

                    if tech_val == '5G': df['Band'] = '78'
                    elif tech_val == '2G':
                        cond_7_8 = df['length'].isin([7, 8])
                        df['Band'] = np.where(cond_7_8, df['CellName'].str[0].map({'C': '5', 'G': '2'}), df['CellName'].str[6].map({'H': '5', 'G': '2', 'D': '8'}))
                    elif tech_val == '3G':
                        dl_col = [col for col in df.columns if 'dl' in col and 'freq' in col]
                        if dl_col:
                            raw_freq = df[dl_col[0]].astype(str).str.replace(',', '', regex=False)
                            df['freq_mhz'] = (pd.to_numeric(raw_freq, errors='coerce') / 5).round(1)
                            conds = [
                                (df['freq_mhz'] >= 2110) & (df['freq_mhz'] <= 2155),
                                (df['freq_mhz'] > 2155) & (df['freq_mhz'] <= 2170),
                                (df['freq_mhz'] >= 1930) & (df['freq_mhz'] <= 1990),
                                (df['freq_mhz'] >= 869) & (df['freq_mhz'] <= 894),
                                (df['freq_mhz'] == 612.4),
                                (df['freq_mhz'].isin([97.4, 92.4])),
                                (df['freq_mhz'] == 201.4)
                            ]
                            df['Band'] = np.select(conds, ['4', '1', '2', '5', '8', '2', '5'], default='UNMAPPED')
                            df['Band'] = df['Band'].replace('UNMAPPED', np.nan)
                        else: df['Band'] = np.nan
                    elif tech_val == '4G':
                        df['Band'] = df['frequency band'].astype(str).str.strip().str.upper() if 'frequency band' in df.columns else np.nan

                    col_lon = buscar_columna_inteligente(df.columns, ['LONG']) or buscar_columna_inteligente(df.columns, ['LON'])
                    col_lat = buscar_columna_inteligente(df.columns, ['LAT'])
                    col_azi = buscar_columna_inteligente(df.columns, ['AZIMUTH']) or buscar_columna_inteligente(df.columns, ['AZIMUT']) or buscar_columna_inteligente(df.columns, ['AZI'])
                    df['Longitude'] = df[col_lon] if col_lon else np.nan
                    df['Latitude'] = df[col_lat] if col_lat else np.nan
                    df['Azimuth'] = df[col_azi] if col_azi else np.nan

                    all_dfs.append(df[['UNICO', 'CellName', 'Region', 'Band', 'Sector', 'Tech', 'Vendor', 'Longitude', 'Latitude', 'Azimuth']])
                except Exception as e:
                    update_callback(15, f"[BATCH] Error in RAN {file}: {str(e)}")
                    continue
    else:
        raise ValueError("For BATCH mode, MASTER EPT, RAN REPORTS, OBJECT TREES, and RND folder are mandatory. - RAN REPORTS folder not found.")

    update_callback(35, "Processing OBJECT TREES...")
    if os.path.exists(p_obj):
        for root, dirs, files in os.walk(p_obj):
            for file in files:
                if not file.upper().endswith('.CSV'): continue
                filepath = os.path.join(root, file)
                tech_val = file.upper().split(' ')[0]
                if tech_val not in ['2G', '3G', '4G', '5G']:
                    tech_val = file.upper()[:2]
                    if tech_val not in ['2G', '3G', '4G', '5G']: continue
                try:
                    try: df = pd.read_csv(filepath, dtype=str, encoding='utf-8-sig', skiprows=2)
                    except UnicodeDecodeError: df = pd.read_csv(filepath, dtype=str, encoding='latin1', skiprows=2)
                    df.columns = df.columns.astype(str).str.replace('\xa0', ' ', regex=False).str.strip().str.lower()
                    cell_col = [col for col in df.columns if 'cell' in col and 'name' in col]
                    if not cell_col: continue
                    df = df.rename(columns={cell_col[0]: 'cellname_raw'})
                    df = df.dropna(subset=['cellname_raw'])
                    df['CellName'] = df['cellname_raw'].astype(str).str.replace('\xa0', ' ', regex=False).str.strip().str.upper()
                    if tech_val == '4G' and 'rat' in df.columns:
                        df = df[df['rat'].astype(str).str.strip().str.upper() != 'NBIOT']
                    df['length'] = df['CellName'].str.len()
                    df = df[df['length'].isin([7, 8, 9, 14])]
                    if df.empty: continue
                    df['Tech'] = tech_val
                    df['Region'] = np.nan
                    cond_7_8 = df['length'].isin([7, 8])
                    df['UNICO'] = np.where(cond_7_8, df['CellName'].str[1:6], df['CellName'].str[:6])
                    df['Vendor'] = np.where(cond_7_8, 'HW_MVS', 'HW_TA')
                    df['Sector'] = df['CellName'].str[-1]

                    if tech_val == '5G': df['Band'] = '78'
                    elif tech_val == '2G':
                        df['Band'] = np.where(cond_7_8, df['CellName'].str[0].map({'C': '5', 'G': '2'}), df['CellName'].str[6].map({'H': '5', 'G': '2', 'D': '8'}))
                    elif tech_val == '3G':
                        dl_col = [col for col in df.columns if 'dl' in col and 'freq' in col]
                        if dl_col:
                            raw_freq = df[dl_col[0]].astype(str).str.replace(',', '', regex=False)
                            df['freq_mhz'] = (pd.to_numeric(raw_freq, errors='coerce') / 5).round(1)
                            conds = [
                                (df['freq_mhz'] >= 2110) & (df['freq_mhz'] <= 2155),
                                (df['freq_mhz'] > 2155) & (df['freq_mhz'] <= 2170),
                                (df['freq_mhz'] >= 1930) & (df['freq_mhz'] <= 1990),
                                (df['freq_mhz'] >= 869) & (df['freq_mhz'] <= 894),
                                (df['freq_mhz'] == 612.4),
                                (df['freq_mhz'].isin([97.4, 92.4])),
                                (df['freq_mhz'] == 201.4)
                            ]
                            df['Band'] = np.select(conds, ['4', '1', '2', '5', '8', '2', '5'], default='UNMAPPED')
                            df['Band'] = df['Band'].replace('UNMAPPED', np.nan)
                        else: df['Band'] = np.nan
                    elif tech_val == '4G':
                        df['Band'] = df['frequency band'].astype(str).str.strip().str.upper() if 'frequency band' in df.columns else np.nan

                    col_lon = buscar_columna_inteligente(df.columns, ['LONG']) or buscar_columna_inteligente(df.columns, ['LON'])
                    col_lat = buscar_columna_inteligente(df.columns, ['LAT'])
                    col_azi = buscar_columna_inteligente(df.columns, ['AZIMUTH']) or buscar_columna_inteligente(df.columns, ['AZIMUT']) or buscar_columna_inteligente(df.columns, ['AZI'])
                    df['Longitude'] = df[col_lon] if col_lon else np.nan
                    df['Latitude'] = df[col_lat] if col_lat else np.nan
                    df['Azimuth'] = df[col_azi] if col_azi else np.nan

                    all_dfs.append(df[['UNICO', 'CellName', 'Region', 'Band', 'Sector', 'Tech', 'Vendor', 'Longitude', 'Latitude', 'Azimuth']])
                except Exception as e:
                    update_callback(35, f"[BATCH] Error in OBJ {file}: {str(e)}")
                    continue
    else:
        raise ValueError("For BATCH mode, MASTER EPT, RAN REPORTS, OBJECT TREES, and RND folder are mandatory. - OBJECT TREES folder not found.")

    update_callback(55, "Processing RND (RSH/SWAP)...")
    rsh_dfs, swap_dfs = [], []
    dict_rnd_lon, dict_rnd_lat, dict_rnd_azi = {}, {}, {}

    def extract_rnd_data(df_raw, sheet_name, file_name, target_list):
        nonlocal dict_rnd_lon, dict_rnd_lat, dict_rnd_azi
        df_raw.columns = df_raw.columns.astype(str).str.strip()
        col_celltype = buscar_columna_inteligente(df_raw.columns, ['CELL', 'TYPE'])
        col_cellname = buscar_columna_inteligente(df_raw.columns, ['CELL', 'NAME'])
        col_unico    = buscar_columna_inteligente(df_raw.columns, ['UNICO'])
        col_freqband = buscar_columna_inteligente(df_raw.columns, ['FREQUENCY', 'BAND']) or buscar_columna_inteligente(df_raw.columns, ['BAND'])
        col_rat      = buscar_columna_inteligente(df_raw.columns, ['RAT'])
        if not all([col_celltype, col_cellname, col_unico, col_freqband, col_rat]): return
        col_lon = buscar_columna_inteligente(df_raw.columns, ['LONG']) or buscar_columna_inteligente(df_raw.columns, ['LON'])
        col_lat = buscar_columna_inteligente(df_raw.columns, ['LAT'])
        col_azi = (buscar_columna_inteligente(df_raw.columns, ['AZIMUTH']) or buscar_columna_inteligente(df_raw.columns, ['AZIMUT']) or buscar_columna_inteligente(df_raw.columns, ['AZI']))

        df_clean = pd.DataFrame()
        df_clean['cell type'] = df_raw[col_celltype]
        df_clean['cell name'] = df_raw[col_cellname]
        df_clean['unico'] = df_raw[col_unico]
        df_clean['frequency band'] = df_raw[col_freqband]
        df_clean['rat'] = df_raw[col_rat]
        df_clean['Longitude'] = df_raw[col_lon] if col_lon else np.nan
        df_clean['Latitude'] = df_raw[col_lat] if col_lat else np.nan
        df_clean['Azimuth'] = df_raw[col_azi] if col_azi else np.nan

        df_clean = df_clean[df_clean['cell type'].astype(str).str.upper().isin(['FDD', 'TDD'])].copy()
        df_clean['CellName'] = df_clean['cell name'].str.strip().str.upper()
        df_clean['len'] = df_clean['CellName'].str.len()
        df_clean['Vendor'] = np.where(df_clean['len'].isin([7, 8]), 'HW_MVS', 'HW_TA')
        df_clean['UNICO']    = df_clean['unico'].str.strip().str.upper()
        df_clean['Band']     = df_clean['frequency band'].str.strip().str.upper()
        df_clean['Sector']   = df_clean['CellName'].str[-1]
        df_clean['Tech']     = df_clean['rat'].str.strip().str.upper()

        if col_lon: dict_rnd_lon.update(build_dict(df_clean, 'CellName', 'Longitude'))
        if col_lat: dict_rnd_lat.update(build_dict(df_clean, 'CellName', 'Latitude'))
        if col_azi: dict_rnd_azi.update(build_dict(df_clean, 'CellName', 'Azimuth'))

        target_list.append(df_clean[['UNICO', 'CellName', 'Band', 'Sector', 'Tech', 'Vendor', 'Longitude', 'Latitude', 'Azimuth']])

    rsh_path = os.path.join(p_rnd, 'RND_RSH.xlsx')
    if os.path.exists(rsh_path):
        try:
            xl = pd.ExcelFile(rsh_path)
            for sheet in xl.sheet_names:
                if 'MVS' in sheet.upper() or 'TP' in sheet.upper():
                    df = pd.read_excel(xl, sheet_name=sheet).astype(str)
                    extract_rnd_data(df, sheet.upper(), 'RND_RSH.xlsx', rsh_dfs)
        except Exception as e: update_callback(55, f"[BATCH] Error in RND_RSH: {str(e)}")

    swap_path = os.path.join(p_rnd, 'RND_SWAP.xlsx')
    if os.path.exists(swap_path):
        try:
            df_swap = pd.read_excel(swap_path, sheet_name='RND').astype(str)
            extract_rnd_data(df_swap, 'RND SHEET', 'RND_SWAP.xlsx', swap_dfs)
        except Exception as e: update_callback(55, f"[BATCH] Error in RND_SWAP: {str(e)}")

    update_callback(70, "Merging and assigning clusters...")
    if not (all_dfs or rsh_dfs or swap_dfs): raise ValueError("No valid data found in RAN, OBJ, or RND.")
    df_master = pd.concat(all_dfs + rsh_dfs + swap_dfs, ignore_index=True)
    
    del all_dfs
    del rsh_dfs
    del swap_dfs
    gc.collect()

    df_master['UNICO'] = clean_series(df_master['UNICO'])
    df_master['Cluster'] = pd.Series(np.nan, index=df_master.index, dtype=object)

    df_master['Cluster'] = df_master['Cluster'].fillna(df_master['UNICO'].map(dict_clus_site))
    df_master['Cluster'] = df_master['Cluster'].fillna(df_master['UNICO'].map(dict_clus_mvs))

    mask_missing = df_master['Cluster'].isna()
    if mask_missing.any() and dict_datos_emg and (dict_clus_site or dict_clus_mvs):
        codigos = df_master.loc[mask_missing, 'UNICO'].map(dict_datos_emg)
        df_master.loc[mask_missing, 'Cluster'] = codigos.map(dict_clus_site)
        mask_missing2 = df_master['Cluster'].isna() & mask_missing
        df_master.loc[mask_missing2, 'Cluster'] = codigos.map(dict_clus_mvs)

    mask_missing = df_master['Cluster'].isna()
    df_master.loc[mask_missing, 'Cluster'] = df_master.loc[mask_missing, 'UNICO'].map(dict_ctrl_sitio)
    mask_missing = df_master['Cluster'].isna()
    df_master.loc[mask_missing, 'Cluster'] = df_master.loc[mask_missing, 'UNICO'].map(dict_ctrl_mvs)
    mask_missing = df_master['Cluster'].isna()
    df_master.loc[mask_missing, 'Cluster'] = df_master.loc[mask_missing, 'UNICO'].map(dict_ctrl_emg)

    if dict_ept_cluster:
        mask_missing = df_master['Cluster'].isna()
        df_master.loc[mask_missing, 'Cluster'] = df_master.loc[mask_missing, 'CellName'].map(dict_ept_cluster)

    if not df_master_ept.empty and col_ept_enodeb in df_master_ept.columns and col_ept_cluster in df_master_ept.columns:
        missing_unicos = df_master[df_master['Cluster'].isna()]['UNICO'].dropna().unique()
        unico_to_cluster = {}
        for _, row in df_master_ept.iterrows():
            enb_val = str(row[col_ept_enodeb]).strip().upper()
            cluster_val = str(row[col_ept_cluster]).strip()
            if not enb_val or not cluster_val: continue
            for unico in missing_unicos:
                if unico in enb_val: unico_to_cluster[unico] = cluster_val
        mask_missing = df_master['Cluster'].isna()
        df_master.loc[mask_missing, 'Cluster'] = df_master.loc[mask_missing, 'UNICO'].map(unico_to_cluster)
        
    del df_master_ept
    gc.collect()

    known_clusters = df_master.dropna(subset=['Cluster']).set_index('UNICO')['Cluster'].to_dict()
    df_master['Cluster'] = df_master['Cluster'].fillna(df_master['UNICO'].map(known_clusters))
    df_master['Cluster'] = df_master['Cluster'].fillna('NO CLUSTER').astype(str)

    conds = [df_master['Cluster'].str.contains(x, na=False) for x in ['AMBA', 'MEDI', 'SUR', 'LITO']]
    df_master['Region'] = np.select(conds, ['AMBA', 'MEDI', 'SUR', 'LITO'], default='NO REGION')
    df_master['Region'] = df_master['Region'].astype('category')

    update_callback(80, "Assigning coordinates and azimuth...")
    for col in ['Longitude', 'Latitude', 'Azimuth']:
        if col not in df_master.columns: df_master[col] = np.nan
        df_master[col] = df_master[col].replace(['NAN', 'NONE', '', 'NaN', 'nan', 'NaT', 'NULL'], np.nan)

    df_master['Azimuth'] = df_master['Azimuth'].fillna(df_master['CellName'].map(dict_ept_azi))
    df_master['Azimuth'] = df_master['Azimuth'].fillna(df_master['CellName'].map(dict_rnd_azi))

    df_master['Longitude'] = df_master['Longitude'].fillna(df_master['UNICO'].map(dict_datos_lon))
    df_master['Latitude'] = df_master['Latitude'].fillna(df_master['UNICO'].map(dict_datos_lat))
    df_master['Longitude'] = df_master['Longitude'].fillna(df_master['CellName'].map(dict_ept_lon))
    df_master['Latitude'] = df_master['Latitude'].fillna(df_master['CellName'].map(dict_ept_lat))
    df_master['Longitude'] = df_master['Longitude'].fillna(df_master['CellName'].map(dict_rnd_lon))
    df_master['Latitude'] = df_master['Latitude'].fillna(df_master['CellName'].map(dict_rnd_lat))
    df_master['Longitude'] = df_master['Longitude'].fillna(df_master['UNICO'].map(dict_clus_lon))
    df_master['Latitude'] = df_master['Latitude'].fillna(df_master['UNICO'].map(dict_clus_lat))

    known_lons = df_master.dropna(subset=['Longitude']).set_index('UNICO')['Longitude'].to_dict()
    known_lats = df_master.dropna(subset=['Latitude']).set_index('UNICO')['Latitude'].to_dict()
    df_master['Longitude'] = df_master['Longitude'].fillna(df_master['UNICO'].map(known_lons))
    df_master['Latitude'] = df_master['Latitude'].fillna(df_master['UNICO'].map(known_lats))

    sector_equiv = {'1':'1','A':'1','G':'1','X':'1','M':'1',
                    '2':'2','B':'2','H':'2','Y':'2','N':'2',
                    '3':'3','C':'3','I':'3','Z':'3','O':'3'}
    df_master['Temp_Sec_Grp'] = df_master['Sector'].astype(str).str.upper().map(sector_equiv).fillna('UNK')
    df_master['Azi_Key'] = df_master['UNICO'] + "_" + df_master['Temp_Sec_Grp']
    known_azi = df_master.dropna(subset=['Azimuth']).set_index('Azi_Key')['Azimuth'].to_dict()
    df_master['Azimuth'] = df_master['Azimuth'].fillna(df_master['Azi_Key'].map(known_azi))
    df_master.drop(columns=['Temp_Sec_Grp', 'Azi_Key'], inplace=True)

    df_master.drop_duplicates(subset=['UNICO', 'CellName', 'Band', 'Sector', 'Tech', 'Vendor'], inplace=True)

    update_callback(90, "Applying MOCN rules...")
    mocn_blocks = []

    if not df_ctrl.empty:
        tech_map = {}
        for t in ['2G', '3G', '4G', '5G']:
            col_found = buscar_columna_inteligente(df_ctrl.columns, [t])
            if col_found:
                tech_map[t] = col_found

        cand_cols = [c for c in [col_ctrl_sitio, col_ctrl_sitio_mvs, col_ctrl_emg] if c and c in df_ctrl.columns]
        
        if tech_map and cand_cols:
            for tech_key, real_col_name in tech_map.items():
                mask_mocn = df_ctrl[real_col_name].astype(str).str.upper().str.contains('MOCN', na=False)
                if not mask_mocn.any(): continue
                
                cands_matrix = df_ctrl.loc[mask_mocn, cand_cols].astype(str).apply(lambda x: x.str.strip().str.upper())
                cands_set = set(cands_matrix.values.flatten())
                cands_set.difference_update({'NAN', 'NONE', '', 'NAT'})
                if not cands_set: continue
                
                mask_tech = (df_master['Tech'] == tech_key) & (df_master['UNICO'].isin(cands_set))
                subset = df_master[mask_tech].copy()
                if not subset.empty:
                    subset['Vendor'] = flip_vendor(subset['Vendor'])
                    mocn_blocks.append(subset)

    mask_sur_3g = (df_master['Region'] == 'SUR') & (df_master['Tech'] == '3G')
    df_sur_3g = df_master[mask_sur_3g].copy()
    if not df_sur_3g.empty:
        df_sur_3g['Vendor'] = flip_vendor(df_sur_3g['Vendor'])
        mocn_blocks.append(df_sur_3g)

    if not df_prog.empty:
        col_sitio = buscar_columna_inteligente(df_prog.columns, ['SITIO']) or 'Sitio'
        col_label = buscar_columna_inteligente(df_prog.columns, ['LABEL'])
        if col_sitio and col_label:
            mask_mocn = df_prog[col_label].astype(str).str.upper().str.contains('MOCN', na=False)
            valid_sites_raw = df_prog.loc[mask_mocn, col_sitio]
            valid_sites = clean_series(valid_sites_raw).unique()
            valid_sites = valid_sites[(valid_sites != 'NAN') & (valid_sites != '')]
            if len(valid_sites) > 0:
                mapped_emgs = [dict_datos_cod.get(s, s) for s in valid_sites]
                search_targets = set(valid_sites).union(set(mapped_emgs))
                mask_sitio = (df_master['Tech'] == '4G') & (df_master['UNICO'].isin(search_targets))
                subset = df_master[mask_sitio].copy()
                if not subset.empty:
                    subset['Vendor'] = flip_vendor(subset['Vendor'])
                    mocn_blocks.append(subset)

    if mocn_blocks:
        df_master = pd.concat([df_master] + mocn_blocks, ignore_index=True)
        df_master.drop_duplicates(subset=['UNICO', 'CellName', 'Band', 'Sector', 'Tech', 'Vendor'], inplace=True)

    del df_ctrl
    del df_prog
    del mocn_blocks
    gc.collect()

    for cat_col in ['Tech', 'Band', 'Sector', 'Vendor']:
        df_master[cat_col] = df_master[cat_col].astype('category')

    update_callback(95, "Exporting final file...")
    MASTER_COLUMNS = ['UNICO', 'CellName', 'Site OSS Name', 'Region', 'Band', 'Sector', 'Tech', 'Cluster', 'Vendor', 'CellID', 'SwapDate', 'Status', 'Longitude', 'Latitude', 'Azimuth']
    for col in MASTER_COLUMNS:
        if col not in df_master.columns: df_master[col] = ""
    df_master = df_master.reindex(columns=MASTER_COLUMNS)

    for col in ['UNICO', 'CellName', 'Band', 'Sector', 'Tech', 'Vendor']:
        df_master[col] = clean_series(df_master[col])
    df_master.drop_duplicates(subset=['UNICO', 'CellName', 'Band', 'Sector', 'Tech', 'Vendor'], inplace=True)
    df_master.sort_values(by=['Cluster', 'Tech', 'UNICO', 'CellName', 'Vendor'], ascending=[True, True, True, True, True], na_position='last', inplace=True)

    out_file_base = os.path.join(out_dir, "MASTERBATCH")
    try:
        if fmt_csv: df_master.to_csv(out_file_base + ".csv", index=False, encoding='latin1', sep=',')
        if fmt_xlsx: df_master.to_excel(out_file_base + ".xlsx", index=False)
    except PermissionError as e:
        raise PermissionError("Permission denied. The file might be open in another program. Please close it and try again.\n\n" + f"({str(e)})")

    update_callback(100, "Generation completed.")
    return out_file_base + (".csv" if fmt_csv else ".xlsx")

# =========================================================================
# OBJECT TREES (REFACTORIZADO)
# =========================================================================
COL_IDX_MAP = {'2G': 8, '3G': 6, '4G': 8, '5G': 5}

def get_file_encoding(filepath):
    try:
        with open(filepath, 'rb') as f: raw = f.read(10000)
        result = chardet.detect(raw)
        enc = result['encoding']
        if enc and enc.lower() in ['big5', 'gb2312', 'gb18030']: return 'latin1'
        return enc or 'latin1'
    except: return 'latin1'

def _process_object_trees(obj_folder, target_name, cells_to_match, out_dir, fmt_csv, fmt_xlsx):
    target_out_dir = os.path.join(out_dir, f"OBJECT_{target_name}")
    os.makedirs(target_out_dir, exist_ok=True)

    for root, dirs, files in os.walk(obj_folder):
        for file in files:
            if 'OBJECT' not in file.upper(): continue
            
            tech = next((t_key for t_key in COL_IDX_MAP.keys() if t_key in file.upper()), None)
            if not tech: continue
            
            filepath = os.path.join(root, file)
            enc = get_file_encoding(filepath)

            with open(filepath, 'r', encoding=enc, errors='replace') as f:
                rows = list(csv.reader(f, delimiter=','))

            if len(rows) < 4: continue
            col_idx = COL_IDX_MAP.get(tech)

            data_rows = []
            for row in rows[3:]:
                if len(row) > col_idx:
                    cell_val = row[col_idx].strip().upper().replace('"', '').replace('\xa0', ' ')
                    if cell_val in cells_to_match: 
                        data_rows.append(row)

            if data_rows:
                df_filtered = pd.DataFrame(rows[:3] + data_rows)
                out_name = f"{os.path.splitext(file)[0]}_{target_name}"
                
                if fmt_csv: 
                    df_filtered.to_csv(os.path.join(target_out_dir, out_name + ".csv"), index=False, header=False, sep=',', quoting=csv.QUOTE_ALL, encoding='utf-8-sig')
                if fmt_xlsx: 
                    df_filtered.to_excel(os.path.join(target_out_dir, out_name + ".xlsx"), index=False, header=False)

def generate_object_trees_regions(obj_folder, batchfile_path, out_dir, fmt_csv, fmt_xlsx, update_callback):
    if not os.path.exists(obj_folder): return
    update_callback(5, "Processing OBJECT TREES - REGIONS...")

    df_batch = read_safe_file(batchfile_path)

    c_cell = buscar_columna_inteligente(df_batch.columns, ['CELL', 'NAME'])
    c_region = buscar_columna_inteligente(df_batch.columns, ['REGION'])
    if not c_cell or not c_region: raise ValueError("Batchfile must contain columns for Cell Name and Region.")

    df_map = pd.DataFrame({
        'CELL': clean_series(df_batch[c_cell]),
        'REGION': clean_series(df_batch[c_region])
    })
    
    del df_batch
    gc.collect()
    
    df_map = df_map[(df_map['CELL'] != '') & (df_map['REGION'] != '')]
    df_map['REGION'] = df_map['REGION'].apply(lambda x: 'MEDI' if x == 'LITO' else x)
    valid_regions = {'AMBA', 'MEDI', 'SUR'}
    df_map = df_map[df_map['REGION'].isin(valid_regions)]
    if df_map.empty: raise ValueError("No valid region mappings found in batchfile (expected AMBA, MEDI, SUR).")

    region_cells = df_map.groupby('REGION')['CELL'].apply(set).to_dict()
    del df_map
    gc.collect()

    for region, cells_set in region_cells.items():
        update_callback(10 + 70 * (list(valid_regions).index(region) / len(valid_regions)), f"Processing region {region}...")
        _process_object_trees(obj_folder, region, cells_set, out_dir, fmt_csv, fmt_xlsx)

    update_callback(100, "Generation completed.")

def generate_object_trees_cluster(obj_folder, batchfile_path, selected_clusters, out_dir, fmt_csv, fmt_xlsx, update_callback, export_mode):
    if not os.path.exists(obj_folder): return
    update_callback(5, "Processing OBJECT TREES - CLUSTER...")

    df_batch = read_safe_file(batchfile_path)

    cells_series = clean_series(df_batch.iloc[:, 1])
    clusters_series = clean_series(df_batch.iloc[:, 7])
    df_map = pd.DataFrame({'CELL': cells_series, 'CLUSTER': clusters_series})
    
    del df_batch
    gc.collect()
    
    df_map = df_map[(df_map['CELL'] != '') & (df_map['CLUSTER'] != '')]
    cell_cluster_map = df_map.set_index('CELL')['CLUSTER'].to_dict()
    del df_map

    targets = []
    if export_mode == "batch":
        cells_in_batch = set()
        for cluster in selected_clusters:
            c_up = str(cluster).strip().upper()
            cells_in_batch.update({c for c, clust in cell_cluster_map.items() if clust == c_up})
        if cells_in_batch: targets = [("BATCH", cells_in_batch)]
    else:
        for cluster in selected_clusters:
            c_up = str(cluster).strip().upper()
            cells = {c for c, clust in cell_cluster_map.items() if clust == c_up}
            if cells: targets.append((c_up, cells))

    for target_name, cells_to_match in targets:
        _process_object_trees(obj_folder, target_name, cells_to_match, out_dir, fmt_csv, fmt_xlsx)

    update_callback(100, "Generation completed.")

def generate_object_trees_site(obj_folder, batchfile_path, selected_unicos, out_dir, fmt_csv, fmt_xlsx, update_callback, export_mode):
    if not os.path.exists(obj_folder): return
    update_callback(5, "Processing OBJECT TREES - SITE...")

    df_batch = read_safe_file(batchfile_path)

    c_unico = buscar_columna_inteligente(df_batch.columns, ['UNICO'])
    if not c_unico: raise KeyError("Batchfile must contain UNICO column.")

    cells_series = clean_series(df_batch.iloc[:, 1])
    unicos_series = clean_series(df_batch[c_unico])

    df_map = pd.DataFrame({'CELL': cells_series, 'UNICO': unicos_series})
    del df_batch
    gc.collect()
    
    df_map = df_map[(df_map['CELL'] != '') & (df_map['UNICO'] != '')]
    unico_cell_map = df_map.groupby('UNICO')['CELL'].apply(set).to_dict()
    del df_map

    targets = []
    if export_mode == "batch":
        cells_in_batch = set()
        for unico in selected_unicos:
            u_up = str(unico).strip().upper()
            cells_in_batch.update(unico_cell_map.get(u_up, set()))
        if cells_in_batch: targets = [("BATCH", cells_in_batch)]
    else:
        for unico in selected_unicos:
            u_up = str(unico).strip().upper()
            cells = unico_cell_map.get(u_up, set())
            if cells: targets.append((u_up, cells))

    for target_name, cells_to_match in targets:
        _process_object_trees(obj_folder, target_name, cells_to_match, out_dir, fmt_csv, fmt_xlsx)

    update_callback(100, "Generation completed.")
