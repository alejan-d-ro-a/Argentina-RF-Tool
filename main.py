import os
import time
import traceback
import threading
import gc
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import webbrowser
import shutil

from config_manager import ConfigManager
from ui_components import (
    COLOR_BG, COLOR_TEXT, COLOR_BTN_DARK, COLOR_BTN_CANCEL, COLOR_BTN_CANCEL_HOVER, 
    FONT_NORMAL, FONT_TITLE, FONT_LARGE, FONT_DYNAMIC, RoundedButton, RoundedProgressBar, 
    TextAnimWrapper, CustomMessageBox, apply_rounded_corners, 
    set_appwindow, resource_path, ClusterSelectionDialog, DatePickerDialog
)
from rf_data_engine import RFDataEngine
import legacy_tools as legacy

# IMPORTACIÓN DEL ACTUALIZADOR
from auto_updater import check_for_updates

class ProcessCancelledException(Exception): pass

class RFProcessorApp:
    def __init__(self, root, scale_factor=1.0):
        self.root = root 
        self.SF = scale_factor
        self.config = ConfigManager.load_config()
        self.engine = RFDataEngine()
        
        self.cancel_requested = False
        self.is_running = False
        self.active_export = None
        self.ept_subtype = None
        self.batch_subtype = None
        self.itsc_subtype = None
        self.object_subtype = None
        self._slide_after_id = None
        self._slide_frame = None
        self.subtype_frame = None
        self.inputs_container = None
        
        self.selected_clusters = set()
        self.selected_obj_clusters = set()
        self.selected_obj_unicos = set()

        self.obj_cluster_mode = tk.StringVar(value="individual")
        self.obj_site_mode = tk.StringVar(value="individual")

        self.root.withdraw()
        self.root.overrideredirect(True)
        self.root.title("Argentina RF Tool")
        
        w, h = int(1200 * self.SF), int(800 * self.SF)
        pw, ph = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        w, h = min(w, pw - 40), min(h, ph - 80)
        x, y = (pw - w) // 2, (ph - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        self.root.configure(bg=COLOR_BG)
        
        try: self.root.iconbitmap(resource_path("argentina.ico"))
        except: pass
        set_appwindow(self.root)

        self._build_title_bar()

        self.dynamic_label = tk.Label(self.root, text="Select an option:", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_DYNAMIC, anchor="center")
        self.dynamic_label.pack(fill="x", padx=int(10 * self.SF), pady=(int(8 * self.SF), int(4 * self.SF)))

        self.p_mvs = tk.StringVar(value=self.config.get('p_mvs', ''))
        self.p_tp = tk.StringVar(value=self.config.get('p_tp', ''))
        self.p_nics = tk.StringVar(value=self.config.get('p_nics', ''))
        self.p_clus = tk.StringVar(value=self.config.get('p_clus', ''))
        self.p_ran = tk.StringVar(value=self.config.get('p_ran', ''))
        self.p_obj = tk.StringVar(value=self.config.get('p_obj', ''))
        self.p_rnd = tk.StringVar(value=self.config.get('p_rnd', ''))
        self.p_dat = tk.StringVar(value=self.config.get('p_dat', ''))
        self.p_ctrl = tk.StringVar(value=self.config.get('p_ctrl', ''))
        self.p_prog = tk.StringVar(value=self.config.get('p_prog', ''))
        self.p_cluster_control = tk.StringVar(value=self.config.get('p_cluster_control', ''))
        self.p_clusterizacion = tk.StringVar(value=self.config.get('p_clusterizacion', ''))
        self.p_rnd_folder = tk.StringVar(value=self.config.get('p_rnd_folder', ''))
        self.p_master_ept = tk.StringVar(value=self.config.get('p_master_ept', ''))
        self.p_itsc_cluster_control = tk.StringVar(value=self.config.get('p_itsc_cluster_control', ''))
        self.p_out = tk.StringVar(value=self.config.get('p_out', ''))
        self.p_obj_trees_folder = tk.StringVar(value=self.config.get('p_obj_trees_folder', ''))
        self.p_obj_batchfile = tk.StringVar(value=self.config.get('p_obj_batchfile', ''))
        self.p_batch_old = tk.StringVar(value=self.config.get('p_batch_old', ''))
        self.p_batch_new = tk.StringVar(value=self.config.get('p_batch_new', ''))
        self.p_it_rnd = tk.StringVar(value=self.config.get('p_it_rnd', ''))
        self.p_it_ctrl = tk.StringVar(value=self.config.get('p_it_ctrl', ''))
        self.p_it_cambios = tk.StringVar(value=self.config.get('p_it_cambios', ''))
        self.p_dv_2g = tk.StringVar(value=self.config.get('p_dv_2g', ''))
        self.p_dv_3g = tk.StringVar(value=self.config.get('p_dv_3g', ''))
        self.p_dv_4g = tk.StringVar(value=self.config.get('p_dv_4g', ''))
        self.p_dv_5g = tk.StringVar(value=self.config.get('p_dv_5g', ''))
        self.p_dv_batch = tk.StringVar(value=self.config.get('p_dv_batch', ''))
        self.p_finder_batch = tk.StringVar(value=self.config.get('p_finder_batch', ''))
        
        self.date_elec = tk.StringVar()
        self.date_mech = tk.StringVar()
        
        self.query_var = tk.StringVar()
        self.result_var = tk.StringVar()

        self.vars_list = [self.p_mvs, self.p_tp, self.p_nics, self.p_clus, self.p_ran, self.p_obj, self.p_rnd,
                          self.p_dat, self.p_ctrl, self.p_prog, self.p_cluster_control, self.p_clusterizacion,
                          self.p_rnd_folder, self.p_master_ept, self.p_itsc_cluster_control, self.p_out,
                          self.p_obj_trees_folder, self.p_obj_batchfile, self.p_batch_old, self.p_batch_new,
                          self.p_it_rnd, self.p_it_ctrl, self.p_it_cambios,
                          self.p_dv_2g, self.p_dv_3g, self.p_dv_4g, self.p_dv_5g, self.p_dv_batch, self.p_finder_batch]
                          
        for v in self.vars_list: v.trace_add("write", lambda *args: self.check_ready_state())

        self.build_ui()
        self.root.deiconify()
        self.root.after(50, lambda: apply_rounded_corners(self.root, int(20 * self.SF)))

    def _build_title_bar(self):
        self.title_bar = tk.Frame(self.root, bg=COLOR_BTN_DARK, relief="raised", bd=0, height=int(32 * self.SF))
        self.title_bar.pack(expand=0, fill="x")
        self.title_bar.pack_propagate(False)
        self.title_bar.bind("<ButtonPress-1>", self.start_move)
        self.title_bar.bind("<B1-Motion>", self.do_move)
        self.title_label = tk.Label(self.title_bar, text="Argentina RF Tool", bg=COLOR_BTN_DARK, fg="white", font=FONT_TITLE)
        self.title_label.place(relx=0.5, rely=0.5, anchor="center")
        
        self.close_btn = tk.Button(self.title_bar, text="✕", bg=COLOR_BTN_DARK, fg="white", bd=0, command=self.root.destroy, font=FONT_TITLE, width=4, activebackground="#A00000", activeforeground="white", cursor="hand2")
        self.close_btn.pack(side="right", fill="y")
        self.close_btn.bind("<Enter>", lambda e: self.close_btn.config(bg="#D00000"))
        self.close_btn.bind("<Leave>", lambda e: self.close_btn.config(bg=COLOR_BTN_DARK))
        
        self.minimize_btn = tk.Button(self.title_bar, text="🗕", bg=COLOR_BTN_DARK, fg="white", bd=0, command=self.minimize_window, font=FONT_TITLE, width=4, cursor="hand2")
        self.minimize_btn.pack(side="right", fill="y")
        self.minimize_btn.bind("<Enter>", lambda e: self.minimize_btn.config(bg="#00509E"))
        self.minimize_btn.bind("<Leave>", lambda e: self.minimize_btn.config(bg=COLOR_BTN_DARK))
        self.root.bind("<Map>", self.frame_mapped)

    def start_move(self, event): self.x, self.y = event.x, event.y
    def do_move(self, event):
        x, y = self.root.winfo_x() + (event.x - self.x), self.root.winfo_y() + (event.y - self.y)
        self.root.geometry(f"+{x}+{y}")
        
    def minimize_window(self):
        self.root.overrideredirect(False)
        self.root.iconify()
        
    def frame_mapped(self, event):
        if str(event.widget) == str(self.root):
            if not self.root.overrideredirect():
                self.root.overrideredirect(True)
                set_appwindow(self.root)
                self.root.after(50, lambda: apply_rounded_corners(self.root, int(20 * self.SF)))

    def build_ui(self):
        f_main = tk.Frame(self.root, bg=COLOR_BG)
        f_main.pack(fill="both", expand=True, padx=int(10 * self.SF), pady=int(5 * self.SF))
        f_main.grid_columnconfigure(0, weight=4, uniform="cols")
        f_main.grid_columnconfigure(1, weight=6, uniform="cols")
        f_main.grid_rowconfigure(0, weight=1)

        self.f_left = tk.Frame(f_main, bg=COLOR_BG)
        self.f_left.grid(row=0, column=0, sticky="nsew", padx=(0, int(5 * self.SF)))
        self.f_left.grid_rowconfigure(0, weight=1)
        self.f_left.grid_columnconfigure(0, weight=1)

        self.f_left_center = tk.Frame(self.f_left, bg=COLOR_BG)
        self.f_left_center.grid(row=0, column=0)

        btn_w, btn_h = 190, 60

        self.btn_download_inputs = RoundedButton(self.f_left_center, text="DOWNLOAD INPUTS", command=self.open_download_link, width=(btn_w * 2 + 16), height=btn_h, sf=self.SF, bg="#28A745", hover_bg="#218838", auto_toggle=False)
        self.btn_download_inputs.grid(row=0, column=0, columnspan=2, padx=8, pady=(8, 15))

        self.btn_ept = RoundedButton(self.f_left_center, text="EXPORT EPT", command=lambda: self.on_export_select('ept'), width=btn_w, height=btn_h, is_toggle=True, sf=self.SF, auto_toggle=False)
        self.btn_ept.grid(row=1, column=0, padx=8, pady=8)
        self.btn_batch = RoundedButton(self.f_left_center, text="EXPORT BATCHFILE", command=lambda: self.on_export_select('batch'), width=btn_w, height=btn_h, is_toggle=True, sf=self.SF, auto_toggle=False)
        self.btn_batch.grid(row=1, column=1, padx=8, pady=8)
        
        self.btn_cluster_friendly = RoundedButton(self.f_left_center, text="EXPORT CLUSTER CONTROL FRIENDLY", command=lambda: self.on_export_select('cluster_friendly'), width=btn_w, height=btn_h, is_toggle=True, sf=self.SF, auto_toggle=False)
        self.btn_cluster_friendly.grid(row=2, column=0, padx=8, pady=8)
        self.btn_itsc = RoundedButton(self.f_left_center, text="EXPORT ITSC OBJECT TEMPLATE", command=lambda: self.on_export_select('itsc'), width=btn_w, height=btn_h, is_toggle=True, sf=self.SF, auto_toggle=False)
        self.btn_itsc.grid(row=2, column=1, padx=8, pady=8)
        
        self.btn_object_trees = RoundedButton(self.f_left_center, text="EXPORT OBJECT TREES", command=lambda: self.on_export_select('object_trees'), width=btn_w, height=btn_h, is_toggle=True, sf=self.SF, auto_toggle=False)
        self.btn_object_trees.grid(row=3, column=0, padx=8, pady=8)
        self.btn_cluster_finder = RoundedButton(self.f_left_center, text="CLUSTER FINDER", command=lambda: self.on_export_select('cluster_finder'), width=btn_w, height=btn_h, is_toggle=True, sf=self.SF, auto_toggle=False)
        self.btn_cluster_finder.grid(row=3, column=1, padx=8, pady=8)

        self.btn_it_final_report = RoundedButton(self.f_left_center, text="EXPORT IT FINAL REPORT", command=lambda: self.on_export_select('it_final_report'), width=btn_w, height=btn_h, is_toggle=True, sf=self.SF, auto_toggle=False)
        self.btn_it_final_report.grid(row=4, column=0, padx=8, pady=8)
        self.btn_data_validation = RoundedButton(self.f_left_center, text="DATA VALIDATION", command=lambda: self.on_export_select('data_validation'), width=btn_w, height=btn_h, is_toggle=True, sf=self.SF, auto_toggle=False)
        self.btn_data_validation.grid(row=4, column=1, padx=8, pady=8)

        self.f_right = tk.Frame(f_main, bg=COLOR_BG)
        self.f_right.grid(row=0, column=1, sticky="nsew", padx=(int(5 * self.SF), 0))
        self.lbl_right_placeholder = tk.Label(self.f_right, text="Select an export type to configure inputs", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
        self.lbl_right_placeholder.place(relx=0.5, rely=0.5, anchor="center")

        f_export = tk.Frame(self.root, bg=COLOR_BG)
        f_export.pack(fill="x", padx=int(10 * self.SF), pady=int(5 * self.SF))
        tk.Label(f_export, bg=COLOR_BG).pack(side="left", expand=True)
        self.lbl_out = tk.Label(f_export, text="EXPORT DESTINATION (*):", anchor="e", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
        self.lbl_out.pack(side="left")
        self.ent_out = tk.Entry(f_export, textvariable=self.p_out, width=45, font=FONT_NORMAL)
        self.ent_out.pack(side="left", padx=int(5 * self.SF))
        self.btn_out = RoundedButton(f_export, text="📁", command=lambda: self.select_folder(self.p_out, 'p_out'), width=45, height=30, radius=8, sf=self.SF)
        self.btn_out.pack(side="left")
        tk.Label(f_export, bg=COLOR_BG).pack(side="left", expand=True)

        self.lbl_inst_fmt = tk.Label(self.root, text="2. Select export format:", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
        self.lbl_inst_fmt.pack(pady=int(5 * self.SF))
        f_fmt = tk.Frame(self.root, bg=COLOR_BG)
        f_fmt.pack(fill="x")
        tk.Label(f_fmt, bg=COLOR_BG).pack(side="left", expand=True)
        self.btn_fmt_csv = RoundedButton(f_fmt, "EXPORT AS CSV", self.check_ready_state, width=420, height=45, is_toggle=True, sf=self.SF)
        self.btn_fmt_csv.pack(side="left", padx=int(10 * self.SF))
        self.btn_fmt_xlsx = RoundedButton(f_fmt, "EXPORT AS XLSX", self.check_ready_state, width=420, height=45, is_toggle=True, sf=self.SF)
        self.btn_fmt_xlsx.pack(side="left", padx=int(10 * self.SF))
        tk.Label(f_fmt, bg=COLOR_BG).pack(side="left", expand=True)

        self.loader = RoundedProgressBar(self.root, width=800, height=35, sf=self.SF)
        self.loader.pack(pady=(int(5 * self.SF), int(2 * self.SF)))
        self.lbl_status = tk.Label(self.root, text="Waiting for instructions...", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_TITLE)
        self.lbl_status.pack()

        f_actions = tk.Frame(self.root, bg=COLOR_BG)
        f_actions.pack(pady=int(2 * self.SF))
        tk.Label(f_actions, bg=COLOR_BG).pack(side="left", expand=True)
        self.btn_run = RoundedButton(f_actions, "START PROCESS", self.start_process, width=380, height=45, sf=self.SF)
        self.btn_run.pack(side="left", padx=int(10 * self.SF))
        self.btn_cancel = RoundedButton(f_actions, "EXIT", self.root.destroy, width=380, height=45, bg=COLOR_BTN_CANCEL, hover_bg=COLOR_BTN_CANCEL_HOVER, sf=self.SF)
        self.btn_cancel.pack(side="left", padx=int(10 * self.SF))
        tk.Label(f_actions, bg=COLOR_BG).pack(side="left", expand=True)

        self.lbl_credits = tk.Label(self.root, text="Developed by: Manuel Alejandro Núñez Ayala | ID: m00963399", bg=COLOR_BG, fg="gray", font=("Comic Sans MS", 8))
        self.lbl_credits.pack(side="bottom", pady=int(2 * self.SF))

    def open_download_link(self):
        def _open():
            try:
                webbrowser.open_new_tab("https://onebox.huawei.com/#eSpaceGroupFile/1/513/19435973")
            except Exception as e:
                print(f"Error opening browser: {e}")
        threading.Thread(target=_open, daemon=True).start()

    def _typewriter_animate_all(self, targets, delay=25):
        if hasattr(self, '_title_animation_id') and self._title_animation_id:
            self.root.after_cancel(self._title_animation_id)
        data = [{'wrap': w, 'curr': str(w.get()), 'final': str(f)} for w, f in targets]
        for item in data:
            if len(item['curr']) > 0:
                item['curr'] = ""
                item['wrap'].set(item['curr'])
        def step_type(idx):
            done = True
            for item in data:
                if len(item['curr']) < len(item['final']):
                    item['curr'] = item['final'][:len(item['curr'])+1]
                    item['wrap'].set(item['curr'])
                    done = False
            if not done: self._title_animation_id = self.root.after(delay, lambda: step_type(idx+1))
            else: self._title_animation_id = None
        self._title_animation_id = self.root.after(delay, lambda: step_type(0))

    def _animate_label_text(self, old_text, new_text):
        wrapper = TextAnimWrapper(self.dynamic_label, "label")
        self.dynamic_label.config(text=old_text)
        self._typewriter_animate_all([(wrapper, new_text)], delay=20)

    def select_file(self, var, var_name):
        path = filedialog.askopenfilename()
        if path:
            var.set(path)
            self.config[var_name] = path
            ConfigManager.save_config(self.config)

    def select_folder(self, var, var_name):
        path = filedialog.askdirectory()
        if path:
            var.set(path)
            self.config[var_name] = path
            ConfigManager.save_config(self.config)
            
    def select_date(self, var):
        dialog = DatePickerDialog(self.root, self.SF)
        if dialog.result:
            var.set(dialog.result)
            
    def download_template(self):
        template_name = "Cluster Final Report_PRE_vs_POST TEMPLATE.xlsx"
        src = resource_path(template_name)
        if not os.path.exists(src):
            CustomMessageBox(self.root, "Error", f"Template file not found.\nEnsure '{template_name}' is bundled in the executable.", "error", self.SF)
            return
            
        dest = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            initialfile=template_name, 
            title="Save Blank Template As", 
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if dest:
            try:
                shutil.copy(src, dest)
                CustomMessageBox(self.root, "Success", "Template downloaded successfully.", "info", self.SF)
            except Exception as e:
                CustomMessageBox(self.root, "Error", f"Failed to save template:\n{str(e)}", "error", self.SF)

    def create_input_row(self, parent, text_lbl, var, var_name, is_folder, row_idx):
        lbl = tk.Label(parent, text=text_lbl, anchor="e", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
        lbl.grid(row=row_idx, column=0, sticky="e", padx=(0, 5), pady=5)
        ent = tk.Entry(parent, textvariable=var, width=38, font=FONT_NORMAL)
        ent.grid(row=row_idx, column=1, sticky="w", padx=2, pady=5)
        cmd = (lambda: self.select_folder(var, var_name)) if is_folder else (lambda: self.select_file(var, var_name))
        btn = RoundedButton(parent, text="📁", command=cmd, width=40, height=28, radius=8, sf=self.SF)
        btn.grid(row=row_idx, column=2, sticky="w", padx=2, pady=5)

    def show_export_inputs(self, export_type):
        if self._slide_after_id:
            self.root.after_cancel(self._slide_after_id)
            self._slide_after_id = None
        if self._slide_frame and self._slide_frame.winfo_exists():
            old_frame = self._slide_frame
            self._slide_frame = None
            self.lbl_right_placeholder.place_forget()
            self._slide_out(old_frame, lambda: self._create_and_animate_new_frame(export_type))
        else:
            self.lbl_right_placeholder.place_forget()
            self._create_and_animate_new_frame(export_type)

    def _slide_out(self, frame, callback):
        steps = 10; delay = 20; start = 0.0; end = -1.0; step = (end - start) / steps
        def anim(i):
            if i > steps:
                frame.place_forget(); frame.destroy(); callback()
                return
            frame.place_configure(rely=start + step * i)
            self._slide_after_id = self.root.after(delay, lambda: anim(i+1))
        anim(0)

    def _slide_in(self, frame):
        steps = 10; delay = 20; start = 1.0; end = 0.0; step = (end - start) / steps
        def anim(i):
            if i > steps:
                frame.place_configure(rely=end); self._slide_after_id = None
                return
            frame.place_configure(rely=start + step * i)
            self._slide_after_id = self.root.after(delay, lambda: anim(i+1))
        anim(0)

    def _create_and_animate_new_frame(self, export_type):
        self._replace_inputs(export_type)
        self._slide_frame.place(relx=0, rely=1.0, relwidth=1.0, relheight=1.0)
        self._slide_in(self._slide_frame)

    def _create_ept_subtype_selector(self, parent):
        frame = tk.Frame(parent, bg=COLOR_BG)
        self.btn_ept_master = RoundedButton(frame, text="MASTER EPT", command=lambda: self.on_ept_subtype_select('master'), width=180, height=45, is_toggle=True, default_active=(self.ept_subtype == 'master'), sf=self.SF, auto_toggle=False)
        self.btn_ept_master.pack(side="left", padx=5)
        self.btn_ept_assistant = RoundedButton(frame, text="EPT ASSISTANT", command=lambda: self.on_ept_subtype_select('assistant'), width=180, height=45, is_toggle=True, default_active=(self.ept_subtype == 'assistant'), sf=self.SF, auto_toggle=False)
        self.btn_ept_assistant.pack(side="left", padx=5)
        self.btn_ept_acp = RoundedButton(frame, text="EPT ACP", command=lambda: self.on_ept_subtype_select('acp'), width=180, height=45, is_toggle=True, default_active=(self.ept_subtype == 'acp'), sf=self.SF, auto_toggle=False)
        self.btn_ept_acp.pack(side="left", padx=5)
        return frame

    def _create_batch_subtype_selector(self, parent):
        frame = tk.Frame(parent, bg=COLOR_BG)
        self.btn_batch_master = RoundedButton(frame, text="MASTERBATCH", command=lambda: self.on_batch_subtype_select('master'), width=180, height=45, is_toggle=True, default_active=(self.batch_subtype == 'master'), sf=self.SF, auto_toggle=False)
        self.btn_batch_master.pack(side="left", padx=5)
        self.btn_batch_audit = RoundedButton(frame, text="AUDIT", command=lambda: self.on_batch_subtype_select('audit'), width=180, height=45, is_toggle=True, default_active=(self.batch_subtype == 'audit'), sf=self.SF, auto_toggle=False)
        self.btn_batch_audit.pack(side="left", padx=5)
        return frame

    def _create_itsc_subtype_selector(self, parent):
        frame = tk.Frame(parent, bg=COLOR_BG)
        self.btn_itsc_rsh = RoundedButton(frame, text="RSH OBJECTS", command=lambda: self.on_itsc_subtype_select('rsh'), width=200, height=45, is_toggle=True, default_active=(self.itsc_subtype == 'rsh'), sf=self.SF, auto_toggle=False)
        self.btn_itsc_rsh.pack(side="left", padx=10)
        self.btn_itsc_swap = RoundedButton(frame, text="SWAP OBJECTS", command=lambda: self.on_itsc_subtype_select('swap'), width=200, height=45, is_toggle=True, default_active=(self.itsc_subtype == 'swap'), sf=self.SF, auto_toggle=False)
        self.btn_itsc_swap.pack(side="left", padx=10)
        return frame

    def _create_obj_subtype_selector(self, parent):
        frame = tk.Frame(parent, bg=COLOR_BG)
        self.btn_obj_cluster = RoundedButton(frame, text="CLUSTER", command=lambda: self.on_object_subtype_select('cluster'), width=180, height=45, is_toggle=True, default_active=(self.object_subtype == 'cluster'), sf=self.SF, auto_toggle=False)
        self.btn_obj_cluster.pack(side="left", padx=5)
        self.btn_obj_region = RoundedButton(frame, text="REGIONS", command=lambda: self.on_object_subtype_select('region'), width=180, height=45, is_toggle=True, default_active=(self.object_subtype == 'region'), sf=self.SF, auto_toggle=False)
        self.btn_obj_region.pack(side="left", padx=5)
        self.btn_obj_site = RoundedButton(frame, text="SITE", command=lambda: self.on_object_subtype_select('site'), width=180, height=45, is_toggle=True, default_active=(self.object_subtype == 'site'), sf=self.SF, auto_toggle=False)
        self.btn_obj_site.pack(side="left", padx=5)
        return frame

    def on_export_select(self, export_type):
        if self.active_export == export_type: return
        old_title = self.dynamic_label.cget("text")
        self.active_export = export_type
        
        for btn_name, value in [('btn_ept', 'ept'), ('btn_batch', 'batch'), ('btn_cluster_friendly', 'cluster_friendly'), ('btn_itsc', 'itsc'), ('btn_object_trees', 'object_trees'), ('btn_cluster_finder', 'cluster_finder'), ('btn_it_final_report', 'it_final_report'), ('btn_data_validation', 'data_validation')]:
            btn = getattr(self, btn_name)
            btn.active = (value == export_type)
            btn.update_colors()
            btn.animate_transition(btn.current_bg, btn.current_fg, btn.pad_normal)
        
        titles = {
            "ept": "EXPORT EPT", "batch": "EXPORT BATCHFILE", "cluster_friendly": "EXPORT CLUSTER CONTROL FRIENDLY",
            "itsc": "EXPORT ITSC OBJECT TEMPLATE", "object_trees": "EXPORT OBJECT TREES",
            "cluster_finder": "INSTANT CLUSTER FINDER", "it_final_report": "EXPORT IT FINAL REPORT", "data_validation": "DATA VALIDATION"
        }
        self._animate_label_text(old_title, titles.get(export_type, ""))
        self.show_export_inputs(export_type)
        self.check_ready_state()

    def on_ept_subtype_select(self, subtype):
        if self.ept_subtype == subtype: return
        self.ept_subtype = subtype
        self.show_export_inputs('ept')

    def on_batch_subtype_select(self, subtype):
        if self.batch_subtype == subtype: return
        self.batch_subtype = subtype
        self.show_export_inputs('batch')

    def on_itsc_subtype_select(self, subtype):
        if self.itsc_subtype == subtype: return
        self.itsc_subtype = subtype
        self.show_export_inputs('itsc')

    def on_object_subtype_select(self, subtype):
        if self.object_subtype == subtype: return
        self.object_subtype = subtype
        self.show_export_inputs('object_trees')

    def _replace_inputs(self, export_type):
        if self._slide_frame and self._slide_frame.winfo_exists():
            self._slide_frame.destroy(); self._slide_frame = None

        self._slide_frame = tk.Frame(self.f_right, bg=COLOR_BG)
        top_section = tk.Frame(self._slide_frame, bg=COLOR_BG, height=90)
        top_section.pack(side="top", fill="x")
        
        center_frame = tk.Frame(self._slide_frame, bg=COLOR_BG)
        center_frame.pack(expand=True, anchor="center")
        self.inputs_container = tk.Frame(center_frame, bg=COLOR_BG)
        self.inputs_container.pack()

        if export_type == 'ept':
            self.subtype_frame = self._create_ept_subtype_selector(top_section)
            self.subtype_frame.place(relx=0.5, rely=0.5, anchor="center")
            if self.ept_subtype == 'master':
                self.create_input_row(self.inputs_container, 'MVS EPT (*):', self.p_mvs, 'p_mvs', False, 0)
                self.create_input_row(self.inputs_container, 'TP EPT (*):', self.p_tp, 'p_tp', False, 1)
                self.create_input_row(self.inputs_container, 'NICS (*):', self.p_nics, 'p_nics', True, 2)
            elif self.ept_subtype in ('assistant', 'acp'):
                self.create_input_row(self.inputs_container, 'MASTER EPT (*):', self.p_master_ept, 'p_master_ept', False, 0)
                tk.Label(self.inputs_container, text="* Clusters will be selected upon generating.", bg=COLOR_BG, fg="gray").grid(row=1, column=0, columnspan=3, pady=10)

        elif export_type == 'batch':
            self.subtype_frame = self._create_batch_subtype_selector(top_section)
            self.subtype_frame.place(relx=0.5, rely=0.5, anchor="center")
            if self.batch_subtype == 'master':
                self.create_input_row(self.inputs_container, 'MASTER EPT (*):', self.p_master_ept, 'p_master_ept', False, 0)
                self.create_input_row(self.inputs_container, 'RAN REPORTS (*):', self.p_ran, 'p_ran', True, 1)
                self.create_input_row(self.inputs_container, 'OBJECT TREES (*):', self.p_obj, 'p_obj', True, 2)
                self.create_input_row(self.inputs_container, 'RND SWAP/RSH (*):', self.p_rnd, 'p_rnd', True, 3)
                self.create_input_row(self.inputs_container, 'CLUSTER CONTROL:', self.p_ctrl, 'p_ctrl', False, 4)
                self.create_input_row(self.inputs_container, 'GENERAL DATA (Opt):', self.p_dat, 'p_dat', False, 5)
                self.create_input_row(self.inputs_container, 'PROGRAMMING (Opt):', self.p_prog, 'p_prog', False, 6)
                self.create_input_row(self.inputs_container, 'CLUSTER MAPPING:', self.p_clus, 'p_clus', False, 7)
            elif self.batch_subtype == 'audit':
                self.create_input_row(self.inputs_container, 'MASTERBATCH OLD (*):', self.p_batch_old, 'p_batch_old', False, 0)
                self.create_input_row(self.inputs_container, 'MASTERBATCH NEW (*):', self.p_batch_new, 'p_batch_new', False, 1)

        elif export_type == 'cluster_friendly':
            self.create_input_row(self.inputs_container, 'CLUSTER CONTROL (*):', self.p_cluster_control, 'p_cluster_control', False, 0)
            self.create_input_row(self.inputs_container, 'CLUSTER MAPPING (*):', self.p_clusterizacion, 'p_clusterizacion', False, 1)

        elif export_type == 'itsc':
            self.subtype_frame = self._create_itsc_subtype_selector(top_section)
            self.subtype_frame.place(relx=0.5, rely=0.5, anchor="center")
            if self.itsc_subtype == 'rsh':
                self.create_input_row(self.inputs_container, 'CLUSTER CONTROL (*):', self.p_itsc_cluster_control, 'p_itsc_cluster_control', False, 0)
                tk.Label(self.inputs_container, text="* Clusters will be selected upon generating.", bg=COLOR_BG, fg="gray").grid(row=1, column=0, columnspan=3, pady=10)
            elif self.itsc_subtype == 'swap':
                lbl_site = tk.Label(self.inputs_container, text="Site (*)", anchor="w", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
                lbl_site.grid(row=0, column=0, sticky="w", padx=2, pady=2)
                lbl_tech = tk.Label(self.inputs_container, text="Tech (*)", anchor="w", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
                lbl_tech.grid(row=0, column=1, sticky="w", padx=2, pady=2)
                self.swap_text = tk.Text(self.inputs_container, height=6, width=45, font=("Courier New", 10), bg="white", fg=COLOR_TEXT)
                self.swap_text.grid(row=1, column=0, columnspan=3, sticky="we", padx=2, pady=5)
                tk.Label(self.inputs_container, text="Format: Site Tech (e.g., XX001 2G_3G)", bg=COLOR_BG, fg="gray").grid(row=2, column=0, columnspan=3, pady=5)

        elif export_type == 'object_trees':
            self.subtype_frame = self._create_obj_subtype_selector(top_section)
            self.subtype_frame.place(relx=0.5, rely=0.5, anchor="center")
            if self.object_subtype in ('cluster', 'site'):
                self.create_input_row(self.inputs_container, 'OBJECT TREES (*):', self.p_obj_trees_folder, 'p_obj_trees_folder', True, 0)
                self.create_input_row(self.inputs_container, 'BATCHFILE (*):', self.p_obj_batchfile, 'p_obj_batchfile', False, 1)
                
                mode_frame = tk.Frame(self.inputs_container, bg=COLOR_BG)
                mode_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(10,5))
                mode_frame.columnconfigure(0, weight=1); mode_frame.columnconfigure(1, weight=1); mode_frame.columnconfigure(2, weight=1)
                
                if self.object_subtype == 'cluster':
                    self.rb_cluster_individual = RoundedButton(mode_frame, text="Individual", command=lambda: self.obj_cluster_mode.set("individual"), width=140, height=35, radius=10, is_toggle=True, default_active=(self.obj_cluster_mode.get() == "individual"), sf=self.SF, auto_toggle=False)
                    self.rb_cluster_individual.grid(row=0, column=1, padx=5)
                    self.rb_cluster_batch = RoundedButton(mode_frame, text="Batch", command=lambda: self.obj_cluster_mode.set("batch"), width=140, height=35, radius=10, is_toggle=True, default_active=(self.obj_cluster_mode.get() == "batch"), sf=self.SF, auto_toggle=False)
                    self.rb_cluster_batch.grid(row=0, column=2, padx=5)
                    
                    def update_cluster_mode_buttons(*args):
                        mode = self.obj_cluster_mode.get()
                        self.rb_cluster_individual.active = (mode == "individual")
                        self.rb_cluster_batch.active = (mode == "batch")
                        self.rb_cluster_individual.update_colors()
                        self.rb_cluster_batch.update_colors()
                        self.rb_cluster_individual.animate_transition(
                            self.rb_cluster_individual.current_bg,
                            self.rb_cluster_individual.current_fg,
                            self.rb_cluster_individual.pad_normal
                        )
                        self.rb_cluster_batch.animate_transition(
                            self.rb_cluster_batch.current_bg,
                            self.rb_cluster_batch.current_fg,
                            self.rb_cluster_batch.pad_normal
                        )
                    
                    self.obj_cluster_mode.trace_add("write", update_cluster_mode_buttons)
                    update_cluster_mode_buttons()
                    
                    tk.Label(self.inputs_container, text="* Clusters will be selected upon generating.", bg=COLOR_BG, fg="gray").grid(row=3, column=0, columnspan=3, pady=10)
                else:  
                    self.rb_site_individual = RoundedButton(mode_frame, text="Individual", command=lambda: self.obj_site_mode.set("individual"), width=140, height=35, radius=10, is_toggle=True, default_active=(self.obj_site_mode.get() == "individual"), sf=self.SF, auto_toggle=False)
                    self.rb_site_individual.grid(row=0, column=1, padx=5)
                    self.rb_site_batch = RoundedButton(mode_frame, text="Batch", command=lambda: self.obj_site_mode.set("batch"), width=140, height=35, radius=10, is_toggle=True, default_active=(self.obj_site_mode.get() == "batch"), sf=self.SF, auto_toggle=False)
                    self.rb_site_batch.grid(row=0, column=2, padx=5)
                    
                    def update_site_mode_buttons(*args):
                        mode = self.obj_site_mode.get()
                        self.rb_site_individual.active = (mode == "individual")
                        self.rb_site_batch.active = (mode == "batch")
                        self.rb_site_individual.update_colors()
                        self.rb_site_batch.update_colors()
                        self.rb_site_individual.animate_transition(
                            self.rb_site_individual.current_bg,
                            self.rb_site_individual.current_fg,
                            self.rb_site_individual.pad_normal
                        )
                        self.rb_site_batch.animate_transition(
                            self.rb_site_batch.current_bg,
                            self.rb_site_batch.current_fg,
                            self.rb_site_batch.pad_normal
                        )
                    
                    self.obj_site_mode.trace_add("write", update_site_mode_buttons)
                    update_site_mode_buttons()
                    
                    tk.Label(self.inputs_container, text="* Sites (UNICOs) will be selected upon generating.", bg=COLOR_BG, fg="gray").grid(row=3, column=0, columnspan=3, pady=10)
            elif self.object_subtype == 'region':
                self.create_input_row(self.inputs_container, 'OBJECT TREES (*):', self.p_obj_trees_folder, 'p_obj_trees_folder', True, 0)
                self.create_input_row(self.inputs_container, 'BATCHFILE (*):', self.p_obj_batchfile, 'p_obj_batchfile', False, 1)

        elif export_type == 'cluster_finder':
            self.create_input_row(self.inputs_container, 'MASTERBATCH DB:', self.p_finder_batch, 'p_finder_batch', False, 0)
            btn_ram = RoundedButton(self.inputs_container, text="Load Masterbatch", command=self.load_ram, width=160, height=32, radius=8, bg="#28A745", sf=self.SF)
            btn_ram.grid(row=1, column=1, sticky="w", padx=2, pady=5)
            
            tk.Label(self.inputs_container, text="Search (UNICO or CellName):", bg=COLOR_BG, font=FONT_LARGE, fg=COLOR_TEXT).grid(row=2, column=0, columnspan=3, pady=(20, 5))
            ent_search = tk.Entry(self.inputs_container, textvariable=self.query_var, font=("Arial", 16), justify="center", width=25)
            ent_search.grid(row=3, column=0, columnspan=3, pady=5)
            ent_search.bind("<KeyRelease>", self.do_search)

            tk.Label(self.inputs_container, text="RESULT (CLUSTER):", bg=COLOR_BG, font=FONT_LARGE, fg=COLOR_TEXT).grid(row=4, column=0, columnspan=3, pady=(15, 5))
            tk.Entry(self.inputs_container, textvariable=self.result_var, font=("Arial", 20, "bold"), justify="center", state="readonly", width=25).grid(row=5, column=0, columnspan=3)

        elif export_type == 'it_final_report':
            self.create_input_row(self.inputs_container, 'RND (*):', self.p_it_rnd, 'p_it_rnd', False, 0)
            self.create_input_row(self.inputs_container, 'CLUSTER CONTROL (*):', self.p_it_ctrl, 'p_it_ctrl', False, 1)
            self.create_input_row(self.inputs_container, 'CHANGES (*):', self.p_it_cambios, 'p_it_cambios', False, 2)

            lbl_elec = tk.Label(self.inputs_container, text="ELECTRICAL CHANGES DATE (*):", anchor="e", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
            lbl_elec.grid(row=3, column=0, sticky="e", padx=(0, 5), pady=5)
            ent_elec = tk.Entry(self.inputs_container, textvariable=self.date_elec, width=38, font=FONT_NORMAL)
            ent_elec.grid(row=3, column=1, sticky="w", padx=2, pady=5)
            btn_elec = RoundedButton(self.inputs_container, text="📅", command=lambda: self.select_date(self.date_elec), width=40, height=28, radius=8, sf=self.SF)
            btn_elec.grid(row=3, column=2, sticky="w", padx=2, pady=5)

            lbl_mech = tk.Label(self.inputs_container, text="MECHANICAL CHANGES DATE:", anchor="e", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
            lbl_mech.grid(row=4, column=0, sticky="e", padx=(0, 5), pady=5)
            ent_mech = tk.Entry(self.inputs_container, textvariable=self.date_mech, width=38, font=FONT_NORMAL)
            ent_mech.grid(row=4, column=1, sticky="w", padx=2, pady=5)
            btn_mech = RoundedButton(self.inputs_container, text="📅", command=lambda: self.select_date(self.date_mech), width=40, height=28, radius=8, sf=self.SF)
            btn_mech.grid(row=4, column=2, sticky="w", padx=2, pady=5)
            
            btn_dl_template = RoundedButton(self.inputs_container, text="📥 DOWNLOAD BLANK TEMPLATE", command=self.download_template, width=320, height=35, radius=8, sf=self.SF)
            btn_dl_template.grid(row=5, column=0, columnspan=3, pady=(20, 5))

        elif export_type == 'data_validation':
            self.create_input_row(self.inputs_container, 'BATCHFILE (*):', self.p_dv_batch, 'p_dv_batch', False, 0)
            self.create_input_row(self.inputs_container, '2G DATA (Opt):', self.p_dv_2g, 'p_dv_2g', False, 1)
            self.create_input_row(self.inputs_container, '3G DATA (Opt):', self.p_dv_3g, 'p_dv_3g', False, 2)
            self.create_input_row(self.inputs_container, '4G DATA (Opt):', self.p_dv_4g, 'p_dv_4g', False, 3)
            self.create_input_row(self.inputs_container, '5G DATA (Opt):', self.p_dv_5g, 'p_dv_5g', False, 4)

        self._slide_frame.update_idletasks()

    def load_ram(self):
        path = self.p_finder_batch.get()
        if not os.path.exists(path):
            CustomMessageBox(self.root, "Error", "Invalid Masterbatch path.", "error", self.SF)
            return
        self.lbl_status.config(text="Loading DB into RAM...")
        self.root.update_idletasks()
        success, msg = self.engine.load_masterbatch_to_memory(path)
        if success:
            CustomMessageBox(self.root, "Success", msg, "info", self.SF)
            self.lbl_status.config(text="DB Loaded. Finder is active.")
        else:
            CustomMessageBox(self.root, "Error", msg, "error", self.SF)

    def do_search(self, event=None):
        res = self.engine.search_cluster_fast(self.query_var.get())
        self.result_var.set(res)

    def check_ready_state(self):
        if self.is_running: return
        self.btn_run.enable()
        self.btn_cancel.update_text("EXIT")
        self.btn_cancel.command = self.root.destroy
        self.btn_cancel.enable()

    def update_progress(self, val, txt):
        if self.cancel_requested: raise ProcessCancelledException()
        self.root.after(0, lambda v=val, t=txt: self._sync_progress(v, t))
        
    def _sync_progress(self, val, txt):
        self.loader.set_target(val)
        self.lbl_status.config(text=txt)

    def start_process(self):
        if self.active_export == 'it_final_report':
            if not self.date_elec.get().strip():
                CustomMessageBox(self.root, "Error", "ELECTRIC DATE is mandatory.", "error", self.SF)
                self.reset_ui_state()
                return

        self.is_running = True
        self.cancel_requested = False
        self.btn_run.disable()
        self.btn_cancel.update_text("CANCEL PROCESS")
        self.loader.percent = 0.0
        
        if self.active_export == 'ept' and self.ept_subtype in ('assistant', 'acp'):
            master_path = self.p_master_ept.get()
            clusters = self.engine.get_clusters_from_master_ept(master_path)
            if not clusters:
                CustomMessageBox(self.root, "Error", "No clusters found in Master EPT.", "error", self.SF)
                self.reset_ui_state()
                return
            dialog = ClusterSelectionDialog(self.root, clusters, self.SF, title="Select Clusters")
            if not dialog.result:
                self.reset_ui_state()
                return
            self.selected_clusters = dialog.selected_clusters

        elif self.active_export == 'itsc' and self.itsc_subtype == 'rsh':
            try:
                df = pd.read_csv(self.p_itsc_cluster_control.get(), dtype=str, on_bad_lines='skip') if self.p_itsc_cluster_control.get().endswith('.csv') else pd.read_excel(self.p_itsc_cluster_control.get(), dtype=str)
                col = legacy.buscar_columna_inteligente(df.columns, ['CLUST', 'ID']) or 'Clust_Id'
                clusters = sorted([c for c in df[col].dropna().astype(str).str.strip().str.upper().unique() if c and c != 'NAN']) if col in df.columns else []
            except: clusters = []
            
            if not clusters:
                CustomMessageBox(self.root, "Error", "No clusters found in Cluster Control.", "error", self.SF)
                self.reset_ui_state()
                return
            dialog = ClusterSelectionDialog(self.root, clusters, self.SF, title="Select Clusters")
            if not dialog.result:
                self.reset_ui_state()
                return
            self.selected_clusters = dialog.selected_clusters

        elif self.active_export == 'object_trees' and self.object_subtype == 'cluster':
            try:
                df = pd.read_csv(self.p_obj_batchfile.get(), dtype=str, on_bad_lines='skip') if self.p_obj_batchfile.get().endswith('.csv') else pd.read_excel(self.p_obj_batchfile.get(), dtype=str)
                col = legacy.buscar_columna_inteligente(df.columns, ['CLUSTER']) or 'Cluster'
                clusters = sorted([c for c in df[col].dropna().astype(str).str.strip().str.upper().unique() if c and c != 'NAN']) if col in df.columns else []
            except: clusters = []
            
            if not clusters:
                CustomMessageBox(self.root, "Error", "No clusters found in Batchfile.", "error", self.SF)
                self.reset_ui_state()
                return
            dialog = ClusterSelectionDialog(self.root, clusters, self.SF, title="Select Clusters")
            if not dialog.result:
                self.reset_ui_state()
                return
            self.selected_obj_clusters = dialog.selected_clusters

        elif self.active_export == 'object_trees' and self.object_subtype == 'site':
            try:
                df = pd.read_csv(self.p_obj_batchfile.get(), dtype=str, on_bad_lines='skip') if self.p_obj_batchfile.get().endswith('.csv') else pd.read_excel(self.p_obj_batchfile.get(), dtype=str)
                col = legacy.buscar_columna_inteligente(df.columns, ['UNICO']) or 'UNICO'
                unicos = sorted([c for c in df[col].dropna().astype(str).str.strip().str.upper().unique() if c and c != 'NAN']) if col in df.columns else []
            except: unicos = []
            
            if not unicos:
                CustomMessageBox(self.root, "Error", "No UNICOs found in Batchfile.", "error", self.SF)
                self.reset_ui_state()
                return
            dialog = ClusterSelectionDialog(self.root, unicos, self.SF, title="Select UNICOs")
            if not dialog.result:
                self.reset_ui_state()
                return
            self.selected_obj_unicos = dialog.selected_clusters

        threading.Thread(target=self.run_thread, daemon=True).start()

    def reset_ui_state(self):
        self.is_running = False
        self.check_ready_state()

    def run_thread(self):
        try:
            out_p = self.p_out.get()
            fmt_csv = self.btn_fmt_csv.active
            fmt_xlsx = self.btn_fmt_xlsx.active
            
            if self.active_export == 'it_final_report':
                out_p = legacy.generate_it_final_report(
                    self.p_it_rnd.get(), 
                    None, 
                    self.p_it_ctrl.get(), 
                    self.p_it_cambios.get(), 
                    self.date_elec.get(), 
                    self.date_mech.get(), 
                    out_p, 
                    fmt_csv, 
                    fmt_xlsx, 
                    self.update_progress
                )
                self.update_progress(100.0, "[IT FINAL REPORT] SUCCESS.")
                
            elif self.active_export == 'ept':
                if self.ept_subtype == 'master':
                    legacy.generate_ept(self.p_mvs.get(), self.p_tp.get(), self.p_nics.get(), out_p, fmt_csv, fmt_xlsx, self.update_progress)
                elif self.ept_subtype == 'assistant':
                    self.engine.export_filtered_ept(self.p_master_ept.get(), out_p, self.selected_clusters, "ASSISTANT")
                elif self.ept_subtype == 'acp':
                    self.engine.export_filtered_ept(self.p_master_ept.get(), out_p, self.selected_clusters, "ACP")
                self.update_progress(100.0, f"[{self.ept_subtype.upper()} EPT] SUCCESS.")
                
            elif self.active_export == 'batch':
                if self.batch_subtype == 'master':
                    legacy.generate_batch(self.p_master_ept.get(), self.p_ran.get(), self.p_obj.get(), self.p_rnd.get(), self.p_dat.get(), self.p_ctrl.get(), self.p_prog.get(), self.p_clus.get(), out_p, fmt_csv, fmt_xlsx, self.update_progress)
                elif self.batch_subtype == 'audit':
                    legacy.generate_batch_audit(self.p_batch_old.get(), self.p_batch_new.get(), out_p, self.update_progress)
                self.update_progress(100.0, "[BATCH] SUCCESS.")

            elif self.active_export == 'cluster_friendly':
                legacy.generate_cluster_friendly(self.p_cluster_control.get(), self.p_clusterizacion.get(), out_p, fmt_csv, fmt_xlsx, self.update_progress)
                self.update_progress(100.0, "[CLUSTER FRIENDLY] SUCCESS.")

            elif self.active_export == 'itsc':
                if self.itsc_subtype == 'rsh':
                    clusters_str = ",".join(self.selected_clusters)
                    legacy.generate_itsc_template(clusters_str, self.p_itsc_cluster_control.get(), out_p, fmt_csv, fmt_xlsx, self.update_progress)
                elif self.itsc_subtype == 'swap':
                    swap_data = self.swap_text.get("1.0", "end-1c")
                    legacy.generate_itsc_swap(swap_data, out_p, fmt_csv, fmt_xlsx, self.update_progress)
                self.update_progress(100.0, "[ITSC] SUCCESS.")

            elif self.active_export == 'object_trees':
                if self.object_subtype == 'cluster':
                    export_mode = self.obj_cluster_mode.get()
                    legacy.generate_object_trees_cluster(self.p_obj_trees_folder.get(), self.p_obj_batchfile.get(), self.selected_obj_clusters, out_p, fmt_csv, fmt_xlsx, self.update_progress, export_mode)
                elif self.object_subtype == 'region':
                    legacy.generate_object_trees_regions(self.p_obj_trees_folder.get(), self.p_obj_batchfile.get(), out_p, fmt_csv, fmt_xlsx, self.update_progress)
                elif self.object_subtype == 'site':
                    export_mode = self.obj_site_mode.get()
                    legacy.generate_object_trees_site(self.p_obj_trees_folder.get(), self.p_obj_batchfile.get(), self.selected_obj_unicos, out_p, fmt_csv, fmt_xlsx, self.update_progress, export_mode)
                self.update_progress(100.0, "[OBJECT TREES] SUCCESS.")

            elif self.active_export == 'data_validation':
                paths = {
                    '2G': self.p_dv_2g.get(),
                    '3G': self.p_dv_3g.get(),
                    '4G': self.p_dv_4g.get(),
                    '5G': self.p_dv_5g.get()
                }
                legacy.generate_data_validation(paths, self.p_dv_batch.get(), out_p, self.update_progress)
                self.update_progress(100.0, "[DATA VALIDATION] SUCCESS.")
                
            time.sleep(1)
            self.root.after(0, lambda: CustomMessageBox(self.root, "Finished", "Generation completed successfully.", "info", self.SF))
            
        except Exception as e:
            err = f"{str(e)}\n\nTraceback:\n{traceback.format_exc()[-300:]}"
            self.root.after(0, lambda m=err: CustomMessageBox(self.root, "Error", m, "error", self.SF))
        finally:
            self.root.after(500, lambda: self.loader.set_target(0.0))
            self.root.after(500, self.reset_ui_state)
            gc.collect()

if __name__ == "__main__":
    check_for_updates()
    
    root = tk.Tk()
    dpi = root.winfo_fpixels('1i')
    SF = dpi / 96.0
    app = RFProcessorApp(root, scale_factor=SF)
    root.mainloop()
