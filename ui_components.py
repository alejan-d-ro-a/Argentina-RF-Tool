import os
import sys
import calendar
import tkinter as tk
from datetime import date

# =========================================================================
# FUNCIONES AUXILIARES
# =========================================================================

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

def set_toplevel_appwindow(toplevel):
    """Hace que un Toplevel (pop-up) aparezca en la barra de tareas."""
    if sys.platform == "win32":
        try:
            import ctypes
            hwnd = toplevel.winfo_id()
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
            if w > 1 and h > 1:
                rgn = ctypes.windll.gdi32.CreateRoundRectRgn(0, 0, w, h, radius, radius)
                ctypes.windll.user32.SetWindowRgn(hwnd, rgn, True)
        except: pass

def resource_path(relative_path):
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def hex_to_rgb(hx):
    hx = hx.lstrip('#')
    return tuple(int(hx[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_hex(rgb):
    return f"#{int(rgb[0]):02x}{int(rgb[1]):02x}{int(rgb[2]):02x}"

def to_hex(c):
    if not c: return "#808080"
    c = str(c).lower()
    if c == "white": return "#ffffff"
    if c in ["gray", "grey"]: return "#808080"
    if c == "black": return "#000000"
    if c.startswith("#"): return c
    return "#808080"

# =========================================================================
# CONFIGURACIÓN VISUAL
# =========================================================================
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

# =========================================================================
# DATE PICKER (ventana estándar)
# =========================================================================
class DatePickerDialog(tk.Toplevel):
    def __init__(self, parent, sf):
        super().__init__(parent)
        self.parent = parent
        self.sf = sf
        self.result = None
        
        self.title("Select Date")
        self.configure(bg=COLOR_BG)
        self.transient(parent)
        try:
            self.iconbitmap(resource_path("argentina.ico"))
        except:
            pass
        
        self.current_date = date.today()
        self.year = self.current_date.year
        self.month = self.current_date.month

        self.setup_ui()
        
        self.update_idletasks()
        req_w, req_h = int(280 * self.sf), int(330 * self.sf)
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        
        if pw < 100 or ph < 100 or px < -10000 or py < -10000:
            pw, ph = parent.winfo_screenwidth(), parent.winfo_screenheight()
            px, py = 0, 0

        x = px + (pw - req_w) // 2
        y = py + (ph - req_h) // 2
        
        self.geometry(f"{req_w}x{req_h}+{x}+{y}")
        self.resizable(False, False)
        self.update_idletasks()
        apply_rounded_corners(self, int(15 * self.sf))
        
        # Hacer que aparezca en la barra de tareas
        set_toplevel_appwindow(self)
        
        # Vincular eventos de la ventana padre para traer este diálogo al frente
        parent.bind("<FocusIn>", self._on_parent_focus, add="+")
        parent.bind("<ButtonPress-1>", self._on_parent_click, add="+")
        
        self.focus_force()
        self.grab_set()
        parent.wait_window(self)

    def _bring_to_front(self):
        if self.winfo_exists():
            self.lift()
            self.attributes('-topmost', True)
            self.update_idletasks()
            self.attributes('-topmost', False)
            self.focus_force()

    def _on_parent_focus(self, event):
        if self.winfo_exists():
            self.after(10, self._bring_to_front)

    def _on_parent_click(self, event):
        if self.winfo_exists():
            self.after(10, self._bring_to_front)

    def setup_ui(self):
        # Barra de título nativa, sin personalizar
        header = tk.Frame(self, bg=COLOR_BG)
        header.pack(fill="x", pady=10)
        tk.Button(header, text="<", command=self.prev_month, bg=COLOR_BTN, fg="white", font=FONT_NORMAL, width=3, relief="flat", cursor="hand2").pack(side="left", padx=10)
        self.lbl_month_year = tk.Label(header, text="", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE)
        self.lbl_month_year.pack(side="left", expand=True)
        tk.Button(header, text=">", command=self.next_month, bg=COLOR_BTN, fg="white", font=FONT_NORMAL, width=3, relief="flat", cursor="hand2").pack(side="right", padx=10)

        self.cal_frame = tk.Frame(self, bg="white")
        self.cal_frame.pack(expand=True, fill="both", padx=15, pady=(0, 15))
        self.update_calendar()

    def prev_month(self):
        self.month -= 1
        if self.month == 0:
            self.month = 12
            self.year -= 1
        self.update_calendar()

    def next_month(self):
        self.month += 1
        if self.month == 13:
            self.month = 1
            self.year += 1
        self.update_calendar()

    def select_day(self, day):
        self.result = date(self.year, self.month, day).strftime("%Y-%m-%d")
        self.destroy()

    def update_calendar(self):
        for widget in self.cal_frame.winfo_children():
            widget.destroy()

        days = ["Mo", "Tu", "We", "Th", "Fr", "Sa", "Su"]
        for i, d in enumerate(days):
            tk.Label(self.cal_frame, text=d, bg="white", fg=COLOR_BTN, font=FONT_NORMAL).grid(row=0, column=i)

        cal = calendar.monthcalendar(self.year, self.month)
        for row_idx, week in enumerate(cal):
            for col_idx, day in enumerate(week):
                if day != 0:
                    btn = tk.Button(self.cal_frame, text=str(day), bg="white", fg=COLOR_TEXT, relief="flat",
                                    command=lambda d=day: self.select_day(d), cursor="hand2")
                    btn.grid(row=row_idx+1, column=col_idx, sticky="nsew", padx=2, pady=2)
                else:
                    tk.Label(self.cal_frame, text="", bg="white").grid(row=row_idx+1, column=col_idx)

        for i in range(7):
            self.cal_frame.grid_columnconfigure(i, weight=1)

        month_name = calendar.month_name[self.month]
        self.lbl_month_year.config(text=f"{month_name} {self.year}")

# =========================================================================
# TOOLTIP
# =========================================================================
class ToolTip(object):
    def __init__(self, widget, text='widget info', sf=1.0):
        self.waittime = 400     
        self.wraplength = int(300 * sf)   
        self.sf = sf
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None; self.tw = None

    def enter(self, event=None): self.schedule()
    def leave(self, event=None): self.unschedule(); self.hidetip()
    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)
    def unschedule(self):
        id_ = self.id
        self.id = None
        if id_: self.widget.after_cancel(id_)
    def showtip(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        
        self.tw.bind("<Configure>", self._on_configure)
        
        label = tk.Label(self.tw, text=self.text, justify='left', background="#FFFFDD", foreground="#000000", relief='solid', borderwidth=1, font=FONT_NORMAL, wraplength=self.wraplength, padx=5, pady=5)
        label.pack(ipadx=1)
        
    def _on_configure(self, event):
        if str(event.widget) == str(self.tw):
            apply_rounded_corners(self.tw, int(8 * self.sf))

    def hidetip(self):
        tw = self.tw; self.tw = None
        if tw: tw.destroy()
    def update_text(self, new_text): self.text = new_text

# =========================================================================
# TEXT ANIMATION WRAPPER
# =========================================================================
class TextAnimWrapper:
    def __init__(self, obj, type_str):
        self.obj = obj
        self.type = type_str
    def get(self):
        if self.type == "label": return self.obj.cget("text")
        elif self.type == "btn": return self.obj.text
        elif self.type == "tt": return self.obj.text
        elif self.type == "entry": return self.obj.get()
    def set(self, val):
        try:
            if self.type == "label": self.obj.config(text=val)
            elif self.type == "btn": self.obj.update_text(val)
            elif self.type == "tt": self.obj.update_text(val)
            elif self.type == "entry":
                state = self.obj.cget("state")
                self.obj.config(state="normal")
                self.obj.delete(0, tk.END)
                self.obj.insert(0, val)
                self.obj.config(state=state)
        except tk.TclError:
            pass

# =========================================================================
# BOTÓN REDONDEADO
# =========================================================================
class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, width=400, height=45, radius=20, bg=COLOR_BTN,
                 hover_bg=COLOR_BTN_HOVER, is_toggle=False, default_active=False, sf=1.0,
                 trace_cb=None, custom_font=None, auto_toggle=True, canvas_bg=COLOR_BG):
        self.sf = sf
        self.width = int(width * sf)
        self.height = int(height * sf)
        self.radius = int(radius * sf)
        self.pad_normal = float(3 * sf)
        self.pad_zoom = 0.0
        self.font_family = custom_font[0] if custom_font else "Comic Sans MS"
        self.font_size = float(custom_font[1] if custom_font else 11)
        super().__init__(parent, width=self.width, height=self.height, bg=canvas_bg, highlightthickness=0)
        self.command = command
        self.trace_cb = trace_cb
        self.base_bg = bg
        self.hover_bg = hover_bg
        self.text = text
        self.is_toggle = is_toggle
        self.auto_toggle = auto_toggle
        self.active = default_active
        self.estado = "normal"
        self._color_anim = None
        self.config(cursor="hand2")
        self.update_colors()
        self.current_drawn_bg = self.current_bg
        self.current_drawn_fg = self.current_fg
        self.current_drawn_pad = self.pad_normal
        self.draw_button_shape()
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Button-1>", self.on_click)

    def update_colors(self):
        if self.is_toggle:
            self.current_bg = self.base_bg if self.active else COLOR_DISABLED_BG
            self.current_fg = "#FFFFFF" if self.active else COLOR_TEXT
        else:
            self.current_bg = self.base_bg
            self.current_fg = "#FFFFFF"

    def create_rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        points = [x1+r, y1, x1+r, y1, x2-r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y1+r, x2, y2-r, x2, y2-r, x2, y2, x2-r, y2, x2-r, y2, x1+r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y2-r, x1, y1+r, x1, y1+r, x1, y1]
        return self.create_polygon(points, **kwargs, smooth=True)

    def draw_button_shape(self):
        try:
            self.delete("all")
            pad = self.current_drawn_pad
            font = (self.font_family, int(round(self.font_size)), "bold")
            self.create_rounded_rect(pad, pad, self.width-pad, self.height-pad, self.radius,
                                     fill=self.current_drawn_bg, outline="")
            max_text_width = self.width - (self.current_drawn_pad * 2) - 10
            self.create_text(self.width/2, self.height/2, text=self.text,
                             fill=self.current_drawn_fg, font=font, justify=tk.CENTER,
                             width=max_text_width)
        except tk.TclError:
            pass

    def animate_transition(self, target_bg, target_fg, target_pad, steps=20, delay=10):
        if self._color_anim is not None:
            self.after_cancel(self._color_anim)
            self._color_anim = None
        curr_bg_r, curr_bg_g, curr_bg_b = hex_to_rgb(to_hex(self.current_drawn_bg))
        targ_bg_r, targ_bg_g, targ_bg_b = hex_to_rgb(to_hex(target_bg))
        curr_fg_r, curr_fg_g, curr_fg_b = hex_to_rgb(to_hex(self.current_drawn_fg))
        targ_fg_r, targ_fg_g, targ_fg_b = hex_to_rgb(to_hex(target_fg))
        dr_bg = (targ_bg_r - curr_bg_r) / steps
        dg_bg = (targ_bg_g - curr_bg_g) / steps
        db_bg = (targ_bg_b - curr_bg_b) / steps
        dr_fg = (targ_fg_r - curr_fg_r) / steps
        dg_fg = (targ_fg_g - curr_fg_g) / steps
        db_fg = (targ_fg_b - curr_fg_b) / steps
        curr_pad = self.current_drawn_pad
        dp_pad = (target_pad - curr_pad) / steps
        def step(step_num):
            try:
                if step_num > steps:
                    self.current_drawn_bg = target_bg
                    self.current_drawn_fg = target_fg
                    self.current_drawn_pad = target_pad
                    self.draw_button_shape()
                    return
                r_bg = int(curr_bg_r + dr_bg * step_num)
                g_bg = int(curr_bg_g + dg_bg * step_num)
                b_bg = int(curr_bg_b + db_bg * step_num)
                self.current_drawn_bg = rgb_to_hex((r_bg, g_bg, b_bg))
                r_fg = int(curr_fg_r + dr_fg * step_num)
                g_fg = int(curr_fg_g + dg_fg * step_num)
                b_fg = int(curr_fg_b + db_fg * step_num)
                self.current_drawn_fg = rgb_to_hex((r_fg, g_fg, b_fg))
                self.current_drawn_pad = curr_pad + dp_pad * step_num
                self.draw_button_shape()
                self._color_anim = self.after(delay, step, step_num + 1)
            except tk.TclError:
                pass
        step(1)

    def update_text(self, new_text):
        self.text = new_text
        self.draw_button_shape()

    def on_enter(self, e):
        if self.estado == "normal":
            self.estado = "hover"
            target_bg = self.hover_bg if (not self.is_toggle or self.active) else "#B0B0B0"
            target_fg = "#FFFFFF" if (not self.is_toggle or self.active) else COLOR_TEXT
            self.animate_transition(target_bg, target_fg, self.pad_zoom)

    def on_leave(self, e):
        if self.estado == "hover":
            self.estado = "normal"
            self.animate_transition(self.current_bg, self.current_fg, self.pad_normal)

    def on_click(self, e):
        if self.estado in ["normal", "hover"]:
            if self.is_toggle and self.auto_toggle:
                self.active = not self.active
                self.update_colors()
                target_bg = self.hover_bg if self.active else "#B0B0B0"
                target_fg = "#FFFFFF" if self.active else COLOR_TEXT
                self.animate_transition(target_bg, target_fg, self.pad_zoom)
            if self.trace_cb: self.trace_cb()
            if self.command: self.command()

    def disable(self):
        self.estado = "disabled"
        self.config(cursor="no")
        if self._color_anim:
            self.after_cancel(self._color_anim)
            self._color_anim = None
        self.current_drawn_bg = COLOR_DISABLED_BG
        self.current_drawn_fg = COLOR_DISABLED_FG
        self.current_drawn_pad = self.pad_normal
        self.draw_button_shape()

    def enable(self):
        self.estado = "normal"
        self.config(cursor="hand2")
        self.update_colors()
        self.animate_transition(self.current_bg, self.current_fg, self.pad_normal)

# =========================================================================
# BARRA DE PROGRESO
# =========================================================================
class RoundedProgressBar(tk.Canvas):
    def __init__(self, parent, width=800, height=35, bg_color="#E0E0E0", fill_color=COLOR_BTN, sf=1.0):
        self.width = int(width * sf)
        self.height = int(height * sf)
        self.radius = int(15 * sf)
        super().__init__(parent, width=self.width, height=self.height, bg=COLOR_BG, highlightthickness=0)
        self.bg_color = bg_color
        self.fill_color = fill_color
        self.percent = 0.0
        self.target_percent = 0.0
        self._animating = False
        self.draw_bar()

    def create_rounded_rect(self, x1, y1, x2, y2, r, **kwargs):
        points = [x1+r, y1, x1+r, y1, x2-r, y1, x2-r, y1, x2, y1, x2, y1+r, x2, y1+r, x2, y2-r, x2, y2-r, x2, y2, x2-r, y2, x2-r, y2, x1+r, y2, x1+r, y2, x1, y2, x1, y2-r, x1, y2-r, x1, y1+r, x1, y1+r, x1, y1]
        return self.create_polygon(points, **kwargs, smooth=True)

    def draw_bar(self):
        self.delete("all")
        self.create_rounded_rect(2, 2, self.width-2, self.height-2, self.radius, fill=self.bg_color, outline="")
        fill_width = max(0, (self.percent / 100.0) * (self.width-4))
        if fill_width > 0:
            if fill_width < 2 * self.radius: fill_width = 2 * self.radius
            self.create_rounded_rect(2, 2, 2+fill_width, self.height-2, self.radius, fill=self.fill_color, outline="")
        text_color = "white" if self.percent > 50 else COLOR_TEXT
        self.create_text(self.width/2, self.height/2, text=f"{self.percent:.2f}%", fill=text_color, font=FONT_LARGE)

    def set_target(self, target):
        self.target_percent = max(0.0, min(100.0, target))
        if not self._animating:
            self._animating = True
            self.animate()

    def animate(self):
        diff = self.target_percent - self.percent
        if abs(diff) < 0.5:
            self.percent = self.target_percent
            self.draw_bar()
            self._animating = False
            return
        self.percent += diff * 0.15
        self.draw_bar()
        self.after(20, self.animate)

# =========================================================================
# CUSTOM MESSAGE BOX (ventana estándar)
# =========================================================================
class CustomMessageBox:
    def __init__(self, parent, title, message, msg_type="info", sf=1.0):
        self.parent = parent
        self.title = title
        self.message = message
        self.msg_type = msg_type
        self.sf = sf
        self.result = False
        
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.configure(bg=COLOR_BG)
        self.top.transient(parent)
        try:
            self.top.iconbitmap(resource_path("argentina.ico"))
        except:
            pass
        
        f_msg = tk.Frame(self.top, bg=COLOR_BG)
        f_msg.pack(expand=True, fill="both", padx=int(20 * self.sf), pady=int(15 * self.sf))
        lbl_msg = tk.Label(f_msg, text=self.message, bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE, wraplength=int(460 * self.sf), justify="center")
        lbl_msg.pack(expand=True)
        
        f_btn = tk.Frame(self.top, bg=COLOR_BG)
        f_btn.pack(side="bottom", pady=int(20 * self.sf))
        if self.msg_type == "yesno":
            btn_yes = RoundedButton(f_btn, "Yes", self.close_yes, width=160, height=45, sf=self.sf)
            btn_yes.pack(side="left", padx=int(10 * self.sf))
            btn_no = RoundedButton(f_btn, "No", self.close_no, width=160, height=45, bg=COLOR_BTN_CANCEL, hover_bg=COLOR_BTN_CANCEL_HOVER, sf=self.sf)
            btn_no.pack(side="left", padx=int(10 * self.sf))
        else:
            btn_bg = COLOR_BTN_CANCEL if self.msg_type == "error" else COLOR_BTN
            btn_hover = COLOR_BTN_CANCEL_HOVER if self.msg_type == "error" else COLOR_BTN_HOVER
            btn_ok = RoundedButton(f_btn, "OK", self.close_yes, width=200, height=45, bg=btn_bg, hover_bg=btn_hover, sf=self.sf)
            btn_ok.pack()
            
        self.top.update_idletasks()
        req_w = max(int(520 * self.sf), self.top.winfo_reqwidth())
        req_h = max(int(250 * self.sf), self.top.winfo_reqheight() + int(20 * self.sf))
        
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        
        if pw < 100 or ph < 100 or px < -10000 or py < -10000:
            if parent.state() == 'iconic':
                parent.deiconify()
            pw = parent.winfo_screenwidth()
            ph = parent.winfo_screenheight()
            px = 0
            py = 0
            
        x = px + (pw - req_w) // 2
        y = py + (ph - req_h) // 2
        
        self.top.geometry(f"{req_w}x{req_h}+{x}+{y}")
        self.top.resizable(False, False)
        self.top.update_idletasks()
        apply_rounded_corners(self.top, int(20 * self.sf))
        
        set_toplevel_appwindow(self.top)
        
        # Vincular eventos de la ventana padre
        parent.bind("<FocusIn>", self._on_parent_focus, add="+")
        parent.bind("<ButtonPress-1>", self._on_parent_click, add="+")
        
        self.top.focus_force()
        self.top.grab_set()
        parent.wait_window(self.top)

    def _bring_to_front(self):
        if self.top.winfo_exists():
            self.top.lift()
            self.top.attributes('-topmost', True)
            self.top.update_idletasks()
            self.top.attributes('-topmost', False)
            self.top.focus_force()

    def _on_parent_focus(self, event):
        if self.top.winfo_exists():
            self.top.after(10, self._bring_to_front)

    def _on_parent_click(self, event):
        if self.top.winfo_exists():
            self.top.after(10, self._bring_to_front)

    def close_yes(self):
        self.result = True
        self.top.destroy()
    def close_no(self):
        self.result = False
        self.top.destroy()

# =========================================================================
# CLUSTER SELECTION DIALOG (ventana estándar)
# =========================================================================
class ClusterSelectionDialog:
    def __init__(self, parent, items, sf=1.0, title="Select Items"):
        self.parent = parent
        self.items = items
        self.sf = sf
        self.title = title
        self.result = False
        self.selected_clusters = set()

        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.configure(bg=COLOR_BG)
        self.top.transient(parent)
        try:
            self.top.iconbitmap(resource_path("argentina.ico"))
        except:
            pass

        f_search = tk.Frame(self.top, bg=COLOR_BG)
        f_search.pack(fill="x", padx=int(10 * self.sf), pady=int(10 * self.sf))
        tk.Label(f_search, text="Search:", bg=COLOR_BG, fg=COLOR_TEXT, font=FONT_LARGE).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.update_list)
        self.entry_search = tk.Entry(f_search, textvariable=self.search_var, font=FONT_NORMAL, width=30)
        self.entry_search.pack(side="left", padx=int(10 * self.sf))

        f_list = tk.Frame(self.top, bg=COLOR_BG)
        f_list.pack(fill="both", expand=True, padx=int(10 * self.sf))
        
        self.scrollbar = tk.Scrollbar(f_list)
        self.scrollbar.pack(side="right", fill="y")
        
        self.listbox = tk.Listbox(f_list, selectmode=tk.MULTIPLE, yscrollcommand=self.scrollbar.set, font=FONT_NORMAL, bg="white", fg=COLOR_TEXT, selectbackground=COLOR_BTN_HOVER)
        self.listbox.pack(side="left", fill="both", expand=True)
        self.scrollbar.config(command=self.listbox.yview)

        f_btn = tk.Frame(self.top, bg=COLOR_BG)
        f_btn.pack(side="bottom", pady=int(15 * self.sf))
        btn_ok = RoundedButton(f_btn, "Confirm", self.close_yes, width=150, height=40, sf=self.sf)
        btn_ok.pack(side="left", padx=int(10 * self.sf))
        btn_no = RoundedButton(f_btn, "Cancel", self.close_no, width=150, height=40, bg=COLOR_BTN_CANCEL, hover_bg=COLOR_BTN_CANCEL_HOVER, sf=self.sf)
        btn_no.pack(side="right", padx=int(10 * self.sf))

        self.update_list()
        
        self.top.update_idletasks()
        req_w, req_h = int(450 * self.sf), int(500 * self.sf)
        pw, ph = parent.winfo_width(), parent.winfo_height()
        px, py = parent.winfo_rootx(), parent.winfo_rooty()
        
        x = px + (pw - req_w) // 2
        y = py + (ph - req_h) // 2
        
        self.top.geometry(f"{req_w}x{req_h}+{x}+{y}")
        self.top.resizable(False, False)
        self.top.update_idletasks()
        apply_rounded_corners(self.top, int(15 * self.sf))
        
        set_toplevel_appwindow(self.top)
        
        # Vincular eventos de la ventana padre
        parent.bind("<FocusIn>", self._on_parent_focus, add="+")
        parent.bind("<ButtonPress-1>", self._on_parent_click, add="+")
        
        self.top.focus_force()
        self.top.grab_set()
        parent.wait_window(self.top)

    def _bring_to_front(self):
        if self.top.winfo_exists():
            self.top.lift()
            self.top.attributes('-topmost', True)
            self.top.update_idletasks()
            self.top.attributes('-topmost', False)
            self.top.focus_force()

    def _on_parent_focus(self, event):
        if self.top.winfo_exists():
            self.top.after(10, self._bring_to_front)

    def _on_parent_click(self, event):
        if self.top.winfo_exists():
            self.top.after(10, self._bring_to_front)

    def sync_selection(self):
        if not hasattr(self, 'listbox') or not self.listbox.winfo_exists():
            return
            
        visible = [self.listbox.get(i) for i in range(self.listbox.size())]
        selected_idx = self.listbox.curselection()
        selected_vals = {self.listbox.get(i) for i in selected_idx}
        
        self.selected_clusters.update(selected_vals)
        unselected = set(visible) - selected_vals
        self.selected_clusters.difference_update(unselected)

    def update_list(self, *args):
        self.sync_selection()
        
        term = self.search_var.get().lower()
        self.listbox.delete(0, tk.END)
        for c in self.items:
            if term in str(c).lower(): 
                self.listbox.insert(tk.END, c)
                
        for i in range(self.listbox.size()):
            if self.listbox.get(i) in self.selected_clusters:
                self.listbox.select_set(i)

    def close_yes(self):
        self.sync_selection()
        if not self.selected_clusters:
            CustomMessageBox(self.top, "Warning", "Please select at least one item.", "error", self.sf)
            return
        self.result = True
        self.top.destroy()
        
    def close_no(self):
        self.result = False
        self.top.destroy()
