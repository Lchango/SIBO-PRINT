import os
import sys
import re
import textwrap
import win32print
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from PIL import Image, ImageTk, ImageDraw, ImageFilter
from io import BytesIO
import ctypes
import math

# Habilitar soporte DPI para evitar textos borrosos (pixelados) en monitores modernos
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# --- OPTIMIZACIONES DE ARRANQUE: REEMPLAZO DE PANDAS (LIGHTWEIGHT DATAFRAME) ---
def notna(val):
    if val is None:
        return False
    if isinstance(val, float):
        if math.isnan(val):
            return False
    return True

class ExcelRow:
    def __init__(self, columns, row_values):
        self.columns = columns
        self.row_values = row_values
        
    @property
    def values(self):
        return self.row_values
        
    def get(self, col_name, default=None):
        try:
            idx = self.columns.index(col_name)
            return self.row_values[idx]
        except ValueError:
            return default

class ILocIndexer:
    def __init__(self, df):
        self.df = df
        
    def __getitem__(self, index):
        return ExcelRow(self.df.columns, self.df.rows[index])

class AtIndexer:
    def __init__(self, df):
        self.df = df
        
    def __getitem__(self, key):
        row_idx, col_name = key
        col_idx = self.df.columns.index(col_name)
        return self.df.rows[row_idx][col_idx]
        
    def __setitem__(self, key, value):
        row_idx, col_name = key
        col_idx = self.df.columns.index(col_name)
        self.df.rows[row_idx][col_idx] = value

class ExcelDataFrame:
    def __init__(self, columns, rows):
        self.columns = list(columns)
        self.rows = [list(row) for row in rows]
        self.iloc = ILocIndexer(self)
        self.at = AtIndexer(self)
        
    @property
    def empty(self):
        return len(self.rows) == 0
        
    def __len__(self):
        return len(self.rows)

def read_excel_light(file_path):
    import openpyxl
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet = wb.active
    
    columns = []
    rows = []
    
    first = True
    for row in sheet.iter_rows(values_only=True):
        if first:
            columns = [str(cell).strip() if cell is not None else f"Column_{i}" for i, cell in enumerate(row)]
            first = False
        else:
            if any(cell is not None for cell in row):
                row_val = list(row)
                if len(row_val) < len(columns):
                    row_val += [None] * (len(columns) - len(row_val))
                elif len(row_val) > len(columns):
                    row_val = row_val[:len(columns)]
                rows.append(row_val)
                
    wb.close()
    return ExcelDataFrame(columns, rows)


class ZebraPrintApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SIBO PRINT")
        
        # Tamaño responsivo dinámico (Ancho compacto, Alto casi completo)
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        
        # Ancho 60% (compacto), Alto 90% (casi completo)
        win_w = int(screen_w * 0.68)
        win_h = int(screen_h * 0.90)
        
        # Limites para evitar desbordamiento en pantallas de baja resolución
        win_w = max(win_w, 1150)
        win_h = max(win_h, 750)
        
        # Calcular centrado horizontal y dar un margen superior (padding top) fijo
        x = int(max(0, (screen_w - win_w) / 2))
        y = 40 # 40 píxeles de espacio desde el borde superior
        
        # Si el alto choca con tu barra de tareas de Windows, ajustar el height
        if win_h + y > screen_h - 50:
            win_h = screen_h - y - 50
            
        self.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.minsize(900, 600) # Tamaño mínimo de seguridad
        
        # Variables lógicas
        self.CHECK_COL = "__sel__"  # Id interno de la columna de checkboxes de la tabla
        self.df = None
        self.plantilla_zpl = None
        self.variables_prn = []
        self.map_columnas = {} # variable -> columna de excel auto-asociada
        
        # Determinar el directorio base (externo si es ejecutable compilado)
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.prn_folder = os.path.join(base_dir, "prn_templates")
        self._preview_timer = None # Control para slider en vivo
        self._current_preview_img = None # Guarda la imagen alta calidad descargada
        self._resize_timer = None # Control para diseño responsive
        self.preview_api_enabled = tk.BooleanVar(value=True) # Control ON/OFF de previsualizacion web (activa por defecto)
        self._switch_anim_timer = None # Timer para animar el interruptor
        self._switch_knob_x = 4 # Posicion inicial de perilla (OFF)
        
        # Construir Interfaz y Estilos
        self._setup_styles()
        self.configure(bg=self.bg_color)
        
        # Configurar icono de la aplicación
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons", "icon1.png")
            if os.path.exists(icon_path):
                icon_img = ImageTk.PhotoImage(Image.open(icon_path))
                self.iconphoto(False, icon_img)
            else:
                self.iconphoto(False, tk.PhotoImage(width=1, height=1))
        except Exception as e:
            print("No se pudo cargar el icono configurado:", e)

        self._build_ui()
        self._load_local_data()

    def _setup_styles(self):
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except:
            pass
        
        # Paleta de colores - Graphite / Petrol Dark Theme (tonos suaves, no azul puro)
        self.bg_color = "#131A21"       # Fondo general (grafito con matiz petróleo)
        self.card_color = "#1C252E"     # Fondo de tarjetas (grafito medio)
        self.border_color = "#2E3A47"   # Bordes sutiles
        self.fg_color = "#F5F7FA"       # Texto principal (blanco suave)
        self.fg_muted = "#9AA7B4"       # Texto secundario (gris claro)

        self.row_even = "#1C252E"       # Fila alterna 1 (gris medio)
        self.row_odd = "#161E25"        # Fila alterna 2 (gris oscuro)
        self.row_hover = "#233241"      # Resalte al pasar el mouse

        self.preview_bg = "#E4E8EC"     # Fondo claro para destacar la etiqueta en la vista previa

        # Acento principal (azul petróleo, no azul puro)
        self.accent_blue = "#1F8A8A"    # Teal petróleo (usado para bordes/acentos/estado)
        self.accent_blue_hover = "#2CA6A6"

        # Degradados de botones (inicio -> fin)
        self.grad_excel = ("#84CC16", "#14B8A6")      # Verde lima -> Turquesa
        self.grad_calibrate = ("#FB923C", "#F59E0B")  # Naranja -> Ámbar suave
        self.grad_print = ("#0D9488", "#2563EB")      # Petróleo -> Azul
        self.grad_reset = ("#55606C", "#3C4650")      # Grafito neutro

        # Compatibilidad: alias usados como pares (inicio, hover) en otras partes del código
        self.accent_green, self.accent_green_hover = self.grad_excel
        self.accent_amber, self.accent_amber_hover = self.grad_calibrate
        self.accent_gray, self.accent_gray_hover = self.grad_reset

        # Configuración básica de elementos ttk
        self.style.configure("TFrame", background=self.bg_color)
        self.style.configure("Panel.TFrame", background=self.card_color)
        
        self.style.configure("TLabel", background=self.bg_color, foreground=self.fg_color, font=("Segoe UI", 9))
        self.style.configure("Panel.TLabel", background=self.card_color, foreground=self.fg_color, font=("Segoe UI", 9))
        self.style.configure("Header.TLabel", background=self.card_color, foreground=self.fg_color, font=("Segoe UI", 10, "bold"))
        self.style.configure("Muted.TLabel", background=self.card_color, foreground=self.fg_muted, font=("Segoe UI", 9))
        
        # Comboboxes estilizados oscuros
        self.style.map('TCombobox',
            fieldbackground=[('readonly', self.card_color)],
            background=[('readonly', self.card_color)],
            foreground=[('readonly', self.fg_color)]
        )
        self.style.configure('TCombobox',
            fieldbackground=self.card_color,
            background=self.border_color,
            foreground=self.fg_color,
            bordercolor=self.border_color,
            lightcolor=self.border_color,
            darkcolor=self.border_color,
            arrowcolor=self.fg_color,
            padding=4,
            insertcolor=self.fg_color,
            font=("Segoe UI", 9)
        )
        
        # Treeview de datos en Slate
        self.style.configure("Treeview",
            background=self.card_color,
            foreground=self.fg_color,
            fieldbackground=self.card_color,
            borderwidth=0,
            rowheight=28,
            font=("Segoe UI", 9)
        )
        self.style.configure("Treeview.Heading",
            background=self.bg_color,
            foreground=self.fg_muted,
            font=("Segoe UI", 9, "bold"),
            borderwidth=0,
            padding=8
        )
        self.style.map("Treeview.Heading",
            background=[("active", self.bg_color)],
            foreground=[("active", self.fg_color)]
        )
        self.style.map("Treeview",
            background=[("selected", self.accent_blue)],
            foreground=[("selected", "#FFFFFF")]
        )
        
        # Sliders
        self.style.configure("TScale",
            background=self.card_color,
            troughcolor=self.bg_color,
            sliderlength=18,
            sliderthickness=12
        )
        
        # Scrollbars modernas estilizadas (delgadas y sin flechas)
        self.style.layout('Dark.Vertical.TScrollbar', 
            [('Vertical.Scrollbar.trough', {'children': 
                [('Vertical.Scrollbar.thumb', {'expand': '1', 'sticky': 'nswe'})], 
            'sticky': 'ns'})])
        self.style.configure('Dark.Vertical.TScrollbar',
            troughcolor=self.bg_color,
            background=self.border_color,
            bordercolor=self.bg_color,
            lightcolor=self.bg_color,
            darkcolor=self.bg_color,
            arrowsize=0,
            width=10
        )
        
        self.style.layout('Dark.Horizontal.TScrollbar', 
            [('Horizontal.Scrollbar.trough', {'children': 
                [('Horizontal.Scrollbar.thumb', {'expand': '1', 'sticky': 'nswe'})], 
            'sticky': 'we'})])
        self.style.configure('Dark.Horizontal.TScrollbar',
            troughcolor=self.bg_color,
            background=self.border_color,
            bordercolor=self.bg_color,
            lightcolor=self.bg_color,
            darkcolor=self.bg_color,
            arrowsize=0,
            width=10
        )

    @staticmethod
    def _hex_to_rgb(hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))

    def _render_gradient_rounded(self, width, height, color_from, color_to, radius=8, shadow=True):
        """Genera (con PIL) un rectángulo redondeado con relleno en degradado horizontal
        y una sombra suave debajo, para simular botones modernos con elevación."""
        scale = 3  # supersampling para bordes redondeados suaves
        pad = 6 if shadow else 0
        canvas_w, canvas_h = width + pad * 2, height + pad * 2

        big = Image.new("RGBA", (canvas_w * scale, canvas_h * scale), (0, 0, 0, 0))

        if shadow:
            shadow_layer = Image.new("RGBA", (canvas_w * scale, canvas_h * scale), (0, 0, 0, 0))
            sd = ImageDraw.Draw(shadow_layer)
            off = 3 * scale
            sd.rounded_rectangle(
                [pad * scale, pad * scale + off, (pad + width) * scale, (pad + height) * scale + off],
                radius=radius * scale, fill=(0, 0, 0, 100)
            )
            shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(radius=max(2, radius // 2) * scale // 2))
            big.alpha_composite(shadow_layer)

        r1, g1, b1 = self._hex_to_rgb(color_from)
        r2, g2, b2 = self._hex_to_rgb(color_to)
        grad = Image.new("RGB", (width, 1))
        px = grad.load()
        for x in range(width):
            t = x / max(1, width - 1)
            px[x, 0] = (round(r1 + (r2 - r1) * t), round(g1 + (g2 - g1) * t), round(b1 + (b2 - b1) * t))
        grad = grad.resize((width * scale, height * scale)).convert("RGBA")

        mask = Image.new("L", (width * scale, height * scale), 0)
        ImageDraw.Draw(mask).rounded_rectangle(
            [0, 0, width * scale - 1, height * scale - 1], radius=radius * scale, fill=255
        )
        grad.putalpha(mask)
        big.alpha_composite(grad, (pad * scale, pad * scale))

        final = big.resize((canvas_w, canvas_h), Image.LANCZOS)
        return final, pad

    def _create_hover_button(self, parent, text, image, color_from, color_to, command, is_bold=True, compound="left", fg_color="#FFFFFF"):
        """Botón moderno: fondo en degradado, esquinas redondeadas y sombra sutil (renderizado con PIL sobre un Canvas)."""
        label_text = text.strip()
        font_spec = ("Segoe UI", 9, "bold" if is_bold else "normal")
        measure_font = tkfont.Font(family="Segoe UI", size=9, weight="bold" if is_bold else "normal")
        text_w = measure_font.measure(label_text)
        text_h = measure_font.metrics("linespace")

        icon_w = (image.width() + 8) if image else 0
        content_w = icon_w + text_w
        btn_w = max(84, content_w + 22)
        btn_h = max(26, text_h + 8)

        normal_img, pad = self._render_gradient_rounded(btn_w, btn_h, color_from, color_to, radius=8, shadow=True)
        hover_img, _ = self._render_gradient_rounded(btn_w, btn_h, color_to, color_from, radius=8, shadow=True)
        normal_photo = ImageTk.PhotoImage(normal_img)
        hover_photo = ImageTk.PhotoImage(hover_img)

        canvas_w, canvas_h = normal_img.width, normal_img.height
        btn = tk.Canvas(parent, width=canvas_w, height=canvas_h, bg=self.bg_color, highlightthickness=0, bd=0, cursor="hand2")
        btn._normal_img = normal_photo
        btn._hover_img = hover_photo
        btn._icon_img = image

        btn.create_image(0, 0, image=normal_photo, anchor="nw", tags="bgimg")

        cx, cy = canvas_w / 2, canvas_h / 2
        if image:
            total_w = image.width() + 8 + text_w
            start_x = cx - total_w / 2
            btn.create_image(start_x + image.width() / 2, cy, image=image, anchor="center", tags="hit")
            btn.create_text(start_x + image.width() + 8 + text_w / 2, cy, text=label_text, fill=fg_color, font=font_spec, tags="hit")
        else:
            btn.create_text(cx, cy, text=label_text, fill=fg_color, font=font_spec, tags="hit")

        def on_enter(e):
            btn.itemconfig("bgimg", image=hover_photo)

        def on_leave(e):
            btn.itemconfig("bgimg", image=normal_photo)

        def on_click(e):
            command()

        for seq, fn in (("<Enter>", on_enter), ("<Leave>", on_leave), ("<Button-1>", on_click)):
            btn.bind(seq, fn)
            btn.tag_bind("hit", seq, fn)

        return btn

    def _create_card(self, parent, title):
        # Card outer container con borde sutil
        card = tk.Frame(
            parent,
            bg=self.card_color,
            highlightbackground=self.border_color,
            highlightcolor=self.border_color,
            highlightthickness=1,
            bd=0
        )

        # Franja superior de acento (look moderno tipo dashboard)
        accent_strip = tk.Frame(card, bg=self.accent_blue, height=3)
        accent_strip.pack(fill="x", side="top")

        # Title bar
        header_frame = tk.Frame(card, bg=self.card_color)
        header_frame.pack(fill="x", padx=14, pady=(9, 7))

        lbl_title = tk.Label(
            header_frame,
            text=title,
            bg=self.card_color,
            fg=self.fg_color,
            font=("Segoe UI", 10, "bold")
        )
        lbl_title.pack(side="left")

        # Content body frame
        content_frame = tk.Frame(card, bg=self.card_color)
        content_frame.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        return card, content_frame

    def _fetch_icon(self, url, filename):
        icon_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
        if not os.path.exists(icon_folder):
            try: os.makedirs(icon_folder)
            except: pass
            
        path = os.path.join(icon_folder, filename)
        if not os.path.exists(path):
            try:
                import requests
                res = requests.get(url, timeout=5)
                if res.status_code == 200:
                    with open(path, 'wb') as f:
                        f.write(res.content)
            except Exception as e:
                print(f"Error descargando {filename}:", e)
        
        if os.path.exists(path):
            try:
                return ImageTk.PhotoImage(Image.open(path))
            except:
                pass
        return tk.PhotoImage(width=1, height=1)

    def _build_ui(self):
        # ---------------------------------------------
        # TOP HEADER: Barra de Herramientas Principal
        # ---------------------------------------------
        self.frame_header = tk.Frame(self, bg=self.bg_color, height=52)
        self.frame_header.pack(fill="x", padx=25, pady=(12, 8))
        self.frame_header.pack_propagate(False)

        # Título / Logo
        title_frame = tk.Frame(self.frame_header, bg=self.bg_color)
        title_frame.pack(side="left", anchor="w")

        lbl_logo = tk.Label(title_frame, text="SIBO", bg=self.bg_color, fg=self.accent_blue, font=("Segoe UI", 15, "bold"))
        lbl_logo.pack(side="left")

        lbl_title = tk.Label(title_frame, text=" PRINT", bg=self.bg_color, fg=self.fg_color, font=("Segoe UI", 15, "normal"))
        lbl_title.pack(side="left")
        
        # Cargar iconos de la web o de caché local
        self.icon_excel = self._fetch_icon("https://img.icons8.com/ios-filled/20/ffffff/ms-excel.png", "excel_btn.png")
        self.icon_print = self._fetch_icon("https://img.icons8.com/ios-filled/20/ffffff/print.png", "print_btn.png")
        self.icon_calib = self._fetch_icon("https://img.icons8.com/ios-filled/20/ffffff/settings.png", "gear_btn.png")
        self.icon_reset = self._fetch_icon("https://img.icons8.com/ios-filled/16/ffffff/restart.png", "reset_btn.png")
        
        actions_frame = tk.Frame(self.frame_header, bg=self.bg_color)
        actions_frame.pack(side="right", anchor="e")
        
        # Botón Abrir Excel (Verde Esmeralda)
        self.btn_excel = self._create_hover_button(
            actions_frame, 
            "  ABRIR EXCEL", 
            self.icon_excel, 
            self.accent_green, 
            self.accent_green_hover, 
            self.load_excel
        )
        self.btn_excel.pack(side="left", padx=5)
        
        # Botón Calibrar (Amber/Naranja)
        self.btn_calibrate = self._create_hover_button(
            actions_frame, 
            "  CALIBRAR", 
            self.icon_calib, 
            self.accent_amber, 
            self.accent_amber_hover, 
            self.calibrate_printer
        )
        self.btn_calibrate.pack(side="left", padx=5)

        # Campo de Copias por etiqueta (Impresión Múltiple)
        copies_frame = tk.Frame(actions_frame, bg=self.bg_color)
        copies_frame.pack(side="left", padx=(5, 10))

        lbl_copies = tk.Label(copies_frame, text="Copias:", bg=self.bg_color, fg=self.fg_muted, font=("Segoe UI", 9, "bold"))
        lbl_copies.pack(side="left", padx=(0, 6))

        vcmd_copies = (self.register(self._validate_copies_input), "%P")
        self.spin_copies = tk.Spinbox(
            copies_frame,
            from_=1,
            to=999,
            width=4,
            justify="center",
            validate="key",
            validatecommand=vcmd_copies,
            bg=self.card_color,
            fg=self.fg_color,
            buttonbackground=self.border_color,
            insertbackground=self.fg_color,
            relief="flat",
            highlightthickness=1,
            highlightbackground=self.border_color,
            highlightcolor=self.accent_blue,
            font=("Segoe UI", 9, "bold")
        )
        self.spin_copies.delete(0, "end")
        self.spin_copies.insert(0, "1")
        self.spin_copies.pack(side="left", ipady=4)

        # Botón Imprimir (degradado Petróleo -> Azul)
        self.btn_print = self._create_hover_button(
            actions_frame,
            "  IMPRIMIR SELECCIÓN",
            self.icon_print,
            self.grad_print[0],
            self.grad_print[1],
            self.print_labels
        )
        self.btn_print.pack(side="left", padx=5)

        # ---------------------------------------------
        # ÁREA CENTRAL: PANEDWINDOW principal (Arriba: Config y Vista Previa, Abajo: Tabla)
        # ---------------------------------------------
        self.paned_main = tk.PanedWindow(self, orient=tk.VERTICAL, bd=0, bg=self.bg_color, sashwidth=10, sashrelief="flat")
        self.paned_main.pack(expand=True, fill="both", padx=25, pady=(5, 10))

        # --- SECCIÓN SUPERIOR: AJUSTES Y VISTA PREVIA ---
        self.frame_top_section = tk.Frame(self.paned_main, bg=self.bg_color)
        self.paned_main.add(self.frame_top_section, minsize=310, stretch="always", height=340)
        
        # Split horizontal dentro de la sección superior: Ajustes y Vista Previa con igual ancho/alto
        self.frame_top_section.columnconfigure(0, weight=1, uniform="top_panels") # 50% ancho para Ajustes
        self.frame_top_section.columnconfigure(1, weight=1, uniform="top_panels") # 50% ancho para Vista Previa
        self.frame_top_section.rowconfigure(0, weight=1)

        # Tarjeta Izquierda: Ajustes
        card_left, body_left = self._create_card(self.frame_top_section, "CONFIGURACIÓN Y AJUSTES DE IMPRESIÓN")
        card_left.grid(row=0, column=0, sticky="nswe", padx=(0, 10), pady=10)
        
        # Tarjeta Derecha: Vista Previa
        card_right, body_right = self._create_card(self.frame_top_section, "VISTA PREVIA DE ETIQUETA RENDERIZADA")
        card_right.grid(row=0, column=1, sticky="nswe", padx=(10, 0), pady=10)

        # Grid para controles en el panel izquierdo
        body_left.columnconfigure(1, weight=1)
        
        # 1. Selector de Plantillas
        lbl_zpl = tk.Label(body_left, text="Plantilla ZPL:", bg=self.card_color, fg=self.fg_muted, font=("Segoe UI", 9, "bold"))
        lbl_zpl.grid(row=0, column=0, sticky="e", pady=5, padx=(0, 10))

        self.combo_prn = ttk.Combobox(body_left, state="readonly")
        self.combo_prn.grid(row=0, column=1, sticky="we", pady=5)
        self.combo_prn.bind("<<ComboboxSelected>>", self.on_prn_selected)

        # 2. Selector de Impresora
        lbl_printer = tk.Label(body_left, text="Impresora:", bg=self.card_color, fg=self.fg_muted, font=("Segoe UI", 9, "bold"))
        lbl_printer.grid(row=1, column=0, sticky="e", pady=5, padx=(0, 10))

        self.combo_printer = ttk.Combobox(body_left, state="readonly")
        self.combo_printer.grid(row=1, column=1, sticky="we", pady=5)

        # Separador sutil
        sep = tk.Frame(body_left, bg=self.border_color, height=1)
        sep.grid(row=2, column=0, columnspan=2, sticky="we", pady=8)

        # 3. Alineación Horizontal
        lbl_x = tk.Label(body_left, text="Alineación Horiz. (X):", bg=self.card_color, fg=self.fg_color, font=("Segoe UI", 9))
        lbl_x.grid(row=3, column=0, sticky="e", pady=4, padx=(0, 10))

        self.slider_x = ttk.Scale(body_left, from_=-15, to=15, orient="horizontal", command=self.on_slider_change)
        self.slider_x.set(0)
        self.slider_x.grid(row=3, column=1, sticky="we", pady=4)

        self.lbl_val_x = tk.Label(body_left, text="0 mm", bg=self.card_color, fg=self.accent_blue, font=("Segoe UI", 9, "bold"), width=6)
        self.lbl_val_x.grid(row=3, column=2, padx=(5, 0))

        # 4. Alineación Vertical
        lbl_y = tk.Label(body_left, text="Alineación Vert. (Y):", bg=self.card_color, fg=self.fg_color, font=("Segoe UI", 9))
        lbl_y.grid(row=4, column=0, sticky="e", pady=4, padx=(0, 10))

        self.slider_y = ttk.Scale(body_left, from_=-15, to=15, orient="horizontal", command=self.on_slider_change)
        self.slider_y.set(0)
        self.slider_y.grid(row=4, column=1, sticky="we", pady=4)

        self.lbl_val_y = tk.Label(body_left, text="0 mm", bg=self.card_color, fg=self.accent_blue, font=("Segoe UI", 9, "bold"), width=6)
        self.lbl_val_y.grid(row=4, column=2, padx=(5, 0))

        # 5. Botón Restaurar Ajustes
        self.btn_reset = self._create_hover_button(
            body_left,
            " Restaurar Ajustes",
            self.icon_reset,
            self.accent_gray,
            self.accent_gray_hover,
            self.reset_sliders
        )
        self.btn_reset.grid(row=5, column=1, sticky="w", pady=(6, 0))

        # Separador para interruptor
        sep2 = tk.Frame(body_left, bg=self.border_color, height=1)
        sep2.grid(row=6, column=0, columnspan=2, sticky="we", pady=8)

        # 6. Switch Vista Previa Web
        lbl_toggle = tk.Label(body_left, text="Vista Previa Web:", bg=self.card_color, fg=self.fg_color, font=("Segoe UI", 9, "bold"))
        lbl_toggle.grid(row=7, column=0, sticky="e", padx=(0, 10))

        switch_frame = tk.Frame(body_left, bg=self.card_color)
        switch_frame.grid(row=7, column=1, sticky="w")
        
        self.preview_switch_canvas = tk.Canvas(
            switch_frame,
            width=50,
            height=26,
            bg=self.card_color,
            bd=0,
            highlightthickness=0,
            cursor="hand2"
        )
        self.preview_switch_canvas.pack(side="left")
        self.preview_switch_canvas.bind("<Button-1>", self.toggle_preview_api)
        
        self.preview_switch_state = tk.Label(
            switch_frame,
            text="OFF",
            bg=self.card_color,
            fg=self.fg_muted,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2"
        )
        self.preview_switch_state.pack(side="left", padx=(10, 0))
        self.preview_switch_state.bind("<Button-1>", self.toggle_preview_api)
        
        self._render_preview_switch()

        # El Canvas de Preview se empaca en el cuerpo derecho
        self.preview_canvas = tk.Label(
            body_right,
            bg=self.preview_bg,
            text="Seleccione un archivo PRN para visualizar",
            fg="#5B6672",
            font=("Segoe UI", 10)
        )
        self.preview_canvas.pack(expand=True, fill="both")
        self.preview_canvas.bind("<Configure>", self.on_preview_resize)

        # --- PANEL DE EXCEL ---
        self.frame_table_section = tk.Frame(self.paned_main, bg=self.bg_color)
        self.paned_main.add(self.frame_table_section, minsize=280, stretch="always", height=560)
        
        # Tarjeta de la Tabla
        card_table, body_table = self._create_card(self.frame_table_section, "REGISTROS DEL EXCEL PARA IMPRESIÓN")
        card_table.pack(expand=True, fill="both", pady=(10, 0))

        # Contenedor para Treeview y sus Scrollbars
        table_container = tk.Frame(body_table, bg=self.card_color)
        table_container.pack(expand=True, fill="both", pady=(5, 5))
        
        # Crear Scrollbars con nuestro estilo Dark
        self.scrolly = ttk.Scrollbar(table_container, orient="vertical", style="Dark.Vertical.TScrollbar")
        self.scrolly.pack(side="right", fill="y")
        
        self.scrollx = ttk.Scrollbar(table_container, orient="horizontal", style="Dark.Horizontal.TScrollbar")
        self.scrollx.pack(side="bottom", fill="x")
        
        self.tree = ttk.Treeview(
            table_container, 
            yscrollcommand=self.scrolly.set, 
            xscrollcommand=self.scrollx.set,
            style="Treeview"
        )
        self.tree.pack(expand=True, fill="both")
        
        self.scrolly.config(command=self.tree.yview)
        self.scrollx.config(command=self.tree.xview)
        
        # Eventos de tabla
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_row_selected)
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-1>", self.on_tree_checkbox_click)
        self.tree.bind("<Motion>", self._on_tree_motion)
        self.tree.bind("<Leave>", self._on_tree_leave)
        self._hover_item = None

        # Barra de estado inferior dentro de la tarjeta
        self.frame_status = tk.Frame(body_table, bg=self.card_color)
        self.frame_status.pack(fill="x", pady=(10, 0))

        self.lbl_status_rows = tk.Label(
            self.frame_status,
            text="Sin datos cargados. Abre un archivo Excel.",
            bg=self.card_color,
            fg=self.fg_muted,
            font=("Segoe UI", 9, "italic")
        )
        self.lbl_status_rows.pack(side="left")

        # Accesos rápidos de selección masiva (equivalen al checkbox del encabezado)
        quick_select_frame = tk.Frame(self.frame_status, bg=self.card_color)
        quick_select_frame.pack(side="left", padx=(18, 0))

        lbl_select_all = tk.Label(
            quick_select_frame, text="Marcar todo", bg=self.card_color,
            fg=self.accent_blue, font=("Segoe UI", 9, "underline"), cursor="hand2"
        )
        lbl_select_all.pack(side="left")
        lbl_select_all.bind("<Button-1>", lambda e: self.tree.selection_set(self.tree.get_children()))

        lbl_sep_quick = tk.Label(quick_select_frame, text=" | ", bg=self.card_color, fg=self.fg_muted, font=("Segoe UI", 9))
        lbl_sep_quick.pack(side="left")

        lbl_deselect_all = tk.Label(
            quick_select_frame, text="Quitar selección", bg=self.card_color,
            fg=self.fg_muted, font=("Segoe UI", 9, "underline"), cursor="hand2"
        )
        lbl_deselect_all.pack(side="left")
        lbl_deselect_all.bind("<Button-1>", lambda e: self.tree.selection_remove(self.tree.get_children()))

        self.lbl_status_selection = tk.Label(
            self.frame_status,
            text="Seleccionados: 0 etiquetas",
            bg=self.bg_color,
            fg=self.accent_blue,
            font=("Segoe UI", 9, "bold"),
            padx=10,
            pady=3
        )
        self.lbl_status_selection.pack(side="right")

    def _load_local_data(self):
        # 1. Impresoras
        try:
            printers = [p[2] for p in win32print.EnumPrinters(2)]
            self.combo_printer['values'] = printers
            if printers:
                self.combo_printer.current(0)
        except Exception as e:
            print("Error cargando impresoras:", e)

        # 2. Plantillas PRN
        if not os.path.exists(self.prn_folder):
            try:
                os.makedirs(self.prn_folder)
            except:
                pass
                
        if os.path.exists(self.prn_folder):
            archivos = sorted(f for f in os.listdir(self.prn_folder) if f.endswith(".prn"))
            self.combo_prn['values'] = archivos
            if archivos:
                self.combo_prn.current(0)

                def _finalize_prn_selection():
                    # Se difiere a después del primer ciclo de mainloop: fuerza a Tk a
                    # repintar el texto seleccionado (evita que el combo se vea vacío
                    # en el primer render) y luego dispara la carga automática.
                    self.combo_prn.set(archivos[0])
                    self.on_prn_selected(None)

                self.after(50, _finalize_prn_selection)

    # =========================================================================
    # EVENTOS DE USUARIO
    # =========================================================================

    def load_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx;*.xls")])
        if not file_path:
            return
            
        try:
            self.df = read_excel_light(file_path)
            
            # Asociar las variables requeridas en el ZPL automáticamente (Invisible)
            self.auto_map_variables()
            
            # Dibujar la cuadrícula
            self._fill_treeview()
            
            # Generar preview de la fila 0 por default
            self.generate_preview(row_index=0)
            
        except Exception as e:
            messagebox.showerror("Error Interno", f"No se pudo leer el archivo Excel:\n{str(e)}")

    def on_prn_selected(self, event):
        prn_file = self.combo_prn.get()
        if not prn_file: return
        
        path = os.path.join(self.prn_folder, prn_file)
        try:
            with open(path, encoding="utf-8", errors="ignore") as f:
                self.plantilla_zpl = f.read()
                
            # Extraer nombres entre corchetes, Ej: [vlsCampo]
            self.variables_prn = list(set(re.findall(r"\[([^\]]+)\]", self.plantilla_zpl)))
            
            # Auto Mapear si ya hay un excel
            self.auto_map_variables()
            
            # Si se hace click justo al abrir, que muestre la fila que esté seleccionada, sino la 0
            self.generate_preview(row_index=self._get_selected_row_index())
            
        except Exception as e:
            messagebox.showerror("Error Lectura", f"Fallo al abrir archivo PRN:\n{str(e)}")

    def auto_map_variables(self):
        """Asocia invisiblemente cada variable del PRN con la cabecera de Excel posicionalmente y por nombre"""
        self.map_columnas.clear()
        if self.df is None or not self.variables_prn:
            return
            
        col_names = list(self.df.columns)
        col_names_lower = [str(c).strip().lower() for c in col_names]
        
        for var in self.variables_prn:
            var_norm = var.strip().lower()
            
            # --- 1. Mapeo Posicional Estricto (Nuevo estándar) ---
            # Si el PRN pide vlscampo1, va a la columna 0 (A). vlscampo2 va a la 1 (B), etc.
            match_num = re.search(r'(\d+)$', var_norm)
            if match_num and "vlscampo" in var_norm:
                idx = int(match_num.group(1)) - 1
                if 0 <= idx < len(col_names):
                    self.map_columnas[var] = col_names[idx]
                    continue
                    
            # Compatibilidad si usaste "vlscampo" clásico (sin número) -> Pasa a Col A
            # Solo aplica si la plantilla NO usa el nuevo estándar "vlscampo1", para evitar confilctos
            if var_norm == 'vlscampo':
                if not any(v.strip().lower() == 'vlscampo1' for v in self.variables_prn):
                    if len(col_names) > 0:
                        self.map_columnas[var] = col_names[0]
                        continue
                else:
                    # Es un huérfano/error de diseño en ZPL, lo dejamos sin mapear para que salga vacío
                    continue
                    
            # --- 2. Mapeo Histórico / Texto Exacto (Ej. "vlsCodigoQR") ---
            if var_norm in col_names_lower:
                idx = col_names_lower.index(var_norm)
                self.map_columnas[var] = col_names[idx]

    def _fill_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._hover_item = None

        if self.df is None or self.df.empty: return

        cols = list(self.df.columns)
        # Columna de checkbox de selección, ubicada junto a la primera columna (CÓDIGO)
        self.tree["columns"] = [self.CHECK_COL] + cols
        self.tree["show"] = "headings"

        self.tree.heading(self.CHECK_COL, text="☐", command=self.toggle_select_all)
        self.tree.column(self.CHECK_COL, width=34, minwidth=34, anchor="center", stretch=False)

        # Ancho de columna auto-ajustado al contenido (máx. primeras 200 filas, por rendimiento)
        sample_size = min(len(self.df), 200)
        for col_idx, col in enumerate(cols):
            self.tree.heading(col, text=col)

            max_len = len(str(col))
            for row_idx in range(sample_size):
                val = self.df.iloc[row_idx].values[col_idx]
                if notna(val):
                    val_len = len(str(val).strip())
                    if val_len > max_len:
                        max_len = val_len

            width = min(260, max(90, max_len * 7 + 20))
            self.tree.column(col, width=width, anchor="w")

        self.tree.tag_configure("evenrow", background=self.row_even)
        self.tree.tag_configure("oddrow", background=self.row_odd)
        self.tree.tag_configure("hoverrow", background=self.row_hover)

        # Para imprimir selecciones manualmente, necesitamos poder ver todas las filas posibles. Límite de seguridad 10000.
        limit = min(10000, len(self.df))
        for index in range(limit):
            valores = [str(x).strip() if notna(x) else "" for x in self.df.iloc[index].values]
            tag = "evenrow" if index % 2 == 0 else "oddrow"
            item_id = self.tree.insert("", "end", values=["☐"] + valores, tags=(str(index), tag))

        self.lbl_status_rows.config(text=f"Total de registros: {len(self.df)} filas.")

    def on_tree_checkbox_click(self, event):
        """Alterna la selección de una fila al hacer clic en su casilla de verificación."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column_id = self.tree.identify_column(event.x)
        if column_id != "#1":  # La columna de checkbox siempre es la primera
            return

        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return

        self.tree.selection_toggle(row_id)
        return "break"  # Evita que el binding por defecto de Treeview colapse la multi-selección

    def _on_tree_motion(self, event):
        """Resalta suavemente la fila bajo el cursor para facilitar la lectura/selección."""
        row_id = self.tree.identify_row(event.y)
        if row_id == self._hover_item:
            return

        self._clear_hover()

        if row_id:
            tags = list(self.tree.item(row_id, "tags"))
            if "hoverrow" not in tags:
                tags.append("hoverrow")
                self.tree.item(row_id, tags=tags)
            self._hover_item = row_id

    def _on_tree_leave(self, event):
        self._clear_hover()

    def _clear_hover(self):
        prev = self._hover_item
        if prev and self.tree.exists(prev):
            tags = [t for t in self.tree.item(prev, "tags") if t != "hoverrow"]
            self.tree.item(prev, tags=tags)
        self._hover_item = None

    def toggle_select_all(self):
        """Marca/desmarca todas las filas visibles (checkbox del encabezado de la columna)."""
        all_items = self.tree.get_children()
        if not all_items:
            return

        if len(self.tree.selection()) == len(all_items):
            self.tree.selection_remove(all_items)
        else:
            self.tree.selection_set(all_items)

    def on_tree_row_selected(self, event):
        selected_set = set(self.tree.selection())

        # Refresca el glifo de checkbox de cada fila según su estado real de selección
        for item in self.tree.get_children():
            self.tree.set(item, self.CHECK_COL, "☑" if item in selected_set else "☐")

        all_items = self.tree.get_children()
        all_selected = bool(all_items) and len(selected_set) == len(all_items)
        self.tree.heading(self.CHECK_COL, text="☑" if all_selected else "☐")

        num_selected = len(selected_set)
        self.lbl_status_selection.config(text=f"Seleccionados: {num_selected} etiqueta(s)")
        self.generate_preview(row_index=self._get_selected_row_index())

    def _get_selected_row_index(self):
        selected = self.tree.selection()
        if selected:
            tags = self.tree.item(selected[0], "tags")
            if tags:
                return int(tags[0])
        return 0

    def on_double_click(self, event):
        # Destruir editor previo si existe (previene superposición)
        if hasattr(self, 'entry_edit') and self.entry_edit.winfo_exists():
            self.entry_edit.destroy()

        # Aislar celda interactuada
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell": return
        
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not row_id or not column_id: return

        if column_id == "#1":  # Columna de checkbox: no es editable como texto
            return

        # Tkinter guarda columnas internamente como #1, #2, etc. Realizamos offset (-1) para array real
        col_index = int(column_id.replace("#", "")) - 1
        
        # Geometría del bounding box (ancho/alto de la celda tocada)
        x, y, width, height = self.tree.bbox(row_id, column_id)
        current_value = self.tree.set(row_id, column_id)
        
        # 1. Insertamos caja texto temporal y absorbemos control
        self.entry_edit = ttk.Entry(self.tree, font=("Segoe UI", 8))
        self.entry_edit.place(x=x, y=y, width=width, height=height)
        self.entry_edit.insert(0, current_value)
        self.entry_edit.select_range(0, tk.END)
        self.entry_edit.focus()
        
        # 2. Reglas de Salida
        self.entry_edit.bind("<Return>", lambda e: self.save_edit(row_id, column_id, col_index))
        # Esc permite escapar si nos equivocamos de doble click
        self.entry_edit.bind("<Escape>", lambda e: self.entry_edit.destroy())

    def save_edit(self, row_id, column_id, col_index):
        if not hasattr(self, 'entry_edit') or not self.entry_edit.winfo_exists():
            return
            
        new_val = self.entry_edit.get()
        self.entry_edit.destroy()
        
        # 1. Salvar FrontEnd GUI Visual
        self.tree.set(row_id, column_id, new_val)
        
        # 2. Salvar BackEnd MEMORIA Pandas
        tags = self.tree.item(row_id, "tags")
        if tags and self.df is not None:
            df_index = int(tags[0])
            real_col_name = self.tree["columns"][col_index]
            self.df.at[df_index, real_col_name] = new_val
            
            # 3. Auto Refrescar y Renderizar para que se note en pantalla grande
            selected_items = self.tree.selection()
            if selected_items and selected_items[0] == row_id:
                self.generate_preview(row_index=df_index)

    def on_preview_resize(self, event):
        if not getattr(self, '_current_preview_img', None):
            return
            
        width = event.width
        height = event.height
        
        if width < 50 or height < 50: return # Ignorar encogimientos irreales
        
        if self._resize_timer is not None:
            self.after_cancel(self._resize_timer)
        # Redimensiona la imagen después de soltar o detener el arrastre de la ventana
        self._resize_timer = self.after(150, lambda: self._apply_resize(width, height))

    def _apply_resize(self, width, height):
        if not getattr(self, '_current_preview_img', None): return
        img = self._current_preview_img.copy()

        try:
            resample_filter = Image.Resampling.LANCZOS # Alta calidad en Pillow modernas
        except AttributeError:
            resample_filter = Image.ANTIALIAS

        margin = 24  # espacio reservado para la sombra alrededor de la etiqueta
        img.thumbnail((max(1, width - margin), max(1, height - margin)), resample_filter)

        composed = self._compose_label_with_shadow(img, width, height)
        img_tk = ImageTk.PhotoImage(composed)
        self.preview_canvas.config(image=img_tk, text="")
        self.preview_canvas.image = img_tk

    def _compose_label_with_shadow(self, label_img, canvas_w, canvas_h):
        """Simula una etiqueta física: sombra suave detrás, sin redondear ni alterar su contenido real."""
        label_img = label_img.convert("RGBA")
        lw, lh = label_img.size
        canvas_w, canvas_h = max(canvas_w, lw + 20), max(canvas_h, lh + 20)

        base = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
        ox, oy = (canvas_w - lw) // 2, (canvas_h - lh) // 2

        shadow = Image.new("RGBA", (canvas_w, canvas_h), (0, 0, 0, 0))
        ImageDraw.Draw(shadow).rectangle([ox + 2, oy + 4, ox + lw + 2, oy + lh + 4], fill=(0, 0, 0, 110))
        shadow = shadow.filter(ImageFilter.GaussianBlur(8))
        base.alpha_composite(shadow)

        base.alpha_composite(label_img, (ox, oy))
        return base

    def on_slider_change(self, *args):
        # Bloqueo de seguridad para evitar errores al crear los widgets al abrir la app
        if not hasattr(self, 'slider_x') or not hasattr(self, 'slider_y') or not hasattr(self, 'lbl_val_y'):
            return
            
        # Actualizar labels visuales
        val_x = int(self.slider_x.get())
        val_y = int(self.slider_y.get())
        self.lbl_val_x.config(text=f"{val_x} mm")
        self.lbl_val_y.config(text=f"{val_y} mm")
        
        if self._preview_timer is not None:
            self.after_cancel(self._preview_timer)
        # Debounce de 400ms para no saturar Labelary mientras el usuario arrastra la perilla
        self._preview_timer = self.after(400, lambda: self.generate_preview(row_index=self._get_selected_row_index()))

    def reset_sliders(self):
        self.slider_x.set(0)
        self.slider_y.set(0)
        self.on_slider_change()

    def _validate_copies_input(self, valor_propuesto):
        # Permite vacío (mientras se edita) o solo dígitos, hasta 3 cifras (máx 999)
        return valor_propuesto == "" or (valor_propuesto.isdigit() and len(valor_propuesto) <= 3)

    def _get_copies_count(self):
        try:
            copias = int(self.spin_copies.get())
        except (ValueError, AttributeError):
            copias = 1
        return max(1, copias)

    def _render_preview_switch(self, knob_x=None):
        if not hasattr(self, 'preview_switch_canvas'):
            return

        is_on = self.preview_api_enabled.get()
        canvas = self.preview_switch_canvas
        canvas.delete("all")

        track_color = "#38A169" if is_on else "#4A5568"
        if knob_x is None:
            knob_x = 28 if is_on else 4

        knob_x = int(round(max(4, min(28, knob_x))))
        self._switch_knob_x = knob_x

        # Pista tipo pastilla (simulada con ovalos + rectangulo)
        canvas.create_oval(2, 2, 26, 26, fill=track_color, outline=track_color)
        canvas.create_rectangle(14, 2, 38, 26, fill=track_color, outline=track_color)
        canvas.create_oval(26, 2, 50, 26, fill=track_color, outline=track_color)

        # Perilla
        canvas.create_oval(knob_x, 4, knob_x + 20, 24, fill="#F7FAFC", outline="#E2E8F0")

        if hasattr(self, 'preview_switch_state'):
            self.preview_switch_state.config(
                text="ON" if is_on else "OFF",
                fg="#C6F6D5" if is_on else "#CBD5E0"
            )

    def _animate_preview_switch(self, target_on):
        if not hasattr(self, 'preview_switch_canvas'):
            return

        if self._switch_anim_timer is not None:
            self.after_cancel(self._switch_anim_timer)
            self._switch_anim_timer = None

        start_x = getattr(self, '_switch_knob_x', 4)
        end_x = 28 if target_on else 4
        steps = 8
        delta = (end_x - start_x) / steps if steps else 0

        def step(frame=0, current_x=start_x):
            self._render_preview_switch(knob_x=current_x)
            if frame < steps:
                self._switch_anim_timer = self.after(16, lambda: step(frame + 1, current_x + delta))
            else:
                self._switch_anim_timer = None
                self._render_preview_switch(knob_x=end_x)

        step()

    def toggle_preview_api(self, event=None):
        self.preview_api_enabled.set(not self.preview_api_enabled.get())
        self.on_preview_toggle(animated=True)

    def on_preview_toggle(self, animated=False):
        if animated:
            self._animate_preview_switch(self.preview_api_enabled.get())
        else:
            self._render_preview_switch()

        if self.preview_api_enabled.get():
            self.generate_preview(row_index=self._get_selected_row_index())
        else:
            self._current_preview_img = None
            self.preview_canvas.config(image='', text="Preview web desactivada. Activa el switch para usar Labelary.", fg="#5B6672")

    def _shift_zpl(self, zpl, mm_x, mm_y):
        if mm_x == 0 and mm_y == 0:
            return zpl
            
        dots_x = int(mm_x * 8) # Asumiendo impresoras típicas de 8dpmm (203 dpi)
        dots_y = int(mm_y * 8)
        
        def mover_punto(match):
            cmd = match.group(1) # Extrae ^FO o ^FT
            try:
                x = max(0, int(match.group(2)) + dots_x)
                y = max(0, int(match.group(3)) + dots_y)
                return f"{cmd}{x:03d},{y:03d}"
            except:
                return match.group(0) # Si algo excepcionalmente falla, dejar original
                
        # Regex captura todos los comandos de coordenadas en ZPL Ej: ^FO100,200 o ^FT30,40
        return re.sub(r'(\^F[OT])(\d+),(\d+)', mover_punto, zpl, flags=re.IGNORECASE)


    # =========================================================================
    # CORE: ENSAMBLADOR ZPL Y CONEXIÓN A LABELARY
    # =========================================================================

    def _get_series_value(self, row_series, col_name):
        val = row_series.get(col_name)
        if notna(val):
            return str(val).strip()
        return ""

    def _resolve_placeholder_value(self, var_name, row_series, max_len=None):
        """Resuelve el valor de un placeholder PRN y soporta variantes tipo vlsCampo2A."""
        # 1) Mapeo directo (compatibilidad actual)
        if var_name in self.map_columnas:
            value = self._get_series_value(row_series, self.map_columnas[var_name])
        else:
            value = ""

        # 2) Fallback para campos posicionales con sufijo de segmento: vlsCampo2A, vlsCampo4C, etc.
        seg_match = re.match(r'(?i)^vlscampo(\d+)([a-z])$', var_name.strip())
        if seg_match:
            base_idx = int(seg_match.group(1)) - 1
            part_idx = ord(seg_match.group(2).upper()) - ord('A')

            if not value and self.df is not None and 0 <= base_idx < len(self.df.columns):
                base_col = self.df.columns[base_idx]
                value = self._get_series_value(row_series, base_col)

            if max_len and max_len > 0:
                chunks = textwrap.wrap(value, width=max_len, break_long_words=True, break_on_hyphens=False)
                return chunks[part_idx] if 0 <= part_idx < len(chunks) else ""

            return value

        # 3) Fallback para campos posicionales simples: vlsCampo1, vlsCampo2, etc.
        num_match = re.match(r'(?i)^vlscampo(\d+)$', var_name.strip())
        if not value and num_match and self.df is not None:
            base_idx = int(num_match.group(1)) - 1
            if 0 <= base_idx < len(self.df.columns):
                base_col = self.df.columns[base_idx]
                value = self._get_series_value(row_series, base_col)

        return value

    def compose_zpl(self, row_series):
        """Compone la macro en base a la línea especificada de los datos de Excel."""
        if not self.plantilla_zpl: return ""
        zpl = self.plantilla_zpl

        # Reemplaza placeholders con posible sufijo de longitud: [vlsCampo2A]34
        # El sufijo ]NN no debe imprimirse; se usa para segmentar texto automáticamente.
        def replace_placeholder(match):
            var_name = match.group(1)
            max_len_str = match.group(2)
            max_len = int(max_len_str) if max_len_str else None
            return self._resolve_placeholder_value(var_name, row_series, max_len=max_len)

        zpl = re.sub(r'\[([^\]]+)\](\d+)?', replace_placeholder, zpl)
            
        # PARCHE SINTAXIS: Resolver comandos ^FD sin cierre que arruinan la API/impresora
        lineas = zpl.split('\n')
        for i in range(len(lineas)):
            if '^FD' in lineas[i] and '^FS' not in lineas[i]:
                lineas[i] = lineas[i].rstrip('\r\n') + '^FS'
        
        zpl = '\n'.join(lineas)
        
        # INYECTAR DESPLAZAMIENTOS CALIBRADOS (Sliders X/Y) a las coordenadas del código
        val_x = int(self.slider_x.get())
        val_y = int(self.slider_y.get())
        zpl = self._shift_zpl(zpl, val_x, val_y)
        
        return zpl

    def generate_preview(self, row_index=0):
        if not self.plantilla_zpl:
            return

        if not self.preview_api_enabled.get():
            self._current_preview_img = None
            self.preview_canvas.config(image='', text="Preview web desactivada. Activa el switch para usar Labelary.", fg="#5B6672")
            return
            
        if self.df is None or self.df.empty:
            # Estado base: Solo PRN cargado -> Mostrar ZPL usando las literales [variables] para confirmar posiciones
            zpl = self.plantilla_zpl
            for var in self.variables_prn:
                # Si una variable tiene corchetes, reescribela con los mismos para identificarla
                zpl = zpl.replace(f"[{var}]", f"[{var}]")
                
            val_x = int(self.slider_x.get())
            val_y = int(self.slider_y.get())
            zpl = self._shift_zpl(zpl, val_x, val_y)
        else:
            if row_index < len(self.df):
                zpl = self.compose_zpl(self.df.iloc[row_index])
            else: return

        self._request_labelary(zpl)

    def _request_labelary(self, zpl):
        try:
            import requests
            self.preview_canvas.config(image='', text="Generando previsualización...", fg="#0D7AA0")
            self.update_idletasks() # Dibujarlo inmediatamente para no congelar app ciegamente
            
            # Aislar únicamente la etiqueta válida del final
            bloques = re.findall(r"\^XA.*?\^XZ", zpl, re.IGNORECASE | re.DOTALL)
            zpl_preview = bloques[-1] if bloques else zpl
            
            # PARCHE FUENTES: Compatibilidad obligatoria en labelary para E:ARIAL.FNT (^A@)
            zpl_preview = re.sub(r'\^A@([NRIBnrib]?)[,]*(\d*)[,]*(\d*)[^\^]*', r'^A0\1,\2,\3', zpl_preview)

            # Auto-dimensiones para que sea siempre la etiqueta exacta sin importar si es 4x2, 10x15, etc
            w_inches, h_inches = 4.0, 4.0
            pw_match = re.search(r"\^PW(\d+)", zpl, re.IGNORECASE)
            ll_match = re.search(r"\^LL(\d+)", zpl, re.IGNORECASE)
            if pw_match and ll_match:
                w_inches = max(1.0, round(int(pw_match.group(1)) / 203.0, 2))
                h_inches = max(1.0, round(int(ll_match.group(1)) / 203.0, 2))

            url = f"http://api.labelary.com/v1/printers/8dpmm/labels/{w_inches}x{h_inches}/0/"
            headers = {"Accept": "image/png"}

            res = requests.post(url, data=zpl_preview.encode('utf-8'), headers=headers, timeout=10)

            if res.status_code == 200:
                self._current_preview_img = Image.open(BytesIO(res.content))
                
                # Proyectar inicialmente sobre el espacio disponible del contenedor
                canvas_w = self.preview_canvas.winfo_width()
                canvas_h = self.preview_canvas.winfo_height()
                if canvas_w < 50: canvas_w = 700
                if canvas_h < 50: canvas_h = 400
                
                self._apply_resize(canvas_w, canvas_h)
            else:
                self._current_preview_img = None
                self.preview_canvas.config(image='', text=f"El servidor rechazó el código ZPL. (Error {res.status_code})", fg="#C53030")

        except Exception as e:
            self._current_preview_img = None
            self.preview_canvas.config(image='', text=f"Error en la conexión a la web:\n{str(e)[:50]}...", fg="#C53030")

    # =========================================================================
    # IMPRESIÓN MASIVA
    # =========================================================================

    def calibrate_printer(self):
        printer_name = self.combo_printer.get()
        if not printer_name:
            messagebox.showwarning("Calibrar", "Selecciona una impresora válida de la lista.")
            return

        respuesta = messagebox.askyesno("Calibrar Impresora", f"¿Estás seguro de enviar el comando de calibración a la impresora '{printer_name}'?")
        if not respuesta: return

        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            try:
                hJob = win32print.StartDocPrinter(hPrinter, 1, ("Calibracion_Zebra", None, "RAW"))
                win32print.StartPagePrinter(hPrinter)
                
                # Comando ZPL para calibrar la etiqueta
                win32print.WritePrinter(hPrinter, b"~JC")
                
                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
            finally:
                win32print.ClosePrinter(hPrinter)
                
            messagebox.showinfo("Calibración", "Comando de calibración (~JC) enviado exitosamente a la impresora.")
            
        except Exception as e:
            messagebox.showerror("Error Físico", f"Falla conectando con el driver Windows Zebra:\n{str(e)}")

    def print_labels(self):
        printer_name = self.combo_printer.get()
        if not printer_name:
            messagebox.showwarning("Imprimir", "Selecciona una impresora válida de la lista.")
            return
            
        if self.df is None or self.df.empty:
            messagebox.showwarning("Imprimir", "Aún no has cargado un set de datos de Excel o está vacío.")
            return

        # Filtrar solo elementos seleccionados por el usuario en la tabla
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Imprimir", "No hay renglones seleccionados.\n\nPor favor, da clic en los renglones de la tabla que deseas mandar a imprimir.")
            return

        copias = self._get_copies_count()
        total_etiquetas = len(selected_items) * copias

        texto_confirmacion = f"¿Estás seguro de enviar a imprimir {len(selected_items)} etiqueta(s) seleccionada(s)"
        if copias > 1:
            texto_confirmacion += f" x {copias} copia(s) cada una (total {total_etiquetas} etiquetas)"
        texto_confirmacion += f" a la impresora '{printer_name}'?"

        respuesta = messagebox.askyesno("Impresión", texto_confirmacion)
        if not respuesta: return

        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            try:
                hJob = win32print.StartDocPrinter(hPrinter, 1, ("Impresion_Etiquetas_PRN", None, "RAW"))
                win32print.StartPagePrinter(hPrinter)

                # Iterar SÓLO las filas que el usuario marcó
                for item in selected_items:
                    tags = self.tree.item(item, "tags")
                    if tags:
                        idx = int(tags[0])
                        row = self.df.iloc[idx]
                        zpl_final = self.compose_zpl(row)
                        zpl_bytes = zpl_final.encode("utf-8")
                        # Impresión múltiple: repite la misma etiqueta N veces (copias)
                        for _ in range(copias):
                            win32print.WritePrinter(hPrinter, zpl_bytes)

                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
            finally:
                win32print.ClosePrinter(hPrinter)

            messagebox.showinfo("Proceso Completo", f"¡Se generaron y enviaron {total_etiquetas} comandos ZPL de forma exitosa hacia la impresora!")
            
        except Exception as e:
            messagebox.showerror("Error Físico", f"Falla conectando con el driver Windows Zebra:\n{str(e)}")


if __name__ == "__main__":
    app = ZebraPrintApp()
    app.mainloop()