#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador de Certificados — Wizard (trimmed preview, no width control)

- Any .ttf font (required).
- Defaults: Auto Snap = ON, Consistent size = OFF.
- Step 2 area selection is optional (auto-detect if skipped).
- Full-page preview shows a tiny 4-up (fallback 3/2) **of the modified page**
  and the preview canvas **shrinks to the exact strip height** (no empty bands).
- Focused preview also shrinks vertically to the image height.
- Fixed window size; big Generate button; progress bar.

pip install pymupdf pillow
"""

import os, sys, csv, json, datetime, re, subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from dataclasses import dataclass
from typing import List, Optional, Tuple
from functools import lru_cache

import fitz  # PyMuPDF
from PIL import Image, ImageTk, ImageFont

APP_TITLE = "Gerador de Certificados — Wizard"
CALIBRATION_SUFFIX = ".calibration.json"

# Brand colors
C_BG      = "#ededed"
C_PRIMARY = "#17c88b"
C_TEAL    = "#009475"
C_DARK    = "#002927"
C_WARM    = "#514a43"

# Window & preview sizing
PREVIEW_TARGET_WIDTH  = 900
PREVIEW_TARGET_HEIGHT = 330   # max height we allow (image will be <= this)
WINDOW_WIDTH  = PREVIEW_TARGET_WIDTH + 120
WINDOW_HEIGHT = 760

# ----------------- Utils -----------------

def sanitize_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]", "_", name)
    return name.strip().strip(".")

def read_names_from_csv(path: str) -> List[str]:
    names: List[str] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(2048); f.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=";,")
            reader = csv.reader(f, dialect)
        except csv.Error:
            reader = csv.reader(f)
        for row in reader:
            if not row: continue
            cell = next((c.strip() for c in row if c and c.strip()), "")
            if not cell: continue
            cell = re.sub(r"[;,\s]+$", "", cell)
            cell = re.split(r"[;,]", cell)[0].strip()
            if cell: names.append(cell)
    if names and names[0].lower() in ("nome","name","aluno","participant","participante"):
        names = names[1:]
    return names

def timestamp() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d_%Hh%M")

@dataclass
class Calibration:
    page_index: int
    rect: Tuple[float, float, float, float]  # x0, y0, x1, y1 (PDF pts)

def load_calibration_for(template_pdf: str) -> Optional[Calibration]:
    meta = template_pdf + CALIBRATION_SUFFIX
    if os.path.exists(meta):
        try:
            with open(meta, "r", encoding="utf-8") as f:
                d = json.load(f)
            return Calibration(page_index=int(d["page_index"]), rect=tuple(d["rect"]))
        except Exception:
            return None
    return None

def save_calibration_for(template_pdf: str, calib: Calibration) -> None:
    meta = template_pdf + CALIBRATION_SUFFIX
    with open(meta, "w", encoding="utf-8") as f:
        json.dump({"page_index": calib.page_index, "rect": list(calib.rect)}, f, ensure_ascii=False, indent=2)

def ensure_ttf_font(font_path: str) -> str:
    if not font_path or not os.path.isfile(font_path):
        raise FileNotFoundError("Escolha um arquivo de fonte .ttf.")
    if not font_path.lower().endswith(".ttf"):
        raise ValueError("A fonte deve ser um arquivo .ttf (TrueType).")
    return font_path

def open_folder(path: str):
    try:
        if os.name == "nt":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass

# ----------------- Text metrics (PIL) -----------------

@lru_cache(maxsize=512)
def _pil_font(font_path: str, size_px: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(font_path, size_px)

def _advance_width(text: str, font_path: str, size_pt: float) -> float:
    size_px = max(1, int(round(size_pt)))
    font = _pil_font(font_path, size_px)
    if hasattr(font, "getlength"):
        return float(font.getlength(text))
    l, t, r, b = font.getbbox(text)
    return float(r - l)

def _line_metrics(font_path: str, size_pt: float) -> Tuple[float,float,float]:
    size_px = max(1, int(round(size_pt)))
    font = _pil_font(font_path, size_px)
    ascent, descent = font.getmetrics()
    return float(ascent), float(descent), float(ascent + descent)

def _fits_in_rect(text: str, rect: fitz.Rect, fontfile: str, size_pt: float) -> bool:
    pad_w = rect.width * 0.02
    pad_h = rect.height * 0.06
    target_w = rect.width - 2*pad_w
    target_h = rect.height - 2*pad_h
    w = _advance_width(text, fontfile, size_pt)
    _, _, line_h = _line_metrics(fontfile, size_pt)
    return (w <= target_w) and (line_h <= target_h)

def autosize_font_to_rect(text: str, rect: fitz.Rect, fontfile: str,
                          min_size=8.0, max_size=200.0) -> float:
    if not text: return 24.0
    lo, hi = float(min_size), float(max_size); best = lo
    while lo <= hi:
        mid = (lo + hi) / 2.0
        if _fits_in_rect(text, rect, fontfile, mid):
            best = mid; lo = mid + 0.5
        else:
            hi = mid - 0.5
    return max(min_size, min(best, max_size))

def common_font_size_for_all(names: List[str], rect: fitz.Rect, fontfile: str,
                             min_size=8.0, max_size=200.0) -> float:
    if not names: return 24.0
    lo, hi = float(min_size), float(max_size); best = lo
    while lo <= hi:
        mid = (lo + hi) / 2.0
        if all(_fits_in_rect(n, rect, fontfile, mid) for n in names):
            best = mid; lo = mid + 0.5
        else:
            hi = mid - 0.5
    return max(min_size, min(best, max_size))

# ----------------- Draw text -----------------

def draw_name_centered_with_size(page: fitz.Page, rect: fitz.Rect, name: str,
                                 ttf_path: str, fontsize: float):
    advance = _advance_width(name, ttf_path, fontsize)
    ascent, descent, _ = _line_metrics(ttf_path, fontsize)
    cx = rect.x0 + rect.width/2.0
    cy = rect.y0 + rect.height/2.0
    x = cx - advance/2.0
    y = cy + (ascent - descent)/2.0
    page.insert_text((x, y), name, fontsize=fontsize, fontfile=ttf_path,
                     fontname="UserFont", color=(0,0,0))

# ----------------- Snap (default ON) -----------------

def _get_text_blocks(page: fitz.Page):
    rects = []
    try:
        blocks = page.get_text("blocks") or []
    except Exception:
        blocks = []
    for b in blocks:
        if len(b) >= 5 and isinstance(b[4], str) and b[4].strip():
            x0, y0, x1, y1 = b[:4]
            rects.append(fitz.Rect(x0, y0, x1, y1))
    return rects

def _overlap_amount(a: fitz.Rect, b: fitz.Rect) -> float:
    return max(0, min(a.y1, b.y1) - max(a.y0, b.y0))

def compute_snapped_rect(page: fitz.Page, rect: fitz.Rect, snap_enabled: bool,
                         tol_pt: float, offset_x: float, offset_y: float) -> fitz.Rect:
    w, h = rect.width, rect.height
    cx, cy = rect.x0 + w/2, rect.y0 + h/2
    if not snap_enabled:
        snapped_cx, snapped_cy = cx, cy
    else:
        candidates_x = [page.rect.width/2]
        blocks = _get_text_blocks(page)
        overlapped = [b for b in blocks if _overlap_amount(b, rect) > 0]
        for b in overlapped:
            candidates_x.extend([b.x0, (b.x0+b.x1)/2, b.x1])
        snapped_cx = min(candidates_x, key=lambda x: abs(x - cx)) if candidates_x else cx
        if abs(snapped_cx - cx) > tol_pt: snapped_cx = cx
        candidates_y = [((b.y0+b.y1)/2) for b in blocks] if blocks else []
        snapped_cy = min(candidates_y, key=lambda y: abs(y - cy)) if candidates_y else cy
        if abs(snapped_cy - cy) > tol_pt: snapped_cy = cy
    snapped_cx += offset_x; snapped_cy += offset_y
    return fitz.Rect(snapped_cx - w/2, snapped_cy - h/2, snapped_cx + w/2, snapped_cy + h/2)

# ----------------- Auto-area detection -----------------

PLACEHOLDER_VARIANTS = [
    "(Seu Nome Aqui)", "Seu Nome Aqui", "( Seu Nome Aqui )",
    "(NOME)", "NOME", "Nome", "(Nome)"
]

def _find_placeholder_rect(page: fitz.Page) -> Optional[fitz.Rect]:
    for text in PLACEHOLDER_VARIANTS:
        try:
            rects = page.search_for(text, quads=False)
        except Exception:
            rects = []
        if rects:
            return max(rects, key=lambda r: r.get_area())
    return None

def _default_center_rect(page: fitz.Page) -> fitz.Rect:
    pw, ph = page.rect.width, page.rect.height
    w = pw * 0.60
    h = ph * 0.10
    cx, cy = pw / 2.0, ph * 0.55
    return fitz.Rect(cx - w/2, cy - h/2, cx + w/2, cy + h/2)

def compute_auto_area(doc: fitz.Document) -> Calibration:
    for i, page in enumerate(doc):
        r = _find_placeholder_rect(page)
        if r:
            pad_x = r.width * 0.4
            pad_y = r.height * 0.9
            rr = fitz.Rect(r.x0 - pad_x, r.y0 - pad_y, r.x1 + pad_x, r.y1 + pad_y)
            rr.intersect(page.rect)
            return Calibration(page_index=i, rect=(rr.x0, rr.y0, rr.x1, rr.y1))
    page = doc[0]
    rr = _default_center_rect(page)
    return Calibration(page_index=0, rect=(rr.x0, rr.y0, rr.x1, rr.y1))

# ----------------- Preview helpers -----------------

def _expand_rect(rect: fitz.Rect, page_rect: fitz.Rect, margin: float = 0.35) -> fitz.Rect:
    dx = rect.width * margin
    dy = rect.height * margin
    r = fitz.Rect(rect.x0 - dx, rect.y0 - dy, rect.x1 + dx, rect.y1 + dy)
    r.x0 = max(page_rect.x0, r.x0); r.y0 = max(page_rect.y0, r.y0)
    r.x1 = min(page_rect.x1, r.x1); r.y1 = min(page_rect.y1, r.y1)
    return r

# ----------------- Calibration UI -----------------

class CalibrateWindow(tk.Toplevel):
    def __init__(self, master, doc_path: str, page_index: int):
        super().__init__(master)
        self.title("Selecionar área do nome (opcional)")
        self.doc_path = doc_path
        self.page_index = page_index
        self.rect_pdf = None

        self.configure(bg=C_BG)
        self.geometry("920x620")
        self.resizable(False, False)

        self.canvas = tk.Canvas(self, bg=C_BG, highlightthickness=1, highlightbackground=C_WARM)
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
        self._load_image()

        self.start_x = self.start_y = None
        self.selection = None
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)

        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=10, pady=8)
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Confirmar", command=self._confirm).pack(side=tk.RIGHT, padx=6)

    def _load_image(self):
        doc = fitz.open(self.doc_path); page = doc[self.page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
        self.img_w, self.img_h = pix.width, pix.height
        self.page_w, self.page_h = page.rect.width, page.rect.height
        self.scale_x = self.img_w / self.page_w; self.scale_y = self.img_h / self.page_h
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.photo = ImageTk.PhotoImage(img)
        self.canvas_img = self.canvas.create_image(0, 0, anchor="nw", image=self.photo)

    def _on_press(self, e):
        self.start_x, self.start_y = e.x, e.y
        if self.selection: self.canvas.delete(self.selection); self.selection = None

    def _on_drag(self, e):
        if self.start_x is None: return
        if self.selection: self.canvas.delete(self.selection)
        self.selection = self.canvas.create_rectangle(self.start_x, self.start_y, e.x, e.y,
                                                      outline=C_TEAL, width=2)

    def _on_release(self, e): pass

    def _confirm(self):
        if not self.selection:
            messagebox.showwarning("Atenção", "Selecione uma área arrastando o mouse (ou cancele para manter automático)."); return
        x0, y0, x1, y1 = self.canvas.coords(self.selection)
        pdf_x0 = x0 / self.scale_x; pdf_y0 = y0 / self.scale_y
        pdf_x1 = x1 / self.scale_x; pdf_y1 = y1 / self.scale_y
        x0, x1 = sorted([pdf_x0, pdf_x1]); y0, y1 = sorted([pdf_y0, pdf_y1])
        self.rect_pdf = (x0, y0, x1, y1); self.destroy()

# ----------------- App -----------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.resizable(False, False)
        self.configure(bg=C_BG)

        style = ttk.Style()
        try: style.theme_use('clam')
        except Exception: pass
        style.configure("App.TLabelframe", background=C_BG, foreground=C_DARK)
        style.configure("App.TLabelframe.Label", background=C_BG, foreground=C_DARK)
        style.configure("App.TLabel", background=C_BG, foreground=C_WARM)
        style.configure("App.TCheckbutton", background=C_BG, foreground=C_WARM)
        style.configure("App.Horizontal.TProgressbar", troughcolor=C_BG)

        # --- State ---
        self.template_pdf = tk.StringVar()
        self.csv_path = tk.StringVar()
        self.font_path = tk.StringVar(value="")
        self.output_dir = tk.StringVar(value="")
        self.evento = tk.StringVar()

        self.use_consistent_size = tk.BooleanVar(value=False)  # OFF
        self.snap_enabled        = tk.BooleanVar(value=True)   # ON
        self.snap_tol = tk.DoubleVar(value=12.0)
        self.offset_x = tk.DoubleVar(value=0.0)
        self.offset_y = tk.DoubleVar(value=0.0)

        self.preview_name = tk.StringVar(value="Seu Nome")
        self.show_full_preview = tk.BooleanVar(value=False)

        self.progress_val = tk.IntVar(value=0)

        default_events = [
            "Fisio Summit BR 2025",
            "Workshop de Captação",
            "Imersão Marketing & Vendas",
            "Treinamento Walkyria Fernandes",
            "Evento Exemplo",
        ]

        # --- Step 1 ---
        step1 = ttk.LabelFrame(self, text="Passo 1 — Arquivos e informações", padding=(10,8), style="App.TLabelframe")
        step1.pack(fill=tk.X, padx=10, pady=(10,8))

        ttk.Label(step1, text="Template PDF:", style="App.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(step1, textvariable=self.template_pdf, width=60).grid(row=0, column=1, sticky="we", padx=5)
        ttk.Button(step1, text="Escolher...", command=self.pick_template).grid(row=0, column=2)

        ttk.Label(step1, text="CSV de Nomes:", style="App.TLabel").grid(row=1, column=0, sticky="w")
        ttk.Entry(step1, textvariable=self.csv_path, width=60).grid(row=1, column=1, sticky="we", padx=5)
        ttk.Button(step1, text="Escolher...", command=self.pick_csv).grid(row=1, column=2)

        ttk.Label(step1, text="Fonte (.ttf obrigatório):", style="App.TLabel").grid(row=2, column=0, sticky="w")
        ttk.Entry(step1, textvariable=self.font_path, width=60).grid(row=2, column=1, sticky="we", padx=5)
        ttk.Button(step1, text="Escolher...", command=self.pick_font).grid(row=2, column=2)

        ttk.Label(step1, text="Evento:", style="App.TLabel").grid(row=3, column=0, sticky="w")
        self.combo_evento = ttk.Combobox(step1, textvariable=self.evento, values=default_events, width=57, state="readonly")
        self.combo_evento.grid(row=3, column=1, sticky="we", padx=5)
        if default_events:
            self.combo_evento.set(default_events[0])

        ttk.Label(step1, text="Pasta de saída (opcional):", style="App.TLabel").grid(row=4, column=0, sticky="w")
        ttk.Entry(step1, textvariable=self.output_dir, width=60).grid(row=4, column=1, sticky="we", padx=5)
        ttk.Button(step1, text="Escolher...", command=self.pick_output_dir).grid(row=4, column=2)
        ttk.Button(step1, text="Abrir", command=lambda: open_folder(self.output_dir.get().strip() or os.getcwd())).grid(row=4, column=3, padx=(5,0))

        # --- Step 2 ---
        step2 = ttk.LabelFrame(self, text="Passo 2 (OPCIONAL) — Selecionar área do nome", padding=(10,8), style="App.TLabelframe")
        step2.pack(fill=tk.X, padx=10, pady=(0,8))
        ttk.Label(step2, text="Se você pular este passo, o app tentará encontrar a área automaticamente.", style="App.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0,6))
        ttk.Button(step2, text="Selecionar Área do Nome…", command=self.on_calibrate).grid(row=1, column=0, sticky="w")
        self.area_status = tk.StringVar(value="Área atual: Automática (sem seleção)")
        ttk.Label(step2, textvariable=self.area_status, style="App.TLabel").grid(row=1, column=1, sticky="w", padx=10)

        # --- Step 3 ---
        step3 = ttk.LabelFrame(self, text="Passo 3 — Pré-visualização e geração", padding=(10,8), style="App.TLabelframe")
        step3.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,8))

        toggles = ttk.Frame(step3); toggles.pack(fill=tk.X)
        ttk.Checkbutton(toggles, text="Snap automático aos textos próximos (recomendado)", variable=self.snap_enabled, command=self.refresh_preview, style="App.TCheckbutton").grid(row=0, column=0, sticky="w")
        ttk.Label(toggles, text="Tolerância (pt):", style="App.TLabel").grid(row=0, column=1, padx=(12,4), sticky="e")
        ttk.Spinbox(toggles, from_=0, to=72, increment=1, width=6, textvariable=self.snap_tol, command=self.refresh_preview).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(toggles, text="Usar tamanho único para todos (consistência)", variable=self.use_consistent_size, command=self.refresh_preview, style="App.TCheckbutton").grid(row=0, column=3, sticky="w", padx=(16,0))

        nudges = ttk.Frame(step3); nudges.pack(fill=tk.X, pady=(6,2))
        ttk.Label(nudges, text="Ajuste Horizontal (pt):", style="App.TLabel").grid(row=0, column=0, padx=(0,4), sticky="e")
        ttk.Spinbox(nudges, from_=-200, to=200, increment=0.5, width=8, textvariable=self.offset_x, command=self.refresh_preview).grid(row=0, column=1, sticky="w")
        ttk.Label(nudges, text="Ajuste Vertical (pt):", style="App.TLabel").grid(row=0, column=2, padx=(16,4), sticky="e")
        ttk.Spinbox(nudges, from_=-200, to=200, increment=0.5, width=8, textvariable=self.offset_y, command=self.refresh_preview).grid(row=0, column=3, sticky="w")

        pv_ctrl = ttk.Frame(step3); pv_ctrl.pack(fill=tk.X, pady=(6,4))
        ttk.Label(pv_ctrl, text="Nome da prévia:", style="App.TLabel").grid(row=0, column=0, sticky="w")
        e = ttk.Entry(pv_ctrl, textvariable=self.preview_name, width=36); e.grid(row=0, column=1, sticky="w", padx=6)
        e.bind("<KeyRelease>", lambda *_: self.refresh_preview())
        ttk.Checkbutton(pv_ctrl, text="Mostrar página inteira", variable=self.show_full_preview, command=self.refresh_preview, style="App.TCheckbutton").grid(row=0, column=2, padx=(16,0), sticky="w")

        # Preview canvas (height will SHRINK to image height)
        self.preview_canvas = tk.Canvas(step3, bg=C_BG, height=PREVIEW_TARGET_HEIGHT, width=PREVIEW_TARGET_WIDTH,
                                        highlightthickness=1, highlightbackground=C_WARM)
        self.preview_canvas.pack(fill=tk.NONE, expand=False, pady=(4,8))

        # Generate button + progress
        gen_row = ttk.Frame(step3); gen_row.pack(fill=tk.X, pady=(0,6))
        self.btn_generate = tk.Button(gen_row, text="GERAR CERTIFICADOS", command=self.generate,
                                      bg=C_PRIMARY, fg="white", activebackground=C_TEAL, activeforeground="white",
                                      relief="flat", padx=12, pady=10)
        self.btn_generate.pack(pady=2)

        actions = ttk.Frame(step3); actions.pack(fill=tk.X, pady=(0,0))
        self.pb = ttk.Progressbar(actions, mode="determinate", maximum=100,
                                  variable=self.progress_val, style="App.Horizontal.TProgressbar")
        self.pb.pack(fill=tk.X, expand=True, side=tk.LEFT)

        # Log
        self.log = tk.Text(self, height=8, wrap="word", bg=C_BG, fg=C_WARM,
                           highlightthickness=1, highlightbackground=C_WARM)
        self.log.pack(fill=tk.BOTH, expand=False, padx=10, pady=(6, 10))
        self.log_insert("1) Escolha Template, CSV, fonte (.ttf) e o Evento. 2) (Opcional) selecione a área. 3) Ajuste a prévia e clique em GERAR.\n")

        self.columnconfigure(0, weight=1)
        self.refresh_area_status()
        self.refresh_preview()

    # ---------- UI helpers ----------
    def log_insert(self, text: str):
        self.log.insert(tk.END, text); self.log.see(tk.END); self.update_idletasks()

    def refresh_area_status(self):
        tpl = self.template_pdf.get().strip()
        calib = load_calibration_for(tpl) if tpl and os.path.isfile(tpl) else None
        self.area_status.set("Área atual: Selecionada manualmente" if calib else "Área atual: Automática (sem seleção)")

    def pick_template(self):
        p = filedialog.askopenfilename(title="Escolha o Template PDF", filetypes=[("PDF","*.pdf")])
        if p:
            self.template_pdf.set(p)
            self.refresh_area_status()
            self.refresh_preview()

    def pick_csv(self):
        p = filedialog.askopenfilename(title="Escolha o CSV de Nomes", filetypes=[("CSV","*.csv")])
        if p:
            self.csv_path.set(p)
            try:
                names = read_names_from_csv(p)
                if names and (not self.preview_name.get() or self.preview_name.get()=="Seu Nome"):
                    self.preview_name.set(names[0])
            except Exception:
                pass
            self.refresh_preview()

    def pick_font(self):
        p = filedialog.askopenfilename(title="Escolha a fonte (.ttf)",
                                       filetypes=[("Fontes TrueType","*.ttf"),("Todos","*.*")])
        if p:
            self.font_path.set(p)
            self.refresh_preview()

    def pick_output_dir(self):
        d = filedialog.askdirectory(title="Escolha a pasta de saída")
        if d: self.output_dir.set(d)

    # ---------- Actions ----------
    def on_calibrate(self):
        tpl = self.template_pdf.get().strip()
        if not tpl:
            messagebox.showwarning("Atenção","Escolha primeiro o Template PDF."); return
        try:
            doc = fitz.open(tpl)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o PDF.\n\n{e}"); return

        page_index = 0
        if len(doc) > 1:
            val = simpledialog.askinteger("Página", f"Selecione a página (1..{len(doc)}):", minvalue=1, maxvalue=len(doc))
            if val is None:
                self.log_insert("Seleção de página cancelada.\n")
                return
            page_index = val - 1

        win = CalibrateWindow(self, tpl, page_index)
        self.wait_window(win)
        if win.rect_pdf:
            calib = Calibration(page_index=page_index, rect=win.rect_pdf)
            save_calibration_for(tpl, calib)
            self.log_insert(f"✓ Área salva (pág {page_index+1}, retângulo {tuple(round(v,2) for v in win.rect_pdf)}).\n")
            self.refresh_area_status()
            self.refresh_preview()
        else:
            self.log_insert("Seleção cancelada (continua automática).\n")
            self.refresh_area_status()

    # ---------- Rendering helpers ----------
    def _render_preview_image(self, doc: fitz.Document, page_idx: int, clip: Optional[fitz.Rect],
                              target_w: int, target_h: int) -> ImageTk.PhotoImage:
        page = doc[page_idx]
        if clip is None:
            s = min(target_w / page.rect.width, target_h / page.rect.height)
        else:
            s = min(target_w / clip.width, target_h / clip.height)
        s = max(0.01, s)
        pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False, clip=clip)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return ImageTk.PhotoImage(img)

    def _render_fullpage_strip(self, page: fitz.Page, tiles_preference=(4,3), gap: int = 10, padding: int = 12) -> ImageTk.PhotoImage:
        """Return a small strip image (width <= PREVIEW_TARGET_WIDTH, height minimal)."""
        page_w, page_h = page.rect.width, page.rect.height

        for n in (*tiles_preference, 2):
            avail_w = PREVIEW_TARGET_WIDTH - 2*padding - gap*(n-1)
            avail_h = PREVIEW_TARGET_HEIGHT - 2*padding
            if avail_w <= 0 or avail_h <= 0:
                continue

            s_w = avail_w / (page_w * n)
            s_h = avail_h / page_h
            s = max(0.01, min(s_w, s_h))
            if s < 0.05 and n > 2:
                continue

            pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
            tile = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            strip_w = n*tile.width + gap*(n-1)
            strip_h = tile.height
            # no vertical padding: canvas == strip size
            canvas = Image.new("RGB", (strip_w, strip_h), (237,237,237))
            x = 0
            for _ in range(n):
                canvas.paste(tile, (x, 0))
                x += tile.width + gap
            return ImageTk.PhotoImage(canvas)

        # Fallback: single centered thumbnail with minimal height
        s = max(0.01, min((PREVIEW_TARGET_WIDTH-2*padding)/page_w, (PREVIEW_TARGET_HEIGHT-2*padding)/page_h))
        pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return ImageTk.PhotoImage(img)

    def _get_active_area(self) -> Optional[Calibration]:
        tpl = self.template_pdf.get().strip()
        if not (tpl and os.path.isfile(tpl)): return None
        calib = load_calibration_for(tpl)
        if calib: return calib
        try:
            doc = fitz.open(tpl)
            auto = compute_auto_area(doc)
            doc.close()
            return auto
        except Exception:
            return None

    # ---------- Preview ----------
    def refresh_preview(self, *_):
        tpl = self.template_pdf.get().strip()
        if not (tpl and os.path.isfile(tpl)): return
        active_area = self._get_active_area()
        if not active_area: return

        show_full = self.show_full_preview.get()
        font_path = self.font_path.get().strip()

        try:
            if show_full:
                base_doc = fitz.open(tpl)
                if font_path and os.path.isfile(font_path) and font_path.lower().endswith(".ttf"):
                    tmp = fitz.open(tpl)
                    p = tmp[active_area.page_index]
                    rect_raw = fitz.Rect(*active_area.rect)
                    rect = compute_snapped_rect(
                        p, rect_raw, self.snap_enabled.get(),
                        float(self.snap_tol.get()), float(self.offset_x.get()), float(self.offset_y.get())
                    )
                    preview_text = (self.preview_name.get().strip() or "Seu Nome")
                    fontsize = autosize_font_to_rect(preview_text, rect, font_path)
                    draw_name_centered_with_size(p, rect, preview_text, font_path, fontsize)
                    photo = self._render_fullpage_strip(p, tiles_preference=(4,3), gap=10, padding=12)
                    tmp.close()
                else:
                    page = base_doc[active_area.page_index]
                    photo = self._render_fullpage_strip(page, tiles_preference=(4,3), gap=10, padding=12)

                self._preview_photo = photo
                self.preview_canvas.delete("all")
                # shrink canvas to the exact image height; keep full width
                self.preview_canvas.config(width=PREVIEW_TARGET_WIDTH,
                                           height=min(photo.height(), PREVIEW_TARGET_HEIGHT))
                x = (PREVIEW_TARGET_WIDTH - photo.width()) // 2
                self.preview_canvas.create_image(max(0, x), 0, anchor="nw", image=self._preview_photo)
                base_doc.close()
                return

            # Focused preview
            doc = fitz.open(tpl)
            page = doc[active_area.page_index]
            rect_raw = fitz.Rect(*active_area.rect)
            rect = compute_snapped_rect(
                page, rect_raw, self.snap_enabled.get(),
                float(self.snap_tol.get()), float(self.offset_x.get()), float(self.offset_y.get())
            )

            if not (font_path and os.path.isfile(font_path) and font_path.lower().endswith(".ttf")):
                clip = _expand_rect(rect, page.rect, margin=0.35)
                photo = self._render_preview_image(doc, active_area.page_index, clip,
                                                   PREVIEW_TARGET_WIDTH, PREVIEW_TARGET_HEIGHT)
                self._preview_photo = photo
                self.preview_canvas.delete("all")
                self.preview_canvas.config(width=PREVIEW_TARGET_WIDTH, height=photo.height())
                x = (PREVIEW_TARGET_WIDTH - photo.width()) // 2
                self.preview_canvas.create_image(max(0, x), 0, anchor="nw", image=self._preview_photo)
                doc.close()
                return

            preview_text = (self.preview_name.get().strip() or "Seu Nome")
            names = []
            if self.use_consistent_size.get() and os.path.isfile(self.csv_path.get()):
                try: names = read_names_from_csv(self.csv_path.get())
                except Exception: names = []
            if self.use_consistent_size.get() and names:
                fontsize = common_font_size_for_all(names, rect, font_path)
            else:
                fontsize = autosize_font_to_rect(preview_text, rect, font_path)

            tmp = fitz.open(tpl)
            p = tmp[active_area.page_index]
            draw_name_centered_with_size(p, rect, preview_text, font_path, fontsize)

            clip = _expand_rect(rect, p.rect, margin=0.35)
            photo = self._render_preview_image(tmp, active_area.page_index, clip,
                                               PREVIEW_TARGET_WIDTH, PREVIEW_TARGET_HEIGHT)
            self._preview_photo = photo
            self.preview_canvas.delete("all")
            self.preview_canvas.config(width=PREVIEW_TARGET_WIDTH, height=photo.height())
            x = (PREVIEW_TARGET_WIDTH - photo.width()) // 2
            self.preview_canvas.create_image(max(0, x), 0, anchor="nw", image=self._preview_photo)

            tmp.close(); doc.close()
        except Exception as e:
            self.log_insert(f"[Prévia] erro: {e}\n")

    # ---------- Generate ----------
    def generate(self):
        tpl = self.template_pdf.get().strip()
        csvp = self.csv_path.get().strip()
        font_path = self.font_path.get().strip()
        evento = self.evento.get().strip()

        if not tpl or not os.path.isfile(tpl):
            messagebox.showwarning("Atenção","Escolha o Template PDF."); return
        if not csvp or not os.path.isfile(csvp):
            messagebox.showwarning("Atenção","Escolha o CSV de Nomes."); return
        try:
            font_path = ensure_ttf_font(font_path)
        except Exception as e:
            messagebox.showwarning("Fonte", str(e)); return
        if not evento:
            messagebox.showwarning("Atenção","Informe o nome do Evento."); return

        try:
            names = read_names_from_csv(csvp)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler CSV: {e}"); return
        if not names:
            messagebox.showwarning("Atenção","Nenhum nome encontrado no CSV."); return

        active_area = self._get_active_area()
        if not active_area:
            messagebox.showwarning("Atenção","Não foi possível definir a área (automática/selecionada)."); return

        outdir = self.output_dir.get().strip()
        if not outdir:
            base = os.path.dirname(tpl) or os.getcwd()
            outdir = os.path.join(base, f"Certificados - {sanitize_filename(evento)} - {timestamp()}")
        os.makedirs(outdir, exist_ok=True)

        try:
            doc_tmp = fitz.open(tpl)
            page_tmp = doc_tmp[active_area.page_index]
            base_rect = compute_snapped_rect(
                page_tmp, fitz.Rect(*active_area.rect),
                self.snap_enabled.get(), float(self.snap_tol.get()),
                float(self.offset_x.get()), float(self.offset_y.get())
            )
            fontsize_common = common_font_size_for_all(names, base_rect, font_path) if self.use_consistent_size.get() else None
            doc_tmp.close()
        except Exception as e:
            self.log_insert(f"Falha ao preparar lote: {e}\n"); return

        total = len(names)
        self.pb["maximum"] = total; self.pb["value"] = 0
        self.progress_val.set(0); self.update_idletasks()

        self.log_insert(f"\n=== Gerando {total} certificados ===\n")
        ok = fail = 0

        for idx, name in enumerate(names, start=1):
            try:
                doc = fitz.open(tpl)
                page = doc[active_area.page_index]
                rect = base_rect
                fontsize = fontsize_common if fontsize_common is not None else autosize_font_to_rect(name, rect, font_path)
                draw_name_centered_with_size(page, rect, name, font_path, fontsize)

                outname = f"Certificado - {sanitize_filename(evento)} - {sanitize_filename(name)}.pdf"
                doc.save(os.path.join(outdir, outname)); doc.close()
                ok += 1; self.log_insert(f"✓ [{idx}] {name} → {outname}\n")
            except Exception as e:
                fail += 1; self.log_insert(f"✗ [{idx}] {name} → ERRO: {e}\n")

            self.pb["value"] = idx
            self.progress_val.set(idx)
            self.update_idletasks()

        self.log_insert(f"Concluído: {ok} ok, {fail} falhas. Saída: {outdir}\n")
        messagebox.showinfo("Pronto", f"Concluído: {ok} ok, {fail} falhas.\nPasta: {outdir}")
        open_folder(outdir)
        self.refresh_preview()

if __name__ == "__main__":
    App().mainloop()
