#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador de Certificados — Wizard

Requisitos:
  pip install pymupdf pillow
"""

import os, sys, csv, json, datetime, re, subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import font as tkfont
from dataclasses import dataclass
from typing import List, Optional, Tuple
from functools import lru_cache

import fitz  # PyMuPDF
from PIL import Image, ImageTk, ImageFont

APP_TITLE = "Gerador de Certificados — Wizard"
CALIBRATION_SUFFIX = ".calibration.json"

# Cores
C_BG      = "#eef8f4"
C_PRIMARY = "#17c88b"
C_TEAL    = "#009475"
C_DARK    = "#002927"
C_WARM    = "#514a43"
C_BORDER  = "#d6e7e2"
C_SURFACE = "#ffffff"
C_MUTED   = "#6b6b6b"

# Layout
SIDEBAR_W      = 540
PREVIEW_MIN_W  = 520
PREVIEW_MIN_H  = 300
BAND_HEIGHT    = 60
BANNER1_MAX_H  = 44
BANNER_SIDE_PAD= 16

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
def asset(*parts: str) -> str: return os.path.join(ASSETS_DIR, *parts)

def sanitize_filename(name: str) -> str:
    return re.sub(r"[\\/:*?\"<>|]", "_", name).strip().strip(".")

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
    rect: Tuple[float, float, float, float]  # x0, y0, x1, y1

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
    with open(template_pdf + CALIBRATION_SUFFIX, "w", encoding="utf-8") as f:
        json.dump({"page_index": calib.page_index, "rect": list(calib.rect)}, f, ensure_ascii=False, indent=2)

def ensure_ttf_font(font_path: str) -> str:
    if not font_path or not os.path.isfile(font_path):
        raise FileNotFoundError("Escolha um arquivo de fonte .ttf.")
    if not font_path.lower().endswith(".ttf"):
        raise ValueError("A fonte deve ser um arquivo .ttf (TrueType).")
    return font_path

def open_folder(path: str):
    try:
        if os.name == "nt": os.startfile(path)
        elif sys.platform == "darwin": subprocess.Popen(["open", path])
        else: subprocess.Popen(["xdg-open", path])
    except Exception:
        pass

@lru_cache(maxsize=512)
def _pil_font(font_path: str, size_px: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(font_path, size_px)

def _advance_width(text: str, font_path: str, size_pt: float) -> float:
    size_px = max(1, int(round(size_pt)))
    f = _pil_font(font_path, size_px)
    if hasattr(f, "getlength"): return float(f.getlength(text))
    l, t, r, b = f.getbbox(text)
    return float(r - l)

def _advance_width_pdf(page: fitz.Page, text: str, ttf_path: str, size_pt: float) -> float:
    try:
        return float(page.get_text_length(text, fontsize=size_pt, fontfile=ttf_path))
    except Exception:
        return _advance_width(text, ttf_path, size_pt)

def _font_metrics_em(ttf_path: str) -> Tuple[float, float]:
    try:
        f = fitz.Font(file=ttf_path)
        return float(f.ascender), float(f.descender)
    except Exception:
        pass
    try:
        size_px = 100
        f = _pil_font(ttf_path, size_px)
        a, d = f.getmetrics()
        total = max(1.0, float(a + d))
        return float(a) / total, -float(d) / total
    except Exception:
        return 0.9, -0.2

def _fits_in_rect(text: str, rect: fitz.Rect, fontfile: str, size_pt: float) -> bool:
    pad_w = rect.width * 0.02
    pad_h = rect.height * 0.06
    target_w = rect.width - 2 * pad_w
    target_h = rect.height - 2 * pad_h
    w = _advance_width(text, fontfile, size_pt)
    asc_em, desc_em = _font_metrics_em(fontfile)
    line_h_pt = (asc_em - desc_em) * size_pt
    return (w <= target_w) and (line_h_pt <= target_h)

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

def draw_name_centered_with_size(page: fitz.Page, rect: fitz.Rect, name: str,
                                 ttf_path: str, fontsize: float):
    w = _advance_width_pdf(page, name, ttf_path, fontsize)
    cx = rect.x0 + rect.width / 2.0
    x = cx - w / 2.0
    asc_em, desc_em = _font_metrics_em(ttf_path)
    cy = rect.y0 + rect.height / 2.0
    y_baseline = cy + (asc_em + desc_em) * fontsize / 2.0
    page.insert_text((x, y_baseline), name, fontsize=fontsize, fontfile=ttf_path,
                     fontname="UserFont", color=(0, 0, 0))

def _overlap_x(a: fitz.Rect, b: fitz.Rect) -> float:
    return max(0.0, min(a.x1, b.x1) - max(a.x0, b.x0))
def _overlap_y(a: fitz.Rect, b: fitz.Rect) -> float:
    return max(0.0, min(a.y1, b.y1) - max(a.y0, b.y0))

def _get_text_blocks(page: fitz.Page):
    try:
        blocks = page.get_text("blocks") or []
    except Exception:
        blocks = []
    rects = []
    for b in blocks:
        if len(b) >= 5 and isinstance(b[4], str) and b[4].strip():
            x0, y0, x1, y1 = b[:4]
            rects.append(fitz.Rect(x0, y0, x1, y1))
    return rects

def compute_snapped_rect(page: fitz.Page, rect: fitz.Rect, snap_enabled: bool,
                         tol_pt: float, offset_x: float, offset_y: float) -> fitz.Rect:
    w, h = rect.width, rect.height
    cx, cy = rect.x0 + w/2, rect.y0 + h/2
    if not snap_enabled:
        return fitz.Rect(cx - w/2 + offset_x, cy - h/2 + offset_y,
                         cx + w/2 + offset_x, cy + h/2 + offset_y)
    candidates = _get_text_blocks(page)
    vy = [b for b in candidates if _overlap_y(b, rect) > 0]
    vx = [b for b in candidates if _overlap_x(b, rect) > 0]
    snapped_cx, snapped_cy = cx, cy
    cx_candidates = [page.rect.width/2] + [ (b.x0 + b.x1)/2 for b in vy ]
    best_cx = min(cx_candidates, key=lambda x: abs(x - cx))
    if abs(best_cx - cx) <= tol_pt: snapped_cx = best_cx
    cy_candidates = [ (b.y0 + b.y1)/2 for b in vx ] or [page.rect.height/2]
    best_cy = min(cy_candidates, key=lambda y: abs(y - cy))
    if abs(best_cy - cy) <= tol_pt: snapped_cy = best_cy
    snapped_cx += offset_x; snapped_cy += offset_y
    return fitz.Rect(snapped_cx - w/2, snapped_cy - h/2, snapped_cx + w/2, snapped_cy + h/2)

PLACEHOLDER_VARIANTS = ["(Seu Nome Aqui)","Seu Nome Aqui","( Seu Nome Aqui )","(NOME)","NOME","Nome","(Nome)"]

def _find_placeholder_rect(page: fitz.Page) -> Optional[fitz.Rect]:
    for text in PLACEHOLDER_VARIANTS:
        try: rects = page.search_for(text, quads=False)
        except Exception: rects = []
        if rects: return max(rects, key=lambda r: r.get_area())
    return None

def _default_center_rect(page: fitz.Page) -> fitz.Rect:
    pw, ph = page.rect.width, page.rect.height
    w = pw * 0.60; h = ph * 0.10
    cx, cy = pw / 2.0, ph * 0.55
    return fitz.Rect(cx - w/2, cy - h/2, cx + w/2, cy + h/2)

def compute_auto_area(doc: fitz.Document) -> Calibration:
    for i, page in enumerate(doc):
        r = _find_placeholder_rect(page)
        if r:
            pad_x = r.width * 0.4; pad_y = r.height * 0.9
            rr = fitz.Rect(r.x0 - pad_x, r.y0 - pad_y, r.x1 + pad_x, r.y1 + pad_y)
            rr.intersect(page.rect)
            return Calibration(page_index=i, rect=(rr.x0, rr.y0, rr.x1, rr.y1))
    page = doc[0]
    rr = _default_center_rect(page)
    return Calibration(page_index=0, rect=(rr.x0, rr.y0, rr.x1, rr.y1))

def _expand_rect(rect: fitz.Rect, page_rect: fitz.Rect, margin: float = 0.35) -> fitz.Rect:
    dx = rect.width * margin; dy = rect.height * margin
    r = fitz.Rect(rect.x0 - dx, rect.y0 - dy, rect.x1 + dx, rect.y1 + dy)
    r.x0 = max(page_rect.x0, r.x0); r.y0 = max(page_rect.y0, r.y0)
    r.x1 = min(page_rect.x1, r.x1); r.y1 = min(page_rect.y1, r.y1)
    return r

class RoundedButton:
    def __init__(self, parent, text, command=None, fill=C_PRIMARY, fg="white", radius=14, padx=22, pady=12):
        self.parent = parent; self.text = text; self.command = command
        self.fill = fill; self.fg = fg; self.radius = radius; self.padx = padx; self.pady = pady
        self.disabled = False
        self.font = tkfont.Font(family="Segoe UI", size=11, weight="bold")
        tw = self.font.measure(self.text); th = self.font.metrics("linespace")
        self.w = tw + self.padx*2; self.h = th + self.pady*2
        self.canvas = tk.Canvas(parent, width=self.w, height=self.h,
                                background=parent.cget("bg") if "bg" in parent.keys() else C_SURFACE,
                                highlightthickness=0)
        self._current_fill = self.fill; self._draw()
        self.canvas.bind("<Enter>", self._on_enter); self.canvas.bind("<Leave>", self._on_leave)
        self.canvas.bind("<ButtonPress-1>", self._on_press); self.canvas.bind("<ButtonRelease-1>", self._on_release)
    def _rounded_rect(self, x, y, w, h, r, fill):  # minimal draw
        c = self.canvas
        c.create_rectangle(x+r, y, x+w-r, y+h, outline="", fill=fill)
        c.create_rectangle(x, y+r, x+w, y+h-r, outline="", fill=fill)
        c.create_arc(x, y, x+2*r, y+2*r, start=90, extent=90, outline="", fill=fill)
        c.create_arc(x+w-2*r, y, x+w, y+2*r, start=0, extent=90, outline="", fill=fill)
        c.create_arc(x, y+h-2*r, x+2*r, y+h, start=180, extent=90, outline="", fill=fill)
        c.create_arc(x+w-2*r, y+h-2*r, x+w, y+h, start=270, extent=90, outline="", fill=fill)
    def _draw(self):
        self.canvas.delete("all")
        self._rounded_rect(0, 0, self.w, self.h, self.radius, self._current_fill)
        self.canvas.create_text(self.w//2, self.h//2, text=self.text, fill=self.fg, font=self.font)
    def _on_enter(self,_): 
        if not self.disabled: self._current_fill = C_TEAL; self._draw()
    def _on_leave(self,_): 
        if not self.disabled: self._current_fill = self.fill; self._draw()
    def _on_press(self,_):
        if not self.disabled: self._current_fill = "#0a6f5c"; self._draw()
    def _on_release(self,_):
        if not self.disabled:
            self._current_fill = C_TEAL; self._draw()
            if callable(self.command): self.command()
    def set_disabled(self, flag: bool):
        self.disabled = flag; self._current_fill = "#bdbdbd" if flag else self.fill; self._draw()
    def pack(self,*a,**k): self.canvas.pack(*a,**k)

def make_card(parent, title: str):
    wrapper = tk.Frame(parent, bg=C_BG)
    head = tk.Frame(wrapper, bg=C_PRIMARY, height=28)
    head.pack(fill=tk.X, side=tk.TOP); head.pack_propagate(False)
    tk.Label(head, text=title, bg=C_PRIMARY, fg="white", font=("Segoe UI Semibold", 11), padx=12).pack(anchor="w", fill=tk.BOTH)
    body = tk.Frame(wrapper, bg=C_SURFACE, highlightthickness=1, highlightbackground=C_BORDER)
    body.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
    return wrapper, body

class CalibrateWindow(tk.Toplevel):
    def __init__(self, master, doc_path: str, page_index: int):
        super().__init__(master)
        self.title("Selecionar área do nome (opcional)")
        self.doc_path = doc_path; self.page_index = page_index; self.rect_pdf = None
        self.configure(bg=C_BG); self.geometry("1100x700"); self.minsize(820,560); self.resizable(True, True)

        top = tk.Frame(self, bg=C_BG); top.pack(fill=tk.X, padx=14, pady=(12,6))
        tk.Label(top, text="Zoom:", bg=C_BG, fg=C_WARM).pack(side=tk.LEFT)
        self.zoom_var = tk.DoubleVar(value=1.0); self.zoom = 1.0
        ttk.Scale(top, from_=0.3, to=3.0, orient="horizontal",
                  variable=self.zoom_var, command=lambda *_: self._apply_zoom()).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        ttk.Button(top, text="–", width=3, command=lambda: self._bump_zoom(-0.1)).pack(side=tk.LEFT, padx=(4,0))
        ttk.Button(top, text="+", width=3, command=lambda: self._bump_zoom(+0.1)).pack(side=tk.LEFT, padx=(4,12))
        ttk.Button(top, text="Largura", command=self._fit_width).pack(side=tk.LEFT, padx=(0,6))
        ttk.Button(top, text="Página", command=self._fit_page).pack(side=tk.LEFT, padx=(0,6))
        ttk.Button(top, text="Reset", command=self._reset_view).pack(side=tk.LEFT)

        tk.Label(self, text="Esquerdo=selecionar • Direito/Meio=arrastar • Roda=zoom",
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=16, pady=(0,4))

        area = tk.Frame(self, bg=C_BG); area.pack(fill=tk.BOTH, expand=True, padx=14, pady=(0,10))
        self.canvas = tk.Canvas(area, background=C_SURFACE, highlightthickness=1, highlightbackground=C_BORDER, cursor="tcross")
        self.hbar = ttk.Scrollbar(area, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.vbar = ttk.Scrollbar(area, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar.grid(row=1, column=0, sticky="ew")
        area.rowconfigure(0, weight=1); area.columnconfigure(0, weight=1)

        doc = fitz.open(self.doc_path); page = doc[self.page_index]
        pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
        self.img_base = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        self.img_w0, self.img_h0 = self.img_base.width, self.img_base.height
        self.page_w, self.page_h = page.rect.width, page.rect.height
        self.scale_x = self.img_w0 / self.page_w; self.scale_y = self.img_h0 / self.page_h
        doc.close()

        self.photo = ImageTk.PhotoImage(self.img_base)
        self.img_item = self.canvas.create_image(0, 0, anchor="nw", image=self.photo)
        self.canvas.config(scrollregion=(0,0,self.img_w0,self.img_h0))

        self.sel_rect = None; self.start_x = self.start_y = None
        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<ButtonPress-2>", lambda e: self.canvas.scan_mark(e.x, e.y))
        self.canvas.bind("<B2-Motion>",   lambda e: self.canvas.scan_dragto(e.x, e.y, gain=1))
        self.canvas.bind("<ButtonPress-3>", lambda e: self.canvas.scan_mark(e.x, e.y))
        self.canvas.bind("<B3-Motion>",     lambda e: self.canvas.scan_dragto(e.x, e.y, gain=1))
        self.canvas.bind("<MouseWheel>", lambda e: self._bump_zoom(+0.1 if e.delta > 0 else -0.1))
        self.canvas.bind("<Button-4>", lambda e: self._bump_zoom(+0.1))
        self.canvas.bind("<Button-5>", lambda e: self._bump_zoom(-0.1))

        btns = tk.Frame(self, bg=C_BG); btns.pack(fill=tk.X, padx=14, pady=(6,12))
        ttk.Button(btns, text="Cancelar", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Confirmar", command=self._confirm).pack(side=tk.RIGHT, padx=6)

        self.after(120, self._fit_width)

    def _apply_zoom(self):
        z = max(0.3, min(3.0, float(self.zoom_var.get())))
        if abs(z - self.zoom) < 1e-3: return
        self.zoom = z
        w = int(self.img_w0 * self.zoom); h = int(self.img_h0 * self.zoom)
        img = self.img_base.resize((w,h), Image.LANCZOS)
        self.photo = ImageTk.PhotoImage(img)
        self.canvas.itemconfig(self.img_item, image=self.photo)
        self.canvas.config(scrollregion=(0,0,w,h))

    def _bump_zoom(self, dv: float):
        self.zoom_var.set(max(0.3, min(3.0, float(self.zoom_var.get()) + dv))); self._apply_zoom()

    def _fit_width(self):
        self.update_idletasks()
        vis_w = max(200, self.canvas.winfo_width())
        self.zoom_var.set(vis_w / self.img_w0); self._apply_zoom()

    def _fit_page(self):
        self.update_idletasks()
        vis_w = max(200, self.canvas.winfo_width())
        vis_h = max(200, self.canvas.winfo_height())
        z = min(vis_w / self.img_w0, vis_h / self.img_h0)
        self.zoom_var.set(z); self._apply_zoom()

    def _reset_view(self):
        self.zoom_var.set(1.0); self._apply_zoom()
        self.canvas.xview_moveto(0); self.canvas.yview_moveto(0)
        if self.sel_rect: self.canvas.delete(self.sel_rect); self.sel_rect = None

    def _on_press(self, e):
        x = self.canvas.canvasx(e.x); y = self.canvas.canvasy(e.y)
        self.start_x, self.start_y = x, y
        if self.sel_rect: self.canvas.delete(self.sel_rect); self.sel_rect = None
        self.sel_rect = self.canvas.create_rectangle(x, y, x, y, outline=C_TEAL, width=2)

    def _on_drag(self, e):
        if self.start_x is None: return
        x = self.canvas.canvasx(e.x); y = self.canvas.canvasy(e.y)
        self.canvas.coords(self.sel_rect, self.start_x, self.start_y, x, y)

    def _on_release(self, e): pass

    def _confirm(self):
        if not self.sel_rect:
            messagebox.showwarning("Atenção", "Selecione uma área arrastando o mouse."); return
        x0, y0, x1, y1 = self.canvas.coords(self.sel_rect)
        w = self.img_w0 * self.zoom; h = self.img_h0 * self.zoom
        x0 = min(max(0, x0), w); x1 = min(max(0, x1), w)
        y0 = min(max(0, y0), h); y1 = min(max(0, y1), h)
        if abs(x1-x0) < 3 or abs(y1-y0) < 3:
            messagebox.showwarning("Atenção", "Área muito pequena."); return
        ix0, iy0 = x0 / self.zoom, y0 / self.zoom
        ix1, iy1 = x1 / self.zoom, y1 / self.zoom
        pdf_x0, pdf_y0 = ix0 / self.scale_x, iy0 / self.scale_y
        pdf_x1, pdf_y1 = ix1 / self.scale_x, iy1 / self.scale_y
        x0, x1 = sorted([pdf_x0, pdf_x1]); y0, y1 = sorted([pdf_y0, pdf_y1])
        self.rect_pdf = (x0, y0, x1, y1); self.destroy()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        try:
            import sv_ttk
            sv_ttk.set_theme("light")
        except Exception:
            pass
        self.configure(bg=C_BG)

        self._brand_after_band = None
        self._brand_after_left = None
        self._banner1_dims = (0, 0)
        self._banner2_dims = (0, 0)

        self.grid_rowconfigure(0, weight=0, minsize=BAND_HEIGHT)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.band = tk.Frame(self, bg=C_PRIMARY, height=BAND_HEIGHT)
        self.band.grid(row=0, column=0, sticky="nsew"); self.band.grid_propagate(False)
        self._band_label = tk.Label(self.band, bg=C_PRIMARY)
        self._band_label.pack(side="left", padx=BANNER_SIDE_PAD, pady=8)

        content = tk.Frame(self, bg=C_BG)
        content.grid(row=1, column=0, sticky="nsew")
        content.grid_rowconfigure(0, weight=1)
        content.grid_columnconfigure(0, weight=1)
        content.grid_columnconfigure(1, weight=0, minsize=SIDEBAR_W)

        self.left = tk.Frame(content, bg=C_BG)
        self.left.grid(row=0, column=0, sticky="nsew", padx=(20,10), pady=(14,12))
        self.right = tk.Frame(content, bg=C_BG)
        self.right.grid(row=0, column=1, sticky="nsew", padx=(10,20), pady=(14,12))

        style = ttk.Style()
        style.configure("TLabel", foreground=C_WARM)
        style.configure("TCheckbutton", foreground=C_WARM)
        style.configure("App.Horizontal.TProgressbar", troughcolor=C_SURFACE)

        self._cancel = False
        self.template_pdf = tk.StringVar(value="")
        self.csv_path     = tk.StringVar(value="")
        self.font_path    = tk.StringVar(value="")
        self.output_dir   = tk.StringVar(value="")
        self.evento       = tk.StringVar(value="Fisio Summit BR 2025")

        self.use_consistent_size = tk.BooleanVar(value=False)
        self.snap_enabled        = tk.BooleanVar(value=True)
        self.snap_tol            = tk.DoubleVar(value=16.0)
        self.offset_x = tk.DoubleVar(value=0.0)
        self.offset_y = tk.DoubleVar(value=0.0)
        self.preview_name = tk.StringVar(value="Seu Nome")
        self.show_full_preview = tk.BooleanVar(value=True)
        self.preview_zoom = tk.DoubleVar(value=1.25)
        self.merge_pdf = tk.BooleanVar(value=False)
        self.progress_val = tk.IntVar(value=0)
        self.status_var = tk.StringVar(value="")

        self._init_icons()
        self._load_brand_images()

        self._build_left()
        self._build_right()
        self.refresh_area_status()

        self._pv_after_id = None
        self.band.bind("<Configure>", self._on_band_configure)
        self.left.bind("<Configure>", self._on_left_configure)

        self.bind_all("<Control-o>", lambda e: self.pick_template())
        self.bind_all("<Control-Shift-o>", lambda e: self.pick_csv())
        self.bind_all("<Control-f>", lambda e: self.pick_font())
        self.bind_all("<Control-g>", lambda e: self.generate())
        self.bind_all("<F1>", lambda e: self._show_help())
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        try:
            self.iconbitmap(asset("logo.ico"))
        except Exception:
            pass

        try:
            self.state("zoomed")
        except Exception:
            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            self.geometry(f"{sw}x{sh}+0+0")

        self.after(120, self._apply_brand_images)
        self.after(350, self.refresh_preview)
        self.after(500, self._show_tutorial)

    def _init_icons(self):
        pass

    def _load_brand_images(self):
        self._banner1_img_raw = None; self._banner2_img_raw = None
        try:
            p1 = asset("banner1.png")
            if os.path.isfile(p1): self._banner1_img_raw = Image.open(p1).convert("RGBA")
        except Exception: self._banner1_img_raw = None
        try:
            p2 = asset("banner2.png")
            if os.path.isfile(p2): self._banner2_img_raw = Image.open(p2).convert("RGBA")
        except Exception: self._banner2_img_raw = None

    def _on_band_configure(self, _):
        if self._brand_after_band:
            try: self.after_cancel(self._brand_after_band)
            except Exception: pass
        self._brand_after_band = self.after(120, self._apply_brand_images)

    def _on_left_configure(self, _):
        if self._brand_after_left:
            try: self.after_cancel(self._brand_after_left)
            except Exception: pass
        self._brand_after_left = self.after(120, self._apply_brand_images)

    def _apply_brand_images(self):
        if self._banner1_img_raw is not None and hasattr(self, "_band_label"):
            bw = max(0, self.band.winfo_width()); bh = max(0, self.band.winfo_height())
            if bw > 4 and bh > 4:
                iw, ih = self._banner1_img_raw.size
                avail_w = max(1, bw - 2*BANNER_SIDE_PAD)
                avail_h = max(1, min(BANNER1_MAX_H, bh - 16))
                scale = min(avail_w / iw, avail_h / ih, 1.0)
                new_w = max(1, int(iw * scale)); new_h = max(1, int(ih * scale))
                if (new_w, new_h) != self._banner1_dims:
                    new = self._banner1_img_raw.resize((new_w, new_h), Image.LANCZOS)
                    self._banner1_photo = ImageTk.PhotoImage(new)
                    self._band_label.configure(image=self._banner1_photo)
                    self._band_label.image = self._banner1_photo
                    self._banner1_dims = (new_w, new_h)

        if self._banner2_img_raw is not None and hasattr(self, "_banner2_label") and self._banner2_label:
            lw = max(0, self.left.winfo_width())
            if lw > 4:
                iw, ih = self._banner2_img_raw.size
                avail = max(1, lw - 2*BANNER_SIDE_PAD)
                scale = min(avail / iw, 1.0)
                new_w = max(1, int(iw * scale)); new_h = max(1, int(ih * scale))
                if (new_w, new_h) != self._banner2_dims:
                    new = self._banner2_img_raw.resize((new_w, new_h), Image.LANCZOS)
                    self._banner2_photo = ImageTk.PhotoImage(new)
                    self._banner2_label.configure(image=self._banner2_photo)
                    self._banner2_label.image = self._banner2_photo
                    self._banner2_dims = (new_w, new_h)

    def _build_left(self):
        hdr = tk.Frame(self.left, bg=C_BG, pady=0); hdr.pack(fill=tk.X, padx=0, pady=(0,6))
        brand2 = tk.Frame(hdr, bg=C_SURFACE, highlightthickness=0); brand2.pack(fill=tk.X, padx=0, pady=0)
        self._banner2_label = tk.Label(brand2, bg=C_SURFACE); self._banner2_label.pack(anchor="w", padx=10, pady=6)

        wrap1, g = make_card(self.left, "Passo 1 — O que você vai usar")
        wrap1.pack(fill=tk.X, padx=0, pady=(12,12))
        for i in range(3): g.columnconfigure(i, weight=1 if i==1 else 0)

        ttk.Label(g, text="Modelo do certificado (PDF):").grid(row=0, column=0, sticky="w", padx=12, pady=(10,4))
        ttk.Entry(g, textvariable=self.template_pdf).grid(row=0, column=1, sticky="ew", padx=6, pady=(10,4))
        ttk.Button(g, text="Escolher...", command=self.pick_template).grid(row=0, column=2, padx=12, pady=(10,4), sticky="e")

        ttk.Label(g, text="Lista de nomes (arquivo .CSV do Excel/Sheets):").grid(row=1, column=0, sticky="w", padx=12, pady=4)
        row_csv = tk.Frame(g, bg=C_SURFACE); row_csv.grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        row_csv.grid_columnconfigure(0, weight=1)
        ttk.Entry(row_csv, textvariable=self.csv_path).grid(row=0, column=0, sticky="ew")
        ttk.Button(row_csv, text="Como salvar CSV?", command=self._explain_csv).grid(row=0, column=1, padx=(6,0))
        ttk.Button(g, text="Escolher...", command=self.pick_csv).grid(row=1, column=2, padx=12, pady=4, sticky="e")

        ttk.Label(g, text="Fonte do nome (arquivo .TTF):").grid(row=2, column=0, sticky="w", padx=12, pady=4)
        ttk.Entry(g, textvariable=self.font_path).grid(row=2, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(g, text="Escolher...", command=self.pick_font).grid(row=2, column=2, padx=12, pady=4, sticky="e")

        ttk.Label(g, text="Nome do evento (aparece no arquivo):").grid(row=3, column=0, sticky="w", padx=12, pady=4)
        self.combo_evento = ttk.Combobox(g, textvariable=self.evento, state="readonly",
                                         values=["Fisio Summit 2025","Liberação Funcional Avançada",
                                                 "Mobilização Neural","RCA360",
                                                 "Congresso RCA 2026"])
        self.combo_evento.grid(row=3, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(g, text="Pasta onde salvar (opcional):").grid(row=4, column=0, sticky="w", padx=12, pady=(4,12))
        ttk.Entry(g, textvariable=self.output_dir).grid(row=4, column=1, sticky="ew", padx=6, pady=(4,12))
        ttk.Button(g, text="Escolher...", command=self.pick_output_dir).grid(row=4, column=2, padx=12, pady=(4,12), sticky="e")

        wrap2, s2 = make_card(self.left, "Passo 2 — Onde o nome deve entrar?")
        wrap2.pack(fill=tk.X, padx=0, pady=(0,12))
        s2.columnconfigure(1, weight=1)
        ttk.Label(s2, text=("Clique em “Selecionar área…” e desenhe um retângulo em cima do local do nome. "
                            "Se preferir, pule e o app tenta adivinhar sozinho.")
                 ).grid(row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(10,6))
        ttk.Button(s2, text="Selecionar área…", command=self.on_calibrate).grid(row=1, column=0, sticky="w", padx=12, pady=(0,12))
        self.area_status = tk.StringVar(value="Área atual: Automática (sem seleção)")
        ttk.Label(s2, textvariable=self.area_status).grid(row=1, column=1, sticky="w", padx=10, pady=(0,12))

        wrap3, step3 = make_card(self.left, "Passo 3 — Ajustes da prévia")
        wrap3.pack(fill=tk.X, padx=0, pady=(0,0))
        toggles = tk.Frame(step3, bg=C_SURFACE); toggles.pack(fill=tk.X, padx=8, pady=(8,0))
        ttk.Checkbutton(
            toggles,
            text="Alinhar com textos do PDF (mantém o nome alinhado com títulos/linhas do modelo)",
            variable=self.snap_enabled, command=self.refresh_preview
        ).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(
            toggles,
            text="Usar o mesmo tamanho de letra para todos os nomes (consistência)",
            variable=self.use_consistent_size, command=self.refresh_preview
        ).grid(row=1, column=0, sticky="w", pady=(6,0))

        nudges = tk.Frame(step3, bg=C_SURFACE); nudges.pack(fill=tk.X, padx=8, pady=(10,8))
        ttk.Label(nudges, text="Ajuste fino horizontal (pt):").grid(row=0, column=0, padx=(0,4), sticky="e")
        ttk.Spinbox(nudges, from_=-200, to=200, increment=0.5, width=8, textvariable=self.offset_x, command=self.refresh_preview).grid(row=0, column=1, sticky="w")
        ttk.Label(nudges, text="Ajuste fino vertical (pt):").grid(row=0, column=2, padx=(16,4), sticky="e")
        ttk.Spinbox(nudges, from_=-200, to=200, increment=0.5, width=8, textvariable=self.offset_y, command=self.refresh_preview).grid(row=0, column=3, sticky="w")

        pv_ctrl = tk.Frame(step3, bg=C_SURFACE); pv_ctrl.pack(fill=tk.X, padx=8, pady=(0,10))
        pv_ctrl.grid_columnconfigure(4, weight=1)
        ttk.Label(pv_ctrl, text="Nome de teste:").grid(row=0, column=0, sticky="w")
        e = ttk.Entry(pv_ctrl, textvariable=self.preview_name); e.grid(row=0, column=1, sticky="w", padx=6)
        e.bind("<KeyRelease>", lambda *_: self.refresh_preview())
        ttk.Checkbutton(pv_ctrl, text="Mostrar página inteira", variable=self.show_full_preview, command=self.refresh_preview).grid(row=0, column=2, padx=(16,12), sticky="w")
        ttk.Label(pv_ctrl, text="Zoom da prévia:").grid(row=0, column=3, sticky="e")
        ttk.Scale(pv_ctrl, from_=0.75, to=1.75, orient="horizontal", variable=self.preview_zoom,
                  command=lambda *_: self.refresh_preview()).grid(row=0, column=4, sticky="we", padx=(8,0))
        ttk.Button(pv_ctrl, text="← Usar o maior nome do CSV", command=self.set_longest_from_csv).grid(row=0, column=5, padx=(12,0), sticky="e")

    def _build_right(self):
        prev_wrap, prev_body = make_card(self.right, "Pré-visualização")
        prev_wrap.pack(fill=tk.BOTH, expand=True, padx=0, pady=(0,12))
        self._pv_wrap = tk.Frame(prev_body, bg=C_SURFACE)
        self._pv_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.preview_canvas = tk.Canvas(self._pv_wrap, background=C_SURFACE,
                                        width=PREVIEW_MIN_W, height=PREVIEW_MIN_H,
                                        highlightthickness=1, highlightbackground=C_BORDER)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        self.preview_img_item = None; self._preview_photo = None
        self.preview_canvas.bind("<Map>", lambda e: self.after(80, self.refresh_preview))
        self.preview_canvas.bind("<Configure>", self._on_canvas_configure)

        actions_wrap, actions_body = make_card(self.right, "Ações")
        actions_wrap.pack(fill=tk.X, padx=0, pady=(0,0))
        pad = 12
        self.btn_generate = RoundedButton(actions_body, "GERAR CERTIFICADOS", command=self.generate)
        self.btn_generate.pack(padx=pad, pady=(pad,8), anchor="n")
        self.btn_cancel = RoundedButton(actions_body, "Cancelar", command=self.cancel_generation, fill="#e0e0e0", fg="#333333")
        self.btn_cancel.set_disabled(True)
        self.btn_cancel.pack(padx=pad, pady=(0,10), anchor="n")
        ttk.Checkbutton(actions_body, text="Juntar tudo em 1 PDF (opcional)", variable=self.merge_pdf).pack(anchor="w", padx=pad, pady=(0,10))
        ttk.Label(actions_body, text="Progresso:").pack(anchor="w", padx=pad, pady=(0,4))
        self.pb = ttk.Progressbar(actions_body, mode="determinate", maximum=100, variable=self.progress_val, style="App.Horizontal.TProgressbar")
        self.pb.pack(fill=tk.X, padx=pad, pady=(0,6))
        ttk.Label(actions_body, textvariable=self.status_var, foreground=C_MUTED).pack(anchor="w", padx=pad, pady=(0,10))
        ttk.Button(actions_body, text="Abrir pasta de saída", command=lambda: open_folder(self.output_dir.get().strip() or os.getcwd())).pack(fill=tk.X, padx=pad, pady=(0,pad))

    def _on_canvas_configure(self, _):
        if getattr(self, "_pv_after_id", None):
            try: self.after_cancel(self._pv_after_id)
            except Exception: pass
        self._pv_after_id = self.after(120, self.refresh_preview)

    def log_insert(self, text: str):
        self.status_var.set(text.strip()); self.update_idletasks()

    def refresh_area_status(self):
        tpl = self.template_pdf.get().strip()
        calib = load_calibration_for(tpl) if tpl and os.path.isfile(tpl) else None
        self.area_status.set("Área atual: Selecionada manualmente" if calib else "Área atual: Automática (sem seleção)")

    def pick_template(self):
        p = filedialog.askopenfilename(title="Escolha o PDF do modelo", filetypes=[("PDF","*.pdf")])
        if p:
            self.template_pdf.set(p); self.refresh_area_status(); self.refresh_preview()

    def pick_csv(self):
        p = filedialog.askopenfilename(title="Escolha a lista de nomes (.CSV)", filetypes=[("CSV","*.csv")])
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
        p = filedialog.askopenfilename(title="Escolha a fonte do nome (.TTF)",
                                       filetypes=[("Fontes TrueType","*.ttf"),("Todos","*.*")])
        if p:
            self.font_path.set(p); self.refresh_preview()

    def pick_output_dir(self):
        d = filedialog.askdirectory(title="Escolha a pasta onde salvar")
        if d: self.output_dir.set(d)

    def _explain_csv(self):
        messagebox.showinfo(
            "Como salvar um CSV",
            "Você pode usar Excel ou Google Sheets:\n\n"
            "• Excel: arquivo com os nomes (uma coluna, cada linha um nome) → Arquivo > Salvar como > CSV.\n"
            "• Google Sheets: Arquivo > Fazer download > Valores separados por vírgulas (.csv).\n\n"
            "Dica: deixe apenas a coluna com o nome. O app ignora o cabeçalho se for 'Nome'/'Name'."
        )

    def set_longest_from_csv(self):
        p = self.csv_path.get().strip()
        if not (p and os.path.isfile(p)):
            messagebox.showinfo("CSV", "Escolha primeiro a lista de nomes (.CSV)."); return
        try:
            names = read_names_from_csv(p)
            if not names: raise ValueError("CSV vazio.")
            longest = max(names, key=lambda s: len(s or ""))
            self.preview_name.set(longest)
            self.refresh_preview()
            self.log_insert(f"Prévia: maior nome do CSV → “{longest}”.")
        except Exception as e:
            messagebox.showerror("CSV", f"Falha ao ler CSV: {e}")

    def on_calibrate(self):
        tpl = self.template_pdf.get().strip()
        if not tpl:
            messagebox.showwarning("Atenção","Escolha primeiro o PDF do modelo."); return
        try:
            doc = fitz.open(tpl)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o PDF.\n\n{e}"); return
        page_index = 0
        if len(doc) > 1:
            val = simpledialog.askinteger("Página", f"Selecione a página (1..{len(doc)}):", minvalue=1, maxvalue=len(doc))
            if val is None:
                self.log_insert("Seleção cancelada (continua automática)."); doc.close(); return
            page_index = val - 1
        doc.close()
        win = CalibrateWindow(self, tpl, page_index)
        self.wait_window(win)
        if win.rect_pdf:
            calib = Calibration(page_index=page_index, rect=win.rect_pdf)
            save_calibration_for(tpl, calib)
            self.log_insert(f"Área salva (pág {page_index+1}).")
            self.refresh_area_status(); self.refresh_preview()
        else:
            self.log_insert("Seleção cancelada (continua automática).")
            self.refresh_area_status()

    def _place_preview_image(self, photo: ImageTk.PhotoImage):
        cw = max(1, self.preview_canvas.winfo_width()); ch = max(1, self.preview_canvas.winfo_height())
        iw, ih = photo.width(), photo.height()
        cx, cy = cw // 2, ch // 2
        x0, y0 = cx - iw // 2, cy - ih // 2
        x1, y1 = x0 + iw, y0 + ih
        self.preview_canvas.delete("all")
        self.preview_canvas.create_rectangle(x0-6, y0-6, x1+6, y1+6, fill="#e9f2ef", outline="")
        self.preview_canvas.create_rectangle(x0-1, y0-1, x1+1, y1+1, outline=C_BORDER, width=1)
        self.preview_img_item = self.preview_canvas.create_image(cx, cy, image=photo, anchor="center")
        self._preview_photo = photo
        self.preview_canvas.image = photo

    def _draw_placeholder(self, msg="Sem prévia"):
        cw = max(1, self.preview_canvas.winfo_width()); ch = max(1, self.preview_canvas.winfo_height())
        self.preview_canvas.delete("all")
        pad = 24
        self.preview_canvas.create_rectangle(0, 0, cw, ch, fill=C_SURFACE, outline=C_BORDER)
        self.preview_canvas.create_rectangle(pad, pad, cw-pad, ch-pad, outline=C_BORDER, dash=(4,3))
        self.preview_canvas.create_text(cw // 2, ch // 2, text=msg, fill=C_MUTED, font=("Segoe UI", 11))
        self.preview_img_item = None; self.preview_canvas.image = None

    def _render_preview_image(self, doc: fitz.Document, page_idx: int, clip: Optional[fitz.Rect],
                              box_w: int, box_h: int, zoom_scale: float) -> ImageTk.PhotoImage:
        page = doc[page_idx]
        src_w = (clip.width if clip else page.rect.width)
        src_h = (clip.height if clip else page.rect.height)
        box_w = max(PREVIEW_MIN_W, int(box_w)); box_h = max(PREVIEW_MIN_H, int(box_h))
        s = min(box_w / src_w, box_h / src_h)
        s = max(0.05, s) * max(0.5, min(2.5, float(zoom_scale)))
        max_pix = 2400
        if src_w * s > max_pix or src_h * s > max_pix:
            s = min(max_pix / src_w, max_pix / src_h)
        pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False, clip=clip)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return ImageTk.PhotoImage(img)

    def _render_fullpage_single(self, page: fitz.Page, box_w: int, box_h: int, zoom_scale: float) -> ImageTk.PhotoImage:
        src_w, src_h = page.rect.width, page.rect.height
        box_w = max(PREVIEW_MIN_W, int(box_w)); box_h = max(PREVIEW_MIN_H, int(box_h))
        s = min(box_w / src_w, box_h / src_h)
        s = max(0.05, s) * max(0.5, min(2.5, float(zoom_scale)))
        max_pix = 2400
        if src_w * s > max_pix or src_h * s > max_pix:
            s = min(max_pix / src_w, max_pix / src_h)
        pix = page.get_pixmap(matrix=fitz.Matrix(s, s), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return ImageTk.PhotoImage(img)

    def _get_active_area(self) -> Optional[Calibration]:
        tpl = self.template_pdf.get().strip()
        if not (tpl and os.path.isfile(tpl)): return None
        calib = load_calibration_for(tpl)
        if calib: return calib
        try:
            doc = fitz.open(tpl); auto = compute_auto_area(doc); doc.close(); return auto
        except Exception:
            return None

    def refresh_preview(self, *_):
        if not hasattr(self, "preview_canvas"): return
        self.update_idletasks()
        cw = self.preview_canvas.winfo_width(); ch = self.preview_canvas.winfo_height()
        if cw <= 10 or ch <= 10:
            self.after(120, self.refresh_preview); return

        tpl = self.template_pdf.get().strip()
        if not (tpl and os.path.isfile(tpl)):
            self._draw_placeholder("Carregue um PDF do modelo para ver a prévia"); return

        active_area = self._get_active_area()
        if not active_area:
            self._draw_placeholder("Defina a área do nome (automático/seleção manual)"); return

        box_w = max(PREVIEW_MIN_W, cw); box_h = max(PREVIEW_MIN_H, ch)
        z = max(0.5, min(2.5, float(self.preview_zoom.get())))
        font_path = self.font_path.get().strip()
        show_full = bool(self.show_full_preview.get())

        base_doc = tmp_doc = None
        try:
            base_doc = fitz.open(tpl)
            page_index = active_area.page_index
            page = base_doc[page_index]

            if show_full:
                photo = None
                if font_path and os.path.isfile(font_path) and font_path.lower().endswith(".ttf"):
                    try:
                        tmp_doc = fitz.open(tpl)
                        p = tmp_doc[page_index]
                        rect_raw = fitz.Rect(*active_area.rect)
                        rect = compute_snapped_rect(
                            p, rect_raw, self.snap_enabled.get(),
                            float(self.snap_tol.get()), float(self.offset_x.get()), float(self.offset_y.get())
                        )
                        preview_text = (self.preview_name.get().strip() or "Seu Nome")
                        names = []
                        if self.use_consistent_size.get() and os.path.isfile(self.csv_path.get()):
                            try:
                                # limitar a 400 para não travar em listas enormes na prévia
                                names_all = read_names_from_csv(self.csv_path.get())
                                names = names_all[:400] if len(names_all) > 400 else names_all
                            except Exception:
                                names = []
                        if self.use_consistent_size.get() and names:
                            fontsize = common_font_size_for_all(names, rect, font_path)
                        else:
                            fontsize = autosize_font_to_rect(preview_text, rect, font_path)
                        draw_name_centered_with_size(p, rect, preview_text, font_path, fontsize)
                        photo = self._render_fullpage_single(p, box_w, box_h, z)
                    except Exception:
                        photo = None
                    finally:
                        if tmp_doc is not None:
                            tmp_doc.close(); tmp_doc = None
                if photo is None:
                    photo = self._render_fullpage_single(page, box_w, box_h, z)
                self._place_preview_image(photo); return

            rect_raw = fitz.Rect(*active_area.rect)
            rect = compute_snapped_rect(
                page, rect_raw, self.snap_enabled.get(),
                float(self.snap_tol.get()), float(self.offset_x.get()), float(self.offset_y.get())
            )
            clip = _expand_rect(rect, page.rect, margin=0.35)
            preview_text = (self.preview_name.get().strip() or "Seu Nome")
            names = []
            if self.use_consistent_size.get() and os.path.isfile(self.csv_path.get()):
                try:
                    names_all = read_names_from_csv(self.csv_path.get())
                    names = names_all[:400] if len(names_all) > 400 else names_all
                except Exception:
                    names = []

            photo = None
            if font_path and os.path.isfile(font_path) and font_path.lower().endswith(".ttf"):
                try:
                    tmp_doc = fitz.open(tpl); p = tmp_doc[page_index]
                    fontsize = common_font_size_for_all(names, rect, font_path) if (self.use_consistent_size.get() and names) else autosize_font_to_rect(preview_text, rect, font_path)
                    draw_name_centered_with_size(p, rect, preview_text, font_path, fontsize)
                    photo = self._render_preview_image(tmp_doc, page_index, clip, box_w, box_h, z)
                except Exception:
                    photo = None
                finally:
                    if tmp_doc is not None:
                        tmp_doc.close(); tmp_doc = None

            if photo is None:
                photo = self._render_preview_image(base_doc, page_index, clip, box_w, box_h, z)
            self._place_preview_image(photo)

        except Exception as e:
            self._draw_placeholder(f"Não foi possível gerar a prévia: {e}")
        finally:
            if base_doc is not None: base_doc.close()

    def cancel_generation(self):
        self._cancel = True; self.log_insert("Cancelando… aguarde a etapa atual.")

    def generate(self):
        tpl = self.template_pdf.get().strip()
        csvp = self.csv_path.get().strip()
        font_path = self.font_path.get().strip()
        evento = self.evento.get().strip()

        if not tpl or not os.path.isfile(tpl):
            messagebox.showwarning("Atenção","Escolha o PDF do modelo."); return
        if not csvp or not os.path.isfile(csvp):
            messagebox.showwarning("Atenção","Escolha a lista de nomes (.CSV)."); return
        try:
            font_path = ensure_ttf_font(font_path)
        except Exception as e:
            messagebox.showwarning("Fonte", str(e)); return
        if not evento:
            messagebox.showwarning("Atenção","Informe o nome do evento."); return

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
            self.log_insert(f"Falha ao preparar lote: {e}"); return

        total = len(names)
        self.pb["maximum"] = total; self.pb["value"] = 0
        self.progress_val.set(0); self.update_idletasks()

        self._cancel = False
        self.btn_generate.set_disabled(True)
        self.btn_cancel.set_disabled(False if total > 0 else True)

        self.log_insert(f"Gerando {total} certificados…")
        ok = fail = 0
        merged_doc = fitz.Document() if self.merge_pdf.get() else None

        for idx, name in enumerate(names, start=1):
            if self._cancel:
                self.log_insert("⚠️ Processo cancelado."); break
            try:
                doc = fitz.open(tpl)
                page = doc[active_area.page_index]
                rect = base_rect
                fontsize = fontsize_common if fontsize_common is not None else autosize_font_to_rect(name, rect, font_path)
                draw_name_centered_with_size(page, rect, name, font_path, fontsize)

                outname = f"Certificado - {sanitize_filename(evento)} - {sanitize_filename(name)}.pdf"
                outpath = os.path.join(outdir, outname)
                doc.save(outpath)
                if merged_doc is not None:
                    try:
                        one = fitz.open(outpath)
                        merged_doc.insert_pdf(one); one.close()
                    except Exception as ie:
                        self.log_insert(f"(!) Falha no PDF único: {ie}")
                doc.close()
                ok += 1; self.log_insert(f"[{idx}/{total}] OK: {name}")
            except Exception as e:
                fail += 1; self.log_insert(f"[{idx}/{total}] ERRO: {name} → {e}")
            self.pb["value"] = idx; self.progress_val.set(idx); self.update_idletasks()

        if (merged_doc is not None) and ok > 0:
            try:
                merged_path = os.path.join(outdir, f"Lote — {sanitize_filename(evento)}.pdf")
                merged_doc.save(merged_path)
                self.log_insert(f"PDF único criado: {merged_path}")
            except Exception as e:
                self.log_insert(f"Falha ao criar PDF único: {e}")
            finally:
                merged_doc.close()

        self.btn_generate.set_disabled(False); self.btn_cancel.set_disabled(True)
        if not self._cancel:
            self.log_insert(f"Concluído: {ok} ok, {fail} falhas.")
            messagebox.showinfo("Pronto", f"Concluído: {ok} ok, {fail} falhas.\nPasta: {outdir}")
            open_folder(outdir)
        else:
            self.log_insert(f"Interrompido: {ok} ok, {fail} falhas. Parcial salvo em: {outdir}")
        self.refresh_preview()

    def _show_help(self):
        messagebox.showinfo("Atalhos",
            "• Ctrl+O: Abrir PDF do modelo\n"
            "• Ctrl+Shift+O: Abrir lista de nomes (.CSV)\n"
            "• Ctrl+F: Escolher Fonte (.TTF)\n"
            "• Ctrl+G: Gerar Certificados\n"
            "• Na seleção de área: zoom com slider/+/–/roda • arrastar com botão direito/meio.")

    def _show_tutorial(self):
        messagebox.showinfo(
            "Como usar (guia rápido)",
            "1) Modelo (PDF): Clique em “Escolher...” e selecione o arquivo do certificado.\n"
            "2) Lista de nomes (CSV): Um arquivo com uma coluna de nomes. Excel/Sheets → salvar como CSV.\n"
            "3) Fonte (.TTF): Escolha a fonte que será usada para escrever o nome.\n"
            "4) Onde escrever: Clique em “Selecionar área...” e desenhe um retângulo onde o nome deve aparecer. "
            "Se não fizer, o app tenta achar sozinho.\n"
            "5) Prévia: Marque “Mostrar página inteira” para ver a página completa ou deixe só a área. Ajuste o zoom.\n"
            "6) Tamanho igual para todos: Ative se quiser que todos os nomes usem o mesmo tamanho (o maior que caiba para todos).\n"
            "7) Gerar: Clique em “GERAR CERTIFICADOS”. Opcional: marcar “Juntar tudo em 1 PDF”."
        )

    def _on_close(self): self.destroy()

if __name__ == "__main__":
    App().mainloop()