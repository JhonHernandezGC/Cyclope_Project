"""
cyclope_gui.py
==============
Interfaz gráfica (tkinter) para el pipeline Cyclope. Pensada para empaquetarse
como un único Cyclope.exe con PyInstaller, de modo que el equipo solo tenga que:

    1. Abrir Cyclope.exe (doble clic).
    2. Seleccionar el/los PDF de la Cámara de Comercio.
    3. (Opcional) elegir la carpeta de salida.
    4. Pulsar "Ejecutar".

Las plantillas (carpeta Requirements) van empaquetadas dentro del .exe, así que
el usuario final no necesita tenerlas a mano.

tkinter viene incluido con Python (no requiere instalación).
"""
from __future__ import annotations

import os
import queue
import sys
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import cyclope_core as core

# Ruta de red por defecto para la Sección 3 (crear carpetas de contraparte)
DEFAULT_NETWORK_PATH = r"\\naeast.ad.jpmorganchase.com\cib2\dipls\NACIB2DIPLSSHARE00001\COB\Counterparties"


def resource_path(relative: str) -> Path:
    """Devuelve la ruta de un recurso, ya sea en desarrollo o dentro del .exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return Path(base) / relative


def default_output_dir() -> Path:
    desktop = Path.home() / "Desktop"
    root = desktop if desktop.exists() else Path.home()
    return root / "Cyclope_Output"


class CyclopeApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.pdf_paths: list[str] = []
        self.templates_dir = resource_path("Requirements")
        self.log_queue: "queue.Queue[str | None]" = queue.Queue()
        self.running = False

        root.title("Cyclope — Pre-Fill")
        root.geometry("760x620")
        root.minsize(680, 540)

        self._build_ui()
        self.root.after(100, self._drain_log)

    # ----------------------------------------------------------------- UI
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        title = tk.Label(self.root, text="Cyclope — Generador de documentos",
                         font=("Segoe UI", 14, "bold"))
        title.pack(anchor="w", **pad)

        # --- PDFs ---
        frm_pdf = ttk.LabelFrame(self.root, text="1) PDF(s) de la Cámara de Comercio")
        frm_pdf.pack(fill="x", **pad)
        self.lbl_pdf = ttk.Label(frm_pdf, text="(ningún archivo seleccionado)")
        self.lbl_pdf.pack(side="left", padx=8, pady=8)
        ttk.Button(frm_pdf, text="Seleccionar PDF(s)…",
                   command=self._pick_pdfs).pack(side="right", padx=8, pady=8)

        # --- Salida ---
        frm_out = ttk.LabelFrame(self.root, text="2) Carpeta de salida")
        frm_out.pack(fill="x", **pad)
        self.var_out = tk.StringVar(value=str(default_output_dir()))
        ttk.Entry(frm_out, textvariable=self.var_out).pack(
            side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(frm_out, text="Cambiar…",
                   command=self._pick_output).pack(side="right", padx=8, pady=8)

        # --- Opción Sección 3 (red) ---
        frm_net = ttk.LabelFrame(self.root, text="3) (Opcional) Crear carpetas de contraparte en la red")
        frm_net.pack(fill="x", **pad)
        self.var_net_on = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_net, text="Activar Sección 3",
                        variable=self.var_net_on,
                        command=self._toggle_net).pack(anchor="w", padx=8, pady=(8, 0))
        self.var_net_path = tk.StringVar(value=DEFAULT_NETWORK_PATH)
        self.ent_net = ttk.Entry(frm_net, textvariable=self.var_net_path, state="disabled")
        self.ent_net.pack(fill="x", padx=8, pady=8)

        # --- Botones ---
        frm_btn = tk.Frame(self.root)
        frm_btn.pack(fill="x", **pad)
        self.btn_run = ttk.Button(frm_btn, text="▶  Ejecutar", command=self._run)
        self.btn_run.pack(side="left")
        ttk.Button(frm_btn, text="Abrir carpeta de resultados",
                   command=self._open_output).pack(side="left", padx=8)

        # --- Log ---
        frm_log = ttk.LabelFrame(self.root, text="Progreso")
        frm_log.pack(fill="both", expand=True, **pad)
        self.txt = scrolledtext.ScrolledText(frm_log, height=12, wrap="word",
                                             font=("Consolas", 9))
        self.txt.pack(fill="both", expand=True, padx=4, pady=4)
        self.txt.configure(state="disabled")

    # ------------------------------------------------------------- acciones
    def _pick_pdfs(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Selecciona el/los PDF de la Cámara",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")])
        if paths:
            self.pdf_paths = list(paths)
            self.lbl_pdf.config(text=f"{len(self.pdf_paths)} archivo(s) seleccionado(s)")

    def _pick_output(self) -> None:
        d = filedialog.askdirectory(title="Carpeta de salida")
        if d:
            self.var_out.set(d)

    def _toggle_net(self) -> None:
        self.ent_net.config(state="normal" if self.var_net_on.get() else "disabled")

    def _open_output(self) -> None:
        out = Path(self.var_out.get())
        if not out.exists():
            messagebox.showinfo("Cyclope", "La carpeta de salida aún no existe.")
            return
        try:
            os.startfile(str(out))  # solo Windows
        except AttributeError:
            messagebox.showinfo("Cyclope", f"Carpeta de resultados:\n{out}")

    def _run(self) -> None:
        if self.running:
            return
        if not self.pdf_paths:
            messagebox.showwarning("Cyclope", "Primero selecciona al menos un PDF.")
            return
        if not self.templates_dir.exists():
            messagebox.showerror(
                "Cyclope",
                f"No se encontró la carpeta de plantillas:\n{self.templates_dir}\n\n"
                "Si ejecutas el .py directamente, coloca la carpeta 'Requirements' "
                "junto a este archivo. En el .exe las plantillas van empaquetadas.")
            return

        self.running = True
        self.btn_run.config(state="disabled")
        self._clear_log()
        out_dir = Path(self.var_out.get())
        net_on = self.var_net_on.get()
        net_path = self.var_net_path.get()

        t = threading.Thread(
            target=self._work, args=(out_dir, net_on, net_path), daemon=True)
        t.start()

    def _work(self, out_dir: Path, net_on: bool, net_path: str) -> None:
        log = self.log_queue.put
        try:
            out_dir.mkdir(parents=True, exist_ok=True)
            log("===== 1) EXTRACCIÓN =====")
            excel = core.run_extraction(self.pdf_paths, out_dir, log=log)
            log("\n===== 2) PRE-LLENADO =====")
            core.run_prefill(excel, self.templates_dir, out_dir, log=log)
            if net_on:
                log("\n===== 3) CARPETAS EN LA RED =====")
                core.create_counterparty_folders(excel, net_path, out_dir, log=log)
            log("\n✔ PROCESO COMPLETADO.")
        except SystemExit as exc:
            log(f"\n✖ ERROR: {exc}")
        except Exception as exc:  # noqa: BLE001
            log(f"\n✖ ERROR: {exc}")
            log(traceback.format_exc())
        finally:
            self.log_queue.put(None)  # señal de fin

    # --------------------------------------------------------------- logging
    def _drain_log(self) -> None:
        try:
            while True:
                item = self.log_queue.get_nowait()
                if item is None:
                    self.running = False
                    self.btn_run.config(state="normal")
                else:
                    self._append(item)
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log)

    def _append(self, text: str) -> None:
        self.txt.configure(state="normal")
        self.txt.insert("end", text + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _clear_log(self) -> None:
        self.txt.configure(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.configure(state="disabled")


def main() -> None:
    root = tk.Tk()
    CyclopeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
