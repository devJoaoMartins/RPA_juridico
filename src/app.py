# pyright: reportArgumentType=false
from __future__ import annotations

import importlib
import logging
import os
import sys
import threading
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Union

import tkinter as tk
from tkinter import BooleanVar, StringVar, filedialog, messagebox, ttk

import config
import post_process
from excel_reader import ExcelReader
from word_writer import WordWriter


# ──────────────────────────────── UI helpers ─────────────────────────────────
def resource_path(relative: str | Path) -> Path:
    """
    Compatível com PyInstaller (onefile).
    Por quê: garante que assets empacotados sejam encontrados no runtime.
    """
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return (base / relative).resolve()


def draw_round_rect(
    canvas: tk.Canvas,
    x1: int,
    y1: int,
    x2: int,
    y2: int,
    r: int,
    *,
    fill: str,
    outline: str = "",
) -> None:
    canvas.create_arc(x1, y1, x1 + 2 * r, y1 + 2 * r, start=90, extent=90, style="pieslice", outline=outline, fill=fill)
    canvas.create_arc(x2 - 2 * r, y1, x2, y1 + 2 * r, start=0, extent=90, style="pieslice", outline=outline, fill=fill)
    canvas.create_arc(x2 - 2 * r, y2 - 2 * r, x2, y2, start=270, extent=90, style="pieslice", outline=outline, fill=fill)
    canvas.create_arc(x1, y2 - 2 * r, x1 + 2 * r, y2, start=180, extent=90, style="pieslice", outline=outline, fill=fill)
    canvas.create_rectangle(x1 + r, y1, x2 - r, y2, outline=outline, fill=fill)
    canvas.create_rectangle(x1, y1 + r, x2, y2 - r, outline=outline, fill=fill)


def set_window_icon(window: tk.Tk | tk.Toplevel) -> None:
    """
    Ícone da janela: prioriza PNG (logo-32.png → iconphoto), fallback ICO.
    Por quê: iconphoto aceita PNG com transparência; fallback evita crash.
    """
    png32 = resource_path(Path("assets") / "logo-32.png")
    png = resource_path(Path("assets") / "logo.png")
    ico = resource_path(Path("assets") / "logo.ico")

    for candidate in (png32, png):
        if candidate.exists():
            try:
                img = tk.PhotoImage(file=str(candidate))
                window.iconphoto(True, img)
                # evita GC do ícone
                if not hasattr(window, "_icon_refs"):
                    window._icon_refs = []  # type: ignore[attr-defined]
                window._icon_refs.append(img)  # type: ignore[attr-defined]
                return
            except Exception:
                pass
    if ico.exists():
        try:
            window.iconbitmap(default=str(ico))
        except Exception:
            pass


def _patch_messagebox() -> None:
    """
    Monkeypatch para nunca exibir 'None' como título/mensagem.
    Por quê: melhora UX em casos de exceções inesperadas.
    """

    def _title(x):
        s = "Erro" if x is None else str(x).strip()
        return s or "Erro"

    def _msg(x):
        try:
            s = "" if x is None else str(x).strip()
        except Exception:
            s = ""
        return s if s and s.lower() != "none" else "Ocorreu um erro inesperado."

    _e, _w, _i = messagebox.showerror, messagebox.showwarning, messagebox.showinfo
    messagebox.showerror = lambda t=None, m=None, *a, **k: _e(_title(t), _msg(m), *a, **k)  # type: ignore
    messagebox.showwarning = lambda t=None, m=None, *a, **k: _w(_title(t), _msg(m), *a, **k)  # type: ignore
    messagebox.showinfo = lambda t=None, m=None, *a, **k: _i(_title(t or "Info"), _msg(m), *a, **k)  # type: ignore


# ────────────────────────────────── App ───────────────────────────────────────
class App:
    def __init__(self, root: tk.Tk) -> None:
        _patch_messagebox()
        config.configure_logging()  

        self.root = root
        self.root.title(config.APP_NAME)
        set_window_icon(self.root)

        self._center(980, 620)
        self.root.minsize(860, 560)

        # Paleta
        self.BG = "#F4F7FA"
        self.CARD_BG = "#FFFFFF"
        self.SHADOW = "#E6ECF5"
        self.PRIMARY = "#0A4F97"
        self.PRIMARY_HOVER = "#0A4687"
        self.MUTED = "#6B7280"
        self.root.configure(bg=self.BG)

        # State
        self.excel_var = StringVar()
        self.outdir_var = StringVar()
        self.open_pdf_var = BooleanVar(value=True)

        # Canvas
        self.canvas = tk.Canvas(self.root, highlightthickness=0, bg=self.BG)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.root.bind("<Configure>", self._on_resize)

        # Logo para canto superior
        self.logo_img: tk.PhotoImage | None = None
        self.logo_small: tk.PhotoImage | None = None
        try:
            logo_path = resource_path(Path("assets") / "logo.png")
            if logo_path.exists():
                self.logo_img = tk.PhotoImage(file=str(logo_path))
        except Exception:
            self.logo_img = None

        # Card
        self.card = ttk.Frame(self.root, padding=24)
        self._style_widgets()
        self._build_card_widgets()

        self._redraw()

    # ─────────────────────────── layout ───────────────────────────
    def _center(self, w: int, h: int) -> None:
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _on_resize(self, _evt=None) -> None:
        self._redraw()

    def _redraw(self) -> None:
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        self.canvas.delete("all")
        self.canvas.create_rectangle(0, 0, w, h, fill=self.BG, outline="")

        # Card central com sombra
        card_w, card_h = 520, 420
        cx, cy = w // 2, h // 2
        x1, y1 = cx - card_w // 2, cy - card_h // 2
        x2, y2 = x1 + card_w, y1 + card_h
        draw_round_rect(self.canvas, x1 + 10, y1 + 12, x2 + 10, y2 + 12, 24, fill=self.SHADOW, outline="")
        draw_round_rect(self.canvas, x1, y1, x2, y2, 24, fill=self.CARD_BG, outline="")
        self.card.place(x=x1 + 24, y=y1 + 24, width=card_w - 48, height=card_h - 48)

        # Logo topo-esquerdo (escalada)
        if self.logo_img is not None:
            target_h = 24
            k = max(1, self.logo_img.height() // target_h)
            try:
                self.logo_small = self.logo_img.subsample(k, k)
            except Exception:
                self.logo_small = self.logo_img
            self.canvas.create_image(20, 20, image=self.logo_small, anchor="w")

    def _style_widgets(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure("Card.TFrame", background=self.CARD_BG)
        style.configure("Card.TLabel", background=self.CARD_BG, foreground="#0F172A", font=("Segoe UI", 11))
        style.configure("Title.TLabel", background=self.CARD_BG, foreground="#0F172A", font=("Segoe UI", 20, "bold"))
        style.configure("Sub.TLabel", background=self.CARD_BG, foreground=self.MUTED, font=("Segoe UI", 10))

        style.configure("Clean.TEntry", fieldbackground="#F6F8FA", bordercolor="#E5E7EB", relief="flat", padding=6)
        style.configure("Clean.TCheckbutton", background=self.CARD_BG)

        style.configure(
            "Primary.TButton",
            background=self.PRIMARY,
            foreground="#FFFFFF",
            font=("Segoe UI", 11, "bold"),
            padding=(22, 10),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", self.PRIMARY_HOVER)],
            relief=[("pressed", "flat"), ("!pressed", "flat")],
        )

        style.configure(
            "Clean.Horizontal.TProgressbar",
            troughcolor="#FDE8D8",
            background="#F59E0B",
            bordercolor="#FDE8D8",
            lightcolor="#F59E0B",
            darkcolor="#F59E0B",
            thickness=10,
        )

        self.card.configure(style="Card.TFrame")

    def _build_card_widgets(self) -> None:
        ttk.Label(self.card, text="Contrato Orçamentário", style="Title.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w"
        )
        ttk.Label(self.card, text="Gere o contrato final a partir da planilha padrão.", style="Sub.TLabel").grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(2, 14)
        )

        # Excel
        ttk.Label(self.card, text="Arquivo Excel", style="Card.TLabel").grid(row=2, column=0, sticky="w")
        self.ent_excel = ttk.Entry(self.card, textvariable=self.excel_var, style="Clean.TEntry")
        self.ent_excel.grid(row=3, column=0, sticky="ew", pady=(4, 10))
        ttk.Button(self.card, text="🗎", width=3, command=self._pick_excel).grid(row=3, column=1, padx=(8, 0))

        # Saída
        ttk.Label(self.card, text="Pasta de saída", style="Card.TLabel").grid(row=4, column=0, sticky="w")
        self.ent_out = ttk.Entry(self.card, textvariable=self.outdir_var, style="Clean.TEntry")
        self.ent_out.grid(row=5, column=0, sticky="ew", pady=(4, 10))
        ttk.Button(self.card, text="📁", width=3, command=self._pick_outdir).grid(row=5, column=1, padx=(8, 0))

        ttk.Checkbutton(
            self.card, text="Abrir PDF ao finalizar", variable=self.open_pdf_var, style="Clean.TCheckbutton"
        ).grid(row=6, column=0, sticky="w", pady=(2, 10))

        self.pbar = ttk.Progressbar(self.card, mode="indeterminate", style="Clean.Horizontal.TProgressbar")
        self.pbar.grid(row=7, column=0, columnspan=2, sticky="ew", pady=(0, 14))

        self.btn_run = ttk.Button(self.card, text="Gerar contrato", style="Primary.TButton", command=self._run_clicked)
        self.btn_run.grid(row=8, column=0, sticky="ew")

        self.card.columnconfigure(0, weight=1)

    # ─────────────────────────── actions ──────────────────────────
    def _pick_excel(self) -> None:
        p = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Planilhas Excel", "*.xlsx;*.xlsm"), ("Todos os arquivos", "*.*")],
        )
        if p:
            self.excel_var.set(p)

    def _pick_outdir(self) -> None:
        p = filedialog.askdirectory(title="Selecione a pasta de saída")
        if p:
            self.outdir_var.set(p)

    def _safe_error(self, title: str, err: Union[str, BaseException]) -> None:
        """Mensagem amigável (fallback)."""
        try:
            s = "" if err is None else str(err).strip()
        except Exception:
            s = ""
        msg = s if s and s.lower() != "none" else "Ocorreu um erro inesperado."
        messagebox.showerror(title or "Erro", msg)

    # ────────────────────────── pipeline ──────────────────────────
    def _run_clicked(self) -> None:
        excel = Path(self.excel_var.get())
        outdir = Path(self.outdir_var.get())
        if not excel.exists():
            self._safe_error("Dados inválidos", "Selecione um arquivo Excel válido.")
            return
        if not outdir.exists():
            self._safe_error("Dados inválidos", "Selecione uma pasta de saída válida.")
            return

        self.btn_run.state(["disabled"])
        self.pbar.start(12)
        threading.Thread(target=self._run_pipeline, args=(excel, outdir), daemon=True).start()

    def _missing_list(self, values: Dict[str, str]) -> List[Tuple[str, str, str]]:
        miss: List[Tuple[str, str, str]] = []
        for marker, (sheet, cell) in config.MAPPING.items():
            v = values.get(marker, "")
            if v is None or str(v).strip() == "":
                miss.append((marker, sheet, cell))
        return miss

    def _run_pipeline(self, user_excel: Path, user_output_dir: Path) -> None:
        try:
            config.set_runtime_paths(user_excel, user_output_dir)

            template_path = resource_path(Path("assets") / "model_contract.docx")
            if not template_path.exists():
                raise FileNotFoundError(f"Template Word não encontrado no pacote: {template_path}")

            replacements: Dict[str, str] = {}
            with ExcelReader(user_excel) as reader:
                for marker, (sheet, cell) in config.MAPPING.items():
                    value = reader.get_cell_value(sheet, cell,)
                    replacements[marker] = "" if value is None else str(value)

            missing = self._missing_list(replacements)
            if missing:
                ts = datetime.now().strftime("%Y-%m-%d_%H%M")
                report = user_output_dir / f"campos_vazios_{ts}.txt"
                try:
                    report.write_text("\n".join(f"{m} — {s}!{c}" for m, s, c in missing), encoding="utf-8")
                except Exception:
                    pass
                msg = "Preencha no Excel antes de continuar:\n\n" + "\n".join(
                    f"• {m} — {s}!{c}" for m, s, c in missing
                )
                self.root.after(0, lambda: self._safe_error("Campos obrigatórios", msg))
                return

            timestamp = datetime.now().strftime("%d-%m-%y_%H-%M")
            filled_docx = user_output_dir / f"ContratoPreenchido_{timestamp}.docx"

            writer = WordWriter(template_path)
            if not writer.replace_in_document(replacements, filled_docx):
                raise RuntimeError("Falha na geração do DOCX.")

            # garante que post_process leia paths atualizados do config
            mod = importlib.reload(post_process)
            final_pdf = mod.build_final_pdf(filled_docx)
            if not final_pdf:
                raise RuntimeError("Pós-processamento falhou.")

            try:
                if filled_docx.exists():
                    filled_docx.unlink()
            except Exception:
                pass

            def _ok():
                messagebox.showinfo("Concluído", f"PDF final gerado:\n{final_pdf}")
                if self.open_pdf_var.get():
                    try:
                        os.startfile(final_pdf)  # Windows
                    except Exception:
                        try:
                            os.startfile(final_pdf.parent)
                        except Exception:
                            pass

            self.root.after(0, _ok)
        except Exception as e:
            logging.getLogger("app").exception("Erro na execução")
            self.root.after(0, lambda: self._safe_error("Erro", e))
        finally:
            self.root.after(0, lambda: self.pbar.stop())
            self.root.after(0, lambda: self.btn_run.state(["!disabled"]))


def main() -> None:
    root = tk.Tk()
    set_window_icon(root)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()