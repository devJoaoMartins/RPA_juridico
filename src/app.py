# pyright: reportArgumentType=false
from __future__ import annotations

import sys
import os
import logging
import threading
import importlib
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Text, StringVar, DISABLED, NORMAL, END

import post_process  # incluído no bundle; daremos reload após ajustar config


# ---------- util ----------
def resource_path(relative: str | Path) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return (base / relative).resolve()


class TkTextHandler(logging.Handler):
    def __init__(self, text_widget: Text):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record) + "\n"
        self.text_widget.configure(state=NORMAL)
        self.text_widget.insert(END, msg)
        self.text_widget.see(END)
        self.text_widget.configure(state=DISABLED)


# ---------- app ----------
class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ContratoRPA")
        self._center(960, 560)
        self.root.minsize(820, 520)

        self.excel_var = StringVar()
        self.outdir_var = StringVar()

        # Canvas de fundo (gradiente + logo)
        self.canvas = tk.Canvas(self.root, highlightthickness=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self.root.bind("<Configure>", self._on_resize)

        # Painel da esquerda (sempre por cima do Canvas)
        self.left_panel = ttk.Frame(self.root, padding=(16, 16))
        self._build_left_panel_widgets()
        self.left_panel.place(x=40, y=110, width=420)  # largura inicial; é ajustada no resize

        # Logging somente na UI
        self._wire_logging_to_ui()

        # Logo
        self.logo_img: tk.PhotoImage | None = None
        logo_path = resource_path(Path("assets") / "logo.png")
        if logo_path.exists():
            try:
                self.logo_img = tk.PhotoImage(file=str(logo_path))
            except Exception:
                self.logo_img = None

        # Primeira pintura
        self._redraw()

    # ---------- layout/draw ----------
    def _center(self, w: int, h: int) -> None:
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = (sw - w) // 2, (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _on_resize(self, _evt) -> None:
        self._redraw()
        # Ajusta largura do painel para caber na área esquerda (55% da janela)
        w = self.root.winfo_width()
        left_w = int(max(360, w * 0.55))
        panel_width = max(320, left_w - 80)
        self.left_panel.place_configure(width=panel_width)

    def _draw_gradient_and_logo(self) -> None:
        w = max(1, self.root.winfo_width())
        h = max(1, self.root.winfo_height())
        self.canvas.delete("all")

        left_w = int(w * 0.55)

        # fundo branco
        self.canvas.create_rectangle(0, 0, w, h, fill="#ffffff", outline="")

        # gradiente vertical (pêssego -> azul)
        start = (246, 218, 204)  # #F6DACC aprox
        end = (208, 223, 255)    # #D0DFFF aprox
        steps = h
        for i in range(steps):
            r = int(start[0] + (end[0] - start[0]) * (i / steps))
            g = int(start[1] + (end[1] - start[1]) * (i / steps))
            b = int(start[2] + (end[2] - start[2]) * (i / steps))
            self.canvas.create_line(0, i, left_w, i, fill=f"#{r:02x}{g:02x}{b:02x}")

        # divisor
        self.canvas.create_rectangle(left_w, 0, left_w + 1, h, fill="#e9ecf3", outline="")

        # logo central no lado direito
        if self.logo_img is not None:
            img = self.logo_img
            ih = img.height()
            target_h = max(1, int(h * 0.45))
            if ih > target_h:
                k = max(1, ih // target_h)
                try:
                    img = self.logo_img.subsample(k, k)
                except Exception:
                    img = self.logo_img
            cx = left_w + (w - left_w) // 2
            cy = h // 2
            self.canvas.create_image(cx, cy, image=img)
            self._current_logo = img  # manter referência

    def _redraw(self) -> None:
        self._draw_gradient_and_logo()

    # ---------- left panel ----------
    def _build_left_panel_widgets(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TButton", padding=(10, 6))
        style.configure("TEntry", padding=4)

        # Linha 1: Excel
        ttk.Label(self.left_panel, text="Arquivo Excel").grid(row=0, column=0, sticky="w")
        self.ent_excel = ttk.Entry(self.left_panel, textvariable=self.excel_var)
        self.ent_excel.grid(row=1, column=0, sticky="ew", pady=(2, 8))
        self.btn_excel = ttk.Button(self.left_panel, text="Selecionar…", command=self._pick_excel)
        self.btn_excel.grid(row=1, column=1, padx=(8, 0))

        # Linha 2: Saída
        ttk.Label(self.left_panel, text="Pasta de saída").grid(row=2, column=0, sticky="w")
        self.ent_out = ttk.Entry(self.left_panel, textvariable=self.outdir_var)
        self.ent_out.grid(row=3, column=0, sticky="ew", pady=(2, 8))
        self.btn_out = ttk.Button(self.left_panel, text="Selecionar…", command=self._pick_outdir)
        self.btn_out.grid(row=3, column=1, padx=(8, 0))

        # Ação
        self.btn_run = ttk.Button(self.left_panel, text="Gerar contrato", command=self._run_clicked)
        self.btn_run.grid(row=4, column=0, sticky="w", pady=(6, 4))

        # Progresso + etapas
        self.pbar = ttk.Progressbar(self.left_panel, mode="indeterminate")
        self.pbar.grid(row=5, column=0, columnspan=2, sticky="ew")

        ttk.Label(self.left_panel, text="Etapas:").grid(row=6, column=0, sticky="w", pady=(6, 0))
        self.txt_log = Text(self.left_panel, height=6, state=DISABLED, borderwidth=0)
        self.txt_log.grid(row=7, column=0, columnspan=2, sticky="nsew")

        self.left_panel.columnconfigure(0, weight=1)
        self.left_panel.rowconfigure(7, weight=1)

    # ---------- actions ----------
    def _wire_logging_to_ui(self) -> None:
        for h in list(logging.getLogger().handlers):
            logging.getLogger().removeHandler(h)
        handler = TkTextHandler(self.txt_log)
        handler.setFormatter(logging.Formatter("%(asctime)s - %(message)s"))
        logging.basicConfig(level=logging.INFO, handlers=[handler])

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

    def _run_clicked(self) -> None:
        excel = Path(self.excel_var.get())
        outdir = Path(self.outdir_var.get())
        if not excel.exists():
            messagebox.showerror("Dados inválidos", "Selecione um arquivo Excel válido.")
            return
        if not outdir.exists():
            messagebox.showerror("Dados inválidos", "Selecione uma pasta de saída válida.")
            return

        self.txt_log.configure(state=NORMAL)
        self.txt_log.delete("1.0", END)
        self.txt_log.configure(state=DISABLED)

        self.btn_run.configure(state=tk.DISABLED)
        self.pbar.start(12)

        threading.Thread(target=self._run_pipeline, args=(excel, outdir), daemon=True).start()

    def _run_pipeline(self, user_excel: Path, user_output_dir: Path) -> None:
        logger = logging.getLogger("app")
        try:
            import config
            config.EXCEL_PATH = user_excel
            config.OUTPUT_DIR = user_output_dir
            config.BASE_DIR = user_output_dir  # tudo na pasta de saída

            template_path = resource_path(Path("assets") / "model_contract.docx")
            if not template_path.exists():
                raise FileNotFoundError(f"Template Word não encontrado no pacote: {template_path}")

            from excel_reader import ExcelReader
            from word_writer import WordWriter

            replacements: dict[str, str] = {}
            with ExcelReader(user_excel) as reader:
                for marker, (sheet, cell) in config.MAPPING.items():
                    value = reader.get_cell_value(sheet, cell)
                    replacements[marker] = "" if value is None else str(value)
                    logger.info("Coletado: %s → %s", marker, replacements[marker])

            timestamp = datetime.now().strftime("%d-%m-%y_%H-%M")
            filled_docx = user_output_dir / f"ContratoPreenchido_{timestamp}.docx"

            writer = WordWriter(template_path)
            if not writer.replace_in_document(replacements, filled_docx):
                raise RuntimeError("Falha na geração do DOCX.")

            logger.info("DOCX gerado. Iniciando pós-processamento (PDF final)…")

            mod = importlib.reload(post_process)
            final_pdf = mod.build_final_pdf(filled_docx)
            if not final_pdf:
                raise RuntimeError("Pós-processamento falhou. Veja as etapas acima.")

            try:
                if filled_docx.exists():
                    filled_docx.unlink()
                    logger.info("DOCX intermediário removido.")
            except Exception:
                pass

            logger.info("Concluído! PDF final: %s", final_pdf)
            self._notify_ok(final_pdf)
        except Exception as e:
            logging.getLogger("app").exception("Erro na execução: %s", e)
            self._notify_err(str(e))
        finally:
            self.pbar.stop()
            self.btn_run.configure(state=tk.NORMAL)

    def _notify_ok(self, final_pdf: Path) -> None:
        def _show():
            messagebox.showinfo("Concluído", f"PDF final gerado:\n{final_pdf}")
            try:
                os.startfile(final_pdf)
            except Exception:
                try:
                    os.startfile(final_pdf.parent)
                except Exception:
                    pass
        self.root.after(0, _show)

    def _notify_err(self, msg: str) -> None:
        self.root.after(0, lambda: messagebox.showerror("Erro", msg))


# ---------- main ----------
def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("TButton", padding=(10, 6))
    style.configure("TEntry", padding=4)

    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
