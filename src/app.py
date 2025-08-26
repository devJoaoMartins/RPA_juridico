# pyright: reportArgumentType=false

from __future__ import annotations

import sys
import threading
import logging
from pathlib import Path
from datetime import datetime
from tkinter import Tk, ttk, filedialog, messagebox, StringVar, Text, DISABLED, NORMAL, END
import importlib
import os


def resource_path(relative: str | Path) -> Path:
    """Resolve arquivos de dados no bundle PyInstaller ou no dev."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return (base / relative).resolve()


class TkTextHandler(logging.Handler):
    """Handler que envia logs para o Text da GUI (por quê: feedback ao usuário)."""
    def __init__(self, text_widget: Text):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record) + "\n"
        self.text_widget.configure(state=NORMAL)
        self.text_widget.insert(END, msg)
        self.text_widget.see(END)
        self.text_widget.configure(state=DISABLED)


class App:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Contrato RPA")
        self._center(720, 420)
        self.root.minsize(660, 360)

        self.excel_var = StringVar()
        self.outdir_var = StringVar()
        self.status_var = StringVar(value="Pronto")

        self._build_ui()
        self._wire_logging_to_ui()

        import post_process as _pp
        self._post_process_modname = "post_process"

    # ---------- UI ----------
    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=(0, 2))

        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self.root)
        frm.pack(fill="both", expand=True)

        # Excel
        ttk.Label(frm, text="Arquivo Excel:").grid(row=0, column=0, sticky="w", **pad)
        ent_excel = ttk.Entry(frm, textvariable=self.excel_var)
        ent_excel.grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Selecionar…", command=self._browse_excel).grid(row=0, column=2, **pad)

        # Saída
        ttk.Label(frm, text="Pasta de saída:").grid(row=1, column=0, sticky="w", **pad)
        ent_out = ttk.Entry(frm, textvariable=self.outdir_var)
        ent_out.grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Selecionar…", command=self._browse_outdir).grid(row=1, column=2, **pad)

        # Ações
        self.btn_run = ttk.Button(frm, text="Gerar contrato", command=self._run_clicked)
        self.btn_run.grid(row=2, column=1, sticky="e", **pad)

        # Progresso
        self.pbar = ttk.Progressbar(frm, mode="indeterminate")
        self.pbar.grid(row=3, column=0, columnspan=3, sticky="ew", padx=10, pady=(0, 6))

        ttk.Label(frm, text="Etapas:").grid(row=4, column=0, sticky="w", padx=10)
        self.txt_log = Text(frm, height=6, state=DISABLED)
        self.txt_log.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))
        s = ttk.Scrollbar(frm, command=self.txt_log.yview)
        s.grid(row=5, column=3, sticky="ns", pady=(0, 10))
        self.txt_log.configure(yscrollcommand=s.set)

        # Statusbar
        status = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status.pack(fill="x", side="bottom")

        # Grid weights
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(5, weight=1)

    def _center(self, w: int, h: int) -> None:
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw // 2) - (w // 2)
        y = (sh // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _wire_logging_to_ui(self) -> None:
        self.ui_handler = TkTextHandler(self.txt_log)
        self.ui_handler.setFormatter(logging.Formatter("%(asctime)s - %(message)s"))

    def _browse_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Planilhas Excel", "*.xlsx;*.xlsm"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.excel_var.set(path)

    def _browse_outdir(self) -> None:
        path = filedialog.askdirectory(title="Selecione a pasta de saída")
        if path:
            self.outdir_var.set(path)

    # ---------- Execução ----------
    def _run_clicked(self) -> None:
        excel = Path(self.excel_var.get())
        outdir = Path(self.outdir_var.get())

        if not excel.exists():
            messagebox.showerror("Dados inválidos", "Selecione um arquivo Excel válido.")
            return
        if not outdir.exists():
            messagebox.showerror("Dados inválidos", "Selecione uma pasta de saída válida.")
            return

        # Limpa log
        self.txt_log.configure(state=NORMAL)
        self.txt_log.delete("1.0", END)
        self.txt_log.configure(state=DISABLED)

        self.status_var.set("Gerando…")
        self.btn_run.configure(state=DISABLED)
        self.pbar.start(12)

        # Thread para não travar a UI
        t = threading.Thread(target=self._run_pipeline, args=(excel, outdir), daemon=True)
        t.start()

    def _setup_logging(self) -> None:
        root_logger = logging.getLogger()
        for h in list(root_logger.handlers):
            root_logger.removeHandler(h)

        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
            handlers=[self.ui_handler],
        )

    def _run_pipeline(self, user_excel: Path, user_output_dir: Path) -> None:
        self._setup_logging()
        logger = logging.getLogger("app")
        try:
            # Override mínimo no config
            import config
            config.EXCEL_PATH = user_excel
            config.OUTPUT_DIR = user_output_dir
            config.BASE_DIR = user_output_dir

            template_path = resource_path(Path("assets") / "model_contract.docx")
            if not template_path.exists():
                raise FileNotFoundError(
                    f"Template Word não encontrado no pacote: {template_path}"
                )

            # Coleta de dados do Excel
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

            logger.info("DOCX gerado com sucesso. Iniciando pós-processamento (PDF final).")

            post_process = importlib.reload(importlib.import_module(self._post_process_modname))
            final_pdf = post_process.build_final_pdf(filled_docx)

            if not final_pdf:
                raise RuntimeError("Pós-processamento falhou. Veja as etapas acima.")

            try:
                if filled_docx.exists():
                    filled_docx.unlink()
                    logger.info("DOCX intermediário removido.")
            except Exception:
                pass

            logger.info("Processo concluído! PDF final: %s", final_pdf)
            self._notify_ok(final_pdf)

        except Exception as e:
            logging.getLogger("app").exception("Erro na execução: %s", e)
            self._notify_err(str(e))
        finally:
            self.status_var.set("Pronto")
            self.btn_run.configure(state=NORMAL)
            self.pbar.stop()

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


def main() -> None:
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
