from __future__ import annotations

import logging
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional, Sequence, Tuple

from config import EXCEL_PATH, OUTPUT_DIR

logger = logging.getLogger("post_process")

_WD_EXPORT_PDF = 17
_XL_TYPE_PDF = 0
_XL_ORIENT_PORTRAIT = 1
_XL_ORIENT_LANDSCAPE = 2

def _ts() -> str:
    return datetime.now().strftime("%d-%m-%y")

@contextmanager
def _word_app():

    import pythoncom

    pythoncom.CoInitialize()
    app = None
    try:
        import win32com.client as win32  # type: ignore

        app = win32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        yield app
    finally:
        try:
            if app is not None:
                app.Quit()
        finally:
            pythoncom.CoUninitialize()


@contextmanager
def _excel_app():
    """
    Excel COM reusável (uma instância para todos os exports).
    Por quê: reduz overhead de abrir/fechar instância para cada aba.
    """
    import pythoncom

    pythoncom.CoInitialize()
    app = None
    try:
        import win32com.client as win32  # type: ignore

        app = win32.DispatchEx("Excel.Application")
        app.Visible = False
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        yield app
    finally:
        try:
            if app is not None:
                app.Quit()
        finally:
            pythoncom.CoUninitialize()


def _convert_docx_to_pdf(docx_path: Path, out_pdf: Path) -> None:
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    with _word_app() as word:
        doc = word.Documents.Open(str(docx_path))
        doc.ExportAsFixedFormat(OutputFileName=str(out_pdf), ExportFormat=_WD_EXPORT_PDF)
        doc.Close(False)
        logger.info(f"DOCX→PDF: {out_pdf}")


def _export_many_excel_ranges_to_pdf(
    xlsm: Path,
    tasks: Sequence[Tuple[str, str, Path, Optional[bool]]],
) -> None:
    with _excel_app() as excel:
        wb = excel.Workbooks.Open(str(xlsm), ReadOnly=True, UpdateLinks=0)
        try:
            for sheet, rng, out_pdf, landscape in tasks:
                out_pdf.parent.mkdir(parents=True, exist_ok=True)
                ws = wb.Worksheets(sheet)
                ws.PageSetup.PrintArea = rng
                ps = ws.PageSetup
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False
                if landscape is not None:
                    ps.Orientation = _XL_ORIENT_LANDSCAPE if landscape else _XL_ORIENT_PORTRAIT
                ws.ExportAsFixedFormat(
                    Type=_XL_TYPE_PDF,
                    Filename=str(out_pdf),
                    Quality=0,
                    IncludeDocProperties=True,
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False,
                )
                logger.info(f"Excel→PDF {sheet}!{rng}: {out_pdf}")
        finally:
            wb.Close(SaveChanges=False)


def _merge_pdfs(pdf_paths: Iterable[Path], out_pdf: Path) -> None:
    from PyPDF2 import PdfMerger

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    merger = PdfMerger()
    try:
        for p in pdf_paths:
            if not p.exists():
                raise FileNotFoundError(f"PDF ausente: {p}")
            merger.append(str(p))
        with open(out_pdf, "wb") as fp:
            merger.write(fp)
        logger.info(f"PDF final: {out_pdf}")
    finally:
        merger.close()


def build_final_pdf(filled_docx: Optional[Path] = None) -> Optional[Path]:
    """
    Executa o pós-processo completo e retorna o caminho do PDF final.
    Por quê: consolida conversões e exportações mantendo side-effects controlados.
    """
    if not EXCEL_PATH.exists():
        logger.error(f"Excel não encontrado: {EXCEL_PATH}")
        return None

    if filled_docx is None:
        docs = sorted(OUTPUT_DIR.glob("ContratoPreenchido_*.docx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not docs:
            logger.error("Nenhum DOCX encontrado em data/output.")
            return None
        filled_docx = docs[0]

    ts = _ts()
    tmp_dir = OUTPUT_DIR / f"_finalData_{ts}"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    pdf_docx = tmp_dir / "01_contrato.docx.pdf"
    pdf_quadro = tmp_dir / "02_quadro.pdf"
    pdf_crono = tmp_dir / "03_cronograma.pdf"
    pdf_check = tmp_dir / "04_checklist.pdf"
    final_pdf = OUTPUT_DIR / f"ContratoFinal_{ts}.pdf"

    try:
        _convert_docx_to_pdf(filled_docx, pdf_docx)
        _export_many_excel_ranges_to_pdf(
            EXCEL_PATH,
            [
                ("QUADRO DE CONCORRENCIA", "A1:K134", pdf_quadro, False),
                ("CRONOGRAMA", "B2:T26", pdf_crono, True),
                ("QUALIFICACAO", "B2:E36", pdf_check, False),
            ],
        )
        _merge_pdfs([pdf_docx, pdf_quadro, pdf_crono, pdf_check], final_pdf)
        return final_pdf
    finally:
        try:
            import shutil

            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
