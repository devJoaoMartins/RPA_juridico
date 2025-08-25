from __future__ import annotations
import logging
from pathlib import Path
from datetime import datetime
from typing import Iterable, Optional
from config import BASE_DIR, EXCEL_PATH, OUTPUT_DIR

LOG_PATH = BASE_DIR / "contrato_rpa.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler(str(LOG_PATH), encoding="utf-8"), logging.StreamHandler()],
)
logger = logging.getLogger("post_process")

def _ts() -> str:
    return datetime.now().strftime("%d-%m-%y")

def _convert_docx_to_pdf(docx_path: Path, out_pdf: Path) -> None:
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    try:
        from docx2pdf import convert
        convert(str(docx_path), str(out_pdf))
        logger.info(f"DOCX→PDF via docx2pdf: {out_pdf}")
        return
    except Exception as e:
        logger.warning(f"docx2pdf falhou ({e}); tentando Word COM...")
    try:
        import win32com.client as win32  # type: ignore
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(docx_path))
        doc.ExportAsFixedFormat(OutputFileName=str(out_pdf), ExportFormat=17)  # PDF
        doc.Close(False)
        word.Quit()
        logger.info(f"DOCX→PDF via Word COM: {out_pdf}")
    except Exception as e:
        logger.error(f"Falha DOCX→PDF: {e}")
        raise

def _export_excel_range_to_pdf(xlsm: Path, sheet: str, rng: str, out_pdf: Path, *, landscape: Optional[bool] = None) -> None:
    import pythoncom
    import win32com.client as win32  # type: ignore
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(xlsm), ReadOnly=True, UpdateLinks=0)
        ws = wb.Worksheets(sheet)
        ws.PageSetup.PrintArea = rng
        ps = ws.PageSetup
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
        if landscape is not None:
            ps.Orientation = 2 if landscape else 1  # por quê: cronograma é horizontal
        ws.ExportAsFixedFormat(
            Type=0,  # PDF
            Filename=str(out_pdf),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False,
        )
        wb.Close(SaveChanges=False)
        logger.info(f"Excel→PDF {sheet}!{rng}: {out_pdf}")
    except Exception as e:
        logger.error(f"Falha Excel→PDF ({sheet}!{rng}): {e}")
        raise
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()

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
    """Executa todo o pós-processo e retorna o caminho do PDF final."""
    if not EXCEL_PATH.exists():
        logger.error(f"Excel não encontrado: {EXCEL_PATH}")
        return None

    if filled_docx is None:
        # fallback: pega o mais recente
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
        _export_excel_range_to_pdf(EXCEL_PATH, "QUADRO DE CONCORRENCIA", "A1:K131", pdf_quadro, landscape=False)
        _export_excel_range_to_pdf(EXCEL_PATH, "CRONOGRAMA", "B2:T26", pdf_crono, landscape=True)
        _export_excel_range_to_pdf(EXCEL_PATH, "CHECKLIST", "B2:E36", pdf_check, landscape=False)
        _merge_pdfs([pdf_docx, pdf_quadro, pdf_crono, pdf_check], final_pdf)
        return final_pdf
    finally:
        # limpeza best-effort
        for f in (pdf_docx, pdf_quadro, pdf_crono, pdf_check):
            try:
                if f.exists():
                    f.unlink()
            except Exception:
                pass
        try:
            if tmp_dir.exists():
                tmp_dir.rmdir()
        except Exception:
            pass