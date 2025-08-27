from __future__ import annotations
import logging
from pathlib import Path
from datetime import datetime
from typing import Iterable, Optional
from config import BASE_DIR, EXCEL_PATH, OUTPUT_DIR

# Use o logger do app
logger = logging.getLogger("post_process")

def _ts() -> str:
    return datetime.now().strftime("%d-%m-%y")

def _convert_docx_to_pdf(docx_path: Path, out_pdf: Path) -> None:
    """Usa Word COM diretamente; por quê: evitar travas do docx2pdf."""
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    import pythoncom
    pythoncom.CoInitialize()
    word = None
    try:
        import win32com.client as win32  # type: ignore
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(docx_path))
        # 17 = wdExportFormatPDF
        doc.ExportAsFixedFormat(OutputFileName=str(out_pdf), ExportFormat=17)
        doc.Close(False)
        logger.info(f"DOCX→PDF via Word COM: {out_pdf}")
    except Exception as e:
        logger.error(f"Falha DOCX→PDF: {e}")
        raise
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

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
            ps.Orientation = 2 if landscape else 1  # cronograma é horizontal
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
    """Executa todo o pós-processo e retorna o caminho do PDF final (apenas ele fica salvo)."""
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
        _export_excel_range_to_pdf(EXCEL_PATH, "QUADRO DE CONCORRENCIA", "A1:K134", pdf_quadro, landscape=False)
        _export_excel_range_to_pdf(EXCEL_PATH, "CRONOGRAMA", "B2:T26", pdf_crono, landscape=True)
        _export_excel_range_to_pdf(EXCEL_PATH, "QUALIFICACAO", "B2:E36", pdf_check, landscape=False)
        _merge_pdfs([pdf_docx, pdf_quadro, pdf_crono, pdf_check], final_pdf)
        return final_pdf
    finally:
        # Limpeza garantida do diretório temporário
        try:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
