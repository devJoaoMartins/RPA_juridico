from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import config
from config import BASE_DIR, EXCEL_PATH, INPUT_DIR, MAPPING, OUTPUT_DIR, TEMPLATE_PATH
from excel_reader import ExcelReader
from post_process import build_final_pdf
from word_writer import WordWriter

config.configure_logging()


def _preflight(logger: logging.Logger) -> bool:
    ok = True
    if not Path(INPUT_DIR).exists():
        logger.error(f"Pasta input não existe: {INPUT_DIR}")
        ok = False
    if not Path(EXCEL_PATH).exists():
        logger.error(f"Excel não encontrado: {EXCEL_PATH}")
        ok = False
    if not Path(TEMPLATE_PATH).exists():
        logger.error(f"Template Word não encontrado: {TEMPLATE_PATH}")
        ok = False
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    return ok


def _missing_list(values: Dict[str, str]) -> List[Tuple[str, str, str]]:
    miss: List[Tuple[str, str, str]] = []
    for marker, (sheet, cell) in MAPPING.items():
        v = values.get(marker, "")
        if v is None or str(v).strip() == "":
            miss.append((marker, sheet, cell))
    return miss


def _safe_msg(e: BaseException | str, default: str = "Ocorreu um erro inesperado.") -> str:
    s = str(e)
    return s if s and s.strip() else default


def main() -> None:
    logger = logging.getLogger(__name__)
    if not _preflight(logger):
        return

    try:
        with ExcelReader(EXCEL_PATH) as reader:
            replacements = reader.read_mapping(MAPPING)
    except Exception as e:
        logger.exception("Falha na leitura do Excel")
        logger.error(_safe_msg(e, "Falha na leitura do Excel."))
        return

    missing = _missing_list(replacements)
    if missing:
        ts = datetime.now().strftime("%Y-%m-%d_%H%M")
        report = OUTPUT_DIR / f"campos_vazios_{ts}.txt"
        try:
            report.write_text("\n".join(f"{m} — {s}!{c}" for m, s, c in missing), encoding="utf-8")
        except Exception:
            pass
        logger.error(
            "Existem %d campos obrigatórios sem preenchimento. Preencha no Excel antes de continuar.\n%s",
            len(missing),
            "\n".join(f" - {m} — {s}!{c}" for m, s, c in missing),
        )
        return

    timestamp = datetime.now().strftime("%d-%m-%y_%H-%M")
    output_path = Path(OUTPUT_DIR) / f"ContratoPreenchido_{timestamp}.docx"

    try:
        writer = WordWriter(TEMPLATE_PATH)
        if not writer.replace_in_document(replacements, output_path):
            raise RuntimeError("Falha na geração do DOCX.")
        final_pdf = build_final_pdf(output_path)
        if not final_pdf:
            raise RuntimeError("Pós-processamento falhou.")
        logger.info("Processo concluído! PDF final: %s", final_pdf)
    except Exception as e:
        logger.exception("Erro na geração")
        logger.error(_safe_msg(e))


if __name__ == "__main__":
    main()