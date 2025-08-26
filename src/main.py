import logging
from datetime import datetime
from pathlib import Path
from config import (MAPPING, EXCEL_PATH, TEMPLATE_PATH, OUTPUT_DIR, INPUT_DIR, BASE_DIR)
from excel_reader import ExcelReader
from word_writer import WordWriter

# pós-processamento automático
from post_process import build_final_pdf

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(str(BASE_DIR / "contrato_rpa.log"), encoding="utf-8"),
        logging.StreamHandler(),
    ],
)

def _preflight(logger: logging.Logger) -> bool:
    logger.info(f"BASE esperada: {BASE_DIR}")
    logger.info(f"Excel esperado : {EXCEL_PATH}")
    logger.info(f"Modelo Word   : {TEMPLATE_PATH}")
    ok = True
    if not Path(INPUT_DIR).exists():
        logger.error(f"Pasta input não existe: {INPUT_DIR}")
        ok = False
    else:
        listing = "\n".join(f" - {p.name}" for p in sorted(Path(INPUT_DIR).iterdir()))
        logger.info("Conteúdo de data/input:\n" + (listing or " <vazio>"))
    if not Path(EXCEL_PATH).exists():
        logger.error("Arquivo Excel NÃO encontrado no path acima.")
        ok = False
    if not Path(TEMPLATE_PATH).exists():
        logger.error("Modelo Word NÃO encontrado no path acima.")
        ok = False
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    return ok

def main() -> None:
    logger = logging.getLogger(__name__)
    logger.info("Iniciando processo de geração de contratos")

    if not _preflight(logger):
        logger.error("Interrompido por paths inválidos.")
        return

    # coleta do Excel
    replacements = {}
    try:
        with ExcelReader(EXCEL_PATH) as reader:
            for marker, (sheet, cell) in MAPPING.items():
                value = reader.get_cell_value(sheet, cell)
                replacements[marker] = value
                logger.info(f"Coletado: {marker} → {value}")
    except Exception as e:
        logger.error(f"Falha na leitura do Excel: {e}")
        return

    timestamp = datetime.now().strftime("%d-%m-%y_%H-%M")
    output_path = Path(OUTPUT_DIR) / f"ContratoPreenchido_{timestamp}.docx"

    writer = WordWriter(TEMPLATE_PATH)
    if writer.replace_in_document(replacements, output_path):
        logger.info("DOCX gerado com sucesso. Iniciando pós-processamento (PDF final).")
        final_pdf = build_final_pdf(output_path)
        if final_pdf:
            logger.info(f"Processo concluído! PDF final: {final_pdf}")
        else:
            logger.error("Pós-processamento falhou.")
    else:
        logger.error("Falha na geração do contrato")

if __name__ == "__main__":
    main()
