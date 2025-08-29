from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, Tuple

APP_NAME = "ContratoRPA"


MAPPING: Dict[str, Tuple[str, str]] = {
    # CONTRATANTE
    "RAZÃO SOCIAL SPE": ("CADASTRO DAS OBRAS", "O27"),
    "nome_completo_contratante": ("CADASTRO DAS OBRAS", "O28"),
    "endereço_completo_contratante": ("CADASTRO DAS OBRAS", "O29"),
    "cidade_contratante": ("CADASTRO DAS OBRAS", "O30"),
    "estado_contratante": ("CADASTRO DAS OBRAS", "O31"),
    "cnpj_contratante": ("CADASTRO DAS OBRAS", "O32"),
    "nacionalidade_contratante": ("CADASTRO DAS OBRAS", "O33"),
    "profissão_contratante": ("CADASTRO DAS OBRAS", "O34"),
    "estado_civil_contratante": ("CADASTRO DAS OBRAS", "O35"),
    "rg_contratante": ("CADASTRO DAS OBRAS", "O36"),
    "cpf_contratante": ("CADASTRO DAS OBRAS", "O37"),
    "telefoneContratante": ("CADASTRO DAS OBRAS", "O38"),
    "emailContratante": ("CADASTRO DAS OBRAS", "O39"),
    # CONTRATADA
    "RAZÃO SOCIAL": ("QUADRO DE CONCORRENCIA", "H9"),
    "cnpj_contratada": ("QUADRO DE CONCORRENCIA", "H10"),
    "endereço_completo_contratada": ("QUADRO DE CONCORRENCIA", "H11"),
    "cidade_contratada": ("QUADRO DE CONCORRENCIA", "H12"),
    "estado_contratada": ("QUADRO DE CONCORRENCIA", "H13"),
    "nacionalidade_contratada": ("QUADRO DE CONCORRENCIA", "H14"),
    "nome_completo_contratada": ("QUADRO DE CONCORRENCIA", "H15"),
    "estado_civil_contratada": ("QUADRO DE CONCORRENCIA", "H16"),
    "cpf_contratada": ("QUADRO DE CONCORRENCIA", "H18"),
    "rg_contratada": ("QUADRO DE CONCORRENCIA", "H17"),
    "profissão_contratada": ("QUADRO DE CONCORRENCIA", "H19"),
    "telefoneContratada": ("QUADRO DE CONCORRENCIA", "H21"),
    "emailContratada": ("QUADRO DE CONCORRENCIA", "H22"),
    # ANEXOS
    "data_anexoI": ("QUADRO DE CONCORRENCIA", "F20"),
    "data_anexo": ("QUADRO DE CONCORRENCIA", "F20"),
    "regime_contratacao": ("QUADRO DE CONCORRENCIA", "H66"),
    # OBJETO
    "concorrencia": ("QUADRO DE CONCORRENCIA", "D8"),
    # LOCAL
    "local_de_servico": ("QUADRO DE CONCORRENCIA", "D22"),
    # PREÇO
    "preço_total": ("QUADRO DE CONCORRENCIA", "H39"),
    # COMPOSIÇÃO
    "maoDeObraPercentual": ("QUADRO DE CONCORRENCIA", "H46"),
    "mao_de_obra": ("QUADRO DE CONCORRENCIA", "H47"),
    "materialPercentual": ("QUADRO DE CONCORRENCIA", "H48"),
    "materiais": ("QUADRO DE CONCORRENCIA", "H49"),
    "equipamentoPercentual": ("QUADRO DE CONCORRENCIA", "H50"),
    "equipamentos": ("QUADRO DE CONCORRENCIA", "H51"),
    # PRAZO
    "dateInicio": ("QUADRO DE CONCORRENCIA", "F20"),
    "dateFim": ("CRONOGRAMA", "Q9"),
    "dateConcluida": ("QUADRO DE CONCORRENCIA", "H53"),
    # PAGAMENTO
    "pagamento": ("QUADRO DE CONCORRENCIA", "H45"),
    # OBS
    "observacao": ("QUADRO DE CONCORRENCIA", "H65"),
    # REAJUSTE
    "R1": ("CONTRATO", "V47"),
    "R2": ("CONTRATO", "V48"),
    "R3": ("CONTRATO", "V51"),
    # RETENÇÃO
    "R4": ("CONTRATO", "V79"),
    "R5": ("CONTRATO", "V80"),
    "numero": ("CONTRATO", "W81"),
    "retencaoMeses": ("CONTRATO", "W82"),
    # PARA CONTRATANTE
    "atencaoContratante": ("CONTRATO", "M56"),
    "contatoContratante": ("CONTRATO", "M57"),
    "endContratante": ("CONTRATO", "M58"),
    "telContratante": ("CONTRATO", "M59"),
    "mailContratante": ("CONTRATO", "M60"),
    # PARA CONTRATADA
    "atencaoContratada": ("CONTRATO", "M64"),
    "contatoContratada": ("CONTRATO", "M65"),
    "endContratada": ("CONTRATO", "M66"),
    "telContratada": ("CONTRATO", "M67"),
    "mailContratada": ("CONTRATO", "M68"),
    # INTERVENIENTE ANUENTE
    "contatoAnuente": ("CONTRATO", "M73"),
    "endAnuente": ("CONTRATO", "M74"),
    "telAnuente": ("CONTRATO", "M75"),
    "mailAnuente": ("CONTRATO", "M76"),
}

BASE_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"

EXCEL_PATH = INPUT_DIR / "template_spreadsheet.xlsx"
TEMPLATE_PATH = INPUT_DIR / "model_contract.docx"


def ensure_dirs() -> None:
    """Garante pastas essenciais (por quê: evita falhas por diretório ausente)."""
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def set_runtime_paths(excel_path: Path, output_dir: Path) -> None:
    """
    Atualiza paths em runtime quando o usuário escolhe planilha e saída.
    Por quê: centraliza side-effect que já existia espalhado.
    """
    global EXCEL_PATH, OUTPUT_DIR, BASE_DIR
    EXCEL_PATH = Path(excel_path)
    OUTPUT_DIR = Path(output_dir)
    BASE_DIR = OUTPUT_DIR


def configure_logging(level: int = logging.INFO) -> None:
    """Configura logging com formatação única para CLI e GUI."""
    if logging.getLogger().handlers:
        return
    ensure_dirs()
    logfile = str((BASE_DIR / "contrato_rpa.log").resolve())
    logging.basicConfig(
        level=level,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[logging.FileHandler(logfile, encoding="utf-8"), logging.StreamHandler()],
    )