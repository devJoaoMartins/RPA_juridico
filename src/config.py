from pathlib import Path


# Mapeamento completo entre marcadores do Word e células do Excel.
MAPPING = {
    # CONTRATANTE
    'RAZÃO SOCIAL SPE': ('CADASTRO DAS OBRAS', 'O27'),
    'nome_completo_contratante': ('CADASTRO DAS OBRAS', '028'),
    'endereço_completo_contratante': ('CADASTRO DAS OBRAS', 'O29'),
    'cidade_contratante': ('CADASTRO DAS OBRAS', 'O30'),
    'estado_contratante': ('CADASTRO DAS OBRAS', 'O31'),
    'cnpj_contratante': ('CADASTRO DAS OBRAS', 'O32'),
    'nacionalidade_contratante': ('CADASTRO DAS OBRAS', 'O33'),
    'profissão_contratante': ('CADASTRO DAS OBRAS', 'O34'),
    'estado_civil_contratante': ('CADASTRO DAS OBRAS', 'O35'),
    'rg_contratante': ('CADASTRO DAS OBRAS', 'O36'),
    'cpf_contratante': ('CADASTRO DAS OBRAS', 'O37'),
    'telefoneContratante': ('CADASTRO DAS OBRAS', 'O38'),
    'emailContratante': ('CADASTRO DAS OBRAS', 'O39'),

    # CONTRATADA
    'RAZÃO SOCIAL': ('QUADRO DE CONCORRENCIA', 'H9'),
    'endereço_completo_contratada': ('QUADRO DE CONCORRENCIA', 'H11'),
    'cidade_contratada': ('QUADRO DE CONCORRENCIA', 'H12'),
    'estado_contratada': ('QUADRO DE CONCORRENCIA', 'H13'),
    'cnpj_contratada': ('QUADRO DE CONCORRENCIA', 'H10'),
    'nome_completo_contratada': ('QUADRO DE CONCORRENCIA', 'H15'),
    'nacionalidade_contratada': ('QUADRO DE CONCORRENCIA', 'H14'),
    'profissão_contratada': ('QUADRO DE CONCORRENCIA', 'H19'),
    'rg_contratada': ('QUADRO DE CONCORRENCIA', 'H17'),
    'estado_civil_contratada': ('QUADRO DE CONCORRENCIA', 'H16'),
    'cpf_contratada': ('QUADRO DE CONCORRENCIA', 'H18'),
    'telefoneContratada': ('QUADRO DE CONCORRENCIA', 'H21'),
    'emailContratada': ('QUADRO DE CONCORRENCIA', 'H22'),

    # ANEXOS
    'data_anexoI': ('QUADRO DE CONCORRENCIA', 'F20'),
    'data_anexoII': ('QUADRO DE CONCORRENCIA', 'F20'),

    # OBJETO DA PRESTAÇÃO DE SERVIÇOS
    'concorrencia': ('QUADRO DE CONCORRENCIA', 'D8'),

    # LOCAL DOS SERVIÇOS
    'local_de_serviço': ('QUADRO DE CONCORRENCIA', 'D21'),

    # PREÇO
    'preço_total': ('QUADRO DE CONCORRENCIA', 'H39'),

    # COMPOSIÇÃO
    'materiais': ('QUADRO DE CONCORRENCIA', 'H49'),
    'materialPercentual': ('QUADRO DE CONCORRENCIA', 'H48'),
    'equipamentos': ('QUADRO DE CONCORRENCIA', 'H51'),
    'equipamentoPercentual': ('QUADRO DE CONCORRENCIA', 'H50'),
    'mao_de_obra': ('QUADRO DE CONCORRENCIA', 'H47'),
    'maoDeObraPercentual': ('QUADRO DE CONCORRENCIA', 'H48'),

    # PRAZO
    'dateInicio': ('QUADRO DE CONCORRENCIA', 'F19'),
    'dateFim': ('CRONOGRAMA', 'Q9'),
    'dateConcluida': ('QUADRO DE CONCORRENCIA', 'H53'),

    # CONDIÇÕES DE PAGAMENTO
    'pagamento': ('QUADRO DE CONCORRENCIA', 'H44'),

    # OBSERVAÇÃO

    # REAJUSTE
    "R1": ('CONTRATO', 'V47'),
    "R2": ('CONTRATO', 'V48'),
    "R3": ('CONTRATO', 'V51'),

    # RETENÇÃO
    'R4': ('CONTRATO', 'V79'),
    'R5': ('CONTRATO', 'V80'),
    'numero': ('CONTRATO', 'W82'),
    'retencaoMeses': ('CONTRATO', 'W83'),

    # PARA CONTRATANTE
    'atencaoContratante': ('CONTRATO', 'M56'),
    'contatoContratante': ('CONTRATO', 'M57'),
    'endContratante': ('CONTRATO', 'M58'),
    'telContratante': ('CONTRATO', 'M59'),
    'mailContratante': ('CONTRATO', 'M560'),

    # PARA CONTRATADA
    'atencaoContratada': ('CONTRATO', 'M64'),
    'contatoContratada': ('CONTRATO', 'M65'),
    'endContratada': ('CONTRATO', 'M66'),
    'telContratada': ('CONTRATO', 'M67'),
    'mailContratada': ('CONTRATO', 'M68'),

    # INTERVENIENTE ANUENTE
    'contatoAnuente': ('CONTRATO', 'M73'),
    'endAnuente': ('CONTRATO', 'M74'),
    'telAnuente': ('CONTRATO', 'M75'),
    'mailAnuente': ('CONTRATO', 'M76'),



}



# raiz do projeto (um nível acima de src)
BASE_DIR = Path(__file__).resolve().parents[1]
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"

EXCEL_PATH = INPUT_DIR / "template_spreadsheet.xlsm"
TEMPLATE_PATH = INPUT_DIR / "model_contract.docx"
