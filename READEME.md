# RPA Jurídico — Gerador Automático de Contratos

## Resumo

RPA Jurídico é uma ferramenta simples e robusta para preencher modelos de contrato em DOCX a partir de dados contidos em uma planilha Excel (.xlsm) e gerar um PDF final combinando o contrato preenchido com exportações do Excel (quadro e cronograma). Foi projetado para fluxo Windows com suporte opcional a conversão via MS Word COM.

## Principais funcionalidades

- Leitura e normalização de valores do Excel (datas, moeda BRL, percentuais).
- Substituição de placeholders/marcadores no modelo Word, inclusive em tabelas, cabeçalhos e rodapés.
- Conversão DOCX→PDF (docx2pdf ou Word COM) e exportação de áreas do Excel como PDF.
- Mesclagem dos PDFs resultantes em um único PDF final pronto para distribuição.

## Arquitetura / arquivos relevantes

- Leitor do Excel: [`excel_reader.ExcelReader`](src/excel_reader.py) — [src/excel_reader.py](src/excel_reader.py)
- Escritor do Word (substituição de marcadores): [`word_writer.WordWriter`](src/word_writer.py) — [src/word_writer.py](src/word_writer.py)
- Pós-processamento (conversão e merge): [`post_process.build_final_pdf`](src/post_process.py) — [src/post_process.py](src/post_process.py)
- Mapeamento entre placeholders e células do Excel: [`config.MAPPING`](src/config.py) — [src/config.py](src/config.py)
- Orquestrador principal: [src/main.py](src/main.py) — ponto de entrada do processo.
- Dependências: [requirements.txt](requirements.txt)
- Regras de Git: [.gitignore](.gitignore)
- Dados de entrada: [data/input](data/input/) — inclui [data/input/template_spreadsheet.xlsm](data/input/template_spreadsheet.xlsm), [data/input/model_contract.docx](data/input/model_contract.docx) e [data/input/checklist.xlsx](data/input/checklist.xlsx)
- Saída gerada: [data/output](data/output/)
- Log principal: [contrato_rpa.log](contrato_rpa.log)

## Requisitos

- Python 3.10+ (testado em 3.11/3.13).
- Windows recomendado se planeja usar a conversão via Word COM. docx2pdf pode funcionar em outros sistemas, mas o suporte a export Excel→PDF exige Excel/COM (Windows).
- Instale as dependências listadas em [requirements.txt](requirements.txt).

## Instalação rápida

1. Clone / copie este repositório e abra o diretório raiz.
2. Crie e ative um virtualenv:
   - Windows:
     python -m venv .venv
     .venv\Scripts\activate
   - Unix/macOS:
     python -m venv .venv
     source .venv/bin/activate
3. Instale dependências:
   pip install -r requirements.txt

## Uso (modo padrão)

1. Coloque a planilha Excel com os dados em: [data/input/template_spreadsheet.xlsm](data/input/template_spreadsheet.xlsm).
2. Ajuste o template Word com os placeholders desejados em: [data/input/model_contract.docx](data/input/model_contract.docx).
3. Altere o mapeamento entre placeholders e células em: [`config.MAPPING`](src/config.py) — [src/config.py](src/config.py). Cada entrada tem a forma 'placeholder': ('NOME_DA_ABA', 'CÉLULA').
4. Execute o orquestrador:
   python src/main.py

Fluxo executado por main.py

- Valida caminhos e presença dos arquivos (ver [src/main.py](src/main.py)).
- Lê valores do Excel com [`excel_reader.ExcelReader`](src/excel_reader.py) — [src/excel_reader.py](src/excel_reader.py).
- Substitui marcadores no DOCX com [`word_writer.WordWriter`](src/word_writer.py) — [src/word_writer.py](src/word_writer.py).
- Salva DOCX preenchido em [data/output](data/output/). Ex.: ContratoPreenchido_DD-MM-AA_HH-MM.docx
- Executa [`post_process.build_final_pdf`](src/post_process.py) — [src/post_process.py](src/post_process.py) para gerar PDFs e mesclar em ContratoFinal_DD-MM-AA.pdf

## Personalização e manutenção

- Mapeamento de placeholders: edite [`config.MAPPING`](src/config.py) — [src/config.py](src/config.py) para adicionar/alterar campos.
- Templates: mantenha uma cópia "limpa" do template Word sem placeholders para referência: [data/input/model_contract.docx](data/input/model_contract.docx).
- Ajuste ranges de export no pós-processo em [`post_process.build_final_pdf`](src/post_process.py) — [src/post_process.py](src/post_process.py) (atualmente "A1:K131" para o quadro e "B2:T26" para o cronograma).
- Logs: verifique [contrato_rpa.log](contrato_rpa.log) para diagnóstico (configuração em [src/main.py](src/main.py) e [src/post_process.py](src/post_process.py)).

## Saída esperada

- DOCX preenchido: data/output/ContratoPreenchido_DD-MM-AA_HH-MM.docx
- PDF final mesclado: data/output/ContratoFinal_DD-MM-AA.pdf
- Arquivos temporários de pós-processo são criados em data/output/\_finalData_DD-MM-YY/ e removidos após execução (best-effort).

## Erros comuns e como resolver

- "Arquivo Excel NÃO encontrado": confirme que [data/input/template_spreadsheet.xlsm](data/input/template_spreadsheet.xlsm) existe.
- Marcadores não substituídos: confirme as chaves em [`config.MAPPING`](src/config.py) e se as células referenciadas possuem valor.
- Falha DOCX→PDF: primeiro tenta usar docx2pdf; se falhar e estiver no Windows, tenta Word COM (MS Word deve estar instalado). Ver [`post_process._convert_docx_to_pdf`](src/post_process.py) — [src/post_process.py](src/post_process.py).
- Falha Excel→PDF: essa operação usa COM (Excel), logo exige Windows + MS Excel. Consulte [`post_process._export_excel_range_to_pdf`](src/post_process.py) — [src/post_process.py](src/post_process.py).

## Boas práticas

- Faça commits pequenos e use branches para features/bugs.
- Teste com um conjunto reduzido de dados antes de processar em massa.
- Evite commitar arquivos sensíveis — veja [.gitignore](.gitignore).
- Mantenha uma cópia do template original e versionada offline.

## Desenvolvimento e testes

- Código principal está em [src/](src/).
- Para adicionar testes, crie uma pasta tests/ e configure pytest no CI/local.
- Verifique logging e mensagens no console para depuração rápida.

## Contribuição

1. Abra uma branch com nome claro (feature/bugfix).
2. Faça PR com descrição e screenshots (se aplicável).
3. Atualize [`config.MAPPING`](src/config.py) se adicionar placeholders novos.

## Licença

Este repositório não contém uma licença definida. Se desejar compartilhá-lo publicamente, adicione um arquivo LICENSE (ex.: MIT) na raiz.

## Contatos e pontos de alteração rápida

- Mapeamentos e paths: [`config.MAPPING`](src/config.py) — [src/config.py](src/config.py)
- Leitura/formatação Excel: [`excel_reader.ExcelReader`](src/excel_reader.py) — [src/excel_reader.py](src/excel_reader.py)
- Substituição Word: [`word_writer.WordWriter`](src/word_writer.py) — [src/word_writer.py](src/word_writer.py)
- Pós-processamento / merge: [`post_process.build_final_pdf`](src/post_process.py) — [src/post_process.py](src/post_process.py)
- Entrada principal: [src/main.py](src/main.py)
- Arquivos de entrada: [data/input/template_spreadsheet.xlsm](data/input/template_spreadsheet.xlsm), [data/input/model_contract.docx](data/input/model_contract.docx)
- Logs: [contrato_rpa.log](contrato_rpa.log)

## Exemplo rápido de execução

1. Ative o ambiente virtual.
2. pip install -r requirements.txt
3. python src/main.py

