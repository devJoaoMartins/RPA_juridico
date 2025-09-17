# RPA Jurídico — Gerador Automático de Contratos

## Resumo

RPA Jurídico preenche modelos DOCX com dados do Excel (.xlsm), converte/combina em PDF e gera um PDF final pronto para distribuição. Projeto pensado para Windows com automação via COM (Word/Excel) e opção de empacotar como .exe.

## Principais pontos

- Entrada: planilha Excel (.xlsm) e template Word (.docx).
- Saída: DOCX preenchido e PDF final mesclado.
- Conversões: usa docx2pdf quando disponível; em Windows tenta Word COM para DOCX→PDF e Excel→PDF.
- Os caminhos principais são controlados por config.EXCEL_PATH e config.OUTPUT_DIR.

## Requisitos

- Windows (recomendado para Excel→PDF via COM).
- Microsoft Word e Excel instalados (exigidos para exportações/conversões COM).
- Para desenvolvimento: Python 3.10+ e dependências em requirements.txt.
- Usuários finais NÃO precisam de Python se receberem o .exe empacotado, mas precisam do Office.

## Arquivos importantes

- Código: src/
- Pós-processo (conversão + merge): src/post_process.py
- Leitura Excel: src/excel_reader.py
- Escrita Word: src/word_writer.py
- Mapeamento: src/config.py (MAPPING, EXCEL_PATH, OUTPUT_DIR)
- Entradas: data/input/
- Saída: data/output/

## Como usar (usuário final)

1. Coloque a planilha Excel (padrão em data/input/) ou selecione-a na UI.
2. Feche Word/Excel antes de executar o programa.
3. Execute o .exe (duplo clique) ou rode python src/main.py num ambiente com Python.
4. Se o Windows mostrar aviso SmartScreen: clicar "Mais informações" → "Executar assim mesmo".
5. Se o executável foi baixado, pode ser necessário desbloqueá‑lo (Propriedades → Desbloquear) ou:
   - PowerShell: Unblock-File .\seu_programa.exe

Observação: O .exe empacotado com PyInstaller inclui o runtime Python — o usuário não precisa instalar Python.

## Empacotamento e distribuição (simples)

- Compactar em ZIP: inclua o .exe, a planilha e este README. ZIP reduz problemas no Google Drive.
- Comando PyInstaller (exemplo):
  ```bash
  pyinstaller --onefile --windowed --add-data "src/assets;assets" src/app.py
  ```
- Para reduzir avisos do Windows: assinar digitalmente o executável (Code Signing).

## Comportamento de arquivos temporários

- O pós-processo cria uma pasta temporária em OUTPUT*DIR chamada \_finalData*{ts} para armazenar PDFs intermediários.
- O código tenta remover essa pasta ao final (best‑effort). Se arquivos continuarem presos, é porque o Windows mantém handles abertos por objetos COM (Word/Excel).
  Soluções:
  - Fechar Word/Excel antes de rodar.
  - Se permanecer, aguardar alguns segundos e excluir manualmente.
  - Melhorias possíveis no código: adicionar gc.collect(), retries e pequenos delays antes de remover a pasta (ver src/post_process.py).

## Nomes de arquivos e sobrescrita

- Os arquivos gerados incluem timestamp com hora/minuto/segundo (formato %d-%m-%y\_%H-%M-%S) — portanto execuções não costumam sobrescrever saídas anteriores.
- Se desejar outra estratégia (UUID, contador sequencial), altere a função \_ts() em src/post_process.py.

## Erros comuns

- "Excel não encontrado": confirme config.EXCEL_PATH aponta para a planilha correta.
- Falha Excel→PDF: exige Excel/Windows/COM.
- Pasta _finalData_\* não removida: fechar Office e tentar novamente; ver item acima sobre remoção.

## Boas práticas para compartilhar

- ZIP com README e planilha.
- Instruir usuários a desbloquear o .exe se necessário.
- Testar o .exe em uma máquina limpa/VM antes de distribuir.
- Para uso corporativo: assinar o executável ou distribuir via repositório/internal share confiável.

## Desenvolvimento

- Recomendado usar virtualenv e instalar requirements.txt.
- Executar localmente com: python src/main.py
- Logs: contrato_rpa.log para diagnóstico.

## Suporte / manutenção

- Para mudar mapeamentos, editar src/config.py (MAPPING).
- Para ajustar ranges exportados, editar src/post_process.py (lista de tasks em build_final_pdf).
- Para reduzir problemas com pastas temporárias, aplicar retry + gc.collect() no bloco de remoção em src/post_process.py.

---

Atualize config.EXCEL_PATH e OUTPUT_DIR conforme seu ambiente antes de distribuir.
