import logging
from typing import Dict
from docx import Document

class WordWriter:
    def __init__(self, template_path):
        self.template_path = str(template_path)
        self.logger = logging.getLogger(__name__)

    def replace_in_document(self, replacements: Dict[str, str], output_path) -> bool:
        try:
            doc = Document(self.template_path)
        except Exception as e:
            self.logger.error(f"Erro ao abrir o modelo Word: {e}")
            return False

        try:
            self._replace_in_paragraphs(doc.paragraphs, replacements)
            for table in doc.tables:
                self._replace_in_table(table, replacements)
            for section in doc.sections:
                self._replace_in_paragraphs(section.header.paragraphs, replacements)
                self._replace_in_paragraphs(section.footer.paragraphs, replacements)
                for t in section.header.tables:
                    self._replace_in_table(t, replacements)
                for t in section.footer.tables:
                    self._replace_in_table(t, replacements)
            doc.save(str(output_path))
            self.logger.info(f"Contrato salvo em: {output_path}")
            return True
        except Exception as e:
            self.logger.error(f"Falha na geração do documento: {e}")
            return False

    def _replace_in_table(self, table, replacements: Dict[str, str]) -> None:
        for row in table.rows:
            for cell in row.cells:
                self._replace_in_paragraphs(cell.paragraphs, replacements)
                for inner in cell.tables:
                    self._replace_in_table(inner, replacements)

    def _replace_in_paragraphs(self, paragraphs, replacements: Dict[str, str]) -> None:
        for p in paragraphs:
            text = p.text
            replaced = False
            for k, v in replacements.items():
                if k in text:
                    text = text.replace(k, "" if v is None else str(v))
                    replaced = True
            if replaced:
                p.text = text  # por quê: captura placeholders quebrados em runs
