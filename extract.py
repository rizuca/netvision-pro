import zipfile
import xml.etree.ElementTree as ET

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def get_text(path):
    try:
        with zipfile.ZipFile(path) as docx:
            tree = ET.XML(docx.read('word/document.xml'))
            text = []
            for p in tree.iter(WORD_NAMESPACE+'p'):
                texts = [node.text for node in p.iter(WORD_NAMESPACE+'t') if node.text]
                if texts:
                    text.append(''.join(texts))
            return '\n'.join(text)
    except Exception as e:
        return str(e)

open('jurnal_text.txt', 'w', encoding='utf-8').write(get_text(r'd:\jarkom\Jurnal ATASI-Jarkom (1).docx'))
open('template_text.txt', 'w', encoding='utf-8').write(get_text(r'd:\jarkom\Template+ATASI+2026.docx'))
print("Extraction complete")
