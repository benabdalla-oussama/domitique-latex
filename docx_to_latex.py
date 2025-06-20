import sys
import zipfile
import xml.etree.ElementTree as ET

NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

LATEX_REPLACEMENTS = {
    '&': r'\&',
    '%': r'\%',
    '$': r'\$',
    '#': r'\#',
    '_': r'\_',
    '{': r'\{',
    '}': r'\}',
    '~': r'\textasciitilde{}',
    '^': r'\^{}',
    '\\': r'\textbackslash{}',
}

def escape_latex(text):
    for k, v in LATEX_REPLACEMENTS.items():
        text = text.replace(k, v)
    return text

def extract_paragraphs(docx_path):
    with zipfile.ZipFile(docx_path) as docx:
        with docx.open('word/document.xml') as f:
            tree = ET.parse(f)
    paragraphs = []
    for p in tree.findall('.//w:p', NAMESPACE):
        texts = [t.text for t in p.findall('.//w:t', NAMESPACE) if t.text]
        if texts:
            paragraphs.append(''.join(texts))
    return paragraphs

def docx_to_latex(docx_path, tex_path):
    paragraphs = extract_paragraphs(docx_path)
    with open(tex_path, 'w', encoding='utf-8') as f:
        f.write("\\documentclass{article}\n\\begin{document}\n")
        for para in paragraphs:
            f.write(escape_latex(para) + "\n\n")
        f.write("\\end{document}\n")

if __name__ == '__main__':
    docx_path = sys.argv[1] if len(sys.argv) > 1 else 'rapport.docx'
    tex_path = sys.argv[2] if len(sys.argv) > 2 else 'rapport.tex'
    docx_to_latex(docx_path, tex_path)
