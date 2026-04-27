from docx import Document
from docx.oxml.ns import qn
from lxml import etree

doc = Document('test_latex.docx')
print('=== 原始文档 run 结构分析 ===')
for i, para in enumerate(doc.paragraphs):
    xml = para._element
    wr = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r'
    wt = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'
    runs = xml.findall('.//' + wr)
    if runs:
        for j, r in enumerate(runs):
            t = r.find(wt)
            txt = t.text if t is not None else ''
            print(f'  para{i+1} run{j+1}: {repr(txt)}')
