from docx import Document
from lxml import etree
import re

doc = Document('test_output2.docx')
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

latex_re = re.compile(r'(\$\$.+?\$\$|\\\[.+?\\\]|\\\(.+?\\\)|\$[^$]+\$)', re.DOTALL)

total_omml = 0
total_remain = 0

print("=== 转换结果验证 ===\n")
for i, para in enumerate(doc.paragraphs):
    xml = para._element
    omath_list = xml.findall(f'.//{{{M}}}oMath')
    t_texts = [t.text or "" for t in xml.findall(f'.//{{{W}}}t')]
    all_text = "".join(t_texts)
    remaining = latex_re.findall(all_text)
    
    if omath_list or remaining:
        status = "[OK]" if not remaining else "[WARN-残留LaTeX]"
        print(f"  段落{i+1:2d} {status}: {len(omath_list)} 个OMML公式  文字='{all_text[:60]}'")
        if remaining:
            print(f"           残留: {remaining}")
        total_omml += len(omath_list)
        total_remain += len(remaining)

print(f"\n总计: OMML公式={total_omml}  残留LaTeX={total_remain}")
print("所有公式均已转换为 Word 原生可编辑公式！" if total_remain == 0 else "仍有公式未能转换，请检查！")
