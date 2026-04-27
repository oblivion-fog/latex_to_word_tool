# latex_to_word_tool
本工具实现了批量处理Word文档中的所有LaTeX公式符号
1. 识别全文中所有LaTeX公式（$ $、$$ $$、\( \)、\[ \]包裹、零散数学符号）
2. 批量转换为Word原生可编辑公式（OMML格式），双击可修改
3. 不改动任何文字、段落、排版，仅转换公式
4. 保留所有数学细节：上下标、分式、根号、矩阵、求和、积分、希腊字母等
5. 直接返回处理好的完整文档内容，无多余说明


## 依赖安装
```bash
pip install python-docx lxml latex2mathml
```
## 使用方法：
```bash
# 基础用法（自动生成 input_converted.docx）
python latex_to_word_formula.py 你的文档.docx

# 指定输出文件名
python latex_to_word_formula.py 你的文档.docx -o 输出文档.docx
```
ps:
1. 在你要转换的word文档所在文件夹下打开终端,执行以上命令
2. 系统要装有python
> 怎么下载python 自行搜索
