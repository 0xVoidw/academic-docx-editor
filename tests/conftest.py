from __future__ import annotations

import base64
import zipfile
from pathlib import Path

import pytest


PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9U9lnB8AAAAASUVORK5CYII=")


def create_docx(path: Path, rich: bool = False) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
    body = '''<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>研究标题 Title</w:t></w:r></w:p>
<w:p><w:r><w:t>本研究</w:t></w:r><w:r><w:t>观察到</w:t></w:r><w:r><w:t>增加。</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>加粗 Bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> 斜体 Italic</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>n</w:t></w:r></w:p>
<w:p><w:r><w:t>重复</w:t></w:r></w:p><w:p><w:r><w:t>重复</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p><w:r><w:t>表格数据 1</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>表 1 示例</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>方法 Methods</w:t></w:r></w:p>'''
    if rich:
        body += '''<w:p><w:hyperlink r:id="rIdHyper"><w:r><w:t>链接文本</w:t></w:r></w:hyperlink></w:p>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF bookmark1 </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
<w:p><m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath></w:p>
<w:p><w:bookmarkStart w:id="1" w:name="bookmark1"/><w:r><w:t>锚点</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>
<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Figure 1" descr="sample"/><a:graphic><a:graphicData><a:blip r:embed="rIdImage"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:r><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/><wp:docPr id="2" name="Floating"/><a:graphic><a:graphicData><a:blip r:embed="rIdFloat"/></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>图 1 示例</w:t></w:r></w:p>'''
    references = '<w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/>' if rich else ''
    body += f'<w:sectPr>{references}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
    files = {
        '[Content_Types].xml': '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        '_rels/.rels': '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/document.xml': f'<?xml version="1.0"?><w:document {ns}><w:body>{body}</w:body></w:document>',
        'word/styles.xml': f'<?xml version="1.0"?><w:styles {ns}><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style><w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="Caption"/></w:style></w:styles>',
    }
    if rich:
        files['[Content_Types].xml'] = files['[Content_Types].xml'].replace('</Types>', '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')
        files['word/_rels/document.xml.rels'] = '''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHyper" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org" TargetMode="External"/><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/><Relationship Id="rIdFloat" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>'''
        files['word/media/image1.png'] = PNG
        files['word/media/image2.png'] = PNG
        files['word/header1.xml'] = f'<?xml version="1.0"?><w:hdr {ns}><w:p><w:r><w:t>论文页眉</w:t></w:r></w:p></w:hdr>'
        files['word/footer1.xml'] = f'<?xml version="1.0"?><w:ftr {ns}><w:p><w:r><w:t>页脚</w:t></w:r></w:p></w:ftr>'
        files['docProps/core.xml'] = '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>独立测试论文</dc:title></cp:coreProperties>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content)
    return path


@pytest.fixture
def simple_docx(tmp_path):
    return create_docx(tmp_path / '中文路径' / '论文.docx')


@pytest.fixture
def rich_docx(tmp_path):
    return create_docx(tmp_path / '复杂文档.docx', rich=True)
