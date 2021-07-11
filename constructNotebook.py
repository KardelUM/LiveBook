import zipfile
import xml.etree.ElementTree as ET
from io import StringIO
import nbformat as nbf


def add_rStyle(content, rStyle):
    if content == "":
        return content
    if 'b' in rStyle:
        content = "**" + content + "**"
    elif 'i' in rStyle:
        content = "*" + content + "*"
    elif "font-monospace" in rStyle:
        content = '`' + content + '`'
    return content + ' '


def nstag(tag, ns):
    _ns, _tag = tag.split(":")
    return "{" + ns[_ns] + "}" + _tag

def add_formula(content, style):
    if style == 'equation':
        return "$" + content + "$ "

def add_block(content, style):
    if style == "title":
        return nbf.v4.new_markdown_cell("# " + content)
    elif style == "text":
        return nbf.v4.new_markdown_cell(content)
    elif style == "code":
        return nbf.v4.new_code_cell(content)
    elif style == "heading":
        return nbf.v4.new_markdown_cell("## " + content)


def build_notebook(source_livescript, dst_notebook):
    nb = nbf.v4.new_notebook()
    nb_blocks = []
    document = zipfile.ZipFile(source_livescript).read("matlab/document.xml").decode()
    ns = dict([node for _, node in ET.iterparse(StringIO(document), events=["start-ns"])])
    root = ET.fromstring(document)

    ls_blocks = root.find("w:body", ns).findall("w:p", ns)
    for ls_block in ls_blocks:
        block = Block()
        block.gainMdFromBlock(ls_block, ns)
        cb = block.generate_md()
        if cb is not None:
            nb_blocks.append(cb)
    nb["cells"] = nb_blocks
    nbf.write(nb, dst_notebook)


class Block:
    def __init__(self):
        self.style = ""
        self.content = ""

    def gainMdFromBlock(self, ls_block, ns):
        for element in list(ls_block):
            if element.tag == nstag("w:pPr", ns):
                if element.find("w:sectPr", ns) is not None:
                    continue
                pStyle = element.find("w:pStyle", ns)
                self.style = pStyle.attrib["{w}val".format(w="{" + ns["w"] + "}")]
            elif element.tag == nstag("mc:AlternateContent", ns):
                pStyle = element.find("mc:Fallback", ns).find("w:pPr", ns).find("w:pStyle", ns)
                self.style = pStyle.attrib["{w}val".format(w="{" + ns["w"] + "}")]
            elif element.tag == nstag("w:r", ns):
                rPr = element.find("w:rPr", ns)
                rStyle = set()
                if rPr is not None:
                    if rPr.find("w:b", ns) is not None:
                        rStyle.add("b")
                    if rPr.find("w:i", ns) is not None:
                        rStyle.add("i")
                    if rPr.find("w:rFonts", ns) is not None:
                        cs = rPr.find("w:rFonts", ns).attrib["{w}cs".format(w="{" + ns["w"] + "}")]
                        if cs == "monospace":
                            rStyle.add("font-monospace")
                self.content += add_rStyle(element.find("w:t", ns).text.strip(), rStyle)
            elif element.tag == nstag("w:customXml", ns):
                if element.attrib["{w}element".format(w="{" + ns["w"] + "}")] == "equation":
                    style = "equation"
                try:
                    self.content += add_formula(element.find("w:r", ns).find("w:t", ns).text.strip(), style)
                except AttributeError:
                    pass

    def generate_md(self):
        return add_block(self.content, self.style)
if __name__ == '__main__':
    build_notebook("data/Lecture_2.mlx", "Lecture2.ipynb")
