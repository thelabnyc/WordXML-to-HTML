from lxml import etree
import tempfile
import subprocess
import zipfile
import sys
import os


# Key/Value Pair of mappings
STYLE_MAPPINGS = {
    'ChapTitlect': 'ChapterTitle',
    'PartTitlept': 'PartTitle',
    'TitlepageAuthorNameau': 'AuthorName',
    'TitlepageBookSubtitlestit': 'BookSubtitle',
    'TitlepageBookSubtitlestit': 'TitleofBook',
    'PageBreakpb': 'PageBreak',
    'Text-Standardtx': 'StandardText',
    'ListBulletbl': 'BulletedList',
    'ListNumnl': 'NumberedList',
    'Extractext': 'Excerpt',
}

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

VALUE_ATTRIB = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'


def remap_styles_in_xml(source_xml):
    etree.register_namespace(
        'w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    q = etree.fromstring(source_xml)
    for tag in q.iterfind(".//w:pStyle", q.nsmap):
        if VALUE_ATTRIB in tag.attrib:
            for key, value in STYLE_MAPPINGS.items():
                if tag.attrib[VALUE_ATTRIB] == value:
                    tag.attrib[VALUE_ATTRIB] = key
    new_xml = etree.tostring(q)
    return new_xml


def convert_manuscript(input_file):
    """ Convert a docx input_file to htmlbook  """
    saxon_exec = 'java -jar saxon9he.jar'
    path_to_script = 'wordtohtml.xsl'
    extension = os.path.splitext(input_file)[1]

    if extension in ('.docx', '.docm'):
        # get the xml file out of the docx zip archive
        with zipfile.ZipFile(input_file) as document:
            xml_content = document.read('word/document.xml')
        xml_content = remap_styles_in_xml(xml_content)
    elif extension in ('.xml'):
        with open(input_file, "r") as document:
            xml_content = input_file.read()
    else:
        raise Exception("Input can be only docx, docm, or xml file")
    tmp_file = tempfile.NamedTemporaryFile()
    tmp_file.write(xml_content)
    html = subprocess.check_output(
        saxon_exec.split() + [tmp_file.name, path_to_script])

    return html.decode('UTF-8')


if __name__ == "__main__":
    html = convert_manuscript(sys.argv[1])
    print(html)
