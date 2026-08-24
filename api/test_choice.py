import zipfile
from lxml import etree
with zipfile.ZipFile('/tmp/tmpfmngpeqk/rebuilt_Đề ngoại tổng hợp - TTK.docx', 'r') as z:
    doc_xml = z.read('word/document.xml')
    tree = etree.fromstring(doc_xml)
    invalid_ac = 0
    for ac in tree.xpath('//mc:AlternateContent', namespaces=tree.nsmap):
        choices = ac.xpath('./mc:Choice', namespaces=tree.nsmap)
        if len(choices) != 1:
            invalid_ac += 1
            print('Invalid AlternateContent! Number of Choice elements:', len(choices))
    print('Total invalid:', invalid_ac)
