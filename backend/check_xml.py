import zipfile
from lxml import etree

with zipfile.ZipFile('/tmp/tmp6ugfrxh8/rebuilt_Đề ngoại tổng hợp - TTK.docx', 'r') as z:
    doc_xml = z.read('word/document.xml')
    tree = etree.fromstring(doc_xml)
    ns = tree.nsmap
    
    invalid_tc = 0
    for tc in tree.xpath('//w:tc', namespaces=ns):
        children = tc.xpath('./w:*', namespaces=ns)
        if not children or children[-1].tag != '{' + ns.get('w', '') + '}p':
            invalid_tc += 1
            print('Found invalid w:tc! Children tags:', [c.tag.split('}')[-1] for c in children])
    
    print('Invalid w:tc elements:', invalid_tc)
    
    # check empty body
    body = tree.xpath('//w:body', namespaces=ns)[0]
    children = body.xpath('./w:*', namespaces=ns)
    if not children or (children[-1].tag != '{' + ns.get('w', '') + '}p' and children[-1].tag != '{' + ns.get('w', '') + '}sectPr'):
        print('Invalid w:body!')

