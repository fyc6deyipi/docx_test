from docxtpl import DocxTemplate,InlineImage



def write_docx(context):
    # doc = DocxTemplate('C:\\Users\\ycf\\Desktop\\write_docx\\模版.docx')
    doc = DocxTemplate('C:\\Users\\Administrator\\Desktop\\write_docx\\模版.docx')
    # context = {'name': 'dhsahiouhdioaujs'}
    doc.render(context)
    # doc.save('C:\\Users\\ycf\\Desktop\\write_docx\\模版1.docx')
    doc.save('C:\\Users\\Administrator\\Desktop\\write_docx\\result.docx')

