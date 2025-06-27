from src.main.com.fyc.excel2word import excel2word

excel_url = 'C:\\Users\\Administrator\\Desktop\\write_docx\\data.xlsx'
word_url = 'C:\\Users\\Administrator\\Desktop\\write_docx\\模版.docx'
write_url = 'C:\\Users\\Administrator\\Desktop\\write_docx\\result.docx'

write=excel2word(excel_url, word_url,write_url)
write.run()
# write.sout_dict()
# write.write()
