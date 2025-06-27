import pandas as pd

df = pd.read_excel('C:\\Users\\Administrator\\Desktop\\write_docx\\data.xlsx', sheet_name='Sheet1',header=0)




index = df[['systemname', 'jczb002', 'jczb007']].sum()
index['aa'] = index['jczb007'] /index['jczb002']
print(index)