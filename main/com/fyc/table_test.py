import pandas as pd

df = pd.read_excel('C:\\Users\\Administrator\\Desktop\\write_docx\\data.xlsx', sheet_name='Sheet1',header=0)



data = df[[ 'jczb002', 'jczb007']].copy()
print(data.head(199))
data.rename(columns={'jczb002':'纳管表数量','jczb007':'技术元数据质量合格表数量'}, inplace=True)
index = data.sum()
index['技术元数据质量合格率'] = round(index['技术元数据质量合格表数量'] /index['纳管表数量']*100,2)
print(index)
# # join = index.join(index, how='inner')
# pd.merge(index, index, how='inner', on='0')