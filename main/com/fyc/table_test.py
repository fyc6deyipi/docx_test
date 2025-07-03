import pandas as pd

# 创建示例数据集
data = {'学生姓名': ['Alice', 'Bob', 'Charlie', 'David', 'Eva'],
        '分数': [85, 70, 95, 60, 75]}

df = pd.DataFrame(data)

# 定义条件和对应的值
conditions = [df['分数'] >= 90, (df['分数'] >= 80) & (df['分数'] < 90), df['分数'] < 80]
values = ['优秀', '良好', '及格']

# 使用 case_when() 方法创建新列
df['等级'] = df['分数'].case_when(conditions, values, default='不及格')

# 输出结果
print(df)

