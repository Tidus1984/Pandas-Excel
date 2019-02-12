import pandas as pd
#python3.6 andaconda pycharm
#学习视频 Python数据分析 - Pandas玩转Excel-Timothy 


#读取文件
df = pd.read_excel(r'C:\Temp\Books.xlsx',skiprows=3,header=2,usecols='C:F',dtype={'ID':object,'InStore':str,'Date':str},index_col='ID')
#skiprows 跳过前3行
#header 从第2行开始读
#usecols 列选择C到F列数据读取
#dtype 重点 pandas把NaN默认flode 如果想下面迭代表达先把空列设置成str类型、或者object
#index_col 把ID列作为DateFrame的index列


#保存文件
df.to_excel(path)
#如果未设置index列系统自动保存，excel打开后 多出一列index，解决方案指定index列
#方法：
df=df.set_index('ID')
df.set_index('ID'，inplace=True)


#单元格填空
df['ID'].at[0] = 100  #Serise 后 at
df.at[0,'ID'] = 100 #DateFrame 后 at
df.loc[0,'ID'] = 100
#ID 列 index 0行 赋值 100 
for i in books.index:
    books['ID'].at[i]=i+1 #i从0开始 单元格赋值从1、2、3开始，注意index
    books['InStore'].at[i]='Yes'if i%2==0 else 'No'


#排序多重排序
df.sort_values(by=['Worthy', 'Price'], ascending=[True, False], inplace=True)
#by 排序的columns
#ascending True 从小到大
#inplace True 直接在df数据上保留修改


#数据筛选、过滤 apply
def age_range(x):
    return 20<=x<=36
def score_range(y):
    return 80<=y<=100
df1=df.loc[df['Age'].apply(age_range)].loc[df['Score'].apply(score_range)]


#多表联合VLOOKUP合并
table = df1.merge(df2,how='left',on='ID').fillna('没找到')
table.Score = table.Score.astype(int)#Score列变成整数
#how=’left‘ 表示依 df1 基础 保留所有df1列信息。默认inner参数
#on=’ID‘ df1与df2都有ID列 前提两张表都有ID列，没有用 left_on与right_on
#.fillna() 表示 在df1中df2没有的数据填下’没找到‘
#merge 不能默认指定index列 必须 on指定
students = pd.read_excel('C:/Temp/Student_score.xlsx', sheet_name='Students', index_col='ID')
scores = pd.read_excel('C:/Temp/Student_score.xlsx', sheet_name='Scores', index_col='ID')
table = students.join(scores, how='left').fillna(0)
table.Score = table.Score.astype(int)
#join 必须指定index_col一样，也有on参数


#数据效验


#分列修改每列替换


#统计分析










