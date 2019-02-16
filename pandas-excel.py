import pandas as pd
#python3.6  pycharm 编辑器
#andaconda 可以用 prompt 命令行：jupyter notebook 启动idle 交互界面
#学习视频 Python数据分析 - Pandas玩转Excel-Timothy 


#读取文件
df = pd.read_excel(r'C:\Temp\Books.xlsx',skiprows=3,header=2,usecols='C:F',dtype={'ID':object,'InStore':str,'Date':str},index_col='ID')
#skiprows 跳过前3行
#header 从第2行开始读
#usecols 列选择C到F列数据读取
#dtype 重点 pandas把NaN默认flode 如果想下面迭代表达先把空列设置成str类型、或者object
#index_col 把ID列作为DateFrame的index列
#sheet_name=‘sheet1’ 把Excel表中读取sheet1 或者sheet2
df = pd.read_csv(path,encoding='gb18030',header=1) #读取文件endcoding注意


#保存文件
df.to_excel(path)
df.to_csv('5.csv',encoding='gb18030')
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
time=df.创建时间.str.split(expand=True)
#str.split()空的表示默认空格分
#expand list分为全咧
#columns 0 1

#统计分析


#数据查看
df.info()#查看每列数据数据类型 str time object等等
df.head(3)#查看头部前三行
df.tail()#查看尾部
df.创建时间.unique()#查看 创建时间 列 不重复值有多少


#分类汇总
#方法1：
import numpy as np
from datetime import date
orders = pd.read_excel('C:/Temp/Orders.xlsx', dtype={'Date': date})#让Date列变成日期类型？始终有问题
orders['Year'] = pd.DatetimeIndex(orders.Date).year
#新增Year列 让Date 2019/5-31 在Year列中显示年份year
pt1 = orders.pivot_table(index='Category', columns='Year', values='Total', aggfunc=np.sum)
#生成数据透视表 row是Category，columns是Year的种类，求值是Total列的数值，目的是aggfunc 求和

#方法2：
groups = orders.groupby(['Category', 'Year'])
#根据Category、Year列分组
s = groups['Total'].sum()
#Total列求和
c = groups['ID'].count()
#ID列求数量
pt2 = pd.DataFrame({'Sum': s, 'Count': c})
#s、c 两个DataFrame 按照列Sum 与 Count 再合并成一个新DateFrame


#行操作集锦
page_001 = pd.read_excel('C:/Temp/Students.xlsx', sheet_name='Page_001')
page_002 = pd.read_excel('C:/Temp/Students.xlsx', sheet_name='Page_002')
# 追加已有：append row 添加002、重新排列index并且放弃原来的两个index
students = page_001.append(page_002).reset_index(drop=True)
# 追加新建：stu新建 row 放到students后面并且忽略index不然报错
stu = pd.Series({'ID': 41, 'Name': 'Abel', 'Score': 90})
students = students.append(stu, ignore_index=True)
# 删除（可切片）：students放弃index39到40 row 并且赋值给 students
students = students.drop(index=[39, 40])
# 插入（切片操作）
stu = pd.Series({'ID': 100, 'Name': 'Bailey', 'Score': 100})
part1 = students[:21]  # .iloc[] is the same
part2 = students[21:]
students = part1.append(stu, ignore_index=True).append(part2).reset_index(drop=True)
# 更改覆盖
stu = pd.Series({'ID': 101, 'Name': 'Danni', 'Score': 101})
students.iloc[39] = stu
# 设置空值
for i in range(5, 15):
    students['Name'].at[i] = ''
# 去掉空值
missing = students.loc[students['Name'] == '']
students.drop(missing.index, inplace=True)


#列操作集锦

# 选择 单据类型列 含有 销售发票 字符的行
df=df[df.单据类型.str.contains('销售发票')]

# 选择 分钟列 大于0 且 小于等于 15。然后给 分钟这一列负值15 int
df.loc[(df.分钟>0)&(df.分钟<=15),'分钟']=15
df.loc[(df.分钟>0)&(df.小时<=15),['分钟','小时']=[15,30]#想一想



#时间轴处理集锦 dt类（有空看一下）
#设置 创建时间 列 数据类型为 datetime64[ns]
df['创建时间'] = pd.to_datetime(df.创建时间)

#查看datetime类型
df['创建时间'].dt.year.unique()
df['创建时间']].dt.month.unique()
df['创建时间']].dt.day.unique()

#查看day=26天的数据有多少
df[df['创建时间'].dt.day == 26]

#查看第6行 创建时间列 与 第200行 创建时间列时间差距
df.loc[6,'创建时间'] - df.loc[200,'创建时间']

#创建新 开票星期几 行 然后 把创建时间列算下星期几或者时间
df['开票星期几'] = pd.DatetimeIndex(df.创建时间).dayofweek


#str方法 有空看一下

#合并两列用 - 区分
df['小时分钟']=df['小时'].str.cat(df.['分钟']，sep='-')






