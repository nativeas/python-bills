import pandas as pd

# 读取第一个 Excel 文件，并设置 "快递单号" 列的 dtype 为 str
df1 = pd.read_excel('data.xlsx', dtype={'快递单号': str})

# 读取第二个 Excel 文件，并设置 "快递单号" 列的 dtype 为 str
df2 = pd.read_excel('data2.xlsx', dtype={'快递单号': str})

# 从 df1 获取 "快递单号" 和对应的 "商家/店铺"
mapping = df1.set_index('快递单号')['商家/店铺'].to_dict()

# 遍历 df2 的每一行，找到对应的 "商家/店铺" 数据并打印出来
for index, row in df2.iterrows():
    kuaidi_no = row['快递单号']
    shop = mapping.get(kuaidi_no, '未找到')
    print(f"快递单号: {kuaidi_no}, 商家/店铺: {shop}")
    df2.at[index, 'Q'] = shop

# 打印 df2 的前三行以查看结果
print(df2.head(3))

# 将 df2 保存到新的 Excel 文件 data3.xlsx
df2.to_excel('data3.xlsx', index=False)