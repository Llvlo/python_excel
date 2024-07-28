import pandas as pd
import time
fn = r"C:\Users\iba\Desktop\3.12\销售表.xlsx"

# 整理一下列
column_1 = "平沙、南水"
column_2 = "小林、三灶、红旗"
column_3 = "香洲（不含唐家/横琴）、斗门、南屏、前山、拱北、湾仔"
column_4 = "横琴、唐家湾、金鼎"


# 处理excel里面的数据
sale = pd.read_excel(fn, sheet_name="销售表")[["规格", "品名", "收货地点", "结算单价", "客户", "日期"]]
quote = pd.read_excel(fn, header=2, sheet_name="报价表")[["规格", "品名", column_1, column_2, column_3, column_4]]


# 因为（不含唐家/横琴） 这里容易导致错误筛选所以不要了
quote = quote.rename(columns={"香洲（不含唐家/横琴）、斗门、南屏、前山、拱北、湾仔": "香洲、斗门、南屏、前山、拱北、湾仔"})
column_3 = "香洲、斗门、南屏、前山、拱北、湾仔"


# 因为这两家公司是有优惠的, 所以就在这里把价格调好
sale.loc[sale["客户"] == "上海找钢网信息科技股份有限公司", "结算单价"] += 3  # loc[boolean, 需要操作的列]
sale.loc[sale["客户"] == "上海钢银电子商务股份有限公司", "结算单价"] += 2  # +=2相当于对符合条件的所有元素+2

# 展示提取结果
print("请检查信息是否准确:")
print(f"销售表共有{sale.shape[0] + 1}行")  # 数据前面有一行所以+1
print(f"报价表共有{quote.shape[0] + 3}行\n")  # 数据前面有三行所以+3

# 通过销售表的规格、品名列定位报价表的行
# merge整合相同的数据
matched_df = pd.merge(sale.reset_index(), quote.reset_index(), on=["规格", "品名"], suffixes=('_sale', '_quote'))

# 通过sale的收货地点定位quote的列
# 通过quote的列拿来对比sale的结算单价
matched_df = matched_df.rename(columns={"index_sale": "销售表", "index_quote": "报价表"})  # 改名
matched_df["销售表"] += 2  # 第一个索引在0但实际上在表中的2
matched_df["报价表"] += 4  # 第一个索引在0但实际上在表中的4

columns = [column_1, column_2, column_3, column_4]
for _, row in matched_df.iterrows():
    zone = row["收货地点"][2:]
    for column in columns:
        if zone in column.split("、") and row[column] != row["结算单价"]:
            data = list(row[["日期", "客户", "规格", "品名", "收货地点"]])
            data[0] = str(data[0])[:10]  # 日期改一下容易看一点

            #  把价格改回来
            if row["客户"] == "上海找钢网信息科技股份有限公司": row["结算单价"] -= 3
            if row["客户"] == "上海钢银电子商务股份有限公司": row["结算单价"] -= 2

            print(f'报价表{row["报价表"]}:{row[column]}\n销售表{row["销售表"]}:{row["结算单价"]}')
            print(data, '\n')

time.sleep(1000000)