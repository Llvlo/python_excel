import time

import pandas as pd
fn = r"C:\Users\Administrator\Desktop\钢材订单.xls"
if input(fn.title()+"\n是的话输入单击回车") != "":
    fn = input("请输入文件路径：")

# 从录单系统获取数据
df = pd.read_excel(fn, sheet_name='录单系统')
column = df[["品名", "规格", "重量", "车号"]]  # 需要找数据的列
record_os = column.values.tolist()
for i in range(len(record_os)):
    if record_os[i][0] == 'HRB400E螺纹钢': record_os[i][0] = '螺纹钢'
    if record_os[i][0] == 'HRB400E盘螺': record_os[i][0] = '热轧带肋钢筋'
    if record_os[i][0] == 'HPB300线材': record_os[i][0] = '热轧光圆钢筋'
    try:
        record_os[i][1] = record_os[i][1].strip()
        record_os[i][1] = record_os[i][1][1:]
    except:
        pass

# 从下单系统获取数据
df = pd.read_excel(fn, sheet_name='下单系统')
column = df[["品名", "规格", "出库重量", "配送车号", "创建时间"]]
order_os = column.values.tolist()
for i in range(len(order_os)):  # 修改一下表的数据φ14X12与Ф25X12会对应不上所以去掉φ和Ф
    order_os[i][1] = order_os[i][1].strip()
    order_os[i][1] = order_os[i][1][1:]

find = []
no_find = []
count_re = 2
count_or = 2
all_exist_ordinal = []

# 寻找相同的数据
for i in record_os:
    for j in order_os:
        j = j[:4]
        if i == j:
            find.append(['在录单系统第', count_re, '行:', i, '在下单系统第', count_or, '行:', j])
            all_exist_ordinal.append(count_or)  # 把两个系统都存在的序号记下来
            break
        count_or += 1  # 下单系统的下标加1
    else:
        no_find.append(f'在录单系统的第{count_re}行：{i}')
    count_or = 2  # 每次循环对于or来说都需要重新找一遍所以归0
    count_re += 1  # 录单系统的下标加1

order_ordinal = list(range(2, len(order_os) + 2))  # 因为表的数据第一行是2所以从2开始，因为是从2开始所以末尾得+2
only_exist_or = list((set(order_ordinal) - set(all_exist_ordinal)))  # 找出仅存在于下单系统的那部分序号

print(f'在录单系统中总共有{len(record_os)}条记录')
print(f'在下单系统中总共有{len(order_os)}条记录')

print(f'异常记录{len(record_os)-len(find)}条。')


print('\n以下数据仅出现在下单系统:')
for i in sorted(only_exist_or):
    if order_os[i - 2][:4] not in record_os:
        print(f'在下单系统第{i}行:{order_os[i - 2]}')

print('\n错误信息,仅出现在录单系统:')
for i in no_find:
    print(i)

print()
# 集合处理重复的数据
t_record = set()
for i in range(len(record_os)):
    if str(record_os[i]) not in t_record:
        t_record.add(str(record_os[i]))
    else:
        print(record_os[i])
        print(f"在录单系统中的第{i+2}条与{record_os[:i].index(record_os[i])+2}条\n")

time.sleep(0.5)
if input('\n\n输入1看所有信息') == '1':
    for i in find:
        print(i[:4])
        print(i[4:])
        print()

time.sleep(1000000)
