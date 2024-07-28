filename = r"C:\Users\iba\Desktop\excel\example.txt"

with open(filename, 'r', encoding='utf-8') as file:
    lines = file.readlines()

p1, p2, p3 = [i.strip().split("=")[1] for i in lines]
print(p1, p2, p3)
# 打印每一行
# for line in lines:
#     print(line.strip())  # `strip()` 去掉每行末尾的换行符
