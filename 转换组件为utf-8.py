# 尝试用gbk读取，用utf-8保存
with open('国泰海通.txt', 'r', encoding='gbk', errors='ignore') as f:
    content = f.read()
with open('国泰海通.txt', 'w', encoding='utf-8') as f:
    f.write(content)
