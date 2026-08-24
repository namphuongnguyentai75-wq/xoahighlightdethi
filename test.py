import re
text='A. Truy?n HCL B. Truy?n Fe C. Dùng'
res=re.split(r'(?=(?:Câu|Bài|CÂU|BÀI)\s+\d+[:\.])|(?=\b[A-Da-d]\.(?:\s|$))', text)
# write to file so we can read it without print encoding error
with open('res.txt', 'w', encoding='utf-8') as f:
    f.write(repr(res))
