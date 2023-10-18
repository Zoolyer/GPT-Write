import json

data = ["第一个元素的内容\n可能包含多行\n如此。",
        "第二个元素，同样可能\n包含多行内容。",
        "第三个元素。"]

with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=4)
