from chatgpt import ChatGPT
import json
from sel import seg
import time
import os
import argparse
import sys
title = "没有"
num = 10
def delete_file_if_exists(filename):
    """删除指定的文件，如果它存在"""
    if os.path.exists(filename):
        os.remove(filename)

def savetxt(txt, name):
    with open(name, "a") as file:
        file.write(txt + "\n")

if __name__ == "__main__":
    # 命令行参数处理
    parser = argparse.ArgumentParser(description="Generate a document based on a given title.")
    parser.add_argument("-title", default=title, help="Title for the document generation")
    args = parser.parse_args()
    title = args.title
    if title == "没有":
        print("请设置标题")
        sys.exit()

    # 删除指定的文件
    delete_file_if_exists('abstract_input.txt')
    delete_file_if_exists('text_input.txt')
    delete_file_if_exists('conclusion_input.txt')
    delete_file_if_exists('output.docx')

    with open('data.json', 'r', encoding='utf-8') as f:
        loaded_data = json.load(f)

    gpt = ChatGPT()
    gpt.start()
    gpt.send('请你充当一名专业学士，写一篇有关于<'+title+'>'+loaded_data[0])
    while not gpt.check_and_print_button_text():
        pass
    # savetxt(gpt.getLastMessage())

    gpt.send("给每一个子标题部分根据重要程度分配字数，然后以每700字为一块分成"+str(num)+"块，总字数在"+str(num*700)+loaded_data[1])

    while not gpt.check_and_print_button_text():
        pass

    content = gpt.getLastMessage()
    for i in range(num):
        if i % 3 == 2:
            sendtext = '接下来开始输出第'+str(i+1)+'块的内容。'+loaded_data[2]+content
        else :
            sendtext = '接下来开始输出第'+str(i+1)+'块的内容。'+loaded_data[2]
        gpt.send(sendtext)
        while not gpt.check_and_print_button_text():
            pass
        txt = gpt.getLastMessage()
        savetxt(txt,'text_input.txt')


    gpt.send('接下来根据上面的文章内容和<'+title+'>'+loaded_data[3])


    while not gpt.check_and_print_button_text():
        pass
    txt = gpt.getLastMessage()
    savetxt(txt,'abstract_input.txt')

    gpt.send(loaded_data[4])
    while not gpt.check_and_print_button_text():
        pass
    txt = gpt.getLastMessage()
    savetxt(txt,'conclusion_input.txt')
    # savetxt(gpt.getLastMessage())
    time.sleep(2)
    import toword
    toword.txt_to_word('abstract_input.txt', 'text_input.txt', 'conclusion_input.txt', 'output.docx',title)
    time.sleep(1)
    seg(title=title)
