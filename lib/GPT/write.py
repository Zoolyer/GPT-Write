import re

from lib.GPT.chatgpt import ChatGPT
import json
from lib.GPT.sel import seg
import time
import os
import sys
import logging

# Setting up logging
log_dir = "log"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)

logging.basicConfig(filename=os.path.join(log_dir, 'execution.log'), level=logging.INFO)

def delete_file_if_exists(filename):
    """删除指定的文件，如果它存在"""
    if os.path.exists(filename):
        os.remove(filename)

def savetxt(txt, name):
    with open(name, "a", encoding='utf-8') as file:
        file.write(txt + "\n")


def trim_txt_content(txt, keywords=None):
    # 删除所有<t之前的部分
    start_pos = txt.find('<t')
    if start_pos != -1:
        txt = txt[start_pos:]

    # 查找最后一个换行符的位置
    if keywords is None:
        keywords = ["这一部分","下一块","以上内容", "部分内容", "此块内容", "块内容", "以上", "论文", "此部分","继续"]
    last_newline_pos = txt.rfind('\n')
    if last_newline_pos == -1:
        return txt  # 如果txt中没有换行符，直接返回原txt

    # 获取从最后一个换行符到txt末尾的内容
    last_segment = txt[last_newline_pos + 1:]

    # 检查是否有形如 "(%d块)" 或 "(块%d)" 的模式
    if re.search(r'\d+块', last_segment) or re.search(r'块\d+', last_segment):
        txt = txt[:last_newline_pos]
        return txt

    # 检查最后一段是否以 "(" 开头，并以 ")" 结尾
    if last_segment.startswith("（") and last_segment.endswith("）"):
        txt = txt[:last_newline_pos]
        return txt
    # 检查最后一段是否以 "(" 开头，并以 ")" 结尾
    if last_segment.startswith("(") and last_segment.endswith(")"):
        txt = txt[:last_newline_pos]
        return txt

    # 遍历关键词列表，判断该段内容中是否含有任何一个关键词
    for keyword in keywords:
        if keyword in last_segment:
            # 如果含有关键词，删除从该换行符到txt末尾的内容
            txt = txt[:last_newline_pos]
            break  # 找到一个关键词后即跳出循环，不再检查其他关键词

    return txt


import time
import sys
import json


# 请确保其他需要的库也已导入

def write(title, num, gpt_num,paper):
    paper[0] = 0  # 设置状态变量
    # 命令行参数处理
    if title == "没有":
        print("请设置标题")
        sys.exit()

    # 删除指定的文件
    delete_file_if_exists('./output/debug/abstract_input.txt')
    delete_file_if_exists('./output/debug/text_input.txt')
    delete_file_if_exists('./output/debug/conclusion_input.txt')
    delete_file_if_exists('./output/debug/output.docx')

    with open('./config/data.json', 'r', encoding='utf-8') as f:
        loaded_data = json.load(f)

    gpt = ChatGPT(gpt_num)
    gpt.start()


    def check_with_timeout(timeout=180):  # 设置3分钟的超时
        start_time = time.time()
        while not gpt.check_and_print_button_text():
            paper[1] = time.time() - start_time
            if time.time() - start_time > timeout:
                raise TimeoutError(f"Operation timed out after {timeout} seconds. State: {state}")
            time.sleep(1)

    paper[0] += 1
    gpt.send('请你充当一名专业学士，写一篇有关于<' + title + '>' + loaded_data[0])
    check_with_timeout()

    gpt.send(
        "给每一个子标题部分根据重要程度分配字数，然后以每700字为一块分成" + str(num) + "块，总字数在" + str(num * 700) +
        loaded_data[1])
    check_with_timeout()
    paper[0] += 1

    content = gpt.getLastMessage()
    for i in range(num):
        sendtext = '接下来开始输出第' + str(i + 1) + '块的内容，700字左右。' + loaded_data[2] + content

        paper[0] += 1
        gpt.send(sendtext)
        check_with_timeout()
        txt = gpt.getLastMessage()
        txt = trim_txt_content(txt)
        savetxt(txt, './output/debug/text_input.txt')

    paper[0] += 1
    gpt.send('接下来根据上面的文章内容和<' + title + '>' + loaded_data[3])
    check_with_timeout()
    txt = gpt.getLastMessage()
    savetxt(txt, './output/debug/abstract_input.txt')

    paper[0] += 1
    gpt.send(loaded_data[4])
    check_with_timeout()

    txt = gpt.getLastMessage()
    savetxt(txt, './output/debug/conclusion_input.txt')

    time.sleep(2)
    gpt.close_browser()
    time.sleep(1)

    import lib.GPT.toword as toword
    toword.txt_to_word('./output/debug/abstract_input.txt', './output/debug/text_input.txt',
                       './output/debug/conclusion_input.txt', './output/debug/output.docx', title)
    time.sleep(1)
    seg(title=title)
    logging.info(f"Execution completed successfully for title: {title} with num: {num} and gpt_num: {gpt_num}")

