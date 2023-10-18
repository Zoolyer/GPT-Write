from chatgpt import ChatGPT
def savetxt(txt):
    with open("output.txt", "a") as file:
        file.write(txt + "\n")
def contains_stop_tag(s):
    return '</stop>' in s

if __name__ == "__main__":

    gpt = ChatGPT()
    gpt.start()
    gpt.send("你好")
    while not gpt.check_and_print_button_text():
        pass
    savetxt(gpt.getLastMessage())
    gpt.send("你是谁")
    while not gpt.check_and_print_button_text():
        pass
    savetxt(gpt.getLastMessage())