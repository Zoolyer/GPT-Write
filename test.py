import lib.GPT.write as write
import lib.GPT.sel as sel
import lib.GPT.toword as tw

if __name__ == "__main__":
    paper = [0,0]
    # write(title_value, num, gpt_value, paper)
    write.write("第三次世界大战的武器装备",20,"4",paper)
    # tw.txt_to_word('./output/debug/abstract_input.txt', './output/debug/text_input.txt', './output/debug/conclusion_input.txt', './output/debug/output.docx',"我是标题")
