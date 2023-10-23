import lib.GPT.write as write
import lib.GPT.sel as sel
import lib.GPT.toword as tw

if __name__ == "__main__":
    paper = [0,0]
    # write.write("123",2,"3.5",paper)
    tw.txt_to_word('./output/debug/abstract_input.txt', './output/debug/text_input.txt', './output/debug/conclusion_input.txt', './output/debug/output.docx',"我是标题")
