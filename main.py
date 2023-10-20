from lib.GPT.chatgpt import ChatGPT
import argparse
from lib.GPT.write import write
import math

title = "没有"

def valid_gpt_value(value):
    if value not in ["3.5", "4"]:
        raise argparse.ArgumentTypeError("Invalid GPT value. It should be '3.5' or '4'.")
    return value

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate a document based on a given title.")
    parser.add_argument("-title", default=title, help="Title for the document generation")
    parser.add_argument("-num", type=int, required=True, help="Number to be processed for 'write'")
    parser.add_argument("-gpt", required=True, type=valid_gpt_value, help="GPT version for 'write' (either '3.5' or '4')")

    args = parser.parse_args()
    title = args.title

    # 把num/700向上取整的结果，但至少为2
    num = max(2, math.ceil(args.num / 700))
    gpt = args.gpt

    write(title, num, gpt)
