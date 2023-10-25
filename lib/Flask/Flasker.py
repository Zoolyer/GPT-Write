import logging
from flask import Flask, request, jsonify, send_from_directory
from lib.GPT.chatgpt import ChatGPT
from lib.GPT.write import write
import math
import threading
import os
import pythoncom

import datetime

# Setting up logging
log_dir = "log"

log_file_path = "./log/requests_log.txt"  # Define log file path

import time


if not os.path.exists(log_dir):
    os.makedirs(log_dir)

logging.basicConfig(filename=os.path.join(log_dir, 'errors.log'), level=logging.ERROR)

app = Flask(__name__)

title = "没有"
thread_status = 0  # 0: 空, 1: 正在运行, -1: 执行错误, 2: 成功执行但还没下载文件
paper = [0,0]
paper[0] = 0
paper[1] = 0



def shutdown_check():
    shutdown_counter = 0
    while True:
        time.sleep(10)  # check every 10 seconds
        if thread_status == 0:
            shutdown_counter += 10
            if shutdown_counter >= 20*60:  # 5 minutes
                logging.error(
                    f"System shutdown")
                os.system("shutdown /s /t 1")  # Shutdown the system (Windows command)
                break
        else:
            shutdown_counter = 0  # reset counter if status changes



def log_request(num_value):
    """
    Log the request to a file with the current timestamp and the calculated num.
    """
    with open(log_file_path, "a") as log_file:  # Use the defined path
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_file.write(f"{timestamp},{num_value + 4}\n")


def get_requests_in_last_duration(duration_hours=3, duration_minutes=20):
    """
    Get the total num of requests in the last specified duration.
    """
    threshold_time = datetime.datetime.now() - datetime.timedelta(hours=duration_hours, minutes=duration_minutes)
    total_num = 0

    # Ensure the directory exists, if not, create it
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # Check if the file exists, if not, return 0
    if not os.path.exists(log_file_path):
        return total_num

    with open(log_file_path, "r") as log_file:  # Use the defined path
        lines = log_file.readlines()
        for line in reversed(lines):  # Start checking from the most recent
            timestamp_str, num_str = line.strip().split(',')
            timestamp = datetime.datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')

            if timestamp < threshold_time:
                break
            total_num += int(num_str)

    return total_num


def valid_gpt_value(value):
    if value not in ["3.5", "4"]:
        return False
    return True

def threaded_write(title_value, num, gpt_value):
    global thread_status
    try:
        pythoncom.CoInitialize()  # Initialize the COM library
        write(title_value, num, gpt_value,paper)
        thread_status = 2  # 标记执行成功但文件尚未下载
    except TimeoutError as te:
        logging.error(
            f"Timeout error occurred with state: {str(te)}")  # Logging the timeout error and its state to the file
        thread_status = -1  # 标记执行错误
    except Exception as e:
        logging.error("An error occurred: " + str(e))  # Logging other errors to the file
        thread_status = -1  # 标记执行错误

@app.route('/write', methods=['POST'])
def generate_document():
    global thread_status

    # Check if the limit is exceeded
    if thread_status == 1:
        return jsonify({"message": "Another task is currently running. Please wait."}), 429

    title_value = request.form.get("title", title)
    num_value = request.form.get("num")
    gpt_value = request.form.get("gpt")


    if not all([title_value, num_value, gpt_value]):
        return jsonify({"error": "Missing required parameters!"}), 400

    if not valid_gpt_value(gpt_value):
        return jsonify({"error": "Invalid GPT value. It should be '3.5' or '4'."}), 400

    num = max(2, math.ceil(int(num_value) / 700))

    if not num:
        log_request(0)  # Log with default value
    else:
        if gpt_value == "4":
            log_request(int(num))
    try:
        thread_status = 1  # 标记线程为正在运行
        threading.Thread(target=threaded_write, args=(title_value, num, gpt_value)).start()
        download_link = f"/download?title={title_value}"
        return jsonify({"message": "Document generation started!", "download_link": download_link}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

from lib.GPT.sel import seg

@app.route('/test', methods=['GET'])
def test():
    try:
        pythoncom.CoInitialize()  # Initialize the COM library
        seg("test")
    except Exception as e:
        logging.error("An error occurred: " + str(e))  # Logging the error to the file
        return "Error", 500

    return "OK", 200


@app.route('/download', methods=['GET'])
def download_file():
    global thread_status
    title_value = request.args.get("title", title)

    # 获取当前Python脚本的绝对目录
    base_directory = os.path.dirname(os.path.abspath(__file__))
    base_directory = os.path.dirname(base_directory)  # 获取父级目录
    base_directory = os.path.dirname(base_directory)  # 获取父级目录
    # 生成完整的文件路径
    filepath = os.path.join(base_directory, "output", f"{title_value}.docx")

    print(filepath)
    if os.path.exists(filepath):
        thread_status = 0  # Reset the thread status after downloading the file
        return send_from_directory(directory=os.path.join(base_directory, "output"), path=f"{title_value}.docx",
                                   as_attachment=True)
    else:
        return jsonify({"error": "File not found!"}), 404


@app.route('/status', methods=['GET'])
def get_status():
    global thread_status
    return jsonify({"status": thread_status})
@app.route('/paper', methods=['GET'])
def get_paper():
    return jsonify({"status": paper[0], "time": paper[1]})

@app.route('/request_count', methods=['GET'])
def get_request_count():
    """
    Get the total number of requests in the last 3 hours and 20 minutes.
    """
    total_requests = get_requests_in_last_duration()
    return jsonify({"request_count": total_requests})


def run():
    # Start the shutdown_check thread
    shutdown_thread = threading.Thread(target=shutdown_check)
    shutdown_thread.start()

    app.run(host='0.0.0.0', port=5000)
