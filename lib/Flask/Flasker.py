import logging
from flask import Flask, request, jsonify, send_from_directory
from lib.GPT.chatgpt import ChatGPT
from lib.GPT.write import write
import math
import threading
import os

import datetime

# Setting up logging
log_dir = "log"

log_file_path = "./log/requests_log.txt"  # Define log file path

if not os.path.exists(log_dir):
    os.makedirs(log_dir)

logging.basicConfig(filename=os.path.join(log_dir, 'errors.log'), level=logging.ERROR)

app = Flask(__name__)

title = "没有"
thread_status = 0  # 0: 空, 1: 正在运行, -1: 执行错误, 2: 成功执行但还没下载文件


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
        write(title_value, num, gpt_value)
        thread_status = 2  # 标记执行成功但文件尚未下载
    except Exception as e:
        logging.error("An error occurred: " + str(e))  # Logging the error to the file
        thread_status = -1  # 标记执行错误


@app.route('/write', methods=['POST'])
def generate_document():
    global thread_status


    # Check if the limit is exceeded
    if get_requests_in_last_duration() > 50:
        return jsonify({"error": "Request limit exceeded!"}), 429
    if thread_status == 1:
        return jsonify({"message": "Another task is currently running. Please wait."}), 429

    data = request.json
    if not data:
        return jsonify({"error": "No data provided!"}), 400

    title_value = data.get("title", title)
    num_value = data.get("num")
    gpt_value = data.get("gpt")

    log_request(num_value)

    if not all([title_value, num_value, gpt_value]):
        return jsonify({"error": "Missing required parameters!"}), 400

    if not valid_gpt_value(gpt_value):
        return jsonify({"error": "Invalid GPT value. It should be '3.5' or '4'."}), 400

    num = max(2, math.ceil(int(num_value) / 700))

    try:
        thread_status = 1  # 标记线程为正在运行
        threading.Thread(target=threaded_write, args=(title_value, num, gpt_value)).start()
        download_link = f"/download?title={title_value}"
        return jsonify({"message": "Document generation started!", "download_link": download_link}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/download', methods=['GET'])
def download_file():
    global thread_status
    title_value = request.args.get("title", title)
    filepath = os.path.join("./output", f"{title_value}.docx")

    if os.path.exists(filepath):
        thread_status = 0  # Reset the thread status after downloading the file
        return send_from_directory(directory="./output", filename=f"{title_value}.docx", as_attachment=True)
    else:
        return jsonify({"error": "File not found!"}), 404

@app.route('/status', methods=['GET'])
def get_status():
    global thread_status
    return jsonify({"status": thread_status})

@app.route('/request_count', methods=['GET'])
def get_request_count():
    """
    Get the total number of requests in the last 3 hours and 20 minutes.
    """
    total_requests = get_requests_in_last_duration()
    return jsonify({"request_count": total_requests})

def run():
    app.run(host='0.0.0.0', port=5000)
