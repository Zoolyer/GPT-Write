import os
import time
import threading
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
import time
import random
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import os
from selenium.webdriver.chrome.service import Service
import win32clipboard as clipboard
import win32con

import os
username1 = os.getlogin()

# 启动Chrome浏览器并启用远程调试

# 配置Chrome选项以连接到远程调试端口

class ChatGPT:

    def __init__(self, gpt):
        self.driver = None
        if gpt == 4:
            self.WEB_URL = 'https://chat.openai.com/?model=gpt-4'
        else:
            self.WEB_URL = 'https://chatgpt.com/?model=text-davinci-002-render-sha'
        self.previous_text = ""

        self.thread = threading.Thread(target=self.start_chrome_with_debugging)
        self.thread.start()

        time.sleep(1)

        # Step 2: 使用Selenium WebDriver连接到这个已启动的浏览器
        self.driver = self.connect_to_chrome_with_debugging()
        # 现在你可以控制已启动的浏览器
        self.driver.get(self.WEB_URL)

    def connect_to_chrome_with_debugging(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

        # 指定chromedriver的位置
        current_directory = os.path.dirname(os.path.abspath(__file__))  # 获取当前文件的目录
        chrome_driver_path = os.path.join(current_directory, 'driver', 'chromedriver.exe')
        service = Service(chrome_driver_path)

        # 启动webdriver并连接到已启动的浏览器
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    def start_chrome_with_debugging(self):
        current_directory = os.path.dirname(os.path.abspath(__file__))  # 获取当前文件的目录
        chrome_path = os.path.join(current_directory, 'driver', 'chrome-win64', 'chrome.exe')
        chrome_user_data = r'C:\Users\\' + username1 + r'\AppData\Local\Google\Chrome for Testing\User Data'
        remote_debugging_port = 9222

        command = f'{chrome_path} --remote-debugging-port={remote_debugging_port} --user-data-dir="{chrome_user_data}"'
        print(command)
        os.system(command)
    def get_all_data_testids(self):
        """
            获取页面中所有包含data-testid属性的div元素的data-testid值
            :param driver: 已经打开的webdriver对象
            :return: 包含所有data-testid值的列表
            """
        # 查找所有包含data-testid属性的div元素
        divs_with_data_testid = self.driver.find_elements(By.XPATH, '//div[@data-testid]')

        # 提取每个div的data-testid值
        data_testid_values = [div.get_attribute('data-testid') for div in divs_with_data_testid]

        return data_testid_values

    def get_content_from_div(self, data_testid_value):
        """
            获取指定data-testid值的div元素内的所有内容
            :param driver: 已经打开的webdriver对象
            :param data_testid_value: 期望的data-testid的值
            :return: div内的所有文本内容
        """
        # 使用XPath定位具有特定data-testid值的div
        div_element = self.driver.find_element(By.XPATH, f'//div[@data-testid="{data_testid_value}"]')

        return div_element.text

    def set_clipboard_text(self,text):
        """设置剪切板内容"""
        clipboard.OpenClipboard()
        clipboard.EmptyClipboard()
        clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
        clipboard.CloseClipboard()

    def send_message_to_chat(self, message):
        # 定位输入框
        textarea = self.driver.find_element(By.ID, "prompt-textarea")
        # 清除输入框中的任何现有内容
        textarea.clear()

        # 将消息复制到剪切板
        self.set_clipboard_text(message)

        # 粘贴消息到输入框
        textarea.send_keys(Keys.CONTROL, 'v')

    def close_browser(self):
        """
        Close the current browser window.
        """
        if self.driver:
            self.driver.quit()

    def submit_message(self):
        # 定位发送按钮
        send_button = self.driver.find_element(By.CSS_SELECTOR, 'button[data-testid="fruitjuice-send-button"]')

        # 点击按钮
        send_button.click()

    def check_and_print_button_text(self):
        # 尝试定位按钮

        try:
            buttons1 = self.driver.find_elements(By.CSS_SELECTOR,  'button[data-testid="fruitjuice-send-button"]')
            buttons2 = self.driver.find_elements(By.CSS_SELECTOR,  'button[data-testid="fruitjuice-stop-button"]')
            if buttons1 :
                return 1
            elif buttons2:
                return 0
            else:
                return -1
        except:
            return -1
    def start(self,driver):
        self.driver = driver
        time.sleep(1)

    def send(self, meg):
        self.send_message_to_chat(meg)
        time.sleep(1.4)
        self.submit_message()

    def getLastMessage(self):
        return self.get_content_from_div(self.get_all_data_testids()[-1])

if __name__ == "__main__":
    # Step 1: 启动Chrome浏览器并启用远程调试
    # 等待几秒钟以确保浏览器已启动并启用了远程调试
    c = ChatGPT(3.5)
    time.sleep(2)
    c.send("你好")
    # time.sleep(3)
    # print(c.check_and_print_button_text())
    # print(c.getLastMessage())
    while True:
        print(c.check_and_print_button_text())
        time.sleep(1)
    time.sleep(100)  # 保持脚本运行以进行测试

    driver.quit()
