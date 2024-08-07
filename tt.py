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

from selenium.webdriver.chrome.options import Options
import os

# 使用os.getlogin()
username1 = os.getlogin()


def set_clipboard_text(text):
    """设置剪切板内容"""
    clipboard.OpenClipboard()
    clipboard.EmptyClipboard()
    clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    clipboard.CloseClipboard()


# Get the current directory of this script
current_directory = os.path.dirname(os.path.abspath(__file__))  # 获取当前文件的目录

# Append the relative path of the Chrome driver
CHROME_DRIVER_PATH = os.path.join(current_directory, 'driver\chromedriver.exe')

class ChatGPT:

    def __init__(self, gpt):
        self.driver = None
        if gpt == "4":
            self.WEB_URL = 'https://chat.openai.com/?model=gpt-4'
        else:
            self.WEB_URL = 'https://chatgpt.com/?model=text-davinci-002-render-sha'
        self.previous_text = ""

    def open_website_with_proxy(self):
        options = webdriver.ChromeOptions()
        # 禁用 window.navigator.webdriver 标志
        CHROME_PATH = os.path.join(current_directory, 'driver\chrome-win64\chrome.exe')
        options.binary_location = CHROME_PATH
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        CHROME_OPTIONS = "user-data-dir=" + r'C:\Users\\' + username1 + r'\AppData\Local\Google\Chrome for Testing\User Data'

        options.add_argument(CHROME_OPTIONS)
        time.sleep(2)

        service = Service(CHROME_DRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)
        # 打开网站并添加随机延迟https://chat.openai.com/?model=text-davinci-002-render-sha
        driver.get(self.WEB_URL)
        # time.sleep(2 + 1 * random.random())  # 随机延迟3到5秒

        # time.sleep(10)
        try:
            # button = WebDriverWait(driver, 8).until(
            #     EC.presence_of_element_located(
            #         (By.CSS_SELECTOR,
            #          'button[data-testid="fruitjuice-send-button"].mb-1.me-1.flex.h-8.w-8.items-center.justify-center.rounded-full.bg-black.text-white.transition-colors.hover\\:opacity-70.focus-visible\\:outline-none.focus-visible\\:outline-black.disabled\\:bg-\\#D7D7D7.disabled\\:text-\\#f4f4f4.disabled\\:hover\\:opacity-100.dark\\:bg-white.dark\\:text-black.dark\\:focus-visible\\:outline-white.disabled\\:dark\\:bg-token-text-quaternary.dark\\:disabled\\:text-token-main-surface-secondary')
            #     )
            # )
            button = WebDriverWait(driver, 8).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'button[data-testid="fruitjuice-send-button"]')
                )
            )
            button.click()
        except:
            print("未能在指定时间内找到按钮，跳过点击操作。")

        return driver

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

    def send_message_to_chat(self, message):
        # 定位输入框
        textarea = self.driver.find_element(By.ID, "prompt-textarea")
        # 清除输入框中的任何现有内容
        textarea.clear()

        # 将消息复制到剪切板
        set_clipboard_text(message)

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
            buttons = self.driver.find_elements(By.CSS_SELECTOR, 'button.btn-neutral')
            # 如果找到了按钮
            if buttons:
                button = buttons[0]
                text = button.text
                # 如果文本内容与上一次的不同
                if text != self.previous_text:
                    self.previous_text = text
                    if "Regenerate" in text:
                        return True
                    if "Continue generating" in text:
                        button.click()  # 如果文本是 "Continue generating" 则点击按钮\
                        time.sleep(1)
                        # print(text)
                        # ls = self.get_all_data_testids()
                        # content = self.get_paragraph_content_from_div(ls[-1])
                        # print(content)

        except StaleElementReferenceException:
            # 当出现StaleElement异常时，您可以重新定位元素或者简单地返回，下次再尝试
            return False

        return False

    def start(self):
        self.driver = self.open_website_with_proxy()
        time.sleep(1)

    def send(self, meg):
        self.send_message_to_chat(meg)
        time.sleep(1.4)
        self.submit_message()

    def getLastMessage(self):
        return self.get_content_from_div(self.get_all_data_testids()[-1])


if __name__ == "__main__":
    gpt = ChatGPT("3.5")
    gpt.start()
    # gpt.send("aaaaaaaaa")
    time.sleep(100)
