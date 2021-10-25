import time
import random
import json
import logging
import pandas as pd
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from tqdm import tqdm
from urllib import request

# 程序终端提醒
def post_webhook():
    # 企业微信机器人链接
    hook_url = ""
    data={"msgtype": "text", "text":{"content": "投放程序中断"}}
    postData = str(json.dumps(data)).encode("utf-8")
    req = request.Request(hook_url, data=postData)
    resp = request.urlopen(req)

logging.basicConfig(level=logging.DEBUG,
                format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S',
                filename='comment_posting.log',
                filemode='w')

users_num = 2

# 需要投送评论的表格
wb = openpyxl.load_workbook("data/xxx.xlsx")
sheets_names = wb.get_sheet_names()
songs_info = []
for sheet_name in sheets_names:
    sheet = pd.read_excel("data/xxx.xlsx", sheet_name)
    for _, row in sheet.iterrows():
        songs_info.append({"id": row["Track ID"], "url": row["SF"], "story": row["Story Review"], "poetic": row["Poetic Review"]})
random.shuffle(songs_info)

review_types = ["story", "poetic"]

drivers = []
for i in range(users_num):
    # Chrome对emoji支持不友好 选择Firefox
    driver = webdriver.Firefox() 
    driver.get("https://y.qq.com/")
    drivers.append(driver)

while True:
    print("请登录QQ音乐账户 输入c继续")
    if input() == "c":
        break
time.sleep(1)

scroll_sizes = [800, 400, 600, 1000, 1200]

try:
    for driver, review_type in zip(drivers, review_types):
        for song_info in tqdm(songs_info):
            url = song_info["url"]
            driver.get(url)
            if "很抱歉，您查看的歌曲已下架" in driver.page_source:
                continue
            # driver.execute_script(f'document.querySelector("div.comment__textarea_input").innerText = "{comment}"')
            for scroll_size in scroll_sizes:
                # 将页面向下滚动800个像素 让评论输入文本框在屏幕上露出
                time.sleep(random.uniform(1, 2))
                driver.execute_script(f'window.scrollBy(0, {scroll_size})')
                time.sleep(random.uniform(1, 2))
                try:
                    # 双击清空文本框
                    ActionChains(driver).double_click(driver.find_elements_by_class_name("comment__textarea_inner")[0]).perform()
                    time.sleep(random.uniform(5, 10))
                    break
                except:
                    driver.execute_script("var q=document.documentElement.scrollTop=0")
            comment = song_info[review_type]
            driver.find_elements_by_class_name("comment__textarea_input")[0].send_keys(comment)
            time.sleep(random.uniform(1, 2))
            driver.find_elements_by_class_name("comment__tool")[0].click()
            track_id = song_info["id"]
            logging.info(f"finish posting track {track_id} {review_type} comment")
            logging.info(f"track url: {url}")
            time.sleep(random.uniform(120, 360))
except:
    post_webhook()