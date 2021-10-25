import datetime
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from tqdm import tqdm
import pandas as pd
import openpyxl
import xlsxwriter

posting_user_names = []
# 筛除热歌歌手
banned_singer = []
date_limit = datetime.datetime(2077, 8, 8)
localtime = time.localtime(time.time())
tm_y = localtime.tm_year
tm_m = localtime.tm_mon
tm_d = localtime.tm_mday

def crawler_comment_info(driver, singer_name):
    """
    options = Options()
    options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" 
    driver = webdriver.Chrome(chrome_options=options) 
    """
    all_comment_num = int(driver.find_elements_by_class_name("part__tit")[0].find_elements_by_class_name("c_tx_thin")[0]\
                .text.split("共")[1].split("条评论")[0])
    if all_comment_num == 0:
        return 0, 0, 0, 0, 0, 0, 0

    for comment_block in driver.find_elements_by_class_name("mod_hot_comment"):
        if "全部评论" in comment_block.find_elements_by_class_name("comment_type__title")[0].text:
            comment_list_block = comment_block.find_elements_by_class_name("comment__list")[0]

    # 将页面滚动到底端 以加载所有评论
    js_bottom = "window.scrollTo(0, document.body.scrollHeight)"
    pre_comment_num = len(comment_list_block.find_elements_by_tag_name("li"))
    while True:
        driver.execute_script(js_bottom)
        time.sleep(2)
        comment_num = len(comment_list_block.find_elements_by_tag_name("li"))
        if comment_num == pre_comment_num:
            break
        pre_comment_num = comment_num
        time.sleep(1)
    comments_list = comment_list_block.find_elements_by_tag_name("li")

    comment_num = 0
    singer_reply_num = 0
    reply_num_sum = 0
    zan_num_sum = 0
    replied_comment_num = 0
    zaned_comment_num = 0
    zaned_more_than_one_comment_num = 0
    comments_info = []
    for com in comments_list:
        content = com.find_elements_by_class_name("comment__text")[0].text
        if "- 该评论已删除 -" == content:
            continue
        user_name_block = com.find_elements_by_class_name("comment__title")
        # 筛掉没有用户名的不合规范评论
        if len(user_name_block) == 0:
            continue
        user_name = user_name_block[0].text
        # 筛掉用于实验投放的账户
        if user_name in posting_user_names:
            continue
        post_date_str = com.find_elements_by_class_name("comment__date")[0].text.split()[0]
        year, month, day = tm_y, tm_m, tm_d
        if "年" in post_date_str:
            year = int(post_date_str.split("年")[0])
            month = int(post_date_str.split("年")[1].split("月")[0])
            day = int(post_date_str.split("月")[1].split("日")[0])
        elif "月" in post_date_str:
            month = int(post_date_str.split("月")[0])
            day = int(post_date_str.split("月")[1].split("日")[0])
        post_date = datetime.datetime(year, month, day)
        if post_date < date_limit:
            continue
        comment_num += 1
        zan_num = com.find_elements_by_class_name("comment__zan")[0].text
        zan_num = int(zan_num) if zan_num != "" else 0
        if zan_num > 0:
            zaned_comment_num += 1
        if zan_num > 1:
            zaned_more_than_one_comment_num += 1

        singer_reply = False
        user_reply = False
        if len(com.find_elements_by_class_name("comment__reply")) > 0:
            reply_block = com.find_elements_by_class_name("comment__reply")[0]
            reply_num = int(com.find_elements_by_class_name("comment__show_all_reply")[0].text.split("查看")[1].split("条回复")[0])

            # 展开评论所有回复
            while True:
                try:
                    reply_block.find_elements_by_class_name("comment__icon_arrow_down")[0].click()
                    time.sleep(1)
                    break
                except:
                    time.sleep(1)
            while reply_block.find_elements_by_class_name("comment__reply_more")[0].text == "显示更多回复":
                while True:
                    try:
                        reply_block.find_elements_by_class_name("comment__icon_reply_more")[0].click()
                        time.sleep(1)
                        break
                    except:
                        time.sleep(1)
            reply_list = reply_block.find_elements_by_class_name("comment__list")[0].find_elements_by_tag_name("li")
            for reply in reply_list:
                if singer_reply and user_reply:
                    break
                # 查看作者是否回复
                if not singer_reply and reply.find_elements_by_class_name("comment__title")[0].text == singer_name:
                    singer_reply_num += 1
                    singer_reply = True
                elif not user_reply and reply.find_elements_by_class_name("comment__title")[0].text != singer_name:
                    replied_comment_num += 1
                    user_reply = True
        else:
            reply_num = 0
        
        zan_num_sum += zan_num
        reply_num_sum += reply_num
        if singer_reply:
            reply_num_sum -= 1

        info = {"user_name": user_name, "zan_num": zan_num, "reply_num": reply_num, "singer_reply": singer_reply}
        comments_info.append(info)

    return comment_num, replied_comment_num, zaned_comment_num, zaned_more_than_one_comment_num, singer_reply_num, reply_num_sum, zan_num_sum


driver = webdriver.Firefox()

book = xlsxwriter.Workbook("data/歌曲评论区情况.xlsx")
wb = openpyxl.load_workbook("data/xxx.xlsx")
sheets_names = wb.get_sheet_names()
songs_info = []
song_cnt = 0
for sheet_name in sheets_names:
    sheet_pd = pd.read_excel("data/xxx.xlsx", sheet_name)
    sheet = book.add_worksheet(sheet_name)
    sheet.set_column('A:A', 15)
    sheet.set_column('B:B', 15)
    sheet.set_column('D:D', 40)
    sheet.set_column('E:E', 15)
    sheet.set_column('F:F', 20)
    sheet.set_column('G:G', 20)
    sheet.set_column('G:G', 20)
    sheet.set_column('H:H', 15)
    sheet.set_column('I:I', 15)
    sheet.set_column('J:J', 15)
    sheet.write(0, 0, "Track ID")
    sheet.write(0, 1, "Track Name")
    sheet.write(0, 2, "Singer")
    sheet.write(0, 3, "SF")
    sheet.write(0, 4, "Comment Num")
    sheet.write(0, 5, "Zaned Comment Num")
    sheet.write(0, 6, "Zaned >1 Comment Num")
    sheet.write(0, 7, "Replied Comment Num")
    sheet.write(0, 8, "Zan Num")
    sheet.write(0, 9, "Reply Num")
    sheet.write(0, 10, "Singer Reply Num")
    line_idx = 1
    for _, row in tqdm(sheet_pd.iterrows()):
        if row["Singer"] in banned_singer:
            continue
        driver.get(row["SF"])
        time.sleep(1)
        if "很抱歉，您查看的歌曲已下架" in driver.page_source:
            continue
        comment_num, replied_comment_num, zaned_comment_num, zaned_more_than_one_comment_num, singer_reply_num, reply_num_sum, zan_num_sum = \
            crawler_comment_info(driver, row["Singer"])
        info = [row["Track ID"], row["Track Name"], row["Singer"], row["SF"], \
            comment_num, zaned_comment_num, zaned_more_than_one_comment_num, replied_comment_num, zan_num_sum, reply_num_sum, singer_reply_num]
        for i, item in enumerate(info):
            sheet.write(line_idx, i, item)
        line_idx += 1
        song_cnt += 1
book.close()

print(f"total {song_cnt} tracks")