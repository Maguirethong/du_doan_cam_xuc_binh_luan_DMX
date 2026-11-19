from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import time
import math
import re
from xlwt import Workbook
import xlrd
import xlwt
from xlutils.copy import copy

REGEX = r"^https://www\.dienmayxanh\.com/.+$"
DRIVER = webdriver.Chrome()
DRIVER.get("https://www.dienmayxanh.com/danh-muc-nhom-hang")
wait = WebDriverWait(DRIVER, 20)


all_id = ["cate-36", "cate-38", "cate-360", "cate-43", "cate-33", "cate-40", "cate-45", "cate-49"]


def chuan_hoa(a):
    if a[-1] == "k":
        a = float(a[:-1])*1000
    elif a[-1] == "M":
        a = float(a[:-1])*1000000

    return str(int(a))

def get_product():
    for i in range(len(all_id)):
        # products = DRIVER.find_elements(By.CSS_SELECTOR, "#{} li".format(id))
        time.sleep(2)
        DRIVER.get("https://www.dienmayxanh.com/danh-muc-nhom-hang")
        time.sleep(2)
        products_category = wait.until(
            EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "#{} li a".format(all_id[i]))
            )
        )
        links = [product.get_attribute('href') for product in products_category]
        for link in links:
            if re.match(REGEX, link):
                link_param = link + r'#o=1&pi=30' #link ten cac san pham
                DRIVER.get(link_param)
                # image_err = wait.until(
                #     EC.presence_of_all_elements_located(
                #         (By.CSS_SELECTOR, "div div img[src='https://cdnv2-tmdt.tgdd.vn/webmwg/production-fe/tmdt-dmx/public/static/images/empty_product_cate.png']")
                #     )
                # )
                time.sleep(10)
                image_err = DRIVER.find_elements(By.CSS_SELECTOR, "div div img[src='https://cdnv2-tmdt.tgdd.vn/webmwg/production-fe/tmdt-dmx/public/static/images/empty_product_cate.png']")
                print("len image error",len(image_err))
                # DRIVER.find_elements(By.CSS_SELECTOR, "div div img[src='https://cdnv2-tmdt.tgdd.vn/webmwg/production-fe/tmdt-dmx/public/static/images/empty_product_cate.png']")
                if len(image_err) == 0:
                    get_comment(link_param)
            # print(len(products))
            # input()
        print("xong danh muc {}".format(i_id))

    
def get_comment(link_param):
    comments = []
    time.sleep(2)
    products = DRIVER.find_elements(By.XPATH, "//div[@class='gap-10px grid grid-cols-5']//div//a[@class='-m-10px relative']")
    links_product = [product.get_attribute("href") for product in products]
    print(len(links_product))
    # input()
    for link_product in links_product:
        new_link_product = link_product + r"/danh-gia"
        DRIVER.get(new_link_product)
        time.sleep(2)
        image_not_found = DRIVER.find_elements(By.XPATH, "//div[@class ='px-16px  mt-56px flex flex-col items-center justify-center md:h-[496px] md:flex-row md:space-x-[64px]']//div//img")
        if not image_not_found:
            try:
                wait.until(
                    EC.presence_of_element_located((By.ID, "chatbot-dmx"))
                )

                DRIVER.execute_script("""
                    let chatbot = document.getElementById('chatbot-dmx');
                    if (chatbot) chatbot.style.display = 'none';
                """)
            except Exception:
                pass
            btn_list = DRIVER.find_elements(By.CSS_SELECTOR, "div.text-gray-5.leading-14")
            print("len btn_list:", len(btn_list))
            btn_count = 0 if len(btn_list) == 0 else math.ceil(int(chuan_hoa(btn_list[-1].text.split()[0]))/20)
            print("page:",btn_count)
            if btn_count == 0:
                page_product_comments = DRIVER.find_elements(By.XPATH, "//div[@class='gap-10px flex max-w-[680px] flex-col']//div//p[2]")
                print("len p_comments", len(page_product_comments))
                comments.extend(comment.text for comment in page_product_comments)
                print("len comment", len(comments))
            else:
                for i_page in range(1, btn_count + 1):
                    page_count = DRIVER.find_elements(By.XPATH, "//button[text()='{}']".format(i_page))
                    if page_count:
                        page_count[0].click() if i_page != 1 else None
                    time.sleep(2)
                    page_product_comments = DRIVER.find_elements(By.XPATH, "//div[@class='gap-10px flex max-w-[680px] flex-col']//div//p[2]")
                    print("len p_comments", len(page_product_comments))
                    comments.extend(comment.text for comment in page_product_comments)
                    print("len comments",len(comments))   
                
        # print("len comments:",len(comments))
    save_excel(comments)

def create_excel():
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Comments")

    ws.write(0, 0, "STT")
    ws.write(0, 1, "Comment")

    wb.save("comments.xls")

def save_excel(Comments):
    rb = xlrd.open_workbook("comments.xls", formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    row = rb.sheet_by_index(0).nrows
    stt = row

    for comment in Comments:
        ws.write(row, 0, stt)
        ws.write(row, 1, comment)

        row += 1
        stt += 1

    wb.save("comments.xls")

if __name__ == "__main__":
    create_excel()
    get_product()
    input()