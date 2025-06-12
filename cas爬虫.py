#用于从https://www.chemradar.com/tools/cis/inv/648fb470e7fff39f78795ebc?爬取数据

import requests
from bs4 import BeautifulSoup
import pandas as pd
import tkinter as tk
from tkinter import messagebox
import os
import pyodbc  # 连接 SQL Server

# SQL Server 连接信息（Windows 身份验证）
server = "localhost"  # SQL Server 服务器名
database = "cas"  # 数据库名称

# 连接 SQL Server（Windows 身份验证）
try:
    conn = pyodbc.connect(f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;")
    cursor = conn.cursor()
    print("成功连接到 SQL Server")
except Exception as e:
    print(f"数据库连接失败: {e}")
    exit()

# 读取 Excel 文件
file_path = "C:/Users/cheei/PycharmProjects/cas/爬虫/参数1.xlsx"
df = pd.read_excel(file_path)

# 请求头和 Cookies（替换为网站返回的）
cookies = {

}

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Referer": "https://www.chemradar.com/",
}

# 弹窗提示函数
def show_popup(message):
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.attributes('-topmost', 1)  # 设置弹窗始终在最前面
    return messagebox.askyesno("滑块验证码", message)

# 保存进度到文件
def save_progress(param1, page_num):
    with open('progress.txt', 'w') as f:
        f.write(f"{param1},{page_num}\n")

# 读取进度文件
def read_progress():
    if os.path.exists('progress.txt'):
        with open('progress.txt', 'r') as f:
            line = f.readline().strip()
            if line:
                param1, page_num = line.split(',')
                return int(param1), int(page_num)
    return None, None

# 等待用户按回车键继续
def wait_for_continue():
    input("按回车继续爬取...")

# 读取进度
last_param1, last_page_num = read_progress()

# 遍历 Excel 表格中的参数1
for index, row in df.iterrows():
    param1 = int(row['参数1'])  # 读取参数1

    # 如果上次中断时的参数1大于当前参数1，跳过
    if last_param1 is not None and param1 < last_param1:
        continue

    print(f"🔄 正在爬取: {param1} 的页面...")

    session = requests.Session()
    url = f"https://www.chemradar.com/tools/cis/inv/648fb470e7fff39f78795ebc?keyword={param1}-&type=2&pageNum=1"

    # 如果当前参数1是上次中断时的参数1，从上次的页码开始
    if last_param1 == param1 and last_page_num is not None:
        print(f"⏸️ 从上次停止的位置继续爬取，恢复到第 {last_page_num} 页。")
        start_page = last_page_num
    else:
        start_page = 1

    def get_page_data():
        """获取页面数据"""
        response = session.get(url, headers=headers, cookies=cookies)
        if response.status_code != 200:
            print(f"❌ 请求失败，状态码: {response.status_code}")
            return None
        response.encoding = "utf-8"
        return BeautifulSoup(response.text, "html.parser")

    # **第一次请求**
    soup = get_page_data()
    if not soup:
        continue  # 跳过当前参数1

    # **检查是否是空页面**
    if soup.find("div", class_="ant-empty-image"):
        if show_popup(f"⚠️ {param1} 可能触发滑块验证码，是否暂停爬取并手动处理？"):
            save_progress(param1, start_page)  # 保存进度
            print("⏸️ 暂停，等待您处理滑块验证码...")
            wait_for_continue()  # 等待用户按回车继续
            print("✅ 继续爬取...")

        soup = get_page_data()  # **第二次请求**

        if soup.find("div", class_="ant-empty-image"):  # 再次无数据，才是真的无数据
            print(f"⚠️ {param1} 真的无数据，跳过。")
            continue
        else:
            print(f"✅ {param1} 滑块处理成功，继续爬取...")

    # **计算最大页数**
    pagination_items = soup.select("ul.ant-pagination li.ant-pagination-item")
    max_page = max(
        int(item["title"]) for item in pagination_items if item["title"].isdigit()) if pagination_items else 1
    print(f"📄 {param1} 共有 {max_page} 页")

    # **逐页爬取**
    for page_num in range(start_page, max_page + 1):
        print(f"  🔍 爬取第 {page_num} 页...")
        url = f"https://www.chemradar.com/tools/cis/inv/648fb470e7fff39f78795ebc?keyword={param1}-&type=2&pageNum={page_num}"
        soup = get_page_data()

        # **检查是否是空页面**
        if soup.find("div", class_="ant-empty-image"):
            if show_popup(f"⚠️ {param1} 可能触发滑块验证码，是否暂停爬取并手动处理？"):
                save_progress(param1, page_num)  # 保存进度
                print("⏸️ 暂停，等待您处理滑块验证码...")
                wait_for_continue()  # 等待用户按回车继续
                print("✅ 继续爬取...")

            soup = get_page_data()  # **第二次请求**

            if soup.find("div", class_="ant-empty-image"):  # 再次无数据，才是真的无数据
                print(f"⚠️ {param1} 无数据，跳过。")
                continue
            else:
                print(f"✅ {param1} 滑块处理成功，继续爬取...")

        if not soup:
            print(f"  ❌ 第 {page_num} 页请求失败，跳过")
            continue

        main_container = soup.find("main", class_="pb-24 container")
        if main_container:
            divs = main_container.find_all("div", class_="relative")
            for div in divs:
                flex_div = div.find("div", class_="flex flex-row text-sm")
                if flex_div:
                    font_divs = flex_div.find_all("div", class_="font-normal")
                    if len(font_divs) >= 2:
                        cas_number = font_divs[0].get_text(strip=True)
                        chemical_name = font_divs[1].get_text(strip=True)
                        if cas_number and chemical_name:
                            # 插入数据库
                            try:
                                cursor.execute("INSERT INTO cas (CAS_Number, English_Substance_Name) VALUES (?, ?)",
                                               (cas_number, chemical_name))
                                conn.commit()
                                print(f"  ✅ 已插入: CAS号: {cas_number}, 化学品名称: {chemical_name}")
                            except Exception as e:
                                print(f"  ❌ 插入失败: {e}")
                        else:
                            print(f"  ❌ 找到 CAS 号，但化学名称为空")
                    else:
                        print(f"  ❌ 找不到有效的 CAS 号或化学名称")
        else:
            print(f"  ❌ 第 {page_num} 页未找到数据")

        # 保存进度
        save_progress(param1, page_num)

    print(f"✅ {param1} 爬取完成\n{'-' * 50}")

# 关闭数据库连接
cursor.close()
conn.close()
print("数据库连接已关闭")
