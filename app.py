# coding:utf-8
from openpyxl import Workbook
from selenium import webdriver
import sys
import os
import re
import requests
import multiprocessing
from tkinter.messagebox import *
from tkinter import scrolledtext, Tk, ttk, filedialog
from mttkinter import mtTkinter as mtk
import threading
import random
import time
from datetime import datetime
from selenium.webdriver.support.wait import WebDriverWait

FILE_PATH_01 = os.path.join(os.path.expanduser("~"), 'Desktop/技工证书/IMAGE').replace("\\", "/")
if not os.path.exists(FILE_PATH_01):
    os.makedirs(FILE_PATH_01)

FILE_PATH_02 = os.path.join(os.path.expanduser("~"), 'Desktop/职称证/IMAGE/02').replace("\\", "/")
if not os.path.exists(FILE_PATH_02):
    os.makedirs(FILE_PATH_02)


class HNProject:

    def __init__(self, master):
        self.root = master
        self.root.geometry("560x380")
        self.root.title("HN工具 1.0")
        self.__create_ui()
        self.__creatBrowser()
        self.__login()
        self.addLog("初始化成功,请进行登录并点击在线办理进入主页面.")

    def __create_ui(self):
        self.msg_box = mtk.LabelFrame(self.root, text="统计信息", fg="blue")
        self.msg_box.place(x=20, y=20, width=250, height=80)

        self.msg_box_0 = mtk.Label(self.msg_box, text="当前无任何数据.")
        self.msg_box_0.place(x=30, y=15, width=150, height=25)

        self.settings_box = mtk.LabelFrame(self.root, text="任务设置", fg="blue")
        self.settings_box.place(x=20, y=120, width=250, height=130)
        # 模板选择
        task_name = mtk.Label(self.settings_box, text="任务名称：")
        task_name.place(x=10, y=10, width=60, height=25)
        self.task_id = ttk.Combobox(self.settings_box)
        self.task_id["values"] = ["技术证", "职称证"]
        self.task_id.place(x=80, y=10, width=150, height=25)
        # 随机间隔时间 10-30-50-60   设置开始时间
        id_setting = mtk.Label(self.settings_box, text="ID设置：")
        id_setting.place(x=10, y=55, width=60, height=25)
        self.id_from = mtk.Entry(self.settings_box)
        self.id_from.place(x=80, y=55, width=55, height=25)
        id_bt = mtk.Label(self.settings_box, text="~")
        id_bt.place(x=135, y=55, width=30, height=25)
        self.id_end = mtk.Entry(self.settings_box)
        self.id_end.place(x=170, y=55, width=55, height=25)
        # 提示信息框
        self.log_box = mtk.LabelFrame(self.root, text="提示信息", fg="blue")
        self.log_box.place(x=290, y=20, width=250, height=330)
        self.logtext = scrolledtext.ScrolledText(self.log_box, fg="green")
        self.logtext.place(x=15, y=5, width=230, height=290)
        # 任务开始栏
        self.task_box = mtk.LabelFrame(self.root)
        self.task_box.place(x=20, y=270, width=250, height=80)
        # 导出数据
        self.load_file_btn = mtk.Button(self.task_box, text="导出数据",
                                        command=lambda: self.thread_it(self.__download_data))
        self.load_file_btn.place(x=10, y=20, width=120, height=40)
        self.task_start_btn = mtk.Button(self.task_box, text="开    始", command=lambda: self.thread_it(self.__start))
        self.task_start_btn.place(x=150, y=20, width=80, height=40)

    def __login(self):
        url = "https://www.hnzwfw.gov.cn/portal/guide/E6F374C00C190CE67FEE90787833B5F4?region=410100000000"
        self.driver.get(url)
        showinfo("提示信息", "请进行登录并点击在线办理进入主页面.")
        return

    def __creatBrowser(self):
        # 创建driver
        try:
            options = webdriver.ChromeOptions()
            # options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            driver = webdriver.Chrome(chrome_options=options)
            driver.set_page_load_timeout(30)
            driver.set_script_timeout(30)
            self.driver = driver
            self.wait = WebDriverWait(self.driver, 600)
        except Exception as e:
            print(e.args)
            return

    @staticmethod
    def thread_it(func, *args):
        t = threading.Thread(target=func, args=args)
        t.setDaemon(True)
        t.start()

    def addLog(self, msg):
        self.logtext.insert(mtk.END, "{} {}\n".format(datetime.now().strftime("%H:%M:%S"), msg))
        self.logtext.yview_moveto(1.0)

    def __get_content(self, task_id, task_name):
        if task_name == "技术证":
            base_url = "http://222.143.38.205:1717/manager/JianZhuYeZiZhi/dghr_JiShuRenYuanMingDan_86202Edit.aspx?zhuid=42103&gvid=mcp_gridlist&pkid={}&ptabid=xnode-67"
        else:
            base_url = "http://222.143.38.205:1717/manager/JianZhuYeZiZhi/dghr_ZhongJiJiYiShangZhiChengRenYuanMingDan_25f0eEdit.aspx?zhuid=-10&gvid=mcp_gridlist&pkid={}&ptabid=xnode-67"
        url = base_url.format(task_id)
        self.driver.get(url)
        time.sleep(random.uniform(0, 1))
        if task_name == "技术证":
            data_item = self.__task_1_parser(task_id, url)
        else:
            data_item = self.__task_2_parser(task_id, url)
        return data_item

    def __task_1_parser(self, task_id, referer):
        html_text = self.driver.page_source
        xingming = re.findall(r'mcp_drpXingMing_Container",width:260,value:"(.*?)"', html_text)
        if not xingming:
            return
        shenfenzheng = re.findall(r'mcp_txtShenFenZhengHao_Container",width:260,value:"(.*?)"', html_text)
        zhuanye = re.findall(r'mcp_txtZhuanYeGongZhong_Container",width:260,value:"(.*?)"', html_text)
        dengji = re.findall(r'mcp_txtJiNengDengJi_Container",width:260,value:"(.*?)"', html_text)
        zhengshu = re.findall(r'mcp_txtGangWeiZhengShuBianHao_Container",width:260,value:"(.*?)"', html_text)
        fazhengdanwei = re.findall(r'mcp_txtFaZhengDanWei_Container",width:260,value:"(.*?)"', html_text)
        image_sfz = re.findall(r'mcp_txtShenFenZhengFJ",xtype:"textfield",width:260,value:"(.*?)"', html_text)
        image_zs = re.findall(r'mcp_txtZhengShuFJ",xtype:"textfield",width:260,value:"(.*?)"', html_text)
        img_li = []
        if image_sfz:
            img_link = "http://222.143.38.205:1717" + image_sfz[0]
            img_name = f"{xingming[0]}_{task_id}_sfz.jpg"
            img_li.append((img_link, img_name, referer))
        if image_zs:
            img_link = "http://222.143.38.205:1717" + image_zs[0]
            img_name = f"{xingming[0]}_{task_id}_zs.jpg"
            img_li.append((img_link, img_name, referer))

        data_item = {
            "xingming": xingming[0],
            "shenfenzheng": shenfenzheng[0] if shenfenzheng else "无",
            "zhuanye": zhuanye[0] if zhuanye else "无",
            "dengji": dengji[0] if dengji else "无",
            "zhengshu": zhengshu[0] if zhengshu else "无",
            "fazhengdanwei": fazhengdanwei[0] if fazhengdanwei else "无",
            "image_info": img_li
        }
        return data_item

    def __task_2_parser(self, task_id, referer):
        html_text = self.driver.page_source
        xingming = re.findall(r'mcp_txtXingMing_Container",width:180,value:"(.*?)"', html_text)
        if not xingming:
            return

        xingbie = re.findall(r'mcp_drpXingBie_Value".*?initSelectedIndex:(\d+)', html_text)
        if xingbie:
            xingbie_ = "男" if xingbie[0] == "0" else "女"
        else:
            xingbie_ = ""
        shenfenzheng = re.findall(r'mcp_txtShenFenZhengHao_Container",width:\d+,value:"(.*?)"', html_text)
        zhiyezhengshu = re.findall(r'mcp_txtZCNum_Container",width:\d+,value:"(.*?)"', html_text)
        xuelizhengshu = re.findall(r'mcp_txtXLNum_Container",width:\d+,value:"(.*?)"', html_text)
        zhuanye = re.findall(r'mcp_txtXueLiZhuanYe_Container",width:\d+,value:"(.*?)"', html_text)
        zhicheng = re.findall(r'mcp_drpJiBie_Value",.*?initSelectedIndex:(\d+)', html_text)
        if zhicheng:
            if zhicheng[0] == "0":
                zhicheng_ = "高级工程师"
            elif zhicheng[0] == "1":
                zhicheng_ = "中级工程师"
            elif zhicheng[0] == "2":
                zhicheng_ = "初级工程师"
            else:
                zhicheng_ = "无"
        else:
            zhicheng_ = "无"

        image_sfz = re.findall(r'mcp_txtShenFenZhengFJ",xtype:"textfield",width:\d+,value:"(.*?)"', html_text)
        image_zczs = re.findall(r'mcp_txtZhengShuFJ",xtype:"textfield",width:\d+,value:"(.*?)"', html_text)
        image_xlzs = re.findall(r'mcp_txtBiYeZhengFJ",xtype:"textfield",width:\d+,value:"(.*?)"', html_text)
        img_li = []
        if image_sfz:
            img_link = "http://222.143.38.205:1717" + image_sfz[0]
            img_name = f"{xingming[0]}_{task_id}_sfz.jpg"
            img_li.append((img_link, img_name, referer))
        if image_zczs:
            img_link = "http://222.143.38.205:1717" + image_zczs[0]
            img_name = f"{xingming[0]}_{task_id}_zczs.jpg"
            img_li.append((img_link, img_name, referer))
        if image_xlzs:
            img_link = "http://222.143.38.205:1717" + image_xlzs[0]
            img_name = f"{xingming[0]}_{task_id}_xlzs.jpg"
            img_li.append((img_link, img_name, referer))

        data_item = {
            "xingming": xingming[0],
            "xingbie": xingbie_,
            "shenfenzheng": shenfenzheng[0] if shenfenzheng else "无",
            "zhiyezhengshu": zhiyezhengshu[0] if zhiyezhengshu else "无",
            "xuelizhengshu": xuelizhengshu[0] if xuelizhengshu else "无",
            "zhuanye": zhuanye[0] if zhuanye else "无",
            "zhicheng": zhicheng_,
            "image_info": img_li
        }
        return data_item

    def __download_data(self):
        if not self.total_data:
            showerror("错误信息", "当前不存在任何数据!")
            return
        excelPath = filedialog.asksaveasfilename(title=u'保存文件', filetypes=[("xlsx", ".xlsx")]) + ".xlsx"
        if excelPath.strip(".xlsx"):
            wb = Workbook()
            ws = wb.active
            for item in self.total_data:
                ws.append(list(item.values())[:-1])
            wb.save(excelPath)
            showinfo("提示信息", "保存成功！")

    def __download_img(self, img_info, FILE_PATH):
        if not img_info:
            return
        for img_link, img_name, referer in img_info:
            try:
                headers = {
                    "Referer": referer,
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.101 Safari/537.36",
                }
                content = requests.get(img_link, headers=headers).content
                image_path = f"{FILE_PATH}/{img_name}"
                with open(image_path, "wb")as file:
                    file.write(content)
            except:
                continue

    def __start(self):
        self.total_data = []
        task_index = 0
        task_name = self.task_id.get()
        id_start = self.id_from.get()
        try:
            id_start_ = int(id_start)
        except:
            showerror("请输入正确的ID")
            return

        id_end = self.id_end.get()
        try:
            id_end_ = int(id_end)
        except:
            showerror("请输入正确的ID")
            return

        if task_name == "技术证":
            FILE_PATH = FILE_PATH_01
        else:
            FILE_PATH = FILE_PATH_02
        # 数据下载
        with open(f"{FILE_PATH.replace('/IMAGE', '')}/{task_name}.txt", "a+", encoding="utf-8")as file:
            for task_id in range(id_start_, id_end_ + 1):
                data_item = self.__get_content(task_id, task_name)
                if data_item:
                    self.total_data.append(data_item)
                    img_info = data_item.pop("image_info")
                    if img_info:
                        try:
                            self.__download_img(img_info, FILE_PATH)
                        except:
                            continue
                        self.addLog("获取数据成功：pkid={}.".format(task_id))
                    file.write("||".join(list(data_item.values())) + "\n")
                else:
                    self.addLog("获取数据失败：pkid={},此pkid无对应信息.".format(task_id))
                self.msg_box_0.config(text="当前共下载数据：{}条.".format(task_index + 1))

                task_index += 1
                time.sleep(random.uniform(1, 3))
        self.addLog("任务完成,请导出数据.")
        showinfo("提示信息", "数据采集完成,请导出数据.")


if __name__ == '__main__':
    if sys.platform.startswith('win'):
        multiprocessing.freeze_support()
    root = Tk()
    app = HNProject(root)
    root.mainloop()
