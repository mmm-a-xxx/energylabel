# 需搭配模板文件“能效网信息提取.xlsx”使用，将此py文件与模板文件放置在同一文件夹中即可。
# 目前仅支持爬取：能效标识备案信息系统v4.0 https://www.energylabel.com.cn/search/pubintro.html

import requests
import pandas as pd
import tkinter as tk
import time

#通过能效网型号获取能效网备案号，支持一个型号对应多个备案号
def get_recordno(model):
    url = f"https://rec.energylabelrecord.com/recordMainController/afficheList?modelno=30&productModel={model}&page=1&limit=10&producerName=&_=1726413699614"
    response = requests.get(url)
    response_list=response.json() #可能包含多个字典，意味着有多个备案号
    list = [] #存放备案号
    list1 = []
    for response_dict in response_list:
        if response_dict["productModel"].lower() == str(model).lower():
            uid = response_dict["uid"]
            #通过产品uid获取详细能效信息（仅适用于“微型计算机”）
            url_result = f"https://rec.energylabelrecord.com/recordMainController/getnot?uid={uid}&_=1726414951380"
            response_result = requests.get(url_result)
            response_result_dict = response_result.json()
            list.append(response_result_dict["recordno"])
            list1.append(response_result_dict["energyLevel"])
        else:
            list.append("失败：未查询到此型号")
            list1.append("失败：未查询到此型号")
            break
    return list,list1

#从excel读取能效网型号并调用get_recordno获取备案号
def recordno_to_excel():
    # 创建空的结果输出表
    data1 = {'产品型号': [], '备案号': [], '能效等级': []}
    df1 = pd.DataFrame(data1)

    df = pd.read_excel("能效网信息提取.xlsx")
    for i in df.values[:,0]:  #读取型号列每个值
        list_recordno,list_level = get_recordno(i)
        for no,level in zip(list_recordno,list_level):
            new_row = pd.DataFrame({'产品型号':[i],'备案号':[no],'能效等级':[level]})
            df1 = df1._append(new_row)

    df1.to_excel('能效信息-根据型号查备案号.xlsx',index=False)


#从能效网备案号查能效网型号
def get_model(recordno):
    list2 = []
    list3 = []
    url = f"https://rec.energylabelrecord.com/recordMainController/afficheList?modelno=30&recordNo={recordno}&page=1&limit=10&_=1726475905995"
    header = {"user-agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"}
    response = requests.get(url,headers=header)
    response_list = response.json()
    if len(response_list) == 1: #如果通过备案号查询出来多条信息，意味着没查到有此条备案号
        response_dic = response_list[0]
        uid = response_dic["uid"]
        url_result = f"https://rec.energylabelrecord.com/recordMainController/getnot?uid={uid}&_=1726414951380"
        response_result = requests.get(url_result,headers=header)
        response_result_dict = response_result.json()
        if response_result_dict["recordno"] == recordno: #二次确认备案号是否与输入的一致
            list2.append(response_result_dict["model"] )
            list3.append(response_result_dict["energyLevel"])
        else:
            list2.append("备案号与输入不匹配")
            list3.append("备案号与输入不匹配")
    else:
        list2.append("备案号查询返回了多个结果")
        list3.append("备案号查询返回了多个结果")
    return list2,list3

#从excel读取能效网备案号并调用get_model获取型号
def model_to_excel():
    data3 = {'产品型号': [], '备案号': [], '能效等级': []}
    df3 = pd.DataFrame(data3)
    df2 = pd.read_excel("能效网信息提取.xlsx")
    for i in df2.values[:, 1]:  # 读取备案号列每个值
        time.sleep(3)
        list_model, list_level1 = get_model(i)
        for model,level1 in zip(list_model, list_level1):
            new_row1 = pd.DataFrame({'产品型号': [model], '备案号': [i], '能效等级': [level1]})
            df3 = df3._append(new_row1)
    df3.to_excel('能效信息-根据备案号查型号.xlsx',index=False)

# 创建窗口
win = tk.Tk()
win.title(u'能效网信息提取')
win.geometry("400x110")

B1 = tk.Button(win,text="根据型号查备案号",activebackground="gray",command=recordno_to_excel)
B1.place(x=70, y=50)
B2 = tk.Button(win,text="根据备案号查型号",activebackground="gray",command=model_to_excel)
B2.place(x=220, y=50)


lb1 = tk.Label(text="按下按钮变为灰色，表示正在进行中，按钮变白即为完成")
lb2 = tk.Label(text="请勿同时按下两个按钮，会发生未知错误")
lb1.place(x=10,y=20)
lb2.place(x=10,y=90)
lb3 = tk.Label(text="请将此程序与填写好的“能效网信息提取.xlsx”放在同一文件夹中")
lb3.place(x=10,y=0)

win.mainloop()
