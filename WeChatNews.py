import requests
from lxml import etree
import pandas as pd
import os
from copy import copy
from urllib3 import encode_multipart_formdata

def page_news():
    # 用于存储新闻标题
    news_titile_zhihu = []
    news_titile_weibo = []
    news_titile_wechat = []
    news_title_pengpai = []
    news_title_baidu = []
    news_title_toutiao = []
    news_title_sougou = []
    news_titile_360 = []
    news_titile_xinjingbao = []

    url = f"https://tophub.today/c/news?p=1"
    headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.72 Safari/537.36"
    }
    response = requests.get(url=url, headers=headers)
    tree = etree.HTML(response.text)
    for i in range(1,51):
        # zhihu
        xpath = f'//*[@id="node-6"]/div/div[2]/div/a[{i}]/div/span[2]'
        news_titile_zhihu.append(tree.xpath(xpath)[0].text)

        # weibo
        xpath = f'//*[@id="node-1"]/div/div[2]/div/a[{i}]/div/span[2]'
        news_titile_weibo.append(tree.xpath(xpath)[0].text)

        # wechat
        if i <= 25:
            xpath = f'//*[@id="node-5"]/div/div[2]/div/a[{i}]/div/span[2]'
            news_titile_wechat.append(tree.xpath(xpath)[0].text)

        # pengpai
        if i <= 20:
            xpath = f'//*[@id="node-51"]/div/div[2]/div/a[{i}]/div/span[2]'
            news_title_pengpai.append(tree.xpath(xpath)[0].text)
        
        # baidu
        xpath = f'//*[@id="node-2"]/div/div[2]/div/a[{i}]/div/span[2]'
        news_title_baidu.append(tree.xpath(xpath)[0].text)

        # toutiao
        if i <= 50:
            xpath = f'//*[@id="node-3608"]/div/div[2]/div/a[{i}]/div/span[2]'
            news_title_toutiao.append(tree.xpath(xpath)[0].text)
        
        # sougou
        if i <= 30:
            xpath = f'//*[@id="node-38"]/div/div[2]/div/a[{i}]/div/span[2]'
            news_title_sougou.append(tree.xpath(xpath)[0].text)
        

    url = f"https://tophub.today/c/news?p=2"
    response = requests.get(url=url, headers=headers)
    tree = etree.HTML(response.text)

    for i in range(1,51):
        # 360 news
        if i <= 40:
            xpath = f'//*[@id="node-69"]/div/div[2]/div/a[{i}]/div/span[2]'
            news_titile_360.append(tree.xpath(xpath)[0].text)
        
        # xinjingbao
        if i <= 10:
            xpath = f'//*[@id="node-2410"]/div/div[2]/div/a[{i}]/div/span[2]'
            news_titile_xinjingbao.append(tree.xpath(xpath)[0].text)


    # 将信息写入到 csv 文件中
    data_zhihu = pd.DataFrame({'热榜': news_titile_zhihu})
    data_weibo = pd.DataFrame({'热搜榜': news_titile_weibo})
    data_wechat = pd.DataFrame({'24h热文榜': news_titile_wechat})
    data_pengpai = pd.DataFrame({'热榜': news_title_pengpai})
    data_baidu = pd.DataFrame({'实时热点': news_title_baidu})
    data_toutiao = pd.DataFrame({'头条热榜': news_title_toutiao})
    data_sougou = pd.DataFrame({'实时热点': news_title_sougou})
    data_360 = pd.DataFrame({'实时热点榜单': news_titile_360})
    data_xinjingbao = pd.DataFrame({'排行': news_titile_xinjingbao})

    with pd.ExcelWriter('news.xlsx') as writer:
        data_zhihu.to_excel(writer, sheet_name='知乎热搜')
        print("------- 知乎热搜已经写入 -------")
        data_weibo.to_excel(writer,sheet_name='微博热搜')
        print("------- 微博热搜已经写入 -------")
        data_wechat.to_excel(writer,sheet_name='微信热文')
        print("------- 微信热文已经写入 -------")
        data_pengpai.to_excel(writer,sheet_name='澎湃热榜')
        print("------- 澎湃热榜已经写入 -------")
        data_baidu.to_excel(writer,sheet_name='百度实时热点')
        print("------- 百度实时热点已经写入 -------")
        data_toutiao.to_excel(writer,sheet_name='头条热榜')
        print("------- 头条热榜已经写入 -------")
        data_sougou.to_excel(writer,sheet_name='搜狗实时热点')
        print("------- 搜狗实时热点已经写入 -------")
        data_360.to_excel(writer,sheet_name='360实时热点')
        print("------- 360实时热点已经写入 -------")
        data_xinjingbao.to_excel(writer,sheet_name='新京报排行')
        print("------- 新京报排行已经写入 -------")




# file_path: e.g /root/data/news.xlsx
# 如果D:\\windows\\ 下面file_name的split需要调整一下
# upload_file 是为了生成 media_id， 供消息使用
def upload_file(file_path, wx_upload_url):
    file_name = file_path.split("/")[-1]
    with open(file_path, 'rb') as f:
        length = os.path.getsize(file_path)
        data = f.read()
    headers = {"Content-Type": "application/octet-stream"}
    params = {
        "filename": file_name,
        "filelength": length,
    }
    file_data = copy(params)
    file_data['file'] = (file_path.split('/')[-1:][0], data)
    encode_data = encode_multipart_formdata(file_data)
    file_data = encode_data[0]
    headers['Content-Type'] = encode_data[1]
    r = requests.post(wx_upload_url, data=file_data, headers=headers)
    print(r.text)
    media_id = r.json()['media_id']
    return media_id

# media_id 通过上一步上传的方法获得
def qi_ye_wei_xin_file(wx_url, media_id):
    headers = {"Content-Type": "text/plain"}
    data = {
        "msgtype": "file",
        "file": {
            "media_id": media_id
        }
    }
    r = requests.post(
        url=wx_url,
        headers=headers, json=data)
    print(r.text)


def push_report():
    test_report = './news.xlsx'
    wx_api_key = "xxxxx" # Webkey
    wx_upload_url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={}&type=file".format(wx_api_key)
    wx_url = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key={}'.format(wx_api_key)
    media_id = upload_file(test_report, wx_upload_url)
    qi_ye_wei_xin_file(wx_url, media_id)
    print("------- 消息已经推送 -------")

if __name__ == '__main__':
    page_news()
    push_report()
