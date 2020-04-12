import requests
from lxml import html
import re
from queue import Queue
import threading
import time
import xlwt
import xlrd
from xlutils.copy import copy
import json
import random

HEADER = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.122 Safari/537.36'
}


class Procuder(threading.Thread):
    def __init__(self, page_queue, img_queue, free_proxy, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.page_queue = page_queue
        self.img_queue = img_queue
        self.free_proxy = free_proxy

    def run(self):
        while True:
            # 结束条件
            if self.page_queue.empty():
                break
            url = self.page_queue.get()
            print(url)
            self.page_urls(url)
            # 防止请求太快
            time.sleep(0.5)
        print(self.img_queue.qsize())

    def page_urls(self, url):
        reponse = requests.get(url, HEADER, proxies=self.free_proxy, stream=True)
        data = reponse.content.decode()
        # 在外面获取视频的时间和链接
        times = re.findall(r'<span title="上传时间" class="so-icon time">.*?</i>(.*?)</span>', data, re.S)
        href = re.findall(r'<li class="video-item matrix".*?<a href="(.*?)"', data, re.S)
        # 视频网页是一个动态网页,不能用以往静态网页的方法,所以我们要抓包来获取数据.
        # 检查网页-network-刷新网页-不难发现我们要获取的数据全在view?cid=…这个包里
        #而这个包的链接就在Headers中General的Request URL,这才是获取数据的链接
        
        # 获取动态网页地址
        url_cids = re.findall(r'<a href=".*?video/(.*?)?from', data, re.S)
        url_videos = []
        for url_cid in url_cids:
            # 取出来末尾会多一个问号，此步骤为了去除问号
            url_cid = url_cid[0:-1]
            # 'https://api.bilibili.com/x/web-interface/view?&bvid='是重点
            url_videos.append('https://api.bilibili.com/x/web-interface/view?&bvid=' + url_cid)
        self.urlls(url_videos, times, href)

    def urlls(self, urlss, times, href):
    # 将爬取下的数据放到列表中
        for index, url in enumerate(urlss):
            dd = []
            response = requests.get(url, HEADER, proxies=self.free_proxy)
            data = response.content.decode()
            content = json.loads(data)
            title = content['data']['title']
            tname = content['data']['tname']
            author = content['data']['owner']['name']
            info = content['data']['desc']
            info = re.sub(r'\n| ', '', info)
            aid = 'av' + str(content['data']['aid'])
            view = content['data']['stat']['view']
            coin = content['data']['stat']['coin']
            share = content['data']['stat']['share']
            like = content['data']['stat']['like']
            favorite = content['data']['stat']['favorite']
            dd.append(title)
            hrefs = 'https:' + href[index]
            dd.append(hrefs)
            times[index] = re.sub(r'\n| ', '', times[index])
            dd.append(times[index])
            dd.append(tname)
            dd.append(author)
            dd.append(aid)
            dd.append(view)
            dd.append(coin)
            dd.append(share)
            dd.append(like)
            dd.append(favorite)
            dd.append(info)
            # 将列表放入队列中
            self.img_queue.put(dd)
            # 及时关闭，防止资源损耗
            response.close()
            break


class Consumer(threading.Thread):
    def __init__(self, page_queue, img_queue, gLock, i, *args, **kwargs):
        super(Consumer, self).__init__(*args, **kwargs)
        self.page_queue = page_queue
        self.img_queue = img_queue
        self.gLock = gLock
        self.i = i

    def run(self):
        dd = 1
        # 定义写入数据的开始行数，并逐次递增
        while True:
        # 这里休眠10秒，防止因为网速延迟问题而导致数据不完整
            if self.img_queue.empty() and self.page_queue.empty():
                time.sleep(10)
            if self.img_queue.empty() and self.page_queue.empty():
                break
            img = self.img_queue.get()
            # 上锁，防止其它干扰
            self.gLock.acquire()
            # 写入表格
            self.export_excel(img, dd)
            self.gLock.release()
            dd += 1

    def export_excel(self, next, xx):
    # 在这里打开之前的excel表格，复制到新的文件中，防止文件被覆盖
        k = 'D:\\表格\\'
        oldWb = xlrd.open_workbook(k + self.i + '.xls')
        newWb = copy(oldWb)
        newWs = newWb.get_sheet(0)
        for j, col in enumerate(next):
            newWs.write(xx, j, col)
        # 保存文件
        newWb.save(k + self.i + '.xls')
        # 输出行数，便于随时监测
        print(xx)


def main():
    gLock = threading.Lock()
    page_queue = Queue(500)
    img_queue = Queue(2000)
    dd = ['简历模板']
    for i in dd:
        # 获取西刺代理ip并检测代理ip是否可用
        response = requests.get('https://www.xicidaili.com/nn/', headers=HEADER)
        data = response.content.decode('utf-8', errors='ignore')
        htmls = html.etree.HTML(data)
        text = htmls.xpath('//tr[@class="odd"]//td[2]//text()')
        tet = htmls.xpath('//tr[@class="odd"]//td[3]//text()')
        lb = []
        for index, url in enumerate(text):
            zd = {}
            dd = url + ':' + tet[index]
            zd['http'] = dd
            lb.append(zd)
        # 检测代理ip是否可用，不可用则剔除
        while True:
            free_proxy = random.choice(lb)
            try:
                requests.get('http://www.baidu.com', headers=HEADER, proxies=free_proxy, timeout=1)
            except:
                lb.remove(free_proxy)
                continue
            break
        print(free_proxy)
        # 先创建一个excel表格，存放表头
        k = 'D:\\表格\\'
        books = xlwt.Workbook(encoding="utf-8")
        sheet = books.add_sheet(i)
        order = ['视频名称', '视频链接', '发行时间', '分区', 'UP主', 'AV号', '播放数', '点赞数', '投硬币数', '收藏数', '分享数', '视频简介']
        for j, col in enumerate(order):
            sheet.write(0, j, col)
        books.save(k + i + '.xls')
        # 写入搜索内容
        url = 'https://search.bilibili.com/all?keyword={}'.format(i)
        reponse = requests.get(url, HEADER, proxies=free_proxy)
        data = reponse.content.decode()
        htmls = html.etree.HTML(data)
        # 统计页数
        shuzi = htmls.xpath('normalize-space(//*[@id="all-list"]/div[1]/div[3]/div/ul/li[8]/button/text())')
        ss = int(shuzi)
        # 将每一页储存在队列中
        for o in range(1, ss + 1):
            url = 'https://search.bilibili.com/all?keyword={}&page={}'.format(i, o)
            page_queue.put(url)
        # 创建生产线程
        for x in range(4):
            t = Procuder(page_queue, img_queue, free_proxy)
            t.start()
        # 这里创建一个线程，因为要写入excel表格，创建多个会出现错误
        w = Consumer(page_queue, img_queue, gLock, i)
        w.start()


if __name__ == '__main__':
    main()
