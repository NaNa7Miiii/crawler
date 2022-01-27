# 一般python程序第一行需要加入
    # -*-coding:utf-8 -*- 或者 # coding = utf-8
    # 这样可以在代码中包含中文
# python文件中可以加入main函数用于测试程序
    # if__name__ == "__main__":
'''
def main():
    print("hello")

if __name__ == "__main__": # 当程序执行时
# 调用函数
    main()

# import bs4 网页解析，获取数据 from bs4 import BeautifulSoup更好
# import re 正则表达式，进行文字匹配
# import urllib.request, urllib.error 制定URL,获取网页数据
# import xlwt 进行excel操作
# import sqlite3 #进行SQLite数据库操作
'''

import urllib.request
from bs4 import BeautifulSoup
import re
import xlwt
import sqlite3

def main():
    baseurl = "https://movie.douban.com/top250?start="
    # 1.爬取网页
    datalist = getData(baseurl)
    # savepath = "豆瓣电影TOP250.xls"
    dbpath = "movie.db"
    # 3.保存数据
    # saveData(datalist, savepath)
    saveData2DB(datalist, dbpath)
    # askURL("https://movie.douban.com/top250?start=")

# link of movies
findlink = re.compile(r'<a href="(.*?)">') # 创建正则表达式对象, 表示规则(字符串的模式）
# 影片图片
findImgSrc = re.compile(r'<img.*src="(.*?)"', re.S) # re.S 让换行符包含在字符中
# 影片片名
findTitle = re.compile(r'<span class="title">(.*)</span>')
# 影片评分
findRating = re.compile(r'span class="rating_num" property="v:average">(.*)</span>')
# 找到评价人数
findJudge = re.compile(r'<span>(\d*)人评价</span>')
# 找到概况
findInq = re.compile(r'<span class="inq">(.*)</span>')
# 找到影片的相关内容
findBd = re.compile(r'<p class="">(.*?)</p>', re.S)

# 爬取网页
def getData(baseurl):
    datalist = []
    for i in range(0, 10):
        url = baseurl + str(i*25)
        html = askURL(url) # store the source code of web
        # 逐一解析数据:
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.find_all('div', class_="item"): # 查找符合要求的字符串，形成列表
            data = [] # 保存一部电影的所有信息
            item = str(item)

            link = re.findall(findlink, item)[0]  #re库用来通过正则表达式查找指定的字符串
            data.append(link)

            imgSrc = re.findall(findImgSrc, item)[0] #添加图片
            data.append(imgSrc)

            titles = re.findall(findTitle, item)
            if len(titles) == 2:
                chinesetitle = titles[0]
                data.append(chinesetitle) # 添加中文名
                othertitle = titles[1].replace("/","") # 去掉无关的符号
                data.append(othertitle) # 添加外国名
            else:
                data.append(titles[0])
                data.append(' ') # 因为外文名没有，所以留空不能让下面内容占位

            rating = re.findall(findRating, item)[0]
            data.append(rating) # 添加评分

            judgeNum = re.findall(findJudge, item)[0]
            data.append(judgeNum) # 添加评价人数

            inq = re.findall(findInq, item)
            if len(inq) != 0:
                inq = inq[0].replace("。", "") # 去掉句号
                data.append(inq)
            else:
                data.append(" ") # 留空

            bd = re.findall(findBd, item)[0]
            bd = re.sub('<br(\s+)?/>(\s+)?', " ", bd) # 去掉<br/>
            bd = re.sub('/', " ", bd) # 替换/
            data.append(bd.strip()) # 去掉前后空格

            datalist.append(data) # 把处理好的一部电影信息放入datalist
    # print(datalist)
    return datalist

# 得到指定一个URL的网页内容
def askURL(url):
    head = { # 模拟浏览器头部信息，向豆瓣服务器发送消息
        "User-Agent": "Mozilla / 5.0(Macintosh; Intel Mac OS X 10_15_7) AppleWebKit / 537.36(KHTML, like Gecko) Chrome / 95.0.4638.69 Safari / 537.36"
    }
    # 用户代理：表示告诉豆瓣服务器，我们是什么类型的机器。本质上是告诉浏览器，我们可以接受什么水平的文件内容
    request = urllib.request.Request(url, headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(request)
        html = response.read().decode("utf-8")
        # print(html)
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html

# 3.保存数据
def saveData(datalist, savepath):
    print("save...")
    book = xlwt.Workbook(encoding="utf-8")  # 创建workbook对象
    sheet = book.add_sheet('豆瓣电影Top250', cell_overwrite_ok=True)  # 创建工作表
    col = ('电影详情链接','图片链接','影片中文名','影片外国名','评分','评价数','概况','相关信息')
    for i in range(0,8):
        sheet.write(0,i,col[i])
    for i in range(0,250):
        print("第%d条" %(i+1))
        data = datalist[i]
        for j in range(0,8):
            sheet.write(i+1,j,data[j]) # 数据

    book.save(savepath) # 保存

def saveData2DB(datalist, dbpath):
    init_db(dbpath)
    conn = sqlite3.connect(dbpath)
    cur = conn.cursor()
    for data in datalist:
        for index in range(len(data)):
            # if index == 4 or index == 5:
                # continue
            data[index] = '"' + data[index] + '"'
        sql = '''
                insert into movie250(
                    info_link, pic_link, chinesetitle, othertitle, score, rated, introduction, info)
                    values(%s)
            '''%",".join(data)
        cur.execute(sql)
        conn.commit()
    cur.close()
    conn.close()

def init_db(dbpath):
    sql = '''
        create table movie250
        (
        id integer primary key autoincrement,
        info_link text,
        pic_link text,
        chinesetitle varchar,
        othertitle varchar,
        score numeric,
        rated numeric,
        introduction text,
        info text
        )
    ''' # 创建数据表
    conn = sqlite3.connect(dbpath)
    cursor = conn.cursor()
    cursor.execute(sql)
    conn.commit()
    conn.close()


# askURL("https://movie.douban.com/top250?start=0")
# if __name__== "__main__":
#     main()
#     print("爬取完毕")

if __name__ == '__main__':
    main()
    print("爬取完毕")

# main()