#! /usr/bin/env python
# -*- coding: UTF-8 –*-
import urllib
import urllib2
import cookielib
import json
import sys
import time
import xlwt
import re

reload(sys)
sys.setdefaultencoding('utf8')

# Select the events to monitor:
# 
# 1-normal 2-important 3-very important
# 
# cid[]=1  meaning foreign exchange
# cid[]=2  meaning stock
# cid[]=9  meaning China
# cid[]=10 meaning American
# 
# such as:
# all event(all country)                      - event = ""
# important and very important(all country)   - event = "importance=2,3"
# very important(all country)                 - event = "importance=3"
# very important in China                     - event = "importance=3&cid[]=9"

# select important and very important event in China
event = ""

# name of file which used for store result, mush be *.xls and *.xlsx
filename = "data.xls"

class WallStreetCnSpider:
    """docstring for WallStreetCnSpider"""

    def __init__(self):
        self.count = 0
        self.reqindex = 0
        self.sheetindex = 1
        self.nextRow = 0
        self.first = True
        # file for save result
        self.filename = filename
        self.initurl = 'http://api.wallstreetcn.com/v2/livenews?limit=100&' + event + '&callback=jQuery213046828437503427267_1433944460455'
        self.url = self.initurl
        self.cookie = cookielib.CookieJar()
        self.opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(self.cookie))

        if self.filename[self.filename.rfind('.')+1:] != 'xls' and self.filename[self.filename.rfind('.')+1:] != 'xlsx':
            print 'Wrong format of filename, please select *.xls and *.xlsx to store result'
            return

        self.excel = xlwt.Workbook(encoding='utf-8')
        self.table = self.excel.add_sheet('wallstreet' + str(self.sheetindex), cell_overwrite_ok=True)

    def getSpiderCount(self):
        respdata = self.opener.open(self.url).read()
        lbracket = respdata.find('(')
        rbracket = respdata.rfind(')') 

        # get json data
        jsondata = json.loads(respdata[lbracket + 1:rbracket])

        # parse the total number of request which we should send 
        lasturl = jsondata['paginator']['last']
        return long(lasturl[lasturl.find('page')+5:lasturl.find('channelId')-1])

    def request(self):
        try:
            respdata = self.opener.open(self.url).read()

            # we should cut out first bracket when first request response, which consist in response string
            if self.first:
                lbracket = respdata.find('(')
                rbracket = respdata.rfind(')') 
                self.first = False
                respdata = respdata[lbracket + 1:rbracket]

            # get json data
            jsondata = json.loads(respdata)
        except ValueError:
            # print e.reason
            # raise e
            print 'end of request\n'
            return
        except urllib2.URLError as e:
            print e.reason
            return

        # parse next url and assign to self.url for next spider
        nexturl = jsondata['paginator']['next']
        self.url = urllib.unquote(nexturl) + '&limit=100'

        self.reqindex += 1
        print str(self.reqindex) + '-th crawl, got items : 100'
        return jsondata['results']

    def parseData(self, results):
        if results is None:
            return

        # write data to excel
        for data in results:

            createTime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(float(data['createdAt'])))
            important = data['importance']
            content = data['contentHtml']
            country = ''
            asset = ''
            cataset = data['categorySet']
            if cataset is not None:
                catagory = cataset.split(',')
                if catagory is not None:       
                    if '9' in catagory:
                        country = '中国'
                    elif '10' in catagory:
                        country = '美国'
                    elif '11' in catagory:
                        country = '欧元区'
                    elif '12' in catagory:
                        country = '日本'
                    elif '13' in catagory:
                        country = '英国'
                    elif '14' in catagory:
                        country = '澳洲'
                    elif '15' in catagory:
                        country = '加拿大'
                    elif '16' in catagory:
                        country = '瑞士'
                    else:
                        country = "其它"
            
                if catagory is not None:
                    if '1' in catagory:
                        asset = '外汇'
                    elif '2' in catagory:
                        asset = '股市'
                    elif '3' in catagory:
                        asset = '商品'
                    elif '4' in catagory:
                        asset = '债市'
                    else:
                        asset = '其它'

                
            # cut out <p> and </p>
            content = re.compile(r'\<\/?p\>').sub('', content)

            try:
                if self.nextRow >= 60000:
                    # reached the maximum of single sheet, create new sheet
                    self.sheetindex += 1
                    self.table = self.excel.add_sheet('wallstreet' + str(self.sheetindex), cell_overwrite_ok=True)
                    self.nextRow = 0

                self.table.write(self.nextRow, 0, createTime)
                self.table.write(self.nextRow, 1, important)
                self.table.write(self.nextRow, 2, country)
                self.table.write(self.nextRow, 3, asset)
                self.table.write(self.nextRow, 4, content)

                self.nextRow += 1
                self.count += 1
            except ValueError, e:
                print e.reason

    def run(self):
        print 'data is being crawl, be patient...'
        print self.getSpiderCount()
        for i in range(1, self.getSpiderCount()):
            self.parseData(self.request())
        print 'grab ' + str(self.count) + ' items'
        self.excel.save(self.filename)
        print 'crawl data successfully'

spider = WallStreetCnSpider()
spider.run()
