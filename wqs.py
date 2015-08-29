# -*- coding: utf-8 -*-
"""
Created on Sat Apr 11 14:44:48 2015

@author: wqs
"""
import urllib.request
import xlwt
import os
import time
#http://quote.eastmoney.com/center/list.html#2850016_0
def getdata():
	print("抓取中，当前时间为："+time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
	book=xlwt.Workbook(encoding='utf-8',style_compression=0)
	sheet=book.add_sheet('wqs',cell_overwrite_ok=True)
	title=['代码','名称','最新价','涨跌额','涨跌幅','振幅','成交量(手)','成交额(万)','昨收','今开','最高','最低','5分钟涨跌','抓取时间']
	for i in range(len(title)):
		sheet.write(0,i,title[i])
	xuhao=1  #像表格中写入数据
	for i in range(3):
		url='http://hqdigi2.eastmoney.com/EM_Quote2010NumericApplication/index.aspx?type=s&sortType=C&sortRule=-1&pageSize=20&page=%d&jsName=quote_123&style=2850016&token=44c9d251add88e27b65ed86506f6e5da&_g=0.6199398094322532' %(i+1)
		data=urllib.request.urlopen(url).read().decode('utf-8')
		x=data.index('[')
		y=data.index(']')
		numdata=data[x+1:y]
		z=str(numdata).strip('"').split('"')
		m=z[::2]
		for n in m:
			#print(n,end='\n')
			shuju=n.split(',')#shuju为['6013901', '601390', '中国中铁', '16.36', '16.29', '18.00', '18.00', '16.02', '1929438', '10943502', '1.64', '10.02%', '17.63', '12.10%', '100.00%', '135514', '35', '6063823', '4879679', '-1', '0', '0.22%', '1.26', '6.40%', '37.50', '001150|002425|003499|003500|003505|003510|003535|003587|003592|003596|003612|003702|003707|003712|5009|50016', '18.00', '0.00', '2015-04-17 15:03:06', '0', '17092509696', '279633469059', '4.8']
			realshuju=shuju[1:3]
			realshuju.append(shuju[5])
			realshuju.append(shuju[10])
			realshuju.append(shuju[11])
			realshuju.append(shuju[13])
			realshuju.append(shuju[9])
			realshuju.append(shuju[8])
			realshuju.append(shuju[3])
			realshuju.append(shuju[4])
			realshuju.append(shuju[6])
			realshuju.append(shuju[7])
			realshuju.append(shuju[21])
			realshuju.append(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
			#print(realshuju)
			for j in range(len(realshuju)):
				sheet.write(xuhao,j,realshuju[j])
			xuhao=xuhao+1
	book.save('.\wqs.xls')
	print("抓取结束，下一次抓取将在十分钟后！")

def main():
	while True: 
		getdata()
		time.sleep(600)

if __name__ == '__main__':
	main()
