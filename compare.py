#!/usr/bin/python3
# -*- coding:utf-8 -*-
'''
Created on 2021年5月20日
@author: yuejing
'''
import zmail
import pymysql
import cx_Oracle
import pandas as pd
from pandas.core.frame import DataFrame

def email(filename=0):
	server = zmail.server('username','password','host')
	recipients=[('别名','test@com')]
	cc=[('别名','test@com')]
	subject='数据对比测试'
	if filename==0:
		contents = 'Dear all：\n\n数据对比行数不一致，请核查！'
		mail = {
	    'subject': subject,  
	    'content_text': contents}
	else:
		contents = 'Dear all：\n\n数据对比存在不一致的结果，附件为不一致记录，请核查！'
		mail = {
		    'subject': subject,  
		    'content_text': contents,  
		    'attachments': filename	}
	try:
		server.send_mail(recipients,mail,cc)
		print('邮件发送成功!')
	except Exception as e:
		print('邮件发送失败!')
		raise e

def read_txt(file_name) :
	with open(file_name,"r",encoding='utf-8') as f:
		text = f.read()
		f.close()
	return text

def MysqlData():
	#读取SQL
	sql=read_txt('mysql.txt')
	#创建mysql连接
	con=pymysql.connect(host='XXX',port=3306,user='XXX',passwd='XXX',db='XXX',use_unicode=True, charset="utf8")
	#查询数据库结果
	result=pd.read_sql(sql,con)
	con.close()
	return result

def OracleData():
	#读取SQL
	sql=read_txt('oracle.txt')
	#创建mysql连接
	con = cx_Oracle.connect('user/passwd@host:port/db')
	#查询数据库结果
	result=pd.read_sql(sql,con)
	con.close()
	return result

if __name__ == "__main__":
	try:

		#DataFrame转换成List
		data1=MysqlData().values.tolist()
		data2=OracleData().values.tolist()
		data1_error=[]
		data2_error=[]
		#验证记录数是否一致
		if len(data1)!=len(data2):
			print('总记录数不一致！')
			email()
		else:
			#每行记录对比
			for i in range(len(data1)):
				if str(data1[i])!=str(data2[i]):
					data1_error.append(data1[i])
					data2_error.append(data2[i])
			#判断对比结果
			if len(data1_error)==0:
				print('数据正常！')
			else:
				#list转成DataFrame
				data1_result=DataFrame(data1_error)
				data2_result=DataFrame(data2_error)
				filename='数据对比不一致结果.xlsx'
				with pd.ExcelWriter(filename) as writer:
					data1_result.to_excel(writer, sheet_name='Sheet1',index=False,header=None)
					data2_result.to_excel(writer, sheet_name='Sheet2',index=False,header=None)
					print('导出Excel成功！')
				email(filename)

	except Exception as e:
		raise e





