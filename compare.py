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

def email():
	#配置邮箱
	server = zmail.server('username','password','host')
	#收件人
	recipients=[('别名','test@com')]
	#抄送人
	cc=[('别名','test@com')]
	#邮件主题
	subject='数据对比测试'
	#邮件正文
	contents = 'Dear all：\n\n数据对比结果见附件，请核查！'
	#邮件附件
	filename='测试结果.xlsx'
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

def Mysql_Example(sql):
	#创建mysql连接
	con=pymysql.connect(host='XXX',port=3306,user='XXX',passwd='XXX',db='XXX',use_unicode=True, charset="utf8")
	#查询数据库结果
	result=pd.read_sql(sql,con)
	con.close()
	return result

def Oracle_Example(sql):
	#创建mysql连接
	con = cx_Oracle.connect('user/passwd@host:port/db')
	#查询数据库结果
	result=pd.read_sql(sql,con)
	con.close()
	return result

if __name__ == "__main__":
	try:
		#获取excel数据
		excel=pd.read_excel('sql.xlsx').values.tolist()
		for i in range(len(excel)):
			#修改测试数据库地址
			test_result=Oracle_Example(excel[i][1]).values.tolist()
			#修改结果数据库地址
			bi_result=Mysql_Example(excel[i][2]).values.tolist()
			row_error=''
			#验证记录数是否一致
			if len(test_result)!=len(bi_result):
				excel[i][3]='总记录数不一致！'
			else:
				#每行记录对比
				for j in range(len(test_result)):
					if str(test_result[j])!=str(bi_result[j]):
						row_error=row_error+str(j+1)+'、'
				#判断对比结果
				if row_error=='':
					excel[i][3]='验证通过！'
				else:
					row_error='第 '+row_error+'行对比不一致！'
					excel[i][3]=row_error
		#list转DataFrame
		result=DataFrame(excel)
		#添加列名
		result.columns = ["序号", "测试数据库SQL", "结果数据库SQL", "测试结果"]
		#导出到excel
		filename='测试结果.xlsx'
		with pd.ExcelWriter(filename) as writer:
			result.to_excel(writer, sheet_name='Sheet1',index=False)
			print('导出Excel成功！')
		#发送邮件
		email()

	except Exception as e:
		print(str(e))












