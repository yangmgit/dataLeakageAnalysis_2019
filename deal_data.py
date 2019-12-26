# -*- coding: utf-8 -*-
import pandas as pd
data = pd.read_excel('data.xlsx')

#泄露原因统计：
data_reason  = data[['泄露原因','泄露事件']].groupby(['泄露原因']).count().reset_index(drop=False).sort_values(by=['泄露事件'], ascending=False).reset_index(drop=True)
data_reason.columns = ['泄露原因', '泄露事件数']


#泄露国家统计
data_country = data[['泄露国家', '泄露事件']].groupby(['泄露国家']).count().reset_index(drop=False).sort_values(by=['泄露事件'], ascending=False).reset_index(drop=True)
data_country.columns = ['泄露国家', '泄露事件数']


#泄露行业统计
data_industry = data[['泄露行业', '泄露事件']].groupby(['泄露行业']).count().reset_index(drop=False).sort_values(by=['泄露事件'], ascending=False).reset_index(drop=True)
data_industry.columns = ['泄露行业', '泄露事件数']

#每月泄露数据
data_month = data[['月份', '泄露事件']].groupby(['月份']).count().reset_index(drop=False).sort_values(by=['月份'], ascending=True).reset_index(drop=True)
data_month.columns = ['月份', '泄露事件数']
data_month['月份'] = data_month['月份'].astype('str') + '月'

#写入excel使用power bi进行可视化
    #创建空表格
excelWriter = pd.ExcelWriter('data_result.xlsx')
data_country.to_excel(excelWriter, sheet_name = "泄露国家统计", index=False, encoding="utf-8")
data_industry.to_excel(excelWriter, sheet_name = "泄露行业统计", index=False, encoding="utf-8")
data_reason.to_excel(excelWriter, sheet_name = "泄露原因统计", index=False, encoding="utf-8")
data_month.to_excel(excelWriter, sheet_name = "泄露月份统计", index=False, encoding="utf-8")
excelWriter.save()