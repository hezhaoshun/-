# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pandas as pd
x=pd.read_excel(r"C:\Users\Administrator\Desktop\1 咪视界内容排行走势_日_表1.xlsx")

x1=x[x['渠道省份']=='剔重汇总'].loc[:,['统计日期',
         '分类',
         '内容名称',
         '播放类型',
         '排名',
         '播放用户数',
         '有效播放用户数',
         '播放次数',
         '有效播放次数',
         '播放时长（小时）',
         '人均播放次数',
         '人均播放时长(小时)',
         '次均播放时长（小时）',
         '达到率']]
cc=['电影',
    '电视剧',
    '动漫',
    '少儿',
    '纪实',
    '综艺',
    '体育']
x2=pd.DataFrame(columns=x1.columns)
for c in cc:
    x2=x2.append(x1[x1['分类']==c].sort_values('播放次数',ascending=False)[:50])
x2=x2.append(x1[x1['分类']=='剔重汇总'].sort_values('播放次数',ascending=False)[:100])
x2.to_excel(r"C:\Users\Administrator\Desktop\排行榜.xlsx",index=False)