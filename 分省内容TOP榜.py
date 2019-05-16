import pandas as pd
from openpyxl import Workbook

x=pd.read_excel(r"C:\Users\Administrator\Desktop\1 咪视界内容排行走势_日_表1.xlsx")
x1=x[x['分类']=='剔重汇总']
means=x1['播放次数'].groupby(x1['渠道省份']).mean()
mean1=means[(means.index!='剔重汇总')&(means.index!='全国')&(means.index!='未知')]
mean2=list(mean1.sort_values(ascending=False).index[:10])

for i in mean2:
    
