import requests
from bs4 import BeautifulSoup
import xlwt
import time

#将二维列表写入Excel表
def WriteExcel(List,Name,SmallName):
    workbook = xlwt.Workbook(encoding="utf-8")
    worksheet = workbook.add_sheet(SmallName)
    for i in range(0, len(List)):
        for j in range(0, len(List[i])):
            worksheet.write(i, j, label=List[i][j])
    workbook.save(Name)

Type=["情感励志","搞笑趣闻","娱乐","旅游","运动","医疗健康","数码科技","汽车","餐饮美食","时尚","房产","母婴","生活常识","时事资讯","政务","财经","地方","职场教育","早教幼教","小学教育","中学教育","大学校园"]
Num=[35,27,31,97,1500,29,30,28,32,33,128,36,1434,26,39952,1443,37,34,42084,42095,42086,39953]

InUse=[]
for n in range(0,len(Type)):
    print("正在",Type[n])
    All=[]
    for i in range(1,9):
        try:
            Final=[]
            url='https://zs.xiguaji.com/BizRank/{}/{}/2020040{}'.format(Type[n],Num[n],i)
            resp = requests.get(url)
            a = resp.content.decode('utf-8')
            soup = BeautifulSoup(a)
            Final=soup.find_all("div",class_="rankMpName")
            for final in Final:
                All.append(str(final)[24:str(final).find("<em>")])
        except:
            print()
    Deduplication=list(set(All))
    for i in range(0,len(Deduplication)):
        N=0
        for j in range(0,len(All)):
            if All[j]==Deduplication[i]:
                N=N+1
        InUse.append([Deduplication[i],N,Type[n]])

print(InUse)
WriteExcel(InUse,"4月超优质公众号.xls",'公众号')