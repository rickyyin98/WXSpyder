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
InUse=[]
for type in Type:
    print("正在",type)
    All=[]
    for i in range(1,9):
        try:
            Final=[]
            url='https://zs.xiguaji.com/BizRank/{}/35/2020040{}'.format(type,i)
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
        InUse.append([Deduplication[i],N,type])

print(InUse)
WriteExcel(InUse,"4月超优质公众号.xls",'公众号')