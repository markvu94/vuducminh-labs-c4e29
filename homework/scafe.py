from urllib.request import urlopen
from bs4 import BeautifulSoup
import pyexcel
from collections import OrderedDict

url = "http://s.cafef.vn/bao-cao-tai-chinh/VNM/IncSta/2017/3/0/0/ket-qua-hoat-dong-kinh-doanh-cong-ty-co-phan-sua-viet-nam.chn"
conn = urlopen(url)
raw_data = conn.read()
html_content = raw_data.decode ("utf8")

with open("scafe.html","wb") as f:
    f.write(raw_data)

soup = BeautifulSoup(html_content, "html.parser")
table_title = soup.find("table", style ="border-top: solid 1px #e6e6e6;border-left:solid 1px #cccccc;")
td_title_list = table_title.find_all("td")
json_title = []

for td in td_title_list:
    if td["class"] == ["b_r_c"]:
        title = td.div.a.string.strip()  #luu y: day la khi chon Truoc, neu chon Sau se phai chon a so 2 (create list)
        json_title.append(title)        
    if td["class"] == ["h_t"]:
        title = td.string.strip()
        json_title.append(title)

json_final = []  

table_content = soup.find("table", id ="tableContent")
table_content_list = table_content.find_all("tr")
for tr in table_content_list:
    json_content = [] 
    td_content_list = tr.find_all("td")
    for td in td_content_list:
        if not td.string is None:
            content = td.string.strip()
            json_content.append(content)
        else:
            content = ""
            json_content.append(content)
    if json_content != [] and ['', '', '', '', '', '', '']:
        json_final.append(OrderedDict(
            {
                json_title[0]:json_content[0],
                json_title[1]:json_content[1],
                json_title[2]:json_content[2],
                json_title[3]:json_content[3],
                json_title[4]:json_content[4],
            })
        )
pyexcel.save_as(records=json_final, dest_file_name="baocaotaichinh.xls")          










