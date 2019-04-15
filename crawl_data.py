from urllib.request import urlopen
from bs4 import BeautifulSoup
import pyexcel
# import collections


#1: open connection
url = "https://dantri.com.vn"
conn = urlopen(url)
raw_data = conn.read()
html_content = raw_data.decode("utf8")

with open("dantri.html","wb") as f:
    f.write(raw_data)

#2: Find ROI (region of internet)
soup = BeautifulSoup(html_content, "html.parser")
ul = soup.find("ul", "ul1 ulnew")

# #3: Extract ROI
li_list = ul.find_all("li")

list_of_dictionary = []
for li in li_list:

    h4 = li.h4
    a = h4.a
    title = a.string.strip()
    link = url + a["href"]
    list_of_dictionary.append(
        {
        "title":title,
        "link":link,
    }
    )
print(list_of_dictionary)

pyexcel.save_as(records=list_of_dictionary, dest_file_name="your_file.xls")





    



