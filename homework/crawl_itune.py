from urllib.request import urlopen
from bs4 import BeautifulSoup
import pyexcel
from youtube_dl import YoutubeDL


options = {
    "default_search" : "ytsearch",
    "max_download" : 1,
}

url = "https://www.apple.com/itunes/charts/songs/"
conn = urlopen(url)
raw_data = conn.read()
html_content = raw_data.decode("utf8")

soup = BeautifulSoup(html_content, "html.parser")
section = soup.find("section", "section chart-grid")
li_list = section.find_all("li")

list_of_dictionary = []
for li in li_list:
    c = li.a
    link = c["href"] 
    h3 = li.h3
    a = h3.a
    title = a.string.strip()
    h4 = li.h4
    b = h4.a
    artist = b.string.strip()
    YoutubeDL(options).download([title])
    
    list_of_dictionary.append(
        {
        "title":title,
        "link":link,
        "artist":artist
    }
    )

pyexcel.save_as(records=list_of_dictionary, dest_file_name="itunes.xls")