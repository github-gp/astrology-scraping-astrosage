#Import Libraries
import pandas as pd
import webbrowser, sys, pyperclip
import requests, bs4
from selenium import webdriver
from selenium.webdriver.common.by import By
import string

#Craete a list of all the personality on astrosage

alphabet = list(string.ascii_uppercase)
list1=[]
for k in range(26):
    try:
        for i in range(1,100):
            url = ('http://www.astrosage.com/celebrity-horoscope/default.asp?page='+str(i)+'&ch='+alphabet[k])
# Create object page
            page = requests.get(url)
            soup = bs4.BeautifulSoup(page.text, 'lxml')
            table1 = soup.find('table', id='')
            texts = table1.find_all('a')
            list2=[]
            for j in [1,4,7,10,13,16,19,22,25,28,31,34,37,40,43,46,49,52,55,58]:
                list2.append(texts[j].text)
            list1.extend(list2)
    except:
        continue    
        
# Gather the data of all the personality collected above in list1

df = pd.DataFrame(columns = ['Name', 'About','Date of Birth', 'Time of Birth','Place of Birth','Longitude','Latitude','Information Source'])
for i in range(len(list1)):
    try:
        text= list1[i]
        url="http://www.astrosage.com/celebrity-horoscope/"
        res = requests.get(url+text.lower().replace(" ", "-")+"-horoscope.asp")
        res.raise_for_status()
        astrosageSoup = bs4.BeautifulSoup(res.text, 'html.parser')
        df=df.append({'Name':astrosageSoup.select('b')[2].next_sibling.strip(),
              'About':astrosageSoup.find_all("div", {"class": "celebcont"})[0].get_text().strip(),
              'Date of Birth': astrosageSoup.select('b')[3].next_sibling.strip(),
               'Time of Birth': astrosageSoup.select('b')[4].next_sibling.strip(),
               'Place of Birth':astrosageSoup.select('b')[5].next_sibling.strip(),
               'Longitude':astrosageSoup.select('b')[6].next_sibling.strip(),
               'Latitude':astrosageSoup.select('b')[7].next_sibling.strip(),
               'Information Source':astrosageSoup.select('b')[9].next_sibling.strip()}
              ,ignore_index=True)
    except:
        continue
        
#Write the data to excel file
writer = pd.ExcelWriter('astro_data_bank.xlsx')
df.to_excel(writer)
writer.save()
print('DataFrame is written successfully to Excel File.')