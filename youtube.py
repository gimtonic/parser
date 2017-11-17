import urllib.request
import  json
import re
import win32com.client

Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(u'Путь\\1.xlsx')
sheet = wb.ActiveSheet

# Количество видео в плейлистах
N=[10,7,15,25,15,19,11,6,9,3,8,3,15,12,2,13]                                                  

tokens=["PLwV_Z7wfOJvtSpkK24hEC9jcoyfNXzu51","PLwV_Z7wfOJvsRU9WeOz7gG_2Vz3X0OT8b",
        "PLwV_Z7wfOJvvh_F2MCG5_D180n4wjfixp","PLwV_Z7wfOJvsYkOkoyizSkkdkg0NUy2HK",
        "PLwV_Z7wfOJvslTv5cXXi1hJ-7ox5wPT0j","PLwV_Z7wfOJvvzuNONMd7teaEbfsjLWpy4",
       "PLwV_Z7wfOJvvJaGXSVAklsnEFRapLiuC5","PLwV_Z7wfOJvuULGX4HoAx0GPTGt5qVWdJ",
       "PLwV_Z7wfOJvuy7UWf6S-Kdz8Z1m9KOcNX","PLwV_Z7wfOJvs6yUMJPeuUpTOjxKtjIwWq",
       "PLwV_Z7wfOJvtbtJ4rG9DZ_UKzYb5Zm_Wg","PLwV_Z7wfOJvuE9Dwkk4RxY64fFI1zwJOh",
       "PLwV_Z7wfOJvsorg6E0OJgmYtTAMpibgWl","PLwV_Z7wfOJvuGmGbVCHMJjATrK7aBejTv",
       "PLwV_Z7wfOJvumIoYklOZqYoIWWN6bjly6","PLwV_Z7wfOJvsSJNBwna0B0p8_z_xw5VjN"
       ]
kategory=["Создание и продвижение лендинга","SEO для бизнеса",
           "Семантическое ядро","Инструменты интернет-маркетинга",
           "Популярные видео","Продвижение ссылками",
          "SEO-тексты, копирайтинг","Коллтрекинг",
          "Зарубежное SEO","Защита SEO-кейсов ТопЭксперт",
          "Отдел продаж","E-mail рассылки, e-mail маркетинг",
          "Конверсия сайта","Передача \"5 кейсов\" от ТопЭксперт",
          "Круглый стол \"Настоящее и будущее SEO\" от ","SEO на Камчатке"
          ]

url_playlist3 = "https://www.googleapis.com/youtube/v3/playlistItems?part=snippet&key=ключ&maxResults=50"
j = 0
for token in tokens:
    i=0
    url3=url_playlist3 + "&playlistId=" + token
    with urllib.request.urlopen(url3) as url:
        data3 = json.loads(url.read().decode())
    print(url3)
    while i < N[j]:
        name = data3["items"][i]["snippet"]["description"] + "..."
        #Название оратора сохраняем как первые два слова description
        if name.replace(' ', '') != "...":
            name = data3["items"][i]["snippet"]["description"]
            name = re.findall(r'\w+', name)
            sheet.Cells(k, 1).value = name[0] + " " + name[1]
        else:
            sheet.Cells(k, 1).value = "Леонид Гроховский"
        title = data3["items"][i]["snippet"]["title"]
        sheet.Cells(k, 2).value = title
        link = data3["items"][i]["snippet"]["resourceId"]["videoId"]
        sheet.Cells(k, 3).value = "https://www.youtube.com/watch?v=" + link
        timedata = data3["items"][i]["snippet"]["publishedAt"]
        sheet.Cells(k, 4).value = timedata[0:10]
        description = data3["items"][i]["snippet"]["description"]
        sheet.Cells(k, 5).value = description[0:120] + "..."
        sheet.Cells(k, 7).value = kategory[j]
        sheet.Cells(k, 8).value = "Общий"
        i = i + 1
        k = k + 1
    j = j+1

wb.Save()
wb.Close()
Excel.Quit()       
