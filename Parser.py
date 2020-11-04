from bs4 import BeautifulSoup
import requests
import xlsxwriter

book = xlsxwriter.Workbook("Music.xlsx")

class Parser(object):

    def __init__(self):
        self.comps = []

    def main(self, start, end, url):
        for n in range(start, end):
            print(n)
            Parser().parse(url + f"page/{n}/", n)

        book.close()

    def parse(self, url, page):
        HEADERS = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36 OPR/71.0.3770.284 (Edition Yx 02)'
        }

        datas = []

        singer = []
        name = []
        link = []
        date = []
        quality = []
        duration = []

        response = requests.get(url, headers=HEADERS)
        soup = BeautifulSoup(response.content, 'html.parser')

        items = soup.find_all('div', class_ = 'music-popular__item')

        for item in items:
            for data in item.find_all('div', class_ = 'popular-download'):
                for i in data.find_all('div', class_ = 'popular-download-date'):
                    datas.append(i.get_text(strip = True))

        for i in range(len(datas)):
            if i%3 == 0:
                date.append(datas[i])
            elif i%3 == 1:
                quality.append(datas[i])
            else:
                duration.append(datas[i])

        for item in items:
            singer.append(item.find('a', class_ = 'popular-play-composition').get_text(strip = True))
            name.append(item.find('a', class_ = 'popular-play-author').get_text(strip = True))
            link.append(item.find('a', class_ = 'popular-play__item').get('data-url'))

        self.comps = zip(singer, name, duration, date, quality, link)

        Parser().saveExcel(page, list(self.comps))

    def saveExcel(self, page, comps):
        sheet = book.add_worksheet("Music" + str(page))
        sheet.set_column(0, 0, 8)
        sheet.set_column(1, 1, 40)
        sheet.set_column(2, 2, 50)
        sheet.set_column(3, 3, 12)
        sheet.set_column(4, 4, 15)
        sheet.set_column(5, 5, 9)
        sheet.set_column(6, 6, 150)

        sheet.write(0, 0, "Номер")
        sheet.write(0, 1, "Исполнитель")
        sheet.write(0, 2, "Название")
        sheet.write(0, 3, "Длительность")
        sheet.write(0, 4, "Дата публикации")
        sheet.write(0, 5, "Качество")
        sheet.write(0, 6, "Ссылка")
        j = 1
        for comp in comps:
            sheet.write(j, 0, j)
            for i,x in enumerate(comp):
                sheet.write(j, i+1, x)
            j += 1

start, end = 0, 0
if __name__ == '__main__':
    start = int(input("Start page "))
    end = int(input("End page "))
    Parser().main(start, end+1, 'https://rus.megapesni.com/russian/')
