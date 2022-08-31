import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
import openpyxl


def brand_choice():
    print(f"Hello Dear Sir!\nCan you pls choice the brand auto?\nPress '1' - 'BMW'\nPress '2' - "
          f"'Mercedes-Benz'\nPress '3' - 'Volkswagen or Lexus'")
    brand = ''
    choice = int(input())
    while choice not in range(1, 4):
        if choice == 1:
            brand = 'BMW'
        elif choice == 2:
            brand = 'MERCEDES'
        elif choice == 3:
            brand = 'TOYOTA%20LEXUS'
        else:
            print("I don't understand, could you pls make a choice")
            choice = int(input())

    return brand


def get_data(url):
    # open the xls file and reading articles
    book = openpyxl.load_workbook(filename='C:\\Users\\chufy\\Desktop\\parser_zzap\\for_pars.xlsx')
    sheet = book.worksheets[0]
    len_sheet = len(sheet['A'])
    result_dict = {}

    brand = brand_choice()

    # approximate time
    approximate_time = time.gmtime((len_sheet - 1) * 13)
    print(f"Have about {time.strftime('%H:%M:%S', approximate_time)}")

    for number_of_str in range(1, len_sheet):
        article = sheet['A'][number_of_str].value
        print(f"{article} {number_of_str}/{len_sheet - 1}")

        # options for webdriver
        options = webdriver.FirefoxOptions()
        options.set_preference("general.useragent.override",
                               "Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:47.0) Gecko/20100101 Firefox/47.0")
        options.set_preference("dom.webdriver.enabled", False)
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--headless")  # headless mode
        s = Service("C:\\Users\\chufy\\Desktop\\parser_zzap\\geckodriver.exe")

        # download html
        try:
            driver = webdriver.Firefox(service=s, options=options)
            driver.get(
                url=url + f"/public/search.aspx#rawdata={article}&class_man={brand}&partnumber={article}&type=1")
            time.sleep(4)

            with open("index.html", "w") as file:
                file.write(driver.page_source)

            # parse html and get prices
            with open("index.html") as file:
                src = file.read()

        except Exception as ex:
            print("Don't download the HTML", ex)

        finally:
            driver.close()
            driver.quit()

            soup = BeautifulSoup(src, "lxml")
            try:

                min_price = soup.find(
                    id="ctl00_BodyPlace_SearchGridView_ctl28_SearchInfoAllPanel_PriceMinOrderLabel").text
                ave_price = soup.find(
                    id="ctl00_BodyPlace_SearchGridView_ctl28_SearchInfoAllPanel_PriceAvgOrderLabel").text
                max_price = soup.find(
                    id="ctl00_BodyPlace_SearchGridView_ctl28_SearchInfoAllPanel_PriceMaxOrderLabel").text
                offers = soup.find(
                    id="ctl00_BodyPlace_SearchGridView_ctl28_SearchInfoAllPanel_PriceCountOrderLabel").text
                # dictionary creation
                result_dict[article] = [min_price.replace(' ', ''), ave_price.replace(' ', ''),
                                        max_price.replace(' ', ''), offers]
            except AttributeError:
                print("Don't find the article")
                result_dict[article] = ["0р.", "0р.", "0р.", 0]

    for number_of_str in range(1, len_sheet):
        article = sheet['A'][number_of_str].value
        sheet['B' + str(number_of_str + 1)] = result_dict[article][0][:-2]
        sheet['C' + str(number_of_str + 1)] = result_dict[article][1][:-2]
        sheet['D' + str(number_of_str + 1)] = result_dict[article][2][:-2]
        sheet['E' + str(number_of_str + 1)] = result_dict[article][3]
        book.save("results.xlsx")
    print('Completed! Open the file with results')


def main():
    get_data("https://www.zzap.ru")


if __name__ == '__main__':
    main()
