from openpyxl import workbook
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import logging
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

logging.basicConfig(
    filename='log.log',
    filemode='w',
    level=logging.INFO,
    encoding='utf-8',
    datefmt="[%d-%m-%Y %H:%M:%S]",
    format='%(asctime)s %(levelname)-8s [%(filename)s:%(lineno)d %(message)s]'
)

class leroy_parsing():
    def __init__(self, url:str):
        self.url = url

    def write_to_file(self, data:list):
        filename = "result.xlsx"
        try:
            wb = load_workbook(filename)
            page = wb.active
            for row in data:
                page.append(row)
            wb.save(filename)
            logging.info(f'Данные записаны в файл: количество строк - {len(data)}')
            print(f'Данные записаны в файл: количество строк - {len(data)}')
        except Exception as ex:
            logging.info(f"Ошибка записи в фай: {ex}")
            print(f"Ошибка записи в фай: {ex}")
            
    def get_url(self):
        try:
            self.driver.get(self.url)
            logging.info(f'В драйвер загружена страница {self.url}')
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[data-qa='product']")))
            print(f'В драйвер загружена страница {self.url}')
        except Exception as ex:
            logging.info(f'Ошибка загрузки страницы {ex}')
            print('Станица загружена в драйвер')
                

    def setup(self):
        options = webdriver.ChromeOptions()
        options.add_argument("user-agent=Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:84.0) Gecko/20100101 Firefox/84.0")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument('--ignore-certificate-errors')
        #options.add_argument("--headless=new")
        options.add_argument("--disable-logging")
        self.driver = webdriver.Chrome(options=options)
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        'source': '''
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
        '''
        })
        self.driver.maximize_window()
    
    def pagination(self):
        count = 0
        while self.driver.find_elements(By.CSS_SELECTOR, "[data-qa-pagination-item='right']"):
            count +=1
            self.parse_page()
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[data-qa='product']")))
            self.driver.find_element(By.CSS_SELECTOR, "[data-qa-pagination-item='right']").click()
            logging.info(f"Отработала пагинация номер {count}")
            print(f"Отработала пагинация номер {count}")
            time.sleep(3)
    
    def parse_page(self):
        items = self.driver.find_elements(By.CSS_SELECTOR, "[data-qa='product']")
        logging.info(f"Найдено количество товаров: {len(items)}")
        print(f"Найдено количество товаров: {len(items)}")
        page_data = []
        for item in items:
            temp_list = []
            try:
                article = item.find_element(By.CSS_SELECTOR, "[data-qa='product-article']").text.split()[1]
            except Exception as ex:
                print("Ошибка считывания артикула")
            try:
                name = item.find_element(By.CSS_SELECTOR, "[data-qa='product-name']").text
            except Exception as ex:
                print("Ошибка считывания названия")
            try:
                base_price = float(item.find_element(By.CSS_SELECTOR, "[data-qa='primary-price-main']").text.replace(" ", ""))
                base_price = round(base_price, 2)
            except Exception as ex:
                base_price = ""
            try:
                best_price = float(item.find_element(By.CSS_SELECTOR, "[data-qa='best-price-main']").text.replace(" ", ""))
                best_price = round(best_price, 2)
            except Exception as ex:
                best_price = ""
            try:
                discount_price = float(item.find_element(By.CSS_SELECTOR, "[data-qa='new-price-main']").text.replace(" ", ""))
                discount_price = round(discount_price, 2)
            except Exception as ex:
                discount_price = ""
            temp_list = [article, name, base_price, best_price, discount_price]
            print(temp_list)
            page_data.append(temp_list)
        self.write_to_file(page_data)
    
    def parse(self):
        self.setup()
        self.get_url()
        self.pagination()
        


if __name__ == "__main__":
    start_time = time.perf_counter()
    urls = [
        "https://kaliningrad.leroymerlin.ru/catalogue/radiatory-otopleniya/",
        "https://kaliningrad.leroymerlin.ru/catalogue/elektroinstrumenty/",
        "https://kaliningrad.leroymerlin.ru/catalogue/suhie-smesi-i-gruntovki/",
    ]
    for url in urls:
        leroy_parsing(url).parse()
    end_time = time.perf_counter()
    print('Время выполнения парсинга Леруа мерлен:', round((float(end_time - start_time) / 60), 1), ' минут!')