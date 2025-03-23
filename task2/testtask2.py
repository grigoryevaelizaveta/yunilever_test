
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import yagmail  
import os  
import time 
from bs4 import BeautifulSoup

def get_themes_from_excel(file_path, sheet_name="Sheet1"):

    try:
        wb = xw.Book(file_path)
        sheet = wb.sheets[sheet_name]
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
        themes = sheet.range(f"A2:A{last_row}").value
        wb.close()
        return themes
    except FileNotFoundError:
        print(f"Ошибка: Файл '{file_path}' не найден.")
        return None
    except Exception as e:
        print(f"Произошла ошибка при чтении Excel: {e}")
        return None


def search_and_get_links(themes, browser_name="chrome"):
    search_results = {}
    try:
        if browser_name.lower() == "chrome":
            chrome_options = Options()
            chrome_options.add_argument("--disable-notifications")  
            chrome_options.add_argument("--disable-popup-blocking") 
            chrome_options.add_argument("--disable-extensions")
            chrome_options.add_argument("--ignore-certificate-errors")
            chrome_options.add_argument("--allow-insecure-localhost")

            driver = webdriver.Chrome(options=chrome_options)
            
        elif browser_name.lower() == "firefox":
            driver = webdriver.Firefox() 
        else:
            print("Неподдерживаемый браузер.  Используется Chrome.")
            driver = webdriver.Chrome()

        driver.get("https://ya.ru")
        time.sleep(2)

        for theme in themes:
            search_results[theme] = []
            search_box = driver.find_element(By.ID, "text") 
            search_box.clear()
            search_box.send_keys(theme)
            search_box.send_keys(Keys.RETURN)
            time.sleep(6) 

            try:
                WebDriverWait(driver, 10).until(
                     EC.presence_of_element_located((By.ID, "search-result"))
                )
            except:
                print(f"Search results didn't load within timeout for '{theme}'")
                continue
            html = driver.page_source
            soup = BeautifulSoup(html, 'lxml')

            try:
               for i in range(1,4):
                  xpath = f'//*[@id="search-result"]/li[{i}]/div/div[2]/div/a' 
                  item = soup.find("a", {'href': True}, string=lambda text: text and len(text)>0 ) 

                  if (item):

                       href_text = item['href'] 
                       search_results[theme].append(href_text)

            except Exception as e:
                print(f"Ошибка при получении ссылок для темы '{theme}': {e}")

        driver.quit() 
        return search_results

    except Exception as e:
        print(f"Произошла ошибка при работе с браузером: {e}")
        return None


def update_excel_with_links(file_path, search_results, sheet_name="Sheet1"):

    try:
        wb = xw.Book(file_path)
        sheet = wb.sheets[sheet_name]
        last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row 
        start_row = last_row + 1

        sheet.range("A1:B1").api.AutoFilter(1) 

        for theme, links in search_results.items():
            for link in links:
                sheet.range(f"A{start_row}").value = theme  
                sheet.range(f"B{start_row}").value = link  
                start_row += 1

        wb.save()
        wb.close()
        print(f"Файл '{file_path}' успешно обновлен.")

    except FileNotFoundError:
        print(f"Ошибка: Файл '{file_path}' не найден.")
    except Exception as e:
        print(f"Произошла ошибка при записи в Excel: {e}")

def send_email(file_path, recipient_email, subject, sender_email, sender_password):

    try:
        yag = yagmail.SMTP(sender_email, sender_password, host='smtp.yandex.ru', port=465, smtp_ssl=True) 
        contents = [
            "Здравствуйте!\n\nВ приложении находится файл со списком тем и ссылками на результаты поиска.",
            file_path
        ]
        yag.send(recipient_email, subject, contents)
        print(f"Email успешно отправлен на {recipient_email}")
    except Exception as e:
        print(f"Произошла ошибка при отправке email: {e}")


if __name__ == "__main__":
    file_path = r"C:\Users\User\Desktop\TestTask2.xlsx"
    recipient_email = "1032199566@yandex.ru" 
    sender_email = "****@yandex.ru"
    sender_password = "****" 
    subject = "Список тем для доклада"
    browser_name = "chrome"  

    themes = get_themes_from_excel(file_path)

    if themes:

        search_results = search_and_get_links(themes, browser_name)

        if search_results:
            
            update_excel_with_links(file_path, search_results)
            send_email(file_path, recipient_email, subject, sender_email, sender_password)
        else:
            print("Не удалось получить результаты поиска.")
    else:
        print("Не удалось получить темы из Excel файла.")
