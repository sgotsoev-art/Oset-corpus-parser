import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import Levenshtein

HEADLESS = False
DELAY = 2

PHRASES_FILE = "ФРАЗЕОЛОГИЗМЫ.xlsx"
LEMMAS_FILE = "ФРАЗЕОЛОГИЗМЫ_С_ЛЕММАМИ.xlsx"
OUTPUT_TEST_FILE = "test_results hadon.xlsx"

FIRST_PHRASE_IDX = 110
TEST_SUBCORPUS = ("Пресса", "press")

def set_lemma_mode(driver, field_num):
    driver.execute_script(f"changeTab('form{field_num}', 'lemma')")
    time.sleep(0.5)

def select_subcorpus_press_only(driver):
    menu_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frame")
    driver.switch_to.frame(menu_frame)
    sub_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Подкорпус')]"))
    )
    sub_btn.click()
    time.sleep(1)

    main_window = driver.current_window_handle
    all_windows = driver.window_handles
    dialog = None
    for w in all_windows:
        if w != main_window:
            dialog = w
            break
    if dialog:
        driver.switch_to.window(dialog)
    else:
        pass

    for g in ["press", "fict_group", "nonfiction_group", "oral_group"]:
        try:
            cb = driver.find_element(By.ID, g)
            if cb.is_selected():
                cb.click()
        except:
            pass

    press_cb = driver.find_element(By.ID, "press")
    if not press_cb.is_selected():
        press_cb.click()

    ok_btn = driver.find_element(By.XPATH, "//input[@class='button' and @value='   Ok   ']")
    ok_btn.click()
    time.sleep(1)

    driver.switch_to.window(main_window)
    driver.switch_to.default_content()
    menu_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frame")
    driver.switch_to.frame(menu_frame)

def get_context(driver, context_url):
    driver.execute_script("window.open('');")
    driver.switch_to.window(driver.window_handles[-1])
    driver.get(context_url)
    time.sleep(DELAY)

    for _ in range(5):
        try:
            expand = driver.find_element(By.XPATH, "//a[contains(text(),'Расширить контекст')]")
            expand.click()
            time.sleep(DELAY)
        except:
            break

    source = ""
    try:
        source_row = driver.find_element(By.XPATH, "/html/body/center/table/tbody/tr[1]/td/table/tbody/tr[1]/td/table/tbody/tr")
        full_text = source_row.text.strip()
        source = full_text
    except:
        pass


    try:
        text = driver.find_element(By.XPATH, "//td[contains(@style,'padding: 3px 10px')]").text.strip()
    except:
        text = ""

    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    return source, text

def levenshtein_dist(s1, s2):
    return Levenshtein.distance(s1.lower(), s2.lower())

def have_common_bigram(w1, w2, min_len=2):
    w1 = w1.lower()
    w2 = w2.lower()
    
    if len(w1) < min_len or len(w2) < min_len:
        return True
    
    for i in range(len(w1) - min_len + 1):
        if w1[i:i+min_len] in w2:
            return True
    return False

def words_similar(w1, w2, max_dist=5):
    dist = levenshtein_dist(w1, w2)
    if dist > max_dist:
        return False
    return have_common_bigram(w1, w2, min_len=2)

def phrase_in_text(phrase_words, text, max_dist=5):
    text_words = re.findall(r'\b\w+\b', text.lower())
    for pw in phrase_words:
        pw_low = pw.lower()
        found = False
        for tw in text_words:
            if words_similar(pw_low, tw, max_dist):
                found = True
                break
        if not found:
            return False
    return True

def tokenize(phrase):
    return re.findall(r'\b\w+\b', phrase)

def find_and_click_pagination(driver, next_page_num):
    """Ищет ссылку пагинации во всех фреймах и кликает по ней"""
    
    def search_in_frame(driver, frame_element=None):
        if frame_element:
            driver.switch_to.frame(frame_element)
        
        try:

            link = driver.find_element(By.LINK_TEXT, str(next_page_num))
            if link and link.is_displayed():
                return link
        except:
            pass
        
        try:
            frames = driver.find_elements(By.TAG_NAME, "frame")
            for frame in frames:
                driver.switch_to.default_content()
                if frame_element:
                    driver.switch_to.frame(frame_element)
                result = search_in_frame(driver, frame)
                if result:
                    return result
        except:
            pass
        
        return None
    
    driver.switch_to.default_content()
    
    link = search_in_frame(driver)
    
    if link:
        driver.execute_script("arguments[0].scrollIntoView(true);", link)
        time.sleep(0.5)
        link.click()
        return True
    
    return False


def main():
    df_phrases = pd.read_excel(PHRASES_FILE)
    df_lemmas = pd.read_excel(LEMMAS_FILE)
    df = pd.merge(df_phrases, df_lemmas, on="Фразеологизм", how="left")

    row = df.iloc[FIRST_PHRASE_IDX]
    phrase = row["Фразеологизм"]
    lemmas_str = row.get("Леммы фраз", "")
    if pd.isna(lemmas_str) or lemmas_str == "":
        print("Нет лемм для первого фразеологизма.")
        return
    lemma_parts = lemmas_str.split()
    if len(lemma_parts) < 2:
        print("Недостаточно лемм (нужно две).")
        return
    lemma1, lemma2 = lemma_parts[0], lemma_parts[1]
    phrase_words = tokenize(phrase)
    print(f"Тестируем фразеологизм: {phrase}")
    print(f"Леммы для поиска: {lemma1} , {lemma2}")
    print(f"Слова фразы: {phrase_words}")
    print(f"Подкорпус: {TEST_SUBCORPUS[0]}")
    print("="*60)

    chrome_options = Options()
    if HEADLESS:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--remote-debugging-port=9222")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(60)

    results = []

    try:
        print("Открываем главную страницу...")
        driver.get("http://corpus.ossetic-studies.org/search/?interface_language=ru")
        driver.switch_to.default_content()
        time.sleep(2)

        print("Выбираем подкорпус...")
        select_subcorpus_press_only(driver)
        print("Подкорпус выбран.")

        print("Устанавливаем режим лемма...")
        set_lemma_mode(driver, 1)
        set_lemma_mode(driver, 2)

        print("Вводим леммы и отправляем...")
        lex1 = driver.find_element(By.ID, "lex1")
        lex1.clear()
        lex1.send_keys(lemma1)
        lex2 = driver.find_element(By.ID, "lex2")
        lex2.clear()
        lex2.send_keys(lemma2)
        lex2.send_keys(Keys.RETURN)
        print("Поиск отправлен.")

        print("Переключаемся во фрейм результатов...")
        driver.switch_to.default_content()
        try:
            results_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frameset/frame[2]")
            driver.switch_to.frame(results_frame)
        except Exception as e:
            print("Ошибка переключения во фрейм результатов:", e)
            return

        print("Ожидаем таблицу результатов...")
        try:
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//table[@class='results_header']")))
            print("Таблица результатов найдена.")
        except TimeoutException:
            print("Таблица результатов НЕ найдена.")
            driver.save_screenshot("debug_screenshot.png")
            print("Сохранён скриншот debug_screenshot.png")
            return
        
        driver.switch_to.default_content()
        results_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frameset/frame[2]")
        driver.switch_to.frame(results_frame)
        print("--- ДИАГНОСТИКА ЗАВЕРШЕНА ---\n")

        page = 1
        while True:
            print(f"\n--- Страница {page} ---")
            
            all_expand_links = driver.find_elements(By.XPATH, "//a[contains(text(),'Расширить контекст')]")
            
            if not all_expand_links:
                print("  Нет ссылок расширения на странице")
            else:
                urls_to_process = []
                for link in all_expand_links:
                    href = link.get_attribute("href")
                    if href:
                        match = re.search(r"ContextOpen\('(.*?)%27", href)
                        if match:
                            rel_url = match.group(1)
                            full_url = "http://corpus.ossetic-studies.org/search/" + rel_url
                            urls_to_process.append(full_url)
                        else:
                            print(f"  Не удалось извлечь URL из href")
                    else:
                        print("  Нет href у ссылки")
                
                for idx, full_url in enumerate(urls_to_process, 1):
                    print(f"  Блок {idx}: получаем контекст...")
                    try:
                        source, context = get_context(driver, full_url)
                    except Exception as e:
                        print(f"    Ошибка: {e}")
                        continue
                    
                    print(f"    >>> КОНТЕКСТ: {context}")
                    print(f"    >>> ИСТОЧНИК: {source[:60]}")
                    print(f"    >>> СЛОВА ФРАЗЫ: {phrase_words}")


                    if phrase_in_text(phrase_words, context, max_dist=5):
                        print(f"    >>> СОВПАДЕНИЕ! Источник: {source[:60]}")
                        results.append({
                            "Фразеологизм": phrase,
                            "Подкорпус": TEST_SUBCORPUS[0],
                            "Источник": source,
                            "Контекст": context
                        })
                    else:
                        print(f"    Нет совпадения.")
                        text_words = re.findall(r'\b\w+\b', context.lower())
                        for pw in phrase_words:
                            pw_low = pw.lower()
                            found = False
                            for tw in text_words:
                                if words_similar(pw_low, tw, max_dist=5):
                                    found = True
                                    break
                            if not found:
                                print(f"      НЕ НАЙДЕНО СЛОВО: {pw_low}")
            
            print("  Пытаемся перейти на следующую страницу...")

            
            try:
                next_page_num = page + 1
                
                driver.switch_to.default_content()
                if find_and_click_pagination(driver, next_page_num):
                    print(f"  Переход на страницу {next_page_num}")
                    time.sleep(DELAY)
                    
                    driver.switch_to.default_content()
                    results_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frameset/frame[2]")
                    driver.switch_to.frame(results_frame)
                    
                    page += 1
                else:
                    print(f"  Больше страниц нет (не найдена ссылка на страницу {next_page_num})")
                    break
                
            except Exception as e:
                print(f"  Ошибка при переходе: {e}")
                break

    finally:
        driver.quit()

    if results:
        df_out = pd.DataFrame(results)
        df_out.to_excel(OUTPUT_TEST_FILE, index=False)
        print(f"\nГотово. Найдено {len(results)} совпадений. Сохранено в {OUTPUT_TEST_FILE}")
    else:
        print("\nСовпадений не найдено.")

if __name__ == "__main__":
    main()