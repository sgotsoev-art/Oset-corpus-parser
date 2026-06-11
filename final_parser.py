import pandas as pd # работа с таблицами
from selenium import webdriver #драйвер для работы с хромом
from selenium.webdriver.common.by import By #для поиска элементов
from selenium.webdriver.common.keys import Keys # для взаимод-ия с элементами
from selenium.webdriver.support.ui import WebDriverWait #модуль с ожиданиями 
from selenium.webdriver.support import expected_conditions as EC #условия для ожиданий
from selenium.webdriver.chrome.options import Options #настройки браузера
from selenium.webdriver.chrome.service import Service #модуль для драйвера
from selenium.common.exceptions import TimeoutException #ошибка при ожидании 
from webdriver_manager.chrome import ChromeDriverManager #для загрузки драйвера
import time #паузы в программе
import re #для извлечения выражений
import Levenshtein #расстояние левенштейа
import os #работа с файловой системой
import traceback #для удобной отладки

# ------------------------------
# конфигурация
# ------------------------------
HEADLESS = True          # False = видно браузер, True = он на фоне
DELAY = 2 #время паузы
SAVE_EVERY = 5 #частота сохранения результатов в бэкап

#тут можно переключать таблицу с которой работаем

PHRASES_FILE = "ФРАЗЕОЛОГИЗМЫ.xlsx"
LEMMAS_FILE = "ФРАЗЕОЛОГИЗМЫ_С_ЛЕММАМИ.xlsx"
OUTPUT_FILE = "final_results.xlsx"
BACKUP_FILE = "final_results_backup.xlsx"

#это просто подкорпуса и их id на сайте
SUBCORPORA = [
    ("Пресса", "press"),
    ("Художественные", "fict_group"),
    ("Нехудожественные", "nonfiction_group"),
    ("Устная речь", "oral_group"),
]

# ------------------------------
# вспомог-ые функции
# ------------------------------
def set_lemma_mode(driver, field_num): #эта функция для скрипта который переключает кнопки режима поиска
    driver.execute_script(f"changeTab('form{field_num}', 'lemma')")
    time.sleep(0.5)

def select_subcorpus(driver, subcorpus_id, retries=3): #эта функция для выбора подкорпуса, три переменные, драйвер, айди подкорпуса и кол-во попыток при ошибке
    for attempt in range(retries): #цикл повторов попыток
        try:
            menu_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frame")#ищем фрейм в котором кнопка выбора подкорпусов
            driver.switch_to.frame(menu_frame)
            sub_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[contains(text(),'Подкорпус')]"))
            )
            sub_btn.click()
            time.sleep(1)

            main_window = driver.current_window_handle #фиксируем айди актуального окна
            all_windows = driver.window_handles #тут получаем список всех открытых окон, т.к. кнопка подкорпус открывается в отдельном окне
            dialog = None 
            for w in all_windows:
                if w != main_window:
                    dialog = w
                    break
            if dialog:
                driver.switch_to.window(dialog) #по итогу находим окно айди которого не айди главного окна и переключаем драйвер на новое окно

            # Ждём, пока загрузится окно с подкорпусами
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "press"))
            )

            for g in ["press", "fict_group", "nonfiction_group", "oral_group"]: #выключаем все подкорпуса чтобы дальше при переключении на след подкорпус мы могли выключить предыдущий по которому уже получили результаты
                try:
                    cb = driver.find_element(By.ID, g)
                    if cb.is_selected():
                        cb.click()
                except:
                    pass

            if subcorpus_id == "press": #далее в мейне мы будем перебирать айди подкорпусов и дальнейший код каждый подкорпус включает по отдельности
                cb = driver.find_element(By.ID, "press")
                if not cb.is_selected():
                    cb.click()
            else:
                group_cb = driver.find_element(By.ID, subcorpus_id)
                if not group_cb.is_selected():
                    group_cb.click()
                if subcorpus_id == "fict_group":
                    driver.execute_script("turnAllChecks('fict');")
                elif subcorpus_id == "nonfiction_group":
                    driver.execute_script("turnAllChecks('nonfiction');")
                elif subcorpus_id == "oral_group":
                    driver.execute_script("turnAllChecks('oral');")

            ok_btn = driver.find_element(By.XPATH, "//input[@class='button' and @value='   Ok   ']")
            ok_btn.click()
            time.sleep(1)
#дальше опять переключаемся в главное окно и работаем в нужном фрейме
            driver.switch_to.window(main_window)
            driver.switch_to.default_content()
            menu_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frame")
            driver.switch_to.frame(menu_frame)
            
            return  # Успех, выходим из функции

        except Exception as e:
            print(f"    Попытка {attempt+1}/{retries} не удалась: {e}")
            if attempt < retries - 1:
                time.sleep(3)
                driver.switch_to.default_content()
            else:
                raise  # Если все попытки не удались, выбрасываем исключение

def get_context(driver, context_url): #данная функция дает нам расширенный контекст каждого блока результатов
    original_window = driver.current_window_handle
    new_window = None
    
    try:
        driver.execute_script("window.open('');") #открываем новую вкладку с результатом
        new_window = driver.window_handles[-1] 
        driver.switch_to.window(new_window) #переключаемся в эту вкладку
        
        try:
            driver.set_page_load_timeout(30) #даем 30сек на загрузку страницы с расширенным контекстом
            driver.get(context_url) 
        except TimeoutException:
            driver.get("about:blank")
            time.sleep(1)
            driver.get(context_url)
        
        time.sleep(DELAY)
        
        for _ in range(5): #максимально расширяем контекст, тут можно регулировать количество попыток, но больше 5 раз вроде не встречал
            try:
                expand = driver.find_element(By.XPATH, "//a[contains(text(),'Расширить контекст')]")
                expand.click()
                time.sleep(DELAY)
            except:
                break
        
        source = "" #строчка для источника
        try:
            source_row = driver.find_element(By.XPATH, "/html/body/center/table/tbody/tr[1]/td/table/tbody/tr[1]/td/table/tbody/tr")
            source = source_row.text.strip() #вносим источник
        except:
            pass
        
        try: #а тут текст из расширенного контекста получаем
            text = driver.find_element(By.XPATH, "//td[contains(@style,'padding: 3px 10px')]").text.strip()
        except:
            text = ""
        
        return source, text
        #и закрываем вкладку с расш контекстом
    finally:
        try:
            if new_window and new_window in driver.window_handles:
                driver.close()
        except:
            pass
        try:
            driver.switch_to.window(original_window)
        except:
            pass

def levenshtein_dist(s1, s2): #ну тут вроде все понятно
    return Levenshtein.distance(s1.lower(), s2.lower())

def have_common_bigram(w1, w2, min_len=2): #тут я добавил еще для длинных слов совпадение по хотя бы двум буквам чтобы они считались как бы однокоренными
    w1 = w1.lower()
    w2 = w2.lower()
    
    # Если хотя бы одно слово короче 2 букв то пропускаем проверку 
    if len(w1) < min_len or len(w2) < min_len:
        return True
    #если слова длиннее то ищем в них две одинаковые буквы 
    for i in range(len(w1) - min_len + 1):
        if w1[i:i+min_len] in w2:
            return True
    return False
#сравниваем теперь слова типа определяем условие если расстояние левенштейна меньше либо равно 5 то тогда ищем среди них одинаковые буквы идущие подряд и если они нахоядтся то слова похожи
def words_similar(w1, w2, max_dist=5):
    dist = levenshtein_dist(w1, w2)
    if dist > max_dist:
        return False
    return have_common_bigram(w1, w2, min_len=2)

def phrase_in_text(phrase_words, text, max_dist=5): #эта функция проверяет все ли слова фраз-ма присутствуют в тексте
    text_words = re.findall(r'\b\w+\b', text.lower()) #получаем список слов из текста
    for pw in phrase_words: #ищем для каждого слова фраз-ма похожие слова в тексте, если для каждого слова фраз-ма нашлось слово то тогда этот блок результата засчитывается
        pw_low = pw.lower()
        found = False
        for tw in text_words:
            if words_similar(pw_low, tw, max_dist):
                found = True
                break
        if not found:
            return False
    return True

def tokenize(phrase): #эта функция будет исп-ся в мейне чтобы получить phrase_words для пред-ей функции
    return re.findall(r'\b\w+\b', phrase)

def find_and_click_pagination(driver, next_page_num): #ищет есть ли след страница и нажимает на нее если она есть
    try:
        driver.switch_to.default_content()
        link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, str(next_page_num)))
        )
        link.click()
        return True
    except:
        return False

# 
# и основная функция мейн
# 
def main():
    df_phrases = pd.read_excel(PHRASES_FILE)
    df_lemmas = pd.read_excel(LEMMAS_FILE)
    df = pd.merge(df_phrases, df_lemmas, on="Фразеологизм", how="left")

    all_results = []
    start_idx = 0 #этот кусок позволяет работать с бэкапом, если какие-то результаты сохранились перед ошибкой
    if os.path.exists(BACKUP_FILE): 
        try:
            backup_df = pd.read_excel(BACKUP_FILE)
            all_results = backup_df.to_dict('records')
            
            # тут мы выгружаем все имеющиеся пары фраз-м и подкорпус
            processed_pairs = set()
            for r in all_results:
                processed_pairs.add((r['Фразеологизм'], r['Подкорпус']))
            
            # Ищем первый фразеологизм, у которого не все подкорпуса обработаны(для этого перебираем фраз-мы начиная с первого и для него проверяем все пары с подкорпусами, и если для фраз-ма нет какой-то пары то парсим его заново(чтобы в случае если вдруг возникла ошибка в середине подкорпуса чтоб не потерять результат)
            for i, row in df.iterrows():
                phrase = row['Фразеологизм']
                all_subs_done = True
                for sub_name, _ in SUBCORPORA:
                    if (phrase, sub_name) not in processed_pairs:
                        all_subs_done = False
                        break
                if not all_subs_done:
                    start_idx = i
                    break
                    
            print(f"Найден бэкап. Продолжаем с фразеологизма {start_idx+1}")
        except Exception as e:
            print(f"Ошибка при загрузке бэкапа: {e}, начинаем сначала")
            all_results = []
            start_idx = 0

    chrome_options = Options()
    if HEADLESS: #тут просто настраиваем браузер, режим отображения окна размер окна скачиваем драйвер и тд
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.set_page_load_timeout(60)

    total_phrases = len(df)  #регистрируем количество фраз-ов
    try:#начинаем основной цикл
        for idx in range(start_idx, total_phrases): #тут включается уже стартовый индекс либо нулевой либо индекс с бэкап файла
            row = df.iloc[idx] #берем строку с соотв-им индексом
            phrase = row["Фразеологизм"] #извлекаем фразеологизм
            lemmas_str = row.get("Леммы фраз", "") #извлекаем леммы

            # если нет лемм то сразу добавляем пустые строки для всех подкорпусов
            if pd.isna(lemmas_str) or lemmas_str == "":
                for sub_name, _ in SUBCORPORA:
                    all_results.append({
                        "Фразеологизм": phrase,
                        "Подкорпус": sub_name,
                        "Источник": "",
                        "Контекст": ""
                    })
                print(f"[{idx+1}/{total_phrases}] Пропуск: {phrase} — нет лемм (добавлены пустые строки)")
                continue
            #также исключаем фраз-мы для которых нет лемм
            lemma_parts = lemmas_str.split()
            if len(lemma_parts) < 2:
                for sub_name, _ in SUBCORPORA:
                    all_results.append({
                        "Фразеологизм": phrase,
                        "Подкорпус": sub_name,
                        "Источник": "",
                        "Контекст": ""
                    })
                print(f"[{idx+1}/{total_phrases}] Пропуск: {phrase} — недостаточно лемм (добавлены пустые строки)")
                continue

            lemma1, lemma2 = lemma_parts[0], lemma_parts[1] #фиксируем леммы в переменные и в терминале выводим информацию об обрабатываемом фраз-ме
            phrase_words = tokenize(phrase) 
            print(f"\n=== [{idx+1}/{total_phrases}] Обработка: {phrase} ===")
            print(f"Леммы: {lemma1}, {lemma2}")

            # через эту переменную будем отслеживать совпадения по подкорпусам
            has_matches = {sub_name: False for sub_name, _ in SUBCORPORA}

            for sub_name, sub_id in SUBCORPORA: #запускаем цикл поиска по подкорпусам
                print(f"  Подкорпус: {sub_name}")
                driver.get("http://corpus.ossetic-studies.org/search/?interface_language=ru") #переходим на страницу
                driver.switch_to.default_content() 
                time.sleep(2)

                select_subcorpus(driver, sub_id) #вызываем вспомогательную функцию и выбираем актуальный подкорпус

                set_lemma_mode(driver, 1) #переключаем режим поиска
                set_lemma_mode(driver, 2)

# тут мы вбиваем леммы в строки поиска и запускаем поиск
                lex1 = driver.find_element(By.ID, "lex1")
                lex1.clear()
                lex1.send_keys(lemma1) 
                lex2 = driver.find_element(By.ID, "lex2")
                lex2.clear()
                lex2.send_keys(lemma2)
                lex2.send_keys(Keys.RETURN)
                time.sleep(DELAY)

                driver.switch_to.default_content() #дальше переключаемся во фрейм результатов
                try:
                    results_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frameset/frame[2]")
                    driver.switch_to.frame(results_frame)
                except Exception as e:
                    print(f"    Ошибка переключения во фрейм результатов: {e}")
                    continue

                try: #проверяем есть ли результаты
                    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//table[@class='results_header']")))
                except TimeoutException:
                    print("    Нет результатов (таблица не загружена)")
                    continue

                page = 1 #если результаты есть то запускаем сбор информации по блокам результатов 
                while True:
                    print(f"    Страница {page}")
                    all_expand_links = driver.find_elements(By.XPATH, "//a[contains(text(),'Расширить контекст')]") #находим блоки по кнопке расширить контекст
                    if all_expand_links: #если есть хотя бы один блок
                        urls_to_process = []
                        for link in all_expand_links: #перебираем все имеющиеся ссылки
                            href = link.get_attribute("href") # берем атрибут href из ссылки
                            if href: 
                                match = re.search(r"ContextOpen\('(.*?)%27", href) #если атрибут есть то берем его 
                                if match:
                                    rel_url = match.group(1) #фиксируем относительный url
                                    full_url = "http://corpus.ossetic-studies.org/search/" + rel_url #получаем окончательную ссылку
                                    urls_to_process.append(full_url) #получаем список всех ссылок на окна с расширенным контекстом
                        for i_url, full_url in enumerate(urls_to_process, 1): #открываем по очереди все полученные ссылки
                            print(f"      Блок {i_url}: получаем контекст...") 
                            try:
                                source, context = get_context(driver, full_url) #пытаемся расширитьь контекст и получаем окончательные источник и текст
                            except Exception as e:
                                print(f"        Ошибка: {e}")
                                continue
                            if phrase_in_text(phrase_words, context, max_dist=5): # проверяем есть ли все слова фразеологизма
                                print(f"        >>> СОВПАДЕНИЕ! Источник: {source[:60]}")
                                has_matches[sub_name] = True
                                all_results.append({ #сохраняем результат в списокк если есть совпадение
                                    "Фразеологизм": phrase,
                                    "Подкорпус": sub_name,
                                    "Источник": source,
                                    "Контекст": context
                                })
                            else:
                                print(f"        Нет совпадения.")
                    # далее переходим на следующую страницу
                    next_page_num = page + 1
                    driver.switch_to.default_content()
                    if find_and_click_pagination(driver, next_page_num):
                        print(f"      Переход на страницу {next_page_num}")
                        time.sleep(DELAY)
                        driver.switch_to.default_content()
                        results_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frameset/frame[2]") #просто ищем ссылку на след страницу если она есть
                        driver.switch_to.frame(results_frame)
                        page += 1
                    else:
                        print("      Достигнут конец результатов")
                        break

            # если совпадений не было для какого-то подкорпуса то в таблицу добавляется пустая строка
            for sub_name, _ in SUBCORPORA:
                if not has_matches[sub_name]:
                    all_results.append({
                        "Фразеологизм": phrase,
                        "Подкорпус": sub_name,
                        "Источник": "",
                        "Контекст": ""
                    })

            # Периодическое сохранение
            if (idx+1) % SAVE_EVERY == 0: #тут просто сохраняем каждый пятый результат
                pd.DataFrame(all_results).to_excel(BACKUP_FILE, index=False)
                print(f"  *** Промежуточное сохранение ({len(all_results)} записей) ***")

    except KeyboardInterrupt: #на случай отмены компиляции или в случае ошибки выводим ее код для отладки
        print("\nПрерывание пользователем. Сохраняю текущие результаты...")
    except Exception as e:
        print(f"\nОшибка: {e}")
        traceback.print_exc()

    finally: #ну и закрываем браузер и сохраняем все имеющиеся результаты если они есть
        driver.quit()
        if all_results:
            df_final = pd.DataFrame(all_results)
            df_final.to_excel(OUTPUT_FILE, index=False)
            print(f"\nСохранено {len(all_results)} записей (включая пустые) в {OUTPUT_FILE}")
            if os.path.exists(BACKUP_FILE):
                os.remove(BACKUP_FILE)
        else:
            print("Нет данных для сохранения.")

if __name__ == "__main__":
    main()