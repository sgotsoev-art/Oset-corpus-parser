import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time
import re


chrome_options = Options()
chrome_options.add_argument("--headless")  
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")

df = pd.read_excel('ФРАЗЕОЛОГИЗМЫ.xlsx')
fr = df['Фразеологизм'].tolist()

driver = webdriver.Chrome(options=chrome_options) 
wait = WebDriverWait(driver, 10)
df['Леммы фраз'] = ""

for i in range(len(fr)):
    phrase = fr[i]
    words = phrase.split()
    first_two = words[:2]

    lemmas = []
    for k, word in enumerate(first_two):
        driver.get('http://corpus.ossetic-studies.org/search/?interface_language=ru')
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "/html/frameset/frameset/frame")))
        search_box = wait.until(EC.presence_of_element_located((By.ID, "lex1")))
        search_box.clear()
        search_box.send_keys(word)
        search_box.submit()

        driver.switch_to.default_content()
        result_frame = wait.until(EC.presence_of_element_located((By.XPATH, "/html/frameset/frameset/frameset/frame[2]")))
        driver.switch_to.frame(result_frame)

        try:
            elem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[class^='result1']")))
            onmouse = elem.get_attribute("onmouseover")
            match = re.search(r"popup\(this,\['([^']*)'", onmouse)
            if match:
                lemma = match.group(1).strip('[]')
            else:
                lemma = ""
        except:
            lemma = ""

        lemmas.append(lemma)
        time.sleep(1)

    lemmas_str = " ".join(lemmas)
    df.at[i, 'Леммы фраз'] = lemmas_str
    print(f"{i}: {phrase} -> {lemmas_str}")

df.to_excel('ФРАЗЕОЛОГИЗМЫ_С_ЛЕММАМИ.xlsx', index=False)
driver.quit()


    
#driver.get('http://corpus.ossetic-studies.org/search/?interface_language=ru')
#wait = WebDriverWait(driver, 10)
#menu_frame = driver.find_element(By.XPATH, "/html/frameset/frameset/frame")
#driver.switch_to.frame(menu_frame)
#search_box = driver.find_element(By.ID, "lex1")
#search_box.send_keys(phrmas[k])
#search_box.submit()


    
#time.sleep(10)

#driver.quit()
