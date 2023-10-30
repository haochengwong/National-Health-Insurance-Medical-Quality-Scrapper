from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import csv

# 設置Chromedriver位置，自行更改
driver_path = 'chromedriver-linux64/chromedriver'
 
# 創建Chrome
driver = webdriver.Chrome(executable_path=driver_path)

#選擇想抓取的時間，可自行微調
#ex: 110年上半年度 → 10年上半 110年全年度 → 10年全
y=['01年上半','01年下半','02年上半','02年下半','03年上半','03年下半','04年上半','04年下半','05年上半','05年下半',
   '06年上半','06年下半','07年上半','07年下半','08年上半','08年下半','09年上半','09年下半','10年上半','10年下半']

for y1 in y:
    url = "https://med.nhi.gov.tw/ihqe0000/IHQE0010S01.aspx?Type=DM"  #網址
    driver.get(url)
    
    dropdown = driver.find_element(By.ID, "drop1")
    dropdown.find_element(By.XPATH, "//option[. = '1"+y1+"年度']").click()  
    element = driver.find_element(By.ID, "drop1")
    actions = ActionChains(driver)
    actions.move_to_element(element).click_and_hold().perform()
    element = driver.find_element(By.ID, "drop1")
    actions = ActionChains(driver)
    actions.move_to_element(element).perform()
    element = driver.find_element(By.ID, "drop1")
    actions = ActionChains(driver)
    actions.move_to_element(element).release().perform()
    driver.find_element(By.ID, "query").click()

    dropdown1 = driver.find_element(By.ID, "DropDA")
    #以下5選一貼入 "//option[. = '糖尿病病人執行檢查率－醣化血紅素（HbA1c）或糖化白蛋白(glycated albumin)']" 中
    #糖尿病病人執行檢查率－醣化血紅素（HbA1c）或糖化白蛋白(glycated albumin)
    #糖尿病病人執行檢查率-空腹血脂
    #糖尿病病人執行檢查率-眼底檢查或眼底彩色攝影
    #糖尿病病人執行檢查率-尿液蛋白質(微量白蛋白)檢查
    #糖尿病病人加入照護方案比率
    dropdown1.find_element(By.XPATH, "//option[. = '糖尿病病人執行檢查率－醣化血紅素（HbA1c）或糖化白蛋白(glycated albumin)']").click() 
    driver.find_element(By.ID, "query").click
    element = driver.find_element(By.ID, "drop1")
    actions = ActionChains(driver)
    actions.move_to_element(element).click_and_hold().perform()

    
    dropdown = driver.find_element(By.NAME, "pep004table_length")
    dropdown.find_element(By.XPATH, "//option[. = '100']").click()
    element = driver.find_element(By.NAME, "pep004table_length")
    actions = ActionChains(driver)
    actions.move_to_element(element).click_and_hold().perform()
    element = driver.find_element(By.NAME, "pep004table_length")
    actions = ActionChains(driver)
    actions.move_to_element(element).perform()
    element = driver.find_element(By.NAME, "pep004table_length")
    actions = ActionChains(driver)
    actions.move_to_element(element).release().perform()

    te=[]
    a=driver.find_elements(By.CLASS_NAME, "paginate_button ")
    for element in a:
        text = element.text.replace("←", "").replace("→", "")
        te.append(text)
    te = list(filter(lambda x: x.strip() != "", te))
    numeric_data = [int(x) if x.isdigit() else x for x in te]
    print(numeric_data)
    for i in range(max(numeric_data)):
        b=driver.find_elements(By.CLASS_NAME, "dt-center")

        text_list = []
        text_list_1 = []
        for element in b:
            text = element.text
            text = text.replace("勾選", "").replace("地區", "").replace("機構名稱", "").replace("特約類別", "").replace("分子", "").replace("分母", "").replace("院所指標值", "").replace("所屬分區業務組指標值", "").replace("\n","")
            text_list.append(text)

        c=driver.find_elements(By.CLASS_NAME, "RowLeft")
        for element in c:
            text = element.text.replace("機構名稱", "")
            text_list_1.append(text)

        text_list_1 = list(filter(lambda x: x.strip() != "", text_list_1))
        text_list = list(filter(lambda x: x.strip() != "", text_list))

        rows = [text_list[i:i+6] for i in range(0, len(text_list), 6)]#將數據分為6個一組
        rows_1 = [text_list_1[i:i+1] for i in range(0, len(text_list), 1)]
        print(rows)
        
        with open('data_'+y1+'.csv', 'a', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(rows)
        with open('診所名稱_'+y1+'.csv', 'a', newline='', encoding='utf-8-sig') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerows(rows_1)

        driver.find_element(By.ID, "pep004table_next").click()

    with open('診所名稱_'+y1+'.csv', 'r', newline='',encoding='utf-8-sig') as csv_file:
        csv_reader = csv.reader(csv_file)
        data = list(csv_reader)

    data = [row for row in data if any(row)]
    with open('診所名稱_'+y1+'.csv', 'w', newline='',encoding='utf-8-sig') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerows(data)

    import pandas as pd

    df1 = pd.read_csv('診所名稱_'+y1+'.csv', header=None)
    df2 = pd.read_csv('data_'+y1+'.csv', header=None)

    merged_df = pd.concat([df1, df2], axis=1)
    merged_df.to_csv('data_'+y1+'.csv', index=False, header=False,encoding='utf-8-sig')

driver.quit()