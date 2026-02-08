# 設定台北時區
import os
from datetime import datetime
import pytz

tz = pytz.timezone("Asia/Taipei")
now = datetime.now(tz)
roc_year = now.year - 1911  # 民國年
date = f"{roc_year}{now.month:02d}{now.day:02d}"

#建立新的資料夾
path=r"C:\Users\trista.huang\Desktop\證券日報表"
folder_path=os.path.join(path,date)
if not os.path.exists(folder_path):
    os.makedirs(folder_path)
    print(f"資料夾'{folder_path}',已建立!")
else:
    print(f"資料夾'{folder_path}',已存在!")

#複製新的檔案到資料夾_智冠
import shutil
base_path = r"C:\Users\trista.huang\Desktop\證券日報表"
folder_path = os.path.join(base_path, date)

# 原始檔案路徑
source_path_sw = os.path.join(base_path, "1150121", f"智冠_1150121.xlsx")

# 目標檔案路徑（複製後的新檔名）
target_path_sw = os.path.join(folder_path, f"智冠_{date}.xlsx")

# 執行複製
shutil.copyfile(source_path_sw, target_path_sw)

print("智冠檔案複製完成！")

#分頁內容更新
import xlwings as xw

if not os.path.exists(target_path_sw):
    print("智冠檔案不存在！")
else:
    wb_sw = xw.Book(target_path_sw)
    
    # 修改分頁名稱
    if len(wb_sw.sheets) >= 2:
        wb_sw.sheets[0].name = f"智冠_{date}"
        wb_sw.sheets[1].name = f"5478_{date}"
        wb_sw.save()
        print("檔案分頁名稱修改完成")
        
        # 清除第一分頁的列表
        sheet1 = wb_sw.sheets[0]
        sheet1.range('A3').expand('table').clear_contents()
        
        # 清除第一分頁的外資買賣超表格
        sheet2 = wb_sw.sheets[0]
        sheet2.range('F4').expand('table').clear_contents()
        
        #清除第一分頁的截圖
        sheet3 = wb_sw.sheets[0]
        for shape in sheet3.shapes:
            shape.delete()

        #清除第一分頁的單日增減變動
        sheet4 = wb_sw.sheets[0]
        sheet4.range('F13').expand('table').clear_contents()

        #清除分點名單
        sheet5 = wb_sw.sheets[0]
        sheet5.range('I13').expand('table').clear_contents()

        #清除第二個分頁的總表
        sheet6=wb_sw.sheets[1]
        sheet6.range('A3').expand('table').clear_contents()
  
        print("原始資料清除完成")
    else:
        print("無法執行清除與命名")

#爬取券商買賣證券日報表查詢系統（一般交易）
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
from bs4 import BeautifulSoup
#爬取網址
options = uc.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
driver = uc.Chrome(version_main=144, options=options)
url = "https://www.tpex.org.tw/zh-tw/mainboard/trading/info/brokerBS.html"
driver.get(url)

#輸入要下載的股票代號
import time 
time.sleep(6)
input_section = driver.find_element(By.ID, "___auto1")
keyword="5478"
input_section.send_keys(keyword)
download_button = driver.find_element(By.XPATH, '//button[text()="下載 CSV (BIG5)"]')
download_button.click()
print("下載完成")

#將下載的檔案讀寫回檔案
import io
import pandas as pd
time.sleep(10)
file_path=f"C:\\Users\\trista.huang\\Downloads\\5478_{date}.csv"
df = pd.read_csv(file_path, encoding='big5',header=2)
df['買進股數'] = df['買進股數'].apply(lambda x: f"{x:,}") 
df['賣出股數'] = df['賣出股數'].apply(lambda x: f"{x:,}")

wb_sw = xw.Book(target_path_sw)
sheet_sw=wb_sw.sheets[1]

sheet_sw.range('A3').value = [df.columns.tolist()] + df.values.tolist() 
print("智冠CSV 資料已成功讀取並寫入 Excel 第二分頁")

#剔除序號和價格
df= df.drop(columns=['序號', '價格'])
#step 1 先改欄位名稱
df = df.rename(columns={ 
    df.columns[1]: "加總-買進股數", 
    df.columns[2]: "加總-賣出股數",
})
#step 2 移除逗號
df['加總-買進股數'] = df['加總-買進股數'].apply(lambda x: int(str(x).replace(',', '')))
df['加總-賣出股數'] = df['加總-賣出股數'].apply(lambda x: int(str(x).replace(',', '')))

# step 3先確保買進股數和賣出股數是int
df['加總-買進股數'] =df['加總-買進股數'].astype(int)
df['加總-賣出股數'] =df['加總-賣出股數'].astype(int)

# step 4合併券商資料
df= df.groupby('券商')[['加總-買進股數', '加總-賣出股數']].sum().reset_index()

# step 5加入一行加總net
df['加總net'] =df['加總-買進股數'] -df['加總-賣出股數']

# step6 加入總計列
summary_row = pd.DataFrame([{
    '券商': '總計',
    '加總-買進股數': df['加總-買進股數'].sum(),
    '加總-賣出股數': df['加總-賣出股數'].sum(),
    '加總net': df['加總net'].sum()
}])
df = pd.concat([df, summary_row], ignore_index=True)

import numpy as np
# 先把「加總net」欄位轉成數字，非數字的轉成 NaN
df_temp = df.copy()
df_temp["加總-買進股數"] = pd.to_numeric(df_temp["加總-買進股數"], errors="coerce")
df_temp["加總-賣出股數"] = pd.to_numeric(df_temp["加總-賣出股數"], errors="coerce")
df_temp["加總net"] = pd.to_numeric(df_temp["加總net"], errors="coerce")

#將代號和券商名稱分開
df[['券商代號', '券商名稱']] =df['券商'].str.split(' ',n=1, expand=True)
# 重新排列欄位順序 
df= df[['券商代號', '券商名稱', '加總-買進股數', '加總-賣出股數', '加總net']]

#先確保券商代號是字串再篩選
df['券商代號'] =df['券商代號'].astype(str)
target_ids = ['1360', '1380', '1440', '1470', '1480', '1520', '1530', '1560','1570', '1590', '1650','8440', '8890', '8900', '8910', '8960','9268']
filtered = df[df['券商代號'].isin(target_ids)]

filtered= filtered.drop(columns=['加總-買進股數', '加總-賣出股數'])
#把加總Net更名
filtered= filtered.rename(columns={"加總net": "單日增(減)變動"})
#計算總和
total_net = filtered['單日增(減)變動'].sum()
# 建立一列總和資料 
summary_row = pd.DataFrame([{ 
    '券商代號': '', 
    '券商名稱': '外資買賣超合計', 
    '單日增(減)變動': int(total_net)
}])
# 合併到原表格底部 
filtered= pd.concat([filtered, summary_row], ignore_index=True)

# 建立一列空白 diff 欄位
diff_row = pd.DataFrame([{ 
    '券商代號': '', 
    '券商名稱': 'DIFF', 
    '單日增(減)變動': '', 
}])
filtered2=pd.concat([filtered, diff_row], ignore_index=True)

#先爬取外資合計買賣超
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = uc.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
driver = uc.Chrome(version_main=144, options=options)

url = f"https://justdata.moneydj.com/z/zc/zcl/zcl.djhtm?a=5478&c={now.year}-{now.month}-{now.day}&d={now.year}-{now.month}-{now.day:02d}"
driver.get(url)

body = WebDriverWait(driver, 10).until(
       EC.presence_of_element_located((By.XPATH, "//td[@class='t3r1']"))
)
value = (int(body.text))*1000

#建立一列外資卷商買賣超小計
all_stocker_net=pd.DataFrame([{ 
    '券商代號': '', 
    '券商名稱': '外資券商買賣超小計', 
    '單日增(減)變動': f"{value}", 
}])
filtered3=pd.concat([filtered2,all_stocker_net], ignore_index=True)

# 先把「單日增(減)變動」欄位轉成數字，非數字的轉成 NaN
filtered3["單日增(減)變動"] = pd.to_numeric(filtered3["單日增(減)變動"], errors="coerce")

# 找出「總和」和「外資券商買賣超小計」的值
total_net = filtered3.loc[filtered3["券商名稱"] == "外資買賣超合計", "單日增(減)變動"].values[0]
subtotal = filtered3.loc[filtered3["券商名稱"] == "外資券商買賣超小計", "單日增(減)變動"].values[0]

# 計算 diff
diff_value = total_net - subtotal

# 更新 DataFrame 的 diff 列
filtered3.loc[filtered3["券商名稱"] == "DIFF", "單日增(減)變動"] = diff_value
filtered3["單日增(減)變動"] = filtered3["單日增(減)變動"].astype(int)

sheet_sw=wb_sw.sheets[0]
sheet_sw.range('F13').options(index=False, header=True).value = filtered3
print("外資券商資料已成功寫入")

#列出關注分點名單
target_ids2 = ['884J', '9811', '700G', '9666', '8843', '965K', '8888', '538P','9A9L']
list_df= df[df['券商代號'].isin(target_ids2)]
#把各股數轉換成int
list_df['加總-買進股數'] =list_df['加總-買進股數'].replace(',', '', regex=True).astype(int)
list_df['加總-賣出股數'] =list_df['加總-賣出股數'].replace(',', '', regex=True).astype(int)
list_df['加總net'] =list_df['加總net'].replace(',', '', regex=True).astype(int)

#計算總和
total_buy = list_df['加總-買進股數'].sum()
total_sale = list_df['加總-賣出股數'].sum()
total_net = list_df['加總net'].sum()
#加入加總欄位
all_list_net=pd.DataFrame([{ 
    '券商代號': '',
    '券商名稱': '分點合計', 
    '加總-買進股數':f"{total_buy}",
    '加總-賣出股數':f"{total_sale}",
    '加總net':f"{total_net}", 
}])
list_df=pd.concat([list_df,all_list_net], ignore_index=True)

for col in ['加總-買進股數', '加總-賣出股數', '加總net']:
    list_df[col] = list_df[col].apply(lambda x: f"{int(x):,}" if str(x).replace(',', '').isdigit() else x)

sheet_sw=wb_sw.sheets[0]
sheet_sw.range('J13').options(index=False, header=True).value = list_df
print("回寫完成")

#法人持股明細
import requests
import requests.packages.urllib3
requests.packages.urllib3.disable_warnings()

def fetch_get(url, encoding="big5"):
    try:
        r = requests.get(url, verify=False)
        r.raise_for_status()          
    except Exception as e:
        print(f"錯誤訊息：{e}")
        return None
    else:
        r.encoding = encoding
        return r.text

url=f"https://justdata.moneydj.com/z/zc/zcl/zcl.djhtm?a=5478&c={now.year}-{now.month}-{now.day}&d={now.year}-{now.month}-{now.day:02d}"
data = fetch_get(url)
    
try:
    dfs = pd.read_html(io.StringIO(data))
except Exception as e:
    print(f"錯誤訊息：{e}")
else:
    print(f"表格數量: {len(dfs)}")

#剔除前面的6行的資料
df_news = dfs[2].iloc[5:].reset_index(drop=True)
df_news = df_news.fillna('')

#數值表示為千分位
def format_thousands(x):
    try:
        return f"{int(x):,}"
    except:
        return x  # 保留非數值或空字串

df_news = df_news.map(format_thousands)

#寫回檔案
sheet_sw=wb_sw.sheets[0]
sheet_sw.range('F4').options(index=False, header=False).value = df_news
print("回寫完成")

# 加框線：Excel 的邊框代碼 7=外框, 9=內框 
last_row = sheet_sw.range('F4').expand().last_cell.row 
last_col = sheet_sw.range('F4').expand().last_cell.column 
data_range = sheet_sw.range((3, 6), (last_row, last_col)) # (row, col) 

for border_id in [7,8,9,10,11,12]:
    data_range.api.Borders(border_id).LineStyle = 1 # 實線 
    data_range.api.Borders(border_id).Weight = 2 # 粗細 (2=中等)

#玩股網截圖
driver = uc.Chrome()
url =f"https://www.wantgoo.com/stock/5478"
driver.get(url)
try:
    ad_close = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.CLASS_NAME, "close-btn"))
    )
    ad_close.click()
except:
    print("沒有廣告或找不到關閉按鈕")

driver.get(url)
time.sleep(10)
pic1=driver.find_element(By.CLASS_NAME,'quotes-hd')
pic1.screenshot('./a1.png')
pic2=driver.find_element(By.CLASS_NAME,'quotes-info')
pic2.screenshot('./a2.png')
pic3=driver.find_element(By.CLASS_NAME,'title')
pic3.screenshot('./a3.png')
pic4=driver.find_element(By.CLASS_NAME,'realtime-card')
pic4.screenshot('./a4.png')

driver.close()
sheet_sw=wb_sw.sheets[0]

file_path = os.path.abspath("./a1.png")
sheet_sw.pictures.add(
    file_path,
    name="pic_a1",   # 指定名稱，方便 update
    update=True,
    left=sheet_sw.range("F26").left,
    top=sheet_sw.range("F26").top
)

file_path2 = os.path.abspath("./a2.png")
sheet_sw.pictures.add(
    file_path2,
    name="pic_a2",   # 指定名稱，方便 update
    update=True,
    left=sheet_sw.range("F28").left,
    top=sheet_sw.range("F28").top
)
file_path3 = os.path.abspath("./a3.png")
sheet_sw.pictures.add(
    file_path3,
    name="pic_a3",   # 指定名稱，方便 update
    update=True,
    left=sheet_sw.range("F32").left,
    top=sheet_sw.range("F32").top
)
file_path4 = os.path.abspath("./a4.png")
sheet_sw.pictures.add(
    file_path4,
    name="pic_a4",   # 指定名稱，方便 update
    update=True,
    left=sheet_sw.range("F34").left,
    top=sheet_sw.range("F34").top
)
print("玩股網截圖貼上完成")

wb_sw.save() 
wb_sw.close()
