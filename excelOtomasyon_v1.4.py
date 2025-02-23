from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import os 
import re 

 # Eğer regex bir eşleşme bulursa, formatı düzeltip float'a çevir, yoksa 0.0 ata
def format_number(num_match):
    if num_match:
        num_str = num_match.group()
        num_str = num_str.replace(".", "")  # Binlik ayıracı kaldır (örnek: "3.642,80" → "3642,80")
        num_str = num_str.replace(",", ".")  # Ondalık ayıracı noktaya çevir (örnek: "3642,80" → "3642.80")
        return float(num_str)  # Sayıyı float'a çevir
    return 0.0

# WebDriver başlat
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)

# UBYS giriş sayfasına git
driver.get("https://portal************")  # Giriş URL'sini doğru yaz
# Kullanıcı adı ve şifre kutularını bul
username_input = driver.find_element(By.ID, "****name")  # Gerçek ID'yi kontrol et
password_input = driver.find_element(By.ID, "****pass")

# Kullanıcı adı ve şifreyi gir
username_input.send_keys("*********")
password_input.send_keys("*********")
password_input.send_keys(Keys.RETURN)  # Enter'a basarak giriş yap

# Sayfanın yüklenmesini bekleyin
time.sleep(5) 

# Transkript sayfasına git
driver.get("https://portal**********")  # Gerçek URL'yi gir

#Sayfanın yüklenmesini bekle
time.sleep(5)

# **1. YIL Seçimi (Eğer gerekliyse)**
try:
    year_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='*****']//a[contains(text(), '2024')]")))
    year_element.click()
    time.sleep(2)  # Sayfanın güncellenmesini bekleyelim
except:
    print("Yıl seçimi gerekli değil veya bulunamadı.")

# 2.Ocak ayını seç**
december_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[*******]")))
december_element.click()
time.sleep(2)  # Sayfanın yenilenmesini bekleyelim

# **Sayfa yenilendiği için öğeyi tekrar bul**
december_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[********]")))
december_element.click()
time.sleep(3)


# Tablodaki tüm satırları seç (tbody içindeki tr etiketleri)
rows = driver.find_elements(By.*****, "tbody tr")

data = []
for row in rows:
    cells = row.find_elements(By.******, "td")
    if len(cells) >= 2:  # En az 3 sütun olmalı
        label = cells[0].text.strip()  # İlk sütun: LABEL
        yield_value = cells[1].text.strip()  # İkinci sütun: YIELD
      #  kwp = cells[2].text.strip()  # Üçüncü sütun: kWh/kWp
        
        # "kWh/h" veya diğer yazıları temizleyip sadece sayı almak
        yield_value = re.search(r"[\d,\.]+", yield_value)  # Sayıyı bul
      #  kwp = re.search(r"[\d,\.]+", kwp)

        yield_value = format_number(yield_value)
     #   kwp = format_number(kwp)

        data.append({
            "LABEL": label,
            "YIELD (kWh/h)": yield_value,
     #       "kWh/kWp": kwp
        })


# Veriyi Excel'e kaydet
df = pd.DataFrame(data)
df.to_excel("yield_cactus_2024aralık.xlsx", index=False)

print("Tablo başarıyla kaydedildi!")

# Kullanıcının İndirilenler klasörünü al
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads", "*******_2024aralık.xlsx")

# Excel dosyasını İndirilenler klasörüne kaydet
df.to_excel(downloads_folder, index=False)

print(f"Tablo başarıyla kaydedildi: {downloads_folder}")

# Dosyayı açks
os.startfile(downloads_folder)

# Tarayıcıyı kapat
driver.quit()

