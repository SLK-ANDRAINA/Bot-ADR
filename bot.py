import time
import os
import re
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

# === 1. Lecture du fichier Excel ===
fichier_excel = "excel/DATA_MIG.xlsx"

if not os.path.exists(fichier_excel):
    print(f"❌ Fichier introuvable : {fichier_excel}")
    exit()

# Lire tout le fichier Excel
df = pd.read_excel(fichier_excel)

# --- Infos de connexion ---
lien_prod = df.iloc[0, 0]   # 1ère colonne = lien principal
username = df.iloc[0, 2]    # 3e colonne
password = df.iloc[0, 3]    # 4e colonne

print(f"Lien: {lien_prod}\nUser: {username}\nPass: {password}")

# === 2. Lancement du navigateur et connexion ===
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(), options=chrome_options)
driver.get(lien_prod)

# --- Clic sur "Open IFS Cloud" ---
time.sleep(5)
try:
    open_btn = driver.find_element(By.XPATH, '//div[text()="Open IFS Cloud"]')
    open_btn.click()
    print("✅ Clic sur 'Open IFS Cloud' réussi")
except Exception as e:
    print("⚠️ Erreur sur clic Open IFS Cloud:", e)

# --- Connexion ---
time.sleep(5)
try:
    driver.find_element(By.ID, "username").send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    driver.find_element(By.ID, "id-ifs-login-btn").click()
    print("✅ Connexion réussie")
except Exception as e:
    print("⚠️ Erreur de connexion:", e)

# === 3. Lecture des ordres de traitement ===
# On suppose que les entêtes sont à la ligne 4 → index 3
df_orders = pd.read_excel(fichier_excel, skiprows=3)

# On prépare une nouvelle colonne pour stocker le View détecté
df_orders["View détecté"] = None

# === 4. Exécution des ordres ===
first_link = df_orders.iloc[0]["Lien"]
print(f"\n➡️ Accès au premier lien : {first_link}")

try:
    driver.get(first_link)
    print("✅ Page du premier lien chargée avec succès")

    # --- Clic sur les initiales (ex: "AN") ---
    time.sleep(5)
    initials_btn = driver.find_element(By.XPATH, "//div[contains(@class,'initials')]")
    initials_btn.click()
    print("✅ Clic sur le bouton 'initiales' réussi")

    # --- Clic sur 'Debug' ---
    time.sleep(2)
    debug_btn = driver.find_element(By.XPATH, "//button[contains(.,'Debug')]")
    debug_btn.click()
    print("✅ Clic sur 'Debug' réussi")

    # --- Clic sur 'Page info' ---
    time.sleep(2)
    page_info_btn = driver.find_element(By.XPATH, "//button[contains(.,'Page info')]")
    page_info_btn.click()
    print("✅ Clic sur 'Page info' réussi")

    # --- Attente du contenu ---
    time.sleep(3)
    html_content = driver.page_source
    soup = BeautifulSoup(html_content, "html.parser")
    markdown_div = soup.find("div", {"class": "markdown-text"})

    if markdown_div:
        text_content = markdown_div.get_text(separator=" ").replace("\n", " ")

        # --- Extraction du premier View ---
        match = re.search(r"View:\s*([A-Z0-9_]+)\s*Component:", text_content, re.IGNORECASE)
        if match:
            view_name = match.group(1).strip()
            print(f"🔍 View détecté : {view_name}")
            df_orders.loc[0, "View détecté"] = view_name
        else:
            print("⚠️ Aucun View trouvé. Vérifie le format de la page info.")
            print("Extrait du texte :", text_content[:300])
    else:
        print("❌ Impossible de trouver la section 'markdown-text'.")

except Exception as e:
    print("❌ Erreur pendant la récupération du View:", e)

# === 5. Sauvegarde du résultat ===
result_path = "excel/DATA_MIG_result.xlsx"
df_orders.to_excel(result_path, index=False)
print(f"💾 Résultat sauvegardé dans {result_path}")

# === 6. Fermeture du navigateur ===
driver.quit()
print("🎉 Script terminé avec succès.")
