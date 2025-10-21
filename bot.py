import time
import os
import re
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException, TimeoutException, StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager

# === 1. Lecture du fichier Excel ===
fichier_excel = "excel/DATA_MIG.xlsx"
if not os.path.exists(fichier_excel):
    print(f"❌ Fichier introuvable : {fichier_excel}")
    exit()

df = pd.read_excel(fichier_excel)

# --- Infos de connexion ---
lien_prod = df.iloc[0, 0]
username = df.iloc[0, 2]
password = df.iloc[0, 3]

print(f"Lien: {lien_prod}\nUser: {username}\nPass: {password}")

# === 2. Configuration Chrome (profil temporaire) ===
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--remote-debugging-port=9222")
chrome_options.add_argument("--ignore-certificate-errors")

profil_temp = os.path.join(os.getcwd(), "chrome_temp_profile")
if not os.path.exists(profil_temp):
    os.makedirs(profil_temp)
chrome_options.add_argument(f"--user-data-dir={profil_temp}")

# === 3. Lancement du navigateur ===
try:
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    print("✅ Chrome lancé avec succès")
except Exception as e:
    print("❌ Erreur lors du lancement de Chrome :", e)
    exit()

# === 4. Connexion IFS ===
driver.get(lien_prod)

try:
    WebDriverWait(driver, 15).until(
        EC.element_to_be_clickable((By.XPATH, '//div[text()="Open IFS Cloud"]'))
    ).click()
    print("✅ Clic sur 'Open IFS Cloud' réussi")
except TimeoutException:
    print("⚠️ Bouton 'Open IFS Cloud' non trouvé — peut-être déjà connecté.")

try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))
    driver.find_element(By.ID, "username").send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    driver.find_element(By.ID, "id-ifs-login-btn").click()
    print("✅ Connexion réussie")
except Exception:
    print("ℹ️ Déjà connecté ou interface différente.")

# === 5. Lecture des ordres ===
df_orders = pd.read_excel(fichier_excel, skiprows=3)
df_orders["View détecté"] = None

# === 6. Exécution du premier lien ===
first_link = df_orders.iloc[0]["Lien"]
print(f"\n➡️ Accès au premier lien : {first_link}")
driver.get(first_link)

try:
    WebDriverWait(driver, 30).until_not(
        EC.presence_of_element_located((By.CSS_SELECTOR, ".spinner-wrapper"))
    )
    print("✅ Page du premier lien chargée avec succès")

    # --- Clic sur les initiales ---
    try:
        initials_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.initials"))
        )
        initials_btn.click()
        print("✅ Clic sur le bouton 'initiales' réussi")
    except ElementClickInterceptedException:
        WebDriverWait(driver, 15).until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".spinner-wrapper"))
        )
        initials_btn = driver.find_element(By.CSS_SELECTOR, "div.initials")
        initials_btn.click()
        print("✅ Clic sur le bouton 'initiales' réussi après attente")

    # --- Clic sur 'Debug' ---
    debug_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Debug')]"))
    )
    debug_btn.click()
    print("✅ Clic sur 'Debug' réussi")

    # --- Clic sur 'Page info' ---
    page_info_btn = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Page info')]"))
    )
    page_info_btn.click()
    print("✅ Clic sur 'Page info' réussi")

    # --- Extraction du View ---
    time.sleep(3)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    markdown_div = soup.find("div", {"class": "markdown-text"})

    if markdown_div:
        text_content = markdown_div.get_text(separator=" ").replace("\n", " ")
        all_views = re.findall(r"View\s*:\s*([A-Za-z0-9_]+)", text_content)
        if all_views:
            first_view = all_views[0].strip()
            print(f"🔍 View détecté : {first_view}")
            df_orders.loc[0, "View détecté"] = first_view
        else:
            print("⚠️ Aucun View trouvé.")
    else:
        print("❌ Section 'markdown-text' introuvable.")
except Exception as e:
    print("❌ Erreur pendant la récupération du View:", e)

# === 7. Sauvegarde intermédiaire ===
result_path = "excel/DATA_MIG_result.xlsx"
df_orders.to_excel(result_path, index=False)
print(f"💾 Résultat sauvegardé dans {result_path}")

# === 8. Accès à la page MigrationJob ===
if 'first_view' in locals() and first_view:
    base_url = lien_prod.split("/landing-page")[0]
    migration_path = "/main/ifsapplications/web/page/MigrationJob/Form;path=0.1656053651.381724595.1872552273.1473091230;"
    migration_url = base_url + migration_path
    driver.get(migration_url)
    print(f"➡️ Accès à la page MigrationJob : {migration_url}")

    try:
        WebDriverWait(driver, 30).until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".spinner-wrapper"))
        )
        print("✅ Page MigrationJob chargée avec succès")
        time.sleep(5)

        # --- Clic sur New ---
        retry_count = 0
        while retry_count < 3:
            try:
                new_btn = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[@title='New']"))
                )
                new_btn.click()
                print("✅ Clic sur le bouton 'New' réussi")
                break
            except StaleElementReferenceException:
                print("⚠️ Élément devenu obsolète, tentative de relocalisation...")
                retry_count += 1
                time.sleep(1)
        if retry_count == 3:
            print("❌ Impossible de cliquer sur 'New' après 3 tentatives")

        # --- Remplissage du Job ID ---
        try:
            time.sleep(5)
            job_id_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Job ID']"))
            )

            type_view = df_orders.iloc[0]["Type"]
            view_name = first_view
            suffix = "_X" if type_view.upper() == "EA" else ""
            base_job_id = f"{type_view}_{view_name}{suffix}"
            job_id_final = base_job_id[:30]  # tronquer à 30 caractères max

            job_id_input.clear()
            job_id_input.send_keys(job_id_final)
            print(f"✅ Job ID rempli : {job_id_final}")

        except TimeoutException:
            print("❌ Champ Job ID introuvable sur la page")
        
        # --- Remplissage du Description et View Name ---
        try:
            # Description (max 50 caractères)
            description_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='Description']"))
            )
            description_value = first_view[:50]  # tronquer à 50 caractères max
            description_input.clear()
            description_input.send_keys(description_value)
            print(f"✅ Description rempli : {description_value}")

            # View Name (max 50 caractères, même valeur)
            view_name_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "input[aria-label='View Name']"))
            )
            view_name_value = first_view[:50]
            view_name_input.clear()
            view_name_input.send_keys(view_name_value)
            print(f"✅ View Name rempli : {view_name_value}")

        except TimeoutException:
            print("❌ Champ Description ou View Name introuvable sur la page")


    except TimeoutException:
        print("⚠️ La page MigrationJob n'a pas fini de charger.")
else:
    print("⚠️ Impossible de générer l'URL MigrationJob — View non détecté.")

# === 9. Fin du script ===
driver.quit()
print("🎉 Script terminé avec succès.")
