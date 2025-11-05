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
group_id_value = str(df.iloc[0, 1]).strip()
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
#chrome_options.add_argument("--headless=new")       # exécuter Chrome en arrière-plan
#chrome_options.add_argument("--window-size=1920,1080")  # simuler une 


# 👉 Profil temporaire
profil_temp = os.path.join(os.getcwd(), "chrome_temp_profile")
os.makedirs(profil_temp, exist_ok=True)
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

# --- Connexion ---
try:
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))
    driver.find_element(By.ID, "username").send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    time.sleep(2)
    driver.find_element(By.ID, "id-ifs-login-btn").click()
    print("✅ Connexion réussie")
except Exception:
    print("ℹ️ Déjà connecté ou interface différente.")

# === 5. Lecture des ordres ===
df_orders = pd.read_excel(fichier_excel, skiprows=3)
df_orders["View détecté"] = None

# === 6. Exécution du traitement ===
for idx, row in df_orders.iterrows():
    first_link = row["Lien"]
    type_traitement = row["Type"]
    predefined_view = str(row.get("View", "")).strip()

    # ✅ 6.1. Si le View est déjà défini dans Excel
    if predefined_view and predefined_view.lower() != "nan":
        print(f"⚡ View déjà défini dans Excel : {predefined_view}")
        first_view = predefined_view
        df_orders.loc[idx, "View détecté"] = first_view

    # 🔍 Sinon, récupération automatique via le lien
    else:
        print("🔍 Aucun View fourni — récupération via le lien...")
        driver.get(first_link)
        time.sleep(4)

        try:
            time.sleep(4)
            WebDriverWait(driver, 30).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".spinner-wrapper"))
            )
            print("✅ Page chargée avec succès")

            # Attente que le spinner disparaisse avant de cliquer
            WebDriverWait(driver, 30).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".spinner-wrapper"))
            )
            time.sleep(3)  # petite pause pour stabilité

            initials_btn = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "div.initials"))
            )
            driver.execute_script("arguments[0].click();", initials_btn)
            print("✅ Clic sur le bouton 'initiales' réussi (via JS)")


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
                    df_orders.loc[idx, "View détecté"] = first_view
                else:
                    print("⚠️ Aucun View trouvé.")
                    first_view = None
            else:
                print("❌ Section 'markdown-text' introuvable.")
                first_view = None

        except Exception as e:
            print("❌ Erreur pendant la récupération du View:", e)
            first_view = None

    # === 7. Sauvegarde intermédiaire ===
    result_path = "excel/DATA_MIG_result.xlsx"
    df_orders.to_excel(result_path, index=False)
    print(f"💾 Résultat sauvegardé dans {result_path}")

    # === 8. Accès à la page MigrationJob ===
    if first_view:
        base_url = lien_prod.split("/landing-page")[0]
        migration_url = base_url + "/main/ifsapplications/web/page/MigrationJob/Form;path=0.1656053651.381724595.1872552273.1473091230;"
        driver.get(migration_url)
        print(f"➡️ Accès à la page MigrationJob : {migration_url}")
        time.sleep(10)
        try:
            WebDriverWait(driver, 30).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".spinner-wrapper"))
            )
            print("✅ Page MigrationJob chargée avec succès")

            time.sleep(2)
            # Clic sur bouton "New"
            new_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@title='New']"))
            )
            new_btn.click()
            print("✅ Clic sur le bouton 'New' réussi")

            # ==== Remplissage automatique ====
            time.sleep(3)
            job_id_value = f"{type_traitement}_{first_view}"[:20]
            if type_traitement == "EA":
                job_id_value += "_X"

            # Champ Job ID
            job_id_field = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//input[@aria-label='Job ID']"))
            )
            job_id_field.send_keys(job_id_value)
            print(f"🆔 Job ID défini : {job_id_value}")

            # Champ Description
            desc_field = driver.find_element(By.XPATH, "//input[@aria-label='Description']")
            desc_field.send_keys(first_view)
            print("📝 Description ajoutée")

            # Champ Procedure Name
            proc_field = driver.find_element(By.XPATH, "//input[@aria-label='Procedure Name']")
            if type_traitement == "EA":
                procedure_value = "EXCEL_MIGRATION"
            else:
                procedure_value = first_view
            proc_field.send_keys(procedure_value[:50])
            print(f"⚙️ Procedure Name défini : {procedure_value}")

            # Champ View Name
            view_field = driver.find_element(By.XPATH, "//input[@aria-label='View Name']")
            view_field.send_keys(first_view)
            print("👁️ View Name ajouté")

            # Champ Group ID
            group_field = driver.find_element(By.XPATH, "//input[@aria-label='Group ID']")
            group_field.send_keys(group_id_value)
            print(f"👥 Group ID ajouté depuis Excel : {group_id_value}")

            # === Clic sur 'Save' ===
            save_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//span[normalize-space()='Save']"))
            )
            save_button.click()
            print("💾 Clic sur 'Save' réussi")

            time.sleep(5)
            WebDriverWait(driver, 30).until_not(
                EC.presence_of_element_located((By.CSS_SELECTOR, ".spinner-wrapper"))
            )
            print("✅ Enregistrement du job terminé avec succès")

            # --- Mise à jour Excel ---
            for col in ["Job ID", "Statut", "Group ID"]:
                if col not in df_orders.columns:
                    df_orders[col] = ""

            df_orders.at[idx, "Job ID"] = job_id_value
            df_orders.at[idx, "Statut"] = "OK"
            df_orders.at[idx, "Group ID"] = group_id_value

            df_orders = df_orders[
                ["Ordre", "Lien", "Type", "View", "View détecté", "Job ID", "Group ID", "Statut"]
            ]
            df_orders.to_excel(result_path, index=False)
            print(f"💾 Excel mis à jour avec succès : {result_path}")

        except Exception as e:
            print("❌ Erreur pendant la création du job migration:", e)
            df_orders.loc[idx, "Statut"] = "ERREUR"
            df_orders.to_excel(result_path, index=False)
            print("🚨 Statut mis à jour dans Excel : ERREUR")

    else:
        print("⚠️ View non détecté, impossible de créer le job migration.")

# === 9. Fin du script ===
driver.quit()
print("🎉 Script terminé avec succès.")