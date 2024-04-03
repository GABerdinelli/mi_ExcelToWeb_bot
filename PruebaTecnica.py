import xlwings as xw
import smtplib
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime

def enviar_correo(correo_electronico, asunto, mensaje):
    # Credenciales SMTP
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_username = "tu_correo_electronico"
    smtp_password = "tu_contrasena"

    with smtplib.SMTP(smtp_server, smtp_port) as smtp:
        smtp.starttls()
        smtp.login(smtp_username, smtp_password)
        smtp.sendmail(smtp_username, correo_electronico, f"Subject: {asunto}\n{mensaje}")

def procesar_excel():
    wb = xw.Book("Base_Seguimiento_Observ_Auditoría_al_30042021.xlsx")
    sheet = wb.sheets["Hoja1"]

    chrome_options = Options()
    chrome_options.add_argument("webdriver.chrome.driver=chromedriver.exe")

    driver = webdriver.Chrome(options=chrome_options)

    try:
        for row in range(2, sheet.used_range.last_cell.row + 1):
            estado = sheet.range(row, 10).value

            if estado == "Regularizado":
                proceso_xpath = "/html/body/div[1]/div/div/form/div/div[1]/div/div[3]/div/select"
                tipo_de_riesgo_xpath = "/html/body/div[1]/div/div/form/div/div[1]/div/div[4]/div/input"
                severidad_observacion_xpath = "/html/body/div[1]/div/div/form/div/div[1]/div/div[5]/div/select"
                responsable_xpath = "/html/body/div[1]/div/div/form/div/div[1]/div/div[6]/div/input"
                fecha_compromiso_xpath = "/html/body/div[1]/div/div/form/div/div[1]/div/div[7]/div/input"
                observacion_xpath = "/html/body/div[1]/div/div/form/div/div[1]/div/div[8]/div/textarea"

                proceso = sheet.range(row, 1).value
                tipo_de_riesgo = sheet.range(row, 3).value
                severidad_observacion = sheet.range(row, 4).value
                responsable = sheet.range(row, 7).value
                fecha_compromiso = sheet.range(row, 6).value
                observacion = sheet.range(row, 2).value

                driver.get("https://roc.myrb.io/s1/forms/M6I8P2PDOZFDBYYG")
                
                proceso_select = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, proceso_xpath)))
                proceso_select.send_keys(proceso)

                tipo_de_riesgo_input = driver.find_element(By.XPATH, tipo_de_riesgo_xpath)
                tipo_de_riesgo_input.send_keys(tipo_de_riesgo)

                severidad_observacion_select = driver.find_element(By.XPATH, severidad_observacion_xpath)
                severidad_observacion_select.send_keys(severidad_observacion)

                responsable_input = driver.find_element(By.XPATH, responsable_xpath)
                responsable_input.send_keys(responsable)

                fecha_compromiso_input = driver.find_element(By.XPATH, fecha_compromiso_xpath)
                fecha_compromiso_input.send_keys(fecha_compromiso.strftime("%d/%m/%Y"))  # Corrección aquí

                observacion_input = driver.find_element(By.XPATH, observacion_xpath)
                observacion_input.send_keys(observacion)

                driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

    finally:
        driver.quit()
        wb.save()
        wb.close()

procesar_excel()
