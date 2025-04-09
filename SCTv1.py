import os
import time
import random
import json
import pandas as pd
import shutil
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

def setup_driver(download_path):
    """Configurar el driver de Chrome con las opciones necesarias"""
    chrome_options = Options()
    
    # Configurar la carpeta de descargas (versión optimizada)
    prefs = {
        "download.default_directory": download_path,
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,  # Para PDFs
        "safebrowsing.enabled": True,  # Mantener la navegación segura
        "download.open_pdf_in_system_reader": False,
        "download.directory_upgrade": True,
        # Asegurarse de que los archivos Excel y PDF se descarguen correctamente
        "browser.helperApps.neverAsk.saveToDisk": "application/vnd.ms-excel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/csv;text/csv;application/pdf"
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Eliminar el mensaje "Un software automatizado está controlando Chrome"
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    
    # Agregar argumento para eliminar el banner de automatización
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    
    # Inicializar el driver
    driver = webdriver.Chrome(options=chrome_options)
    
    # Ejecutar JavaScript para eliminar el banner de automatización
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        })
        """
    })
    
    return driver

def read_credentials(excel_path):
    """Leer credenciales desde el archivo Excel"""
    try:
        df = pd.read_excel(excel_path)
        # Verificar que existan las columnas necesarias
        required_columns = ['CUIT Ingreso', 'CLAVE Ingreso', 'CUIT Contribuyente']
        for col in required_columns:
            if col not in df.columns:
                print(f"Error: El archivo Excel debe contener la columna '{col}'")
                return []
        
        # Convertir CUIT a string para asegurar formato correcto
        df['CUIT Ingreso'] = df['CUIT Ingreso'].astype(str)
        df['CUIT Contribuyente'] = df['CUIT Contribuyente'].astype(str)
        
        return df[required_columns].values.tolist()
    except Exception as e:
        print(f"Error al leer el archivo Excel: {str(e)}")
        return []

def human_typing(element, text):
    """Simular escritura humana tecla por tecla con pausas aleatorias"""
    for char in text:
        time.sleep(random.uniform(0.05, 0.15))
        element.send_keys(char)
        # Pausa adicional aleatoria ocasional para simular pensamiento
        if random.random() < 0.2:  # 20% de probabilidad
            time.sleep(random.uniform(0.1, 0.3))

def login_afip(driver, cuit, clave, wait):
    """Realizar el login en AFIP simulando comportamiento humano"""
    try:
        # Ingresar CUIT
        cuit_input = wait.until(EC.element_to_be_clickable((By.ID, "F1:username")))
        cuit_input.clear()
        # Escribir CUIT tecla por tecla
        human_typing(cuit_input, cuit)
        
        # Pequeña pausa antes de hacer clic en el botón
        time.sleep(random.uniform(0.3, 0.6))
        
        # Click en botón siguiente
        next_button = wait.until(EC.element_to_be_clickable((By.ID, "F1:btnSiguiente")))
        next_button.click()
        
        # Esperar a que aparezca el campo de contraseña
        time.sleep(random.uniform(1.0, 1.5))
        
        # Ingresar Clave
        clave_input = wait.until(EC.element_to_be_clickable((By.ID, "F1:password")))
        clave_input.clear()
        # Escribir clave tecla por tecla
        human_typing(clave_input, clave)
        
        # Pequeña pausa antes de hacer clic en el botón de login
        time.sleep(random.uniform(0.25, 0.5))
        
        # Click en botón de login
        login_button = wait.until(EC.element_to_be_clickable((By.ID, "F1:btnIngresar")))
        login_button.click()
        
        # Esperar a que cargue la página después del login
        time.sleep(random.uniform(2.0, 4.0))
        return True
    except Exception as e:
        print(f"Error en el login: {str(e)}")
        return False

def check_authentication_error(driver):
    """Verificar si hay un error de autenticación en la página"""
    try:
        # Verificar si hay un mensaje de error de autenticación
        error_text = driver.find_element(By.TAG_NAME, "body").text
        if "HTTP Status 401" in error_text and "AUTHENTICATION_ALREADY_PRESENT" in error_text:
            print("Detectado error de autenticación: HTTP Status 401 - AUTHENTICATION_ALREADY_PRESENT")
            return True
        return False
    except:
        return False

def check_authentication_error_message(driver):
    """Verificar si hay un mensaje de error de autenticación en español y refrescar la página usando el botón del navegador"""
    try:
        # Verificar si hay un mensaje de error de autenticación en español
        error_text = driver.find_element(By.TAG_NAME, "body").text
        if "Ha ocurrido un error al autenticar" in error_text or "intente nuevamente" in error_text:
            print("Detectado error de autenticación en español: 'Ha ocurrido un error al autenticar, intente nuevamente.'")
            print("Haciendo clic en el botón de refrescar del navegador...")
            
            try:
                # Intentar encontrar y hacer clic en el botón de refrescar del navegador
                # Primero intentamos con el selector CSS
                try:
                    refresh_button = driver.find_element(By.CSS_SELECTOR, "button[title='Actualizar']")
                    refresh_button.click()
                    print("Botón de refrescar encontrado por CSS y clickeado")
                except:
                    # Si no funciona, intentamos con XPath
                    try:
                        # XPath para el botón de refrescar en Chrome
                        refresh_button = driver.find_element(By.XPATH, "//div[@id='reload-button' or @class='reload-button']")
                        refresh_button.click()
                        print("Botón de refrescar encontrado por XPath y clickeado")
                    except:
                        # Si no podemos encontrar el botón, usamos el método refresh() como respaldo
                        print("No se pudo encontrar el botón de refrescar. Usando driver.refresh() como alternativa")
                        driver.refresh()
                
                time.sleep(random.uniform(3.0, 5.0))
                return True
            except Exception as e:
                print(f"Error al intentar refrescar la página: {str(e)}. Usando driver.refresh() como alternativa")
                driver.refresh()
                time.sleep(random.uniform(3.0, 5.0))
                return True
        return False
    except:
        return False

def navigate_to_sct(driver, wait, cuit, max_attempts=3):
    """Navegar al Sistema de Cuentas Tributarias usando el buscador con reintentos"""
    for attempt in range(1, max_attempts + 1):
        try:
            print(f"Navegando al Sistema de Cuentas Tributarias para CUIT: {cuit} (Intento {attempt}/{max_attempts})")
            
            # Esperar a que la página principal cargue completamente
            time.sleep(random.uniform(1.0, 2.0))
            
            # Buscar el campo de búsqueda
            print("Buscando el campo de búsqueda...")
            search_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='buscadorInput']")))
            
            # Mover el mouse al elemento antes de hacer clic
            actions = ActionChains(driver)
            actions.move_to_element(search_input).pause(random.uniform(0.25, 0.75)).perform()
            search_input.click()
            time.sleep(random.uniform(0.5, 1.0))
            
            # Limpiar el campo de búsqueda
            search_input.clear()
            
            # Escribir "SISTEMA DE CUENTAS TRIBUTARIAS" tecla por tecla
            print("Escribiendo 'SISTEMA DE CUENTAS TRIBUTARIAS' en el buscador...")
            human_typing(search_input, "SISTEMA DE CUENTAS TRIBUTARIAS")
            time.sleep(random.uniform(1.5, 2.0))
            
            # Esperar a que aparezcan los resultados de búsqueda
            print("Esperando resultados de búsqueda...")
            
            # Buscar y hacer clic en el resultado del Sistema de Cuentas Tributarias
            try:
                # Intentar encontrar el resultado específico (primer resultado)
                result_item = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//*[@id='rbt-menu-item-0']/a/div/div/div[1]/div/p")))
                
                # Verificar que el texto del resultado contenga "Sistema de Cuentas Tributarias"
                if "Sistema de Cuentas Tributarias" in result_item.text:
                    print(f"Resultado encontrado: {result_item.text}")
                    
                    # Mover el mouse al elemento antes de hacer clic
                    actions = ActionChains(driver)
                    actions.move_to_element(result_item).pause(random.uniform(0.5, 1.0)).perform()
                    result_item.click()
                    time.sleep(random.uniform(2.0, 3.0))
                    
                    # Cambiar a la nueva pestaña que se abre
                    print("Cambiando a la nueva pestaña...")
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        time.sleep(random.uniform(2.0, 3.0))
                        
                        # Verificar si hay error de autenticación (HTTP Status 401)
                        if check_authentication_error(driver):
                            print("Cerrando pestaña con error HTTP 401 y reintentando...")
                            driver.close()
                            driver.switch_to.window(driver.window_handles[0])
                            time.sleep(random.uniform(2.0, 3.0))
                            continue  # Reintentar
                        
                        # Verificar si hay error de autenticación en español
                        if check_authentication_error_message(driver):
                            print("Se refrescó la página debido al error de autenticación")
                            time.sleep(random.uniform(2.0, 3.0))
                            # No cerramos la pestaña, solo esperamos a que se refresque
                            # Verificamos si después del refresh sigue el error
                            if check_authentication_error_message(driver):
                                print("El error persiste después de refrescar. Cerrando pestaña y reintentando...")
                                driver.close()
                                driver.switch_to.window(driver.window_handles[0])
                                time.sleep(random.uniform(2.0, 3.0))
                                continue  # Reintentar
                        
                        # Intentar cerrar la ventana emergente si existe
                        try_close_popup(driver, wait)
                        
                        return True
                    else:
                        print("No se abrió una nueva pestaña para el Sistema de Cuentas Tributarias")
                        return False
                else:
                    print(f"El resultado no coincide con 'Sistema de Cuentas Tributarias': {result_item.text}")
                    return False
            except TimeoutException:
                print("No se encontró el resultado específico. Reintentando...")
                continue
            
        except Exception as e:
            print(f"Error al navegar al Sistema de Cuentas Tributarias (Intento {attempt}): {str(e)}")
            
            # Si hay pestañas abiertas, cerrarlas y volver a la principal
            if len(driver.window_handles) > 1:
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                time.sleep(random.uniform(1.5, 2.0))
    
    print(f"No se pudo navegar al Sistema de Cuentas Tributarias después de {max_attempts} intentos")
    return False

def try_close_popup(driver, wait, max_attempts=3):
    """Intentar cerrar el popup si existe, con múltiples intentos"""
    for attempt in range(max_attempts):
        try:
            print(f"Intentando cerrar ventana emergente (intento {attempt+1}/{max_attempts})...")
            
            # Esperar a que el popup esté visible
            close_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
                (By.XPATH, "//*[@id='noticias']/div/a")))
            
            # Mover el mouse al elemento antes de hacer clic
            actions = ActionChains(driver)
            actions.move_to_element(close_button).pause(random.uniform(0.5, 1.0)).perform()
            close_button.click()
            time.sleep(random.uniform(0.5, 1.0))
            print("Ventana emergente cerrada correctamente")
            
            # Verificar si el popup realmente se cerró
            try:
                # Si el botón ya no está visible, el popup se cerró correctamente
                WebDriverWait(driver, 2).until_not(EC.presence_of_element_located(
                    (By.XPATH, "//*[@id='noticias']/div/a")))
                print("Confirmado: el popup se cerró correctamente")
                return True
            except:
                # Si el botón sigue visible, intentar nuevamente
                print("El popup parece seguir visible. Reintentando...")
                continue
                
        except Exception as popup_error:
            print(f"No se encontró ventana emergente o no se pudo cerrar: {str(popup_error)}")
            # Si no hay popup o no se puede cerrar, continuamos
            return False
    
    print(f"No se pudo cerrar el popup después de {max_attempts} intentos")
    return False

def select_option_by_text(driver, element, text):
    """Seleccionar una opción de un select por su texto visible"""
    try:
        # Usar el selector correcto basado en el HTML proporcionado
        elementos = driver.find_elements(By.XPATH, f"//select[@name='$PropertySelection']/option[contains(text(), '{text}')]")
        if len(elementos) > 0:
            inner = elementos[0].get_attribute('innerHTML')
            select = Select(element)
            select.select_by_visible_text(inner)
            print(f"Opción seleccionada: {inner}")
            return True
        else:
            print(f"No se encontró la opción con texto: {text}")
            return False
    except Exception as e:
        print(f"Error al seleccionar opción por texto: {str(e)}")
        return False

def select_cuit_contribuyente(driver, wait, cuit_contribuyente):
    """Seleccionar el CUIT Contribuyente del desplegable"""
    try:
        print(f"Seleccionando CUIT Contribuyente: {cuit_contribuyente}")
        
        # Esperar a que el desplegable esté disponible usando el selector correcto
        # El select está dentro de un div con ID "cuitForm"
        select_element = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='cuitForm']/select[@name='$PropertySelection']")))
        
        # Verificar si el CUIT ya está seleccionado
        selected_option = select_element.find_element(By.XPATH, ".//option[@selected='selected']")
        if selected_option and cuit_contribuyente in selected_option.text:
            print(f"El CUIT {cuit_contribuyente} ya está seleccionado")
            
            # Verificar si hay un popup y cerrarlo
            try_close_popup(driver, wait)
            
            return True
        
        # Seleccionar el CUIT del desplegable
        if select_option_by_text(driver, select_element, cuit_contribuyente):
            # Esperar a que la página se actualice después de seleccionar el CUIT
            # El select tiene onchange="javascript:this.form.submit();" que envía el formulario automáticamente
            time.sleep(random.uniform(1.5, 3.0))
            
            # Verificar si hay un diálogo de resubmisión y manejarlo
            try:
                resubmit_button = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(text(), 'Continue') or contains(text(), 'Continuar')]")))
                
                print("Detectado diálogo de resubmisión. Haciendo clic en Continuar...")
                actions = ActionChains(driver)
                actions.move_to_element(resubmit_button).pause(random.uniform(0.3, 0.7)).perform()
                resubmit_button.click()
                time.sleep(random.uniform(2.0, 3.0))
            except:
                # No apareció el diálogo, continuamos normalmente
                pass
            
            # Verificar si hay un popup y cerrarlo después de seleccionar el CUIT
            try_close_popup(driver, wait)
            
            return True
        else:
            # Si no se pudo seleccionar por texto, intentar seleccionar por índice
            try:
                print("Intentando seleccionar por índice...")
                select = Select(select_element)
                options = select_element.find_elements(By.TAG_NAME, "option")
                
                for i, option in enumerate(options):
                    if cuit_contribuyente in option.text:
                        select.select_by_index(i)
                        print(f"CUIT seleccionado por índice: {option.text}")
                        time.sleep(random.uniform(2.0, 3.0))
                        
                        # Verificar si hay un popup y cerrarlo después de seleccionar el CUIT
                        try_close_popup(driver, wait)
                        
                        return True
                
                print(f"No se encontró el CUIT {cuit_contribuyente} en las opciones disponibles")
                return False
            except Exception as e:
                print(f"Error al seleccionar por índice: {str(e)}")
                return False
    except Exception as e:
        print(f"Error al seleccionar CUIT Contribuyente: {str(e)}")
        return False

def expandir_impuestos_y_exportar(driver, wait, cuit_contribuyente, download_path):
    """Expandir impuestos y exportar a XLSX"""
    try:
        print("Expandiendo impuestos y exportando a XLSX...")

        # 1. Hacer clic en el botón de XLSX
        print("Haciendo clic en el botón de XLSX...")
        pdf_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='DataTables_Table_0_wrapper']/div[1]/a[2]")))
        
        # Obtener el nombre del archivo original antes de hacer clic
        original_files = os.listdir(download_path)
        
        # Mover el mouse al elemento antes de hacer clic
        actions = ActionChains(driver)
        actions.move_to_element(pdf_button).pause(random.uniform(0.5, 1.0)).perform()
        pdf_button.click()
        
        # Esperar a que se descargue el archivo
        print("Esperando a que se descargue el archivo XLSX...")
        time.sleep(random.uniform(3.0, 5.0))
        
        # Verificar que se haya descargado el archivo
        new_files = os.listdir(download_path)
        downloaded_files = [f for f in new_files if f not in original_files]
        
        if downloaded_files:
            # Renombrar el archivo descargado
            for file in downloaded_files:
                if file.endswith('.xlsx'):
                    old_path = os.path.join(download_path, file)
                    new_path = os.path.join(download_path, f"{cuit_contribuyente}_pantalla inicial sct.xlsx")
                    
                    # Si ya existe un archivo con ese nombre, eliminarlo
                    if os.path.exists(new_path):
                        os.remove(new_path)
                        
                    os.rename(old_path, new_path)
                    print(f"PDF guardado como: {cuit_contribuyente}_pantalla inicial sct.xlsx")
                    return True
            
            print("Se descargó un archivo, pero no tiene extensión de XLSX (.xlsx)")
            return False
        else:
            print("No se detectó ningún archivo descargado")
            return False
        
    except Exception as e:
        print(f"Error al expandir impuestos y exportar a xlsx: {str(e)}")
        return False

def close_sct_tab(driver):
    """Cerrar la pestaña del SCT y volver a la pestaña principal"""
    try:
        # Cerrar la pestaña actual (SCT)
        print("Cerrando la pestaña del SCT...")
        driver.close()
        
        # Cambiar a la pestaña original (ARCA)
        driver.switch_to.window(driver.window_handles[0])
        time.sleep(random.uniform(1.0, 2.0))
        
        return True
    except Exception as e:
        print(f"Error al cerrar la pestaña del SCT: {str(e)}")
        return False

def logout_afip(driver, wait):
    """Cerrar sesión en AFIP/ARCA"""
    try:
        # Cerrar sesión en ARCA
        print("Cerrando sesión en ARCA...")
        
        # Hacer clic en el icono de usuario
        user_icon = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='userIconoChico']")))
        
        # Mover el mouse al icono de usuario antes de hacer clic
        actions = ActionChains(driver)
        actions.move_to_element(user_icon).pause(random.uniform(0.5, 1.0)).perform()
        user_icon.click()
        time.sleep(random.uniform(1.0, 2.0))
        
        # Hacer clic en el botón de cerrar sesión
        logout_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//*[@id='contBtnContribuyente']/div[6]/button/div/div[2]")))
        
        # Mover el mouse al botón de cerrar sesión antes de hacer clic
        actions = ActionChains(driver)
        actions.move_to_element(logout_button).pause(random.uniform(0.5, 1.0)).perform()
        logout_button.click()
        time.sleep(random.uniform(2.0, 3.0))
        
        print("Sesión cerrada correctamente")
        return True
    except Exception as e:
        print(f"Error al cerrar sesión: {str(e)}")
        return False

def main():
    """Función principal"""
    # Configuración de rutas
    excel_path = r"C:\Users\eze\Downloads\CREDENCIALES.xlsx"
    download_path = r"C:\Users\eze\Downloads"
    
    # Verificar que el archivo Excel existe
    if not os.path.exists(excel_path):
        print(f"Error: No se encontró el archivo Excel en {excel_path}")
        return
    
    # Leer credenciales
    credentials = read_credentials(excel_path)
    if not credentials:
        print("No se pudieron obtener credenciales válidas. Verifique el archivo Excel.")
        return
    
    print(f"Se encontraron {len(credentials)} registros para procesar.")
    
    # Inicializar el driver una sola vez para todos los CUIT
    driver = None
    
    try:
        # Configurar el driver
        driver = setup_driver(download_path)
        
        # Configurar espera explícita
        wait = WebDriverWait(driver, 20)
        
        # Navegar a la página de AFIP
        driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")
        driver.maximize_window()
        
        # Procesar cada CUIT
        for i, (cuit_ingresar, password, cuit_contribuyente) in enumerate(credentials):
            print(f"\nProcesando CUIT: {cuit_ingresar} ({i+1}/{len(credentials)})")
            print(f"CUIT Contribuyente a seleccionar: {cuit_contribuyente}")
            
            try:
                # Si no es el primer CUIT y ya estamos logueados, cerrar sesión primero
                if i > 0:
                    # Cerrar sesión
                    if not logout_afip(driver, wait):
                        print(f"No se pudo cerrar la sesión anterior. Refrescando la página...")
                        driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")
                        time.sleep(random.uniform(3.0, 5.0))
                
                # Login en AFIP
                if not login_afip(driver, cuit_ingresar, password, wait):
                    print(f"No se pudo completar el login para el CUIT {cuit_ingresar}. Continuando con el siguiente.")
                    continue
                
                # Navegar al Sistema de Cuentas Tributarias con manejo de errores de autenticación
                if not navigate_to_sct(driver, wait, cuit_ingresar, max_attempts=3):
                    print(f"No se pudo navegar al Sistema de Cuentas Tributarias para el CUIT {cuit_ingresar}. Continuando con el siguiente.")
                    continue
                
                # Seleccionar el CUIT Contribuyente del desplegable
                if not select_cuit_contribuyente(driver, wait, cuit_contribuyente):
                    print(f"No se pudo seleccionar el CUIT Contribuyente {cuit_contribuyente}. Continuando con el siguiente.")
                    # Cerrar la pestaña del SCT y volver a la pestaña principal
                    close_sct_tab(driver)
                    continue
                
                # NUEVO FLUJO: Expandir impuestos y exportar a PDF
                if not expandir_impuestos_y_exportar(driver, wait, cuit_contribuyente, download_path):
                    print(f"No se pudo expandir impuestos y exportar a PDF para el CUIT {cuit_contribuyente}. Continuando con el siguiente.")
                    # Continuar aunque no se pueda expandir impuestos y exportar a PDF
                
                # Cerrar la pestaña del SCT y volver a la pestaña principal
                if not close_sct_tab(driver):
                    print(f"No se pudo cerrar la pestaña del SCT. Continuando con el siguiente CUIT.")
                    # Si hay un problema, intentar cerrar todas las pestañas excepto la primera
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                        time.sleep(1)
                    driver.switch_to.window(driver.window_handles[0])
                
            except Exception as e:
                print(f"Error procesando CUIT {cuit_ingresar}: {str(e)}")
                
                # Intentar recuperarse para el siguiente CUIT
                try:
                    # Cerrar todas las pestañas excepto la primera
                    while len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                        driver.close()
                        time.sleep(1)
                    driver.switch_to.window(driver.window_handles[0])
                    
                    # Volver a la página de inicio de AFIP
                    driver.get("https://auth.afip.gob.ar/contribuyente_/login.xhtml")
                    time.sleep(random.uniform(3.0, 5.0))
                except Exception as e:
                    print(f"Error al intentar recuperarse: {str(e)}")
    except Exception as e:
        print(f"Error general: {str(e)}")
    finally:
        # Cerrar el navegador al finalizar todos los CUIT
        if driver:
            driver.quit()
    
    print("\nProceso completado.")

if __name__ == "__main__":
    main()