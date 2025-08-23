import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains

# --- 1. Conectarse al Chrome abierto en modo depuración ---
options = webdriver.ChromeOptions()
options.debugger_address = "127.0.0.1:9222"  # Chrome abierto con --remote-debugging-port=9222
driver = webdriver.Chrome(options=options)

# --- 2. Ir al formulario principal ---
driver.get("https://sut.trabajo.gob.ec/mrl/empresa/actas/registroActaFrm.xhtml")
time.sleep(5)

# --- 3. Leer Excel ---
df_datos = pd.read_excel(r"autofill_sut\datos.xlsx", dtype=str)  # todas las columnas como texto
if 'Enviado' not in df_datos.columns:
    df_datos['Enviado'] = ""

# --- 4. Tomar únicamente la primera fila no enviada ---
fila = df_datos[df_datos['Enviado'] != "Sí"].head(1)

if fila.empty:
    print("❌ No hay registros pendientes para procesar")
else:
    row = fila.iloc[0]

    identificacion = row['Identificacion']
    remuneracion = row['Remuneracion']
    causa = row['Causa']
    mes = row['Mes']
    anio = row['Año']
    salario_pendiente = row['Salario_pendiente']
    sueldo_nominal = row['Sueldo_nominal']
    horas_suplementarias = row['Horas_suplementarias']
    horas_extraordinarias = row['Horas_extraordinarias']
    horas_nocturnas = row['Horas_nocturnas']

    # --- 4a. Seleccionar "Identificación" en el combo de búsqueda ---
    tipo_busqueda = driver.find_element(By.ID, "frmLegal:tipoDiscapacidad_input")
    driver.execute_script("arguments[0].value = 'I';", tipo_busqueda)
    driver.execute_script(
        "PrimeFaces.ab({s:'frmLegal:tipoDiscapacidad',e:'valueChange',f:'frmLegal',p:'frmLegal:fldFiltro',u:'frmLegal:fldFiltro',ps:true});"
    )
    time.sleep(1)

    # --- 4b. Escribir la identificación ---
    campo_ident = driver.find_element(By.ID, "frmLegal:j_idt81")
    campo_ident.clear()
    campo_ident.send_keys(identificacion)

    # --- 4c. Presionar Buscar ---
    btn_buscar = driver.find_element(By.ID, "frmLegal:j_idt83")
    btn_buscar.click()
    time.sleep(5)

    # --- 4d. Paso intermedio: Generar Acta Finiquito ---
    btn_generar = driver.find_element(By.ID, "frmLegal:j_idt98:0:j_idt115")
    btn_generar.click()
    time.sleep(5)

    # --- 4e. Validar que la identificación coincide ---
    input_ident_form = driver.find_element(By.ID, "frmLegal:identificacion")
    ident_formulario = input_ident_form.get_attribute("value")
    if ident_formulario != identificacion:
        print(f"❌ Error: la identificación en el formulario ({ident_formulario}) no coincide con la del Excel ({identificacion})")
        driver.quit()
        raise ValueError("Identificación no coincide")
    else:
        print(f"✅ Identificación validada: {ident_formulario}")
        
    # --- 4e. Rellenar remuneración principal ---
    input_remu = driver.find_element(By.ID, "frmLegal:remuneracion")
    driver.execute_script(f"arguments[0].value = '{remuneracion}';", input_remu)

    # --- 4f. Seleccionar causa ---
    fila_causa = driver.find_element(By.XPATH, f"//tbody[@id='frmLegal:j_idt374_data']/tr/td[text()='{causa}']")
    ActionChains(driver).move_to_element(fila_causa).click().perform()
    time.sleep(1)

    # --- 4g. Agregar remuneración pendiente ---
    btn_agregar = driver.find_element(By.ID, "frmLegal:j_idt574")
    btn_agregar.click()
    time.sleep(1)

    # Ingresar Mes y Año
    mes_label = driver.find_element(By.ID, "frmLegal:dttRemu001:0:j_idt578_label")
    anio_label = driver.find_element(By.ID, "frmLegal:dttRemu001:0:j_idt580_label")
    driver.execute_script(f"arguments[0].innerText='{mes}';", mes_label)
    driver.execute_script(f"arguments[0].innerText='{anio}';", anio_label)
    time.sleep(1)

    # --- 4h. Componer remuneración ---
    btn_componer = driver.find_element(By.ID, "frmLegal:dttRemu001:0:btRemu000001")
    btn_componer.click()
    time.sleep(1)

    # --- 4i. Llenar el diálogo ---
    driver.execute_script(f"arguments[0].value='{salario_pendiente}';", driver.find_element(By.ID, "frmLegal:txtDlg001"))
    driver.execute_script(f"arguments[0].value='{sueldo_nominal}';", driver.find_element(By.ID, "frmLegal:txtDlg0012"))
    driver.execute_script(f"arguments[0].value='{horas_suplementarias}';", driver.find_element(By.ID, "frmLegal:txtDlg002"))
    driver.execute_script(f"arguments[0].value='{horas_extraordinarias}';", driver.find_element(By.ID, "frmLegal:txtDlg004"))
    driver.execute_script(f"arguments[0].value='{horas_nocturnas}';", driver.find_element(By.ID, "frmLegal:txtDlg004n"))
    time.sleep(1)

    # --- 4j. Guardar diálogo ---
    btn_aceptar = driver.find_element(By.ID, "frmLegal:btnDlg003")
    btn_aceptar.click()
    time.sleep(2)

    # --- 4k. Marcar fila como enviada ---
    df_datos.at[fila.index[0], 'Enviado'] = "Sí"
    df_datos.to_excel(r"autofill_sut\datos.xlsx", index=False)
    print(f"✅ Registro con Identificación {identificacion} procesado")

driver.quit()
