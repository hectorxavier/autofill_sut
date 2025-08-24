import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException, TimeoutException, StaleElementReferenceException
)

# --- Funciones auxiliares ---
def wait_and_click(driver, by, selector, timeout=10):
    def _clickable(d):
        try:
            elem = d.find_element(by, selector)
            if elem.is_displayed() and elem.is_enabled():
                elem.click()
                return True
            return False
        except StaleElementReferenceException:
            return False
    WebDriverWait(driver, timeout).until(_clickable)

def safe_send_keys(driver, campo_id, valor, intentos=3):
    for intento in range(intentos):
        try:
            elem = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, campo_id))
            )
            elem.clear()
            elem.send_keys(valor)
            return True
        except (StaleElementReferenceException, TimeoutException):
            print(f"⚠️ Intento {intento+1}: no se pudo escribir en {campo_id}, reintentando...")
            time.sleep(0.5)
    raise RuntimeError(f"❌ No se pudo escribir en {campo_id} después de {intentos} intentos")

# --- Conexión al Chrome en modo depuración ---
options = webdriver.ChromeOptions()
options.debugger_address = "127.0.0.1:9222"
driver = webdriver.Chrome(options=options)

driver.get("https://sut.trabajo.gob.ec/mrl/empresa/actas/registroActaFrm.xhtml")
time.sleep(3)

# --- Leer Excel ---
df_datos = pd.read_excel(r"autofill_sut\datos.xlsx", dtype=str)
if 'Enviado' not in df_datos.columns:
    df_datos['Enviado'] = ""

# --- Tomar primera fila no enviada ---
fila = df_datos[df_datos['Enviado'] != "Sí"].head(1)
if fila.empty:
    print("❌ No hay registros pendientes para procesar")
    driver.quit()
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
    fondo_reserva = row['Fondo de reserva'].strip().lower()
    valor_fr = row['Valor FR']
    xiii = row['XIII'].strip().lower()
    fecha_xiii = row['Fecha XIII'] if 'Fecha XIII' in row else ''
    obs_xiii = row['Obs XIII'] if 'Obs XIII' in row else ''

    # --- Pasos críticos con reintento ---
    MAX_INTENTOS = 3
    for intento in range(MAX_INTENTOS):
        try:
            # Seleccionar "Identificación"
            tipo_busqueda = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "frmLegal:tipoDiscapacidad_input"))
            )
            driver.execute_script(
                "arguments[0].value='I'; arguments[0].dispatchEvent(new Event('change'));", tipo_busqueda
            )
            driver.execute_script(
                "PrimeFaces.ab({s:'frmLegal:tipoDiscapacidad',e:'valueChange',f:'frmLegal',p:'frmLegal:fldFiltro',u:'frmLegal:fldFiltro',ps:true});"
            )
            time.sleep(0.5)

            # Escribir identificación
            safe_send_keys(driver, "frmLegal:j_idt81", identificacion)

            # Presionar Buscar
            btn_buscar = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "frmLegal:j_idt83"))
            )
            btn_buscar.click()

            # Esperar fila con la identificación
            fila_encontrada = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, f"//td[contains(text(), '{identificacion}')]"))
            )

            # Generar Acta Finiquito
            btn_generar = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "frmLegal:j_idt98:0:j_idt115"))
            )
            btn_generar.click()

            # Validar formulario
            input_ident_form = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "frmLegal:identificacion"))
            )
            if input_ident_form.get_attribute("value") == identificacion:
                print(f"✅ Pasos críticos completados para identificación {identificacion}")
                break
        except Exception as e:
            print(f"⚠️ Intento {intento+1} fallido: {e}")
            time.sleep(1)
    else:
        raise RuntimeError("❌ No se pudieron completar los pasos críticos después de 3 intentos")

    # --- Rellenar remuneración principal ---
    input_remu = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "frmLegal:remuneracion"))
    )
    driver.execute_script("""
        arguments[0].value = arguments[1];
        arguments[0].dispatchEvent(new Event('input'));
        arguments[0].dispatchEvent(new Event('change'));
    """, input_remu, remuneracion)

    # --- Seleccionar causa ---
    def seleccionar_causa(driver, causa_num, intentos=3):
        for intento in range(intentos):
            try:
                fila_causa = driver.find_element(
                    By.XPATH,
                    f"//tbody[@id='frmLegal:j_idt374_data']/tr/td[1][text()='{causa_num}']/.."
                )
                if fila_causa.get_attribute("aria-selected") == "true":
                    print(f"✅ Causa {causa_num} ya aplicada")
                    return True
                ActionChains(driver).move_to_element(fila_causa).click().perform()
                time.sleep(1.5)
                print(f"✅ Causa {causa_num} aplicada correctamente")
                return True
            except Exception as e:
                print(f"⚠️ Intento {intento+1}: no se pudo seleccionar la causa {causa_num}: {e}")
                time.sleep(1)
        raise RuntimeError(f"❌ No se pudo aplicar la causa {causa_num} después de {intentos} intentos")

    seleccionar_causa(driver, causa)

    # --- Agregar Remuneración pendiente ---
    def agregar_remuneracion(driver, salario_pendiente, mes, anio, sueldo_nominal,
                            horas_suplementarias, horas_extraordinarias, horas_nocturnas):
        if float(salario_pendiente) <= 0:
            print("ℹ️ Salario pendiente <= 0, se omite agregar remuneración pendiente")
            return False

        wait_and_click(driver, By.ID, "frmLegal:j_idt574")
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "frmLegal:dttRemu001_data"))
        )

        # --- Desplegar y confirmar Mes ---
        mes_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "frmLegal:dttRemu001:0:j_idt578_label"))
        )
        mes_inicial = mes_label.text
        if mes_inicial != mes:
            mes_label.click()
            print("📌 Lista de Mes desplegada. Por favor selecciona el mes manualmente.")
            WebDriverWait(driver, 300).until(
                lambda d: d.find_element(By.ID, "frmLegal:dttRemu001:0:j_idt578_label").text != mes_inicial
            )
        mes_seleccionado = driver.find_element(By.ID, "frmLegal:dttRemu001:0:j_idt578_label").text
        print(f"✅ Mes confirmado: {mes_seleccionado}")

        # --- Desplegar y confirmar Año ---
        anio_label = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "frmLegal:dttRemu001:0:j_idt580_label"))
        )
        anio_inicial = anio_label.text
        if anio_inicial != anio:
            ActionChains(driver).move_to_element(anio_label).click().perform()
            print("📌 Lista de Año desplegada. Por favor selecciona el año manualmente.")
            WebDriverWait(driver, 300).until(
                lambda d: d.find_element(By.ID, "frmLegal:dttRemu001:0:j_idt580_label").text != anio_inicial
            )
        anio_seleccionado = driver.find_element(By.ID, "frmLegal:dttRemu001:0:j_idt580_label").text
        print(f"✅ Año confirmado: {anio_seleccionado}")

        wait_and_click(driver, By.ID, "frmLegal:dttRemu001:0:btRemu000001")

        # --- Llenar diálogo de remuneración ---
        campos_script = {
            "frmLegal:txtDlg001": salario_pendiente,
            "frmLegal:txtDlg0012": sueldo_nominal
        }
        for campo_id, valor in campos_script.items():
            elem = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, campo_id)))
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input'));
                arguments[0].dispatchEvent(new Event('change'));
            """, elem, valor)

        campos_sendkeys = {
            "frmLegal:txtDlg002": horas_suplementarias,
            "frmLegal:txtDlg004": horas_extraordinarias,
            "frmLegal:txtDlg004n": horas_nocturnas
        }
        for campo_id, valor in campos_sendkeys.items():
            safe_send_keys(driver, campo_id, valor)

        wait_and_click(driver, By.ID, "frmLegal:btnDlg003")
        print("✅ Remuneración pendiente procesada")
        return True

    # --- Fondo de Reserva ---
    def procesar_fondo_reserva(driver, fondo_reserva, valor_fr):
        fondo_reserva = fondo_reserva.strip().lower()

        if fondo_reserva == "si":
            # --- Seleccionar radio 'Sí' vía JS ---
            radio_si = driver.find_element(By.ID, "frmLegal:j_idt590:0")
            driver.execute_script("arguments[0].checked = true; arguments[0].dispatchEvent(new Event('change'));", radio_si)
            print("✅ Radio 'Sí' seleccionado automáticamente (JS)")

            # Presionar botón para agregar Fondo de Reserva
            btn_agregar_fr = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "frmLegal:j_idt598"))
            )
            btn_agregar_fr.click()
            time.sleep(0.5)

            # --- Desplegar y confirmar Mes FR ---
            mes_label = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "frmLegal:j_idt600:0:j_idt602_label"))
            )
            mes_inicial = mes_label.text
            if mes_inicial != mes:
                mes_label.click()
                print("📌 Lista de Mes del FR desplegada. Por favor selecciona el mes manualmente.")
                WebDriverWait(driver, 300).until(
                    lambda d: d.find_element(By.ID, "frmLegal:j_idt600:0:j_idt602_label").text != mes_inicial
                )
            mes_seleccionado = driver.find_element(By.ID, "frmLegal:j_idt600:0:j_idt602_label").text
            print(f"✅ Mes FR confirmado: {mes_seleccionado}")

            # --- Desplegar y confirmar Año FR ---
            anio_label = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "frmLegal:j_idt600:0:j_idt604_label"))
            )
            anio_inicial = anio_label.text
            if anio_inicial != anio:
                ActionChains(driver).move_to_element(anio_label).click().perform()
                print("📌 Lista de Año del FR desplegada. Por favor selecciona el año manualmente.")
                WebDriverWait(driver, 300).until(
                    lambda d: d.find_element(By.ID, "frmLegal:j_idt600:0:j_idt604_label").text != anio_inicial
                )
            anio_seleccionado = driver.find_element(By.ID, "frmLegal:j_idt600:0:j_idt604_label").text
            print(f"✅ Año FR confirmado: {anio_seleccionado}")

            # --- Ingresar valor del Fondo de Reserva ---
            input_valor_fr = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "frmLegal:j_idt600:0:j_idt607"))
            )
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input'));
                arguments[0].dispatchEvent(new Event('change'));
            """, input_valor_fr, valor_fr)
            print(f"✅ Fondo de Reserva procesado: {valor_fr}")

        else:
            # --- Seleccionar radio 'No' vía JS ---
            radio_no = driver.find_element(By.ID, "frmLegal:j_idt590:1")
            driver.execute_script("arguments[0].checked = true; arguments[0].dispatchEvent(new Event('change'));", radio_no)
            print("✅ Fondo de Reserva: No aplica")

    # --- Décima Tercera ---
    def procesar_xiii(driver, xiii, fecha_xiii, obs_xiii):
        import datetime
        import pandas as pd
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC

        def formatear_fecha_xiii(fecha):
            if pd.isna(fecha) or fecha == "":
                return None
            if isinstance(fecha, str) and fecha.isdigit():
                return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(fecha))
            fecha_dt = pd.to_datetime(fecha, errors="coerce")
            if pd.isna(fecha_dt):
                raise ValueError(f"❌ No se pudo interpretar la fecha: {fecha}")
            return fecha_dt

        xiii_flag = xiii.strip().lower() == "si"

        # --- Selección de radio con JS ---
        radio_id = "frmLegal:j_idt616:0" if xiii_flag else "frmLegal:j_idt616:1"
        radio_otra = "frmLegal:j_idt616:1" if xiii_flag else "frmLegal:j_idt616:0"

        driver.execute_script(f"""
            var radio = document.getElementById('{radio_id}');
            radio.checked = true;
            radio.dispatchEvent(new Event('change'));
            var parentDiv = radio.closest('.ui-radiobutton');
            parentDiv.classList.add('ui-state-active');
            parentDiv.querySelector('.ui-radiobutton-icon').classList.remove('ui-icon-blank');
            parentDiv.querySelector('.ui-radiobutton-icon').classList.add('ui-icon-bullet');
            var otra = document.getElementById('{radio_otra}').closest('.ui-radiobutton');
            otra.classList.remove('ui-state-active');
            otra.querySelector('.ui-radiobutton-icon').classList.remove('ui-icon-bullet');
            otra.querySelector('.ui-radiobutton-icon').classList.add('ui-icon-blank');
        """)

        if xiii_flag:
            WebDriverWait(driver, 10).until(
                lambda d: len(d.find_element(By.ID, "frmLegal:dgrDCR0003").find_elements(By.CSS_SELECTOR, "*")) > 0
            )

            # --- Manejo de fecha ---
            if fecha_xiii:
                fecha_dt = formatear_fecha_xiii(fecha_xiii)
                if fecha_dt:
                    calendar_icon = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable(
                            (By.CSS_SELECTOR, "#frmLegal\\:fechaInicioD3_input + button.ui-datepicker-trigger")
                        )
                    )
                    calendar_icon.click()
                    WebDriverWait(driver, 5).until(
                        EC.visibility_of_element_located((By.ID, "ui-datepicker-div"))
                    )

                    # Seleccionar año
                    year_select = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "ui-datepicker-year"))
                    )
                    year_select.send_keys(str(fecha_dt.year))

                    # Seleccionar mes
                    month_select = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "ui-datepicker-month"))
                    )
                    month_select.send_keys(fecha_dt.strftime("%b"))

                    # Seleccionar día
                    mes_attr = fecha_dt.month - 1
                    year_attr = fecha_dt.year
                    dia = fecha_dt.day
                    day_xpath = f"//td[@data-handler='selectDay' and @data-month='{mes_attr}' and @data-year='{year_attr}']/a[text()='{dia}']"

                    WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, day_xpath))
                    ).click()

                    print(f"✅ Fecha Décima Tercera establecida: {fecha_dt.strftime('%d/%m/%Y')}")

            # --- Manejo de observación ---
            if obs_xiii:
                safe_send_keys(driver, "frmLegal:j_idt626", obs_xiii)
                print(f"✅ Observación Décima Tercera: {obs_xiii}")

            # --- Nuevo bloque: registrar salario pendiente en Agosto 2025 ---
            try:
                # Tomar el salario pendiente de la fila 0
                salario_input = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "frmLegal:dttRemu001:0:txtFilaRE0001"))
                )
                # Limpiar USD y espacios
                salario_pendiente = salario_input.get_attribute("value").replace("USD", "").strip()

                # Presionar el botón para habilitar el campo de Agosto 2025
                boton_agregar = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "frmLegal:j_idt1053"))
                )
                boton_agregar.click()
                print("✅ Botón 'Registrar Ingreso' presionado para habilitar Agosto 2025")

                # Ingresar el valor en Agosto 2025
                agosto_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "frmLegal:txtSueldo20257"))
                )
                agosto_input.clear()
                agosto_input.send_keys(salario_pendiente)

                # Disparar evento onchange de PrimeFaces
                driver.execute_script("""
                    var input = document.getElementById('frmLegal:txtSueldo20257');
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                """)

                print(f"✅ Salario pendiente registrado en Agosto 2025: {salario_pendiente}")
            except Exception as e:
                print(f"❌ Error registrando salario pendiente en Agosto 2025: {e}")
                
    # --- Uso dentro del flujo principal ---
    agregar_remuneracion(driver, salario_pendiente, mes, anio, sueldo_nominal,
                        horas_suplementarias, horas_extraordinarias, horas_nocturnas)

    procesar_fondo_reserva(driver, fondo_reserva, valor_fr)

    procesar_xiii(driver, xiii, fecha_xiii, obs_xiii)

    # --- Marcar fila como enviada ---
    df_datos.at[fila.index[0], 'Enviado'] = "Sí"
    df_datos.to_excel(r"autofill_sut\datos.xlsx", index=False)
    print(f"✅ Registro con Identificación {identificacion} procesado")

df_datos.info()

driver.quit()