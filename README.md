import requests, time, os
import csv
from joblib import Parallel, delayed
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from os import path
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver import DesiredCapabilities
import PyPDF2
import pandas as pd
import lxml.html
from lxml.etree import XPath
from threading import Thread
import base64

#improv: intento de multi-hilos

###########################

def cierraDrivers(listaDrivers):
    for driver in listaDrivers:
        driver.close()
        driver.quit()
    return

def creaDrivers():
    print("Iniciando Drivers..")
    op = Options()
    op.page_load_strategy = 'normal'
    op.add_argument("--headless")  # Agrego la opcion de sin cabecera para que no abra el navegador 
    op.add_argument("window-size=1920,1920") # Le asigno un tamaño predeterminado a la ventana del navegador aunque este no sea visible
    op.add_experimental_option('excludeSwitches', ['enable-logging'])   # Quito los comentarios en exceso del modo headless en el terminal

    d0 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")#, desired_capabilities=capabilities)
    d1 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")#, desired_capabilities=capabilities)
    d2 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")
    d3 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")
    d4 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")
    d5 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")
    d6 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")
    d7 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")
    d8 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")
    d9 = webdriver.Chrome(options=op, executable_path=r"C:\chromedriver.exe")

 
    return d0,d1,d2,d3,d4,d5,d6,d7,d8,d9


#################################

def log(archivo, mes, año, nomina, rut_empresa, rs):
    with open('C:\\Users\\gmendoza\\Desktop\\Descarga_PDF\\' + archivo, 'a', newline='') as file:
        mywriter = csv.writer(file, delimiter=';')
        mywriter.writerow([mes, año, nomina, rut_empresa, rs])
    file.close()

    return

def creaDrivers():
    print("Iniciando Drivers..")
    op = Options()
    op.page_load_strategy = 'normal'
    op.add_argument("--headless")  # Agrego la opcion de sin cabecera para que no abra el navegador 
    op.add_argument("window-size=1920,1080") # Le asigno un tamaño predeterminado a la ventana del navegador aunque este no sea visible
    op.add_experimental_option('excludeSwitches', ['enable-logging'])   # Quito los comentarios en exceso del modo headless en el terminal

    
    d0 = driver = webdriver.Chrome()
    #d0 = webdriver.Chrome(options=op, executable_path="C:/users/gmendoza/desktop/descarga_pdf/chromedriver.exe")#, desired_capabilities=capabilities)
    #d1 = webdriver.Chrome(options=op, executable_path="C:/users/gmendoza/desktop/descarga_pdf/chromedriver.exe")#, desired_capabilities=capabilities)
    #d2 = webdriver.Chrome(options=op, executable_path=r"C:/DESCARGAS DE PDF/chromedriver.exe")
    #d3 = webdriver.Chrome(options=op, executable_path=r"C:/DESCARGAS DE PDF/chromedriver.exe")
    #d4 = webdriver.Chrome(options=op, executable_path=r"C:/DESCARGAS DE PDF/chromedriver.exe")

    return d0,d1#,d2,d3,d4#,d5,d6,d7,d8,d9

def login(user, psw, driver): #ME LOGUEO
    CAMPO_RUT = '/html/body/div[1]/div[2]/div[1]/div/form/div/div/div[1]/input'
    CAMPO_PSW = '/html/body/div[1]/div[2]/div[1]/div/form/div/div/div[2]/input'
    BOTON_INGRESAR = '/html/body/div[1]/div[2]/div[1]/div/form/div/div/div[3]/button'
    TEXTO_CERRAR_SESION = '//*[@id="dermarca"]/ul/li[1]/a' 
    SECCION_EMPRESAS = '/html/body/form/div/div[3]/ul/li[3]/div'
    TEXTOP_INGRESO_USUARIOS = '//*[@id="form1"]/h2'
    texto_cerrar_sesion = ''
    for intento in range(5):
        try:
            texto_cerrar_sesion = driver.find_element(By.XPATH,TEXTO_CERRAR_SESION).text
            if "Ayuda" in texto_cerrar_sesion:
                print("LOGIN: Ya logueado:",texto_cerrar_sesion)
                return True  #Si encuentro el texto cerrar sesion, estoy logueado
        except Exception as err: 
            print("ERROR: Login -> No logueado:", texto_cerrar_sesion)
            pass
        driver.delete_all_cookies() #borro cookies
        #checkear si estoy logueado, sino, me logueo
        cargada = False
        while not cargada:
            try:
                driver.get('https://www.previred.com/wPortal/login/login.jsp')
                texto_ingreso_usuarios = driver.find_element(By.XPATH, TEXTOP_INGRESO_USUARIOS)
                if texto_ingreso_usuarios.text == "Ingreso de usuarios":break
            except:
                print("Aun no carga la pagina...., espero 5 segundos para reintentar....")
                time.sleep(5)
                pass

        driver.find_element(By.XPATH,CAMPO_RUT).send_keys(user)
        time.sleep(1)
        driver.find_element(By.XPATH,CAMPO_PSW).send_keys(psw)
        time.sleep(1)
        driver.find_element(By.XPATH,BOTON_INGRESAR).click()
        #seccion_empresa = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, SECCION_EMPRESAS)))
        #seccion_empresa.click()# driver.find_element(By.XPATH,SECCION_EMPRESAS).click()
    return True

def select_company(usr, psw, company_name, company_rut, nombre_usuario, driver):
    if login(usr, psw, driver):
        SECCION_EMPRESAS = '/html/body/form/div/div[3]/ul/li[3]/div'
        TABLA_EMPRESAS = '//*[@id="tablaResultados"]/tbody'
        BOTON_REMUNERACIONES = '/html/body/form/div[2]/div[4]/div/div[2]/div'
        VISTA_BUSCAR_PLANILLAS = '//*[@id="contenido_trabajadores"]/h2'
        MENU_IMPRIMIR_DOCUMENTOS = '//*[@id="menu_sitio"]/li[6]/a'
        MENU_VER_PLANILLAS_PAGADAS = '//*[@id="ver_planillas_pagadas"]'
        EMPRESA_ACTUAL = '//*[@id="barrauser"]/div/div[1]/strong'
        LINK_INICIO = '//*[@id="dermarca"]/ul/li[2]/a'
        BOTON_NUEVA_BUSQUEDA = '//*[@id="buscar"]/span'
        MENSAJE_VALIDACION = '//*[@id="contenido_trabajadores"]/div[2]/div[1]/table/tbody/tr/td[2]/p[1]/strong'
        try:
            empresa_actual = driver.find_element(By.XPATH, EMPRESA_ACTUAL).text
            print("Buscando:", company_rut+company_name, " -- ", empresa_actual)
            if empresa_actual.replace(" ", "") == (company_rut+company_name).replace(" ", ""):
                return True
            
            if nombre_usuario not in empresa_actual:
                time.sleep(1)
                #driver.find_element(By.XPATH, LINK_LISTA_EMPRESAS).click()
                lista_empresas = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, LINK_INICIO)))
                lista_empresas.click()# driver.find_element(By.XPATH,SECCION_EMPRESAS).click()
                print("Usuario:", nombre_usuario, "no esta en", empresa, "le hago click a lista de empresas")
                #exit()
        except Exception as err:
            print("Seleccion de empresa: Error al intentar ingresar a LISTA EMPRESAS")#, err)
            #exit()
            return False
        try:
            time.sleep(1)
            seccion_empresa = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, SECCION_EMPRESAS)))
            seccion_empresa.click()# driver.find_element(By.XPATH,SECCION_EMPRESAS).click()
            WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, TABLA_EMPRESAS))) #original
            #time.sleep(3)
            column_headers = []
            content = lxml.html.fromstring(driver.page_source)
            #print(content)
            tr_elements = content.xpath('//tr') #/html/body/form/div[2]/div[2]/table/tbody/tr[3]
            #print(len(tr_elements))
            for column in tr_elements[0]:
                name = column.text_content()
                column_headers.append((name, []))
            #print(column_headers)
            row_count = 0
            founded = False
            for row in range(1, len(tr_elements)):
                row_count = row_count + 1
                table_tr = tr_elements[row]
                column_count = 0
                fila_completa = []
                for column in table_tr.iterchildren():
                    #print(column.text_content())
                    data = column.text_content()
                    fila_completa.append(data.replace("\t","").replace("\n", ""))#column_headers[column_count][1].append(data)
                    column_count += 1
                #print("Buscando:", company_rut, company_name, "en", fila_completa)
                if fila_completa[1].lower().replace(" ","") == company_rut.lower().replace(" ", "") and fila_completa[2].lower().replace(" ", "") == company_name.lower().replace(" ", ""):
                    #print("Encontrada en ROW", row_count, fila_completa)
                    link = '/html/body/form/div[2]/div[2]/table/tbody/tr['+str(row_count+1)+']/td[1]/button'
                    founded = True
                    break
                if founded:
                    break
            if not founded:
                print("ERROR: No se ha encontrado la empresa", company_rut, company_name)
                exit()
      
        except Exception as err:
            print(err)
            return False
       
        try:
            driver.find_element(By.XPATH, link).click() #aca hago click al link de la empresa que me cosntrui y estoy buscando...
            time.sleep(1)
            driver.find_element(By.XPATH,BOTON_REMUNERACIONES).click()
            time.sleep(2)
            #en algunos casos, al hacer click en "remuneraciones" me muestra una pantalla de aviso de "Valide si la AFP ingresada por usted es la correcta.", debo saltarme esta rs y guardar log
            try:
                mensaje_validacion = driver.find_element(By.XPATH,MENSAJE_VALIDACION)
                if 'Valide si la AFP ingresada por usted es la correcta.' == mensaje_validacion.text:
                    print("OMITIDA:", mes, ano, company_name, company_rut)
                    return "omitir"

            except:pass
            
            #driver.find_element(By.XPATH,MENU_IMPRIMIR_DOCUMENTOS).click() 
            driver.find_element(By.XPATH, ("//a[text()='Imprimir Documentos']")).click() #("//*[contains(text(), 'My Button')]")
            time.sleep(1)
            driver.find_element(By.XPATH,MENU_VER_PLANILLAS_PAGADAS).click()  #si aparece la opcion del menu, le hago click para asegurarme de que me aparezcan los desplegables (esto porque ciertas empresas tienen mensajes intermedios en la misma vista)
            time.sleep(1)
            try:
                boton_nueva_busqueda = driver.find_element(By.XPATH,BOTON_NUEVA_BUSQUEDA) #EN ALGUNOS CASOS, EL CLICK ANTERIOR ME LLEVA DE INMEDIATO A LOS FOLIOS, HACIENDO CLICK EN NUEVA BUSQUEDA ME ASEGURO DE LLEGAR A LOS DESPLEGABLES
                if "Nueva Búsqueda" == boton_nueva_busqueda.text:
                    boton_nueva_busqueda.click()
                    time.sleep(2)
            except: pass
            texto_buscar_planillas = driver.find_element(By.XPATH,VISTA_BUSCAR_PLANILLAS).text
            if "Buscar Planillas Pagadas para Imprimir" in texto_buscar_planillas:
                return True
        except Exception as err:
            print("Select_company: ERROR:", err)
            return False
        print("ERROR: Seleccion de empresa - > EMPRESA NO ENCONTRADA.")
        exit() #debemos hacer algo menos radical para estos casos
        return False
    else:print("ERROR: Seleccion de empresa -> no estoy logueado")
    return False

def select_periodo(usr, psw, company, rut, month, year, nombre_usuario, driver):

    selecciona_compania = select_company(usr, psw, company, rut, nombre_usuario, driver) #If comany is selected ok, do:
    #CHECK_COMPANY
    #print("Ex ok, selecciono mes y año...")
    if selecciona_compania == True:
        LISTA_ANO_ID = 'yearR0'
        LISTA_MES_ID = 'mesR0'
        try:
            ano = Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, LISTA_ANO_ID))))
            #ano = Select(driver.find_element(By.ID,LISTA_ANO_ID))   
            mes = Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, LISTA_MES_ID))))
            #mes = Select(driver.find_element(By.ID,LISTA_MES_ID))
            # select by visible text
        except:return False
        if year != ano.first_selected_option.get_attribute("value"):
            ano.select_by_value(year)
            time.sleep(1)
        # select by value 
        mes.select_by_value(month)
    elif selecciona_compania == 'omitir':
        return "omitir"
    else:
        print("ERROR: Seleccion de periodods -> Empresa NO SELECCIONADA")
        return False
    return True

def list_nomina(usr, psw, empresa, rut, mes, ano, nombre_usuario, driver):
    #time.sleep(1)
    selecciona_periodo = select_periodo(usr, psw, empresa, rut, mes, ano, nombre_usuario, driver)
    if selecciona_periodo == 'omitir': return 'omitir'
    cont_intentos = 0
    while selecciona_periodo == False:
        selecciona_periodo = select_periodo(usr, psw, empresa, rut, mes, ano, nombre_usuario, driver)
        if selecciona_periodo == 'omitir': return 'omitir'
        cont_intentos += 1
        if cont_intentos == 10: return "reiniciar"
        continue
    LISTA_NOMINAS = 'combo_nominas'
    LISTA_TIPO_INSTITUCION = 'combo_tipo_institucion'
    while True:
        try: 
            nominas = Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, LISTA_NOMINAS))))
            #print(nominas.select_by_visible_text)
            time.sleep(1)    
            options = nominas.options
            lista_nominas = []
            for index in options:
                index.click()
                #print(index.text)
                #tipo_institucion = Select(driver.find_element(By.ID,LISTA_TIPO_INSTITUCION))
                if index.text != "Seleccione una Nómina":lista_nominas.append(index.text)
            #print(index)
            return lista_nominas
        except Exception as err:
            #print(err)
            continue
    return lista_nominas

def select_nomina(nomina, driver):
    time.sleep(1)
        
    TEXTO_NO_TIMBRADAS = '//*[@id="mensaje_nuevo_trabajor"]'

    LISTA_NOMINAS = 'combo_nominas'
    LISTA_TIPO_INSTITUCION = 'combo_tipo_institucion'
    BOTON_BUSCAR = '/html/body/form/div/div[2]/div[3]/div[2]/button/span'
    BOTO_ERROR = '/html/body/div[3]/div[3]/div/button'
    
    
    #ACA VERIFICO QUE LA NOMINA QUE ACABO DE INTENTAR DESCARGAR ESTE TIMBRADA, SINO, DEVUELVO UN FLAG Y QUE FORZARA A DEJAR LOG Y PASAR A LA SIGUIENTE NOMINA
    try:
        if "Es posible que sus planillas aún no estén timbradas" in (driver.find_element(By.XPATH, TEXTO_NO_TIMBRADAS).text):
            return "no_timbrada"
    except:pass
    
    try:
        nominas = Select(WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.ID, LISTA_NOMINAS))))
        nominas.select_by_visible_text(nomina)
    except:
        print("ERROR: Select nomina -> no seleccionada:", nomina)
        return False
    time.sleep(1)
    tipo_institucion = Select(driver.find_element(By.ID,LISTA_TIPO_INSTITUCION))
    time.sleep(1)
    tipo_institucion.select_by_visible_text('Todos los Tipos')
    time.sleep(1)
    try: driver.find_element(By.XPATH, BOTO_ERROR).click()
    except:pass
    driver.find_element(By.XPATH,BOTON_BUSCAR).click()
    return True

def create_periods(inicio, fin):
    periodos =  []
    ano_inicio = int(inicio[:4])
    mes_inicio = int(inicio[-2:])
    ano_fin = int(fin[:4])
    mes_fin = int(fin[-2:])
    mes=mes_inicio
    ano=ano_inicio
    while True:
        if mes+1==14:
            ano=ano+1
            mes=1
        if len(str(mes))==1:mess='0'+str(mes)
        else:mess=str(mes)
        periodos.append([str(ano),mess])
        if ano == ano_fin and mes==mes_fin:break
        mes=mes+1

    return periodos

def check_file(file, nomina, nombre_planilla): #Verifico que el archivo descargsdo sea un PDF readable
    existe = os.path.exists(file)
    if existe:
        try:
            f = open(file, "rb")
            PyPDF2.PdfReader(f)
            f.close()
            return True
        except Exception as err:
            f.close()
            print("PDF Invalido:", file, ". Lo borro. ERR:", err)
            #os.remove(file) #si el archivo es invalido, lo borro
            return False
        else:
            pass
        #print("Existe:", existe, "Nomina:", nomina, "Nombre_Planilla:", nombre_planilla)
    return False

def download_pdf(driver, rut, empresa, year, month, nomina):
    periodo=year+month
    time.sleep(3)

    BOTON_NUEVA_BUSQUEDA = '/html/body/form/div/div[2]/div[3]/div[2]/button[1]/span'
    PLANILLAS = '/html/body/form/div/div[2]/div[3]/div[2]/button[2]'
    RUT_PAGADOR = '//*[@id="web_rut_pagador"]'
    WEB_ID_CONTEXT = '//*[@id="web_id_context"]'
    NOMBRE_PLANILLA = '//*[@id="cuerpo_resultado"]/table/tbody/tr[1]/td[1]/div'

    try:
        planillas = str(driver.find_element(By.XPATH, PLANILLAS).get_attribute('id')).replace("planillas_masivas#", "")
    except:
        return False
    rut_pagador = str(driver.find_element(By.XPATH, RUT_PAGADOR).get_attribute('value'))
    web_id_context = str(driver.find_element(By.XPATH, WEB_ID_CONTEXT).get_attribute('value'))
    nombre_planilla = str(driver.find_element(By.XPATH, NOMBRE_PLANILLA).text).replace("Nombre Nómina: ","")



    url = "https://www.previred.com/wEmpresas/CtrlFce"
    s = requests.Session()
    
    payload = "reqName=prgcheckpdflineabatch&web_rol=TE&web_periodo=202204&web_rut_pagador="+rut_pagador+"&web_cod_division=00&web_periodo_desde=&web_periodo_hasta=&web_primera_vez=&web_mes=&web_year=&web_id_nomina=36418561&web_institucion=&web_tipo_institucion=&web_periodo_busqueda=201901&maxtipoinstitucion=&web_glosa_institucion=&web_folios_gravamenes_impresion="+planillas+"&web_centros_costo_impresion=&web_opcion_busqueda=0&web_planillas_masivas=1&pdfmasivo_lineas=5000&pdfmasivo_lineas_cc=10000&habilitacion_pdf_masivo=1&web_id_context="+web_id_context+"&web_prg_destino=&web_rut=19165598&web_accion=login&web_barra_accion=&periodo_busqueda=&web_opcion_pago=&web_accion_pagador=&web_gth=0&web_opcion_menu=&web_opcion_item=nominas_actuales&web_app_modal=&web_func_modal=&web_funes_empresa=&prgSalida=&web_url_chat=http://200.29.71.231/WebAPI802/Previred_Chat/index.jsp?cod=&web_correo_envio=klabrin@fiabilis.cl&web_texto=Imprimir / Descargar Planillas Masivas"
    headers = {
    'authority': 'www.previred.com',
    'accept': '*/*',
    'accept-language': 'es-ES,es;q=0.9',
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'origin': 'https://www.previred.com',
    'referer': 'https://www.previred.com/wEmpresas/CtrlFce',
    'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Google Chrome";v="100"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
    'x-requested-with': 'XMLHttpRequest'
    }

    cokies = driver.get_cookies()
    for cokie in cokies:
        if cokie['name'] == 'JSESSIONID':JSESSIONID =cokie['value']
        if cokie['name'] == 'BIGipServerpool_previred.com_PUB_nuevo':BIG = cokie['value']
        if cokie['name'] == '_ga': GA = cokie['value']
        if cokie['name'] == '_gid': GID = cokie['value']
    for cookie in cokies:
        try:s.cookies.set(name=cookie['name'], value=cookie['value'], expires=1951067631, secure=cookie['secure'], rest=cookie['httpOnly'],path=cookie['path'],domain=cookie['domain'])
        except:s.cookies.set(name=cookie['name'], value=cookie['value'], secure=cookie['secure'], rest=cookie['httpOnly'],path=cookie['path'],domain=cookie['domain'])
    
    response = s.request("POST", url, headers=headers, data=payload)#, cookies = s.cookies)
    time.sleep(1)
    
    cokies = driver.get_cookies()
    for cokie in cokies:
        if cokie['name'] == 'JSESSIONID':JSESSIONID =cokie['value']
        if cokie['name'] == 'BIGipServerpool_previred.com_PUB_nuevo':BIG = cokie['value']
        if cokie['name'] == '_ga': GA = cokie['value']
        if cokie['name'] == '_gid': GID = cokie['value']
    payload = "reqName=prgobtienepdf&web_rol=TE&web_periodo=202204&web_rut_pagador="+rut_pagador+"&web_cod_division=00&web_periodo_desde=&web_periodo_hasta=&web_primera_vez=&web_mes=&web_year=&web_id_nomina=36418561&web_institucion=&web_tipo_institucion=&web_periodo_busqueda=201901&maxtipoinstitucion=&web_glosa_institucion=&web_folios_gravamenes_impresion="+planillas+"&web_centros_costo_impresion=&web_opcion_busqueda=0&web_planillas_masivas=1&pdfmasivo_lineas=5000&pdfmasivo_lineas_cc=10000&habilitacion_pdf_masivo=1&web_id_context="+web_id_context+"&web_prg_destino=&web_rut=19165598&web_accion=login&web_barra_accion=&periodo_busqueda=&web_opcion_pago=&web_accion_pagador=&web_gth=0&web_opcion_menu=&web_opcion_item=nominas_actuales&web_app_modal=&web_func_modal=&web_funes_empresa=&prgSalida=&web_url_chat=http://200.29.71.231/WebAPI802/Previred_Chat/index.jsp?cod=&web_correo_envio=klabrin@fiabilis.cl&web_texto=Imprimir / Descargar Planillas Masivas"
    headers = {
    'authority': 'www.previred.com',
    'accept': '*/*',
    'accept-language': 'es-ES,es;q=0.9',
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'cookie': """JSESSIONID="""+JSESSIONID+"""; BIGipServerpool_previred.com_PUB_nuevo="""+str(BIG)+"""; _ga="""+GA+"""; _gid="""+GID+""";_gat=1""",
    'origin': 'https://www.previred.com',
    'referer': 'https://www.previred.com/wEmpresas/CtrlFce',
    'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Google Chrome";v="100"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
    'x-requested-with': 'XMLHttpRequest'
    }

    # ruta = "D:/WORK/PDFS - ALL/"+rut+"_"+empresa+"/"+periodo+"/"
    ruta = RUTA_DESTINO
    if not os.path.exists(ruta):
        os.makedirs(ruta)

    #codigo original alexis
    #response = s.request("POST", url, headers=headers, data=payload)#, cookies = s.cookies)
    #pdf = s.get('https://www.previred.com/wEmpresas/CtrlPdf')
    #with open(ruta+nombre_planilla+".pdf", 'wb') as f:
    #    f.write(pdf.content)

    response = s.request("POST", url, headers=headers, data=payload)#, cookies = s.cookies)
    pdf_viewer = s.get('https://www.previred.com/wEmpresas/CtrlPdf')
    # pdf viewer en formato binario
    pdf_viewer_content = pdf_viewer.content
    # Get pdf bytes indexes from the pdf viewer
    start = pdf_viewer_content.find(b"base64,") + 7
    end = pdf_viewer_content.find(b"' type='application/pdf'")
    # Decodificar la cadena base64 a bytes
    pdf_bytes = base64.b64decode(pdf_viewer_content[start:end])
    print(start, "->", end)
    print(pdf_bytes)
    with open(ruta+nombre_planilla+".pdf", 'wb') as f:
        f.write(pdf_bytes)

    driver.find_element(By.XPATH, BOTON_NUEVA_BUSQUEDA).click()
    
    if not check_file(ruta+nombre_planilla+".pdf", nomina, nombre_planilla):
        print("ERROR: Download -> Archivo no descargado.", nombre_planilla)

    return

def generate_folio(usr, psw, empresa, rut, period, nombre_usuario, driver):
    mes = period[1]
    ano = period[0]
    planillas = ""
    intentos = 1
    lista_nominas = []
    lista_nominas = list_nomina(usr, psw, empresa, rut, mes, ano, nombre_usuario, driver)
    if lista_nominas == 'reiniciar': return 'reiniciar'
    if lista_nominas == 'omitir':return 'omitir'
    if not lista_nominas:
        print("INFO: Sin nominas para:", ano, mes)
        return "sin_nomina"
    for nomina in lista_nominas:
        print("Descargando:", ano, mes, nomina)
        while True:
            if intentos > 1:print("ERROR: Generate folio, voy denuevo ->", ano, mes, nomina, " -> try:",intentos)
            intentos = intentos + 1
            time.sleep(2)
            if intentos == 10: 
                print("Relogueandome....")
                try:driver.get('https://www.previred.com/wPortal/login/login.jsp')
                except:pass
                intentos = 0
            try:
                resultado = select_nomina(nomina, driver)
                if resultado == "no_timbrada":
                    #aca debo dejar log
                    print("Nomina NO TIMBRADA:", ano,mes, nomina)
                    log("NO_TIMBRADAS.csv", mes, ano, nomina, rut, empresa)
                    #print("Q nominas", len(lista_nominas))
                    if len(lista_nominas) == 1: #aca atajo el error de descargar folios de una RS donde la unica nomina no esta timbrada (descargaba basura, y el scrip de verificacion de PDFs obligaba a reintentar entrando en un loop infinito)
                        print("Es la unica!!!...., paso a la siguiente RS sin descargar nada")
                        #time.sleep(18000)
                        return "unica_no_timbrada"
                    #aca debo pasar a la siguiente nomina con break. Este es para todos los demas casos donde hay una nomina no timbrada, pero no es la unica.
                    break
            except:continue #reinicio el loop   
            #BOTON_DESCARGA_MASIVA = '/html/body/form/div/div[2]/div[3]/div[2]/button[2]'
            #BOTON_NUEVA_BUSQUEDA = '/html/body/form/div/div[2]/div[3]/div[2]/button[1]/span' #antiugua forma
            BOTON_NUEVA_BUSQUEDA = '//*[@id="buscar"]/span'
            PLANILLAS = '/html/body/form/div/div[2]/div[3]/div[2]/button[2]'
            RUT_PAGADOR = '//*[@id="web_rut_pagador"]'
            WEB_ID_CONTEXT = '//*[@id="web_id_context"]'
            NOMBRE_PLANILLA = '//*[@id="cuerpo_resultado"]/table/tbody/tr[1]/td[1]/div'

            try:
                time.sleep(1)
                nuevas =  str(driver.find_element(By.XPATH, PLANILLAS).get_attribute('id')).replace("planillas_masivas#", "")
                #print(nuevas)
                if len(nuevas) < 4: #si el folio es de menos de 4 digitos, error!!!
                    print("ERROR: Generate Folio - > Folio raro:", nuevas)
                    continue #reinicio loop 
                planillas = planillas + nuevas #voy acumulando las planillas
                driver.find_element(By.XPATH, BOTON_NUEVA_BUSQUEDA).click() #me devuelvo
                select_periodo(usr, psw, empresa, rut, mes, ano, nombre_usuario,  driver)
                intentos = 1
                break #me salgo del loop
            except Exception as err:
                print("ERROR: Generate Folio - > No obtuve los folios:", ano,mes, nomina)#, "ERROR:", err)
                #exit()
                select_periodo(usr, psw, empresa, rut, mes, ano, nombre_usuario, driver)
    return planillas
    
def download_big_pdf(driver, folios, ano, mes, rut, company, company_code): #Recibe los folios y los descarga en 1 solo PDF
    BOTON_NUEVA_BUSQUEDA = '/html/body/form/div/div[2]/div[3]/div[2]/button[1]/span'
    PLANILLAS = '/html/body/form/div/div[2]/div[3]/div[2]/button[2]'
    RUT_PAGADOR = '//*[@id="web_rut_pagador"]'
    WEB_ID_CONTEXT = '//*[@id="web_id_context"]'
    NOMBRE_PLANILLA = '//*[@id="cuerpo_resultado"]/table/tbody/tr[1]/td[1]/div'

    rut_pagador = str(driver.find_element(By.XPATH, RUT_PAGADOR).get_attribute('value'))
    web_id_context = str(driver.find_element(By.XPATH, WEB_ID_CONTEXT).get_attribute('value'))

    url = "https://www.previred.com/wEmpresas/CtrlFce"
    s = requests.Session()
    
    payload = "reqName=prgcheckpdflineabatch&web_rol=TE&web_periodo=202204&web_rut_pagador="+rut_pagador+"&web_cod_division=00&web_periodo_desde=&web_periodo_hasta=&web_primera_vez=&web_mes=&web_year=&web_id_nomina=36418561&web_institucion=&web_tipo_institucion=&web_periodo_busqueda=201901&maxtipoinstitucion=&web_glosa_institucion=&web_folios_gravamenes_impresion="+folios+"&web_centros_costo_impresion=&web_opcion_busqueda=0&web_planillas_masivas=1&pdfmasivo_lineas=5000&pdfmasivo_lineas_cc=10000&habilitacion_pdf_masivo=1&web_id_context="+web_id_context+"&web_prg_destino=&web_rut=19165598&web_accion=login&web_barra_accion=&periodo_busqueda=&web_opcion_pago=&web_accion_pagador=&web_gth=0&web_opcion_menu=&web_opcion_item=nominas_actuales&web_app_modal=&web_func_modal=&web_funes_empresa=&prgSalida=&web_url_chat=http://200.29.71.231/WebAPI802/Previred_Chat/index.jsp?cod=&web_correo_envio=klabrin@fiabilis.cl&web_texto=Imprimir / Descargar Planillas Masivas"
    headers = {
    'authority': 'www.previred.com',
    'accept': '*/*',
    'accept-language': 'es-ES,es;q=0.9',
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'origin': 'https://www.previred.com',
    'referer': 'https://www.previred.com/wEmpresas/CtrlFce',
    'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Google Chrome";v="100"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
    'x-requested-with': 'XMLHttpRequest'
    }

    cokies = driver.get_cookies()
    for cokie in cokies:
        if cokie['name'] == 'JSESSIONID':JSESSIONID =cokie['value']
        if cokie['name'] == 'BIGipServerpool_previred.com_PUB_nuevo':BIG = cokie['value']
        if cokie['name'] == '_ga': GA = cokie['value']
        if cokie['name'] == '_gid': GID = cokie['value']

    for cookie in cokies:
        try:s.cookies.set(name=cookie['name'], value=cookie['value'], expires=1951067631, secure=cookie['secure'], rest=cookie['httpOnly'],path=cookie['path'],domain=cookie['domain'])
        except:s.cookies.set(name=cookie['name'], value=cookie['value'], secure=cookie['secure'], rest=cookie['httpOnly'],path=cookie['path'],domain=cookie['domain'])
    
    response = s.request("POST", url, headers=headers, data=payload)#, cookies = s.cookies)
    time.sleep(1)
    
    cokies = driver.get_cookies()
    for cokie in cokies:
        if cokie['name'] == 'JSESSIONID':JSESSIONID =cokie['value']
        if cokie['name'] == 'BIGipServerpool_previred.com_PUB_nuevo':BIG = cokie['value']
        if cokie['name'] == '_ga': GA = cokie['value']
        if cokie['name'] == '_gid': GID = cokie['value']
    payload = "reqName=prgobtienepdf&web_rol=TE&web_periodo=202204&web_rut_pagador="+rut_pagador+"&web_cod_division=00&web_periodo_desde=&web_periodo_hasta=&web_primera_vez=&web_mes=&web_year=&web_id_nomina=36418561&web_institucion=&web_tipo_institucion=&web_periodo_busqueda=201901&maxtipoinstitucion=&web_glosa_institucion=&web_folios_gravamenes_impresion="+folios+"&web_centros_costo_impresion=&web_opcion_busqueda=0&web_planillas_masivas=1&pdfmasivo_lineas=5000&pdfmasivo_lineas_cc=10000&habilitacion_pdf_masivo=1&web_id_context="+web_id_context+"&web_prg_destino=&web_rut=19165598&web_accion=login&web_barra_accion=&periodo_busqueda=&web_opcion_pago=&web_accion_pagador=&web_gth=0&web_opcion_menu=&web_opcion_item=nominas_actuales&web_app_modal=&web_func_modal=&web_funes_empresa=&prgSalida=&web_url_chat=http://200.29.71.231/WebAPI802/Previred_Chat/index.jsp?cod=&web_correo_envio=klabrin@fiabilis.cl&web_texto=Imprimir / Descargar Planillas Masivas"
    headers = {
    'authority': 'www.previred.com',
    'accept': '*/*',
    'accept-language': 'es-ES,es;q=0.9',
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'cookie': """JSESSIONID="""+JSESSIONID+"""; BIGipServerpool_previred.com_PUB_nuevo="""+str(BIG)+"""; _ga="""+GA+"""; _gid="""+GID+""";_gat=1""",
    'origin': 'https://www.previred.com',
    'referer': 'https://www.previred.com/wEmpresas/CtrlFce',
    'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="100", "Google Chrome";v="100"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.75 Safari/537.36',
    'x-requested-with': 'XMLHttpRequest'
    }

    ruta = os.path.join(RUTA_DESTINO,company_code)
    if not os.path.exists(ruta):
        os.makedirs(ruta)

    
    #Codigo original alexis
    #response = s.request("POST", url, headers=headers, data=payload)#, cookies = s.cookies)
    #pdf = s.get('https://www.previred.com/wEmpresas/CtrlPdf')
    #with open(ruta+"/"+rut+"-"+company+"-"+str(ano)+str(mes)+".pdf", 'wb') as f:
    #    f.write(pdf.content)

    response = s.request("POST", url, headers=headers, data=payload)#, cookies = s.cookies)
    pdf_viewer = s.get('https://www.previred.com/wEmpresas/CtrlPdf')
    # pdf viewer en formato binario
    pdf_viewer_content = pdf_viewer.content
    # Get pdf bytes indexes from the pdf viewer
    start = pdf_viewer_content.find(b"base64,") + 7
    end = pdf_viewer_content.find(b"' type='application/pdf'")
    # Decodificar la cadena base64 a bytes
    pdf_bytes = base64.b64decode(pdf_viewer_content[start:end])
    #print(start, "->", end)
    #print(pdf_bytes)
    with open(ruta+"/"+rut+"-"+company+"-"+str(ano)+str(mes)+".pdf", 'wb') as f:
        f.write(pdf_bytes)


    if check_file(ruta+"/"+rut+"-"+company+"-"+str(ano)+str(mes)+".pdf", "nara", "nara"):
        f.close()
        return True

    return False

def create_trabajos(empresas, periods, drivers): #recibe los periodos, drivers y crea los trabajos para ejecutar en paralelo
    jobs = []
    cont = 0
    for period in periods:
        #print(cont)
        jobs.append([period, drivers[cont]])
        #jobs.append([period, drivers])
        cont = cont + 1
        if cont > len(drivers)-1: cont = 0
    return jobs

if __name__ == '__main__':
    empresas = []
    file = pd.read_excel('c:/users/gmendoza/desktop/descarga_pdf/para_descargar.xlsx')
    for idx in file.index:
        empresas.append([file['Rut'][idx], file['Nombre'][idx], file['Usuario'][idx], file['Desde'][idx], file['Hasta'][idx], file['Company'][idx]])

    # El periodo inicial que la página permite es 198001

    # DATOS DE ENTRADA
    RUTA_DESTINO        = "C:/users/gmendoza/desktop/descarga_pdf/PDFS/"
    op = Options()
    op.page_load_strategy = 'normal'
    #op.add_argument("--headless")  # Agrego la opcion de sin cabecera para que no abra el navegador 
    op.add_argument("window-size=1920,1920") # Le asigno un tamaño predeterminado a la ventana del navegador aunque este no sea visible
    op.add_experimental_option('excludeSwitches', ['enable-logging'])   # Quito los comentarios en exceso del modo headless en el terminal}
    capabilities = DesiredCapabilities.CHROME
    capabilities["loggingPrefs"] = {"performance": "ALL"}  # newer: goog:loggingPrefs
   
    # VARIABLES GLOBALES para una sola empresa. Lo he reemplazado par que funcione multiempresa:
    #empresas = [EMPRESA_A_DESCARGAR]

    cont = 0
    while True:
        cont_peri = 0
        if cont == len(empresas): break#si el largo del cont es es igual a len de empresas, estoy fuera, me salgo, llegue al final
        empresa = empresas[cont]
        #print("Descargando:", empresa[0], empresa[1])
        #exit()
        if "Pascale" in empresa[2]:           
            usr = '18.103.977-9'
            psw = 'Pascalemaucor93'
        elif "Nathalia" in empresa[2]:
            usr = '24.684.392-9'
            psw = 'Pfizer2012'
        elif "Katherine" in empresa[2]:
            usr = '19.165.598-2'
            psw = 'Fiabilis2022'
        elif "Pamela" in empresa[2]:
            usr = '13.479.053-9'
            psw = 'IsiFran123'
        elif "Monica" in empresa[2]:
            usr = '6.272.325-4'
            psw = 'victoria2020'
        elif "Gustavo" in empresa[2]:
            usr = '7.439.448-5'
            psw = 'unonortehda'
        elif "Carla" in empresa[2]:
            usr = '10.883.300-9'
            psw = 'co2546'
        elif "Victoria" in empresa[2]:
            usr = '17.170.406-5'
            psw = 'Caffa2023'
        elif "Marcela" in empresa[2]:
            usr = '9.895.739-1'
            psw = '2022CESO'
        elif "Margarita" in empresa[2]:
            usr = '6867705-k'
            psw = 'APROF123'
        elif "TGF" in empresa[2]:
            usr = '77983790-4'
            psw = '19MONTAU14'

        else:
            print("ERROR: Usuario:", empresa[2], "no reconocido.")
            exit()
        period_ini = str(int(empresa[3]))
        period_fin = str(int(empresa[4]))
        periodos = create_periods(period_ini, period_fin)
        #creo driver
        
        #RUTINA DE DESCARGA

        while True:   
            driver = webdriver.Chrome(options=op)    
             #driver = webdriver.Chrome(options=op, executable_path=r"C:/users/gmendoza/desktop/DESCARGA_PDF/chromedriver.exe")#, desired_capabilities=capabilities)       #print(empresa[1], empresa[0])
            #driver.minimize_window()
            if cont_peri == len(periodos):
                cont = cont + 1
                break    
            periodo = periodos[cont_peri]
            ano = periodo[0]
            mes = periodo[1]
            big_planillas = ''
            folios = ''
            try: folios = generate_folio(usr, psw, empresa[1], empresa[0], periodo, empresa[2], driver)
            except:         
                try:
                    print("ERROR: Error desconocido en generate_folio. Cierro Driver y reintento. ")
                    driver.close()
                    driver.quit
                except:pass
            if folios == "unica_no_timbrada":
                cont_peri = cont_peri + 1
                continue
            if folios == "reiniciar":
                cont_peri == cont_peri - 1
                continue
            if folios == "sin_nomina":
                log("SIN_NOMINAS.csv", mes, ano, "sin_planillas", empresa[0], empresa[1])
                cont_peri = cont_peri + 1
                continue
            if folios == 'omitir':
                log('OMITIDAS.csv', mes, ano, "sin_info", empresa[0], empresa[1])
                cont_peri = cont_peri + 1
                continue

            big_planillas = big_planillas+folios


           # HE COMPROBADO QUE SI SE RELANZA EL PROCESO, AL ENCONTRAR UN PDFS CON EL MISMO NOMBRE, ESTE SE SOBREESCRIBE, Y NO HACE APPEND ;)
            
            try:
                #threading.Thread(target=descarga, args=(resultados, cookies, ruta, rut)).start()
                if download_big_pdf(driver, big_planillas, ano, mes, empresa[0], empresa[1], empresa[5]): 
                    cont_peri = cont_peri + 1#solo si la descarga de este periodo fue correcta, avanza de empresa, de lo contrario, se queda pegado en la misma empresa/periodo...
                else:
                    print("Ha habido algun tipo de error al descargar el pdf, se reintentara periodo...")
                    try:
                        driver.close()
                        driver.quit()
                        
                    except:driver = webdriver.Chrome()
                        #driver = webdriver.Chrome(options=op, executable_path=r"C:/users/gmendoza/desktop/DESCARGA_PDF/chromedriver.exe")#, desired_capabilities=capabilities)       #print(empresa[1], empresa[0])
            except:
                print("Except:Ha habido algun tipo de error al descargar el pdf, se reintentara periodo...")
                try:
                    driver.close()
                    driver.quit()
                except:
                    driver = webdriver.Chrome(options=op, executable_path=r"C:/users/gmendoza/desktop/DESCARGA_PDF/chromedriver.exe")#, desired_capabilities=capabilities)       #print(empresa[1], empresa[0])
                    pass
                continue
            
            #cont_peri =+1
        f = open('c:/users/gmendoza/desktop/descarga_pdf/folios.csv', 'a', newline='', encoding="utf-8")


        with f:
            writer = csv.writer(f, delimiter = ";")
            writer.writerow([ano, mes, empresa, folios])
        f.close()
        
        driver.close()
        driver.quit()
    exit()
