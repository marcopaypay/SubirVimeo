from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import pyodbc
from datetime import datetime

server = 'PC-CASA\SQLEXPRESS'  
database = 'SICEX_INSPECCION'
username = 'sa'
password = 'daniel261516'

# Crea la conexión
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 17 for SQL Server};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password}'
)
cursor = conn.cursor()

service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()

driver = webdriver.Chrome(service=service, options=options)

try:
    driver.get("https://cej.pj.gob.pe/")

    time.sleep(5)
    WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH,
                                        '//*[@id="distritoJudicial"]/option[19]')))\
    .click()

    WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH,
                                        '//*[@id="organoJurisdiccional"]/option[3]')))\
    .click()

    WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH,
                                        '//*[@id="anio"]/option[2]')))\
    .click()

    WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH,
                                        '//*[@id="especialidad"]/option[8]')))\
    .click()

    time.sleep(20)
    #Guardar la data de la priemra tabla
    try:
        select_element = driver.find_element_by_id("distritoJudicial")
        id_distrito_judicial = int(select_element.get_attribute("value"))
        print("Distrituto Judicial")
        print(id_distrito_judicial)
        select_element2 = driver.find_element_by_id("organoJurisdiccional")
        id_instancia = int(select_element2.get_attribute("value"))
        print("Instancia")
        print(id_instancia)
        select_element3 = driver.find_element_by_id("especialidad")
        id_especialidad = int(select_element3.get_attribute("value"))
        print("Especializada")
        print(id_especialidad)
        select_element4 = driver.find_element_by_id("anio")
        id_ano = int(select_element4.get_attribute("value"))
        print("Año")
        print(id_ano)
        imputCorelativo= driver.find_element_by_id("numeroExpediente")
        correlativo = int(imputCorelativo.get_attribute("value"))
        print("Correlativo")
        print(correlativo)
    except Exception as e:
        print(e)
    try:
        insert_query = '''
                        INSERT INTO EXPEDIENTE_CABECERA (id_distrito_judicial,id_instancia, id_especialidad, año, correlativo)
                        OUTPUT inserted.id_secuencia
                        VALUES (?,?,?,?,?);
                        '''
        cursor.execute(insert_query, id_distrito_judicial, id_instancia, id_especialidad, id_ano, correlativo)
        row = cursor.fetchone()
        id_secuencia = row.id_secuencia
        cursor.commit()
    except Exception as e:
        print(e)

    WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH,
                                        '//*[@id="consultarExpedientes"]')))\
    .click()

    WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.XPATH,
                                        '//*[@id="command"]/button')))\
    .click()
    #Primero registro de arriba
    registrosPrimero=[]
    try:
        datosDelExpediente = driver.find_elements_by_xpath("//div[@class='divRepExp']/div")
    except Exception as e:
        print(e)
    for i in range(0, len(datosDelExpediente), 2):
        try:
            nombre = datosDelExpediente[i].text
            valor = datosDelExpediente[i+1].text
            registrosPrimero.append({"name": nombre, "valor": valor})
        except Exception as e:
            print(e)
    print("Datos de expedición")
    print(registrosPrimero)
    #guardar data de Expediente Detalle

    numero_expediente=None
    organo_jurisdiccional=None
    juez=None
    observacion=None
    materia=None
    etapa_procesal=None
    ubicacion=None
    especialista_legal=None
    proceso=None
    estado=None
    fecha_conclusion=None
    motivo_conclusion=None
    sumilla=None
    for item in registrosPrimero:
        if item['name'] == 'Expediente N°:':
            numero_expediente = item['valor']
        elif item['name'] == 'Órgano Jurisdiccional:':
            organo_jurisdiccional = item['valor']
        elif item['name'] == 'Distrito Judicial:':
            distrito = item['valor']
        elif item['name'] == 'Juez:':
            juez = item['valor']
        elif item['name'] == 'Especialista Legal:':
            especialista = item['valor']
        elif item['name'] == 'Fecha de Inicio:':
            fecha_inicio = item['valor']
            fecha_inicio_dt = datetime.strptime(fecha_inicio.strip(), '%d/%m/%Y')
            fecha_inicio = fecha_inicio_dt.strftime('%Y-%m-%d')
        elif item['name'] == 'Proceso:':
            proceso = item['valor']
        elif item['name'] == 'Observación:':
            observacion = item['valor']
        elif item['name'] == 'Especialidad:':
            especialista_legal = item['valor']
        elif item['name'] == 'Materia(s):':
            materia = item['valor']
        elif item['name'] == 'Estado:':
            estado = item['valor']
        elif item['name'] == 'Etapa Procesal:':
            etapa_procesal = item['valor']
        elif item['name'] == 'Fecha Conclusión:':
            fecha_conclusion = item['valor']
            if fecha_conclusion==' ':
                fecha_conclusion=None
            else:
                fecha_conclusion_dt = datetime.strptime(fecha_conclusion.strip(), '%d/%m/%Y')
                fecha_conclusion = fecha_conclusion_dt.strftime('%Y-%m-%d')
        elif item['name'] == 'Ubicación:':
            ubicacion = item['valor']
        elif item['name'] == 'Motivo Conclusión:':
            motivo_conclusion = item['valor']
        elif item['name'] == 'Sumilla:':
            sumilla = item['valor']
    try:
        insert_query = '''
                    INSERT INTO EXPEDIENTE_DETALLE (numero_expediente,organo_jurisdiccional, juez, fecha_inicio, observacion,materia,etapa_procesal,ubicacion,especialista_legal,proceso,estado,fecha_conclusion,motivo_conclusion,sumilla, id_secuencia)
                    VALUES (?,?, ?, ?, ?,?,?,?, ?, ?, ?,?,?,?,?)
                        '''
        cursor.execute(insert_query, numero_expediente, organo_jurisdiccional, juez, fecha_inicio, observacion,materia,etapa_procesal,ubicacion,especialista_legal,proceso
                        ,estado,fecha_conclusion,motivo_conclusion,sumilla,id_secuencia)
        cursor.commit()
    except Exception as e :
        print(e)
    #segundo registro
    partes=[]
    partesData = driver.find_elements_by_xpath("//div[@class='partes']")
    for part in partesData:
        dataPartes=part.text.split("\n")
        partes.append(dataPartes)
    partes.pop(0)
    print("PArtes:")
    print(partes)
    #Guardar partes

    parte=" "
    tipo_persona=" "
    nombres=" "
    apellido_paterno=" "
    apellido_materno=" "
    razon_social=" "
    id_parte=0
    
    for part in partes:
        id_parte=id_parte+1
        if len(part)==5:
            parte=part[0]
            tipo_persona=part[1]
            apellido_paterno=part[2]
            apellido_materno=part[3]
            nombres=part[4]
            numero_expediente=numero_expediente
            razon_social=" "
        elif len(part)==3:
            parte=part[0]
            tipo_persona=part[1]
            razon_social=part[2]
            numero_expediente=numero_expediente
            nombres=" "
            apellido_paterno=" "
            apellido_materno=" "
        try:
            insert_query = '''
                INSERT INTO EXPEDIENTE_PARTES (parte,tipo_persona, nombres, apellido_paterno, apellido_materno,razon_social,numero_expediente)
                VALUES (?,?,?,?,?,?,?)
                    '''
            cursor.execute(insert_query, parte,tipo_persona, nombres, apellido_paterno, apellido_materno,razon_social,numero_expediente)
            cursor.commit()
        except Exception as e:
            print(e)
    #tercera parte Segumiento de Expedeinte
    cabezera = driver.find_elements_by_xpath("//*[@id='divResol']//div[contains(@class, 'roptionss')]")
    dataDeSeguimientoExpediente = driver.find_elements_by_xpath("//*[@id='divResol']/div/div[@class='row']//div[contains(@class, 'fleft')]")
    resultados=[]
    for cabecera, dataSeg in zip(cabezera, dataDeSeguimientoExpediente):
        resultados.append({"name": cabecera.text, "valor": dataSeg.text})
    print(resultados)
    tamano=int(len(resultados)/8)
    tamano_parte = len(resultados) // tamano
    partesSeguimientoExp = [resultados[i:i+tamano_parte] for i in range(0, len(resultados), tamano_parte)]
    if len(partesSeguimientoExp) > tamano:
        partesSeguimientoExp[-2] = partesSeguimientoExp[-2] + partesSeguimientoExp[-1]
        partesSeguimientoExp.pop()

    print(partesSeguimientoExp)
    #Guardar Expediente Seguimiento
    fecha=None
    resolucion=" "
    tipo_notificacion=" "
    sumilla=" "
    descripcion_usuario=" "
    acto=" "
    folio=" "
    proveido=" "
    id_seguimiento=0
    for x in partesSeguimientoExp:
        id_seguimiento=id_seguimiento+1
        for item in x:
            if item['name'] == 'Fecha de Ingreso:':
                fecha = item['valor']
                fecha_dt = datetime.strptime(fecha.strip(), '%d/%m/%Y %H:%M')
                fecha = fecha_dt.strftime('%Y-%m-%d %H:%M:%S')

            elif item['name'] == 'Resolución:':
                resolucion = item['valor']
            elif item['name'] == 'Tipo de Notificación:':
                tipo_notificacion = item['valor']
            elif item['name'] == 'Acto:':
                acto = item['valor']
            elif item['name'] == 'Folios:':
                folio = item['valor']
            elif item['name'] == 'Proveido:':
                proveido = item['valor']
                if proveido=='No Proveido':
                    proveido=None
                else:
                    proveido_dt = datetime.strptime(proveido.strip(), '%d/%m/%Y')
                    proveido = fecha_dt.strftime('%Y-%m-%d')
            elif item['name'] == 'Sumilla:':
                sumilla = item['valor']
            elif item['name'] == 'Descripción de Usuario:':
                descripcion_usuario = item['valor']
        try:
            insert_query = '''
                INSERT INTO EXPEDIENTE_SEGUIMIENTO (fecha,resolucion, tipo_notificacion, sumilla, descripcion_usuario,acto,folio,proveido,numero_expediente)
                VALUES (?,?,?,?,?,?,?,?,?)
                    '''
            cursor.execute(insert_query, fecha,resolucion, tipo_notificacion, sumilla, descripcion_usuario,acto,folio,proveido,numero_expediente)
            cursor.commit()
        except Exception as e:
            print(e)
    print("Termino")
    conn.close()
finally:
    # Cerrar el navegador
    driver.quit()
    


