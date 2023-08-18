import pandas as pd
import os
from selenium.webdriver.common.keys import Keys
# from selenium import webdriver
# from selenium.webdriver.chrome.options import Options
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
from seleniumbase import Driver
import time
from colorama import Fore, init

init()

driver = Driver(uc=True)
driver.get('https://dynamicoos.my.site.com/prestadores/s/login/?ec=302&startURL=%2Fprestadores%2Fs%2F')
time.sleep(5)

# Encontrar los campos de usuario y contraseña
usuario_input = driver.find_element("id", "48:2;a")  # Reemplaza "username" con el ID real del campo de usuario
contraseña_input = driver.find_element("id", "61:2;a")  # Reemplaza "password" con el ID real del campo de contraseña

# Ingresar el usuario y la contraseña
usuario_input.send_keys("confimed@coosalud.com")  # Reemplaza "tu_usuario" con tu usuario real
contraseña_input.send_keys("Danielito2023/")  # Reemplaza "tu_contraseña" con tu contraseña real

time.sleep(1)

# Encontrar y hacer clic en el botón de inicio de sesión
boton_login = driver.find_element("xpath", "//span[@class=' label bBody']")
boton_login.click()

time.sleep(10)

#------------------------------------------------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------

# # Ruta del archivo de texto
# ruta_archivo = 'CITOLOGIA.txt'

# # Nombres de las columnas
# columnas = ['Factura', 'Codigo', 'TipoDocumento', 'NoDocumento', 'Fecha', 'Col6', 'Cups', 'Col8', 'Col9', 'Col10', 'Diagnostico', 'Col12', 'Col13', 'Col14', 'Col15']

# # Leer el archivo de texto y convertirlo en un DataFrame
# df = pd.read_csv(ruta_archivo, delimiter=',', names=columnas, usecols=[2, 3, 4, 6, 10])

# numero_documento_lista = df['NoDocumento']
# fecha_proceso_lista = df['Fecha']
# codigo_cups_lista = df['Cups']
# diagnostico_lista = df['Diagnostico']

# -------------------------------------------------------------------------------------------------------------------------------

# # Ruta del archivo de texto
# ruta_archivo = 'ESPECIALISTAS.txt'

# # Nombres de las columnas
# columnas = ['Factura', 'Codigo', 'TipoDocumento', 'NoDocumento', 'Fecha', 'Col6', 'Cups', 'Col8', 'Col9', 'Diagnostico', 'Col11', 'Col12', 'Col13', 'Col14', 'Col15', 'Col16', 'Col17' ]

# # Leer el archivo de texto y convertirlo en un DataFrame
# df = pd.read_csv(ruta_archivo, delimiter=',', names=columnas, usecols=[2, 3, 4, 6, 9])

# numero_documento_lista = df['NoDocumento']
# fecha_proceso_lista = df['Fecha']
# codigo_cups_lista = df['Cups']
# diagnostico_lista = df['Diagnostico']

# cantidad = len(numero_documento_lista)
# cantidad = str(cantidad)

# ---------------------------------------------------------------------------

# Ruta del archivo de texto
ruta_archivo = 'POTENCIALES.txt'

# Nombres de las columnas
columnas = ['Factura', 'Codigo', 'TipoDocumento', 'NoDocumento', 'Fecha', 'Col6', 'Cups', 'Col8', 'Col9',  'Col10', 'Diagnostico', 'Col12', 'Col13', 'Col14', 'Col15']

# Leer el archivo de texto y convertirlo en un DataFrame
df = pd.read_csv(ruta_archivo, delimiter=',', names=columnas, usecols=[2, 3, 4, 6, 10])

numero_documento_lista = df['NoDocumento']
fecha_proceso_lista = df['Fecha']
codigo_cups_lista = df['Cups']
diagnostico_lista = df['Diagnostico']

cantidad = len(numero_documento_lista)
cantidad = str(cantidad)

contador = 0
numeros_agendas = []

for NumeroDocumento, Fecha, CodigoCups, diagnostico_v in zip(numero_documento_lista, fecha_proceso_lista, codigo_cups_lista, diagnostico_lista):
    try:
        contador += 1
        print(Fore.CYAN + "\n[!] Paciente Numero: " + str(contador), end='\r')
        # Empezar a subir los datos a la pagina
        documento = str(NumeroDocumento)
        #Buscando la etiqueta para ingresar el numero de cedula
        searc_input = driver.find_element("id", "178:0")

        #borrar el contenido que hay en buscar
        searc_input.clear()
        time.sleep(1)

        #ingresar numero de cedula
        searc_input.send_keys(NumeroDocumento)
        time.sleep(1)
        searc_input.send_keys(Keys.ENTER)
        time.sleep(12)


        #selecccionar paciente
        person = driver.find_element("xpath", "//a[@data-aura-class='forceOutputLookup']")
        time.sleep(4)
        person.click()
        time.sleep(6)

        #-----------------------------------------------------------------------------------------------------
        #-----------------------------------------------------------------------------------------------------

        #verificar si el usuario esta activo
        estado = driver.find_elements("xpath", "//div[@class='slds-form-element__static slds-truncate']")

        if len(estado) > 4:
            afiliacion = estado[2].text

        if afiliacion == "Activo":

            #seleccionar solicitudes ambulatorias
            request = driver.find_elements("xpath", "//span[@class='title']")

            if len(request) > 1:
                solicitud = request[1]
                solicitud.click()
                time.sleep(4)
            else:
                print('No se encuentra el xpath')

            #Ingresar numero telefonico
            phone = driver.find_element("xpath", "//input[@class='slds-input']")
            phone.send_keys(NumeroDocumento)
            time.sleep(1)
            button = driver.find_element("xpath", "//button[@title='Siguiente']")
            button.click()
            time.sleep(6)

            codigo_cups = f'{CodigoCups}'

            if codigo_cups == '898001':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    

                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys("Z000")
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("ESTUDIO DE COLORACIÓN BÁSICA EN CITOLOGÍA VAGINAL TUMORAL O FUNCIONAL")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")

                codigo = cups[-1]
                codigo.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)

                

            #Proceso para primera vez medicina interna
            elif codigo_cups == '890266':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    

                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("CONSULTA PRIMERA VEZ MEDICINA INTERNA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")


                codigo = cups[-15]
                codigo.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)

            #Proceso para control medicina general
            elif codigo_cups == '890366':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    

                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("CONSULTA DE CONTROL MEDICINA INTERNA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                #elegir el servicio correcto
                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")
                last_cups = cups[-1]
                last_cups.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)

            #Proceso para primera vez ginecologia
            elif codigo_cups == '890250':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    
                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("CONSULTA DE PRIMERA VEZ POR ESPECIALISTA EN GINECOLOGÍA Y OBSTETRICIA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                #elegir el servicio correcto
                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")
                last_cups = cups[-12]
                last_cups.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)

            #Proceso para control ginecologia
            elif codigo_cups == '890350':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    
                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("CONSULTA DE CONTROL O DE SEGUIMIENTO POR ESPECIALISTA EN GINECOLOGÍA Y OBSTETRICIA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                #elegir el servicio correcto
                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")
                last_cups = cups[-5]
                last_cups.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)

            #Proceso para primera vez nutricion
            elif codigo_cups == '890206':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    
                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("CONSULTA DE CONTROL O DE SEGUIMIENTO POR ESPECIALISTA EN GINECOLOGÍA Y OBSTETRICIA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                #elegir el servicio correcto
                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")
                last_cups = cups[-4]
                last_cups.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)

            #Proceso para control nutricion
            elif codigo_cups == '890306':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    
                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("CONSULTA DE CONTROL O DE SEGUIMIENTO POR ESPECIALISTA EN GINECOLOGÍA Y OBSTETRICIA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                #elegir el servicio correcto
                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")
                last_cups = cups[-1]
                last_cups.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)
            #Proceso para primera vez pediatria
            elif codigo_cups == '890283':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    
                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("CONSULTA DE CONTROL O DE SEGUIMIENTO POR ESPECIALISTA EN GINECOLOGÍA Y OBSTETRICIA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                #elegir el servicio correcto
                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")
                last_cups = cups[-3]
                last_cups.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)
            #Proceso para control pediatria
            elif codigo_cups == '890383':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    
                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("CONSULTA DE CONTROL O DE SEGUIMIENTO POR ESPECIALISTA EN GINECOLOGÍA Y OBSTETRICIA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                #elegir el servicio correcto
                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")
                last_cups = cups[-1]
                last_cups.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)

                #Dar click en el boton direccionar
                direccionar = driver.find_element("xpath", "//button[@title='Direccionar']")
                direccionar.click()
                time.sleep(6)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)

                #borrar el contenido que hay en buscar
                searc_input.clear()
                time.sleep(1)

                #Ingresar el numero de agenda
                searc_input.send_keys(numero_agenda)
                time.sleep(3)

                #Seleccionar la agenda correcta
                agenda2 = driver.find_element("xpath", "//div[@class='slds-truncate uiOutputRichText']")
                agenda2.click()
                time.sleep(5)

                #Dar click en el boton asistida
                direccionar = driver.find_element("xpath", "//button[@class='slds-button slds-button_success']")
                direccionar.click()
                print(Fore.GREEN + "Cargado exitosamente el paciente: " + documento, end='\r')
                time.sleep(8)

            elif codigo_cups == '954629':
                #Llenar solicitud
                ips = driver.find_element("xpath", "//input[@placeholder='Selecciona una ips']")
                ips.send_keys("CONFIMED")
                time.sleep(3)

                #Seleccionar confimed ips
                confimed = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(confimed) > 13:
                    sede = confimed[14]
                    sede.click()
                    time.sleep(3)
                    
                #seleccionar diagnostico
                dx = driver.find_element("xpath", "//input[@placeholder='Buscar Diagnósticos de cuidados sanitarios...']")
                dx.send_keys(f'{diagnostico_v}')
                time.sleep(3)

                diagnostico = driver.find_elements("xpath", "//div[@class='slds-media__body']")

                if len(diagnostico) > 18:
                    dx = diagnostico[19]
                    dx.click()

                #seleccionar ubicacion
                ubicacion = driver.find_element("xpath", "//input[@placeholder='Selecciona una ubicacion']")
                ubicacion.send_keys("FLORIDABLANCA")
                time.sleep(3)

                place = driver.find_elements("xpath", "//span[@class='slds-lookup__item-action slds-media']")

                if len(diagnostico) > 2:
                    city = place[3]
                    city.click()
                    time.sleep(3)

                #Ingresar la fecha de la orden
                day = driver.find_element("xpath", "//input[@class='slds-input']")
                day.send_keys(Fecha)
                time.sleep(1)

                #Ingresar descripcion del servicio
                descripcion = driver.find_element("xpath", "//textarea[@class='slds-textarea']")
                descripcion.send_keys("POTENCIALES EVOCADOS AUDITIVOS DE CORTA LATENCIA CON CURVA FUNCIÓN INTENSIDAD-LATENCIA")
                time.sleep(1)

                #Ingresar el servicio
                servicio = driver.find_element("xpath", "//input[@placeholder='Selecciona un servicio']")
                servicio.send_keys(codigo_cups)
                time.sleep(4)

                #elegir el servicio correcto
                cups = driver.find_elements("xpath", "//div[@class='slds-lookup__result-text']")
                last_cups = cups[-1]
                last_cups.click()
                time.sleep(5)

                agregar = driver.find_element("xpath", "//button[@class='slds-button slds-button_brand']")
                agregar.click()
                time.sleep(5)

                #Dar click en el boton siguiente
                siguiente = driver.find_elements("xpath", "//button[@class='slds-button slds-button_brand']")

                if len(siguiente) > 1:
                    next = siguiente[1]
                    next.click()
                    time.sleep(5)

                #Dar click en el boton verde o bolita
                no_modelo = driver.find_element("xpath", "//span[@class='slds-radio_faux']")
                no_modelo.click()
                time.sleep(2)

                #dar click en el boton validar
                validar = driver.find_element("xpath", "//button[@title='Validar']")
                validar.click()
                time.sleep(5)
                print(Fore.GREEN + '\r' + "Cargado exitosamente el paciente: " + documento, end='')
                time.sleep(5)

                #Copiar el numero de agenda
                agenda = driver.find_element("xpath", "//span[@class='slds-pill__label sl-pill__label']")
                numero_agenda = agenda.text

                #agregar numero de agenda a la lista
                numeros_agendas.append(numero_agenda)
            
        else:
            print(Fore.YELLOW + f'[!] El paciente {documento} esta {afiliacion}')
    except:
        print(Fore.RED + f'\n[-] El paciente {NumeroDocumento} no fue agregado \n')



# Cierra el navegador
driver.quit()

df['Agendas'] = numeros_agendas

df.to_excel('agendas.xlsx', index=False)




