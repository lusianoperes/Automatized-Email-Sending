from Funciones import *
# Necesario: pip install pandas y pip install openpyxl
import pandas as pd


# Restablezco la lista de errores como vacía al inicio de la ejecución
with open('Lista_Errores.txt', 'w+', encoding='utf-8') as archivo:
    archivo.write('')


# Guarda el contenido de los excels en variables tipo DataFrames
# Un DataFrame es una estructura de datos bidimensional (tabla) de "pandas"
listadoCorreos = pd.read_excel('listadoCorreos.xlsx')
nominaAumento = pd.read_excel('Nomina_Aumento.xlsx')

# El comando with cierra el archivo luego de realizar la iteracion
with open('Mail_Incremento.txt', 'r', encoding='utf-8') as archivo:
    contenido = archivo.read()


# Itera por cada uno de los registros del excel
for i in range(0, len(nominaAumento)):
    # Extrae de la nomina de aumento el nombre del Empleado Supervisor y Counselor
    # Con otros valores de la nomina , genera el mensaje a enviar para cada Empleado, Supervisor y Counselor
    asunto, cuerpo, nombreEmpleado, nombreSupervisor, nombreCounselor, remitentes = extraer_valores(i, nominaAumento, contenido)

    # A partir de los nombres, busca sus emails en la lista de Correos y los agrega a Remitentes
    # En caso de que no se encuentre algún valor, se genera/modifica un archivo .txt donde se indica que falta
    chequeo_error(listadoCorreos, remitentes, (i+1), nombreEmpleado, nombreEmpleado, 'Empleado')
    chequeo_error(listadoCorreos, remitentes, (i+1), nombreEmpleado, nombreSupervisor, 'Supervisor')
    chequeo_error(listadoCorreos, remitentes, (i+1), nombreEmpleado, nombreCounselor, 'Counselor')

    # Analiza si el empleado está en situación de promoción, si lo está añade un contenido especial al correo
    cuerpo = promocion(i, nominaAumento, cuerpo)

    # Con la lista de remitentes y sus emails, establece conexión con un servidor SMTP para enviar los correos.
    email = enviar_mail(asunto, cuerpo, remitentes)

    print('-' * 100)
