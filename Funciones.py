import math
import time
import locale
import ssl
import smtplib
from email.message import EmailMessage


def cambiar_contenido(cuerpo, asunto, nombreEmpleado, salarioFinal, incrementoMayo):
    asunto = asunto.replace("[Nombre del Empleado]", nombreEmpleado)
    cuerpo = cuerpo.replace("[Nombre del Empleado]", nombreEmpleado)
    cuerpo = cuerpo.replace("[Nuevo Salario]", salarioFinal)
    cuerpo = cuerpo.replace("[Incremento de Mayo]", incrementoMayo)

    return asunto, cuerpo


def extraer_valores(index, archivoDatos, contenido):
    cuerpo = contenido
    asunto = "DDC - Ajuste Salarial Mayo 2023 - [Nombre del Empleado]"
    print("\n")

    # La funcion iloc es de la biblioteca pandas y funciona para obtener un valor
    # de una "casilla" de un DataFrame especificandole fila y columna
    nombreEmpleado = archivoDatos.iloc[index]['Nombre de empleado']
    print(nombreEmpleado)

    salarioFinal = archivoDatos.iloc[index]['Salario Final']
    locale.setlocale(locale.LC_ALL, 'de')
    salarioFinal = "$" + locale.format_string("%.2f", salarioFinal, True) + " "
    print(salarioFinal)

    incrementoMayo = archivoDatos.iloc[index]['Incremento de Mayo']
    incrementoMayo = str(int(incrementoMayo * 100)) + "%"
    print(incrementoMayo)

    asunto, cuerpo = cambiar_contenido(cuerpo, asunto, nombreEmpleado, salarioFinal, incrementoMayo)
    
    nombreSupervisor = archivoDatos.iloc[index]['Supervisor Name']

    nombreCounselor = archivoDatos.iloc[index]['Counselor']

    emailDDC = ''

    remitentes = [emailDDC]

    return asunto, cuerpo, nombreEmpleado, nombreSupervisor, nombreCounselor, remitentes


def chequeo_error(archivoCorreos, remitentes, index, nombreEmpleado, nombre, string):
    filtro = archivoCorreos['Name'] == nombre

    try:
        email = archivoCorreos.loc[filtro, 'Correo'].iloc[0]
    except IndexError:
        index_error(index, nombreEmpleado, nombre, string)
    else:
        remitentes.append(email)


def index_error(index, nombreEmpleado, nombre, string):

    match string:
        case 'Empleado':
            with open('Lista_Errores.txt', 'a', encoding='utf-8') as archivo:
                if isinstance(nombre, float) and math.isnan(nombre):
                    archivo.write('El empleado ' + str(index) + ' no posee nombre en la nomina de aumentos\n')
                    print('El empleado ' + str(index) + ' no posee nombre en la nomina de aumentos\n')
                else:
                    archivo.write('El Empleado ' + nombreEmpleado + ' no se encontró en el listado de correos\n')
                    print('El Empleado ' + nombreEmpleado + ' no se encontró en el listado de correos\n')
        
        case 'Supervisor':
            with open('Lista_Errores.txt', 'a', encoding='utf-8') as archivo:
                if isinstance(nombre, float) and math.isnan(nombre):
                    archivo.write(
                        'El empleado ' + str(index) + ' no posee el nombre de su Supervisor en la nomina de aumentos\n')
                    print('El empleado ' + str(index) + ' no posee el nombre de su Supervisor en la nomina de aumentos\n')
                else:
                    archivo.write(
                        'Para empleado ' + nombreEmpleado + ' no se encontró su Supervisor de nombre ' + nombre + ' en el listado de correos\n')
                    print(
                        'Para empleado ' + nombreEmpleado + ' no se encontró su Supervisor de nombre ' + nombre + ' en el listado de correos\n')
        
        case "Counselor":
            with open('Lista_Errores.txt', 'a', encoding='utf-8') as archivo:
                if isinstance(nombre, float) and math.isnan(nombre):
                    archivo.write(
                        'El empleado ' + str(index) + ' no posee el nombre de su Counselor en la nomina de aumentos \n')
                    print('El empleado ' + str(index) + ' no posee el nombre de su Counselor en la nomina de aumentos \n')
                else:
                    archivo.write(
                        'Para empleado ' + nombreEmpleado + ' no se encontró su Counselor de nombre ' + nombre + ' en el listado de correos\n')
                    print(
                        'Para empleado ' + nombreEmpleado + ' no se encontró su Counselor de nombre ' + nombre + ' en el listado de correos\n')
            
            
def promocion(index, archivo, cuerpo):
    if archivo.iloc[index]['Promocion '] == 'P':
        print('P')

        nuevaPosicion = archivo.iloc[index]['Position Final']
        print(nuevaPosicion)

        incrementoAcumulado = archivo.iloc[index]['Incremento Acumulado']
        incrementoAcumulado = str(int(incrementoAcumulado * 100)) + "%"
        print(incrementoAcumulado)

        textoPromocion = "También queremos felicitarte por la promoción a [Nueva Posición], que significó un 15%. Tu acumulado anual es de [Incremento Acumulado]."
        textoPromocion = textoPromocion.replace("[Nueva Posición]", nuevaPosicion)
        textoPromocion = textoPromocion.replace("[Incremento Acumulado]", incrementoAcumulado)

        cuerpo = cuerpo.replace(
            "[También queremos felicitarte por la promoción a [Nueva Posición], que significó un 15%. Tu acumulado anual es de [Incremento Acumulado].]",
            textoPromocion)
        print(cuerpo)
    else:
        cuerpo = cuerpo.replace(
            "[También queremos felicitarte por la promoción a [Nueva Posición], que significó un 15%. Tu acumulado anual es de [Incremento Acumulado].]",
            "")
        cuerpo = cuerpo.replace("\n\n", "\n")
        print(cuerpo)
    return cuerpo


def enviar_mail(asunto, cuerpo, remitentes):
    # Defino el mail y contraseña del mail utilizado para ENVIAR los correos
    email_emisor = 'luchoperez004@gmail.com'
    email_contrasena = ''

    # Crea un objeto de tipo mensaje de correo de la clase EmailMessage de la biblioteca email.message
    em = EmailMessage()
    em['From'] = email_emisor
    em['Subject'] = asunto
    em.set_content(cuerpo)

    # Crea un objeto de tipo contexto SSL (protocolo) el cual contiene infomacion sobre config de seguridad SSL/TLS
    # para establecer una conexion segura con un servidor remoto
    contexto = ssl.create_default_context()

    # Establece una conexion segura con un servidor SMTP "Simple Mail Transfer Protocol" y los parámetros
    # son la dirección del servidor, el puerto del servidor y el objeto SSL
    """
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=contexto) as smtp:
        smtp.login(email_emisor, email_contrasena)
        for remit in remitentes:
            print(remit)
            # em['To'] = 'luchoperez004@gmail.com' NO ES NECESARIO YA QUE EL RECEPTOR SE PONE EN EL SMTP.SENDMAIL
            smtp.sendmail(email_emisor, remit, em.as_string())
            time.sleep(3)
    """
