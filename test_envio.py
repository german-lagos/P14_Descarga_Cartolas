import os
import win32com.client as win32

from pathlib import Path
from dotenv import load_dotenv


# ==========================================
# 1. CONFIGURACIÓN
# ==========================================

BASE_DIR = Path(__file__).resolve().parent

load_dotenv(BASE_DIR / ".env")

EMAIL_CUENTA = os.getenv("EMAIL_CUENTA")
EMAIL_DESTINATARIO = os.getenv("EMAIL_DESTINATARIO")

EMAIL_ASUNTO_ENVIO = os.getenv(
    "EMAIL_ASUNTO_ENVIO",
    "PRUEBA P14 - CARTOLA PROCESADA"
)

EMAIL_MENSAJE = os.getenv(
    "EMAIL_MENSAJE",
    "Se adjunta archivo TXT procesado."
)

OUTPUT_PATH = BASE_DIR / os.getenv(
    "OUTPUT_PATH",
    "procesados"
)

EMPRESA = os.getenv("EMPRESA")
FECHA_CARTOLA = os.getenv("FECHA_CARTOLA")

MODO_PRUEBA = (
    os.getenv("MODO_PRUEBA", "true").lower() == "true"
)


# ==========================================
# 2. VALIDAR CONFIGURACIÓN
# ==========================================

def validar_configuracion():

    print("\n[1] Validando configuración...")

    if not EMAIL_CUENTA:
        raise ValueError(
            "EMAIL_CUENTA no está configurada."
        )

    if not EMAIL_DESTINATARIO:
        raise ValueError(
            "EMAIL_DESTINATARIO no está configurado."
        )

    if not EMPRESA or not FECHA_CARTOLA:
        raise ValueError(
            "Faltan EMPRESA o FECHA_CARTOLA."
        )

    print("Configuración validada correctamente.")


# ==========================================
# 3. IDENTIFICAR ARCHIVO TXT
# ==========================================

def obtener_archivo_txt():

    print("\n[2] Buscando archivo TXT procesado...")

    nombre_empresa = "".join(
        caracter if caracter.isalnum() else "_"
        for caracter in EMPRESA.upper()
    )

    fecha_nombre = FECHA_CARTOLA.replace("-", "")

    nombre_archivo = (
        f"CARTOLA_{nombre_empresa}_{fecha_nombre}.txt"
    )

    ruta_archivo = OUTPUT_PATH / nombre_archivo

    if not ruta_archivo.is_file():
        raise FileNotFoundError(
            f"No se encontró el archivo: {ruta_archivo}"
        )

    if ruta_archivo.stat().st_size == 0:
        raise ValueError(
            "El archivo TXT está vacío."
        )

    print(f"Archivo identificado: {ruta_archivo.name}")

    return ruta_archivo


# ==========================================
# 4. CONECTAR CON OUTLOOK
# ==========================================

def conectar_outlook():

    print("\n[3] Conectando con Outlook...")

    outlook = win32.Dispatch(
        "Outlook.Application"
    )

    namespace = outlook.GetNamespace("MAPI")

    for cuenta in namespace.Accounts:

        direccion = cuenta.SmtpAddress

        if (
            direccion
            and direccion.lower() == EMAIL_CUENTA.lower()
        ):

            print(f"Cuenta identificada: {direccion}")

            return outlook, cuenta

    raise RuntimeError(
        f"No se encontró la cuenta {EMAIL_CUENTA}"
    )


# ==========================================
# 5. PREPARAR Y ENVIAR CORREO
# ==========================================

def enviar_correo(outlook, cuenta, ruta_archivo):

    print("\n[4] Preparando correo electrónico...")

    mail = outlook.CreateItem(0)

    # Seleccionar explícitamente la cuenta remitente
    mail.SendUsingAccount = cuenta

    mail.To = EMAIL_DESTINATARIO
    mail.Subject = EMAIL_ASUNTO_ENVIO

    mail.Body = (
        f"{EMAIL_MENSAJE}\n\n"
        f"Empresa: {EMPRESA}\n"
        f"Fecha de cartola: {FECHA_CARTOLA}\n\n"
        "Mensaje generado durante las pruebas del proceso P14."
    )

    mail.Attachments.Add(
        str(ruta_archivo.resolve())
    )

    print(f"Destinatario: {EMAIL_DESTINATARIO}")
    print(f"Asunto: {EMAIL_ASUNTO_ENVIO}")
    print(f"Adjunto: {ruta_archivo.name}")

    if MODO_PRUEBA:

        mail.Save()

        print("\nMODO PRUEBA ACTIVADO")
        print("El mensaje fue guardado como borrador.")
        print("No se envió ningún correo.")

    else:

        mail.Send()

        print("\nOutlook aceptó la solicitud de envío.")
        print("Verifica el resultado en Elementos enviados.")


# ==========================================
# 6. EJECUCIÓN PRINCIPAL
# ==========================================

def main():

    print("=" * 55)
    print(" P14 - PRUEBA ENVÍO TXT MEDIANTE OUTLOOK")
    print("=" * 55)

    try:

        validar_configuracion()

        ruta_txt = obtener_archivo_txt()

        outlook, cuenta = conectar_outlook()

        enviar_correo(
            outlook,
            cuenta,
            ruta_txt
        )

        print("\nProceso finalizado correctamente.")

    except Exception as error:

        print("\nERROR EN EL PROCESO:")
        print(f"{type(error).__name__}: {error}")


if __name__ == "__main__":
    main()