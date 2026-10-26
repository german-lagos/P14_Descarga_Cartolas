import os
import win32com.client as win32

from pathlib import Path
from dotenv import load_dotenv


# ==========================================
# CONFIGURACIÓN
# ==========================================

load_dotenv()

EMAIL_CUENTA = os.getenv("EMAIL_CUENTA")
EMAIL_REMITENTE = os.getenv("EMAIL_REMITENTE")
EMAIL_ASUNTO = os.getenv("EMAIL_ASUNTO")

DOWNLOAD_PATH = Path(
    os.getenv("DOWNLOAD_PATH", "descargas")
)

DOWNLOAD_PATH.mkdir(
    parents=True,
    exist_ok=True
)

def validar_configuracion():

    variables = {
        "EMAIL_CUENTA": EMAIL_CUENTA,
        "EMAIL_REMITENTE": EMAIL_REMITENTE,
        "EMAIL_ASUNTO": EMAIL_ASUNTO
    }

    for nombre, valor in variables.items():

        if not valor:

            raise ValueError(
                f"La variable {nombre} no está "
                f"configurada en el archivo .env"
            )

    print("[0] Configuración validada correctamente.")

# ==========================================
# DESCARGAR ARCHIVO DAT
# ==========================================

def descargar_dat():

    print("=" * 55)
    print(" P14 - PRUEBA DESCARGA DAT DESDE OUTLOOK")
    print("=" * 55)

    try:

        # Conectar con Outlook
        outlook = win32.Dispatch(
            "Outlook.Application"
        )

        namespace = outlook.GetNamespace("MAPI")

        print("\n[1] Conexión con Outlook establecida.")

        # Buscar cuenta configurada
        cuenta_encontrada = False

        for cuenta in namespace.Accounts:

            if cuenta.SmtpAddress.lower() == EMAIL_CUENTA.lower():

                cuenta_encontrada = True
                break

        if not cuenta_encontrada:

            raise Exception(
                "No se encontró la cuenta configurada."
            )

        print(f"[2] Cuenta identificada: {EMAIL_CUENTA}")

        # Obtener bandeja de entrada de la cuenta
        bandeja = cuenta.DeliveryStore.GetDefaultFolder(6)

        print("[3] Accediendo a bandeja de entrada...")

        mensajes = bandeja.Items

        # Ordenar mensajes por fecha descendente
        mensajes.Sort("[ReceivedTime]", True)

        print("[4] Buscando correo con archivo DAT...")

        for mensaje in mensajes:

            # Verificar que sea un correo
            if mensaje.Class != 43:
                continue

            # Obtener información
            remitente = mensaje.SenderEmailAddress
            asunto = mensaje.Subject

            # Validar remitente
            if remitente.lower() != EMAIL_REMITENTE.lower():
                continue

            # Validar asunto
            if EMAIL_ASUNTO.lower() not in asunto.lower():
                continue

            print("\nCorreo encontrado:")
            print(f"Remitente: {remitente}")
            print(f"Asunto: {asunto}")

            # Buscar adjuntos
            for adjunto in mensaje.Attachments:

                nombre = adjunto.FileName

                if nombre.lower().endswith(".dat"):

                    ruta = DOWNLOAD_PATH / nombre

                    adjunto.SaveAsFile(
                        str(ruta.resolve())
                    )

                    print("\nArchivo descargado correctamente:")
                    print(ruta.resolve())

                    return ruta

        print("\nNo se encontró un correo con archivo DAT.")

        return None

    except Exception as error:

        print("\nERROR:")
        print(error)

        return None


# ==========================================
# EJECUCIÓN
# ==========================================

if __name__ == "__main__":

    validar_configuracion()
    descargar_dat()