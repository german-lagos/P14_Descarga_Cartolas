import win32com.client


def probar_conexion_outlook():

    print("=" * 55)
    print(" PRUEBA DE CONEXIÓN CON MICROSOFT OUTLOOK")
    print("=" * 55)

    try:

        outlook = win32com.client.Dispatch(
            "Outlook.Application"
        )

        namespace = outlook.GetNamespace("MAPI")

        print("\nConexión establecida correctamente.")

        print("\nCuentas configuradas:")

        for cuenta in namespace.Accounts:

            print(f"- {cuenta.SmtpAddress}")

        print("\nCarpetas disponibles:")

        for carpeta in namespace.Folders:

            print(f"- {carpeta.Name}")

        print("\nPrueba finalizada correctamente.")

    except Exception as error:

        print("\nERROR AL CONECTAR CON OUTLOOK:")
        print(error)


if __name__ == "__main__":

    probar_conexion_outlook()