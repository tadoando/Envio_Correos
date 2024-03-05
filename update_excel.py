
#%%
import win32com.client
import time
# Inicializar una instancia de Excel

def actualizar():
    excel = win32com.client.Dispatch("Excel.Application")

    # Abrir el libro de Excel
    workbook = excel.Workbooks.Open(r"C:\Users\manuel.amado\Documents\CUMPLIMIENTO_REPORTING_v.4.xlsm")

    # Ejecutar un macro en el libro de Excel
    excel.Application.Run("actulizar")
    print("Actulizando....Excel")

    time.sleep(60)

    # Guardar y cerrar el libro de Excel
    workbook.Save()
    workbook.Close()
    print("Cierre de docuemnto.")
    # Cerrar Excel
    excel.Quit()
    time.sleep(20)

