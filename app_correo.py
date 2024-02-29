#%%
import win32com.client
#%%
outlook = win32com.client.Dispatch("Outlook.Application")
#%%
def envio_corre(Destinatarios, titulo, cuerpo):
    mail = outlook.CreateItem(0)
    mail.To = Destinatarios
    mail.Subject = titulo
    mail.Body = cuerpo
    mail.Send()