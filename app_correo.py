#%%
import win32com.client
import sys
#sys.path.append(r'C:\Users\manuel.amado\Documents\Codigos_procesos') 
#from PARAMETROS import DESTINATARIOS, TITULO, CUERPO_CORREO, ARCHIVO


def enviar_correo(destinatarioS, asunto, cuerpo, adjunto):
    outlook = win32com.client.Dispatch("Outlook.Application")
    for destinatario in destinatarioS:
        try:
            correo = outlook.CreateItem(0)
            correo.To = destinatario
            correo.Subject = asunto
            correo.HTMLBody = cuerpo
            
            # Adjuntar archivo
            adjunto_path = adjunto
            correo.Attachments.Add(adjunto_path)
            
            # Enviar correo
            correo.Send()
            print("Correo enviado correctamente a", destinatario)
        except Exception as e:
            print(f"Error al enviar correo a {destinatario}: {e}")

# Ejemplo de uso
#if __name__ == "__main__":
    
#    enviar_correo(DESTINATARIOS, TITULO, CUERPO_CORREO, ARCHIVO)

