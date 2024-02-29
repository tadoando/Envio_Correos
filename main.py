#%%
from  app_correo import envio_corre
import PARAMETROS

Des= PARAMETROS.DESTINATARIOS
titulo = PARAMETROS.TITULO
cuerpo = PARAMETROS.CUERPO_CORREO

envio_corre(Des, titulo, cuerpo)

print("Correo Enviado")
# %%
