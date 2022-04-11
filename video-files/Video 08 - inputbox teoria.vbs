' /*-------------------------------[ Inputbox ]----------------------------------/
'  Autor       : Gonen09
'  Descripción : Definición, teoria y operación de Inputbox
'  Versión     : 1.0
'  Fecha       : 22/05/2013
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

'variable = inputbox ("Mensaje","Titulo","Respuesta por defecto",corx,cory) 'version completa

'El tamaño de "Mensaje" puede llegar hasta 1024 caracteres
'corx= Posicion horizontal del mensaje, ingresar valores numericos enteros
'cory= Posicion vertical del mensaje, ingresar valores numericos enteros

'La posicion del mensaje va desde (0,0) [Esquina superior izquiera] hasta la resolucion de tu pantalla [Esquina inferior derecha]

'El valor maximo para la resolucion 1024 x 768 es (9800,8500)

'El valor maximo para la resolucion 1600 x 900 es (18500,10500)


'Otras formas del inputbox 
'--------------------------


'variable = inputbox ("Mensaje","Titulo","Respuesta por defecto") 'version sin coordenadas, aparece al centro de la pantalla

'variable = inputbox ("Mensaje","Titulo") 'version que no incluye la respuesta por defecto

'variable = inputbox ("Mensaje") 'version que no incluye la respuesta por defecto ni el titulo


'Otras formas del inputbox incluyendo la constante empty (Tambien puede ser "")
'-------------------------------------------------------


'variable = inputbox (empty,"Titulo","Respuesta por defecto") 'Version sin mensaje'

'variable = inputbox ("Mensaje",empty,"Respuesta por defecto") 'Version sin titulo

'variable = inputbox ("Mensaje","Titulo",empty) 'Version sin respuesta por defecto