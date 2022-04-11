' /*-----------------------------[ Tabla ASCII ]---------------------------------/
'  Autor       : Gonen09
'  Descripción : Generar lista codigo ascii
'  Versión     : 2.0
'  Fecha       : 11/07/2013
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/


i=0

while (i<>256)
 
  a = msgbox ("chr("&i&") = "&chr(i),0,"Caracter " &i)
  i = i + 1

wend