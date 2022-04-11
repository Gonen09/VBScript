' /*----------------------------[ Inputbox Suma ]--------------------------------/
'  Autor       : Gonen09
'  Descripción : Sumar variables con inputbox
'  Versión     : 1.0
'  Fecha       : 14/02/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/
 
a = inputbox ("Ingresa primer sumando: ","TEST")  

a = cint(a)  'Convertidor de string o char a numero entero

b = inputbox ("Ingresa segundo sumando: ","TEST")  

b = cint(b)

suma = a+b

msgbox "resultado:  "&suma