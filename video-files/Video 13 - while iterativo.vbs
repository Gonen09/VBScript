' /*--------------------------[ While Iterativo ]--------------------------------/
'  Autor       : Gonen09
'  Descripción : Definición y operación con ciclo While de forma iterativa
'  Versión     : 1.0
'  Fecha       : 10/03/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

'WHILE ITERATIVO

'tamano = 3
'indice = 0

'while (indice < tamano)

   'CODIGO A EJECUTAR

   'indice = indice + 1

'wend



'EJEMPLO

indice = 0

while (indice <= 5)

   msgbox indice
   indice = indice + 1	
   
wend



'EJEMPLO CON ARREGLOS

dim nombres(3)
i = 0   	'PARA NO USAR INDICE ABREVIAMOS CON LA I

while (i < 3)

  nombres(i) = inputbox("Ingrese nombre en posicion: "&i)
  i = i+1  
 
wend

i = 0   'VOLVEMOS EL CONTADOR A CERO, PARA NO CREAR OTRA VARIABLE CONTADOR

while (i < 3)

  msgbox "Posicion: "&i&" --> Nombre: "&nombres(i) 
  i = i+1
 
wend
