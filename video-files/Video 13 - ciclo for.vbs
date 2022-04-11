' /*------------------------------[ Ciclo For ]----------------------------------/
'  Autor       : Gonen09
'  Descripción : Definición y operación con ciclo For
'  Versión     : 1.0
'  Fecha       : 10/03/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

'CICLO FOR


'for indice = inicio to final
 
  'CODIGO A EJECUTAR

'next



'EJEMPLO


for indice = 0 to 5

   msgbox indice

next




'EJEMPLO CON ARREGLOS

dim nombres() 	'DECLARAMOS UN ARREGLO DINAMICO

cantidad = inputbox ("Ingrese cantidad de contactos para guardar: ") 'PARA ALMACENAR CANTIDAD DE CONTACTOS

redim nombres(cantidad)  'ASIGNAMOS LA DIMENSION AL ARREGO (REDIMENSIONAR)


for i = 0 to (cantidad - 1)	'CICLO FOR PARA INGRESAR NOMBRES

  nombres(i) = inputbox("Ingrese nombre en posicion: "&i)

next


for i = 0 to (cantidad - 1)     'CICLO FOR PARA MOSTRAR NOMBRES

  msgbox "Posicion: "&i&" --> Nombre: "&nombres(i) 

next

  
