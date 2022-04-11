' /*---------------------------[ Modulos ejemplo ]-------------------------------/
'  Autor       : Gonen09
'  Descripción : Ejemplo de procedimientos y funciones
'  Versión     : 2.0
'  Fecha       : 07/09/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

principal () ' Es necesario llamar a main y se puede en cualquier parte

'------ Programa principal ---------'

Sub principal ()
    
  a = 5^3
  b = 10   

  mensaje () ' llamada procedimiento

  sumatoria = suma(a,b) 'asignacion de funcion

  msgbox ("Suma = "&sumatoria)


End Sub 


'----------------------------------'



'---------- Procedimiento ----------'

sub mensaje ()  'no es necesario el parentesis cuando es sin parametros pero para distinguir entre variable y funcion

  msgbox ("hola Gonen ")

end sub

'-----------------------------------'


'------------ Funcion --------------'

function suma (a,b)

  add = a+b

  suma = add 'asignacion del valor

end function

'-----------------------------------'


