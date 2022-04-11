' /*-------------------------------[ Modulos ]-----------------------------------/
'  Autor       : Gonen09
'  Descripción : Definición y operación con procedimientos y funciones
'  Versión     : 1.0
'  Fecha       : 07/09/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

'----------- Sintaxis -----------'

'Procedimiento
'-------------


'sub nombre (parametro1,parametro2,...,parametroN)

 'Sentencias

'end sub


'Funcion
'-------

'function nombre (parametro1,parametro2,...,parametroN)

  'sentencias
  'nombre = algunValor ' aqui se asigna el valor de la funcion

'end function



'Ejemplo
'--------


principal () ' Es necesario llamar a main y se puede en cualquier parte


'--------- Funcion suma ------------'

function suma (num1,num2)

  resultado = num1+num2

  suma = resultado

end function


'------- Procedimiento suma --------'

sub sumas (num1,num2)
 
  resultado = num1+num2
  msgbox ("Respuesta = "&resultado)

end sub

'------ Programa principal ---------'

Sub principal ()
    
  a=5
  b=10   
  
  call sumas (a,b)
  sumatoria = suma(a,b) 'asignacion de funcion

  msgbox ("Suma= "&sumatoria)


End Sub 


'----------------------------------'
