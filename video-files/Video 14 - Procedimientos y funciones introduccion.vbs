
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
  sumatoria= suma(a,b) 'asignacion de funcion

  msgbox ("Suma= "&sumatoria)


End Sub 


'----------------------------------'
