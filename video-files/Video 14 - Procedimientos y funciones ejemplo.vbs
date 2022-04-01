principal () ' Es necesario llamar a main y se puede en cualquier parte



'------ Programa principal ---------'

Sub principal ()
    
  a=5^3
  b=10   

  mensaje () ' llamada procedimiento

  sumatoria= suma(a,b) 'asignacion de funcion

  msgbox ("Suma= "&sumatoria)


End Sub 


'----------------------------------'



'---------- Procedimiento ----------'

sub mensaje ()  'no es necesario el parentesis cuando es sin parametros pero para distinguir entre variable y funcion

msgbox ("hola Gonen ")

end sub

'-----------------------------------'


'------------ Funcion --------------'

function suma (a,b)

naty=a+b

suma=naty 'asignacion del valor

end function

'-----------------------------------'


