' /*----------------------------[ Calculadora ]----------------------------------/
'  Autor       : Gonen09
'  Descripción : Calculadora simple utilizando procedimientos y funciones
'  Versión     : 1.0
'  Fecha       : 07/09/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

principal () ' Es necesario llamar a main y se puede en cualquier parte

'------- Procedimiento suma --------'

sub suma ( )
 
  num1 = valor("ingrese primer sumando")
  num2 = valor("ingrese segundo sumando")
  resultado = num1+num2
  msgbox ("Respuesta = "&resultado)

end sub

'------------------------------------'



'------- Procedimiento resta --------'

sub resta ( )
 
  num1 = valor("ingrese minuendo")
  num2 = valor("ingrese sustraendo") 
  resultado = num1-num2
  msgbox ("Respuesta = "&resultado)

end sub

'------------------------------------'



'--- Procedimiento multiplicacion ---'

sub multi ( )
 
  num1 = valor("ingrese primer factor")
  num2 = valor("ingrese segundo factor") 
  resultado = num1*num2
  msgbox ("Respuesta = "&resultado)

end sub

'------------------------------------'


'------ Procedimiento division ------'

sub div ( )

  num1 = valor("ingrese dividendo")
  num2 = valor("ingrese divisor") 
 
  if (num2 <> 0) then
  
  	resultado = num1/num2
  	msgbox ("Respuesta = "&resultado)
  else
   
        msgbox ("No se puede dividir por cero") 

  end if  

end sub

'------------------------------------'


'----- Funcion obtener valores ------'

function valor (mensaje)

  numero = inputbox (mensaje,"Captando valor")

  if (isnumeric(numero) = true) then

	  valor = Cint(numero)   

  else
        msgbox ("ingrese solo números")
		valor = 0	
  end if

end function

'-----------------------------------'



'------ Programa principal ---------'

Sub principal ()
    
   menu = "Calculadora"&chr(10)&"------------------"&chr(10)&"1) Adicion"&chr(10)&"2) Sustraccion"&chr(10)&"3) Multiplicacion"&chr(10)&"4) Division"&chr(10)&"5) Salir"

   opcion = 0

   while (opcion <> 5 )

        opcion = valor (menu&chr(10)&chr(10)&"Ingrese opcion: ")
        
		if (opcion = 1) then
           
			suma()  
           
		end if 

		if (opcion = 2) then

			resta()

		end if 

		if (opcion = 3) then

			multi()

		end if 

		if (opcion = 4) then

			div()

		end if 

   wend

   fin = msgbox("Calculadora creada por Gonen09",64,"Info")  

End Sub 

'----------------------------------'