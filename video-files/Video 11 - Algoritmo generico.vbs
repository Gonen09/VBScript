' /*------------------------------[ Ciclo While ]--------------------------------/
'  Autor       : Gonen09
'  Descripci�n : Ejemplo de uso ciclo While
'  Versi�n     : 2.0
'  Fecha       : 14/02/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

'Solo es un ejemplo lo importante es que vaya dentro
'de un while para que se repita hasta ingresar la opci�n
'correcta, en este caso, que lo ingresado sea un n�mero

salida = false

while (salida = false)

   
    variable = inputbox ("Ingresa solo numeros: ","TEST")  
							   
    if variable = EMPTY then				    
	 						
  	 msgbox "No ingresaste nada"

    else

        if isnumeric(variable) = true then 	'Funcion para comprobar n�mero

             msgbox "Muy bien"
             salida = true
 
        else
      
             msgbox "Error, no ingresaste un n�mero"    
 
    	end if 

    end if 

wend 