

while (salida = false)

   
    variable = inputbox ("Ingresa solo numeros: ","TEST")  'Solo es un ejemplo lo importante es que vaya dentro
							   'de un while para que se repita hasta ingresar la opci�n
    if variable = EMPTY then				   'correcta, en este caso, que lo ingresado fuera un n�mero 
	 						
  	 msgbox "No ingresaste nada"

    else

        if isnumeric(variable)=TRUE then 

             msgbox "Muy bien"
             salida = true
 
        else
      
             msgbox "Error, no ingresaste un n�mero"    
 
    	end if 

    end if 

wend 