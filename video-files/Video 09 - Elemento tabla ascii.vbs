

salida=false

while (salida=false)

  var = inputbox ("Introduce un n�mero entre 0 y 255 :","Transformar n�mero a c�digo ASCII")

  if isnumeric(var) then

       if (var<256 and var>=0) then

             salida=true
     
       else

             msj= msgbox ("Ingrese un n�mero entre 0 y 255",16,"Transformar n�mero a c�digo ASCII") 
             salida=false

       end if

  else
  
       msj= msgbox ("Ingrese solo n�meros",16,"Transformar n�mero a c�digo ASCII") 
       salida=false
   
  end if 

wend

b= msgbox ("N�mero:   "&var&chr(10)&"Caracter:   ' "&chr(var)&" '",0,"Respuesta")

