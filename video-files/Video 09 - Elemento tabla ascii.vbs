

salida=false

while (salida=false)

  var = inputbox ("Introduce un número entre 0 y 255 :","Transformar número a código ASCII")

  if isnumeric(var) then

       if (var<256 and var>=0) then

             salida=true
     
       else

             msj= msgbox ("Ingrese un número entre 0 y 255",16,"Transformar número a código ASCII") 
             salida=false

       end if

  else
  
       msj= msgbox ("Ingrese solo números",16,"Transformar número a código ASCII") 
       salida=false
   
  end if 

wend

b= msgbox ("Número:   "&var&chr(10)&"Caracter:   ' "&chr(var)&" '",0,"Respuesta")

