' /*---------------------------[ Elemento ASCII ]--------------------------------/
'  Autor       : Gonen09
'  Descripción : Obtener un valor y generar un codigo ascii
'  Versión     : 1.0
'  Fecha       : 11/07/2013
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

salida = false

while (salida = false)

  var = inputbox ("Introduce un número entre 0 y 255 :","Transformar número a código ASCII")

  if isnumeric(var) then

       if (var<256 and var>=0) then

             salida=true
     
       else

             msj= msgbox ("Ingrese un número entre 0 y 255",16,"Transformar número a código ASCII") 
             salida=false

       end if

  else
  
       msj = msgbox ("Ingrese solo números",16,"Transformar número a código ASCII") 
       salida=false
   
  end if 

wend

b = msgbox ("Número:   "&var&chr(10)&"Caracter:   ' "&chr(var)&" '",0,"Respuesta")

