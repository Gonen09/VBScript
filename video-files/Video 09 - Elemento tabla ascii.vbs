' /*---------------------------[ Elemento ASCII ]--------------------------------/
'  Autor       : Gonen09
'  Descripci�n : Obtener un valor y generar un codigo ascii
'  Versi�n     : 1.0
'  Fecha       : 11/07/2013
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

salida = false

while (salida = false)

  var = inputbox ("Introduce un n�mero entre 0 y 255 :","Transformar n�mero a c�digo ASCII")

  if isnumeric(var) then

       if (var<256 and var>=0) then

             salida=true
     
       else

             msj= msgbox ("Ingrese un n�mero entre 0 y 255",16,"Transformar n�mero a c�digo ASCII") 
             salida=false

       end if

  else
  
       msj = msgbox ("Ingrese solo n�meros",16,"Transformar n�mero a c�digo ASCII") 
       salida=false
   
  end if 

wend

b = msgbox ("N�mero:   "&var&chr(10)&"Caracter:   ' "&chr(var)&" '",0,"Respuesta")

