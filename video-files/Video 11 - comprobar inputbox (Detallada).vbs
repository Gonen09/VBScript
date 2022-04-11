' /*-------------------------[ Inputbox avanzado ]-------------------------------/
'  Autor       : Gonen09
'  Descripción : Operaciones con inputbox avanzadas
'  Versión     : 1.0
'  Fecha       : 14/02/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

t="texto" 

a= inputbox ("Ingresa algo: ","TEST")

i= msgbox ("Ingresaste:  ' "&a&" '",64,"INFO")


if a = EMPTY then                                     'SI NO INGRESO NADA, EMPTY = constante que representa cadena vacía

     z = msgbox ("No ingresaste nada",64,"INFO")

else
  
    if isnumeric(a)=TRUE then                         'SI INGRESO UN NUMERO
 
        b= msgbox ("Ingresaste un número",64,"INFO")

        if a>100 then
    
            b= msgbox (a&" es mayor que 100",64,"INFO")

        else

            c= msgbox (a&" es menor que 100",64,"INFO")
        
        end if


    else                                              'DE LO CONTRARIO SE ASUME UNA CADENA O CARACTER
  
        if len (a)=1 then			      'len() = funcion que cuenta numero de caracteres	

            d= msgbox ("Ingresaste un caracter",64,"INFO")

        else

            e= msgbox ("Ingresaste una cadena",64,"INFO")

            if a=t or a="Magister" then

               f= msgbox ("Ingresaste una palabra secreta",64,"INFO")

            end if 

        end if 

     end if

end if