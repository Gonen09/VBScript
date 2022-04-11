' /*-------------------------[ Inputbox Simple ]---------------------------------/
'  Autor       : Gonen09
'  Descripción : Operaciones con inputbox simples
'  Versión     : 2.0
'  Fecha       : 14/02/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

q= inputbox ("¿Cual es tu consola favorita?"+chr(10)+chr(10)+"a) Nintendo Wii U"&chr(10)&"b) PS4"+chr(10)+"c) Xbox one"+chr(10)+"d) Todas :3"+chr(10),"Pregunta","Escribe tu respuesta aqui")


'opcion a + aceptar

if (q="a") then

  msgbox ("Compañia: Nintendo")

end if



'opcion b + aceptar

if (q="b") then

  msgbox ("Compañia: Sony")

end if



'opcion c + aceptar

if (q="c") then

  msgbox ("Compañia: Microsoft")

end if



'opcion d + aceptar


if (q="d") then

  msgbox ("Compañia: La mejor eleccion :P")
 
end if



'respuesta por defecto + aceptar

if (q="Escribe tu respuesta aqui") then

  msgbox ("Para la proxima escribe una opcion :P")

end if


'vacio (dejar en blanco) + aceptar o cancelar

if (q=empty) then  'La constante Empty=vacio y no utiliza "" por ser constante

  msgbox ("Para la proxima ingresa una opcion :P")

end if



'Escribir algo + aceptar

if (q<>"a" and q<>"b" and q<>"c" and q<>"d" and q<>"Escribe tu respuesta aqui" and q<>empty) then ' si q es distinto de todas estas

  msgbox ("Para la proxima ingresa una opcion valida :P")                                         

end if                                                                                            

'Simbologia
' igual -->  = 
' distinto --> <>
' Y --> AND
' O --> OR
