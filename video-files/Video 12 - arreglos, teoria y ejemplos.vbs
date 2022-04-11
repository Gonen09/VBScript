' /*------------------------------[ Arreglos ]----------------------------------/
'  Autor       : Gonen09
'  Descripción : Teoría, definición y operación con arreglos
'  Versión     : 1.0
'  Fecha       : 03/03/2014
'  Correo      : gonen.rt@gmail.com
'  GitHub      : Gonen09
'  Licencia    : GNU GPL v3
'  Derechos    : Copyright Gonen09, todos los derechos reservados.
' /-----------------------------------------------------------------------------*/

'DECLARAR UN ARREGLO
'-------------------


'DIM NOMBRE_ARREGLO(DIMENSION_ARREGLO)


'DIM: PALABRA RESERVADA PARA DECLARAR UN ARREGLO

'NOMBRE_ARREGLO: NOMBRE DE LA VARIABLE DEL ARREGLO

'DIMENSION_ARREGLO: CANTIDAD DE DATOS PARA ALMACENAR EN EL ARREGLO 



'EJEMPLO
'-------

dim miarreglo(5)  


'DESCRIPCION
'-----------


'ARREGLO DE TAMAÑO 5, CON LAS POSICIONES 0,1,2,3,4

'LA PRIMERA POSICION DEL ARREGLO EMPIEZA EN CERO Y 
'LA ULTIMA POSICION ES LA DIMENSION - 1 (EN ESTE CASO 5-1 = 4)

'EN GENERAL LOS ARREGLOS SE MANEJAN CON UN INDICE DENOTADO POR LA
'LETRA "i" QUE INDICA LA POSICION Y LO MAS IMPORTANTE UN ARREGLO
'DEBE CONTENER EL MISMO TIPO DE DATOS PARA TODOS SUS ELEMENTOS Y
'LOS ELEMENTOS NO DEBEN ESTAR FUERA DEL RANGO DEL ARREGLO



'CODIGO EJEMPLO: MINI AGENDA DE CONTACTOS
'----------------------------------------

'FORMA SIN ARREGLOS

c1n = inputbox ("Ingrese nombre para contacto uno:","Datos del contacto") 
c1t = inputbox ("Ingrese telefono para contacto uno:","Datos del contacto") 

c2n = inputbox ("Ingrese nombre para contacto dos:","Datos del contacto") 
c2t = inputbox ("Ingrese telefono para contacto dos:","Datos del contacto") 

msgbox "C1 = "+c1n+" "+c1t
msgbox "C2 = "+c2n+" "+c2t


'FORMA CON ARREGLOS

DIM nombres(2)   'ARREGLO DE TIPO STRING (CADENA) PARA NOMBRES DE LOS CONTACTOS
DIM telefonos(2) 'ARREGLO DE TIPO INT (NUMERO ENTERO) PARA TELEFONO DE LOS CONTACTOS 

nombres(0) = inputbox ("Ingrese nombre para contacto uno:","Datos del contacto") 
telefonos(0) = inputbox ("Ingrese telefono para contacto uno:","Datos del contacto") 

nombres(1) = inputbox ("Ingrese nombre para contacto dos:","Datos del contacto") 
telefonos(1) = inputbox ("Ingrese telefono para contacto dos:","Datos del contacto") 

msgbox "C1 = "+nombres(0)+" "+telefonos(0) 
msgbox "C2 = "+nombres(1)+" "+telefonos(1) 


