'''<General>
''' Creado por Shedark.
''' En caso de dudas/arreglos/bugs sobre el programa contactar a:
'''<Contacto>
''' GSZone: www.gs-zone.com.ar, mensaje privado a Shed.
''' Shedark@live.com.ar
'''<Uso>
''' Atención, los parches tienen que subirse con el siguiente nombre: "Parche" Numero de parche ".zip". Ejemplo "Parche2.zip"
''' Tienen que crear en el cliente el siguiente archivo \INIT\Update.ini con la información:
'''     [INIT]
'''     X=0 '1 es el numero de actualizacion
''' Ahora en su host web tienen que subir el parche con el nombre ya explicado, y
''' un archivo de texto que contenga los siguientes datos http://suhost.com\VEREXE.txt :
'''     0
''' Este último es el número de actualización del server, si quieren que descarge un nuevo
''' parche deben subir este archivo con el numero "1" y el "Parche1.zip". Y por cada parche, un numero nuevo.
''' Reemplazar los links en el codigo que dicen "www.tuweb.com" por la pagina donde uploadeaste los archivos
'''<Proceso>
''' 1. Busca archivos.
''' 2. Descarga archivo.
''' 3. Unzipea el archivo.
''' 4. Finaliza.

<Explicacion de las Imagenes>
-INICIO-
-Paso1:
 .Colocar en el cliente, en la carpeta INIT, el archivo que ahi se muestra.
-Paso2:
 .Subir el archivo VEREXE.txt como ahi se muestra a www.tuweb.com

-Agregar parche-
-Actualizar:
 .Modificar el VEREXE.txt con el numero de actualizacion por el que vamos contando.
 .Los nombre de los archivos son ParcheX.zip
 .Poner el contenido en el zip como:
	GRAFICOS/10000.bmp
	Cliente.exe
	MIDIS/asd.mid
