Attribute VB_Name = "History"
Option Explicit

'-A�adido Global (; y el texto para hablar) Para activarlo hay que pone /Glob AC(Funcional)
'-A�adido Conectar directo desde el frmconnect (Funcional)
'-A�adido Sistema de Noche (Funcional)
'-Eliminado la inactividad (Funcional)
'-A�adido sistema de pasajes (Funcional)
'-Agregado Sistema de Duelos (Funcional)
'-Agregado Launcher, al iniciar el Cliente de ejecuta automaticamente (Funcional)
'-El Sacerdote cura el Veneno (Funcional)
'-Al hacercarse al Sacerdote autoresucita o autocura (Funcional)
'-Cambiado el /Passwd (Funcional)
'-El nombre Siempre se ve(Funcional)
'-Arreglado la Invisibilidad(Funcional)
'-Objetos de Newbie segun la clase (funcional)
'-A�adido Anti-Macros (Funcional)
'-A�adido Gran Poder(Funcional)
'-Al hacerse invisible muestra el tiempo que queda para volver a la normalidad(Funcional)
'-Al paralizarse muertra el tiempo que queda para volver a la normalidad (Funcional)
'-Agregado el Comando invansion (solo GM) (funcional)
'-Arreglado Bug que los GM hablaban en blanco
'-Arreglado bug que los bichos se metian en el cuerpo de los Pj
'-Arreglado bug de al atacar a un NPC se salia
'-Borrado el "�Has recuperado x Puntos de Mana!"
'-Borrado el "No ves nada interesante"
'-Cuando llega al Nivel Maximo aparece un cartelito que pone �Nivel Maximo! (Funcional)
'-A�adido �As ganado X de monedas de Oro! (Funcional)
'-Ahora al subir de nivel dan 5 skill libres
'-Agregado Atributos Asignables (Eliminado, provocaba un bug que no dejaba crear PJ)
'-Agregado conectar desde el FrmConnect direcctamente (Funcional)
'-Ip ocultada (para cambiar la ip modificar Public Const IpServidor As String = "127.0.0.1") (Funcional)
'-A�adido LIMPIARMUNDO (funcional)
'-Agregado Cirujano, sugerido por Lectral (Funcional)
'-A�adido Anticheat
'-Cambiado version de la 0.11.5 a version personalizada 1.0.0 (Funcional)
'-Cuando un Pj o npc falla el Golpe aparece arriba de la cabeza "Falla" (funciona)
'-Reprogramado el FrmMain
'-Reprogramado FrmConnect
'-Mejorado Sistema de Skills (funcional)
'-Agregado Transparencia de techos (Funcional)
'-Arreglado bug de que seguia andando cuando paralizaba
'-Al logearse se ve el estado en el Msn mas el nombre del Pj (Funcional)
'-Agregado Control del Volumen de la m�sica desde las opciones (Eliminado)
'-Agregado mapa al pulsara la Tecla "Q" (Funcional)
'-Agregado bordes negros en las letras del juego (funcional)
'-Agregado transparencia al hacerse invisible (Funcional)
'-Agregado sistema de canjes de puntos de torneos (Funcional)
'-Agregado nuevas ciudades de Inicio y arreglado bug que siempre aparecia en la misma (Funcional)
'-Detectado y Arreglado bug en el sistema de pasajes, no dejaba ir a ningun sitio
'-Agregado para que muestre el nombre del mapa en el FrmMain (Funcional)
'-A�adido Raza Orca (Deshabilitado)
'-Al pasar el cursor por encima de un objeto del inventario nos muestra el nombre en un Label (Funcional)
'-Agregado barra de progreso de nivel (Funcional)
'-Agregado Transparencia en el Mapa, se puede controlar el nivel de transparencia (Funcional)
'-Agregado Minimapa (funcional)
'-Agregado estado del usuario
'-Eliminado seguro de autodestruccion del Servidor.
'-Agregado Sistema de Monturas (Beta)
'-Agregado Poci�n Anti-Estupidez
'-Solucionado bug con el oro, si el npc tiraba 100k o mas no lo daba (Solucionado)
'-A�adido opcion de jugar en modo ventana sin ningun problema (Funcional)
'-A�adido Sistema que no deja abrir mas de 1 server (Funcional)
'-Agregado nuevas funciones en el estado del msn (Funcional)
'-Agregado Sistema de castillos por clanes (Funcional)
'-Arreglado bug que se veia la Espada y Escudo en el caballo (Funcional)
'-Agregado Torneos programables desde un .ini (Funcional)
'-Reparado el modo Party
'-A�adido poder activar y desactivar el Efecto Noche (Funcional)
'-Modificado Sistema noche, agregado Ma�ana, Tarde y Noche, arreglado la lentitud y si cerrabas el cliente y volvias a entrar no aparecia el efecto noche (Funcional)
'-Arreglado Modo Party que habia
'-Agregado cuando vas con el banquero te sale un menu (funcional)
'-Implantacion de la libreria vbDABL.dll (Funcional)
'-Agregado magia Portal Tridimensional, crea un portal (Funcional)
'-Todas las interfaces fueron cambiada de la carpeta Graficos a la carpeta Interfaces del cliente (Funcional)
'-A�adido que requiera Anillo del Poder para poder tirar algunos hechizos
'-Graficos Encriptados con contrase�a (Funcional)
'-A�adido Sistema de noticias en el FrmConnect (Funcional)
'-Agregado otro Anti Cheat (Funcional)
'-Minimapa Desactivable.
'-Agregado Boton para fundar clan en la lista de clanes (Funcional)
'-Agregado Estado del server en el Launcher (Funcional)
'-Agregado Comando /PV NICK y Mensaje, para hablar por privado (No Funciona)
'-Agregado Sistema de Macros (Funcional)
'-Ahora se puede elegir la cabeza deseada al crear un nuevo personaje. (Funcional)
'-Agregado Consola en el server como el de AOReady (Funcional)
'-Agregado Hechizos exclusivos para algunas clases (Funcional)
'-Agregado Sistema de Cuentas (Funcional)
'-Ahora solo se podra comenzar en Ramx (Funcional)
'-Agregado Sistema de emoticonos (Funciona)
'-Agregado Sistema de Dropeo (Funcional)
'-Agregado Cuenta Regresiva (Funcional)
'-Reducido Paquetes al Navegar.
'-Al pegar el Arma se mueve. (Funcional)
'-Agregado Sistema de Subastas, Para subastar se hace /Subastar SLOT@CANTIDAD@PRECIO, Para ofertar /Ofertar CANTIDAD, Para ver informacion /Infosubasta, Para cerrar la subasta /CerrarSubasta (Funcional)
'-Agregado nuevo Anti-Macros
'-A�adido Limite de nick, Minimo 3, maximo 12 (Funcional)
'-Agregado Anti-Doble cliente (Eliminado)
'-Agregado Armas que paralizan (Funcional)
