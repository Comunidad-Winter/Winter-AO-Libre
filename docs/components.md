# Componentes de Winter-AO

> 📚 *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Módulos del servidor (`server/Codigo/`)

El servidor contiene **70 archivos** distribuidos en módulos `.bas`, clases `.cls` y formularios `.frm`.

### Módulos principales de lógica

| Archivo | Tamaño | Función |
|---------|--------|---------|
| `TCP.bas` | ~210 KB | Protocolo de red completo: envío y recepción de paquetes. Módulo más grande del servidor. |
| `TCP_HandleData1.bas` | ~90 KB | Despacho de paquetes entrantes del cliente (primera mitad del protocolo). |
| `TCP_HandleData2.bas` | ~93 KB | Despacho de paquetes entrantes del cliente (segunda mitad del protocolo). |
| `praetorians.bas` | ~87 KB | Sistema avanzado de NPCs cooperativos (tipo Pretorianos/guardianes). El módulo más complejo de IA. |
| `Modulo_UsUaRiOs.bas` | ~74 KB | Gestión completa de usuarios conectados: login, logout, movimiento, estadísticas. |
| `FileIO.bas` | ~84 KB | Lectura y escritura de todos los archivos de datos (personajes, NPCs, objetos, mapas). |
| `InvUsuario.bas` | ~64 KB | Sistema de inventario del jugador: equipar, desequipar, usar objetos, mover items. |
| `modHechizos.bas` | ~62 KB | Sistema de magia: casting, efectos, colisión, hechizos especiales de clase. |
| `SistemaCombate.bas` | ~57 KB | Combate PvP y PvE: cálculo de daño, críticos, defensa, golpe/fallo. |
| `Trabajo.bas` | ~60 KB | Sistema de trabajos: minería, pesca, tala, carpintería, herrería. |
| `modGuilds.bas` | ~54 KB | Sistema de clanes: fundación, administración, guerras, castillos. |
| `AI_NPC.bas` | ~39 KB | Inteligencia artificial básica de NPCs: movimiento, persecución, ataque. |
| `Declares.bas` | ~37 KB | Declaraciones globales: constantes, tipos, variables de aplicación, APIs de Windows. |
| `General.bas` | ~43 KB | Funciones generales del servidor: inicialización, utilidades compartidas. |
| `wsksock.bas` | ~39 KB | Implementación de sockets TCP custom (capa de bajo nivel). |
| `wskapiAO.bas` | ~22 KB | API de Winsock adaptada para el protocolo AO. |
| `ModFacciones.bas` | ~17 KB | Sistema de facciones: Imperio y Caos, armaduras imperiales, puntuaciones. |
| `Admin.bas` | ~12 KB | Comandos de administración: ban, kick, advertencias, teleport, spawn. |
| `Comercio.bas` | ~16 KB | Comercio con NPCs mercaderes. |
| `mdlCOmercioConUsuario.bas` | ~9 KB | Comercio entre jugadores. |
| `ModAreas.bas` | ~21 KB | Sistema de zonas/áreas del mapa con reglas específicas (PK, seguras, etc.). |
| `PathFinding.bas` | ~7 KB | Algoritmo de búsqueda de caminos para NPCs. |
| `SecurityIp.bas` | ~11 KB | Verificación de IP por GM: restringe acceso a cuentas de staff por IP. |
| `SebaSecurity.bas` | ~10 KB | Seguridad adicional del servidor. |
| `modBanco.bas` | ~10 KB | Sistema de banco: depósito, retiro de oro y objetos. |
| `modSubasta.bas` | ~10 KB | Sistema de subastas entre jugadores. |
| `modQuests.bas` | ~14 KB | Sistema de misiones (quests). |
| `mdParty.bas` | ~13 KB | Lógica de grupos (party): formación, experiencia compartida, combate. |
| `History.bas` | ~6 KB | Historial de acciones del servidor. |
| `Acciones.bas` | ~13 KB | Acciones de personaje: uso de objetos, interacción con el mundo. |
| `GameLogic.bas` | ~26 KB | Lógica extra del juego: nivel, experiencia, habilidades, reputación. |
| `Matematicas.bas` | ~2 KB | Funciones matemáticas auxiliares. |
| `Modulo_InventANDobj.bas` | ~6 KB | Inventario de NPCs y objetos del mundo. |
| `Mod_Climas.bas` | ~4 KB | Sistema de clima dinámico. |
| `Mod_Guerras.bas` | ~8 KB | Guerras entre facciones: conquista, territorios. |
| `modInvisibles.bas` | ~2 KB | Gestión de personajes invisibles. |
| `modNuevoTimer.bas` | ~3 KB | Sistema de timers de precisión mejorado. |
| `modCentinela.bas` | ~11 KB | Centinela anti-cheat: detección de patrones de comportamiento sospechoso. |
| `modHexaStrings.bas` | ~1 KB | Utilidades de strings hexadecimales. |
| `Modulo_SysTray.bas` | ~2 KB | Icono en la bandeja del sistema. |
| `Queue.bas` | ~1 KB | Cola genérica. |

### Clases del servidor

| Archivo | Tamaño | Función |
|---------|--------|---------|
| `clsClan.cls` | ~26 KB | Clase de clan: datos, operaciones, persistencia de un clan. |
| `clsParty.cls` | ~13 KB | Clase de grupo de personajes. |
| `clsIniReader.cls` | ~12 KB | Lector de archivos INI. |
| `ConsultasPopulares.cls` | ~7 KB | Sistema de consultas frecuentes de jugadores. |
| `ModCola.cls` | ~3 KB | Cola de operaciones. |
| `clsEstadisticasIPC.cls` | ~3 KB | Estadísticas de comunicación entre procesos. |
| `clsAntiMassClon.cls` | ~1 KB | Anti-clonación masiva (anti-macro de creación). |
| `clsMapSoundManager.cls` | ~3 KB | Gestor de sonidos por zona del mapa. |
| `clsdicc.cls` | ~3 KB | Diccionario genérico. |
| `clsLeerInis.cls` | ~4 KB | Lector de INIs específico del servidor. |
| `cColaArray.cls` | ~2 KB | Cola basada en array. |
| `cSolicitud.cls` | ~0.5 KB | Solicitud de clan/guild. |
| `cGarbage.cls` | ~0.5 KB | Recolección de objetos descartados. |
| `Cls_InterGTC.cls` | ~1 KB | Interfaz GTC (Game Tracking Component). |

### Formularios del servidor

| Formulario | Función |
|------------|---------|
| `frmMain.frm` | Consola principal del servidor (log de eventos, estado). |
| `frmServidor.frm` | Panel de control del servidor con herramientas de administración. |
| `FrmInterv.frm` | Panel de configuración de intervalos de timers en runtime. |
| `frmAdmin.frm` | Formulario de administración avanzada. |
| `frmEstadisticas.frm` | Panel de estadísticas del servidor. |
| `frmUserList.frm` | Lista de usuarios conectados. |
| `frmDebugNpc.frm` | Depuración de NPCs en runtime. |
| `frmDebugSocket.frm` | Depuración de sockets de red. |
| `frmTrafic.frm` | Monitor de tráfico de red. |
| `frmConID.frm` | Formulario de identificación de conexión. |
| `FrmStat.frm` | Estadísticas simplificadas. |
| `frmCargando.frm` | Pantalla de carga al iniciar el servidor. |

---

## Módulos del cliente DX7 (`Cliente/CODIGO/`)

El cliente DX7 contiene **135 archivos**.

### Módulos principales

| Archivo | Tamaño | Función |
|---------|--------|---------|
| `clsTileEngineX.cls` | ~195 KB | Motor gráfico principal DirectX 7: renderizado de mapas por capas, personajes, efectos, animaciones. Módulo más grande del cliente. |
| `clsCustGui.cls` | ~99 KB | Sistema de GUI personalizada: controles, ventanas, botones, barras de progreso dibujadas con DX. |
| `TCP.bas` | ~65 KB | Protocolo TCP del cliente: serialización y deserialización de todos los paquetes del protocolo AO. |
| `modCompression.bas` | ~72 KB | Descompresión de archivos de gráficos del juego (formato propio con zlib). |
| `Mod_WAO.bas` | ~50 KB | Extensiones propias del mod: sistema de noche, pasajes, torneos, funciones extra. |
| `General.bas` | ~41 KB | Lógica general del cliente: variables globales, inicialización, funciones compartidas. |
| `Declares.bas` | ~14 KB | Declaraciones de API de Windows, constantes del protocolo, tipos de datos. |
| `frmMain.frm` | ~81 KB | HUD principal del juego: input, renderizado del juego, manejo de eventos. Formulario más grande. |
| `clsAudio.cls` | ~14 KB | Sistema de audio: DirectSound, reproducción de música y efectos. |
| `clsCustomKeys.cls` | ~12 KB | Sistema de teclas personalizadas: lectura, almacenamiento, remapeo. |
| `clsGrapchicalInventory.cls` | ~15 KB | Inventario gráfico: drag & drop de objetos con DirectX. |
| `clsSurfaceManDyn.cls` | ~13 KB | Gestor de superficies dinámicas DirectX (sprites animados). |
| `clsSurfaceManStatic.cls` | ~10 KB | Gestor de superficies estáticas DirectX (fondos, tiles). |
| `TileEngine.bas` | ~81 KB | Funciones globales del motor de tiles: helpers de renderizado. |
| `DX_InIt.bas` | ~4 KB | Inicialización de DirectX (dispositivo, ventana, modo de video). |
| `Procesos.bas` | ~5 KB | Gestión de procesos del cliente (detección de múltiples instancias). |
| `cDialogos.cls` | ~10 KB | Sistema de diálogos flotantes sobre los personajes en pantalla. |
| `clsMP3Player.cls` | ~4 KB | Reproductor MP3 vía Windows Media Player integrado. |

### Formularios del cliente DX7 (principales)

| Formulario | Función |
|------------|---------|
| `FrmLanzador.frm` | **Entry point**: Lanzador del juego, verifica actualizaciones. |
| `frmConnect.frm` | Pantalla de conexión: login, registro, noticias del servidor vía HTTP. |
| `frmMain.frm` | HUD principal: mapa, chat, inventario, barra de stats. |
| `frmCrearPersonaje.frm` | Creación de personaje: clase, raza, cabeza, alineación. |
| `frmSkills3.frm` | Panel de habilidades del personaje. |
| `FrmEstadisticas.frm` | Estadísticas detalladas del personaje. |
| `frmOpciones.frm` | Opciones del juego: audio, video, teclas. |
| `frmPanelGm.frm` | Panel de Game Master. |
| `frmComerciar.frm` | Interfaz de comercio con NPC. |
| `frmComerciarUsu.frm` | Interfaz de comercio entre jugadores. |
| `FrmBoveda.frm` | Bóveda / almacenamiento privado del personaje. |
| `frmBanco.frm` / `frmBancoObj.frm` | Banco de oro y banco de objetos. |
| `frmGuild*.frm` (×8) | Formularios del sistema de clanes. |
| `frmAmigos.frm` | Lista de amigos. |
| `FrmMap.frm` | Mini-mapa del juego. |
| `frmMacros.frm` | Gestor de macros de usuario. |
| `frmCustomKeys.frm` | Configuración de teclas personalizadas. |
| `frmCanjes.frm` | Canje de puntos de torneo. |
| `frmConsolaTorneo.frm` / `frmConsolaTorneoUS.frm` | Consola del sistema de torneos (GM / usuario). |
| `frmCuent.frm` | Gestión de cuenta de usuario. |
| `frmSubasta.frm` | Interfaz de subastas. |
| `CreandoCuenta.frm` | Formulario de creación de cuenta. |
| `Form2.frm` | Formulario de selección de personaje tras login. |
| `frmQuests.frm` | Panel de misiones. |
| `frmReproductor.frm` | Reproductor de música integrado. |
| `frmCharInfo.frm` | Información detallada de personaje (vista de terceros). |
| `FrmCredits.frm` | Créditos del juego. |
| `FrmHogar.frm` | Sistema de hogar/casa. |
| `frmEligeAlineacion.frm` | Selección de alineación (Imperio/Caos). |

---

## Módulos adicionales del cliente DX8 (`WAO DX8/CODIGO/`)

El cliente DX8 contiene **149 archivos** y comparte la mayoría de módulos con DX7. Las diferencias son:

| Archivo (DX8 only) | Tamaño | Función |
|--------------------|--------|---------|
| `clsDX8Engine.cls` | ~94 KB | Motor gráfico DirectX 8: reemplaza el sistema DX7, soporta FPS libres y efectos adicionales. |
| `modDX8Fifo.bas` | ~7 KB | Buffer FIFO de comandos para el motor DX8. |
| `modDX8Requires.bas` | ~4 KB | Verificación de requisitos de DirectX 8 al arrancar. |
| `Transparencia.bas` | ~2 KB | Soporte de transparencia de superficies (alpha blending). |
| `Volumen.bas` | ~13 KB | Control de volumen de audio (eliminado en DX7 por bug). |
| `msn.bas` | ~2 KB | Estado de MSN Messenger al conectarse al juego. |
| `AntiCheatEngine.bas` | ~0.5 KB | Anti-cheat engine (stub o implementación básica). |
| `Antidoble.bas` | ~1 KB | Prevención de doble cliente. |
| `Carteles.bas` | ~2 KB | Sistema de carteles/letreros en el mapa. |
| `modHexaStrings.bas` | ~1 KB | Utilidades hexadecimales (también en DX7 pero con variaciones). |
| `frmRadio.frm` | — | Formulario de radio integrada (eliminada en DX7). |
| `frmRecuperar.frm` | — | Recuperación de contraseña. |
| `GameIni.bas` | ~1.5 KB | Lectura de configuración del juego desde INI. |
| `EleccionCabezas.bas` | ~0.1 KB | Módulo stub para elección de cabezas al crear personaje. |
| `MODOS_DE_VIDEO.bas` | ~1 KB | Gestión de modos de video DirectX 8. |
| `ModAreas.bas` | ~1 KB | Versión reducida del módulo de áreas (solo cliente). |
| `Mod_ErrorLOG.bas` | ~0.1 KB | Log de errores del cliente (stub). |
| `Mod_Macros.bas` | ~2 KB | Módulo de macros de cliente. |
| `APIdeclaraciones.bas` | ~2 KB | Declaraciones adicionales de API de Windows. |

---

## Sistema de Auto Update (`Auto Update/`)

| Archivo | Función |
|---------|---------|
| `frmMain.frm` | Formulario principal: descarga, descompresión y barra de progreso. |
| `ModGeneral.bas` | Lógica de actualización: petición HTTP, verificación de versión, llamada a Unzip32. |
| `AutoUpdate.vbp` | Proyecto VB6. Entry point: `frmMain`. |
| `Unzip32.dll` | Biblioteca de descompresión ZIP nativa (32-bit). |
| `vbalProgBar6.ocx` | Control de barra de progreso personalizado. |

---

## Archivos de datos del juego (`server/Dat/`)

| Archivo | Tamaño | Formato | Contenido |
|---------|--------|---------|-----------|
| `obj.dat` | ~196 KB | INI por secciones | Catálogo completo de objetos del juego (índice, nombre, tipo, stats). |
| `NPCs.dat` | ~42 KB | INI por secciones | Definiciones de NPCs: nombre, gráfico, stats, loot, comportamiento. |
| `NPCs-HOSTILES.dat` | ~38 KB | INI por secciones | Tabla de NPCs hostiles para zonas específicas. |
| `Hechizos.dat` | ~31 KB | INI por secciones | Definiciones de hechizos: nombre, tipo, efecto, mana, requerimientos. |
| `AreasStats.dat` | ~37 KB | INI por secciones | Estadísticas y configuración de cada área del mapa. |
| `ArmadurasHerrero.dat` | ~4 KB | Custom INI | Recetas de armaduras para el herrero. |
| `ArmasHerrero.dat` | ~1 KB | Custom INI | Recetas de armas para el herrero. |
| `ObjCarpintero.dat` | ~1 KB | Custom INI | Recetas del carpintero. |
| `Invokar.dat` | ~1 KB | Custom INI | Tabla de invocaciones (mascotas/criaturas invocables). |
| `Head.dat` | ~0.2 KB | Custom | Datos de cabezas de personaje disponibles. |
| `Map.dat` | ~0.1 KB | Custom | Metadatos de mapas (nombres, restricciones). |
| `Ciudades.Dat` | ~26 B | Custom | Lista de ciudades y coordenadas de inicio. |
| `Motd.ini` | ~1 KB | INI | Mensaje del día mostrado al conectar. |
| `Propagandas.ini` | ~0.4 KB | INI | Mensajes de propaganda/anuncios periódicos. |
| `BanIps.dat` | ~48 B | Lista | IPs baneadas del servidor. |
| `castillos.dat` | ~71 B | Custom INI | Estado de conquista de castillos de clanes. |
| `CastillosEO.dat` | ~35 B | Custom | Datos adicionales de castillos. |
| `consultas.dat` | ~252 B | Custom | Consultas populares registradas. |
| `apuestas.dat` | ~56 B | Custom | Datos del sistema de apuestas. |
| `bkNPCs.dat` | ~24 KB | INI | Backup de NPCs (datos de respaldo). |
| `NombresInvalidos.txt` | ~732 B | Lista | Nombres de personaje prohibidos. |
| `Help.dat` | ~773 B | Texto | Texto de ayuda del juego. |

---

## Herramienta ZlibWAO (`ZlibWAO/`)

Utilidad de compresión/descompresión de gráficos del cliente.

| Archivo | Función |
|---------|---------|
| `ZlibWao.exe` | Ejecutable de compresión. Opera sobre las carpetas `GRAFICOS/` y `GRAFICOS COMPRIMIDOS/`. |
| `zlib.dll` | Biblioteca zlib nativa (32-bit). |
| `GRAFICOS/` | Carpeta de entrada: gráficos originales sin comprimir. |
| `GRAFICOS COMPRIMIDOS/` | Carpeta de salida: gráficos comprimidos para distribuir con el cliente. |

---

## Archivos de diseño (`PSD/`)

| Archivo | Contenido |
|---------|-----------|
| `PSD Interfaces WAO.rar` | ~12 MB comprimido. Archivos fuente Photoshop (`.psd`) de todas las interfaces gráficas del cliente Winter-AO. Recurso valioso para re-crear o modificar las interfaces originales. |
