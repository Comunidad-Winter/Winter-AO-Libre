# Componentes y Módulos de Winter-AO

> **Esta documentación fue generada por y para Comunidad-Winter con el objetivo de preservar recursos de Argentum Online.**

Desglose detallado de cada módulo del código fuente, verificado directamente contra los archivos `.vbp` y el contenido de las carpetas `Codigo/`.

---

## Servidor (`server/Codigo/`) — 73 archivos, ~2 MB

### Módulos principales (.bas)

| Módulo | Archivo | Tamaño | Función verificada |
|--------|---------|--------|-------------------|
| `General` | `General.bas` | 49 KB | **Entry point** (`Sub Main` línea 208), inicialización global, carga de datos |
| `Declaraciones` | `Declares.bas` | 39 KB | Todos los `Type`, `Enum`, `Const` y variables globales del servidor |
| `Protocol` | `Protocol.bas` | **680 KB** | Serialización/deserialización de paquetes. Archivo más grande del proyecto |
| `ES` (FileIO) | `FileIO.bas` | 87 KB | Lectura/escritura de archivos INI, `.dat` y binarios legacy |
| `UsUaRiOs` | `Modulo_UsUaRiOs.bas` | 80 KB | Gestión de usuarios conectados, login, logout, estadísticas |
| `Trabajo` | `Trabajo.bas` | 67 KB | Sistemas de recolección: tala, minería, pesca, herrería, carpintería |
| `TCP` | `TCP.bas` | 64 KB | Gestión de conexiones TCP, aceptación de clientes, envío de datos |
| `SistemaCombate` | `SistemaCombate.bas` | 61 KB | Cálculos de daño, evasión, probabilidad de impacto PvP/PvE |
| `modGuilds` | `modGuilds.bas` | 60 KB | Clanes: creación, administración, guerras, alianzas |
| `Extra` (GameLogic) | `GameLogic.bas` | 47 KB | Lógica miscelánea del juego, intervalos, restricciones |
| `wsksock` | `wsksock.bas` | 41 KB | Wrapper directo sobre la API Winsock de Windows |
| `AI` | `AI_NPC.bas` | 39 KB | IA de NPCs: persecución, patrullaje, agresión |
| `NPCs` | `MODULO_NPCs.bas` | 34 KB | Spawn, muerte, drops, propiedades de NPCs |
| `ModFacciones` | `ModFacciones.bas` | 31 KB | Facciones (Ejército Imperial / Legión Oscura), rangos, armaduras |
| `modSendData` | `modSendData.bas` | 23 KB | Broadcast de paquetes a áreas, mapas o todos los usuarios |
| `modQuestSystem` | `modQuestSystem.bas` | 21 KB | Misiones: asignación, progreso, recompensas |
| `mdParty` | `mdParty.bas` | 19 KB | Sistema de grupos/parties |
| `wskapiAO` | `wskapiAO.bas` | 19 KB | Declaraciones de la API Winsock nativa |
| `Admin` | `Admin.bas` | 17 KB | Comandos de administración (ban, kick, teleport, etc.) |
| `ModAreas` | `ModAreas.bas` | 17 KB | Sistema de áreas visibles y broadcasting por zona |
| `AutoTorneos` | `AutoTorneos.bas` | 17 KB | Torneos PvP automatizados (máquina de estados) |
| `Acciones` | `Acciones.bas` | 16 KB | Acciones del usuario (equipar, usar, tirar, etc.) |
| `modHechizos` | `modHechizos.bas` | 84 KB | Resolución de hechizos: daño, curación, estados alterados |
| `Comercio` | `Comercio.bas` | 14 KB | Comercio con NPCs |
| `modCentinela` | `modCentinela.bas` | 13 KB | Anti-macro automático |
| `Statistics` | `Statistics.bas` | 13 KB | Estadísticas del servidor |
| `mdlComercioConUsuario` | `mdlCOmercioConUsuario.bas` | 13 KB | Comercio entre jugadores |
| `SecurityIp` | `SecurityIp.bas` | 11 KB | Baneos por IP, límites de conexión |
| `modBanco` | `modBanco.bas` | 10 KB | Sistema de bóveda/banco |
| `Makro` | `Makro.bas` | 10 KB | Detección de macros |
| `InvNpc` | `Modulo_InventANDobj.bas` | 8 KB | Inventario de NPCs y objetos del mundo |
| `InvUsuario` | `InvUsuario.bas` | 69 KB | Sistema completo de inventario del jugador |
| `Mod_Cuentas` | `Mod_Cuentas.bas` | 7 KB | Sistema de cuentas de usuario |
| `ModHistorial` | `ModHistorial.bas` | 7 KB | Log de acciones de jugadores |
| `PathFinding` | `PathFinding.bas` | 7 KB | Pathfinding A* para NPCs |
| `modNuevoTimer` | `modNuevoTimer.bas` | 6 KB | Sistema de timers del servidor |
| `SysTray` | `Modulo_SysTray.bas` | 3 KB | Icono en bandeja del sistema |
| `Matematicas` | `Matematicas.bas` | 2 KB | Funciones matemáticas auxiliares |
| `LwKClimas` | `LwKClimas.bas` | 2 KB | Sistema de clima por mapa |
| `Queue` | `Queue.bas` | 2 KB | Cola de procesamiento |
| `Eventos` | `Eventos.bas` | 1 KB | Sistema de eventos (stub) |
| `Characters` | `Characters.bas` | 1 KB | Utilidades de personaje |
| `Mod_InterGTC` | `Mod_InterGTC.bas` | 2 KB | Inter-servidor (parcial) |

### Clases (.cls)

| Clase | Archivo | Propósito |
|-------|---------|-----------|
| `clsByteQueue` | `clsByteQueue.cls` | Buffer circular para cola de paquetes de red |
| `clsClan` | `clsClan.cls` | Entidad Clan con toda su lógica |
| `clsIniManager` | `clsIniManager.cls` | Escritor/lector optimizado de archivos INI |
| `clsIniReader` | `clsIniReader.cls` | Lector rápido de archivos INI |
| `clsParty` | `clsParty.cls` | Entidad Party/grupo |
| `clsByteBuffer` | `clsByteBuffer.cls` | Buffer de bytes para serialización |
| `clsAntiMassClon` | `clsAntiMassClon.cls` | Detección de clonación masiva |
| `clsMapSoundManager` | `clsMapSoundManager.cls` | Sonido ambiental por mapa |
| `cCola` | `ModCola.cls` | Cola genérica |
| `cGarbage` | `cGarbage.cls` | Recolector de objetos expirados |
| `cSolicitud` | `cSolicitud.cls` | Solicitudes de clan |
| `clsInterGTC` | `clsInterGTC.cls` | Comunicación inter-servidor |
| `diccionario` | `clsdicc.cls` | Diccionario clave-valor |

### Formularios del servidor

| Formulario | Propósito |
|-----------|-----------|
| `frmMain.frm` | Ventana principal del servidor |
| `frmServidor.frm` | Panel de configuración |
| `frmCargando.frm` | Splash de carga |
| `FrmInterv.frm` | Configuración de intervalos |
| `FrmStat.frm` | Estadísticas en tiempo real |
| `frmAdmin.frm` | Panel administrativo |
| `frmUserList.frm` | Lista de usuarios conectados |
| `frmTrafic.frm` | Monitor de tráfico de red |
| `frmConID.frm` | Consulta por ID de conexión |
| `frmDebugNpc.frm` | Debug de NPCs |

---

## Cliente (`Cliente/CODIGO/`) — 77 archivos fuente, ~1.4 MB

### Módulos principales (.bas)

| Módulo | Archivo | Tamaño | Función verificada |
|--------|---------|--------|-------------------|
| `Protocol` | `Protocol.bas` | **321 KB** | Protocolo completo cliente-servidor |
| `Mod_TileEngine` | `TileEngine.bas` | 120 KB | Motor de renderizado 2D isométrico |
| `ProtocolCmdParse` | `ProtocolCmdParse.bas` | 72 KB | Parser de comandos de texto del usuario |
| `Mod_General` | `General.bas` | 44 KB | **Entry point** (`Sub Main` línea 558), inicialización |
| `modCompression` | `modCompression.bas` | 38 KB | Descompresión zlib de archivos `.WAO` |
| `Multimod` | `MultiMod.bas` | 21 KB | Funciones multi-propósito |
| `Mod_Declaraciones` | `Declares.bas` | 20 KB | Types, enums, constantes, APIs Win32 |
| `ModCarga` | `ModCarga.bas` | 16 KB | Carga de recursos y datos de inicio |
| `ModSeguridad` | `ModSeguridad.bas` | 11 KB | Anti-cheat: detección de procesos y ventanas |

### Clases (.cls)

| Clase | Archivo | Tamaño | Propósito |
|-------|---------|--------|-----------|
| `clsJpeg` | `clsJpeg.cls` | 71 KB | Decodificación de imágenes JPEG |
| `clsByteQueue` | `clsByteQueue.cls` | 41 KB | Buffer de paquetes de red |
| `clsAudio` | `clsAudio.cls` | 31 KB | Música (MP3 vía Quartz) y efectos de sonido |
| `clsLight` | `clsLight.cls` | 17 KB | Sistema de iluminación dinámica |
| `clsIniReader` | `clsIniReader.cls` | 14 KB | Lector de archivos INI |
| `clsGrapchicalInventory` | `clsGrapchicalInventory.cls` | 13 KB | Inventario gráfico renderizado en DX8 |
| `clsCustomKeys` | `clsCustomKeys.cls` | 12 KB | Teclas configurables |
| `clsSurfaceManDyn` | `clsSurfaceManDyn.cls` | 8 KB | Gestor dinámico de texturas DX8 |
| `clsDialogs` | `clsDialogs.cls` | 8 KB | Diálogos flotantes sobre personajes |
| `clsTimer` | `MainTimer.cls` | 7 KB | Timer de alta resolución |
| `clsGuildDlg` | `clsGuildDlg.cls` | 4 KB | Diálogo de clanes |
| `cAssociate` | `cAssociate.cls` | 2 KB | Asociación clave-valor |
| `clsSurfaceManager` | `clsSurfaceManager.cls` | 1 KB | Gestor de surfaces (interfaz) |

---

## WorldEditor (`WorldEditor/Codigo/`) — 52 archivos

Editor visual de mapas con motor DirectX 8 propio. Autor: Lorwik (RincondelAO).

### Módulos clave

| Módulo | Propósito |
|--------|-----------|
| `modGeneral.bas` | Lógica principal del editor |
| `modEdicion.bas` (41 KB) | Todas las operaciones de edición de tiles |
| `modMapIO.bas` (34 KB) | Lectura/escritura de archivos `.map` binarios |
| `modRenderer.bas` (31 KB) | Pipeline de renderizado DX8 del editor |
| `TileEngine.bas` (27 KB) | Motor de tiles adaptado para el editor |
| `modPaneles.bas` | Gestión de paneles de herramientas |
| `modIndices.bas` | Manejo de índices de gráficos (Grh) |
| `modCompression.bas` | Descompresión de recursos |
| `clsDX8Engine.cls` (82 KB) | Motor DirectX 8 completo |
| `clsSurfaceManager.cls` | Gestión de texturas |
| `clsLight.cls` | Iluminación en el editor |

### Formularios

| Formulario | Propósito |
|-----------|-----------|
| `frmMain.frm` (170 KB) | Ventana principal con lienzo, paneles y toolbars |
| `frmUnionAdyasente.frm` (44 KB) | Herramienta de unión de mapas adyacentes |
| `frmMapInfo.frm` | Propiedades del mapa |
| `frmConfigSup.frm` | Configuración de superficies/capas |
| `frmOptimizar.frm` | Optimización de mapas |
| `frmGRHaBMP.frm` | Exportador de gráficos a BMP |
| `frmMusica.frm` | Asignación de música por mapa |
| `frmInformes.frm` | Reportes sobre el mapa |

---

## Editor de Partículas (`Editor de Particulas/Codigos/`)

Herramienta visual autónoma para diseñar efectos de partículas. Autor: Lorwik (RincondelAO).

- **Motor**: DirectX 8 (`clsDX8Engine.cls`, `modDX8Requires.bas`)
- **Renderizado**: `TileEngine.bas` adaptado para previsualización
- **Gestión de texturas**: `clsSurfaceManDynDX8.cls`
- **UI**: `frmMain.frm` (editor principal), `Form1.frm` (auxiliar), `frmCargando.frm`

---

## Indexador (`Indexador/`)

Herramienta para generar los índices de gráficos (Grh) que mapean IDs numéricos a coordenadas dentro de texturas. Autor: LwK Project.

- **Módulos**: `LoadGrh.bas` (13 KB, carga y parseo de Grh), `General.bas` (9 KB)
- **Formularios**: `frmmain.frm` (interfaz principal), `frmExtra.frm` (opciones avanzadas), `frmAbout.frm`, `frmcarga.frm`

---

## LwK-Universal (`LwK-Universal/Codigos/`)

Compresor/descompresor de archivos `.WAO`. Empaqueta carpetas de assets en contenedores comprimidos con zlib.

- **Módulo principal**: `modCompression.bas` — lógica de compresión/descompresión
- **UI**: `frmmain.frm` — interfaz de selección de carpetas y operación
- **Dependencia**: `zlib.dll` (incluida)
- **Herramientas adicionales incluidas**: `Conversor.exe`, `AO 0.12.X Minimap Color Finder.exe`, `LwK-Particle Editor.exe`

---

## PathHelper (`PathHelper/`)

Utilidad de diagnóstico para usuarios con problemas de parcheo. Autor: Lorwik.

- **Módulo**: `General.bas` (2 KB) — verificación de rutas y dependencias
- **UI**: `Form1.frm` — formulario único
- **Dependencia**: `MSINET.OCX` (transferencia de archivos por Internet)

---

## Registrador de Librerías (`Registrador de librerias/`)

Herramienta para registrar componentes COM (`.ocx`, `.dll`) en el sistema Windows. Autor: Lorwik (LwK-Projects).

- **Módulo**: `General.bas` — lógica de registro via `regsvr32`
- **UI**: `frmMain.frm` — interfaz con log de operaciones (`RICHTX32.OCX`)
- **Dependencia**: `Microsoft Scripting Runtime` (FileSystemObject)

---

## Interfaces (`Interfaces/`)

Carpeta con **61 archivos de diseño** (`.psd` de Adobe Photoshop) de todas las ventanas del cliente. No contiene código ejecutable.

Archivos destacados: `Conectar.psd`, `CrearPersonaje.psd`, `Inventario.psd`, `Comercio.psd`, `Estadisticas.psd`, `Mapa.psd`, `Launcher.psd`, `GuildBrief.psd`, `VntQuest.psd`, entre otros. El archivo más grande es `ModernLeft.psd` (~37 MB).

---

## Datos del Servidor (`server/Dat/`)

| Archivo | Tamaño | Contenido |
|---------|--------|-----------|
| `obj.dat` | 190 KB | ~600+ definiciones de objetos del juego |
| `NPCs.dat` | 119 KB | Definiciones de NPCs (stats, IA, drops) |
| `Hechizos.dat` | 63 KB | Definiciones de hechizos |
| `AreasStats.dat` | 35 KB | Estadísticas por área del mapa |
| `QUESTS.DAT` | 3 KB | Definiciones de misiones |
| `ArmadurasHerrero.dat` | 3 KB | Recetas de armaduras |
| `Balance.dat` | 2 KB | Parámetros de balance del juego |
| `ArmasHerrero.dat` | 1 KB | Recetas de armas |
| `ObjCarpintero.dat` | 1 KB | Recetas de carpintería |
| `ObjCanjes.dat` | 1 KB | Objetos canjeables |
| `Motd.ini` | 1 KB | Mensaje del día |
| `Invokar.dat` | 1 KB | Criaturas invocables |
| `Map.dat` | 122 B | Configuración de mapas |
