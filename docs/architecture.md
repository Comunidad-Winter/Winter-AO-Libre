# Arquitectura de Winter-AO

> 📚 *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Visión general

Winter-AO es un sistema cliente-servidor basado íntegramente en **Visual Basic 6**. Sigue la arquitectura original de Argentum Online: comunicación TCP binaria directa sin capa de abstracción intermedia, persistencia en archivos planos, y lógica de juego ejecutada exclusivamente en el servidor.

```
┌─────────────────────────────────────────────────────────┐
│                  WINTER-AO — ECOSYSTEM                  │
├─────────────────────────────────────────────────────────┤
│                                                         │
│  ┌──────────────┐     ┌──────────────┐                  │
│  │  Auto Update │────▶│  Cliente VB6 │                  │
│  │  (Launcher)  │     │  DX7 o DX8   │                  │
│  └──────────────┘     └──────┬───────┘                  │
│                              │ TCP :7500                │
│                              ▼                          │
│                    ┌──────────────────┐                 │
│                    │  Servidor VB6    │                 │
│                    │  (AOWinter.exe)  │                 │
│                    └──────┬───────────┘                 │
│                           │                             │
│              ┌────────────┼────────────┐                │
│              ▼            ▼            ▼                │
│           Dat/          Maps/       guilds/             │
│         (datos)        (mapas)     (clanes)             │
│                                                         │
│  ┌──────────────┐                                       │
│  │   ZlibWAO   │  (herramienta offline de compresión)  │
│  └──────────────┘                                       │
└─────────────────────────────────────────────────────────┘
```

---

## Stack tecnológico

| Componente | Tecnología | Versión |
|------------|------------|---------|
| Lenguaje | Visual Basic 6 | 6.0 SP6 (inferido) |
| Motor gráfico (Cliente DX7) | DirectX 7 for VB (`DX7VB.DLL`) | 7.0 |
| Motor gráfico (Cliente DX8) | DirectX 8 (`clsDX8Engine.cls`) | 8.x |
| Red (cliente) | Winsock OCX (`MSWINSCK.OCX`) + CSWSK32 | — |
| Red (servidor) | Sockets custom (`wsksock.bas`, `wskapiAO.bas`) | — |
| Internet (cliente) | `MSINET.OCX` (noticias/actualizaciones) | — |
| Compresión | zlib nativa (`zlib.dll`) vía ZlibWAO | — |
| UI controles | `COMCTL32.OCX`, `RICHTX32.OCX` | — |
| Auto-actualización | `Unzip32.dll`, `vbalProgBar6.ocx` | — |

---

## Arquitectura del servidor

### Modelo de ejecución

El servidor VB6 es **monohilo** (VB6 single-thread por diseño). La concurrencia se simula mediante timers de Windows y la API Winsock asíncrona. El entry point es `Sub Main`, definido en el proyecto `SERVER.VBP`.

```
Sub Main (General.bas)
    │
    ├── Carga Server.ini y datos de Dat/
    ├── Inicializa arrays de usuarios, NPCs, objetos
    ├── Inicia el servidor Winsock (wsksock.bas)
    └── Inicia timers (intervalos definidos en Server.ini [INTERVALOS])
            │
            ├── Timer NPC AI     → AI_NPC.bas / MODULO_NPCs.bas
            ├── Timer combate    → SistemaCombate.bas
            ├── Timer usuarios   → Modulo_UsUaRiOs.bas
            ├── Timer trabajos   → Trabajo.bas
            └── Timer limpieza   → General.bas
```

### Capas del servidor

```
┌─────────────────────────────────────────────────────┐
│                   SERVIDOR VB6                      │
├─────────────────────────────────────────────────────┤
│  Red         │ TCP.bas, TCP_HandleData1/2.bas        │
│              │ wsksock.bas, wskapiAO.bas             │
├─────────────────────────────────────────────────────┤
│  Protocolo   │ TCP_HandleData1.bas (paquetes 1/2)   │
│  entrante    │ TCP_HandleData2.bas (paquetes 2/2)   │
├─────────────────────────────────────────────────────┤
│  Lógica      │ SistemaCombate.bas (combate PvP/PvE) │
│  de juego    │ AI_NPC.bas (IA de NPCs)              │
│              │ Trabajo.bas (sistema de trabajos)    │
│              │ modHechizos.bas (sistema de magia)   │
│              │ Comercio.bas (NPC + jugador-jugador) │
│              │ modGuilds.bas/clsClan.cls (clanes)   │
│              │ mdParty.bas/clsParty.cls (grupos)    │
│              │ ModFacciones.bas (Imperio/Caos)      │
│              │ modQuests.bas (sistema de quests)    │
│              │ praetorians.bas (NPCs avanzados)     │
├─────────────────────────────────────────────────────┤
│  Datos       │ FileIO.bas (lectura/escritura E/S)   │
│              │ InvUsuario.bas (inventario usuario)  │
│              │ Modulo_InventANDobj.bas (obj NPC)    │
├─────────────────────────────────────────────────────┤
│  Seguridad   │ SecurityIp.bas (IP por GM)           │
│              │ SebaSecurity.bas (seguridad general) │
│              │ clsAntiMassClon.cls (anti-macros)    │
│              │ modCentinela.bas (centinela)         │
├─────────────────────────────────────────────────────┤
│  Servicios   │ modBanco.bas (banco)                 │
│  adicionales │ modSubasta.bas (subastas)            │
│              │ Mod_Climas.bas (clima)               │
│              │ Mod_Guerras.bas (guerras de facciones│
│              │ PathFinding.bas (A* o similar)       │
│              │ History.bas (historial de acciones)  │
├─────────────────────────────────────────────────────┤
│  Persistencia│ Dat/*.dat (lectura)                  │
│              │ Maps/* (lectura de mapas)            │
│              │ guilds/* (lectura/escritura)         │
│              │ charfile/* (lectura/escritura PJs)   │
└─────────────────────────────────────────────────────┘
```

### Manejo de paquetes TCP

El servidor recibe paquetes de los clientes y los despacha en dos módulos:

- **`TCP_HandleData1.bas`** (~90 KB): Primera mitad del protocolo entrante.
- **`TCP_HandleData2.bas`** (~93 KB): Segunda mitad del protocolo entrante.

La división en dos módulos es por limitación de tamaño de módulo en VB6.

Los paquetes de salida (servidor → cliente) se envían desde todos los módulos de lógica, directamente mediante funciones de `TCP.bas`.

---

## Arquitectura del cliente

### Versión DX7 (`Cliente/`)

```
FrmLanzador (entry point)
    │
    └── frmConnect (pantalla de conexión + noticias vía MSINET)
            │
            └── frmMain (HUD principal del juego)
                    │
                    ├── TileEngine.bas / clsTileEngineX.cls (renderizado DX7)
                    ├── TCP.bas (protocolo de red)
                    ├── General.bas (lógica cliente general)
                    ├── Mod_WAO.bas (extensiones propias del mod)
                    ├── clsAudio.cls (audio DirectSound)
                    ├── clsCustGui.cls (GUI custom ~99 KB)
                    └── [frm* de interfaces secundarias]
```

### Versión DX8 (`WAO DX8/`)

Comparte la misma estructura lógica que DX7 con estas diferencias clave:

- `clsDX8Engine.cls` (~94 KB): nuevo motor gráfico DirectX 8 que reemplaza al DX7.
- `modDX8Fifo.bas`: buffer FIFO para comandos del motor DX8.
- `modDX8Requires.bas`: inicialización de requisitos DX8.
- `Transparencia.bas`: soporte de transparencia para efectos visuales.
- `Volumen.bas`: control de volumen (en DX7 esta función fue eliminada según changelog).
- `msn.bas`: integración con estado de MSN Messenger.

### Capas del cliente

```
┌─────────────────────────────────────────────────────┐
│                   CLIENTE VB6                       │
├─────────────────────────────────────────────────────┤
│  Presentación │ frmMain.frm (HUD principal)         │
│  (UI)         │ frmConnect.frm (login + noticias)   │
│               │ frmCrearPersonaje.frm               │
│               │ frmPanelGm.frm (panel GM)           │
│               │ frmSkills3.frm (habilidades)        │
│               │ +30 formularios adicionales         │
├─────────────────────────────────────────────────────┤
│  Motor        │ clsTileEngineX.cls (DX7, ~195 KB)   │
│  gráfico      │ clsDX8Engine.cls (DX8, ~94 KB)      │
│               │ TileEngine.bas (funciones globales) │
│               │ clsSurfaceMan*.cls (superficies)    │
├─────────────────────────────────────────────────────┤
│  Red          │ TCP.bas (~65 KB) — protocolo AO     │
│               │ Declares.bas (constantes + API)     │
├─────────────────────────────────────────────────────┤
│  Audio        │ clsAudio.cls / clsMP3Player.cls     │
├─────────────────────────────────────────────────────┤
│  Lógica       │ General.bas (lógica general)        │
│  cliente      │ Mod_WAO.bas (extensiones propias)   │
│               │ Procesos.bas (procesos del cliente) │
│               │ modCompression.bas (descompresión)  │
│               │ clsGrapchicalInventory.cls          │
│               │ cDialogos.cls (diálogos in-game)    │
│               │ clsCustomKeys.cls (teclas config)   │
├─────────────────────────────────────────────────────┤
│  Seguridad    │ AntiCheatEngine.bas (DX8 only)      │
│               │ Antidoble.bas (anti-doble cliente)  │
└─────────────────────────────────────────────────────┘
```

---

## Protocolo de red (cliente-servidor)

El protocolo es **binario TCP** heredado del AO 0.11.5. Los paquetes se estructuran con un byte de tipo de operación seguido de datos variables.

- **Encoding**: Windows-1252 (compatibilidad con VB6 strings)
- **Byte order**: Little-Endian (nativo de VB6)
- **Puerto**: 7500 (configurable en `Server.ini`)
- **Cifrado**: Configurable (`Encriptar=1` en `Server.ini`), con CRC (`CrcSubKey=12345`)

El cliente también usa `MSINET.OCX` para peticiones HTTP independientes del protocolo del juego (noticias en `frmConnect`, verificación de actualizaciones).

---

## Sistema de datos (persistencia)

No hay base de datos relacional. Toda la persistencia es basada en archivos:

| Tipo | Formato | Acceso |
|------|---------|--------|
| Objetos del juego | `obj.dat` (INI-like) | Solo lectura al inicio |
| NPCs | `NPCs.dat` / `NPCs-HOSTILES.dat` | Solo lectura al inicio |
| Hechizos | `Hechizos.dat` | Solo lectura al inicio |
| Mapas | `Maps/*.map` (binario) | Solo lectura al inicio |
| Personajes | `charfile/<nombre>.chr` | Lectura/escritura runtime |
| Clanes | `guilds/<nombre>.guild` | Lectura/escritura runtime |
| Configuración | `Server.ini` (INI estándar) | Lectura al inicio |
| Ranking | `Configuracion.ini` | Lectura/escritura runtime |
| Torneos | `Torneos/*.ini` | Lectura/escritura runtime |
| Sugerencias | `SUGS/*.txt` | Solo escritura |
| Ban de IPs | `Dat/BanIps.dat` | Lectura/escritura runtime |

---

## Sistemas específicos de Winter-AO

Estos sistemas **no estaban en el AO 0.11.5 base** y fueron añadidos por el autor:

| Sistema | Módulo principal | Descripción |
|---------|-----------------|-------------|
| Ciclo día/noche | `Mod_WAO.bas` / cliente | Mañana, tarde y noche con efectos visuales |
| Torneos | `Torneos/*.ini` + server | Torneos programables desde .ini |
| Sistema de monturas | `SistemaCombate.bas` + cliente | Monturas con skills propios (beta) |
| Castillos de clanes | `clsClan.cls` + `castillos.dat` | Sistema de conquista de castillos |
| Duelos | Servidor + cliente | Ring de duelos con árbitro NPC |
| Anti-Macros | `clsAntiMassClon.cls` + `modCentinela.bas` | Detección de macros |
| Sistema de cuentas | `frmCuent.frm` + servidor | Cuentas de usuario con múltiples personajes |
| Pasajes | Servidor | Sistema de viaje entre ciudades (requiere skill de navegación) |
| Advertencias GM | `Admin.bas` | Sistema de advertencias con ban automático a las 5 |
| Radio WAO | `frmRadio.frm` (DX8) | Radio integrada desde las opciones |
| Anti-Cheat IP | `SecurityIp.bas` | Restricción de IP por GM |
| Canjes de puntos | `frmcanjes.frm` + servidor | Canje de puntos de torneo |
| Macro configurable | `frmMacros.frm` + `clsCustomKeys.cls` | Macros de usuario y trabajo |
| Chat global | `Mod_WAO.bas` | Canal de chat global con prefijo `;` |
| Gran Poder | Servidor | Habilidad especial de clase |
| Cirujano | Servidor | Personaje NPC médico |
| Portal Tridimensional | `modHechizos.bas` | Hechizo de portal entre mapas |
