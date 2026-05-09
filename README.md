# Winter-AO

> **Mod del servidor y cliente de Argentum Online 0.11.5 con soporte DirectX 8 y múltiples extensiones de gameplay.**

---

> 📚 *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

## Qué es

**Winter-AO** es un mod de código fuente completo (cliente + servidor) del MMORPG argentino *Argentum Online*, basado en la versión **0.11.5**. El proyecto fue desarrollado por "Lorwik" y distribuido públicamente como recurso libre para la comunidad.

Está escrito íntegramente en **Visual Basic 6** y extiende significativamente el AO base con sistemas propios: ciclo día/noche, sistema de torneos, monturas, castillos por clanes, anti-cheat, sistema de cuentas, macros configurables, radio integrada, entre otros.

> **Estado del proyecto**: Preservación histórica. El autor liberó el código al dejar de desarrollarlo. No está en desarrollo activo.

---

## Estructura del repositorio

```
Winter-AO/
├── Cliente/            ← Proyecto VB6 del cliente (motor DX7 / versión 1.x)
│   ├── Client.vbp      ← Archivo de proyecto VB6 (entry point: FrmLanzador)
│   └── CODIGO/         ← Código fuente del cliente (135 archivos: .frm, .bas, .cls)
│
├── WAO DX8/            ← Proyecto VB6 del cliente con motor DirectX 8 (versión 2.0)
│   ├── Client.vbp      ← Proyecto VB6 actualizado
│   └── CODIGO/         ← Código fuente DX8 (149 archivos: .frm, .bas, .cls)
│
├── server/             ← Proyecto VB6 del servidor
│   ├── SERVER.VBP      ← Archivo de proyecto VB6 (entry point: Sub Main)
│   ├── Server.ini      ← Configuración principal del servidor
│   ├── Configuracion.ini ← Configuración de ranking en memoria
│   ├── AOWinter.exe    ← Binario compilado del servidor (incluido en repo)
│   ├── Codigo/         ← Código fuente del servidor (70 archivos: .bas, .cls, .frm)
│   ├── Dat/            ← Archivos de datos del juego (NPCs, objetos, hechizos, mapas)
│   ├── Maps/           ← Archivos de mapa binarios
│   ├── guilds/         ← Archivos de clanes (generados en runtime)
│   ├── Torneos/        ← Configuración de torneos (.ini)
│   ├── criticos/       ← Archivos de log o datos críticos
│   └── SUGS/           ← Sugerencias recibidas vía comando /SUG
│
├── Auto Update/        ← Proyecto VB6 del sistema de auto-actualización
│   ├── AutoUpdate.vbp  ← Proyecto VB6 del launcher
│   ├── frmMain.frm     ← Formulario principal del launcher
│   └── ModGeneral.bas  ← Lógica de descarga y actualización
│
├── ZlibWAO/            ← Herramienta de compresión de gráficos
│   ├── ZlibWao.exe     ← Compresor/descompresor ejecutable
│   ├── zlib.dll        ← Biblioteca zlib nativa
│   ├── GRAFICOS/       ← Carpeta de gráficos descomprimidos
│   └── GRAFICOS COMPRIMIDOS/ ← Salida comprimida
│
├── PSD/                ← Archivos fuente de diseño de interfaces
│   └── PSD Interfaces WAO.rar ← Archivos PSD de todas las interfaces (~12 MB)
│
└── README.md           ← Este archivo
```

---

## Componentes principales

### 1. Servidor (`server/`)

Servidor de juego escrito en VB6. Gestiona la lógica de juego completa: personajes, NPCs, combate, inventario, clanes, hechizos, mapas y comunicación de red.

- **Entry point**: `Sub Main` en `SERVER.VBP`
- **Ejecutable compilado**: `AOWinter.exe`
- **Puerto por defecto**: 7500 (TCP, configurable en `Server.ini`)
- **Máximo de usuarios**: 100 (configurable)
- **Versión del protocolo**: 2.0.1

### 2. Cliente DX7 (`Cliente/`)

Cliente original con motor DirectX 7. Es la versión de referencia del mod (versión 1.x).

- **Entry point**: `FrmLanzador` (lanzador integrado con control de actualizaciones)
- **Módulos clave**: `frmMain.frm` (HUD principal), `TCP.bas` (red), `TileEngine.bas` (motor gráfico), `clsTileEngineX.cls` (engine extendido)
- **Versión del proyecto VB6**: 2.0.0

### 3. Cliente DX8 (`WAO DX8/`)

Variante experimental del cliente con motor **DirectX 8**, liberada como versión 2.0 incompleta. Añade FPS libres y un motor gráfico actualizado (`clsDX8Engine.cls`).

- **Diferencias respecto al DX7**: Motor DX8 (`clsDX8Engine.cls`), módulos adicionales (`modDX8Fifo.bas`, `modDX8Requires.bas`), soporte de transparencia y anti-cheat mejorado.
- **Estado**: Liberado con bugs conocidos, no considerado estable por el autor.

### 4. Auto-Actualización (`Auto Update/`)

Launcher independiente que descarga actualizaciones del cliente antes de ejecutarlo. Usa `Unzip32.dll` para descomprimir paquetes y muestra una barra de progreso.

- **Dependencias**: `vbalProgBar6.ocx`, `Unzip32.dll`
- **Entry point**: `frmMain.frm`

### 5. Herramienta ZlibWAO (`ZlibWAO/`)

Utilidad de línea de comandos para comprimir y descomprimir los gráficos del cliente. Utiliza `zlib.dll` nativa. Los gráficos del juego se distribuyen comprimidos.

---

## Cómo funciona (arquitectura básica)

```
[Launcher / Auto Update]
        │
        ▼
[Cliente VB6 (DX7 o DX8)]
        │  TCP (puerto 7500)
        ▼
[Servidor VB6 (AOWinter.exe)]
        │
        ├── Dat/      ← Lee NPCs, objetos, hechizos, mapas
        ├── Maps/     ← Lee archivos de mapa binarios
        ├── guilds/   ← Lee/escribe datos de clanes
        └── charfile/ ← Lee/escribe personajes (no incluido en repo)
```

- La comunicación cliente-servidor es **TCP binario**, compatible con el protocolo AO 0.11.5.
- El servidor lee todos los datos del juego desde archivos `.dat` e `.ini` en la carpeta `Dat/`.
- No hay base de datos: la persistencia es 100% basada en archivos.
- Los gráficos del cliente se distribuyen comprimidos con ZlibWAO y se descomprimen en runtime o al instalar.

### Flujo del servidor

1. `Sub Main` → carga `Server.ini` y datos de `Dat/`
2. Inicializa timers y el sistema de red (Winsock / sockets custom `wsksock.bas`)
3. Acepta conexiones en el puerto 7500
4. Despacha paquetes en `TCP_HandleData1.bas` y `TCP_HandleData2.bas`
5. Corre lógica de juego mediante timers configurados en `Server.ini` (`[INTERVALOS]`)

### Flujo del cliente

1. `FrmLanzador` → verifica actualizaciones y lanza `frmConnect`
2. `frmConnect` → conexión al servidor y login
3. `frmMain` → HUD principal del juego (renderizado, input, red)
4. El motor gráfico (`TileEngine.bas` / `clsTileEngineX.cls`) renderiza los mapas en capas

---

## Instalación / Compilación / Ejecución

> ⚠️ Este proyecto requiere **Visual Basic 6** (IDE + runtime) y las dependencias de DirectX/OCX propias de la era Windows XP/2000. No compila con herramientas modernas sin adaptación.

### Servidor

1. Abrir `server/SERVER.VBP` en el IDE de VB6.
2. Compilar → genera `AOWinter.exe`.
3. Colocar el ejecutable en la carpeta `server/` junto a `Server.ini` y la carpeta `Dat/`.
4. Ejecutar `AOWinter.exe`. El servidor escucha en `127.0.0.1:7500` por defecto.
5. Para producción: cambiar `ServerIp` en `Server.ini` y configurar los GMs en `[Admines]` / `[IPGM]`.

**Dependencias del servidor**:
- `COMCTL32.OCX` (incluida en `server/`)
- `msinet.ocx` (incluida en `server/`)
- Runtime VB6 estándar

### Cliente (DX7)

1. Abrir `Cliente/Client.vbp` en el IDE de VB6.
2. Compilar → genera `Winter AO.exe`.
3. Colocar el ejecutable junto a los recursos gráficos y de audio del cliente.
4. **Cambiar la IP del servidor** en `CODIGO/Declares.bas` (constante `IpServidor`).

### Cliente (DX8)

1. Abrir `WAO DX8/Client.vbp` en VB6.
2. Compilar con soporte de DirectX 8 SDK.
3. Mismo proceso de configuración de IP que el DX7.

### Auto Update

1. Abrir `Auto Update/AutoUpdate.vbp` en VB6.
2. Compilar.
3. Configurar la URL de actualizaciones en el código (`ModGeneral.bas`).
4. El launcher descarga y descomprime paquetes `.zip` antes de lanzar el cliente.

---

## Datos del servidor (`server/Dat/`)

| Archivo | Contenido |
|---------|-----------|
| `NPCs.dat` | Definiciones de todos los NPCs del juego |
| `NPCs-HOSTILES.dat` | Tabla de NPCs hostiles |
| `obj.dat` | Definiciones de objetos (~196 KB, catálogo completo) |
| `Hechizos.dat` | Definiciones de hechizos (~31 KB) |
| `AreasStats.dat` | Estadísticas de áreas del mapa |
| `ArmadurasHerrero.dat` | Recetas de herrero (armaduras) |
| `ArmasHerrero.dat` | Recetas de herrero (armas) |
| `ObjCarpintero.dat` | Recetas de carpintero |
| `Ciudades.Dat` | Configuración de ciudades |
| `Head.dat` | Datos de cabezas de personaje |
| `Invokar.dat` | Tabla de invocaciones |
| `Map.dat` | Metadatos de mapas |
| `Hechizos.dat` | Definiciones de hechizos |
| `Motd.ini` | Mensaje del día del servidor |
| `Propagandas.ini` | Mensajes de propaganda/anuncios |
| `BanIps.dat` | IPs baneadas |
| `castillos.dat` | Estado de los castillos de clanes |
| `consultas.dat` | Consultas populares registradas |

---

## Estado del proyecto

| Componente | Estado |
|------------|--------|
| Servidor VB6 | ✅ Código fuente completo + binario compilado incluido |
| Cliente DX7 | ✅ Código fuente completo |
| Cliente DX8 | ⚠️ Código fuente completo pero con bugs conocidos (liberado sin terminar) |
| Auto Update | ✅ Código fuente completo |
| ZlibWAO | ✅ Herramienta compilada + fuentes (implícito por `.exe` incluido) |
| PSD de interfaces | ✅ Archivos fuente de diseño (.rar) |
| Datos del juego | ✅ Carpeta `Dat/` y `Maps/` incluidas |
| Personajes (`charfile/`) | ❌ No incluida (datos de runtime) |

> **Este es un proyecto de preservación histórica.** El autor lo liberó públicamente al dejar el desarrollo. No está en mantenimiento activo.

---

## Créditos y licencia

- **Autor principal**: Erwin (identificado en `Server.ini` y archivos de configuración)
- **Colaboradores mencionados**: Santo, Hennox, Stick, Mortis, Lectral (mencionados en changelog del README original)
- **Base**: Argentum Online 0.11.5 — MMORPG argentino desarrollado originalmente por Pablo Márquez (1999)
- **Copyright declarado en el proyecto**: `Copyright 2007 - Winter AO` (según `Client.vbp`)
- **Licencia**: No se incluye archivo de licencia formal en el repositorio. El código fue liberado públicamente por el autor para uso comunitario. Ver archivo `LICENSE` en la raíz del repositorio.

---

## Galería

![Captura de Winter-AO en juego](https://github.com/user-attachments/assets/452a45b1-14ad-4cb5-8031-8666bd95ed8b)

---

## Notas

- Los archivos `.frx` son recursos binarios de formularios VB6 (imágenes embebidas, etc.) y no son legibles directamente.
- El archivo `praetorians.bas` (~87 KB) es el módulo más grande del servidor y gestiona el sistema NPC cooperativo avanzado.
- La carpeta `guilds/` y `charfile/` son generadas en runtime y no se incluyen en el repositorio.
- Ver [`docs/notes.md`](docs/notes.md) para advertencias técnicas detalladas y deuda técnica conocida.
- Ver [`docs/components.md`](docs/components.md) para el detalle por módulo de cliente y servidor.
- Ver [`docs/build-and-run.md`](docs/build-and-run.md) para instrucciones de compilación extendidas.
