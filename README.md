# Winter-AO — Argentum Online (Motor VB6)

> **Esta documentación fue generada por y para Comunidad-Winter con el objetivo de preservar recursos de Argentum Online.**

Winter-AO es una modificación integral del MMORPG Argentum Online, basada en el motor clásico v0.12.x, desarrollada en Visual Basic 6 con renderizado DirectX 8.

---

## Qué es

**Winter AO Ultimate** es un servidor y cliente de Argentum Online creado por el equipo **LwK-Projects** (Lorwik, MaxTus, Hennox, entre otros). El proyecto amplía el motor original con:

- **Sistema de Cuentas** unificado (reemplaza el login directo por personaje del AO clásico).
- **Empaquetado de recursos** en formato propietario `.WAO` (compresión zlib).
- **Motor gráfico DirectX 8** con soporte de partículas, iluminación dinámica y clima.
- **Sistemas de juego automatizados**: torneos PvP, misiones (quests), centinela anti-macros.
- **Interfaz gráfica personalizada** con skins vía formularios VB6 estilizados.

### Contexto histórico

Este es un **proyecto de preservación**. Winter-AO fue un servidor activo de la comunidad hispanohablante de Argentum Online durante la primera mitad de la década de 2010. El código refleja las prácticas de desarrollo de esa era: estado global mutable, archivos de texto plano como persistencia, y dependencias COM de Windows de 32 bits. No es un proyecto activo de desarrollo.

---

## Estructura del Repositorio

```
Winter-AO/
├── Cliente/              ← Cliente del juego (VB6 + DirectX 8)
│   ├── Client.vbp        ← Proyecto VB6 principal del cliente
│   ├── CODIGO/            ← 77 archivos fuente (.bas, .cls, .frm) ~1.4 MB
│   ├── Recursos/          ← Assets empaquetados en formato .WAO (~127 MB)
│   ├── Librerias/         ← DLLs y OCXs requeridos (DirectX 8, Winsock, zlib)
│   └── Init/              ← Configuración local (BindKeys.bin, Config.cfg)
│
├── server/               ← Servidor del juego (VB6)
│   ├── SERVER.VBP         ← Proyecto VB6 principal del servidor
│   ├── Codigo/            ← 73 archivos fuente (.bas, .cls, .frm) ~2 MB
│   ├── Dat/               ← Datos del mundo (objetos, NPCs, hechizos, quests)
│   ├── Maps/              ← 172 mapas binarios del mundo
│   ├── Charfile/          ← Archivos de personajes (generados en runtime)
│   ├── Cuentas/           ← Archivos de cuentas de usuario
│   ├── guilds/            ← Datos de clanes
│   └── Server.ini         ← Configuración del servidor
│
├── Launcher/             ← Launcher con actualizador (VB6)
├── WorldEditor/          ← Editor de mapas con motor DX8 (VB6)
├── Editor de Particulas/ ← Editor visual de efectos de partículas (VB6)
├── Indexador/            ← Generador de índices de gráficos/animaciones (VB6)
├── LwK-Universal/        ← Compresor/descompresor de archivos .WAO (VB6)
├── PathHelper/           ← Utilidad de diagnóstico de rutas de parcheo (VB6)
├── Registrador de librerias/ ← Registrador de componentes COM en Windows (VB6)
├── Interfaces/           ← Archivos fuente de diseño (.psd) de la UI (~61 archivos)
├── LICENSE               ← GNU General Public License v3
└── docs/                 ← Documentación técnica detallada
```

---

## Componentes Principales

| Componente | Proyecto VB6 | Ejecutable | Descripción |
|-----------|-------------|-----------|-------------|
| **Cliente** | `Client.vbp` | `Winter AO Ultimate.exe` | Motor de juego con DX8, UI completa, protocolo de red |
| **Servidor** | `SERVER.VBP` | `Winter-AO Server.exe` | Lógica de mundo, combate, NPCs, persistencia |
| **Launcher** | `Launcher.vbp` | `WinterAO Ultimate Launcher.exe` | Actualizador con contenido Flash (SWF) |
| **WorldEditor** | `WorldEditor.vbp` | `WorldEditorDX8.exe` | Editor visual de mapas con DX8 |
| **Editor de Partículas** | `Proyecto1.vbp` | `Particle Editor - RincondelAO.exe` | Diseño de efectos visuales |
| **Indexador** | `Proyecto1.vbp` | `Indexador RincondelAO.exe` | Generación de índices de gráficos (Grh) |
| **LwK-Universal** | `LwKUniversal.vbp` | `Comprensor.Descomprensor LwK Universal.exe` | Empaquetado/desempaquetado de archivos `.WAO` |
| **PathHelper** | `Proyecto1.vbp` | `PathHelper.exe` | Diagnóstico de rutas y parcheo |
| **Registrador** | `Registrador de Librerias.vbp` | `Registrador de librerias.exe` | Registro de OCX/DLL en el sistema |

---

## Cómo Funciona (Arquitectura Básica)

```
┌─────────────────┐        TCP/IP (puerto 7666)        ┌─────────────────┐
│    CLIENTE       │◄──────────────────────────────────►│    SERVIDOR     │
│                  │     Protocolo binario (1 byte ID)  │                 │
│  DX8 Renderer    │                                    │  Game Logic     │
│  TileEngine      │                                    │  AI / NPCs      │
│  UI (VB6 Forms)  │                                    │  Persistencia   │
│  Audio (DSound)  │                                    │  (.dat/.ini)    │
│  Assets (.WAO)   │                                    │  172 Mapas      │
└─────────────────┘                                    └─────────────────┘
```

- **Red**: Sockets TCP asíncronos vía controles Winsock (`CSWSK32.OCX` / `MSWINSCK.OCX`). Protocolo binario con ID de paquete de 1 byte. Serialización/deserialización en `Protocol.bas` (ambos lados).
- **Persistencia**: Todo en archivos planos. Objetos, NPCs y hechizos en `Dat/*.dat`. Personajes en `Charfile/`. Cuentas en `Cuentas/`. Mapas en binario en `Maps/`.
- **Game Loop**: No hay un bucle `while(true)`. El servidor es event-driven: timers de Windows + eventos de socket controlan el tick del mundo.
- **Renderizado**: Motor 2D isométrico (`TileEngine.bas`) sobre Direct3D 8 (`dx8vb.dll`). Texturas cargadas bajo demanda desde archivos `.WAO` comprimidos.

---

## Instalación / Compilación / Ejecución

> Para instrucciones detalladas, ver [docs/build-and-run.md](docs/build-and-run.md).

### Requisitos

- **Windows** (32 bits o 64 bits con modo de compatibilidad).
- **Microsoft Visual Basic 6.0** (SP6) para compilar desde fuente.
- Componentes COM registrados (DirectX 8, Winsock, zlib). Ver `Registrador de librerias/`.

### Pasos rápidos

1. Registrar librerías COM ejecutando `Registrar Librerias.exe` como Administrador.
2. **Servidor**: Abrir `server/SERVER.VBP` → `File → Make Winter-AO Server.exe`. Configurar `Server.ini`.
3. **Cliente**: Abrir `Cliente/Client.vbp` → `File → Make Winter AO Ultimate.exe`. Requiere `Recursos/*.WAO`.

---

## Estado del Proyecto

| Aspecto | Estado |
|---------|--------|
| **Propósito** | 🏛️ Preservación histórica |
| **Desarrollo activo** | ❌ No |
| **Compilable** | ⚠️ Requiere VB6 + librerías COM de 32 bits |
| **Completitud** | ✅ Cliente y servidor funcionales con sistemas de juego completos |
| **Calidad del código** | ⚠️ Deuda técnica típica de mods VB6 de la época |

---

## Créditos y Licencia

> **Esta documentación fue generada por y para Comunidad-Winter con el objetivo de preservar recursos de Argentum Online.**

### Autoría verificada en el código

| Entidad | Rol | Fuente |
|---------|-----|--------|
| **Pablo Ignacio Márquez (Morgolock)** | Creador original de Argentum Online | Cabeceras AGPL en módulos del servidor |
| **LwK-Projects / Lorwik** | Desarrollo de Winter-AO, WorldEditor, herramientas | Metadatos de todos los `.vbp` |
| **MaxTus, Hennox, Sigfrido, Harry, Tefo, Hioryz** | Staff del servidor | `Server.ini` (secciones Admines/Dioses/SemiDioses) |
| **RincondelAO.com.ar** | Comunidad de origen de herramientas | Metadatos de WorldEditor, Indexador, Particle Editor |

### Licencia

- **Archivo `LICENSE`**: GNU General Public License v3 (GPLv3).
- **Cabeceras del código servidor**: Múltiples módulos heredados contienen encabezados AGPL (Affero GPL) referenciando el código original de Argentum Online.
- **Nota**: Existe una discrepancia entre el archivo `LICENSE` (GPLv3) y las cabeceras del código (AGPL). Se preserva tal cual se encontró en el repositorio.

---

## Notas

- Para documentación técnica detallada, consultar la carpeta [`docs/`](docs/):
  - [Arquitectura](docs/architecture.md) — Visión de alto nivel del sistema.
  - [Componentes](docs/components.md) — Desglose módulo por módulo.
  - [Compilación y ejecución](docs/build-and-run.md) — Guía paso a paso.
  - [Notas técnicas](docs/notes.md) — Advertencias, deuda técnica y elementos no verificados.
- El código contiene hardcoding extensivo, variables globales y prácticas propias del desarrollo comunitario de Argentum Online circa 2010-2015.
- Los binarios precompilados (`.exe`) incluidos en el repositorio no han sido verificados respecto al código fuente.
