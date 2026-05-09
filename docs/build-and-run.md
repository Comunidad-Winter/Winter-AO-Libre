# Compilación y Ejecución de Winter-AO

> **Esta documentación fue generada por y para Comunidad-Winter con el objetivo de preservar recursos de Argentum Online.**

Guía paso a paso para compilar y ejecutar el proyecto, deducida del análisis de los archivos `.vbp`, dependencias COM y estructura de carpetas.

---

## 1. Requisitos del Sistema

### Para compilar desde fuente

| Requisito | Detalle |
|-----------|---------|
| **SO** | Windows (32 bits, o 64 bits con compatibilidad) |
| **IDE** | Microsoft Visual Basic 6.0 SP6 (Enterprise o Professional) |
| **DirectX** | DirectX 8 SDK con `dx8vb.dll` registrada en `SysWOW64` |
| **Privilegios** | Administrador (para registrar componentes COM) |

### Para ejecutar binarios precompilados

| Requisito | Detalle |
|-----------|---------|
| **SO** | Windows XP / 7 / 8 / 10 / 11 (32 bits o con WoW64) |
| **Librerías** | Componentes COM registrados (ver sección 2) |
| **Assets** | Carpeta `Recursos/` con archivos `.WAO` (solo cliente) |

> ⚠️ En Windows 10/11 de 64 bits, el IDE VB6 puede requerir modo de compatibilidad y ejecución como Administrador.

---

## 2. Registro de Librerías COM

El repositorio incluye la herramienta `Registrador de librerias/` para automatizar este paso. Alternativamente, registrar manualmente:

### Librerías del cliente (verificadas en `Client.vbp`)

```batch
:: Ejecutar como Administrador en cmd de 32 bits
:: Los archivos están en Cliente/Librerias/

C:\Windows\SysWOW64\regsvr32.exe CSWSK32.OCX
C:\Windows\SysWOW64\regsvr32.exe MSWINSCK.OCX
C:\Windows\SysWOW64\regsvr32.exe RICHTX32.OCX
C:\Windows\SysWOW64\regsvr32.exe MSCOMCTL.OCX
C:\Windows\SysWOW64\regsvr32.exe MSINET.OCX
C:\Windows\SysWOW64\regsvr32.exe COMDLG32.OCX
C:\Windows\SysWOW64\regsvr32.exe vbalProgBar6.ocx
C:\Windows\SysWOW64\regsvr32.exe dx8vb.dll
C:\Windows\SysWOW64\regsvr32.exe quartz.dll
```

> **Nota**: `zlib.dll` NO se registra con regsvr32; se carga dinámicamente. Debe estar en la misma carpeta que el ejecutable.

### Librerías del servidor (verificadas en `SERVER.VBP`)

El servidor requiere menos dependencias COM. Solo necesita `MSCOMCTL.OCX` de controles externos. El resto son módulos estándar de VB6.

---

## 3. Compilación del Servidor

1. Abrir `server/SERVER.VBP` con Visual Basic 6.
2. Menú `File → Make Winter-AO Server.exe`.
3. El ejecutable se genera en la misma carpeta.

### Estructura de carpetas requerida para ejecución

```
server/
├── Winter-AO Server.exe    ← Ejecutable compilado
├── Server.ini              ← Configuración (puerto, versión, staff, intervalos)
├── Dat/                    ← OBLIGATORIO: obj.dat, NPCs.dat, Hechizos.dat, etc.
├── Maps/                   ← OBLIGATORIO: 172 archivos .map
├── Charfile/               ← Se crea en runtime (personajes)
├── Cuentas/                ← Se crea en runtime (cuentas)
├── guilds/                 ← Se crea en runtime (clanes)
├── Logs/                   ← Se crea en runtime (logs)
└── Reportes/               ← Se crea en runtime (reportes)
```

### Configuración básica (`Server.ini`)

Parámetros clave verificados en el archivo:

```ini
[INIT]
StartPort=7666              ; Puerto TCP del servidor
MaxUsers=550                ; Máximo de conexiones simultáneas
Version=4.0.2               ; Versión esperada del cliente
PuedeCrearPersonajes=1      ; Habilitar creación de personajes
AllowMultiLogins=1           ; Permitir múltiples sesiones
Encriptar=1                  ; Activar encriptación del protocolo
```

---

## 4. Compilación del Cliente

1. Abrir `Cliente/Client.vbp` con Visual Basic 6.
2. Verificar que todas las referencias y controles se resuelven correctamente (menú `Project → References` y `Project → Components`).
3. Menú `File → Make Winter AO Ultimate.exe`.

### Estructura de carpetas requerida para ejecución

```
Cliente/
├── Winter AO Ultimate.exe  ← Ejecutable compilado
├── zlib.dll                ← OBLIGATORIO: descompresión de .WAO
├── Init/                   ← Configuración local (BindKeys, Config)
└── Recursos/               ← OBLIGATORIO: assets empaquetados
    ├── Graphics.WAO         ← Sprites (~36 MB)
    ├── Interface.WAO        ← Texturas UI (~7 MB)
    ├── Maps.WAO             ← Mapas (~458 KB)
    ├── Musics.WAO           ← Música (~31 MB)
    ├── Sounds.WAO           ← Efectos de sonido (~48 MB)
    └── Scripts.WAO          ← Datos/índices (~198 KB)
```

> ⚠️ **Sin los archivos `.WAO` y `zlib.dll`, el cliente lanzará errores fatales** (Automation Error) antes de mostrar la pantalla de carga.

### Conexión al servidor

La IP/puerto del servidor se configura desde el formulario de conexión (`frmConnect.frm`) en runtime, o puede estar hardcodeada en `General.bas` o `Declares.bas` según la versión.

Para pruebas locales, usar `127.0.0.1` con el puerto `7666`.

---

## 5. Compilación de Herramientas

### WorldEditor

1. Abrir `WorldEditor/WorldEditor.vbp`.
2. Requiere `COMDLG32.OCX`, `dx8vb.dll`, `progressbar-xp.ocx` registrados.
3. Compilar: `File → Make WorldEditorDX8.exe`.
4. Necesita la carpeta `Init/` con índices de gráficos y `Maps/` para funcionar.

### Editor de Partículas

1. Abrir `Editor de Particulas/Proyecto1.vbp`.
2. Requiere `dx8vb.dll`, `RICHTX32.OCX`, `MSCOMCTL.OCX`, `COMDLG32.OCX`.
3. Compilar: `File → Make Particle Editor - RincondelAO.exe`.

### Indexador

1. Abrir `Indexador/Proyecto1.vbp`.
2. Sin dependencias COM externas adicionales.
3. Compilar: `File → Make Indexador RincondelAO.exe`.
4. Necesita la carpeta `Graficos/` con los BMP/PNG de sprites.

### LwK-Universal (Compresor .WAO)

1. Abrir `LwK-Universal/LwKUniversal.vbp`.
2. Requiere `vbalProgBar6.ocx` registrado.
3. Compilar: `File → Make Comprensor.Descomprensor LwK Universal.exe`.
4. Necesita `zlib.dll` en la misma carpeta para funcionar.

---

## 6. Ejecución Paso a Paso (Prueba Local)

```
1. Registrar librerías COM (una sola vez)
2. Ejecutar el servidor:
   → Doble clic en Winter-AO Server.exe
   → Esperar a que cargue mapas, NPCs y objetos
   → La ventana muestra "Servidor iniciado"

3. Ejecutar el cliente:
   → Doble clic en Winter AO Ultimate.exe
   → En el formulario de conexión, ingresar: 127.0.0.1 : 7666
   → Crear cuenta → Crear personaje → Jugar

4. Para obtener privilegios de administrador:
   → Crear un personaje normalmente
   → Cerrar el servidor
   → Editar server/Cuentas/<cuenta>/... y server/Server.ini
   → Agregar el nombre del personaje en [Admines]
   → Reiniciar el servidor
```

---

## 7. Notas sobre Debugging en VB6

- **No usar el botón "Stop" del IDE** mientras el cliente está corriendo. DirectX 8 mantiene control exclusivo del dispositivo gráfico; una interrupción abrupta puede dejar la pantalla del SO en un estado corrupto (resolución incorrecta, pantalla negra).
- Usar el comando in-game `/SALIR` o cerrar la ventana del cliente normalmente para que se ejecuten las rutinas de limpieza (`Class_Terminate` de `clsDX8Engine`).
- El servidor puede ejecutarse de forma más segura desde el IDE ya que no usa DirectX.
- Para depurar el protocolo, inspeccionar los `Debug.Print` esparcidos en `Protocol.bas` de ambos lados.
