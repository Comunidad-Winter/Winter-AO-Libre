# Compilación y Ejecución — Winter-AO

> 📚 *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

> ⚠️ **Proyecto de preservación histórica.** Winter-AO es código VB6 de aproximadamente 2007. Requiere herramientas de la era Windows XP/2000. Los pasos descritos a continuación están inferidos del análisis de los archivos de proyecto `.vbp`, las dependencias declaradas y la estructura del repositorio.

---

## Requisitos generales

| Requisito | Versión mínima | Notas |
|-----------|---------------|-------|
| Visual Basic 6 IDE | SP6 (recomendado) | Solo disponible con licencia MSDN/VBA |
| Windows | XP / Vista / 7 (32-bit) | En Windows 10/11 puede requerir compatibilidad |
| DirectX 7 (cliente DX7) | DirectX 7 runtime | Incluido en Windows XP |
| DirectX 8 (cliente DX8) | DirectX 8 SDK + runtime | Necesario solo para `WAO DX8/` |
| VB6 Runtime | VB6SP6 runtime | Instalable en sistemas modernos |

### Dependencias OCX / DLL requeridas

Deben estar registradas en el sistema (`regsvr32`) o copiadas junto al ejecutable:

**Servidor**:
- `COMCTL32.OCX` — incluida en `server/`
- `msinet.ocx` — incluida en `server/`
- `VB6 Runtime` — instalable por separado

**Cliente DX7**:
- `DX7VB.DLL` — DirectX 7 for VB (en `system32` de Windows XP)
- `RICHTX32.OCX` — RichTextBox control
- `CSWSK32.OCX` — Socket custom
- `MSINET.OCX` — Internet Transfer Control
- `MSWINSCK.OCX` — Winsock control
- `ieframe.dll` — Internet Explorer (para el control Web integrado)

**Cliente DX8** (añadidos sobre DX7):
- DirectX 8 SDK headers/libs para compilar

**Auto Update**:
- `vbalProgBar6.ocx` — incluida en `Auto Update/`
- `Unzip32.dll` — incluida en `Auto Update/`

---

## Compilar y ejecutar el servidor

### 1. Compilación con VB6 IDE

```
1. Abrir Visual Basic 6 IDE
2. File → Open Project → server/SERVER.VBP
3. Verificar que COMCTL32.OCX y msinet.ocx estén registrados
4. File → Make AOWinter.exe
5. Guardar en server/AOWinter.exe (ya existe un binario compilado)
```

> El repositorio ya incluye `server/AOWinter.exe` compilado. Si no se modifica el código, se puede usar directamente.

### 2. Configuración antes de ejecutar

Editar `server/Server.ini`:

```ini
[INIT]
ServerIp=127.0.0.1        ; Cambiar por IP pública para servidor en producción
StartPort=7500             ; Puerto TCP de escucha
MaxUsers=100               ; Máximo de conexiones simultáneas
Version=2.0.1              ; Versión del protocolo (debe coincidir con el cliente)
PuedeCrearPersonajes=1     ; 1=permite crear personajes nuevos

[Admines]
Admin1=NombreGM            ; Reemplazar "Erwin" por el nombre del GM admin

[IPGM]
NombreGMIp=xxx.xxx.xxx.xxx ; IP del GM (sistema de seguridad por IP)
```

### 3. Ejecución

```
1. Ejecutar server/AOWinter.exe
2. El servidor mostrará frmMain (consola) y frmServidor (panel de control)
3. Estado del servidor visible en frmEstadisticas
4. Para detener: cerrar la aplicación o usar el panel de control
```

### 4. Directorios requeridos en runtime

El servidor espera estas carpetas relativas a su ubicación. **Crearlas si no existen**:

```
server/
├── Dat/           ← ✅ Incluida en el repo
├── Maps/          ← ✅ Incluida en el repo
├── guilds/        ← ✅ Incluida en el repo (vacía)
├── charfile/      ← ❌ NO incluida — debe crearse manualmente
├── logs/          ← ❌ NO incluida — se crea automáticamente (inferido)
├── Torneos/       ← ✅ Incluida
└── SUGS/          ← ✅ Incluida
```

> **Crítico**: La carpeta `charfile/` debe existir antes de ejecutar el servidor, de lo contrario fallará al intentar guardar personajes. Crear manualmente como directorio vacío.

---

## Compilar y ejecutar el cliente DX7

### 1. Compilación con VB6 IDE

```
1. Abrir Visual Basic 6 IDE
2. File → Open Project → Cliente/Client.vbp
3. Verificar que todas las referencias OCX estén registradas:
   - DX7VB.DLL (DirectX 7 for VB)
   - RICHTX32.OCX
   - CSWSK32.OCX
   - MSINET.OCX
   - MSWINSCK.OCX
4. File → Make "Winter AO.exe"
```

### 2. Configurar la IP del servidor

Antes de compilar, editar `Cliente/CODIGO/Declares.bas`:

```vb
Public Const IpServidor As String = "127.0.0.1"  ' Cambiar por IP del servidor
```

> El README original advierte que si no se cambia también la IP de comprobación de actualizaciones, el cliente dará errores. Esta constante se encuentra en `Declares.bas` o `Mod_WAO.bas`.

### 3. Estructura de recursos del cliente

El cliente necesita sus recursos gráficos y de audio junto al ejecutable. La distribución exacta no está en el repositorio (solo el código fuente). Estructura esperada (inferida del código):

```
[Directorio del cliente]/
├── Winter AO.exe          ← Ejecutable compilado
├── Graficos/              ← Gráficos descomprimidos (o comprimidos vía ZlibWAO)
├── Musica/                ← Archivos de música (.mid, .mp3)
├── Sonidos/               ← Efectos de sonido (.wav)
├── Interfaces/            ← Recursos de interfaz gráfica
└── Mapas/                 ← Archivos de mapa del cliente
```

### 4. Ejecución

```
1. Ejecutar "Winter AO.exe"
2. Aparece FrmLanzador (el launcher interno)
3. El lanzador verifica actualizaciones si está configurado
4. Al hacer clic en Jugar: aparece frmConnect
5. Introducir IP del servidor, usuario y contraseña
6. Si el servidor está activo: se accede al juego (frmMain)
```

---

## Compilar y ejecutar el cliente DX8

El proceso es idéntico al DX7 con las siguientes diferencias:

```
1. Abrir WAO DX8/Client.vbp en VB6
2. Requiere DirectX 8 SDK instalado para resolver referencias de clsDX8Engine.cls
3. Compilar → genera el ejecutable del cliente DX8
4. Mismo proceso de configuración de IP en Declares.bas
```

> **Estado**: El autor liberó esta versión con bugs conocidos (ver `WAO DX8/Leeme Importante !!!.txt`). No se garantiza estabilidad.

---

## Compilar el launcher de Auto Update

```
1. Abrir Auto Update/AutoUpdate.vbp en VB6
2. Verificar que vbalProgBar6.ocx esté registrado
3. File → Make AutoUpdate.exe
4. Configurar la URL de actualizaciones en ModGeneral.bas
5. Colocar junto al ejecutable del cliente
```

---

## Usar ZlibWAO (compresión de gráficos)

```
1. Colocar los gráficos originales en ZlibWAO/GRAFICOS/
2. Ejecutar ZlibWAO/ZlibWao.exe
3. Los gráficos comprimidos aparecen en ZlibWAO/GRAFICOS COMPRIMIDOS/
4. Distribuir los gráficos comprimidos con el cliente
```

El cliente usa `modCompression.bas` para descomprimir en runtime automáticamente.

---

## Seguridad del servidor: sistema de IP por GM

El servidor incluye un sistema de restricción de IP para cuentas de Game Master (`SecurityIp.bas`):

```
1. Editar Server.ini, sección [IPGM]:
   NombreGMIp=xxx.xxx.xxx.xxx

2. El nombre de la variable debe coincidir con el nombre del GM + "Ip"
   Ejemplo: Admin1=Erwin → ErwinIp=192.168.1.100

3. Si el GM intenta conectar desde otra IP, el servidor rechaza la conexión.
```

> El README original dice: *"buscan ERWIN y sustituyen en nombre de Erwin por el nombre del GM"*. Esto implica buscar referencias a `Erwin` en el código del servidor si se desea cambiar el admin por defecto.

---

## Verificación MD5 del cliente

El servidor puede verificar la integridad del cliente mediante hashes MD5 (`Server.ini`, sección `[MD5Hush]`):

```ini
[MD5Hush]
Activado=0                                ; 0=desactivado, 1=activado
MD5Aceptados=2
Md5Aceptado1=2024E4FDE84DBD6765D2C1B4FC8D9C70   ; Hash del cliente normal
Md5Aceptado2=22EF81267FBC05B17F8948DF49F47D96   ; Hash del cliente AlphaBlending
```

Si se recompila el cliente, el hash cambiará y habrá que actualizar estos valores (o desactivar la verificación con `Activado=0`).

---

## Configuración de torneos

Los torneos se configuran mediante archivos `.ini` en `server/Torneos/`. El servidor los lee para gestionar eventos programados. El formato exacto no ha sido verificado en este análisis — se debe inspeccionar el módulo `Mod_WAO.bas` del servidor para los detalles.

---

## Problemas conocidos al compilar/ejecutar

| Problema | Causa probable | Solución |
|----------|---------------|----------|
| Error "componente no registrado" | OCX no registrada | `regsvr32 nombre.ocx` como administrador |
| Error al arrancar el servidor | Falta carpeta `charfile/` | Crear la carpeta manualmente |
| Cliente no conecta | IP incorrecta en `Declares.bas` | Editar y recompilar |
| Error de DX7/DX8 | DirectX no instalado o versión incorrecta | Instalar DirectX End-User Runtime |
| Hash MD5 rechazado | Cliente recompilado, hash desactualizado | Desactivar MD5 en `Server.ini` o actualizar hashes |
| NPC de duelo no funciona | Bug conocido (reportado en README original) | Sin solución documentada |
| Minimapa cortado | Bug conocido (reportado en README original) | Sin solución documentada |
| Inestabilidad del motor gráfico | Bug conocido en DX8 (reportado por el autor) | Usar versión DX7 para mayor estabilidad |
