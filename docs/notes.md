# Notas técnicas — Winter-AO

> 📚 *Esta documentación fue generada por y para **Comunidad-Winter** con el objetivo de preservar recursos de Argentum Online.*

---

> ℹ️ Este documento recoge advertencias técnicas, deuda de código, partes no verificadas y observaciones del análisis del repositorio. Todo lo marcado como **[inferido]** o **[no verificado]** no pudo ser confirmado directamente en el código durante este análisis.

---

## Estado general del proyecto

Este repositorio es un **proyecto de preservación histórica**. El autor (Erwin) lo liberó públicamente al dejar el desarrollo activo. El código tiene características típicas de proyectos VB6 de su época:

- Módulos muy grandes sin separación de responsabilidades clara (TCP.bas con ~210 KB, clsTileEngineX.cls con ~195 KB).
- Mezcla de lógica de presentación y lógica de negocio en formularios VB6.
- Sin tests automatizados.
- Sin sistema de gestión de dependencias (las DLLs y OCX se referencian directamente por ruta del sistema).
- Persistencia 100% basada en archivos planos.
- Comentarios escasos o inexistentes en la mayor parte del código.

---

## Bugs conocidos (declarados por el autor)

Los siguientes bugs fueron declarados en el README original del proyecto y **no han sido corregidos** en el código del repositorio:

| Bug | Componente | Estado |
|-----|-----------|--------|
| NPC de duelo no funciona | Servidor + cliente | Sin fix |
| Minimapa cortado por la mitad en algunos mapas | Cliente | Sin fix |
| Bug en el casco de Hierro | Cliente (gráficos) | Sin fix |
| Banderas de la Armada bugeadas | Cliente (gráficos) | Sin fix |
| Inestabilidad del motor gráfico DX8 | Cliente WAO DX8 | Sin fix — el autor liberó la versión sin terminar |

---

## Funcionalidades eliminadas o no funcionales

El README original del proyecto lista explícitamente funcionalidades que fueron **eliminadas o no funcionan**:

| Funcionalidad | Estado declarado |
|--------------|-----------------|
| Atributos Asignables | Eliminado (provocaba bug que no dejaba crear personajes) |
| Control de Volumen de música desde opciones | Eliminado |
| Anti-Doble cliente | Eliminado |
| Sistema de Subasta | Eliminado |
| Comando `/PV NICK Mensaje` (chat privado) | **No funciona** |

---

## Observaciones técnicas por componente

### Servidor

1. **División del protocolo en dos módulos**: `TCP_HandleData1.bas` y `TCP_HandleData2.bas` suman ~183 KB de código de manejo de paquetes. Esta división es una workaround a la limitación de tamaño de módulo de VB6, no un diseño intencional.

2. **`praetorians.bas` (~87 KB)**: Módulo de IA de NPCs cooperativos muy extenso. Su complejidad y tamaño lo hacen propenso a bugs. El nombre "Praetorians" sugiere que gestiona NPCs especiales tipo guardianes o élites.

3. **Seguridad por IP de GM** (`SecurityIp.bas`): Sistema simple que verifica la IP de la conexión del GM. Fácilmente bypasseable si el atacante controla su IP. La contraseña de GM no es suficiente sin esta verificación activada.

4. **`Encriptar=1` en `Server.ini`**: El servidor tiene cifrado del protocolo activado por defecto (`CrcSubKey=12345`). El mecanismo exacto de cifrado no ha sido verificado en este análisis — se infiere que es un XOR o cifrado por byte con la clave como semilla. **[inferido]**

5. **`AllowMultiLogins=1`**: El servidor permite múltiples logins simultáneos por defecto. Esto puede ser un vector de abuso si se deja en producción sin revisar.

6. **`charfile/` no incluida**: Los archivos de personaje (`charfile/*.chr`) son datos de runtime y no están en el repositorio. El servidor fallará al intentar guardar personajes si esta carpeta no existe.

7. **Facciones hardcodeadas**: Los IDs de armaduras imperiales y del caos están hardcodeados en `Server.ini` con valores numéricos (ej.: `ArmaduraImperial1=370`). Si los datos de `obj.dat` cambian, hay que actualizar estos valores manualmente.

8. **Puerto de estadísticas**: El servidor expone un segundo puerto (`StartPortEstadisticas=7669`) para consultas de estadísticas. El módulo `clsEstadisticasIPC.cls` maneja esto. **[no verificado en profundidad]**

### Cliente DX7

1. **`clsTileEngineX.cls` (~195 KB)**: El motor de renderizado más grande del cliente. Contiene lógica mezclada de renderizado, física simple y lógica de juego. Es el módulo más crítico y más difícil de mantener.

2. **`TCP.bas` del cliente (~65 KB)**: Implementa el protocolo completo de comunicación AO. Cualquier cambio en el protocolo del servidor requiere cambios correspondientes aquí.

3. **`modCompression.bas` (~72 KB)**: Módulo de descompresión de gráficos. El formato de compresión es propio del proyecto (basado en zlib). Los gráficos del cliente deben distribuirse comprimidos.

4. **`Form2.frm` (~17 KB, recursos binarios ~662 KB)**: Formulario de selección de personaje tras el login. Los `.frx` son recursos binarios (imágenes embebidas) que no pueden modificarse sin el IDE de VB6.

5. **IP del servidor hardcodeada**: La IP del servidor está definida como constante en `Declares.bas`. Para cambiarla hay que recompilar el cliente. El README original confirma esto.

6. **Sistema de comprobación de actualizaciones**: El cliente usa `MSINET.OCX` para verificar actualizaciones en una URL. Si la URL ya no existe, el cliente generará errores en `frmConnect`. El README original advierte sobre esto.

7. **`clsMP3Player.cls`**: Usa Windows Media Player como backend. En sistemas sin WMP instalado puede fallar.

### Cliente DX8

1. **Versión liberada sin terminar**: El autor liberó explícitamente la versión DX8 como beta no finalizada. No debe usarse en producción.

2. **`clsDX8Engine.cls` (~94 KB)**: Motor DX8 nuevo. Al no haber terminado la versión, pueden existir funcionalidades del DX7 no migradas.

3. **`msn.bas`**: Módulo de integración con MSN Messenger para mostrar el estado del jugador. MSN Messenger ya no existe. Este módulo es no funcional en entornos modernos.

4. **`frmRadio.log`**: Archivo de log de errores del formulario de radio. Indica que el componente de radio tuvo errores de compilación o runtime registrados. **[no verificado]**

### Auto Update

1. **URL de actualización no documentada**: La URL del servidor de actualizaciones estaba en `ModGeneral.bas` pero no ha sido verificada en este análisis. Es probable que el servidor de actualizaciones original ya no exista.

2. **`Unzip32.dll`**: Biblioteca ZIP de 32 bits. Compatible con Windows XP–10 en modo 32 bits. Puede dar problemas en Windows 11 de 64 bits sin modo de compatibilidad.

### ZlibWAO

1. **Sin código fuente visible del ejecutable**: `ZlibWao.exe` está incluido como binario compilado pero no se encuentra su `.vbp` de origen en la carpeta. El código fuente del compresor podría haberse perdido o no haberse incluido en el repositorio.

2. **`zlib.dll` de 32 bits**: Compatible con VB6 y Windows de 32 bits. En Windows 64 bits funciona si el proceso es de 32 bits.

---

## Partes no verificadas en este análisis

Las siguientes partes del repositorio no fueron analizadas en profundidad durante la generación de esta documentación:

| Componente | Razón |
|------------|-------|
| Contenido exacto de `Maps/` | Binarios, no inspeccionables con texto |
| Formato binario de archivos `.map` | No documentado internamente |
| Mecanismo exacto de cifrado del protocolo | Requeriría análisis de `TCP.bas` completo |
| `modQuests.bas` — sistema de quests | No verificado si está funcional o es un stub |
| `Mod_Guerras.bas` — guerras de facciones | No verificado su estado funcional |
| Archivos `guilds/` | Carpeta existe pero sin archivos de ejemplo |
| `Bugs/` en `server/` | Carpeta con bugs registrados, no analizada |
| `frmConsolaTorneoUS.frm` | Versión de usuario de la consola de torneos, no analizada |

---

## Deuda técnica heredada del AO base

Winter-AO hereda la deuda técnica del AO 0.11.5 original:

- **Sin separación cliente/servidor clara en datos**: El servidor y el cliente tienen sus propios archivos de datos independientes que deben mantenerse sincronizados manualmente.
- **Protocolo binario frágil**: Sin versioning, sin comprobación de integridad por paquete (solo CRC global opcional). Un cambio en el protocolo rompe compatibilidad.
- **Arquitectura monohilo en servidor**: VB6 no soporta multihilo real. En picos de carga o con lógica lenta, el servidor puede quedar bloqueado.
- **Sin logging estructurado**: Los logs son texto plano en archivos o mensajes en la consola del servidor, sin niveles de log diferenciados.

---

## Notas sobre el contexto histórico

- El proyecto data de aproximadamente **2007** (según el copyright declarado en `Client.vbp`).
- El servidor VB6 original de Argentum Online (`v0.11.5`) es la base. Winter-AO añade extensiones sobre esa base.
- VB6 dejó de tener soporte oficial de Microsoft en 2008. Las herramientas para compilar este código son difíciles de obtener legalmente hoy.
- El binario `server/AOWinter.exe` incluido en el repositorio es la versión compilada original y puede usarse directamente en Windows XP/7/10 (32 bits) sin necesidad de recompilar.
- Los archivos `*.log` en los directorios de código (ej. `frmMain.log`, `frmConnect.log`) son logs de errores del IDE de VB6, no logs de ejecución. Registran errores de compilación históricos y no son relevantes para el uso actual.
- Los archivos `FOXUSER.DBF` y `FOXUSER.FPT` son artefactos de FoxPro generados por algunas herramientas de la época. No tienen función en el proyecto VB6 y pueden ignorarse.
- Los archivos `MSSCCPRJ.SCC` son archivos de control de versiones de SourceSafe (sistema de VCS de Microsoft). Indican que el proyecto fue desarrollado con Visual SourceSafe.
