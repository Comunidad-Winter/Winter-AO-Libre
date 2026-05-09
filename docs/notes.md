# Notas Técnicas y Advertencias

> **Esta documentación fue generada por y para Comunidad-Winter con el objetivo de preservar recursos de Argentum Online.**

Observaciones, inconsistencias, elementos no verificados y deuda técnica detectados durante el análisis del repositorio. Todo marcado según su nivel de certeza.

---

## 1. Contradicciones de Versiones

**Estado: Verificado** — Existen tres versiones distintas declaradas en el proyecto:

| Fuente | Versión declarada |
|--------|------------------|
| `Server.ini` → `[INIT] Version` | `4.0.2` |
| `SERVER.VBP` → `MajorVer.MinorVer.RevisionVer` | `0.12.359` |
| `Client.vbp` → `MajorVer.MinorVer.RevisionVer` | `4.0.3` |
| `Client.vbp` → `ExeName32` | `Winter AO Ultimate.exe` |

**Interpretación**: El `.vbp` del servidor declara `0.12.359`, lo que confirma que la base del motor es **Argentum Online v0.12.x**. La versión `4.0.x` del `Server.ini` y del cliente es el versionado propio de Winter-AO, usado para validar compatibilidad cliente-servidor. El `RevisionVer=359` con `AutoIncrementVer=1` indica que el servidor fue compilado al menos 359 veces durante su desarrollo.

---

## 2. Discrepancia de Licencia

**Estado: Verificado**

- El archivo `LICENSE` en la raíz del repositorio es **GNU General Public License v3 (GPLv3)**.
- Sin embargo, múltiples módulos del servidor contienen cabeceras **AGPL** (Affero GPL) heredadas del código original de Argentum Online. Archivos afectados: `TCP.bas`, `FileIO.bas`, `Queue.bas`, `Modulo_SysTray.bas`, `MODULO_NPCs.bas`, `modGuilds.bas`, `modCentinela.bas`, `modBanco.bas`, `ModAreas.bas`, `mdParty.bas`, `mdlComercioConUsuario.bas`, `Matematicas.bas`, y otros.

Estas cabeceras referencian `http://www.affero.org/oagpl.html` y atribuyen el código a **Pablo Ignacio Márquez** y la comunidad de Argentum Online.

**No está claro si la GPLv3 del LICENSE fue aplicada intencionalmente sobre código AGPL.** Se preserva tal cual.

---

## 3. Archivos `.WAO` — Formato Propietario

**Estado: Parcialmente verificado**

- Los archivos `.WAO` en `Cliente/Recursos/` son contenedores comprimidos con **zlib**.
- La lógica de compresión/descompresión está en `modCompression.bas` (presente en el cliente, WorldEditor y LwK-Universal).
- La herramienta `LwK-Universal` permite empaquetar y desempaquetar estos archivos.
- **No verificado**: Si existe alguna capa adicional de ofuscación o encriptación más allá de zlib. El nombre del archivo fuente sugiere solo compresión (`modCompression`), pero sin ejecutar las herramientas no es posible confirmar al 100%.

**Implicación**: Para modificar gráficos, mapas o sonidos, es necesario desempaquetar con LwK-Universal, editar los assets crudos, y volver a empaquetar.

---

## 4. Charfiles y Cuentas Vacíos

**Estado: Verificado**

- Las carpetas `server/Charfile/` y `server/Cuentas/` están presentes pero vacías (o con datos mínimos de prueba).
- No hay una cuenta de administrador preconfigurada.
- Para obtener privilegios de GM, hay que:
  1. Crear una cuenta y personaje en el juego.
  2. Editar manualmente el archivo de cuenta/personaje generado.
  3. Añadir el nombre del personaje en la sección `[Admines]` del `Server.ini`.

**Fuente**: `Server.ini` declara `Admin1=Lorwik`, `Dios1=Hennox`, `Dios2=Sigfrido`, etc., pero estos personajes no existen en los charfiles del repositorio.

---

## 5. Seguridad — Limitaciones Conocidas

**Estado: Verificado en código**

### Cliente (`ModSeguridad.bas`, ~11 KB)
- Usa `EnumWindows` y `GetWindowText` de la API Win32 para buscar ventanas con títulos sospechosos (editores de memoria, trainers).
- **Obsoleto**: Este método es trivialmente bypasseable en Windows modernos. Es una defensa superficial.

### Servidor
- **CRC**: `CrcSubKey=12345` en `Server.ini` — clave estática.
- **Encriptación**: `Encriptar=1` activa cifrado del protocolo, pero la implementación es un XOR simple o similar (verificar `Protocol.bas`).
- **Anti-macro**: `modCentinela.bas` implementa un sistema de preguntas periódicas al jugador para detectar bots. Sistema funcional pero rudimentario.
- **IP**: `SecurityIp.bas` implementa baneos por IP y límites de conexiones desde una misma dirección.

**Conclusión**: La seguridad es la típica de un servidor privado de AO de la época. No es apta para un entorno de producción moderno.

---

## 6. Binarios Precompilados — No Verificados

**Estado: Advertencia**

El repositorio incluye ejecutables precompilados que **no han sido verificados** contra el código fuente:

| Binario | Ubicación |
|---------|-----------|
| `Winter AO Ultimate.exe` (1.4 MB) | `Cliente/` |
| `Winter-AO Server.exe` (1.7 MB) | `server/` |
| `WinterAO Ultimate Launcher.exe` (147 KB) | `Cliente/` |
| `PathHelper.exe` (37 KB) | `Cliente/` |
| `RadioXtreme.exe` (33 KB) | `Cliente/` |
| `Registrar Librerias.exe` (29 KB) | `Cliente/` |
| `WorldEditorDX8.exe` (1.4 MB) | `WorldEditor/` |
| `Particle Editor - RincondelAO.exe` (254 KB) | `WorldEditor/` y `Editor de Particulas/` |
| `Indexador RincondelAO.exe` (221 KB) | `Indexador/` |
| Varios `.exe` en `LwK-Universal/` | `LwK-Universal/` |

> ⚠️ Recomendación: recompilar desde fuente si se planea ejecutar cualquiera de estos programas.

---

## 7. Deuda Técnica y Limitaciones del Motor

### Inherentes a VB6
- **Single-threaded**: No hay soporte nativo de concurrencia. El servidor atiende a todos los usuarios en un solo hilo.
- **Sin garbage collection real**: La memoria se gestiona con `Set obj = Nothing` manual. Fugas de memoria potenciales.
- **Límite de 32 bits**: Sin acceso a más de ~1.5 GB de RAM por proceso.
- **Arrays globales**: Todo el estado del mundo vive en arrays globales mutables (`UserList()`, `NpcList()`, etc.), sin encapsulamiento.

### Específicas de Winter-AO
- **Protocol.bas monolítico**: 680 KB en el servidor (el archivo más grande). Contiene toda la serialización en un solo módulo, lo que dificulta el mantenimiento.
- **Hardcoding extensivo**: IDs de objetos de facciones (armaduras imperiales/caos) están hardcodeados en `Server.ini` con números mágicos.
- **Sin tests**: No hay tests unitarios ni de integración en ningún componente.
- **Backup de NPCs**: Existe `Dat/bkNPCs.dat` (~60 KB) junto a `NPCs.dat` (~119 KB), sugiriendo que se hicieron cambios significativos sin control de versiones granular.

---

## 8. Archivo "Ideas y Bugs.txt" — Contexto Histórico

**Estado: Verificado como archivo de texto** — NO como verdad técnica.

El archivo `server/Ideas y Bugs.txt` contiene notas del equipo de desarrollo:

| Entrada | Tipo | Estado inferido |
|---------|------|-----------------|
| Bug en creación de ítems de herrero/carpintero | Bug | No verificado si fue resuelto |
| Flechitas de las quest | Bug | Probable referencia a indicadores visuales de misiones |
| Deathmatch | Idea | `AutoTorneos.bas` existe pero no se verificó si implementa este modo |
| Inteligencia Artificial | Idea | `AI_NPC.bas` implementa IA básica (persecución, patrullaje) |
| Retos 1vs1 y 2vs2 | Idea | No se encontró implementación dedicada |
| Items con Luz | Idea | `clsLight.cls` existe en el cliente, posible implementación parcial |
| Susurro a larga distancia | Idea | No verificado |

> Este archivo se usa como **contexto histórico** del estado del desarrollo, no como especificación técnica.

---

## 9. Herramientas de Terceros Incluidas

Archivos de origen externo detectados en el repositorio:

| Archivo | Origen | Notas |
|---------|--------|-------|
| `LWAO.swf` (868 KB) | Flash (SWF) | Contenido del launcher, probablemente noticias/web embebida |
| `RadioXtreme.exe` (33 KB) | Terceros | Reproductor de radio incluido con el cliente, origen no verificado |
| `Libreria de Render.exe` | Terceros | En carpeta WorldEditor, propósito no verificado |
| `AO 0.12.X Minimap Color Finder.exe` | Comunidad AO | Herramienta auxiliar en LwK-Universal |
| `Conversor.exe` | Terceros | En LwK-Universal, propósito no verificado |

---

## 10. Elementos No Verificables

Los siguientes elementos **no se pueden confirmar** solo con el análisis del código:

- Si el protocolo binario del cliente es compatible con servidores AO estándar v0.12.x, o si Winter-AO usa un protocolo modificado que requiere su propio servidor.
- El estado funcional real de los `.WAO` incluidos (si los assets están completos y no corruptos).
- Si las herramientas precompiladas corresponden exactamente al código fuente presente.
- La compatibilidad real con versiones de Windows posteriores a Windows 7.
- Si el sistema de cuentas es una adición original de Winter-AO o fue adaptado de otro proyecto de la comunidad.
