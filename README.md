# ⚡ VMware Inventory System

> Aplicación de escritorio Python para extracción, visualización y exportación de inventario VMware vCenter/ESXi. Reemplaza scripts PowerShell con una interfaz gráfica profesional, soporte multi-conexión y exportación Excel consolidada.

![Python](https://img.shields.io/badge/Python-3.12+-3776AB?style=flat-square&logo=python&logoColor=white)
![Tkinter](https://img.shields.io/badge/UI-Tkinter%20%2B%20ttkbootstrap-darkly?style=flat-square)
![pyVmomi](https://img.shields.io/badge/VMware-pyVmomi-607078?style=flat-square&logo=vmware)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Status](https://img.shields.io/badge/Status-Production-brightgreen?style=flat-square)

---

## 📋 Tabla de Contenidos

- [Descripción](#-descripción)
- [Características](#-características)
- [Arquitectura](#-arquitectura)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Requisitos](#-requisitos)
- [Instalación](#-instalación)
- [Uso](#-uso)
- [Módulos](#-módulos)
- [Exportación Excel](#-exportación-excel)
- [Seguridad](#-seguridad)
- [Solución de Problemas](#-solución-de-problemas)

---

## 📖 Descripción

VMware Inventory System es una aplicación de escritorio **100% Python** que se conecta a entornos VMware (vCenter Server y hosts ESXi standalone) para extraer automáticamente el inventario completo de infraestructura virtual.

**Problema que resuelve:** La gestión manual de inventario VMware mediante scripts PowerShell dispersos, sin interfaz gráfica, sin soporte multi-entorno y con exportaciones inconsistentes.

**Solución:** Una única aplicación portable que centraliza conexiones, automatiza la extracción y genera reportes Excel estructurados con una sola acción.

---

## ✨ Características

### Conexión
- ✅ Soporte para **vCenter Server** y **Hosts ESXi standalone**
- ✅ Validación SSL configurable (producción/laboratorio)
- ✅ **Perfiles de conexión guardados** con cifrado Fernet (AES-128)
- ✅ Prueba de conectividad antes de extraer

### Multi-Conexión
- ✅ **Tabla de conexiones** con gestión Add / Edit / Remove
- ✅ Escaneo **paralelo o secuencial** configurable por el usuario
- ✅ Workers configurables (1–10 conexiones simultáneas)
- ✅ Barra de progreso individual por fuente
- ✅ Estado visual en tiempo real por conexión

### Extracción de Inventario
- ✅ **Máquinas Virtuales** — 20+ campos incluyendo NICs, discos, SO, estado
- ✅ **Hosts ESXi** — hardware, CPU, RAM, versión, cluster, serie
- ✅ **Datastores** — capacidad, espacio libre/usado, tipo
- ✅ **Redes** — tipo (Standard/Distributed), VLAN, switch
- ✅ Procesador de VM obtenido desde el host físico padre via cache pre-cargado

### Exportación Excel
- ✅ Exportación **individual** por conexión
- ✅ Exportación **consolidada multi-fuente** en un solo archivo
- ✅ Hoja de Resumen ejecutivo con estado de cada fuente
- ✅ Una hoja por fuente con todas sus secciones
- ✅ 4 hojas consolidadas al final
- ✅ Formato profesional con colores por estado
- ✅ Nombre automático con timestamp

---

## 🏗️ Arquitectura

```
┌─────────────────────────────────────────────────────────────────┐
│                    CAPA DE PRESENTACIÓN (UI)                     │
│                                                                  │
│   app.py (ventana principal)    multi_connection_panel.py        │
│   ├── Sidebar conexión          ├── Tabla Add/Edit/Remove        │
│   ├── Treeviews preview         ├── ScanConfigDialog             │
│   ├── Tabs VMs/Hosts/DS/Redes   └── Barra progreso por fuente    │
│   └── Log consola                                                │
│                     multi_tab.py (frame integrador)              │
└──────────────────────────┬──────────────────────────────────────┘
                           │ callbacks + threading.after()
┌──────────────────────────▼──────────────────────────────────────┐
│                     CAPA DE SERVICIOS                            │
│                                                                  │
│   vmware_service.py              connection_manager.py           │
│   ├── connect() / disconnect()   ├── add/remove profile          │
│   ├── extract_vms()              ├── test_connection()           │
│   │   └── host_cpu_map cache     ├── start_scan()                │
│   ├── extract_hosts()            ├── _scan_sequential()          │
│   ├── extract_datastores()       ├── _scan_parallel()            │
│   └── extract_networks()         │   └── ThreadPoolExecutor      │
│        └── pyVmomi API           └── ConsolidatedResult          │
└──────────────────────────┬──────────────────────────────────────┘
                           │ PropertyCollector (eficiente)
┌──────────────────────────▼──────────────────────────────────────┐
│               INFRAESTRUCTURA VMWARE                             │
│                                                                  │
│    vCenter-Prod    vCenter-Dev    ESXi Standalone ...            │
│    (puerto 443)    (puerto 443)   (puerto 443)                   │
└──────────────────────────┬──────────────────────────────────────┘
                           │ datos extraídos
┌──────────────────────────▼──────────────────────────────────────┐
│                       CAPA DE DATOS                              │
│                                                                  │
│   vm_model.py               connection_profile.py               │
│   ├── VMModel                ├── ConnectionProfile               │
│   ├── HostModel              ├── ScanConfig                      │
│   ├── DatastoreModel         ├── ConnectionType                  │
│   ├── NetworkModel           └── ConnectionStatus                │
│   ├── NicInfo                                                    │
│   └── DiskInfo               SimpleInventory (runtime DTO)       │
└──────────────────────────┬──────────────────────────────────────┘
                           │ DataFrames + openpyxl
┌──────────────────────────▼──────────────────────────────────────┐
│                    CAPA DE EXPORTACIÓN                           │
│                                                                  │
│   excel_exporter.py             multi_exporter.py               │
│   └── Conexión simple           ├── Hoja Resumen ejecutivo       │
│       4 hojas estándar          ├── N hojas por fuente           │
│       Colores por estado        ├── 4 hojas consolidadas         │
│                                 ├── Auto-fit columnas            │
│                                 └── Freeze panes + bordes        │
└─────────────────────────────────────────────────────────────────┘
```

### Flujo de operación — Escaneo Multi-Conexión

```
Usuario           UI Panel         ConnectionManager      VMwareService
   │                 │                    │                    │
   ├─ Agrega ───────►│                    │                    │
   │   conexiones    │                    │                    │
   ├─ Configura ────►│                    │                    │
   │   modo scan     │                    │                    │
   ├─ "Escanear" ───►│                    │                    │
   │                 ├─ start_scan() ────►│                    │
   │                 │                   ├─ Thread #1 ────────►│
   │                 │                   │                     ├─ connect()
   │                 │                   │                     ├─ extract_vms()
   │                 │                   │                     ├─ extract_hosts()
   │                 │◄─ on_progress() ──┤◄────────────────────┤
   │◄─ UI update ────┤  (% por fuente)   │                    │
   │                 │                   ├─ _tag_inventory() ──┤
   │                 │                   │  (inyecta "Fuente") │
   │                 │◄─ on_complete() ──┤                    │
   │                 │  ConsolidatedResult│                    │
   ├─ "Exportar" ───►│                    │                    │
   │                 ├─ MultiSourceExporter.export()           │
   │◄─ archivo .xlsx─┤                    │                    │
```

---

## 📁 Estructura del Proyecto

```
vmware_inventory/
│
├── main.py                          # Punto de entrada
│
├── ui/
│   ├── app.py                       # Ventana principal (Tkinter + ttkbootstrap)
│   ├── multi_connection_panel.py    # Panel gestión multi-conexión + diálogos
│   └── multi_tab.py                 # Frame integrador pestaña multi-conexión
│
├── services/
│   ├── vmware_service.py            # Conexión y extracción VMware via pyVmomi
│   └── connection_manager.py        # Orquestador multi-conexión + ThreadPoolExecutor
│
├── models/
│   ├── vm_model.py                  # Dataclasses: VMModel, HostModel, etc.
│   └── connection_profile.py        # ConnectionProfile, ScanConfig, enums
│
├── exporters/
│   ├── excel_exporter.py            # Excel para conexión individual
│   └── multi_exporter.py            # Excel consolidado multi-fuente
│
├── utils/
│   ├── credentials.py               # Funciones de gestión de perfiles
│   └── security.py                  # CredentialManager con cifrado Fernet
│
├── setup_multi_connection.py        # Instalación/integración automática
├── build_exe.py                     # Compilación a .exe con PyInstaller
└── requirements.txt
```

---

## 🔧 Requisitos

| Componente | Versión |
|---|---|
| Python | 3.12+ |
| pyVmomi | 8.0.0+ |
| ttkbootstrap | 1.10.0+ |
| pandas | 2.0.0+ |
| openpyxl | 3.1.0+ |
| cryptography | 41.0.0+ |

**Sistema Operativo:** Windows 10/11 · Linux Ubuntu 20+ · macOS 12+

---

## 🚀 Instalación

```bash
# 1. Clonar
git clone https://github.com/tu-usuario/vmware-inventory-system.git
cd vmware-inventory-system/vmware_inventory

# 2. Entorno virtual
python -m venv venv
venv\Scripts\activate        # Windows
source venv/bin/activate     # Linux/macOS

# 3. Dependencias
pip install -r requirements.txt

# 4. Ejecutar
python main.py
```

### Integración automática (multi-conexión)

```bash
# Coloca los archivos nuevos junto a setup_multi_connection.py y ejecuta:
python setup_multi_connection.py
```

El script: valida entorno → crea backup de app.py → copia archivos → parchea app.py → verifica sintaxis.

---

## 📖 Uso

### Conexión Simple

1. Selecciona tipo (`vCenter` o `Host ESXi`) en el panel izquierdo
2. Ingresa `IP/FQDN`, `Puerto`, `Usuario`, `Contraseña`
3. **Probar Conexión** → **Extraer Inventario Completo**
4. **Exportar a Excel**

### Multi-Conexión

1. Pestaña **🌐 Multi-Conexión**
2. **➕ Agregar** — registra cada vCenter/ESXi
3. **⚙ Configurar Escaneo** — modo paralelo/secuencial, workers, timeout
4. **🚀 Escanear Todo** — progreso en tiempo real por fuente
5. **💾 Exportar Excel Consolidado**

### Perfiles de Conexión

- Activa **"Guardar como:"** antes de extraer para cifrar y guardar credenciales
- Archivo en: `~/.vmware_inventory/profiles.enc`
- Clave en: `~/.vmware_inventory/.key` (oculto en Windows)

---

## 📦 Módulos Clave

### `vmware_service.py` — Extracción VMware

Usa `PropertyCollector` (no traversal recursivo) para máxima eficiencia en entornos grandes.

**Técnica del procesador:** Las VMs no tienen campo CPU propio en VMware. El sistema pre-carga `{ host_moId → cpuModel }` en una sola query a `HostSystem`, luego cruza `runtime.host._moId` por VM. En 500 VMs sobre 10 hosts: 10 lecturas de CPU en vez de 500.

### `connection_manager.py` — Orquestación

`ThreadPoolExecutor` con semáforo configurable. Callbacks de progreso enviados al hilo UI via `root.after(0, fn)` para thread-safety en Tkinter.

### `multi_exporter.py` — Excel Multi-Fuente

Convierte `ConsolidatedResult` a Excel con pandas + openpyxl. Los row-converters (`_vm_to_row`, `_host_to_row`) leen los atributos reales del modelo (ej. `ram_mb`, `vcpu`, `os_name`) y los mapean a nombres de columna legibles.

---

## 📊 Campos Exportados — VMs

| Columna Excel | Campo VMModel | Fuente pyVmomi |
|---|---|---|
| Hostname | `hostname` | `name` |
| IP | via `nics[].ip_addresses` | `guest.net` |
| MAC | via `nics[].mac_address` | `config.hardware.device` |
| vCPU | `vcpu` | `config.hardware.numCPU` |
| RAM (GB) | `ram_mb / 1024` | `config.hardware.memoryMB` |
| Procesador | `processor` | `host_cpu_map[runtime.host._moId]` |
| Sistema Operativo | `os_name` | `guest.guestFullName` |
| Edición SO | `os_edition` | `config.guestFullName` |
| Discos | `disks[].size_gb` | `config.hardware.device` (VirtualDisk) |
| Estado | `power_state` | `runtime.powerState` |
| VMware Tools | `tools_status` | `guest.toolsStatus` |
| Versión HW | `hw_version` | `config.version` |
| Fuente | `source_name` | Inyectado por `connection_manager` |

---

## 🔒 Seguridad

| Escenario | Mecanismo |
|---|---|
| `cryptography` instalado | Fernet AES-128-CBC + HMAC-SHA256 |
| Sin `cryptography` | Base64 (instalar `cryptography` para producción) |
| Hash verificación | SHA-256 independiente por perfil |

> ⚠️ Las contraseñas **nunca** se almacenan en texto plano.

---

## 🛠️ Compilar a .exe

```bash
pip install pyinstaller
python build_exe.py
# Resultado: dist/VMwareInventory.exe (portable, sin Python requerido)
```

---

## 🐛 Solución de Problemas

**`ImportError: cannot import name 'VMInfo'`**
→ Corregir `models/__init__.py`:
```python
from .vm_model import VMModel, HostModel, DatastoreModel, NetworkModel, NicInfo, DiskInfo
```

**`VMwareService.__init__() got an unexpected keyword argument 'host'`**
→ Las credenciales van en `.connect()`, no en el constructor:
```python
svc = VMwareService()
svc.connect(host=..., user=..., password=..., port=..., ignore_ssl=...)
```

**Procesador vacío en Excel**
→ Verificar que `vmware_service.py` tiene el método `extract_vms()` con construcción del `host_cpu_map` antes del loop de VMs.

**SSL Error al conectar**
→ Activar **"Ignorar certificado SSL"**. Los entornos de laboratorio usan certificados autofirmados.

**UI se congela durante extracción**
→ Toda operación de red debe correr en `threading.Thread(daemon=True)`. Actualizar UI solo con `root.after(0, callback)`.

---

## 📄 Licencia

MIT License — libre para uso personal y comercial.

---

*Python 3.12 · pyVmomi 8.x · ttkbootstrap 1.10 · pandas 2.x · openpyxl 3.1*
