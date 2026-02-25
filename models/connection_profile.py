"""
models/connection_profile.py
Modelo de perfil de conexión para gestión multi-vCenter/ESXi
"""
from dataclasses import dataclass, field
from enum import Enum
import uuid


def _short_uuid() -> str:
    """Genera un ID corto de 8 caracteres. Wrapper nombrado para evitar lambda."""
    return str(uuid.uuid4())[:8]


class ConnectionType(Enum):
    VCENTER = "vCenter"
    ESXI = "ESXi Host"


class ConnectionStatus(Enum):
    PENDING   = "Pendiente"
    TESTING   = "Probando..."
    OK        = "Conectado"
    SCANNING  = "Escaneando..."
    DONE      = "Completado"
    ERROR     = "Error"
    SKIPPED   = "Omitido"


@dataclass
class ConnectionProfile:
    """
    Representa una conexión individual a vCenter o ESXi.
    Cada perfil tiene un ID único para rastrearlo en la UI y en threads.
    """
    host: str
    username: str
    password: str
    connection_type: ConnectionType = ConnectionType.VCENTER
    port: int = 443
    ignore_ssl: bool = True
    alias: str = ""                          # Nombre amigable opcional

    # Campos de runtime (no se persisten)
    id: str = field(default_factory=_short_uuid)
    status: ConnectionStatus = field(default=ConnectionStatus.PENDING)
    error_message: str = ""
    vms_found: int = 0
    hosts_found: int = 0
    datastores_found: int = 0
    networks_found: int = 0


@dataclass
class ScanConfig:
    """Configuración para el escaneo multi-conexión."""
    parallel: bool = False
    max_workers: int = 3
    timeout: int = 30
    include_vms: bool = True
    include_hosts: bool = True
    include_datastores: bool = True
    include_networks: bool = True
    export_partial: bool = True
