"""
Utilidades: credenciales seguras, perfiles, helpers
"""
import json
import base64
from pathlib import Path
from typing import Optional, Dict

try:
    import bcrypt
    BCRYPT_AVAILABLE = True
except ImportError:
    BCRYPT_AVAILABLE = False

try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False

APP_DATA_DIR = Path.home() / ".vmware_inventory"
PROFILES_FILE = APP_DATA_DIR / "profiles.enc"
KEY_FILE = APP_DATA_DIR / ".key"

def ensure_app_dir():
    APP_DATA_DIR.mkdir(exist_ok=True)

def get_or_create_key() -> bytes:
    ensure_app_dir()
    if KEY_FILE.exists():
        return KEY_FILE.read_bytes()
    key = Fernet.generate_key()
    KEY_FILE.write_bytes(key)
    KEY_FILE.chmod(0o600)
    return key

def encrypt_password(password: str) -> str:
    if not CRYPTO_AVAILABLE:
        return base64.b64encode(password.encode()).decode()
    f = Fernet(get_or_create_key())
    return f.encrypt(password.encode()).decode()

def decrypt_password(encrypted: str) -> str:
    if not CRYPTO_AVAILABLE:
        return base64.b64decode(encrypted.encode()).decode()
    f = Fernet(get_or_create_key())
    return f.decrypt(encrypted.encode()).decode()

def hash_password(password: str) -> str:
    """
    Hashea una contrasena con bcrypt (coste=12) o PBKDF2 como fallback.
    FIX CodeQL py/weak-sensitive-data-hashing (linea 45).
    """
    if BCRYPT_AVAILABLE:
        return bcrypt.hashpw(
            password.encode("utf-8"),
            bcrypt.gensalt(rounds=12),
        ).decode("utf-8")
    import hashlib as _hl, os as _os
    salt = _os.urandom(32)
    dk = _hl.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 600_000)
    return "pbkdf2:" + salt.hex() + ":" + dk.hex()


def verify_password(password: str, stored_hash: str) -> bool:
    """Verifica contrasena contra hash almacenado (bcrypt o PBKDF2)."""
    if not stored_hash:
        return False
    try:
        if stored_hash.startswith("pbkdf2:"):
            import hashlib as _hl, hmac as _hmac
            _, salt_hex, dk_hex = stored_hash.split(":")
            salt = bytes.fromhex(salt_hex)
            dk_exp = bytes.fromhex(dk_hex)
            dk_act = _hl.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, 600_000)
            return _hmac.compare_digest(dk_act, dk_exp)
        if BCRYPT_AVAILABLE:
            return bcrypt.checkpw(password.encode("utf-8"), stored_hash.encode("utf-8"))
    except Exception as _e:
        pass  # Error de verificacion — hash invalido o incompatible
    return False

def save_profile(name: str, host: str, user: str, password: str,
                 port: int = 443, conn_type: str = "vcenter", ignore_ssl: bool = True):
    ensure_app_dir()
    profiles = load_all_profiles()
    profiles[name] = {
        "host": host,
        "user": user,
        "password_enc": encrypt_password(password),
        "password_hash": hash_password(password),
        "port": port,
        "conn_type": conn_type,
        "ignore_ssl": ignore_ssl,
    }
    if CRYPTO_AVAILABLE:
        f = Fernet(get_or_create_key())
        data = f.encrypt(json.dumps(profiles).encode())
        PROFILES_FILE.write_bytes(data)
    else:
        PROFILES_FILE.write_text(json.dumps(profiles))

def load_all_profiles() -> Dict:
    if not PROFILES_FILE.exists():
        return {}
    try:
        if CRYPTO_AVAILABLE:
            f = Fernet(get_or_create_key())
            data = f.decrypt(PROFILES_FILE.read_bytes())
            return json.loads(data)
        else:
            return json.loads(PROFILES_FILE.read_text())
    except Exception:
        return {}

def load_profile(name: str) -> Optional[Dict]:
    profiles = load_all_profiles()
    p = profiles.get(name)
    if p and "password_enc" in p:
        p = dict(p)
        p["password"] = decrypt_password(p["password_enc"])
    return p

def delete_profile(name: str):
    profiles = load_all_profiles()
    profiles.pop(name, None)
    if CRYPTO_AVAILABLE:
        f = Fernet(get_or_create_key())
        data = f.encrypt(json.dumps(profiles).encode())
        PROFILES_FILE.write_bytes(data)
    else:
        PROFILES_FILE.write_text(json.dumps(profiles))

def list_profiles():
    return list(load_all_profiles().keys())

def format_bytes(b: int) -> str:
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if b < 1024:
            return f"{b:.1f} {unit}"
        b /= 1024
    return f"{b:.1f} PB"
