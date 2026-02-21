import os
import json
import base64
from pathlib import Path
from typing import Optional, Dict

try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False

try:
    import bcrypt
    BCRYPT_AVAILABLE = True
except ImportError:
    BCRYPT_AVAILABLE = False


APP_DATA_DIR  = Path.home() / ".vmware_inventory"
PROFILES_FILE = APP_DATA_DIR / "profiles.enc"
KEY_FILE      = APP_DATA_DIR / ".key"


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
    Hashea una contrasena usando bcrypt (coste=12) o PBKDF2-SHA256 como fallback.

    FIX CodeQL py/weak-sensitive-data-hashing (linea 52):
      Reemplaza hashlib.sha256(password.encode()).hexdigest() que es vulnerable
      a ataques de fuerza bruta por ser demasiado rapido (~10B hashes/seg con GPU).

    bcrypt con coste=12:
      - ~20.000 hashes/segundo (250ms por verificacion)
      - Sal de 16 bytes generada automaticamente por contrasena
      - Resistente a rainbow tables y ataques de GPU

    IMPORTANTE: bcrypt genera un hash diferente cada vez por la sal aleatoria.
      Para verificar usa verify_password(), no compares hashes directamente.
    """
    if BCRYPT_AVAILABLE:
        hashed = bcrypt.hashpw(
            password.encode("utf-8"),
            bcrypt.gensalt(rounds=12),
        )
        return hashed.decode("utf-8")

    # Fallback PBKDF2-SHA256 — tambien resuelve la alerta CodeQL
    # OWASP 2024: minimo 600.000 iteraciones para PBKDF2-SHA256
    import hashlib
    salt = os.urandom(32)
    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        iterations=600_000,
    )
    return "pbkdf2:" + salt.hex() + ":" + dk.hex()


def verify_password(password: str, stored_hash: str) -> bool:
    """
    Verifica una contrasena contra su hash almacenado.
    Resistente a timing attacks via bcrypt.checkpw / hmac.compare_digest.
    """
    if not stored_hash:
        return False
    try:
        if stored_hash.startswith("pbkdf2:"):
            import hashlib, hmac
            _, salt_hex, dk_hex = stored_hash.split(":")
            salt = bytes.fromhex(salt_hex)
            dk_expected = bytes.fromhex(dk_hex)
            dk_actual = hashlib.pbkdf2_hmac(
                "sha256",
                password.encode("utf-8"),
                salt,
                iterations=600_000,
            )
            return hmac.compare_digest(dk_actual, dk_expected)

        if BCRYPT_AVAILABLE:
            return bcrypt.checkpw(
                password.encode("utf-8"),
                stored_hash.encode("utf-8"),
            )
    except Exception:
        pass
    return False


def save_profile(name: str, host: str, user: str, password: str,
                 port: int = 443, conn_type: str = "vcenter", ignore_ssl: bool = True):
    ensure_app_dir()
    profiles = load_all_profiles()
    profiles[name] = {
        "host":         host,
        "user":         user,
        "password_enc": encrypt_password(password),
        "password_hash": hash_password(password),   # bcrypt / PBKDF2
        "port":         port,
        "conn_type":    conn_type,
        "ignore_ssl":   ignore_ssl,
    }
    if CRYPTO_AVAILABLE:
        f = Fernet(get_or_create_key())
        PROFILES_FILE.write_bytes(f.encrypt(json.dumps(profiles).encode()))
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
        PROFILES_FILE.write_bytes(f.encrypt(json.dumps(profiles).encode()))
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
