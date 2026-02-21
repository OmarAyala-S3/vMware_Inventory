import json
import base64
from pathlib import Path
from typing import Optional

# ── Cifrado Fernet ────────────────────────────────────────────────────────────
try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False

# ── Hashing seguro de contrasenas ─────────────────────────────────────────────
try:
    import bcrypt
    BCRYPT_AVAILABLE = True
except ImportError:
    BCRYPT_AVAILABLE = False


class CredentialManager:
    """
    Gestiona el almacenamiento seguro de perfiles de conexion VMware.

    Seguridad implementada:
      - Contrasenas cifradas con Fernet (AES-128-CBC + HMAC-SHA256)
      - Hash de contrasena con bcrypt (coste=12, sal automatica por contrasena)
      - Archivo de perfiles cifrado en disco (.enc)
      - Clave de cifrado oculta en el sistema de archivos
    """

    PROFILES_FILE = Path.home() / ".vmware_inventory" / "profiles.enc"
    KEY_FILE      = Path.home() / ".vmware_inventory" / ".key"

    # Factor de coste de bcrypt — mayor = mas lento y seguro
    # 12 es el minimo recomendado por OWASP (2024)
    BCRYPT_COST = 12

    def __init__(self):
        self._ensure_dir()
        self._key = self._load_or_create_key()

    # ── Setup ─────────────────────────────────────────────────────────────────

    def _ensure_dir(self):
        self.PROFILES_FILE.parent.mkdir(parents=True, exist_ok=True)

    def _load_or_create_key(self) -> Optional[bytes]:
        if not CRYPTO_AVAILABLE:
            return None
        if self.KEY_FILE.exists():
            return self.KEY_FILE.read_bytes()
        key = Fernet.generate_key()
        self.KEY_FILE.write_bytes(key)
        # Ocultar archivo en Windows
        try:
            import ctypes
            ctypes.windll.kernel32.SetFileAttributesW(str(self.KEY_FILE), 2)
        except Exception:
            pass
        return key

    # ── API publica ───────────────────────────────────────────────────────────

    def save_profile(
        self,
        name: str,
        host: str,
        user: str,
        password: str,
        port: int = 443,
        conn_type: str = "vcenter",
    ) -> bool:
        """
        Guarda un perfil de conexion cifrado.
        La contrasena se almacena cifrada (Fernet) y hasheada (bcrypt).
        El hash sirve para verificar la contrasena sin descifrarla.
        """
        try:
            profiles = self.load_profiles()
            profiles[name] = {
                "host":         host,
                "user":         user,
                "port":         port,
                "conn_type":    conn_type,
                # Contrasena cifrada — para recuperacion y uso
                "password_enc": self._encrypt(password),
                # Hash bcrypt — para verificacion sin descifrar
                # FIX CodeQL py/weak-sensitive-data-hashing:
                #   Reemplaza hashlib.sha256 por bcrypt (coste=12, sal automatica)
                "password_hash": self._hash_password(password),
            }
            self._write_profiles(profiles)
            return True
        except Exception:
            return False

    def load_profile(self, name: str) -> Optional[dict]:
        """Carga un perfil y descifra la contrasena."""
        profiles = self.load_profiles()
        if name not in profiles:
            return None
        profile = profiles[name].copy()
        if profile.get("password_enc"):
            profile["password"] = self._decrypt(profile["password_enc"])
        return profile

    def verify_profile_password(self, name: str, password: str) -> bool:
        """
        Verifica una contrasena contra el hash almacenado sin descifrarla.
        Usa bcrypt.checkpw() que es resistente a timing attacks.
        """
        profiles = self.load_profiles()
        profile = profiles.get(name)
        if not profile or not profile.get("password_hash"):
            return False
        return self._verify_password(password, profile["password_hash"])

    def load_profiles(self) -> dict:
        """Carga todos los perfiles (sin contrasenas descifradas)."""
        if not self.PROFILES_FILE.exists():
            return {}
        try:
            data = self.PROFILES_FILE.read_bytes()
            if CRYPTO_AVAILABLE and self._key:
                f = Fernet(self._key)
                decrypted = f.decrypt(data)
                return json.loads(decrypted.decode())
            return json.loads(data.decode())
        except Exception:
            return {}

    def delete_profile(self, name: str) -> bool:
        """Elimina un perfil guardado."""
        profiles = self.load_profiles()
        if name in profiles:
            del profiles[name]
            self._write_profiles(profiles)
            return True
        return False

    # ── Hashing de contrasenas (bcrypt) ───────────────────────────────────────

    def _hash_password(self, password: str) -> str:
        """
        Hashea una contrasena con bcrypt.

        bcrypt incluye:
          - Sal aleatoria de 16 bytes generada automaticamente por contrasena
          - Factor de coste configurable (BCRYPT_COST=12 => ~250ms por hash)
          - Resistencia a ataques de GPU y rainbow tables

        FIX CodeQL py/weak-sensitive-data-hashing (lineas 62, 127):
          Reemplaza hashlib.sha256(password.encode()).hexdigest()
        """
        if not BCRYPT_AVAILABLE:
            # Fallback documentado — instalar bcrypt para produccion
            # No usar en entornos con datos reales sin bcrypt instalado
            import hashlib
            import os
            salt = os.urandom(32)
            dk = hashlib.pbkdf2_hmac(
                "sha256",
                password.encode("utf-8"),
                salt,
                iterations=600_000,    # OWASP 2024: minimo 600k para PBKDF2-SHA256
            )
            return "pbkdf2:" + salt.hex() + ":" + dk.hex()

        hashed = bcrypt.hashpw(
            password.encode("utf-8"),
            bcrypt.gensalt(rounds=self.BCRYPT_COST),
        )
        return hashed.decode("utf-8")

    def _verify_password(self, password: str, stored_hash: str) -> bool:
        """
        Verifica una contrasena contra un hash almacenado.
        Soporta hashes bcrypt y el fallback PBKDF2.
        """
        if not stored_hash:
            return False

        try:
            if stored_hash.startswith("pbkdf2:"):
                # Verificacion del fallback PBKDF2
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

    @staticmethod
    def hash_password(password: str) -> str:
        """
        Metodo estatico publico para hash de contrasena.

        FIX CodeQL py/weak-sensitive-data-hashing (linea 127):
          Reemplaza hashlib.sha256 por bcrypt o PBKDF2 segun disponibilidad.

        Nota: Este metodo genera un nuevo hash cada vez (sal nueva).
              Para verificar, usa verify_password() — no comparar hashes directamente.
        """
        if BCRYPT_AVAILABLE:
            hashed = bcrypt.hashpw(
                password.encode("utf-8"),
                bcrypt.gensalt(rounds=12),
            )
            return hashed.decode("utf-8")

        # Fallback PBKDF2 sin bcrypt
        import hashlib, os
        salt = os.urandom(32)
        dk = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt,
            iterations=600_000,
        )
        return "pbkdf2:" + salt.hex() + ":" + dk.hex()

    @staticmethod
    def verify_password(password: str, stored_hash: str) -> bool:
        """
        Verifica una contrasena contra un hash (estatico).
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
                    "sha256", password.encode("utf-8"), salt, iterations=600_000
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

    # ── Cifrado Fernet ────────────────────────────────────────────────────────

    def _encrypt(self, text: str) -> str:
        if not CRYPTO_AVAILABLE or not self._key:
            return base64.b64encode(text.encode()).decode()
        f = Fernet(self._key)
        return f.encrypt(text.encode()).decode()

    def _decrypt(self, encrypted: str) -> str:
        if not CRYPTO_AVAILABLE or not self._key:
            return base64.b64decode(encrypted.encode()).decode()
        f = Fernet(self._key)
        return f.decrypt(encrypted.encode()).decode()

    def _write_profiles(self, profiles: dict):
        data = json.dumps(profiles, indent=2).encode()
        if CRYPTO_AVAILABLE and self._key:
            f = Fernet(self._key)
            self.PROFILES_FILE.write_bytes(f.encrypt(data))
        else:
            self.PROFILES_FILE.write_bytes(data)
