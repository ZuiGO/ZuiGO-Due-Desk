import os
import secrets
import json
from pathlib import Path

from argon2 import PasswordHasher
from cryptography.fernet import Fernet, InvalidToken

from database import IS_SQLITE


HERE = Path(__file__).parent
PASSWORDS = PasswordHasher()


def _local_key(filename: str) -> bytes:
    path = HERE / filename
    try:
        return path.read_bytes().strip()
    except FileNotFoundError:
        value = secrets.token_bytes(32)
        path.write_bytes(value)
        return value


def session_secret() -> str:
    value = os.getenv("SESSION_SECRET_KEY")
    if value:
        if os.getenv("APP_ENV") == "production" and len(value) < 32:
            raise RuntimeError("SESSION_SECRET_KEY must contain at least 32 characters in production.")
        return value
    if os.getenv("APP_ENV") == "production":
        raise RuntimeError("SESSION_SECRET_KEY must be configured for production.")
    return _local_key(".session.key").hex()


def credential_cipher() -> Fernet:
    value = os.getenv("CREDENTIALS_ENCRYPTION_KEY")
    if not value:
        if os.getenv("APP_ENV") == "production" or not IS_SQLITE:
            raise RuntimeError("CREDENTIALS_ENCRYPTION_KEY must be configured outside local development.")
        value = Fernet.generate_key().decode("ascii")
        key_path = HERE / ".credentials.key"
        try:
            value = key_path.read_text(encoding="ascii").strip()
        except FileNotFoundError:
            key_path.write_text(value, encoding="ascii")
    try:
        return Fernet(value.encode("ascii"))
    except (ValueError, TypeError) as error:
        raise RuntimeError("CREDENTIALS_ENCRYPTION_KEY must be a valid Fernet key.") from error


def encrypt_secrets(values: dict[str, str]) -> bytes:
    cipher = credential_cipher()
    return cipher.encrypt(json.dumps(values, sort_keys=True).encode("utf-8"))


def decrypt_secrets(value: bytes | None) -> dict[str, str]:
    if not value:
        return {}
    try:
        result = json.loads(credential_cipher().decrypt(value).decode("utf-8"))
        if not isinstance(result, dict) or any(not isinstance(key, str) or not isinstance(secret, str) for key, secret in result.items()):
            raise ValueError("stored credentials have an invalid format")
        return result
    except (InvalidToken, ValueError, UnicodeDecodeError) as error:
        raise RuntimeError("Stored channel credentials could not be decrypted; check the encryption key.") from error
