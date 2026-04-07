# -*- coding: utf-8 -*-
import base64
import hashlib
import hmac
import secrets
from typing import Any, Dict, List, Optional
from .helpers import AppError

class PasswordManager:
    ITERATIONS = 200_000

    @staticmethod
    def hash_password(password: str, salt: Optional[bytes] = None):
        if salt is None:
            salt = secrets.token_bytes(16)
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PasswordManager.ITERATIONS)
        return base64.b64encode(digest).decode("ascii"), base64.b64encode(salt).decode("ascii")

    @staticmethod
    def verify_password(password: str, stored_hash: str, stored_salt: str) -> bool:
        if not stored_hash or not stored_salt:
            return False
        try:
            salt = base64.b64decode(stored_salt.encode("ascii"))
            expected_hash = base64.b64decode(stored_hash.encode("ascii"))
        except Exception:
            return False
        current = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, PasswordManager.ITERATIONS)
        return hmac.compare_digest(current, expected_hash)

class UserManager:
    @staticmethod
    def list_users(config: Dict[str, Any]) -> List[Dict[str, str]]:
        users = config.get("security", {}).get("users", [])
        return [u for u in users if u.get("username")]

    @staticmethod
    def find_user(config: Dict[str, Any], username: str) -> Optional[Dict[str, str]]:
        name = (username or "").strip().lower()
        for user in UserManager.list_users(config):
            if user.get("username", "").strip().lower() == name:
                return user
        return None

    @staticmethod
    def add_user(config: Dict[str, Any], username: str, password: str):
        if UserManager.find_user(config, username):
            raise AppError("Já existe um usuário com esse nome.")
        pwd_hash, salt = PasswordManager.hash_password(password)
        config["security"].setdefault("users", []).append({
            "username": username.strip(),
            "password_hash": pwd_hash,
            "password_salt": salt,
        })

    @staticmethod
    def update_user_password(config: Dict[str, Any], username: str, new_password: str):
        user = UserManager.find_user(config, username)
        if not user:
            raise AppError("Usuário não encontrado.")
        pwd_hash, salt = PasswordManager.hash_password(new_password)
        user["password_hash"] = pwd_hash
        user["password_salt"] = salt

    @staticmethod
    def remove_user(config: Dict[str, Any], username: str):
        config["security"]["users"] = [
            u for u in config.get("security", {}).get("users", [])
            if u.get("username", "").strip().lower() != username.strip().lower()
        ]

    @staticmethod
    def validate_login(config: Dict[str, Any], username: str, password: str) -> bool:
        user = UserManager.find_user(config, username)
        if not user:
            return False
        return PasswordManager.verify_password(
            password,
            user.get("password_hash", ""),
            user.get("password_salt", ""),
        )
