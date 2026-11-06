#!/usr/bin/env python3
"""Guarda cifrado el permiso de OneDrive entre ejecuciones automáticas.

  python estado.py abrir     descifra estado/onedrive.enc hacia ONEDRIVE_TOKEN_FILE
  python estado.py guardar   lo vuelve a cifrar si ya pasó una semana (Microsoft renueva el permiso al usarlo)
  python estado.py iniciar   cifra por primera vez el permiso local

La llave es TOKEN_KEY (32 bytes en base64), que vive solo en los secretos del repositorio.
"""
import base64
import datetime as dt
import json
import os
import pathlib
import secrets
import sys

from cryptography.hazmat.primitives.ciphers.aead import AESGCM

BASE = pathlib.Path(__file__).resolve().parent
SOBRE = BASE / "estado" / "onedrive.enc"
PERMISO = pathlib.Path(os.environ.get("ONEDRIVE_TOKEN_FILE", pathlib.Path.home() / ".liquidaciones-tesoro" / "onedrive_token.json"))
CADA = dt.timedelta(days=7)


def llave():
    return AESGCM(base64.b64decode(os.environ["TOKEN_KEY"]))


def cifrar():
    iv = secrets.token_bytes(12)
    dato = llave().encrypt(iv, PERMISO.read_bytes(), None)
    SOBRE.parent.mkdir(exist_ok=True)
    SOBRE.write_text(json.dumps(dict(
        fecha=dt.datetime.now(dt.timezone.utc).isoformat(),
        iv=base64.b64encode(iv).decode(), dato=base64.b64encode(dato).decode())), encoding="utf-8")


def main():
    accion = sys.argv[1]
    if accion == "abrir":
        s = json.loads(SOBRE.read_text(encoding="utf-8"))
        PERMISO.parent.mkdir(parents=True, exist_ok=True)
        PERMISO.write_bytes(llave().decrypt(base64.b64decode(s["iv"]), base64.b64decode(s["dato"]), None))
    elif accion == "guardar":
        s = json.loads(SOBRE.read_text(encoding="utf-8"))
        if dt.datetime.now(dt.timezone.utc) - dt.datetime.fromisoformat(s["fecha"]) >= CADA:
            cifrar()
            print("Permiso renovado")
    elif accion == "iniciar":
        cifrar()
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
