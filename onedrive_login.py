#!/usr/bin/env python3
"""Autoriza una sola vez la lectura de OneDrive, sin compartir la contraseña.

La persona inicia sesión en la página de Microsoft con el código que muestra este
script. El permiso concedido es de solo lectura (Files.Read) y se puede retirar en
https://account.live.com/consent/Manage (cuentas personales) o en
https://myaccount.microsoft.com (cuentas de trabajo).
El permiso se guarda fuera del proyecto y nunca se imprime.
"""
import json
import os
import pathlib
import subprocess
import sys
import time

import requests

# Aplicación pública de Microsoft (Microsoft Graph Command Line Tools).
CLIENT_ID = os.environ.get("ONEDRIVE_CLIENT_ID", "14d82eec-204b-4c2f-b7e8-296a70dab67e")
AUTORIDAD = "https://login.microsoftonline.com/common/oauth2/v2.0"
PERMISOS = "offline_access Files.Read User.Read"
GUARDADO = pathlib.Path.home() / ".liquidaciones-tesoro" / "onedrive_token.json"


def guardar(token):
    GUARDADO.parent.mkdir(parents=True, exist_ok=True)
    GUARDADO.write_text(json.dumps({
        "client_id": CLIENT_ID,
        "refresh_token": token["refresh_token"],
    }), encoding="utf-8")
    if os.name == "nt":
        # Primero se concede el acceso propio; solo si funciona se retira la herencia.
        usuario = subprocess.run(["whoami"], capture_output=True, text=True, check=True).stdout.strip()
        subprocess.run(["icacls", str(GUARDADO), "/grant:r", f"{usuario}:F"], capture_output=True, check=True)
        subprocess.run(["icacls", str(GUARDADO), "/inheritance:r"], capture_output=True, check=True)


def main():
    r = requests.post(f"{AUTORIDAD}/devicecode", data={"client_id": CLIENT_ID, "scope": PERMISOS}, timeout=30)
    if r.status_code != 200:
        sys.exit(f"Microsoft no entregó un código: {r.json().get('error_description', r.text)[:300]}")
    d = r.json()
    print(f"PAGINA: {d['verification_uri']}", flush=True)
    print(f"CODIGO: {d['user_code']}", flush=True)
    print(f"VENCE_EN_MINUTOS: {int(d['expires_in']) // 60}", flush=True)

    limite = time.time() + int(d["expires_in"])
    espera = int(d.get("interval", 5))
    while time.time() < limite:
        time.sleep(espera)
        t = requests.post(f"{AUTORIDAD}/token", data={
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": CLIENT_ID,
            "device_code": d["device_code"],
        }, timeout=30)
        cuerpo = t.json()
        error = cuerpo.get("error")
        if t.status_code == 200:
            guardar(cuerpo)
            cab = {"Authorization": "Bearer " + cuerpo["access_token"]}
            yo = requests.get("https://graph.microsoft.com/v1.0/me", headers=cab, timeout=30).json()
            unidad = requests.get("https://graph.microsoft.com/v1.0/me/drive", headers=cab, timeout=30).json()
            print(f"AUTORIZADO: {yo.get('userPrincipalName') or yo.get('mail')}", flush=True)
            print(f"TIPO_UNIDAD: {unidad.get('driveType')}", flush=True)
            print(f"PERMISO_GUARDADO: {GUARDADO}", flush=True)
            return
        if error == "authorization_pending":
            print(f"ESPERANDO: faltan {int(limite - time.time()) // 60} min", flush=True)
            continue
        if error == "slow_down":
            espera += 5
            continue
        sys.exit(f"No se completó la autorización: {error}: {cuerpo.get('error_description', '')[:300]}")
    sys.exit("El código venció sin que se completara el inicio de sesión.")


if __name__ == "__main__":
    main()
