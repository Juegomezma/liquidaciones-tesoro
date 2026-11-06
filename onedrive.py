#!/usr/bin/env python3
"""Acceso de solo lectura a OneDrive con el permiso guardado por onedrive_login.py."""
import json
import os
import pathlib
import urllib.parse

import requests

AUTORIDAD = "https://login.microsoftonline.com/common/oauth2/v2.0"
GRAPH = "https://graph.microsoft.com/v1.0"
GUARDADO = pathlib.Path(os.environ.get("ONEDRIVE_TOKEN_FILE", pathlib.Path.home() / ".liquidaciones-tesoro" / "onedrive_token.json"))
PERMISOS = "offline_access Files.Read User.Read"


class OneDrive:
    def __init__(self):
        guardado = json.loads(GUARDADO.read_text(encoding="utf-8"))
        r = requests.post(f"{AUTORIDAD}/token", data={
            "grant_type": "refresh_token", "client_id": guardado["client_id"],
            "refresh_token": guardado["refresh_token"], "scope": PERMISOS}, timeout=30)
        if r.status_code != 200:
            raise RuntimeError(f"Microsoft rechazó el permiso guardado: {r.json().get('error')}. Vuelve a ejecutar onedrive_login.py")
        t = r.json()
        guardado["refresh_token"] = t.get("refresh_token", guardado["refresh_token"])
        GUARDADO.write_text(json.dumps(guardado), encoding="utf-8")
        self.cab = {"Authorization": "Bearer " + t["access_token"]}

    def get(self, ruta, **kw):
        url = ruta if ruta.startswith("http") else GRAPH + ruta
        r = requests.get(url, headers=self.cab, timeout=60, **kw)
        r.raise_for_status()
        return r

    def paginas(self, ruta):
        while ruta:
            d = self.get(ruta).json()
            yield from d.get("value", [])
            ruta = d.get("@odata.nextLink")

    def buscar(self, texto):
        return list(self.paginas("/me/drive/root/search(q='" + urllib.parse.quote(texto) + "')?$top=200&$select=id,name,size,lastModifiedDateTime,parentReference,file,folder"))

    def hijos(self, item_id="root"):
        return list(self.paginas(f"/me/drive/items/{item_id}/children?$top=200&$select=id,name,size,lastModifiedDateTime,parentReference,file,folder"))

    def descargar(self, item_id, destino):
        destino = pathlib.Path(destino)
        destino.write_bytes(self.get(f"/me/drive/items/{item_id}/content").content)
