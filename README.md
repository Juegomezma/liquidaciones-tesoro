# Tablero de liquidaciones

Tablero web para seguir las ventas, las máquinas, los gastos y el cierre de un casino a partir de los Excel diarios de contadores guardados en OneDrive.

## Cómo funciona

1. `importar.py` lee los Excel desde OneDrive con permiso de solo lectura. Nunca los modifica.
2. Revisa cada mes y genera los datos del tablero, cifrados con una clave (`docs/datos.enc.json`).
3. Una tarea de GitHub Actions repite el proceso cada hora y publica los cambios.
4. `docs/index.html` es la página. Pide la clave y descifra los datos en el navegador.

Ninguna cifra se guarda sin cifrar en este repositorio. Los Excel y los datos locales están excluidos en `.gitignore`.

## Uso local

```
pip install -r requirements.txt
python onedrive_login.py     # solo la primera vez
python importar.py           # genera docs/datos.local.js
```

Después abre `docs/index.html` en el navegador.

## Secretos del repositorio

| Secreto | Para qué sirve |
|---|---|
| `DATOS_CLAVE` | Clave con la que se cifran los datos y se abre el tablero |
| `TOKEN_KEY` | Llave de 32 bytes en base64 que protege el permiso de OneDrive en `estado/onedrive.enc` |

Para cambiar la cuenta de OneDrive o retirar el acceso, revoca el permiso en la cuenta de Microsoft y vuelve a ejecutar `onedrive_login.py` y `estado.py iniciar`.
