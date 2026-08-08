#!/usr/bin/env python3
"""Genera el hash de una contraseña para un jefe de becarios nuevo.

Uso (desde la raíz del proyecto, con el entorno virtual activado):

    python3 scripts/generar_password.py

Pide la contraseña en texto plano y muestra el hash bcrypt que va en Secrets,
en la sección [jefes.<usuario>] (campo `password`). Nunca guardes la
contraseña en texto plano en Secrets ni en ningún archivo del repo.
"""

import getpass

import streamlit_authenticator as stauth


def main() -> None:
    password = getpass.getpass("Contraseña para el jefe: ")
    confirm = getpass.getpass("Repite la contraseña: ")
    if password != confirm:
        print("Las contraseñas no coinciden.")
        raise SystemExit(1)

    hashed = stauth.Hasher().hash(password)
    print("\nCopia esta línea en Secrets, dentro de [jefes.<usuario>]:\n")
    print(f'password = "{hashed}"')


if __name__ == "__main__":
    main()
