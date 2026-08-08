"""Login de jefes de becarios (streamlit-authenticator).

Las credenciales viven en `st.secrets`, nunca en el Excel ni en el repo:

    [auth]
    cookie_name = "becarios_up_auth"
    cookie_key = "una-clave-larga-y-aleatoria"
    cookie_expiry_days = 30

    [jefes.mlopez]
    name = "Mariana López"          # debe escribirse igual que en la columna
                                     # ENCARGADO del Excel (acentos/mayúsculas no importan)
    password = "$2b$12$..."         # hash generado con scripts/generar_password.py

Si `st.secrets` no tiene ningún jefe configurado, el login simplemente no se
muestra y el dashboard sigue funcionando igual que antes.
"""

from __future__ import annotations

import streamlit as st
import streamlit_authenticator as stauth

from data_loader import normalize_text


def _credenciales() -> dict:
    jefes = st.secrets.get("jefes", {})
    return {
        "usernames": {
            username: {"name": info["name"], "password": info["password"]}
            for username, info in jefes.items()
        }
    }


def configured_jefe_names() -> set[str]:
    """Nombres (normalizados) de los jefes que ya tienen cuenta en Secrets."""
    jefes = st.secrets.get("jefes", {})
    return {normalize_text(info["name"]) for info in jefes.values()}


def _authenticator() -> stauth.Authenticate | None:
    auth_cfg = st.secrets.get("auth", {})
    cookie_key = auth_cfg.get("cookie_key")
    if not cookie_key:
        st.sidebar.error("Falta `cookie_key` en Secrets (sección [auth]) para poder iniciar sesión.")
        return None
    return stauth.Authenticate(
        _credenciales(),
        auth_cfg.get("cookie_name", "becarios_up_auth"),
        cookie_key,
        auth_cfg.get("cookie_expiry_days", 30),
    )


def render_login_sidebar() -> tuple[str | None, str | None]:
    """Dibuja el login de jefes en la barra lateral.

    Devuelve (nombre_del_jefe, username) si hay una sesión activa, o (None, None)
    si no hay jefes configurados o nadie ha iniciado sesión.
    """
    if not st.secrets.get("jefes"):
        return None, None
    authenticator = _authenticator()
    if authenticator is None:
        return None, None

    st.sidebar.markdown("#### Acceso de jefes de becarios")
    authenticator.login(location="sidebar")
    status = st.session_state.get("authentication_status")
    if status is False:
        st.sidebar.error("Usuario o contraseña incorrectos.")
    elif status is None:
        st.sidebar.caption(
            "¿Cómo funciona? Cada jefe tiene su propio usuario. Al iniciar sesión, ves "
            "solo los eventos donde apareces como ENCARGADO en el Excel, y ahí puedes "
            "confirmar o corregir las horas de los becarios que te apoyaron. Si no tienes "
            "usuario todavía, pídele a quien administra el dashboard que te dé de alta."
        )
    else:
        st.sidebar.success(f"Sesión: {st.session_state.get('name')}")
        authenticator.logout("Cerrar sesión", location="sidebar")

    if status:
        return st.session_state.get("name"), st.session_state.get("username")
    return None, None
