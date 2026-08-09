from __future__ import annotations

from typing import Any

from src.config import get_app_env, initial_admin_config, usar_postgres
from src.repositories.audit_repository import registrar_auditoria
from src.repositories.user_repository import (
    actualizar_usuario,
    autenticar_usuario,
    cambiar_password,
    crear_usuario,
    listar_usuarios,
)
from src.security import ROLES_PERMISOS


def pantalla_login(st: Any):
    st.set_page_config(page_title="Login | Gestor FD-GC71 / FD-GC72", layout="centered")
    st.title("Ingreso al gestor académico")
    st.caption("FD-GC71 / FD-GC72 con control de acceso por perfiles.")

    with st.form("login_form", clear_on_submit=False):
        usuario = st.text_input("Usuario", value="admin")
        password = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Entrar", use_container_width=True)

    if entrar:
        try:
            user = autenticar_usuario(usuario, password)
            if user:
                st.session_state["auth_user"] = user
                registrar_auditoria("Login", "Ingreso correcto", user=user)
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
        except Exception as exc:
            st.error(str(exc))

    with st.expander("Primera instalación"):
        if get_app_env() == "production":
            st.info("Las credenciales iniciales se leen desde los secretos INITIAL_ADMIN_USER e INITIAL_ADMIN_PASSWORD.")
        else:
            admin_user, _, _, _ = initial_admin_config()
            st.write(f"Usuario inicial: `{admin_user}`")
            st.write("Contraseña inicial: definida en secretos o `Admin123*`.")
        st.warning("Cambie la contraseña apenas ingrese.")


def ui_mi_cuenta(st: Any):
    user = st.session_state.get("auth_user", {})
    st.header("Mi cuenta")
    st.write(f"**Nombre:** {user.get('nombre_completo', '')}")
    st.write(f"**Usuario:** {user.get('usuario', '')}")
    st.write(f"**Correo:** {user.get('email', '')}")
    st.write(f"**Perfil:** {user.get('rol', '')}")

    with st.form("form_cambiar_clave"):
        nueva = st.text_input("Nueva contraseña", type="password")
        confirmar = st.text_input("Confirmar contraseña", type="password")
        enviar = st.form_submit_button("Actualizar contraseña", use_container_width=True)

    if enviar:
        if nueva != confirmar:
            st.error("Las contraseñas no coinciden.")
        else:
            try:
                cambiar_password(int(user["id"]), nueva, forzar_cambio=False)
                st.session_state["auth_user"]["debe_cambiar_clave"] = False
                registrar_auditoria("Cambio de contraseña", "Contraseña propia actualizada", user=user)
                st.success("Contraseña actualizada correctamente.")
            except Exception as e:
                st.error(str(e))


def ui_admin_usuarios(st: Any):
    st.header("Gestión de usuarios y perfiles")
    st.caption("Administra accesos y asigna roles del sistema.")

    df_users = listar_usuarios()
    st.dataframe(df_users, use_container_width=True, hide_index=True)

    with st.expander("Crear nuevo usuario"):
        with st.form("crear_usuario_form"):
            nu = st.text_input("Usuario")
            nom = st.text_input("Nombre completo")
            em = st.text_input("Correo electrónico")
            rol = st.selectbox("Perfil", list(ROLES_PERMISOS.keys()), index=2)
            pw = st.text_input("Contraseña inicial", type="password")
            crear = st.form_submit_button("Crear usuario", use_container_width=True)
        if crear:
            try:
                crear_usuario(nu, nom, em, rol, pw, debe_cambiar=True)
                registrar_auditoria("Crear usuario", f"Usuario {nu} creado con rol {rol}")
                st.success(f"Usuario {nu} creado exitosamente.")
                st.rerun()
            except Exception as e:
                st.error(str(e))
