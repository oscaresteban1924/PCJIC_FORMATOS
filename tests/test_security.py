import pytest
from src.security import hash_password, verificar_password, tiene_permiso, ROLES_PERMISOS


def test_hash_and_verify_password():
    password = "ClaveDePrueba123*"
    salt, p_hash = hash_password(password)
    assert salt is not None and len(salt) > 0
    assert p_hash is not None and len(p_hash) > 0
    assert verificar_password(password, salt, p_hash) is True
    assert verificar_password("ClaveInvalida", salt, p_hash) is False


def test_tiene_permiso():
    admin_user = {"usuario": "admin", "rol": "Administrador"}
    docente_user = {"usuario": "docente", "rol": "Docente"}

    assert tiene_permiso("FD-GC71 - Planeación", admin_user) is True
    assert tiene_permiso("Usuarios y perfiles", admin_user) is True
    assert tiene_permiso("Usuarios y perfiles", docente_user) is False
