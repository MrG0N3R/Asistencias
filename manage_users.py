"""Administración local: python manage_users.py --help."""

import argparse
from contextlib import closing
from getpass import getpass

import psycopg2

from hr_dashboard.auth_store import create_user, hash_password, initialize_database, normalize_username
from hr_dashboard.database import DatabaseConfigurationError, get_connection


def read_password():
    password = getpass("Contraseña (mínimo 12 caracteres): ")
    if password != getpass("Repite la contraseña: "):
        raise ValueError("Las contraseñas no coinciden.")
    return password


def main():
    parser = argparse.ArgumentParser(description="Administrar accesos de Dashboard Checadores")
    commands = parser.add_subparsers(dest="command", required=True)
    commands.add_parser("init-db", help="Crear la tabla de usuarios si no existe")
    create = commands.add_parser("create-user", help="Crear un usuario; pide la contraseña oculta")
    create.add_argument("username")
    create.add_argument("--name", required=True, help="Nombre que aparece en la aplicación")
    for name, help_text in (
        ("reset-password", "Cambiar la contraseña y quitar el bloqueo temporal"),
        ("disable-user", "Desactivar el acceso de un usuario"),
        ("enable-user", "Reactivar el acceso de un usuario"),
    ):
        command = commands.add_parser(name, help=help_text)
        command.add_argument("username")
    args = parser.parse_args()
    try:
        if args.command == "init-db":
            initialize_database()
            print("Tabla de usuarios lista.")
        elif args.command == "create-user":
            password = read_password()
            initialize_database()
            create_user(args.username, args.name, password)
            print("Usuario creado.")
        else:
            encoded = hash_password(read_password()) if args.command == "reset-password" else None
            with closing(get_connection()) as conn, conn, conn.cursor() as cur:
                if encoded is not None:
                    cur.execute("""
                        UPDATE public.checadores_users
                        SET password_hash = %s, failed_attempts = 0, locked_until = NULL
                        WHERE username = %s
                    """, (encoded, normalize_username(args.username)))
                else:
                    cur.execute("""
                        UPDATE public.checadores_users
                        SET is_active = %s, failed_attempts = 0, locked_until = NULL
                        WHERE username = %s
                    """, (args.command == "enable-user", normalize_username(args.username)))
                if not cur.rowcount:
                    raise ValueError("No se encontró el usuario.")
            print("Usuario actualizado.")
    except psycopg2.errors.UniqueViolation:
        parser.exit(1, "Ese usuario ya existe. Usa otro nombre o reset-password.\n")
    except (ValueError, DatabaseConfigurationError) as exc:
        parser.exit(1, f"{exc}\n")
    except psycopg2.Error:
        parser.exit(1, "No se pudo acceder a PostgreSQL. Revisa la conexión y ejecuta init-db.\n")


if __name__ == "__main__":
    main()
