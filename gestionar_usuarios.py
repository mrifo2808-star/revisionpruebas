"""
CLI local de administración de cuentas para la plataforma "Herramientas
Docentes" -- NO se despliega, corre solo en la máquina del administrador.
El grupo de usuarios es cerrado a propósito (sin auto-registro público):
las cuentas se precargan acá, una por una.

Requiere el mismo .streamlit/secrets.toml que usa `streamlit run
app_revisor.py` en local (ver registro_sheets.py para el formato exacto de
credenciales y los encabezados que debe tener la planilla).

Uso:
    py gestionar_usuarios.py --agregar ana --nombre "Ana Soto" --rol docente
    py gestionar_usuarios.py --listar
    py gestionar_usuarios.py --desactivar ana
"""
import argparse
import getpass
import sys

import registro_sheets as rs


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    grupo = ap.add_mutually_exclusive_group(required=True)
    grupo.add_argument("--agregar", metavar="USUARIO", help="crea una cuenta nueva")
    grupo.add_argument("--listar", action="store_true", help="lista las cuentas existentes")
    grupo.add_argument("--desactivar", metavar="USUARIO", help="revoca el acceso sin borrar historial")
    ap.add_argument("--nombre", default=None, help="nombre para mostrar (solo con --agregar)")
    ap.add_argument("--rol", default="docente", help="rol de la cuenta (solo con --agregar, default: docente)")
    args = ap.parse_args()

    if args.agregar:
        nombre = args.nombre or args.agregar
        password = getpass.getpass(f"Clave para '{args.agregar}': ")
        confirmacion = getpass.getpass("Repite la clave: ")
        if password != confirmacion:
            print("Las claves no coinciden -- no se creó la cuenta.", file=sys.stderr)
            sys.exit(1)
        if len(password) < 8:
            print("La clave debe tener al menos 8 caracteres.", file=sys.stderr)
            sys.exit(1)
        rs.agregar_usuario(args.agregar, password, nombre, args.rol)
        print(f"Cuenta '{args.agregar}' creada (rol: {args.rol}).")

    elif args.listar:
        usuarios = rs.listar_usuarios()
        if not usuarios:
            print("No hay cuentas registradas todavía.")
            return
        print(f"{'usuario':<20}{'nombre':<25}{'rol':<12}{'activo'}")
        for u in usuarios:
            print(f"{str(u['usuario']):<20}{str(u['nombre']):<25}{str(u['rol']):<12}{u['activo']}")

    elif args.desactivar:
        if rs.desactivar_usuario(args.desactivar):
            print(f"Cuenta '{args.desactivar}' desactivada.")
        else:
            print(f"No se encontró la cuenta '{args.desactivar}'.", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()
