from __future__ import annotations

import argparse
import os
from pathlib import Path

from .auth import DeviceRegistry
from .server import ROOT, _load_env


def registry() -> DeviceRegistry:
    _load_env(ROOT / "lugest.env")
    configured = str(os.environ.get("LUGEST_MOBILE_DEVICE_STORE", "") or "").strip()
    path = Path(configured) if configured else ROOT / "mobile_gateway_data" / "devices.json"
    return DeviceRegistry(path.resolve())


def main() -> None:
    parser = argparse.ArgumentParser(description="Gerir dispositivos LuGEST Field")
    commands = parser.add_subparsers(dest="command", required=True)
    create = commands.add_parser("create")
    create.add_argument("--name", required=True)
    revoke = commands.add_parser("revoke")
    revoke.add_argument("device_id")
    rotate = commands.add_parser("rotate")
    rotate.add_argument("device_id")
    imported = commands.add_parser("import-legacy")
    imported.add_argument("--name", default="Telemóvel principal")
    commands.add_parser("list")
    args = parser.parse_args()

    store = registry()
    if args.command == "create":
        device, token = store.create(args.name)
        print(f"Dispositivo: {device['name']} ({device['id']})")
        print(f"Chave: {token}")
        print("Guarda esta chave agora; o servidor conserva apenas o hash.")
    elif args.command == "revoke":
        device = store.revoke(args.device_id)
        print(f"Dispositivo revogado: {device['name']} ({device['id']})")
    elif args.command == "rotate":
        device, token = store.rotate(args.device_id)
        print(f"Nova chave para {device['name']} ({device['id']}):")
        print(token)
    elif args.command == "import-legacy":
        token = str(os.environ.get("LUGEST_MOBILE_API_TOKEN", "") or "").strip()
        if not token:
            raise SystemExit("LUGEST_MOBILE_API_TOKEN não está definido.")
        device = store.import_token(args.name, token)
        print(f"Chave existente associada a: {device['name']} ({device['id']})")
    else:
        rows = store.list()
        if not rows:
            print("Nenhum dispositivo registado.")
        for row in rows:
            state = "ativo" if row.get("enabled") else "revogado"
            print(f"{row.get('id')} | {row.get('name')} | {state} | último acesso: {row.get('last_seen_at') or '-'}")


if __name__ == "__main__":
    main()
