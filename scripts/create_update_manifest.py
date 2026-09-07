from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_core.updates import sha256_file, validate_update_manifest
from lugest_infra.config import AtomicJsonStore


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Cria um manifesto de atualização luGEST com hashes SHA-256 obrigatórios."
    )
    parser.add_argument("--version", required=True)
    parser.add_argument("--package-file", type=Path, required=True)
    parser.add_argument("--package-url", required=True)
    parser.add_argument("--bootstrap-file", type=Path, required=True)
    parser.add_argument("--bootstrap-url", required=True)
    parser.add_argument("--channel", choices=("stable", "beta", "pilot"), default="stable")
    parser.add_argument("--notes", default="")
    parser.add_argument("--output", type=Path, required=True)
    return parser.parse_args()


def build_manifest(args: argparse.Namespace) -> dict[str, object]:
    package_file = args.package_file.resolve()
    bootstrap_file = args.bootstrap_file.resolve()
    if not package_file.is_file():
        raise FileNotFoundError(f"Pacote não encontrado: {package_file}")
    if not bootstrap_file.is_file():
        raise FileNotFoundError(f"Reparador não encontrado: {bootstrap_file}")
    payload: dict[str, object] = {
        "schema_version": 1,
        "version": str(args.version).strip(),
        "channel": args.channel,
        "package_url": str(args.package_url).strip(),
        "sha256": sha256_file(package_file),
        "bootstrap_url": str(args.bootstrap_url).strip(),
        "bootstrap_sha256": sha256_file(bootstrap_file),
        "notes": str(args.notes or "").strip(),
        "created_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }
    validate_update_manifest(payload)
    return payload


def main() -> int:
    args = parse_args()
    payload = build_manifest(args)
    output = args.output.resolve()
    AtomicJsonStore(output, max_bytes=1024 * 1024).save(payload)
    print(f"Manifesto criado: {output}")
    print(f"Pacote SHA-256: {payload['sha256']}")
    print(f"Reparador SHA-256: {payload['bootstrap_sha256']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
