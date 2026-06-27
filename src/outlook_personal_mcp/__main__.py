from __future__ import annotations
import sys
from .config import load_settings
from .auth import TokenProvider
from .server import build_server


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    cmd = argv[0] if argv else "serve"
    settings = load_settings()
    if cmd == "login":
        TokenProvider(settings).login()
        print("Login successful; token cached.", file=sys.stderr)
        return 0
    if cmd in ("serve", ""):
        build_server(settings).run()
        return 0
    print(f"Unknown command: {cmd}. Use 'serve' or 'login'.", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
