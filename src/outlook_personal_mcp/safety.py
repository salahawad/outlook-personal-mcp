from __future__ import annotations

import os
from pathlib import Path
from urllib.parse import quote


def graph_id(value: str, *, name: str = "id") -> str:
    value = str(value)
    if not value:
        raise ValueError(f"{name} is required")
    return quote(value, safe="")


def _root(settings) -> Path:
    return Path(settings.file_root).expanduser().resolve(strict=False)


def _ensure_under_root(path: Path, root: Path, *, label: str) -> None:
    try:
        path.relative_to(root)
    except ValueError as exc:
        raise ValueError(f"{label} must be under OUTLOOK_MCP_FILE_ROOT ({root})") from exc


def _reject_symlink_parts(path: Path, root: Path, *, label: str) -> None:
    current = path
    while True:
        if current.exists() and current.is_symlink():
            raise ValueError(f"{label} must not include symlinks")
        if current == root:
            return
        if current.parent == current:
            return
        current = current.parent


def _raw_under_root(settings, value: str, *, label: str) -> tuple[Path, Path]:
    root = _root(settings)
    requested = Path(value).expanduser()
    raw = requested if requested.is_absolute() else root / requested
    raw_abs = Path(os.path.abspath(raw))
    _ensure_under_root(raw_abs, root, label=label)
    return root, raw_abs


def resolve_read_file(settings, value: str) -> Path:
    root, raw_abs = _raw_under_root(settings, value, label="file_path")
    _reject_symlink_parts(raw_abs, root, label="file_path")
    resolved = raw_abs.resolve(strict=True)
    _ensure_under_root(resolved, root, label="file_path")
    if not resolved.is_file():
        raise ValueError("file_path must point to a regular file")
    size = resolved.stat().st_size
    if size > settings.max_file_bytes:
        raise ValueError(
            f"file_path exceeds OUTLOOK_MCP_MAX_FILE_BYTES ({settings.max_file_bytes})"
        )
    return resolved


def resolve_write_file(settings, value: str) -> Path:
    root, raw_abs = _raw_under_root(settings, value, label="save_path")
    if raw_abs.exists():
        raise FileExistsError(f"save_path already exists: {raw_abs}")
    raw_abs.parent.mkdir(parents=True, exist_ok=True)
    _reject_symlink_parts(raw_abs.parent, root, label="save_path")
    resolved_parent = raw_abs.parent.resolve(strict=True)
    _ensure_under_root(resolved_parent, root, label="save_path")
    return raw_abs


def ensure_size_within_limit(settings, data: bytes, *, label: str) -> None:
    if len(data) > settings.max_file_bytes:
        raise ValueError(f"{label} exceeds OUTLOOK_MCP_MAX_FILE_BYTES ({settings.max_file_bytes})")
