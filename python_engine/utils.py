import re
from pathlib import Path
from uuid import uuid4


def ensure_directory(path: str | Path) -> Path:
    directory = Path(path)
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def safe_slug(value: str, fallback: str = "file") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    cleaned = cleaned.strip("._-")
    return cleaned or fallback


def unique_output_path(directory: str | Path, filename: str) -> str:
    directory_path = ensure_directory(directory)
    source = Path(filename)
    stem = safe_slug(source.stem, "output")
    suffix = source.suffix or ".dat"
    return str(directory_path / f"{stem}_{uuid4().hex[:8]}{suffix}")


def is_within_directory(path: str | Path, directory: str | Path) -> bool:
    candidate = Path(path).resolve()
    root = Path(directory).resolve()
    try:
        candidate.relative_to(root)
        return True
    except ValueError:
        return False
