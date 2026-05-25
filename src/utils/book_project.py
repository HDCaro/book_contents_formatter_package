import json
import os
from pathlib import Path


def find_project_root(anchor_path=None):
    current = Path(anchor_path or __file__).resolve()
    start = current if current.is_dir() else current.parent

    for parent in [start] + list(start.parents):
        if (parent / "src").exists():
            return parent

    raise RuntimeError("Project root not found")


def load_book_project(project_root=None):
    project_root = Path(project_root) if project_root else find_project_root()
    config_path = project_root / "books" / "book_project.json"

    config = {}
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as handle:
            config = json.load(handle)

    active_book = os.environ.get("BOOK_SLUG") or config.get("active_book") or "hits_and_happiness"
    books_dir = Path(config.get("books_dir", "books"))
    if not books_dir.is_absolute():
        books_dir = project_root / books_dir

    return {
        "project_root": project_root,
        "config_path": config_path,
        "books_dir": books_dir,
        "active_book": active_book,
        "book_root": books_dir / active_book,
    }


def get_active_book_root(project_root=None):
    return load_book_project(project_root)["book_root"]


def resolve_book_path(book_root, configured_path):
    path = Path(configured_path)
    if path.is_absolute():
        return path
    return Path(book_root) / path
