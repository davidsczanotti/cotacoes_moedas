from __future__ import annotations

import os
from pathlib import Path
import shutil
import uuid

import pytest


@pytest.fixture
def tmp_path() -> Path:
    """
    Custom tmp_path to avoid os.mkdir(path, mode=0o700) on Windows.

    In some locked-down environments, creating directories with an explicit
    `mode` can result in unreadable directories. We create temp dirs using
    os.mkdir() without the mode argument and keep them under the repo.
    """

    base = Path.cwd() / ".pytest-tmp"
    if not base.exists():
        os.mkdir(base)

    path = base / f"tmp-{uuid.uuid4().hex}"
    os.mkdir(path)
    try:
        yield path
    finally:
        shutil.rmtree(path, ignore_errors=True)
