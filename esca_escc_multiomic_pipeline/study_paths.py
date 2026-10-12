"""Resolve study / repo paths for scripts (GitHub clone vs full local workspace)."""
from __future__ import annotations

import os
from pathlib import Path

_PIPELINE = Path(__file__).resolve().parent
_REPO_ROOT = _PIPELINE.parent


def pipeline_dir() -> Path:
    return Path(os.environ.get("ESCA_PIPELINE_DIR", _PIPELINE))


def repo_root() -> Path:
    return _REPO_ROOT


def study_root() -> Path:
    """Full TCGA ESCA workspace (contains Cancer Medicine submission_package)."""
    if env := os.environ.get("ESCA_STUDY_ROOT"):
        return Path(env)
    # Default: parent of code/repo when nested under the study folder.
    candidate = _REPO_ROOT.parent.parent
    if (candidate / "Cancer Medicine submission_package").is_dir():
        return candidate
    return _REPO_ROOT


def submission_package_dir() -> Path:
    if env := os.environ.get("CM_SUBMISSION_PACKAGE"):
        return Path(env)
    return study_root() / "Cancer Medicine submission_package"
