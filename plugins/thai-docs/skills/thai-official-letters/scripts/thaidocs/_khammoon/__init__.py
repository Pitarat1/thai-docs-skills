# Vendored from khammoon 0.1.0 (https://github.com/warit-s/khammoon-segmenter, PyPI khammoon==0.1.0)
# Original path: khammoon/__init__.py
# Copyright 2026 The khammoon authors. License: Apache-2.0 (code)
# Model files in ./model/ are licensed separately: CC BY-NC-SA 4.0 (see model/LICENSE)
# MODIFIED by thaidocs: reduced to the numpy backend; module-level segment() helper not re-exported
# SPDX-License-Identifier: Apache-2.0
"""khammoon — Thai morpheme (คำมูล) segmentation with a small CPU model (numpy backend only)."""
from .segmenter import MODEL_DIR, Segmenter

__version__ = "0.1.0"
__all__ = ["MODEL_DIR", "Segmenter", "__version__"]
