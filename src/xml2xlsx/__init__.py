# Released under GPL3 License.
# Copyright (c) 2026 Ladislav Bartos and Robert Vacha Lab

from .core import Summary, convert, read_plates

try:
    from ._version import __version__
except ImportError:
    __version__ = "0+unknown"

__all__ = ["Summary", "convert", "read_plates", "__version__"]
