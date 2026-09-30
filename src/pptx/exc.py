"""Exceptions used with python-pptx.

The base exception class is PythonPptxError.
"""

from __future__ import annotations


class PythonPptxError(Exception):
    """Generic error class."""


class PackageNotFoundError(PythonPptxError):
    """
    Raised when a package cannot be found at the specified path.
    """


class InvalidXmlError(PythonPptxError):
    """
    Raised when a value is encountered in the XML that is not valid according
    to the schema.
    """


class UnrepresentableTransformError(PythonPptxError, ValueError):
    """Raised when a shape cannot be reparented without altering its appearance.

    The transform required to preserve the shape's on-slide geometry when moving
    it to a new container includes shear or degenerate (zero) scaling, which cannot
    be expressed by a shape's `a:xfrm` (offset, extents, rotation, and axis-aligned
    flips only). Raised instead of silently approximating the shape's geometry.
    """
