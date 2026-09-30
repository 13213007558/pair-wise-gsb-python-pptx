"""Support for moving shapes between shape trees without changing their appearance.

A shape's position on the slide is determined by its own `a:xfrm` composed with the
transforms of any group shapes that contain it. Group transforms contribute offset,
scaling (`chOff`/`chExt` mapped to `off`/`ext`), rotation, and flips. When a
shape is moved to a different container (reparented), its own transform must be
rewritten so the composed result is unchanged.

A shape transform can only express translation, scaling, rotation, and axis-aligned
flips. When the required composed transform includes shear or degenerate (zero)
scaling, the shape's geometry cannot be reproduced exactly and
|UnrepresentableTransformError| is raised rather than silently approximating.
"""

from __future__ import annotations

import math
from typing import TYPE_CHECKING, NamedTuple

from pptx.exc import UnrepresentableTransformError
from pptx.oxml.ns import qn
from pptx.util import Emu

if TYPE_CHECKING:
    from pptx.oxml.shapes import ShapeElement
    from pptx.oxml.shapes.groupshape import CT_GroupShape


class XfrmValues(NamedTuple):
    """New `a:xfrm` attribute values for a reparented shape."""

    x: int
    y: int
    cx: int
    cy: int
    rot: float
    flip_h: bool
    flip_v: bool


class _AffineMap:
    """2-D affine transform `p -> M @ p + t` operating on EMU coordinates."""

    __slots__ = ("m00", "m01", "m10", "m11", "tx", "ty")

    def __init__(
        self,
        m00: float = 1.0,
        m01: float = 0.0,
        m10: float = 0.0,
        m11: float = 1.0,
        tx: float = 0.0,
        ty: float = 0.0,
    ):
        self.m00, self.m01, self.m10, self.m11 = m00, m01, m10, m11
        self.tx, self.ty = tx, ty

    def applied_to(self, x: float, y: float) -> tuple[float, float]:
        """Return the image of point (`x`, `y`) under this transform."""
        return (
            self.m00 * x + self.m01 * y + self.tx,
            self.m10 * x + self.m11 * y + self.ty,
        )

    def composed_with(self, other: "_AffineMap") -> "_AffineMap":
        """Return the map equivalent to applying `other` first, then this map."""
        return _AffineMap(
            self.m00 * other.m00 + self.m01 * other.m10,
            self.m00 * other.m01 + self.m01 * other.m11,
            self.m10 * other.m00 + self.m11 * other.m10,
            self.m10 * other.m01 + self.m11 * other.m11,
            self.m00 * other.tx + self.m01 * other.ty + self.tx,
            self.m10 * other.tx + self.m11 * other.ty + self.ty,
        )

    @property
    def inverse(self) -> "_AffineMap":
        """The inverse of this map.

        Raises |UnrepresentableTransformError| if the map is degenerate (singular),
        which occurs for example when a group in the chain has a zero `chExt`.
        """
        det = self.m00 * self.m11 - self.m01 * self.m10
        if abs(det) < 1e-12:
            raise UnrepresentableTransformError(
                "a group in the destination hierarchy has a degenerate (zero-size)"
                " transform; its coordinate mapping cannot be inverted"
            )
        return _AffineMap(
            self.m11 / det,
            -self.m01 / det,
            -self.m10 / det,
            self.m00 / det,
            (self.m01 * self.ty - self.m11 * self.tx) / det,
            (self.m10 * self.tx - self.m00 * self.ty) / det,
        )


def new_xfrm_values(elm: ShapeElement, dest_container_elm: CT_GroupShape) -> XfrmValues:
    """Compute `a:xfrm` values for `elm` preserving its page geometry as a child of
    `dest_container_elm`.

    Raises |UnrepresentableTransformError| if the required transform cannot be
    expressed with offset, extents, rotation, and flips (i.e. it includes shear or
    degenerate scaling). Does not modify the XML.
    """
    page_map = _container_map(elm.getparent()).composed_with(_element_map(elm))
    new_map = _container_map(dest_container_elm).inverse.composed_with(page_map)
    return _decompose(elm, new_map)


def apply_xfrm_values(elm: ShapeElement, values: XfrmValues) -> None:
    """Write `values` to the `a:xfrm` of `elm` (created when not already present).

    For a group shape, `chOff` and `chExt` are left untouched so the coordinates
    of its member shapes remain valid.
    """
    xfrm = elm.get_or_add_xfrm()
    xfrm.x = Emu(values.x)
    xfrm.y = Emu(values.y)
    xfrm.cx = Emu(values.cx)
    xfrm.cy = Emu(values.cy)
    xfrm.rot = values.rot
    xfrm.flipH = values.flip_h
    xfrm.flipV = values.flip_v


def reparent_element(elm: ShapeElement, dest_container_elm: CT_GroupShape) -> None:
    """Move `elm` into `dest_container_elm`, preserving its on-slide geometry."""
    apply_xfrm_values(elm, new_xfrm_values(elm, dest_container_elm))
    dest_container_elm.insert_element_before(elm, "p:extLst")


def _element_map(elm: ShapeElement) -> _AffineMap:
    """Affine map from the shape's content coordinates to its parent's coordinates.

    For a group shape, content coordinates are child-shape coordinates, mapped through
    `chOff`/`chExt` onto `off`/`ext`. For other shapes, content coordinates are
    local coordinates in the range `(0, 0)` to `(cx, cy)`. Flips and rotation are
    applied about the center of the `off`/`ext` bounding box.
    """
    xfrm = elm.xfrm
    if xfrm is None:
        return _AffineMap()

    ox, oy = float(xfrm.x or 0), float(xfrm.y or 0)
    cx, cy = float(xfrm.cx or 0), float(xfrm.cy or 0)
    rot = float(xfrm.rot or 0.0)
    flip_h, flip_v = bool(xfrm.flipH), bool(xfrm.flipV)

    if elm.tag == qn("p:grpSp"):
        ch_off, ch_ext = xfrm.chOff, xfrm.chExt
        ch_ox, ch_oy = (float(ch_off.x), float(ch_off.y)) if ch_off is not None else (0.0, 0.0)
        ch_cx, ch_cy = (float(ch_ext.cx), float(ch_ext.cy)) if ch_ext is not None else (cx, cy)
        if ch_cx == 0.0 or ch_cy == 0.0:
            raise UnrepresentableTransformError(
                "group has a zero-size chExt; its child-coordinate mapping is degenerate"
            )
        sx, sy = cx / ch_cx, cy / ch_cy
    else:
        ch_ox = ch_oy = 0.0
        sx = sy = 1.0

    theta = math.radians(rot)
    cos, sin = math.cos(theta), math.sin(theta)
    fh = -1.0 if flip_h else 1.0
    fv = -1.0 if flip_v else 1.0
    # --- K = R(rot) @ F, the orthogonal (rotation + flip) part of the transform ---
    k00, k01 = cos * fh, -sin * fv
    k10, k11 = sin * fh, cos * fv

    cen_x, cen_y = ox + cx / 2.0, oy + cy / 2.0
    # --- translation: K @ (off - S @ chOff - center) + center ---
    vx, vy = ox - sx * ch_ox - cen_x, oy - sy * ch_oy - cen_y
    return _AffineMap(
        k00 * sx,
        k01 * sy,
        k10 * sx,
        k11 * sy,
        k00 * vx + k01 * vy + cen_x,
        k10 * vx + k11 * vy + cen_y,
    )


def _container_map(container_elm) -> _AffineMap:
    """Affine map from the child-coordinate space of `container_elm` to page coordinates.

    Composes the transforms of `container_elm` (when it is itself a group shape) and
    all its ancestor group shapes, innermost first.
    """
    result = _AffineMap()
    node = container_elm
    while node is not None and node.tag == qn("p:grpSp"):
        result = _element_map(node).composed_with(result)
        node = node.getparent()
    return result


def _decompose(elm: ShapeElement, new_map: _AffineMap) -> XfrmValues:
    """Express `new_map` as `a:xfrm` values for `elm`.

    The linear part of `new_map` must decompose as `K @ S` with `K` orthogonal
    (rotation + flips) and `S` a positive diagonal (scaling absorbed into the shape's
    extents). Raises |UnrepresentableTransformError| when the columns of the linear
    part are not orthogonal, i.e. the transform includes shear.
    """
    is_group = elm.tag == qn("p:grpSp")

    # --- content-space origin and size; these are preserved for a group so its
    # --- member shapes keep their coordinates ---
    if is_group:
        xfrm = elm.xfrm
        ch_off = xfrm.chOff if xfrm is not None else None
        ch_ext = xfrm.chExt if xfrm is not None else None
        oc_x = float(ch_off.x) if ch_off is not None else 0.0
        oc_y = float(ch_off.y) if ch_off is not None else 0.0
        ec_x = float(ch_ext.cx) if ch_ext is not None else 0.0
        ec_y = float(ch_ext.cy) if ch_ext is not None else 0.0
    else:
        xfrm = elm.xfrm
        oc_x = oc_y = 0.0
        ec_x = float(xfrm.cx) if xfrm is not None and xfrm.cx is not None else 0.0
        ec_y = float(xfrm.cy) if xfrm is not None and xfrm.cy is not None else 0.0

    d1 = math.hypot(new_map.m00, new_map.m10)
    d2 = math.hypot(new_map.m01, new_map.m11)
    if d1 < 1e-12 or d2 < 1e-12:
        raise UnrepresentableTransformError(
            "the composed transform collapses the shape to zero size and cannot be" " represented"
        )
    dot = new_map.m00 * new_map.m01 + new_map.m10 * new_map.m11
    if abs(dot) > 1e-9 * d1 * d2:
        raise UnrepresentableTransformError(
            "the composed transform includes shear, which cannot be represented by a"
            " shape's offset, extents, rotation, and flips; refusing to approximate"
        )

    # --- K_full = M @ S^-1, the orthogonal rotation+flip matrix ---
    kf00, kf01 = new_map.m00 / d1, new_map.m01 / d2
    kf10, kf11 = new_map.m10 / d1, new_map.m11 / d2

    det = kf00 * kf11 - kf01 * kf10
    flip_h = det < 0.0
    # --- rotation matrix R = K_full @ F^-1 with F canonicalized to flipH ---
    r00, r10 = (-kf00, -kf10) if flip_h else (kf00, kf10)
    rot = math.degrees(math.atan2(r10, r00)) % 360.0

    new_cx, new_cy = d1 * ec_x, d2 * ec_y
    # --- off' = t + K_full @ (S @ chOff) + (K_full - I) @ (ext' / 2) ---
    so_x, so_y = d1 * oc_x, d2 * oc_y
    ex, ey = new_cx / 2.0, new_cy / 2.0
    new_x = new_map.tx + (kf00 * so_x + kf01 * so_y) + (kf00 - 1.0) * ex + kf01 * ey
    new_y = new_map.ty + (kf10 * so_x + kf11 * so_y) + kf10 * ex + (kf11 - 1.0) * ey

    return XfrmValues(
        x=int(round(new_x)),
        y=int(round(new_y)),
        cx=max(0, int(round(new_cx))),
        cy=max(0, int(round(new_cy))),
        rot=rot,
        flip_h=flip_h,
        flip_v=False,
    )
