"""Test suite for geometry-preserving reparenting of shapes between containers."""

from __future__ import annotations

import math

import pytest

from pptx import Presentation
from pptx.enum.shapes import MSO_CONNECTOR_TYPE, MSO_SHAPE
from pptx.exc import UnrepresentableTransformError
from pptx.oxml.ns import qn
from pptx.util import Emu


def page_point(elm, pt):
    """Map content-local point `pt` of `elm` to slide (page) coordinates.

    Independent reference implementation: walks the XML tree applying each
    element's xfrm (off/ext, chOff/chExt scaling, flips, rotation) in turn.
    """
    x, y = pt
    node = elm
    shape_tags = (qn("p:grpSp"), qn("p:sp"), qn("p:cxnSp"), qn("p:pic"), qn("p:graphicFrame"))
    while node is not None and node.tag in shape_tags:
        xfrm = node.xfrm
        ox, oy = float(xfrm.x or 0), float(xfrm.y or 0)
        cx, cy = float(xfrm.cx or 0), float(xfrm.cy or 0)
        rot = float(xfrm.rot or 0.0)
        flip_h, flip_v = bool(xfrm.flipH), bool(xfrm.flipV)
        if node.tag == qn("p:grpSp"):
            ch_off, ch_ext = xfrm.chOff, xfrm.chExt
            sx = cx / float(ch_ext.cx)
            sy = cy / float(ch_ext.cy)
            x = ox + (x - float(ch_off.x)) * sx
            y = oy + (y - float(ch_off.y)) * sy
        else:
            x += ox
            y += oy
        cen_x, cen_y = ox + cx / 2.0, oy + cy / 2.0
        if flip_h:
            x = 2 * cen_x - x
        if flip_v:
            y = 2 * cen_y - y
        theta = math.radians(rot)
        dx, dy = x - cen_x, y - cen_y
        x = cen_x + dx * math.cos(theta) - dy * math.sin(theta)
        y = cen_y + dx * math.sin(theta) + dy * math.cos(theta)
        node = node.getparent()
    return x, y


def page_corners(shape):
    """Page coordinates of the four content-bbox corners of `shape`."""
    elm = shape.element
    if elm.tag == qn("p:grpSp"):
        xfrm = elm.grpSpPr.get_or_add_xfrm()
        ox, oy = float(xfrm.chOff.x), float(xfrm.chOff.y)
        cx, cy = float(xfrm.chExt.cx), float(xfrm.chExt.cy)
    else:
        ox, oy = 0.0, 0.0
        cx, cy = float(elm.cx), float(elm.cy)
    corners = [(ox, oy), (ox + cx, oy), (ox, oy + cy), (ox + cx, oy + cy)]
    return [page_point(elm, corner) for corner in corners]


def assert_same_geometry(before, after, tolerance=2.0):
    """Assert each page point in `before` matches the corresponding one in `after`."""
    assert len(before) == len(after)
    for (bx, by), (ax, ay) in zip(before, after):
        assert math.hypot(bx - ax, by - ay) < tolerance


def set_group_xfrm(group, off, ext, ch_off, ch_ext, rot=0.0, flip_h=False, flip_v=False):
    """Directly set the full transform of `group`, including chOff/chExt scaling."""
    xfrm = group.element.grpSpPr.get_or_add_xfrm()
    xfrm.x, xfrm.y = Emu(off[0]), Emu(off[1])
    xfrm.cx, xfrm.cy = Emu(ext[0]), Emu(ext[1])
    xfrm.chOff.x, xfrm.chOff.y = Emu(ch_off[0]), Emu(ch_off[1])
    xfrm.chExt.cx, xfrm.chExt.cy = Emu(ch_ext[0]), Emu(ch_ext[1])
    xfrm.rot = rot
    xfrm.flipH = flip_h
    xfrm.flipV = flip_v


@pytest.fixture
def slide():
    prs = Presentation()
    return prs.slides.add_slide(prs.slide_layouts[6])


@pytest.fixture
def nested(slide):
    """Two nested groups collectively having non-uniform scaling and rotation.

    Outer group is rotated 30 degrees, flipped horizontally, and scaled (2.0, 0.5);
    inner group is flipped vertically and scaled (1.5, 0.75). Returns
    `(slide, outer, inner, member)` where `member` is a shape in the inner group.
    """
    shapes = slide.shapes
    member = shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Emu(300000), Emu(400000), Emu(500000), Emu(200000)
    )
    other = shapes.add_shape(MSO_SHAPE.OVAL, Emu(900000), Emu(100000), Emu(300000), Emu(600000))
    outer = shapes.add_group_shape([member, other])
    set_group_xfrm(
        outer,
        off=(1000000, 800000),
        ext=(2000000, 500000),
        ch_off=(0, 0),
        ch_ext=(1000000, 1000000),
        rot=30.0,
        flip_h=True,
    )
    inner = outer.shapes.add_group_shape([member])
    set_group_xfrm(
        inner,
        off=(100000, 50000),
        ext=(600000, 300000),
        ch_off=(0, 0),
        ch_ext=(400000, 400000),
        flip_v=True,
    )
    return slide, outer, inner, member


class DescribeReparent(object):

    def it_preserves_page_geometry_through_nested_scaled_rotated_groups(self, nested):
        slide, outer, inner, member = nested
        before = page_corners(member)

        slide.shapes.reparent(member)

        assert member.element.getparent() is slide.shapes._spTree
        assert_same_geometry(before, page_corners(member))

    def it_preserves_page_geometry_when_moving_a_group(self, nested):
        slide, outer, inner, member = nested
        before = page_corners(member)

        slide.shapes.reparent(inner)

        assert inner.element.getparent() is slide.shapes._spTree
        assert_same_geometry(before, page_corners(member))

    def it_preserves_page_geometry_when_moving_into_a_group(self, nested):
        slide, outer, inner, member = nested
        before = page_corners(member)

        slide.shapes.reparent(member)
        inner.shapes.reparent(member)

        assert member.element.getparent() is inner.element
        assert_same_geometry(before, page_corners(member))

    def it_survives_repeated_moves_in_and_out(self, nested):
        slide, outer, inner, member = nested
        before = page_corners(member)

        for _ in range(3):
            slide.shapes.reparent(member)
            assert_same_geometry(before, page_corners(member))
            inner.shapes.reparent(member)
            assert_same_geometry(before, page_corners(member))

    def it_does_not_disturb_remaining_members_of_the_source_group(self, nested):
        slide, outer, inner, member = nested
        sibling = outer.shapes[0]
        sibling_before = page_corners(sibling)
        outer_xfrm_before = outer.element.grpSpPr.get_or_add_xfrm().xml
        sibling_xfrm_before = sibling.element.get_or_add_xfrm().xml

        slide.shapes.reparent(member)

        # ---remaining member keeps identical stored coordinates and page geometry---
        assert sibling.element.get_or_add_xfrm().xml == sibling_xfrm_before
        assert_same_geometry(sibling_before, page_corners(sibling))
        # ---source group extents are not recalculated---
        assert outer.element.grpSpPr.get_or_add_xfrm().xml == outer_xfrm_before

    def it_rejects_shear_transforms_without_changing_anything(self, slide):
        shapes = slide.shapes
        member = shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Emu(300000), Emu(400000), Emu(500000), Emu(200000)
        )
        # ---non-uniformly scaled group containing a rotated shape: the composed
        #    transform includes shear and cannot be represented---
        group = shapes.add_group_shape([member])
        set_group_xfrm(
            group,
            off=(1000000, 800000),
            ext=(2000000, 500000),
            ch_off=(0, 0),
            ch_ext=(1000000, 1000000),
        )
        member.rotation = 45.0
        xfrm_before = member.element.get_or_add_xfrm().xml

        with pytest.raises(UnrepresentableTransformError):
            shapes.reparent(member)

        assert member.element.getparent() is group.element
        assert member.element.get_or_add_xfrm().xml == xfrm_before

    def it_rejects_a_rotated_group_inside_a_nonuniformly_scaled_group(self, nested):
        slide, outer, inner, member = nested
        inner.rotation = 45.0

        with pytest.raises(UnrepresentableTransformError):
            slide.shapes.reparent(member)

        assert member.element.getparent() is inner.element

    def it_preserves_connector_endpoints_and_bindings(self, slide):
        shapes = slide.shapes
        begin_shape = shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Emu(100000), Emu(100000), Emu(400000), Emu(300000)
        )
        end_shape = shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Emu(1500000), Emu(900000), Emu(400000), Emu(300000)
        )
        connector = shapes.add_connector(
            MSO_CONNECTOR_TYPE.STRAIGHT, Emu(300000), Emu(250000), Emu(1500000), Emu(1050000)
        )
        connector.begin_connect(begin_shape, 2)
        connector.end_connect(end_shape, 0)
        group = shapes.add_group_shape([connector])
        set_group_xfrm(
            group,
            off=(500000, 600000),
            ext=(2400000, 900000),
            ch_off=(0, 0),
            ch_ext=(1200000, 1200000),
            rot=15.0,
        )
        begin_before = page_point(connector.element, (0.0, 0.0))
        end_before = page_point(
            connector.element,
            (float(connector.element.cx), float(connector.element.cy)),
        )
        id_before = connector.element.shape_id

        shapes.reparent(connector)

        # ---endpoint page positions are preserved---
        assert_same_geometry([begin_before], [page_point(connector.element, (0.0, 0.0))])
        assert_same_geometry(
            [end_before],
            [
                page_point(
                    connector.element,
                    (float(connector.element.cx), float(connector.element.cy)),
                )
            ],
        )
        # ---endpoint bindings are preserved and shape ids remain valid and unique---
        cNvCxnSpPr = connector.element.nvCxnSpPr.cNvCxnSpPr
        assert cNvCxnSpPr.stCxn.id == begin_shape.shape_id
        assert cNvCxnSpPr.endCxn.id == end_shape.shape_id
        assert connector.element.shape_id == id_before
        all_ids = [int(id_) for id_ in slide.shapes._spTree.xpath("//p:cNvPr/@id")]
        assert len(all_ids) == len(set(all_ids))


class DescribeUngroup(object):

    def it_moves_members_to_the_parent_tree_and_removes_the_group(self, nested):
        slide, outer, inner, member = nested
        member_before = page_corners(member)

        moved = slide.shapes.ungroup(outer)

        assert outer.element.getparent() is None
        assert inner.element.getparent() is slide.shapes._spTree
        assert_same_geometry(member_before, page_corners(member))
        assert len(moved) == 2

    def it_ungroups_via_the_group_shape_itself(self, nested):
        slide, outer, inner, member = nested
        member_before = page_corners(member)

        moved = inner.ungroup()

        assert inner.element.getparent() is None
        assert member.element.getparent() is outer.element
        assert_same_geometry(member_before, page_corners(member))
        assert list(moved)[0].element is member.element

    def it_simply_removes_an_empty_group(self, slide):
        shapes = slide.shapes
        group = shapes.add_group_shape()

        moved = shapes.ungroup(group)

        assert moved == []
        assert group.element.getparent() is None

    def it_leaves_an_explicitly_emptied_group_in_place_until_ungrouped(self, nested):
        slide, outer, inner, member = nested

        slide.shapes.reparent(member)

        # ---the emptied inner group remains in place, unchanged---
        assert inner.element.getparent() is outer.element
        assert len(inner.shapes) == 0
        # ---and can be removed explicitly---
        slide.shapes.ungroup(inner)
        assert inner.element.getparent() is None


class DescribeAddGroupShape(object):

    def it_preserves_page_geometry_when_grouping_nested_shapes(self, nested):
        slide, outer, inner, member = nested
        member_before = page_corners(member)

        new_group = slide.shapes.add_group_shape([member])

        assert member.element.getparent() is new_group.element
        assert_same_geometry(member_before, page_corners(member))

    def it_rejects_grouping_a_shape_with_its_own_ancestor(self, nested):
        slide, outer, inner, member = nested

        with pytest.raises(ValueError):
            slide.shapes.add_group_shape([inner, member])

    def it_rejects_shapes_from_a_different_slide(self, nested):
        slide, outer, inner, member = nested
        prs = Presentation()
        other_slide = prs.slides.add_slide(prs.slide_layouts[6])

        with pytest.raises(ValueError):
            other_slide.shapes.reparent(member)
        with pytest.raises(ValueError):
            other_slide.shapes.add_group_shape([member])
        with pytest.raises(ValueError):
            other_slide.shapes.ungroup(inner)

    def it_rejects_reparenting_a_group_into_itself(self, nested):
        slide, outer, inner, member = nested

        with pytest.raises(ValueError):
            inner.shapes.reparent(inner)
        with pytest.raises(ValueError):
            inner.shapes.reparent(outer)
