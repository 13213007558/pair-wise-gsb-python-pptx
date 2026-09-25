"""Integration tests for combo (multi-plot) chart creation and update."""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from pptx import Presentation
from pptx.chart.data import ComboChartData
from pptx.chart.plot import PlotTypeInspector
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches


def _bar_line_chart_data():
    chart_data = ComboChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    bar_plot = chart_data.add_plot(XL_CHART_TYPE.COLUMN_CLUSTERED)
    bar_plot.add_series("Sales", (100, 200, 300, 400), number_format="#,##0")
    line_plot = chart_data.add_plot(XL_CHART_TYPE.LINE, secondary_axis=True)
    line_plot.add_series("Margin", (0.12, 0.15, 0.11, 0.18), number_format="0.0%")
    return chart_data


def _two_line_chart_data():
    chart_data = ComboChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    left_plot = chart_data.add_plot(XL_CHART_TYPE.LINE)
    left_plot.add_series("Revenue", (10, 20, 30, 40))
    right_plot = chart_data.add_plot(XL_CHART_TYPE.LINE_MARKERS, secondary_axis=True)
    right_plot.add_series("Growth", (0.5, 0.6, 0.55, 0.7), number_format="0%")
    return chart_data


def _presentation_with_chart(chart_data):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    graphic_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1),
        Inches(8),
        Inches(5),
        chart_data,
    )
    return prs, graphic_frame.chart


def _save_and_reopen(prs):
    stream = io.BytesIO()
    prs.save(stream)
    stream.seek(0)
    return Presentation(stream)


def _chart_of(prs):
    return prs.slides[0].shapes[0].chart


def _child_val(el, tag):
    """Return the ``val`` attribute of the *tag* child of *el*, as a string."""
    child = el.find(qn(tag))
    return None if child is None else child.get("val")


def _values_ref(ser):
    """Return the worksheet range referenced by the ``c:val`` of *ser*."""
    return ser.find(qn("c:val")).find(qn("c:numRef")).find(qn("c:f")).text


def _xlsx_cells(chart):
    """Return dict mapping cell reference to value in the embedded workbook."""
    blob = chart.part.chart_workbook.xlsx_part.blob
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        shared = (
            zf.read("xl/sharedStrings.xml").decode("utf-8")
            if "xl/sharedStrings.xml" in zf.namelist()
            else ""
        )
    cells = dict(re.findall(r'<c r="([A-Z]+[0-9]+)"[^>]*><v>([^<]*)</v></c>', sheet_xml))
    return cells, shared


class DescribeComboChartCreation:
    def it_creates_a_bar_and_line_plot_sharing_a_category_axis(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())

        assert len(chart.plots) == 2
        assert PlotTypeInspector.chart_type(chart.plots[0]) == XL_CHART_TYPE.COLUMN_CLUSTERED
        assert PlotTypeInspector.chart_type(chart.plots[1]) == XL_CHART_TYPE.LINE
        assert [s.name for s in chart.series] == ["Sales", "Margin"]
        assert [list(s.values) for s in chart.series] == [
            [100.0, 200.0, 300.0, 400.0],
            [0.12, 0.15, 0.11, 0.18],
        ]

    def it_assigns_consistent_axis_ids_and_crossAx_references(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())
        chartSpace = chart._chartSpace

        catAxs = chartSpace.xpath(".//c:catAx")
        valAxs = chartSpace.xpath(".//c:valAx")
        assert len(catAxs) == 1
        assert len(valAxs) == 2

        cat_ax_id = _child_val(catAxs[0], "c:axId")
        prim_val_ax_id, sec_val_ax_id = (_child_val(ax, "c:axId") for ax in valAxs)
        assert len({cat_ax_id, prim_val_ax_id, sec_val_ax_id}) == 3

        # ---each plot references the shared category axis and its value axis---
        barChart, lineChart = chartSpace.xpath(".//c:barChart"), chartSpace.xpath(".//c:lineChart")
        assert [e.get("val") for e in barChart[0].xpath("./c:axId")] == [cat_ax_id, prim_val_ax_id]
        assert [e.get("val") for e in lineChart[0].xpath("./c:axId")] == [cat_ax_id, sec_val_ax_id]

        # ---crossAx references form a consistent graph---
        assert _child_val(catAxs[0], "c:crossAx") == prim_val_ax_id
        assert _child_val(valAxs[0], "c:crossAx") == cat_ax_id
        assert _child_val(valAxs[1], "c:crossAx") == cat_ax_id

    def it_places_the_secondary_value_axis_on_the_right_with_percent_format(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())
        valAxs = chart._chartSpace.xpath(".//c:valAx")

        prim_ax, sec_ax = valAxs
        assert _child_val(prim_ax, "c:axPos") == "l"
        assert _child_val(sec_ax, "c:axPos") == "r"
        assert sec_ax.numFmt.formatCode == "0.0%"
        assert _child_val(sec_ax, "c:crosses") == "max"

    def it_assigns_series_idx_and_order_unique_across_plots(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())
        sers = chart._chartSpace.plotArea.sers

        assert [ser.idx.val for ser in sers] == [0, 1]
        assert [ser.order.val for ser in sers] == [0, 1]

    def it_writes_series_to_the_embedded_workbook_in_plot_order(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())
        cells, shared = _xlsx_cells(chart)

        # ---series names occupy columns B and C of row 1---
        assert "Sales" in shared
        assert "Margin" in shared
        # ---categories in column A, values in columns B and C---
        assert "Q1" in shared
        assert "Q4" in shared
        assert cells["B2"] == "100"
        assert cells["B5"] == "400"
        assert cells["C2"] == "0.12"
        assert cells["C5"] == "0.18"

    def it_supports_two_line_plots_on_left_and_right_axes(self):
        prs, chart = _presentation_with_chart(_two_line_chart_data())

        assert len(chart.plots) == 2
        assert PlotTypeInspector.chart_type(chart.plots[0]) == XL_CHART_TYPE.LINE
        assert PlotTypeInspector.chart_type(chart.plots[1]) == XL_CHART_TYPE.LINE_MARKERS

        lineCharts = chart._chartSpace.xpath(".//c:lineChart")
        assert len(lineCharts) == 2
        valAxs = chart._chartSpace.xpath(".//c:valAx")
        assert [_child_val(ax, "c:axPos") for ax in valAxs] == ["l", "r"]
        assert valAxs[1].numFmt.formatCode == "0%"

    def it_round_trips_through_save_and_reopen(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())
        chart2 = _chart_of(_save_and_reopen(prs))

        assert len(chart2.plots) == 2
        assert [s.name for s in chart2.series] == ["Sales", "Margin"]
        assert [list(s.values) for s in chart2.series] == [
            [100.0, 200.0, 300.0, 400.0],
            [0.12, 0.15, 0.11, 0.18],
        ]
        assert list(chart2.plots[0].categories) == ["Q1", "Q2", "Q3", "Q4"]

    def it_raises_when_a_plot_has_no_series(self):
        chart_data = ComboChartData()
        chart_data.categories = ["Q1"]
        chart_data.add_plot(XL_CHART_TYPE.COLUMN_CLUSTERED)

        with pytest.raises(ValueError, match="no series"):
            chart_data.xml_bytes(XL_CHART_TYPE.COLUMN_CLUSTERED)

    def it_raises_when_there_are_no_plots(self):
        chart_data = ComboChartData()
        chart_data.categories = ["Q1"]

        with pytest.raises(ValueError, match="no plots"):
            chart_data.xml_bytes(XL_CHART_TYPE.COLUMN_CLUSTERED)


class DescribeComboChartReplaceData:
    def it_replaces_values_in_both_plots_and_the_workbook(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())

        new_data = _bar_line_chart_data()
        new_data.plots[0][0]._data_points[0]._value = 111
        chart.replace_data(new_data)

        assert [list(s.values) for s in chart.series] == [
            [111.0, 200.0, 300.0, 400.0],
            [0.12, 0.15, 0.11, 0.18],
        ]
        cells, _ = _xlsx_cells(chart)
        assert cells["B2"] == "111"

    def it_preserves_other_plot_axes_styles_and_data_labels(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())
        bar_plot, line_plot = chart.plots
        bar_plot.has_data_labels = True
        bar_plot.data_labels.number_format = "#,##0"
        bar_plot.data_labels.number_format_is_linked = False
        chart.series[1].smooth = True

        chart.replace_data(_bar_line_chart_data())

        bar_plot2, line_plot2 = chart.plots
        assert bar_plot2.has_data_labels is True
        assert bar_plot2.data_labels.number_format == "#,##0"
        assert chart.series[1].smooth is True
        # ---axes are untouched by data replacement---
        assert chart._chartSpace.xpath(".//c:catAx") != []
        assert len(chart._chartSpace.xpath(".//c:valAx")) == 2
        valAxs = chart._chartSpace.xpath(".//c:valAx")
        assert [_child_val(ax, "c:axPos") for ax in valAxs] == ["l", "r"]
        assert valAxs[1].numFmt.formatCode == "0.0%"

    def it_removes_a_series_without_disturbing_the_remaining_plot(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())
        chart.plots[0].has_data_labels = True

        new_data = ComboChartData()
        new_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        bar_plot = new_data.add_plot(XL_CHART_TYPE.COLUMN_CLUSTERED)
        bar_plot.add_series("Sales", (101, 201, 301, 401))
        chart.replace_data(new_data)

        assert len(chart.plots) == 1
        assert [s.name for s in chart.series] == ["Sales"]
        assert list(chart.series[0].values) == [101.0, 201.0, 301.0, 401.0]
        assert chart.plots[0].has_data_labels is True
        # ---remaining axes are still intact and consistent---
        catAxs = chart._chartSpace.xpath(".//c:catAx")
        assert len(catAxs) == 1
        barChart = chart._chartSpace.xpath(".//c:barChart")[0]
        assert _child_val(catAxs[0], "c:axId") in [e.get("val") for e in barChart.xpath("./c:axId")]

        chart2 = _chart_of(_save_and_reopen(prs))
        assert [s.name for s in chart2.series] == ["Sales"]
        assert list(chart2.series[0].values) == [101.0, 201.0, 301.0, 401.0]

    def it_adds_a_series_to_an_existing_plot_keeping_refs_consistent(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())

        new_data = ComboChartData()
        new_data.categories = ["Q1", "Q2", "Q3", "Q4"]
        bar_plot = new_data.add_plot(XL_CHART_TYPE.COLUMN_CLUSTERED)
        bar_plot.add_series("Sales", (100, 200, 300, 400))
        bar_plot.add_series("Returns", (5, 6, 7, 8))
        line_plot = new_data.add_plot(XL_CHART_TYPE.LINE, secondary_axis=True)
        line_plot.add_series("Margin", (0.12, 0.15, 0.11, 0.18))
        chart.replace_data(new_data)

        assert [s.name for s in chart.series] == ["Sales", "Returns", "Margin"]
        sers = chart._chartSpace.plotArea.sers
        # ---idx and order values are unique across the whole chart---
        assert sorted(ser.idx.val for ser in sers) == [0, 1, 2]
        assert sorted(ser.order.val for ser in sers) == [0, 1, 2]
        # ---the new series stays in the bar plot and points at column C---
        barChart = chart._chartSpace.xpath(".//c:barChart")[0]
        assert len(barChart.sers) == 2
        assert _values_ref(barChart.sers[1]) == "Sheet1!$C$2:$C$5"
        lineChart = chart._chartSpace.xpath(".//c:lineChart")[0]
        assert _values_ref(lineChart.sers[0]) == "Sheet1!$D$2:$D$5"

        chart2 = _chart_of(_save_and_reopen(prs))
        assert [list(s.values) for s in chart2.series] == [
            [100.0, 200.0, 300.0, 400.0],
            [5.0, 6.0, 7.0, 8.0],
            [0.12, 0.15, 0.11, 0.18],
        ]

    def it_supports_replacing_data_again_after_save_and_reopen(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())
        chart.replace_data(_bar_line_chart_data())

        prs2 = _save_and_reopen(prs)
        chart2 = _chart_of(prs2)

        newer = ComboChartData()
        newer.categories = ["A", "B"]
        bar_plot = newer.add_plot(XL_CHART_TYPE.COLUMN_CLUSTERED)
        bar_plot.add_series("Sales", (1, 2))
        line_plot = newer.add_plot(XL_CHART_TYPE.LINE, secondary_axis=True)
        line_plot.add_series("Margin", (0.5, 0.6), number_format="0.0%")
        chart2.replace_data(newer)

        assert [list(s.values) for s in chart2.series] == [[1.0, 2.0], [0.5, 0.6]]
        assert list(chart2.plots[0].categories) == ["A", "B"]
        assert len(chart2._chartSpace.xpath(".//c:valAx")) == 2

        chart3 = _chart_of(_save_and_reopen(prs2))
        assert [list(s.values) for s in chart3.series] == [[1.0, 2.0], [0.5, 0.6]]
        cells, shared = _xlsx_cells(chart3)
        assert cells["B2"] == "1"
        assert cells["C3"] == "0.6"

    def it_raises_when_chart_data_has_more_plots_than_the_chart(self):
        prs, chart = _presentation_with_chart(_bar_line_chart_data())

        three_plot_data = ComboChartData()
        three_plot_data.categories = ["Q1"]
        for _ in range(3):
            plot = three_plot_data.add_plot(XL_CHART_TYPE.LINE)
            plot.add_series("S", (1,))

        with pytest.raises(ValueError, match="cannot add plots"):
            chart.replace_data(three_plot_data)
