"""Integration tests for multi-plot category charts."""

from __future__ import annotations

import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

import pytest

from pptx import Presentation
from pptx.chart.data import ComboChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def combo_data(categories, sales, margin, margin_name="Margin"):
    chart_data = ComboChartData()
    chart_data.categories = categories
    column_plot = chart_data.add_plot(XL_CHART_TYPE.COLUMN_CLUSTERED)
    column_plot.add_series("Sales", sales)
    line_plot = chart_data.add_plot(
        XL_CHART_TYPE.LINE, secondary_axis=True, number_format="0%"
    )
    line_plot.add_series(margin_name, margin)
    return chart_data


def two_line_data(categories, left, right=None):
    chart_data = ComboChartData()
    chart_data.categories = categories
    primary_plot = chart_data.add_plot(XL_CHART_TYPE.LINE)
    primary_plot.add_series("Left", left)
    secondary_plot = chart_data.add_plot(XL_CHART_TYPE.LINE, secondary_axis=True)
    if right is not None:
        secondary_plot.add_series("Right", right)
    return chart_data


def xlsx_cells(prs_blob):
    with zipfile.ZipFile(BytesIO(prs_blob)) as package:
        xlsx_name = next(
            name for name in package.namelist() if name.startswith("ppt/embeddings/")
        )
        with zipfile.ZipFile(BytesIO(package.read(xlsx_name))) as workbook:
            shared_strings = [
                node.text
                for node in ET.fromstring(
                    workbook.read("xl/sharedStrings.xml")
                ).iter(f"{{{XLSX_NS}}}t")
            ]
            root = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))

    cells = {}
    for cell in root.iter(f"{{{XLSX_NS}}}c"):
        value = cell.find(f"{{{XLSX_NS}}}v")
        if value is None:
            continue
        if cell.get("t") == "s":
            cells[cell.get("r")] = shared_strings[int(value.text)]
        else:
            cells[cell.get("r")] = float(value.text)
    return cells


class DescribeComboChart(object):
    def but_it_requires_categories_before_a_plot_can_be_added(self):
        with pytest.raises(ValueError):
            ComboChartData().add_plot(XL_CHART_TYPE.LINE)

    def it_creates_a_column_line_combo_with_a_secondary_percentage_axis(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(8),
            Inches(4),
            combo_data(("Jan", "Feb"), (10, 20), (0.1, 0.2)),
        ).chart

        assert len(chart.plots) == 2
        assert [plot.series[0].name for plot in chart.plots] == ["Sales", "Margin"]
        assert chart.series[0].values == (10.0, 20.0)
        assert chart.series[1].values == (0.1, 0.2)

        chart_space = chart._chartSpace
        bar_chart, line_chart = chart_space.plotArea.xCharts
        assert bar_chart.xpath("./c:axId/@val") == ["1", "2"]
        assert line_chart.xpath("./c:axId/@val") == ["1", "3"]
        assert [ser.idx.val for ser in chart_space.plotArea.sers] == [0, 1]
        assert [ser.order.val for ser in chart_space.plotArea.sers] == [0, 1]
        cat_axis, primary_axis, secondary_axis = chart_space.plotArea.xpath(
            "./c:catAx|./c:valAx"
        )
        assert cat_axis.xpath("./c:crossAx/@val") == ["2"]
        assert primary_axis.xpath("./c:crossAx/@val") == ["1"]
        assert secondary_axis.xpath("./c:crossAx/@val") == ["1"]
        assert secondary_axis.xpath("./c:numFmt/@formatCode") == ["0%"]

        line_chart_data = two_line_data(("Jan", "Feb"), (1, 2), (3, 4))
        line_chart = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(1),
            Inches(1),
            Inches(8),
            Inches(4),
            line_chart_data,
        ).chart
        right_axis = line_chart._chartSpace.plotArea.valAx_lst[1]
        assert right_axis.xpath("./c:numFmt/@formatCode") == []

        blob = BytesIO()
        prs.save(blob)
        cells = xlsx_cells(blob.getvalue())
        assert cells == {
            "A2": "Jan",
            "A3": "Feb",
            "B1": "Sales",
            "B2": 10.0,
            "B3": 20.0,
            "C1": "Margin",
            "C2": 0.1,
            "C3": 0.2,
        }

    def it_updates_a_combo_twice_without_rebuilding_plot_formatting(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        graphic_frame = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED,
            Inches(1),
            Inches(1),
            Inches(8),
            Inches(4),
            combo_data(("Jan", "Feb"), (10, 20), (0.1, 0.2)),
        )
        chart = graphic_frame.chart
        chart.plots[0].has_data_labels = True
        chart.plots[1].has_data_labels = True
        axis_ids = [
            axis.axId.val
            for axis in chart._chartSpace.plotArea.xpath("./c:catAx|./c:valAx")
        ]

        chart.replace_data(combo_data(("Mar", "Apr"), (30, 40), (0.3, 0.4)))
        blob = BytesIO()
        prs.save(blob)
        prs = Presentation(blob)
        chart = prs.slides[0].shapes[0].chart
        chart.replace_data(combo_data(("May", "Jun"), (50, 60), (0.5, 0.6)))

        assert [plot.has_data_labels for plot in chart.plots] == [True, True]
        post_axis_ids = [
            axis.axId.val
            for axis in chart._chartSpace.plotArea.xpath("./c:catAx|./c:valAx")
        ]
        assert post_axis_ids == axis_ids
        assert chart.series[0].values == (50.0, 60.0)
        assert chart.series[1].values == (0.5, 0.6)
        assert chart._chartSpace.xpath('.//c:f[text()="Sheet1!$B$2:$B$3"]')
        assert chart._chartSpace.xpath('.//c:f[text()="Sheet1!$C$2:$C$3"]')

        final_blob = BytesIO()
        prs.save(final_blob)
        assert xlsx_cells(final_blob.getvalue()) == {
            "A2": "May",
            "A3": "Jun",
            "B1": "Sales",
            "B2": 50.0,
            "B3": 60.0,
            "C1": "Margin",
            "C2": 0.5,
            "C3": 0.6,
        }

    def it_removes_a_series_from_one_plot_and_preserves_the_other_plot(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        chart = slide.shapes.add_chart(
            XL_CHART_TYPE.LINE,
            Inches(1),
            Inches(1),
            Inches(8),
            Inches(4),
            two_line_data(("A", "B", "C"), (1, 2, 3), (0.1, 0.2, 0.3)),
        ).chart
        plot_axis_ids = [
            x_chart.xpath("./c:axId/@val")
            for x_chart in chart._chartSpace.plotArea.xCharts
        ]
        assert plot_axis_ids == [["1", "2"], ["1", "3"]]
        chart.plots[0].has_data_labels = True
        chart.plots[1].has_data_labels = True
        axis_ids = [
            axis.axId.val
            for axis in chart._chartSpace.plotArea.xpath("./c:catAx|./c:valAx")
        ]

        chart.replace_data(two_line_data(("D", "E"), (4, 5)))

        assert len(chart.plots) == 2
        assert len(chart.plots[0].series) == 1
        assert len(chart.plots[1].series) == 0
        assert [plot.has_data_labels for plot in chart.plots] == [True, True]
        assert len(chart._chartSpace.plotArea.valAx_lst) == 2
        assert chart.plots[0].series[0].values == (4.0, 5.0)

        chart.replace_data(two_line_data(("F", "G"), (6, 7), (0.6, 0.7)))

        assert len(chart.plots[1].series) == 1
        assert chart.plots[1].series[0].values == (0.6, 0.7)
        assert [plot.has_data_labels for plot in chart.plots] == [True, True]
        assert [
            axis.axId.val for axis in chart._chartSpace.plotArea.xpath("./c:catAx|./c:valAx")
        ] == axis_ids
