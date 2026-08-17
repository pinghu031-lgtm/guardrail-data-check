from openpyxl import Workbook

from guardrail_core import (
    CATEGORY_BOLT,
    CATEGORY_HEIGHT,
    ISSUE_BOLT_ABNORMAL_VALUE,
    ISSUE_ZERO_VALUE_MARK_ERROR,
    InspectionProcessor,
)


def build_sample_workbook(path):
    workbook = Workbook()

    height_sheet = workbook.active
    height_sheet.title = "护栏高度"
    height_sheet.append(
        [
            "地市",
            "路线编号",
            "方向",
            "原始桩号",
            "护栏类型",
            "梁板中心高度-离地(mm)",
            "路缘石高度-离地(mm)",
            "梁板中心高度(mm)",
            "异常标记",
        ]
    )
    height_sheet.append(["广州", "G15", "上行", "K1+000", "三波护栏", 0, 0, 0, ""])

    bolt_sheet = workbook.create_sheet("螺栓缺失")
    bolt_sheet.append(
        [
            "地市",
            "路线编号",
            "方向",
            "原始桩号",
            "护栏类型",
            "备注标记",
            "拼接螺栓缺失数量（颗）",
            "连接螺栓缺失数量（颗）",
        ]
    )
    bolt_sheet.append(["广州", "G15", "上行", "K1+010", "三波护栏", "", -1, 0])
    workbook.save(path)


def test_inspection_processor_preserves_height_zero_value_rule(tmp_path):
    source = tmp_path / "sample.xlsx"
    build_sample_workbook(source)

    groups, stats = InspectionProcessor().process_files([str(source)])

    assert stats["files_success"] == 1
    assert any(
        group.category == CATEGORY_HEIGHT and group.issue_type == ISSUE_ZERO_VALUE_MARK_ERROR
        for group in groups
    )


def test_inspection_processor_preserves_negative_bolt_rule(tmp_path):
    source = tmp_path / "sample.xlsx"
    build_sample_workbook(source)

    groups, _stats = InspectionProcessor().process_files([str(source)])

    assert any(
        group.category == CATEGORY_BOLT and group.issue_type == ISSUE_BOLT_ABNORMAL_VALUE
        for group in groups
    )
