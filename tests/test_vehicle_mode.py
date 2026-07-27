"""vehicle_mode：交通关键词判定 + 提示词/工具的条件注入。"""

from __future__ import annotations

from office_agent.agent.prompts import build_system_prompt
from office_agent.domain.vehicle_data import is_vehicle_related


class TestIsVehicleRelated:
    def test_hits(self):
        assert is_vehicle_related("写一份交通事故报告，车牌京A12345")
        assert is_vehicle_related("车辆评估报告")
        assert is_vehicle_related("车险理赔材料")

    def test_misses(self):
        assert not is_vehicle_related("写一份项目周报")
        assert not is_vehicle_related("做一份季度销售数据的 Excel 表格")
        assert not is_vehicle_related("")
        assert not is_vehicle_related(None)


class TestVehicleRulesInjection:
    def test_default_no_vehicle_rules(self):
        prompt = build_system_prompt("/tmp/a.docx")
        assert "query_vehicle" not in prompt
        assert "交通类文档专项" not in prompt

    def test_vehicle_mode_injects_rules(self):
        prompt = build_system_prompt("/tmp/a.docx", vehicle_mode=True)
        assert "query_vehicle" in prompt
        assert "交通类文档专项" in prompt

    def test_vehicle_mode_works_for_all_kinds(self):
        for ext in ("docx", "xlsx", "pptx"):
            prompt = build_system_prompt(f"/tmp/a.{ext}", vehicle_mode=True)
            assert "query_vehicle" in prompt
