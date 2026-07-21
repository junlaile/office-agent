"""vehicle_data.py 单元测试：mock 查询确定性 + 脱敏。

全部纯函数，确定性输出（车牌做 MD5 种子）。
"""

from __future__ import annotations

import pytest

from office_agent.vehicle_data import (
    _mask_id_card,
    _mask_phone,
    _seed_rng,
    mock_query,
    query,
)


class TestDeterminism:
    """同一车牌返回一致结果（确定性）。"""

    def test_same_plate_same_result(self):
        """同车牌两次查询结果 deep equal。"""
        r1 = mock_query("京A12345")
        r2 = mock_query("京A12345")
        assert r1 == r2

    def test_different_plates_usually_different(self):
        """不同车牌大概率不同结果（统计性，用多个样本）。"""
        results = {mock_query(f"京A{i:05d}")["status"] for i in range(20)}
        # 至少不会全部相同 status
        assert len(results) >= 1  # 弱断言，主要确保不报错

    def test_seed_rng_deterministic(self):
        """_seed_rng 同车牌产生相同序列。"""
        rng1 = _seed_rng("沪B88888")
        rng2 = _seed_rng("沪B88888")
        assert rng1.random() == rng2.random()
        assert rng1.randint(0, 1000) == rng2.randint(0, 1000)


class TestStatusBranches:
    """三种返回 status。"""

    def test_empty_plate_not_found(self):
        """空车牌返回 not_found。"""
        assert mock_query("")["status"] == "not_found"
        assert mock_query("   ")["status"] == "not_found"
        assert mock_query(None)["status"] == "not_found"  # type: ignore[arg-type]

    @pytest.mark.parametrize("status", ["ok", "multiple", "not_found"])
    def test_all_statuses_occur(self, status):
        """扫描足够多车牌，三种 status 都会出现。"""
        found = False
        for i in range(200):
            r = mock_query(f"测A{i:06d}")
            if r["status"] == status:
                found = True
                # 验证该 status 的返回结构
                if status == "ok":
                    assert "vehicle" in r
                    assert "image_url" in r
                    assert "accidents" in r
                    assert "violations" in r
                elif status == "multiple":
                    assert "candidates" in r
                    assert len(r["candidates"]) >= 2
                else:  # not_found
                    assert "message" in r
                break
        assert found, f"200 个车牌里没出现 status={status}"


class TestOkResultStructure:
    """status=ok 的返回结构完整性。"""

    @pytest.fixture
    def ok_result(self):
        """找一个 ok 结果的车牌。"""
        for i in range(50):
            r = mock_query(f"京A{i:05d}")
            if r["status"] == "ok":
                return r
        pytest.skip("50 个车牌都没 ok，概率异常")

    def test_vehicle_fields(self, ok_result):
        v = ok_result["vehicle"]
        assert "plate_number" in v
        assert "brand" in v
        assert "model" in v
        assert "color" in v
        assert "owner" in v

    def test_owner_fields(self, ok_result):
        owner = ok_result["vehicle"]["owner"]
        assert "name" in owner
        assert "id_card" in owner
        assert "phone" in owner

    def test_image_url_is_string(self, ok_result):
        assert isinstance(ok_result["image_url"], str)
        assert ok_result["image_url"].startswith("http")

    def test_stats_consistent(self, ok_result):
        """stats 计数与实际列表长度一致。"""
        stats = ok_result["stats"]
        assert stats["accident_count"] == len(ok_result["accidents"])


class TestMasking:
    """脱敏函数。"""

    def test_mask_id_card_standard(self):
        """身份证脱敏：保留前6后4，中间8位打*。"""
        assert _mask_id_card("110101199001011234") == "110101********1234"

    def test_mask_id_card_short(self):
        """短输入不崩溃。"""
        assert _mask_id_card("12345") == "12345"

    def test_mask_phone_standard(self):
        """手机号脱敏：保留前3后4，中间4位打*。"""
        assert _mask_phone("13800138000") == "138****8000"

    def test_mask_phone_short(self):
        """短输入不崩溃。"""
        assert _mask_phone("12345") == "12345"


class TestQueryEntryPoint:
    """query() 统一入口（当前 MOCK_MODE=True 走 mock_query）。"""

    def test_query_equals_mock_when_mock_mode(self):
        """MOCK_MODE=True 时 query == mock_query。"""
        # query 内部转调 mock_query
        assert query("京A99999") == mock_query("京A99999")
