"""车辆号牌查询的 mock 数据层。

当前是写死的随机数据（用车牌做种子保证同一车牌每次查询结果一致）。
后期接入真实交管接口时，只需替换 mock_query 内的实现（保留相同返回结构）。

返回结构（三种 status）:
    ok        — 唯一匹配，含完整车辆/所有人/图片/事故/违法信息
    multiple  — 多辆匹配，含候选清单（需调 ask_user 让用户选，再用 id 重查）
    not_found — 无匹配

MOCK_MODE 开关：True 走随机 mock，False 走真实接口（预留，当前未实现）。
"""

from __future__ import annotations

import hashlib
import random
from datetime import date, timedelta
from typing import Any

# 真实接口开关：True=mock（当前），False=真实接口（待实现）
MOCK_MODE = True

# 多车匹配概率（模拟车牌模糊/重号场景）
_MULTIPLE_PROB = 0.15
# 无匹配概率
_NOT_FOUND_PROB = 0.05

# 数据池
_BRANDS_MODELS = [
    ("比亚迪", ["汉EV", "秦PLUS", "宋Pro", "海豚", "元PLUS"]),
    ("丰田", ["凯美瑞", "卡罗拉", "RAV4荣放", "亚洲龙", "汉兰达"]),
    ("大众", ["朗逸", "帕萨特", "途观L", "迈腾", "高尔夫"]),
    ("宝马", ["3系", "5系", "X3", "X5", "1系"]),
    ("奔驰", ["C级", "E级", "GLC", "A级", "S级"]),
    ("本田", ["雅阁", "思域", "CR-V", "飞度", "皓影"]),
    ("奥迪", ["A4L", "A6L", "Q5L", "A3", "Q3"]),
    ("特斯拉", ["Model 3", "Model Y", "Model S", "Model X"]),
]
_COLORS = ["白色", "黑色", "银色", "灰色", "红色", "蓝色", "其他"]
_FUEL_TYPES = ["汽油", "柴油", "纯电动", "插电混动", "油电混合"]
_SURNAMES = ["张", "王", "李", "赵", "刘", "陈", "杨", "黄", "周", "吴", "徐", "孙"]
_GIVEN_NAMES = ["伟", "芳", "娜", "敏", "静", "强", "磊", "军", "洋", "勇", "艳", "杰"]
_PROVINCES = ["北京市", "上海市", "广东省", "江苏省", "浙江省", "山东省", "四川省", "湖北省"]
_CITIES = ["海淀区", "朝阳区", "浦东新区", "天河区", "鼓楼区", "西湖区", "历下区", "武侯区"]
_ACCIDENT_TYPES = ["追尾", "侧面碰撞", "剐蹭", "单车事故", "行人碰撞", "车辆失控"]
_LIABILITY = ["全责", "主责", "同责", "次责", "无责"]
_VIOLATION_TYPES = [
    ("超速行驶", 3, 200),
    ("闯红灯", 6, 200),
    ("违章停车", 0, 100),
    ("不按规定车道行驶", 2, 100),
    ("违规变道", 3, 200),
    ("未系安全带", 1, 50),
    ("打电话驾驶", 2, 200),
    ("逆行", 3, 200),
    ("违反禁令标志", 3, 200),
]


def _seed_rng(plate: str) -> random.Random:
    """用车牌生成确定性种子，保证同一车牌每次查询结果一致。"""
    h = hashlib.md5(plate.encode("utf-8")).hexdigest()
    return random.Random(int(h[:8], 16))


def _mask_id_card(idc: str) -> str:
    """身份证脱敏：保留前6后4，中间用 * 代替。"""
    if len(idc) >= 10:
        return idc[:6] + "*" * (len(idc) - 10) + idc[-4:]
    return idc


def _mask_phone(phone: str) -> str:
    """电话脱敏：保留前3后4。"""
    if len(phone) >= 7:
        return phone[:3] + "****" + phone[-4:]
    return phone


def _gen_id_card(rng: random.Random) -> str:
    """生成形似身份证号的 18 位字符串。"""
    body = "".join(str(rng.randint(0, 9)) for _ in range(17))
    check = str(rng.randint(0, 9))
    return body + check


def _gen_phone(rng: random.Random) -> str:
    prefix = rng.choice(["138", "139", "150", "158", "186", "188", "199"])
    suffix = "".join(str(rng.randint(0, 9)) for _ in range(8))
    return prefix + suffix


def _gen_vehicle(plate: str, rng: random.Random) -> dict[str, Any]:
    """生成单辆车的完整信息。"""
    brand, models = rng.choice(_BRANDS_MODELS)
    model = rng.choice(models)
    color = rng.choice(_COLORS)
    fuel = rng.choice(_FUEL_TYPES)
    # 注册日期：1-12 年前
    reg_days_ago = rng.randint(365, 365 * 12)
    reg_date = (date.today() - timedelta(days=reg_days_ago)).isoformat()
    vin = "L" + "".join(
        str(rng.randint(0, 9)) if rng.random() > 0.3 else rng.choice("ABCDEFGHJKLMNPRSTUVWXYZ")
        for _ in range(16)
    )
    engine_no = "".join(str(rng.randint(0, 9)) for _ in range(8))

    surname = rng.choice(_SURNAMES)
    given = rng.choice(_GIVEN_NAMES)
    if rng.random() > 0.5:
        given += rng.choice(_GIVEN_NAMES)
    owner = surname + given
    province = rng.choice(_PROVINCES)
    city = rng.choice(_CITIES)

    return {
        "plate_number": plate,
        "brand": brand,
        "model": model,
        "color": color,
        "fuel_type": fuel,
        "register_date": reg_date,
        "vin": vin,
        "engine_no": engine_no,
        "use_type": rng.choice(["非营运(私家车)", "非营运(私家车)", "非营运(私家车)", "营运"]),
        "inspection_expire": (date.today() + timedelta(days=rng.randint(30, 700))).isoformat(),
        "insurance_status": rng.choice(["在保", "在保", "在保", "已过期"]),
        "owner": {
            "name": owner,
            "id_card": _mask_id_card(_gen_id_card(rng)),
            "phone": _mask_phone(_gen_phone(rng)),
            "address": f"{province}{city}{rng.randint(1, 999)}号",
        },
    }


def _gen_accidents(rng: random.Random, n: int) -> list[dict[str, Any]]:
    accs = []
    for _ in range(n):
        days_ago = rng.randint(7, 365 * 3)
        accs.append(
            {
                "date": (date.today() - timedelta(days=days_ago)).isoformat(),
                "location": f"{rng.choice(_PROVINCES)}{rng.choice(_CITIES)}"
                f"{rng.choice(['路口', '快速路', '高速', '辅路'])}",
                "type": rng.choice(_ACCIDENT_TYPES),
                "liability": rng.choice(_LIABILITY),
                "severity": rng.choice(["轻微", "轻微", "一般", "较大"]),
                "settled": rng.choice([True, True, False]),
            }
        )
    accs.sort(key=lambda x: x["date"], reverse=True)  # type: ignore[arg-type,return-value]
    return accs


def _gen_violations(rng: random.Random, n: int) -> list[dict[str, Any]]:
    vios = []
    for _ in range(n):
        vtype, points, fine = rng.choice(_VIOLATION_TYPES)
        days_ago = rng.randint(7, 365 * 2)
        vios.append(
            {
                "date": (date.today() - timedelta(days=days_ago)).isoformat(),
                "type": vtype,
                "points": points,
                "fine": fine,
                "location": f"{rng.choice(_CITIES)}{rng.choice(['路段', '路口', '隧道'])}",
                "status": rng.choice(["已处理", "已处理", "未处理"]),
            }
        )
    vios.sort(key=lambda x: x["date"], reverse=True)  # type: ignore[arg-type,return-value]
    return vios


def _placeholder_image_url(plate: str) -> str:
    """占位车辆图片 URL（后期可换成真实车辆图）。"""
    # placehold.co 支持 text 参数；URL 编码车牌里的中文
    from urllib.parse import quote

    return f"https://placehold.co/400x300/EEE/333?text={quote(plate)}"


def mock_query(plate: str) -> dict[str, Any]:
    """mock 查询入口。同一车牌返回一致结果。

    返回 status:
        "ok"        — 唯一匹配，含完整信息
        "multiple"  — 多辆匹配，含 candidates 清单
        "not_found" — 无匹配
    """
    plate = (plate or "").strip()
    if not plate:
        return {"status": "not_found", "message": "车牌号为空"}

    rng = _seed_rng(plate)

    # 分支：无匹配
    if rng.random() < _NOT_FOUND_PROB:
        return {"status": "not_found", "message": f"未查询到车牌 {plate} 的登记信息"}

    # 分支：多辆匹配
    if rng.random() < _MULTIPLE_PROB:
        n = rng.randint(2, 3)
        candidates = []
        for i in range(n):
            v = _gen_vehicle(f"{plate}-{i + 1}" if i > 0 else plate, rng)
            candidates.append(
                {
                    "id": f"{plate}#{i + 1}",
                    "plate_number": v["plate_number"],
                    "brand": v["brand"],
                    "model": v["model"],
                    "color": v["color"],
                    "owner": v["owner"]["name"],
                }
            )
        return {
            "status": "multiple",
            "message": f"车牌 {plate} 匹配到 {n} 辆车，请确认是哪一辆",
            "candidates": candidates,
        }

    # 分支：唯一匹配（完整信息）
    vehicle = _gen_vehicle(plate, rng)
    n_acc = rng.choices([0, 1, 2, 3], weights=[30, 40, 20, 10])[0]
    n_vio = rng.choices([0, 1, 2, 3, 4], weights=[10, 25, 30, 25, 10])[0]
    return {
        "status": "ok",
        "vehicle": vehicle,
        "image_url": _placeholder_image_url(plate),
        "accidents": _gen_accidents(rng, n_acc),
        "violations": _gen_violations(rng, n_vio),
        "stats": {
            "accident_count": n_acc,
            "violation_count": n_vio,
            "total_points": sum(v["points"] for v in _gen_violations(rng, n_vio)),
        },
    }


def query(plate: str) -> dict[str, Any]:
    """对外统一查询入口：根据 MOCK_MODE 走 mock 或真实接口。

    后期实现真实接口时，把 MOCK_MODE 设为 False 并实现 _real_query。
    """
    if MOCK_MODE:
        return mock_query(plate)
    # TODO: 接入真实交管接口
    # return _real_query(plate)
    raise NotImplementedError("真实车辆查询接口尚未实现，请保持 MOCK_MODE=True")


# 需求文本里出现任一关键词 → 判定为交通/车辆相关会话。
# 用于决定是否绑定 query_vehicle 工具、注入交通类文档专项提示词。
# 宁可多判（多带一个工具+一段提示词无害），不可漏判（漏了查不到车辆信息）。
_VEHICLE_HINT_KEYWORDS = (
    "车牌",
    "车辆",
    "汽车",
    "机动车",
    "交通事故",
    "车祸",
    "车险",
    "理赔",
    "违章",
    "违法记录",
    "肇事",
    "驾驶",
)


def is_vehicle_related(text: str) -> bool:
    """判断需求文本是否与车辆/交通相关（决定 query_vehicle 工具与提示词注入）。"""
    t = text or ""
    return any(k in t for k in _VEHICLE_HINT_KEYWORDS)
