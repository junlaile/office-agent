# 新增车辆号牌查询功能

## 目标
生成交通相关文档时，用户在 ask_user 表单填入车牌后，agent 自主调用 `query_vehicle` 查询车牌对应的详细信息（基本信息/所有人/图片/最近事故/违法记录），信息**先 mock 随机返回**（后期你补真实接口）。查到多辆车时返回清单给 LLM，由 LLM 调 ask_user 让用户确认选哪辆，再查该车的详细信息。

## 数据流

```
用户需求(交通文档) 
  → agent 调 ask_user 收集车牌(表单)
  → 用户填车牌
  → agent 逐个调 query_vehicle(车牌)
       ├─ 唯一匹配 → 返回完整车辆信息
       └─ 多辆匹配 → 返回清单(含区分信息) → agent 调 ask_user 让用户选 → 再调 query_vehicle(选定ID)
  → agent 用返回的信息写文档(调 add_paragraph/add_table/add_image)
```

## 工具设计

### 1. `query_vehicle`（新工具）
```python
@tool
def query_vehicle(plate_number: str) -> dict:
    """根据车牌号查询车辆详细信息（基本信息/所有人/图片/事故/违法）。
    用于生成交通事故报告等交通类文档。
    返回:
      - 唯一匹配: {status:"ok", vehicle:{plate,brand,model,color,owner,...}, image_url, accidents:[...], violations:[...]}
      - 多辆匹配: {status:"multiple", candidates:[{id,plate,owner,brand},...]} 
                  → 此时需调 ask_user 让用户选，再用选定 id 重新查询
      - 无匹配:   {status:"not_found"}
    """
```
返回 dict（结构化），agent 据此组织文档内容。

### 2. Mock 数据模块（`src/office_agent/vehicle_data.py`，新）
- `mock_query(plate: str) -> dict`：先写死随机返回。用 plate 做种子保证同车牌返回一致（可复现）。
- 基本信息随机生成：品牌(比亚迪/丰田/大众/宝马...)、车型、颜色、注册日期、车架号、发动机号
- 所有人：姓名、身份证(脱敏)、电话(脱敏)、地址
- 图片 URL：用一个占位图 URL（如 `https://placehold.co/400x300?text=<车牌>`），后期可换成真实车辆图
- 事故记录：随机 0-3 条（日期/地点/性质/责任）
- 违法记录：随机 0-3 条（日期/违法项/扣分/罚款）
- **多车逻辑**：约 15% 概率返回"多辆匹配"（模拟车牌模糊/重号场景），给 2-3 个候选
- 全部用 `random.seed(hash(plate))` 保证同一车牌每次查结果一致

预留真实接口接入点：函数内用 `if MOCK_MODE:` 分支，真实接口实现后替换即可。

### 3. `add_image`（新工具，让车辆图片能进文档）
新增 officecli 图片插入能力（上游已确认支持 `--type picture`）。在 DocTool 加 `add_image` 方法，tools.py 暴露为 @tool。
```python
@tool
def add_image(url_or_path: str, width: str = "8cm", caption: str = "") -> str:
    """插入图片到文档（车辆照片等）。支持本地路径/URL。可选加图注。"""
```

## 提示词更新（prompts.py）
在 AGENT_SYSTEM_PROMPT 加交通文档专项指引：
```
## 交通类文档专项（事故报告/车辆评估等）
当生成交通相关文档时:
1. 用 ask_user 收集车牌（字段 key 用 plate_a/plate_b 等）
2. 用户填完车牌后，【逐个】调 query_vehicle(车牌) 查询每辆车的详细信息
3. 若 query_vehicle 返回 status:"multiple"，调 ask_user 让用户从候选里选，再用选定信息
4. 把查回的信息（车辆基本/所有人/事故/违法）结构化写入文档:
   - 车辆基本信息用表格
   - 车辆图片用 add_image
   - 事故/违法记录用列表或表格
```

## main.py 工具调用展示
`_format_tool_call` 加 `query_vehicle` 和 `add_image` 的格式化分支：
```
🔧 query_vehicle("京A12345")
   ↳ 查询到: 比迪汉EV · 张三 · 2条违法
🔧 add_image("https://placehold.co/...", caption="车辆照片")
```

## 文件清单
```
src/office_agent/
├── vehicle_data.py   # 新: mock_query 随机数据 + 真实接口预留点
├── officecli.py      # 改: DocTool 加 add_image 方法
├── tools.py          # 改: 加 query_vehicle + add_image 工具
└── prompts.py        # 改: 加交通文档专项指引
main.py               # 改: _format_tool_call 加新工具格式化
```

## 验证
1. 单元：mock_query 一致性（同车牌同结果）、多车/无车分支
2. officecli 插图实测：create docx → add picture → view（确认图片进文档）
3. 端到端：`python main.py "写一份交通事故报告，车牌京A12345和京B67890"` 
   → ask_user 收集 → query_vehicle 逐个查 → 多车确认分支 → 写文档含图片/表格/记录

## 关键设计决策
- **返回清单给 LLM 而非工具内 interrupt**：符合你选择的"返回清单给 LLM"，职责清晰；query_vehicle 纯查询无副作用，多车确认走标准 ask_user 流程。
- **mock 用种子保证一致性**：同车牌多次查结果相同，避免 agent 困惑。
- **图片用占位 URL**：先验证流程跑通，真实图片后期补。
- **真实接口预留**：mock_query 内留 `MOCK_MODE` 开关，实现真实接口时只改这一个函数。