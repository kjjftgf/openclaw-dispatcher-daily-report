# 数据日报生成器

自动分析 Aeolus 实时出勤数据，生成骑手当日业绩达标报告（.xlsx）

## 这是干嘛的

每天从 Aeolus 导出实时出勤数据，一键算清楚：

- **今天排班多少人、多少人达标、谁没达标**
- **9个时段出勤率和达标率** — 早高峰、午高峰、晚高峰、夜宵等
- **档位分布** — 标准5到不足标准1各档多少人
- **谁差一点就能升档** — 差1单、差2单、差几分钟的明细
- **谁有冲三档机会** — 标准2骑手里最有希望冲上标准3的

输出一份完整 Excel，站长直接能看。

## 数据来源

从 Aeolus 网页手动导出：
- 实时出勤数据 → `~/Downloads/2026-MM-DD_HH_MM_SS实时出勤数据.xlsx`
- 排班数据 → `~/Downloads/2026-MM-DD到2026-MM-DD排班数据.xlsx`

脚本自动取 `~/Downloads/` 里最新的文件。

## 使用方法

```bash
# 生成日报
cd ~/.openclaw/workspace/data-analyst
python3 generate_report.py

# 验证报告格式
python3 validate_report.py
```

输出：`~/Desktop/配送日报_MM-DD.xlsx`

## 核心脚本

| 文件 | 作用 |
|:----|:----:|
| `generate_report.py` | 日报生成主入口（六板块） |
| `validate_report.py` | 格式校验，跑通才算生成成功 |
| `scripts/attendance_analysis.py` | 各时段出勤率、达标率计算 |
| `scripts/level_analysis.py` | 档位分布、接近升级骑手 |
| `scripts/focus_analysis.py` | 冲三档+全部升档机会 |
| `scripts/fetch_aeolus_data.py` | Aeolus 数据解析 |

## 日报六板块

1. **全天达成** — 谁达基线（≥6h & ≥18单），谁没达
2. **时段未达标** — 9个时段的出勤、达标情况
3. **档位达成** — 各档位分布 + 接近升级骑手 + 低于基线名单
4. **重点跟进** — 冲三档 + 全部升档机会 + 汇总
5. **0-2档不够明细** — 低档骑手到下一档还差多少
6. **档位明细** — 每位骑手到下一档的差距明细

## 技术栈

- Python 3 + openpyxl
- 输出 WPS 兼容格式
- 配色、字体、对齐等已适配 WPS 办公
