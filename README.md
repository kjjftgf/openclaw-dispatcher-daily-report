# 配送日报自动生成技能

> OpenClaw 技能 — dispatcher-data-analysis

## 功能

从 Aeolus 实时出勤数据生成完整的配送日报 Excel（六板块）：

| 板块 | 内容 |
|:----:|:----:|
| 全天达成 | 基线达标/未达标骑手统计 |
| 时段未达标 | 9个时段出勤率、达标率分析 |
| 档位达成 | 标准0~5档位分布+接近升级骑手 |
| 重点跟进 | 冲三档+全部升档机会+汇总 |
| 0-2档不够明细 | 低档骑手到下一档差距 |
| 档位明细 | 0~4档每位骑手晋升差距 |

## 使用方法

```bash
# 1. 从 Aeolus 网页导出实时出勤 Excel 到 ~/Downloads/

# 2. 生成日报
cd ~/.openclaw/workspace/data-analyst
python3 generate_report.py

# 3. 验证报告
python3 validate_report.py
```

输出位置：`~/Desktop/配送日报_{日期}.xlsx`

## 脚本说明

| 文件 | 说明 |
|:----|:----:|
| `generate_report.py` | 日报生成主脚本（六板块） |
| `validate_report.py` | 格式验证脚本 |
| `scripts/attendance_analysis.py` | 时段考勤分析 |
| `scripts/level_analysis.py` | 档位达成分析 |
| `scripts/focus_analysis.py` | 重点跟进分析 |
| `scripts/fetch_aeolus_data.py` | Aeolus 数据获取 |
| `references/*.md` | 标准/规则参考文档 |
