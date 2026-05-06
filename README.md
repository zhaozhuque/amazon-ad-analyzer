# Amazon Ad Analyzer 亚马逊广告诊断工具

Upload your Amazon advertising report, get a full diagnosis in 3 seconds.

上传亚马逊广告报告，3秒生成全面诊断，自动生成否定词清单。

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Version](https://img.shields.io/badge/Version-1.1-brightgreen.svg)

---

## 这个工具能帮你做什么？

你从亚马逊后台下载搜索词报告，拖进工具，3秒后得到一份完整的诊断Excel：

- **哪些词在烧钱** → 花费高但0转化，立即否定
- **哪些词该加投** → ACoS低转化高，提高竞价抢量
- **否定词清单** → 按Campaign分组，可直接复制到后台（v1.1 新增）
- **哪些词有潜力** → ACoS临界可优化，调整竞价或listing
- **哪个Campaign亏钱** → 活动维度汇总，问题一目了然
- **哪个ASIN在亏** → 商品维度盈亏分析

完全离线运行，数据不上传任何服务器。

---

## 快速开始

### 方式一：下载 EXE（推荐，无需安装 Python）

1. 进入 [Releases](../../releases) 页面
2. 下载最新版 `AmazonAdAnalyzer.zip`
3. 解压后双击 `AmazonAdAnalyzer.exe` 运行

### 方式二：Python 运行

```bash
git clone https://github.com/zhaozhuque/amazon-ad-analyzer.git
cd amazon-ad-analyzer
pip install -r requirements.txt
python amazon_ad_tool.py
```

命令行模式（不打开界面）：

```bash
python amazon_ad_tool.py 搜索词报告.xlsx 推广商品报告.xlsx
```

---

## 支持的报告类型

| 报告 | 格式 | 语言 |
|------|------|------|
| 商品推广 - 搜索词报告 | .xlsx / .csv | 中文 / English |
| 商品推广 - 推广的商品报告 | .xlsx / .csv | 中文 / English |

搜索词报告为必选，推广商品报告为可选。两份都上传可以获得 ASIN 维度分析。

---

## 诊断报告包含 9 个 Sheet

| Sheet | 说明 |
|-------|------|
| 📊 总览仪表盘 | 核心KPI + 关键词健康分布 + 优化建议 |
| 🔴 烧钱词&否定词 | 花费高无转化的词，按花费排序 |
| 🟢 优质词&潜力词 | 转化好的词，按销售额排序 |
| 📊 Campaign分析 | 各广告活动效果对比 + 问题词计数 |
| 🔤 匹配类型分析 | 精准/广泛/自动效果对比 |
| 📦 推广商品分析 | ASIN维度盈亏诊断 |
| ❌ 否词推荐 | 按Campaign分组的否定词清单，标注紧急程度（v1.1） |
| 📋 否定词复制列表 | 纯净词表，直接复制粘贴到亚马逊后台（v1.1） |
| 📋 搜索词明细 | 全部数据 + 诊断标签 + 颜色标记 |

---

## 否词推荐逻辑（v1.1 新增）

工具会自动识别以下类型的无效搜索词并推荐否定：

| 条件 | 紧急程度 | 建议否定类型 |
|------|----------|-------------|
| 花费 ≥ $5 且 0 订单 | ‼️ 立即否定 | 精准否定 |
| 点击 ≥ 8 次且 0 订单 | ⚠️ 建议否定 | 精准否定 |
| ACoS > 80% 严重亏损 | ⚠️ 建议否定 | 精准否定 |
| 高曝光 + 0 转化（流量不精准） | 📌 可选否定 | 短语否定 |

否词清单按 Campaign 分组展示，方便你逐个去后台操作。同时提供独立的"否定词复制列表"Sheet，选中第一列直接复制粘贴即可。

---

## 自定义阈值

工具界面上可以调整，适配不同产品利润率：

| 参数 | 默认值 | 含义 |
|------|--------|------|
| 高ACoS阈值 | 40% | 超过此值标记为高ACoS |
| 优质词ACoS上限 | 20% | 低于此值标记为优质词 |
| 烧钱花费阈值 | $5 | 花费超过且0单标记为烧钱词 |

---

## 自行打包 EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name AmazonAdAnalyzer amazon_ad_tool.py
```

打包完成后 `dist/AmazonAdAnalyzer.exe` 即为独立程序。

---

## 隐私说明

- 所有分析在本地完成，不联网、不上传数据
- 不收集任何用户信息
- 开源代码，可自行审查

---

## 更新日志

### v1.1
- 新增"否词推荐"Sheet，按Campaign分组，标注紧急程度和否定类型
- 新增"否定词复制列表"Sheet，可直接复制到亚马逊后台
- 支持4种否词识别逻辑（高花费无单、多点击无单、极高ACoS、高曝光无转化）

### v1.0
- 首个发布版本
- 支持搜索词报告 + 推广商品报告分析
- 7个维度诊断报告
- 图形界面 + 命令行双模式

---

## License

MIT License - 可自由使用、修改和分发。

## Author

[@zhaozhuque](https://github.com/zhaozhuque)
