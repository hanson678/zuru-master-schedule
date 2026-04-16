# ZURU 总排期入单系统

将 ZURU 客户的 PO Excel 文件自动录入 `Z盘` 总排期生产表，支持修改单、新单、双排期货号展开、黑名单过滤、整行填充汇总等业务功能。

## 核心功能

| 功能 | 说明 |
|------|------|
| **批量上传 PO** | 拖放多个 Excel PO 文件，一键解析并写入总排期 |
| **修改单识别** | 按 `PO号 + 货号 + line_no` 三元组匹配，已有行只更新变化字段并蓝色标记 |
| **新单生成** | 未匹配到的新货号不写 Z 盘，生成独立 Excel 按分排期 sheet 分组供粘贴 |
| **重复 PO 自动去重** | 同 PO 号多个版本（R1/R2/Rev.1 等），自动保留最新版，橙色提示 |
| **货号黑名单** | `data/ignore_items.json` 配置的货号前缀直接跳过，不入排期 |
| **双排期货号展开** | `data/dual_schedule_map.json` 配置的货号一行拆多行，自动插入系列号 |
| **有填充行汇总** | 扫描总排期整行填充的行（连续≥6格非白填充），按分排期 sheet 分类统计 |
| **分排期归属查找** | 上传后展示每个货号属于哪个分排期文件 |
| **切换路径** | 前端直接切换总排期文件位置，支持本地/网络盘 |

## 技术栈

- **Python 3.12** + Flask 3.x
- **openpyxl**：只读建索引（快速查找已有行）
- **WPS COM (Ket.Application)**：写入 Z 盘排期（保留公式与格式，仅限 Windows）
- **Bootstrap 5.3**：前端

## 运行前提

| 条件 | 说明 |
|------|------|
| **Windows 系统** | COM 操作依赖，仅支持 Windows |
| **WPS Office** | 需包含 WPS 表格组件（Ket.Application COM 接口） |
| **共享盘权限** | 能访问总排期所在目录（通常为 Z:\各客排期\ZURU生产排期） |
| **Python 3.12+** | 源码版运行需要 |

## 快速启动

### 源码方式

```bash
# 1. 克隆仓库
git clone https://github.com/<your-user>/zuru-master-schedule.git
cd zuru-master-schedule

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动服务
python app.py
```

浏览器访问 http://localhost:5003

### Windows 一键启动

双击 `一键启动.vbs`（静默启动）或 `启动系统.bat`（带控制台）。

首次使用前先运行 `安装依赖.bat` 安装 Python 依赖。

> **为什么不提供 Docker 部署？** 总排期入单依赖 WPS COM 接口（`Ket.Application`）写入 Excel，此接口只能在 Windows 宿主机上调用，Docker 容器内无法运行。因此只提供 Windows 原生部署方式。

## 目录结构

```
zuru-master-schedule/
├── app.py                       # Flask 主程序（上传、去重、黑名单过滤、API路由）
├── master_schedule.py           # 写入核心（openpyxl建索引 + WPS COM写入）
├── excel_po_parser.py           # ZURU PO Excel 解析器
├── generate_yellow_summary.py   # 有填充行汇总扫描
├── base_path.py                 # PyInstaller 打包路径辅助
├── requirements.txt
│
├── templates/
│   └── master.html              # 主页面（拖拽上传 + 结果展示）
│
├── data/
│   ├── config.json              # 系统配置（Z盘路径、端口）
│   ├── ignore_items.json        # 货号黑名单（不入排期的货号）
│   ├── dual_schedule_map.json   # 双排期货号配置
│   ├── item_cn_name_map.json    # 货号 → 中文名直查表
│   └── sub_schedule_map.json    # 货号 → 分排期归属
│
├── 启动系统.bat                 # Windows 控制台启动
├── 一键启动.vbs                 # Windows 静默启动
├── 安装依赖.bat                 # 首次 Python 依赖安装
│
├── DEPLOY.md                    # 详细部署文档
└── README.md
```

## 使用流程

1. **首次启动**
   - 访问 http://localhost:5003
   - 如果 Z 盘路径与默认不同，点击"切换路径"指向你的总排期 xlsx 文件
2. **上传 PO**
   - 拖拽多个 Excel PO 文件到上传区
   - 点击"写入总排期"
3. **查看结果**
   - 修改单明细：显示修改的字段和行号
   - 新单列表：下载生成的新单 Excel（按分排期 sheet 分组）
   - 重复 PO 去重报告、黑名单忽略报告、分排期归属图
4. **有填充行汇总**
   - 点击"有填充行汇总"按钮
   - 扫描总排期中所有整行填充的行，按分排期 sheet 分类统计

## 配置说明

### `data/config.json`

```json
{
  "z_drive_path": "Z:\\各客排期\\ZURU生产排期",
  "port": 5003
}
```

### `data/ignore_items.json` — 黑名单

```json
{
  "ignore_items": ["15790", "XXXX"]
}
```

匹配规则：按货号开头的数字前缀匹配，如 `15790` 会匹配 `15790-S001`、`15790-S003` 等所有规格。

### `data/dual_schedule_map.json` — 双排期货号

```json
{
  "77785": {"targets": ["77673", "77711"], "mode": "mid_insert"},
  "92123": {"targets": ["9298", "92104"], "mode": "append"}
}
```

Mode 说明：
- `append`：末尾追加（`92123-S001` → `92123-S001-9298`）
- `slt_insert`：SLT后插入
- `mid_insert`：S00x后插入
- `s_insert`：S00x前插入
- `none`：不加系列号

## 关键业务规则

### 总排期列布局（COL）

```
A=接单期  B=客户   C=走货国  D=PO号    E=客PO   F=SKU    G=货号#  H=中文名
I=数量    J=内箱   K=外箱    L=总箱    M=出货期  N=验货期
...
Y=备注    Z=跟单   AA=单价   AB=金额
```

### 验货日期

- **Fuggler (157 开头)**：出货 - 2 天
- **其他货号**：出货 - 4 天
- **河源工厂**：周六→周五、周日→周一
- **非河源**：周日→周一

### 匹配键

- 修改单识别：`(PO号, 货号, line_no)` 三元组
- 新单回落：`(PO号, 货号)` 二元组兜底

## 更多文档

- [DEPLOY.md](DEPLOY.md) — 详细部署步骤与故障排查
