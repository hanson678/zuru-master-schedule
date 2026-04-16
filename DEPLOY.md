# ZURU 总排期入单系统 — 部署指南

本文档介绍如何在目标机器上部署本系统。推荐使用 **Windows 原生方式**（Python 源码 + WPS）。

---

## 一、前提条件

| 条件 | 必须 |
|------|------|
| Windows 10/11 操作系统 | ✅ |
| WPS Office（包含 WPS 表格组件 Ket.Application） | ✅ |
| Python 3.12+ | ✅（源码方式） |
| 对总排期 Excel 文件的读写权限（通常为 Z 盘或局域网共享盘） | ✅ |

---

## 二、源码安装（推荐）

### 1. 克隆仓库

```powershell
git clone https://github.com/<your-user>/zuru-master-schedule.git
cd zuru-master-schedule
```

### 2. 安装 Python 依赖

运行 `安装依赖.bat`，或手动：

```powershell
pip install -r requirements.txt
```

依赖包括：flask、openpyxl、pdfplumber、apscheduler、pywin32（Windows COM）。

### 3. 配置总排期文件路径

编辑 `data/config.json`：

```json
{
  "z_drive_path": "Z:\\各客排期\\ZURU生产排期",
  "port": 5003
}
```

- `z_drive_path`：总排期 xlsx 所在目录（也可通过前端"切换路径"按钮动态修改）
- `port`：Web 服务端口（默认 5003，被占用可改其他）

### 4. 启动服务

三种方式任选其一：

```powershell
# 方式A：控制台（可看日志）
启动系统.bat

# 方式B：静默启动（无窗口）
一键启动.vbs

# 方式C：手动
python app.py
```

浏览器访问 http://localhost:5003

---

## 三、防火墙与远程访问

系统默认监听 `0.0.0.0:5003`，局域网内其他电脑可通过 `http://<本机IP>:5003` 访问。

如果无法访问，检查：
1. Windows 防火墙是否放行 TCP 5003 端口
2. 本机 IP（运行 `ipconfig` 查看）
3. 如启用了企业域/VPN，需确认策略允许

---

## 四、常见问题

### Q1. 启动报错 `ImportError: No module named 'win32com'`

```
解决：pip install pywin32
```

### Q2. 总排期文件显示"被占用"

同事在 WPS 中打开了总排期文件，让对方关闭后点"刷新"按钮。

### Q3. 前端报 "COM 错误" 或 "Ket.Application 未注册"

确认 WPS Office 已安装并至少运行过一次（初次安装需激活 COM 注册）。

### Q4. 分排期归属显示"未识别"

货号没在 `data/sub_schedule_map.json` 里。如需补充映射，编辑该 JSON，格式：
```json
{
  "9548UQ1": [{"file": "2025年ZURU...xlsx", "sheet": "9548"}]
}
```

### Q5. 如何添加黑名单货号

编辑 `data/ignore_items.json`：
```json
{
  "ignore_items": ["15790", "其他货号"]
}
```
保存即生效，无需重启（系统自动 mtime 检测）。

### Q6. 端口被占用

改 `data/config.json` 的 `port` 字段，或修改 `app.py` 最后一行。

---

## 五、日志与故障排查

- 运行日志：`data/ops.log`
- Flask 控制台日志：运行 `启动系统.bat` 方式可直接看
- 上传失败的 PO 文件会保留在 `uploads/` 目录

---

## 六、映射表维护

| 文件 | 维护场景 | 格式 |
|------|---------|------|
| `data/ignore_items.json` | 新增不入排期的货号 | 见 [README.md](README.md) |
| `data/dual_schedule_map.json` | 新增双排期货号 | 见 [README.md](README.md) |
| `data/item_cn_name_map.json` | 货号中文名补全 | 自动从总排期扫描构建 |
| `data/sub_schedule_map.json` | 货号 → 分排期文件归属 | 运行扫描脚本生成 |

---

*华登玩具集团 · ZURU 总排期入单系统*
