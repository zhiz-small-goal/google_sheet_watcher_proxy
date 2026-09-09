# Google Sheet Watcher — 多文件 + 代理版

截至 2026-09-09。

支持三种模式：

```text
1. 单文件
2. --all 自动发现全部 Google Sheets
3. --watchlist 只监控指定文件
```

每个 Spreadsheet 都会导出：

```text
完整 Spreadsheet -> 1 个 .xlsx
每个 tab           -> 1 个 .csv
```

---

## 安装 / 更新依赖

激活你的 `.venv` 后：

```powershell
python -m pip install -r requirements.txt
```

---

## 推荐：你的场景直接用 --all

```powershell
python sheet_watcher.py --all
```

默认每 60 秒：

1. 重新扫描账号当前可见的 Google Sheets
2. 读取每个文件的 Drive `version`
3. 只有 version 变化的文件才重新下载
4. 新建的 Google Sheet 会在下一轮自动发现

如果你只想备份自己拥有的文件，不要备份别人分享给你的 Sheet：

```powershell
python sheet_watcher.py --all --owned-only
```

---

## 只运行一次

先测试：

```powershell
python sheet_watcher.py --all --once
```

确认下载目录正常后，再去掉 `--once` 长期运行。

---

## 指定 watchlist

复制：

```text
watchlist.example.json
```

为：

```text
watchlist.json
```

把 URL 换成真实 Sheet URL：

```json
{
  "sheets": [
    "https://docs.google.com/spreadsheets/d/学习评测系统ID/edit",
    "https://docs.google.com/spreadsheets/d/做饭记录ID/edit"
  ]
}
```

运行：

```powershell
python sheet_watcher.py --watchlist watchlist.json
```

watchlist 每轮都会重新读取，因此你编辑 `watchlist.json` 后不需要重启程序。

---

## 单文件模式仍兼容

```powershell
python sheet_watcher.py "Google Sheets URL"
```

---

## 代理

默认：

```text
http://127.0.0.1:7890
```

例如你的代理是 7897：

```powershell
python sheet_watcher.py --all --proxy http://127.0.0.1:7897
```

不使用代理：

```powershell
python sheet_watcher.py --all --proxy ""
```

---

## 输出目录

例如账号有：

```text
学习评测系统
做饭记录
```

会得到类似：

```text
downloads/
├── 学习评测系统__1cJwzW8m/
│   ├── 学习评测系统.xlsx
│   ├── 学习评测系统__Learning Items.csv
│   ├── 学习评测系统__Daily Evaluation.csv
│   └── ...
│
└── 做饭记录__AbCdEf12/
    ├── 做饭记录.xlsx
    ├── 做饭记录__菜谱.csv
    ├── 做饭记录__复盘.csv
    └── ...
```

目录名附带 file_id 前 8 位，用于避免两个同名 Google Sheets 覆盖彼此。

---

## 多文件状态

状态文件升级为：

```json
{
  "schema_version": 2,
  "files": {
    "FILE_ID_A": {
      "name": "学习评测系统",
      "version": "123"
    },
    "FILE_ID_B": {
      "name": "做饭记录",
      "version": "56"
    }
  }
}
```

旧版单文件状态会自动迁移。

---

## 可靠性设计

```text
扫描目标
  ↓
逐文件读取 Drive version
  ↓
version 没变化 ──────> 跳过
  ↓ 有变化
导出完整 XLSX
  ↓
导出所有 tab CSV
  ↓
全部成功
  ↓
原子更新状态中的该 file_id/version
```

关键行为：

- 一个 Sheet 同步失败，不会阻塞其他 Sheet。
- 失败文件不会推进 version，下轮会重试。
- XLSX / CSV 使用临时文件完成后再 `os.replace()`。
- 网络 timeout、连接错误、HTTP 408/429/5xx 指数退避。
- `--all` 每轮重新扫描，所以后来新建的 Sheet 会自动加入。
- CSV 包含隐藏 tab。
- CSV 使用 UTF-8 BOM，方便 Windows Excel 打开中文。
- CSV 保存格式化显示值；公式和格式以 XLSX 为主。

---

## 建议你的启动命令

先验证一次：

```powershell
python sheet_watcher.py --all --owned-only --once
```

如果 `downloads/` 里能看到「学习评测系统」和「做饭记录」，长期运行：

```powershell
python sheet_watcher.py --all --owned-only
```
