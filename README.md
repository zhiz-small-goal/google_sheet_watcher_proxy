# Google Sheet Watcher — 多文件 + 代理 + 日志版

## 你的场景：无需填写 example

如果你希望自动扫描自己账号拥有的全部 Google Sheets：

```powershell
python sheet_watcher.py --all --owned-only
```

`watchlist.example.json` **不用填，也不用使用**。

它只是可选模式：只有当你以后想“只监控指定的几份 Sheet”时，才复制成
`watchlist.json` 并填写 URL / file_id。

---

## 推荐先验证一次

```powershell
python sheet_watcher.py --all --owned-only --once
```

程序会自动通过 Google Drive API 扫描：

```text
当前账号拥有的 Google Sheets
├── 学习评测系统
├── 做饭记录
└── 以后新建的其他 Sheet
```

长期运行：

```powershell
python sheet_watcher.py --all --owned-only
```

每轮都会重新扫描，因此以后新增 Sheet 不需要改配置。

---

## 日志

默认写入：

```text
logs/
└── sheet_watcher.log
```

同时终端仍然显示日志。

日志轮换策略：

```text
sheet_watcher.log      当前日志
sheet_watcher.log.1    历史
sheet_watcher.log.2
...
sheet_watcher.log.5
```

每个日志接近 5 MiB 时自动轮换，最多保留 5 份历史日志，避免长期自启导致日志无限增长。

自定义日志路径：

```powershell
python sheet_watcher.py --all --owned-only --log "D:\Logs\sheet_watcher.log"
```

---

## 开机 / 登录自启推荐参数

Windows 任务计划程序中：

程序：

```text
E:\downld\google_sheet_watcher\.venv\Scripts\python.exe
```

参数：

```text
"E:\downld\google_sheet_watcher\sheet_watcher.py" --all --owned-only
```

起始于：

```text
E:\downld\google_sheet_watcher
```

这样即使没有终端窗口，也可以直接检查：

```text
E:\downld\google_sheet_watcher\logs\sheet_watcher.log
```

判断程序是否启动、扫描了多少 Sheet、哪些文件同步成功、哪些失败。

---

## 代理

默认：

```text
http://127.0.0.1:7890
```

如果代理端口不同：

```powershell
python sheet_watcher.py --all --owned-only --proxy http://127.0.0.1:7897
```

---

## 输出

每个 Google Sheets 文件独立目录：

```text
downloads/
├── 学习评测系统__xxxxxxxx/
│   ├── 学习评测系统.xlsx
│   └── 每个 tab.csv
└── 做饭记录__xxxxxxxx/
    ├── 做饭记录.xlsx
    └── 每个 tab.csv
```

每个文件独立记录 Drive version，只有发生变化才重新导出。
