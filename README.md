# Google Sheet Watcher — Proxy 版

这个版本会显式让 Google Drive API / Google Sheets API 通过代理访问。

## 先更新依赖

在工具目录执行：

```powershell
python -m pip install -r requirements.txt
```

## 默认代理

默认：

```text
http://127.0.0.1:7890
```

因此正常可直接：

```powershell
python sheet_watcher.py "你的 Google Sheets URL"
```

启动日志应该出现：

```text
Google API 使用代理：http://127.0.0.1:7890
```

## 如果你的代理端口不是 7890

例如本机 HTTP / Mixed Port 是 7897：

```powershell
python sheet_watcher.py "你的 Google Sheets URL" --proxy http://127.0.0.1:7897
```

v2rayN 常见 HTTP 端口如果是 10809：

```powershell
python sheet_watcher.py "你的 Google Sheets URL" --proxy http://127.0.0.1:10809
```

以你代理软件实际显示的 HTTP / Mixed Port 为准。

## 临时关闭显式代理

```powershell
python sheet_watcher.py "你的 Google Sheets URL" --proxy ""
```

## 这版的网络路径

```text
sheet_watcher
    |
    +--> OAuth token refresh (requests)
    |       |
    |       +--> HTTP_PROXY / HTTPS_PROXY
    |
    +--> Drive API / Sheets API
            |
            +--> google-auth-httplib2
                    |
                    +--> httplib2.Http
                            |
                            +--> 显式 ProxyInfo
```

此外，以下网络错误现在也会进入指数退避重试：

- socket timeout
- WinError / OSError
- connection error
- httplib2 transport error
- HTTP 408 / 429 / 500 / 502 / 503 / 504

默认单次连接 timeout 为 30 秒，单轮最多重试 5 次。
