# kdocs-cli 安装与使用

`kdocs-cli` 是金山文档官方 CLI（[仓库](https://github.com/kdocs-app/kdocs-skill)），
预编译二进制，**不在 npm 上**。npm 上的 `docli` / `kdocs` / `wps-cli` 都是无关的第三方包。

## 安装（Windows）

官方 `setup.ps1` 的下载规律：

```
{CDN}/v{版本}/releases/kdocs-cli-{版本}-windows-{arch}.zip
CDN 默认: https://wpsai.wpscdn.cn/skillhub/pro
```

版本号取自 SKILL.md 的 `version` 字段。手动安装流程（含 SHA256 校验）：

```powershell
$CdnBase='https://wpsai.wpscdn.cn/skillhub/pro'; $Version='2.7.1'; $Arch='amd64'
$InstallDir=Join-Path $env:LOCALAPPDATA 'kdocs-cli'
$Archive="kdocs-cli-$Version-windows-$Arch.zip"
$Tmp=Join-Path ([System.IO.Path]::GetTempPath()) "kdocs-install-$PID"
New-Item -ItemType Directory -Path $Tmp -Force | Out-Null
Invoke-WebRequest "$CdnBase/v$Version/releases/$Archive" -OutFile (Join-Path $Tmp $Archive) -UseBasicParsing
Invoke-WebRequest "$CdnBase/v$Version/releases/checksums.txt" -OutFile (Join-Path $Tmp 'checksums.txt') -UseBasicParsing
$expected=((Get-Content (Join-Path $Tmp 'checksums.txt') | Where-Object { $_ -match [regex]::Escape($Archive) }) -split '\s+')[0].ToLower()
$actual=(Get-FileHash (Join-Path $Tmp $Archive) -Algorithm SHA256).Hash.ToLower()
if($expected -ne $actual){ throw "SHA256 mismatch" }
Expand-Archive (Join-Path $Tmp $Archive) -DestinationPath $Tmp -Force
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
Copy-Item (Get-ChildItem $Tmp -Recurse -Filter 'kdocs-cli.exe' | Select-Object -First 1).FullName (Join-Path $InstallDir 'kdocs-cli.exe') -Force
[Environment]::SetEnvironmentVariable('PATH', "$InstallDir;" + [Environment]::GetEnvironmentVariable('PATH','User'), 'User')
```

装完 `kdocs-cli version` 应输出版本号。

## 登录

```bash
kdocs-cli auth login      # 打印授权链接，浏览器确认后自动存 Token 到系统密钥链
kdocs-cli auth status     # authenticated: true 即成功
```

Token 存系统密钥链，不要写进文件或对话。有效期约 8760 小时。

## 常用命令

```bash
# 工作表列表（拿 sheetId）
kdocs-cli sheet get-sheets-info url="<URL>" --compact

# 读区域（rowFrom/rowTo/colFrom/colTo 均 0-based，range 必须是对象）
kdocs-cli sheet get-range-data '{"url":"<URL>","worksheet_id":11,
  "range":{"rowFrom":0,"rowTo":59,"colFrom":0,"colTo":9}}'

# 写区域
kdocs-cli sheet update-range-data '{"url":"<URL>","worksheet_id":11,
  "rangeData":[{"rowFrom":0,"rowTo":0,"colFrom":0,"colTo":0,"formula":[["日期"]]}]}'
```

**中文/长参数不要用 `key=value`**，PowerShell 会破坏 UTF-8；用 `--file payload.json`。
生成 JSON 用 Python/Node，不要用 `ConvertTo-Json`（会带 BOM）。

## 错误速查

| 现象 | 处理 |
|---|---|
| `400001 range 须为对象` | `range` 要包成对象，不能传数组 |
| `400006` / 鉴权失败 | `kdocs-cli auth login` 重新登录 |
| `429001` / `429002` | 限频/熔断，立即停止请求，等到恢复时间再试 |
| `403` / 无权限 | 停止重试，告知用户无权限 |
| `unknown action` | `kdocs-cli upgrade -y` 升级 |
