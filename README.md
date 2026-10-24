# 德育分自动化系统

班级德育分（加减分）日常记录、周纪实表 / 月度总表自动生成，附带寝室星级打分与 AI（MCP）对接。

- 单文件 Tkinter 应用（`软件说明/德育分自动化系统.py`），Windows 平台
- 数据全部落在程序目录（`名单模板.xlsx` + `20XX年X月/`），无外部服务依赖
- 更新走 GitHub Releases 热更新，`名单模板.xlsx` 与历史数据永不被覆盖

## 普通用户：安装与更新

1. 到 [Releases](https://github.com/xiaohaifale-QWQ/moral-edu-automation/releases) 下载
   `deyufen-setup-x.x.x.exe`，双击安装（默认装到用户目录，无需管理员）。
2. 首次运行会引导「班级名单」导入：从 Excel 复制（学号、姓名、寝室号、角色）粘贴即可，**或直接把名单 xlsx/csv 文件拖进软件窗口**。
3. 以后有新版本：每次启动软件会自动检查并弹窗询问（无新版/网络不通不打扰）；也可在左下角「设置 → 检查更新」手动触发，自动下载、校验 sha256、替换、重启。
   更新不动 `名单模板.xlsx`、`20XX年X月/`、`更新配置.json`。
4. 国内网络打不开 GitHub 时：「设置 → 更新源」把镜像换成
   `https://ghfast.top/https://raw.githubusercontent.com/...` 这类加速地址，或到 Releases 手动下载。

### 首次运行提示（SmartScreen）

exe 未购买代码签名证书，Windows 可能弹「已保护你的电脑」：
点 **更多信息 → 仍要运行** 即可。这是未签名程序的正常现象，不是病毒。

### AI 对接（可选）

安装包已内置 `AI对接\mcp_server.exe`，用户机不需要装 Python。
把 `AI对接\MCP配置.json` 里的内容加到支持 MCP 的客户端（Cherry Studio / Claude Desktop 等）即可。
挪动安装目录后，在左下角「设置 → 重新生成 MCP 配置」修复路径。

## 开发者：构建与发布

```bash
# 源码自检（改完代码必跑）
python -m py_compile 软件说明/德育分自动化系统.py
python _ui-work/verify_ai_list.py        # 在仓库外的 _ui-work 目录
python tools/test_roster_write.py        # 名单导入写表逻辑
python tools/test_updater.py             # 更新器 A–E 用例
python tools/test_blank_run.py           # 空白模板 GUI 验收

# 三个 PyInstaller 产物（主程序 / mcp_server / update）
python -m PyInstaller --clean --noconfirm 软件说明/德育分自动化系统.spec   # 产物在 软件说明/dist
cd src && python -m PyInstaller --clean --noconfirm mcp_server.spec
cd src && python -m PyInstaller --clean --noconfirm update.spec

# 组装 payload → 打更新包（自动回填 version.json）
python tools/make_blank_template.py && python tools/check_blank_template.py
python tools/make_payload.py
python tools/make_update_zip.py 1.7.0 --notes "……"

# 安装器（二选一）
#   A：Inno Setup —— ISCC installer/setup.iss
#   B：自写安装器 —— cd installer && python -m PyInstaller installer.spec

# 发布自检
python tools/check_manifest.py
```

发版 SOP：改 `APP_VERSION` / `setup.iss` / `version.json` 三处版本一致 → 打包 →
GitHub Releases 建 tag `vX.Y.Z` 上传 `release/deyufen-X.Y.Z-win64.zip` → push main。

## 隐私红线

- `名单模板.xlsx`（真实名单）、`20XX年X月/`（分数）、`deepseek_config.json`（API Key）、
  `_ai_bridge/`、`名单模板.自动备份-*.xlsx` 一律 **不入库、不进分发包**（.gitignore 已配好，
  打包脚本只做白名单复制）。
- 分发用空白模板由 `tools/make_blank_template.py` 从真实模板生成，自检不过禁止打包。

## 目录

```
软件说明/            主程序源码 + spec + 使用文档
AI对接/              mcp_server.py（MCP 服务，stdio）
src/                 update.py 更新器 + 两个小件的 spec
installer/           setup.iss（Inno）与 installer.py（自写安装器）
tools/               空白模板 / payload / 更新包 / 自检脚本
version.json         热更新清单（客户端唯一契约）
release/             本地产出的更新包（不入库）
```
