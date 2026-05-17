# 足球彩票分析工具

通过 Playwright 浏览器自动化从 live.500.com 抓取实时比赛数据，生成足球彩票分析 Excel 模板。

## 分发给不懂技术的用户（推荐方式）

将整个应用打包成一个独立文件夹，**对方无需安装 Python 或任何东西**，解压后双击即可运行：

```bash
# 首次需安装 PyInstaller
uv sync --extra build

# 确保 Chromium 已安装
uv run playwright install chromium

# 打包
uv run python build_package.py
```

打包结果在 `dist/足球分析工具/`，将整个文件夹压缩成 zip 发给 Windows 用户，对方：

1. 解压 zip
2. 双击 `足球分析工具.exe`
3. 浏览器自动打开，开始使用

> 注意：PyInstaller 不支持交叉编译，给 Windows 用户需要在 Windows 上打包，macOS 同理。

## 自己开发/使用

### 环境要求

- Python 3.12+
- [uv](https://github.com/astral-sh/uv) 包管理器

### 安装

```bash
uv sync
uv run playwright install chromium
```

### 启动

```bash
uv run launcher                    # Web 界面 + 自动打开浏览器
uv run web-app --port 5100         # 仅启动服务
```

浏览器访问 `http://127.0.0.1:5100`。

### 命令行模式

```bash
uv run generate-browser-template
uv run generate-browser-template -o output.xlsx --max-matches 10
uv run generate-browser-template --asian-handicap
uv run generate-browser-template --no-headless
```

| 参数 | 说明 |
|------|------|
| `-o, --output` | 输出 Excel 文件名 |
| `-u, --url` | 要抓取的 URL |
| `-m, --max-matches` | 最多比赛数量 |
| `--asian-handicap` | 亚洲盘口 + 百家欧赔凯利数据 |
| `--no-headless` | 显示浏览器窗口（调试用） |

## Web 界面功能

- 实时比赛数据抓取和预览
- 多次抓取累积保存到"今日表格"
- 批量管理（删除单次记录、清空全部）
- 下载本次 / 今日累积 Excel 模板
- 页面内一键检测/安装 Chromium 浏览器
- 支持亚洲盘口 + 百家欧赔凯利数据获取

## 数据字段

| 字段 | 说明 |
|------|------|
| 场次 | 周一001, 周二002 等 |
| 联赛 | 德甲, 意甲 等 |
| 轮次 | 第17轮, 半决赛 等 |
| 比赛时间 | MM-DD HH:MM |
| 状态 | 未, 进行中, 完场 |
| 主队/排名 | 主队名称及排名 |
| 让球 | 半球, 球半, 受半球 等 |
| 客队/排名 | 客队名称及排名 |
| 胜负奖金 / 让球奖金 | 赔率数据 |
| 平均欧赔 / 威廉 / 澳彩 / 365 / 皇者 | 各家赔率 |

## 项目结构

```
├── web_app.py                   # Flask Web 应用
├── launcher.py                  # 跨平台启动器
├── browser_bet_scraper.py       # 浏览器自动化抓取核心
├── build_package.py             # PyInstaller 打包脚本
├── data/                        # 导出的 Excel 文件（已 gitignore）
├── templates/
│   └── index.html               # Web 前端页面
├── 启动足球分析工具.command       # macOS 启动脚本
├── 启动足球分析工具.bat           # Windows 启动脚本
└── pyproject.toml               # 项目配置
```
