# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python CLI tool that generates comprehensive Excel templates for soccer betting analysis with real match data scraped from 500.com. The project uses a flat-layout structure with top-level modules.

**Package Manager:** `uv` (fast Python package manager)

**Python Version:** 3.12+

**Main Dependencies:** `openpyxl` for Excel file generation, `playwright` for browser automation

## Commands

### Development Setup
```bash
# Install dependencies and sync the package
uv sync

# Install Playwright browsers (first time only)
uv run playwright install chromium
```

### Running the Tool
```bash
# Generate template using browser automation (Playwright)
uv run generate-browser-template

# Browser automation with options
uv run generate-browser-template -o custom_output.xlsx --max-matches 5
uv run generate-browser-template --no-headless  # Run with visible browser for debugging
uv run generate-browser-template --enhanced-odds  # Fetch detailed odds data

# The tool fetches live match data from live.500.com and displays it in terminal:
# - 场次 (Match ID): 周一001, 周二002, etc.
# - 赛事 (League): 德甲, 意甲, 非洲杯, etc.
# - 轮次 (Round): 第17轮, 半决赛, etc.
# - 比赛时间 (Match Time): 01-15 01:30
# - 状态 (Status): 未, 进行中, 完场
# - 主队/客队 (Home/Away Team) with rankings: [04]那不勒斯
# - 让球 (Handicap): 半球, 球半, 受半球, etc.
# - 赔率数据 (Odds): 胜负奖金, 让球奖金, 平均欧赔, 威廉, 澳彩, 365, 皇者
```

### Custom Claude Code Commands

```bash
# Git commit helper - interactive commit creation
/commit

# This command provides:
# - Current git status display
# - Unstaged and staged changes preview
# - Recent commit messages for style reference
# - Automatic co-authorship (Co-Authored-By: Claude <noreply@anthropic.com>)
```

**Commit Message Style:**
- Use conventional commit format: `Type: Description`
- Types: Add, Fix, Refactor, Update, Docs, Remove
- Examples:
  - `Add: Asian handicap analysis data fetching from detail pages`
  - `Fix: Handle missing odds data gracefully`
  - `Refactor: Improve error handling in browser automation`

### Building Executable

```bash
# Install PyInstaller (first time only)
uv sync --dev

# Build executable directly
uv run pyinstaller --onefile --name "足球彩票分析工具" --console browser_bet_scraper.py
# Output: dist/足球彩票分析工具 (or 足球彩票分析工具.exe on Windows)
```

### After Adding New Modules
When adding new `.py` files as CLI modules, update `pyproject.toml`:
1. Add the module to `[tool.setuptools].py-modules` list
2. Add a new entry in `[project.scripts]` mapping the command to `module:main`
3. Run `uv sync` to rebuild the package

Example:
```toml
[project.scripts]
new-command = "new_module:main"
generate-browser-template = "browser_bet_scraper:main"

[tool.setuptools]
py-modules = ["new_module", "browser_bet_scraper"]
```

## Architecture

### Module Structure

#### `browser_bet_scraper.py` - Browser Automation Template Generator (Playwright)

**Purpose:** Generates betting analysis templates using browser automation for reliable data extraction from live.500.com.

**Key Features:**
- Uses Playwright browser automation (Chromium) to render JavaScript
- Fetches comprehensive match data from live.500.com Jingcai score table
- Desktop browser mode for full table data access
- Displays formatted match table in terminal with odds information
- Supports headless and visible browser modes
- Optional enhanced odds data fetching from detail pages
- Fallback to sequential match IDs if scraping fails

**Extracted Data Fields:**
- `match_id` - 场次: 周一001, 周二002, etc.
- `league` - 赛事: 德甲, 意甲, 非洲杯, etc.
- `round` - 轮次: 第17轮, 半决赛, etc.
- `match_time` - 比赛时间: 01-15 01:30
- `status` - 状态: 未, 进行中, 完场
- `home_team` / `home_rank` - 主队及排名: [04]那不勒斯
- `handicap` - 让球: 半球, 球半, 受半球, etc.
- `away_team` / `away_rank` - 客队及排名: 帕尔马[14]
- `halftime_score` - 半场比分
- `win_odds` - 胜负奖金
- `let_odds` - 让球奖金
- `avg_euro` - 平均欧赔
- `william_odds` - 威廉赔率
- `aust_odds` - 澳彩赔率
- `bet365_odds` - 365赔率
- `royal_odds` - 皇者赔率

**Functions:**
- `fetch_matches_with_browser(url, headless, timeout)` - Async function to fetch match data using Playwright
- `fetch_enhanced_odds_data(match_ids)` - Async function to fetch detailed odds data for specific matches
- `_generate_fallback_matches(count)` - Generates fallback MatchData objects when scraping fails
- `add_match_data(ws, start_row, match)` - Adds match data to Excel worksheet
- `generate_browser_template(filename, url, headless, max_matches, fetch_enhanced_odds)` - Main entry point with CLI support via argparse

**Template Structure:**
- Two-row header system with merged cells
- 18 columns for detailed odds analysis (亚盘盘口, 百家初凯, 横纵分析, 左右格局警示, 主流凯利, 平局预警, 平赔数据)
- Time-based tracking: Each match has 9 time points (初盘3点, 赛前1, 赛前2, 临场一小时, 临场半小时, 临场15分钟, 临场10分钟, 临场, 完场)
- 10 rows per match (1 match ID row + 9 time point rows)
- Color-coded styling: blue headers, yellow match IDs, gray time labels

### Template Structure

**Header Rows (Row 1-2):**
- Row 1: Main section headers with merged cells (亚盘盘口, 横纵分析, 左右格局警示, 主流凯利, 平局预警, 平赔数据)
- Row 2: Sub-headers (比赛/时间, 初始亚凯对比, 主, 客, 对比, etc.)

**Match Structure:**
Each match consists of 10 rows:
1. Match ID row (e.g., "001", "002") - displays match ID and handicap
2-10. Time point rows: 初盘3点, 赛前1, 赛前2, 临场一小时, 临场半小时, 临场15分钟, 临场10分钟, 临场, 完场

### Styling System

Uses `openpyxl.styles` for consistent formatting:
- **Headers:** Blue background (#366092), white bold text, centered
- **Match ID rows:** Yellow background (#FFF2CC), bold text
- **Time label rows:** Gray background (#E7E6E6)
- **Data cells:** White background, centered, with thin black borders
- **Borders:** Thin black borders on all cells

### Localization

All text is in Chinese. When modifying content:
- Use Chinese for headers, labels, and data
- Time labels: 初盘3点, 赛前1, 赛前2, 临场一小时, 临场半小时, 临场15分钟, 临场10分钟, 临场, 完场

## Key Patterns

### Adding a New CLI Module

1. Create new `.py` file with `main()` function as entry point
2. Follow existing patterns: constants at top, then styling definitions, then functions, then `main()` at bottom
3. Import from `openpyxl.styles`: Font, PatternFill, Border, Side, Alignment
4. For browser automation, import `asyncio` and `from playwright.async_api import async_playwright`
5. Update `pyproject.toml` as shown above
6. Run `uv sync`

### Modifying Column Structures

- Column indices are 1-18 (A-R columns)
- Column widths are defined in `COLUMN_WIDTHS` dictionary mapping column letters to widths
- Merged cells for headers are defined in `MERGED_CELLS` list of tuples

## Constants Reference

**MAIN_HEADERS:** Main section headers for row 1 (18 columns)
**SUB_HEADERS:** Sub-headers for row 2 (18 columns)
**COLUMN_WIDTHS:** Dictionary mapping column letters to width values
**MERGED_CELLS:** List of tuples defining merged cell ranges for main sections
**TIME_LABELS:** List of 9 time tracking labels for each match

**Styling Constants:**
- `HEADER_FONT`, `HEADER_FILL`, `HEADER_ALIGNMENT` - Header row styling
- `BORDER_STYLE` - Thin black border style
- `MATCH_ID_FILL`, `MATCH_ID_FONT` - Match ID row styling
- `TIME_LABEL_FILL`, `TIME_LABEL_FONT` - Time label row styling
- `DATA_ALIGNMENT` - Default cell alignment

## Web Scraping

### Browser Automation Scraping (`browser_bet_scraper.py`)
Uses Playwright for browser automation:
- Desktop viewport (1920x1080) with realistic user agent
- Fetches data from https://live.500.com/ (Jingcai score page)
- Multiple selector strategies for robust data extraction
- Extracts comprehensive match info: teams, leagues, times, handicaps, odds, rankings
- Handles JavaScript-rendered content that curl cannot access
- Displays formatted table output in terminal

**JavaScript Injection:** Data is extracted via `page.evaluate()` executing custom JavaScript in the browser context, allowing access to rendered DOM and dynamic content.

**Terminal Output:** The tool displays a formatted table with all match data including odds information for easy viewing.
