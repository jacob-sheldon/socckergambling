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
# Generate betting analysis template with live match data from 500.com (curl-based)
uv run generate-live-template

# Generate template using browser automation (Playwright) - more reliable for dynamic content
uv run generate-browser-template

# Browser automation with options
uv run generate-browser-template -o custom_output.xlsx --max-matches 5
uv run generate-browser-template --no-headless  # Run with visible browser for debugging
uv run generate-browser-template --enhanced-odds  # Fetch detailed odds data

# The above commands create 'live_betting_template.xlsx' with real match IDs
```

### Building Executable

**Note:** The PyInstaller spec file (`betting_analysis_template.spec`) currently references a non-existent source file. To build an executable with the actual `live_bet_scraper.py` module:

```bash
# Install PyInstaller (first time only)
uv sync --dev

# Build executable directly
uv run pyinstaller --onefile --name "足球彩票分析工具" --console live_bet_scraper.py
# Output: dist/足球彩票分析工具 (or 足球彩票分析工具.exe on Windows)
```

Or update the spec file to reference `live_bet_scraper.py` instead of `betting_analysis_template.py`.

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
py-modules = ["live_bet_scraper", "new_module", "browser_bet_scraper"]
```

## Architecture

### Module Structure

#### `live_bet_scraper.py` - Live Match Data Template Generator (curl-based)

**Purpose:** Generates a professional betting analysis template with real match IDs scraped from 500.com.

**Key Features:**
- Scrapes live match data from 500.com (mobile version: https://live.m.500.com/home/zq/jczq/cur)
- Uses `curl` via subprocess to avoid SSL library issues
- Extracts match IDs in format like "001", "002" from Chinese sports betting website
- Generates Excel templates with comprehensive odds analysis structure

**Functions:**
- `fetch_jingcai_matches(url, timeout)` - Scrapes match IDs from 500.com using curl
- `_generate_fallback_match_ids(count)` - Generates sequential match IDs as fallback
- `generate_live_template(filename, url, max_matches)` - Main entry point for template generation

#### `browser_bet_scraper.py` - Browser Automation Template Generator (Playwright)

**Purpose:** Generates betting analysis templates using browser automation for more reliable data extraction from dynamic websites.

**Key Features:**
- Uses Playwright browser automation (Chromium) to render JavaScript
- Extracts comprehensive match data: teams, leagues, match times, handicaps, odds
- Anti-detection features: realistic user agent, mobile viewport
- Supports headless and visible browser modes
- Optional enhanced odds data fetching from detail pages
- Fallback to sequential match IDs if scraping fails

**Functions:**
- `fetch_matches_with_browser(url, headless, timeout)` - Async function to fetch match data using Playwright
- `fetch_enhanced_odds_data(match_ids)` - Async function to fetch detailed odds data for specific matches
- `_generate_fallback_matches(count)` - Generates fallback MatchData objects when scraping fails
- `add_match_data(ws, start_row, match)` - Adds match data to Excel worksheet
- `generate_browser_template(filename, url, headless, max_matches, fetch_enhanced_odds)` - Main entry point with CLI support via argparse

**Data Structure:**
- `MatchData` dataclass with fields: match_id, league, home_team, away_team, match_time, handicap, home_odds, draw_odds, away_odds, asian_handicap_home, asian_handicap_away, over_under, kelly_home, kelly_draw, kelly_away

**Template Structure:**
- Two-row header system with merged cells
- 18 columns for detailed odds analysis (亚盘盘口, 百家初凯, 横纵分析, 左右格局警示, 主流凯利, 平局预警, 平赔数据)
- Time-based tracking: Each match has 9 time points (初盘3点, 赛前1, 赛前2, 临场一小时, 临场半小时, 临场15分钟, 临场10分钟, 临场, 完场)
- 10 rows per match (1 match ID row + 9 time point rows)
- Color-coded styling: blue headers, yellow match IDs, gray time labels

**Functions:**
- `fetch_jingcai_matches(url, timeout)` - Scrapes match IDs from 500.com
- `generate_live_template(filename, url, max_matches)` - Main entry point for template generation

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

### curl-based Scraping (`live_bet_scraper.py`)
The module scrapes match data from 500.com using `curl` via subprocess:
- Mobile version URL is simpler to parse than desktop
- Uses regex pattern `周[一二三四五六七日天日]\d{3}` to find match IDs
- Extracts last 3 characters as numeric match ID
- Includes error handling for timeouts, missing curl, and network failures

**Note:** If the website structure changes, the regex pattern in `fetch_jingcai_matches()` may need updating.

### Browser Automation Scraping (`browser_bet_scraper.py`)
Uses Playwright for browser automation:
- Executes JavaScript in browser to render dynamic content
- Mobile viewport (375x812) with realistic user agent
- Multiple selector strategies for robust data extraction
- Extracts comprehensive match info: teams, leagues, times, handicaps, odds
- Handles JavaScript-rendered content that curl cannot access

**JavaScript Injection:** Data is extracted via `page.evaluate()` executing custom JavaScript in the browser context, allowing access to rendered DOM and dynamic content.
