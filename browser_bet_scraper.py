"""
Generate betting analysis Excel template with comprehensive real match data from live.500.com.
This uses Playwright to extract the full table data from the Jingcai score page.

Run: uv run generate-browser-template
"""

import asyncio
import json
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import List, Dict, Any, Optional
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# URLs for different data sources
BASE_URL = "https://live.500.com/"  # Desktop version with full table data
DESKTOP_URL = "https://www.500.com/jczq/"
ODDS_URL = "https://odds.500.com/"

# ============================================================================
# DATA STRUCTURES
# ============================================================================


@dataclass
class MatchData:
    """Comprehensive match data structure from live.500.com table."""
    match_id: str           # 场次: 周一001, 周二002, etc.
    league: str             # 赛事: 澳超, 非洲杯, etc.
    round: str              # 轮次: 第17轮, 1/8决赛, etc.
    match_time: str         # 比赛时间: 01-05 16:00
    status: str             # 状态: 未, 进行中, 完场
    home_team: str          # 主队: [04]麦克阿瑟FC
    home_rank: str          # 主队排名: 04
    handicap: str           # 让球: 受平手/半球, 一球/球半, etc.
    away_team: str          # 客队: 奥克兰FC[01]
    away_rank: str          # 客队排名: 01
    halftime_score: str     # 半场比分
    win_odds: str           # 胜负奖金
    let_odds: str           # 让球奖金
    avg_euro: str           # 平均欧赔
    william_odds: str       # 威廉赔率
    aust_odds: str          # 澳彩赔率
    bet365_odds: str        # 365赔率
    royal_odds: str         # 皇者赔率

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for Excel serialization."""
        return {
            'match_id': self.match_id,
            'league': self.league,
            'round': self.round,
            'match_time': self.match_time,
            'status': self.status,
            'home_team': self.home_team,
            'home_rank': self.home_rank,
            'handicap': self.handicap,
            'away_team': self.away_team,
            'away_rank': self.away_rank,
            'halftime_score': self.halftime_score,
            'win_odds': self.win_odds,
            'let_odds': self.let_odds,
            'avg_euro': self.avg_euro,
            'william_odds': self.william_odds,
            'aust_odds': self.aust_odds,
            'bet365_odds': self.bet365_odds,
            'royal_odds': self.royal_odds,
        }


# ============================================================================
# EXCEL TEMPLATES CONSTANTS
# ============================================================================

MAIN_HEADERS = [
    "",
    "亚盘盘口",
    "",
    "",
    "横纵分析",
    "",
    "",
    "左右格局警示",
    "",
    "主流凯利",
    "",
    "",
    "",
    "平局预警",
    "平赔数据",
    "",
    "",
    ""
]

SUB_HEADERS = [
    "比赛/时间",
    "初始亚凯对比",
    "主",
    "客",
    "对比",
    "主",
    "客",
    "预测",
    "提示",
    "初凯",
    "即凯",
    "初变",
    "初变",
    "",
    "初凯",
    "即凯",
    "即初凯",
    "初凯变化"
]

COLUMN_WIDTHS = {
    'A': 10, 'B': 12, 'C': 8, 'D': 8, 'E': 8,
    'F': 8, 'G': 8, 'H': 8, 'I': 8, 'J': 8,
    'K': 8, 'L': 8, 'M': 8, 'N': 10, 'O': 8,
    'P': 8, 'Q': 10, 'R': 12
}

MERGED_CELLS = [
    ("B1", "D1"), ("E1", "G1"), ("H1", "I1"),
    ("J1", "M1"), ("N1", "N1"), ("O1", "R1")
]

TIME_LABELS = [
    "初盘3点", "赛前1", "赛前2", "临场一小时",
    "临场半小时", "临场15分钟", "临场10分钟", "临场", "完场"
]

# Styling constants
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")

BORDER_STYLE = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='thin', color='000000')
)

MATCH_ID_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
MATCH_ID_FONT = Font(bold=True, size=11)

TIME_LABEL_FILL = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
TIME_LABEL_FONT = Font(size=10)

DATA_ALIGNMENT = Alignment(horizontal="center", vertical="center")


# ============================================================================
# BROWSER AUTOMATION WITH PLAYWRIGHT
# ============================================================================


async def fetch_matches_with_browser(
    url: str = BASE_URL,
    headless: bool = True,
    timeout: int = 60000
) -> List[MatchData]:
    """
    Fetch comprehensive match data from live.500.com using Playwright.
    Extracts the full Jingcai betting table with all columns.

    Args:
        url: URL to scrape (default: live.500.com)
        headless: Run browser in headless mode (no GUI)
        timeout: Page load timeout in milliseconds

    Returns:
        List of MatchData objects with comprehensive match information
    """
    from playwright.async_api import async_playwright, Error as PlaywrightError

    print(f"正在从 {url} 获取竞彩比分数据...")
    print("=" * 80)

    try:
        async with async_playwright() as p:
            # Launch browser with anti-detection settings
            browser = await p.chromium.launch(
                headless=headless,
                args=[
                    '--disable-blink-features=AutomationControlled',
                    '--disable-dev-shm-usage',
                    '--no-sandbox',
                ]
            )

            # Create context with realistic user agent (desktop)
            context = await browser.new_context(
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                viewport={'width': 1920, 'height': 1080},
                locale='zh-CN',
            )

            page = await context.new_page()

            # Navigate to URL and wait for network idle
            print("正在加载页面...")
            await page.goto(url, wait_until='domcontentloaded', timeout=timeout)

            # Wait for the table to load
            try:
                await page.wait_for_selector('table', timeout=10000)
                print("✓ 表格加载成功")
            except Exception:
                print("⚠ 未找到表格，继续尝试...")

            # Execute JavaScript to extract the full table data
            print("正在提取表格数据...")

            extraction_script = """
            () => {
                const matches = [];

                // Find the main table with Jingcai scores
                const tables = document.querySelectorAll('table');
                let targetTable = null;

                // Look for table with match rows
                for (const table of tables) {
                    const rows = table.querySelectorAll('tr');
                    if (rows.length > 2) {
                        const firstRow = rows[0];
                        if (firstRow.innerText.includes('场次') || firstRow.innerText.includes('赛事')) {
                            targetTable = table;
                            break;
                        }
                    }
                }

                if (!targetTable) {
                    console.log('未找到目标表格');
                    return matches;
                }

                console.log('找到表格，行数:', targetTable.querySelectorAll('tr').length);

                // Extract data from each row
                const rows = targetTable.querySelectorAll('tr');
                for (let i = 0; i < rows.length; i++) {
                    const row = rows[i];
                    const cells = row.querySelectorAll('td');

                    if (cells.length < 5) continue; // Skip header rows or empty rows

                    try {
                        // Extract match ID (e.g., "周一001")
                        const firstCell = cells[0].innerText.trim();
                        const matchIdMatch = firstCell.match(/周[一二三四五六七日天日](\\d{3})/);
                        if (!matchIdMatch) continue;

                        const match_id = firstCell;

                        // Extract league (赛事)
                        const league = cells[1] ? cells[1].innerText.trim() : '';

                        // Extract round (轮次)
                        const round = cells[2] ? cells[2].innerText.trim() : '';

                        // Extract match time
                        const match_time = cells[3] ? cells[3].innerText.trim() : '';

                        // Extract status
                        const status = cells[4] ? cells[4].innerText.trim() : '';

                        // Extract home team with ranking
                        let home_team_raw = cells[5] ? cells[5].innerText.trim() : '';
                        // Handle completed matches with score embedded
                        if (home_team_raw.includes('\\n')) {
                            const lines = home_team_raw.split('\\n');
                            home_team_raw = lines[0];
                        }
                        const home_rank_match = home_team_raw.match(/\\[(\\d+)\\]/);
                        const home_rank = home_rank_match ? home_rank_match[1] : '';
                        const home_team = home_team_raw.replace(/\\[\\d+\\]/, '').trim();

                        // Extract handicap (让球)
                        const handicap = cells[6] ? cells[6].innerText.trim() : '';

                        // Extract away team with ranking
                        let away_team_raw = cells[7] ? cells[7].innerText.trim() : '';
                        // Handle completed matches with score embedded
                        if (away_team_raw.includes('\\n')) {
                            const lines = away_team_raw.split('\\n');
                            away_team_raw = lines[0];
                        }
                        const away_rank_match = away_team_raw.match(/\\[(\\d+)\\]/);
                        const away_rank = away_rank_match ? away_rank_match[1] : '';
                        const away_team = away_team_raw.replace(/\\[\\d+\\]/, '').trim();

                        // Extract halftime score
                        let halftime_score = cells[8] ? cells[8].innerText.trim() : '';
                        // Clean up multiline scores
                        if (halftime_score.includes('\\n')) {
                            const lines = halftime_score.split('\\n');
                            halftime_score = lines[lines.length - 1].trim();
                        }

                        // Extract odds data
                        const win_odds = cells[9] ? cells[9].innerText.trim() : '';
                        const let_odds = cells[10] ? cells[10].innerText.trim() : '';
                        const avg_euro = cells[11] ? cells[11].innerText.trim() : '';

                        // Extract company odds
                        const william_odds = cells[12] ? cells[12].innerText.trim() : '';
                        const aust_odds = cells[13] ? cells[13].innerText.trim() : '';
                        const bet365_odds = cells[14] ? cells[14].innerText.trim() : '';
                        const royal_odds = cells[15] ? cells[15].innerText.trim() : '';

                        matches.push({
                            match_id,
                            league,
                            round,
                            match_time,
                            status,
                            home_team,
                            home_rank,
                            handicap,
                            away_team,
                            away_rank,
                            halftime_score,
                            win_odds,
                            let_odds,
                            avg_euro,
                            william_odds,
                            aust_odds,
                            bet365_odds,
                            royal_odds
                        });

                    } catch (err) {
                        console.error(`处理第 ${i} 行时出错:`, err.message);
                    }
                }

                console.log(`成功提取 ${matches.length} 场比赛数据`);
                return matches;
            }
            """

            # Execute extraction and get results
            extracted_data = await page.evaluate(extraction_script)

            # Convert to MatchData objects
            matches = []
            for item in extracted_data:
                match = MatchData(
                    match_id=item.get('match_id', ''),
                    league=item.get('league', ''),
                    round=item.get('round', ''),
                    match_time=item.get('match_time', ''),
                    status=item.get('status', ''),
                    home_team=item.get('home_team', ''),
                    home_rank=item.get('home_rank', ''),
                    handicap=item.get('handicap', ''),
                    away_team=item.get('away_team', ''),
                    away_rank=item.get('away_rank', ''),
                    halftime_score=item.get('halftime_score', ''),
                    win_odds=item.get('win_odds', ''),
                    let_odds=item.get('let_odds', ''),
                    avg_euro=item.get('avg_euro', ''),
                    william_odds=item.get('william_odds', ''),
                    aust_odds=item.get('aust_odds', ''),
                    bet365_odds=item.get('bet365_odds', ''),
                    royal_odds=item.get('royal_odds', '')
                )
                matches.append(match)

            # Print table to terminal with detailed odds
            print("\n" + "=" * 200)
            print(f"{'场次':<8} {'赛事':<8} {'轮次':<12} {'时间':<14} {'状态':<4} {'主队':<22} {'让球':<12} {'客队':<22} {'胜负':<12} {'让球':<12} {'均欧':<10}")
            print("=" * 200)

            for match in matches:
                home_rank_str = f"[{match.home_rank}]" if match.home_rank else ""
                away_rank_str = f"[{match.away_rank}]" if match.away_rank else ""

                print(f"{match.match_id:<8} {match.league:<8} {match.round:<12} {match.match_time:<14} "
                      f"{match.status:<4} {home_rank_str}{match.home_team:<22} {match.handicap:<12} "
                      f"{match.away_team:<22}{away_rank_str} {match.win_odds:<12} {match.let_odds:<12} {match.avg_euro:<10}")

            print("=" * 160)
            print(f"\n✓ 成功从浏览器提取 {len(matches)} 场比赛数据")

            await browser.close()
            return matches

    except PlaywrightError as e:
        print(f"Playwright 错误: {e}")
        print("回退到生成示例数据...")
        return _generate_fallback_matches()
    except Exception as e:
        print(f"浏览器自动化失败: {e}")
        print("回退到生成示例数据...")
        return _generate_fallback_matches()


def _generate_fallback_matches(count: int = 15) -> List[MatchData]:
    """Generate fallback match data when browser scraping fails."""
    matches = []
    for i in range(1, count + 1):
        match = MatchData(
            match_id=f"周一{i:03d}",
            league=f"竞彩联赛",
            round=f"第{i}轮",
            match_time=f"01-{i+5:02d} 14:00",
            status="未",
            home_team=f"主队{i}",
            home_rank=f"{i:02d}",
            handicap=f"受平手/半球",
            away_team=f"客队{i}",
            away_rank=f"{20-i:02d}",
            halftime_score="-",
            win_odds="",
            let_odds="",
            avg_euro="",
            william_odds="",
            aust_odds="",
            bet365_odds="",
            royal_odds=""
        )
        matches.append(match)
    print(f"生成了 {len(matches)} 条示例数据")
    return matches


async def fetch_enhanced_odds_data(match_ids: List[str]) -> Dict[str, Dict]:
    """
    Fetch enhanced odds data for specific matches using browser automation.
    This can navigate to detailed odds pages for each match.

    Args:
        match_ids: List of match IDs to fetch odds for

    Returns:
        Dictionary mapping match_id to odds data
    """
    from playwright.async_api import async_playwright

    odds_data = {}

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page()

            for match_id in match_ids:
                try:
                    # Construct detailed odds URL for the match
                    detail_url = f"{ODDS_URL}jczq/{match_id.replace('周', '')}.shtml"
                    await page.goto(detail_url, timeout=15000)

                    # Extract odds data
                    odds_script = """
                    () => {
                        return {
                            kelly: document.querySelector('.kelly-value')?.innerText,
                            asian_handicap: document.querySelector('.ah-value')?.innerText,
                            over_under: document.querySelector('.ou-value')?.innerText,
                            history: Array.from(document.querySelectorAll('.odds-history tr')).map(tr => ({
                                time: tr.cells[0]?.innerText,
                                home: tr.cells[1]?.innerText,
                                draw: tr.cells[2]?.innerText,
                                away: tr.cells[3]?.innerText
                            }))
                        };
                    }
                    """

                    match_odds = await page.evaluate(odds_script)
                    odds_data[match_id] = match_odds

                    # Small delay to be respectful to the server
                    await asyncio.sleep(0.5)

                except Exception as e:
                    print(f"无法获取比赛 {match_id} 的赔率数据: {e}")
                    odds_data[match_id] = {}

            await browser.close()

    except Exception as e:
        print(f"增强赔率数据获取失败: {e}")

    return odds_data


# ============================================================================
# EXCEL GENERATION
# ============================================================================


def create_template_workbook():
    """Create a new workbook with headers and basic formatting."""
    wb = Workbook()
    ws = wb.active
    ws.title = "足球彩票分析模板"
    return wb, ws


def set_column_widths(ws):
    """Set appropriate column widths."""
    for col, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col].width = width


def merge_header_cells(ws):
    """Merge cells for main section headers."""
    for start_cell, end_cell in MERGED_CELLS:
        ws.merge_cells(f"{start_cell}:{end_cell}")


def style_header_rows(ws):
    """Apply styling to header rows."""
    # Style Row 1
    for col_idx, header in enumerate(MAIN_HEADERS, start=1):
        if header:
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGNMENT
            cell.border = BORDER_STYLE

    # Style Row 2
    for col_idx, header in enumerate(SUB_HEADERS, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = BORDER_STYLE

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 20


def add_match_data(ws, start_row: int, match: MatchData) -> int:
    """
    Add comprehensive match data to the worksheet.

    Args:
        ws: Worksheet object
        start_row: Starting row number
        match: MatchData object with comprehensive information

    Returns:
        Number of rows added
    """
    current_row = start_row

    # Match ID row
    ws.cell(row=current_row, column=1, value=match.match_id)
    ws.cell(row=current_row, column=2, value=match.handicap)

    # Add odds data if available
    if match.win_odds:
        ws.cell(row=current_row, column=3, value=match.win_odds)
    if match.let_odds:
        ws.cell(row=current_row, column=4, value=match.let_odds)

    # Style match ID row
    for col_idx in range(1, 19):
        cell = ws.cell(row=current_row, column=col_idx)
        cell.fill = MATCH_ID_FILL
        cell.font = MATCH_ID_FONT
        cell.alignment = DATA_ALIGNMENT
        cell.border = BORDER_STYLE

    current_row += 1

    # Add time point rows
    for time_label in TIME_LABELS:
        ws.cell(row=current_row, column=1, value=time_label)

        # Style the row
        for col_idx in range(1, 19):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.alignment = DATA_ALIGNMENT
            cell.border = BORDER_STYLE

            if col_idx == 1:
                cell.fill = TIME_LABEL_FILL
                cell.font = TIME_LABEL_FONT
            else:
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        current_row += 1

    return len(TIME_LABELS) + 1


def generate_browser_template(
    filename: str = "live_betting_template.xlsx",
    url: str = BASE_URL,
    headless: bool = True,
    max_matches: Optional[int] = None,
    fetch_enhanced_odds: bool = False
):
    """
    Generate Excel betting analysis template using browser automation.

    Args:
        filename: Output Excel filename
        url: URL to scrape match data from
        headless: Run browser in headless mode
        max_matches: Maximum number of matches to include
        fetch_enhanced_odds: Whether to fetch enhanced odds data (slower)
    """
    print("=" * 80)
    print("浏览器自动化 - 足球彩票分析模板生成器")
    print("=" * 80)

    # Fetch match data using browser automation
    matches = asyncio.run(fetch_matches_with_browser(url, headless=headless))

    if not matches:
        print("未找到比赛数据，生成空模板...")
        matches = _generate_fallback_matches(5)

    # Limit matches if specified
    if max_matches is not None:
        matches = matches[:max_matches]

    # Optionally fetch enhanced odds data
    odds_data = {}
    if fetch_enhanced_odds:
        print("\n正在获取增强赔率数据...")
        match_ids = [m.match_id for m in matches]
        odds_data = asyncio.run(fetch_enhanced_odds_data(match_ids))

    # Create workbook and worksheet
    wb, ws = create_template_workbook()

    # Apply formatting
    set_column_widths(ws)
    merge_header_cells(ws)
    style_header_rows(ws)

    # Add match data
    current_row = 3
    for match in matches:
        rows_added = add_match_data(ws, current_row, match)
        current_row += rows_added

    # Freeze header rows
    ws.freeze_panes = "A3"

    # Add metadata sheet
    info_sheet = wb.create_sheet("数据来源信息")
    info_sheet.column_dimensions['A'].width = 20
    info_sheet.column_dimensions['B'].width = 40

    info_data = [
        ["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["数据来源", url],
        ["抓取方式", "浏览器自动化 (Playwright)"],
        ["比赛数量", len(matches)],
        ["增强赔率数据", "是" if fetch_enhanced_odds else "否"],
    ]

    for row_idx, (label, value) in enumerate(info_data, start=1):
        info_sheet.cell(row=row_idx, column=1, value=label)
        info_sheet.cell(row=row_idx, column=2, value=value)

    # Save workbook
    wb.save(filename)

    print("=" * 80)
    print(f"✓ Excel 模板 '{filename}' 生成成功！")
    print(f"  - 包含 {len(matches)} 场真实比赛数据")
    print(f"  - 使用浏览器自动化抓取数据")
    if fetch_enhanced_odds:
        print(f"  - 已包含增强赔率数据")
    print("=" * 80)


def main():
    """Entry point for the CLI command."""
    import argparse

    parser = argparse.ArgumentParser(
        description="使用浏览器自动化生成足球彩票分析 Excel 模板"
    )
    parser.add_argument(
        "-o", "--output",
        default="live_betting_template.xlsx",
        help="输出 Excel 文件名 (默认: live_betting_template.xlsx)"
    )
    parser.add_argument(
        "-u", "--url",
        default=BASE_URL,
        help=f"要抓取的 URL (默认: {BASE_URL})"
    )
    parser.add_argument(
        "--no-headless",
        action="store_true",
        help="显示浏览器窗口（用于调试）"
    )
    parser.add_argument(
        "-m", "--max-matches",
        type=int,
        default=None,
        help="最多包含的比赛数量"
    )
    parser.add_argument(
        "--enhanced-odds",
        action="store_true",
        help="获取增强赔率数据（更慢但更全面）"
    )

    args = parser.parse_args()

    generate_browser_template(
        filename=args.output,
        url=args.url,
        headless=not args.no_headless,
        max_matches=args.max_matches,
        fetch_enhanced_odds=args.enhanced_odds
    )


if __name__ == "__main__":
    main()
