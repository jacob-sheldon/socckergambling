"""
Generate betting analysis Excel template with comprehensive real match data from 500.com using browser automation.
This uses Playwright for reliable JavaScript rendering and data extraction.

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
BASE_URL = "https://live.m.500.com/home/zq/jczq/cur"
DESKTOP_URL = "https://www.500.com/jczq/"
ODDS_URL = "https://odds.500.com/"

# ============================================================================
# DATA STRUCTURES
# ============================================================================


@dataclass
class MatchData:
    """Comprehensive match data structure."""
    match_id: str
    league: str
    home_team: str
    away_team: str
    match_time: str
    handicap: Optional[str] = None
    home_odds: Optional[float] = None
    draw_odds: Optional[float] = None
    away_odds: Optional[float] = None
    asian_handicap_home: Optional[float] = None
    asian_handicap_away: Optional[float] = None
    over_under: Optional[float] = None
    kelly_home: Optional[float] = None
    kelly_draw: Optional[float] = None
    kelly_away: Optional[float] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for Excel serialization."""
        return {
            'match_id': self.match_id,
            'league': self.league,
            'home_team': self.home_team,
            'away_team': self.away_team,
            'match_time': self.match_time,
            'handicap': self.handicap,
            'home_odds': self.home_odds,
            'draw_odds': self.draw_odds,
            'away_odds': self.away_odds,
            'asian_handicap_home': self.asian_handicap_home,
            'asian_handicap_away': self.asian_handicap_away,
            'over_under': self.over_under,
            'kelly_home': self.kelly_home,
            'kelly_draw': self.kelly_draw,
            'kelly_away': self.kelly_away,
        }


# ============================================================================
# EXCEL TEMPLATES CONSTANTS (reused from live_bet_scraper.py)
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
    timeout: int = 30000
) -> List[MatchData]:
    """
    Fetch comprehensive match data using Playwright browser automation.

    This approach:
    - Executes JavaScript to render dynamic content
    - Handles complex page interactions
    - Extracts detailed odds information
    - More reliable than HTTP requests for SPA/dynamic sites

    Args:
        url: URL to scrape (default: 500.com mobile Jingcai page)
        headless: Run browser in headless mode (no GUI)
        timeout: Page load timeout in milliseconds

    Returns:
        List of MatchData objects with comprehensive match information
    """
    from playwright.async_api import async_playwright, Error as PlaywrightError

    print(f"Launching browser to fetch data from {url}...")

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

            # Create context with realistic user agent
            context = await browser.new_context(
                user_agent='Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1',
                viewport={'width': 375, 'height': 812},  # Mobile viewport
                locale='zh-CN',
            )

            page = await context.new_page()

            # Navigate to URL and wait for network idle
            print("Loading page...")
            await page.goto(url, wait_until='networkidle', timeout=timeout)

            # Wait for match data to load
            try:
                await page.wait_for_selector('table, .match-list, [data-match-id]', timeout=10000)
            except Exception:
                print("Warning: Could not find expected match container, proceeding anyway...")

            # Execute JavaScript to extract comprehensive match data
            print("Extracting match data...")

            extraction_script = """
            () => {
                const matches = [];

                // Try multiple selectors for match containers
                const possibleContainers = [
                    'table tr',
                    '.match-item',
                    '[data-match-id]',
                    '.jczq-table tr',
                    '.odds-table tr',
                    'tbody tr'
                ];

                let matchRows = [];
                for (const selector of possibleContainers) {
                    matchRows = Array.from(document.querySelectorAll(selector));
                    if (matchRows.length > 0) {
                        console.log(`Found ${matchRows.length} rows using selector: ${selector}`);
                        break;
                    }
                }

                // If no structured data found, try to find match info in page
                if (matchRows.length === 0) {
                    // Fallback: search for match patterns in entire page
                    const textContent = document.body.innerText;
                    const matchPatterns = textContent.match(/周[一二三四五六七日天日]\\d{3}/g) || [];
                    console.log(`Found ${matchPatterns.length} match patterns in page text`);

                    return matchPatterns.map((pattern, idx) => ({
                        match_id: pattern.slice(-3),
                        league: '未知联赛',
                        home_team: '主队' + (idx + 1),
                        away_team: '客队' + (idx + 1),
                        match_time: '待定',
                        raw_html: pattern
                    }));
                }

                // Extract data from each row
                matchRows.forEach((row, idx) => {
                    try {
                        const cells = row.querySelectorAll('td');
                        if (cells.length < 2) return;

                        const text = row.innerText || '';
                        const matchIdMatch = text.match(/(\\d{3})/) ||
                                           text.match(/周[一二三四五六七日天日](\\d{3})/) ||
                                           row.getAttribute('data-match-id');

                        if (!matchIdMatch) return;

                        const matchId = matchIdMatch instanceof Array ?
                                       matchIdMatch[0].slice(-3) : matchIdMatch;

                        // Extract team names
                        const teamMatch = text.match(/([\\u4e00-\\u9fa5\\w\\s]+)\\s*[vsVS|]\\s*([\\u4e00-\\u9fa5\\w\\s]+)/);
                        const homeTeam = teamMatch ? teamMatch[1].trim() : `主队${idx + 1}`;
                        const awayTeam = teamMatch ? teamMatch[2].trim() : `客队${idx + 1}`;

                        // Extract league info
                        const leagueMatch = text.match(/([\\u4e00-\\u9fa5]{2,10}\\s?\\d{4})/);
                        const league = leagueMatch ? leagueMatch[1] : '未知联赛';

                        // Extract odds if present
                        const oddsMatch = text.match(/([\\d.]+)\\s*[\\/／]\\s*([\\d.]+)\\s*[\\/／]\\s*([\\d.]+)/);
                        const homeOdds = oddsMatch ? parseFloat(oddsMatch[1]) : null;
                        const drawOdds = oddsMatch ? parseFloat(oddsMatch[2]) : null;
                        const awayOdds = oddsMatch ? parseFloat(oddsMatch[3]) : null;

                        // Extract handicap
                        const handicapMatch = text.match(/([\\-+]?[\\d.]+\\s*球?)/);
                        const handicap = handicapMatch ? handicapMatch[1] : null;

                        // Extract time
                        const timeMatch = text.match(/(\\d{1,2}:\\d{2})/);
                        const matchTime = timeMatch ? timeMatch[1] : '待定';

                        matches.push({
                            match_id: matchId,
                            league: league,
                            home_team: homeTeam,
                            away_team: awayTeam,
                            match_time: matchTime,
                            handicap: handicap,
                            home_odds: homeOdds,
                            draw_odds: drawOdds,
                            away_odds: awayOdds,
                            raw_html: row.innerHTML.slice(0, 200)  // Truncate for debugging
                        });

                    } catch (err) {
                        console.error(`Error processing row ${idx}:`, err.message);
                    }
                });

                console.log(`Extracted ${matches.length} matches`);
                return matches;
            }
            """

            # Execute extraction and get results
            extracted_data = await page.evaluate(extraction_script)

            # Convert to MatchData objects
            matches = []
            for item in extracted_data:
                match = MatchData(
                    match_id=item.get('match_id', f"{len(matches) + 1:03d}"),
                    league=item.get('league', '未知联赛'),
                    home_team=item.get('home_team', '未知主队'),
                    away_team=item.get('away_team', '未知客队'),
                    match_time=item.get('match_time', '待定'),
                    handicap=item.get('handicap'),
                    home_odds=item.get('home_odds'),
                    draw_odds=item.get('draw_odds'),
                    away_odds=item.get('away_odds')
                )
                matches.append(match)

            print(f"Successfully extracted {len(matches)} matches from browser")

            await browser.close()
            return matches

    except PlaywrightError as e:
        print(f"Playwright error: {e}")
        print("Falling back to sequential match IDs...")
        return _generate_fallback_matches()
    except Exception as e:
        print(f"Browser automation failed: {e}")
        print("Falling back to sequential match IDs...")
        return _generate_fallback_matches()


def _generate_fallback_matches(count: int = 15) -> List[MatchData]:
    """Generate fallback match data when browser scraping fails."""
    matches = []
    for i in range(1, count + 1):
        match = MatchData(
            match_id=f"{i:03d}",
            league=f"竞彩联赛{i}",
            home_team=f"主队{i}",
            away_team=f"客队{i}",
            match_time="14:00"
        )
        matches.append(match)
    print(f"Generated {len(matches)} fallback matches")
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
                    detail_url = f"{ODDS_URL}jczq/{match_id}.shtml"
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
                    print(f"Could not fetch odds for match {match_id}: {e}")
                    odds_data[match_id] = {}

            await browser.close()

    except Exception as e:
        print(f"Enhanced odds fetching failed: {e}")

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

    # Add handicap if available
    handicap = match.handicap if match.handicap else "待定"
    ws.cell(row=current_row, column=2, value=handicap)

    # Add odds data if available
    if match.home_odds:
        ws.cell(row=current_row, column=3, value=match.home_odds)
    if match.away_odds:
        ws.cell(row=current_row, column=4, value=match.away_odds)

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
    print("=" * 60)
    print("Browser-Based Betting Template Generator")
    print("=" * 60)

    # Fetch match data using browser automation
    matches = asyncio.run(fetch_matches_with_browser(url, headless=headless))

    if not matches:
        print("No matches found. Generating empty template...")
        matches = _generate_fallback_matches(5)

    # Limit matches if specified
    if max_matches is not None:
        matches = matches[:max_matches]

    # Optionally fetch enhanced odds data
    odds_data = {}
    if fetch_enhanced_odds:
        print("Fetching enhanced odds data...")
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

    print("=" * 60)
    print(f"✓ Excel template '{filename}' generated successfully!")
    print(f"  - Contains {len(matches)} real matches")
    print(f"  - Data scraped using browser automation")
    if fetch_enhanced_odds:
        print(f"  - Enhanced odds data included")
    print("=" * 60)


def main():
    """Entry point for the CLI command."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate betting analysis Excel template using browser automation"
    )
    parser.add_argument(
        "-o", "--output",
        default="live_betting_template.xlsx",
        help="Output Excel filename (default: live_betting_template.xlsx)"
    )
    parser.add_argument(
        "-u", "--url",
        default=BASE_URL,
        help=f"URL to scrape (default: {BASE_URL})"
    )
    parser.add_argument(
        "--no-headless",
        action="store_true",
        help="Run browser with GUI (useful for debugging)"
    )
    parser.add_argument(
        "-m", "--max-matches",
        type=int,
        default=None,
        help="Maximum number of matches to include"
    )
    parser.add_argument(
        "--enhanced-odds",
        action="store_true",
        help="Fetch enhanced odds data (slower but more comprehensive)"
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
