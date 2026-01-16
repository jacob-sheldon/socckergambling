"""
Match table widget for displaying scraped match data.
"""

from PyQt6.QtWidgets import QTableView
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QColor
from typing import List, Optional


class MatchTable(QTableView):
    """
    Table view for displaying match data with 18 columns.
    """

    # Column definitions
    COLUMNS = [
        "场次", "赛事", "轮次", "时间", "状态",
        "主队", "排名", "让球", "客队", "排名",
        "半场", "胜负", "让球", "均欧", "威廉", "澳彩", "365", "皇者"
    ]

    def __init__(self, parent=None):
        super().__init__(parent)

        self._init_ui()
        self.matches = []

    def _init_ui(self):
        """Initialize the table view."""
        # Create model
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(self.COLUMNS)
        self.setModel(self.model)

        # Configure table appearance
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.horizontalHeader().setStretchLastSection(True)

        # Set column widths
        self._set_column_widths()

    def _set_column_widths(self):
        """Set appropriate column widths."""
        widths = [80, 80, 100, 120, 60, 150, 50, 80, 150, 50, 80, 80, 80, 80, 80, 80, 80, 80]
        for i, width in enumerate(widths):
            self.setColumnWidth(i, width)

    def clear_matches(self):
        """Clear all matches from the table."""
        self.model.clear()
        self.model.setHorizontalHeaderLabels(self.COLUMNS)
        self.matches = []

    def add_match(self, match):
        """
        Add a MatchData object to the table.

        Args:
            match: MatchData object from browser_bet_scraper
        """
        self.matches.append(match)

        # Create row data
        row_data = [
            match.match_id,
            match.league,
            match.round,
            match.match_time,
            match.status,
            match.home_team,
            match.home_rank,
            match.handicap,
            match.away_team,
            match.away_rank,
            match.halftime_score,
            match.win_odds,
            match.let_odds,
            match.avg_euro,
            match.william_odds,
            match.aust_odds,
            match.bet365_odds,
            match.royal_odds,
        ]

        # Add Asian handicap data if available
        if match.asian_handicap:
            # Replace handicap column with Asian handicap
            row_data[7] = f"{match.handicap} (亚:{match.asian_handicap})"
        if match.home_water:
            row_data[11] = match.home_water  # Replace win_odds with home_water
        if match.away_water:
            row_data[12] = match.away_water  # Replace let_odds with away_water

        # Create items and add to model
        items = []
        for data in row_data:
            item = QStandardItem(str(data) if data else "")
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            items.append(item)

        self.model.appendRow(items)

        # Scroll to bottom
        self.scrollToBottom()

    def get_all_matches(self) -> List:
        """Get all MatchData objects currently displayed."""
        return self.matches

    def get_match_count(self) -> int:
        """Get the number of matches in the table."""
        return self.model.rowCount()

    def refresh_table(self):
        """Refresh the table display with current matches."""
        self.clear_matches()
        for match in self.matches:
            self.add_match(match)
