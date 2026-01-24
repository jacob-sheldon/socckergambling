"""
Match table widget for displaying scraped data in the Excel template layout.
"""

import re
from typing import List

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QStandardItemModel, QStandardItem, QColor
from PyQt6.QtWidgets import QTableView, QAbstractItemView

from browser_bet_scraper import (
    MAIN_HEADERS,
    SUB_HEADERS,
    COLUMN_WIDTHS,
    MERGED_CELLS,
    TIME_LABELS,
)


HEADER_BG = "#366092"
HEADER_FG = "#000000"
MATCH_BG = "#FFF2CC"
TIME_LABEL_BG = "#E7E6E6"
TEXT_FG = "#000000"


class MatchTable(QTableView):
    """
    Table view that mirrors the live_betting_template.xlsx layout.
    """

    def __init__(self, parent=None):
        super().__init__(parent)

        self.column_count = len(SUB_HEADERS)
        self.header_rows = 2
        self.matches: List = []
        self._row_map = {}

        self._init_ui()

    def _init_ui(self):
        """Initialize the table view."""
        self.model = QStandardItemModel(0, self.column_count, self)
        self.setModel(self.model)

        # Configure table appearance
        self.setAlternatingRowColors(False)
        self.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.setEditTriggers(QTableView.EditTrigger.NoEditTriggers)
        self.horizontalHeader().setVisible(False)
        self.verticalHeader().setVisible(False)
        self.horizontalHeader().setStretchLastSection(True)
        self.setWordWrap(False)
        self.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)

        self._init_headers()
        self._set_column_widths()

    def _init_headers(self):
        """Render the two header rows (main + sub headers)."""
        self.model.setRowCount(self.header_rows)
        self.model.setColumnCount(self.column_count)

        for col_idx in range(self.column_count):
            main_header = MAIN_HEADERS[col_idx] or ""
            sub_header = SUB_HEADERS[col_idx] or ""
            self.model.setItem(0, col_idx, self._make_item(main_header, header=True))
            self.model.setItem(1, col_idx, self._make_item(sub_header, header=True))

        self._apply_header_spans()
        self.setRowHeight(0, 28)
        self.setRowHeight(1, 22)

    def _apply_header_spans(self):
        """Merge cells to match the Excel template header layout."""
        self.clearSpans()
        for start_cell, end_cell in MERGED_CELLS:
            start_col = self._excel_cell_to_col(start_cell)
            end_col = self._excel_cell_to_col(end_cell)
            if end_col > start_col:
                self.setSpan(0, start_col, 1, end_col - start_col + 1)

    def _excel_cell_to_col(self, cell_ref: str) -> int:
        """Convert Excel cell ref like 'B1' to zero-based column index."""
        match = re.match(r"([A-Z]+)", cell_ref)
        if not match:
            return 0
        letters = match.group(1)
        col_index = 0
        for ch in letters:
            col_index = col_index * 26 + (ord(ch) - ord("A") + 1)
        return col_index - 1

    def _set_column_widths(self):
        """Set column widths based on the Excel template sizing."""
        font_metrics = self.fontMetrics()
        char_width = max(font_metrics.horizontalAdvance("0"), 8)
        for idx, letter in enumerate("ABCDEFGHIJKLMNOPQR"):
            width = COLUMN_WIDTHS.get(letter, 8)
            self.setColumnWidth(idx, int(width * char_width * 1.4 + 16))

    def clear_matches(self):
        """Clear all matches from the table while preserving headers."""
        self.model.clear()
        self.matches = []
        self._row_map = {}
        self._init_headers()
        self._set_column_widths()

    def add_match(self, match):
        """
        Add a MatchData object to the table.

        Args:
            match: MatchData object from browser_bet_scraper
        """
        if match.match_id in self._row_map:
            start_row = self._row_map[match.match_id]
            for idx, existing in enumerate(self.matches):
                if existing.match_id == match.match_id:
                    self.matches[idx] = match
                    break
            self._update_match_rows(start_row, match)
            return

        self.matches.append(match)

        start_row = self.model.rowCount()
        self.model.insertRows(start_row, 1 + len(TIME_LABELS))
        self._row_map[match.match_id] = start_row

        self._update_match_rows(start_row, match)

    def _update_match_rows(self, start_row: int, match):
        """Update rows for a match starting at the given row index."""
        match_values = [""] * self.column_count
        match_values[0] = match.match_id

        if match.asian_handicap:
            match_values[1] = match.asian_handicap
            match_values[2] = match.home_water or ""
            match_values[3] = match.away_water or ""
        else:
            match_values[1] = match.handicap
            match_values[2] = match.win_odds or ""
            match_values[3] = match.let_odds or ""

        self._set_row_items(start_row, match_values, row_type="match")

        for idx, label in enumerate(TIME_LABELS):
            row_idx = start_row + 1 + idx
            time_values = [""] * self.column_count
            time_values[0] = label

            if idx == 0:
                if match.euro_kelly_win:
                    time_values[2] = match.euro_kelly_win
                if match.euro_kelly_lose:
                    time_values[3] = match.euro_kelly_lose
                if match.euro_kelly_win_2:
                    time_values[5] = match.euro_kelly_win_2
                if match.euro_kelly_lose_2:
                    time_values[6] = match.euro_kelly_lose_2
                if match.euro_kelly_draw:
                    time_values[8] = match.euro_kelly_draw
                if match.euro_kelly_draw_2:
                    time_values[9] = match.euro_kelly_draw_2

            self._set_row_items(row_idx, time_values, row_type="time")

    def _set_row_items(self, row_idx: int, values: List[str], row_type: str):
        """Populate a row with values and apply template styling."""
        for col_idx, value in enumerate(values):
            text = "" if value is None else str(value)
            item = self._make_item(text)

            if row_type == "match":
                self._apply_match_style(item)
            elif row_type == "time":
                if col_idx == 0:
                    self._apply_time_label_style(item)
                else:
                    self._apply_data_style(item)
            else:
                self._apply_data_style(item)

            self.model.setItem(row_idx, col_idx, item)

    def _make_item(self, text: str, header: bool = False) -> QStandardItem:
        """Create a table item with default alignment and optional header styling."""
        item = QStandardItem(text)
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        if text:
            item.setToolTip(text)

        if header:
            font = item.font()
            font.setBold(True)
            item.setFont(font)
            item.setBackground(QColor(HEADER_BG))
            item.setForeground(QColor(HEADER_FG))
        else:
            item.setForeground(QColor(TEXT_FG))

        return item

    def _apply_match_style(self, item: QStandardItem):
        font = item.font()
        font.setBold(True)
        item.setFont(font)
        item.setBackground(QColor(MATCH_BG))
        item.setForeground(QColor(TEXT_FG))

    def _apply_time_label_style(self, item: QStandardItem):
        item.setBackground(QColor(TIME_LABEL_BG))
        item.setForeground(QColor(TEXT_FG))

    def _apply_data_style(self, item: QStandardItem):
        item.setBackground(QColor("#FFFFFF"))
        item.setForeground(QColor(TEXT_FG))

    def get_all_matches(self) -> List:
        """Get all MatchData objects currently displayed."""
        return self.matches

    def get_match_count(self) -> int:
        """Get the number of matches currently displayed."""
        return len(self.matches)

    def refresh_table(self):
        """Refresh the table display with current matches."""
        matches = list(self.matches)
        self.clear_matches()
        for match in matches:
            self.add_match(match)
