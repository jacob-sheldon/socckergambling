"""
GUI Workers Package

Contains background worker threads for async operations.
"""

from .scraping_worker import ScrapingWorker

__all__ = ["ScrapingWorker"]
