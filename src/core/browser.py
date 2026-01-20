from playwright.async_api import async_playwright, Browser, Page, Playwright
from loguru import logger
from typing import Optional

class BrowserManager:
    _instance = None
    _playwright: Optional[Playwright] = None
    _browser: Optional[Browser] = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(BrowserManager, cls).__new__(cls)
        return cls._instance

    async def start_browser(self, headless: bool = False) -> Page:
        """Inicia o browser e retorna uma página context."""
        if self._playwright is None:
            self._playwright = await async_playwright().start()
        
        if self._browser is None:
            self._browser = await self._playwright.chromium.launch(headless=headless)
            logger.info(f"Browser iniciado (Headless: {headless})")
            
        context = await self._browser.new_context()
        page = await context.new_page()
        return page

    async def stop_browser(self):
        if self._browser:
            await self._browser.close()
            self._browser = None
        if self._playwright:
            await self._playwright.stop()
            self._playwright = None
            logger.info("Browser finalizado.")
