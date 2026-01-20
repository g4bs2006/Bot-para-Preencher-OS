from playwright.async_api import Page
from loguru import logger
from src.config.settings import settings

class LoginPage:
    def __init__(self, page: Page):
        self.page = page
        self.input_usuario = '//*[@id="login"]'
        self.input_senha = '//*[@id="senha"]'
        self.btn_entrar = '//*[@id="formusuario"]/div[3]'

    async def navegar(self):
        """Acessa a URL inicial"""
        logger.info(f"Acessando: {settings.NEOVERO_URL}")
        await self.page.goto(settings.NEOVERO_URL)

    async def realizar_login(self):
        """Executa o fluxo de login completo"""
        logger.info("Preenchendo credenciais...")
        
        # 1. Espera o campo de usuário aparecer 
        await self.page.wait_for_selector(self.input_usuario, state="visible")
        
        # 2. Preenche Usuário
        await self.page.fill(self.input_usuario, settings.NEOVERO_USER)
        
        # 3. Preenche Senha
        await self.page.fill(self.input_senha, settings.NEOVERO_PASS)
        
        # 4. Clica em Entrar e espera a navegação acontecer
        logger.info("Clicando em Entrar...")
        
        # O Playwright espera a ação de clique disparar uma navegação (carregamento de página)
        # Se o login falhar (usuário inválido), não navega. 
        # Tenta esperar navegação, mas se não navegar (erro login), segue
        try:
            async with self.page.expect_navigation(timeout=10000): # 10s de timeout
                await self.page.click(self.btn_entrar)
            logger.success("Navegação pós-login detectada.")
        except Exception:
            logger.warning("Navegação não detectada ou timeout. Verifique se o login foi bem sucedido.")

