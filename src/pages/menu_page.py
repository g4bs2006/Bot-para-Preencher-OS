from playwright.async_api import Page, expect
from loguru import logger
from src.core.exceptions import AutomacaoOSError

class MenuPage:
    def __init__(self, page: Page):
        self.page = page
        self.input_busca_equipamento = '//*[@id="side-menu"]/div[2]/nv-atalhos/div/div[2]/form/input'

    async def buscar_ativo(self, tag: str):
        """
        Busca um equipamento pela TAG usando a barra lateral.
        """
        logger.info(f"🔍 Buscando ativo: {tag}...")

        try:
            # 1. Garante que o menu está visível e o input existe
            locator_busca = self.page.locator(self.input_busca_equipamento)
            
            try:
                await expect(locator_busca).to_be_visible(timeout=10000)
            except AssertionError:
                logger.warning("Campo de busca não apareceu em 10s.")
                raise

            # 2. Limpa o campo e preenche
            await locator_busca.fill("") 
            await locator_busca.fill(tag)

            # 3. Pressiona ENTER para iniciar a busca
            await locator_busca.press("Enter")
            
            # 4. Aguarda feedback da aplicação
            try:
                await self.page.wait_for_load_state("networkidle", timeout=5000)
            except Exception:
                pass 

            logger.debug(f"Busca por {tag} disparada com sucesso.")

        except Exception as e:
            logger.error(f"❌ Erro ao buscar a tag {tag} no menu: {e}")
            # Repassa o erro para o controlador principal tomar decisão (abortar OS ou tentar de novo)
            raise e
