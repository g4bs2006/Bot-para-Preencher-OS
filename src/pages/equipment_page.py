import asyncio
import os
from playwright.async_api import Page, Frame, Locator, expect
from loguru import logger
from src.core.exceptions import AutomacaoOSError
from src.config.settings import settings

class EquipmentPage:
    def __init__(self, page: Page):
        self.page = page
        self.btn_abrir_os = '//*[@id="btnAbrirOS_text"]'
        self.btn_fechar = '//*[@id="btnFechar_text"]'
        # Texto chave para verificação
        self.texto_desativacao = "DESATIVAÇÃO-INTERNA"

    # ... existing methods ...



    async def _encontrar_elemento_em_frames(self, seletor: str, timeout: int = 5000) -> tuple[Frame, Locator] | None:
        """
        Varre o frame principal e TODOS os iframes filhos
        procurando pelo seletor. Retorna o Frame e o Locator se achar.
        """
        # 1. Tenta na página principal primeiro
        locator_main = self.page.locator(seletor)
        try:
            # Verifica se existe E se está visível
            if await locator_main.count() > 0 and await locator_main.first.is_visible():
                return self.page.main_frame, locator_main.first
        except:
            pass

        # 2. Varre todos os iframes carregados
        for frame in self.page.frames:
            try:
                locator = frame.locator(seletor)
                # Verifica se existe e se está visível
                if await locator.count() > 0:
                    return frame, locator.first
            except Exception:
                continue
        
        return None

    async def verificar_desativacao_existente(self) -> bool:
        """
        Verifica se já existe uma ordem de desativação ativa para o equipamento.
        """
        logger.info("Verificando histórico de Ordens...")
        await asyncio.sleep(3) 

        encontrou = False
        
        for frame in self.page.frames:
            try:
                frame_visible_text = await frame.inner_text("body")
                
                if "DESATIVAÇÃO" in frame_visible_text.upper():
                    # Verifica status da linha encontrada
                    locator_gen = frame.locator('//*[contains(translate(text(), "abcdefghijklmnopqrstuvwxyz", "ABCDEFGHIJKLMNOPQRSTUVWXYZ"), "DESATIVAÇÃO")]')
                    
                    count = await locator_gen.count()
                    if count > 0:
                        for i in range(count):
                            el = locator_gen.nth(i)
                            if await el.is_visible():
                                txt = await el.text_content()
                                status_ignorados = ["ENCERRADA", "CANCELADA", "FECHADA", "CONCLUÍDA"]
                                txt_upper = txt.upper()
                                
                                if not any(s in txt_upper for s in status_ignorados):
                                    logger.warning(f"Desativação ativa encontrada: '{txt.strip()}'")
                                    encontrou = True
                                else:
                                    pass
            except Exception:
                pass
            
            if encontrou: 
                break

        return encontrou

    async def clicar_abrir_os(self):
        """
        Localiza e clica no botão 'Abrir OS'.
        """
        logger.info("Tentando abrir Nova OS...")

        resultado = await self._encontrar_elemento_em_frames(self.btn_abrir_os)
        
        if not resultado:
            resultado = await self._encontrar_elemento_em_frames("text=Abrir OS")

        if resultado:
            _, locator = resultado
            await locator.click(force=True)
            logger.info("Botão Abrir OS clicado.")
        else:
            logger.error("Botão Abrir OS não encontrado.")
            raise AutomacaoOSError("Falha ao localizar botão Abrir OS.")
