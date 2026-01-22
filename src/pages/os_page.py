import asyncio
import os
from playwright.async_api import Page, Frame, expect, TimeoutError as PlaywrightTimeoutError
from loguru import logger
from src.core.exceptions import AutomacaoOSError
from src.config.settings import settings
from src.models import OrdemServico

class OsPage:
    def __init__(self, page: Page):
        self.page = page
        
        # --- SELETORES (Mapeados) ---
        self.input_data_inicio = '//*[@id="txtdataabertura"]'
        self.input_hora_inicio = '//*[@id="txthoraabertura"]'
        self.btn_fechar_agora = '//*[@id="btnDataFechamentoHoje"]' 
        self.btn_salvar = '//*[@id="btnsalvar_container"]'
        self.btn_fechar_janela_os = 'xpath=/html/body/nv-root/nv-desktop/div/div[2]/nv-window[2]/div/div[1]/div[1]/div[3]/a[4]'

        # Dropdowns
        self.select_oficina = '//*[@id="cboOficina"]'
        self.select_tipo_ordem = '//*[@id="cbotipomanutencao"]'
        self.select_complexidade = '//*[@id="cbocomplexidadeos"]'
        self.select_reclamante = '//*[@id="cboUsuario"]'
        self.select_tipo_ocorrencia = '//*[@id="cboOcorrencia"]'
        self.select_causa_ocorrencia = '//*[@id="cboCausa"]'
        self.select_tecnico = '//*[@id="cbofuncionario"]'
        self.select_servico = '//*[@id="ddlservico"]'
        self.check_mao_obra = '//*[@id="chkOcorrenciaResolvidaMaoDeObra"]'
        
        # Campos de Texto
        self.input_observacoes = '//*[@id="txtObservacaoOcorrencia"]'


    async def _encontrar_frame_ativo(self):
        """
        Encontra o frame onde o formulário de OS está carregado.
        Varre todos os frames buscando um campo chave (Data Abertura).
        """
        # Tenta na main primeiro
        try:
            if await self.page.locator(self.input_data_inicio).count() > 0:
                 return self.page
        except:
            pass

        for frame in self.page.frames:
            try:
                if await frame.locator(self.input_data_inicio).count() > 0:
                    logger.debug(f"Formulário OS encontrado no frame: {frame.name or frame.url}")
                    return frame
            except:
                continue
        return self.page

    async def preencher_dropdown_inteligente(self, frame, seletor: str, texto_excel: str):
        """
        Busca e seleciona a opção no Dropdown que contém o texto extraído do Excel.
        """
        if not texto_excel: 
            return

        try:
            # 1. Espera o select aparecer
            await frame.wait_for_selector(seletor, state="visible", timeout=5000)
            locator_select = frame.locator(seletor)
            
            # 2. Pega todas as opções (Labels) disponíveis
            opcoes = await locator_select.evaluate("el => Array.from(el.options).map(o => o.text)")
            
            # 3. Busca parcial (Case Insensitive)
            valor_para_selecionar = None
            texto_excel_clean = str(texto_excel).strip().upper()
            
            for opcao in opcoes:
                if texto_excel_clean in opcao.strip().upper():
                    valor_para_selecionar = opcao
                    break
            
            # 4. Seleciona
            if valor_para_selecionar:
                await locator_select.select_option(label=valor_para_selecionar)
                logger.debug(f"Dropdown {seletor}: '{texto_excel}' -> '{valor_para_selecionar}'")
            else:
                logger.warning(f"⚠️ Opção '{texto_excel}' não encontrada em {seletor}. Tentando valor original.")
                # Tenta selecionar pelo valor original como fallback
                try:
                    await locator_select.select_option(label=texto_excel)
                except:
                    pass

        except Exception as e:
            logger.error(f"Erro no dropdown {seletor} (Valor: {texto_excel}): {e}")

    async def _aguardar_fechamento_modal(self, timeout: int = 15000) -> bool:
        """
        Aguarda a modal/iframe da OS desaparecer após o salvamento.
        Retorna True se fechou, False se timeout.
        """
        logger.info("⏳ Aguardando fechamento automático da janela OS...")
        
        try:
            # Estratégia 1: Espera o formulário desaparecer
            await self.page.wait_for_selector(
                self.input_data_inicio, 
                state="hidden", 
                timeout=timeout
            )
            logger.success("✅ Janela OS fechada automaticamente!")
            return True
            
        except PlaywrightTimeoutError:
            logger.warning("⚠️ Janela não fechou automaticamente no tempo esperado.")
            return False

    async def _fechar_modal_forcado(self) -> bool:
        """
        Tenta fechar a modal/janela de OS manualmente clicando no botão X.
        Retorna True se conseguiu fechar, False caso contrário.
        """
        logger.info("🔧 Tentando fechar janela manualmente...")
        
        # Lista de possíveis seletores para o botão fechar
        seletores_fechar = [
            self.btn_fechar_janela_os,
            "//a[contains(@class, 'fechar') or contains(@class, 'close')]",
            "//button[contains(text(), 'Fechar')]",
            "//a[@title='Fechar']",
            "//*[contains(@class, 'nv-window')]//a[contains(@class, 'close')]"
        ]
        
        for seletor in seletores_fechar:
            try:
                # Procura em todos os frames
                for frame in [self.page] + self.page.frames:
                    locator = frame.locator(seletor)
                    if await locator.count() > 0:
                        elemento = locator.first
                        if await elemento.is_visible():
                            await elemento.click()
                            logger.info(f"✅ Clicado botão fechar: {seletor}")
                            await asyncio.sleep(1)
                            
                            # Verifica se realmente fechou
                            if await self.page.locator(self.input_data_inicio).count() == 0:
                                logger.success("✅ Janela fechada com sucesso!")
                                return True
            except Exception as e:
                logger.debug(f"Seletor {seletor} falhou: {e}")
                continue
        
        logger.error("❌ Não conseguiu fechar a janela manualmente!")
        return False

    async def _garantir_estado_limpo(self):
        """
        Garante que não há modais/iframes flutuantes antes de prosseguir.
        """
        logger.info("🧹 Verificando estado limpo...")
        
        # Verifica se o formulário ainda está visível
        formulario_visivel = await self.page.locator(self.input_data_inicio).count() > 0
        
        if formulario_visivel:
            logger.warning("⚠️ Formulário ainda detectado! Limpando estado...")
            
            # Tenta fechar
            fechou = await self._fechar_modal_forcado()
            
            if not fechou:
                # Último recurso: fecha todas as janelas nv-window exceto a principal
                logger.warning("⚠️ Usando método de emergência: fechando janelas flutuantes...")
                try:
                    await self.page.evaluate("""
                        () => {
                            const windows = document.querySelectorAll('nv-window');
                            windows.forEach((win, idx) => {
                                // Mantém apenas a primeira janela (menu principal)
                                if (idx > 0) {
                                    win.remove();
                                }
                            });
                        }
                    """)
                    await asyncio.sleep(1)
                    logger.info("✅ Janelas flutuantes removidas via JavaScript")
                except Exception as e:
                    logger.error(f"❌ Falha ao remover janelas: {e}")
        else:
            logger.success("✅ Estado limpo confirmado!")

    async def _clicar_area_neutra(self):
        """
        Clica em uma área neutra (1,1) da página para remover qualquer foco residual.
        Evita que popups/assinaturas apareçam por elementos ainda focados.
        """
        logger.info("🎯 Clicando em área neutra (1,1) para sanitização de foco...")
        try:
            await self.page.mouse.click(1, 1)
            await asyncio.sleep(0.5)
            logger.success("✅ Foco sanitizado com sucesso")
        except Exception as e:
            logger.warning(f"⚠️ Erro ao clicar em área neutra: {e}")

    async def _fechar_janela_os_manualmente(self):
        """
        Fecha explicitamente a janela/modal de OS clicando no botão 'X'.
        Estratégia: Tenta seletor mapeado primeiro, depois genéricos.
        """
        logger.info("🔧 Fechando janela de OS manualmente (clique no botão X)...")
        
        janela_fechada = False
        
        # ESTRATÉGIA 1: Seletor específico mapeado
        seletores_fechar = [
            self.btn_fechar_janela_os,  # Seletor principal mapeado
            # Fallbacks genéricos caso o principal não funcione
            '//a[contains(@class, "close") or contains(@class, "fechar")]',
            '//button[contains(@class, "close")]',
            '//*[contains(@class, "nv-window")]//a[contains(@class, "close")]',
            '//a[@title="Fechar"]',
            '//*[@id="btnFechar"]',
        ]
        
        # Tenta em todos os frames (página + iframes)
        frames_para_verificar = [self.page] + self.page.frames
        
        for seletor in seletores_fechar:
            if janela_fechada:
                break
                
            for frame in frames_para_verificar:
                try:
                    locator = frame.locator(seletor)
                    count = await locator.count()
                    
                    if count > 0:
                        # Procura por elemento visível
                        for i in range(count):
                            elemento = locator.nth(i)
                            
                            try:
                                if await elemento.is_visible(timeout=1000):
                                    logger.info(f"✅ Botão fechar encontrado: {seletor}")
                                    await elemento.click()
                                    janela_fechada = True
                                    logger.success("✅ Janela de OS fechada manualmente")
                                    return
                            except:
                                continue
                except:
                    continue
        
        if not janela_fechada:
            logger.warning("⚠️ Botão fechar não encontrado! Tentando fallback JavaScript...")
            # Último recurso: remove janelas via JavaScript
            try:
                await self.page.evaluate("""
                    () => {
                        const windows = document.querySelectorAll('nv-window');
                        windows.forEach((win, idx) => {
                            if (idx > 0) win.remove();
                        });
                    }
                """)
                logger.info("✅ Janela removida via JavaScript")
            except Exception as e:
                logger.error(f"❌ Falha ao fechar janela: {e}")
                raise AutomacaoOSError("Não foi possível fechar a janela de OS")

    async def preencher_nova_os(self, os_data: OrdemServico):
        """Executa o preenchimento completo da OS com sequência rigorosa de encerramento."""
        logger.info(f"📝 Preenchendo OS: {os_data.tag} | Padrão: {os_data.padrao}")
        
        # 1. Localizar Frame
        frame = await self._encontrar_frame_ativo()

        try:
            # Garante que o form carregou
            await frame.wait_for_selector(self.input_data_inicio, timeout=10000)

            # 2. Datas e Horas (Tratando campos Any/Raw)
            data_val = os_data.data_inicio.strftime("%d/%m/%Y") if hasattr(os_data.data_inicio, 'strftime') else str(os_data.data_inicio)
            hora_val = os_data.hora_inicio.strftime("%H:%M") if hasattr(os_data.hora_inicio, 'strftime') else str(os_data.hora_inicio)

            await frame.fill(self.input_data_inicio, data_val)
            await frame.fill(self.input_hora_inicio, hora_val)
            await frame.press(self.input_hora_inicio, "Tab")

            # 3. Dropdowns
            await self.preencher_dropdown_inteligente(frame, self.select_oficina, os_data.tipo_oficina)
            await self.preencher_dropdown_inteligente(frame, self.select_tipo_ordem, os_data.tipo_ordem)
            await self.preencher_dropdown_inteligente(frame, self.select_complexidade, os_data.complexidade)
            await self.preencher_dropdown_inteligente(frame, self.select_reclamante, os_data.reclamante)
            await self.preencher_dropdown_inteligente(frame, self.select_tipo_ocorrencia, os_data.tipo_ocorrencia)
            await self.preencher_dropdown_inteligente(frame, self.select_causa_ocorrencia, os_data.causa_ocorrencia)

            # 4. Observação
            if os_data.observacoes:
                await frame.fill(self.input_observacoes, str(os_data.observacoes))

            # 5. Fechamento "NOW"
            is_now = False
            if isinstance(os_data.data_fechamento, str) and os_data.data_fechamento.upper() == 'NOW':
                is_now = True
            
            if is_now:
                logger.info("Botão 'Agora' acionado (Fechamento Imediato).")
                await frame.click(self.btn_fechar_agora)
                await asyncio.sleep(1)

            # 6. Mão de Obra / Técnico / Serviço 
            if os_data.mao_de_obra_finalizada:
                is_checked = await frame.locator(self.check_mao_obra).is_checked()
                if not is_checked:
                    logger.info("Marcando 'Mão de Obra Resolvida'...")
                    await frame.click(self.check_mao_obra)
            
            if os_data.tecnico:
                await self.preencher_dropdown_inteligente(frame, self.select_tecnico, os_data.tecnico)
                
            if os_data.servico_executado:
                await self.preencher_dropdown_inteligente(frame, self.select_servico, os_data.servico_executado)

            # === Screenshot antes de salvar ===
            logger.info("📸 Capturando screenshot antes do salvamento...")
            await self.page.screenshot(path=os.path.join(settings.LOGS_DIR, f"debug_antes_salvar_{os_data.tag}.png"))

            # ═══════════════════════════════════════════════════════════════════
            # SEQUÊNCIA RIGOROSA DE ENCERRAMENTO (STRICT SEQUENCE)
            # ═══════════════════════════════════════════════════════════════════
            
            logger.info("=" * 80)
            logger.info("🔄 INICIANDO SEQUÊNCIA RIGOROSA DE ENCERRAMENTO")
            logger.info("=" * 80)
            
            # === AÇÃO 1: SALVAR ===
            logger.info("💾 [1/5] AÇÃO 1: Salvando OS...")
            
            btn_salvar_id = '//*[@id="btnsalvar"]'
            
            # Wait explícito
            await frame.wait_for_selector(btn_salvar_id, state='visible', timeout=10000)
            btn_salvar = frame.locator(btn_salvar_id)
            
            # Validação
            is_enabled = await btn_salvar.is_enabled()
            if not is_enabled:
                logger.warning("⚠️ Botão salvar está desabilitado!")
                raise AutomacaoOSError("Botão salvar desabilitado")
            
            # Clique
            await btn_salvar.click()
            logger.success("✅ Botão Salvar clicado!")
            
            # === WAIT 1: PROCESSAMENTO ===
            logger.info("⏳ [2/5] WAIT: Aguardando processamento do sistema (5 segundos)...")
            await asyncio.sleep(5)
            logger.success("✅ Processamento concluído")
            
            # === AÇÃO 2: FECHAR JANELA ===
            logger.info("🔧 [3/5] AÇÃO 2: Fechando janela de OS manualmente...")
            await self._fechar_janela_os_manualmente()
            
            # === WAIT 2: ESTABILIZAÇÃO ===
            logger.info("⏳ [4/5] WAIT: Aguardando estabilização da animação (3 segundos)...")
            await asyncio.sleep(3)
            logger.success("✅ Janela estabilizada")
            
            # === AÇÃO 3: SANITIZAÇÃO DE FOCO ===
            logger.info("🎯 [5/5] AÇÃO 3: Sanitizando foco (clique em área neutra)...")
            await self._clicar_area_neutra()
            
            logger.info("=" * 80)
            logger.success("✅ SEQUÊNCIA DE ENCERRAMENTO CONCLUÍDA COM SUCESSO")
            logger.info("=" * 80)

            # === Screenshot depois de salvar ===
            logger.info("📸 Capturando screenshot após salvamento...")
            await self.page.screenshot(path=os.path.join(settings.LOGS_DIR, f"debug_depois_salvar_{os_data.tag}.png"))

            logger.success(f"✅ OS {os_data.tag} processada e finalizada!")

        except Exception as e:
            logger.error(f"❌ Erro no preenchimento da OS: {e}")
            screenshot_path = os.path.join(settings.LOGS_DIR, f"erro_preenchimento_{os_data.tag}.png")
            await self.page.screenshot(path=screenshot_path)
            raise AutomacaoOSError(f"Erro ao preencher formulário: {e}")
