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
        self.texto_desativacao = "DESATIVAÇÃO-INTERNA"

    async def _encontrar_elemento_em_frames(self, seletor: str, timeout: int = 5000) -> tuple[Frame, Locator] | None:
        """
        Varre o frame principal e TODOS os iframes filhos
        procurando pelo seletor. Retorna o Frame e o Locator se achar.
        """
        # 1. Tenta na página principal primeiro
        locator_main = self.page.locator(seletor)
        try:
            if await locator_main.count() > 0 and await locator_main.first.is_visible():
                return self.page.main_frame, locator_main.first
        except:
            pass

        # 2. Varre todos os iframes carregados
        for frame in self.page.frames:
            try:
                locator = frame.locator(seletor)
                if await locator.count() > 0:
                    return frame, locator.first
            except Exception:
                continue
        
        return None

    async def verificar_desativacao_existente(self) -> bool:
        """
        Verifica se existe QUALQUER registro de desativação no histórico do equipamento.
        REGRA DE NEGÓCIO ABSOLUTA: Não distingue status (Aberta/Fechada).
        Se encontrar "DESATIV", considera duplicidade imediatamente.
        """
        logger.info("🔍 Verificando histórico de Ordens (regra absoluta: qualquer DESATIVAÇÃO = duplicidade)...")
        await asyncio.sleep(3)
        
        total_linhas_analisadas = 0
        total_frames_verificados = 0

        # Varre TODOS os frames (página principal + iframes)
        frames_para_verificar = [self.page] + self.page.frames
        
        for frame_idx, frame in enumerate(frames_para_verificar):
            total_frames_verificados += 1
            
            try:
                # SELETORES GENÉRICOS: Busca TODAS as linhas de tabela
                seletores_tabela = [
                    'table tr',
                    'tr',
                    'tbody tr',
                ]
                
                for seletor in seletores_tabela:
                    try:
                        locator_linhas = frame.locator(seletor)
                        count = await locator_linhas.count()
                        
                        if count == 0:
                            continue
                        
                        logger.debug(f"📋 Frame {frame_idx}: {count} linha(s) com seletor '{seletor}'")
                        
                        # ITERAÇÃO COM LOG DETALHADO
                        for i in range(count):
                            total_linhas_analisadas += 1
                            linha = locator_linhas.nth(i)
                            
                            try:
                                # Verifica visibilidade
                                try:
                                    is_visible = await linha.is_visible(timeout=500)
                                    if not is_visible:
                                        continue
                                except:
                                    continue
                                
                                # Extrai texto cru
                                try:
                                    texto_linha = await linha.inner_text(timeout=500)
                                except:
                                    continue
                                
                                # LOG DEBUG obrigatório
                                logger.debug(f"   Linha {i+1}: '{texto_linha.strip()}'")
                                
                                # REGRA ABSOLUTA: Converte para maiúsculo e verifica "DESATIV"
                                texto_linha_upper = texto_linha.upper()
                                
                                if "DESATIV" in texto_linha_upper:
                                    # DUPLICIDADE DETECTADA - RETORNA IMEDIATAMENTE
                                    logger.warning("⚠️ WARNING: Histórico de Desativação encontrado!")
                                    logger.warning(f"   Texto: '{texto_linha.strip()}'")
                                    logger.warning(f"   Frame: {frame.name or frame.url[:100]}")
                                    logger.warning("❌ DUPLICIDADE DETECTADA (regra absoluta)")
                                    return True
                            
                            except Exception as e_linha:
                                # Tolerância a falhas
                                logger.debug(f"   ⚠️ Erro ao processar linha {i+1}: {e_linha} - continuando...")
                                continue
                        
                    except Exception as e_seletor:
                        logger.debug(f"Erro com seletor '{seletor}': {e_seletor}")
                        continue
                    
            except Exception as e_frame:
                logger.debug(f"⚠️ Erro ao processar frame {frame_idx}: {e_frame}")
                continue

        # Se chegou aqui, não encontrou nenhuma desativação
        logger.info(f"📊 Varredura completa: {total_linhas_analisadas} linha(s) analisadas em {total_frames_verificados} frame(s)")
        logger.success("✅ Nenhum registro de desativação encontrado. Pode prosseguir.")
        return False

    async def clicar_abrir_os(self):
        """
        Localiza e clica no botão 'Abrir OS'.
        Implementa verificação prévia para evitar cliques duplos.
        Verifica o carregamento do formulário em qualquer frame (até 10s).
        """
        logger.info("🔧 Tentando abrir Nova OS...")
        
        input_data_abertura = '//*[@id="txtdataabertura"]'
        
        # VERIFICAÇÃO ANTI-DUPLO CLIQUE
        # Verifica se já não existe uma janela de OS aberta (busca em frames)
        resultado_existente = await self._encontrar_elemento_em_frames(input_data_abertura, timeout=1000)
        if resultado_existente:
            logger.warning("⚠️ Janela de OS já está aberta! Pulando clique...")
            return
        
        # LOCALIZA E CLICA NO BOTÃO
        resultado = await self._encontrar_elemento_em_frames(self.btn_abrir_os)
        
        if not resultado:
            resultado = await self._encontrar_elemento_em_frames("text=Abrir OS")

        if resultado:
            _, locator = resultado
            
            # Verifica se o botão está habilitado
            is_enabled = await locator.is_enabled()
            if not is_enabled:
                logger.warning("⚠️ Botão 'Abrir OS' está desabilitado!")
                raise AutomacaoOSError("Botão Abrir OS desabilitado")
            
            await locator.click()
            logger.info("✅ Botão Abrir OS clicado.")
            
            # === VERIFICAÇÃO ROBUSTA: Aguarda formulário aparecer em qualquer frame ===
            logger.info("⏳ Aguardando janela de OS carregar (timeout: 10s)...")
            
            janela_carregada = False
            tempo_inicio = asyncio.get_event_loop().time()
            timeout_segundos = 10
            
            # Loop de retentativa com verificação em frames
            while (asyncio.get_event_loop().time() - tempo_inicio) < timeout_segundos:
                # Busca o elemento em todos os frames usando o helper
                resultado_formulario = await self._encontrar_elemento_em_frames(
                    input_data_abertura,
                    timeout=500  # 500ms por tentativa
                )
                
                if resultado_formulario:
                    frame_encontrado, _ = resultado_formulario
                    logger.success(f"✅ Janela de OS aberta com sucesso! (Frame: {frame_encontrado.name or 'main'})")
                    janela_carregada = True
                    break
                
                # Pequena pausa antes de tentar novamente
                await asyncio.sleep(0.5)
            
            # Valida se conseguiu carregar
            if not janela_carregada:
                logger.error("❌ Janela de OS não abriu após 10s de espera!")
                raise AutomacaoOSError("Timeout: Janela de OS não carregou em nenhum frame")
                
        else:
            logger.error("❌ Botão Abrir OS não encontrado.")
            raise AutomacaoOSError("Falha ao localizar botão Abrir OS.")

    async def fechar_janela(self):
        """
        Fecha explicitamente janelas/modais abertas.
        CORREÇÃO CRÍTICA: Prioriza interação nativa (botões Fechar/Cancelar) e só usa 
        JavaScript como último recurso (fallback).
        """
        logger.info("🧹 Executando limpeza de janelas abertas...")
        
        janela_fechada = False
        
        # === ESTRATÉGIA 1: BUSCAR E CLICAR EM BOTÕES NATIVOS ===
        # Lista de seletores conhecidos para botões de fechar (ordem de prioridade)
        seletores_fechar_nativos = [
            self.btn_fechar,  # Botão Fechar mapeado
            '//*[@id="btnCancelar_text"]',  # Botão Cancelar
            '//*[@id="btnCancelar"]',
            '//button[contains(text(), "Fechar")]',
            '//button[contains(text(), "Cancelar")]',
            '//a[contains(@class, "close") or contains(@class, "fechar")]',
            '//a[@title="Fechar"]',
            '//*[contains(@class, "nv-window")]//a[contains(@class, "close")]',
            '//*[contains(@class, "btn-close")]'
        ]
        
        logger.debug("🔍 Tentando localizar botões nativos de fechar...")
        
        for seletor in seletores_fechar_nativos:
            try:
                # Procura em todos os frames (página principal + iframes)
                frames_para_verificar = [self.page] + self.page.frames
                
                for frame in frames_para_verificar:
                    try:
                        locator = frame.locator(seletor)
                        count = await locator.count()
                        
                        if count > 0:
                            # Tenta clicar em cada ocorrência visível
                            for i in range(count):
                                elemento = locator.nth(i)
                                
                                try:
                                    if await elemento.is_visible():
                                        logger.info(f"✅ Clicando em botão nativo: {seletor} (ocorrência {i+1})")
                                        await elemento.click()
                                        await asyncio.sleep(1)
                                        janela_fechada = True
                                        
                                        # Verifica se realmente fechou
                                        # (checa se o formulário de OS desapareceu)
                                        if await self.page.locator('//*[@id="txtdataabertura"]').count() == 0:
                                            logger.success("✅ Janela fechada com sucesso via botão nativo!")
                                            return
                                except:
                                    continue
                    except:
                        continue
                        
            except Exception as e:
                logger.debug(f"Seletor {seletor} não encontrado: {e}")
                continue
        
        # Se chegou aqui e janela_fechada é True mas ainda detecta modal, continua
        if janela_fechada:
            logger.info("⚠️ Botão foi clicado mas modal ainda pode estar presente. Verificando...")
            await asyncio.sleep(1)
        
        # === ESTRATÉGIA 2: VERIFICAR SE AINDA HÁ JANELAS ABERTAS ===
        janelas_ainda_abertas = False
        try:
            # Verifica se ainda há janelas nv-window além da principal
            num_janelas = await self.page.evaluate("""
                () => document.querySelectorAll('nv-window').length
            """)
            
            if num_janelas > 1:
                janelas_ainda_abertas = True
                logger.warning(f"⚠️ Ainda há {num_janelas} janelas nv-window abertas!")
        except:
            pass
        
        # === ESTRATÉGIA 3: JAVASCRIPT FALLBACK (ÚLTIMO RECURSO) ===
        if not janela_fechada or janelas_ainda_abertas:
            logger.warning("⚠️ Botões nativos não funcionaram ou janelas ainda abertas. Usando fallback JavaScript...")
            
            try:
                # Remove todas as janelas nv-window exceto a primeira (menu principal)
                resultado = await self.page.evaluate("""
                    () => {
                        const windows = document.querySelectorAll('nv-window');
                        let removidas = 0;
                        
                        windows.forEach((win, idx) => {
                            // Mantém apenas a primeira janela (índice 0 = menu principal)
                            if (idx > 0) {
                                win.remove();
                                removidas++;
                            }
                        });
                        
                        return removidas;
                    }
                """)
                
                if resultado > 0:
                    logger.success(f"✅ {resultado} janela(s) removida(s) via JavaScript (fallback)")
                    await asyncio.sleep(1)
                else:
                    logger.info("ℹ️ Nenhuma janela adicional detectada para remover")
                    
            except Exception as e_js:
                logger.error(f"❌ Erro ao executar fallback JavaScript: {e_js}")
        
        # === VALIDAÇÃO FINAL ===
        try:
            num_janelas_final = await self.page.evaluate("""
                () => document.querySelectorAll('nv-window').length
            """)
            
            if num_janelas_final <= 1:
                logger.success(f"✅ Estado limpo confirmado ({num_janelas_final} janela(s) restante(s))")
            else:
                logger.warning(f"⚠️ Ainda há {num_janelas_final} janelas abertas após limpeza!")
        except:
            pass
