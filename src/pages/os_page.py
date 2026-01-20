import asyncio
import os
from playwright.async_api import Page, Frame, expect
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
        Encontra o frame onde o forumalário de OS está carregado.
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

    async def preencher_nova_os(self, os_data: OrdemServico):
        """Executa o preenchimento completo da OS."""
        logger.info(f"📝 Preenchendo OS: {os_data.tag} | Padrão: {os_data.padrao}")
        
        # 1. Localizar Frame
        frame = await self._encontrar_frame_ativo()

        try:
            # Garante que o form carregou
            await frame.wait_for_selector(self.input_data_inicio, timeout=10000)

            # 2. Datas e Horas (Tratando campos Any/Raw)
            # Se vier objeto date/datetime, formata. Se vier string, manda direto.
            data_val = os_data.data_inicio.strftime("%d/%m/%Y") if hasattr(os_data.data_inicio, 'strftime') else str(os_data.data_inicio)
            hora_val = os_data.hora_inicio.strftime("%H:%M") if hasattr(os_data.hora_inicio, 'strftime') else str(os_data.hora_inicio)

            await frame.fill(self.input_data_inicio, data_val)
            await frame.fill(self.input_hora_inicio, hora_val)
            await frame.press(self.input_hora_inicio, "Tab") # Trigger de validação

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
            # Precisamos tratar o campo 'data_fechamento' que agora é Any
            is_now = False
            if isinstance(os_data.data_fechamento, str) and os_data.data_fechamento.upper() == 'NOW':
                is_now = True
            
            if is_now:
                logger.info("Botão 'Agora' acionado (Fechamento Imediato).")
                await frame.click(self.btn_fechar_agora)
                await asyncio.sleep(1) # Espera sistema preencher Data Fim

            # 6. Mão de Obra / Técnico / Serviço 
            
            # 6.1 Checkbox "Mão de Obra Finalizada"
            if os_data.mao_de_obra_finalizada:
                is_checked = await frame.locator(self.check_mao_obra).is_checked()
                if not is_checked:
                    logger.info("Marcando 'Mão de Obra Resolvida'...")
                    await frame.click(self.check_mao_obra)
            
            # 6.2 Técnico (Dropdown)
            if os_data.tecnico:
                await self.preencher_dropdown_inteligente(frame, self.select_tecnico, os_data.tecnico)
                
            # 6.3 Serviço (Dropdown)
            if os_data.servico_executado:
                await self.preencher_dropdown_inteligente(frame, self.select_servico, os_data.servico_executado)

            # 7. Salvar
            if await btn_real.count() > 0 and await btn_real.first.is_visible():
                await btn_real.first.click()
            else:
                await btn_container.click()
            
            logger.info("Botão Salvar clicado. Aguardando processamento...")
            await asyncio.sleep(5) 

        except Exception as e:
            logger.error(f"Erro no preenchimento da OS: {e}")
            screenshot_path = os.path.join(settings.LOGS_DIR, f"erro_preenchimento_{os_data.tag}.png")
            await self.page.screenshot(path=screenshot_path)
            raise AutomacaoOSError(f"Erro ao preencher formulário: {e}")

