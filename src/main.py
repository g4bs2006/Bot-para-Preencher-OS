import asyncio
import sys
import os
from loguru import logger

# Garante que o diretório raiz esteja no path
sys.path.append(os.getcwd())

from src.config.settings import settings
from src.core.browser import BrowserManager
from src.pages.login_page import LoginPage
from src.pages.menu_page import MenuPage
from src.pages.equipment_page import EquipmentPage
from src.pages.os_page import OsPage
from src.services.excel_loader import carregar_planilha

async def run_automation():
    logger.info("=" * 80)
    logger.info("🚀 Iniciando Automação de OS - Estratégia State-Clean (Sem Reload)")
    logger.info("=" * 80)
    
    # 1. Carregar Dados do Excel
    input_file = os.path.join(settings.INPUT_DIR, "dados.xlsx")
    if not os.path.exists(input_file):
        logger.error(f"❌ Arquivo não encontrado: {input_file}")
        return

    ordens = carregar_planilha(input_file)
    if not ordens:
        logger.error("❌ Nenhuma ordem carregada da planilha!")
        return

    logger.info(f"📊 Total de {len(ordens)} ordem(ns) carregada(s) da planilha")

    # 2. Setup Browser
    browser_manager = BrowserManager()
    page = await browser_manager.start_browser()
    
    # Injeta script para prevenir roubo de foco
    await page.add_init_script("window.focus = function() { return false; }")
    logger.info("🔒 Script anti-foco injetado no navegador")
    
    # Estatísticas de execução
    stats = {"sucesso": 0, "falha": 0, "pulado": 0}
    
    try:
        # Instancia Páginas
        login_page = LoginPage(page)
        menu_page = MenuPage(page)
        equipment_page = EquipmentPage(page)
        os_page = OsPage(page)
        
        # === LOGIN ===
        logger.info("🔐 Iniciando processo de login...")
        await login_page.navegar()
        await login_page.realizar_login()
        logger.success("✅ Login realizado com sucesso")
        await asyncio.sleep(3)
        
        # === LOOP PRINCIPAL ===
        logger.info(f"\n{'=' * 80}")
        logger.info(f"🔄 Iniciando processamento de {len(ordens)} ordem(ns)")
        logger.info(f"{'=' * 80}\n")
        
        for i, os_data in enumerate(ordens):
            num_ordem = i + 1
            logger.info(f"\n{'─' * 80}")
            logger.info(f"📌 ORDEM {num_ordem}/{len(ordens)} | TAG: {os_data.tag}")
            logger.info(f"{'─' * 80}")
            
            try:
                # ═══════════════════════════════════════════════════════════════
                # MOMENTO 1: LIMPEZA PRÉVIA (Início de cada iteração)
                # Remove resquícios da OS anterior antes de buscar novo ativo
                # ═══════════════════════════════════════════════════════════════
                logger.info("🧹 [MOMENTO 1] Limpeza prévia: removendo resquícios da iteração anterior...")
                await equipment_page.fechar_janela()
                await asyncio.sleep(1)
                
                # === PASSO 1: BUSCAR ATIVO ===
                logger.info(f"🔍 Buscando ativo com TAG: {os_data.tag}")
                await menu_page.buscar_ativo(os_data.tag)
                await asyncio.sleep(2)  # Aguarda sistema processar busca
                
                # === PASSO 2: VERIFICAÇÃO DE DUPLICIDADE (Apenas para Desativações) ===
                is_desativacao = (
                    "DESATIV" in str(os_data.tipo_ordem).upper() or 
                    "DESATIV" in str(os_data.tipo_oficina).upper()
                )
                
                if is_desativacao:
                    logger.info("🔎 Tipo identificado como DESATIVAÇÃO. Verificando duplicidade...")
                    tem_duplicidade = await equipment_page.verificar_desativacao_existente()
                    
                    if tem_duplicidade:
                        # ═══════════════════════════════════════════════════════════════
                        # MOMENTO 2: LIMPEZA AO PULAR (Condicional de duplicidade)
                        # Fecha janela de equipamento ao detectar duplicidade
                        # ═══════════════════════════════════════════════════════════════
                        logger.warning(f"⏭️ PULANDO ordem {os_data.tag}: Desativação ativa já existente!")
                        stats["pulado"] += 1
                        
                        logger.info("🧹 [MOMENTO 2] Fechando janela de equipamento (duplicidade)...")
                        await equipment_page.fechar_janela()
                        await asyncio.sleep(1)
                        
                        logger.info(f"📊 Status atual: ✅ {stats['sucesso']} | ⏭️ {stats['pulado']} | ❌ {stats['falha']}")
                        continue  # Pula para próxima ordem
                else:
                    logger.debug("ℹ️ Não é desativação. Pulando verificação de duplicidade.")
                
                # === PASSO 3: ABRIR NOVA OS ===
                logger.info("🆕 Abrindo formulário de Nova OS...")
                await equipment_page.clicar_abrir_os()
                await asyncio.sleep(2)  # Aguarda iframe/modal carregar
                
                # === PASSO 4: PREENCHER E SALVAR OS ===
                logger.info("📝 Preenchendo formulário da OS...")
                await os_page.preencher_nova_os(os_data)
                
                stats["sucesso"] += 1
                logger.success(f"✅ OS {os_data.tag} processada com sucesso!")
                logger.info(f"📊 Status atual: ✅ {stats['sucesso']} | ⏭️ {stats['pulado']} | ❌ {stats['falha']}")

            except Exception as e_os:
                # ═══════════════════════════════════════════════════════════════
                # MOMENTO 3: LIMPEZA DE ERRO (Bloco except)
                # Garante que falhas não deixem janelas órfãs
                # ═══════════════════════════════════════════════════════════════
                stats["falha"] += 1
                logger.error(f"❌ ERRO ao processar OS {os_data.tag}: {e_os}")
                
                # Screenshot de debug
                try:
                    screenshot_path = os.path.join(settings.LOGS_DIR, f"erro_{os_data.tag}.png")
                    await page.screenshot(path=screenshot_path)
                    logger.info(f"📸 Screenshot salvo: {screenshot_path}")
                except Exception as e_screenshot:
                    logger.debug(f"Não foi possível capturar screenshot: {e_screenshot}")
                
                # LIMPEZA DE EMERGÊNCIA
                logger.warning("🧹 [MOMENTO 3] Limpeza de emergência após erro...")
                try:
                    await equipment_page.fechar_janela()
                    await asyncio.sleep(2)  # Pausa maior para estabilização após erro
                except Exception as e_cleanup:
                    logger.error(f"❌ Falha na limpeza de emergência: {e_cleanup}")
                    
                    # Último recurso: força limpeza via JavaScript direto
                    try:
                        logger.warning("⚠️ Executando limpeza JavaScript direta (último recurso)...")
                        await page.evaluate("""
                            () => {
                                const windows = document.querySelectorAll('nv-window');
                                windows.forEach((win, idx) => {
                                    if (idx > 0) win.remove();
                                });
                            }
                        """)
                        await asyncio.sleep(1)
                        logger.info("✅ Limpeza JavaScript concluída")
                    except Exception as e_js:
                        logger.error(f"❌ Falha crítica na limpeza JavaScript: {e_js}")
                
                logger.info(f"📊 Status atual: ✅ {stats['sucesso']} | ⏭️ {stats['pulado']} | ❌ {stats['falha']}")
            
            # Pequena pausa entre iterações para estabilidade do sistema
            await asyncio.sleep(0.5)

        # === RELATÓRIO FINAL ===
        logger.info(f"\n{'=' * 80}")
        logger.info("📋 RELATÓRIO FINAL DE EXECUÇÃO")
        logger.info(f"{'=' * 80}")
        logger.success(f"✅ Ordens Processadas com Sucesso: {stats['sucesso']}")
        logger.warning(f"⏭️ Ordens Puladas (Duplicidade):  {stats['pulado']}")
        logger.error(f"❌ Ordens com Falha:               {stats['falha']}")
        logger.info(f"📊 Total Processado:                {stats['sucesso'] + stats['pulado'] + stats['falha']}/{len(ordens)}")
        logger.info(f"{'=' * 80}")
        
        if stats['falha'] == 0:
            logger.success("🎉 Automação concluída SEM FALHAS!")
        else:
            logger.warning(f"⚠️ Automação concluída com {stats['falha']} falha(s). Verifique os logs.")

    except Exception as e_fatal:
        logger.critical(f"💥 ERRO FATAL na execução: {e_fatal}")
        
        try:
            fatal_screenshot = os.path.join(settings.LOGS_DIR, "fatal_error.png")
            await page.screenshot(path=fatal_screenshot)
            logger.info(f"📸 Screenshot de erro fatal salvo: {fatal_screenshot}")
        except:
            pass
        
        raise  # Re-lança exceção para debugging
        
    finally:
        logger.info("\n🔌 Encerrando navegador...")
        await browser_manager.stop_browser()
        logger.info("✅ Navegador encerrado com sucesso")

if __name__ == "__main__":
    # Configura logger com rotação de arquivos
    logger.add(
        os.path.join(settings.LOGS_DIR, "execution.log"),
        rotation="1 MB",
        retention="7 days",
        level="DEBUG"
    )
    
    try:
        asyncio.run(run_automation())
    except KeyboardInterrupt:
        logger.warning("\n⚠️ Execução interrompida pelo usuário (Ctrl+C)")
        sys.exit(0)
    except Exception as e:
        logger.critical(f"💥 Erro não tratado: {e}")
        sys.exit(1)