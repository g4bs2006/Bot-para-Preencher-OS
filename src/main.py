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
    logger.info("Iniciando Automação de OS (Estratégia Reload)...")
    
    # 1. Carregar Dados
    input_file = os.path.join(settings.INPUT_DIR, "dados.xlsx")
    if not os.path.exists(input_file):
        logger.error(f"Arquivo não encontrado: {input_file}")
        return

    ordens = carregar_planilha(input_file)
    if not ordens: return

    # 2. Setup Browser
    browser_manager = BrowserManager()
    page = await browser_manager.start_browser(headless=settings.HEADLESS)
    
    # Estatísticas
    stats = {"sucesso": 0, "falha": 0, "pulado": 0}
    
    try:
        # Instancia Páginas
        login_page = LoginPage(page)
        menu_page = MenuPage(page)
        equipment_page = EquipmentPage(page)
        os_page = OsPage(page)
        
        # --- LOGIN ---
        await login_page.navegar()
        await login_page.realizar_login()
        logger.info("Login realizado. Aguardando 3s...")
        await asyncio.sleep(3) 
        
        # 3. Loop Principal
        logger.info(f"Iniciando loop para {len(ordens)} ordens...")
        
        for i, os_data in enumerate(ordens):
            logger.info(f"👉 [{i+1}/{len(ordens)}] Processando TAG: {os_data.tag}...")
            
            try:
                # PASSO 1: Busca Ativo
                await menu_page.buscar_ativo(os_data.tag)
                
                # PASSO 2: Verifica Duplicidade (Só se for Desativação)
                is_desativacao = "DESATIV" in str(os_data.tipo_ordem).upper() or "DESATIV" in str(os_data.tipo_oficina).upper()
                
                if is_desativacao:
                    tem_duplicidade = await equipment_page.verificar_desativacao_existente()
                    if tem_duplicidade:
                        logger.warning(f"⏩ Pulando {os_data.tag}: Ordem de desativação já existente.")
                        stats["pulado"] += 1
                        # O 'reload' no finally vai limpar tudo para o próximo
                        continue
                
                # PASSO 3: Abrir Nova OS
                await equipment_page.clicar_abrir_os()
                
                # Aguarda iframe carregar
                await asyncio.sleep(2)
                
                # PASSO 4: Preencher OS
                # O preenchimento inclui clicar em Salvar e esperar 5s
                await os_page.preencher_nova_os(os_data)
                
                stats["sucesso"] += 1
                logger.success(f"OS {os_data.tag} finalizada com sucesso!")

            except Exception as e_os:
                logger.error(f"Erro ao processar OS {os_data.tag}: {e_os}")
                stats["falha"] += 1
                try:
                    await page.screenshot(path=os.path.join(settings.LOGS_DIR, f"erro_{os_data.tag}.png"))
                except:
                    pass
                
            finally:
                # Reseta o contexto da página para evitar conflitos na próxima iteração
                try:
                    await page.reload()
                    await page.wait_for_load_state("networkidle", timeout=10000)
                except Exception:
                    pass
                
                await asyncio.sleep(1)

        logger.info(f"""
=== RELATÓRIO FINAL ===
✅ Sucesso: {stats['sucesso']}
⚠️ Pulados: {stats['pulado']}
❌ Falhas:  {stats['falha']}
=======================
""")
        logger.success("Processamento finalizar.")

    except Exception as e:
        logger.error(f"Erro fatal na execução: {e}")
        try:
            await page.screenshot(path=os.path.join(settings.LOGS_DIR, "fatal_error.png"))
        except:
            pass
        
    finally:
        await browser_manager.stop_browser()

if __name__ == "__main__":
    logger.add(os.path.join(settings.LOGS_DIR, "execution.log"), rotation="1 MB")
    try:
        asyncio.run(run_automation())
    except KeyboardInterrupt:
        pass