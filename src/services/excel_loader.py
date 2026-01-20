import polars as pl
from src.models import OrdemServico
from loguru import logger
from typing import List
from datetime import date, time, datetime

def _parse_time(val):
    if val is None:
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    if isinstance(val, str):
        val = val.strip()
        try:
            return datetime.strptime(val, "%H:%M").time()
        except ValueError:
            try:
                return datetime.strptime(val, "%H:%M:%S").time()
            except ValueError:
                pass # Let Pydantic handle or fail
    return val

def _parse_date(val):
    if val is None:
        return None
    if isinstance(val, date):
        return val # Covers date and datetime
    if isinstance(val, datetime):
        return val.date()
    
    # Tratamento para datas numéricas "DDMMYYYY" (ex: 19012026.0)
    if isinstance(val, (int, float)):
        s_val = str(int(val))
        if len(s_val) == 8: # DDMMYYYY
            try:
                # DD = s_val[:2], MM = s_val[2:4], YYYY = s_val[4:]
                return date(int(s_val[4:]), int(s_val[2:4]), int(s_val[:2]))
            except ValueError:
                pass
    
    if isinstance(val, str):
        val = val.strip()
        # Tenta formatos comuns
        formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]
        for fmt in formatos:
            try:
                return datetime.strptime(val, fmt).date()
            except ValueError:
                continue
    return val

def _safe_str(val):
    """Converte None para string vazia"""
    if val is None:
        return ""
    return str(val)

def carregar_planilha(caminho_arquivo: str) -> List[OrdemServico]:
    """
    Lê um arquivo Excel e converte suas linhas em objetos OrdemServico.
    Ignora linhas inválidas, logando warnings.
    """
    logger.info(f"Lendo arquivo: {caminho_arquivo}...")
    
    try:
        # Lê o Excel usando fastexcel diretamente 
        import fastexcel
        excel_reader = fastexcel.read_excel(caminho_arquivo)
        # Carrega a primeira aba (default) e converte para Polars
        df = excel_reader.load_sheet(0).to_polars()
    except Exception as e:
        logger.error(f"Erro crítico ao abrir arquivo: {e}")
        raise

    ordens_validas = []
    
    # Itera sobre cada linha da planilha
    for row in df.iter_rows(named=True):
        try:
            # 1. Tratamento Especial para "NOW" e Datas
            hora_fim_raw = row.get("Hora Fim")
            data_inicio_raw = row.get("Data Início")
            
            # Lógica: Se Hora Fim for "NOW", a Data Fim também vira "NOW" para o modelo
            if isinstance(hora_fim_raw, str) and hora_fim_raw.strip().upper() == "NOW":
                data_fechamento_processada = "NOW"
                hora_fechamento_processada = None
            else:
                # Se não for NOW, assumimos que fecha no mesmo dia que abriu
                data_fechamento_processada = data_inicio_raw
                hora_fechamento_processada = hora_fim_raw

            # 2. Mapeamento Exato (Coluna Excel -> Campo Modelo)
            dados_os = {
                "tag": _safe_str(row.get("Tag")),
                "padrao": _safe_str(row.get("Padrão")),
                
                "data_inicio": data_inicio_raw,
                "hora_inicio": row.get("Hora Início"),
                
                # Campos calculados acima
                "data_fechamento": data_fechamento_processada,
                "hora_fechamento": hora_fechamento_processada,
                
                "tipo_oficina": _safe_str(row.get("Tipo de Oficina")),
                "tipo_ordem": _safe_str(row.get("Tipo de Ordem")),
                "complexidade": _safe_str(row.get("Complexidade")),
                "reclamante": _safe_str(row.get("Reclamante")),
                "tipo_ocorrencia": _safe_str(row.get("Tipo de Ocorrência")),
                "causa_ocorrencia": _safe_str(row.get("Causa da ocorrência")),
                "observacoes": _safe_str(row.get("Observações")),
                
                # Conversão para bool
                "mao_de_obra_finalizada": bool(row.get("Check Mão de Obra")),
                
                "tecnico": _safe_str(row.get("Técnico Responsável")),
                "servico_executado": _safe_str(row.get("Serviço Realizado"))
            }
            
            # 3. Validação Pydantic
            logger.debug(f"Tentando criar OS com: {dados_os}")
            os_obj = OrdemServico(**dados_os)
            ordens_validas.append(os_obj)
            
        except Exception as e:
            tag_erro = row.get("Tag", "LINHA SEM TAG")
            logger.warning(f"Ignorando OS '{tag_erro}' por dados inválidos: {e}")

    logger.success(f"Sucesso! {len(ordens_validas)} ordens prontas para processar.")
    return ordens_validas
