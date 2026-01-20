# tests/test_loader.py
import pytest
import polars as pl
from datetime import date, time
from src.services.excel_loader import carregar_planilha

# Excel falso 
def criar_excel_mock(tmp_path):
    dados = {
        "Tag": ["TAG-01", "TAG-02"],
        "Padrão": ["PREV", "CORR"],
        "Data Início": [date(2026, 1, 20), date(2026, 1, 21)],
        "Hora Início": [time(8, 0), time(9, 30)],
        "Hora Fim": ["18:00", "NOW"],
        "Tipo de Oficina": ["ELETRICA", "MECANICA"],
        "Tipo de Ordem": ["ROTINA", "URGENTE"],
        "Complexidade": ["BAIXA", "ALTA"],
        "Reclamante": ["JOAO", "MARIA"],
        "Tipo de Ocorrência": ["FALHA", "QUEBRA"],
        "Causa da ocorrência": ["USO", "ACIDENTE"],
        "Observações": ["Ok", "Check"],
        "Check Mão de Obra": [True, False],
        "Técnico Responsável": ["TEC1", "TEC2"],
        "Serviço Realizado": ["TROCA", "REPARO"]
    }
    
    df = pl.DataFrame(dados)
    caminho = tmp_path / "teste_os.xlsx"
    df.write_excel(caminho)
    return str(caminho)

def test_leitura_excel_com_now(tmp_path):
    """Testa se o loader lê o Excel e entende o NOW na coluna Hora Fim"""
    arquivo_mock = criar_excel_mock(tmp_path)
    
    # Executa a função real
    lista_os = carregar_planilha(arquivo_mock)
    
    assert len(lista_os) == 2
    
    # Verifica a OS 1 (Normal)
    os1 = lista_os[0]
    assert os1.tag == "TAG-01"
    assert os1.is_closing_now is False
    assert os1.data_fechamento == date(2026, 1, 20)

    # Verifica a OS 2 (NOW)
    os2 = lista_os[1]
    assert os2.tag == "TAG-02"
    assert os2.is_closing_now is True 