# tests/test_models.py
from datetime import date, time
import pytest
from pydantic import ValidationError
from src.models import OrdemServico

def test_criacao_os_valida():
    """Testa se conseguimos criar uma OS com dados perfeitos e se a limpeza funciona."""
    dados = {
        "tag": "  tag-123  ",  # Testando espaços extras
        "padrao": "preventiva", # Testando minúsculas
        "data_inicio": date(2026, 1, 20),
        "hora_inicio": time(8, 0),
        "data_fechamento": date(2026, 1, 20),
        "hora_fechamento": time(10, 0),
        "tipo_oficina": "clínica",
        "tipo_ordem": "corretiva",
        "complexidade": "baixa",
        "reclamante": "joao silva",
        "tipo_ocorrencia": "falha",
        "causa_ocorrencia": "desgaste",
        "observacoes": "Teste de obs",
        "mao_de_obra_finalizada": True,
        "tecnico": "maria souza",
        "servico_executado": "troca de peça"
    }

    os = OrdemServico(**dados)
    
    assert os.tag == "TAG-123"          # Deve estar maiúsculo e sem espaço
    assert os.padrao == "PREVENTIVA"    # Deve estar maiúsculo
    assert os.tecnico == "MARIA SOUZA"  # Deve estar maiúsculo
    assert os.is_closing_now is False   # Não é NOW

def test_logica_now():
    """Testa se a flag is_closing_now ativa quando passamos 'NOW'."""
    dados = {
        "tag": "TAG-999",
        "padrao": "PREV",
        "data_inicio": date(2026, 1, 20),
        "hora_inicio": time(14, 0),
        "data_fechamento": "NOW",       # O pulo do gato
        "tipo_oficina": "CLINICA",
        "tipo_ordem": "CORRETIVA",
        "complexidade": "MEDIA",
        "reclamante": "USER",
        "tipo_ocorrencia": "ERRO",
        "causa_ocorrencia": "DANO",
        "mao_de_obra_finalizada": False,
        "tecnico": "TEC",
        "servico_executado": "SERV"
    }

    os = OrdemServico(**dados)
    assert os.is_closing_now is True

def test_falha_campo_obrigatorio():
    """Testa se o sistema GRITA quando falta um campo (ex: Tag)."""
    dados_incompletos = {
        "padrao": "PREV",
        # Falta a TAG e outros campos
    }

    # Esperamos que levante um erro de validação
    with pytest.raises(ValidationError):
        OrdemServico(**dados_incompletos)