from datetime import date, time
from typing import Optional, Union, Any
from pydantic import BaseModel, Field, field_validator, ValidationInfo

class OrdemServico(BaseModel):
    """
    Representa uma Ordem de Serviço com validações e normalização de dados.
    """
    # --- Identificadores ---
    tag: str = Field(..., description="Identificador único do equipamento (TAG)")
    padrao: str = Field(..., description="Padrão do equipamento")
    data_inicio: Any
    hora_inicio: Any
    data_fechamento: Any
    hora_fechamento: Any = None
    tipo_oficina: str
    tipo_ordem: str
    complexidade: str
    reclamante: str
    tipo_ocorrencia: str
    causa_ocorrencia: str
    observacoes: Optional[str] = Field(default="", description="Observações adicionais")
    mao_de_obra_finalizada: bool
    tecnico: str
    servico_executado: str
    # --- Validadores ---

    @field_validator(
        'tag', 'padrao', 'tipo_oficina', 'tipo_ordem', 
        'complexidade', 'reclamante', 'tipo_ocorrencia', 
        'causa_ocorrencia', 'tecnico', 'servico_executado',
        mode='before'
    )
    @classmethod
    def limpar_strings_upper(cls, v: any, info: ValidationInfo) -> any:
        """Remove espaços extras e converte para maiúsculo."""
        if isinstance(v, str):
            return v.strip().upper()
        return v

    @field_validator('data_fechamento', mode='before')
    @classmethod
    def validar_data_fechamento(cls, v: any) -> any:
        """
        Valida o campo data_fechamento.
        Pode ser um objeto date, uma string de data isoformat, ou a string 'NOW'.
        """
        if isinstance(v, str):
            v_upper = v.strip().upper()
            if v_upper == 'NOW':
                return 'NOW'
        return v

    @property
    def is_closing_now(self) -> bool:
        """
        Detecta se a Data de Fechamento é 'NOW'.
        Isso acionará o clique no botão 'Agora' do Neovero.
        """
        if isinstance(self.data_fechamento, str):
            return self.data_fechamento == "NOW"
        return False
