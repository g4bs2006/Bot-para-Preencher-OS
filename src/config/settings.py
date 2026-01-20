import os
from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    NEOVERO_URL: str
    NEOVERO_USER: str
    NEOVERO_PASS: str
    HEADLESS: bool = False

    # Paths
    BASE_DIR: str = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    
    @property
    def DATA_DIR(self) -> str:
        return os.path.join(self.BASE_DIR, "data")
    
    @property
    def INPUT_DIR(self) -> str:
        return os.path.join(self.DATA_DIR, "input")

    @property
    def OUTPUT_DIR(self) -> str:
        return os.path.join(self.DATA_DIR, "output")
        
    @property
    def LOGS_DIR(self) -> str:
        return os.path.join(self.DATA_DIR, "logs")

    class Config:
        env_file = ".env"
        env_file_encoding = 'utf-8'

settings = Settings()
