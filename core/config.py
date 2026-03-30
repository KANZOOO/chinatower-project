# core/config.py
from pydantic_settings import BaseSettings
from pathlib import Path

class Settings(BaseSettings):
    # 自定义路径属性（作为模型字段）
    index: Path = Path(__file__).parent.parent
    # 其他配置字段
    log_level: str
    json_logs: bool
    ai_api_key: str
    ai_api_url: str
    alibaba_cloud_access_key_id: str
    alibaba_cloud_access_key_secret: str
    db_host: str
    db_port: int
    db_user: str
    db_password: str
    db_name: str
    chrome_binary_path:str
    chrome_user_data_dir:str
    playwright_user_data_dir:str
    model_config = {
        "env_file": Path(__file__).parent.parent / ".env",
        "env_file_encoding": "utf-8",
    }
    @property
    def db_url(self) -> str:
        return (
            {
            'url':f'mysql+pymysql://{self.db_user}:{self.db_password}@{self.db_host}:{self.db_port}/{self.db_name}?charset=utf8mb4',  # 数据库
            'pool_size': 24,
            'max_overflow': 20,
            'pool_recycle': 3600,
            'isolation_level': "READ UNCOMMITTED",
            'pool_pre_ping': True,
            'echo': False
            }
        )


    @classmethod
    def for_db(cls, db_name: str) -> "Settings":
        """
        创建一个指定数据库名的配置实例。
        其他字段（host/user/password 等）仍从 .env 或环境变量加载。
        """
        return cls(db_name=db_name)

    def resolve_path(self, relative_path: str) -> Path:
        parts = relative_path.split('/')
        return self.index.joinpath(*parts)

# 保留默认全局实例（用于大多数场景）
settings = Settings()