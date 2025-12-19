# config.py
import json
import os
import shutil

SECRETS_DIR = os.environ.get("SECRETS_DIR", "/etc/secrets")
from dataclasses import dataclass
from typing import Any, Dict

CONFIG_LOCAL = "config.local.json"
CONFIG_EXAMPLE = "config.example.json"

@dataclass
class AppConfig:
    caminho_planilha: str
    pasta_base_saida: str
    pasta_download_temp: str
    pasta_imagens_cert: str
    delay_acao: float = 3.5

def _read_json(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def load_config() -> AppConfig:
    # Se estiver rodando em nuvem (ex.: Render), secret files ficam em /etc/secrets
    secret_cfg = os.path.join(SECRETS_DIR, CONFIG_LOCAL)
    if (not os.path.exists(CONFIG_LOCAL)) and os.path.exists(secret_cfg):
        try:
            shutil.copy(secret_cfg, CONFIG_LOCAL)
        except Exception:
            # Se não der para copiar, seguimos lendo diretamente do secret
            pass

    if os.path.exists(CONFIG_LOCAL):
        return AppConfig(**_read_json(CONFIG_LOCAL))

    if os.path.exists(CONFIG_EXAMPLE):
        data = _read_json(CONFIG_EXAMPLE)
    else:
        data = {
            "caminho_planilha": "./planilhas/ACESSO_PORTAL_NACIONAL.xlsx",
            "pasta_base_saida": "./saidas",
            "pasta_download_temp": "./downloads_temp",
            "pasta_imagens_cert": "./imagens",
            "delay_acao": 3.5,
        }

    with open(CONFIG_LOCAL, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    return AppConfig(**data)

def save_config(cfg: AppConfig) -> None:
    data = {
        "caminho_planilha": cfg.caminho_planilha,
        "pasta_base_saida": cfg.pasta_base_saida,
        "pasta_download_temp": cfg.pasta_download_temp,
        "pasta_imagens_cert": cfg.pasta_imagens_cert,
        "delay_acao": float(cfg.delay_acao),
    }
    with open(CONFIG_LOCAL, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
