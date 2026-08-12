"""Utilitários compartilhados pela aplicação desktop de escalas.

O módulo não depende de Tkinter nem do Outlook para permitir testes em ambientes
sem interface gráfica ou sem Microsoft Office instalado.
"""

from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Callable


DEFAULT_CONFIG = {
    "outputs": {
        "etapa_1": "Check_2405_Gerado.xlsx",
        "etapa_2": "Check_Pre_Envio_Gerado.xlsx",
        "html_dir": "escalas_geradas_html",
        "history": "historico_execucoes.jsonl",
        "log": "gerador_escalas.log",
        "contacts": "contatos_nova_versao.xlsx",
    },
    "defaults": {
        "report_2405": "Check pre envio-Macro Colunas - 15052026.xlsm",
        "report_2468": "(2468) Esporte - Atividades de Equipe – Sub-Atividades_v2_ (9).xlsx",
        "contacts": "contatos_nova_versao.xlsx",
    },
}


def load_app_config(base_dir: str | None = None) -> dict[str, Any]:
    """Carrega a configuração externa, preservando valores padrão ausentes."""

    root = Path(base_dir or os.path.dirname(__file__))
    config_path = root / "app_config.json"
    config = json.loads(json.dumps(DEFAULT_CONFIG))
    try:
        with config_path.open("r", encoding="utf-8") as config_file:
            loaded = json.load(config_file)
        for section, values in loaded.items():
            if isinstance(values, dict) and isinstance(config.get(section), dict):
                config[section].update(values)
    except (OSError, json.JSONDecodeError):
        pass
    return config


class OutputFileLockedError(PermissionError):
    """Indica que o arquivo de saída e uma alternativa também estão bloqueados."""

    def __init__(self, original_path: str, fallback_path: str | None = None) -> None:
        self.original_path = original_path
        self.fallback_path = fallback_path
        message = (
            f"O arquivo de saída está em uso: {original_path}. "
            "Feche-o no Excel e tente novamente."
        )
        if fallback_path:
            message += f" Também não foi possível gravar a alternativa: {fallback_path}."
        super().__init__(message)


def fallback_path(path: str, suffix: str = "_novo") -> str:
    """Cria um caminho alternativo preservando extensão e diretório."""

    source = Path(path)
    return str(source.with_name(f"{source.stem}{suffix}{source.suffix}"))


def save_with_fallback(writer: Callable[[str], Any], path: str) -> tuple[str, bool]:
    """Executa uma gravação e tenta um nome alternativo se o arquivo estiver bloqueado.

    Retorna ``(caminho_utilizado, usou_fallback)``. Erros diferentes de
    ``PermissionError`` são propagados para não mascarar falhas de dados ou de
    estrutura da planilha.
    """

    try:
        writer(path)
        return path, False
    except PermissionError as original_error:
        alternate = fallback_path(path)
        try:
            writer(alternate)
            return alternate, True
        except PermissionError as fallback_error:
            raise OutputFileLockedError(path, alternate) from fallback_error
        except Exception:
            raise original_error


def append_execution_history(
    base_dir: str,
    etapa: str,
    status: str,
    resumo: str,
    detalhes: str = "",
) -> str | None:
    """Acrescenta uma execução em formato JSON Lines e retorna o caminho salvo."""

    config = load_app_config()
    history_name = config.get("outputs", {}).get("history", "historico_execucoes.jsonl")
    history_path = os.path.join(base_dir, history_name)
    entry = {
        "data_hora": datetime.now().astimezone().isoformat(timespec="seconds"),
        "etapa": etapa,
        "status": status,
        "resumo": resumo,
        "detalhes": detalhes,
    }
    try:
        os.makedirs(base_dir, exist_ok=True)
        with open(history_path, "a", encoding="utf-8") as history_file:
            history_file.write(json.dumps(entry, ensure_ascii=False) + "\n")
        return history_path
    except OSError:
        # O histórico não deve impedir a execução principal, por exemplo, se a
        # pasta do programa estiver somente para leitura.
        return None


def safe_filename(value: Any, default: str = "sem_nome") -> str:
    """Normaliza um valor para uso seguro como nome de arquivo."""

    text = str(value or "").strip()
    text = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text)
    text = text.strip(" ._")
    return text or default


def is_valid_email(value: str) -> bool:
    """Validação simples para impedir criação de rascunhos sem destinatário."""

    text = str(value or "").strip()
    return bool(re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", text))


__all__ = [
    "OutputFileLockedError",
    "append_execution_history",
    "fallback_path",
    "is_valid_email",
    "load_app_config",
    "safe_filename",
    "save_with_fallback",
]
