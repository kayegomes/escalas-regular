import unicodedata

import numpy as np
import pandas as pd


def _norm_text(value):
    text = str(value if value is not None else "").strip().upper()
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _get_first(row, keys, default=""):
    for key in keys:
        if key in row.index:
            value = row.get(key)
            if pd.notna(value):
                return value
    return default


def _row_text(row):
    parts = [
        _get_first(row, ["Tipo de Atividade ", "Tipo de Atividade", "Tipo Atividade "]),
        _get_first(row, ["Row Display"]),
        _get_first(row, ["Sub-Atividade", "Sub-Atividade (shift)"]),
        _get_first(row, ["Descrição", "Evento/Programa", "Atividade/DescriÇõÇœo", "Evento"]),
        _get_first(row, ["Produto (WO/Quick Hold)", "Produto (WO/Shift)"]),
        _get_first(row, ["Quick Hold Job Info"]),
        _get_first(row, ["Quick Hold Job Details"]),
        _get_first(row, ["Event Group"]),
    ]
    return " | ".join(_norm_text(part) for part in parts)


def _is_ge_tv_row(row):
    """Identifica linhas do canal GE TV, cujo envio é feito separadamente."""
    for col in ["Canal", "Plataforma", "Canal (Master Room)"]:
        if col in row.index:
            value = _norm_text(row.get(col))
            if value.replace(" ", "").replace("-", "") == "GETV":
                return True
    return False


def _is_folga_row(row):
    combined = _row_text(row)
    if "VIAGEM" in combined:
        return False
    return any(
        token in combined
        for token in ["FOLGA", "DAY OFF", "VACATION", "FERIAS", "FERI", "COMP DAY"]
    )


def _is_quickhold_in_scale(row):
    tipo = _norm_text(_get_first(row, ["Tipo de Atividade ", "Tipo de Atividade", "Tipo Atividade "]))
    if "QUICK HOLD" not in tipo:
        return False

    combined = _row_text(row)
    if "VIAGEM" in combined:
        return False
    if "AUSENCIA MEDICA" in combined or ("AUSENCIA" in combined and "MEDICA" in combined):
        return False

    return (
        "PODCAST" in combined
        or "CABINE DO JOGO" in combined
        or ("PARTICIPACAO" in combined and "CABINE" in combined and "JOGO" in combined)
    )


def _ensure_columns(df, columns, fill_value=""):
    for col in columns:
        if col not in df.columns:
            df[col] = fill_value


def _team_key(row):
    """Chave da equipe na mesma WO e na mesma janela do evento."""
    def value(keys):
        raw = _get_first(row, keys, "")
        if raw is None or pd.isna(raw):
            return ""
        return str(raw).strip()

    date_value = value(["Data", "Data_raw"])
    parsed_date = pd.to_datetime(date_value, errors="coerce", dayfirst=True)
    date_key = parsed_date.strftime("%Y-%m-%d") if pd.notna(parsed_date) else _norm_text(date_value)
    platform_key = _norm_text(value(["Plataforma", "Canal (Master Room)", "Canal"]))
    event_key = _norm_text(value(["Evento/Programa", "Descrição", "Atividade/Descrição", "Evento", "Event Group"]))
    start_key = _norm_text(value(["Início", "Inicio", "InÇðcio", "Air Start Time"]))
    end_key = _norm_text(value(["Fim", "FIM", "Air End Time"]))
    wo_key = value(["WO#"])
    return (wo_key, date_key, platform_key, event_key, start_key, end_key)


def process_2468_base(file_path):
    """
    Reads the 2468 report and consolidates roles by WO#.
    Preserves folgas and selected Quick Holds without WO# so they reach the check.
    """
    xl = pd.ExcelFile(file_path)
    sheet_to_read = None
    for sheet_name in ["base (2)", "base MP", "Base MP"]:
        if sheet_name in xl.sheet_names:
            sheet_to_read = sheet_name
            break

    if not sheet_to_read:
        sheet_to_read = xl.sheet_names[0]

    df = pd.DataFrame()
    for header_row in [0, 1, 2]:
        temp_df = xl.parse(sheet_to_read, header=header_row)
        if "WO #" in temp_df.columns:
            temp_df = temp_df.rename(columns={"WO #": "WO#"})
        if "WO#" in temp_df.columns or "Nome" in temp_df.columns or "Nome " in temp_df.columns:
            df = temp_df
            break

    if "WO#" not in df.columns:
        raise ValueError("Coluna 'WO#' nao encontrada no relatorio 2468.")

    if "Plataforma" not in df.columns and "Canal" in df.columns:
        df["Plataforma"] = df["Canal"]

    # GE TV possui fluxo de envio separado e não deve entrar nesta etapa.
    if not df.empty:
        ge_tv_mask = df.apply(_is_ge_tv_row, axis=1)
        df = df[~ge_tv_mask].copy()

    keep_extra_mask = pd.Series(False, index=df.index)
    if "Nome" in df.columns:
        keep_extra_mask = df.apply(lambda row: _is_folga_row(row) or _is_quickhold_in_scale(row), axis=1)

    df_extra = df[keep_extra_mask & df["Nome"].notna()].copy() if "Nome" in df.columns else pd.DataFrame()
    df_main = df[~keep_extra_mask].copy()
    df_main = df_main.dropna(subset=["Nome", "WO#"]) if "Nome" in df_main.columns else df_main

    if "Nome" in df_main.columns and "Função" not in df_main.columns:
        df_main["Função"] = ""
    if "Função" in df_main.columns:
        df_main["Função"] = df_main["Função"].fillna("")

    role_columns = ["Narrador", "Comentarista", "Repórter", "Coordenador", "Produtor", "Elenco"]
    _ensure_columns(df_main, role_columns, "")

    if "Nome" in df_main.columns:
        df_main["Nome"] = df_main["Nome"].fillna("")

    df_main["__team_key"] = df_main.apply(_team_key, axis=1)
    wo_team = {}
    for _, row in df_main.iterrows():
        wo = row["__team_key"]
        nome = str(row.get("Nome", "")).strip()
        funcao = _norm_text(_get_first(row, ["Função", "FunÇõÇœo"])).lower()

        if wo not in wo_team:
            wo_team[wo] = {
                "Narrador": [],
                "Comentarista": [],
                "Repórter": [],
                "Coordenador": [],
                "Produtor": [],
                "Elenco": [],
            }

        if "narrador" in funcao:
            wo_team[wo]["Narrador"].append(nome)
        elif "coment" in funcao:
            wo_team[wo]["Comentarista"].append(nome)
        elif "rep" in funcao or "reporter" in funcao:
            wo_team[wo]["Repórter"].append(nome)
        elif "coord" in funcao:
            wo_team[wo]["Coordenador"].append(nome)
        elif "produtor" in funcao:
            wo_team[wo]["Produtor"].append(nome)
        else:
            wo_team[wo]["Elenco"].append(nome)

    for wo in wo_team:
        for role in wo_team[wo]:
            wo_team[wo][role] = " ; ".join(sorted(set(filter(None, wo_team[wo][role]))))

    for role in role_columns:
        df_main[role] = df_main["__team_key"].apply(lambda w: wo_team.get(w, {}).get(role, ""))
    df_main = df_main.drop(columns=["__team_key"])

    if "Data" in df_main.columns:
        df_main["Data_raw"] = df_main["Data"]

    if not df_extra.empty:
        for col in ["Função", *role_columns, "Plataforma", "Data_raw"]:
            if col not in df_extra.columns:
                df_extra[col] = ""

        if "Nome" in df_extra.columns:
            df_extra["Nome"] = df_extra["Nome"].fillna("")
        if "Função" in df_extra.columns:
            df_extra["Função"] = df_extra["Função"].fillna("")
        if "Data" in df_extra.columns:
            df_extra["Data_raw"] = df_extra["Data"]

        for role in role_columns:
            df_extra[role] = ""

        if "Status" not in df_extra.columns:
            df_extra["Status"] = ""
        else:
            df_extra["Status"] = df_extra["Status"].fillna("")

        df_main = pd.concat([df_main, df_extra], ignore_index=True, sort=False)

    return df_main
