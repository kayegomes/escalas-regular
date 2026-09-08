import pandas as pd
import numpy as np

def _find_header_row(df_raw, keywords=("EVENTO", "DATA")):
    for i, row in df_raw.iterrows():
        row_text = " | ".join(str(v) for v in row.values).upper()
        if all(keyword in row_text for keyword in keywords):
            return i
    return 0

def flatten_sportv_grade(file_path):
    """
    Reads the horizontal SporTV grid and flattens it.
    """
    try:
        xl = pd.ExcelFile(file_path)
        # Assume first sheet is the grade
        df_raw = xl.parse(xl.sheet_names[0], header=None)

        # In the SporTV grid, row 3 (index 2) usually contains the dates? No, let's search for "DATA GRADE"
        start_row = 0
        for i, row in df_raw.iterrows():
            if 'DATA' in str(row.values).upper() or 'EVENTO' in str(row.values).upper():
                start_row = i
                break

        df = xl.parse(xl.sheet_names[0], header=start_row)

        date_col = None
        for c in df.columns:
            if "DATA REAL" in str(c).upper() or "DATA GRADE" in str(c).upper():
                date_col = c
                break

        if date_col is not None:
            df[date_col] = df[date_col].ffill()

        flat_events = []

        # In the horizontal layout, we have repeating blocks for each channel.
        # We look for columns that have 'Evento' or 'Programa' in their name
        evento_cols = [c for c in df.columns if 'EVENTO' in str(c).upper() or 'PROGRAMA' in str(c).upper()]

        for c in evento_cols:
            # We need to find the related Data, Hora, V/I for this specific Evento column.
            # Usually, they are the columns immediately to the left.
            idx = df.columns.get_loc(c)
            # Ensure idx is an integer, if there are duplicate column names, get_loc returns a slice or boolean array
            if isinstance(idx, (slice, np.ndarray)):
                # If there are duplicates, we need to iterate over indices. 
                # Pandas handles duplicate columns by adding .1, .2, etc. if read properly, 
                # but just in case, we grab the first match if it's an array.
                if isinstance(idx, slice):
                    idx_list = range(idx.start, idx.stop, idx.step or 1)
                else:
                    idx_list = np.where(idx)[0]

                for i in idx_list:
                    extract_sportv_channel_block(df, i, flat_events, date_col)
            else:
                extract_sportv_channel_block(df, idx, flat_events, date_col)

        df_flat = pd.DataFrame(flat_events)
        if not df_flat.empty:
            # Clean up
            df_flat = df_flat.dropna(subset=['Evento'])
            df_flat = df_flat[df_flat['Evento'] != '']
        return df_flat
    except Exception as e:
        print(f"Erro processando Grade SporTV: {e}")
        return pd.DataFrame()

def _parse_hour(value):
    if value is None or pd.isna(value):
        return None
    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.notna(dt):
            return dt.hour
    except Exception:
        pass
    text = str(value).strip()
    if ":" in text:
        part = text.split(" ")[-1].split(":")[0]
        if part.isdigit():
            return int(part)
    return None


def _time_key(value):
    if value is None or pd.isna(value):
        return None
    if hasattr(value, "hour") and hasattr(value, "minute"):
        return int(value.hour) * 60 + int(value.minute)
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "nat"}:
        return None
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed):
            return parsed.hour * 60 + parsed.minute
    except Exception:
        pass
    parts = text.split(":")
    if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
        return int(parts[0]) * 60 + int(parts[1])
    return None


def _event_key(value):
    return " ".join(str(value or "").upper().split())


def _ppv_match_key(data_key, mandante, visitante, evento):
    """Chave do confronto PPV, independente do canal textual da linha."""
    home = _event_key(mandante)
    away = _event_key(visitante)
    if home and away and home not in {"NAN", "NONE"} and away not in {"NAN", "NONE"}:
        return f"{data_key}|{home}|{away}"
    return f"{data_key}|{_event_key(evento)}"


def _elapsed_minutes(start, end):
    start_key = _time_key(start)
    end_key = _time_key(end)
    if start_key is None or end_key is None:
        return None
    return (end_key - start_key) % (24 * 60)


def _extend_grade_windows_to_next_event(block_events, max_gap_minutes=90):
    """Estende blocos curtos até o próximo evento distinto da mesma grade.

    Algumas versões da grade exibem um evento em vários marcadores sucessivos,
    enquanto o parser recebe apenas o primeiro marcador com `V`. Se o próximo
    evento distinto começa pouco depois do fim calculado, esse início é o fim
    real da janela anterior.
    """

    for index, event in enumerate(block_events):
        event_vi = str(event.get("V/I", "")).strip().upper()
        if event_vi not in {"V", "I", "AO VIVO", "LIVE"}:
            continue

        event_start = _time_key(event.get("Início"))
        event_end_elapsed = _elapsed_minutes(event.get("Início"), event.get("Fim"))
        if event_start is None or event_end_elapsed is None:
            continue

        event_key = (
            str(event.get("Plataforma", "")).upper(),
            str(event.get("Data", ""))[:10],
        )
        for next_event in block_events[index + 1:]:
            next_key = (
                str(next_event.get("Plataforma", "")).upper(),
                str(next_event.get("Data", ""))[:10],
            )
            if next_key != event_key:
                continue
            next_vi = str(next_event.get("V/I", "")).strip().upper()
            if next_vi not in {"V", "R", "REPRISE", "I", "AO VIVO", "LIVE"}:
                continue
            if _event_key(next_event.get("Evento")) == _event_key(event.get("Evento")):
                continue

            next_start_elapsed = _elapsed_minutes(event.get("Início"), next_event.get("Início"))
            if next_start_elapsed is None:
                break

            # Uma reprise distinta também pode encerrar a janela anterior.
            # Exemplo: Argentina x Venezuela termina quando a grade inicia
            # uma reprise às 12:30, mesmo que uma linha V posterior esteja
            # disponível às 13:30.
            if next_vi in {"R", "REPRISE"}:
                if next_start_elapsed <= event_end_elapsed or next_start_elapsed - event_end_elapsed <= max_gap_minutes:
                    next_boundary = next_event.get("Pré") if next_event.get("Pré") is not None else next_event.get("Início")
                    event["Fim"] = next_boundary
                break

            if next_start_elapsed <= event_end_elapsed:
                break
            if next_start_elapsed - event_end_elapsed <= max_gap_minutes:
                next_boundary = next_event.get("Pré") if next_event.get("Pré") is not None else next_event.get("Início")
                event["Fim"] = next_boundary
            break

    return block_events


def _merge_repeated_grade_windows(block_events):
    """Consolida marcadores repetidos do mesmo evento em uma janela única.

    Grades de eventos esportivos podem repetir o mesmo título a cada bloco de
    30/60 minutos. Quando o próximo marcador do mesmo evento começa exatamente
    no ``Fim`` calculado do marcador anterior, as ocorrências são partes de uma
    mesma transmissão. Mantemos a primeira linha e estendemos o ``Fim`` até o
    último marcador antes do próximo evento distinto.
    """

    merged = []
    last_index_by_key = {}
    for event in block_events:
        data_key = str(event.get("Data", ""))[:10]
        key = (
            str(event.get("Plataforma", "")).upper(),
            data_key,
            _event_key(event.get("Evento")),
        )
        previous_index = last_index_by_key.get(key)
        if previous_index is not None:
            previous = merged[previous_index]
            previous_end = _time_key(previous.get("Fim"))
            current_start = _time_key(event.get("Início"))
            if previous_end is not None and current_start is not None and current_start <= previous_end:
                current_end = _time_key(event.get("Fim"))
                if current_end is not None and (previous_end is None or current_end > previous_end):
                    previous["Fim"] = event.get("Fim")
                continue

        last_index_by_key[key] = len(merged)
        merged.append(dict(event))

    return merged


def extract_sportv_channel_block(df, evento_col_idx, flat_events, date_col=None):
    col_evento = df.columns[evento_col_idx]
    col_obs = df.columns[evento_col_idx + 1] if (evento_col_idx + 1) < len(df.columns) else None

    # Try to find Data, Hora, V/I to the left
    col_data = df.columns[evento_col_idx - 3] if evento_col_idx >= 3 else None
    col_hora = df.columns[evento_col_idx - 2] if evento_col_idx >= 2 else None
    col_vi = df.columns[evento_col_idx - 1] if evento_col_idx >= 1 else None
    col_canal = df.columns[evento_col_idx + 3] if (evento_col_idx + 3) < len(df.columns) else None

    # Os blocos horizontais usam cinco colunas por canal:
    # Data, Hora, V/I, Evento/Programa e Observação. A coluna seguinte
    # ao Evento pertence ao bloco seguinte e não deve ser usada como canal.
    # O canal é determinado pela posição do bloco na grade.
    block_number = max(0, (evento_col_idx - 10) // 5)
    canal = {
        0: "SPORTV",
        1: "SPORTV2",
        2: "SPORTV3",
        3: "SPORTV4",
    }.get(block_number, f"SPORTV{block_number + 1}")

    last_aquecimento_by_channel = {}
    last_aquecimento_date_by_channel = {}
    current_date_by_channel = {}
    block_events = []

    for _, row in df.iterrows():
        # Não ler col_canal: no layout mensal essa posição normalmente é a
        # coluna Data do próximo bloco e causava a perda de linhas válidas.
        row_channel = canal

        evento = str(row[col_evento]).strip()
        if evento.lower() in ['nan', 'none', '']:
            continue

        obs = str(row[col_obs]).strip() if col_obs and pd.notna(row[col_obs]) else ""
        if obs and obs.lower() not in ['nan', 'none', '']:
            if obs.upper() not in evento.upper():
                evento = f"{evento} - {obs}"

        row_date = None
        if date_col and pd.notna(row[date_col]):
            row_date = row[date_col]
        elif col_data and pd.notna(row[col_data]):
            row_date = row[col_data]

        hora = row[col_hora] if col_hora else None
        hour_val = _parse_hour(hora)

        current_date = current_date_by_channel.get(row_channel)
        if row_date is not None and str(row_date).strip() not in ["", "nan", "NaT"]:
            if not (hour_val is not None and hour_val < 6 and current_date is not None):
                current_date = row_date
                current_date_by_channel[row_channel] = current_date

        vi = row[col_vi] if col_vi else "V"

        ev_upper = evento.upper().strip()
        is_pre_row = (
            'AQUECIMENTO' in ev_upper
            or ev_upper in {'PRÉ', 'PRE', 'PRÉ-JOGO', 'PRE-JOGO', 'PRÉ JOGO', 'PRE JOGO'}
            or ev_upper.startswith('PRÉ-JOGO')
            or ev_upper.startswith('PRE-JOGO')
        )
        if is_pre_row:
            # O bloco de Pré/Pré-Jogo também encerra o conteúdo imediatamente
            # anterior na mesma plataforma e data. Como grades longas repetem
            # o mesmo título em várias linhas, o limite deve ser aplicado a
            # toda a sequência consecutiva, não apenas à última repetição.
            boundary_time = _time_key(hora)
            previous_event_key = _event_key(block_events[-1].get("Evento")) if block_events else None
            if boundary_time is not None and previous_event_key:
                for previous in reversed(block_events):
                    same_context = (
                        str(previous.get("Plataforma", "")).upper() == str(row_channel).upper()
                        and str(previous.get("Data", ""))[:10] == str(current_date)[:10]
                        and _event_key(previous.get("Evento")) == previous_event_key
                        and _time_key(previous.get("Início")) is not None
                        and boundary_time > _time_key(previous.get("Início"))
                    )
                    if same_context:
                        # O primeiro AQUECIMENTO/Pré já encerra a sequência.
                        # Aquecimentos posteriores do mesmo bloco não podem
                        # estender novamente a janela até o jogo seguinte.
                        if _time_key(previous.get("Fim")) is None:
                            previous["Fim"] = hora
                    elif str(previous.get("Data", ""))[:10] == str(current_date)[:10]:
                        break
            previous_pre_date = last_aquecimento_date_by_channel.get(row_channel)
            current_date_key = str(current_date)[:10] if current_date is not None else ""
            previous_pre_date_key = str(previous_pre_date)[:10] if previous_pre_date is not None else ""
            if last_aquecimento_by_channel.get(row_channel) is None or current_date_key != previous_pre_date_key:
                last_aquecimento_by_channel[row_channel] = hora
                last_aquecimento_date_by_channel[row_channel] = current_date
            continue

        pre = last_aquecimento_by_channel.get(row_channel)

        block_events.append({
            'Plataforma': row_channel,
            'Data': current_date,
            'Início': hora,
            'Pré': pre,
            'Fim': None,
            'Evento': evento,
            'V/I': vi
        })

        last_aquecimento_by_channel[row_channel] = None
        last_aquecimento_date_by_channel[row_channel] = None

    # Calculate Fim: skip filler/generic blocks (no V/R marker) and find
    # the next live (V) or reprise (R) broadcast event. The Pré of that
    # next event is the real end of the current broadcast block.
    for i in range(len(block_events)):
        # A linha de AQUECIMENTO/Pré já pode ter definido o fim real do
        # evento anterior. Não substituí-lo pelo próximo V/I mais distante.
        if _time_key(block_events[i].get('Fim')) is not None:
            continue
        for j in range(i + 1, len(block_events)):
            current = block_events[i]
            nxt = block_events[j]
            if str(nxt.get('Plataforma', '')).upper() != str(current.get('Plataforma', '')).upper():
                continue
            if str(nxt.get('Data', ''))[:10] != str(current.get('Data', ''))[:10]:
                continue
            vi_val = str(nxt.get('V/I', '')).strip().upper()
            if vi_val in ('V', 'R', 'AO VIVO', 'REPRISE', 'I'):
                next_time = nxt['Pré'] if nxt['Pré'] is not None else nxt['Início']
                block_events[i]['Fim'] = next_time
                break

    block_events = _merge_repeated_grade_windows(block_events)
    block_events = _extend_grade_windows_to_next_event(block_events)
    flat_events.extend(block_events)

def _consolidate_grade_dataframe_windows(df):
    """Consolida e estende janelas depois que todos os blocos foram reunidos."""

    if df is None or df.empty or not {"Plataforma", "Data", "Início", "Fim", "Evento"}.issubset(df.columns):
        return df

    consolidated = []
    work = df.copy()
    work["_grade_date_key"] = work["Data"].apply(lambda value: str(value)[:10])
    group_cols = ["Plataforma", "_grade_date_key"]
    for _, part in work.groupby(group_cols, sort=False, dropna=False):
        part = part.copy()
        part["_grade_start_key"] = part["Início"].apply(_time_key)
        part = part.sort_values("_grade_start_key", kind="stable")
        events = part.drop(columns=["_grade_date_key", "_grade_start_key"]).to_dict("records")
        events = _merge_repeated_grade_windows(events)
        events = _extend_grade_windows_to_next_event(events)
        consolidated.extend(events)

    return pd.DataFrame(consolidated)


def process_premiere_grade(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        df_raw = xl.parse(xl.sheet_names[0], header=None)
        start_row = _find_header_row(df_raw, keywords=("EVENTO", "DATA", "CANAL"))
        df = xl.parse(xl.sheet_names[0], header=start_row)

        events = []
        pending_pre = {}
        for _, row in df.iterrows():
            evento = str(row.get('EVENTO', '')).strip()
            if evento.lower() in ['nan', 'none', '']:
                continue

            data_value = row.get('DATA')
            data_key = str(data_value)[:10]
            canal = str(row.get('CANAL', 'PREMIERE'))
            channel_key = canal.strip().upper() or 'PREMIERE'
            inicio_value = row.get('HORA') if 'HORA' in df.columns else row.get(':ORA')
            mandante = str(row.get('MANDANTE', '')).strip()
            visitante = str(row.get('VISITANTE', '')).strip()
            matchup_key = _ppv_match_key(data_key, mandante, visitante, evento)

            # No PPV, um Pré com mais de 30 minutos pode aparecer como uma
            # linha separada (por exemplo, FLUMINENSE X PALMEIRAS - PRÉ-HORA).
            # Essa linha é um marcador de Pré, não uma atividade independente.
            evento_norm = evento.upper().replace('É', 'E').replace('-', ' ').strip()
            if 'PRE HORA' in evento_norm or evento_norm in {'PRE', 'AQUECIMENTO'}:
                pre_value = inicio_value if pd.notna(inicio_value) else row.get('PRÉ')
                pending_pre[matchup_key] = pre_value
                continue

            # Construct Event name from Mandante X Visitante if Evento is just "BRASILEIRO"
            if mandante and visitante and mandante.lower() != 'nan' and visitante.lower() != 'nan':
                evento_full = f"{mandante} X {visitante} - {evento}"
            else:
                evento_full = evento

            pre_value = pending_pre.pop(matchup_key, None)
            if pre_value is None or pd.isna(pre_value):
                pre_value = row.get('PRÉ')

            events.append({
                'Plataforma': 'PREMIERE', # we'll extract TV Globo later
                'Data': data_value,
                'Início': inicio_value,
                'Pré': pre_value,
                'Fim': row.get('PÓS'),
                'Evento': evento_full,
                'V/I': 'V',
                'Raw_Canal': canal
            })

        return pd.DataFrame(events)
    except Exception as e:
        print(f"Erro processando Grade Premiere: {e}")
        return pd.DataFrame()

def process_combate_grade(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        df_raw = xl.parse(xl.sheet_names[0], header=None)
        start_row = _find_header_row(df_raw, keywords=("EVENTO", "DATA", "COMBATE"))
        df = xl.parse(xl.sheet_names[0], header=start_row)

        events = []
        for _, row in df.iterrows():
            evento = str(row.get('EVENTO', '')).strip()
            if evento.lower() in ['nan', 'none', '']:
                continue

            canal_sportv = str(row.get('SPORTV', ''))

            events.append({
                'Plataforma': 'COMBATE',
                'Data': row.get('DATA'),
                'Início': row.get('INÍCIO'),
                'Pré': row.get('PRÉ'),
                'Fim': row.get('FIM'),
                'Evento': evento,
                'V/I': 'V',
                'Raw_Sportv': canal_sportv
            })

        return pd.DataFrame(events)
    except Exception as e:
        print(f"Erro processando Grade Combate: {e}")
        return pd.DataFrame()

def process_all_grades(path_sp1, path_sp2, path_pr1, path_pr2, path_co1, path_co2):
    """ Processes all grids, extracts TV Globo shared events, and returns a single DataFrame """
    import os

    def load_grade(path, func):
        if path and os.path.exists(path):
            return func(path)
        return pd.DataFrame()

    df_sp1 = load_grade(path_sp1, flatten_sportv_grade)
    df_sp2 = load_grade(path_sp2, flatten_sportv_grade)
    df_pr1 = load_grade(path_pr1, process_premiere_grade)
    df_pr2 = load_grade(path_pr2, process_premiere_grade)
    df_co1 = load_grade(path_co1, process_combate_grade)
    df_co2 = load_grade(path_co2, process_combate_grade)

    df_sp = pd.concat([df for df in [df_sp1, df_sp2] if not df.empty], ignore_index=True) if not df_sp1.empty or not df_sp2.empty else pd.DataFrame()
    df_sp = _consolidate_grade_dataframe_windows(df_sp)
    df_pr = pd.concat([df for df in [df_pr1, df_pr2] if not df.empty], ignore_index=True) if not df_pr1.empty or not df_pr2.empty else pd.DataFrame()
    df_co = pd.concat([df for df in [df_co1, df_co2] if not df.empty], ignore_index=True) if not df_co1.empty or not df_co2.empty else pd.DataFrame()

    all_events = []

    if not df_sp.empty:
        all_events.append(df_sp)

    if not df_pr.empty:
        # Extract TV Globo from Premiere
        # e.g., 'CANAL' == 'PRE/TVG/GE TV'
        df_pr_globo = df_pr[df_pr['Raw_Canal'].astype(str).str.contains('TVG|GLOBO|GE TV', case=False, na=False)].copy()
        if not df_pr_globo.empty:
            df_pr_globo['Plataforma'] = 'TV GLOBO'
            all_events.append(df_pr_globo)
        all_events.append(df_pr)

    if not df_co.empty:
        # Extract TV Globo / SporTV from Combate
        df_co_sp = df_co[df_co['Raw_Sportv'].astype(str).str.contains('SPORTV', case=False, na=False)].copy()
        if not df_co_sp.empty:
            # We can't know exactly if it's SP1, SP2 or SP3 without parsing the text carefully,
            # but usually it says 'COMPARTILHADO COM O SPORTV 3'
            df_co_sp['Plataforma'] = 'SPORTV' # generic
            all_events.append(df_co_sp)
        all_events.append(df_co)

    if not all_events:
        return pd.DataFrame()

    df_all = pd.concat(all_events, ignore_index=True)

    # Filter out Replays
    if 'V/I' in df_all.columns:
        df_all = df_all[~df_all['V/I'].astype(str).str.upper().isin(['R', 'REPRISE'])]

    return df_all
