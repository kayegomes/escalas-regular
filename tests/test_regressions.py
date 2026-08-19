import json
import os
import tempfile
import unittest

import pandas as pd

from app_support import (
    OutputFileLockedError,
    append_execution_history,
    fallback_path,
    is_valid_email,
    safe_filename,
    save_with_fallback,
)
from engine_2405 import _event_score, _normalize_channel, _normalize_text
from engine_2468 import _is_ge_tv_row
from engine_cross import _score_grade_match
from engine_grades import _consolidate_grade_dataframe_windows, _extend_grade_windows_to_next_event, _merge_repeated_grade_windows, extract_sportv_channel_block, process_premiere_grade
try:
    from gerador_escalas_desktop import GeradorEscalasApp
except ModuleNotFoundError:
    GeradorEscalasApp = None
from engine_cross import (
    _append_pre_review_alerts,
    _mark_missing_grade,
    _find_best_grade_match,
    _is_folga_row,
    _is_programa,
    _is_quickhold_in_scale,
    _is_valid_time_str,
)


class AppSupportTests(unittest.TestCase):
    def test_safe_filename_removes_invalid_windows_characters(self):
        self.assertEqual(safe_filename('Ana: "Teste"/2026'), "Ana_Teste_2026")
        self.assertEqual(safe_filename("   "), "sem_nome")

    def test_email_validation(self):
        self.assertTrue(is_valid_email("pessoa@example.com"))
        self.assertFalse(is_valid_email("sem-email"))
        self.assertFalse(is_valid_email(""))

    def test_save_with_fallback_when_original_is_locked(self):
        calls = []

        def writer(path):
            calls.append(path)
            if len(calls) == 1:
                raise PermissionError("locked")
            with open(path, "w", encoding="utf-8") as file:
                file.write("ok")

        with tempfile.TemporaryDirectory() as directory:
            original = os.path.join(directory, "saida.xlsx")
            path, used_fallback = save_with_fallback(writer, original)
            self.assertEqual(path, fallback_path(original))
            self.assertTrue(used_fallback)
            self.assertTrue(os.path.exists(path))

    def test_save_with_fallback_raises_specific_error_when_both_are_locked(self):
        def writer(_path):
            raise PermissionError("locked")

        with tempfile.TemporaryDirectory() as directory:
            original = os.path.join(directory, "saida.xlsx")
            with self.assertRaises(OutputFileLockedError):
                save_with_fallback(writer, original)

    def test_history_is_json_lines(self):
        with tempfile.TemporaryDirectory() as directory:
            history = append_execution_history(directory, "Teste", "SUCESSO", "1 item")
            self.assertIsNotNone(history)
            with open(history, encoding="utf-8") as file:
                entry = json.loads(file.readline())
            self.assertEqual(entry["etapa"], "Teste")
            self.assertEqual(entry["status"], "SUCESSO")


class EngineRegressionTests(unittest.TestCase):
    def test_channel_and_text_normalization(self):
        self.assertEqual(_normalize_channel("Sportv 1"), "SPORTV")
        self.assertEqual(_normalize_channel("TV Globo Rede"), "TV GLOBO")
        self.assertEqual(_normalize_text("Irã x Alemanha"), "IRA X ALEMANHA")

    @unittest.skipIf(GeradorEscalasApp is None, "Tkinter não disponível neste ambiente")
    def test_html_preserves_elenco_product_and_event_fields(self):
        from pathlib import Path
        with tempfile.TemporaryDirectory() as tmp:
            df = pd.DataFrame([{
                "Nome": "André Felipe",
                "Plataforma": "Sportv 2",
                "Data": "10/08/2026",
                "Dia": "Monday",
                "Pré": "-",
                "Início": "12:00",
                "Fim": "13:45",
                "Evento/Programa": "EQUADOR X BRASIL",
                "Produto (WO/Quick Hold)": "CAMPEONATO SUL-AMERICANO MASCULINO DE FUTSAL SUB-17/2026/NA",
                "Local": "Internacional - Paraguai",
                "Elenco": "",
                "Narrador": "André Felipe",
                "Comentarista": "Vander Carioca",
                "Repórter": "",
                "Coordenador": "Tomaz Leão",
                "Produtor": "-",
            }])
            GeradorEscalasApp.gerar_html(object(), "André Felipe", df, tmp)
            html_path = next(Path(tmp).glob("escala_*.html"))
            html = html_path.read_text(encoding="utf-8")
            self.assertIn("André Felipe ; Vander Carioca", html)
            self.assertIn("CAMPEONATO SUL-AMERICANO MASCULINO DE FUTSAL SUB-17/2026/NA", html)
            self.assertIn("EQUADOR X BRASIL", html)

    def test_gecom_platform_matches_main_sportv_grade(self):
        row = pd.Series({
            "Plataforma": "GE.com 01",
            "Evento/Programa": "PANELA SPORTV",
            "Produto (WO/Quick Hold)": "PANELA SPORTV/NA/NA",
            "Event Group": "PANELA SPORTV",
            "Data_raw": "10/08/2026",
            "Início": "17:00",
            "Air Start Time": "17:00",
            "Tipo de Produção": "Exibição - Internet",
        })
        grades = pd.DataFrame([
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-10"), "Início": "23:30", "Pré": None, "Fim": "01:30", "Evento": "PANELA", "V/I": "I"},
        ])
        match = _find_best_grade_match(row, grades)
        self.assertIsNotNone(match)
        self.assertEqual(str(match["Plataforma"]), "SPORTV")
        self.assertEqual(str(match["Evento"]), "PANELA")

    def test_ppv_separate_pre_hora_attaches_to_following_event(self):
        from pathlib import Path
        ppv = Path('/home/ubuntu/escalas_regular/agosto_2026/PPV2026(Ago11ªversão).xlsx')
        if not ppv.exists():
            self.skipTest('Arquivo PPV de agosto não disponível')
        grades = process_premiere_grade(str(ppv))
        match = grades[(grades['Evento'].astype(str).str.upper().str.contains('FLUMINENSE X PALMEIRAS', na=False)) & (grades['Data'].astype(str).str.startswith('2026-08-15'))]
        self.assertEqual(len(match), 1)
        row = match.iloc[0]
        self.assertEqual(str(row['Início']), '16:30')
        self.assertEqual(str(row['Pré']), '15:30')
        self.assertEqual(str(row['Fim']), '18:40')

    def test_vt_inedito_futsal_matches_sportv2_same_day(self):
        row = pd.Series({
            "Plataforma": "Sportv 2",
            "Evento/Programa": "EQUADOR X BRASIL",
            "Produto (WO/Quick Hold)": "CAMPEONATO SUL-AMERICANO MASCULINO DE FUTSAL SUB-17/2026/NA",
            "Event Group": "CAMPEONATO SUL-AMERICANO MASCULINO DE FUTSAL SUB-17",
            "Data_raw": "10/08/2026",
            "Início": "12:00",
            "Air Start Time": "12:00",
            "Tipo de Produção": "Exibição - VT no Controle",
        })
        grades = pd.DataFrame([
            {"Plataforma": "SPORTV2", "Data": pd.Timestamp("2026-08-10"), "Início": "12:00", "Pré": None, "Fim": "13:45", "Evento": "CAMPEONATO SUL-AMERICANO MASCULINO DE FUTSAL SUB-17 - EQUADOR X BRASIL", "V/I": "I"},
        ])
        match = _find_best_grade_match(row, grades)
        self.assertIsNotNone(match)
        self.assertEqual(str(match["V/I"]), "I")
        self.assertEqual(str(match["Plataforma"]), "SPORTV2")

    def test_date_only_row_uses_report_start_time_and_same_day_match(self):
        row = pd.Series({
            "Plataforma": "Sportv",
            "Evento/Programa": "Sportv News",
            "Produto (WO/Quick Hold)": "SPORTV NEWS/1ª ED",
            "Event Group": "SPORTV NEWS",
            "Data_raw": "15/08/2026",
            "Início": "13:30",
            "Air Start Time": "13:30",
        })
        grades = pd.DataFrame([
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-14"), "Início": "21:30", "Pré": None, "Fim": "23:30", "Evento": "SPORTV NEWS", "V/I": "V"},
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-15"), "Início": "13:30", "Pré": None, "Fim": "14:00", "Evento": "SPORTV NEWS", "V/I": "V"},
        ])
        match = _find_best_grade_match(row, grades)
        self.assertIsNotNone(match)
        self.assertEqual(pd.to_datetime(match["Data"]).strftime("%Y-%m-%d"), "2026-08-15")
        self.assertEqual(str(match["Início"]), "13:30")

    def test_brazilian_report_date_matches_same_grade_date(self):
        row = pd.Series({
            "Plataforma": "Sportv",
            "Evento/Programa": "TROCA DE PASSES",
            "Produto (WO/Quick Hold)": "TROCA DE PASSES/NA/NA",
            "Event Group": "TROCA DE PASSES",
            "Data_raw": "10/08/2026",
        })
        grades = pd.DataFrame([
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-10"), "Início": "22:00", "Pré": None, "Fim": "23:30", "Evento": "TROCA DE PASSES", "V/I": "V"},
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-10-08"), "Início": "22:00", "Pré": None, "Fim": "23:30", "Evento": "TROCA DE PASSES", "V/I": "V"},
        ])
        match = _find_best_grade_match(row, grades)
        self.assertIsNotNone(match)
        self.assertEqual(pd.to_datetime(match["Data"]).strftime("%Y-%m-%d"), "2026-08-10")

    def test_ge_tv_channel_is_excluded(self):
        self.assertTrue(_is_ge_tv_row(pd.Series({"Canal": "GE TV", "Plataforma": "GE TV"})))
        self.assertTrue(_is_ge_tv_row(pd.Series({"Canal": "GE-TV"})))
        self.assertFalse(_is_ge_tv_row(pd.Series({"Canal": "Sportv", "Plataforma": "Sportv 2"})))

    def test_daytime_event_does_not_match_same_title_on_other_date(self):
        row = pd.Series({
            "Plataforma": "Sportv 3",
            "Evento/Programa": "COPA DO MUNDO DE CANOAGEM VELOCIDADE E PARACANOAGEM - SPRINT,PARACANOAGEM E SPRINT 5KM",
            "Produto (WO/Quick Hold)": "COPA DO MUNDO DE CANOAGEM DE VELOCIDADE/2026/NA",
            "Event Group": "COPA DO MUNDO DE CANOAGEM VELOCIDADE E PARACANOAGEM",
            "Data_raw": pd.Timestamp("2026-07-15 14:00:00"),
        })
        grades = pd.DataFrame([
            {
                "Plataforma": "SPORTV3",
                "Data": pd.Timestamp("2026-07-11"),
                "Início": "17:00",
                "Pré": None,
                "Fim": "19:00",
                "Evento": "COPA DO MUNDO DE CANOAGEM VELOCIDADE E PARACANOAGEM - SPRINT E PARACANOAGEM",
                "V/I": "I",
            }
        ])
        self.assertIsNone(_find_best_grade_match(row, grades))

    def test_generic_vt_grade_does_not_match_specific_event(self):
        grade = pd.Series({"Evento": "VT DE EVENTO", "Início": "14:00", "Pré": None})
        score = _score_grade_match(grade, "COPA DO MUNDO DE CANOAGEM VELOCIDADE E PARACANOAGEM", 14 * 3600)
        self.assertEqual(score, -100)

    def test_event_score_prefers_exact_match(self):
        row = pd.Series({"W/O Description": "LIGA DAS NAÇÕES MASCULINA DE VÔLEI - IRÃ X ALEMANHA"})
        self.assertEqual(
            _event_score("LIGA DAS NAÇÕES MASCULINA DE VÔLEI - IRÃ X ALEMANHA", row),
            1.0,
        )

    def test_folga_and_quick_hold_rules(self):
        folga = pd.Series({"Tipo de Atividade": "Booking", "Descrição": "FOLGA"})
        quick_hold = pd.Series({"Tipo de Atividade": "QUICK HOLD", "Descrição": "PODCAST"})
        viagem = pd.Series({"Tipo de Atividade": "QUICK HOLD", "Descrição": "PODCAST VIAGEM"})
        self.assertTrue(_is_folga_row(folga))
        self.assertTrue(_is_quickhold_in_scale(quick_hold))
        self.assertFalse(_is_quickhold_in_scale(viagem))

    def test_time_validation_rejects_undefined_values(self):
        self.assertTrue(_is_valid_time_str("01:30"))
        self.assertFalse(_is_valid_time_str("A Definir"))
        self.assertFalse(_is_valid_time_str("-"))

    def test_vt_in_grade_is_not_treated_as_program(self):
        row = pd.Series({
            "Tipo de Produção": "Exibição - VT no Controle",
            "Tipo de Atividade": "Booking",
            "Evento/Programa": "COPA DO MUNDO DE CANOAGEM VELOCIDADE E PARACANOAGEM",
        })
        grade = pd.Series({"Evento": "COPA DO MUNDO DE CANOAGEM VELOCIDADE E PARACANOAGEM", "V/I": "I"})
        self.assertFalse(_is_programa(row, grade))

    def test_vt_in_grade_with_live_indicator_is_not_treated_as_program(self):
        row = pd.Series({
            "Tipo de Produção": "Exibição - VT no Controle",
            "Tipo de Atividade": "Booking",
            "Evento/Programa": "EVENTO ESPORTIVO AO VIVO",
        })
        grade = pd.Series({"Evento": "EVENTO ESPORTIVO AO VIVO", "V/I": "V"})
        self.assertFalse(_is_programa(row, grade))

    def test_sportv_parser_uses_prejogo_as_previous_window_boundary(self):
        df = pd.DataFrame([
            ["SÁB", "2026-08-15", "13:30", "V", "SPORTV NEWS", "", "PROGRAMA", "SPORTV", "", 1/24, ""],
            ["SÁB", "2026-08-15", "14:00", "V", "PRÉ-JOGO", "", "PROGRAMA", "SPORTV", "", 1/48, ""],
        ], columns=["dia", "data", "hora", "vi", "evento", "obs", "tipo", "canal", "extra", "duracao", "extra2"])
        events = []
        extract_sportv_channel_block(df, 4, events)
        news = [event for event in events if event["Evento"] == "SPORTV NEWS"][0]
        self.assertEqual(str(news["Fim"]), "14:00")

    def test_repeated_sportv_news_window_reaches_1430(self):
        grades = pd.DataFrame([
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-15"), "Início": "13:30", "Fim": "14:30", "Evento": "SPORTV NEWS", "V/I": "V"},
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-15"), "Início": "14:00", "Fim": "14:30", "Evento": "SPORTV NEWS", "V/I": "I"},
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-15"), "Início": "14:30", "Fim": "14:30", "Evento": "PRÉ-JOGO", "V/I": "V"},
        ])
        consolidated = _consolidate_grade_dataframe_windows(grades)
        news = consolidated[consolidated["Evento"] == "SPORTV NEWS"].iloc[0]
        self.assertEqual(str(news["Início"]), "13:30")
        self.assertEqual(str(news["Fim"]), "14:30")

    def test_dataframe_window_uses_next_event_boundary(self):
        grades = pd.DataFrame([
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-10"), "Início": "22:00", "Fim": "22:30", "Evento": "TROCA DE PASSES", "V/I": "V"},
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-10"), "Início": "22:30", "Fim": "23:30", "Evento": "HELLO LA", "V/I": "R"},
            {"Plataforma": "SPORTV", "Data": pd.Timestamp("2026-08-10"), "Início": "23:30", "Fim": "23:30", "Evento": "PANELA", "V/I": "I"},
        ])
        consolidated = _consolidate_grade_dataframe_windows(grades)
        troca = consolidated[consolidated["Evento"] == "TROCA DE PASSES"].iloc[0]
        self.assertEqual(str(troca["Fim"]), "23:30")

    def test_program_window_extends_to_next_distinct_event(self):
        events = [
            {"Plataforma": "SPORTV", "Data": "2026-08-10", "Início": "22:00", "Fim": "22:30", "Evento": "TROCA DE PASSES", "V/I": "V"},
            {"Plataforma": "SPORTV", "Data": "2026-08-10", "Início": "23:30", "Fim": "23:30", "Evento": "PANELA", "V/I": "I"},
        ]
        extended = _extend_grade_windows_to_next_event(events)
        self.assertEqual(extended[0]["Fim"], "23:30")

    def test_repeated_uts_markers_are_merged_until_next_distinct_event(self):
        events = [
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "18:00", "Fim": "19:00", "Evento": "UTS - 1ª RODADA", "V/I": "V"},
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "18:30", "Fim": "19:00", "Evento": "UTS - RIO DE JANEIRO, BRASIL", "V/I": None},
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "19:00", "Fim": "20:00", "Evento": "UTS - 1ª RODADA", "V/I": "V"},
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "19:30", "Fim": "20:00", "Evento": "UTS - RIO DE JANEIRO, BRASIL", "V/I": None},
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "20:00", "Fim": "21:00", "Evento": "UTS - 1ª RODADA", "V/I": "V"},
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "20:30", "Fim": "21:00", "Evento": "UTS - RIO DE JANEIRO, BRASIL", "V/I": None},
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "21:00", "Fim": "22:00", "Evento": "UTS - 1ª RODADA", "V/I": "V"},
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "21:30", "Fim": "22:00", "Evento": "UTS - RIO DE JANEIRO, BRASIL", "V/I": None},
            {"Plataforma": "SPORTV3", "Data": "2026-07-16", "Início": "22:00", "Fim": "23:00", "Evento": "LIGA DAS NAÇÕES MASCULINA DE VÔLEI - EUA X BRASIL", "V/I": "V"},
        ]
        merged = _merge_repeated_grade_windows(events)
        uts = [event for event in merged if event["Evento"] == "UTS - 1ª RODADA"]
        self.assertEqual(len(uts), 1)
        self.assertEqual(uts[0]["Início"], "18:00")
        self.assertEqual(uts[0]["Fim"], "22:00")
        self.assertEqual(len(merged), 6)

    def test_vt_grade_without_vi_indicator_is_treated_as_program(self):
        row = pd.Series({
            "Tipo de Produção": "Exibição - VT no Controle",
            "Tipo de Atividade": "Booking",
            "Evento/Programa": "VT DE EVENTO ESPECIAL",
        })
        grade = pd.Series({"Evento": "VT DE EVENTO ESPECIAL", "V/I": None})
        self.assertTrue(_is_programa(row, grade))

    def test_vt_without_grade_remains_program(self):
        row = pd.Series({
            "Tipo de Produção": "Exibição - VT no Controle",
            "Tipo de Atividade": "Booking",
            "Evento/Programa": "VT DE EVENTO ESPECIAL",
        })
        self.assertTrue(_is_programa(row, None))

    def test_event_without_pre_requires_review(self):
        out_row = {"Pré": "-", "Início": "22:00", "Fim": "02:00"}
        alerts = []
        severity = _append_pre_review_alerts(out_row, alerts, "OK", is_prog=False, has_valid_grade_time=True)
        self.assertEqual(severity, "YELLOW")
        self.assertEqual(alerts, ["Conferir Pré"])

    def test_missing_grade_is_flagged_even_for_program(self):
        out_row = {"Pré": "", "Início": "09:00", "Fim": "13:00"}
        alerts = []
        severity = _mark_missing_grade(out_row, alerts, "OK")
        self.assertEqual(severity, "YELLOW")
        self.assertEqual(alerts, ["Horário não encontrado na Grade"])
        self.assertEqual(out_row["Início"], "09:00")
        self.assertEqual(out_row["Fim"], "13:00")
        self.assertEqual(out_row["Pré"], "-")

    def test_program_without_pre_remains_ok(self):
        out_row = {"Pré": "-", "Início": "22:00", "Fim": "02:00"}
        alerts = []
        severity = _append_pre_review_alerts(out_row, alerts, "OK", is_prog=True, has_valid_grade_time=True)
        self.assertEqual(severity, "OK")
        self.assertEqual(alerts, [])

    def test_event_with_pre_does_not_require_pre_review(self):
        out_row = {"Pré": "21:30", "Início": "22:00", "Fim": "02:00"}
        alerts = []
        severity = _append_pre_review_alerts(out_row, alerts, "OK", is_prog=False, has_valid_grade_time=True)
        self.assertEqual(severity, "OK")
        self.assertEqual(alerts, [])

    def test_one_fight_night_45_matches_combate_grade(self):
        base_row = pd.Series(
            {
                "Plataforma": "Combate",
                "Data_raw": pd.Timestamp("2026-07-17 22:00:00"),
                "Evento/Programa": "ONE FIGHT NIGHT 45",
                "Produto (WO/Quick Hold)": "ONE FRIDAY FIGHTS/2026/NA",
                "Event Group": "ONE FIGHT NIGHT",
            }
        )
        grades = pd.DataFrame(
            [
                {
                    "Plataforma": "COMBATE",
                    "Data": pd.Timestamp("2026-07-17"),
                    "Início": "22:00:00",
                    "Pré": "X",
                    "Fim": "02:00:00",
                    "Evento": "ONE FIGHT NIGHT 45",
                    "V/I": "V",
                }
            ]
        )
        match = _find_best_grade_match(base_row, grades)
        self.assertIsNotNone(match)
        self.assertEqual(match["Evento"], "ONE FIGHT NIGHT 45")


if __name__ == "__main__":
    unittest.main()
