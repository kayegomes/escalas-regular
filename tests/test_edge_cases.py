"""Teste manual legado da Etapa 3.

Este arquivo dependia de uma sessão Tkinter e da assinatura anterior de
``process_etapa2``. Os casos de horários continuam cobertos pelas funções de
regressão independentes em ``test_regressions.py``.
"""

import unittest


@unittest.skip("Teste manual legado: a Etapa 3 agora recebe filtros e opções explicitamente.")
class LegacyEdgeCaseTests(unittest.TestCase):
    def test_legacy_edge_cases(self):
        self.fail("Teste manual legado não deve ser executado automaticamente.")
