"""Testes legados preservados apenas como referência histórica.

O fluxo antigo chamava ``process_data``, método que não existe mais na
aplicação atual. A cobertura vigente está em ``test_regressions.py``.
"""

import unittest


@unittest.skip("Teste legado: process_data foi substituído pelo fluxo em três etapas.")
class LegacyProcessTests(unittest.TestCase):
    def test_legacy_process_flow(self):
        self.fail("Teste legado não deve ser executado na versão atual.")
