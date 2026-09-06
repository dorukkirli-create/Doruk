"""Ucuncu tur: okuyucu katmaninin uyarisi satiri OTOMATIK'ten dusurur."""
from __future__ import annotations

import unittest
from datetime import date

from masraf.masraf_merkezi import _durum_belirle, _okuyucu_uyarilari
from masraf.modeller import GiderSatiri


def _satir(**ek) -> GiderSatiri:
    return GiderSatiri(
        kaynak_dosya="x.xlsx", kaynak_tip="genel", satir_no=1, kisi_ham="ALI VELI",
        sicil_ham=None, tckn_ham=None, tutar=10.0, para_birimi="USD",
        belge_tarihi=date(2026, 7, 1), gider_tipi="Bilet", aciklama="",
        masraf_merkezi_kaynak=None, ek=dict(ek),
    )


class OkuyucuUyarisiTest(unittest.TestCase):
    def test_uyari_yoksa_bos(self):
        self.assertEqual(_okuyucu_uyarilari(_satir()), [])

    def test_uyari_listeye_girer_ve_incele_yapar(self):
        u = _okuyucu_uyarilari(_satir(okuyucu_uyarisi="ozel okuyucu sablonu tanimadi"))
        self.assertEqual(len(u), 1)
        self.assertTrue(u[0].startswith("OKUYUCU:"))
        # Guven tam olsa bile uyari varken OTOMATIK olamaz.
        self.assertEqual(_durum_belirle(1.0, u, "GPP", 0.90, 0.50), "INCELE")
        self.assertEqual(_durum_belirle(1.0, [], "GPP", 0.90, 0.50), "OTOMATIK")

    def test_ek_sozluk_degilse_cokmez(self):
        s = _satir()
        s.ek = None
        self.assertEqual(_okuyucu_uyarilari(s), [])


if __name__ == "__main__":
    unittest.main()
