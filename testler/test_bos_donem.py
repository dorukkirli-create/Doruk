"""Donemi BOS personel satirlarinin en guncel kaydi bozmadigini dogrular.

Personel ana verisi her ay yeniden uretilir. Bir satirin 'Donem' hucresi bos
gelirse bu satir donem siralamasinda SONA dizilir. Duzeltme oncesinde
``sicil_ile()`` ve tarihsiz ``donem_kaydi()`` listenin son elemanini
dondurdugu icin donemi bos satir gercek son donemin (orn. 2026-07) yerine
geciyor ve YANLIS masraf merkezi uretiyordu.

Bu testler dogru davranisi sabitler: donemi DOLU olan son kayit tercih edilir.
"""

from __future__ import annotations

import unittest
from datetime import date

import pandas as pd

from masraf.kayit import PersonelDefteri

SICIL = 900123


def _defter(satirlar: list[dict]) -> PersonelDefteri:
    """Verilen satirlardan bellek ici bir personel defteri kurar."""
    return PersonelDefteri(pd.DataFrame(satirlar))


def _satir(gorev_yeri: str, donem: str | None) -> dict:
    return {
        "Sicil": SICIL,
        "Adı Soyadı": "Test Kisi",
        "Görev Yeri": gorev_yeri,
        "Dönem": donem,
        "Kategori": "Aktif",
        "Şirket 2": "RHI",
    }


class BosDonemTesti(unittest.TestCase):
    """Donemi bos satir bulunan bir sicilin sorgulari."""

    def setUp(self) -> None:
        # Donemi bos satir KASTEN sona konmustur; siralama onu zaten sona
        # atacaktir, sira bagimliligi olmadigi da boylece test edilir.
        self.defter = _defter([
            _satir("GPP Project", "2026-05-01"),
            _satir("GPP Project", "2026-06-01"),
            _satir("Udokan (GMK)", "2026-07-01"),
            _satir("ESKI PROJE", None),
        ])

    def test_sicil_ile_donemi_dolu_son_kaydi_dondurur(self) -> None:
        kayit = self.defter.sicil_ile(str(SICIL))
        self.assertIsNotNone(kayit)
        self.assertEqual(kayit["gorev_yeri"], "Udokan (GMK)")
        self.assertEqual(kayit["donem"], date(2026, 7, 1))

    def test_tarihsiz_gider_donemi_dolu_son_kaydi_kullanir(self) -> None:
        kayit = self.defter.donem_kaydi(str(SICIL), None)
        self.assertIsNotNone(kayit)
        self.assertEqual(kayit["gorev_yeri"], "Udokan (GMK)")
        self.assertEqual(kayit["_donem_eslesme"], "tarihsiz")

    def test_tarihli_gider_dogru_donemi_secmeye_devam_eder(self) -> None:
        """Duzeltme, tarihli normal secimi bozmamalidir."""
        for gun, beklenen in (
            (date(2026, 7, 15), "Udokan (GMK)"),
            (date(2026, 6, 15), "GPP Project"),
            (date(2026, 5, 15), "GPP Project"),
        ):
            with self.subTest(gun=gun):
                kayit = self.defter.donem_kaydi(str(SICIL), gun)
                self.assertEqual(kayit["gorev_yeri"], beklenen)
                self.assertEqual(kayit["_donem_eslesme"], "tam")


class TumDonemlerBosTesti(unittest.TestCase):
    """Hicbir kaydin donemi yoksa sorgular yine de kayit dondurmelidir."""

    def setUp(self) -> None:
        self.defter = _defter([
            _satir("ILK PROJE", None),
            _satir("SON PROJE", None),
        ])

    def test_sicil_ile_cokmez(self) -> None:
        kayit = self.defter.sicil_ile(str(SICIL))
        self.assertIsNotNone(kayit)
        self.assertEqual(kayit["gorev_yeri"], "SON PROJE")

    def test_donem_kaydi_cokmez(self) -> None:
        for tarih in (None, date(2026, 7, 15)):
            with self.subTest(tarih=tarih):
                kayit = self.defter.donem_kaydi(str(SICIL), tarih)
                self.assertIsNotNone(kayit)
                self.assertEqual(kayit["_donem_eslesme"], "tarihsiz")


class DonemiBosSatirYokTesti(unittest.TestCase):
    """Normal veride (donemi bos satir yok) davranis degismemelidir."""

    def test_en_guncel_donem_secilir(self) -> None:
        defter = _defter([
            _satir("GPP Project", "2026-06-01"),
            _satir("Udokan (GMK)", "2026-07-01"),
        ])
        self.assertEqual(defter.sicil_ile(str(SICIL))["gorev_yeri"], "Udokan (GMK)")
        self.assertEqual(
            defter.donem_kaydi(str(SICIL), None)["gorev_yeri"], "Udokan (GMK)"
        )


if __name__ == "__main__":
    unittest.main()
