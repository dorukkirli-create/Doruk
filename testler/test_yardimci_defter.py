"""1C personel listesi (grup sirketleri defteri) testleri.

Ana veri sadece RHI ve UST LUGA tuzel kisilerini kapsar. Renservis,
Renstroydetal, RC, One Tower, Top Tower personeli ancak 1C listesinde bulunur.

Altin ornekler gercek Temmuz 2026 faturasindan alinmistir. Bunlar once
hicbir yerde bulunamiyordu; artik sirketleri ve projeleriyle cozuluyor.
"""

from __future__ import annotations

import unittest
from pathlib import Path

try:
    from masraf.metin import isim_normalize
    from masraf.yardimci_defter import YardimciDefter

    MODUL_VAR = True
except ImportError:
    MODUL_VAR = False

LISTE = Path("ornek_veri/personel/1C_Personnel_List_31082026.xlsx")

#: (aranan isim, beklenen sicil, beklenen sirket, projede gecmesi beklenen metin)
ALTIN_ORNEKLER = (
    ("Celer Ahmet", "534561", "RENSTROYDETAL", "Renstroydetal"),
    ("Boynuegri Vedat", "101074", "RENSERVIS", "Renservis"),
    ("Erdur Koray", "408755", "RC", "Top Tower"),
    ("Menetlioglu Gokhan", "404389", "RC", "One Tower"),
    ("Ozcan Mustafa", "475837", "RENSTROYDETAL", "Lytkarino"),
    ("Surul Tolga", "644705", "RC", "One Tower"),
    ("Gundogdu Ali", "643032", "RENSTROYDETAL", "Renstroydetal"),
    ("Kaiyrbekov Azat", "442829", "RENSERVIS", "Renservis"),
)


@unittest.skipUnless(MODUL_VAR, "masraf.yardimci_defter bulunamadi")
class YardimciDefterTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not LISTE.is_file():
            raise unittest.SkipTest("1C personel listesi ornek_veri'de yok")
        cls.defter = YardimciDefter.yukle(LISTE)

    def test_grup_sirketleri_kapsaniyor(self):
        ist = self.defter.istatistik()
        self.assertGreater(ist["benzersiz_sicil"], 30000)
        self.assertGreater(ist["benzersiz_isim"], 15000)
        sirketler = set(ist["sirket_dagilimi"])
        for beklenen in ("RSS", "RC", "RHI", "UST LUGA"):
            self.assertIn(beklenen, sirketler, f"{beklenen} kapsanmiyor")

    def test_altin_ornekler_bulunuyor(self):
        for ad, sicil, sirket, proje_parcasi in ALTIN_ORNEKLER:
            with self.subTest(kisi=ad):
                adaylar = self.defter.isimle_adaylar(isim_normalize(ad))
                self.assertIn(sicil, adaylar, f"{ad} icin {sicil} bulunamadi")
                kayit = self.defter.sicil_ile(sicil)
                self.assertIsNotNone(kayit)
                self.assertEqual(kayit["sirket"], sirket)
                self.assertIn(proje_parcasi.lower(), (kayit["gorev_yeri"] or "").lower())

    def test_sicil_normalize_ediliyor(self):
        """Sicil float okunsa bile ('534561.0') ayni kayda ulasilmali."""
        self.assertIsNotNone(self.defter.sicil_ile("534561"))
        self.assertIsNotNone(self.defter.sicil_ile("534561.0"))
        self.assertIsNotNone(self.defter.sicil_ile(" 534561 "))

    def test_donem_kaydi_tek_snapshot_isaretler(self):
        """1C listesi tek tarihlidir; donem dogrulanamaz olarak isaretlenmeli."""
        import datetime

        kayit = self.defter.donem_kaydi("534561", datetime.date(2026, 7, 15))
        self.assertIsNotNone(kayit)
        self.assertEqual(kayit["_donem_eslesme"], "yardimci_defter")
        self.assertTrue(kayit["_donem_tahmini"])

    def test_olmayan_sicil_none_doner(self):
        self.assertIsNone(self.defter.sicil_ile("bu-sicil-yok"))
        self.assertIsNone(self.defter.donem_kaydi("bu-sicil-yok", None))

    def test_excluded_kategorisi_isim_indeksine_girmez(self):
        """'Excluded' kategorisindeki satirlar gercek kisi degildir."""
        for kayit in list(self.defter._sicil.values())[:5000]:
            if (kayit.get("kategori_1c") or "").strip().lower() == "excluded":
                ad = kayit.get("ad_soyad")
                if ad:
                    self.assertNotIn(
                        kayit["sicil"],
                        self.defter.isimle_adaylar(isim_normalize(ad)),
                        "Excluded kayit isim indeksine girmis",
                    )

    def test_olmayan_dosya_hata_verir(self):
        with self.assertRaises(FileNotFoundError):
            YardimciDefter.yukle("olmayan_liste.xlsx")


@unittest.skipUnless(MODUL_VAR, "masraf.yardimci_defter bulunamadi")
class BoruEntegrasyonTest(unittest.TestCase):
    """1C listesi boru hattina baglandiginda kapsam artmali."""

    ANA = Path("ornek_veri/personel/2025_2026_giris_cikis.xlsx")
    MESAJ = Path("ornek_veri/posta/ornek_mail.msg")

    @classmethod
    def setUpClass(cls):
        if not (cls.ANA.is_file() and LISTE.is_file() and cls.MESAJ.is_file()):
            raise unittest.SkipTest("ornek veri eksik")

    def test_1c_ile_kapsam_artiyor(self):
        from masraf.boru import Boru, CalismaAyarlari

        def olc(yardimci):
            a = CalismaAyarlari(
                personel_yolu=self.ANA, yardimci_personel_yolu=yardimci,
                veri_dizini="veri", cikti_dizini="cikti",
                defterleri_besle=False, ogrenmeyi_kaydet=False,
            )
            b = Boru(a)
            sonuclar = b.isle([self.MESAJ])
            return sum(1 for s in sonuclar if s.masraf_merkezi), len(sonuclar)

        onceki, toplam = olc(None)
        sonraki, toplam2 = olc(LISTE)
        self.assertEqual(toplam, toplam2, "satir sayisi degismemeli")
        self.assertGreater(
            sonraki, onceki,
            f"1C listesi kapsami artirmali ({onceki} -> {sonraki})",
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
