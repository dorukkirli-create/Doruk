"""Harita onerisi: tanimsiz gorev yerini 1C 'Firm 2' ile sirkete baglar.

Bir gorev yeri masraf merkezi haritasinda yoksa kod onu metin olarak tasir ve
isaretler. Operatorun o projeyi elle arastirip tuzel kisisini bulmasi
gerekirdi. Oysa bilgi 1C listesinin 'Firm 2' kolonunda zaten duruyor ve
haritanin kendi 'sirket' kolonuyla AYNI sozlugu kullaniyor (RHI, UST LUGA,
RSS, RC, BSK). Bu modul o baglantiyi kurar.
"""

from __future__ import annotations

import unittest
from pathlib import Path

from masraf.harita_onerisi import kod_uret, oneri_uret


class KodUretTest(unittest.TestCase):
    def test_okunakli_kod(self):
        self.assertEqual(kod_uret("One Tower"), "ONE-TOWER")

    def test_uzun_ad_kisaltilir(self):
        kod = kod_uret("Bsk Management Group")
        self.assertIn("BSK", kod)
        self.assertLessEqual(len(kod), 24)

    def test_turkce_karakterler_asciye_iner(self):
        kod = kod_uret("Ölçme Değerlendirme Şirketi")
        self.assertTrue(all(ord(c) < 128 for c in kod), kod)

    def test_proje_kelimesi_atilir(self):
        self.assertNotIn("PROJECT", kod_uret("GPP Project"))

    def test_cakisan_kod_benzersizlestirilir(self):
        kullanilan = {"ONE-TOWER"}
        self.assertNotEqual(kod_uret("One Tower", kullanilan), "ONE-TOWER")

    def test_bos_girdi_cokmez(self):
        self.assertTrue(kod_uret(""))


class SahteHarita:
    def __init__(self, tanimlilar=()):
        self._t = {t.upper() for t in tanimlilar}

    def coz(self, yer):
        return {"masraf_merkezi_kodu": "X"} if str(yer).upper() in self._t else None

    def kod_adlari(self):
        return {}


class SahteDefter:
    def __init__(self, kayitlar):
        self._kayitlar = kayitlar


class OneriUretTest(unittest.TestCase):
    def _sonuc(self, gorev_yeri, tutar, haritada):
        from masraf.modeller import Eslesme, GiderSatiri, Sonuc

        satir = GiderSatiri(
            kaynak_dosya="f.xlsx", kaynak_tip="genel", satir_no=1,
            belge_tarihi=None, aciklama="", kisi_ham="A B", sicil_ham=None,
            tckn_ham=None, tutar=tutar, para_birimi="USD",
            masraf_merkezi_kaynak=None, gider_tipi="Bilet",
            ek={"masraf_merkezi_haritada": haritada},
        )
        return Sonuc(
            satir=satir,
            eslesme=Eslesme(sicil="1", ad_soyad="A B", yontem="sicil", guven=1.0,
                            aday_sayisi=1, aciklama=""),
            donem=None, gorev_yeri=gorev_yeri, masraf_merkezi=gorev_yeri,
            sirket=None, sirket2=None, statu=None, kategori=None,
            cikis_tarihi=None, durum="OTOMATIK",
        )

    def test_firm2_sirket_olarak_kullanilir(self):
        defter = SahteDefter([
            {"gorev_yeri": "One Tower", "sirket": "KASIM KOS", "sirket2": "RC"},
            {"gorev_yeri": "One Tower", "sirket": "LLC X", "sirket2": "RC"},
        ])
        oneriler = oneri_uret([self._sonuc("One Tower", 500.0, False)],
                              SahteHarita(), yardimci=defter)
        self.assertEqual(len(oneriler), 1)
        self.assertEqual(oneriler[0].sirket, "RC", "Firm 2 (sirket2) kullanilmali")
        self.assertEqual(oneriler[0].kisi_sayisi, 2)

    def test_haritada_tanimli_olan_onerilmez(self):
        oneriler = oneri_uret([self._sonuc("GPP Project", 100.0, True)], SahteHarita())
        self.assertEqual(oneriler, [])

    def test_tutar_buyukten_kucuge_sirali(self):
        s = [self._sonuc("A Projesi", 100.0, False),
             self._sonuc("B Projesi", 900.0, False)]
        oneriler = oneri_uret(s, SahteHarita())
        self.assertEqual([o.gorev_yeri for o in oneriler], ["B Projesi", "A Projesi"])

    def test_sirket_bulunamazsa_isaretlenir(self):
        oneriler = oneri_uret([self._sonuc("Bilinmeyen Proje", 10.0, False)],
                              SahteHarita())
        self.assertEqual(oneriler[0].sirket, "")
        self.assertIn("bulunamadi", oneriler[0].kaynak)

    def test_csv_satiri_haritaya_uygun(self):
        defter = SahteDefter([{"gorev_yeri": "Top Tower", "sirket2": "RC"}])
        o = oneri_uret([self._sonuc("Top Tower", 50.0, False)],
                       SahteHarita(), yardimci=defter)[0]
        parcalar = o.csv_satiri().split(",")
        self.assertEqual(len(parcalar), 5, "harita 5 kolonlu olmali")
        self.assertEqual(parcalar[0], "Top Tower")
        self.assertEqual(parcalar[3], "RC")
        self.assertEqual(parcalar[4], "E")


class GercekVeriTest(unittest.TestCase):
    """Gercek 1C listesiyle: dokuz grup sirketi projesi haritada olmali."""

    HARITA = Path("veri/masraf_merkezi_haritasi.csv")
    GRUP_PROJELERI = {
        "Renstroydetal - Ust-Luga GPC": "RSS",
        "Renservis - Lytkarino - Renservis": "RSS",
        "Renstroydetal - Lytkarino": "RSS",
        "One Tower": "RC",
        "Top Tower": "RC",
        "Icity Business Center": "RC",
        "Lakhta 2 Fit-out": "RC",
        "Bsk Management Group": "BSK",
    }

    @classmethod
    def setUpClass(cls):
        if not cls.HARITA.is_file():
            raise unittest.SkipTest("harita dosyasi yok")

    def test_grup_sirketi_projeleri_haritada_ve_dogru_sirkette(self):
        from masraf.masraf_merkezi import MasrafMerkeziHaritasi

        harita = MasrafMerkeziHaritasi.yukle(self.HARITA)
        for proje, beklenen in self.GRUP_PROJELERI.items():
            with self.subTest(proje=proje):
                cozum = harita.coz(proje)
                self.assertIsNotNone(cozum, f"{proje} haritada yok")
                self.assertEqual(cozum["sirket"], beklenen)


if __name__ == "__main__":
    unittest.main(verbosity=2)
