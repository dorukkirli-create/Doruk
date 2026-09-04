"""TC kimlik koprusu ve dogum tarihi ile dogrulanmis alias turetme testleri.

Altin ornekler gercek Temmuz 2026 verisinden alinmistir. Dort tanesi
dogum tarihi kontrolu olmadan YANLIS eslesen vakalardir; testler bunlarin
uretilmedigini dogrular.
"""

from __future__ import annotations

import datetime
import unittest
from dataclasses import dataclass, field
from pathlib import Path

try:
    from masraf.kopru import (
        KopruAdayi,
        alias_turet,
        aliaslari_deftere_yaz,
        kopru_ozeti,
        kopru_turet,
        kopruyu_deftere_yaz,
    )

    MODUL_VAR = True
except ImportError:
    MODUL_VAR = False

PERSONEL = Path("ornek_veri/personel/2025_2026_giris_cikis.xlsx")
SAGLIK = Path("ornek_veri/energo/SAGLIK_KONTROL_LISTE.xlsx")


@dataclass
class SahteSatir:
    """GiderSatiri'nin kopru testleri icin yeterli olan sade taklidi."""

    kisi_ham: str | None = None
    tckn_ham: str | None = None
    kaynak_tip: str = "test"
    ek: dict = field(default_factory=dict)


class SahteDefter:
    """Kucuk, elle kurulmus bir personel defteri taklidi."""

    def __init__(self, kayitlar: dict[str, dict]):
        self._k = kayitlar
        self._isim = {}
        for sicil, kayit in kayitlar.items():
            self._isim.setdefault(kayit["ad_soyad_norm"], []).append(sicil)
        self._isim_index = self._isim  # kopru._tum_siciller bunu kullanir

    def sicil_ile(self, sicil):
        return self._k.get(sicil)

    def isimle_adaylar(self, norm):
        return list(self._isim.get(norm, ()))

    def token_ile_adaylar(self, tokenlar):
        return [s for s, k in self._k.items()
                if frozenset(k["ad_soyad_norm"].split()) == tokenlar]

    def soyad_ile_adaylar(self, soyad):
        return [s for s, k in self._k.items()
                if k["ad_soyad_norm"].split()[0] == soyad]


@unittest.skipUnless(MODUL_VAR, "masraf.kopru bulunamadi")
class KopruTuretmeTest(unittest.TestCase):
    """Dogum tarihi ile daraltmanin dogru calistigini sahte veriyle dogrular."""

    def setUp(self):
        self.defter = SahteDefter({
            "100": {"ad_soyad": "Yilmaz Ahmet", "ad_soyad_norm": "YILMAZ AHMET",
                    "dogum_tarihi": datetime.date(1980, 1, 1)},
            "200": {"ad_soyad": "Yilmaz Ahmet", "ad_soyad_norm": "YILMAZ AHMET",
                    "dogum_tarihi": datetime.date(1990, 5, 5)},
            "300": {"ad_soyad": "Kaya Veli", "ad_soyad_norm": "KAYA VELI",
                    "dogum_tarihi": datetime.date(1975, 3, 3)},
        })

    def test_dogum_tarihi_cakisan_ismi_ayirir(self):
        s = SahteSatir(kisi_ham="Ahmet Yilmaz", tckn_ham="11111111111",
                       ek={"dogum_tarihi": datetime.date(1990, 5, 5)})
        a = kopru_turet([s], self.defter)
        self.assertEqual(len(a), 1)
        self.assertEqual(a[0].sicil, "200")
        self.assertEqual(a[0].yontem, "isim_dogum")

    def test_dogum_tarihi_tutmuyorsa_kopru_kurulmaz(self):
        s = SahteSatir(kisi_ham="Ahmet Yilmaz", tckn_ham="11111111111",
                       ek={"dogum_tarihi": datetime.date(2000, 12, 12)})
        self.assertEqual(kopru_turet([s], self.defter), [])

    def test_dogum_yoksa_varsayilan_olarak_atlanir(self):
        s = SahteSatir(kisi_ham="Veli Kaya", tckn_ham="22222222222")
        self.assertEqual(kopru_turet([s], self.defter), [])

    def test_dogum_zorunlu_kapaliyken_tek_aday_kabul_edilir(self):
        s = SahteSatir(kisi_ham="Veli Kaya", tckn_ham="22222222222")
        a = kopru_turet([s], self.defter, dogum_zorunlu=False)
        self.assertEqual(len(a), 1)
        self.assertEqual(a[0].sicil, "300")
        self.assertLess(a[0].guven, 0.8, "dogum tarihi yokken guven dusuk olmali")

    def test_ayni_tckn_bir_kez_dondurulur(self):
        s = SahteSatir(kisi_ham="Veli Kaya", tckn_ham="22222222222",
                       ek={"dogum_tarihi": datetime.date(1975, 3, 3)})
        self.assertEqual(len(kopru_turet([s, s, s], self.defter)), 1)

    def test_tckn_yoksa_atlanir(self):
        s = SahteSatir(kisi_ham="Veli Kaya",
                       ek={"dogum_tarihi": datetime.date(1975, 3, 3)})
        self.assertEqual(kopru_turet([s], self.defter), [])


@unittest.skipUnless(MODUL_VAR, "masraf.kopru bulunamadi")
class GercekVeriTest(unittest.TestCase):
    """Gercek personel verisi ve saglik listesiyle altin ornekleri dogrular."""

    @classmethod
    def setUpClass(cls):
        if not (PERSONEL.is_file() and SAGLIK.is_file()):
            raise unittest.SkipTest("ornek_veri bulunamadi")
        from masraf.kayit import PersonelDefteri
        from masraf.okuyucular.kesif import oku
        cls.defter = PersonelDefteri.yukle(PERSONEL)
        cls.satirlar = oku(SAGLIK)

    def test_tckn_koprusu_kurulur(self):
        a = kopru_turet(self.satirlar, self.defter)
        self.assertGreaterEqual(len(a), 20, "beklenen kopru sayisina ulasilamadi")
        for k in a:
            self.assertEqual(len(k.tckn), 11)
            self.assertTrue(k.sicil)
            self.assertGreater(k.guven, 0.9)

    def test_dogru_eslesmeler_uretilir(self):
        """Soyadi farkli yazilmis ama dogum tarihi tutan kisiler bulunmali."""
        uretilen = {a.ad_soyad_kaynak: a.sicil for a in alias_turet(self.satirlar, self.defter)}
        beklenen = {
            "ŞEHRİBAN ÖZKAN": "D20296",     # personelde Seriban Ozkan
            "YAŞAR MERT YANAR": "642586",   # personelde Banar Yasar Mert
            "MEHMET DALKILIÇ": "490707",    # personelde Dalkilinc Mehmet
        }
        for ad, sicil in beklenen.items():
            self.assertIn(ad, uretilen, f"{ad} icin alias uretilmedi")
            self.assertEqual(uretilen[ad], sicil, f"{ad} yanlis sicile baglandi")

    def test_yanlis_eslesmeler_uretilmez(self):
        """Isim benzer ama dogum tarihi farkli olanlar reddedilmeli.

        Bu dort vaka dogum tarihi kontrolu olmadan yanlis eslesiyordu.
        """
        uretilen = {a.ad_soyad_kaynak for a in alias_turet(self.satirlar, self.defter)}
        for ad in ("GÖKHAN GÜZEL", "MEHMET EKREM NERGİZ", "BARIŞ GÖÇEDEN", "OGÜN BİZ"):
            self.assertNotIn(ad, uretilen,
                             f"{ad} icin yanlis alias uretildi, dogum tarihi tutmuyor")

    def test_ozet_tutarli(self):
        a = kopru_turet(self.satirlar, self.defter)
        o = kopru_ozeti(self.satirlar, self.defter, a)
        self.assertEqual(o["tckn_tasiyan_kisi"],
                         o["isimle_bulunan"] + o["personel_verisinde_yok"])
        self.assertLessEqual(o["kopru_kurulan"], o["isimle_bulunan"])

    def test_alias_turetme_hizli(self):
        import time
        t = time.time()
        alias_turet(self.satirlar, self.defter)
        self.assertLess(time.time() - t, 5.0, "alias turetme cok yavas")


if __name__ == "__main__":
    unittest.main(verbosity=2)
