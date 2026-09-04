"""Gider ayi ile personel donemi eslesmesi testleri.

Mahsuplasmanin temel kurali: bir gider, giderin YAPILDIGI AYDAKI personel
kaydina gore mahsuplasir. Kisinin projesi ve masraf merkezi aydan aya
degisebilir, bu yuzden "en guncel kayit" degil "o ayki kayit" kullanilir.

Ele alinan dort durum:

``tam``
    Gider ayi ile donem ayi ayni. Normal durum.
``onceki_donem``
    Kisinin gider ayinda kaydi yok; daha eski bir donem kullanildi. Pratikte
    kisi o tarihten once ayrilmistir ve kullanilan donem CIKIS AYIDIR. Cikis
    bileti gibi masraflar dogru sekilde cikis santiyesine gider.
``ilk_donem_oncesi``
    Gider, kisinin ilk kaydindan once yapilmis. Kisi henuz ise baslamamis;
    mobilizasyon veya aday seyahati.
``tarihsiz``
    Gider satirinda tarih yok.
"""

from __future__ import annotations

import datetime
import unittest
from pathlib import Path

try:
    from masraf.kayit import PersonelDefteri
    from masraf.masraf_merkezi import MasrafMerkeziHaritasi, masraf_merkezi_coz
    from masraf.modeller import DURUM_INCELE, Eslesme, GiderSatiri

    MODUL_VAR = True
except ImportError:
    MODUL_VAR = False

PERSONEL = Path("ornek_veri/personel/2025_2026_giris_cikis.xlsx")
HARITA = Path("veri/masraf_merkezi_haritasi.csv")


def _satir(tarih: datetime.date | None, kisi: str = "test") -> "GiderSatiri":
    return GiderSatiri(
        kaynak_dosya="test", kaynak_tip="antik_cari", satir_no=1,
        belge_tarihi=tarih, aciklama=kisi, kisi_ham=kisi, sicil_ham=None,
        tckn_ham=None, tutar=500.0, para_birimi="USD",
        masraf_merkezi_kaynak=None, gider_tipi="Bilet",
    )


def _eslesme(sicil: str, ad: str = "Test Kisi") -> "Eslesme":
    return Eslesme(sicil=sicil, ad_soyad=ad, yontem="sicil", guven=1.0,
                   aday_sayisi=1, aciklama="test")


@unittest.skipUnless(MODUL_VAR, "masraf modulleri bulunamadi")
class DonemEslesmesiTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not PERSONEL.is_file():
            raise unittest.SkipTest("ornek_veri/personel bulunamadi")
        cls.defter = PersonelDefteri.yukle(PERSONEL)
        cls.harita = (MasrafMerkeziHaritasi.yukle(HARITA) if HARITA.is_file()
                      else MasrafMerkeziHaritasi.yukle(Path("veri") / "yok.csv"))
        # Butun donemlerde kaydi olan, ayrilmamis bir kisi: Ozakay Mustafa Kemal
        cls.aktif_sicil = "100003"

    # ---------------- donem_kaydi seviyesi ----------------

    def test_ayni_ay_tam_isaretlenir(self):
        k = self.defter.donem_kaydi(self.aktif_sicil, datetime.date(2026, 7, 15))
        self.assertEqual(k["_donem_eslesme"], "tam")
        self.assertEqual(k["donem"].year, 2026)
        self.assertEqual(k["donem"].month, 7)
        self.assertFalse(k["_donem_tahmini"])

    def test_her_ay_kendi_donemine_duser(self):
        """Ocak faturasi Ocak donemine, Haziran faturasi Haziran donemine."""
        for ay in (11, 12):
            k = self.defter.donem_kaydi(self.aktif_sicil, datetime.date(2025, ay, 20))
            self.assertEqual((k["donem"].year, k["donem"].month), (2025, ay))
            self.assertEqual(k["_donem_eslesme"], "tam")
        for ay in range(1, 8):
            k = self.defter.donem_kaydi(self.aktif_sicil, datetime.date(2026, ay, 20))
            self.assertEqual((k["donem"].year, k["donem"].month), (2026, ay))
            self.assertEqual(k["_donem_eslesme"], "tam")

    def test_ilk_donemden_once_isaretlenir(self):
        k = self.defter.donem_kaydi(self.aktif_sicil, datetime.date(2020, 1, 1))
        self.assertEqual(k["_donem_eslesme"], "ilk_donem_oncesi")
        self.assertTrue(k["_donem_tahmini"])

    def test_son_donemden_sonra_onceki_donem_olur(self):
        """Veri araligindan sonraki bir tarih son donemi kullanir ve isaretlenir."""
        k = self.defter.donem_kaydi(self.aktif_sicil, datetime.date(2030, 1, 1))
        self.assertEqual(k["_donem_eslesme"], "onceki_donem")
        self.assertTrue(k["_donem_tahmini"])

    def test_tarihsiz_isaretlenir(self):
        k = self.defter.donem_kaydi(self.aktif_sicil, None)
        self.assertEqual(k["_donem_eslesme"], "tarihsiz")

    # ---------------- masraf_merkezi_coz seviyesi ----------------

    def test_ayrilmis_kisi_cikis_donemine_baglanir(self):
        """Cikis sonrasi bir gider, kisinin CIKIS AYINDAKI santiyesine gider.

        Gercek ornek: Tkacheva Natalia (C92620) 27.11.2025'te ayrilmis.
        Temmuz 2026 dokumunde bir satiri var. Donem 2025-11 olmali ve satir
        uyari ile incelemeye alinmali.
        """
        kayit = self.defter.sicil_ile("C92620")
        if kayit is None:
            self.skipTest("C92620 ornek veride yok")
        s = _satir(datetime.date(2026, 7, 28), "Tkacheva Natalia")
        sonuc = masraf_merkezi_coz(s, _eslesme("C92620"), self.defter, self.harita)
        self.assertEqual(sonuc.donem_eslesme, "onceki_donem")
        self.assertEqual(sonuc.donem.year, 2025)
        self.assertEqual(sonuc.donem.month, 11)
        self.assertEqual(sonuc.durum, DURUM_INCELE)
        self.assertTrue(sonuc.uyarilar)
        birlesik = " ".join(sonuc.uyarilar).lower()
        self.assertIn("temmuz 2026", birlesik, "gider ayi uyarida yazmali")
        self.assertIn("kasim 2025", birlesik, "kullanilan donem uyarida yazmali")

    def test_ayni_ay_uyari_uretmez(self):
        s = _satir(datetime.date(2026, 7, 15), "Ozakay Mustafa Kemal")
        sonuc = masraf_merkezi_coz(s, _eslesme(self.aktif_sicil), self.defter, self.harita)
        self.assertEqual(sonuc.donem_eslesme, "tam")
        donem_uyarisi = [u for u in sonuc.uyarilar if "donem" in u.lower() or "ay" in u.lower()]
        self.assertFalse(donem_uyarisi, f"ayni ayda gereksiz uyari: {donem_uyarisi}")

    def test_ise_girmeden_once_uyari_verir(self):
        s = _satir(datetime.date(2020, 3, 10), "Ozakay Mustafa Kemal")
        sonuc = masraf_merkezi_coz(s, _eslesme(self.aktif_sicil), self.defter, self.harita)
        self.assertEqual(sonuc.donem_eslesme, "ilk_donem_oncesi")
        self.assertEqual(sonuc.durum, DURUM_INCELE)
        self.assertTrue(any("baslamamis" in u.lower() or "once" in u.lower()
                            for u in sonuc.uyarilar))

    def test_gorev_yeri_degisen_kisi_dogru_aya_baglanir(self):
        """Donemler arasinda gorev yeri degistiren bir kisi bulunur ve
        her iki ayda da o ayin gorev yerine baglandigi dogrulanir.

        Bu test projenin varlik sebebini korur: kisilerin yuzde 1,4'unun
        gorev yeri degisir ve tam o kisiler yanlis mahsuplasmaya yol acar.
        """
        aday = None
        for sicil in list(getattr(self.defter, "_isim_index", {}).values())[:4000]:
            for sic in sicil:
                yerler = set()
                for ay in (11, 12):
                    k = self.defter.donem_kaydi(sic, datetime.date(2025, ay, 15))
                    if k and k.get("_donem_eslesme") == "tam":
                        yerler.add(k.get("gorev_yeri"))
                for ay in range(1, 8):
                    k = self.defter.donem_kaydi(sic, datetime.date(2026, ay, 15))
                    if k and k.get("_donem_eslesme") == "tam":
                        yerler.add(k.get("gorev_yeri"))
                if len([y for y in yerler if y]) > 1:
                    aday = sic
                    break
            if aday:
                break
        if aday is None:
            self.skipTest("gorev yeri degistiren kisi bulunamadi")

        gorulen = {}
        for yil, ay in [(2025, 11), (2025, 12)] + [(2026, a) for a in range(1, 8)]:
            k = self.defter.donem_kaydi(aday, datetime.date(yil, ay, 15))
            if k and k.get("_donem_eslesme") == "tam":
                gorulen[(yil, ay)] = k.get("gorev_yeri")
        self.assertGreater(len(set(gorulen.values())), 1,
                           "bu kisi gercekten gorev yeri degistirmis olmali")
        # Her ay icin cozulen masraf merkezi o ayin gorev yerinden gelmeli
        for (yil, ay), yer in gorulen.items():
            s = _satir(datetime.date(yil, ay, 15))
            sonuc = masraf_merkezi_coz(s, _eslesme(aday), self.defter, self.harita)
            self.assertEqual(sonuc.gorev_yeri, yer,
                             f"{yil}-{ay:02d}: o ayin gorev yeri kullanilmadi")


if __name__ == "__main__":
    unittest.main(verbosity=2)
