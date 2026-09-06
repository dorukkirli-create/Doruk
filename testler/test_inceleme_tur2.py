"""Ikinci tur bagimsiz incelemede olculen bulgularin gerileme testleri.

Kimlik eslestirme (bulanik miknatis, 1C sirasi, alias es isimli, alt kume
soyad, ters alt kume, kesik bitisik ad, kanit kaynagi) ve donem / masraf
merkezi (NaT donem, personel dosyasi eski, ilk donem mesaji, 1C DIKKAT
karsilastirmasi, paylasim etiketi kanonik sirket, harita yukleme).
"""

from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

import pandas as pd

from masraf import eslestirici as E
from masraf import masraf_merkezi as MM
from masraf.kayit import PersonelDefteri, _tarihe_cevir
from masraf.modeller import DURUM_INCELE, YONTEMLER, Eslesme, GiderSatiri
from testler.altin import altin, altin_veya_none

KOK = Path(__file__).resolve().parents[1]
PERSONEL = KOK / "ornek_veri" / "personel" / "2025_2026_giris_cikis.xlsx"
YARDIMCI = KOK / "ornek_veri" / "personel" / "1C_Personnel_List_31082026.xlsx"


def _gider(kisi, tip="Bilet", tckn=None, kaynak_tip="antik_cari", tarih=date(2026, 7, 15)):
    return GiderSatiri(kaynak_dosya="t.xls", kaynak_tip=kaynak_tip, satir_no=1, belge_tarihi=tarih,
                       aciklama=kisi or "", kisi_ham=kisi, sicil_ham=None, tckn_ham=tckn, tutar=10.0,
                       para_birimi="USD", masraf_merkezi_kaynak=None, gider_tipi=tip)


class SabitlerTest(unittest.TestCase):
    def test_incele_esigi_uretim_esigiyle_ayni(self):
        self.assertEqual(E.INCELE_ESIGI, MM.GUVEN_ESIGI)

    def test_bulanik_tavani_otomatik_esiginin_altinda(self):
        self.assertLess(E.BULANIK_TAVAN, MM.GUVEN_ESIGI)

    def test_yardimci_defter_yontem_sozlesmede(self):
        self.assertIn("yardimci_defter", YONTEMLER)


class TariheCevirTest(unittest.TestCase):
    def test_nat_none_olur(self):
        self.assertIsNone(_tarihe_cevir(pd.NaT))
        self.assertIsNone(_tarihe_cevir(float("nan")))

    def test_bos_donem_hucresi_defteri_cokertmez(self):
        df = pd.DataFrame([
            {"Sicil": 1, "Adı Soyadı": "Test Kisi", "Görev Yeri": "GPP Project",
             "Dönem": pd.Timestamp("2026-06-01"), "Kategori": "Aktif", "Şirket 2": "RHI"},
            {"Sicil": 1, "Adı Soyadı": "Test Kisi", "Görev Yeri": "ESKI", "Dönem": pd.NaT,
             "Kategori": "Aktif", "Şirket 2": "RHI"},
            {"Sicil": 2, "Adı Soyadı": "Tek Satir", "Görev Yeri": "GPP Project", "Dönem": pd.NaT,
             "Kategori": "Aktif", "Şirket 2": "RHI"},
        ])
        defter = PersonelDefteri(df)
        self.assertEqual(defter.sicil_ile("1")["gorev_yeri"], "GPP Project")
        k = defter.donem_kaydi("1", date(2026, 7, 15))
        self.assertEqual(k["gorev_yeri"], "GPP Project")
        self.assertIsNotNone(defter.donem_kaydi("2", date(2026, 7, 15)))


class HaritaYardimcilariTest(unittest.TestCase):
    def test_en_dash_anahtar(self):
        self.assertEqual(MM._anahtar("Ust-Luga–Reshetnikova Office"), "UST LUGA RESHETNIKOVA OFFICE")
        self.assertEqual(MM._anahtar("Ust Luga - Reshetnikova Office"), "UST LUGA RESHETNIKOVA OFFICE")

    def test_tuzel_kisi_tam_token(self):
        h = MM.MasrafMerkeziHaritasi([])
        for e in ("RHI", "RSS", "BSK", "RHI 1/3 - RENSTROYDETAL 2/3", "ONE TOWER 1/2 - RC 1/2", "YAKA LLC"):
            self.assertTrue(h.tuzel_kisi_mi(e), e)
        for e in ("R", "U", "SA", "TOP", "GPP Project", "Ust Luga GPP", "RSS PROJESI X"):
            self.assertFalse(h.tuzel_kisi_mi(e), e)

    def test_sirket_kanonik(self):
        self.assertEqual(MM.sirket_kanonik("RENSTROYDETAL"), "RSS")
        self.assertEqual(MM.sirket_kanonik("Renservis"), "RSS")
        self.assertEqual(MM.sirket_kanonik("One Tower"), "RC")
        self.assertEqual(MM.sirket_kanonik("ust luga"), "UST LUGA")
        self.assertIsNone(MM.sirket_kanonik("GPP Project"))

    def test_cp1254_harita_okunur_ve_uyarir(self):
        with tempfile.TemporaryDirectory() as d:
            yol = Path(d) / "h.csv"
            yol.write_bytes("gorev_yeri,masraf_merkezi_kodu,masraf_merkezi_adi,sirket,aktif\nŞantiye Üç,SU3,Santiye Uc,rhi,E\n".encode("cp1254"))
            h = MM.MasrafMerkeziHaritasi.yukle(yol)
        self.assertIsNotNone(h.coz("Şantiye Üç"))
        self.assertEqual(h.coz("Şantiye Üç")["sirket"], "RHI")
        self.assertIn("Windows-1254", h.yukleme_uyarisi or "")

    def test_ayni_gorev_yeri_iki_satir_uyarir_ve_ilk_kazanir(self):
        with tempfile.TemporaryDirectory() as d:
            yol = Path(d) / "h.csv"
            yol.write_text("gorev_yeri,masraf_merkezi_kodu,masraf_merkezi_adi,sirket,aktif\n"
                           "GPP Project,ESKI,Eski,RHI,H\nGPP Project,YENI,Yeni,RHI,E\n", encoding="utf-8")
            h = MM.MasrafMerkeziHaritasi.yukle(yol)
        self.assertEqual(h.coz("GPP Project")["masraf_merkezi_kodu"], "ESKI")
        self.assertIn("birden fazla satirda", h.yukleme_uyarisi or "")

    def test_kod_gorev_yeriyle_cakisinca_gorev_yeri_kazanir(self):
        with tempfile.TemporaryDirectory() as d:
            yol = Path(d) / "h.csv"
            yol.write_text("gorev_yeri,masraf_merkezi_kodu,masraf_merkezi_adi,sirket,aktif\n"
                           "Udokan (GMK),Udokan,Udokan GMK,RHI,E\nUdokan,UDK-CAMP,Udokan Kamp,RHI,E\n", encoding="utf-8")
            h = MM.MasrafMerkeziHaritasi.yukle(yol)
        self.assertEqual(h.coz("Udokan")["masraf_merkezi_kodu"], "UDK-CAMP")
        self.assertIn("cakisiyor", h.yukleme_uyarisi or "")


class DonemMesajlariTest(unittest.TestCase):
    """Sentetik defter: personel dosyasi eski / ilk donemden once mesajlari."""

    def setUp(self):
        self.defter = PersonelDefteri(pd.DataFrame([
            {"Sicil": 7, "Adı Soyadı": "Test Kisi", "Görev Yeri": "GPP Project", "Dönem": pd.Timestamp("2026-05-01"),
             "Kategori": "Aktif", "Şirket 2": "RHI", "RHI İşe Giriş Tarihi": pd.Timestamp("2015-03-01")},
            {"Sicil": 7, "Adı Soyadı": "Test Kisi", "Görev Yeri": "GPP Project", "Dönem": pd.Timestamp("2026-06-01"),
             "Kategori": "Aktif", "Şirket 2": "RHI", "RHI İşe Giriş Tarihi": pd.Timestamp("2015-03-01")},
        ]))
        self.harita = MM.MasrafMerkeziHaritasi([MM.MasrafMerkezi(gorev_yeri="GPP Project", kod="GPP", ad="GPP Project", sirket="UST LUGA", aktif=True)])
        self.eslesme = Eslesme(sicil="7", ad_soyad="Test Kisi", yontem="tam_isim", guven=0.95, aday_sayisi=1,
                               aciklama="", aday_siciller=["7"])

    def test_personel_dosyasi_eski(self):
        s = MM.masraf_merkezi_coz(_gider("TEST KISI", tarih=date(2026, 8, 15)), self.eslesme, self.defter,
                                  self.harita, son_donem=date(2026, 6, 1))
        self.assertEqual(s.donem_eslesme, "personel_dosyasi_eski")
        self.assertEqual(s.durum, DURUM_INCELE)
        metin = " ".join(s.uyarilar).lower()
        self.assertIn("guncelleyip", metin)
        self.assertNotIn("cikis masrafi", metin)

    def test_aktif_kisi_kayit_kesilmis(self):
        s = MM.masraf_merkezi_coz(_gider("TEST KISI", tarih=date(2026, 7, 15)), self.eslesme, self.defter,
                                  self.harita, son_donem=date(2026, 7, 1))
        self.assertEqual(s.donem_eslesme, "onceki_donem")
        metin = " ".join(s.uyarilar).lower()
        self.assertIn("cikis kaydi da yok", metin)
        self.assertNotIn("cikis masrafi", metin)

    def test_ilk_donemden_once_ama_calisan(self):
        s = MM.masraf_merkezi_coz(_gider("TEST KISI", tarih=date(2025, 10, 15)), self.eslesme, self.defter,
                                  self.harita, son_donem=date(2026, 6, 1))
        self.assertEqual(s.donem_eslesme, "ilk_donem_oncesi")
        metin = " ".join(s.uyarilar).lower()
        self.assertIn("tarihinden beri calisiyor", metin)
        self.assertNotIn("henuz ise baslamamis", metin)


class AyniProjeTest(unittest.TestCase):
    def test_harita_koduyla_karsilastirir(self):
        h = MM.MasrafMerkeziHaritasi([MM.MasrafMerkezi(gorev_yeri="Ust-Luga – Reshetnikova Office", kod="UL-RESHET",
                                                        ad="Reshetnikova", sirket="RHI", aktif=True)])
        self.assertTrue(MM._ayni_proje(h, "Ust-Luga - St. Petersburg Office", "Ust-Luga – Reshetnikova Office"))
        self.assertFalse(MM._ayni_proje(h, "GPP Project", "Ust-Luga – Reshetnikova Office"))


@unittest.skipUnless(PERSONEL.is_file(), "ornek personel verisi yok")
class GercekVeriDavranisTest(unittest.TestCase):
    """Ikinci tur incelemede degisen gercek satirlar (adlar altin.json'da)."""

    @classmethod
    def setUpClass(cls):
        cls.davranis = altin_veya_none("davranis")
        if not cls.davranis:
            raise unittest.SkipTest("altin davranis kayitlari yok")
        from masraf.defter import Defterler
        from masraf.yardimci_defter import YardimciDefter
        cls.defter = PersonelDefteri.yukle(PERSONEL)
        cls.yardimci = YardimciDefter.yukle(YARDIMCI) if YARDIMCI.is_file() else None
        cls.gecici = tempfile.mkdtemp(prefix="tur2_")
        veri = KOK / "veri"
        for ad in ("aliases.csv", "harici_kisiler.csv", "ek_kisiler.csv", "tckn_sicil.csv", "masraf_merkezi_haritasi.csv"):
            if (veri / ad).is_file():
                (Path(cls.gecici) / ad).write_bytes((veri / ad).read_bytes())
        cls.defterler = Defterler(Path(cls.gecici))

    def _esle(self, anahtar):
        d = self.davranis[anahtar]
        e = E.Eslestirici(self.defter, self.defterler, yardimci=self.yardimci)
        return d, e.esle(_gider(d["kisi_ham"], d["gider_tipi"]))

    def test_bulanik_ad_farki_sicil_doldurmaz(self):
        d, s = self._esle("bulanik_ad_farki")
        self.assertEqual(s.yontem, "bulanik"); self.assertIsNone(s.sicil); self.assertEqual(s.aday_siciller, d["aday"])

    def test_ters_alt_kume_ucuncu_ad(self):
        d, s = self._esle("ters_alt_kume_ucuncu_ad")
        self.assertEqual(s.yontem, "alt_kume"); self.assertIsNone(s.sicil); self.assertLessEqual(s.guven, 0.5)

    def test_1c_birebir_tahminden_once(self):
        if self.yardimci is None:
            self.skipTest("1C listesi yok")
        d, s = self._esle("yardimci_once_tahmin")
        self.assertEqual(s.yontem, "yardimci_defter"); self.assertEqual(s.sicil, d["sicil"])

    def test_bitisik_kesik_ad_prefix(self):
        d, s = self._esle("prefix_bitisik_kesik")
        self.assertEqual(s.yontem, "prefix"); self.assertEqual(s.sicil, d["sicil"]); self.assertEqual(s.guven, 0.85)

    def test_alias_es_isimli_incelemeye_duser(self):
        if not (Path(self.gecici) / "aliases.csv").is_file():
            self.skipTest("alias defteri yok")
        d, s = self._esle("alias_es_isimli")
        self.assertEqual(s.yontem, "alias"); self.assertEqual(s.sicil, d["sicil"])
        self.assertLess(s.guven, MM.GUVEN_ESIGI); self.assertGreater(s.aday_sayisi, 1)

    def test_alt_kume_soyad_tutmayinca_incele(self):
        d, s = self._esle("alt_kume_soyad_tutmuyor")
        self.assertEqual(s.yontem, "alt_kume"); self.assertEqual(s.sicil, d["sicil"]); self.assertEqual(s.guven, 0.72)

    def test_kutuk_kaydi_aile_kaniti_uretmez(self):
        e = E.Eslestirici(self.defter, self.defterler, yardimci=self.yardimci)
        o = altin("aktif_sicil")
        kesin = Eslesme(sicil=o["sicil"], ad_soyad=o["ad_soyad"], yontem="tam_isim", guven=0.95, aday_sayisi=1, aciklama="")
        e._ogren(kesin, _gider(o["ad_soyad"], kaynak_tip="energo_saglik"))
        self.assertFalse(e._kesin_soyadlar, "kutuk kaynagi kanit uretmemeli")
        e._ogren(Eslesme(sicil=o["sicil"], ad_soyad=o["ad_soyad"], yontem="alt_kume", guven=0.90, aday_sayisi=1, aciklama=""),
                 _gider(o["ad_soyad"]))
        self.assertFalse(e._kesin_soyadlar, "tahmin kademesi kanit uretmemeli")
        e._ogren(kesin, _gider(o["ad_soyad"]))
        self.assertTrue(e._kesin_soyadlar)

    def test_tek_tokenli_ad_bulanik_havuzunda_kazanmaz(self):
        e = E.Eslestirici(self.defter, self.defterler, yardimci=self.yardimci)
        for isim in ("MANOHAR LAL", "JAHIRUL ISLAM", "CELENK YUNUS"):
            s = e.esle(_gider(isim, "Otel"))
            with self.subTest(isim=isim):
                self.assertFalse(s.yontem == "bulanik" and s.guven >= MM.GUVEN_ESIGI, (s.yontem, s.guven))
                if s.yontem == "bulanik" and s.sicil:
                    self.assertGreaterEqual(len(e._isim_tokenlar.get(
                        e._defter.sicil_ile(s.sicil).get("ad_soyad_norm", ""), frozenset())), 2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
