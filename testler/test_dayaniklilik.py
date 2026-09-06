"""Sessizce yanlis sonuc ureten durumlara karsi gerileme testleri.

Her test, dayaniklilik incelemesinde OLCULEREK bulunmus bir vakayi temsil
eder. Gercek veri gerektirmez; sentetik dosyalarla calisir.
"""

from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from masraf.mahsuplasma import DAGITILAMAYAN, KontrolSatiri, mahsuplasma_uret
from masraf.modeller import (
    DURUM_ESLESMEDI,
    DURUM_OTOMATIK,
    Eslesme,
    GiderSatiri,
    Sonuc,
)


def gider(kaynak="a.xlsx", kisi="AHMET YILMAZ", tutar=100.0, tarih=date(2026, 7, 5),
          satir_no=1, tip="genel", para="USD", ek=None):
    return GiderSatiri(kaynak_dosya=kaynak, kaynak_tip=tip, satir_no=satir_no,
                       belge_tarihi=tarih, aciklama=kisi or "", kisi_ham=kisi,
                       sicil_ham=None, tckn_ham=None, tutar=tutar, para_birimi=para,
                       masraf_merkezi_kaynak=None, gider_tipi="Vize", ek=dict(ek or {}))


def sonuc(satir, merkez="GPP", durum=DURUM_OTOMATIK, sirket="RHI"):
    satir.ek.setdefault("masraf_merkezi_adi", "GPP Project")
    satir.ek.setdefault("masraf_merkezi_haritada", True)
    return Sonuc(satir=satir,
                 eslesme=Eslesme(sicil="1", ad_soyad=satir.kisi_ham, yontem="tam_isim",
                                 guven=1.0, aday_sayisi=1, aciklama=""),
                 donem=date(2026, 7, 1), gorev_yeri="Ust-Luga", masraf_merkezi=merkez,
                 sirket=sirket, sirket2=sirket, statu="Aktif", kategori="Aktif",
                 cikis_tarihi=None, durum=durum)


class FarkliKisilerEslenmezTest(unittest.TestCase):
    """Ayni gun ayni sabit ucret, iki dosyada BASKA kisiler: yinelenen degildir."""

    def test_esit_sayida_farkli_kisiler_korunur(self):
        a = [sonuc(gider("a.xlsx", k, 80.0, satir_no=i)) for i, k in enumerate(("ALI VELI", "AYSE KAYA", "CAN DEMIR"), 1)]
        b = [sonuc(gider("b.xlsx", k, 80.0, satir_no=i)) for i, k in enumerate(("HASAN AK", "ZEYNEP OZ", "MURAT SU"), 1)]
        t = mahsuplasma_uret(a + b)
        self.assertEqual(t.yinelenen_sayisi, 0)
        self.assertEqual(round(sum(m.tutar for m in t.satirlar), 2), 480.0)
        self.assertTrue(t.kapali_mi)

    def test_ayni_kisinin_iki_gercek_islemi_teke_inmez(self):
        a = [sonuc(gider("a.xlsx", "ALI VELI", 120.0, satir_no=1)),
             sonuc(gider("a.xlsx", "ALI VELI", 120.0, satir_no=2))]
        b = [sonuc(gider("b.xlsx", "ALI VELI", 120.0, satir_no=1))]
        t = mahsuplasma_uret(a + b)
        self.assertEqual(t.yinelenen_sayisi, 1)
        self.assertEqual(round(sum(m.tutar for m in t.satirlar), 2), 240.0)

    def test_isimsiz_kurumsal_kalemler_hala_eslenir(self):
        a = [sonuc(gider("a.xlsx", None, 500.0))]
        b = [sonuc(gider("b.xlsx", "", 500.0))]
        t = mahsuplasma_uret(a + b)
        self.assertEqual(t.yinelenen_sayisi, 1)


class EslesmediSirketTasimazTest(unittest.TestCase):
    def test_dagitilamayan_sirketsiz(self):
        t = mahsuplasma_uret([sonuc(gider(), merkez="GPP", durum=DURUM_ESLESMEDI, sirket="UST LUGA")])
        (m,) = t.satirlar
        self.assertEqual(m.masraf_merkezi, DAGITILAMAYAN)
        self.assertIsNone(m.sirket)
        gruplar = t.sirket_ozeti()
        self.assertEqual([g["sirket"] for g in gruplar], ["(dagitilamayan)"])


class BeyanToleransiTest(unittest.TestCase):
    def test_kurus_farki_tedarikci_yuvarlamasi_kapali(self):
        k = KontrolSatiri(kaynak="x", para_birimi="USD", gelen=6505.64, dagitilan=6505.64,
                          dagitilamayan=0.0, satir_sayisi=6, beyan_toplam=6505.63)
        self.assertTrue(k.kapali_mi, k.acik_sebebi)

    def test_buyuk_fark_acik(self):
        k = KontrolSatiri(kaynak="x", para_birimi="USD", gelen=1000.0, dagitilan=1000.0,
                          dagitilamayan=0.0, satir_sayisi=6, beyan_toplam=1076.78)
        self.assertFalse(k.kapali_mi)
        self.assertIn("beyan", k.acik_sebebi)

    def test_tolerans_ustu_sinirli(self):
        k = KontrolSatiri(kaynak="x", para_birimi="USD", gelen=100.0, dagitilan=100.0,
                          dagitilamayan=0.0, satir_sayisi=500, beyan_toplam=100.60)
        self.assertFalse(k.kapali_mi)   # 0,50'den buyuk fark yuvarlama olamaz


class GenelOkuyucuTest(unittest.TestCase):
    def _xlsx(self, klasor, ad, basliklar, satirlar):
        import openpyxl
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sayfa1"
        ws.append(basliklar)
        for r in satirlar:
            ws.append(r)
        yol = Path(klasor) / ad
        wb.save(yol)
        return yol

    def test_doviz_kolonu_okunur(self):
        from masraf.okuyucular.genel import genel_oku
        with tempfile.TemporaryDirectory() as d:
            yol = self._xlsx(d, "doviz.xlsx", ["Ad Soyad", "Tarih", "Tutar", "Para Birimi"],
                             [["ALI VELI", "05.07.2026", 100, "USD"], ["AYSE KAYA", "05.07.2026", 9000, "RUB"]])
            satirlar = genel_oku(yol)
        self.assertEqual([s.para_birimi for s in satirlar], ["USD", "RUB"])

    def test_tutar_basligindaki_doviz(self):
        from masraf.okuyucular.genel import genel_oku
        with tempfile.TemporaryDirectory() as d:
            yol = self._xlsx(d, "usd.xlsx", ["Ad Soyad", "Tarih", "Tutar (USD)"],
                             [["ALI VELI", "05.07.2026", 100]])
            (s,) = genel_oku(yol)
        self.assertEqual(s.para_birimi, "USD")

    def test_soyadi_genel_olan_kisi_atilmaz(self):
        from masraf.okuyucular.genel import genel_oku
        with tempfile.TemporaryDirectory() as d:
            yol = self._xlsx(d, "genel.xlsx", ["Ad Soyad", "Tutar"],
                             [["GENEL AHMET", 10], ["TOTAL MEHMET", 10], ["IADE AYSE", 10],
                              ["OZET ALI", 10], ["YILMAZ AHMET", 10], ["TOPLAM", 50], ["GENEL TOPLAM", 50]])
            satirlar = genel_oku(yol)
        self.assertEqual(sorted(s.kisi_ham for s in satirlar),
                         ["GENEL AHMET", "IADE AYSE", "OZET ALI", "TOTAL MEHMET", "YILMAZ AHMET"])

    def test_baslik_12_satirda_olsa_da_okunur(self):
        from masraf.okuyucular.genel import genel_oku
        with tempfile.TemporaryDirectory() as d:
            ust = [[f"ust bilgi {i}"] for i in range(11)]
            yol = self._xlsx(d, "baslik12.xlsx", ["fatura"], ust + [["Ad Soyad", "Tutar"], ["ALI VELI", 5]])
            satirlar = genel_oku(yol)
        self.assertEqual([s.kisi_ham for s in satirlar], ["ALI VELI"])


class AyniAdliDosyaTest(unittest.TestCase):
    """Boru: ayni adli ikinci dosya ya atlanir ya ayri fatura olarak etiketlenir."""

    def _boru(self):
        from masraf.boru import Boru, CalismaAyarlari
        ayarlar = CalismaAyarlari(personel_yolu=Path("yok.xlsx"), veri_dizini="veri",
                                  cikti_dizini="cikti", defterleri_besle=False, ogrenmeyi_kaydet=False)
        return Boru(ayarlar)

    def test_ayni_icerik_atlanir(self):
        b = self._boru()
        gorulen = {}
        ilk = b._ayni_adli_dosyayi_ele([gider("m1.msg > EK.xlsx", "A B", 10.0), gider("m1.msg > EK.xlsx", "C D", 20.0)], gorulen)
        ikinci = b._ayni_adli_dosyayi_ele([gider("EK.xlsx", "A B", 10.0), gider("EK.xlsx", "C D", 20.0)], gorulen)
        self.assertEqual(len(ilk), 2)
        self.assertEqual(ikinci, [])
        self.assertTrue(any("iki kez verildi" in u for u in b.uyarilar))

    def test_farkli_icerik_ayri_etiketlenir(self):
        b = self._boru()
        gorulen = {}
        b._ayni_adli_dosyayi_ele([gider("m1.msg > EK.xlsx", "A B", 10.0)], gorulen)
        ikinci = b._ayni_adli_dosyayi_ele([gider("m2.msg > EK.xlsx", "A B", 10.0), gider("m2.msg > EK.xlsx", "C D", 5.0)], gorulen)
        self.assertEqual(len(ikinci), 2)
        self.assertEqual(ikinci[0].ek["kaynak_etiketi"], "m2.msg > EK.xlsx")
        self.assertTrue(any("FARKLI dosya" in u for u in b.uyarilar))
        # Mahsuplasmada ayri kontrol satiri olur.
        t = mahsuplasma_uret([sonuc(s) for s in ikinci] + [sonuc(gider("m1.msg > EK.xlsx", "A B", 10.0))])
        self.assertEqual(sorted(k.kaynak for k in t.kontrol), ["EK.xlsx", "m2.msg > EK.xlsx"])


class HaritaKolonUyarisiTest(unittest.TestCase):
    def test_eksik_kolon_uyarisi(self):
        from masraf.masraf_merkezi import MasrafMerkeziHaritasi
        with tempfile.TemporaryDirectory() as d:
            yol = Path(d) / "harita.csv"
            yol.write_text("site,code\nUst Luga,GPP\n", encoding="utf-8")
            h = MasrafMerkeziHaritasi.yukle(yol)
        self.assertTrue(h.kaynak_var)
        self.assertIn("gorev_yeri", h.yukleme_uyarisi or "")


class PostaAtlananEklerTest(unittest.TestCase):
    def test_pdf_ekler_listelenir(self):
        try:
            import extract_msg  # noqa: F401
        except ImportError:
            self.skipTest("extract_msg yok")
        from masraf.okuyucular import kesif
        mesaj = Path("ornek_veri/posta/ornek_mail.msg")
        if not mesaj.is_file():
            self.skipTest("ornek mail yok")
        atlananlar: list[str] = []
        with tempfile.TemporaryDirectory() as d:
            kesif.oku(mesaj, cikarma_dizini=d, atlanan_ekler=atlananlar)
        self.assertTrue(atlananlar)
        self.assertTrue(any(str(a).lower().split("  [")[0].endswith(".pdf") for a in atlananlar))


if __name__ == "__main__":
    unittest.main(verbosity=2)
