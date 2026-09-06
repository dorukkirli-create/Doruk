"""Mahsuplasma incelemesi (ikinci tur) gerileme testleri.

Yineleme elemesinin sessiz tarafi: siki/gevsek dosya ciftleri, yineleme
suphesi raporu, kismi paylasim, para birimi normalizasyonu, (DAGITILAMAYAN)
bolunmez, isimsiz kalemler, merkez ozeti sirketleri.
"""

from __future__ import annotations

import unittest
from datetime import date

from masraf.mahsuplasma import DAGITILAMAYAN, mahsuplasma_uret
from masraf.modeller import DURUM_ESLESMEDI, DURUM_OTOMATIK, Eslesme, GiderSatiri, Sonuc


def gider(kaynak="a.xlsx", kisi="AHMET YILMAZ", tutar=100.0, tarih=date(2026, 7, 5), satir_no=1,
          tip="genel", para="USD", ek=None, santiye=None, gider_tipi="Vize"):
    return GiderSatiri(kaynak_dosya=kaynak, kaynak_tip=tip, satir_no=satir_no, belge_tarihi=tarih,
                       aciklama=kisi or "", kisi_ham=kisi, sicil_ham=None, tckn_ham=None, tutar=tutar,
                       para_birimi=para, masraf_merkezi_kaynak=santiye, gider_tipi=gider_tipi, ek=dict(ek or {}))


def sonuc(satir, merkez="GPP", durum=DURUM_OTOMATIK, sirket="UST LUGA", sicil="1"):
    satir.ek.setdefault("masraf_merkezi_adi", "GPP Project")
    satir.ek.setdefault("masraf_merkezi_haritada", True)
    return Sonuc(satir=satir, eslesme=Eslesme(sicil=sicil, ad_soyad=satir.kisi_ham, yontem="tam_isim", guven=1.0,
                                              aday_sayisi=1, aciklama=""),
                 donem=date(2026, 7, 1), gorev_yeri="GPP", masraf_merkezi=merkez, sirket=sirket, sirket2=sirket,
                 statu="Aktif", kategori="Aktif", cikis_tarihi=None, durum=durum)


class SikiGevsekEslemeTest(unittest.TestCase):
    def test_bagimsiz_dosyalarda_tek_ortak_kelime_elenmez(self):
        a = sonuc(gider("VIZE_A.xlsx", "MEHMET YILMAZ", 120.0), sicil="1")
        b = sonuc(gider("VIZE_B.xlsx", "MEHMET DEMIR", 120.0), merkez="AGPP", sicil="2")
        t = mahsuplasma_uret([a, b])
        self.assertEqual(t.yinelenen_sayisi, 0)
        self.assertEqual(round(t.toplamlar()["USD"]["net"], 2), 240.0)

    def test_ham_elle_ciftinde_tek_ortak_kelime_elenir_ve_suphe_yazilir(self):
        a = sonuc(gider("ENERGO.xls", "YILMAZ MEHMET", 120.0, tip="antik_cari"))
        b = sonuc(gider("YUZYIL.xlsx", "MEHMET YILMAS", 120.0, tip="yuzyil_dagitilmis"))
        t = mahsuplasma_uret([a, b])
        self.assertEqual(t.yinelenen_sayisi, 1)
        self.assertEqual(len(t.supheler), 1)
        self.assertIn("bir kelimesi ortak", t.supheler[0].sebep)
        k = {x.kaynak: x for x in t.kontrol}["YUZYIL.xlsx"]
        self.assertEqual(k.suphe_satir, 1)
        self.assertTrue(k.kapali_mi)
        self.assertIn("YINELEME SUPHESI", k.suphe_notu)

    def test_ayni_mailin_ekleri_gevsek(self):
        a = sonuc(gider("m.msg > A.xlsx", "YILMAZ MEHMET", 50.0))
        b = sonuc(gider("m.msg > B.xlsx", "MEHMET YILMAS", 50.0))
        self.assertEqual(mahsuplasma_uret([a, b]).yinelenen_sayisi, 1)

    def test_isimsizler_kendi_aralarinda_eslenir(self):
        a = [sonuc(gider("HAM.xls", None, 50.0, satir_no=1, tip="antik_cari")),
             sonuc(gider("HAM.xls", None, 50.0, satir_no=2, tip="antik_cari"))]
        b = [sonuc(gider("ELLE.xlsx", None, 50.0, satir_no=1, tip="yuzyil_dagitilmis")),
             sonuc(gider("ELLE.xlsx", "AYSE KAYA", 50.0, satir_no=2, tip="yuzyil_dagitilmis"))]
        t = mahsuplasma_uret(a + b)
        self.assertEqual(t.yinelenen_sayisi, 1)
        self.assertEqual(round(t.toplamlar()["USD"]["net"], 2), 150.0)


class KismiYinelemeSuphesiTest(unittest.TestCase):
    def test_tutari_degisen_kalem_suphe_olur(self):
        ham = [sonuc(gider("HAM.xls", "ALI VELI", 100.0, satir_no=1, tip="antik_cari")),
               sonuc(gider("HAM.xls", "AYSE KAYA", 200.0, satir_no=2, tip="antik_cari"))]
        elle = [sonuc(gider("ELLE.xlsx", "ALI VELI", 100.0, satir_no=1, tip="yuzyil_dagitilmis")),
                sonuc(gider("ELLE.xlsx", "AYSE KAYA", 201.0, satir_no=2, tip="yuzyil_dagitilmis"))]
        t = mahsuplasma_uret(ham + elle)
        self.assertEqual(t.yinelenen_sayisi, 1)
        self.assertTrue(any("tutar farkli" in s.sebep for s in t.supheler), t.supheler)
        self.assertTrue(any("YINELEME SUPHESI" in u for u in t.uyarilar), t.uyarilar)
        k = {x.kaynak: x for x in t.kontrol}["ELLE.xlsx"]
        self.assertEqual(k.suphe_satir, 1)
        self.assertEqual(k.suphe_tutar, 201.0)

    def test_tarihsiz_kalem_suphe_olur(self):
        ham = [sonuc(gider("HAM.xls", "ALI VELI", 100.0, satir_no=1, tip="antik_cari")),
               sonuc(gider("HAM.xls", "AYSE KAYA", 77.0, satir_no=2, tarih=None, tip="antik_cari"))]
        elle = [sonuc(gider("ELLE.xlsx", "ALI VELI", 100.0, satir_no=1, tip="yuzyil_dagitilmis")),
                sonuc(gider("ELLE.xlsx", "AYSE KAYA", 77.0, satir_no=2, tarih=None, tip="yuzyil_dagitilmis"))]
        t = mahsuplasma_uret(ham + elle)
        self.assertTrue(any("Tarihsiz" in s.sebep for s in t.supheler), t.supheler)
        k = {x.kaynak: x for x in t.kontrol}["ELLE.xlsx"]
        self.assertEqual(k.tarihsiz_satir, 1)
        self.assertIn("tarihsiz", k.suphe_notu)


class PaylasimTur2Test(unittest.TestCase):
    def _pay(self, etiket, pay, bolen):
        return {"masraf_merkezi": etiket, "pay": pay, "bolen": bolen, "oran": pay / bolen}

    def test_kismi_pay_kalan_kendi_sirketinde(self):
        t = mahsuplasma_uret([sonuc(gider(tutar=100.0, ek={"paylasim": [self._pay("RHI", 1, 2)]}))])
        satirlar = sorted(t.satirlar, key=lambda m: m.sirket)
        self.assertEqual([(m.sirket, round(m.tutar, 2)) for m in satirlar], [("RHI", 50.0), ("UST LUGA", 50.0)])
        self.assertTrue(any("(kalan" in (m.pay_notu or "") for m in satirlar))

    def test_fazla_oran_paylasim_uygulanmaz(self):
        t = mahsuplasma_uret([sonuc(gider(tutar=100.0, ek={"paylasim": [self._pay("RHI", 1, 2), self._pay("RSD", 2, 3)]}))])
        self.assertEqual(len(t.satirlar), 1)
        self.assertEqual(t.satirlar[0].sirket, "UST LUGA")
        self.assertTrue(any("toplami" in u for u in t.uyarilar), t.uyarilar)

    def test_taninmayan_etiket_pay_sayilmaz(self):
        class BosHarita:
            """Ne proje ne tuzel kisi taniyan harita: her etiket taninmaz."""
            def coz(self, etiket):
                return None
            def tuzel_kisi_mi(self, etiket):
                return False
        # Harita VARKEN taninmayan etiket pay sayilmaz. (Harita yokken etiket
        # dogrulanamaz ve sirket adi sayilir; bkz. test_ucte_bir_bolme_kurus_birakmaz.)
        t = mahsuplasma_uret([sonuc(gider(tutar=100.0, ek={"paylasim": [self._pay("BILET", 2, 2)]}))], BosHarita())
        self.assertEqual(len(t.satirlar), 1)
        self.assertEqual(t.satirlar[0].sirket, "UST LUGA")
        self.assertTrue(any("BILET" in u for u in t.uyarilar))

    def test_dagitilamayan_bolunmez(self):
        t = mahsuplasma_uret([sonuc(gider(tutar=99.0, ek={"paylasim": [self._pay("RHI", 1, 3), self._pay("RSD", 2, 3)]}),
                                    durum=DURUM_ESLESMEDI)])
        self.assertEqual(len(t.satirlar), 1)
        self.assertEqual(t.satirlar[0].masraf_merkezi, DAGITILAMAYAN)
        self.assertIsNone(t.satirlar[0].sirket)
        self.assertEqual([g["sirket"] for g in t.sirket_ozeti()], ["(dagitilamayan)"])

    def test_devralinan_paylasim_ham_satiri_kalici_degistirmez(self):
        ham = sonuc(gider("HAM.xls", "ALI VELI", 201.42, tip="antik_cari"))
        elle = sonuc(gider("ELLE.xlsx", "ALI VELI", 201.42, tip="yuzyil_dagitilmis",
                           ek={"paylasim": [self._pay("RHI", 1, 3), self._pay("RENSTROYDETAL", 2, 3)]}))
        t1 = mahsuplasma_uret([ham, elle])
        self.assertEqual(len(t1.satirlar), 2)
        self.assertNotIn("paylasim", ham.satir.ek)
        t2 = mahsuplasma_uret([ham, elle], yinelenenleri_ele=False)
        ham_satirlari = [m for m in t2.satirlar if m.kaynak == "HAM.xls"]
        self.assertEqual(len(ham_satirlari), 1)
        self.assertIsNone(ham_satirlari[0].pay_notu)


class ParaBirimiTest(unittest.TestCase):
    def test_kucuk_harf_ayni_kova(self):
        a = sonuc(gider("ENERGO.xls", "ALI VELI", 100.0, para="usd ", tip="antik_cari"))
        b = sonuc(gider("YUZYIL.xlsx", "ALI VELI", 100.0, para="USD", tip="yuzyil_dagitilmis"))
        t = mahsuplasma_uret([a, b])
        self.assertEqual(t.yinelenen_sayisi, 1)
        self.assertEqual(list(t.toplamlar()), ["USD"])

    def test_para_birimi_yoksa_acik(self):
        t = mahsuplasma_uret([sonuc(gider(para=None))])
        (k,) = t.kontrol
        self.assertEqual(k.para_birimi, "?")
        self.assertFalse(k.kapali_mi)
        self.assertIn("para birimi", k.acik_sebebi)


class ElleEtiketTest(unittest.TestCase):
    def test_elle_etiket_ve_sirket_uyusmazligi(self):
        s = sonuc(gider("ELLE.xlsx", "ALI VELI", 100.0, tip="yuzyil_dagitilmis", santiye="RHI"), sirket="UST LUGA")
        t = mahsuplasma_uret([s])
        self.assertEqual(t.satirlar[0].elle_etiket, "RHI")
        self.assertEqual(t.sirket_uyusmazligi["USD"]["satir"], 1)
        self.assertIn("RHI -> UST LUGA", t.sirket_uyusmazligi["USD"]["ciftler"])
        # Yansitma bilgisi uyari degildir: 'uyarilar' listesine girmez.
        self.assertFalse(any("UYUSMAZ" in u.upper() for u in t.uyarilar))

    def test_evrak_no_tasinir(self):
        s = sonuc(gider(ek={"evrak_no": "VF123"}))
        self.assertEqual(mahsuplasma_uret([s]).satirlar[0].evrak_no, "VF123")


class MerkezOzetiTest(unittest.TestCase):
    def test_coklu_sirket_birlestirilir_ve_pay_mutlak(self):
        t = mahsuplasma_uret([sonuc(gider(tutar=300.0, satir_no=1), sirket="UST LUGA"),
                              sonuc(gider(tutar=50.0, satir_no=2), sirket="RHI"),
                              sonuc(gider(tutar=-100.0, satir_no=3), merkez="UDOKAN", sirket="RHI")])
        oz = {o["masraf_merkezi"]: o for o in t.merkez_ozeti()}
        self.assertEqual(oz["GPP"]["sirket"], "RHI; UST LUGA")
        self.assertGreaterEqual(oz["UDOKAN"]["pay_yuzde"], 0)
        self.assertLessEqual(oz["GPP"]["pay_yuzde"], 100)


if __name__ == "__main__":
    unittest.main(verbosity=2)
