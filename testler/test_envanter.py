"""Dosya envanteri: hangi dosya okundu, hangisi atlandi, neden (masraf.envanter)."""

from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from masraf.envanter import (
    ATLANDI, AYNI_ICERIK, DETAY_LISTESI, KUTUK, MAIL, OKUNDU, SATIR_YOK, DosyaKaydi,
    envanter_ozeti, satirlardan_kayit,
    FATURA_PDF,
    TARANMIS,
)
from masraf.modeller import GiderSatiri


def gider(tip="genel", tutar=10.0, para="USD"):
    return GiderSatiri(kaynak_dosya="x", kaynak_tip=tip, satir_no=1, belge_tarihi=date(2026, 7, 1),
                       aciklama="", kisi_ham="A B", sicil_ham=None, tckn_ham=None, tutar=tutar,
                       para_birimi=para, masraf_merkezi_kaynak=None, gider_tipi="Bilet")


class KayitTuretmeTest(unittest.TestCase):
    def test_okundu(self):
        k = satirlardan_kayit("a.xlsx", "m.msg", "genel", [gider(), gider(tutar=5.5)])
        self.assertEqual((k.durum, k.satir, k.tutarli_satir, k.tutar, k.para_birimi), (OKUNDU, 2, 2, 15.5, "USD"))
        self.assertTrue(k.dagilima_girdi)

    def test_tutarsiz_satir_notu(self):
        k = satirlardan_kayit("a.xlsx", "", "genel", [gider(), gider(tutar=None)])
        self.assertEqual(k.durum, OKUNDU)
        self.assertIn("1 satirda tutar okunamadi", k.sebep)

    def test_kutuk(self):
        k = satirlardan_kayit("s.xlsx", "", "energo_saglik", [gider("energo_saglik", None)] * 3)
        self.assertEqual(k.durum, KUTUK)
        self.assertFalse(k.dagilima_girdi)

    def test_detay_listesi(self):
        k = satirlardan_kayit("ASS1.xlsx", "", "energo_assessment", [gider("energo_assessment_detay", None)])
        self.assertEqual(k.durum, DETAY_LISTESI)

    def test_satir_yok(self):
        k = satirlardan_kayit("bos.xlsx", "", "genel", [])
        self.assertEqual(k.durum, SATIR_YOK)

    def test_ozet_sayimi(self):
        o = envanter_ozeti([DosyaKaydi(ad="a", durum=OKUNDU), DosyaKaydi(ad="b", durum=ATLANDI),
                            DosyaKaydi(ad="c", durum=ATLANDI)])
        self.assertEqual(o, {OKUNDU: 1, ATLANDI: 2})


class GercekMailEnvanteriTest(unittest.TestCase):
    MESAJ = Path("ornek_veri/posta/ornek_mail.msg")

    @classmethod
    def setUpClass(cls):
        if not cls.MESAJ.is_file():
            raise unittest.SkipTest("ornek mail yok")
        from masraf.okuyucular import kesif
        cls.envanter: list = []
        with tempfile.TemporaryDirectory() as d:
            cls.satirlar = kesif.oku(cls.MESAJ, cikarma_dizini=d, gorulen_ozetler=set(), envanter=cls.envanter)

    def test_her_ek_bir_kayit(self):
        sayim = envanter_ozeti(self.envanter)
        # Kok mail tek; ic mailler (ekli .msg) de MAIL satiri olarak envantere girer.
        self.assertEqual(sum(1 for k in self.envanter if k.durum == MAIL and not k.kaynak), 1, sayim)
        self.assertGreaterEqual(sayim.get(MAIL, 0), 1, sayim)
        # PDF'ler artik atlanmiyor: basligi okunanlar FATURA (PDF), metin
        # katmani olmayanlar TARANMIS PDF olarak kayda giriyor.
        self.assertGreaterEqual(sayim.get(FATURA_PDF, 0), 10, sayim)
        self.assertGreaterEqual(sayim.get(TARANMIS, 0), 3, sayim)
        self.assertEqual(sayim.get(ATLANDI, 0), 0, sayim)
        self.assertGreaterEqual(sayim.get(OKUNDU, 0), 4, sayim)       # tutarli tablolar
        self.assertGreaterEqual(sayim.get(KUTUK, 0), 3, sayim)        # katilimci, saglik, sigorta
        self.assertEqual(sayim.get(DETAY_LISTESI, 0), 4, sayim)       # ASS fatura detaylari
        self.assertGreaterEqual(sayim.get(AYNI_ICERIK, 0), 1, sayim)  # ic mailde tekrar eden ekler

    def test_okunan_tutarlar_satirlarla_tutuyor(self):
        for k in self.envanter:
            if k.durum != OKUNDU:
                continue
            beklenen = round(sum(float(s.tutar) for s in self.satirlar
                                 if s.kaynak_dosya.split("> ")[-1] == k.ad and s.tutar is not None), 2)
            self.assertAlmostEqual(k.tutar, beklenen, places=2, msg=k.ad)

    def test_pdf_kayitlari_kaynak_zinciri_tasir(self):
        pdfler = [k for k in self.envanter
                  if k.durum in (FATURA_PDF, TARANMIS) and k.ad.lower().endswith(".pdf")]
        self.assertTrue(pdfler)
        for k in pdfler:
            self.assertTrue(k.kaynak.startswith(self.MESAJ.name), k.kaynak)
            self.assertTrue(k.sebep)

    def test_okunan_pdf_fatura_no_ve_tutar_tasir(self):
        """Basligi okunan her PDF fatura no ve belge tutari yazmali.

        Kanit zincirinin ilk halkasi: kullanici Excel'de bir tutari gorup
        'hangi belgeden' diye sordugunda cevap bu kayittadir.
        """
        okunan = [k for k in self.envanter if k.durum == FATURA_PDF]
        self.assertGreaterEqual(len(okunan), 10, [k.ad for k in okunan])
        for k in okunan:
            self.assertIn("fatura ", k.sebep, k.ad)
            self.assertIn("belge tutari", k.sebep, k.ad)

    def test_taranmis_pdf_durustce_isaretli(self):
        """Metin katmani olmayan PDF sessizce okunmus sayilmamali."""
        taranmis = [k for k in self.envanter if k.durum == TARANMIS]
        self.assertEqual(len(taranmis), 3, [k.ad for k in taranmis])
        for k in taranmis:
            self.assertIn("OCR", k.sebep, k.ad)


class BoruEnvanteriTest(unittest.TestCase):
    def test_bozuk_dosya_okunamadi_kaydi(self):
        from masraf.boru import Boru, CalismaAyarlari
        ana = Path("ornek_veri/personel/2025_2026_giris_cikis.xlsx")
        if not ana.is_file():
            self.skipTest("ornek personel verisi yok")
        with tempfile.TemporaryDirectory() as d:
            bozuk = Path(d) / "BOZUK.xlsx"
            bozuk.write_bytes(b"bozuk")
            kopya = Path(d) / "KOPYA.xlsx"
            kopya.write_bytes(b"bozuk")   # ayni icerik
            b = Boru(CalismaAyarlari(personel_yolu=ana, veri_dizini="veri", cikti_dizini=d,
                                     defterleri_besle=False, ogrenmeyi_kaydet=False))
            b.isle([bozuk, kopya])
        durumlar = {k.ad: k.durum for k in b.envanter}
        self.assertEqual(durumlar.get("BOZUK.xlsx"), "OKUNAMADI")
        self.assertEqual(durumlar.get("KOPYA.xlsx"), AYNI_ICERIK)


if __name__ == "__main__":
    unittest.main(verbosity=2)
