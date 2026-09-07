"""PDF fatura okuyucusu: normalizasyon, alan cikarimi, oz-denetim, hata yolu.

Buradaki en kritik test yumusak tire (U+00AD) testidir: o karakter ekranda
tire gibi gorunur, regex'i sessizce kirar ve alan bos doner. Bu sinifta bir
hata kimsenin gozune carpmaz, bu yuzden birim testiyle kilitlenir.

Ornek metinler gercek sablonlardan turetilmistir ama kisi adi, TCKN ve IBAN
gibi kisisel veri ICERMEZ; tutarlar da degistirilmistir.
"""

from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from masraf.envanter import FATURA_PDF, OKUNAMADI, TARANMIS, satirlardan_kayit
from masraf.modeller import KAYNAK_TIPLERI
from masraf.okuyucular.fatura_pdf import fatura_basligi_coz, fatura_pdf_oku
from masraf.okuyucular.pdf_metin import metni_normalize

# Etiket ve deger AYNI satirda (AS Olcme / Kerem Kockesen sablonu).
SATIR_ICI = """SAYIN
ORNEK TAAHHUT INSAAT VE TICARET ANONIM SIRKETI
VKN: 1234567890
Senaryo: TICARIFATURA
Fatura Tipi: SATIS
Fatura No: ASS2026000009999
Fatura Tarihi: 22-05-2026
VKN: 0987654321
e-FATURA
ETTN: EEA26779-377C-44BA-BAA9-BD1933DF2993
Mal Hizmet Toplam Tutari 41.250,00 TL
Toplam Indirim 0,00 TL
Hesaplanan KDV(%20) 8.250,00 TL
Vergiler Dahil Toplam Tutar 49.500,00 TL
Odenecek Tutar 49.500,00 TL
"""

# Etiket ve deger AYRI satirda (A PLUS / arabuluculuk sablonu).
AYRI_SATIR = """ORNEK ARABULUCULUK VE EGITIM A.S.
Vergi No
2233445566
e-Fatura
Sayin
ORNEK TAAHHUT INSAAT VE TICARET
Vergi No
1234567890
Fatura Numarasi
EF02026000009999
Fatura Tarihi
02.06.2026
- 09:38:43
Senaryo
TICARIFATURA
ETTN
0a8fc750-bbb9-4393-b379-aed42b1210f7
Mal Hizmet Toplam Tutari
35.000,00 TL
Toplam Indirim
0,00 TL
Hesaplanan KDV GERCEK (%20.0)
7.000,00 TL
Vergiler Dahil Toplam Tutar
42.000,00 TL
Odenecek Tutar
42.000,00 TL
"""

# Noktasiz 'i', bitisik para birimi, ayracsiz tarih (Unvest / universite sablonu).
BITISIK = """SAYIN
ORNEK TAAHHUT INSAAT VE TICARET ANONIM SIRKETI
VKN: 1234567890
Senaryo: TEMELFATURA
Fatura No: UNF2026000009999
Fatura Tarihi: 09062026
VKN: 3344556677
eFATURA
Mal Hizmet Toplam Tutari 869.272,16TRY
Toplam Iskonto 0,00TRY
Matrah 869.272,16TRY
Hesaplanan KDV(%20) 173.854,43TRY
Vergiler Dahil Toplam Tutar 1.043.126,59TRY
Odenecek Tutar 1.043.126,59TRY
"""


class NormalizasyonTest(unittest.TestCase):
    def test_yumusak_tire_temizlenir(self):
        """U+00AD ekranda tire gibi gorunur ama regex'i sessizce kirar."""
        ham = "Fatura No: ASS­2026­000009999"
        self.assertIn("­", ham)
        self.assertEqual(metni_normalize(ham), "Fatura No: ASS2026000009999")

    def test_sifir_genislikli_bosluk_silinir(self):
        self.assertEqual(metni_normalize("A\u200bB"), "AB")

    def test_bolunmez_bosluk_duz_bosluga_doner(self):
        self.assertEqual(metni_normalize("1.234,56\u00a0TL"), "1.234,56 TL")

    def test_uzun_tire_duz_tireye_doner(self):
        self.assertEqual(metni_normalize("22–05–2026"), "22-05-2026")

    def test_turkce_harfler_korunur(self):
        self.assertEqual(metni_normalize("Ödenecek Tutarı"), "Ödenecek Tutarı")

    def test_bos_girdi(self):
        self.assertEqual(metni_normalize(""), "")


class AlanCikarimiTest(unittest.TestCase):
    def test_satir_ici_sablon(self):
        b = fatura_basligi_coz(SATIR_ICI)
        self.assertEqual(b.fatura_no, "ASS2026000009999")
        self.assertEqual(b.tarih, date(2026, 5, 22))
        self.assertEqual((b.matrah, b.kdv, b.toplam), (41250.0, 8250.0, 49500.0))
        self.assertEqual(b.para_birimi, "TRY")
        self.assertEqual(b.aritmetik, "tamam")
        self.assertEqual(b.bulunamayan, [])

    def test_ayri_satirli_sablon(self):
        """Etiket bir satirda, deger bir alt satirda: A PLUS sablonu."""
        b = fatura_basligi_coz(AYRI_SATIR)
        self.assertEqual(b.fatura_no, "EF02026000009999")
        self.assertEqual(b.tarih, date(2026, 6, 2))
        self.assertEqual((b.matrah, b.kdv, b.toplam), (35000.0, 7000.0, 42000.0))
        self.assertEqual(b.aritmetik, "tamam")

    def test_bitisik_para_birimi_ve_ayracsiz_tarih(self):
        b = fatura_basligi_coz(BITISIK)
        self.assertEqual(b.fatura_no, "UNF2026000009999")
        self.assertEqual(b.tarih, date(2026, 6, 9))
        self.assertEqual(b.toplam, 1043126.59)
        self.assertEqual(b.para_birimi, "TRY")
        self.assertEqual(b.aritmetik, "tamam")

    def test_yumusak_tire_alan_kaybettirmez(self):
        """Yumusak tire iki katmanda birden temizlenir: metni_normalize ve
        alan cikariminin kendi ASCII katlamasi. Ikisi de calismali."""
        bozuk = SATIR_ICI.replace("ASS2026000009999", "ASS\u00ad2026\u00ad000009999")
        self.assertIn("\u00ad", bozuk)
        self.assertEqual(fatura_basligi_coz(bozuk).fatura_no, "ASS2026000009999")
        self.assertEqual(fatura_basligi_coz(metni_normalize(bozuk)).fatura_no,
                         "ASS2026000009999")
        self.assertNotIn("\u00ad", metni_normalize(bozuk))

    def test_vkn_listesi_toplanir(self):
        b = fatura_basligi_coz(SATIR_ICI)
        self.assertEqual(b.vkn_listesi, ["1234567890", "0987654321"])

    def test_aritmetik_tutmayinca_isaretlenir(self):
        """matrah + KDV toplama esit degilse okuma supheli sayilmali."""
        bozuk = SATIR_ICI.replace("Vergiler Dahil Toplam Tutar 49.500,00 TL",
                                  "Vergiler Dahil Toplam Tutar 50.000,00 TL")
        b = fatura_basligi_coz(bozuk)
        self.assertNotEqual(b.aritmetik, "tamam")
        self.assertIn("-500.00", b.aritmetik)

    def test_alan_bulunamayinca_listelenir(self):
        b = fatura_basligi_coz("Bu bir fatura degil, duz bir metin.")
        self.assertFalse(b.okundu_mu)
        self.assertIn("fatura_no", b.bulunamayan)
        self.assertIn("toplam", b.bulunamayan)

    def test_bos_metin_cokmez(self):
        b = fatura_basligi_coz("")
        self.assertIsNone(b.fatura_no)
        self.assertFalse(b.okundu_mu)


class OkuyucuSozlesmesiTest(unittest.TestCase):
    """kesif.PARSERLAR sozlesmesi: tek argument, liste doner, ISTISNA FIRLATMAZ."""

    def test_kaynak_tip_kayitli(self):
        self.assertIn("fatura_pdf", KAYNAK_TIPLERI)

    def test_bozuk_dosya_istisna_firlatmaz(self):
        with tempfile.TemporaryDirectory() as d:
            yol = Path(d) / "bozuk.pdf"
            yol.write_bytes(b"bu bir PDF degil")
            satirlar = fatura_pdf_oku(yol)
        self.assertEqual(len(satirlar), 1)
        self.assertIsNone(satirlar[0].tutar)
        self.assertEqual(satirlar[0].kaynak_tip, "fatura_pdf")

    def test_olmayan_dosya_istisna_firlatmaz(self):
        satirlar = fatura_pdf_oku("/yok/boyle/bir/dosya.pdf")
        self.assertEqual(len(satirlar), 1)
        self.assertTrue(satirlar[0].ek.get("pdf_taranmis") or satirlar[0].aciklama)

    def test_tutar_bilerek_bostur(self):
        """Belge tutari YEREL para biriminde; dagitima girmemeli.

        Tutar dolu birakilsaydi 1.043.126,59 TRY dogrudan mahsuba girer ve
        USD tablosunda anlamsiz bir kalem olusurdu.
        """
        with tempfile.TemporaryDirectory() as d:
            yol = Path(d) / "x.pdf"
            yol.write_bytes(b"%PDF-1.4 sahte")
            s = fatura_pdf_oku(yol)[0]
        self.assertIsNone(s.tutar)
        self.assertIsNone(s.para_birimi)


class EnvanterDurumuTest(unittest.TestCase):
    def _satir(self, **ek):
        with tempfile.TemporaryDirectory() as d:
            yol = Path(d) / "x.pdf"
            yol.write_bytes(b"%PDF-1.4 sahte")
            s = fatura_pdf_oku(yol)[0]
        s.ek.update(ek)
        return s

    def test_taranmis_pdf_kendi_durumunu_alir(self):
        """Taranmis PDF OKUNAMADI sayilmaz: kullanici onu duzeltemez ve
        boru hatti okunamayan eki sert hataya cevirip maili elde birakir."""
        s = self._satir(pdf_taranmis=True)
        kayit = satirlardan_kayit("x.pdf", "mail", "pdf", [s])
        self.assertEqual(kayit.durum, TARANMIS)
        self.assertIn("OCR", kayit.sebep)

    def test_okunan_fatura_no_ve_tutari_sebebe_yazar(self):
        s = self._satir(pdf_taranmis=False, fatura_no="ASS2026000009999",
                        fatura_toplam_yerel=49500.0, para_birimi_yerel="TRY")
        kayit = satirlardan_kayit("x.pdf", "mail", "pdf", [s])
        self.assertEqual(kayit.durum, FATURA_PDF)
        self.assertIn("ASS2026000009999", kayit.sebep)
        self.assertIn("49,500.00 TRY", kayit.sebep)

    def test_basligi_okunamayan_pdf_okunamadi(self):
        s = self._satir(pdf_taranmis=False, fatura_no=None,
                        pdf_bulunamayan_alanlar=["fatura_no", "toplam"])
        kayit = satirlardan_kayit("x.pdf", "mail", "pdf", [s])
        self.assertEqual(kayit.durum, OKUNAMADI)
        self.assertIn("fatura_no", kayit.sebep)

    def test_belge_satiri_mahsuba_girmez(self):
        """Fatura belgesi bir gider satiri degildir: mahsupta ayri sayilir."""
        from masraf.envanter import BELGE_TIPLERI
        self.assertIn("fatura_pdf", BELGE_TIPLERI)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
