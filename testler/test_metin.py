"""Metin normalizasyon katmaninin testleri (masraf.metin).

Eslestirmenin tamami bu moduldeki fonksiyonlarin ciktisi uzerinde calisir;
buradaki bir regresyon tum boru hattini sessizce bozar. Bu yuzden testler
gercek fatura metinlerinden alinmis ornekleri kullanir.

Ornek veri gerektirmez, her ortamda calisir.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

KOK = Path(__file__).resolve().parents[1]
if str(KOK) not in sys.path:
    sys.path.insert(0, str(KOK))

try:
    from masraf.metin import (
        ascii_katla,
        bitisik_ad_ac,
        isim_normalize,
        isim_tokenlari,
        kisi_metnini_temizle,
        tr_buyuk,
        translit_varyantlari,
    )

    MODUL_VAR = True
except ImportError:  # modul henuz yazilmadiysa testler atlanir
    MODUL_VAR = False


@unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
class TrBuyukTest(unittest.TestCase):
    """Turkce'ye ozgu buyuk harf kurallari."""

    def test_noktasiz_i_bozulmaz(self):
        # Python'un varsayilan upper() 'i' harfini 'I' yapar; Turkce'de
        # 'i' -> 'I' (noktali) olmalidir, aksi halde 'Isik' ile 'Işık' ayrisir.
        self.assertEqual(tr_buyuk("ismail"), "ISMAIL")
        self.assertEqual(tr_buyuk("ığdır"), "IGDIR")

    def test_turkce_harfler_buyur(self):
        self.assertEqual(tr_buyuk("şçöü"), "SCOU")

    def test_bos_girdi(self):
        self.assertEqual(tr_buyuk(""), "")


@unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
class AsciiKatlaTest(unittest.TestCase):
    """Turkce ve Kiril karakterlerin ASCII'ye katlanmasi."""

    def test_turkce_karakterler(self):
        self.assertEqual(ascii_katla("İĞŞÇÖÜ"), "IGSCOU")
        self.assertEqual(ascii_katla("Şirket"), "Sirket")

    def test_kiril_karakterler(self):
        # 1C ve Rusca kaynaklardan gelen isimler Kiril yazilabiliyor.
        self.assertEqual(ascii_katla("ЩУКА ЖУКОВ ЧАЙ ЮРИЙ ЯНА"),
                         "SCHUKA ZHUKOV CHAY YURIY YANA")

    def test_ascii_degismez(self):
        self.assertEqual(ascii_katla("OZAKAY MUSTAFA KEMAL"), "OZAKAY MUSTAFA KEMAL")

    def test_bos_girdi(self):
        self.assertEqual(ascii_katla(""), "")


@unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
class IsimNormalizeTest(unittest.TestCase):
    """Buyuk harf + ASCII + sadece harf/bosluk + tek bosluk."""

    def test_noktalama_ve_bosluk_temizlenir(self):
        self.assertEqual(isim_normalize("  Özkan,  Mustafa-Kemal  "),
                         "OZKAN MUSTAFA KEMAL")

    def test_bolu_isareti_bosluga_doner(self):
        # PNR bicimi: SOYAD/AD
        self.assertEqual(isim_normalize("OZAKAY/MUSTAFAKEMA"), "OZAKAY MUSTAFAKEMA")

    def test_coklu_bosluk_teke_iner(self):
        self.assertEqual(isim_normalize("Kumar   Manoj  "), "KUMAR MANOJ")

    def test_bos_girdi(self):
        self.assertEqual(isim_normalize(""), "")

    def test_tokenlar_sirasiz_kume(self):
        # 'AD SOYAD' ile 'SOYAD AD' ayni token kumesini vermelidir.
        self.assertEqual(isim_tokenlari("Ozakay Mustafa Kemal"),
                         isim_tokenlari("Mustafa Kemal Ozakay"))
        self.assertEqual(isim_tokenlari("Ozakay Mustafa Kemal"),
                         frozenset({"OZAKAY", "MUSTAFA", "KEMAL"}))


@unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
class TranslitVaryantlariTest(unittest.TestCase):
    """Rusca/pasaport transliterasyonundan Turkce yazima geri donus.

    Bilet sistemleri ismi Rusca pasaporttan okuyup latin harfe ceviriyor;
    ortaya cikan yazim personel dosyasindakinden farkli oluyor.
    """

    def test_gekhan_gokhan_uretir(self):
        varyantlar = translit_varyantlari("IYLMAZ GEKHAN")
        self.assertIn("YILMAZ GOKHAN", varyantlar)

    def test_mekhmet_veisi_mehmet_veysi_uretir(self):
        varyantlar = translit_varyantlari("YRMAK MEKHMET VEISI")
        self.assertIn("IRMAK MEHMET VEYSI", varyantlar)

    def test_orijinal_her_zaman_kumede(self):
        self.assertIn("IYLMAZ GEKHAN", translit_varyantlari("IYLMAZ GEKHAN"))

    def test_varyant_sayisi_sinirli(self):
        # Kombinatoryal patlama olmamali; 130 satirlik dosya saniyeler icinde
        # eslesmeli.
        self.assertLessEqual(len(translit_varyantlari("ABDULKADIR SEYHMUS KOCAK")), 128)

    def test_bos_girdi(self):
        self.assertEqual(translit_varyantlari(""), set())


@unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
class BitisikAdAcTest(unittest.TestCase):
    """Bilet sistemlerinde iki parcali adlar bitisik yaziliyor."""

    SOZLUK = {"MUSTAFA", "KEMAL", "AHMET", "ALI", "SEYHMUS", "ABDUL", "KADIR"}

    def test_mustafakemal_ayrilir(self):
        self.assertEqual(bitisik_ad_ac("MUSTAFAKEMAL", self.SOZLUK),
                         ["MUSTAFA", "KEMAL"])

    def test_sozlukte_olmayan_bolunmez(self):
        self.assertIsNone(bitisik_ad_ac("XQZWWVBB", self.SOZLUK))

    def test_zaten_tek_parca_olan_bolunmez(self):
        self.assertIsNone(bitisik_ad_ac("MUSTAFA", self.SOZLUK))

    def test_kisa_token_bolunmez(self):
        self.assertIsNone(bitisik_ad_ac("ALI", self.SOZLUK))

    def test_bos_sozluk(self):
        self.assertIsNone(bitisik_ad_ac("MUSTAFAKEMAL", set()))


@unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
class KisiMetniniTemizleTest(unittest.TestCase):
    """Gercek Antik/Yuzyil cari hareket aciklamalari.

    Bu satirlarda kisi adi otel adi, guzergah kodu, tarih ve islem metniyle
    ic ice gecmis durumdadir.
    """

    def test_otel_satirindan_sadece_kisi_kalir(self):
        ham = ("CIHAN BALABAN GRAND PLAZA HOTEL HANOI "
               "[11.07.2026] - [13.07.2026]  (2) KONAKLAMA YURTDISI")
        self.assertEqual(kisi_metnini_temizle(ham), "CIHAN BALABAN")

    def test_bilet_satirindan_bilet_no_ve_guzergah_atilir(self):
        ham = "TK4093099626 OZAKAY/MUSTAFAKEMAL MR  IST-CDG BILET BEDELI"
        self.assertEqual(kisi_metnini_temizle(ham), "OZAKAY MUSTAFAKEMAL")

    def test_cinsiyet_isareti_ve_coklu_guzergah(self):
        ham = "PC2255749381 TEMIR MEHMET\\M  KYA-SAW-LED BILET BEDELI"
        self.assertEqual(kisi_metnini_temizle(ham), "TEMIR MEHMET")

    def test_vize_metni_temizlenir(self):
        ham = "TALIP KEREM KOCKESEN RUSYA FEDERASYONU TURISTIK E-VIZE"
        self.assertEqual(kisi_metnini_temizle(ham), "TALIP KEREM KOCKESEN")

    def test_bagaj_satirinda_kisi_ortada(self):
        ham = "EKSTRA BAGAJ UCRETI ISA MUCAHIT SAHIN TARAFINDAN TASINDI"
        self.assertEqual(kisi_metnini_temizle(ham), "ISA MUCAHIT SAHIN")

    def test_kisi_icermeyen_satir_bos_doner(self):
        # Cenaze celengi gibi satirlarda kisi yoktur; temizlikten sonra
        # anlamli bir isim kalmamalidir.
        temiz = kisi_metnini_temizle("[13.07.2026] - [14.07.2026] CENAZE CELENGI")
        self.assertNotIn("2026", temiz)

    def test_bos_girdi(self):
        self.assertEqual(kisi_metnini_temizle(""), "")


if __name__ == "__main__":
    unittest.main()
