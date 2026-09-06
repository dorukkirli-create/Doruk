"""Kesif ve ozel okuyucularin 3. tur denetim bulgulari (sentetik dosyalarla).

Her test, bagimsiz incelemenin olctugu bir hatayi kucuk bir openpyxl
dosyasiyla yeniden uretir ve duzeltmeyi kanitlar:

1. Assessment: sayfa adi / 'Katilimci' / 'Paket' basligi degisince dosya
   genel okuyucuya dusup TL 'Toplam' kolonunu USD gibi dagitiyordu.
2. Antik: aciklamasinda 'TOPLAM' gecen veri satiri toplam sanilip siliniyordu.
3. Kesif kural 7: tutar kolonu olan 'Sicil No + Masraf Merkezi' dosyasi
   kutuk sayilip dagitilmiyordu.
4. Kesif kural 4: veri hucresindeki 'ARABULUCU' tum dosyayi arabuluculuk
   yapiyordu.
5. Arabuluculuk: fatura detay sayfasi/kolonlari degisince tutar None kaliyor,
   aciklama yanlis sebep soyluyordu.
6. Envanter 'Tur' alani fiilen kullanilan okuyucuyu gostermiyordu.
7. 'IKRAM' adli kisiler hizmet kelimesi sanilip ESLESMEDI cikiyordu.

Gercek kisi adi ve gercek veri kullanilmaz; adlar uydurmadir. Gercek ornek
dosyalar (``ornek_veri/``) varsa davranisin korundugu ayrica dogrulanir.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

KOK = Path(__file__).resolve().parents[1]
if str(KOK) not in sys.path:
    sys.path.insert(0, str(KOK))

from masraf.modeller import GiderSatiri
from masraf.okuyucular import kesif
from masraf.okuyucular.antik import (
    _HIZMET_KELIMELERI,
    _bilet_kisi,
    _diger_kisi,
    _kuyruk_kirp,
    _temizle_ve_dogrula,
    _vize_kisi,
    antik_cari_oku,
)
from masraf.okuyucular.energo import arabulucu_oku, assessment_oku
from masraf.okuyucular.kesif import (
    OKUYUCU_UYARISI,
    dosya_tip_adaylari,
    dosya_tipini_bul,
    oku,
    okuyucu_turu,
)

ORNEK = KOK / "ornek_veri"
ANTIK = ORNEK / "antik_travel" / "ANTIK_CARI_TEMMUZ_2026.xls"
ASSESSMENT = ORNEK / "energo" / "ASSESSMENT_YANSITMA_2026_05_06.xlsx"
ARABULUCU = ORNEK / "energo" / "ARABULUCULUK_2026_06_07.xlsx"
SAGLIK = ORNEK / "energo" / "SAGLIK_KONTROL_LISTE.xlsx"


class SentetikDosyaTemeli(unittest.TestCase):
    """Gecici dizinde openpyxl ile kucuk sentetik calisma kitaplari uretir."""

    def setUp(self):
        self._dizin = tempfile.TemporaryDirectory()
        self.kok = Path(self._dizin.name)

    def tearDown(self):
        self._dizin.cleanup()

    def xlsx(self, ad: str, sayfalar: dict[str, list[list]]) -> Path:
        import openpyxl

        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sayfa_adi, satirlar in sayfalar.items():
            ws = wb.create_sheet(sayfa_adi)
            for satir in satirlar:
                ws.append(satir)
        yol = self.kok / ad
        wb.save(yol)
        return yol


# --------------------------------------------------------------------------
# 1) Assessment: sayfa ve baslik ICERIKLE bulunur
# --------------------------------------------------------------------------

def _assessment_sayfalari(kisi_sayfasi="Kişi Listesi", katilimci="Katılımcı", paket="Paket"):
    """Iki katilimcili, iki faturali kucuk assessment yansitmasi.

    TL 'Toplam' 49.500 x 2 = 99.000; USD 1.087,25 x 2; Energo Payi 1.304,70 x 2
    = 2.609,40 USD. Dogru tutar Energo Payi'dir; TL toplam asla dagitilmamali.
    """
    return {
        "Fatura Detay": [
            ["Belge tarihi", "UP cinsinden tutar", "Ulusal para birimi", "Tutar (UPB3)",
             "UPB 3", "Energo Payı", "Metin"],
            [" 22.05.2026", 49500, "TRY", 1087.25, "USD", 1304.7, " ASS2026000009001 - AS Olcme"],
            [" 08.06.2026", 49500, "TRY", 1087.25, "USD", 1304.7, " ASS2026000009002 - AS Olcme"],
            [None, None, None, None, None, 2609.4, None],
        ],
        kisi_sayfasi: [
            ["Tarih", katilimci, "Pozisyon", "Firma Yetkilisi", "Uygulama Türü", "Uygulama Yeri",
             paket, "Fatura Numarası", "Fatura Tarihi", "Toplam", "USD ", "Energo Payı", "Yansıtma"],
            ["16.05.2026", "ALI VELI", "Muhendis", "YETKILI KISI", "Online DM", "Online",
             "MANAGER", "ASS2026000009001", "22.05.2026", 49500, 1087.25, 1304.7, "RHI"],
            ["06.06.2026", "AYSE FATMA", "Uzman", "YETKILI KISI", "Online DM", "Online",
             "MANAGER", "ASS2026000009002", "08.06.2026", 49500, 1087.25, 1304.7, "RSD"],
        ],
    }


class AssessmentIcerikTespitiTest(SentetikDosyaTemeli):
    """Bulgu 1(a): sayfa adi ya da baslik degisse de assessment okuyucusu calisir."""

    def _dogrula(self, yol: Path):
        self.assertEqual(dosya_tipini_bul(yol), "energo_assessment")
        env: list = []
        satirlar = oku(yol, envanter=env)
        self.assertEqual(len(satirlar), 2)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"energo_assessment"})
        self.assertEqual({s.para_birimi for s in satirlar}, {"USD"})
        self.assertAlmostEqual(sum(s.tutar for s in satirlar), 2609.4, places=2)
        self.assertEqual([s.kisi_ham for s in satirlar], ["ALI VELI", "AYSE FATMA"])
        self.assertEqual([s.masraf_merkezi_kaynak for s in satirlar], ["RHI", "RSD"])
        for s in satirlar:
            self.assertNotIn(OKUYUCU_UYARISI, s.ek, "sablon taninmisken uyari olmamali")
            self.assertEqual(s.ek["tutar_yontemi"], "kisi satirindaki Energo Payi")
        self.assertEqual(env[-1].tur, "energo_assessment")
        self.assertAlmostEqual(env[-1].tutar, 2609.4, places=2)
        # Fatura detay sayfasi da icerikle bulunmus, beyan yazilmis olmali.
        self.assertEqual(satirlar[0].ek["fatura_ozeti"], {"ASS2026000009001": 1304.7,
                                                           "ASS2026000009002": 1304.7})

    def test_sablon_oldugu_gibi(self):
        self._dogrula(self.xlsx("Yansitma.xlsx", _assessment_sayfalari()))

    def test_kisi_listesi_sayfasi_liste_olmus(self):
        self._dogrula(self.xlsx("Yansitma.xlsx", _assessment_sayfalari(kisi_sayfasi="Liste")))

    def test_katilimci_basligi_katilimci_adi_olmus(self):
        self._dogrula(self.xlsx("Yansitma.xlsx", _assessment_sayfalari(katilimci="Katılımcı Adı")))

    def test_paket_basligi_program_olmus(self):
        self._dogrula(self.xlsx("Yansitma.xlsx", _assessment_sayfalari(paket="Program")))

    def test_basta_bos_sayfa(self):
        sayfalar = {"Bos": []}
        sayfalar.update(_assessment_sayfalari())
        self._dogrula(self.xlsx("Yansitma.xlsx", sayfalar))

    def test_fatura_detay_sayfasi_kisi_listesine_karismaz(self):
        # Kisi listesi 'Fatura Numarasi' tasir; 'fatura' adayina uyar ama
        # kisi kolonu olmayan gercek detay sayfasi secilmeli.
        satirlar = assessment_oku(self.xlsx("Yansitma.xlsx", _assessment_sayfalari()))
        self.assertIn("fatura_ozeti", satirlar[0].ek)
        self.assertEqual(satirlar[0].ek["beyan_yontemi"],
                         "'Fatura Detay' sayfasindaki fatura bazli Energo Payi")


# --------------------------------------------------------------------------
# 1b / 4 / 6) Ozel okuyucu bos donerse: zincir, uyari, envanter turu
# --------------------------------------------------------------------------

class OzelOkuyucuDususuTest(SentetikDosyaTemeli):
    """Bulgu 1(b), 4, 6: sessiz dusus yok; sonraki adaylar denenir; tur dogru."""

    def _taninmayan_assessment(self) -> Path:
        # Basligi 'Katilimci' + 'Paket' tasidigi icin kesif assessment der;
        # ama sablonun geri kalani yok, assessment okuyucusu bos doner.
        return self.xlsx("Katilimci_Odemeleri.xlsx", {"Sayfa1": [
            ["Katılımcı", "Paket", "Tutar (USD)", "Tarih"],
            ["ALI VELI", "TEMEL", 120.0, "01.07.2026"],
            ["AYSE FATMA", "ILERI", 80.0, "02.07.2026"],
        ]})

    def test_adaylar_oncelik_sirali_ve_genel_ile_biter(self):
        yol = self._taninmayan_assessment()
        adaylar = dosya_tip_adaylari(yol)
        self.assertEqual(adaylar[0], "energo_assessment")
        self.assertEqual(adaylar[-1], "genel")
        self.assertEqual(dosya_tipini_bul(yol), adaylar[0])

    def test_genel_dususu_uyari_tasir(self):
        yol = self._taninmayan_assessment()
        env: list = []
        satirlar = oku(yol, envanter=env)
        self.assertEqual(len(satirlar), 2)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"genel"})
        self.assertAlmostEqual(sum(s.tutar for s in satirlar), 200.0)
        for s in satirlar:
            self.assertIn(OKUYUCU_UYARISI, s.ek)
            self.assertIn("genel okuyucuyla okundu", s.ek[OKUYUCU_UYARISI])
            self.assertIn("elle dogrulanmali", s.ek[OKUYUCU_UYARISI])
            self.assertIn("energo_assessment", s.ek[OKUYUCU_UYARISI])
            self.assertEqual(s.ek["okuyucu_turu"], "energo_assessment -> genel")
        kayit = env[-1]
        self.assertEqual(kayit.tur, "energo_assessment -> genel")
        self.assertIn("elle dogrulanmali", kayit.sebep)
        self.assertEqual(kayit.durum, "OKUNDU")

    def test_okuyucu_turu_mail_eki_icin_satirlardan_turetilir(self):
        # _msg_oku_icerik ek satirlarini oku() ile alir; tur satirlardan
        # cikarilabilmeli (ek['okuyucu_turu'] notu).
        satirlar = oku(self._taninmayan_assessment())
        self.assertEqual(okuyucu_turu(satirlar, "energo_assessment"), "energo_assessment -> genel")

    def test_arabulucu_bedeli_basligi_genel_dosya(self):
        # 'Arabulucu Bedeli' basligi kesifte arabuluculuk ipucu verir; ama
        # sablon (personel / proje / TCKN) yok, arabulucu okuyucusu bos doner.
        # Genel okuyucu tutari okur (100 USD), satir uyari tasir.
        yol = self.xlsx("Hakem.xlsx", {"Sayfa1": [
            ["Ad Soyad", "Arabulucu Bedeli (USD)", "Tarih"],
            ["ALI VELI", 100, "01.07.2026"],
        ]})
        self.assertEqual(dosya_tipini_bul(yol), "energo_arabulucu")
        self.assertEqual(dosya_tip_adaylari(yol), ["energo_arabulucu", "genel"])
        (s,) = oku(yol)
        self.assertEqual(s.kaynak_tip, "genel")
        self.assertEqual(s.tutar, 100.0)
        self.assertEqual(s.para_birimi, "USD")
        self.assertIn(OKUYUCU_UYARISI, s.ek)
        self.assertEqual(s.ek["okuyucu_turu"], "energo_arabulucu -> genel")

    def test_ozel_okuyucudan_ozel_okuyucuya_zincir_uyarisiz(self):
        # Sayfa adi 'Arabulucu Egitimi' arabuluculuk ipucu verir; icerik Koc
        # katilimci listesidir. Arabulucu okuyucusu bos doner, sonraki aday
        # (koc_katilimci) okur; sablon taninmis sayilir, uyari yok.
        yol = self.xlsx("Katilimci.xlsx", {"Arabulucu Egitimi": [
            ["ID", "Ad Soyad", "Pozisyon", "Alt Fonksiyon", "Katılım Tarihi"],
            [629001, "VELI ALI", "Muhendis", "Saha", "01.07.2026"],
        ]})
        self.assertEqual(dosya_tip_adaylari(yol)[:2], ["energo_arabulucu", "koc_katilimci"])
        env: list = []
        (s,) = oku(yol, envanter=env)
        self.assertEqual(s.kaynak_tip, "koc_katilimci")
        self.assertNotIn(OKUYUCU_UYARISI, s.ek)
        self.assertEqual(env[-1].tur, "energo_arabulucu -> koc_katilimci")
        self.assertEqual(env[-1].durum, "KUTUK")

    def test_detay_listesi_turu_envanterde_detay(self):
        yol = self.xlsx("ASS2026000009001 010720261000 ENERGO Fatura Detayı.xlsx", {"Sayfa1": [
            ["Tarih", "Firma", "Katılımcı", "Pozisyon", "Firma Yetkilisi", "Uygulama Türü",
             "Uygulama Yeri", "Paket"],
            ["25.06.2026", "RHI", "ALI VELI", "Muhendis", "YETKILI", "Online DM", "Online", "MANAGER"],
        ]})
        env: list = []
        (s,) = oku(yol, envanter=env)
        self.assertEqual(s.kaynak_tip, "energo_assessment_detay")
        self.assertEqual(env[-1].tur, "energo_assessment_detay")
        self.assertEqual(env[-1].durum, "DETAY LISTESI")
        self.assertNotIn(OKUYUCU_UYARISI, s.ek)

    def test_okuyucu_turu_yardimcisi(self):
        def gider(tip):
            return GiderSatiri(kaynak_dosya="x", kaynak_tip=tip, satir_no=1, belge_tarihi=None,
                               aciklama="", kisi_ham="A B", sicil_ham=None, tckn_ham=None,
                               tutar=None, para_birimi=None, masraf_merkezi_kaynak=None,
                               gider_tipi=None)
        self.assertEqual(okuyucu_turu([gider("energo_assessment_detay")], "energo_assessment"),
                         "energo_assessment_detay")
        self.assertEqual(okuyucu_turu([gider("genel")], "energo_assessment", "genel"),
                         "energo_assessment -> genel")
        self.assertEqual(okuyucu_turu([gider("genel")], "genel", "genel"), "genel")
        self.assertEqual(okuyucu_turu([], "energo_saglik", "genel"), "energo_saglik -> genel")
        self.assertEqual(okuyucu_turu([], "genel"), "genel")


# --------------------------------------------------------------------------
# 4) Kural 4: veri hucresindeki 'ARABULUCU' dosya tipini degistirmez
# --------------------------------------------------------------------------

class SaglikListesiGorevArabulucuTest(SentetikDosyaTemeli):
    def _saglik(self, gorev: str) -> Path:
        return self.xlsx("SAGLIK KONTROL LISTE.xlsx", {
            "BORDROLU LISTE": [
                ["S.NO", "ADI SOYADI", "TCKN", "DOĞUM TARİHİ", "ÜLKE", "GÖREVİ", "ŞANTİYE",
                 "FİRMA & EKİP & FORMEN", "TALEP TARIHI", "SAĞLIK KONTROL TARİHİ"],
                [1, "ALI VELI", "10000000146", "01.01.1990", "RUSYA", gorev,
                 "Ust Luga GPP projesi", "RHI", "01.07.2026", "05.07.2026"],
                [2, "AYSE FATMA", "10000000278", "02.02.1991", "RUSYA", "Kaynakci",
                 "Ust Luga GPP projesi", "RHI", "01.07.2026", "05.07.2026"],
            ],
            "BORDROSUZ LİSTE": [
                ["S NO", "ADI SOYADI", "TC KIMLIK NO", "DOGUM TARIHI", "SANTIYE", "GOREVI"],
                [1, "CAN DELI", "10000000401", "03.03.1992", "Ust Luga GPP projesi", "Usta"],
            ],
        })

    def test_gorev_hucresi_arabulucu_olsa_da_saglik(self):
        yol = self._saglik("ARABULUCU")
        self.assertEqual(dosya_tipini_bul(yol), "energo_saglik")
        self.assertEqual(dosya_tipini_bul(self._saglik("Kaynakci")), "energo_saglik")
        env: list = []
        satirlar = oku(yol, envanter=env)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"energo_saglik"})
        self.assertEqual({s.gider_tipi for s in satirlar}, {"Saglik"})
        self.assertEqual(len(satirlar), 2)  # bordrolu sayfa
        self.assertEqual(env[-1].durum, "KUTUK")
        self.assertEqual(env[-1].tur, "energo_saglik")

    def test_baslik_hucresindeki_arabulucu_ipucu_sayilir(self):
        # Ipucu basliktan gelir: kisi listesi sablonu taninmali.
        yol = self.xlsx("Arabuluculuk.xlsx", _arabulucu_sayfalari())
        self.assertEqual(dosya_tipini_bul(yol), "energo_arabulucu")

    def test_assessment_pozisyonu_arabulucu_olsa_da_assessment(self):
        sayfalar = _assessment_sayfalari()
        sayfalar["Kişi Listesi"][1][2] = "Arabulucu"
        yol = self.xlsx("Yansitma.xlsx", sayfalar)
        self.assertEqual(dosya_tipini_bul(yol), "energo_assessment")
        self.assertNotIn("energo_arabulucu", dosya_tip_adaylari(yol))


# --------------------------------------------------------------------------
# 3) Kural 7: tutar kolonu varsa kutuk degil gider dosyasi
# --------------------------------------------------------------------------

class Kural7TutarKolonuTest(SentetikDosyaTemeli):
    def test_tutar_kolonu_ve_sayisal_deger_varsa_gider(self):
        yol = self.xlsx("Sigorta_Primi.xlsx", {"Sayfa1": [
            ["Sicil No", "Ad Soyad", "Masraf Merkezi", "Tutar (USD)", "Tarih"],
            *[[629001 + i, f"KISI {i} VELI", "GPP Project", 100.0 + i, "01.07.2026"] for i in range(5)],
        ]})
        self.assertEqual(dosya_tipini_bul(yol), "genel")
        env: list = []
        satirlar = oku(yol, envanter=env)
        self.assertEqual(len(satirlar), 5)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"genel"})
        self.assertAlmostEqual(sum(s.tutar for s in satirlar), 510.0)
        self.assertEqual({s.para_birimi for s in satirlar}, {"USD"})
        self.assertFalse(any(s.ek.get("referans_liste") for s in satirlar))
        self.assertEqual(env[-1].durum, "OKUNDU")
        self.assertAlmostEqual(env[-1].tutar, 510.0)

    def test_tutar_kolonu_yoksa_kutuk_ve_sebep_yazilir(self):
        yol = self.xlsx("Ferdi_Kaza_Listesi.xlsx", {"Sayfa1": [
            ["Sicil No", "Ad Soyad", "Masraf Merkezi", "Dogum Tarihi"],
            *[[629001 + i, f"KISI {i} VELI", "GPP Project", "01.01.1990"] for i in range(5)],
        ]})
        self.assertEqual(dosya_tipini_bul(yol), "referans_liste")
        env: list = []
        satirlar = oku(yol, envanter=env)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"referans_liste"})
        for s in satirlar:
            self.assertTrue(s.ek["referans_liste"])
            self.assertIn("tutar kolonu yok", s.ek["kutuk_sebebi"])
        self.assertEqual(env[-1].durum, "KUTUK")
        self.assertIn("tutar kolonu yok", env[-1].sebep)
        self.assertNotIn("kolon_esanlamlilari", env[-1].sebep)

    def test_tutar_basligi_var_ama_hucreler_bos_ise_kutuk(self):
        yol = self.xlsx("Liste.xlsx", {"Sayfa1": [
            ["Sicil No", "Ad Soyad", "Masraf Merkezi", "Tutar"],
            *[[629001 + i, f"KISI {i} VELI", "GPP Project", None] for i in range(5)],
        ]})
        self.assertEqual(dosya_tipini_bul(yol), "referans_liste")

    def test_cok_satirli_tutarsiz_genel_dosya_kutuk_sebebiyle(self):
        yol = self.xlsx("Liste.xlsx", {"Sayfa1": [
            ["Ad Soyad", "Sicil", "Görev Yeri"],
            *[[f"KISI {i} VELI", 600000 + i, "GPP"] for i in range(kesif._KUTUK_SATIR_ESIGI)],
        ]})
        self.assertEqual(dosya_tipini_bul(yol), "genel")
        env: list = []
        satirlar = oku(yol, envanter=env)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"referans_liste"})
        self.assertIn("hicbirinde tutar yok", env[-1].sebep)
        self.assertEqual(env[-1].durum, "KUTUK")


# --------------------------------------------------------------------------
# 5) Arabuluculuk: fatura detay sayfasi/kolonlari icerikle bulunur
# --------------------------------------------------------------------------

def _arabulucu_sayfalari(kisi_sayfasi="Kişi Listesi", detay_sayfasi="Fatura Detay",
                         yer="Masraf yeri", pay="Energo Payı", detay_var=True,
                         kisiler=None):
    """Uc kisilik arabuluculuk yansitmasi: RHI 2 kisi (10,50), RC 1 kisi (21,00)."""
    kisiler = kisiler or [("ALI VELI", "RHI"), ("AYSE FATMA", "RHI"), ("CAN DELI", "RC")]
    sayfalar: dict[str, list[list]] = {}
    if detay_var:
        sayfalar[detay_sayfasi] = [
            ["Kayıt tarihi", "Fatura TutarI ( TL )", "Fatura Tutarı ( USD )", pay, yer, "Metin"],
            ["01.07.2026", 1000, 30, 10.5, "RHI", "EF02026000000150 - A P"],
            ["01.07.2026", 2000, 60, 21.0, "RC", "EF02026000000150 - A P"],
            [None, None, None, 31.5, None, None],
        ]
    sayfalar[kisi_sayfasi] = [
        ["TARİH", "YETKİLİ", "PERSONEL ", "PERSONEL T.C.", "PROJE", "İLGİLİ ŞİRKET",
         "ARABULUCU", "Fatura No", " Fatura Tarihi"],
        *[["15.06.2026", "YETKILI KISI", ad, f"1000000{i}146", "GPP Project", sirket,
           "ARABULUCU KISI", "EF02026000000150", "01.07.2026"]
          for i, (ad, sirket) in enumerate(kisiler)],
    ]
    return sayfalar


class ArabulucuDetaySayfasiTest(SentetikDosyaTemeli):
    def _dogrula(self, yol: Path):
        self.assertEqual(dosya_tipini_bul(yol), "energo_arabulucu")
        satirlar = oku(yol)
        self.assertEqual(len(satirlar), 3)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"energo_arabulucu"})
        self.assertEqual([s.tutar for s in satirlar], [5.25, 5.25, 21.0])
        for s in satirlar:
            self.assertNotIn(OKUYUCU_UYARISI, s.ek)
            self.assertIn("masraf yeri toplami", s.ek["tutar_yontemi"])

    def test_sablon_oldugu_gibi(self):
        self._dogrula(self.xlsx("Arabuluculuk.xlsx", _arabulucu_sayfalari()))

    def test_fatura_detay_sayfasi_ozet_olmus(self):
        self._dogrula(self.xlsx("Arabuluculuk.xlsx", _arabulucu_sayfalari(detay_sayfasi="Özet")))

    def test_detay_kolonlari_sirket_ve_pay_usd(self):
        self._dogrula(self.xlsx("Arabuluculuk.xlsx",
                                _arabulucu_sayfalari(yer="Şirket", pay="Pay (USD)")))

    def test_kisi_listesi_sayfasi_personel_olmus(self):
        self._dogrula(self.xlsx("Arabuluculuk.xlsx", _arabulucu_sayfalari(kisi_sayfasi="Personel")))

    def test_detay_sayfasi_yoksa_gercek_sebep(self):
        yol = self.xlsx("Arabuluculuk.xlsx", _arabulucu_sayfalari(detay_var=False))
        satirlar = arabulucu_oku(yol)
        self.assertEqual(len(satirlar), 3)
        for s in satirlar:
            self.assertIsNone(s.tutar)
            self.assertTrue(s.ek["tutar_yontemi"].startswith("fatura detay sayfasi bulunamadi"))
            self.assertIn("beklenen kolonlar", s.ek["tutar_yontemi"])
            self.assertIn("Masraf yeri", s.ek["tutar_yontemi"])
            self.assertIn("Energo Payi", s.ek["tutar_yontemi"])
            self.assertNotIn("eslesen masraf yeri yok", s.ek["tutar_yontemi"])

    def test_detay_tutarlari_bos_ise_gercek_sebep(self):
        sayfalar = _arabulucu_sayfalari()
        for satir in sayfalar["Fatura Detay"][1:]:
            satir[3] = None
        satirlar = arabulucu_oku(self.xlsx("Arabuluculuk.xlsx", sayfalar))
        for s in satirlar:
            self.assertIsNone(s.tutar)
            self.assertIn("bulundu ama hicbir satirda", s.ek["tutar_yontemi"])

    def test_eslesmeyen_masraf_yeri_sebebi_detaydaki_yerleri_sayar(self):
        kisiler = [("ALI VELI", "RHI"), ("AYSE FATMA", "RC"), ("CAN DELI", "XYZ")]
        satirlar = arabulucu_oku(self.xlsx("Arabuluculuk.xlsx", _arabulucu_sayfalari(kisiler=kisiler)))
        self.assertEqual([s.tutar for s in satirlar], [10.5, 21.0, None])
        yontem = satirlar[2].ek["tutar_yontemi"]
        self.assertIn("eslesen masraf yeri yok", yontem)
        self.assertIn("'XYZ'", yontem)
        self.assertIn("RC, RHI", yontem)


# --------------------------------------------------------------------------
# 2) Antik: aciklamadaki 'TOPLAM' satiri toplam yapmaz
# --------------------------------------------------------------------------

def _antik_satirlari(toplam_etiketi_kolonu=9, toplam_etiketi="TOPLAM"):
    """Antik cari dokumunun kucuk bir kopyasi (3 veri + 1 toplam satiri)."""
    bos = [""] * 13
    baslik = ["İşlem Tarihi", "", "", "İşlem", "", "Evrak No", "Açıklama", "", "", "", "Döviz", "Borç", "Alacak"]
    toplam = list(bos)
    toplam[toplam_etiketi_kolonu] = toplam_etiketi
    toplam[10], toplam[11] = "USD", 350.0
    return [
        list(bos),
        ["", "", "", "", "", "", "", "", "Cari Hareket Dökümü Detayı", "", "", "", ""],
        baslik,
        ["ENERGO-USD-ENERGO-USD", "", "", "", "", "", "", "", "", "", "", "", ""],
        [date(2026, 7, 1), "", "", "Bilet İşlem", "", "TK 1234567890",
         "TK1234567890 VELI/ALI MR  IST-LED BILET BEDELI", "", "", "", "USD", 100.0, None],
        [date(2026, 7, 2), "", "", "Otel İşlemleri", "", "OT 1",
         "ALI VELI ; AYSE FATMA GRAND HOTEL [10.07.2026] - [11.07.2026]  (1) KONAKLAMA YURTICI (TOPLAM 2 KISI)",
         "", "", "", "USD", 250.0, None],
        [date(2026, 7, 3), "", "", "Bilet İşlem", "", "TK 1234567891",
         "TK1234567891 VELI/ALI MR  IST-LED BILET IADE", "", "", "", "USD", None, 16.0],
        toplam,
    ]


class AntikToplamSatiriTest(SentetikDosyaTemeli):
    def _dogrula(self, satirlar):
        self.assertEqual(len(satirlar), 3, [s.aciklama for s in satirlar])
        self.assertAlmostEqual(sum(s.tutar for s in satirlar), 334.0)
        otel = [s for s in satirlar if "(TOPLAM 2 KISI)" in s.aciklama]
        self.assertEqual(len(otel), 1)
        self.assertEqual(otel[0].tutar, 250.0)
        self.assertEqual(otel[0].ek["kisiler"], ["ALI VELI", "AYSE FATMA"])
        self.assertEqual(satirlar[0].ek["fatura_ozeti"],
                         {"TOPLAM (borc)": 350.0, "okunan alacak": -16.0})

    def test_aciklamadaki_toplam_veri_satirini_silmez(self):
        yol = self.xlsx("ENERGO CARI.xlsx", {"Sayfa1": _antik_satirlari()})
        self.assertEqual(dosya_tipini_bul(yol), "antik_cari")
        self._dogrula(antik_cari_oku(yol))

    def test_toplam_etiketi_aciklama_kolonunda(self):
        yol = self.xlsx("ENERGO CARI.xlsx",
                        {"Sayfa1": _antik_satirlari(toplam_etiketi_kolonu=6, toplam_etiketi="GENEL TOPLAM")})
        self._dogrula(antik_cari_oku(yol))

    def test_bakiye_etiketi_de_toplam_satiridir(self):
        yol = self.xlsx("ENERGO CARI.xlsx", {"Sayfa1": _antik_satirlari(toplam_etiketi="BAKIYE")})
        self._dogrula(antik_cari_oku(yol))

    def test_tarih_tasiyan_satir_toplam_sayilmaz(self):
        # Aciklamasi yalnizca 'TOPLAM' olsa bile tarih/evrak varsa veri satiridir.
        satirlar = _antik_satirlari()
        satirlar.insert(7, [date(2026, 7, 4), "", "", "Diğer Hizmetler", "", "DH 1",
                            "TOPLAM", "", "", "", "USD", 5.0, None])
        okunan = antik_cari_oku(self.xlsx("ENERGO CARI.xlsx", {"Sayfa1": satirlar}))
        self.assertEqual(len(okunan), 4)
        self.assertEqual(okunan[0].ek["fatura_ozeti"]["TOPLAM (borc)"], 350.0)


# --------------------------------------------------------------------------
# 7) IKRAM: kisi adi olabilen hizmet kelimesi
# --------------------------------------------------------------------------

class IkramAdiTest(unittest.TestCase):
    def test_ikram_kosulsuz_hizmet_kelimesi_degil(self):
        self.assertNotIn("IKRAM", _HIZMET_KELIMELERI)

    def test_kuyruk_kirpma(self):
        self.assertEqual(_kuyruk_kirp("IKRAM YILMAZ"), "IKRAM YILMAZ")
        self.assertEqual(_kuyruk_kirp("ALI VELI IKRAM"), "ALI VELI IKRAM")
        self.assertEqual(_kuyruk_kirp("ALI VELI IKRAM BEDELI"), "ALI VELI")
        self.assertEqual(_kuyruk_kirp("ALI VELI IKRAM HIZMETI"), "ALI VELI")
        # Kesme oncesi en az iki ad tokeni yoksa IKRAM'da kesilmez.
        self.assertEqual(_kuyruk_kirp("IKRAM BEDELI"), "IKRAM")
        self.assertEqual(_kuyruk_kirp("ALI IKRAM BEDELI"), "ALI IKRAM")
        # Diger hizmet kelimeleri eskisi gibi kosulsuz keser.
        self.assertEqual(_kuyruk_kirp("OMER CAN CETIR KONSOLOSLUK UCRETI"), "OMER CAN CETIR")

    def test_kisi_cikarma_yollari(self):
        self.assertEqual(_temizle_ve_dogrula("IKRAM YILMAZ"), "IKRAM YILMAZ")
        self.assertEqual(_vize_kisi("IKRAM YILMAZ RUSYA FEDERASYONU TURISTIK E-VIZE"), "IKRAM YILMAZ")
        self.assertEqual(_bilet_kisi("TK1234567890 YILMAZ/IKRAM MR  IST-LED BILET BEDELI"), "YILMAZ IKRAM")
        self.assertEqual(_diger_kisi("ALI VELI IKRAM BEDELI [10.07.2026] TOPLANTI"), "ALI VELI")
        self.assertIsNone(_diger_kisi("IKRAM BEDELI [10.07.2026] TOPLANTI"))


# --------------------------------------------------------------------------
# Gercek ornek veri (varsa): davranis korunmus olmali
# --------------------------------------------------------------------------

def _veri_gerek(test: unittest.TestCase, *yollar: Path) -> None:
    for yol in yollar:
        if not yol.exists():
            test.skipTest(f"Ornek veri bulunamadi: {yol}")


class GercekVeriKorunmasiTest(unittest.TestCase):
    def test_antik_satir_net_ve_beyan(self):
        _veri_gerek(self, ANTIK)
        satirlar = antik_cari_oku(ANTIK)
        self.assertEqual(len(satirlar), 134)
        self.assertAlmostEqual(sum(s.tutar for s in satirlar if s.tutar is not None), 48946.59, places=2)
        self.assertAlmostEqual(satirlar[0].ek["fatura_ozeti"]["TOPLAM (borc)"], 48962.59, places=2)

    def test_assessment_uyarisiz_ve_usd(self):
        _veri_gerek(self, ASSESSMENT)
        env: list = []
        satirlar = oku(ASSESSMENT, envanter=env)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"energo_assessment"})
        self.assertEqual({s.para_birimi for s in satirlar}, {"USD"})
        self.assertFalse(any(OKUYUCU_UYARISI in s.ek for s in satirlar))
        self.assertEqual(env[-1].tur, "energo_assessment")

    def test_arabulucu_her_satirda_tutar(self):
        _veri_gerek(self, ARABULUCU)
        satirlar = oku(ARABULUCU)
        self.assertEqual({s.kaynak_tip for s in satirlar}, {"energo_arabulucu"})
        self.assertTrue(all(s.tutar is not None for s in satirlar))

    def test_saglik_kutuk(self):
        _veri_gerek(self, SAGLIK)
        self.assertEqual(dosya_tipini_bul(SAGLIK), "energo_saglik")
        env: list = []
        oku(SAGLIK, envanter=env)
        self.assertEqual(env[-1].durum, "KUTUK")


if __name__ == "__main__":
    unittest.main()
