"""Ucuncu tur inceleme bulgularinin testleri (genel okuyucu, kolon sozlugu,
sayi/tarih cozumu, kutuk beslemesi, Yuzyil paylasim deseni, arayuz yukleme
klasoru temizligi).

Tum dosyalar openpyxl ile uretilen kucuk sentetik calisma kitaplaridir;
gercek kisi adi, TC kimlik numarasi ya da gercek veri icermez.
"""

from __future__ import annotations

import re
import sys
import tempfile
import types
import unittest
from datetime import date
from pathlib import Path

KOK = Path(__file__).resolve().parents[1]
if str(KOK) not in sys.path:
    sys.path.insert(0, str(KOK))

import openpyxl  # noqa: E402

from masraf.defter import Defterler, besleme_aciklamasi  # noqa: E402
from masraf.modeller import GiderSatiri  # noqa: E402
from masraf.okuyucular.genel import (  # noqa: E402
    BINLIK_VARSAYILDI,
    genel_oku,
    hucre_sayisi,
    hucre_tarihi,
    kisi_anahtari,
    kolon_ara,
    kolon_haritasi,
    sayi_coz,
    sayfa_tekrarlarini_isaretle,
    sicil_bicimi_mi,
)


def _kitap_yaz(yol: Path, sayfalar: dict[str, list[list]]) -> Path:
    """{sayfa adi: satirlar} sozlugunu xlsx olarak yazar."""
    kitap = openpyxl.Workbook()
    kitap.remove(kitap.active)
    for ad, satirlar in sayfalar.items():
        sayfa = kitap.create_sheet(ad)
        for satir in satirlar:
            sayfa.append(satir)
    kitap.save(yol)
    return yol


def _gider(kaynak_tip: str, kisi: str | None, sicil: str | None = None,
           tckn: str | None = None, sayfa: str = "Data", satir_no: int = 2,
           merkez: str | None = "GPP Project") -> GiderSatiri:
    return GiderSatiri(
        kaynak_dosya="liste.xlsx", kaynak_tip=kaynak_tip, satir_no=satir_no,
        belge_tarihi=None, aciklama=kisi or "", kisi_ham=kisi, sicil_ham=sicil,
        tckn_ham=tckn, tutar=None, para_birimi=None, masraf_merkezi_kaynak=merkez,
        gider_tipi="Diger", ek={"sayfa": sayfa},
    )


# --------------------------------------------------------------------------
# Bulgu 3: Turkce sayi yazimi
# --------------------------------------------------------------------------

class SayiCozTest(unittest.TestCase):
    """Metin hucrelerindeki Turkce/Ingilizce tutar yazimlari."""

    def test_tek_nokta_uc_hane_binlik_ve_not(self):
        # Turkce tedarikci dosyasinda '2.500' iki bin bes yuzdur; karar not dusulur.
        self.assertEqual(sayi_coz("2.500"), (2500.0, BINLIK_VARSAYILDI))
        self.assertEqual(sayi_coz("12.500 TL"), (12500.0, BINLIK_VARSAYILDI))
        self.assertEqual(hucre_sayisi("2.500"), 2500.0)

    def test_coklu_nokta_binlik(self):
        self.assertEqual(sayi_coz("1.234.567"), (1234567.0, None))

    def test_ondalik_kalir(self):
        # Sifirla baslayan ya da 3 hane olmayan kesirler ondaliktir, not yok.
        self.assertEqual(sayi_coz("0.500"), (0.5, None))
        self.assertEqual(sayi_coz("12.5"), (12.5, None))
        self.assertEqual(sayi_coz("1234.500"), (1234.5, None))

    def test_virgul_ondalik_ve_ingilizce(self):
        self.assertEqual(sayi_coz("12,5"), (12.5, None))
        self.assertEqual(sayi_coz("1.234,56"), (1234.56, None))
        self.assertEqual(sayi_coz("1,234.56"), (1234.56, None))
        self.assertEqual(sayi_coz("1,234,567.89"), (1234567.89, None))
        self.assertEqual(sayi_coz("36 000,00"), (36000.0, None))

    def test_parantez_ve_sondaki_eksi_negatif(self):
        self.assertEqual(sayi_coz("(100)"), (-100.0, None))
        self.assertEqual(sayi_coz("(1.234,56)"), (-1234.56, None))
        self.assertEqual(sayi_coz("100-"), (-100.0, None))
        self.assertEqual(sayi_coz("-1.234,56"), (-1234.56, None))

    def test_bilimsel_gosterim_dogru_cozulur(self):
        # Eski surum 'e' harfini yutup 1.53 uretiyordu.
        self.assertEqual(sayi_coz("1.5e3"), (1500.0, None))
        self.assertEqual(sayi_coz("2,5E+02"), (250.0, None))

    def test_tarih_gibi_metinler_sayi_degil(self):
        self.assertIsNone(hucre_sayisi("15.07.2026"))
        self.assertIsNone(hucre_sayisi("2026-07-15"))
        self.assertIsNone(hucre_sayisi("Tutar"))
        self.assertIsNone(hucre_sayisi("--"))

    def test_sayisal_hucreye_dokunulmaz(self):
        self.assertEqual(sayi_coz(2.5), (2.5, None))
        self.assertEqual(sayi_coz(2500), (2500.0, None))
        self.assertEqual(sayi_coz(None), (None, None))
        self.assertEqual(sayi_coz(True), (None, None))

    def test_genel_oku_binlik_notunu_eke_yazar(self):
        with tempfile.TemporaryDirectory() as d:
            yol = _kitap_yaz(Path(d) / "tutar.xlsx", {"Sheet": [
                ["Ad Soyad", "Tutar"],
                ["Aaaa Bbbb", "2.500"],
                ["Cccc Dddd", 2.5],
                ["Eeee Ffff", "1.234,50"],
            ]})
            satirlar = genel_oku(yol)
        self.assertEqual([s.tutar for s in satirlar], [2500.0, 2.5, 1234.5])
        self.assertTrue(satirlar[0].ek.get(BINLIK_VARSAYILDI))
        self.assertNotIn(BINLIK_VARSAYILDI, satirlar[1].ek)
        self.assertNotIn(BINLIK_VARSAYILDI, satirlar[2].ek)


# --------------------------------------------------------------------------
# Bulgu 6: tarih cozumu
# --------------------------------------------------------------------------

class HucreTarihiTest(unittest.TestCase):
    def test_tek_basina_yil_tarih_degil(self):
        # Eski surum 2026'yi Excel seri numarasi sanip 1905-07-18 uretiyordu.
        self.assertIsNone(hucre_tarihi(2026))
        self.assertIsNone(hucre_tarihi(2026.0))
        self.assertIsNone(hucre_tarihi("2026"))
        self.assertIsNone(hucre_tarihi(1900))
        self.assertIsNone(hucre_tarihi(2100))

    def test_excel_seri_numarasi_hala_cozulur(self):
        self.assertEqual(hucre_tarihi(45000), date(2023, 3, 15))
        self.assertEqual(hucre_tarihi("45000"), date(2023, 3, 15))

    def test_saatli_bicimler(self):
        self.assertEqual(hucre_tarihi("15.07.2026 14:30"), date(2026, 7, 15))
        self.assertEqual(hucre_tarihi("15/07/2026 09:05:00"), date(2026, 7, 15))
        self.assertEqual(hucre_tarihi("2026-07-15 00:00:00"), date(2026, 7, 15))

    def test_iso_t_bicimi(self):
        self.assertEqual(hucre_tarihi("2026-07-15T00:00:00"), date(2026, 7, 15))
        self.assertEqual(hucre_tarihi("2026-07-15T14:30:00.000Z"), date(2026, 7, 15))

    def test_ay_gun_belirsizligi(self):
        # Gun 12'den buyukse ay/gun (Amerikan) okunur; degilse gun/ay onceliklidir.
        self.assertEqual(hucre_tarihi("07/15/2026"), date(2026, 7, 15))
        self.assertEqual(hucre_tarihi("05/07/2026"), date(2026, 7, 5))

    def test_mevcut_bicimler_korunur(self):
        self.assertEqual(hucre_tarihi("22.05.2026"), date(2026, 5, 22))
        self.assertEqual(hucre_tarihi("2026-05-22"), date(2026, 5, 22))
        self.assertEqual(hucre_tarihi("1.7.26"), date(2026, 7, 1))
        self.assertIsNone(hucre_tarihi("31.06.2026"))
        self.assertIsNone(hucre_tarihi("Tarih"))


# --------------------------------------------------------------------------
# Bulgu 1: kolon_ara kisa anahtarlar
# --------------------------------------------------------------------------

class KolonAraKisaAnahtarTest(unittest.TestCase):
    SICIL = ("sicil", "sicil no", "personel no", "id", "employee id", "tabel", "kod")
    TCKN = ("tckn", "tc kimlik no", "tc kimlik", "kimlik no", "tc no", "tc")

    def _sec(self, basliklar, adaylar, **ek):
        i = kolon_ara(kolon_haritasi(basliklar), *adaylar, **ek)
        return None if i is None else basliklar[i]

    def test_id_alt_dizi_olarak_eslesmez(self):
        self.assertIsNone(self._sec(["Yolcu", "Provider", "Paid", "Valid Until"], self.SICIL))

    def test_tc_batch_no_ile_eslesmez(self):
        self.assertIsNone(self._sec(["Ad Soyad", "Batch No", "Fatura Tutarı"], self.TCKN))

    def test_kod_bilesikleri_sicil_degil(self):
        for baslik in ("Proje Kodu", "Posta Kodu", "Masraf Merkezi Kodu", "Invoice ID", "Fatura No"):
            self.assertIsNone(self._sec(["Ad Soyad", baslik], self.SICIL), baslik)

    def test_kelime_olarak_gecen_kisa_anahtar_eslesir(self):
        self.assertEqual(self._sec(["Ad Soyad", "Personel ID"], self.SICIL), "Personel ID")
        self.assertEqual(self._sec(["Ad Soyad", "Sicil No"], self.SICIL), "Sicil No")
        self.assertEqual(self._sec(["Ad Soyad", "Personel Kodu"], self.SICIL), "Personel Kodu")
        self.assertEqual(self._sec(["Ad Soyad", "ID"], self.SICIL), "ID")
        self.assertEqual(self._sec(["Personel", "Personel TC"], self.TCKN), "Personel TC")
        self.assertEqual(self._sec(["Personel", "TC Kimlik"], self.TCKN), "TC Kimlik")

    def test_uzun_adaylar_alt_dizi_olarak_eslesmeye_devam_eder(self):
        self.assertEqual(self._sec(["Fatura Tutarı (USD)"], ("tutar",)), "Fatura Tutarı (USD)")
        self.assertEqual(self._sec(["Energo Payı (USD)"], ("pay",)), "Energo Payı (USD)")
        self.assertEqual(self._sec(["Toplam USD"], ("usd",)), "Toplam USD")

    def test_haric_listesi_korunur(self):
        basliklar = ["Doğum Tarihi", "Fatura Tarihi"]
        self.assertEqual(self._sec(basliklar, ("tarih",), haric=("dogum",)), "Fatura Tarihi")

    def test_baslik_listesi_de_kabul_edilir(self):
        # Bulgu 2: tanilama liste veriyordu; artik liste de harita gibi calisir.
        self.assertEqual(kolon_ara(["Ad Soyad", "Tutar"], "tutar"), 1)


class SicilKolonuDegerDogrulamaTest(unittest.TestCase):
    def test_sicil_bicimi(self):
        self.assertTrue(sicil_bicimi_mi(632481))
        self.assertTrue(sicil_bicimi_mi("632481"))
        self.assertTrue(sicil_bicimi_mi(632481.0))
        self.assertTrue(sicil_bicimi_mi("RHI-1234"))
        self.assertFalse(sicil_bicimi_mi("GPP Project"))
        self.assertFalse(sicil_bicimi_mi("15.07.2026"))
        self.assertFalse(sicil_bicimi_mi("12345678901"))  # 11 hane: TCKN uzunlugu
        self.assertFalse(sicil_bicimi_mi(None))

    def test_degerleri_sicil_olmayan_kolon_secilmez(self):
        with tempfile.TemporaryDirectory() as d:
            yol = _kitap_yaz(Path(d) / "kodu.xlsx", {"Sheet": [
                ["Adı Soyadı", "Kodu", "Ücret"],
                ["Aaaa Bbbb", "Standart Paket", 100],
                ["Cccc Dddd", "Premium Paket", 200],
                ["Eeee Ffff", "Standart Paket", 100],
            ]})
            notlar: list[str] = []
            satirlar = genel_oku(yol, notlar=notlar)
        self.assertEqual(len(satirlar), 3)
        self.assertTrue(all(s.sicil_ham is None for s in satirlar))
        self.assertEqual(satirlar[0].ek.get("sicil_kolonu_reddedildi"), "kodu")
        self.assertTrue(any("sicil bicimine uymuyor" in n for n in notlar))

    def test_degerleri_sicil_olan_kolon_secilir(self):
        with tempfile.TemporaryDirectory() as d:
            yol = _kitap_yaz(Path(d) / "sicil.xlsx", {"Sheet": [
                ["Ad Soyad", "Sicil", "Tutar"],
                ["Aaaa Bbbb", 600001, 10],
                ["Cccc Dddd", 600002, 20],
            ]})
            satirlar = genel_oku(yol)
        self.assertEqual([s.sicil_ham for s in satirlar], ["600001", "600002"])
        self.assertNotIn("sicil_kolonu_reddedildi", satirlar[0].ek)


# --------------------------------------------------------------------------
# Ek not 1: 'Ucreti' hizmet bedeli tutardir, 'Ucret' / 'Brut Ucret' maas degildir
# --------------------------------------------------------------------------

class UcretKolonuTest(unittest.TestCase):
    def test_hizmet_ucreti_tutar_sayilir(self):
        with tempfile.TemporaryDirectory() as d:
            yol = _kitap_yaz(Path(d) / "arabulucu.xlsx", {"Sheet": [
                ["Ad Soyad", "Arabulucu Ücreti", "Tarih"],
                ["Aaaa Bbbb", 1500, "01.07.2026"],
                ["Cccc Dddd", "2.500", "01.07.2026"],
            ]})
            satirlar = genel_oku(yol)
        self.assertEqual([s.tutar for s in satirlar], [1500.0, 2500.0])

    def test_kutukteki_maas_kolonu_tutar_sayilmaz(self):
        # 250 satirlik personel kutugu; 'Brut Ucret' ve 'Ucret' maas kolonlaridir.
        with tempfile.TemporaryDirectory() as d:
            satirlar_xlsx = [["Sicil No", "Ad Soyad", "Masraf Merkezi", "Brüt Ücret", "Ücret"]]
            satirlar_xlsx += [[600000 + i, f"Kisi{i:03d} Soyad", "GPP Project", 50000 + i, 40000 + i]
                              for i in range(250)]
            yol = _kitap_yaz(Path(d) / "kutuk.xlsx", {"Liste": satirlar_xlsx})
            satirlar = genel_oku(yol)
        self.assertEqual(len(satirlar), 250)
        self.assertTrue(all(s.tutar is None for s in satirlar))
        self.assertIsNone(satirlar[0].ek["cozulen_kolonlar"]["tutar"])

    def test_kesif_kutugu_fatura_sanmaz(self):
        # Ayni kutuk kesiften gecince referans_liste kalir (tutar yok).
        try:
            from masraf.okuyucular.kesif import oku
        except ImportError as e:
            self.skipTest(f"kesif yok: {e}")
        with tempfile.TemporaryDirectory() as d:
            satirlar_xlsx = [["Ad Soyad", "Görev Yeri", "Ücret"]]
            satirlar_xlsx += [[f"Kisi{i:03d} Soyad", "GPP", 40000 + i] for i in range(250)]
            yol = _kitap_yaz(Path(d) / "maas.xlsx", {"Liste": satirlar_xlsx})
            satirlar = oku(yol)
        self.assertTrue(satirlar)
        self.assertTrue(all(s.kaynak_tip == "referans_liste" for s in satirlar))


# --------------------------------------------------------------------------
# Bulgu 2: kolon_sozlugu.tanilama
# --------------------------------------------------------------------------

class TanilamaTest(unittest.TestCase):
    def test_bulunan_alanlar_dolu(self):
        from masraf.kolon_sozlugu import tanilama

        with tempfile.TemporaryDirectory() as d:
            yol = _kitap_yaz(Path(d) / "fatura.xlsx", {"Sheet": [
                ["Ad Soyad", "Tutar (USD)", "Tarih", "Proje Kodu", "Company Name"],
                ["Aaaa Bbbb", 10, "01.07.2026", "GPP", "Zzzz Ltd"],
            ]})
            sonuc = tanilama(yol, veri_dizini=Path(d) / "veri")
        self.assertTrue(sonuc["islenebilir"], sonuc["mesaj"])
        self.assertEqual(sonuc["bulunan"].get("kisi"), "Ad Soyad")
        self.assertEqual(sonuc["bulunan"].get("tutar"), "Tutar (USD)")
        self.assertEqual(sonuc["bulunan"].get("tarih"), "Tarih")
        # 'Proje Kodu' sicil degil; 'Company Name' kisi degil.
        self.assertNotIn("sicil", sonuc["bulunan"])
        self.assertIn("sicil", sonuc["eksik"])

    def test_kisi_kolonu_yoksa_islenemez(self):
        from masraf.kolon_sozlugu import tanilama

        with tempfile.TemporaryDirectory() as d:
            yol = _kitap_yaz(Path(d) / "liste.xlsx", {"Sheet": [
                ["Company Name", "Proje Kodu", "Tutar"],
                ["Zzzz Ltd", "GPP", 10],
            ]})
            sonuc = tanilama(yol, veri_dizini=Path(d) / "veri")
        self.assertFalse(sonuc["islenebilir"])
        self.assertIn("bulunamadi", sonuc["mesaj"])


# --------------------------------------------------------------------------
# Bulgu 4b: pivot sayfalari ve tekrar eden sayfalar
# --------------------------------------------------------------------------

class PivotVeTekrarEdenSayfaTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dizin = tempfile.TemporaryDirectory()
        kisiler = [[600000 + i, f"Kisi{i:02d} Soyad{i:02d}", "GPP Project"] for i in range(30)]
        baslik = ["SIRA NO", "SICIL NO", "PERSONEL", "MASRAF MERKEZI"]
        cls.yol = _kitap_yaz(Path(cls.dizin.name) / "kutuk.xlsx", {
            "Ozet": [
                [None, None, None],
                ["Count of PERSONEL", "Column Labels", None],
                ["Row Labels", "Rusya / Ekspat", "Grand Total"],
                ["GPP Project", 10, 10],
                ["ALNG2-GBS Project", 20, 20],
                ["Grand Total", 30, 30],
            ],
            "Лист1": [
                ["Названия строк", "Количество по полю Категория"],
                ["GPP Project", 30],
                ["Общий итог", 30],
            ],
            "Data": [baslik] + [[i + 1, *k] for i, k in enumerate(kisiler)],
            "Rusya Ekspat": [baslik] + [[i + 1, *k] for i, k in enumerate(kisiler[:25])],
            "Rusya Yerel": [baslik] + [[i + 1, 700000 + i, f"Yerel{i:02d} Soyad", "GPP Project"] for i in range(20)],
        })
        cls.notlar: list[str] = []
        cls.satirlar = genel_oku(cls.yol, notlar=cls.notlar)

    @classmethod
    def tearDownClass(cls):
        cls.dizin.cleanup()

    def test_pivot_sayfalari_kisi_uretmez(self):
        sayfalar = {s.ek.get("sayfa") for s in self.satirlar}
        self.assertNotIn("Ozet", sayfalar)
        self.assertNotIn("Лист1", sayfalar)
        self.assertFalse(any(s.kisi_ham in ("Grand Total", "Row Labels", "GPP Project")
                             for s in self.satirlar))
        self.assertTrue(any("'Ozet' sayfasi pivot" in n for n in self.notlar), self.notlar)

    def test_tekrar_eden_sayfa_isaretlenir(self):
        ekspat = [s for s in self.satirlar if s.ek.get("sayfa") == "Rusya Ekspat"]
        self.assertEqual(len(ekspat), 25)
        self.assertTrue(all(s.ek.get("sayfa_tekrari") == "Data" for s in ekspat))
        data = [s for s in self.satirlar if s.ek.get("sayfa") == "Data"]
        self.assertFalse(any(s.ek.get("sayfa_tekrari") for s in data))
        yerel = [s for s in self.satirlar if s.ek.get("sayfa") == "Rusya Yerel"]
        self.assertFalse(any(s.ek.get("sayfa_tekrari") for s in yerel))
        self.assertTrue(any("tekrar ediyor" in n for n in self.notlar), self.notlar)

    def test_benzersiz_kisi_sayisi_sismez(self):
        self.assertEqual(len(self.satirlar), 30 + 25 + 20)
        benzersiz = {kisi_anahtari(s) for s in self.satirlar if kisi_anahtari(s)}
        self.assertEqual(len(benzersiz), 50)

    def test_kucuk_sayfalar_tekrar_sayilmaz(self):
        satirlar = [_gider("genel", f"Kisi {i}", sicil=str(i), sayfa="A") for i in range(5)]
        satirlar += [_gider("genel", f"Kisi {i}", sicil=str(i), sayfa="B") for i in range(5)]
        self.assertEqual(sayfa_tekrarlarini_isaretle(satirlar), {})


# --------------------------------------------------------------------------
# Bulgu 4a: kutuk satirlari defteri beslemiyor, mesaj gercegi soylesin
# --------------------------------------------------------------------------

class KutukBeslemeTest(unittest.TestCase):
    def test_referans_listesi_beslemez_ve_nedeni_yazilir(self):
        with tempfile.TemporaryDirectory() as d:
            defterler = Defterler(d)
            satirlar = [_gider("referans_liste", f"Kisi{i} Soyad", sicil=str(600000 + i)) for i in range(5)]
            ozet = defterler.yardimci_kaynaktan_besle(satirlar)
        self.assertEqual(ozet["ek_kisi"], 0)
        self.assertEqual(ozet["atlanan"], 5)
        self.assertEqual(ozet["kaynak_disi"], {"referans_liste": 5})
        self.assertEqual(ozet["tip_ozeti"]["referans_liste"]["satir"], 5)
        metin = besleme_aciklamasi(ozet, "referans_liste")
        self.assertTrue(metin.startswith("defteri beslemedi:"), metin)
        self.assertIn("sicil tasiyan personel kutugu", metin)

    def test_besleyen_kaynak_sayilari(self):
        with tempfile.TemporaryDirectory() as d:
            defterler = Defterler(d)
            satirlar = [_gider("energo_saglik", f"Kisi{i} Soyad", merkez="GPP Project") for i in range(3)]
            satirlar.append(_gider("energo_saglik", None))  # kimliksiz
            ozet = defterler.yardimci_kaynaktan_besle(satirlar)
        self.assertEqual(ozet["ek_kisi"], 3)
        self.assertEqual(ozet["kimliksiz"], 1)
        self.assertEqual(ozet["islenen"], 3)
        tip = ozet["tip_ozeti"]["energo_saglik"]
        self.assertEqual((tip["satir"], tip["ek_kisi"], tip["kimliksiz"]), (4, 3, 1))
        metin = besleme_aciklamasi(ozet, "energo_saglik")
        self.assertTrue(metin.startswith("defteri besledi:"), metin)
        self.assertIn("3 kisi ek kisi defterine yazildi", metin)

    def test_besleme_kapaliysa_soylenir(self):
        self.assertIn("kapali", besleme_aciklamasi(None, "energo_saglik", besleme_acik=False))
        self.assertTrue(besleme_aciklamasi(None, "energo_saglik").startswith("defteri beslemedi"))

    def test_boru_kutuk_uyarisi_gercek_sayilari_yazar(self):
        from masraf.boru import Boru
        from masraf.envanter import KUTUK, DosyaKaydi

        boru = Boru.__new__(Boru)
        boru.ayarlar = types.SimpleNamespace(defterleri_besle=True)
        boru.envanter = [DosyaKaydi(ad="liste.xlsx", tur="referans_liste", durum=KUTUK,
                                    sebep="kisi listesi; defter beslemesinde kullanildi, dagilima girmedi")]
        referanslar = [_gider("referans_liste", f"Kisi{i} Soyad", sicil=str(600000 + i), sayfa="Data") for i in range(30)]
        referanslar += [_gider("referans_liste", f"Kisi{i} Soyad", sicil=str(600000 + i), sayfa="Ekspat") for i in range(25)]
        for s in referanslar:
            # kesif kutuk nedenini ve cozulen kolonlari satira yazar; tutar kolonu VAR ama bos.
            s.ek["kutuk_sebebi"] = "55 satirin hicbirinde tutar yok"
            s.ek["cozulen_kolonlar"] = {"tutar": 3}
        sayfa_tekrarlarini_isaretle(referanslar)
        with tempfile.TemporaryDirectory() as d:
            ozet = Defterler(d).yardimci_kaynaktan_besle(referanslar)
        metin = boru._kutuk_uyarisi(referanslar, besleme=ozet)
        self.assertIn("55 satir (30 benzersiz kisi)", metin)
        self.assertIn("neden: 55 satirin hicbirinde tutar yok", metin)
        self.assertIn("Ekspat 25 ('Data' sayfasindaki kisileri tekrar ediyor)", metin)
        self.assertIn("Kutuk defteri beslemedi:", metin)
        self.assertNotIn("defter beslemesinde kullanildi", metin)
        # Tutar kolonu cozulmusse (bos da olsa) sozluk tavsiyesi gosterilmez.
        self.assertNotIn("kolon_esanlamlilari.csv", metin)
        # Tutar kolonu hic cozulmediyse fatura olabilir: tavsiye gorunur.
        for s in referanslar:
            s.ek["kutuk_sebebi"] = "sicil ve masraf merkezi kolonlari var, tutar kolonu yok"
            s.ek["cozulen_kolonlar"] = {"tutar": None}
        metin = boru._kutuk_uyarisi(referanslar, besleme=ozet)
        self.assertIn("neden: sicil ve masraf merkezi kolonlari var, tutar kolonu yok", metin)
        self.assertIn("kolon_esanlamlilari.csv", metin)
        boru._kutuk_envanterini_guncelle(ozet)
        self.assertIn("defteri beslemedi", boru.envanter[0].sebep)
        self.assertNotIn("defter beslemesinde kullanildi", boru.envanter[0].sebep)
        # Besleme kapaliyken de gercek soylenir.
        boru.ayarlar = types.SimpleNamespace(defterleri_besle=False)
        self.assertIn("kapali", boru._kutuk_uyarisi(referanslar, besleme=None))


# --------------------------------------------------------------------------
# Bulgu 5: Yuzyil paylasim deseni
# --------------------------------------------------------------------------

class PaylasimDeseniTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            from masraf.okuyucular.antik import _paylasim_ayristir
        except ImportError as e:  # antik.py baska bir ajan tarafindan degistiriliyor olabilir
            raise unittest.SkipTest(f"antik._paylasim_ayristir yok: {e}")
        cls.ayristir = staticmethod(_paylasim_ayristir)

    def _ozet(self, *metinler):
        return [(p["masraf_merkezi"], p["pay"], p["bolen"]) for p in self.ayristir(*metinler)]

    def test_gercek_paylasim_korunur(self):
        self.assertEqual(self._ozet("RHI 1/3- RENSTROYDETAL 2/3"), [("RHI", 1, 3), ("RENSTROYDETAL", 2, 3)])
        self.assertEqual(self._ozet("UST LUGA 1/2 - RHI 1/2"), [("UST LUGA", 1, 2), ("RHI", 1, 2)])

    def test_tarih_kesir_sayilmaz(self):
        self.assertEqual(self._ozet("RHI", "GIRIS 10/07/2026"), [])
        self.assertEqual(self._ozet("RHI 01/07/2026"), [])
        self.assertEqual(self._ozet("RHI 10/07 14:30"), [])

    def test_makul_sinirlar(self):
        self.assertEqual(self._ozet("RHI 3/2"), [])    # pay > bolen
        self.assertEqual(self._ozet("RHI 2/25"), [])   # bolen > 12
        self.assertEqual(self._ozet("RHI 0/3"), [])    # pay 0

    def test_taninmayan_etiket_paylasim_uretmez(self):
        self.assertEqual(self._ozet("ODA 2/3 KISI"), [])
        self.assertEqual(self._ozet("GECE 1/2"), [])


# --------------------------------------------------------------------------
# Bulgu 7: arayuz yukleme klasoru temizligi
# --------------------------------------------------------------------------

class YuklemeDiziniTemizligiTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            import app  # noqa: F401
        except Exception as e:  # noqa: BLE001  streamlit yoksa atla
            raise unittest.SkipTest(f"app.py yuklenemedi: {e}")
        cls.app = app

    class _Yuklenen:
        def __init__(self, ad: str, veri: bytes):
            self.name = ad
            self._veri = veri

        def getbuffer(self):
            return memoryview(self._veri)

    def _sahte_st(self, dizin: str):
        return types.SimpleNamespace(session_state={"yukleme_dizini": dizin},
                                     error=lambda *a, **k: None)

    def test_onceki_parti_silinir_mevcut_parti_kalir(self):
        app = self.app
        eski_st = app.st
        with tempfile.TemporaryDirectory() as d:
            app.st = self._sahte_st(d)
            try:
                parti1 = app._yuklenenleri_kaydet([self._Yuklenen("a.xlsx", b"111")])
                parti1_tekrar = app._yuklenenleri_kaydet([self._Yuklenen("a.xlsx", b"111")])
                parti2 = app._yuklenenleri_kaydet([self._Yuklenen("b.xlsx", b"22"), self._Yuklenen("c.csv", b"3")])
                self.assertEqual(parti1, parti1_tekrar)  # yeniden calistirma yeni klasor acmaz
                self.assertNotEqual(parti1[0].parent, parti2[0].parent)
                self.assertTrue(all(p.is_file() for p in parti1 + parti2))
                silinen = app._eski_yuklemeleri_temizle(parti2)
                self.assertEqual(silinen, 1)
                self.assertFalse(parti1[0].exists())
                self.assertTrue(all(p.is_file() for p in parti2))
                app._yukleme_dizinini_sil(d)
                self.assertFalse(Path(d).exists())
            finally:
                app.st = eski_st
            Path(d).mkdir(exist_ok=True)  # TemporaryDirectory temizligi icin

    def test_oturum_baslangicinda_atexit_kaydi(self):
        app = self.app
        eski_st, eski_atexit = app.st, app.atexit
        kayitlar: list = []
        app.st = types.SimpleNamespace(session_state={})
        app.atexit = types.SimpleNamespace(register=lambda f, *a: kayitlar.append((f, a)))
        try:
            app._oturum_hazirla()
            dizin = app.st.session_state["yukleme_dizini"]
            self.assertTrue(Path(dizin).is_dir())
            self.assertEqual(kayitlar, [(app._yukleme_dizinini_sil, (dizin,))])
            app._yukleme_dizinini_sil(dizin)
            self.assertFalse(Path(dizin).exists())
        finally:
            app.st, app.atexit = eski_st, eski_atexit


# --------------------------------------------------------------------------
# Ek bulgu 1: isim kolonu yokken aciklamaya kimlik/dogum/telefon dusmesin
# --------------------------------------------------------------------------

class AciklamaGizlilikTest(unittest.TestCase):
    TCKN = "12345678901"  # uydurma, gecersiz kontrol haneli

    def test_kimlik_dogum_telefon_aciklamaya_girmez(self):
        with tempfile.TemporaryDirectory() as d:
            yol = _kitap_yaz(Path(d) / "uye.xlsx", {"Sheet": [
                ["Sira", "Uye", "TCKN", "Dogum Tarihi", "Telefon", "Tutar (USD)"],
                [1, "Aaaa Bbbb", self.TCKN, "01.01.1990", "+7 921 123 45 67", 120.5],
                [2, "Cccc Dddd", int(self.TCKN), date(1985, 5, 5), "0 532 123 45 67", "2.500"],
            ]})
            satirlar = genel_oku(yol)
        self.assertEqual(len(satirlar), 2)
        for s in satirlar:
            self.assertIsNotNone(s.tckn_ham)          # TCKN alani dolu, ama aciklamada yok
            self.assertNotIn(self.TCKN, s.aciklama)
            self.assertIsNone(re.search(r"\d{11}", s.aciklama), s.aciklama)
            self.assertNotIn("1990", s.aciklama)
            self.assertNotIn("1985", s.aciklama)
            self.assertNotIn("921", s.aciklama)
            self.assertNotIn("532", s.aciklama)
        self.assertIn("Aaaa Bbbb", satirlar[0].aciklama)
        self.assertIn("120.5", satirlar[0].aciklama)

    def test_geriye_bir_sey_kalmazsa_kimlik_maskelenir(self):
        with tempfile.TemporaryDirectory() as d:
            yol = _kitap_yaz(Path(d) / "kimlik.xlsx", {"Sheet": [
                ["TCKN", "Dogum Tarihi", "Telefon"],
                [self.TCKN, "01.01.1990", "+7 921 123 45 67"],
            ]})
            satirlar = genel_oku(yol)
        self.assertEqual(len(satirlar), 1)
        self.assertEqual(satirlar[0].aciklama, "TCKN 123******")


# --------------------------------------------------------------------------
# Ek bulgu 2: excel_yaz hatasi yutulmasin, dosya '_EKSIK' isaretlensin
# --------------------------------------------------------------------------

class ExcelUretHataTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            import app  # noqa: F401
        except Exception as e:  # noqa: BLE001
            raise unittest.SkipTest(f"app.py yuklenemedi: {e}")
        cls.app = app

    def test_cekirdek_yazici_hata_verince_uyari_ve_eksik_dosya(self):
        app = self.app

        class _BozukCikti:
            @staticmethod
            def excel_yaz(*a, **k):
                raise RuntimeError("sentetik yazici hatasi")

        eski = app.CIKTI_MODULU
        app.CIKTI_MODULU = _BozukCikti
        try:
            with tempfile.TemporaryDirectory() as d:
                hatalar: list[str] = []
                hedef = app.excel_uret([], Path(d) / "masraf_merkezi_x.xlsx", {}, hatalar=hatalar)
                self.assertEqual(hedef.name, "masraf_merkezi_x_EKSIK.xlsx")
                self.assertTrue(hedef.is_file())
                self.assertEqual(len(hatalar), 1)
                self.assertIn("sentetik yazici hatasi", hatalar[0])
                self.assertIn("EKSİK", hatalar[0])
        finally:
            app.CIKTI_MODULU = eski

    def test_cekirdek_yazici_calisirsa_ad_degismez(self):
        app = self.app

        class _Cikti:
            @staticmethod
            def excel_yaz(sonuclar, yol, ozet, mahsup):
                Path(yol).write_bytes(b"ok")

        eski = app.CIKTI_MODULU
        app.CIKTI_MODULU = _Cikti
        try:
            with tempfile.TemporaryDirectory() as d:
                hatalar: list[str] = []
                hedef = app.excel_uret([], Path(d) / "tam.xlsx", {}, hatalar=hatalar)
                self.assertEqual(hedef.name, "tam.xlsx")
                self.assertEqual(hatalar, [])
        finally:
            app.CIKTI_MODULU = eski


# --------------------------------------------------------------------------
# Ek bulgu 3: bozuk defter dosyasi sessizce bos okunmasin
# --------------------------------------------------------------------------

class DefterBozukDosyaTest(unittest.TestCase):
    def _defter(self, d: str, dosya: str, icerik: bytes) -> Defterler:
        (Path(d) / dosya).write_bytes(icerik)
        return Defterler(d)

    ALIAS_METNI = "isim_norm;sicil;ad_soyad;kaynak;eklenme_tarihi\nAAAA BBBB;600001;Aaaa Bbbb;inceleme;2026-01-01\n"

    def test_bomlu_utf16_dosya_kurtarilir(self):
        # Excel 'Unicode metin' olarak kaydedince BOM'lu UTF-16 uretir; eski
        # surum bunu sessizce bos okuyordu. Artik dogru cozulur, kayit kaybolmaz.
        with tempfile.TemporaryDirectory() as d:
            defterler = self._defter(d, "aliases.csv", self.ALIAS_METNI.encode("utf-16"))
        self.assertEqual(defterler.aliases, {"AAAA BBBB": "600001"})
        self.assertEqual(defterler.uyarilar, [])

    def test_bomsuz_utf16_dosya_uyari_verir(self):
        with tempfile.TemporaryDirectory() as d:
            defterler = self._defter(d, "aliases.csv", self.ALIAS_METNI.encode("utf-16-le"))
        self.assertEqual(defterler.aliases, {})
        self.assertTrue(any("aliases.csv okunamadi/bozuk" in u for u in defterler.uyarilar), defterler.uyarilar)
        self.assertEqual(defterler.istatistik()["yuklenen"]["aliases.csv"], 0)

    def test_rastgele_bayt_uyari_verir(self):
        with tempfile.TemporaryDirectory() as d:
            defterler = self._defter(d, "harici_kisiler.csv", bytes(range(256)) * 4)
        self.assertEqual(defterler.harici, {})
        self.assertTrue(any("harici_kisiler.csv okunamadi/bozuk" in u for u in defterler.uyarilar))

    def test_yanlis_baslik_uyari_verir(self):
        with tempfile.TemporaryDirectory() as d:
            defterler = self._defter(d, "aliases.csv", "Ad;Soyad\nAaaa;Bbbb\n".encode("utf-8-sig"))
        self.assertTrue(any("beklenen basliklar" in u for u in defterler.uyarilar), defterler.uyarilar)

    def test_virgul_ayiricili_dosya_sorunsuz_okunur(self):
        with tempfile.TemporaryDirectory() as d:
            metin = "isim_norm,sicil,ad_soyad,kaynak,eklenme_tarihi\nAAAA BBBB,600001,Aaaa Bbbb,inceleme,2026-01-01\n"
            defterler = self._defter(d, "aliases.csv", metin.encode("utf-8-sig"))
        self.assertEqual(defterler.aliases, {"AAAA BBBB": "600001"})
        self.assertEqual(defterler.uyarilar, [])
        ist = defterler.istatistik()
        self.assertEqual(ist["okunan"]["aliases.csv"], 1)
        self.assertEqual(ist["yuklenen"]["aliases.csv"], 1)

    def test_satirlarin_cogu_ayristirilamazsa_uyari(self):
        with tempfile.TemporaryDirectory() as d:
            satirlar = ["tckn;sicil;ad_soyad;kaynak;eklenme_tarihi"]
            satirlar += [f";;Kisi{i};inceleme;2026-01-01" for i in range(8)]  # tckn ve sicil bos
            satirlar += ["11111111110;600001;Aaaa;inceleme;2026-01-01"]
            defterler = self._defter(d, "tckn_sicil.csv", ("\n".join(satirlar) + "\n").encode("utf-8-sig"))
        self.assertEqual(len(defterler.tckn_sicil), 1)
        self.assertTrue(any("tckn_sicil.csv bozuk olabilir" in u for u in defterler.uyarilar), defterler.uyarilar)

    def test_saglam_dosyalar_uyari_uretmez(self):
        with tempfile.TemporaryDirectory() as d:
            defterler = Defterler(d)
            defterler.alias_ekle("AAAA BBBB", "600001")
            defterler.harici_ekle("CCCC DDDD", "Cccc Dddd", "Dis Firma", "HQ")
            defterler.kaydet()
            yeniden = Defterler(d)
        self.assertEqual(yeniden.uyarilar, [])
        self.assertEqual(yeniden.istatistik()["yuklenen"], {
            "aliases.csv": 1, "harici_kisiler.csv": 1, "ek_kisiler.csv": 0, "tckn_sicil.csv": 0,
        })


if __name__ == "__main__":
    unittest.main()
