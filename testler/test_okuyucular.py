"""Kaynak dosya okuyucularinin testleri (masraf.okuyucular).

Her fatura dosyasi ailesi icin: dogru satir sayisi cikiyor mu, kisi adi
cikarma orani kabul edilebilir mi, kimlik anahtarlari (sicil / TCKN) dolu mu.

Ornek veri (``ornek_veri/``) repoya girmez; yoksa testler atlanir.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

KOK = Path(__file__).resolve().parents[1]
if str(KOK) not in sys.path:
    sys.path.insert(0, str(KOK))

try:
    from masraf.modeller import GIDER_TIPLERI, KAYNAK_TIPLERI, GiderSatiri
    from masraf.okuyucular.antik import antik_cari_oku, yuzyil_dagitilmis_oku
    from masraf.okuyucular.energo import (
        arabulucu_oku,
        assessment_oku,
        koc_katilimci_oku,
        saglik_oku,
    )
    from masraf.okuyucular.kesif import dosya_tipini_bul

    MODUL_VAR = True
except ImportError:
    MODUL_VAR = False

ORNEK = KOK / "ornek_veri"
ANTIK = ORNEK / "antik_travel" / "ANTIK_CARI_TEMMUZ_2026.xls"
YUZYIL = ORNEK / "antik_travel" / "YUZYIL_TEMMUZ_2026_ELLE_DAGITILMIS.xlsx"
ASSESSMENT = ORNEK / "energo" / "ASSESSMENT_YANSITMA_2026_05_06.xlsx"
ARABULUCU = ORNEK / "energo" / "ARABULUCULUK_2026_06_07.xlsx"
SAGLIK = ORNEK / "energo" / "SAGLIK_KONTROL_LISTE.xlsx"
KOC = ORNEK / "energo" / "KOC_UNI_KATILIMCI_LISTESI.xlsx"


def veri_gerek(test: unittest.TestCase, *yollar: Path) -> None:
    """Ornek veri yoksa testi atlar."""
    for yol in yollar:
        if not yol.exists():
            test.skipTest(f"Ornek veri bulunamadi: {yol}")


class OkuyucuTemeli(unittest.TestCase):
    """Her okuyucunun ciktisi icin ortak sozlesme kontrolleri."""

    def sozlesmeyi_dogrula(self, satirlar, beklenen_tip: str) -> None:
        self.assertTrue(satirlar, "Okuyucu hic satir uretmedi")
        for satir in satirlar:
            self.assertIsInstance(satir, GiderSatiri)
            self.assertEqual(satir.kaynak_tip, beklenen_tip)
            self.assertIn(satir.kaynak_tip, KAYNAK_TIPLERI)
            self.assertGreaterEqual(satir.satir_no, 1)
            self.assertIsInstance(satir.aciklama, str)
            if satir.gider_tipi is not None:
                self.assertIn(satir.gider_tipi, GIDER_TIPLERI)


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular bulunamadi")
class AntikCariTest(OkuyucuTemeli):
    """Seyahat acentesi ham cari hareket dokumu (.xls, xlrd)."""

    @classmethod
    def setUpClass(cls):
        if not ANTIK.exists():
            cls.satirlar = None
            return
        cls.satirlar = antik_cari_oku(ANTIK)

    def setUp(self):
        veri_gerek(self, ANTIK)

    def test_satir_sayisi(self):
        # Dosyada r10..r143 arasi 134 veri satiri var; r144 TOPLAM satiridir
        # ve veriye dahil EDILMEMELIDIR.
        self.assertEqual(len(self.satirlar), 134)

    def test_toplam_satiri_alinmamis(self):
        for satir in self.satirlar:
            self.assertNotIn("TOPLAM", (satir.aciklama or "").upper()[:10])

    def test_kisi_cikarma_orani(self):
        kisili = [s for s in self.satirlar if s.kisi_ham]
        self.assertGreaterEqual(
            len(kisili), 120,
            f"Antik dosyasindan en az 120 kisili satir beklenir, {len(kisili)} cikti",
        )
        oran = len(kisili) / len(self.satirlar)
        self.assertGreaterEqual(oran, 0.90)

    def test_sozlesme(self):
        self.sozlesmeyi_dogrula(self.satirlar, "antik_cari")

    def test_bilet_satirinda_isim_temizlenmis(self):
        # 'TK4093099626 OZAKAY/MUSTAFAKEMAL MR  IST-CDG BILET BEDELI'
        satir = next(s for s in self.satirlar if s.satir_no == 10)
        self.assertEqual(satir.gider_tipi, "Bilet")
        self.assertEqual(satir.kisi_ham, "OZAKAY MUSTAFAKEMAL")

    def test_otel_satirinda_otel_adi_atilmis(self):
        # 'CIHAN BALABAN GRAND PLAZA HOTEL HANOI [...] KONAKLAMA YURTDISI'
        satir = next(s for s in self.satirlar if s.satir_no == 12)
        self.assertEqual(satir.gider_tipi, "Otel")
        self.assertEqual(satir.kisi_ham, "CIHAN BALABAN")

    def test_gider_tipleri_dagilimi(self):
        tipler = {s.gider_tipi for s in self.satirlar}
        self.assertIn("Bilet", tipler)
        self.assertIn("Otel", tipler)

    def test_tarih_cozulmus(self):
        # Excel seri numarasi (46204.0) gercek tarihe cevrilmis olmali.
        tarihli = [s for s in self.satirlar if s.belge_tarihi is not None]
        self.assertGreaterEqual(len(tarihli), len(self.satirlar) - 5)
        for satir in tarihli:
            self.assertEqual(satir.belge_tarihi.year, 2026)

    def test_masraf_merkezi_kolonu_yok(self):
        # Ham cari dokumunde santiye bilgisi YOKTUR; tum degeri buradan
        # cikarmak eslestiricinin isidir.
        self.assertFalse(any(s.masraf_merkezi_kaynak for s in self.satirlar))


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular bulunamadi")
class YuzyilDagitilmisTest(OkuyucuTemeli):
    """Elle santiyeye dagitilmis referans dosya (dogruluk olcumu icin)."""

    @classmethod
    def setUpClass(cls):
        if not YUZYIL.exists():
            cls.satirlar = None
            return
        cls.satirlar = yuzyil_dagitilmis_oku(YUZYIL)

    def setUp(self):
        veri_gerek(self, YUZYIL)

    def test_satir_sayisi(self):
        # 139 satirin son 4'u TOPLAM/IADE/ODENECEK ozet satiridir.
        self.assertGreaterEqual(len(self.satirlar), 130)
        self.assertLessEqual(len(self.satirlar), 135)

    def test_kisi_cikarma_orani(self):
        kisili = [s for s in self.satirlar if s.kisi_ham]
        self.assertGreaterEqual(len(kisili), 120)

    def test_santiye_kolonu_okunmus(self):
        # Bu dosyanin degeri SANTIYESI kolonundadir; olcum referansi budur.
        merkezli = [s for s in self.satirlar if s.masraf_merkezi_kaynak]
        self.assertGreaterEqual(len(merkezli), 120)
        etiketler = {s.masraf_merkezi_kaynak for s in merkezli}
        self.assertTrue(
            any("RHI" in (e or "").upper() for e in etiketler),
            f"RHI etiketi bekleniyordu, gorulen: {sorted(etiketler)[:10]}",
        )

    def test_sozlesme(self):
        self.sozlesmeyi_dogrula(self.satirlar, "yuzyil_dagitilmis")


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular bulunamadi")
class EnergoAssessmentTest(OkuyucuTemeli):
    """Assessment yansitma dosyasi (Kisi Listesi sayfasi)."""

    @classmethod
    def setUpClass(cls):
        if not ASSESSMENT.exists():
            cls.satirlar = None
            return
        cls.satirlar = assessment_oku(ASSESSMENT)

    def setUp(self):
        veri_gerek(self, ASSESSMENT)

    def test_satir_ve_kisi(self):
        self.assertGreaterEqual(len(self.satirlar), 5)
        # Katilimci kolonu her satirda dolu; kisi cikarma orani %100 olmali.
        self.assertEqual(len([s for s in self.satirlar if s.kisi_ham]),
                         len(self.satirlar))

    def test_sozlesme(self):
        self.sozlesmeyi_dogrula(self.satirlar, "energo_assessment")


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular bulunamadi")
class EnergoArabulucuTest(OkuyucuTemeli):
    """Arabuluculuk dosyasi: TCKN ve PROJE kolonlari var."""

    @classmethod
    def setUpClass(cls):
        if not ARABULUCU.exists():
            cls.satirlar = None
            return
        cls.satirlar = arabulucu_oku(ARABULUCU)

    def setUp(self):
        veri_gerek(self, ARABULUCU)

    def test_satir_sayisi(self):
        self.assertEqual(len(self.satirlar), 25)

    def test_her_satirda_kisi_var(self):
        self.assertEqual(len([s for s in self.satirlar if s.kisi_ham]), 25)

    def test_her_satirda_tckn_var(self):
        tcknli = [s for s in self.satirlar if s.tckn_ham]
        self.assertEqual(len(tcknli), 25)
        for satir in tcknli:
            self.assertEqual(len(satir.tckn_ham), 11, satir.tckn_ham)
            self.assertTrue(satir.tckn_ham.isdigit())

    def test_proje_kolonu_masraf_merkezi_kaynagina_gecmis(self):
        merkezli = [s for s in self.satirlar if s.masraf_merkezi_kaynak]
        self.assertGreaterEqual(len(merkezli), 20)

    def test_sozlesme(self):
        self.sozlesmeyi_dogrula(self.satirlar, "energo_arabulucu")


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular bulunamadi")
class EnergoSaglikTest(OkuyucuTemeli):
    """Saglik kontrol listesi: ek kisi defterini besleyen ana kaynak."""

    @classmethod
    def setUpClass(cls):
        if not SAGLIK.exists():
            cls.satirlar = None
            return
        cls.satirlar = saglik_oku(SAGLIK)

    def setUp(self):
        veri_gerek(self, SAGLIK)

    def test_satir_sayisi(self):
        self.assertEqual(len(self.satirlar), 50)

    def test_tckn_ve_santiye_dolu(self):
        self.assertEqual(len([s for s in self.satirlar if s.tckn_ham]), 50)
        self.assertEqual(len([s for s in self.satirlar if s.kisi_ham]), 50)
        merkezli = [s for s in self.satirlar if s.masraf_merkezi_kaynak]
        self.assertGreaterEqual(len(merkezli), 40)

    def test_sozlesme(self):
        self.sozlesmeyi_dogrula(self.satirlar, "energo_saglik")


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular bulunamadi")
class KocKatilimciTest(OkuyucuTemeli):
    """Koc Universitesi katilimci listesi: ID kolonu SICIL numarasidir."""

    @classmethod
    def setUpClass(cls):
        if not KOC.exists():
            cls.satirlar = None
            return
        cls.satirlar = koc_katilimci_oku(KOC)

    def setUp(self):
        veri_gerek(self, KOC)

    def test_satir_sayisi(self):
        self.assertEqual(len(self.satirlar), 50)

    def test_her_satirda_sicil_var(self):
        # En kolay ve en guvenilir eslesme yolu; hicbir satir sicilsiz olmamali.
        sicilsiz = [s.satir_no for s in self.satirlar if not s.sicil_ham]
        self.assertEqual(sicilsiz, [], f"Sicilsiz satirlar: {sicilsiz}")

    def test_sicil_metin_olarak_ve_temiz(self):
        for satir in self.satirlar:
            self.assertIsInstance(satir.sicil_ham, str)
            self.assertNotIn(".0", satir.sicil_ham)
            self.assertEqual(satir.sicil_ham, satir.sicil_ham.strip())

    def test_her_satirda_kisi_var(self):
        self.assertEqual(len([s for s in self.satirlar if s.kisi_ham]), 50)

    def test_sozlesme(self):
        self.sozlesmeyi_dogrula(self.satirlar, "koc_katilimci")


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular bulunamadi")
class DosyaTipiKesfiTest(unittest.TestCase):
    """Kullanici dosyayi surukleyip biraktiginda tip otomatik bulunmalidir."""

    BEKLENEN = (
        (ANTIK, "antik_cari"),
        (YUZYIL, "yuzyil_dagitilmis"),
        (ASSESSMENT, "energo_assessment"),
        (ARABULUCU, "energo_arabulucu"),
        (SAGLIK, "energo_saglik"),
        (KOC, "koc_katilimci"),
    )

    def test_tum_ornek_dosyalar_taninir(self):
        for yol, beklenen in self.BEKLENEN:
            with self.subTest(dosya=yol.name):
                veri_gerek(self, yol)
                self.assertEqual(dosya_tipini_bul(yol), beklenen)


if __name__ == "__main__":
    unittest.main()
