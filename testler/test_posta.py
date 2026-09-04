"""Outlook .msg ek cikarma testleri."""

from __future__ import annotations

import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

try:
    from masraf.okuyucular.posta import (
        CikarilanEk,
        _guvenli_ad,
        _zip_ac,
        msg_aciklarini_cikar,
        msg_mi,
    )

    MODUL_VAR = True
except ImportError:
    MODUL_VAR = False

ORNEK_MSG = Path("ornek_veri/posta")


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular.posta bulunamadi")
class GuvenliAdTest(unittest.TestCase):
    def test_zip_kacislari_cozulur(self):
        # Zip araclari Turkce karakterleri #Uxxxx olarak kacisliyor
        self.assertEqual(_guvenli_ad("Kat#U0131l#U0131mc#U0131 Listesi.xlsx"),
                         "Katılımcı Listesi.xlsx")
        self.assertEqual(_guvenli_ad("As #U00d6l#U00e7me.xlsx"), "As Ölçme.xlsx")

    def test_yol_ayiricilari_temizlenir(self):
        self.assertNotIn("/", _guvenli_ad("a/b/c.xlsx"))
        self.assertNotIn("\\", _guvenli_ad("a\\b.xlsx"))

    def test_outlook_konu_adlari(self):
        # Outlook ekleri '>>: Konu' gibi adlar tasiyabiliyor
        ad = _guvenli_ad(">>: Energo Taahhüt Mayıs - Haziran Yansıtma")
        self.assertNotIn(">", ad)
        self.assertNotIn(":", ad)
        self.assertTrue(ad)

    def test_bos_ad_varsayilana_duser(self):
        self.assertEqual(_guvenli_ad(""), "adsiz")
        self.assertEqual(_guvenli_ad("   "), "adsiz")
        self.assertEqual(_guvenli_ad("...", "yedek"), "yedek")


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular.posta bulunamadi")
class ZipGuvenlikTest(unittest.TestCase):
    def setUp(self):
        self.gecici = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.gecici, ignore_errors=True)

    def test_normal_arsiv_acilir(self):
        z = self.gecici / "a.zip"
        with zipfile.ZipFile(z, "w") as zf:
            zf.writestr("klasor/veri.xlsx", b"icerik")
        cikanlar = _zip_ac(z, self.gecici / "acilmis")
        self.assertEqual(len(cikanlar), 1)
        self.assertEqual(cikanlar[0].name, "veri.xlsx")
        self.assertEqual(cikanlar[0].read_bytes(), b"icerik")

    def test_yol_gecisi_engellenir(self):
        """../ iceren girdi arsiv dizininin disina yazamamali."""
        z = self.gecici / "kotu.zip"
        with zipfile.ZipFile(z, "w") as zf:
            zf.writestr("../../kacan.xlsx", b"zararli")
        hedef = self.gecici / "acilmis"
        _zip_ac(z, hedef)
        self.assertFalse((self.gecici.parent / "kacan.xlsx").exists())
        for c in hedef.rglob("*"):
            self.assertTrue(str(c.resolve()).startswith(str(hedef.resolve())))

    def test_bozuk_arsiv_coktermez(self):
        z = self.gecici / "bozuk.zip"
        z.write_bytes(b"bu bir zip degil")
        self.assertEqual(_zip_ac(z, self.gecici / "acilmis"), [])


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular.posta bulunamadi")
class MsgTanimaTest(unittest.TestCase):
    def test_uzanti_tanima(self):
        self.assertTrue(msg_mi("gelen.msg"))
        self.assertTrue(msg_mi("GELEN.MSG"))
        self.assertFalse(msg_mi("tablo.xlsx"))

    def test_olmayan_dosya_hata_verir(self):
        with self.assertRaises(FileNotFoundError):
            msg_aciklarini_cikar("olmayan_dosya.msg", tempfile.mkdtemp())


@unittest.skipUnless(MODUL_VAR, "masraf.okuyucular.posta bulunamadi")
class GercekMsgTest(unittest.TestCase):
    """Gercek bir Outlook mesaji varsa ic ice cikarmayi dogrular."""

    def setUp(self):
        if not ORNEK_MSG.is_dir():
            self.skipTest("ornek_veri/posta dizini yok")
        self.msgler = sorted(ORNEK_MSG.glob("*.msg"))
        if not self.msgler:
            self.skipTest("ornek .msg dosyasi yok")
        self.gecici = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(getattr(self, "gecici", Path(tempfile.mkdtemp())), ignore_errors=True)

    def test_ic_ice_ekler_cikarilir(self):
        ekler = msg_aciklarini_cikar(self.msgler[0], self.gecici)
        self.assertGreater(len(ekler), 5, "ic ice eklerin cogu bulunamadi")
        adlar = {e.ad for e in ekler}
        # Ana mailin dogrudan eki
        self.assertTrue(any(a.endswith(".xls") for a in adlar), "ham .xls bulunamadi")
        # Zip icinden cikan dosya (derinlik >= 2)
        self.assertTrue(any(e.derinlik >= 2 for e in ekler), "arsiv icindekiler cikarilmadi")
        # Her ek kaynak zincirini tasimali
        for e in ekler:
            self.assertTrue(e.kaynak_aciklamasi)
            self.assertTrue(e.yol.is_file())
            self.assertGreater(e.yol.stat().st_size, 0)

    def test_gorsel_ekler_atlanir(self):
        ekler = msg_aciklarini_cikar(self.msgler[0], self.gecici)
        for e in ekler:
            self.assertNotIn(e.yol.suffix.lower(), {".png", ".jpg", ".jpeg", ".gif"})

    def test_ayni_dosya_tekrarlanmaz(self):
        ekler = msg_aciklarini_cikar(self.msgler[0], self.gecici)
        anahtarlar = [(e.ad, e.yol.stat().st_size) for e in ekler]
        self.assertEqual(len(anahtarlar), len(set(anahtarlar)), "yinelenen ek dondu")


if __name__ == "__main__":
    unittest.main(verbosity=2)
