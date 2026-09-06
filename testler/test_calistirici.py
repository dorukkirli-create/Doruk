"""Klasor tabanli calistirici (paketle/calistir.py) testleri.

Iki davranis muhasebe icin kritik:

1. Islenen faturalar arsive TASINIR. 1_FATURALAR bosaltilmazsa gelecek ay ayni
   dosyalar tekrar islenir ve tutarlar iki kez sayilir. Kullanicinin
   hatirlamasina birakilmaz.
2. Outlook eklerinin cikarildigi gecici dizin is bitince SILINIR. Ekler kisisel
   veri tasir; Windows'ta %TEMP% altinda birikmesi kabul edilemez. Bu makinede
   olculdu: temizlik eklenmeden once 123 kalinti dizin vardi.
"""

from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

KOK = Path(__file__).resolve().parent.parent


def _calistir_modulu():
    yol = KOK / "paketle" / "calistir.py"
    spec = importlib.util.spec_from_file_location("calistir_test_modulu", yol)
    modul = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = modul
    spec.loader.exec_module(modul)
    return modul


class ArsivTest(unittest.TestCase):
    def setUp(self):
        self.m = _calistir_modulu()
        self.gecici = tempfile.TemporaryDirectory()
        self.kok = Path(self.gecici.name)
        self.m.dizinleri_hazirla(self.kok)
        self.fatura = self.kok / self.m.FATURA_DIZINI

    def tearDown(self):
        self.gecici.cleanup()

    def test_islenen_dosyalar_arsive_tasinir_alt_klasor_korunur(self):
        (self.fatura / "temmuz").mkdir()
        a = self.fatura / "a.msg"; a.write_bytes(b"x")
        b = self.fatura / "temmuz" / "b.xlsx"; b.write_bytes(b"y")
        tasinan, hatalar = self.m.islenenleri_arsivle(self.kok, [a, b], "20260906_1200")
        self.assertEqual(hatalar, [])
        self.assertEqual(len(tasinan), 2)
        arsiv = self.kok / self.m.ARSIV_DIZINI / "20260906_1200"
        self.assertTrue((arsiv / "a.msg").is_file())
        self.assertTrue((arsiv / "temmuz" / "b.xlsx").is_file())
        self.assertFalse(a.exists())
        self.assertFalse((self.fatura / "temmuz").exists(), "bosalan alt klasor silinmeli")

    def test_isaret_dosyasi_yerinde_kalir(self):
        isaret = self.fatura / "BURAYA_FATURA_ATIN.txt"
        isaret.write_text("isaret", encoding="utf-8")
        a = self.fatura / "a.msg"; a.write_bytes(b"x")
        self.m.islenenleri_arsivle(self.kok, [a, isaret], "d")
        self.assertTrue(isaret.is_file())

    def test_disaridan_gelen_dosyaya_dokunulmaz(self):
        """Surukle-birak ile verilen dosya kullanicinin kendi klasorundedir."""
        with tempfile.TemporaryDirectory() as baska:
            dis = Path(baska) / "dis.msg"; dis.write_bytes(b"z")
            tasinan, _ = self.m.islenenleri_arsivle(self.kok, [dis], "d")
            self.assertEqual(tasinan, [])
            self.assertTrue(dis.is_file())

    def test_ayni_adli_dosya_ezilmez(self):
        arsiv = self.kok / self.m.ARSIV_DIZINI / "d"
        arsiv.mkdir(parents=True)
        (arsiv / "a.msg").write_bytes(b"eski")
        a = self.fatura / "a.msg"; a.write_bytes(b"yeni")
        tasinan, _ = self.m.islenenleri_arsivle(self.kok, [a], "d")
        self.assertEqual((arsiv / "a.msg").read_bytes(), b"eski")
        self.assertEqual(len(tasinan), 1)
        self.assertEqual(tasinan[0].read_bytes(), b"yeni")

    def test_calistirma_kaydi_yazilir(self):
        class S:  # minimal Sonuc
            def __init__(self, d): self.durum = d
        kayit = self.m.calistirma_kaydi_yaz(
            self.kok, "d", [self.fatura / "a.msg"], None, None,
            [S("OTOMATIK"), S("INCELE")], None, "x.xlsx", [])
        self.assertIsNotNone(kayit)
        metin = kayit.read_text(encoding="utf-8")
        self.assertIn("Okunan satir   : 2", metin)
        self.assertTrue((self.kok / self.m.ARSIV_DIZINI / "d" / "OZET.txt").is_file())


class GeciciDizinTemizligiTest(unittest.TestCase):
    """Outlook eklerinin cikarildigi gecici dizin okuma bitince silinmeli."""

    MESAJ = KOK / "ornek_veri" / "posta" / "ornek_mail.msg"

    def test_msg_okunduktan_sonra_gecici_dizin_kalmaz(self):
        if not self.MESAJ.is_file():
            self.skipTest("ornek mesaj yok")
        import glob, logging
        logging.disable(logging.WARNING)
        from masraf.okuyucular.kesif import oku
        once = set(glob.glob(str(Path(tempfile.gettempdir()) / "masraf_msg_*")))
        satirlar = oku(self.MESAJ)
        self.assertGreater(len(satirlar), 0)
        sonra = set(glob.glob(str(Path(tempfile.gettempdir()) / "masraf_msg_*")))
        self.assertEqual(sonra - once, set(), "gecici dizin geride kaldi")

    def test_kullanicinin_verdigi_dizine_dokunulmaz(self):
        if not self.MESAJ.is_file():
            self.skipTest("ornek mesaj yok")
        import logging
        logging.disable(logging.WARNING)
        from masraf.okuyucular.kesif import oku
        with tempfile.TemporaryDirectory() as g:
            oku(self.MESAJ, cikarma_dizini=g)
            self.assertTrue(any(Path(g).iterdir()), "kullanicinin dizini silinmemeli")


if __name__ == "__main__":
    unittest.main(verbosity=2)
