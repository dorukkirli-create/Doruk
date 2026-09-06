"""Kapsam olcumu ve aile kurali iyilestirmelerinin gerileme testleri.

Buradaki ornekler ``testler/kapsam_olc.py`` ile yapilan kapsam olcumunde
ALGORITMA SORUNU olarak tespit edilmis ve duzeltilmis GERCEK vakalardir
(adlar ve siciller depoya girmez; ``ornek_veri/altin.json``den okunur):

    soyad_sonda   'AD AD SOYAD' bilet satiri: soyad SONDA (gider tipi yaniltiyor)
    disi_soyad    Rusca kadin soyadi eki, defterde erkek hali var
    cok_aday      soyad SONDA, birden cok aday listelenmeli
    rusca_ad      GERILEME KORUMASI: Rusca bir AD soyad sanilmamali

Ilk ucu duzeltirken dorduncusunun bozulmamasi sarttir; iki uctan da soyad
aramak, Rusca bir adi soyad sanip alakasiz bir calisana baglama riskini
dogurur. Bu yuzden dorduncu ornek testte acikca yer alir.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
import unittest
from pathlib import Path

KOK = Path(__file__).resolve().parents[1]
if str(KOK) not in sys.path:
    sys.path.insert(0, str(KOK))

from testler.altin import altin, altin_veya_none

try:
    from masraf.defter import Defterler
    from masraf.eslestirici import AILE_SOYAD_BELIRGIN, Eslestirici
    from masraf.kayit import PersonelDefteri
    from masraf.metin import rus_disi_soyad_erkek_hali
    from masraf.modeller import GiderSatiri

    MODUL_VAR = True
except ImportError:  # pragma: no cover - modul eksikse testler atlanir
    MODUL_VAR = False

PERSONEL = KOK / "ornek_veri" / "personel" / "2025_2026_giris_cikis.xlsx"

_ORTAM: dict = {}


def setUpModule() -> None:
    """Personel defterini ve BOS ogrenen defterleri bir kez hazirlar."""
    if not MODUL_VAR or not PERSONEL.exists():
        return
    gecici = tempfile.mkdtemp(prefix="masraf_test_kapsam_")
    _ORTAM["gecici"] = gecici
    _ORTAM["defter"] = PersonelDefteri.yukle(PERSONEL)
    _ORTAM["defterler"] = Defterler(gecici)


def tearDownModule() -> None:
    """Gecici defter dizinini siler."""
    gecici = _ORTAM.pop("gecici", None)
    if gecici:
        shutil.rmtree(gecici, ignore_errors=True)
    _ORTAM.clear()


def gider(kisi: str, gider_tipi: str = "Bilet") -> "GiderSatiri":
    """Test icin tek kisilik bir gider satiri uretir."""
    return GiderSatiri(
        kaynak_dosya="test.xls",
        kaynak_tip="antik_cari",
        satir_no=1,
        belge_tarihi=None,
        aciklama=kisi,
        kisi_ham=kisi,
        sicil_ham=None,
        tckn_ham=None,
        tutar=None,
        para_birimi="USD",
        masraf_merkezi_kaynak=None,
        gider_tipi=gider_tipi,
        ek={},
    )


class DisiSoyadTest(unittest.TestCase):
    """Rusca kadin soyadi ekinin erkek haline cevrilmesi."""

    @unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
    def test_bilinen_ekler_cevrilir(self):
        for kadin, erkek in (
            ("NOVOSELOVA", "NOVOSELOV"),
            ("SHTEYNIKOVA", "SHTEYNIKOV"),
            ("DOBRYNINA", "DOBRYNIN"),
            ("ZHUKOVSKAYA", "ZHUKOVSKIY"),
        ):
            with self.subTest(kadin=kadin):
                self.assertEqual(rus_disi_soyad_erkek_hali(kadin), erkek)

    @unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
    def test_turkce_isimler_bozulmaz(self):
        """Turkce ad ve soyadlar kadin soyadi eki sanilmamalidir."""
        for token in ("MUSTAFA", "KOCAK", "GOZUKARA", "SIMSEK", "AYSE"):
            with self.subTest(token=token):
                self.assertIsNone(rus_disi_soyad_erkek_hali(token))

    @unittest.skipUnless(MODUL_VAR, "masraf.metin bulunamadi")
    def test_kisa_tokenlar_atlanir(self):
        """'EVA' bir ADDIR; -EVA kurali kisa tokenlara uygulanmamalidir."""
        self.assertIsNone(rus_disi_soyad_erkek_hali("EVA"))
        self.assertIsNone(rus_disi_soyad_erkek_hali("INA"))


@unittest.skipUnless(MODUL_VAR, "masraf.eslestirici bulunamadi")
class AileSoyadKonumuTest(unittest.TestCase):
    """Aile kuralinin soyadi DOGRU uctan almasi."""

    def setUp(self):
        if not PERSONEL.exists():
            self.skipTest(f"Ornek veri bulunamadi: {PERSONEL}")
        self.eslestirici = Eslestirici(_ORTAM["defter"], _ORTAM["defterler"])

    def esle(self, kisi: str, gider_tipi: str = "Bilet"):
        return self.eslestirici.esle(gider(kisi, gider_tipi))

    def test_soyad_olasiligi_ad_ve_soyadi_ayirir(self):
        """'GOZUKARA' soyaddir, 'HASAN' addir; olcu bunu gostermelidir."""
        self.assertGreaterEqual(
            self.eslestirici._soyad_olasiligi("GOZUKARA"), AILE_SOYAD_BELIRGIN)
        self.assertLess(
            self.eslestirici._soyad_olasiligi(altin("kapsam", "soyad_sonda", "ad")),
            AILE_SOYAD_BELIRGIN)

    def test_soyad_sonda_olsa_da_bulunur(self):
        """'AD AD SOYAD' bilet satiri ama isim AD SOYAD sirali.

        Gider tipi 'Bilet' oldugu icin tercih ILK tokendir; ancak ilk token
        bir ADDIR. Dogru cevap sondaki soyaddir.
        """
        o = altin("kapsam", "soyad_sonda")
        sonuc = self.esle(o["isim"])
        self.assertEqual(sonuc.yontem, "aile")
        self.assertEqual(sonuc.sicil, o["sicil"])
        self.assertIn(o["soyad"], sonuc.aciklama)

    def test_disi_soyad_erkek_haliyle_eslesir(self):
        """'SOYADOVA AD' -> defterdeki 'Soyadov ...' (erkek hali)."""
        o = altin("kapsam", "disi_soyad")
        sonuc = self.esle(o["isim"])
        self.assertEqual(sonuc.yontem, "aile")
        self.assertEqual(sonuc.sicil, o["sicil"])
        self.assertIn(o["soyad"], sonuc.aciklama)

    def test_sondaki_soyad_adaylari_listeler(self):
        """Ayni soyadda 4 kisi var, kimlik belirsiz ama adaylar sunulur."""
        o = altin("kapsam", "cok_aday")
        sonuc = self.esle(o["isim"])
        self.assertEqual(sonuc.yontem, "aile")
        self.assertGreater(len(sonuc.aday_siciller), 1)
        self.assertIn(o["soyad"], sonuc.aciklama)

    def test_rusca_ad_soyad_sanilmaz(self):
        """GERILEME KORUMASI: Rusca bir AD soyad sanilmamali.

        Defterde ayni adi tasiyan, AD-SOYAD sirali girilmis tek bir kayit
        vardir; iki uctan da soyad arayan naif bir kural satiri bu alakasiz
        kisiye baglar. Adin soyad olasiligi dusuk oldugu icin eslesme
        OLMAMALIDIR.
        """
        sonuc = self.esle(altin("kapsam", "rusca_ad", "isim"))
        self.assertEqual(sonuc.yontem, "yok")
        self.assertIsNone(sonuc.sicil)

    def test_tercih_edilen_uc_korunur(self):
        """Soyadi ONDE olan normal PNR satirlari degismemelidir.

        Aile bireyi tek satir olarak islendiginde ayni soyadli iki calisan
        arasinda secim yapilamaz; sicil BOS kalir ve ikisi de aday listelenir.
        (Gercek is akisinda ayni dosyada calisanin kendisi kesin eslestigi
        icin secim netlesir; bu test yalniz soyadin ILK uctan alindigini dogrular.)
        """
        o = altin("kapsam", "aile_tek")
        sonuc = self.esle(o["isim"])
        self.assertEqual(sonuc.yontem, "aile")
        self.assertIn(o["soyad"], sonuc.aciklama)
        self.assertIn(o["sicil"], sonuc.aday_siciller)


@unittest.skipUnless(MODUL_VAR, "masraf.eslestirici bulunamadi")
class KapsamOlcumuTest(unittest.TestCase):
    """Olcum aracinin kendisi calisir durumda mi."""

    def setUp(self):
        if not PERSONEL.exists():
            self.skipTest(f"Ornek veri bulunamadi: {PERSONEL}")
        from testler import kapsam_olc

        self.modul = kapsam_olc
        self.arayici = kapsam_olc.ElleArayici(_ORTAM["defter"])

    def test_elle_arama_aday_dondurur(self):
        adaylar = self.arayici.ara(altin("kapsam", "elle_arama"))
        self.assertTrue(adaylar)
        self.assertGreaterEqual(adaylar[0][2], self.modul.KESIN_ESIK)

    def test_soyad_kontrolu_adi_soyad_saymaz(self):
        """Olcum, motorun bilerek reddettigi ADLARI 'soyad var' saymamalidir."""
        k = altin("kapsam")
        self.assertIsNone(self.arayici.soyad_var_mi(k["soyad_ad_degil"]))
        self.assertIsNone(self.arayici.soyad_var_mi(k["rusca_ad"]["isim"]))
        self.assertEqual(self.arayici.soyad_var_mi(k["cok_aday"]["isim"]), k["cok_aday"]["soyad"])

    def test_isim_izi_kurumsal_gideri_ayirir(self):
        """Kisi barindirmayan satir PARSER sorunu sayilmamalidir."""
        self.assertIsNone(
            self.arayici.isim_izi("[13.07.2026] - [14.07.2026] CENAZE CELENK GONDERIMI"))
        self.assertIsNotNone(self.arayici.isim_izi("EKSTRA BAGAJ UCRETI GOZUKARA"))


if __name__ == "__main__":
    unittest.main()
