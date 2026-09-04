"""Harici kisiler defteri: calisan olmayanlarin masraf merkezine baglanmasi.

Kurumsal gelisim kocu, dis konusmaci, danisman gibi kisiler personel ana
verisinde yoktur ve HICBIR ZAMAN otomatik eslesmez. Dogru davranis onlari
elle tutulan kucuk bir deftere yazmaktir.

Kritik nokta: fatura metinlerinde ayni kisi birden fazla yazimla gecer.
Gercek ornek, Temmuz 2026 faturasi:

    ENERGO ham dokum       -> 'KOCKESEN TALIPKEREM'   (soyad once, ad bitisik)
    Yuzyil elle dagitilmis -> 'TALIP KEREM KOCKESEN'  (duz yazim)

Defterin token kumesi bunlari ayni saymaz cunku 'TALIPKEREM' tek token.
Kullanicinin her yazim icin ayri satir eklemesi beklenemez; unutulan yazim
sessizce dagitilamayan tutara duser. Bu yuzden harf imzasi indeksi var.
"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from masraf.metin import isim_imzasi


class IsimImzasiTest(unittest.TestCase):
    def test_sira_ve_bitisiklik_ayni_imzayi_verir(self):
        self.assertEqual(
            isim_imzasi("KOCKESEN TALIPKEREM"),
            isim_imzasi("TALIP KEREM KOCKESEN"),
        )

    def test_soyad_one_alinmis_hali(self):
        self.assertEqual(
            isim_imzasi("OZAKAY MUSTAFAKEMAL"),
            isim_imzasi("MUSTAFA KEMAL OZAKAY"),
        )

    def test_farkli_kisiler_farkli_imza(self):
        self.assertNotEqual(isim_imzasi("AHMET YILMAZ"), isim_imzasi("MEHMET KAYA"))

    def test_bos_girdi(self):
        self.assertEqual(isim_imzasi(""), "")
        self.assertEqual(isim_imzasi("   "), "")


class HariciDefterEslestirmeTest(unittest.TestCase):
    """Tek bir defter satiri butun yazimlari cozmeli."""

    def _eslestirici(self, veri_dizini: Path):
        from masraf.defter import Defterler
        from masraf.eslestirici import Eslestirici
        from masraf.kayit import PersonelDefteri

        ana = Path("ornek_veri/personel/2025_2026_giris_cikis.xlsx")
        if not ana.is_file():
            raise unittest.SkipTest("ornek personel verisi yok")
        defter = PersonelDefteri.yukle(ana)
        defterler = Defterler(veri_dizini)
        return Eslestirici(defter, defterler)

    def _satir(self, isim: str):
        from masraf.modeller import GiderSatiri

        return GiderSatiri(
            kaynak_dosya="test.xls", kaynak_tip="antik_cari", satir_no=1,
            belge_tarihi=None, aciklama=isim, kisi_ham=isim,
            sicil_ham=None, tckn_ham=None, tutar=100.0, para_birimi="USD",
            masraf_merkezi_kaynak=None, gider_tipi="Bilet",
        )

    def test_tek_satir_iki_yazimi_da_cozer(self):
        with tempfile.TemporaryDirectory() as gecici:
            veri = Path(gecici)
            (veri / "harici_kisiler.csv").write_text(
                "isim_norm;ad_soyad;kurum;masraf_merkezi;aciklama;kaynak;eklenme_tarihi\n"
                "TALIP KEREM KOCKESEN;Talip Kerem Kockesen;Dis konusmaci;"
                "RHI Russia - Headquarter (Moscow);konusmaci;elle;04.09.2026\n",
                encoding="utf-8-sig",
            )
            e = self._eslestirici(veri)
            for yazim in ("TALIP KEREM KOCKESEN", "KOCKESEN TALIPKEREM",
                          "Talip Kerem Kockesen", "KOCKESEN TALIP KEREM"):
                with self.subTest(yazim=yazim):
                    eslesme = e.esle(self._satir(yazim))
                    self.assertEqual(eslesme.yontem, "harici",
                                     f"'{yazim}' harici defterinde bulunamadi")
                    self.assertIn("Moscow", eslesme.aciklama)

    def test_defterde_olmayan_kisi_eslesmez(self):
        with tempfile.TemporaryDirectory() as gecici:
            veri = Path(gecici)
            (veri / "harici_kisiler.csv").write_text(
                "isim_norm;ad_soyad;kurum;masraf_merkezi;aciklama;kaynak;eklenme_tarihi\n"
                "TALIP KEREM KOCKESEN;Talip Kerem Kockesen;Dis konusmaci;"
                "RHI Russia - Headquarter (Moscow);konusmaci;elle;04.09.2026\n",
                encoding="utf-8-sig",
            )
            e = self._eslestirici(veri)
            eslesme = e.esle(self._satir("ZZZZ BILINMEYEN KISI"))
            self.assertNotEqual(eslesme.yontem, "harici")


class GercekFaturaTest(unittest.TestCase):
    """Gercek Temmuz 2026 faturasinda Kockesen'in butun satirlari cozulmeli."""

    ANA = Path("ornek_veri/personel/2025_2026_giris_cikis.xlsx")
    #: Kockesen seyahat dosyasinda gecer; o da bu mesajin ekinde.
    MESAJ = Path("ornek_veri/posta_doruk/yuzyil_temmuz.msg")

    @classmethod
    def setUpClass(cls):
        if not (cls.ANA.is_file() and cls.MESAJ.is_file()):
            raise unittest.SkipTest("ornek veri eksik")

    def test_harici_kayit_butun_satirlari_moskovaya_baglar(self):
        from masraf.boru import Boru, CalismaAyarlari

        with tempfile.TemporaryDirectory() as gecici:
            veri = Path(gecici) / "veri"
            veri.mkdir()
            harita = Path("veri/masraf_merkezi_haritasi.csv")
            if harita.is_file():
                (veri / harita.name).write_bytes(harita.read_bytes())
            (veri / "harici_kisiler.csv").write_text(
                "isim_norm;ad_soyad;kurum;masraf_merkezi;aciklama;kaynak;eklenme_tarihi\n"
                "TALIP KEREM KOCKESEN;Talip Kerem Kockesen;Dis konusmaci;"
                "RHI Russia - Headquarter (Moscow);konusmaci;elle;04.09.2026\n",
                encoding="utf-8-sig",
            )
            ayarlar = CalismaAyarlari(
                personel_yolu=self.ANA, veri_dizini=str(veri),
                cikti_dizini=gecici, defterleri_besle=False,
                ogrenmeyi_kaydet=False,
            )
            sonuclar = Boru(ayarlar).isle([self.MESAJ])
            kockesen = [s for s in sonuclar
                        if s.satir.kisi_ham and "KOCKESEN" in s.satir.kisi_ham.upper()]
            self.assertGreater(len(kockesen), 0, "Kockesen satiri bulunamadi")
            for s in kockesen:
                with self.subTest(yazim=s.satir.kisi_ham):
                    self.assertEqual(s.masraf_merkezi, "HQ-MOSCOW")
                    self.assertEqual(s.eslesme.yontem, "harici")


if __name__ == "__main__":
    unittest.main(verbosity=2)
