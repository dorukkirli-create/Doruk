"""Kademeli eslestiricinin altin ornek testleri (masraf.eslestirici).

Buradaki ornekler Temmuz 2026 seyahat dosyasindan ELLE dogrulanmis gercek
vakalardir; adlar ve siciller depoya girmez, ``ornek_veri/altin.json``
dosyasindan okunur (bkz. testler/altin.py). Her biri farkli bir kademeyi
temsil eder:

    tam_isim         AD SOYAD birebir
    bitisik ad       'SOYAD ADAD' (iki ad bitisik yazilmis)
    transliterasyon  Kiril'den geri cevrilmis yazim (GE->GO, KH->H, IY->YI)
    prefix (kesik)   bilet sisteminin 20 karakterde kestigi isim
    aile             es/cocuk, calisan uzerinden
    yok              dis danisman, eslesmemeli
    cakisma          ayni isimde cok calisan, otomatik kabul EDILMEZ

Testler ogrenen defterlerin BOS bir kopyasiyla calisir; boylece
``veri/aliases.csv`` icindeki birikmis duzeltmeler sonucu maskelemez.
Yani her altin ornek, sifirdan kurulan bir sistemde de gecmelidir.
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
    from masraf.eslestirici import Eslestirici, durum_belirle
    from masraf.kayit import PersonelDefteri
    from masraf.modeller import DURUM_OTOMATIK, YONTEMLER, GiderSatiri

    MODUL_VAR = True
except ImportError:
    MODUL_VAR = False

PERSONEL = KOK / "ornek_veri" / "personel" / "2025_2026_giris_cikis.xlsx"

#: Personel ana verisi 24 MB'dir; modul basina bir kez yuklenir.
_ORTAM: dict = {}


def setUpModule() -> None:
    """Personel defterini ve BOS ogrenen defterleri bir kez hazirlar."""
    if not MODUL_VAR or not PERSONEL.exists():
        return
    gecici = tempfile.mkdtemp(prefix="masraf_test_eslestirici_")
    harita = KOK / "veri" / "masraf_merkezi_haritasi.csv"
    if harita.exists():
        shutil.copy(harita, Path(gecici) / harita.name)
    _ORTAM["gecici"] = gecici
    _ORTAM["defter"] = PersonelDefteri.yukle(PERSONEL)
    _ORTAM["defterler"] = Defterler(gecici)


def tearDownModule() -> None:
    """Gecici defter dizinini siler."""
    gecici = _ORTAM.pop("gecici", None)
    if gecici:
        shutil.rmtree(gecici, ignore_errors=True)
    _ORTAM.clear()


def gider(kisi: str, gider_tipi: str = "Bilet", satir_no: int = 1,
          **ek) -> "GiderSatiri":
    """Test icin tek kisilik bir gider satiri uretir."""
    return GiderSatiri(
        kaynak_dosya="test.xls",
        kaynak_tip="antik_cari",
        satir_no=satir_no,
        belge_tarihi=ek.pop("belge_tarihi", None),
        aciklama=ek.pop("aciklama", kisi),
        kisi_ham=kisi,
        sicil_ham=ek.pop("sicil_ham", None),
        tckn_ham=ek.pop("tckn_ham", None),
        tutar=ek.pop("tutar", None),
        para_birimi=ek.pop("para_birimi", "USD"),
        masraf_merkezi_kaynak=ek.pop("masraf_merkezi_kaynak", None),
        gider_tipi=gider_tipi,
        ek=ek,
    )


@unittest.skipUnless(MODUL_VAR, "masraf.eslestirici bulunamadi")
class EslestiriciTemeli(unittest.TestCase):
    """Modul duzeyinde yuklenen defteri tum test siniflarina dagitir."""

    def setUp(self):
        if not PERSONEL.exists():
            self.skipTest(f"Ornek veri bulunamadi: {PERSONEL}")
        # Her test dosya ici aile kanitindan bagimsiz baslar.
        self.eslestirici = Eslestirici(_ORTAM["defter"], _ORTAM["defterler"])

    def esle(self, kisi: str, gider_tipi: str = "Bilet"):
        return self.eslestirici.esle(gider(kisi, gider_tipi))


class AltinOrneklerTest(EslestiriciTemeli):
    """Bilinen dogru eslesmeler."""

    def test_tam_isim(self):
        o = altin("eslestirici", "tam_isim")
        sonuc = self.esle(o["isim"], "Otel")
        self.assertEqual(sonuc.sicil, o["sicil"])
        self.assertEqual(sonuc.yontem, "tam_isim")
        self.assertGreaterEqual(sonuc.guven, 0.90)
        self.assertEqual(sonuc.aday_sayisi, 1)

    def test_isim_sirasi_onemsiz(self):
        # Otel satirlarinda 'AD SOYAD', bilet satirlarinda 'SOYAD AD' gelir;
        # ikisi de ayni sicile dusmelidir.
        o = altin("eslestirici", "bitisik")
        duz = self.esle(o["duz"], "Otel")
        ters = self.esle(o["ters"], "Bilet")
        self.assertEqual(duz.sicil, o["sicil"])
        self.assertEqual(ters.sicil, o["sicil"])

    def test_bitisik_ad(self):
        # 'ADAD' -> 'AD AD' acilmali.
        o = altin("eslestirici", "bitisik")
        sonuc = self.esle(o["bitisik"])
        self.assertEqual(sonuc.sicil, o["sicil"])
        self.assertEqual(sonuc.ad_soyad, o["ad_soyad"])
        self.assertGreaterEqual(sonuc.guven, 0.85)

    def test_kesik_isim_ayni_dosyada_cozulur(self):
        # Bilet sistemi ismi 20 karakterde kesebiliyor: 'SOYAD/ADAD' -> 'SOYAD/ADA'.
        # Ayni dosyada tam yazim da varsa satir dogru sicile baglanmalidir.
        o = altin("eslestirici", "bitisik")
        satirlar = [
            gider(o["bitisik"], "Bilet", 1),
            gider(o["kesik"], "Bilet", 2),
        ]
        sonuclar = self.eslestirici.esle_toplu(satirlar)
        self.assertEqual(sonuclar[0].sicil, o["sicil"])
        self.assertEqual(sonuclar[1].sicil, o["sicil"])
        # Kesik isim kesin degildir; otomatik kabul edilmemelidir.
        self.assertNotEqual(durum_belirle(sonuclar[1]), DURUM_OTOMATIK)

    def test_kesik_isim_tek_basina_adaylari_verir(self):
        # Tam yazim dosyada yoksa bile dogru kisi ADAY olarak listelenmelidir.
        o = altin("eslestirici", "bitisik")
        sonuc = self.esle(o["kesik"])
        self.assertIn(o["sicil"], sonuc.aday_siciller)
        self.assertNotEqual(durum_belirle(sonuc), DURUM_OTOMATIK)

    def test_prefix_kesik_isim(self):
        o = altin("eslestirici", "prefix")
        sonuc = self.esle(o["isim"])
        self.assertEqual(sonuc.sicil, o["sicil"])
        self.assertEqual(sonuc.yontem, "prefix")

    def test_transliterasyon_ge_go_kh_h(self):
        # GE->GO, KH->H, IY->YI
        o = altin("eslestirici", "translit1")
        sonuc = self.esle(o["isim"])
        self.assertEqual(sonuc.sicil, o["sicil"])
        self.assertEqual(sonuc.yontem, "transliterasyon")
        self.assertEqual(sonuc.ad_soyad, o["ad_soyad"])

    def test_transliterasyon_yr_ir_ei_ey(self):
        # YR->IR, KH->H, EI->EY
        o = altin("eslestirici", "translit2")
        sonuc = self.esle(o["isim"])
        self.assertEqual(sonuc.sicil, o["sicil"])
        self.assertEqual(sonuc.yontem, "transliterasyon")

    def test_sicil_dogrudan_verilirse_kullanilir(self):
        # Koc katilimci listesinde ID = sicil; isim eslestirmeye gerek yok.
        o = altin("aktif_sicil")
        satir = gider(o["ad_soyad"], "Egitim", sicil_ham=o["sicil"])
        sonuc = self.eslestirici.esle(satir)
        self.assertEqual(sonuc.sicil, o["sicil"])
        self.assertEqual(sonuc.yontem, "sicil")
        self.assertGreaterEqual(sonuc.guven, 0.95)

    def test_yontem_adlari_sozlesmede_tanimli(self):
        e = altin("eslestirici")
        for kisi in (e["tam_isim"]["isim"], e["translit1"]["isim"],
                     e["harici"]["isim"], e["bitisik"]["bitisik"]):
            with self.subTest(kisi=kisi):
                self.assertIn(self.esle(kisi).yontem, YONTEMLER)

    def test_aciklama_turkce_ve_dolu(self):
        # Kullanici neden o sonuca varildigini gormeli.
        e = altin("eslestirici")
        for kisi in (e["tam_isim"]["isim"], e["translit1"]["isim"], e["harici"]["isim"]):
            with self.subTest(kisi=kisi):
                aciklama = self.esle(kisi).aciklama
                self.assertTrue(aciklama and aciklama.strip())
                self.assertGreaterEqual(len(aciklama), 20)


class AileKuraliTest(EslestiriciTemeli):
    """Es ve cocuklarin biletleri calisanin masraf merkezine yazilir."""

    def test_es_ve_cocuk_aile_bireyi(self):
        o = altin("eslestirici", "aile1")
        satirlar = [
            gider(o["calisan"], "Bilet", 1),
            gider(o["es"], "Bilet", 2),
            gider(o["cocuk"], "Bilet", 3),
        ]
        sonuclar = self.eslestirici.esle_toplu(satirlar)
        calisan, es, cocuk = sonuclar

        self.assertEqual(calisan.sicil, o["sicil"])
        self.assertEqual(calisan.yontem, "tam_isim")

        for sonuc in (es, cocuk):
            self.assertEqual(sonuc.yontem, "aile")
            self.assertIn(o["sicil"], sonuc.aday_siciller)
            # Aile bireyi kesin degildir: guven dusuk, otomatik kabul yok.
            self.assertLess(sonuc.guven, 0.80)
            self.assertNotEqual(durum_belirle(sonuc), DURUM_OTOMATIK)

    def test_ad_soyad_sirali_aile_bireyi(self):
        o = altin("eslestirici", "aile2")
        satirlar = [
            gider(o["calisan"], "Bilet", 1),
            gider(o["uye"], "Otel", 2),
        ]
        sonuclar = self.eslestirici.esle_toplu(satirlar)
        self.assertEqual(sonuclar[0].sicil, o["sicil"])
        self.assertEqual(sonuclar[1].yontem, "aile")
        self.assertIn(o["sicil"], sonuclar[1].aday_siciller)

    def test_esle_toplu_sira_bagimsiz(self):
        # Aile bireyi calisandan ONCE gelse de ayni sonuc cikmali.
        o = altin("eslestirici", "aile1")
        ters = [
            gider(o["es"], "Bilet", 1),
            gider(o["calisan"], "Bilet", 2),
        ]
        sonuclar = self.eslestirici.esle_toplu(ters)
        self.assertEqual(sonuclar[0].yontem, "aile")
        self.assertIn(o["sicil"], sonuclar[0].aday_siciller)
        self.assertEqual(sonuclar[1].sicil, o["sicil"])


class EslesmeyenlerTest(EslestiriciTemeli):
    """Personel ana verisinde gercekten olmayan kisiler."""

    def test_dis_danisman_eslesmez(self):
        # Kurumsal gelisim koclugu veren dis danisman; personel degil.
        sonuc = self.esle(altin("eslestirici", "harici", "isim"), "Vize")
        self.assertEqual(sonuc.yontem, "yok")
        self.assertIsNone(sonuc.sicil)
        self.assertEqual(sonuc.guven, 0.0)

    def test_uydurma_isim_eslesmez(self):
        sonuc = self.esle("QWXZVBN PLKJHGF")
        self.assertEqual(sonuc.yontem, "yok")
        self.assertIsNone(sonuc.sicil)

    def test_kisi_adi_olmayan_satir(self):
        satir = gider("", "Diger", aciklama="CENAZE CELENGI")
        sonuc = self.eslestirici.esle(satir)
        self.assertEqual(sonuc.yontem, "yok")
        self.assertIsNone(sonuc.sicil)


class CakismaTest(EslestiriciTemeli):
    """Ayni isimde birden fazla calisan varsa otomatik kabul YASAK."""

    def test_cok_adayli_isim_otomatik_kabul_edilmez(self):
        # Ayni isim 8 farkli sicilde geciyor (Hindistanli iscilerde
        # isim cakismasi %10,9). Sistem tahmin YURUTMEMELI.
        sonuc = self.esle(altin("eslestirici", "cakisma", "isim"), "Otel")
        self.assertGreater(sonuc.aday_sayisi, 1,
                           "Cakisan isim icin birden fazla aday beklenir")
        self.assertLess(sonuc.guven, 0.80)
        self.assertNotEqual(durum_belirle(sonuc), DURUM_OTOMATIK)
        self.assertTrue(sonuc.aday_siciller,
                        "Incelemeye giden satirda adaylar listelenmeli")

    def test_cakisma_adaylari_incelemede_gosterilir(self):
        sonuc = self.esle(altin("eslestirici", "cakisma", "isim2"), "Otel")
        self.assertGreater(sonuc.aday_sayisi, 1)
        self.assertNotEqual(durum_belirle(sonuc), DURUM_OTOMATIK)


if __name__ == "__main__":
    unittest.main()
