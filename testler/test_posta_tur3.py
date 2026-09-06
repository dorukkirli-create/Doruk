"""Tur 3 posta dayanikliligi (masraf.okuyucular.posta + kesif._msg_oku_icerik).

Bulgular:
1. Tekrar tespiti ad + boyutla degil ICERIK ozetiyle (sha256) yapilir.
2. Sifreli / bozuk zip mailin tamamini dusurmez; AtlananEk olarak raporlanir.
3. Zip icindeki zip acilir; guvenlik sinirlari mail genelinde uygulanir.
4. Bozuk ek, derinlik asimi ve ic mail hatasi AtlananEk olur; ic mailler ve
   arsivler envanterde MAIL / ARSIV bilgi satiri olarak gorunur.
5. Raporlarda diske yazilan kisaltilmis / '_2' ekli ad degil ORIJINAL ek adi.

Gercek .msg uretmek Outlook ister; bu yuzden sahte mesaj nesneleri kullanilir
ve `extract_msg.openMsg` yamalanir. Gercek kisi adi ya da gercek veri YOKTUR.
"""

from __future__ import annotations

import io
import shutil
import struct
import tempfile
import unittest
import zipfile
from datetime import datetime
from pathlib import Path
from unittest import mock

try:
    import extract_msg
    EXTRACT_MSG_VAR = True
except ImportError:  # pragma: no cover - ortama bagli
    extract_msg = None
    EXTRACT_MSG_VAR = False

from masraf.envanter import ARSIV, ATLANDI, AYNI_ICERIK, MAIL, OKUNAMADI, OKUNDU
from masraf.okuyucular import posta
from masraf.okuyucular.posta import AtlananEk, Kapsayici, msg_aciklarini_cikar

ORNEK_MAIL = Path("ornek_veri/posta/ornek_mail.msg")


# ---------------------------------------------------------------------------
# Sahte Outlook nesneleri ve yardimcilar
# ---------------------------------------------------------------------------

class SahteEk:
    """extract_msg Attachment gibi: longFilename / shortFilename / data."""

    def __init__(self, ad, veri):
        self.longFilename = ad
        self.shortFilename = ad
        self._veri = veri

    @property
    def data(self):
        if isinstance(self._veri, BaseException):
            raise self._veri
        return self._veri


class SahteMesaj:
    """extract_msg Message gibi: subject / sender / date / attachments / close."""

    def __init__(self, konu, ekler, gonderen="tedarikci@ornek.test"):
        self.subject = konu
        self.sender = gonderen
        self.date = datetime(2026, 7, 1)
        self.attachments = ekler

    def close(self):
        pass


class BozukMesaj:
    """attachments ozelligi AttributeError DISINDA bir hata firlatan ic mail."""

    subject = "bozuk"
    sender = "tedarikci@ornek.test"
    date = datetime(2026, 7, 1)

    @property
    def attachments(self):
        raise RuntimeError("OLE akisi okunamadi")

    def close(self):
        pass


def csv_bayt(kisi: str, tutar: int) -> bytes:
    """Ayni uzunlukta ad + ayni basamakli tutar => ayni bayt boyutu, farkli icerik."""
    return f"Ad Soyad,Tutar\n{kisi},{tutar}\n".encode("utf-8")


def zip_bayt(girdiler: list[tuple[str, bytes]]) -> bytes:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        for ad, veri in girdiler:
            z.writestr(ad, veri)
    return b.getvalue()


def sifreli_zip_bayt(girdiler: list[tuple[str, bytes]]) -> bytes:
    """Sifreli bayragi (bit 0) set edilmis zip; pyzipper gerekmez.

    zipfile.writestr bayraklari sifirladigi icin ham baytlar sonradan yamalanir.
    zipfile bu girdiyi acarken 'password required' RuntimeError firlatir;
    gercek parola korumali arsivle birebir ayni davranis.
    """
    ham = bytearray(zip_bayt(girdiler))
    for imza, ofset in ((b"PK\x03\x04", 6), (b"PK\x01\x02", 8)):
        i = 0
        while True:
            i = ham.find(imza, i)
            if i < 0:
                break
            (bayrak,) = struct.unpack("<H", ham[i + ofset:i + ofset + 2])
            ham[i + ofset:i + ofset + 2] = struct.pack("<H", bayrak | 1)
            i += 4
    return bytes(ham)


A_CSV = csv_bayt("ALI VELI", 100)
B_CSV = csv_bayt("CAN DELI", 900)
X_CSV = csv_bayt("AYSE ATA", 10)


class _Temel(unittest.TestCase):
    def setUp(self):
        self.gecici = Path(tempfile.mkdtemp(prefix="posta_tur3_"))

    def tearDown(self):
        shutil.rmtree(self.gecici, ignore_errors=True)

    def yuru(self, msg):
        """_msg_yuru'yu dogrudan calistirir; (tablolar, atlananlar, kapsayicilar)."""
        sonuc: list = []
        atl: list = []
        kaps: list = []
        y = posta._Yuruyus(sonuc=sonuc, atlananlar=atl, kapsayicilar=kaps)
        posta._msg_yuru(msg, self.gecici, [], 0, sonuc, atl, yuruyus=y)
        return sonuc, atl, kaps

    def _msg_dosyasi(self) -> Path:
        yol = self.gecici / "sahte.msg"
        yol.write_bytes(b"sahte outlook mesaji")
        return yol

    def cikar(self, msg, dosya_mesajlari: dict | None = None, **kw):
        """msg_aciklarini_cikar'i sahte mesajla ucundan ucuna calistirir.

        ``dosya_mesajlari``: disk adi -> mesaj; bayt olarak eklenmis .msg
        dosyalari icin openMsg bu haritadan doner.
        """
        if not EXTRACT_MSG_VAR:
            self.skipTest("extract_msg yok")
        yol = self._msg_dosyasi()

        def _ac(p):
            return (dosya_mesajlari or {}).get(Path(p).name, msg)

        with mock.patch.object(extract_msg, "openMsg", side_effect=_ac):
            return msg_aciklarini_cikar(yol, self.gecici / "acilan", **kw)

    def kesif_oku(self, msg, envanter: list, atlananlar: list | None = None,
                  gorulen: set | None = None):
        """kesif._msg_oku_icerik'i sahte mesajla calistirir; satirlari dondurur."""
        if not EXTRACT_MSG_VAR:
            self.skipTest("extract_msg yok")
        from masraf.okuyucular import kesif
        yol = self._msg_dosyasi()
        with mock.patch.object(extract_msg, "openMsg", return_value=msg):
            return kesif._msg_oku_icerik(yol, self.gecici / "acilan", gorulen, atlananlar, envanter)


# ---------------------------------------------------------------------------
# 1. Ayni ad + ayni boyut, farkli icerik
# ---------------------------------------------------------------------------

class SaTekillestirmeTest(_Temel):
    def test_on_kosul_ayni_boyut_farkli_icerik(self):
        self.assertEqual(len(A_CSV), len(B_CSV))
        self.assertNotEqual(A_CSV, B_CSV)

    def _iki_fatura_maili(self, a=A_CSV, b=B_CSV, ad_a="Fatura Detayi.csv", ad_b="Fatura Detayi.csv"):
        return SahteMesaj("ust", [
            SahteEk("m1.msg", SahteMesaj("RE: 1", [SahteEk(ad_a, a)])),
            SahteEk("m2.msg", SahteMesaj("RE: 2", [SahteEk(ad_b, b)])),
        ])

    def test_farkli_icerik_ikisi_de_okunur(self):
        """Tedarikcinin fatura basina gonderdigi tek satirlik sablon: ikincisi para tasir."""
        atl: list = []
        ekler = self.cikar(self._iki_fatura_maili(), atlananlar=atl)
        self.assertEqual(len(ekler), 2, [str(a) for a in atl])
        self.assertEqual([e.gosterim_adi for e in ekler], ["Fatura Detayi.csv"] * 2)
        self.assertNotEqual(ekler[0].yol, ekler[1].yol)
        self.assertEqual({e.yol.read_bytes() for e in ekler}, {A_CSV, B_CSV})
        self.assertEqual([a for a in atl if a.tekrar], [])
        # Her ek icerik ozetini tasir; ikisi farkli
        self.assertTrue(all(e.ozet for e in ekler))
        self.assertNotEqual(ekler[0].ozet, ekler[1].ozet)

    def test_ayni_icerik_tekrar_sayilir(self):
        atl: list = []
        ekler = self.cikar(self._iki_fatura_maili(A_CSV, A_CSV), atlananlar=atl)
        self.assertEqual(len(ekler), 1)
        tekrarlar = [a for a in atl if a.tekrar]
        self.assertEqual(len(tekrarlar), 1)
        self.assertEqual(tekrarlar[0].ad, "Fatura Detayi.csv")
        self.assertIn("icerigi", tekrarlar[0].sebep)
        self.assertNotIn("ad ve boyut", tekrarlar[0].sebep)   # eski yanlis gerekce
        self.assertEqual(tekrarlar[0].zincir, ["ust", "RE_ 2"])

    def test_ayni_icerik_farkli_ad_da_tekrar(self):
        """Ad farkli olsa da icerik ayniysa para iki kez sayilmaz; ilk kopya belirtilir."""
        atl: list = []
        ekler = self.cikar(self._iki_fatura_maili(A_CSV, A_CSV, "Temmuz.csv", "Temmuz (kopya).csv"),
                           atlananlar=atl)
        self.assertEqual([e.gosterim_adi for e in ekler], ["Temmuz.csv"])
        self.assertEqual(len(atl), 1)
        self.assertTrue(atl[0].tekrar)
        self.assertIn("Temmuz.csv", atl[0].sebep)

    def test_kesif_envanterde_iki_fatura_da_okunur(self):
        env: list = []
        satirlar = self.kesif_oku(self._iki_fatura_maili(), env, gorulen=set())
        self.assertEqual(sorted(s.tutar for s in satirlar), [100.0, 900.0])
        okunan = [k for k in env if k.durum == OKUNDU]
        self.assertEqual([k.ad for k in okunan], ["Fatura Detayi.csv"] * 2)
        self.assertEqual([k for k in env if k.durum == AYNI_ICERIK], [])
        # Ad tabanli kontroller (boru._ayni_adli_dosyayi_ele) maildeki adi gorur
        self.assertEqual({s.kaynak_dosya for s in satirlar}, {"sahte.msg > Fatura Detayi.csv"})
        self.assertEqual(sorted(k.kaynak for k in okunan),
                         ["sahte.msg > ust > RE_ 1", "sahte.msg > ust > RE_ 2"])

    def test_kesif_ayni_icerik_envanterde_ayni_icerik(self):
        env: list = []
        satirlar = self.kesif_oku(self._iki_fatura_maili(A_CSV, A_CSV), env, gorulen=set())
        self.assertEqual([s.tutar for s in satirlar], [100.0])
        ayni = [k for k in env if k.durum == AYNI_ICERIK]
        self.assertEqual(len(ayni), 1)
        self.assertIn("icerigi", ayni[0].sebep)
        kok = next(k for k in env if k.durum == MAIL and not k.kaynak)
        self.assertIn("1 tablo eki", kok.sebep)
        self.assertIn("1 tekrar eden ek", kok.sebep)


# ---------------------------------------------------------------------------
# 2. Sifreli / bozuk arsiv
# ---------------------------------------------------------------------------

class SifreliVeBozukArsivTest(_Temel):
    def test_sifreli_zip_maili_dusurmez(self):
        msg = SahteMesaj("k", [
            SahteEk("once.csv", A_CSV),
            SahteEk("sifreli.zip", sifreli_zip_bayt([("gizli.csv", B_CSV)])),
            SahteEk("sonra.csv", X_CSV),
        ])
        sonuc, atl, kaps = self.yuru(msg)   # istisna firlatmamali
        self.assertEqual([e.gosterim_adi for e in sonuc], ["once.csv", "sonra.csv"])
        self.assertEqual([a.ad for a in atl], ["gizli.csv"])
        self.assertIn("sifreli", atl[0].sebep)
        self.assertIn("parola", atl[0].sebep)
        self.assertEqual(atl[0].zincir, ["k", "sifreli.zip"])
        self.assertEqual(atl[0].boyut, len(B_CSV))
        arsivler = [k for k in kaps if k.tur == "arsiv"]
        self.assertEqual([k.ad for k in arsivler], ["sifreli.zip"])
        self.assertIn("sifreli", arsivler[0].aciklama)

    def test_sifreli_girdi_digerlerini_engellemez(self):
        z = sifreli_zip_bayt([("gizli.csv", B_CSV)])
        # Ayni arsive sifresiz bir girdi de ekle: ham baytlari birlestirmek yerine
        # iki arsivi ayri ekleyip 'girdi bazinda' davranisi _zip_ac_ayrintili ile olc.
        yol = self.gecici / "karisik.zip"
        b = io.BytesIO()
        with zipfile.ZipFile(b, "w") as zf:
            zf.writestr("acik.csv", A_CSV)
            zf.writestr("gizli.csv", B_CSV)
        ham = bytearray(b.getvalue())
        # yalnizca 'gizli.csv' basliklarini sifreli isaretle
        for imza, ad_ofset, bayrak_ofset in ((b"PK\x03\x04", 30, 6), (b"PK\x01\x02", 46, 8)):
            i = 0
            while True:
                i = ham.find(imza, i)
                if i < 0:
                    break
                if ham[i + ad_ofset:i + ad_ofset + 9] == b"gizli.csv":
                    (bayrak,) = struct.unpack("<H", ham[i + bayrak_ofset:i + bayrak_ofset + 2])
                    ham[i + bayrak_ofset:i + bayrak_ofset + 2] = struct.pack("<H", bayrak | 1)
                i += 4
        yol.write_bytes(bytes(ham))
        del z
        sonuc = posta._zip_ac_ayrintili(yol, self.gecici / "acilmis")
        self.assertIsNone(sonuc.hata)
        self.assertEqual([g.orijinal_ad for g in sonuc.cikanlar], ["acik.csv"])
        self.assertEqual([(ad, boyut) for ad, _, boyut in sonuc.atlananlar], [("gizli.csv", len(B_CSV))])
        self.assertEqual(sonuc.atlananlar[0][1], posta.SEBEP_SIFRELI)

    def test_bozuk_zip_atlanir_digerleri_okunur(self):
        msg = SahteMesaj("k", [
            SahteEk("bozuk.zip", b"bu bir zip degil"),
            SahteEk("tablo.csv", A_CSV),
        ])
        sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual([e.gosterim_adi for e in sonuc], ["tablo.csv"])
        self.assertEqual([a.ad for a in atl], ["bozuk.zip"])
        self.assertIn("acilamadi", atl[0].sebep)
        self.assertEqual([k.ad for k in kaps], ["bozuk.zip"])
        self.assertIn("acilamadi", kaps[0].aciklama)

    def test_zip_ac_geriye_uyumlu(self):
        """Eski `_zip_ac` sarmali: liste doner, sifreli arsivde istisna firlatmaz."""
        normal = self.gecici / "n.zip"
        normal.write_bytes(zip_bayt([("k/veri.csv", A_CSV)]))
        self.assertEqual([p.name for p in posta._zip_ac(normal, self.gecici / "n")], ["veri.csv"])
        sifreli = self.gecici / "s.zip"
        sifreli.write_bytes(sifreli_zip_bayt([("gizli.csv", A_CSV)]))
        self.assertEqual(posta._zip_ac(sifreli, self.gecici / "s"), [])

    def test_ucundan_ucuna_sifreli_zip(self):
        """msg_aciklarini_cikar + kesif: sifreli zip OKUNAMADI degil ATLANDI satiri, mail okunur."""
        msg = SahteMesaj("k", [
            SahteEk("once.csv", A_CSV),
            SahteEk("sifreli.zip", sifreli_zip_bayt([("gizli.csv", B_CSV)])),
        ])
        env: list = []
        atl: list = []
        satirlar = self.kesif_oku(msg, env, atl)
        self.assertEqual([s.tutar for s in satirlar], [100.0])
        atlanan = [k for k in env if k.durum == ATLANDI]
        self.assertEqual([(k.ad, k.tur) for k in atlanan], [("gizli.csv", "csv")])
        self.assertIn("parola", atlanan[0].sebep)
        self.assertEqual(atlanan[0].kaynak, "sahte.msg > k > sifreli.zip")
        self.assertEqual([k.ad for k in env if k.durum == ARSIV], ["sifreli.zip"])
        self.assertTrue(any(str(a).startswith("gizli.csv") for a in atl))


# ---------------------------------------------------------------------------
# 3. Zip icinde zip ve guvenlik sinirlari
# ---------------------------------------------------------------------------

class IcIceZipTest(_Temel):
    def test_zip_icinde_zip_acilir(self):
        msg = SahteMesaj("k", [SahteEk("dis.zip", zip_bayt([("ic.zip", zip_bayt([("fatura.csv", A_CSV)]))]))])
        sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual([e.gosterim_adi for e in sonuc], ["fatura.csv"])
        self.assertEqual(sonuc[0].zincir, ["k", "dis.zip", "ic.zip"])
        self.assertEqual(sonuc[0].derinlik, 2)
        self.assertEqual(sonuc[0].yol.read_bytes(), A_CSV)
        self.assertEqual(atl, [])
        self.assertEqual([(k.ad, k.zincir) for k in kaps], [("dis.zip", ["k"]), ("ic.zip", ["k", "dis.zip"])])
        self.assertTrue(all(k.tur == "arsiv" for k in kaps))

    def test_uc_seviye_ve_yan_dosyalar(self):
        ic = zip_bayt([("en_ic.csv", A_CSV), ("not.txt", b"x")])
        orta = zip_bayt([("ic.zip", ic), ("orta.csv", B_CSV)])
        msg = SahteMesaj("k", [SahteEk("dis.zip", zip_bayt([("orta.zip", orta), ("logo.png", b"x")]))])
        sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual(sorted(e.gosterim_adi for e in sonuc), ["en_ic.csv", "orta.csv"])
        self.assertEqual([(a.ad, a.zincir) for a in atl], [("not.txt", ["k", "dis.zip", "orta.zip", "ic.zip"])])
        self.assertEqual([k.ad for k in kaps], ["dis.zip", "orta.zip", "ic.zip"])

    def test_derinlik_siniri_arsivde_de_gecerli(self):
        icic = zip_bayt([("derin.csv", A_CSV)])
        ic = zip_bayt([("icic.zip", icic)])
        msg = SahteMesaj("k", [SahteEk("dis.zip", zip_bayt([("ic.zip", ic), ("yakin.csv", B_CSV)]))])
        with mock.patch.object(posta, "AZAMI_DERINLIK", 1):
            sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual([e.gosterim_adi for e in sonuc], ["yakin.csv"])
        self.assertEqual([a.ad for a in atl], ["icic.zip"])
        self.assertIn("derin", atl[0].sebep)

    def test_girdi_sayisi_siniri_raporlanir(self):
        cok = zip_bayt([(f"d{i}.csv", A_CSV) for i in range(3)])
        msg = SahteMesaj("k", [SahteEk("cok.zip", cok), SahteEk("tek.csv", B_CSV)])
        with mock.patch.object(posta, "AZAMI_DOSYA_SAYISI", 2):
            sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual([e.gosterim_adi for e in sonuc], ["tek.csv"])
        self.assertEqual([a.ad for a in atl], ["cok.zip"])
        self.assertIn("guvenlik siniri", atl[0].sebep)
        self.assertIn("guvenlik siniri", kaps[0].aciklama)

    def test_boyut_siniri_raporlanir(self):
        buyuk = zip_bayt([("b.csv", b"0" * 4096)])
        msg = SahteMesaj("k", [SahteEk("buyuk.zip", buyuk)])
        with mock.patch.object(posta, "AZAMI_ACILMIS_BOYUT", 1024):
            sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual(sonuc, [])
        self.assertEqual([a.ad for a in atl], ["buyuk.zip"])
        self.assertIn("MB", atl[0].sebep)

    def test_butce_mail_genelinde_uygulanir(self):
        """Ic ice arsivlerde toplam acilan girdi sayisi izlenir (zip bombasi)."""
        z1 = zip_bayt([("a.csv", A_CSV), ("b.csv", B_CSV)])
        z2 = zip_bayt([("c.csv", X_CSV), ("d.txt", b"x")])
        msg = SahteMesaj("k", [SahteEk("z1.zip", z1), SahteEk("z2.zip", z2)])
        with mock.patch.object(posta, "AZAMI_DOSYA_SAYISI", 3):
            sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual(sorted(e.gosterim_adi for e in sonuc), ["a.csv", "b.csv"])
        self.assertEqual([a.ad for a in atl], ["z2.zip"])
        self.assertIn("toplami", atl[0].sebep)

    def test_zip_icindeki_msg_ve_tekrar(self):
        """Zip icindeki tablo, dogrudan ekle ayni icerikse tekrar sayilir."""
        msg = SahteMesaj("k", [SahteEk("a.csv", A_CSV), SahteEk("z.zip", zip_bayt([("kopya/a.csv", A_CSV)]))])
        atl: list = []
        ekler = self.cikar(msg, atlananlar=atl)
        self.assertEqual(len(ekler), 1)
        self.assertEqual([(a.ad, a.tekrar, a.zincir) for a in atl], [("a.csv", True, ["k", "z.zip"])])


# ---------------------------------------------------------------------------
# 4. Bozuk ek, derinlik asimi, ic mail hatasi, kapsayici envanteri
# ---------------------------------------------------------------------------

class BozukEkVeDerinlikTest(_Temel):
    def test_bozuk_ek_yalnizca_kendini_dusurur(self):
        msg = SahteMesaj("k", [
            SahteEk("a.docx", b"x"), SahteEk("logo.png", b"x"), SahteEk("ekadsiz", b"x"),
            SahteEk("bozuk.csv", OSError("OLE akisi okunamadi")), SahteEk("x.csv", A_CSV),
        ])
        sonuc, atl, kaps = self.yuru(msg)   # istisna yok
        self.assertEqual([e.gosterim_adi for e in sonuc], ["x.csv"])
        self.assertEqual([a.ad for a in atl], ["a.docx", "ekadsiz", "bozuk.csv"])
        bozuk = atl[-1]
        self.assertIn("okunamadi", bozuk.sebep)
        self.assertIn("OSError", bozuk.sebep)
        self.assertNotIn("logo.png", [a.ad for a in atl])   # gorseller tasarim geregi sessiz

    def test_veri_tasimayan_ek_raporlanir(self):
        msg = SahteMesaj("k", [SahteEk("bulut.xlsx", None), SahteEk("x.csv", A_CSV)])
        sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual([e.gosterim_adi for e in sonuc], ["x.csv"])
        self.assertEqual([a.ad for a in atl], ["bulut.xlsx"])
        self.assertIn("icerigi mailde yok", atl[0].sebep)

    def test_derinlik_asimi_raporlanir(self):
        m = SahteMesaj("d8", [SahteEk("derin.csv", A_CSV)])
        for i in range(7, -1, -1):
            m = SahteMesaj(f"d{i}", [SahteEk(f"d{i}.msg", m)])
        sonuc, atl, kaps = self.yuru(m)
        self.assertEqual(sonuc, [])
        self.assertEqual(len(atl), 1)
        self.assertEqual(atl[0].ad, "d6.msg")
        self.assertIn("derin", atl[0].sebep)
        self.assertEqual(atl[0].zincir, [f"d{i}" for i in range(7)])
        # Sinira kadar olan ic mailler kapsayici olarak kaydedildi
        self.assertEqual([k.ad for k in kaps], [f"d{i}.msg" for i in range(6)])

    def test_ic_mail_hatasi_ana_maili_dusurmez(self):
        msg = SahteMesaj("k", [SahteEk("m1.msg", BozukMesaj()), SahteEk("x.csv", A_CSV)])
        sonuc, atl, kaps = self.yuru(msg)
        self.assertEqual([e.gosterim_adi for e in sonuc], ["x.csv"])
        self.assertEqual([a.ad for a in atl], ["m1.msg"])
        self.assertIn("RuntimeError", atl[0].sebep)

    def test_bayt_olarak_ekli_bozuk_msg_dosyasi(self):
        """Bir .msg dosyasi bayt olarak eklenmis ama acilmiyor: yalnizca o ek atlanir."""
        if not EXTRACT_MSG_VAR:
            self.skipTest("extract_msg yok")
        kok = SahteMesaj("k", [SahteEk("bozuk.msg", b"bozuk"), SahteEk("x.csv", A_CSV)])
        yol = self._msg_dosyasi()

        def _ac(p):
            if Path(p).name == "bozuk.msg":
                raise ValueError("OLE degil")
            return kok

        atl: list = []
        with mock.patch.object(extract_msg, "openMsg", side_effect=_ac):
            ekler = msg_aciklarini_cikar(yol, self.gecici / "acilan", atlananlar=atl)
        self.assertEqual([e.gosterim_adi for e in ekler], ["x.csv"])
        self.assertEqual([a.ad for a in atl], ["bozuk.msg"])
        self.assertIn("acilamadi", atl[0].sebep)

    def test_kok_mail_acilamiyorsa_hata_yine_verilir(self):
        if not EXTRACT_MSG_VAR:
            self.skipTest("extract_msg yok")
        yol = self._msg_dosyasi()
        with mock.patch.object(extract_msg, "openMsg", side_effect=ValueError("OLE degil")):
            with self.assertRaises(posta.MesajAcilamadi):
                msg_aciklarini_cikar(yol, self.gecici / "acilan")

    def test_eski_konumsal_imza_calisir(self):
        """_msg_yuru(msg, hedef, zincir, derinlik, sonuc, atlananlar) eski cagri bicimi."""
        sonuc: list = []
        atl: list = []
        posta._msg_yuru(SahteMesaj("k", [SahteEk("a.csv", A_CSV), SahteEk("b.pdf", b"x")]),
                        self.gecici, [], 0, sonuc, atl)
        self.assertEqual([e.gosterim_adi for e in sonuc], ["a.csv"])
        self.assertEqual([a.ad for a in atl], ["b.pdf"])


class KapsayiciEnvanterTest(_Temel):
    def _mail(self):
        return SahteMesaj("ust", [
            SahteEk("m1.msg", SahteMesaj("RE: 1", [SahteEk("Fatura Detayi.csv", A_CSV), SahteEk("f.pdf", b"x")])),
            SahteEk("m2.msg", SahteMesaj("RE: 2", [SahteEk("Fatura Detayi.csv", B_CSV)])),
            SahteEk("dis.zip", zip_bayt([("ic.zip", zip_bayt([("gomulu.csv", X_CSV)])),
                                         ("logo.png", b"x"), ("liste.csv", X_CSV)])),
            SahteEk("imza.png", b"x"),
            SahteEk("bos.xlsx", b"bozuk"),
        ])

    def test_kapsayicilar_ve_ekler_envanterde(self):
        env: list = []
        atl: list = []
        satirlar = self.kesif_oku(self._mail(), env, atl, gorulen=set())
        self.assertEqual(sorted(s.tutar for s in satirlar), [10.0, 100.0, 900.0])
        durumlar = {}
        for k in env:
            durumlar.setdefault(k.durum, []).append(k.ad)
        self.assertEqual(sorted(durumlar[MAIL]), ["m1.msg", "m2.msg", "sahte.msg"])
        self.assertEqual(sorted(durumlar[ARSIV]), ["dis.zip", "ic.zip"])
        self.assertEqual(durumlar[ATLANDI], ["f.pdf"])                 # gorseller haric
        self.assertEqual(durumlar[AYNI_ICERIK], ["liste.csv"])          # gomulu.csv ile ayni icerik
        self.assertEqual(sorted(durumlar[OKUNDU]), ["Fatura Detayi.csv", "Fatura Detayi.csv", "gomulu.csv"])
        self.assertEqual(durumlar[OKUNAMADI], ["bos.xlsx"])
        self.assertEqual(len(env), 11)
        # Kaynak zincirleri
        kaynaklar = {(k.durum, k.ad): k.kaynak for k in env}
        self.assertEqual(kaynaklar[(MAIL, "sahte.msg")], "")
        self.assertEqual(kaynaklar[(MAIL, "m1.msg")], "sahte.msg > ust")
        self.assertEqual(kaynaklar[(ARSIV, "dis.zip")], "sahte.msg > ust")
        self.assertEqual(kaynaklar[(ARSIV, "ic.zip")], "sahte.msg > ust > dis.zip")
        self.assertEqual(kaynaklar[(OKUNDU, "gomulu.csv")], "sahte.msg > ust > dis.zip > ic.zip")
        self.assertEqual(kaynaklar[(ATLANDI, "f.pdf")], "sahte.msg > ust > RE_ 1")
        # Kok mail satiri tam sayim verir
        kok = next(k for k in env if k.durum == MAIL and not k.kaynak)
        for parca in ("4 tablo eki", "1 okunmayan ek", "1 tekrar eden ek", "2 ekli mail", "2 arsiv"):
            self.assertIn(parca, kok.sebep)
        # Ic mail ve arsiv satirlari bilgi tasir
        m1 = next(k for k in env if k.durum == MAIL and k.ad == "m1.msg")
        self.assertIn("2 ek", m1.sebep)
        self.assertEqual(m1.tur, "outlook_msg")
        dis = next(k for k in env if k.durum == ARSIV and k.ad == "dis.zip")
        self.assertIn("3 girdi", dis.sebep)
        self.assertEqual(dis.tur, "zip")
        # Kapsayicilar 'okunmayan ek' listesine girmez (Kontrol sayfasi)
        self.assertEqual([str(a).split("  [")[0] for a in atl if not a.tekrar], ["f.pdf"])

    def test_kapsayicilar_atlananlara_karismaz(self):
        msg = SahteMesaj("k", [SahteEk("m1.msg", SahteMesaj("ic", [SahteEk("a.csv", A_CSV)])),
                               SahteEk("z.zip", zip_bayt([("b.csv", B_CSV)]))])
        atl: list = []
        kaps: list = []
        ekler = self.cikar(msg, atlananlar=atl, kapsayicilar=kaps)
        self.assertEqual(sorted(e.gosterim_adi for e in ekler), ["a.csv", "b.csv"])
        self.assertEqual(atl, [])
        self.assertEqual([(k.ad, k.tur, k.ek_sayisi) for k in kaps], [("m1.msg", "mail", 1), ("z.zip", "arsiv", 1)])
        self.assertTrue(all(isinstance(k, Kapsayici) for k in kaps))


# ---------------------------------------------------------------------------
# 5. Orijinal ek adi
# ---------------------------------------------------------------------------

class OrijinalAdTest(_Temel):
    UZUN = "ASS2026000009999 - 010720261200 ENERGO Fatura Detayi Listesi Uzun Uzun Uzun Ad.csv"

    def test_uzun_ad_diskte_kisa_raporda_tam(self):
        self.assertGreater(len(self.UZUN), 60)
        sonuc, atl, kaps = self.yuru(SahteMesaj("k", [SahteEk(self.UZUN, A_CSV)]))
        (ek,) = sonuc
        self.assertLessEqual(len(ek.ad), 60)
        self.assertTrue(ek.ad.endswith(".csv"))
        self.assertEqual(ek.gosterim_adi, self.UZUN)
        self.assertEqual(ek.orijinal_ad, self.UZUN)
        self.assertTrue(ek.kaynak_aciklamasi.endswith(self.UZUN))
        self.assertTrue(ek.yol.is_file())

    def test_es_adli_ekler_ayni_klasorde(self):
        """Ayni mailde iki 'Fatura Detayi.csv': diskte _2, raporda orijinal ad."""
        sonuc, atl, kaps = self.yuru(SahteMesaj("k", [SahteEk("Fatura Detayi.csv", A_CSV),
                                                       SahteEk("Fatura Detayi.csv", B_CSV)]))
        self.assertEqual(sorted(e.ad for e in sonuc), ["Fatura Detayi.csv", "Fatura Detayi_2.csv"])
        self.assertEqual([e.gosterim_adi for e in sonuc], ["Fatura Detayi.csv"] * 2)

    def test_kesif_es_adli_ekler_orijinal_adla(self):
        env: list = []
        satirlar = self.kesif_oku(SahteMesaj("k", [SahteEk("Fatura Detayi.csv", A_CSV),
                                                   SahteEk("Fatura Detayi.csv", B_CSV)]), env, gorulen=set())
        self.assertEqual({s.kaynak_dosya for s in satirlar}, {"sahte.msg > Fatura Detayi.csv"})
        self.assertEqual([k.ad for k in env if k.durum == OKUNDU], ["Fatura Detayi.csv"] * 2)
        self.assertNotIn("_2", " ".join(k.ad for k in env))
        for s in satirlar:
            self.assertEqual(s.ek["mail_zinciri"], "k > Fatura Detayi.csv")

    def test_kesif_uzun_ad_envanter_ve_satirda(self):
        env: list = []
        satirlar = self.kesif_oku(SahteMesaj("k", [SahteEk(self.UZUN, A_CSV)]), env, gorulen=set())
        self.assertEqual([k.ad for k in env if k.durum == OKUNDU], [self.UZUN])
        self.assertEqual({s.kaynak_dosya for s in satirlar}, {f"sahte.msg > {self.UZUN}"})

    def test_yasak_karakterli_ad(self):
        ad = "Fatura: Temmuz/2026 <taslak>.csv"
        sonuc, atl, kaps = self.yuru(SahteMesaj("k", [SahteEk(ad, A_CSV)]))
        (ek,) = sonuc
        self.assertEqual(ek.gosterim_adi, ad)
        self.assertNotIn("/", ek.ad)
        self.assertNotIn(":", ek.ad)

    def test_zip_kacislari_gosterim_adinda_cozulur(self):
        z = zip_bayt([("Kat#U0131l#U0131mc#U0131 Listesi.csv", A_CSV)])
        sonuc, atl, kaps = self.yuru(SahteMesaj("k", [SahteEk("liste.zip", z)]))
        self.assertEqual([e.gosterim_adi for e in sonuc], ["Katılımcı Listesi.csv"])

    def test_atlanan_ek_orijinal_adi_tasir(self):
        uzun_pdf = "ENERGO 2026 Temmuz Fatura Ekleri Toplu Liste Uzun Uzun Uzun Uzun Ad Dosyasi.pdf"
        self.assertGreater(len(uzun_pdf), 60)
        sonuc, atl, kaps = self.yuru(SahteMesaj("k", [SahteEk(uzun_pdf, b"%PDF")]))
        self.assertEqual([a.ad for a in atl], [uzun_pdf])
        self.assertTrue(str(atl[0]).startswith(uzun_pdf))
        self.assertIsInstance(atl[0], AtlananEk)

    def test_msg_aciklarini_cikar_imzasi_geriye_uyumlu(self):
        msg = SahteMesaj("k", [SahteEk("a.csv", A_CSV), SahteEk("b.pdf", b"x")])
        self.assertEqual([e.gosterim_adi for e in self.cikar(msg)], ["a.csv"])
        atl: list = []
        if not EXTRACT_MSG_VAR:
            self.skipTest("extract_msg yok")
        yol = self._msg_dosyasi()
        with mock.patch.object(extract_msg, "openMsg", return_value=msg):
            ekler = msg_aciklarini_cikar(yol, self.gecici / "acilan2", atl)   # konumsal atlananlar
        self.assertEqual([e.gosterim_adi for e in ekler], ["a.csv"])
        self.assertEqual([a.ad for a in atl], ["b.pdf"])


# ---------------------------------------------------------------------------
# 6. Tablo olmayan eklerin (PDF) tekrari
# ---------------------------------------------------------------------------

class TabloOlmayanTekrarTest(_Temel):
    PDF_A = b"%PDF-1.4 sahte fatura A"
    PDF_B = b"%PDF-1.4 sahte fatura B"

    def _mail(self, ic_mailde=PDF_A, zipte=PDF_A):
        """Ayni adli PDF hem ic mailde hem zip icinde."""
        return SahteMesaj("ust", [
            SahteEk("m1.msg", SahteMesaj("RE: 1", [SahteEk("fatura.pdf", ic_mailde), SahteEk("a.csv", A_CSV)])),
            SahteEk("ekler.zip", zip_bayt([("fatura.pdf", zipte), ("b.csv", B_CSV)])),
        ])

    def test_ayni_pdf_ikincisi_tekrar(self):
        atl: list = []
        ekler = self.cikar(self._mail(), atlananlar=atl)
        self.assertEqual(sorted(e.gosterim_adi for e in ekler), ["a.csv", "b.csv"])
        self.assertEqual([(a.ad, a.tekrar) for a in atl], [("fatura.pdf", False), ("fatura.pdf", True)])
        self.assertTrue(atl[0].ozet)
        self.assertEqual(atl[0].ozet, atl[1].ozet)
        self.assertIn("birebir ayni", atl[1].sebep)
        self.assertEqual(atl[1].zincir, ["ust", "ekler.zip"])
        # Kapak/Kontrol (cikti.py) tekrar olmayanlari sayar: 1 farkli PDF
        self.assertEqual(len([a for a in atl if not a.tekrar]), 1)

    def test_farkli_pdf_ikisi_de_sayilir(self):
        atl: list = []
        self.cikar(self._mail(self.PDF_A, self.PDF_B), atlananlar=atl)
        self.assertEqual([(a.ad, a.tekrar) for a in atl], [("fatura.pdf", False), ("fatura.pdf", False)])
        self.assertNotEqual(atl[0].ozet, atl[1].ozet)

    def test_kesif_envanterde_pdf_tekrari_ayni_icerik(self):
        env: list = []
        atl: list = []
        self.kesif_oku(self._mail(), env, atl, gorulen=set())
        self.assertEqual([k.ad for k in env if k.durum == ATLANDI], ["fatura.pdf"])
        ayni = [k for k in env if k.durum == AYNI_ICERIK]
        self.assertEqual([(k.ad, k.tur) for k in ayni], [("fatura.pdf", "pdf")])
        self.assertTrue(ayni[0].ozet)
        kok = next(k for k in env if k.durum == MAIL and not k.kaynak)
        self.assertIn("1 okunmayan ek", kok.sebep)
        self.assertIn("1 tekrar eden ek", kok.sebep)

    def test_mailler_arasi_pdf_tekrari(self):
        gorulen: set = set()
        env: list = []
        atl: list = []
        m1 = SahteMesaj("k1", [SahteEk("fatura.pdf", self.PDF_A), SahteEk("a.csv", A_CSV)])
        m2 = SahteMesaj("k2", [SahteEk("fatura.pdf", self.PDF_A), SahteEk("b.csv", B_CSV)])
        self.kesif_oku(m1, env, atl, gorulen)
        self.kesif_oku(m2, env, atl, gorulen)
        self.assertEqual([(a.ad, a.tekrar) for a in atl], [("fatura.pdf", False), ("fatura.pdf", True)])
        self.assertIn("baska bir mailde", atl[1].sebep)
        self.assertEqual([k.ad for k in env if k.durum == ATLANDI], ["fatura.pdf"])
        self.assertEqual([k.ad for k in env if k.durum == AYNI_ICERIK], ["fatura.pdf"])
        # Tablolar etkilenmedi: a.csv ve b.csv farkli, ikisi de okundu
        self.assertEqual(sorted(k.ad for k in env if k.durum == OKUNDU), ["a.csv", "b.csv"])

    def test_sifreli_girdinin_ozeti_yok_tekrar_sayilmaz(self):
        z = sifreli_zip_bayt([("gizli.pdf", self.PDF_A)])
        msg = SahteMesaj("k", [SahteEk("s1.zip", z), SahteEk("s2.zip", z)])
        atl: list = []
        self.cikar(msg, atlananlar=atl)
        self.assertEqual([(a.ad, a.tekrar, a.ozet) for a in atl],
                         [("gizli.pdf", False, None), ("gizli.pdf", False, None)])


# ---------------------------------------------------------------------------
# 7. Kisa cikarma dizinleri (Windows MAX_PATH) ve yazma hatasi
# ---------------------------------------------------------------------------

class KisaDizinTest(_Temel):
    KISA_KLASOR = r"^[mz]\d{2}_[0-9a-f]{6}$"
    UZUN_KONU = "Cok uzun bir mail konusu, tedarikci yazismasi " * 5   # ~230 karakter

    def test_klasor_adlari_kisa_ve_konudan_bagimsiz(self):
        ic = SahteMesaj(self.UZUN_KONU + " ic", [
            SahteEk(OrijinalAdTest.UZUN, A_CSV),
            SahteEk("arsiv.zip", zip_bayt([("a/b/c/d/e/f/g/h/veri.csv", B_CSV)])),
        ])
        m = ic
        for i in range(4):
            m = SahteMesaj(f"{self.UZUN_KONU} {i}", [SahteEk(f"ic{i}.msg", m)])
        sonuc, atl, kaps = self.yuru(m)
        self.assertEqual(sorted(e.gosterim_adi for e in sonuc), sorted([OrijinalAdTest.UZUN, "veri.csv"]))
        self.assertEqual(atl, [])
        klasorler = [p for p in self.gecici.rglob("*") if p.is_dir()]
        self.assertGreaterEqual(len(klasorler), 6)
        for p in klasorler:
            self.assertRegex(p.name, self.KISA_KLASOR)
        en_uzun = max(len(str(p.relative_to(self.gecici))) for p in self.gecici.rglob("*"))
        self.assertLess(en_uzun, 120, en_uzun)
        for e in sonuc:
            self.assertLessEqual(len(e.ad), posta.DISK_AD_AZAMI)
            # Kullaniciya gorunen zincir konulardan olusur, klasor adlarindan degil
            self.assertTrue(all(z.startswith("Cok uzun") or z == "arsiv.zip" for z in e.zincir), e.zincir)
            self.assertNotRegex(e.zincir[0], self.KISA_KLASOR)

    def test_ayni_konulu_ic_mailler_ayri_klasor(self):
        m = SahteMesaj("ust", [
            SahteEk("m1.msg", SahteMesaj("RE: ayni", [SahteEk("Fatura.csv", A_CSV)])),
            SahteEk("m2.msg", SahteMesaj("RE: ayni", [SahteEk("Fatura.csv", B_CSV)])),
        ])
        sonuc, atl, kaps = self.yuru(m)
        self.assertEqual(len({e.yol.parent for e in sonuc}), 2)
        self.assertEqual([e.ad for e in sonuc], ["Fatura.csv", "Fatura.csv"])   # _2 gerekmedi
        self.assertEqual([e.zincir for e in sonuc], [["ust", "RE_ ayni"]] * 2)

    def test_zip_girdileri_duzlestirilir(self):
        z = zip_bayt([("Klasor Adi Uzun/Alt/veri.csv", A_CSV), ("Klasor Adi Uzun/Alt2/veri.csv", B_CSV)])
        sonuc, atl, kaps = self.yuru(SahteMesaj("k", [SahteEk("a.zip", z)]))
        self.assertEqual([e.gosterim_adi for e in sonuc], ["veri.csv", "veri.csv"])
        self.assertEqual(sorted(e.ad for e in sonuc), ["veri.csv", "veri_2.csv"])
        self.assertEqual(len({e.yol.parent for e in sonuc}), 1)
        self.assertRegex(sonuc[0].yol.parent.name, self.KISA_KLASOR)
        self.assertEqual({e.yol.read_bytes() for e in sonuc}, {A_CSV, B_CSV})

    def test_yazma_hatasi_yalnizca_o_eki_dusurur(self):
        gercek = Path.write_bytes

        def sahte(self_, veri):
            if self_.name.startswith("uzun"):
                raise OSError(206, "The filename or extension is too long")
            return gercek(self_, veri)

        msg = SahteMesaj("k", [SahteEk("uzun.csv", A_CSV), SahteEk("kisa.csv", B_CSV), SahteEk("uzun.pdf", b"%PDF")])
        with mock.patch.object(Path, "write_bytes", sahte):
            sonuc, atl, kaps = self.yuru(msg)   # istisna yok
        self.assertEqual([e.gosterim_adi for e in sonuc], ["kisa.csv"])
        self.assertEqual([a.ad for a in atl], ["uzun.csv", "uzun.pdf"])
        for a in atl:
            self.assertIn("gecici dizine yazilamadi", a.sebep)
            self.assertIn("OSError", a.sebep)
        self.assertEqual(atl[0].boyut, len(A_CSV))

    def test_arsiv_dizini_olusturulamazsa_hata(self):
        z = self.gecici / "a.zip"
        z.write_bytes(zip_bayt([("v.csv", A_CSV)]))
        engel = self.gecici / "engel"
        engel.write_bytes(b"dosya")          # dizin yerine dosya: mkdir OSError verir
        sonuc = posta._zip_ac_ayrintili(z, engel)
        self.assertEqual(sonuc.cikanlar, [])
        self.assertIn("gecici dizine acilamadi", sonuc.hata)


# ---------------------------------------------------------------------------
# 8. Outlook 365 bulut eki (extract_msg WebAttachment)
# ---------------------------------------------------------------------------

class BulutEk:
    """extract_msg WebAttachment gibi: ad ve veri NotImplementedError firlatir."""

    url = "https://ornek.sharepoint.test/sites/finans/Fatura%20Temmuz.xlsx?web=1"

    @property
    def longFilename(self):
        raise NotImplementedError("Cannot get the filename of a web attachment.")

    @property
    def shortFilename(self):
        raise NotImplementedError("Cannot get the filename of a web attachment.")

    @property
    def data(self):
        raise NotImplementedError("Cannot get the data of a web attachment.")


class AdsizBulutEk(BulutEk):
    url = None


class AdliBulutEk(BulutEk):
    """Adi okunabilen ama verisi olmayan bulut eki."""

    longFilename = "Sozlesme.pdf"
    shortFilename = "Sozlesme.pdf"


class BulutEkiTest(_Temel):
    def test_bulut_eki_atlanir_digerleri_okunur(self):
        msg = SahteMesaj("k", [SahteEk("a.csv", A_CSV), BulutEk(), SahteEk("b.csv", B_CSV)])
        sonuc, atl, kaps = self.yuru(msg)   # getattr NotImplementedError'i yutmaz; istisna olmamali
        self.assertEqual([e.gosterim_adi for e in sonuc], ["a.csv", "b.csv"])
        self.assertEqual(len(atl), 1)
        self.assertEqual(atl[0].ad, "Fatura Temmuz.xlsx")          # baglantidan turetildi
        self.assertIn("bulut baglantisi", atl[0].sebep)
        self.assertIn("elle indirin", atl[0].sebep)
        self.assertEqual(atl[0].zincir, ["k"])
        self.assertFalse(atl[0].tekrar)

    def test_adsiz_bulut_eki(self):
        sonuc, atl, kaps = self.yuru(SahteMesaj("k", [AdsizBulutEk()]))
        self.assertEqual(sonuc, [])
        self.assertEqual([(a.ad, a.sebep) for a in atl], [("bulut eki", posta.SEBEP_BULUT)])

    def test_adli_bulut_eki(self):
        sonuc, atl, kaps = self.yuru(SahteMesaj("k", [AdliBulutEk()]))
        self.assertEqual([(a.ad, a.sebep) for a in atl], [("Sozlesme.pdf", posta.SEBEP_BULUT)])

    def test_kesif_envanterde_bulut_eki(self):
        env: list = []
        atl: list = []
        satirlar = self.kesif_oku(SahteMesaj("k", [BulutEk(), SahteEk("a.csv", A_CSV)]), env, atl)
        self.assertEqual([s.tutar for s in satirlar], [100.0])
        atlanan = [k for k in env if k.durum == ATLANDI]
        self.assertEqual([(k.ad, k.tur) for k in atlanan], [("Fatura Temmuz.xlsx", "xlsx")])
        self.assertIn("bulut baglantisi", atlanan[0].sebep)
        kok = next(k for k in env if k.durum == MAIL and not k.kaynak)
        self.assertIn("1 okunmayan ek", kok.sebep)


# ---------------------------------------------------------------------------
# Gercek mail (varsa)
# ---------------------------------------------------------------------------

class GercekMailTur3Test(unittest.TestCase):
    def setUp(self):
        if not EXTRACT_MSG_VAR:
            self.skipTest("extract_msg yok")
        if not ORNEK_MAIL.is_file():
            self.skipTest("ornek mail yok")
        self.gecici = Path(tempfile.mkdtemp(prefix="posta_tur3_gercek_"))

    def tearDown(self):
        shutil.rmtree(getattr(self, "gecici", Path(tempfile.mkdtemp())), ignore_errors=True)

    def test_ozet_kapsayici_ve_orijinal_ad(self):
        atl: list = []
        kaps: list = []
        ekler = msg_aciklarini_cikar(ORNEK_MAIL, self.gecici, atlananlar=atl, kapsayicilar=kaps)
        self.assertGreater(len(ekler), 5)
        ozetler = [e.ozet for e in ekler]
        self.assertTrue(all(ozetler))
        self.assertEqual(len(ozetler), len(set(ozetler)), "ayni icerik iki kez dondu")
        for e in ekler:
            self.assertTrue(e.orijinal_ad)
            self.assertEqual(Path(e.orijinal_ad).suffix.lower(), e.yol.suffix.lower())
        self.assertTrue(kaps, "ic mail / arsiv kapsayicisi bulunamadi")
        self.assertTrue(all(k.tur in ("mail", "arsiv") and k.aciklama for k in kaps))
        for a in atl:
            self.assertTrue(a.ad and a.sebep)
        # Tablo olmayan tekrarlar: farkli PDF sayisi = farkli ozet sayisi
        pdfler = [a for a in atl if a.ad.lower().endswith(".pdf")]
        self.assertTrue(pdfler)
        self.assertEqual(len({a.ozet for a in pdfler if a.ozet}), len([a for a in pdfler if not a.tekrar]))
        # Windows yol siniri: kisa klasor adlari, kisa goreli yollar
        for p in self.gecici.rglob("*"):
            if p.is_dir():
                self.assertRegex(p.name, KisaDizinTest.KISA_KLASOR)
        en_uzun = max(len(str(p.relative_to(self.gecici))) for p in self.gecici.rglob("*"))
        self.assertLess(en_uzun, 130, en_uzun)


if __name__ == "__main__":
    unittest.main(verbosity=2)
