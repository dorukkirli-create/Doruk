"""Kaynak dosyadaki santiye kolonunun PROJE olarak karsilastirilmasi.

Isletme karari: kaynak dosyadaki santiye kolonu proje demektir. Personel
kaydindan cozulen proje ile karsilastirilir.

Uc davranis test edilir:

1. Kaynak proje ile cozulen proje AYNI ise uyari uretilmez.
2. FARKLI ise ``PROJE UYUSMAZLIGI`` uyarisi uretilir ve satir incelemeye duser.
3. Kaynakta proje yerine SIRKET adi yazilmissa (RHI, RENSERVIS) satir
   incelemeye DUSURULMEZ; celiski yoktur, sadece isaretlenir ve ozette
   toplu raporlanir.
"""

from __future__ import annotations

import datetime
import unittest
from pathlib import Path

try:
    from masraf.boru import _kaynak_proje_ozeti
    from masraf.kayit import PersonelDefteri
    from masraf.masraf_merkezi import MasrafMerkeziHaritasi, masraf_merkezi_coz
    from masraf.modeller import DURUM_INCELE, Eslesme, GiderSatiri

    MODUL_VAR = True
except ImportError:
    MODUL_VAR = False

PERSONEL = Path("ornek_veri/personel/2025_2026_giris_cikis.xlsx")
HARITA = Path("veri/masraf_merkezi_haritasi.csv")

#: Ozakay Mustafa Kemal - RHI Russia Headquarter (Moscow), tum donemlerde aktif
SICIL = "100003"
TARIH = datetime.date(2026, 7, 15)


def _satir(kaynak_etiket: str | None) -> "GiderSatiri":
    return GiderSatiri(
        kaynak_dosya="test", kaynak_tip="yuzyil_dagitilmis", satir_no=1,
        belge_tarihi=TARIH, aciklama="test", kisi_ham="Ozakay Mustafa Kemal",
        sicil_ham=None, tckn_ham=None, tutar=100.0, para_birimi="USD",
        masraf_merkezi_kaynak=kaynak_etiket, gider_tipi="Bilet",
    )


def _eslesme() -> "Eslesme":
    return Eslesme(sicil=SICIL, ad_soyad="Ozakay Mustafa Kemal", yontem="sicil",
                   guven=1.0, aday_sayisi=1, aciklama="test")


@unittest.skipUnless(MODUL_VAR, "masraf modulleri bulunamadi")
class KaynakProjeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not PERSONEL.is_file():
            raise unittest.SkipTest("ornek_veri/personel bulunamadi")
        cls.defter = PersonelDefteri.yukle(PERSONEL)
        cls.harita = MasrafMerkeziHaritasi.yukle(HARITA)

    def _coz(self, etiket):
        return masraf_merkezi_coz(_satir(etiket), _eslesme(), self.defter, self.harita)

    def test_ayni_proje_uyari_uretmez(self):
        sonuc = self._coz("RHI Russia - Headquarter (Moscow)")
        uyusmazlik = [u for u in sonuc.uyarilar if "UYUSMAZLIGI" in u]
        self.assertFalse(uyusmazlik, f"ayni projede uyari uretildi: {uyusmazlik}")

    def test_esanlamli_proje_adi_da_uyari_uretmez(self):
        """'MOSKOVA' gibi kisayazimlar esanlamli tablosundan cozulmeli."""
        sonuc = self._coz("Moskova")
        uyusmazlik = [u for u in sonuc.uyarilar if "UYUSMAZLIGI" in u]
        self.assertFalse(uyusmazlik, f"esanlamli cozulemedi: {sonuc.uyarilar}")

    def test_farkli_proje_uyusmazlik_uretir(self):
        sonuc = self._coz("GPP Proje")
        uyusmazlik = [u for u in sonuc.uyarilar if "PROJE UYUSMAZLIGI" in u]
        self.assertTrue(uyusmazlik, "farkli projede uyari uretilmedi")
        self.assertEqual(sonuc.durum, DURUM_INCELE)
        self.assertIn("GPP", uyusmazlik[0])

    def test_sirket_adi_incelemeye_dusurmez(self):
        """Kaynakta proje yerine sirket adi varsa satir OTOMATIK kalabilmeli."""
        for etiket in ("RHI", "RENSERVIS", "ONE TOWER"):
            with self.subTest(etiket=etiket):
                sonuc = self._coz(etiket)
                uyusmazlik = [u for u in sonuc.uyarilar if "UYUSMAZLIGI" in u]
                self.assertFalse(uyusmazlik,
                                 f"'{etiket}' proje sanildi: {uyusmazlik}")
                ek = sonuc.satir.ek if isinstance(sonuc.satir.ek, dict) else {}
                self.assertEqual(ek.get("kaynak_proje_yerine_sirket"), etiket,
                                 "sirket etiketi isaretlenmedi")

    def test_gpc_etiketi_gpp_projesine_cozulur(self):
        """Olculdu: saglik listesinde bu etiketi tasiyan 7 kisiden 6'si GPP Project."""
        cozum = self.harita.coz("Ust Luga Gas Processing Complex - GPC")
        self.assertIsNotNone(cozum, "GPC etiketi cozulemedi")
        self.assertEqual(cozum["masraf_merkezi_kodu"], "GPP")

    def test_ozet_gruplari_toplami_tutar(self):
        sonuclar = [
            self._coz("RHI Russia - Headquarter (Moscow)"),  # uyusan
            self._coz("GPP Proje"),                          # uyusmayan
            self._coz("RHI"),                                # proje yok
            self._coz(None),                                 # etiketsiz
        ]
        o = _kaynak_proje_ozeti(sonuclar)
        self.assertEqual(o["uyusan"], 1)
        self.assertEqual(o["uyusmayan"], 1)
        self.assertEqual(o["proje_yok"], 1)
        self.assertEqual(o["etiketsiz"], 1)
        self.assertEqual(o["sirket_etiketleri"], {"RHI": 1})
        self.assertEqual(
            o["uyusan"] + o["uyusmayan"] + o["proje_yok"] + o["etiketsiz"],
            len(sonuclar),
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
