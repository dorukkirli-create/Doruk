"""Mail govdesi okuyucusu: ozet tablo kurali ve yesil katilim isaretleri.

Sentetik HTML kullanilir; kisi adi ve kimlik numarasi uydurmadir. Mutasyon
senaryolari gercek mail uzerinde olculen hata modlarini kilitler: kurussuz
tutar sessizce dusmemeli, TOPLAM satiri kalem sayilmamali, tarife karti
ozet sanilmamali.
"""

from __future__ import annotations

import unittest
from types import SimpleNamespace

from masraf.envanter import BELGE_TIPLERI
from masraf.modeller import KAYNAK_TIPLERI
from masraf.okuyucular.govde import govde_satirlari, ozet_tablosu, para_coz, yesil_katilimcilar


def _tablo(satirlar, baslik=("", "RHI (Mayis - Haziran)")):
    tr = lambda a, b: f"<tr><td>{a}</td><td>{b}</td></tr>"
    govde = tr(*baslik) if baslik else ""
    govde += "".join(tr(a, b) for a, b in satirlar)
    return f"<html><body><table>{govde}</table></body></html>"


OZET = _tablo([("Egitim", "$ 27.203,48"), ("Danismanlik", "$\xa0\xa0 17.299,00"),
               ("Saglik", "$ 3.000,00"), ("Yemek", "$ 675,70")])


class ParaCozTest(unittest.TestCase):
    def test_turkce_bicim_ve_isaret(self):
        self.assertEqual(para_coz("$ 5.211,77"), (5211.77, "USD"))
        self.assertEqual(para_coz("3.000,00 TL"), (3000.0, "TRY"))
        self.assertEqual(para_coz("21 997,00 USD"), (21997.0, "USD"))
        self.assertEqual(para_coz("$\xa0\xa0675,70"), (675.7, "USD"))

    def test_isaretsiz_tutar_reddedilir(self):
        """Para birimi isareti zorunlu: yanlis pozitiflerin hepsi isaretsizdi."""
        self.assertIsNone(para_coz("5.211,77"))

    def test_abd_bicimi_reddedilir(self):
        self.assertIsNone(para_coz("$ 1,234.56"))

    def test_kurussuz_reddedilir(self):
        self.assertIsNone(para_coz("$ 3.000"))


class OzetTablosuTest(unittest.TestCase):
    def test_ozet_tablo_okunur(self):
        o = ozet_tablosu(OZET)
        self.assertIsNotNone(o)
        self.assertEqual(len(o.kalemler), 4)
        self.assertEqual(o.para_birimi, "USD")
        self.assertEqual(o.toplam, 48178.18)
        self.assertEqual(o.baslik, "RHI (Mayis - Haziran)")
        self.assertIsNone(o.beyan_toplam)

    def test_toplam_satiri_kalem_degil_kontrol_toplamidir(self):
        html = _tablo([("Egitim", "$ 100,00"), ("Saglik", "$ 50,00"), ("TOPLAM", "$ 150,00")])
        o = ozet_tablosu(html)
        self.assertEqual(len(o.kalemler), 2)
        self.assertEqual(o.beyan_toplam, 150.0)
        self.assertTrue(o.beyan_uyusuyor)

    def test_yanlis_toplam_satiri_isaretlenir(self):
        html = _tablo([("Egitim", "$ 100,00"), ("Saglik", "$ 50,00"), ("Toplam", "$ 999,00")])
        o = ozet_tablosu(html)
        self.assertFalse(o.beyan_uyusuyor)

    def test_kurussuz_hucre_tabloyu_reddettirir(self):
        """Sessiz kayip yok: '3.000' cozulemiyorsa tablo bastan reddedilir."""
        html = _tablo([("Egitim", "$ 100,00"), ("Saglik", "$ 3.000")])
        self.assertIsNone(ozet_tablosu(html))

    def test_uc_kolon_reddedilir(self):
        html = ("<table><tr><td>Sirket</td><td>Toplam</td><td>Avans</td></tr>"
                "<tr><td>RHI</td><td>$ 1,00</td><td>$ 2,00</td></tr>"
                "<tr><td>RC</td><td>$ 3,00</td><td>$ 4,00</td></tr></table>")
        self.assertIsNone(ozet_tablosu(html))

    def test_karisik_para_birimi_reddedilir(self):
        html = _tablo([("Egitim", "$ 100,00"), ("Saglik", "50,00 TL")])
        self.assertIsNone(ozet_tablosu(html))

    def test_tekrarlayan_etiket_tarife_kartidir(self):
        """Gercek mailde olculdu: 'premium per person per month' x2 ozet degildir."""
        html = _tablo([("Prim", "$ 2,75"), ("Prim", "$ 1,95")], baslik=None)
        self.assertIsNone(ozet_tablosu(html))

    def test_basliksiz_ve_bos_satirli_tablo_okunur(self):
        html = ("<table><tr><td>Egitim</td><td>$ 100,00</td></tr><tr><td></td><td></td></tr>"
                "<tr><td>Saglik</td><td>$ 50,00</td></tr><tr><td>Yemek</td><td>$ 5,00</td></tr></table>")
        o = ozet_tablosu(html)
        self.assertEqual(len(o.kalemler), 3)

    def test_duz_metin_yedegi(self):
        duz = "Ozet\n\nEgitim\n$ 100,00\n\nSaglik\n$ 50,00\n"
        o = ozet_tablosu("", duz)
        self.assertIsNotNone(o)
        self.assertEqual([k.etiket for k in o.kalemler], ["Egitim", "Saglik"])

    def test_bos_govde(self):
        self.assertIsNone(ozet_tablosu("", ""))


class YesilKatilimTest(unittest.TestCase):
    HTML = (
        "<p>03.06.2026 / 2 kisi</p>"
        "<table><tr><td>ID</td><td>Ad Soyad</td></tr>"
        "<tr><td><span style='background:lime'>100001</span></td><td>Ornek Bir</td></tr>"
        "<tr><td>100002</td><td>Ornek Iki</td></tr>"
        "<tr><td><span style='background:lime'>100003</span></td><td>Ornek Uc</td></tr></table>"
        "<p>04.06.2026 / 1 kisi</p>"
        "<table><tr><td>ID</td><td>Ad Soyad</td></tr>"
        "<tr><td><span style='background:lime'>100001</span></td><td>Ornek Bir</td></tr>"
        "<tr><td>100003</td><td>Ornek Uc</td></tr></table>"
    )

    def test_yesil_satirlar_ve_gunler(self):
        k = {x.kimlik: x for x in yesil_katilimcilar(self.HTML)}
        self.assertEqual(set(k), {"100001", "100003"})
        self.assertEqual(k["100001"].gunler, ["2026-06-03", "2026-06-04"])
        self.assertEqual(k["100003"].gunler, ["2026-06-03"])
        self.assertEqual(sum(x.gun_sayisi for x in k.values()), 3)

    def test_yesil_yoksa_bos(self):
        self.assertEqual(yesil_katilimcilar("<table><tr><td>100001</td><td>A</td></tr></table>"), [])


class TasiyiciSatirTest(unittest.TestCase):
    def test_tasiyicilar_belge_tipi_ve_tutarsiz(self):
        govde = SimpleNamespace(konu="Yansitma", gonderen=None, tarih=None, zincir=[],
                                derinlik=0, html=OZET + YesilKatilimTest.HTML, duz="",
                                kaynak_aciklamasi="Yansitma")
        satirlar = govde_satirlari(govde, "mail.msg")
        tipler = {s.kaynak_tip for s in satirlar}
        self.assertEqual(tipler, {"govde_kalemi", "govde_katilim"})
        self.assertTrue(tipler <= KAYNAK_TIPLERI)
        self.assertTrue(tipler <= BELGE_TIPLERI)          # mahsuba girmez
        self.assertTrue(all(s.tutar is None for s in satirlar))
        kalemler = [s for s in satirlar if s.kaynak_tip == "govde_kalemi"]
        self.assertAlmostEqual(sum(s.ek["govde_tutar"] for s in kalemler), 48178.18, places=2)
        katilim = next(s for s in satirlar if s.kaynak_tip == "govde_katilim")
        self.assertEqual(len(katilim.ek["katilimcilar"]), 2)


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
