"""Excel ve CSV cikti uretimi.

Uretilen Excel dosyasi finans ekibinin dogrudan calisacagi belgedir. Sayfalar
IS AKISI SIRASINDADIR: once muhasebeye gidecek olan, sonra kaniti.

    Ozet        - KAPAK. Tutarlar, mutabakat durumu, sirket kirilimi, dikkat
                  notlari ve sayfa rehberi. Dosya acilinca ilk gorulen sayfa.
    Mahsuplasma - NIHAI CIKTI. Her fatura icin hangi projeye ne kadar
                  yazilacagi. Muhasebeye giden tablo budur.
    Kontrol     - Mutabakat. Her fatura icin okunan / yinelenen / dagitilan /
                  dagitilamayan tutar. 'Fark' sutunu sifir olmak zorundadir.
    Sonuc       - tum satirlar, tum kolonlar (mahsuplasmanin dayanagi)
    Incele      - durum = INCELE (guven dusuk veya uyari var)
    Eslesmedi   - durum = ESLESMEDI (kisi bulunamadi)
    Ozet        - durum/yontem dagilimi, masraf merkezi bazinda tutar toplami

Tasarim ilkesi: kullanici HER satirda neden o sonuca varildigini gorebilmeli.
Bu yuzden 'Eslestirme Yontemi', 'Guven', 'Eslestirme Aciklamasi' ve 'Uyarilar'
kolonlari ciktida her zaman yer alir ve satirlar duruma gore renklendirilir.

Mahsuplasma sayfasi DUZ bir tablodur, gruplu bir rapor degil. Sebebi pratiktir:
finans bu sayfada filtreleyip pivot yapar. Alt toplamlar Kontrol sayfasindadir.

Yalniz ``xlsxwriter`` ve standart kutuphane kullanilir; modul tek basina
import edilebilir.
"""

from __future__ import annotations

import csv
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

from masraf.modeller import DURUM_ESLESMEDI, DURUM_INCELE, DURUM_OTOMATIK, Sonuc

__all__ = [
    "excel_yaz", "csv_yaz", "mahsuplasma_csv_yaz",
    "KOLONLAR", "MAHSUP_KOLONLARI", "KONTROL_KOLONLARI", "varsayilan_cikti_adi",
]

#: (baslik, tip, genislik). Tip: metin | tarih | sayi | tamsayi | yuzde
KOLONLAR: tuple[tuple[str, str, int], ...] = (
    ("Kaynak Dosya", "metin", 28),
    ("Satir", "tamsayi", 7),
    ("Belge Tarihi", "tarih", 12),
    ("Gider Tipi", "metin", 12),
    ("Aciklama", "metin", 58),
    ("Cikarilan Kisi", "metin", 26),
    ("Sicil", "metin", 10),
    ("Ad Soyad", "metin", 26),
    ("Eslestirme Yontemi", "metin", 16),
    ("Guven", "yuzde", 8),
    ("Aday Sayisi", "tamsayi", 11),
    ("Eslestirme Aciklamasi", "metin", 62),
    ("Donem", "tarih", 11),
    ("Donem Eslesmesi", "metin", 18),
    ("Gorev Yeri", "metin", 32),
    ("Masraf Merkezi Kodu", "metin", 20),
    ("Masraf Merkezi Adi", "metin", 34),
    ("Sirket", "metin", 12),
    ("Statu", "metin", 14),
    ("Kategori", "metin", 10),
    ("Cikis Tarihi", "tarih", 12),
    ("Tutar", "sayi", 13),
    ("Para Birimi", "metin", 10),
    ("Kaynak Dosyadaki Santiye", "metin", 24),
    ("Evrak / Fatura No", "metin", 22),
    ("Mail Konusu", "metin", 40),
    ("Durum", "metin", 12),
    ("Uyarilar", "metin", 70),
)

#: Mahsuplasma sayfasi kolonlari: (baslik, tip, genislik).
#: Sira muhasebecinin okuma sirasidir: once hangi fatura, sonra hangi proje,
#: sonra ne kadar, en sonda kalite isaretleri.
MAHSUP_KOLONLARI: tuple[tuple[str, str, int], ...] = (
    ("Fatura / Kaynak Dosya", "metin", 34),
    # Sirket PROJEDEN ONCE gelir: muhasebe once tuzel kisiyi, sonra o tuzel
    # kisinin altindaki projeyi okur. Kaynagi 1C listesindeki 'Firm 2'dir.
    ("Sirket", "metin", 16),
    ("Masraf Merkezi Kodu", "metin", 22),
    ("Masraf Merkezi Adi", "metin", 34),
    ("Gider Tipi", "metin", 13),
    ("Paylasim", "metin", 20),
    ("Tutar", "sayi", 14),
    ("Para Birimi", "metin", 10),
    ("Fatura Payi", "yuzde", 11),
    ("Satir", "tamsayi", 7),
    ("Kisi", "tamsayi", 7),
    ("Otomatik", "tamsayi", 9),
    ("Incele", "tamsayi", 8),
    ("Eslesmedi", "tamsayi", 10),
    ("Gider Donemi", "metin", 18),
    ("Durum", "metin", 34),
)

#: Kontrol (mutabakat) sayfasi kolonlari.
KONTROL_KOLONLARI: tuple[tuple[str, str, int], ...] = (
    ("Fatura / Kaynak Dosya", "metin", 34),
    ("Para Birimi", "metin", 10),
    ("Okunan Tutar", "sayi", 15),
    ("Yinelenen (baska dosyada sayildi)", "sayi", 18),
    ("Net Tutar", "sayi", 15),
    ("Dagitilan", "sayi", 15),
    ("Dagitilamayan", "sayi", 15),
    ("Fark", "sayi", 11),
    ("Satir", "tamsayi", 7),
    ("Yinelenen Satir", "tamsayi", 14),
    ("Dagitim Orani", "yuzde", 13),
    ("Mutabakat", "metin", 16),
)

#: Mahsup satirinin 'Durum' sutununda gosterilecek kisa uyarilar.
MAHSUP_DURUM_ETIKETLERI: tuple[tuple[str, str], ...] = (
    ("dagitilamadi", "MASRAF MERKEZI YOK"),
    ("haritada_yok", "HARITADA TANIMLI DEGIL"),
    ("incele", "INCELENECEK SATIR VAR"),
    ("eslesmedi", "ESLESMEYEN SATIR VAR"),
)

#: Mahsuplasma sayfasindaki satir renkleri (duruma gore).
MAHSUP_RENKLERI: dict[str, str] = {
    "tamam": "#E2EFDA",
    "uyari": "#FFF2CC",
    "engel": "#FCE4D6",
    # Sirket kirilimindeki ust seviye satirlar
    "baslik": "#D6E4F0",
}

#: Duruma gore satir arka plan renkleri.
DURUM_RENKLERI: dict[str, str] = {
    DURUM_OTOMATIK: "#E2EFDA",
    DURUM_INCELE: "#FFF2CC",
    DURUM_ESLESMEDI: "#FCE4D6",
}

#: Gider ayi ile personel donemi iliskisinin okunakli karsiliklari.
DONEM_ESLESME_ETIKETLERI: dict[str, str] = {
    "tam": "Ayni ay",
    "onceki_donem": "Onceki donem (ayrilmis)",
    "ilk_donem_oncesi": "Ise girmeden once",
    "tarihsiz": "Tarih yok",
    "yok": "",
}

# RHI kurumsal kimligi (McKinsey 2019 sistemine dayanan sirket sablonu).
# Basliklar Georgia, govde Arial; lacivert ve beyaz tasiyici, canli mavi tek vurgu.
LACIVERT = "#051C2C"
BEYAZ = "#FFFFFF"
CANLI_MAVI = "#1F40E6"
CAMGOBEGI = "#00A9F4"
KOYU_GRI = "#4D4D4D"
ORTA_GRI = "#7F7F7F"
KENAR_GRI = "#D0D0D0"
ACIK_GRI = "#F2F4F6"        # zebra; sablondaki E6E6E6 tablo icinde agir kaliyor
BASLIK_YAZI = "Georgia"
GOVDE_YAZI = "Arial"

BASLIK_RENGI = LACIVERT     # geriye uyumluluk
TARIH_BICIMI = "DD.MM.YYYY"
TUTAR_BICIMI = "#,##0.00"
YUZDE_BICIMI = "0.0%"
TAMSAYI_BICIMI = "#,##0"


def varsayilan_cikti_adi(onek: str = "masraf_dagitimi") -> str:
    """Zaman damgali varsayilan dosya adi uretir."""
    return f"{onek}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"


def _metin(deger: Any) -> str:
    """Hucre degerini duz metne cevirir."""
    if deger is None:
        return ""
    metin = str(deger).strip()
    if metin.lower() in {"nan", "nat", "none"}:
        return ""
    return metin


def _tarih(deger: Any) -> date | None:
    """Tarih benzeri degeri date'e cevirir; cozulemezse None.

    pandas NaT degeri ``datetime`` alt sinifidir ve ``.year`` erisiminde hata
    verir; 'kendine esit degil' testiyle elenir (NaT != NaT dogrudur).
    """
    if deger is None:
        return None
    try:
        if deger != deger:  # NaT / NaN
            return None
    except (TypeError, ValueError):
        return None
    if isinstance(deger, datetime):
        return deger.date()
    if isinstance(deger, date):
        return deger
    return None


def satir_degerleri(sonuc: Sonuc) -> list[Any]:
    """Bir ``Sonuc`` kaydini KOLONLAR sirasina gore degerlere cevirir."""
    satir = sonuc.satir
    eslesme = sonuc.eslesme
    ek = satir.ek if isinstance(satir.ek, dict) else {}
    return [
        _metin(Path(satir.kaynak_dosya).name if satir.kaynak_dosya else ""),
        satir.satir_no,
        _tarih(satir.belge_tarihi),
        _metin(satir.gider_tipi),
        _metin(satir.aciklama),
        _metin(satir.kisi_ham),
        _metin(eslesme.sicil),
        _metin(eslesme.ad_soyad),
        _metin(eslesme.yontem),
        eslesme.guven,
        eslesme.aday_sayisi,
        _metin(eslesme.aciklama),
        _tarih(sonuc.donem),
        DONEM_ESLESME_ETIKETLERI.get(
            getattr(sonuc, "donem_eslesme", "yok"),
            getattr(sonuc, "donem_eslesme", "") or "",
        ),
        _metin(sonuc.gorev_yeri),
        _metin(sonuc.masraf_merkezi),
        _metin(satir.ek.get("masraf_merkezi_adi") if isinstance(satir.ek, dict) else ""),
        _metin(sonuc.sirket or sonuc.sirket2),
        _metin(sonuc.statu),
        _metin(sonuc.kategori),
        _tarih(sonuc.cikis_tarihi),
        satir.tutar,
        _metin(satir.para_birimi),
        _metin(satir.masraf_merkezi_kaynak),
        # Denetim izi: bir satir sorgulandiginda hangi belgeye ait oldugu
        # gorulsun. Eslestirmede KULLANILMAZ, sadece raporlanir.
        _metin(ek.get("evrak_no") or ek.get("fatura_no")),
        _metin(ek.get("mail_konusu")),
        _metin(sonuc.durum),
        " | ".join(str(u) for u in (sonuc.uyarilar or [])),
    ]


class _Bicimler:
    """Duruma ve hucre tipine gore xlsxwriter bicimlerini uretir ve saklar.

    Tek kaynak: butun sayfalar ayni yazi tipi, ayni lacivert baslik, ayni
    zebra ve ayni sayi bicimini buradan alir. Sayfalar arasinda gorunum
    farki olmasin diye biçimler burada merkezilestirildi.
    """

    GOVDE = {"font_name": GOVDE_YAZI, "font_size": 10, "font_color": KOYU_GRI}

    def __init__(self, calisma: Any) -> None:
        self._calisma = calisma
        self._onbellek: dict[tuple, Any] = {}

        def f(**ozellik) -> Any:
            return calisma.add_format({**self.GOVDE, **ozellik})

        # Tablo basligi: lacivert zemin, beyaz kalin Arial.
        self.baslik = f(bold=True, font_color=BEYAZ, bg_color=LACIVERT,
                        border=1, border_color=LACIVERT, align="left",
                        valign="vcenter", text_wrap=True)
        # Bolum basligi (kapak ve aciklama bloklari): lacivert zemin.
        self.bolum = f(bold=True, font_color=BEYAZ, bg_color=LACIVERT, valign="vcenter")
        # Sayfa basligi: Georgia, buyuk, lacivert.
        self.sayfa_basligi = calisma.add_format({
            "font_name": BASLIK_YAZI, "font_size": 18, "bold": True,
            "font_color": LACIVERT, "valign": "vcenter"})
        self.alt_baslik = f(font_size=11, font_color=ORTA_GRI, valign="vcenter")
        self.kalin = f(bold=True, font_color=LACIVERT)
        self.etiket = f(font_size=9, font_color=ORTA_GRI)
        self.not_metni = f(font_size=9, font_color=ORTA_GRI, text_wrap=True, valign="top")
        # Toplam satiri: kalin lacivert, ustte ince lacivert cizgi.
        self.toplam_metin = f(bold=True, font_color=LACIVERT, top=1, border_color=LACIVERT)
        self.toplam_sayi = f(bold=True, font_color=LACIVERT, top=1, border_color=LACIVERT,
                             num_format=TUTAR_BICIMI)
        self.toplam_tamsayi = f(bold=True, font_color=LACIVERT, top=1, border_color=LACIVERT,
                                num_format=TAMSAYI_BICIMI)
        # Kapak sayfasi
        self.ozet_metin = f(align="left", valign="vcenter")
        self.ozet_sayi = f(num_format=TAMSAYI_BICIMI)
        self.ozet_tutar = f(num_format=TUTAR_BICIMI)
        self.ozet_yuzde = f(num_format=YUZDE_BICIMI)
        self.kpi_sayi = calisma.add_format({
            "font_name": BASLIK_YAZI, "font_size": 20, "bold": True,
            "font_color": LACIVERT, "num_format": TUTAR_BICIMI, "valign": "vcenter"})
        self.kpi_tamsayi = calisma.add_format({
            "font_name": BASLIK_YAZI, "font_size": 20, "bold": True,
            "font_color": LACIVERT, "num_format": TAMSAYI_BICIMI, "valign": "vcenter"})
        self.kpi_metin = calisma.add_format({
            "font_name": BASLIK_YAZI, "font_size": 20, "bold": True,
            "font_color": LACIVERT, "valign": "vcenter"})
        self.kpi_etiket = f(font_size=9, font_color=ORTA_GRI, valign="top")
        self.durum_iyi = f(bold=True, font_color=BEYAZ, bg_color="#2E7D32", align="center",
                           valign="vcenter")
        self.durum_kotu = f(bold=True, font_color=BEYAZ, bg_color="#C62828", align="center",
                            valign="vcenter")
        self.vurgu = f(bold=True, font_color=BEYAZ, bg_color=CANLI_MAVI, valign="vcenter",
                       text_wrap=True)
        self.kapak_bant = calisma.add_format({"bg_color": LACIVERT})
        self.kapak_baslik = calisma.add_format({
            "font_name": BASLIK_YAZI, "font_size": 22, "bold": True,
            "font_color": BEYAZ, "bg_color": LACIVERT, "valign": "vcenter", "indent": 1})
        self.kapak_alt = calisma.add_format({
            "font_name": GOVDE_YAZI, "font_size": 11, "font_color": "#AAE6F0",
            "bg_color": LACIVERT, "valign": "vcenter", "indent": 1})

    def _hucre(self, onek: str, tip: str, arka: str | None, zebra: bool) -> Any:
        anahtar = (onek, tip, arka, zebra)
        if anahtar in self._onbellek:
            return self._onbellek[anahtar]
        ozellikler: dict[str, Any] = {**self.GOVDE, "bottom": 1, "bottom_color": KENAR_GRI,
                                      "valign": "vcenter"}
        if arka:
            ozellikler["bg_color"] = arka
        elif zebra:
            ozellikler["bg_color"] = ACIK_GRI
        if tip == "tarih":
            ozellikler["num_format"] = TARIH_BICIMI
            ozellikler["align"] = "center"
        elif tip == "sayi":
            ozellikler["num_format"] = TUTAR_BICIMI
        elif tip == "tamsayi":
            ozellikler["num_format"] = TAMSAYI_BICIMI
            ozellikler["align"] = "center"
        elif tip == "yuzde":
            ozellikler["num_format"] = YUZDE_BICIMI
        bicim = self._calisma.add_format(ozellikler)
        self._onbellek[anahtar] = bicim
        return bicim

    def al(self, tip: str, durum: str, zebra: bool = False) -> Any:
        """Satir dokumu sayfalari: durum rengi varsa o, yoksa zebra."""
        return self._hucre("satir", tip, DURUM_RENKLERI.get(durum), zebra)

    def mahsup(self, tip: str, renk_anahtari: str, zebra: bool = False) -> Any:
        """Mahsuplasma/Kontrol/Sirket sayfalari: kalite rengi varsa o, yoksa zebra."""
        return self._hucre("mahsup", tip, MAHSUP_RENKLERI.get(renk_anahtari), zebra)


def _sayfa_hazirla(sayfa: Any, kolonlar: Sequence[tuple[str, str, int]], bicimler: _Bicimler) -> None:
    """Her tablo sayfasinin ortak gorunumu: gizli kilavuz, donmus baslik, genislikler."""
    sayfa.hide_gridlines(2)
    sayfa.freeze_panes(1, 0)
    sayfa.set_row(0, 30)
    sayfa.set_default_row(18)
    for sutun, (baslik, _tip, genislik) in enumerate(kolonlar):
        sayfa.write_string(0, sutun, baslik, bicimler.baslik)
        sayfa.set_column(sutun, sutun, genislik)
    sayfa.set_landscape()
    sayfa.fit_to_pages(1, 0)
    sayfa.repeat_rows(0)


def _hucre_yaz(sayfa: Any, satir: int, sutun: int, tip: str, deger: Any, bicim: Any) -> None:
    if deger is None or deger == "":
        sayfa.write_blank(satir, sutun, None, bicim)
    elif tip == "tarih":
        sayfa.write_datetime(satir, sutun, datetime(deger.year, deger.month, deger.day), bicim)
    elif tip in ("sayi", "tamsayi", "yuzde"):
        try:
            sayfa.write_number(satir, sutun, float(deger), bicim)
        except (TypeError, ValueError):
            sayfa.write_string(satir, sutun, str(deger), bicim)
    else:
        sayfa.write_string(satir, sutun, str(deger), bicim)


def _sayfa_yaz(calisma: Any, ad: str, sonuclar: Sequence[Sonuc], bicimler: _Bicimler) -> None:
    """Bir satir dokumu sayfasini basliklari, filtresi ve bicimleriyle yazar."""
    sayfa = calisma.add_worksheet(ad)
    _sayfa_hazirla(sayfa, KOLONLAR, bicimler)

    for indeks, sonuc in enumerate(sonuclar, start=1):
        degerler = satir_degerleri(sonuc)
        zebra = indeks % 2 == 0
        for sutun, ((_baslik, tip, _g), deger) in enumerate(zip(KOLONLAR, degerler)):
            _hucre_yaz(sayfa, indeks, sutun, tip, deger, bicimler.al(tip, sonuc.durum, zebra))

    son_satir = max(1, len(sonuclar))
    sayfa.autofilter(0, 0, son_satir, len(KOLONLAR) - 1)
    if not sonuclar:
        sayfa.write_string(1, 0, "(Bu sayfada satir yok)", bicimler.ozet_metin)


def _ozet_yaz(
    calisma: Any,
    ozet: dict,
    sonuclar: Sequence[Sonuc],
    bicimler: _Bicimler,
    mahsup: Any = None,
    oneriler: Any = None,
) -> None:
    """'Ozet' kapak sayfasi: bir bakista butun calisma.

    Finans muduru dosyayi actiginda ilk bu sayfayi gorur. Sirasi okuma
    sirasidir: para nereden geldi nereye gitti, mutabakat kapandi mi, hangi
    sirkete ne dustu, nelere dikkat edilmeli, hangi sayfada ne var.
    """
    sayfa = calisma.add_worksheet("Ozet")
    sayfa.hide_gridlines(2)
    sayfa.set_column(0, 0, 3)          # sol bosluk
    sayfa.set_column(1, 1, 34)
    sayfa.set_column(2, 7, 17)
    sayfa.set_column(8, 8, 3)
    sayfa.set_landscape()
    sayfa.fit_to_pages(1, 0)

    # --- Ust bant --------------------------------------------------------
    for r in range(0, 4):
        sayfa.set_row(r, 22 if r else 8)
        for c in range(0, 9):
            sayfa.write_blank(r, c, None, bicimler.kapak_bant)
    sayfa.set_row(1, 34)
    sayfa.merge_range(1, 1, 1, 7, "Masraf Merkezi Dagitimi", bicimler.kapak_baslik)
    donem = _gider_donemi(sonuclar)
    dosya_sayisi = ozet.get("dosya_sayisi", 0)
    alt = (f"{donem}  |  {dosya_sayisi} dosya, {len(sonuclar)} satir  |  "
           f"uretim {datetime.now():%d.%m.%Y %H:%M}  |  Rencons Heavy Industries")
    sayfa.merge_range(2, 1, 2, 7, alt, bicimler.kapak_alt)

    satir = 5

    # --- Para birimi bazinda KPI satiri ---------------------------------
    toplamlar = mahsup.toplamlar() if mahsup is not None else {}
    if not toplamlar:
        sayfa.write_string(satir, 1, "Dagitilacak tutarli satir bulunamadi.", bicimler.ozet_metin)
        satir += 2
    for para, d in toplamlar.items():
        kpi = (
            ("Okunan", d["gelen"], bicimler.kpi_sayi, "dosyalarda gorulen toplam"),
            ("Yinelenen", d["yinelenen"], bicimler.kpi_sayi, "baska dosyada zaten sayildi"),
            ("Net", d["net"], bicimler.kpi_sayi, "gercekten dagitilacak"),
            ("Dagitilan", d["dagitilan"], bicimler.kpi_sayi, "projelere yazildi"),
            ("Dagitilamayan", d["dagitilamayan"], bicimler.kpi_sayi, "kisi/merkez bulunamadi"),
            ("Dagitim orani", d["oran"] / 100.0, None, "dagitilan / net"),
        )
        sayfa.set_row(satir, 14)
        sayfa.write_string(satir, 1, f"TUTARLAR ({para})", bicimler.kalin)
        satir += 1
        sayfa.set_row(satir, 30)
        for i, (ad, deger, bicim, aciklama) in enumerate(kpi):
            c = 1 + i
            if bicim is None:
                yuzde = calisma.add_format({
                    "font_name": BASLIK_YAZI, "font_size": 20, "bold": True,
                    "font_color": CANLI_MAVI, "num_format": YUZDE_BICIMI, "valign": "vcenter"})
                sayfa.write_number(satir, c, float(deger), yuzde)
            else:
                sayfa.write_number(satir, c, float(deger), bicim)
        satir += 1
        sayfa.set_row(satir, 24)
        for i, (ad, _d, _b, aciklama) in enumerate(kpi):
            sayfa.write_string(satir, 1 + i, f"{ad}\n{aciklama}", bicimler.kpi_etiket)
        satir += 2

    # --- Mutabakat durumu ------------------------------------------------
    if mahsup is not None:
        sayfa.set_row(satir, 26)
        if mahsup.kapali_mi:
            sayfa.merge_range(satir, 1, satir, 7,
                              "MUTABAKAT KAPALI  -  Okunan = Yinelenen + Dagitilan + Dagitilamayan. "
                              "Para kaybolmadi; tablo muhasebeye gonderilebilir.",
                              bicimler.durum_iyi)
        else:
            acik = ", ".join(f"{k.kaynak} ({k.fark:+.2f})" for k in mahsup.acik_kontroller)
            sayfa.merge_range(satir, 1, satir, 7,
                              f"MUTABAKAT ACIK  -  {acik}. Tablo muhasebeye GONDERILMEMELI.",
                              bicimler.durum_kotu)
        satir += 2

    # --- Satir durumu ----------------------------------------------------
    def tablo_basligi(basliklar: Sequence[str], ilk_genis: bool = True) -> None:
        nonlocal satir
        sayfa.set_row(satir, 20)
        for i, b in enumerate(basliklar):
            sayfa.write_string(satir, 1 + i, b, bicimler.baslik)
        satir += 1

    def bolum(baslik: str) -> None:
        nonlocal satir
        satir += 1
        sayfa.set_row(satir, 22)
        sayfa.merge_range(satir, 1, satir, 7, baslik, bicimler.bolum)
        satir += 1

    bolum("Satir durumu")
    tablo_basligi(("Durum", "Satir", "Oran", "Ne demek"))
    aciklama = {
        DURUM_OTOMATIK: "Kimlik kesin, uyari yok; oldugu gibi kaydedilebilir",
        DURUM_INCELE: "Sistem sonuc buldu ama emin degil; gerekcesi satirda yazili",
        DURUM_ESLESMEDI: "Kisi bulunamadi; tutar (DAGITILAMAYAN) satirinda duruyor",
    }
    dagilim = ozet.get("durum_dagilimi", {})
    toplam_satir = max(1, sum(dagilim.get(d, 0) for d in aciklama))
    for i, durum in enumerate(aciklama):
        adet = dagilim.get(durum, 0)
        zebra = i % 2 == 1
        sayfa.write_string(satir, 1, durum, bicimler.al("metin", durum))
        sayfa.write_number(satir, 2, float(adet), bicimler.al("tamsayi", durum))
        sayfa.write_number(satir, 3, adet / toplam_satir, bicimler.al("yuzde", durum))
        sayfa.merge_range(satir, 4, satir, 7, aciklama[durum], bicimler.al("metin", durum))
        satir += 1

    # --- Sirket kirilimi -------------------------------------------------
    if mahsup is not None and hasattr(mahsup, "sirket_ozeti"):
        gruplar = mahsup.sirket_ozeti()
        if gruplar:
            bolum("Sirket kirilimi  (tuzel kisi ustte, en buyuk projesi yaninda)")
            tablo_basligi(("Sirket", "Tutar", "Para", "Pay", "Kisi", "En buyuk proje", "Proje payi"))
            for i, g in enumerate(gruplar):
                zebra = i % 2 == 1
                en = g["projeler"][0] if g["projeler"] else None
                sayfa.write_string(satir, 1, g["sirket"], bicimler.mahsup("metin", "", zebra))
                sayfa.write_number(satir, 2, g["tutar"], bicimler.mahsup("sayi", "", zebra))
                sayfa.write_string(satir, 3, g["para_birimi"], bicimler.mahsup("metin", "", zebra))
                sayfa.write_number(satir, 4, g["pay_yuzde"] / 100.0, bicimler.mahsup("yuzde", "", zebra))
                sayfa.write_number(satir, 5, g["kisi_sayisi"], bicimler.mahsup("tamsayi", "", zebra))
                sayfa.write_string(satir, 6, (en["masraf_merkezi"] if en else ""), bicimler.mahsup("metin", "", zebra))
                sayfa.write_number(satir, 7, (en["pay_yuzde"] / 100.0 if en else 0.0), bicimler.mahsup("yuzde", "", zebra))
                satir += 1

        # --- En buyuk masraf merkezleri ---------------------------------
        merkezler = mahsup.merkez_ozeti()[:10]
        if merkezler:
            bolum("En buyuk masraf merkezleri  (ilk 10)")
            tablo_basligi(("Masraf merkezi", "Tutar", "Para", "Pay", "Kisi", "Sirket", "Durum"))
            for i, m in enumerate(merkezler):
                zebra = i % 2 == 1
                renk = "" if m["haritada_var"] else "uyari"
                sayfa.write_string(satir, 1, str(m["masraf_merkezi"]), bicimler.mahsup("metin", renk, zebra))
                sayfa.write_number(satir, 2, m["tutar"], bicimler.mahsup("sayi", renk, zebra))
                sayfa.write_string(satir, 3, m["para_birimi"], bicimler.mahsup("metin", renk, zebra))
                sayfa.write_number(satir, 4, m["pay_yuzde"] / 100.0, bicimler.mahsup("yuzde", renk, zebra))
                sayfa.write_number(satir, 5, m["kisi_sayisi"], bicimler.mahsup("tamsayi", renk, zebra))
                sayfa.write_string(satir, 6, str(m.get("sirket") or ""), bicimler.mahsup("metin", renk, zebra))
                sayfa.write_string(satir, 7, "" if m["haritada_var"] else "haritada tanimli degil",
                                   bicimler.mahsup("metin", renk, zebra))
                satir += 1

    # --- Dikkat ----------------------------------------------------------
    dikkat: list[str] = []
    if mahsup is not None:
        for c in getattr(mahsup, "isaret_celiskileri", []) or []:
            dikkat.append("Isaret celiskisi: " + c.aciklama())
        if getattr(mahsup, "tutarsiz_satir_sayisi", 0):
            dikkat.append(f"{mahsup.tutarsiz_satir_sayisi} satirda tutar okunamadi; dagilima girmedi.")
    if oneriler:
        dikkat.append(f"{len(oneriler)} gorev yeri masraf merkezi haritasinda tanimli degil. "
                      "'Harita Onerileri' sayfasinda hazir satirlar var.")
    for u in (ozet.get("uyarilar") or []):
        metin = str(u)
        if "Ogrenildi" in metin or "kisi kutugu" in metin:
            continue   # bilgi mesaji, dikkat gerektirmiyor
        if metin not in dikkat:
            dikkat.append(metin)
    for h in (ozet.get("hatalar") or []):
        dikkat.append("HATA: " + str(h))
    if dikkat:
        bolum("Dikkat")
        for i, metin in enumerate(dikkat[:12]):
            sayfa.set_row(satir, 30)
            sayfa.merge_range(satir, 1, satir, 7, metin, bicimler.not_metni)
            satir += 1
        if len(dikkat) > 12:
            sayfa.merge_range(satir, 1, satir, 7, f"... ve {len(dikkat) - 12} not daha (Kontrol ve Sonuc sayfalarinda).",
                              bicimler.not_metni)
            satir += 1

    # --- Islenen dosyalar ------------------------------------------------
    dosyalar = ozet.get("dosyalar") or []
    if dosyalar:
        bolum("Islenen dosyalar")
        for i, yol in enumerate(dosyalar):
            zebra = i % 2 == 1
            sayfa.merge_range(satir, 1, satir, 7, Path(str(yol)).name, bicimler.mahsup("metin", "", zebra))
            satir += 1
    bolum("Kaynak veri")
    for ad, deger in (
        ("Personel ana verisi", Path(str(ozet.get("personel_dosyasi") or "")).name or "-"),
        ("Personel son donemi", ozet["son_donem"].strftime("%d.%m.%Y") if isinstance(ozet.get("son_donem"), date) else "-"),
        ("Otomatik kabul esigi", f"guven >= {ozet.get('guven_esigi', '')} ve uyari yok"),
    ):
        sayfa.write_string(satir, 1, ad, bicimler.ozet_metin)
        sayfa.merge_range(satir, 2, satir, 7, str(deger), bicimler.ozet_metin)
        satir += 1

    # --- Sayfa rehberi ---------------------------------------------------
    bolum("Bu dosyada ne var")
    rehber = (
        ("Mahsuplasma", "MUHASEBEYE GIDEN TABLO. Her fatura icin sirket ve projeye ne kadar yazilacagi."),
        ("Sirket Kirilimi", "Tuzel kisi ustte, projeleri altinda."),
        ("Kontrol", "Mutabakat. Fark sutunu sifir olmak zorunda."),
        ("Harita Onerileri", "Tanimsiz gorev yerleri ve haritaya yapistirmaya hazir satirlar."),
        ("Sonuc", "Butun satirlar: kisinin nasil bulundugu, guven, gerekce, evrak no, mail konusu."),
        ("Incele", "Elle bakilacak satirlar."),
        ("Eslesmedi", "Kisi bulunamayan satirlar."),
    )
    for i, (ad, aciklama) in enumerate(rehber):
        zebra = i % 2 == 1
        sayfa.write_url(satir, 1, f"internal:'{ad}'!A1", bicimler.mahsup("metin", "", zebra), ad)
        sayfa.merge_range(satir, 2, satir, 7, aciklama, bicimler.mahsup("metin", "", zebra))
        satir += 1

    satir += 1
    sayfa.merge_range(satir, 1, satir, 7,
                      "Otomasyon cevrimdisi calisir; personel ve fatura verisi bilgisayardan disari cikmaz. "
                      "Bu dosya kisisel veri icerir, paylasirken dikkat edin.",
                      bicimler.not_metni)


def _gider_donemi(sonuclar: Sequence[Sonuc]) -> str:
    """Satirlardaki belge tarihlerinden okunakli donem etiketi ('Temmuz 2026')."""
    aylar = ("Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran", "Temmuz",
             "Agustos", "Eylul", "Ekim", "Kasim", "Aralik")
    tarihler = [s.satir.belge_tarihi for s in sonuclar
                if getattr(s.satir, "belge_tarihi", None) and getattr(s.satir, "tutar", None) is not None]
    if not tarihler:
        return "Donem belirsiz"
    from collections import Counter
    (yil, ay), _ = Counter((t.year, t.month) for t in tarihler).most_common(1)[0]
    return f"{aylar[ay - 1]} {yil}"


def mahsup_durumu(satir: Any) -> tuple[str, str]:
    """Bir mahsup satirinin kalite durumu: (renk anahtari, okunakli etiket).

    Muhasebeci bu sutunu okuyup satiri oldugu gibi kaydedip kaydedemeyecegini
    anlar. Bos etiket 'kontrol gerekmez' demektir.
    """
    from masraf.mahsuplasma import DAGITILAMAYAN

    sorunlar: list[str] = []
    if satir.masraf_merkezi == DAGITILAMAYAN:
        sorunlar.append("MASRAF MERKEZI YOK")
    elif not satir.haritada_var:
        sorunlar.append("HARITADA TANIMLI DEGIL")
    if satir.eslesmedi:
        sorunlar.append(f"{satir.eslesmedi} eslesmeyen satir")
    if satir.incele:
        sorunlar.append(f"{satir.incele} satir incelenecek")

    if not sorunlar:
        return "tamam", ""
    engel = (satir.masraf_merkezi == DAGITILAMAYAN) or bool(satir.eslesmedi)
    return ("engel" if engel else "uyari"), " | ".join(sorunlar)


def mahsup_satir_degerleri(satir: Any, fatura_toplami: float) -> list[Any]:
    """Bir ``MahsupSatiri``ni MAHSUP_KOLONLARI sirasina cevirir."""
    _renk, etiket = mahsup_durumu(satir)
    pay = (satir.tutar / fatura_toplami) if fatura_toplami else 0.0   # kesir; Excel 0.0% gosterir
    return [
        satir.kaynak,
        satir.sirket or "(sirket yok)",
        satir.masraf_merkezi,
        satir.masraf_merkezi_adi or "",
        satir.gider_tipi,
        satir.pay_notu or "",
        round(satir.tutar, 2),
        satir.para_birimi,
        round(pay, 4),
        satir.satir_sayisi,
        satir.kisi_sayisi,
        satir.otomatik,
        satir.incele,
        satir.eslesmedi,
        satir.gider_donemi,
        etiket,
    ]


def kontrol_satir_degerleri(kontrol: Any) -> list[Any]:
    """Bir ``KontrolSatiri``ni KONTROL_KOLONLARI sirasina cevirir."""
    return [
        kontrol.kaynak,
        kontrol.para_birimi,
        kontrol.gelen,
        kontrol.yinelenen_tutar,
        kontrol.net,
        kontrol.dagitilan,
        kontrol.dagitilamayan,
        kontrol.fark,
        kontrol.satir_sayisi,
        kontrol.yinelenen_satir,
        round(kontrol.dagitim_orani / 100.0, 4),
        "KAPANDI" if kontrol.kapali_mi else "ACIK - KONTROL EDIN",
    ]


def _tablo_yaz(
    calisma: Any,
    ad: str,
    kolonlar: Sequence[tuple[str, str, int]],
    satirlar: Sequence[Sequence[Any]],
    renkler: Sequence[str],
    bicimler: "_Bicimler",
    bos_mesaj: str = "(Bu sayfada satir yok)",
    toplam_sutunlari: Sequence[int] = (),
) -> Any:
    """Basliklari, filtresi, zebra/durum renkleri ve toplam satiri olan tablo yazar."""
    sayfa = calisma.add_worksheet(ad)
    _sayfa_hazirla(sayfa, kolonlar, bicimler)

    for indeks, (degerler, renk) in enumerate(zip(satirlar, renkler), start=1):
        zebra = indeks % 2 == 0
        for sutun, ((_baslik, tip, _g), deger) in enumerate(zip(kolonlar, degerler)):
            _hucre_yaz(sayfa, indeks, sutun, tip, deger, bicimler.mahsup(tip, renk, zebra))

    son = len(satirlar)
    sayfa.autofilter(0, 0, max(1, son), len(kolonlar) - 1)
    if not satirlar:
        sayfa.write_string(1, 0, bos_mesaj, bicimler.ozet_metin)
        return sayfa

    if toplam_sutunlari:
        satir_no = son + 1
        sayfa.write_string(satir_no, 0, "TOPLAM", bicimler.toplam_metin)
        for sutun in range(1, len(kolonlar)):
            tip = kolonlar[sutun][1]
            if sutun in toplam_sutunlari:
                harf = _sutun_harfi(sutun)
                bicim = bicimler.toplam_tamsayi if tip == "tamsayi" else bicimler.toplam_sayi
                sayfa.write_formula(satir_no, sutun, f"=SUM({harf}2:{harf}{son + 1})", bicim)
            else:
                sayfa.write_blank(satir_no, sutun, None, bicimler.toplam_metin)
    return sayfa


def _sutun_harfi(indeks: int) -> str:
    """0 tabanli sutun indeksini Excel harfine cevirir (0 -> A)."""
    harfler = ""
    indeks += 1
    while indeks:
        indeks, kalan = divmod(indeks - 1, 26)
        harfler = chr(65 + kalan) + harfler
    return harfler


SIRKET_KOLONLARI: tuple[tuple[str, str, int], ...] = (
    ("Sirket / Proje", "metin", 42),
    ("Seviye", "metin", 10),
    ("Masraf Merkezi Kodu", "metin", 22),
    ("Tutar", "sayi", 15),
    ("Para Birimi", "metin", 10),
    ("Pay", "yuzde", 9),
    ("Satir", "tamsayi", 8),
    ("Kisi", "tamsayi", 8),
    ("Durum", "metin", 26),
)


def _sirket_ozeti_yaz(calisma: Any, tablo: Any, bicimler: "_Bicimler") -> None:
    """'Sirket Kirilimi' sayfasi: tuzel kisi ustte, projeleri altinda.

    Muhasebenin okuma sirasi budur. Sirket satirlari koyu, proje satirlari
    girintili yazilir; boylece tek bakista hangi projenin hangi sirkete ait
    oldugu gorulur.
    """
    if not hasattr(tablo, "sirket_ozeti"):
        return
    degerler: list[list[Any]] = []
    renkler: list[str] = []
    for grup in tablo.sirket_ozeti():
        degerler.append([
            grup["sirket"], "SIRKET", "",
            grup["tutar"], grup["para_birimi"], grup["pay_yuzde"] / 100.0,
            grup["satir_sayisi"], grup["kisi_sayisi"], "",
        ])
        renkler.append("baslik")
        for proje in grup["projeler"]:
            uyari = "" if proje["haritada_var"] else "HARITADA TANIMLI DEGIL"
            degerler.append([
                "    " + str(proje["masraf_merkezi_adi"] or proje["masraf_merkezi"]),
                "proje", proje["masraf_merkezi"],
                proje["tutar"], grup["para_birimi"], proje["pay_yuzde"] / 100.0,
                proje["satir_sayisi"], proje["kisi_sayisi"], uyari,
            ])
            renkler.append("tamam" if proje["haritada_var"] else "uyari")
    _tablo_yaz(
        calisma, "Sirket Kirilimi", SIRKET_KOLONLARI, degerler, renkler, bicimler,
        bos_mesaj="(Dagitilacak satir yok)",
    )


def _harita_onerisi_yaz(calisma: Any, oneriler: Any, bicimler: "_Bicimler") -> None:
    """'Harita Onerileri' sayfasi: haritaya eklenecek satirlar, hazir halde."""
    from masraf.harita_onerisi import ONERI_BASLIKLARI, oneri_satir_degerleri

    degerler = [oneri_satir_degerleri(o) for o in (oneriler or [])]
    sayfa = _tablo_yaz(
        calisma, "Harita Onerileri", ONERI_BASLIKLARI, degerler,
        ["uyari"] * len(degerler), bicimler,
        bos_mesaj="(Butun gorev yerleri haritada tanimli. Yapilacak bir sey yok.)",
    )
    satir = len(degerler) + 3
    sayfa.write_string(satir, 0, "Bu sayfa ne icin", bicimler.bolum)
    satir += 1
    for metin in (
        "Asagidaki gorev yerleri masraf merkezi haritasinda tanimli degil. Kod "
        "onlari metin olarak tasidi ve isaretledi; finans koduna cevrilmeden "
        "muhasebeye gitmemeliler.",
        "SIRKET sutunu 1C personel listesindeki 'Firm 2' kolonundan gelir. "
        "Haritadaki mevcut satirlarla ayni sozlugu kullanir (RHI, UST LUGA, "
        "RSS, RC, BSK), bu yuzden guvenilir.",
        "ONERILEN KOD bir baslangic degeridir, finansin onayina tabidir. "
        "Kendi kodunuzu kullanin.",
        "Eklemek icin: veri\\masraf_merkezi_haritasi.csv dosyasini acin ve her "
        "satir icin su bicimde bir satir ekleyin:",
        "    gorev_yeri,masraf_merkezi_kodu,masraf_merkezi_adi,sirket,aktif",
        "Bir kez eklemek yeterlidir; sonraki aylarda otomatik cozulur.",
    ):
        sayfa.write_string(satir, 0, metin, bicimler.ozet_metin)
        satir += 1
    if degerler:
        satir += 1
        sayfa.write_string(satir, 0, "Yapistirmaya hazir satirlar", bicimler.bolum)
        satir += 1
        for o in oneriler:
            sayfa.write_string(satir, 0, o.csv_satiri(), bicimler.ozet_metin)
            satir += 1


def _mahsuplasma_yaz(calisma: Any, tablo: Any, bicimler: "_Bicimler") -> None:
    """'Mahsuplasma' ve 'Kontrol' sayfalarini yazar."""
    fatura_toplami: dict[tuple[str, str], float] = {}
    for m in tablo.satirlar:
        anahtar = (m.kaynak, m.para_birimi)
        fatura_toplami[anahtar] = fatura_toplami.get(anahtar, 0.0) + m.tutar

    degerler, renkler = [], []
    for m in tablo.satirlar:
        renk, _etiket = mahsup_durumu(m)
        degerler.append(mahsup_satir_degerleri(m, fatura_toplami[(m.kaynak, m.para_birimi)]))
        renkler.append(renk)
    _tablo_yaz(
        calisma, "Mahsuplasma", MAHSUP_KOLONLARI, degerler, renkler, bicimler,
        bos_mesaj="(Dagitilacak tutarli satir bulunamadi)",
        toplam_sutunlari=(6, 9, 10, 11, 12, 13),
    )

    k_degerler = [kontrol_satir_degerleri(k) for k in tablo.kontrol]
    k_renkler = ["tamam" if k.kapali_mi else "engel" for k in tablo.kontrol]
    sayfa = _tablo_yaz(
        calisma, "Kontrol", KONTROL_KOLONLARI, k_degerler, k_renkler, bicimler,
        bos_mesaj="(Kontrol edilecek fatura yok)",
        toplam_sutunlari=(2, 3, 4, 5, 6, 8, 9),   # Fark (7) ve oran (10) toplanmaz
    )

    # Kontrol sayfasinin altina aciklamalar ve isaret celiskileri.
    satir = len(k_degerler) + 3
    sayfa.write_string(satir, 0, "Nasil okunur", bicimler.bolum)
    satir += 1
    for metin in (
        "Okunan Tutar = dosyada gorulen her seyin toplami.",
        "Yinelenen = ayni islem baska bir dosyada zaten sayildigi icin bu "
        "dosyadan dusuldu. Ayni mailde hem acentenin ham dokumu hem de o "
        "islemlerin elle dagitilmis hali gelirse para cift sayilir; bu sutun "
        "onu engeller.",
        "Net Tutar = Okunan - Yinelenen. Bu dosyanin gercekten kattigi tutar.",
        "Fark = Okunan - (Yinelenen + Dagitilan + Dagitilamayan). SIFIR olmali. "
        "Sifir degilse dagitimda kayip var demektir, muhasebeye gonderilmemeli.",
        "Dagitilamayan = kisi veya masraf merkezi bulunamadigi icin projeye "
        "yazilamayan tutar. Silinmez; Mahsuplasma sayfasinda "
        "'(DAGITILAMAYAN)' satiri olarak durur.",
    ):
        sayfa.write_string(satir, 0, metin, bicimler.ozet_metin)
        satir += 1

    if getattr(tablo, "isaret_celiskileri", None):
        satir += 1
        sayfa.write_string(satir, 0, "Isaret celiskileri", bicimler.bolum)
        satir += 1
        for celiski in tablo.isaret_celiskileri:
            sayfa.write_string(satir, 0, celiski.aciklama(), bicimler.ozet_metin)
            satir += 1

    if tablo.kutuk_satir_sayisi or tablo.tutarsiz_satir_sayisi:
        satir += 1
        sayfa.write_string(satir, 0, "Dagilima girmeyen satirlar", bicimler.bolum)
        satir += 1
        if tablo.kutuk_satir_sayisi:
            sayfa.write_string(
                satir, 0,
                f"{tablo.kutuk_satir_sayisi} satir kisi kutugunden geldi "
                "(katilimci listesi, saglik kontrol listesi). Bunlar fatura "
                "degildir, tutar tasimazlar.",
                bicimler.ozet_metin)
            satir += 1
        if tablo.tutarsiz_satir_sayisi:
            sayfa.write_string(
                satir, 0,
                f"{tablo.tutarsiz_satir_sayisi} satirda tutar okunamadi; "
                "dagilima girmediler. Kaynak dosyada tutar kolonu bos olabilir "
                "ya da kolon adi taninmamis olabilir.",
                bicimler.ozet_metin)
            satir += 1


def excel_yaz(
    sonuclar: list[Sonuc],
    yol: str,
    ozet: dict,
    mahsup: Any = None,
    harita_onerileri: Any = None,
) -> str:
    """Sonuclari bicimlendirilmis cok sayfali Excel dosyasina yazar.

    Args:
        sonuclar: Boru hattinin urettigi sonuc listesi.
        yol: Yazilacak .xlsx dosyasinin yolu.
        ozet: ``Boru.ozet()`` ciktisi.
        mahsup: ``mahsuplasma_uret`` ciktisi olan ``MahsupTablosu``. Verilirse
            'Mahsuplasma' ve 'Kontrol' sayfalari EN BASA eklenir; muhasebeye
            giden tablo odur, satir dokumu onun dayanagidir.

    Returns:
        Yazilan dosyanin tam yolu (str).
    """
    import xlsxwriter

    hedef = Path(yol)
    hedef.parent.mkdir(parents=True, exist_ok=True)

    calisma = xlsxwriter.Workbook(
        str(hedef), {"default_date_format": TARIH_BICIMI, "constant_memory": False}
    )
    try:
        bicimler = _Bicimler(calisma)
        # Kapak once: dosya acildiginda ilk gorulen sayfa. Sonra muhasebeye
        # giden tablo, sonra kaniti.
        _ozet_yaz(calisma, ozet or {}, list(sonuclar), bicimler, mahsup, harita_onerileri)
        if mahsup is not None:
            _mahsuplasma_yaz(calisma, mahsup, bicimler)
            _sirket_ozeti_yaz(calisma, mahsup, bicimler)
        if harita_onerileri is not None:
            _harita_onerisi_yaz(calisma, harita_onerileri, bicimler)
        _sayfa_yaz(calisma, "Sonuc", list(sonuclar), bicimler)
        _sayfa_yaz(
            calisma, "Incele", [s for s in sonuclar if s.durum == DURUM_INCELE], bicimler
        )
        _sayfa_yaz(
            calisma, "Eslesmedi", [s for s in sonuclar if s.durum == DURUM_ESLESMEDI], bicimler
        )
    finally:
        calisma.close()

    return str(hedef)


def csv_yaz(sonuclar: Iterable[Sonuc], yol: str) -> str:
    """Sonuclari UTF-8 BOM'lu CSV olarak yazar (Excel dogrudan acar).

    Ayirici noktali virguldur; Rusca/Turkce yerel ayarli Excel kurulumlari
    bu dosyayi kolon kaymasi olmadan acar.
    """
    hedef = Path(yol)
    hedef.parent.mkdir(parents=True, exist_ok=True)

    with hedef.open("w", encoding="utf-8-sig", newline="") as akis:
        yazici = csv.writer(akis, delimiter=";")
        yazici.writerow([baslik for baslik, _tip, _g in KOLONLAR])
        for sonuc in sonuclar:
            hucreler: list[str] = []
            for (_baslik, tip, _g), deger in zip(KOLONLAR, satir_degerleri(sonuc)):
                if deger is None or deger == "":
                    hucreler.append("")
                elif tip == "tarih":
                    hucreler.append(deger.strftime("%d.%m.%Y"))
                elif tip == "yuzde":
                    hucreler.append(f"{float(deger) * 100:.0f}%")
                elif tip == "sayi":
                    hucreler.append(f"{float(deger):.2f}")
                else:
                    hucreler.append(str(deger))
            yazici.writerow(hucreler)

    return str(hedef)


def mahsuplasma_csv_yaz(mahsup: Any, yol: str) -> str:
    """Mahsuplasma tablosunu UTF-8 BOM'lu CSV olarak yazar.

    Excel'i olmayan ya da tabloyu baska bir sisteme aktaracak kullanicilar
    icin. Ayirici noktali virguldur.
    """
    fatura_toplami: dict[tuple[str, str], float] = {}
    for m in mahsup.satirlar:
        anahtar = (m.kaynak, m.para_birimi)
        fatura_toplami[anahtar] = fatura_toplami.get(anahtar, 0.0) + m.tutar

    hedef = Path(yol)
    hedef.parent.mkdir(parents=True, exist_ok=True)
    with hedef.open("w", encoding="utf-8-sig", newline="") as akis:
        yazici = csv.writer(akis, delimiter=";")
        yazici.writerow([baslik for baslik, _tip, _g in MAHSUP_KOLONLARI])
        for m in mahsup.satirlar:
            degerler = mahsup_satir_degerleri(
                m, fatura_toplami[(m.kaynak, m.para_birimi)]
            )
            hucreler: list[str] = []
            for (_baslik, tip, _g), deger in zip(MAHSUP_KOLONLARI, degerler):
                if deger is None or deger == "":
                    hucreler.append("")
                elif tip == "tarih":
                    hucreler.append(deger.strftime("%d.%m.%Y"))
                elif tip == "yuzde":
                    hucreler.append(f"{float(deger) * 100:.1f}%")
                elif tip == "sayi":
                    hucreler.append(f"{float(deger):.2f}")
                else:
                    hucreler.append(str(deger))
            yazici.writerow(hucreler)
    return str(hedef)
