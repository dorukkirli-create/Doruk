"""Excel ve CSV cikti uretimi.

Uretilen Excel dosyasi finans ekibinin dogrudan calisacagi belgedir. Sayfalar
IS AKISI SIRASINDADIR: once muhasebeye gidecek olan, sonra kaniti.

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
    ("Durum", "metin", 12),
    ("Uyarilar", "metin", 70),
)

#: Mahsuplasma sayfasi kolonlari: (baslik, tip, genislik).
#: Sira muhasebecinin okuma sirasidir: once hangi fatura, sonra hangi proje,
#: sonra ne kadar, en sonda kalite isaretleri.
MAHSUP_KOLONLARI: tuple[tuple[str, str, int], ...] = (
    ("Fatura / Kaynak Dosya", "metin", 34),
    ("Masraf Merkezi Kodu", "metin", 22),
    ("Masraf Merkezi Adi", "metin", 34),
    ("Sirket", "metin", 18),
    ("Gider Tipi", "metin", 13),
    ("Paylasim", "metin", 20),
    ("Tutar", "sayi", 14),
    ("Para Birimi", "metin", 10),
    ("Fatura Payi %", "sayi", 12),
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
    ("Dagitim Orani %", "sayi", 14),
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

BASLIK_RENGI = "#1F3864"
TARIH_BICIMI = "DD.MM.YYYY"
TUTAR_BICIMI = "#,##0.00"
YUZDE_BICIMI = "0%"


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
        _metin(sonuc.durum),
        " | ".join(str(u) for u in (sonuc.uyarilar or [])),
    ]


class _Bicimler:
    """Duruma ve hucre tipine gore xlsxwriter bicimlerini uretir ve saklar."""

    def __init__(self, calisma: Any) -> None:
        self._calisma = calisma
        self._onbellek: dict[tuple[str, str], Any] = {}

        self.baslik = calisma.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": BASLIK_RENGI,
            "border": 1,
            "border_color": BASLIK_RENGI,
            "align": "left",
            "valign": "vcenter",
            "text_wrap": True,
        })
        self.bolum = calisma.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": BASLIK_RENGI,
        })
        self.kalin = calisma.add_format({"bold": True})
        self.toplam_metin = calisma.add_format({
            "bold": True, "top": 6, "border_color": BASLIK_RENGI,
        })
        self.toplam_sayi = calisma.add_format({
            "bold": True, "top": 6, "border_color": BASLIK_RENGI,
            "num_format": TUTAR_BICIMI,
        })
        self.ozet_metin = calisma.add_format({"align": "left"})
        self.ozet_sayi = calisma.add_format({"num_format": "#,##0"})
        self.ozet_tutar = calisma.add_format({"num_format": TUTAR_BICIMI})
        self.ozet_yuzde = calisma.add_format({"num_format": "0.0"})

    def al(self, tip: str, durum: str) -> Any:
        """Verilen hucre tipi ve satir durumu icin bicimi dondurur."""
        anahtar = (tip, durum)
        if anahtar in self._onbellek:
            return self._onbellek[anahtar]

        ozellikler: dict[str, Any] = {"border": 1, "border_color": "#D9D9D9"}
        renk = DURUM_RENKLERI.get(durum)
        if renk:
            ozellikler["bg_color"] = renk
        if tip == "tarih":
            ozellikler["num_format"] = TARIH_BICIMI
        elif tip == "sayi":
            ozellikler["num_format"] = TUTAR_BICIMI
        elif tip == "tamsayi":
            ozellikler["num_format"] = "0"
        elif tip == "yuzde":
            ozellikler["num_format"] = YUZDE_BICIMI

        bicim = self._calisma.add_format(ozellikler)
        self._onbellek[anahtar] = bicim
        return bicim

    def mahsup(self, tip: str, renk_anahtari: str) -> Any:
        """Mahsuplasma/Kontrol sayfalari icin hucre bicimi.

        ``al`` durum kodlarina baglidir; bu sayfalarda satirin durumu farkli
        bir eksende olculur (dagitildi mi, haritada var mi), o yuzden ayri.
        """
        anahtar = (f"mahsup:{tip}", renk_anahtari)
        if anahtar in self._onbellek:
            return self._onbellek[anahtar]
        ozellikler: dict[str, Any] = {"border": 1, "border_color": "#D9D9D9"}
        renk = MAHSUP_RENKLERI.get(renk_anahtari)
        if renk:
            ozellikler["bg_color"] = renk
        if tip == "tarih":
            ozellikler["num_format"] = TARIH_BICIMI
        elif tip == "sayi":
            ozellikler["num_format"] = TUTAR_BICIMI
        elif tip == "tamsayi":
            ozellikler["num_format"] = "0"
        elif tip == "yuzde":
            ozellikler["num_format"] = YUZDE_BICIMI
        bicim = self._calisma.add_format(ozellikler)
        self._onbellek[anahtar] = bicim
        return bicim


def _sayfa_yaz(calisma: Any, ad: str, sonuclar: Sequence[Sonuc], bicimler: _Bicimler) -> None:
    """Bir veri sayfasini basliklari, filtresi ve bicimleriyle yazar."""
    sayfa = calisma.add_worksheet(ad)
    sayfa.freeze_panes(1, 0)
    sayfa.set_row(0, 30)

    for sutun, (baslik, _tip, genislik) in enumerate(KOLONLAR):
        sayfa.write_string(0, sutun, baslik, bicimler.baslik)
        sayfa.set_column(sutun, sutun, genislik)

    for indeks, sonuc in enumerate(sonuclar, start=1):
        degerler = satir_degerleri(sonuc)
        durum = sonuc.durum
        for sutun, ((_baslik, tip, _g), deger) in enumerate(zip(KOLONLAR, degerler)):
            bicim = bicimler.al(tip, durum)
            if deger is None or deger == "":
                sayfa.write_blank(indeks, sutun, None, bicim)
            elif tip == "tarih":
                sayfa.write_datetime(indeks, sutun, datetime(deger.year, deger.month, deger.day), bicim)
            elif tip in ("sayi", "tamsayi", "yuzde"):
                try:
                    sayfa.write_number(indeks, sutun, float(deger), bicim)
                except (TypeError, ValueError):
                    sayfa.write_string(indeks, sutun, str(deger), bicim)
            else:
                sayfa.write_string(indeks, sutun, str(deger), bicim)

    son_satir = max(1, len(sonuclar))
    sayfa.autofilter(0, 0, son_satir, len(KOLONLAR) - 1)
    if not sonuclar:
        sayfa.write_string(1, 0, "(Bu sayfada satir yok)", bicimler.ozet_metin)


def _ozet_yaz(calisma: Any, ozet: dict, sonuclar: Sequence[Sonuc], bicimler: _Bicimler) -> None:
    """'Ozet' sayfasini yazar."""
    sayfa = calisma.add_worksheet("Ozet")
    sayfa.set_column(0, 0, 42)
    sayfa.set_column(1, 1, 18)
    sayfa.set_column(2, 2, 18)
    sayfa.set_column(3, 3, 18)
    sayfa.set_column(4, 4, 18)
    satir = 0

    def bolum(baslik: str) -> None:
        nonlocal satir
        satir += 1
        sayfa.write_string(satir, 0, baslik, bicimler.bolum)
        for sutun in range(1, 5):
            sayfa.write_blank(satir, sutun, None, bicimler.bolum)
        satir += 1

    def cift(ad: str, deger: Any, bicim: Any | None = None) -> None:
        nonlocal satir
        sayfa.write_string(satir, 0, ad, bicimler.ozet_metin)
        if isinstance(deger, (int, float)) and not isinstance(deger, bool):
            sayfa.write_number(satir, 1, float(deger), bicim or bicimler.ozet_sayi)
        else:
            sayfa.write_string(satir, 1, _metin(deger), bicim or bicimler.ozet_metin)
        satir += 1

    sayfa.write_string(0, 0, "MASRAF MERKEZI DAGITIM OZETI", bicimler.bolum)
    for sutun in range(1, 5):
        sayfa.write_blank(0, sutun, None, bicimler.bolum)
    satir = 1

    bolum("Genel")
    cift("Uretim zamani", datetime.now().strftime("%d.%m.%Y %H:%M"))
    cift("Islenen dosya sayisi", ozet.get("dosya_sayisi", 0))
    cift("Toplam satir", ozet.get("satir_sayisi", len(sonuclar)))
    cift("Otomatik cozulme orani (%)", ozet.get("otomatik_orani", 0.0), bicimler.ozet_yuzde)
    cift("Guven esigi", ozet.get("guven_esigi", ""))
    cift("Personel dosyasi", ozet.get("personel_dosyasi", ""))
    son_donem = ozet.get("son_donem")
    cift("Personel son donemi", son_donem.strftime("%d.%m.%Y") if isinstance(son_donem, date) else "")

    dosyalar = ozet.get("dosyalar") or []
    if dosyalar:
        bolum("Islenen dosyalar")
        for yol in dosyalar:
            cift(Path(str(yol)).name, "")

    bolum("Durum dagilimi")
    sayfa.write_string(satir, 0, "Durum", bicimler.kalin)
    sayfa.write_string(satir, 1, "Satir", bicimler.kalin)
    sayfa.write_string(satir, 2, "Oran (%)", bicimler.kalin)
    satir += 1
    oranlar = ozet.get("durum_orani", {})
    for durum in (DURUM_OTOMATIK, DURUM_INCELE, DURUM_ESLESMEDI):
        adet = ozet.get("durum_dagilimi", {}).get(durum, 0)
        bicim = bicimler.al("metin", durum)
        sayfa.write_string(satir, 0, durum, bicim)
        sayfa.write_number(satir, 1, float(adet), bicimler.ozet_sayi)
        sayfa.write_number(satir, 2, float(oranlar.get(durum, 0.0)), bicimler.ozet_yuzde)
        satir += 1

    def dagilim(baslik: str, veri: dict, birinci: str) -> None:
        nonlocal satir
        if not veri:
            return
        bolum(baslik)
        sayfa.write_string(satir, 0, birinci, bicimler.kalin)
        sayfa.write_string(satir, 1, "Satir", bicimler.kalin)
        satir += 1
        for ad, adet in veri.items():
            cift(str(ad), adet)

    dagilim("Eslestirme yontemi dagilimi", ozet.get("yontem_dagilimi", {}), "Yontem")
    dagilim("Gider tipi dagilimi", ozet.get("gider_tipi_dagilimi", {}), "Gider tipi")
    dagilim("Kaynak dosya tipi dagilimi", ozet.get("kaynak_dagilimi", {}), "Kaynak tipi")

    # Masraf merkezi bazinda tutar toplami (para birimi bazinda).
    merkezler = ozet.get("masraf_merkezi_ozeti") or []
    if merkezler:
        bolum("Masraf merkezi bazinda tutar")
        for sutun, baslik in enumerate(
            ("Kod", "Ad", "Satir", "Tutar", "Para Birimi")
        ):
            sayfa.write_string(satir, sutun, baslik, bicimler.kalin)
        satir += 1
        for kayit in merkezler:
            tutarlar = kayit.get("tutarlar") or {}
            if not tutarlar:
                tutarlar = {"": None}
            ilk = True
            for para, tutar in sorted(tutarlar.items()):
                sayfa.write_string(satir, 0, _metin(kayit.get("kod")), bicimler.ozet_metin)
                sayfa.write_string(satir, 1, _metin(kayit.get("ad")), bicimler.ozet_metin)
                if ilk:
                    sayfa.write_number(satir, 2, float(kayit.get("adet", 0)), bicimler.ozet_sayi)
                    ilk = False
                else:
                    sayfa.write_blank(satir, 2, None, bicimler.ozet_sayi)
                if tutar is None:
                    sayfa.write_blank(satir, 3, None, bicimler.ozet_tutar)
                else:
                    sayfa.write_number(satir, 3, float(tutar), bicimler.ozet_tutar)
                sayfa.write_string(satir, 4, _metin(para), bicimler.ozet_metin)
                satir += 1

    paralar = ozet.get("para_birimi_toplamlari") or {}
    if paralar:
        bolum("Para birimi bazinda genel toplam")
        for para, tutar in sorted(paralar.items()):
            cift(str(para), float(tutar), bicimler.ozet_tutar)

    eksik = ozet.get("eksik_masraf_merkezleri") or []
    if eksik:
        bolum("Masraf merkezi haritasinda tanimsiz gorev yerleri")
        for ad in eksik:
            cift(str(ad), "veri/masraf_merkezi_haritasi.csv dosyasina ekleyin")

    eslesmeyen = ozet.get("eslesmeyen_kisiler") or {}
    if eslesmeyen:
        bolum("Eslesmeyen kisiler (tekrar sayisi)")
        for ad, adet in eslesmeyen.items():
            cift(str(ad), adet)

    for baslik, anahtar in (("Uyarilar", "uyarilar"), ("Hatalar", "hatalar")):
        kayitlar = ozet.get(anahtar) or []
        if kayitlar:
            bolum(baslik)
            for kayit in kayitlar:
                cift(str(kayit), "")


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
    pay = (satir.tutar / fatura_toplami * 100.0) if fatura_toplami else 0.0
    return [
        satir.kaynak,
        satir.masraf_merkezi,
        satir.masraf_merkezi_adi or "",
        satir.sirket or "",
        satir.gider_tipi,
        satir.pay_notu or "",
        round(satir.tutar, 2),
        satir.para_birimi,
        round(pay, 2),
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
        round(kontrol.dagitim_orani, 2),
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
    """Basliklari, filtresi, renkleri ve istege bagli toplam satiri olan tablo yazar."""
    sayfa = calisma.add_worksheet(ad)
    sayfa.freeze_panes(1, 0)
    sayfa.set_row(0, 30)
    for sutun, (baslik, _tip, genislik) in enumerate(kolonlar):
        sayfa.write_string(0, sutun, baslik, bicimler.baslik)
        sayfa.set_column(sutun, sutun, genislik)

    for indeks, (degerler, renk) in enumerate(zip(satirlar, renkler), start=1):
        for sutun, ((_baslik, tip, _g), deger) in enumerate(zip(kolonlar, degerler)):
            bicim = bicimler.mahsup(tip, renk)
            if deger is None or deger == "":
                sayfa.write_blank(indeks, sutun, None, bicim)
            elif tip == "tarih":
                sayfa.write_datetime(
                    indeks, sutun, datetime(deger.year, deger.month, deger.day), bicim
                )
            elif tip in ("sayi", "tamsayi", "yuzde"):
                try:
                    sayfa.write_number(indeks, sutun, float(deger), bicim)
                except (TypeError, ValueError):
                    sayfa.write_string(indeks, sutun, str(deger), bicim)
            else:
                sayfa.write_string(indeks, sutun, str(deger), bicim)

    son = len(satirlar)
    sayfa.autofilter(0, 0, max(1, son), len(kolonlar) - 1)
    if not satirlar:
        sayfa.write_string(1, 0, bos_mesaj, bicimler.ozet_metin)
        return sayfa

    if toplam_sutunlari:
        satir_no = son + 1
        sayfa.write_string(satir_no, 0, "TOPLAM", bicimler.toplam_metin)
        for sutun in range(1, len(kolonlar)):
            if sutun in toplam_sutunlari:
                harf = _sutun_harfi(sutun)
                sayfa.write_formula(
                    satir_no, sutun, f"=SUM({harf}2:{harf}{son + 1})",
                    bicimler.toplam_sayi,
                )
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
        toplam_sutunlari=(2, 3, 4, 5, 6, 7, 8, 9),
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
        if mahsup is not None:
            _mahsuplasma_yaz(calisma, mahsup, bicimler)
        _sayfa_yaz(calisma, "Sonuc", list(sonuclar), bicimler)
        _sayfa_yaz(
            calisma, "Incele", [s for s in sonuclar if s.durum == DURUM_INCELE], bicimler
        )
        _sayfa_yaz(
            calisma, "Eslesmedi", [s for s in sonuclar if s.durum == DURUM_ESLESMEDI], bicimler
        )
        _ozet_yaz(calisma, ozet or {}, list(sonuclar), bicimler)
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
                elif tip == "sayi":
                    hucreler.append(f"{float(deger):.2f}")
                else:
                    hucreler.append(str(deger))
            yazici.writerow(hucreler)
    return str(hedef)
