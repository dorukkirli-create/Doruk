"""Excel ve CSV cikti uretimi.

Uretilen Excel dosyasi finans ekibinin dogrudan calisacagi belgedir; dort
sayfa icerir:

    Sonuc     - tum satirlar, tum kolonlar
    Incele    - durum = INCELE (guven dusuk veya uyari var)
    Eslesmedi - durum = ESLESMEDI (kisi bulunamadi)
    Ozet      - durum/yontem dagilimi, masraf merkezi bazinda tutar toplami

Tasarim ilkesi: kullanici HER satirda neden o sonuca varildigini gorebilmeli.
Bu yuzden 'Eslestirme Yontemi', 'Guven', 'Eslestirme Aciklamasi' ve 'Uyarilar'
kolonlari ciktida her zaman yer alir ve satirlar duruma gore renklendirilir.

Yalniz ``xlsxwriter`` ve standart kutuphane kullanilir; modul tek basina
import edilebilir.
"""

from __future__ import annotations

import csv
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

from masraf.modeller import DURUM_ESLESMEDI, DURUM_INCELE, DURUM_OTOMATIK, Sonuc

__all__ = ["excel_yaz", "csv_yaz", "KOLONLAR", "varsayilan_cikti_adi"]

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


def excel_yaz(sonuclar: list[Sonuc], yol: str, ozet: dict) -> str:
    """Sonuclari bicimlendirilmis dort sayfali Excel dosyasina yazar.

    Args:
        sonuclar: Boru hattinin urettigi sonuc listesi.
        yol: Yazilacak .xlsx dosyasinin yolu.
        ozet: ``Boru.ozet()`` ciktisi.

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
