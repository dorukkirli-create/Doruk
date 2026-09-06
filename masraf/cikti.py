"""Excel ve CSV cikti uretimi.

Uretilen Excel dosyasi finans ekibinin dogrudan calisacagi belgedir. Sayfalar
IS AKISI SIRASINDADIR: once muhasebeye gidecek olan, sonra kaniti.

    Ozet        - KAPAK. Tutarlar, mutabakat durumu, sirket kirilimi, dikkat
                  notlari ve sayfa rehberi. Dosya acilinca ilk gorulen sayfa.
    Mahsuplasma - NIHAI CIKTI. Her fatura icin hangi projeye ne kadar
                  yazilacagi. Muhasebeye giden tablo budur.
    Kontrol     - Mutabakat. Her fatura icin okunan / yinelenen / dagitilan /
                  dagitilamayan tutar. 'Fark' sutunu sifir olmak zorundadir.
                  Altinda: fatura detay listeleri capraz kontrolu, okunmayan
                  ekler, isaret celiskileri, dagilima girmeyen satirlar.
    Dosyalar    - Dosya envanteri. Ust duzey dosyalar ve mail eklerinin her
                  biri: okundu / kutuk / detay listesi / atlandi (PDF) / ayni
                  icerik / okunamadi, satir sayisi ve okunan tutar. Elle
                  kontrol icin 'hangi dosyadan ne cikti' listesi.
    Sirket Kirilimi - tuzel kisi ustte, projeleri altinda (1C 'Firm 2').
    Harita Onerileri - haritada tanimsiz gorev yerleri, hazir CSV satiri
                  (yalnizca oneri varsa yazilir).
    Sonuc       - tum satirlar, tum kolonlar (mahsuplasmanin dayanagi)
    Incele      - durum = INCELE (guven dusuk veya uyari var)
    Eslesmedi   - durum = ESLESMEDI (kisi bulunamadi)

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
    # Elle dagitilmis dosyada insanin yazdigi santiye/sirket etiketi. Dagitimda
    # KULLANILMAZ (personel kaydi esastir); tablonun sirketiyle celisirse
    # muhasebe burada gorur. Evrak no denetim izidir.
    ("Elle Dagitim Etiketi", "metin", 20),
    ("Evrak / Fatura No", "metin", 22),
    ("Durum", "metin", 34),
)

#: Kontrol (mutabakat) sayfasi kolonlari.
#: 'Dosyalar' sayfasi: dosya envanteri.
DOSYA_KOLONLARI: tuple[tuple[str, str, int], ...] = (
    ("Sira", "tamsayi", 6),
    ("Dosya", "metin", 46),
    ("Nereden Geldi", "metin", 48),
    ("Tur", "metin", 22),
    ("Durum", "metin", 15),
    ("Satir", "tamsayi", 8),
    ("Tutarli Satir", "tamsayi", 12),
    ("Okunan Tutar", "sayi", 14),
    ("Para", "metin", 6),
    ("Boyut (KB)", "sayi", 11),
    ("Not / Sebep", "metin", 80),
)

#: Envanter durumu -> satir rengi.
_ENVANTER_RENKLERI = {
    "OKUNDU": "tamam", "KUTUK": "", "DETAY LISTESI": "", "MAIL": "baslik",
    "ATLANDI": "uyari", "SATIR YOK": "uyari", "AYNI ICERIK": "uyari",
    "OKUNAMADI": "engel", "PERSONEL": "uyari",
}


def dosya_satir_degerleri(sira: int, k: dict) -> list[Any]:
    """Bir envanter kaydini DOSYA_KOLONLARI sirasina cevirir."""
    boyut = k.get("boyut")
    return [
        sira,
        _metin(k.get("ad")),
        _metin(k.get("kaynak")),
        _metin(k.get("tur")),
        _metin(k.get("durum")),
        int(k.get("satir") or 0),
        int(k.get("tutarli_satir") or 0),
        k.get("tutar"),
        _metin(k.get("para_birimi")),
        round(boyut / 1024, 1) if isinstance(boyut, (int, float)) else None,
        _metin(k.get("sebep")),
    ]


#: Envanter gosterim sirasi: once dagilima girenler, sonra karar gerektirenler,
#: en sonda bilgi amacli olanlar (PDF'ler, tekrarlar).
_ENVANTER_SIRASI = {
    "OKUNDU": 0, "KUTUK": 1, "DETAY LISTESI": 2, "SATIR YOK": 3, "OKUNAMADI": 4,
    "PERSONEL": 5, "AYNI ICERIK": 6, "ATLANDI": 7, "MAIL": 8,
}


def envanteri_sirala(envanter: Sequence[dict]) -> list[dict]:
    """Duruma, sonra kaynak zincirine ve ada gore siralar (kararli)."""
    return sorted(
        envanter,
        key=lambda k: (_ENVANTER_SIRASI.get(str(k.get("durum")), 9),
                       str(k.get("kaynak") or ""), str(k.get("ad") or "").lower()),
    )


def _dosyalar_yaz(calisma: Any, envanter: Sequence[dict], bicimler: "_Bicimler") -> None:
    """'Dosyalar' sayfasi: her dosya ve ek icin okundu / atlandi / neden."""
    envanter = envanteri_sirala(envanter)
    degerler = [dosya_satir_degerleri(i, k) for i, k in enumerate(envanter, 1)]
    renkler = [_ENVANTER_RENKLERI.get(str(k.get("durum")), "") for k in envanter]
    sayfa = _tablo_yaz(
        calisma, "Dosyalar", DOSYA_KOLONLARI, degerler, renkler, bicimler,
        bos_mesaj="(Dosya envanteri yok)", toplam_sutunlari=(5, 6, 7),
    )
    satir = len(degerler) + 3
    sayfa.write_string(satir, 0, "Nasil okunur", bicimler.bolum)
    satir += 1
    for metin in (
        "OKUNDU = gider satiri uretti, dagilima girdi. 'Okunan Tutar' o dosyadan okunan tutarlarin "
        "toplamidir; dosyayi acip kendi toplamiyla karsilastirabilirsiniz.",
        "KUTUK = kisi listesi (katilimci, saglik, sigorta). Defter beslemesinde kullanildi, dagilima girmedi.",
        "DETAY LISTESI = tutar kolonu olmayan fatura detay listesi; kisileri yansitma dosyasiyla capraz kontrol edildi.",
        "ATLANDI = tablo olmayan ek (PDF, docx). Acilmadi. Tutari yalnizca burada olan bir fatura varsa "
        "bu tabloda YOKTUR; elle eklenmeli.",
        "AYNI ICERIK = daha once okunan bir dosyayla birebir ayni; cift sayim olmasin diye atlandi.",
        "SATIR YOK = acildi ama gider satiri cikmadi; kolon adlari taninmamis olabilir "
        "(veri/kolon_esanlamlilari.csv).",
        "OKUNAMADI = acilamadi (bozuk, parola korumali, bos). Dosya 1_FATURALAR'da birakilir.",
        "MAIL = Outlook mesaji; ekleri ayri satirlardadir.",
    ):
        sayfa.write_string(satir, 0, metin, bicimler.ozet_metin)
        satir += 1


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
    ("Tutarsiz Satir", "tamsayi", 13),
    ("Faturada Yazan Toplam", "sayi", 18),
    ("Beyan Farki", "sayi", 12),
    ("Dagitim Orani", "yuzde", 13),
    ("Mutabakat", "metin", 40),
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
    "onceki_donem": "Onceki donem (kayit yok)",
    "personel_dosyasi_eski": "Personel dosyasi eski",
    "ilk_donem_oncesi": "Ilk donemden once",
    "yardimci_defter": "1C listesi (donem yok)",
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


def _kaynak_santiye(kaynak: Any, ek: dict) -> str:
    """Satirin kendi santiye etiketi; yoksa yinelenen elle dosyadan devralinan."""
    if kaynak not in (None, ""):
        return _metin(kaynak)
    devralinan = ek.get("kaynak_santiye_devralindi") if isinstance(ek, dict) else None
    return f"{_metin(devralinan)} (elle dosyadan)" if devralinan not in (None, "") else ""


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
        round(satir.tutar, 2) if isinstance(satir.tutar, (int, float)) else satir.tutar,
        _metin(satir.para_birimi),
        _kaynak_santiye(satir.masraf_merkezi_kaynak, ek),
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
            "font_name": BASLIK_YAZI, "font_size": 15, "bold": True,
            "font_color": LACIVERT, "num_format": TUTAR_BICIMI, "valign": "vcenter"})
        self.kpi_tamsayi = calisma.add_format({
            "font_name": BASLIK_YAZI, "font_size": 15, "bold": True,
            "font_color": LACIVERT, "num_format": TAMSAYI_BICIMI, "valign": "vcenter"})
        self.kpi_metin = calisma.add_format({
            "font_name": BASLIK_YAZI, "font_size": 15, "bold": True,
            "font_color": LACIVERT, "valign": "vcenter"})
        self.kpi_etiket = f(font_size=9, font_color=ORTA_GRI, valign="top", text_wrap=True)
        self.durum_iyi = f(bold=True, font_color=BEYAZ, bg_color="#2E7D32", align="center",
                           valign="vcenter")
        self.durum_kotu = f(bold=True, font_color=BEYAZ, bg_color="#C62828", align="center",
                            valign="vcenter")
        self.durum_uyari = f(bold=True, font_size=10, font_color=KOYU_GRI, bg_color="#FFF2CC", border=1, border_color="#BF9000", valign="vcenter", text_wrap=True)
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


def _kapak_dikkat_nedenleri(ozet: dict, mahsup: Any, oneriler: Any) -> list[str]:
    """Mutabakat kapali olsa da onaydan once bakilmasi gereken seyler.

    Kapak durumunu yesilden sariya ceviren nedenler. Her biri Kontrol ya da
    Ozet sayfasinda ayrintili olarak yazilidir; burada yalnizca sayilir.
    """
    nedenler: list[str] = []
    supheler = list(getattr(mahsup, "supheler", None) or [])
    if supheler:
        nedenler.append(f"{len(supheler)} kalemde yineleme şüphesi")
    tarihsiz = sum(int(getattr(k, "tarihsiz_satir", 0) or 0) for k in (getattr(mahsup, "kontrol", None) or []))
    if tarihsiz:
        nedenler.append(f"{tarihsiz} satır tarihsiz (yineleme kontrolüne girmedi)")
    if getattr(mahsup, "isaret_celiskileri", None):
        nedenler.append(f"{len(mahsup.isaret_celiskileri)} işaret çelişkisi")
    detay = [d for d in (getattr(mahsup, "detay_kontrolleri", None) or []) if not d.tutarli_mi]
    if detay:
        nedenler.append(f"{len(detay)} fatura detay listesiyle uyuşmuyor")
    if getattr(mahsup, "tutarsiz_satir_sayisi", 0):
        nedenler.append(f"{mahsup.tutarsiz_satir_sayisi} satırda tutar okunamadı")
    if any(not m.get("haritada_var", True) for m in (mahsup.merkez_ozeti() if mahsup is not None else [])):
        nedenler.append("haritada tanımsız masraf merkezi var")
    if oneriler:
        nedenler.append(f"{len(oneriler)} görev yeri haritada yok")
    atlanan = [a for a in (ozet.get("atlanan_ekler") or []) if not getattr(a, "tekrar", False)]
    if atlanan:
        nedenler.append(f"{len(atlanan)} ek okunmadı")
    onemli_uyari = [
        u for u in (ozet.get("uyarilar") or [])
        if "Ogrenildi" not in str(u) and "kisi kutugu" not in str(u)
    ]
    if onemli_uyari:
        nedenler.append(f"{len(onemli_uyari)} uyarı")
    return nedenler


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
    sayfa.set_column(2, 7, 19)
    sayfa.set_column(8, 8, 3)
    sayfa.set_landscape()
    sayfa.fit_to_pages(1, 0)

    # --- Ust bant --------------------------------------------------------
    for r in range(0, 4):
        sayfa.set_row(r, 22 if r else 8)
        for c in range(0, 9):
            sayfa.write_blank(r, c, None, bicimler.kapak_bant)
    sayfa.set_row(1, 34)
    sayfa.merge_range(1, 1, 1, 7, "Masraf Merkezi Dağıtımı", bicimler.kapak_baslik)
    donem = _gider_donemi(sonuclar)
    dosya_sayisi = ozet.get("dosya_sayisi", 0)
    alt = (f"{donem}   |   {dosya_sayisi} kaynak dosya, {len(sonuclar)} satır   |   "
           f"üretim {datetime.now():%d.%m.%Y %H:%M}   |   Rencons Heavy Industries")
    sayfa.merge_range(2, 1, 2, 7, alt, bicimler.kapak_alt)

    satir = 5

    # --- Para birimi bazinda KPI satiri ---------------------------------
    toplamlar = mahsup.toplamlar() if mahsup is not None else {}
    if not toplamlar:
        sayfa.write_string(satir, 1, "Dağıtılacak tutarlı satır bulunamadı.", bicimler.ozet_metin)
        satir += 2
    for para, d in toplamlar.items():
        kpi = (
            ("Okunan", d["gelen"], bicimler.kpi_sayi, "dosyalarda görülen toplam"),
            ("Yinelenen", d["yinelenen"], bicimler.kpi_sayi, "başka dosyada zaten sayıldı"),
            ("Net", d["net"], bicimler.kpi_sayi, "gerçekten dağıtılacak"),
            ("Dağıtılan", d["dagitilan"], bicimler.kpi_sayi, "projelere yazıldı"),
            ("Dağıtılamayan", d["dagitilamayan"], bicimler.kpi_sayi, "kişi / merkez bulunamadı"),
            ("Dağıtım oranı", d["oran"] / 100.0, None, "dağıtılan / net"),
        )
        sayfa.set_row(satir, 14)
        sayfa.write_string(satir, 1, f"TUTARLAR ({para})", bicimler.kalin)
        satir += 1
        sayfa.set_row(satir, 30)
        for i, (ad, deger, bicim, aciklama) in enumerate(kpi):
            c = 1 + i
            if bicim is None:
                yuzde = calisma.add_format({
                    "font_name": BASLIK_YAZI, "font_size": 15, "bold": True,
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
        incele_sayisi = sum(1 for s in sonuclar if s.durum == DURUM_INCELE)
        eslesmedi_sayisi = sum(1 for s in sonuclar if s.durum == DURUM_ESLESMEDI)
        hatalar = list(ozet.get("hatalar") or [])
        dikkat_nedenleri = _kapak_dikkat_nedenleri(ozet, mahsup, oneriler)
        if hatalar:
            sayfa.merge_range(satir, 1, satir, 7,
                              f"MUTABAKAT AÇIK   |   {len(hatalar)} dosya okunamadı veya işlenirken hata verdi; "
                              "tutarları bu tabloda yok. Tablo muhasebeye GÖNDERİLMEMELİ. Liste 'Dikkat' bölümünde.",
                              bicimler.durum_kotu)
        elif mahsup.kapali_mi and not incele_sayisi and not eslesmedi_sayisi and not dikkat_nedenleri:
            sayfa.merge_range(satir, 1, satir, 7,
                              "MUTABAKAT KAPALI   |   Okunan = Yinelenen + Dağıtılan + Dağıtılamayan. "
                              "Para kaybolmadı; inceleme bekleyen satır yok. Tablo onaya hazır.",
                              bicimler.durum_iyi)
        elif mahsup.kapali_mi and not incele_sayisi and not eslesmedi_sayisi:
            sayfa.merge_range(satir, 1, satir, 7,
                              "MUTABAKAT KAPALI, DİKKAT   |   Para kaybolmadı; inceleme bekleyen satır yok. "
                              f"Onaydan önce bakılmalı: {'; '.join(dikkat_nedenleri)}.",
                              bicimler.durum_uyari)
        elif mahsup.kapali_mi:
            # Para kaybolmadi ama kimlik/merkez karari bekleyen satirlar var:
            # bu bir TASLAKTIR. 'Gonderilebilir' demek onay akisini atlatir.
            bekleyen = " + ".join(
                p for p in (
                    f"{incele_sayisi} satır inceleme" if incele_sayisi else "",
                    f"{eslesmedi_sayisi} satır kişi bulunamadı ((DAGITILAMAYAN))" if eslesmedi_sayisi else "",
                ) if p)
            ek_not = f" Ayrıca: {'; '.join(dikkat_nedenleri)}." if dikkat_nedenleri else ""
            sayfa.merge_range(satir, 1, satir, 7,
                              f"MUTABAKAT KAPALI, TASLAK   |   Para kaybolmadı; {bekleyen} "
                              "bekliyor. 'Incele' ve 'Eslesmedi' sayfaları görülmeden onaya sunulmamalı." + ek_not,
                              bicimler.durum_uyari)
        else:
            acik = ", ".join(
                f"{k.kaynak} ({k.acik_sebebi or f'{k.fark:+.2f}'})" for k in mahsup.acik_kontroller)
            sayfa.merge_range(satir, 1, satir, 7,
                              f"MUTABAKAT AÇIK   |   {acik}. Tablo muhasebeye GÖNDERİLMEMELİ.",
                              bicimler.durum_kotu)
        satir += 1
        # Onay akisi: durum satiri kimin ne yaptigini soylemez; bu satirlar
        # elle doldurulur. Otomasyon 'hazirlar', insan 'kontrol eder' ve 'onaylar'.
        sayfa.set_row(satir, 16)
        sayfa.write_string(satir, 1, "Hazırlayan: otomasyon", bicimler.etiket)
        sayfa.write_string(satir, 3, "Kontrol eden: ____________   tarih: ________", bicimler.etiket)
        sayfa.write_string(satir, 6, "Onaylayan: ____________   tarih: ________", bicimler.etiket)
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

    bolum("Satır durumu")
    sayfa.set_row(satir, 20)
    for i, b in enumerate(("Durum", "Satır", "Oran")):
        sayfa.write_string(satir, 1 + i, b, bicimler.baslik)
    sayfa.merge_range(satir, 4, satir, 7, "Ne demek", bicimler.baslik)
    satir += 1
    aciklama = {
        DURUM_OTOMATIK: "Kimlik kesin, uyarı yok; olduğu gibi kaydedilebilir",
        DURUM_INCELE: "Sistem sonuç buldu ama emin değil; gerekçesi satırda yazılı",
        DURUM_ESLESMEDI: "Kişi bulunamadı; tutar (DAGITILAMAYAN) satırında duruyor",
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
            bolum("Şirket kırılımı   (tüzel kişi ve en büyük projesi)")
            tablo_basligi(("Şirket", "Tutar", "Para", "Pay", "Kişi", "En büyük proje", "Proje payı"))
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
            bolum("En büyük masraf merkezleri   (ilk 10)")
            tablo_basligi(("Masraf merkezi", "Tutar", "Para", "Pay", "Kişi", "Şirket", "Durum"))
            for i, m in enumerate(merkezler):
                zebra = i % 2 == 1
                dagitilamayan = str(m["masraf_merkezi"]) == "(DAGITILAMAYAN)"
                renk = "" if m["haritada_var"] else ("engel" if dagitilamayan else "uyari")
                durum_metni = "" if m["haritada_var"] else ("kişi bulunamadı" if dagitilamayan else "haritada yok")
                sayfa.write_string(satir, 1, str(m["masraf_merkezi"]), bicimler.mahsup("metin", renk, zebra))
                sayfa.write_number(satir, 2, m["tutar"], bicimler.mahsup("sayi", renk, zebra))
                sayfa.write_string(satir, 3, m["para_birimi"], bicimler.mahsup("metin", renk, zebra))
                sayfa.write_number(satir, 4, m["pay_yuzde"] / 100.0, bicimler.mahsup("yuzde", renk, zebra))
                sayfa.write_number(satir, 5, m["kisi_sayisi"], bicimler.mahsup("tamsayi", renk, zebra))
                sayfa.write_string(satir, 6, str(m.get("sirket") or ""), bicimler.mahsup("metin", renk, zebra))
                sayfa.write_string(satir, 7, durum_metni, bicimler.mahsup("metin", renk, zebra))
                satir += 1

    # --- Dikkat ----------------------------------------------------------
    dikkat: list[str] = []
    if mahsup is not None:
        for c in getattr(mahsup, "isaret_celiskileri", []) or []:
            dikkat.append("İşaret çelişkisi: " + c.kisa_aciklama())
        if getattr(mahsup, "tutarsiz_satir_sayisi", 0):
            dikkat.append(f"{mahsup.tutarsiz_satir_sayisi} satırda tutar okunamadı; dağılıma girmedi.")
    for pb, u in sorted((getattr(mahsup, "sirket_uyusmazligi", None) or {}).items()) if mahsup is not None else []:
        ciftler = u.get("ciftler") or {}
        en = next(iter(ciftler.items()), None)
        ornek = f" En büyüğü {en[0]}: {en[1]['satir']} satır, {en[1]['tutar']:,.2f}." if en else ""
        dikkat.append(
            f"Kaynak dosyadaki şirket etiketi {u['satir']} satırda tablonun şirketinden farklı "
            f"({pb} {u['tutar']:,.2f}).{ornek} Etiket faturanın kesildiği tarafsa bu beklenen "
            "yansıtmadır; kişinin işvereniyse personel kaydıyla çelişir. Çiftler 'Kontrol' sayfasında."
        )
    if oneriler:
        dikkat.append(f"{len(oneriler)} görev yeri masraf merkezi haritasında tanımlı değil. "
                      "'Harita Onerileri' sayfasında hazır satırlar var.")
    atlanan_ekler = [a for a in (ozet.get("atlanan_ekler") or []) if not getattr(a, "tekrar", False)]
    if atlanan_ekler:
        # PDF faturalar ve diger tablo olmayan ekler okunmaz; sessiz atlanmaz.
        pdf_sayisi = sum(1 for a in atlanan_ekler if str(a).lower().split("  [")[0].endswith(".pdf"))
        dikkat.append(
            f"{len(atlanan_ekler)} ek okunmadı ({pdf_sayisi} PDF). PDF'ler tutarı kişi kırılımı olmadan "
            "taşır; tutar yalnızca PDF'te olan bir fatura varsa bu tabloda YOKTUR. Liste 'Kontrol' sayfasında."
        )
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
            sayfa.merge_range(satir, 1, satir, 7, f"... ve {len(dikkat) - 12} not daha (Kontrol ve Sonuc sayfalarında).",
                              bicimler.not_metni)
            satir += 1

    # --- Dosya envanteri -------------------------------------------------
    envanter = list(ozet.get("dosya_envanteri") or [])
    if envanter:
        from collections import Counter
        sayim = Counter(str(k.get("durum")) for k in envanter)
        bolum("Dosyalar   (hangisi okundu, hangisi atlandı)")
        sayfa.merge_range(
            satir, 1, satir, 7,
            "   ".join(f"{d}: {n}" for d, n in sayim.most_common())
            + "   |   tam liste ve sebepler 'Dosyalar' sayfasında",
            bicimler.ozet_metin)
        satir += 1
        tablo_basligi(("Dosya", "Nereden geldi", "Durum", "Satır", "Okunan tutar", "Para", "Not"))
        gosterilecek = [k for k in envanteri_sirala(envanter) if str(k.get("durum")) != "MAIL"][:30]
        for i, k in enumerate(gosterilecek):
            zebra = i % 2 == 1
            renk = _ENVANTER_RENKLERI.get(str(k.get("durum")), "")
            sayfa.write_string(satir, 1, _metin(k.get("ad")), bicimler.mahsup("metin", renk, zebra))
            sayfa.write_string(satir, 2, _metin(k.get("kaynak")), bicimler.mahsup("metin", renk, zebra))
            sayfa.write_string(satir, 3, _metin(k.get("durum")), bicimler.mahsup("metin", renk, zebra))
            sayfa.write_number(satir, 4, int(k.get("satir") or 0), bicimler.mahsup("tamsayi", renk, zebra))
            if isinstance(k.get("tutar"), (int, float)):
                sayfa.write_number(satir, 5, float(k["tutar"]), bicimler.mahsup("sayi", renk, zebra))
            else:
                sayfa.write_string(satir, 5, "", bicimler.mahsup("metin", renk, zebra))
            sayfa.write_string(satir, 6, _metin(k.get("para_birimi")), bicimler.mahsup("metin", renk, zebra))
            sayfa.write_string(satir, 7, _metin(k.get("sebep"))[:120], bicimler.mahsup("metin", renk, zebra))
            satir += 1
        if len(envanter) - sum(1 for k in envanter if str(k.get("durum")) == "MAIL") > 30:
            sayfa.merge_range(satir, 1, satir, 7, "... devamı 'Dosyalar' sayfasında.", bicimler.not_metni)
            satir += 1
        satir += 1
    else:
        dosyalar = ozet.get("dosyalar") or []
        if dosyalar:
            bolum("İşlenen dosyalar")
            for i, yol in enumerate(dosyalar):
                zebra = i % 2 == 1
                sayfa.merge_range(satir, 1, satir, 7, Path(str(yol)).name, bicimler.mahsup("metin", "", zebra))
                satir += 1
    bolum("Kaynak veri")
    for ad, deger in (
        ("Personel ana verisi", Path(str(ozet.get("personel_dosyasi") or "")).name or "-"),
        ("1C personel listesi", Path(str(ozet.get("yardimci_dosyasi") or "")).name or "- (kullanılmadı)"),
        ("Masraf merkezi haritası", Path(str(ozet.get("harita_dosyasi") or "")).name or "-"),
        ("Personel son dönemi", ozet["son_donem"].strftime("%d.%m.%Y") if isinstance(ozet.get("son_donem"), date) else "-"),
        ("Otomatik kabul eşiği", f"güven ≥ {ozet.get('guven_esigi', '')} ve uyarı yok"),
    ):
        sayfa.write_string(satir, 1, ad, bicimler.ozet_metin)
        sayfa.merge_range(satir, 2, satir, 7, str(deger), bicimler.ozet_metin)
        satir += 1

    # --- Sayfa rehberi ---------------------------------------------------
    bolum("Bu dosyada ne var   (sayfa adına tıklayın)")
    rehber = (
        ("Mahsuplasma", "MUHASEBEYE GİDEN TABLO. Her fatura için şirkete ve projeye ne kadar yazılacağı."),
        ("Sirket Kirilimi", "Tüzel kişi üstte, projeleri altında."),
        ("Kontrol", "Mutabakat. Fark sütunu sıfır olmak zorunda."),
        ("Dosyalar", "Dosya envanteri: hangi dosya ve ek okundu, hangisi atlandı, neden; satır ve tutar."),
        ("Harita Onerileri", "Tanımsız görev yerleri ve haritaya yapıştırmaya hazır satırlar."),
        ("Sonuc", "Bütün satırlar: kişinin nasıl bulunduğu, güven, gerekçe, evrak no, mail konusu."),
        ("Incele", "Elle bakılacak satırlar."),
        ("Eslesmedi", "Kişi bulunamayan satırlar."),
    )
    if oneriler is None:
        # Sayfa yazilmadiysa rehberde kirik baglanti olmasin.
        rehber = tuple(r for r in rehber if r[0] != "Harita Onerileri")
    if not ozet.get("dosya_envanteri"):
        rehber = tuple(r for r in rehber if r[0] != "Dosyalar")
    for i, (ad, aciklama) in enumerate(rehber):
        zebra = i % 2 == 1
        sayfa.write_url(satir, 1, f"internal:'{ad}'!A1", bicimler.mahsup("metin", "", zebra), ad)
        sayfa.merge_range(satir, 2, satir, 7, aciklama, bicimler.mahsup("metin", "", zebra))
        satir += 1

    satir += 1
    sayfa.merge_range(satir, 1, satir, 7,
                      "Otomasyon çevrimdışı çalışır; personel ve fatura verisi bilgisayardan dışarı çıkmaz. "
                      "Bu dosya kişisel veri içerir, paylaşırken dikkat edin.",
                      bicimler.not_metni)


def _gider_donemi(sonuclar: Sequence[Sonuc]) -> str:
    """Satirlardaki belge tarihlerinden okunakli donem etiketi ('Temmuz 2026')."""
    aylar = ("Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", "Temmuz",
             "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık")
    tarihler = [s.satir.belge_tarihi for s in sonuclar
                if getattr(s.satir, "belge_tarihi", None) and getattr(s.satir, "tutar", None) is not None]
    if not tarihler:
        return "Dönem belirsiz"
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
        getattr(satir, "elle_etiket", None) or "",
        getattr(satir, "evrak_no", None) or "",
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
        kontrol.tutarsiz_satir,
        kontrol.beyan_toplam,
        kontrol.beyan_farki,
        round(kontrol.dagitim_orani / 100.0, 4),
        kontrol_durum_metni(kontrol),
    ]


def kontrol_durum_metni(kontrol: Any) -> str:
    """Mutabakat sutunu: KAPANDI / KAPANDI, DIKKAT - ... / ACIK - ..."""
    if not kontrol.kapali_mi:
        return f"ACIK - {kontrol.acik_sebebi}"
    suphe = getattr(kontrol, "suphe_notu", "") or ""
    return f"KAPANDI, DIKKAT - {suphe}" if suphe else "KAPANDI"


def kontrol_rengi(kontrol: Any) -> str:
    if not kontrol.kapali_mi:
        return "engel"
    return "uyari" if (getattr(kontrol, "suphe_notu", "") or "") else "tamam"


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
    from masraf.mahsuplasma import DAGITILAMAYAN

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
            if proje["masraf_merkezi"] == DAGITILAMAYAN:
                # Harita sorunu degil: kisi ya da merkez bulunamadi.
                uyari, renk = "KISI / MERKEZ BULUNAMADI", "engel"
            elif proje["haritada_var"]:
                uyari, renk = "", "tamam"
            else:
                uyari, renk = "HARITADA TANIMLI DEGIL", "uyari"
            degerler.append([
                "    " + str(proje["masraf_merkezi_adi"] or proje["masraf_merkezi"]),
                "proje", proje["masraf_merkezi"],
                proje["tutar"], grup["para_birimi"], proje["pay_yuzde"] / 100.0,
                proje["satir_sayisi"], proje["kisi_sayisi"], uyari,
            ])
            renkler.append(renk)
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
    k_renkler = [kontrol_rengi(k) for k in tablo.kontrol]
    sayfa = _tablo_yaz(
        calisma, "Kontrol", KONTROL_KOLONLARI, k_degerler, k_renkler, bicimler,
        bos_mesaj="(Kontrol edilecek fatura yok)",
        toplam_sutunlari=(2, 3, 4, 5, 6, 8, 9, 10, 11),   # Fark, Beyan Farki ve oran toplanmaz
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
        "Tutarsiz Satir = dosyada var ama tutari okunamayan satirlar. Sifir "
        "degilse fatura toplaminin bir kismi dagilima girmemistir; mutabakat "
        "ACIK kalir.",
        "Faturada Yazan Toplam = kaynak dosyanin kendi beyan ettigi toplam "
        "(varsa). Beyan Farki = bu toplam ile okunan arasindaki fark; sifir "
        "degilse okuyucu bir seyi kacirmistir.",
    ):
        sayfa.write_string(satir, 0, metin, bicimler.ozet_metin)
        satir += 1

    if getattr(tablo, "detay_kontrolleri", None):
        satir += 1
        sayfa.write_string(satir, 0, "Fatura detay listeleri (capraz kontrol)", bicimler.bolum)
        satir += 1
        sayfa.write_string(
            satir, 0,
            "Tedarikci her fatura icin tutarsiz bir katilimci listesi gonderir; "
            "tutar yansitma dosyasindadir. Iki listedeki kisiler ayni olmali.",
            bicimler.ozet_metin,
        )
        satir += 1
        for dk in tablo.detay_kontrolleri:
            bicim = bicimler.durum_iyi if dk.tutarli_mi else bicimler.durum_kotu
            sayfa.write_string(satir, 0, "TAMAM" if dk.tutarli_mi else "UYUSMUYOR", bicim)
            sayfa.write_string(satir, 1, f"{dk.aciklama()}  [{dk.kaynak}]", bicimler.ozet_metin)
            satir += 1

    supheler = list(getattr(tablo, "supheler", None) or [])
    if supheler:
        satir += 1
        sayfa.write_string(satir, 0, "Yineleme suphesi (insan bakmali)", bicimler.bolum)
        satir += 1
        sayfa.write_string(
            satir, 0,
            "Otomatik karar verilemeyen kalemler. Uc tur: ayni kisi ve tarih ama tutar farkli "
            "(ikisi de dagitima girdi, biri fazla olabilir); tarihsiz satir (yineleme kontrolune "
            "giremedi); ayni gun ve tutar ama isimlerin tek kelimesi ortak (yinelenen sayilip elendi, "
            "farkli kisiyse geri eklenmeli). Her satirda karar ve gerekce yazili.",
            bicimler.ozet_metin)
        satir += 1
        for sp in supheler:
            sayfa.write_string(satir, 0, sp.aciklama(), bicimler.ozet_metin)
            satir += 1

    uyusmazlik = dict(getattr(tablo, "sirket_uyusmazligi", None) or {})
    if uyusmazlik:
        satir += 1
        sayfa.write_string(satir, 0, "Kaynak dosyadaki sirket etiketi ile tablo sirketi farkli olan satirlar",
                           bicimler.bolum)
        satir += 1
        sayfa.write_string(
            satir, 0,
            "Acente cogu zaman faturanin KESILDIGI tarafi yazar (orn. RHI); tablo ise kisinin personel "
            "kaydindaki tuzel kisiyi kullanir. Fark sirketler arasi yansitmanin kendisidir ve hata "
            "degildir. Ama etiket kisinin isverenini kastediyorsa personel kaydi ile celisir; o zaman "
            "1C listesi ya da kaynak dosya duzeltilmeli. Etiketler Mahsuplasma sayfasinda "
            "'Elle Dagitim Etiketi' kolonunda.",
            bicimler.ozet_metin)
        satir += 1
        for pb, u in sorted(uyusmazlik.items()):
            sayfa.write_string(satir, 0, f"{pb}: toplam {u['satir']} satir, {u['tutar']:,.2f}", bicimler.kalin)
            satir += 1
            for ad, c in (u.get("ciftler") or {}).items():
                sayfa.write_string(satir, 0, f"    {ad}", bicimler.ozet_metin)
                sayfa.write_string(satir, 1, f"{c['satir']} satir", bicimler.ozet_metin)
                sayfa.write_number(satir, 2, float(c["tutar"]), bicimler.ozet_tutar)
                satir += 1

    if getattr(tablo, "isaret_celiskileri", None):
        satir += 1
        sayfa.write_string(satir, 0, "Isaret celiskileri", bicimler.bolum)
        satir += 1
        for celiski in tablo.isaret_celiskileri:
            sayfa.write_string(satir, 0, celiski.aciklama(), bicimler.ozet_metin)
            satir += 1

    atlanan_ekler = list(getattr(tablo, "atlanan_ekler", None) or [])
    if atlanan_ekler:
        satir += 1
        sayfa.write_string(satir, 0, "Okunmayan ekler (PDF ve tablo olmayan dosyalar)", bicimler.bolum)
        satir += 1
        sayfa.write_string(
            satir, 0,
            "Bu ekler acilmadi; PDF faturalar kisi kirilimi tasimaz, tutar yanlarindaki Excel'den "
            "okunur. Tutari YALNIZCA PDF'te olan bir fatura varsa bu tabloda yoktur; elle eklenmeli.",
            bicimler.ozet_metin)
        satir += 1
        for ad in atlanan_ekler:
            sayfa.write_string(satir, 0, str(ad), bicimler.ozet_metin)
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
            try:
                mahsup.atlanan_ekler = list((ozet or {}).get("atlanan_ekler") or [])
            except Exception:  # noqa: BLE001 - salt okunur nesne olabilir
                pass
            _mahsuplasma_yaz(calisma, mahsup, bicimler)
        envanter = list((ozet or {}).get("dosya_envanteri") or [])
        if envanter:
            _dosyalar_yaz(calisma, envanter, bicimler)
        if mahsup is not None:
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
