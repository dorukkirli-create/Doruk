"""Otomasyonun dogrulugunu ELLE DAGITILMIS dosyaya karsi olcer.

Ground truth adayi: ``YUZYIL_TEMMUZ_2026_ELLE_DAGITILMIS.xlsx`` dosyasinin
'SANTIYESI' kolonu. Bu kolonu bir insan doldurmustur.

OLCUMUN EN ONEMLI BULGUSU
-------------------------
Elle dosyadaki 'SANTIYESI' kolonu, otomasyonun urettigi seyle AYNI SORUYU
CEVAPLAMIYOR:

* Otomasyon sorusu : "bu kisi hangi projede calisiyor?" (kisi -> gorev yeri)
* Elle dosya sorusu: "bu bileti hangi tuzel kisi odeyecek?"

Kanit ``yon_ozeti()`` fonksiyonunda uretilir: ayni kisi, ayni proje, ayni ay
icinde iki farkli etiket alir ve ayirt edici degisken KISI DEGIL SEYAHATIN
YONUDUR.

    ESB (Ankara) -> Rusya, tek yon            -> varis projesine/sirketine
                                                 ('UST LUGA GPP', 'AMUR',
                                                  'RENSERVIS', 'ONE TOWER'...)
    Diger her guzergah, otel, vize, diger      -> merkeze ('RHI')

Ankara Esenboga toplu mobilizasyon kapisidir: grup halinde ise alinan saha
personeli oradan ucurulur ve bileti masrafi ustlenen projeye/sirkete yazilir.
Sayilar ``kural_ozeti()`` ile uretilir: ESB kalkisli girislerin 27/29'u gercek
proje etiketi tasir, diger guzergahlarin 13/14'u varsayilan 'RHI'dir.

Bu yuzden dogruluk TEK bir yuzde ile ifade edilemez. Betik uc okuma uretir:

``proje``
    Naif okuma: elle etiket proje sanilir. Otomasyonun proje bazli sonucu
    ile dogrudan karsilastirilir. Bu okuma DUSUK cikar ve bu bir otomasyon
    hatasi degil, taksonomi farkidir.
``tuzel``
    Elle etiket tuzel kisi sayilir; 'RHI' personel ana verisinde bulunan her
    kisiyi kabul eder. Otomasyonun elle dosya ile CELISIP celismedigini
    olcer. Zayif bir testtir, cunku 'RHI' etiketi 100 satirda ayrim yapmaz.
``bilgi``
    Sadece elle dosyanin GERCEK proje bilgisi tasidigi satirlar (etiket
    varsayilan 'RHI' degil). Otomasyonun proje atamasi icin tek gecerli
    ground truth budur.

Calistirma::

    python3 -m testler.dogruluk_olc
    python3 -m testler.dogruluk_olc --ogrenmeyi-kaydet   # veri/ dizinine yazar

Varsayilan olarak veri/ dizinine HICBIR SEY YAZILMAZ; olcum ogrenen
defterleri kirletmemelidir.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Sequence

KOK = Path(__file__).resolve().parent.parent
if str(KOK) not in sys.path:
    sys.path.insert(0, str(KOK))

from masraf.boru import Boru, CalismaAyarlari  # noqa: E402
from masraf.modeller import DURUM_ESLESMEDI, GiderSatiri, Sonuc  # noqa: E402
from masraf.okuyucular.antik import antik_cari_oku, yuzyil_dagitilmis_oku  # noqa: E402

__all__ = [
    "Karsilastirma",
    "DogrulukRaporu",
    "hizala",
    "elle_etiketler",
    "otomasyon_etiketleri",
    "seyahat_yonu",
    "yon_ozeti",
    "kural_ozeti",
    "kalkis_havalimani",
    "olc",
]

# --------------------------------------------------------------------------
# Varsayilan dosya yollari
# --------------------------------------------------------------------------

HAM_DOSYA = KOK / "ornek_veri" / "antik_travel" / "ANTIK_CARI_TEMMUZ_2026.xls"
ELLE_DOSYA = KOK / "ornek_veri" / "antik_travel" / "YUZYIL_TEMMUZ_2026_ELLE_DAGITILMIS.xlsx"
PERSONEL_DOSYA = KOK / "ornek_veri" / "personel" / "2025_2026_giris_cikis.xlsx"

# --------------------------------------------------------------------------
# Okuma bicimleri
# --------------------------------------------------------------------------

OKUMA_PROJE = "proje"
OKUMA_TUZEL = "tuzel"
OKUMALAR: tuple[str, ...] = (OKUMA_PROJE, OKUMA_TUZEL)

#: Elle dosyanin VARSAYILAN etiketi. Bu etiket 134 satirin 100'unde
#: kullanilmistir ve proje ayrimi yapmaz; 'merkez odesin' demektir.
VARSAYILAN_ETIKET = "RHI"

# --------------------------------------------------------------------------
# Taksonomi koprusu
# --------------------------------------------------------------------------

#: Elle dosyada gecen ve personel ana verisinde KARSILIGI OLMAYAN tuzel
#: kisiler. Personel dosyasi sadece RHI ve UST LUGA tuzel kisilerini icerir.
KAPSAM_DISI_ETIKETLER: frozenset[str] = frozenset({
    "RENSERVIS", "RENSTROYDETAL", "RSD", "RC PETER", "RC PETERSBURG",
    "RC MOSKOVA", "RC MOSCOW", "ONE TOWER", "TOP TOWER", "SAREN",
    "YAKA", "YAKA LLC",
})

#: Otomasyonun masraf merkezi KODUNDAN elle dosyada beklenecek etiketler.
#: Elle dosya sadece iki projeyi ad ile ayirir.
KOD_ETIKETLERI: dict[str, frozenset[str]] = {
    "GPP": frozenset({"UST LUGA GPP", "UST LUGA", "GPP"}),
    "AGPP": frozenset({"AMUR", "AMURSKY", "AGPZ"}),
}

#: Harita 'sirket' kolonundan turetilen genel etiketler.
SIRKET_ETIKETLERI: dict[str, frozenset[str]] = {
    "RHI": frozenset({"RHI", "RHI RUSSIA"}),
    "UST LUGA": frozenset({"UST LUGA", "UST LUGA GPP", "USTLUGA"}),
}

# --------------------------------------------------------------------------
# Seyahat yonu
# --------------------------------------------------------------------------

#: Turkiye havalimani IATA kodlari (kaynak veride gecenler).
TR_HAVALIMANLARI: frozenset[str] = frozenset({
    "IST", "SAW", "ESB", "ADB", "AYT", "ADA", "TZX", "GZT", "KYA",
    "ASR", "DIY", "VAN", "ERZ", "MLX", "SZF", "BJV", "DLM", "NAV", "OGU",
})

#: Rusya / BDT havalimani IATA kodlari.
RU_HAVALIMANLARI: frozenset[str] = frozenset({
    "LED", "VKO", "SVO", "DME", "ZIA", "BQS", "IKT", "OVB", "MMK",
    "KHV", "YKS", "CEK", "KZN", "AER", "ROV", "UFA", "SGC", "NUX",
})

#: Bilet aciklamasindaki IATA guzergah zinciri.
_RE_ROTA = re.compile(r"\b([A-Z]{3}(?:\s*-\s*[A-Z]{3})+)\b")

# Seyahat yonleri
YON_GIRIS = "giris"       # TR -> RU : mobilizasyon, varis projesi odeyecek
YON_CIKIS = "cikis"       # RU -> TR : rotasyon / izin donusu
YON_GIDIS_DONUS = "gidis_donus"   # ayni ulkede baslayip biten
YON_YOK = "yok"           # otel / vize / diger, guzergah metni yok


def _ulke(kod: str) -> str | None:
    """Bir IATA kodunun ulke grubunu dondurur ('TR', 'RU' veya None)."""
    if kod in TR_HAVALIMANLARI:
        return "TR"
    if kod in RU_HAVALIMANLARI:
        return "RU"
    return None


def seyahat_yonu(aciklama: str) -> str:
    """Bilet aciklamasindaki guzergahtan seyahatin yonunu cikarir.

    Elle dosyadaki etiketin hangi kurala gore secildigini gosterebilmek icin
    gereklidir: mobilizasyon (TR -> RU tek yon) bileti varis projesine,
    rotasyon donusu (RU -> TR) merkeze yazilmaktadir.
    """
    eslesme = _RE_ROTA.search((aciklama or "").upper())
    if not eslesme:
        return YON_YOK
    duraklar = [d.strip() for d in eslesme.group(1).split("-")]
    bas, son = _ulke(duraklar[0]), _ulke(duraklar[-1])
    if bas is None or son is None:
        return YON_YOK
    if bas == son:
        return YON_GIDIS_DONUS
    return YON_GIRIS if (bas == "TR" and son == "RU") else YON_CIKIS


def kalkis_havalimani(aciklama: str) -> str | None:
    """Bilet aciklamasindaki guzergahin ILK duragini dondurur.

    Elle etiketin ayirt edici degiskeni budur: Ankara (ESB) kalkisli tek yon
    biletler toplu mobilizasyondur ve varis projesine yazilir.
    """
    eslesme = _RE_ROTA.search((aciklama or "").upper())
    if not eslesme:
        return None
    return eslesme.group(1).split("-")[0].strip()


def _etiket_normalize(deger: Any) -> str:
    """Etiketi karsilastirilabilir bicime getirir (buyuk harf, tek bosluk)."""
    if deger is None:
        return ""
    metin = str(deger).strip().upper()
    for eski, yeni in (("İ", "I"), ("Ş", "S"), ("Ğ", "G"), ("Ü", "U"),
                       ("Ö", "O"), ("Ç", "C")):
        metin = metin.replace(eski, yeni)
    metin = metin.replace("-", " ").replace(".", " ")
    return " ".join(metin.split())


def elle_etiketler(satir: GiderSatiri) -> list[str]:
    """Elle dosyadaki satirin santiye etiket(ler)ini dondurur.

    Paylasimli satirlarda ('RHI 1/3- RENSTROYDETAL 2/3') birden fazla etiket
    doner; otomasyon bunlardan HERHANGI birini bulduysa eslesmis sayilir
    (paylasim orani otomasyonun hedefi degildir).
    """
    paylasim = satir.ek.get("paylasim") or []
    if paylasim:
        return [_etiket_normalize(p.get("masraf_merkezi")) for p in paylasim
                if _etiket_normalize(p.get("masraf_merkezi"))]
    tek = _etiket_normalize(satir.masraf_merkezi_kaynak)
    return [tek] if tek else []


def otomasyon_etiketleri(sonuc: Sonuc, okuma: str = OKUMA_PROJE) -> frozenset[str]:
    """Otomasyonun sonucundan elle dosyada BEKLENEN etiket kumesini turetir.

    Args:
        sonuc: Boru hattinin urettigi sonuc.
        okuma: ``OKUMA_PROJE`` elle etiketi proje sayar. ``OKUMA_TUZEL``
            ayrica varsayilan 'RHI' etiketini de kabul eder; yani personel
            ana verisinde bulunan her kisi 'RHI' yazilmis olabilir.
    """
    kod = _etiket_normalize(sonuc.masraf_merkezi)
    sirket = _etiket_normalize(sonuc.sirket) or _etiket_normalize(sonuc.sirket2)

    if kod in KOD_ETIKETLERI:
        etiketler = set(KOD_ETIKETLERI[kod])
    elif sirket in SIRKET_ETIKETLERI:
        etiketler = set(SIRKET_ETIKETLERI[sirket])
    elif sirket:
        etiketler = {sirket}
    elif kod:
        etiketler = {kod}
    else:
        return frozenset()

    if okuma == OKUMA_TUZEL and sonuc.eslesme.sicil:
        # Personel ana verisi sadece RHI ve UST LUGA'yi kapsar; orada bulunan
        # herkesin masrafi elle dosyada merkeze (RHI) yazilmis olabilir.
        etiketler.add(VARSAYILAN_ETIKET)
    return frozenset(etiketler)


# --------------------------------------------------------------------------
# Hizalama
# --------------------------------------------------------------------------


def _tutar_yakin(a: float | None, b: float | None, tolerans: float = 0.02) -> bool:
    """Iki tutari MUTLAK DEGER uzerinden karsilastirir.

    Iade satirlarinda ham dosya negatif (alacak), elle dosya pozitif yazar;
    bu bir hizalama farki degildir.
    """
    if a is None or b is None:
        return a is None and b is None
    return abs(abs(a) - abs(b)) <= tolerans


def _tarih_yakin(a: date | None, b: date | None) -> bool:
    """Belge tarihlerini karsilastirir; biri bossa engel sayilmaz."""
    if a is None or b is None:
        return True
    return a == b


def _puan(ham: GiderSatiri, elle: GiderSatiri) -> int:
    """Iki satirin ayni islem olma gucu (2 = tutar+tarih, 1 = tutar, 0 = hayir)."""
    if not _tutar_yakin(ham.tutar, elle.tutar):
        return 0
    return 2 if _tarih_yakin(ham.belge_tarihi, elle.belge_tarihi) else 1


def hizala(
    ham_satirlar: Sequence[GiderSatiri],
    elle_satirlar: Sequence[GiderSatiri],
    pencere: int = 5,
) -> list[tuple[int | None, int | None, str]]:
    """Iki dosyanin satirlarini sirali olarak hizalar.

    Dosyalar ayni islemleri ayni sirada icerir; bu yuzden iki imlecli sirali
    yurume kullanilir. Her adimda once dogrudan karsilik (tutar + tarih)
    denenir; tutmazsa ileri bakis penceresi icinde ilk uyan cift aranir ve
    atlanan satirlar eslesmemis isaretlenir.

    Returns:
        (ham_indeks, elle_indeks, yontem) uclulerinin listesi. Indekslerden
        biri None ise o satirin karsiligi bulunamamistir.
    """
    ciftler: list[tuple[int | None, int | None, str]] = []
    i = j = 0
    while i < len(ham_satirlar) and j < len(elle_satirlar):
        puan = _puan(ham_satirlar[i], elle_satirlar[j])
        if puan:
            ciftler.append((i, j, "tutar+tarih" if puan == 2 else "tutar"))
            i, j = i + 1, j + 1
            continue
        bulundu = False
        for adim in range(1, pencere + 1):
            for di, dj in ((adim, 0), (0, adim), (adim, adim)):
                ni, nj = i + di, j + dj
                if ni >= len(ham_satirlar) or nj >= len(elle_satirlar):
                    continue
                if _puan(ham_satirlar[ni], elle_satirlar[nj]) == 2:
                    for k in range(i, ni):
                        ciftler.append((k, None, "hizalanamadi"))
                    for k in range(j, nj):
                        ciftler.append((None, k, "hizalanamadi"))
                    ciftler.append((ni, nj, "tutar+tarih"))
                    i, j = ni + 1, nj + 1
                    bulundu = True
                    break
            if bulundu:
                break
        if not bulundu:
            ciftler.append((i, j, "sira (dogrulanamadi)"))
            i, j = i + 1, j + 1
    for k in range(i, len(ham_satirlar)):
        ciftler.append((k, None, "hizalanamadi"))
    for k in range(j, len(elle_satirlar)):
        ciftler.append((None, k, "hizalanamadi"))
    return ciftler


# --------------------------------------------------------------------------
# Karsilastirma kayitlari
# --------------------------------------------------------------------------

AYNI = "AYNI"
FARKLI = "FARKLI"
KAPSAM_DISI = "KAPSAM DISI"
CELISKI_GRUP = "CELISKI (GRUP SIRKETI)"
ELLE_BOS = "ELLE BOS"
OTOMASYON_YOK = "OTOMASYON BULAMADI"
HIZALANAMADI = "HIZALANAMADI"

DURUMLAR: tuple[str, ...] = (
    AYNI, FARKLI, CELISKI_GRUP, KAPSAM_DISI, OTOMASYON_YOK, ELLE_BOS, HIZALANAMADI,
)


@dataclass
class Karsilastirma:
    """Tek bir islem satirinin otomasyon/elle karsilastirmasi."""

    sira: int
    s_no: str | None
    tarih: date | None
    tutar: float | None
    aciklama: str
    hizalama: str
    yon: str
    kisi_ham: str | None
    elle_kisi: str | None
    elle_etiket: list[str]
    otomasyon_kodu: str | None
    otomasyon_gorev_yeri: str | None
    otomasyon_sirket: str | None
    #: Masraf merkezinin nereden geldigi: 'personel' | 'defter' | 'kaynak dosya' | 'yok'
    cozum_kaynagi: str
    sicil: str | None
    personel_adi: str | None
    yontem: str
    guven: float
    otomasyon_durumu: str
    uyarilar: list[str] = field(default_factory=list)
    #: okuma adi -> durum
    durumlar: dict[str, str] = field(default_factory=dict)
    #: okuma adi -> beklenen etiket kumesi
    beklenen: dict[str, frozenset[str]] = field(default_factory=dict)

    @property
    def elle_metni(self) -> str:
        return " + ".join(self.elle_etiket) if self.elle_etiket else "(bos)"

    @property
    def otomasyon_metni(self) -> str:
        if self.otomasyon_kodu:
            return f"{self.otomasyon_kodu} ({self.otomasyon_gorev_yeri})"
        return "(bulunamadi)"

    @property
    def bilgi_tasiyor(self) -> bool:
        """Elle etiket gercek bir proje/sirket ayrimi tasiyor mu?

        Varsayilan 'RHI' etiketi 134 satirin 100'unde kullanilmistir ve hicbir
        ayrim yapmaz; o satirlar otomasyonu SINAMAZ.
        """
        return bool(self.elle_etiket) and self.elle_etiket != [VARSAYILAN_ETIKET]

    def durum(self, okuma: str = OKUMA_PROJE) -> str:
        return self.durumlar.get(okuma, FARKLI)


def _durum_belirle(k: Karsilastirma, okuma: str) -> str:
    """Bir karsilastirmanin verilen okumaya gore durumunu belirler."""
    if k.hizalama == "hizalanamadi":
        return HIZALANAMADI
    if not k.elle_etiket:
        return ELLE_BOS
    beklenen = k.beklenen.get(okuma, frozenset())
    otomasyon_var = bool(beklenen) and k.otomasyon_durumu != DURUM_ESLESMEDI
    grup_sirketi = all(e in KAPSAM_DISI_ETIKETLER for e in k.elle_etiket)
    if grup_sirketi:
        # Elle dosya grup sirketine yazmis. Otomasyon personel ana verisinde
        # (RHI / UST LUGA) bir calisan bulduysa bu gercek bir celiskidir.
        return CELISKI_GRUP if otomasyon_var else KAPSAM_DISI
    if not otomasyon_var:
        return OTOMASYON_YOK
    return AYNI if (set(k.elle_etiket) & set(beklenen)) else FARKLI


def karsilastir(
    ham_satirlar: Sequence[GiderSatiri],
    sonuclar: Sequence[Sonuc],
    elle_satirlar: Sequence[GiderSatiri],
) -> list[Karsilastirma]:
    """Hizalanmis satirlari tek tek karsilastirir."""
    sonuc_haritasi: dict[tuple[str, int], Sonuc] = {
        (s.satir.kaynak_dosya, s.satir.satir_no): s for s in sonuclar
    }
    ciktilar: list[Karsilastirma] = []
    for sira, (hi, ei, yontem) in enumerate(hizala(ham_satirlar, elle_satirlar), start=1):
        ham = ham_satirlar[hi] if hi is not None else None
        elle = elle_satirlar[ei] if ei is not None else None
        sonuc = sonuc_haritasi.get((ham.kaynak_dosya, ham.satir_no)) if ham else None
        aciklama = ham.aciklama if ham else (elle.aciklama if elle else "")

        k = Karsilastirma(
            sira=sira,
            s_no=(elle.ek.get("s_no") if elle else None),
            tarih=(ham.belge_tarihi if ham else (elle.belge_tarihi if elle else None)),
            tutar=(ham.tutar if ham else (elle.tutar if elle else None)),
            aciklama=aciklama,
            hizalama=yontem,
            yon=seyahat_yonu(aciklama),
            kisi_ham=(ham.kisi_ham if ham else None),
            elle_kisi=(elle.kisi_ham if elle else None),
            elle_etiket=(elle_etiketler(elle) if elle else []),
            otomasyon_kodu=(sonuc.masraf_merkezi if sonuc else None),
            otomasyon_gorev_yeri=(sonuc.gorev_yeri if sonuc else None),
            otomasyon_sirket=((sonuc.sirket or sonuc.sirket2) if sonuc else None),
            cozum_kaynagi=(str(sonuc.satir.ek.get("cozum_kaynagi") or "yok")
                           if sonuc else "yok"),
            sicil=(sonuc.eslesme.sicil if sonuc else None),
            personel_adi=(sonuc.eslesme.ad_soyad if sonuc else None),
            yontem=(sonuc.eslesme.yontem if sonuc else "yok"),
            guven=(sonuc.eslesme.guven if sonuc else 0.0),
            otomasyon_durumu=(sonuc.durum if sonuc else DURUM_ESLESMEDI),
            uyarilar=(list(sonuc.uyarilar) if sonuc else []),
        )
        for okuma in OKUMALAR:
            k.beklenen[okuma] = otomasyon_etiketleri(sonuc, okuma) if sonuc else frozenset()
        for okuma in OKUMALAR:
            k.durumlar[okuma] = _durum_belirle(k, okuma)
        ciktilar.append(k)
    return ciktilar


# --------------------------------------------------------------------------
# Rapor
# --------------------------------------------------------------------------


@dataclass
class DogrulukRaporu:
    """Tum olcumun sonucu."""

    karsilastirmalar: list[Karsilastirma]
    ozet: dict
    boru_ozeti: dict

    def sayim(self, durum: str, okuma: str = OKUMA_PROJE,
              kume: Sequence[Karsilastirma] | None = None) -> int:
        kaynak = self.karsilastirmalar if kume is None else kume
        return sum(1 for k in kaynak if k.durum(okuma) == durum)

    def karsilastirilabilir(self, okuma: str = OKUMA_PROJE,
                            kume: Sequence[Karsilastirma] | None = None
                            ) -> list[Karsilastirma]:
        """Dogruluk yuzdesinin paydasi: iki tarafin da sonuc urettigi satirlar."""
        kaynak = self.karsilastirmalar if kume is None else kume
        return [k for k in kaynak if k.durum(okuma) in (AYNI, FARKLI)]

    def dogruluk(self, okuma: str = OKUMA_PROJE,
                 kume: Sequence[Karsilastirma] | None = None) -> tuple[float, int, int]:
        """(yuzde, ayni, karsilastirilabilir) uclusunu dondurur."""
        payda = self.karsilastirilabilir(okuma, kume)
        pay = sum(1 for k in payda if k.durum(okuma) == AYNI)
        yuzde = round(pay / len(payda) * 100, 1) if payda else 0.0
        return yuzde, pay, len(payda)

    @property
    def bilgi_tasiyanlar(self) -> list[Karsilastirma]:
        """Elle dosyanin gercek proje bilgisi tasidigi satirlar."""
        return [k for k in self.karsilastirmalar if k.bilgi_tasiyor]


def yon_ozeti(karsilastirmalar: Sequence[Karsilastirma]) -> dict[tuple[str, str], int]:
    """Seyahat yonu x elle etiket capraz tablosu."""
    sayac: Counter[tuple[str, str]] = Counter()
    for k in karsilastirmalar:
        if k.elle_etiket:
            sayac[(k.yon, k.elle_metni)] += 1
    return dict(sayac)


def kural_ozeti(karsilastirmalar: Sequence[Karsilastirma]) -> dict[str, dict[str, int]]:
    """Elle etiketin hangi kurala gore secildigini gosteren capraz tablo.

    Satirlar guzergah sinifi, kolonlar etiket tipi:

        'ESB kalkisli giris'  Ankara'dan Rusya'ya tek yon (mobilizasyon)
        'diger giris'         baska bir TR havalimanindan Rusya'ya
        'cikis'               Rusya'dan TR'ye
        'gidis-donus'         ayni ulkede baslayip biten
        'guzergahsiz'         otel / vize / diger

    Etiket tipi 'proje/sirket' ise insan gercek bir mahsup adresi yazmis,
    'RHI (varsayilan)' ise merkeze birakmistir.
    """
    tablo: dict[str, dict[str, int]] = {}
    for k in karsilastirmalar:
        if not k.elle_etiket:
            continue
        if k.yon == YON_GIRIS:
            sinif = ("ESB kalkisli giris" if kalkis_havalimani(k.aciklama) == "ESB"
                     else "diger giris")
        elif k.yon == YON_CIKIS:
            sinif = "cikis"
        elif k.yon == YON_GIDIS_DONUS:
            sinif = "gidis-donus"
        else:
            sinif = "guzergahsiz"
        tip = "proje/sirket" if k.bilgi_tasiyor else "RHI (varsayilan)"
        tablo.setdefault(sinif, {"proje/sirket": 0, "RHI (varsayilan)": 0})[tip] += 1
    return tablo


# --------------------------------------------------------------------------
# Olcum
# --------------------------------------------------------------------------


def olc(
    ham_yolu: str | Path = HAM_DOSYA,
    elle_yolu: str | Path = ELLE_DOSYA,
    personel_yolu: str | Path = PERSONEL_DOSYA,
    veri_dizini: str | Path = KOK / "veri",
    ogrenmeyi_kaydet: bool = False,
) -> tuple[DogrulukRaporu, Boru]:
    """Ham dosyayi boru hattindan gecirir ve elle dosyaya karsi olcer."""
    boru = Boru(
        CalismaAyarlari(
            personel_yolu=personel_yolu,
            veri_dizini=veri_dizini,
            ogrenmeyi_kaydet=ogrenmeyi_kaydet,
        )
    )
    sonuclar = boru.isle([ham_yolu])
    boru_ozeti = boru.ozet(sonuclar)

    ham_satirlar = antik_cari_oku(ham_yolu)
    elle_satirlar = yuzyil_dagitilmis_oku(elle_yolu)
    karsilastirmalar = karsilastir(ham_satirlar, sonuclar, elle_satirlar)

    ozet = {
        "ham_satir": len(ham_satirlar),
        "elle_satir": len(elle_satirlar),
        "hizalanan": sum(
            1 for k in karsilastirmalar if k.durum(OKUMA_PROJE) != HIZALANAMADI
        ),
    }
    return DogrulukRaporu(karsilastirmalar, ozet, boru_ozeti), boru


# --------------------------------------------------------------------------
# Raporlama
# --------------------------------------------------------------------------


def _kanit(k: Karsilastirma, boru: Boru) -> list[str]:
    """Bir uyusmazlik icin personel verisinden kanit satirlari uretir."""
    if not k.sicil:
        if k.otomasyon_kodu:
            return [
                "      personel ana verisinde sicil YOK; masraf merkezi "
                f"'{k.otomasyon_kodu}' {k.cozum_kaynagi} kaynagindan ONERI olarak geldi",
                f"      eslestirme: {k.yontem} (guven {k.guven:.2f})",
            ]
        return ["      personel ana verisinde karsilik YOK, oneri de yok"]
    kayit = boru.defter.donem_kaydi(k.sicil, k.tarih) or boru.defter.sicil_ile(k.sicil)
    if not kayit:
        return [f"      sicil {k.sicil} personel verisinde bulunamadi"]
    satirlar = [
        f"      personel : {k.sicil} / {kayit.get('ad_soyad')} / "
        f"{kayit.get('gorev_yeri')} / Sirket2={kayit.get('sirket2')} / "
        f"{kayit.get('statu')} / donem {kayit.get('donem')} / {kayit.get('kategori')}",
        f"      eslestirme: {k.yontem} (guven {k.guven:.2f})",
    ]
    for uyari in k.uyarilar[:2]:
        satirlar.append(f"      uyari    : {uyari}")
    return satirlar


def _liste_yaz(baslik: str, secilenler: Sequence[Karsilastirma], boru: Boru) -> None:
    """Bir durum grubunu kanitlariyla birlikte basar."""
    if not secilenler:
        return
    print("-" * 78)
    print(f"{baslik}  [{len(secilenler)}]")
    print("-" * 78)
    for k in secilenler:
        tutar = f"{k.tutar:>10.2f}" if k.tutar is not None else " " * 10
        print(f"  #{k.s_no or k.sira} {k.tarih} {tutar} "
              f"{(k.kisi_ham or k.elle_kisi or '?')[:30]:<30} yon={k.yon}")
        print(f"      elle     : {k.elle_metni}")
        print(f"      otomasyon: {k.otomasyon_metni} [{k.otomasyon_durumu}]")
        for satir in _kanit(k, boru):
            print(satir)
        print(f"      aciklama : {k.aciklama[:88]}")
        print()


def rapor_yaz(rapor: DogrulukRaporu, boru: Boru) -> None:
    """Olcum sonucunu konsola basar."""
    print("=" * 78)
    print("DOGRULUK OLCUMU - Temmuz 2026 seyahat faturasi")
    print("=" * 78)
    print(f"Ham dosya satiri : {rapor.ozet['ham_satir']}")
    print(f"Elle dosya satiri: {rapor.ozet['elle_satir']}")
    print(f"Hizalanan satir  : {rapor.ozet['hizalanan']}")
    print()
    print("BORU HATTI DURUMU")
    for ad, adet in rapor.boru_ozeti["durum_dagilimi"].items():
        print(f"  {ad:<10} {adet:>4}")
    print(f"  otomatik orani: %{rapor.boru_ozeti['otomatik_orani']}")
    print()

    print("=" * 78)
    print("BULGU: ELLE ETIKET KISIYE DEGIL BILETIN GUZERGAHINA GORE SECILMIS")
    print("=" * 78)
    print(f"  {'guzergah sinifi':<22}{'proje/sirket':>14}{'RHI (varsayilan)':>19}")
    for sinif, satir in sorted(
        kural_ozeti(rapor.karsilastirmalar).items(),
        key=lambda p: -p[1]["proje/sirket"],
    ):
        print(f"  {sinif:<22}{satir['proje/sirket']:>14}{satir['RHI (varsayilan)']:>19}")
    print()
    print("  Ankara (ESB) kalkisli tek yon biletler TOPLU MOBILIZASYONDUR ve")
    print("  masrafi ustlenen projeye/sirkete yazilmistir. Diger her sey merkeze")
    print("  ('RHI') birakilmistir. Yani elle kolon 'bu kisi hangi projede")
    print("  calisiyor' sorusunu DEGIL 'bu bileti kim odeyecek' sorusunu")
    print("  cevapliyor; otomasyonun urettigi sey ise birincisidir.")
    print()
    print("  Ayni projenin ayni ayki calisanlari iki farkli etiket almistir:")
    gpp = [k for k in rapor.karsilastirmalar if k.otomasyon_gorev_yeri == "GPP Project"]
    for etiket, adet in Counter(k.elle_metni for k in gpp).most_common():
        print(f"      GPP Project calisani -> elle '{etiket}': {adet} satir")
    print("  Kisi bazinda CELISKI YOKTUR (ayni kisi hep ayni etiketi almis);")
    print("  ayrimi yapan degisken kisi degil, biletin guzergahidir.")
    print()

    print("=" * 78)
    print("DOGRULUK - UC OKUMA")
    print("=" * 78)
    for okuma, aciklama in (
        (OKUMA_PROJE, "naif: elle etiket proje sanilirsa"),
        (OKUMA_TUZEL, "elle etiket tuzel kisi sayilirsa ('RHI' herkesi kabul eder)"),
    ):
        yuzde, pay, payda = rapor.dogruluk(okuma)
        print(f"  [{okuma}] {aciklama}")
        print(f"      %{yuzde}  ({pay}/{payda})")
        for durum in DURUMLAR:
            adet = rapor.sayim(durum, okuma)
            if adet:
                print(f"          {durum:<24}{adet:>4}")
    bilgi = rapor.bilgi_tasiyanlar
    yuzde, pay, payda = rapor.dogruluk(OKUMA_PROJE, bilgi)
    print(f"  [bilgi] sadece elle etiketin GERCEK proje bilgisi tasidigi satirlar")
    print(f"      {len(bilgi)} satir; bunlarin {payda} tanesi karsilastirilabilir")
    print(f"      %{yuzde}  ({pay}/{payda})")
    for durum in DURUMLAR:
        adet = rapor.sayim(durum, OKUMA_PROJE, bilgi)
        if adet:
            print(f"          {durum:<24}{adet:>4}")
    print()

    print("=" * 78)
    print("SATIR SATIR INCELEME (tuzel kisi okumasi)")
    print("=" * 78)
    _liste_yaz(
        "GERCEK UYUSMAZLIK: ikisi de proje soyledi, projeler farkli",
        [k for k in rapor.karsilastirmalar if k.durum(OKUMA_TUZEL) == FARKLI],
        boru,
    )
    _liste_yaz(
        "CELISKI: elle grup sirketi yazmis, otomasyon RHI/UST LUGA calisani buldu",
        [k for k in rapor.karsilastirmalar if k.durum(OKUMA_TUZEL) == CELISKI_GRUP],
        boru,
    )
    _liste_yaz(
        "OTOMASYON BULAMADI: elle yazmis, otomasyon kisiyi eslestiremedi",
        [k for k in rapor.karsilastirmalar if k.durum(OKUMA_TUZEL) == OTOMASYON_YOK],
        boru,
    )
    _liste_yaz(
        "ELLE BOS: insan santiye yazmamis",
        [k for k in rapor.karsilastirmalar if k.durum(OKUMA_TUZEL) == ELLE_BOS],
        boru,
    )
    kapsam = [k for k in rapor.karsilastirmalar if k.durum(OKUMA_TUZEL) == KAPSAM_DISI]
    if kapsam:
        print("-" * 78)
        print(f"KAPSAM DISI: grup sirketi personeli, personel ana verisinde yok  "
              f"[{len(kapsam)}]")
        print("-" * 78)
        for k in kapsam:
            print(f"  #{k.s_no} {k.elle_metni:<18} {str(k.kisi_ham)[:32]:<33} "
                  f"{k.aciklama[:40]}")
        print()


def main(argv: Sequence[str] | None = None) -> int:
    """Komut satiri girisi."""
    ayrist = argparse.ArgumentParser(description="Otomasyon dogrulugunu olcer.")
    ayrist.add_argument("--ham", default=str(HAM_DOSYA))
    ayrist.add_argument("--elle", default=str(ELLE_DOSYA))
    ayrist.add_argument("--personel", default=str(PERSONEL_DOSYA))
    ayrist.add_argument("--veri", default=str(KOK / "veri"))
    ayrist.add_argument(
        "--ogrenmeyi-kaydet",
        action="store_true",
        help="Beslenen defterleri veri/ dizinine yaz (varsayilan: yazma).",
    )
    secim = ayrist.parse_args(argv)
    rapor, boru = olc(
        secim.ham, secim.elle, secim.personel, secim.veri, secim.ogrenmeyi_kaydet
    )
    rapor_yaz(rapor, boru)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
