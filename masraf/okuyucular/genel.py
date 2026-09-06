"""Genel amacli dosya okuma yardimcilari ve taninmayan dosyalar icin parser.

Bu modul okuyucular paketinin EN ALT katmanidir; paket icindeki hicbir
modulu import etmez (dairesel bagimlilik yok). antik.py, energo.py ve
kesif.py buradaki yardimcilari kullanir.

Iki is yapar:
    1. Excel (.xls / .xlsx / .xlsm) ve CSV dosyalarini tek bir sade
       "sayfa adi -> satir listesi" yapisina indirger (calisma_oku).
    2. Sablonunu tanimadigimiz dosyalar icin baslik satirini ve kolon
       anlamlarini anahtar kelimeyle tahmin eden bir yedek parser sunar
       (genel_oku).
"""

from __future__ import annotations

import csv
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

from masraf.metin import ascii_katla, isim_normalize
from masraf.modeller import GiderSatiri

__all__ = [
    "Calisma",
    "calisma_oku",
    "kolon_anahtari",
    "kolon_haritasi",
    "kolon_ara",
    "sayfa_sec",
    "baslik_satiri_bul",
    "hucre_metni",
    "hucre_sayisi",
    "sayi_coz",
    "hucre_tarihi",
    "sicil_bicimi_mi",
    "kisi_anahtari",
    "sayfa_tekrarlarini_isaretle",
    "tckn_normalize",
    "dolu_hucre_sayisi",
    "genel_oku",
]

# Bos sayilan hucre metinleri.
_BOS_METINLER = frozenset({"", "nan", "nat", "none", "null", "-", "#n/a", "na"})

# Excel seri tarih araligi (1970-01-01 ~ 25569, 2100-01-01 ~ 73051).
_SERI_ALT = 1.0
_SERI_UST = 80000.0

_TARIH_BICIMLERI: tuple[str, ...] = (
    "%d.%m.%Y", "%d/%m/%Y", "%d-%m-%Y",
    "%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d",
    "%d.%m.%y", "%d/%m/%y", "%d-%m-%y",
    "%Y%m%d",
    "%d %B %Y", "%B %d, %Y", "%d %b %Y", "%b %d, %Y",
)
#: Gun/ay olarak cozulemeyen ('07/15/2026': 15. ay yok) metinler icin ay/gun
#: denemesi. Gun <= 12 ise Turkce okuma (gun/ay) her zaman once kazanir.
_AY_GUN_BICIMLERI: tuple[str, ...] = ("%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y")
#: '15.07.2026 14:30', '2026-07-15T00:00:00', '... 14:30:00.000Z' saat kuyrugu.
_SAAT_KUYRUGU = re.compile(
    r"(?:[T ]\s*\d{1,2}:\d{2}(?::\d{2}(?:[.,]\d+)?)?\s*(?:Z|[+-]\d{2}:?\d{2})?)$"
)
#: Bu araliktaki tam sayilar Excel seri numarasi degil YILDIR ('2026' hucresi
#: 1905-07-18 olarak cozuluyordu). Tek basina yil tarih sayilmaz.
_YIL_ALT, _YIL_UST = 1900, 2100

_SADECE_RAKAM = re.compile(r"\D+")
_COKLU_BOSLUK = re.compile(r"\s+")


# --------------------------------------------------------------------------
# Hucre degeri donusturucleri
# --------------------------------------------------------------------------

def hucre_metni(deger: Any) -> str | None:
    """Hucre degerini duz metne cevirir; anlamli icerik yoksa None doner.

    Bolunemez bosluk (\\xa0) ve coklu bosluklar tek boslugu indirgenir.
    """
    if deger is None:
        return None
    if isinstance(deger, str):
        metin = deger
    elif isinstance(deger, (datetime, date)):
        return deger.isoformat()
    elif isinstance(deger, float):
        # 4989.0 gibi degerleri '4989' olarak yaz, ondalikli ise oldugu gibi birak.
        if deger != deger:  # NaN
            return None
        metin = str(int(deger)) if deger.is_integer() else str(deger)
    else:
        metin = str(deger)
    metin = metin.replace("\xa0", " ").replace("​", "")
    metin = _COKLU_BOSLUK.sub(" ", metin).strip()
    if metin.lower() in _BOS_METINLER:
        return None
    return metin


#: Bilimsel gosterim: '1.5e3', '1,23457E+10'.
_BILIMSEL = re.compile(r"^[-+]?\d+(?:[.,]\d+)?[eE][-+]?\d+$")
#: Gecerli binlik gruplamasi: '1.234.567' evet, '15.07.2026' (son grup 4 hane) hayir.
_BINLIK_NOKTA = re.compile(r"^\d{1,3}(?:\.\d{3})+$")
_BINLIK_VIRGUL = re.compile(r"^\d{1,3}(?:,\d{3})+$")

#: sayi_coz'un dondurdugu not: metin hucresindeki tek nokta binlik sayildi.
BINLIK_VARSAYILDI = "binlik_varsayildi"


def sayi_coz(deger: Any) -> tuple[float | None, str | None]:
    """Hucre degerini ondalik sayiya cevirir; ``(sayi, not)`` dondurur.

    Sayisal hucreler oldugu gibi gecer (not None). Metin hucrelerinde Turkce
    tedarikci dosyalarinin yazim aliskanliklari uygulanir:

    * '1.234,56' Turkce (nokta binlik, virgul ondalik); '1,234.56' Ingilizce.
    * Yalniz nokta, birden cok grup: '1.234.567' -> 1234567 (binlik).
    * Tek nokta + tam 3 hane + virgul yok: '2.500' -> 2500. Bu alanda tutarlar
      2 ondalikli yazilir, 3 ondalikli tutar gorulmez; yine de karar
      ``BINLIK_VARSAYILDI`` notuyla dondurulur, okuyucu satirin ek'ine yazar.
      '0.500' ve '12.5' ondaliktir.
    * 'x,yy' virgul ondalik; '36 000,00' bosluk binlik; "1'234.50" kesme binlik.
    * Muhasebe eksisi: '(100)' -> -100, '100-' -> -100.
    * Bilimsel gosterim ('1.5e3') dogru cozulur (1500); 'e' yutulup 1.53
      uretilmez.
    * Gruplari 3'er hane olmayan noktali/virgullu metinler ('15.07.2026',
      '2026-07-15') sayi DEGILDIR: None.
    """
    if deger is None or isinstance(deger, bool):
        return None, None
    if isinstance(deger, (int, float)):
        return (None if deger != deger else float(deger)), None
    metin = hucre_metni(deger)
    if metin is None:
        return None, None
    negatif = False
    if metin.startswith("(") and metin.endswith(")"):
        negatif, metin = True, metin[1:-1]
    # Para birimi, bosluk ve kesme isareti atilir; 'e' yalnizca bilimsel
    # gosterimde kalir ('EUR 12,5' -> '12,5').
    govde = re.sub(r"[^\d,.\-+eE]", "", metin)
    if not _BILIMSEL.match(govde):
        govde = re.sub(r"[eE]", "", govde)
    if not govde or not any(ch.isdigit() for ch in govde):
        return None, None
    if _BILIMSEL.match(govde):
        try:
            sayi = float(govde.replace(",", "."))
        except ValueError:
            return None, None
        return (-sayi if negatif else sayi), None
    # Eksi isareti yalnizca basta ya da sonda olabilir; ortadaki tire tarih
    # ('2026-07-15') ya da aralik belirtir, sayi degildir.
    if govde.startswith("-") or govde.endswith("-"):
        negatif = True
    govde = govde.strip("+-")
    if "-" in govde or "+" in govde:
        return None, None
    not_ = None
    n_nokta, n_virgul = govde.count("."), govde.count(",")
    if n_nokta and n_virgul:
        if govde.rfind(",") > govde.rfind("."):
            tam = govde.replace(".", "").replace(",", ".")  # Turkce
        else:
            tam = govde.replace(",", "")  # Ingilizce
    elif n_nokta > 1:
        if not _BINLIK_NOKTA.match(govde):
            return None, None
        tam = govde.replace(".", "")
    elif n_nokta == 1:
        bas, son = govde.split(".")
        if len(son) == 3 and 1 <= len(bas) <= 3 and bas != "0":
            tam, not_ = bas + son, BINLIK_VARSAYILDI
        else:
            tam = govde
    elif n_virgul > 1:
        if not _BINLIK_VIRGUL.match(govde):
            return None, None
        tam = govde.replace(",", "")
    elif n_virgul == 1:
        tam = govde.replace(",", ".")  # Turkce ondalik
    else:
        tam = govde
    try:
        sayi = float(tam)
    except ValueError:
        return None, None
    return (-sayi if negatif else sayi), not_


def hucre_sayisi(deger: Any) -> float | None:
    """Hucre degerini ondalik sayiya cevirir; cevrilemezse None.

    Kurallar icin ``sayi_coz``'a bakin; bu sarmalayici yalnizca sayiyi
    dondurur (binlik varsayimi notunu atar).
    """
    return sayi_coz(deger)[0]


def hucre_tarihi(deger: Any, datemode: int = 0) -> date | None:
    """Hucre degerini tarihe cevirir; cevrilemezse None.

    datetime/date nesnelerini, Excel seri numaralarini (xlrd datemode ile)
    ve yaygin metin bicimlerini destekler: '22.05.2026', '2026-05-22',
    '15.07.2026 14:30', '2026-07-15T00:00:00', '07/15/2026' (gun 12'den
    buyukse ay/gun), 'July 15, 2026'.

    Tek basina yil ('2026' ya da 2026 sayisi) tarih DEGILDIR; None doner.
    Eski surum bunu Excel seri numarasi sanip 1905'e cozuyordu.
    """
    if deger is None or isinstance(deger, bool):
        return None
    if isinstance(deger, datetime):
        return deger.date()
    if isinstance(deger, date):
        return deger
    if isinstance(deger, (int, float)):
        if deger != deger:
            return None
        sayi = float(deger)
        if sayi.is_integer() and _YIL_ALT <= int(sayi) <= _YIL_UST:
            return None
        if not (_SERI_ALT <= sayi <= _SERI_UST):
            return None
        try:
            import xlrd

            return xlrd.xldate_as_datetime(sayi, datemode).date()
        except Exception:
            return None
    metin = hucre_metni(deger)
    if metin is None:
        return None
    govde = _SAAT_KUYRUGU.sub("", metin).strip()
    for bicim in _TARIH_BICIMLERI:
        try:
            return datetime.strptime(govde, bicim).date()
        except ValueError:
            continue
    for bicim in _AY_GUN_BICIMLERI:
        try:
            return datetime.strptime(govde, bicim).date()
        except ValueError:
            continue
    # Salt sayi iceren metin Excel seri numarasi olabilir ('45000'); yil
    # ('2026') yukaridaki sayi kuraliyla elenir.
    try:
        return hucre_tarihi(float(govde.replace(",", ".")), datemode)
    except ValueError:
        return None


def tckn_normalize(deger: Any) -> str | None:
    """TC Kimlik Numarasini dogrular ve 11 haneli metin olarak dondurur.

    Rakam disi karakterler atilir. Sonuc tam 11 hane degilse veya '0' ile
    basliyorsa None doner (gecersiz TCKN).
    """
    metin = hucre_metni(deger)
    if metin is None:
        return None
    if isinstance(deger, float) and deger.is_integer():
        metin = str(int(deger))
    rakamlar = _SADECE_RAKAM.sub("", metin)
    if len(rakamlar) != 11 or rakamlar[0] == "0":
        return None
    return rakamlar


def dolu_hucre_sayisi(satir: Sequence[Any]) -> int:
    """Satirdaki anlamli (bos olmayan) hucre sayisi."""
    return sum(1 for h in satir if hucre_metni(h) is not None)


def _metinsel_mi(deger: Any) -> bool:
    """Hucre saf metin mi (sayi veya tarih olarak yorumlanamiyor mu)?

    Baslik satirini veri satirlarindan ayirmak icin kullanilir: basliklar
    metin, veri satirlari genellikle sayi/tarih icerir. CSV'de her hucre
    metin olarak geldigi icin '15.07.2026' gibi degerler de veri sayilir.
    """
    if hucre_metni(deger) is None:
        return False
    return hucre_sayisi(deger) is None and hucre_tarihi(deger) is None


# --------------------------------------------------------------------------
# Kolon adi cozumleme
# --------------------------------------------------------------------------

def kolon_anahtari(ad: Any) -> str:
    """Kolon veya sayfa adini karsilastirilabilir ASCII/kucuk bicime cevirir.

    kayit._kolon_anahtari ile ayni davranisi gosterir; bolunemez bosluk,
    satir sonu, alt cizgi ve ayirici isaretler tek boslugu indirgenir.

    >>> kolon_anahtari("ADI SOYADI\\xa0\\xa0")
    'adi soyadi'
    """
    if ad is None:
        return ""
    metin = str(ad).replace("\xa0", " ")
    metin = unicodedata.normalize("NFKC", metin)
    metin = ascii_katla(metin)
    for isaret in ("\n", "\r", "\t", "_", "/", "-", ".", "(", ")", ",", ";", ":"):
        metin = metin.replace(isaret, " ")
    return " ".join(metin.lower().split())


def kolon_haritasi(baslik: Sequence[Any]) -> dict[str, int]:
    """Baslik satirindan 'normalize kolon adi -> indeks' haritasi uretir.

    Ayni ada sahip birden fazla kolon varsa ILK gecen kazanir.
    """
    harita: dict[str, int] = {}
    for i, hucre in enumerate(baslik):
        anahtar = kolon_anahtari(hucre)
        if anahtar and anahtar not in harita:
            harita[anahtar] = i
    return harita


#: Bu uzunlukta ve daha kisa adaylar ('id', 'kod', 'tc', 'no', 'usd', 'pb')
#: 'icerir' modunda yalnizca KELIME olarak aranir. Eskiden alt dizi aranirdi:
#: 'id' -> 'Provider', 'Valid', 'Paid'; 'tc' -> 'Batch No'; 'kod' -> 'Proje
#: Kodu' sicil/TCKN kolonu sayiliyor, degeri bir sicile denk gelen satir
#: bambaska bir personele baglaniyordu (olculdu).
_KISA_ADAY_SINIRI = 3
#: Kisa aday kelimeye Turkce iyelik eki gelmis olabilir: 'Kodu', 'Sicil Nosu'.
_KISA_ADAY_EKLERI = frozenset({"", "u", "i", "su", "si", "nu", "ni", "lari", "leri"})
#: Kisa aday kelime olarak gecse bile bu bilesiklerde baska bir seyi
#: numaralandirir: 'Proje Kodu' proje, 'Invoice ID' fatura. Kisi kimligi degil.
_KISA_ADAY_YASAKLARI: dict[str, tuple[str, ...]] = {
    "kod": (
        "proje", "posta", "masraf", "ulke", "para birimi", "doviz", "vergi", "hesap",
        "urun", "firma", "sirket", "departman", "tedarikci", "musteri", "birim", "is kod",
        "gider", "hizmet", "santiye", "kod adi", "banka", "iban", "swift",
    ),
    "id": (
        "invoice", "fatura", "belge", "kayit", "transaction", "islem", "order", "siparis",
        "booking", "document", "record", "ticket", "bilet", "payment", "odeme", "batch",
    ),
    "no": (
        "invoice", "fatura", "belge", "evrak", "kayit", "islem", "order", "siparis",
        "ticket", "bilet", "batch", "sira", "s no", "telefon", "phone", "hesap", "iban",
        "pasaport", "passport", "police", "policy", "seri", "oda", "room", "ucus", "flight",
    ),
    "tc": ("batch", "etc",),
}


def _kelime_olarak_gecer(aday: str, ad: str) -> bool:
    """Kisa aday, kolon adinda ayri bir kelime (ya da ek almis hali) mi?

    Bilesik yasak listesindeki basliklar ('proje kodu', 'invoice id') kelime
    olarak gecse de eslesmez.

    >>> _kelime_olarak_gecer("id", "personel id")
    True
    >>> _kelime_olarak_gecer("id", "provider")
    False
    >>> _kelime_olarak_gecer("kod", "personel kodu")
    True
    >>> _kelime_olarak_gecer("kod", "proje kodu")
    False
    """
    if any(yasak in ad for yasak in _KISA_ADAY_YASAKLARI.get(aday, ())):
        return False
    for kelime in ad.split():
        if kelime.startswith(aday) and kelime[len(aday):] in _KISA_ADAY_EKLERI:
            return True
    return False


def kolon_ara(
    harita: dict[str, int] | Sequence[Any],
    *adaylar: str,
    icerir: bool = True,
    haric: Sequence[str] = (),
) -> int | None:
    """Aday adlardan biriyle eslesen kolonun indeksini dondurur.

    Once tam eslesme, sonra (icerir=True ise) 'aday, kolon adinin icinde
    geciyor mu' kontrolu yapilir. Adaylar oncelik sirasindadir. Uc ve daha
    az harfli adaylar icerme kontrolunde yalnizca kelime olarak aranir
    ('sicil no', 'personel id', 'tc kimlik' evet; 'provider', 'valid',
    'batch' hayir).

    Args:
        harita: ``kolon_haritasi`` ciktisi. Baslik listesi verilirse harita
            burada uretilir (eski surumde liste verilince AttributeError
            yutuluyor ve her dosya 'kolon bulunamadi' cikiyordu).
        haric: bu parcalari iceren kolon adlari hicbir zaman secilmez
            (orn. tarih ararken 'dogum tarihi' kolonunu elemek icin).
    """
    if not isinstance(harita, dict):
        harita = kolon_haritasi(list(harita))
    yasakli = [kolon_anahtari(h) for h in haric if kolon_anahtari(h)]

    def uygun(ad: str) -> bool:
        return not any(y in ad for y in yasakli)

    for aday in adaylar:
        anahtar = kolon_anahtari(aday)
        if anahtar in harita and uygun(anahtar):
            return harita[anahtar]
    if not icerir:
        return None
    for aday in adaylar:
        anahtar = kolon_anahtari(aday)
        if not anahtar:
            continue
        kisa = len(anahtar) <= _KISA_ADAY_SINIRI
        for ad, i in harita.items():
            if not uygun(ad):
                continue
            if kisa:
                if _kelime_olarak_gecer(anahtar, ad):
                    return i
            elif anahtar in ad:
                return i
    return None


def sayfa_sec(sayfa_adlari: Iterable[str], *adaylar: str) -> str | None:
    """Sayfa adlarindan aday anahtarlarla eslesenin gercek adini dondurur.

    Turkce karakter farklarini yok saymak icin ascii katlanmis karsilastirma
    yapar ('Kisi Listesi' <-> 'Kişi Listesi').
    """
    adlar = list(sayfa_adlari)
    anahtarlar = {ad: kolon_anahtari(ad) for ad in adlar}
    for aday in adaylar:
        hedef = kolon_anahtari(aday)
        for ad in adlar:
            if anahtarlar[ad] == hedef:
                return ad
    for aday in adaylar:
        hedef = kolon_anahtari(aday)
        if not hedef:
            continue
        for ad in adlar:
            if hedef in anahtarlar[ad]:
                return ad
    return None


def baslik_satiri_bul(
    satirlar: Sequence[Sequence[Any]],
    aranan: Sequence[str] = (),
    sinir: int = 15,
) -> int:
    """Baslik satirinin 0 tabanli indeksini bulur.

    'aranan' verilmisse, ilk 'sinir' satir icinde bu anahtarlardan HERHANGI
    birini iceren ilk satir dondurulur (dinamik baslik tespiti). Bulunamazsa
    veya 'aranan' bos ise en cok dolu hucreye sahip ilk satir secilir.
    Hicbir dolu satir yoksa -1 doner.
    """
    ust = min(len(satirlar), sinir)
    if aranan:
        hedefler = {kolon_anahtari(a) for a in aranan if kolon_anahtari(a)}
        for i in range(ust):
            anahtarlar = {kolon_anahtari(h) for h in satirlar[i]}
            if hedefler & anahtarlar:
                return i
    # Puanlama: metin hucreleri baslik lehine, sayi/tarih hucreleri aleyhine.
    # Boylece 'veri satiri basliktan daha dolu' oldugu dosyalarda (orn. Yuzyil
    # dagitilmis) yanlislikla veri satiri baslik sanilmaz.
    en_iyi_i, en_iyi_puan = -1, 0.0
    for i in range(ust):
        satir = satirlar[i]
        metinsel = sum(1 for h in satir if _metinsel_mi(h))
        veri = dolu_hucre_sayisi(satir) - metinsel
        if metinsel < 2:
            continue
        puan = metinsel - 0.5 * veri
        if puan > en_iyi_puan:
            en_iyi_i, en_iyi_puan = i, puan
    return en_iyi_i


# --------------------------------------------------------------------------
# Dosya okuma
# --------------------------------------------------------------------------

@dataclass
class Calisma:
    """Bir Excel/CSV dosyasinin sade temsili.

    sayfalar: sayfa adi -> satir listesi (her satir hucre degerleri listesi).
    datemode: sadece .xls dosyalarinda anlamli (xlrd tarih taban modu).
    """

    yol: Path
    sayfalar: dict[str, list[list[Any]]] = field(default_factory=dict)
    datemode: int = 0

    @property
    def sayfa_adlari(self) -> list[str]:
        return list(self.sayfalar.keys())

    def satirlar(self, sayfa: str | None = None) -> list[list[Any]]:
        """Verilen sayfanin satirlarini dondurur; sayfa None ise ilk sayfa."""
        if not self.sayfalar:
            return []
        if sayfa is None:
            return next(iter(self.sayfalar.values()))
        return self.sayfalar.get(sayfa, [])


def _xls_oku(yol: Path, satir_siniri: int | None) -> Calisma:
    import xlrd

    kitap = xlrd.open_workbook(str(yol), formatting_info=False)
    try:
        sayfalar: dict[str, list[list[Any]]] = {}
        for sayfa in kitap.sheets():
            ust = sayfa.nrows if satir_siniri is None else min(sayfa.nrows, satir_siniri)
            sayfalar[sayfa.name] = [
                [sayfa.cell_value(r, c) for c in range(sayfa.ncols)] for r in range(ust)
            ]
        return Calisma(yol=yol, sayfalar=sayfalar, datemode=kitap.datemode)
    finally:
        kitap.release_resources()


def _xlsx_oku(yol: Path, satir_siniri: int | None) -> Calisma:
    import openpyxl

    kitap = openpyxl.load_workbook(str(yol), data_only=True, read_only=True)
    try:
        sayfalar: dict[str, list[list[Any]]] = {}
        for ad in kitap.sheetnames:
            sayfa = kitap[ad]
            satirlar: list[list[Any]] = []
            for i, satir in enumerate(sayfa.iter_rows(values_only=True)):
                if satir_siniri is not None and i >= satir_siniri:
                    break
                satirlar.append(list(satir))
            sayfalar[ad] = satirlar
        return Calisma(yol=yol, sayfalar=sayfalar, datemode=0)
    finally:
        kitap.close()


def _csv_oku(yol: Path, satir_siniri: int | None) -> Calisma:
    ham: str | None = None
    for kodlama in ("utf-8-sig", "cp1254", "latin-1"):
        try:
            ham = yol.read_text(encoding=kodlama)
            break
        except UnicodeDecodeError:
            continue
    if ham is None:
        ham = yol.read_text(encoding="utf-8", errors="replace")
    ornek = ham[:8192]
    try:
        lehce = csv.Sniffer().sniff(ornek, delimiters=";,\t|")
        ayirici = lehce.delimiter
    except csv.Error:
        ayirici = ";" if ornek.count(";") > ornek.count(",") else ","
    satirlar: list[list[Any]] = []
    for i, satir in enumerate(csv.reader(ham.splitlines(), delimiter=ayirici)):
        if satir_siniri is not None and i >= satir_siniri:
            break
        satirlar.append(list(satir))
    return Calisma(yol=yol, sayfalar={yol.stem: satirlar}, datemode=0)


def calisma_oku(yol: str | Path, satir_siniri: int | None = None) -> Calisma:
    """Excel veya CSV dosyasini sayfa adi -> satir listesi yapisina cevirir.

    satir_siniri verilirse her sayfadan sadece o kadar satir okunur; dosya
    tipi tespiti (kesif modulu) icin hizli on izleme saglar.

    Raises:
        FileNotFoundError: dosya yoksa.
        ValueError: uzanti desteklenmiyorsa.
    """
    p = Path(yol)
    if not p.exists():
        raise FileNotFoundError(f"Dosya bulunamadi: {p}")
    uzanti = p.suffix.lower()
    if uzanti == ".xls":
        return _xls_oku(p, satir_siniri)
    if uzanti in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        return _xlsx_oku(p, satir_siniri)
    if uzanti in {".csv", ".txt", ".tsv"}:
        return _csv_oku(p, satir_siniri)
    raise ValueError(f"Desteklenmeyen dosya uzantisi: {p.suffix} ({p.name})")


# --------------------------------------------------------------------------
# Taninmayan dosyalar icin yedek parser
# --------------------------------------------------------------------------

# Kolon anlami -> aday anahtar kelimeler (oncelik sirasinda).
_ISIM_ADAYLARI = (
    "ad soyad", "adi soyadi", "ad ve soyad", "isim soyisim", "isim",
    "personel adi", "personel", "katilimci", "calisan", "yolcu",
    "full name", "name", "aciklama",
)
_SICIL_ADAYLARI = (
    "sicil", "sicil no", "personel no", "personel numarasi", "id",
    "employee id", "tabel", "kod",
)
_TCKN_ADAYLARI = (
    "tckn", "tc kimlik no", "tc kimlik", "personel t c", "t c kimlik",
    "kimlik no", "tc no", "tc",
)
# 'ucreti' (iyelik: 'Arabulucu Ucreti', 'Hizmet Ucreti') bir HIZMET BEDELIDIR;
# yalin 'Ucret' ve 'Brut/Net Ucret' ise kutuk listelerindeki MAAS kolonudur.
# Maas tutar sayilirsa 20 bin satirlik personel kutugu fatura gibi dagitilir;
# bu yuzden yalin 'ucret' aday DEGILDIR (gerekirse kullanici sozlugune
# eklenir) ve maas bilesikleri _TUTAR_HARIC ile elenir.
_TUTAR_ADAYLARI = (
    "tutar", "satis", "borc", "amount", "toplam", "bedel", "fiyat",
    "energo payi", "usd", "total",
    "hizmet ucreti", "servis ucreti", "islem ucreti", "arabulucu ucreti",
    "arabuluculuk ucreti", "danismanlik ucreti", "egitim ucreti", "vize ucreti",
    "bagaj ucreti", "konaklama ucreti", "ucreti",
)
# Tutar kolonu ararken asla secilmemesi gereken kolonlar: maas/ucret skalasi.
_TUTAR_HARIC = (
    "brut ucret", "net ucret", "aylik ucret", "gunluk ucret", "saat ucreti",
    "saatlik ucret", "maas", "salary", "wage", "ucret skalasi", "baz ucret",
    "temel ucret", "asgari ucret",
)
_TARIH_ADAYLARI = (
    "belge tarihi", "fatura tarihi", "islem tarihi", "kayit tarihi",
    "tarih", "date",
)
_MERKEZ_ADAYLARI = (
    "santiye", "santiyesi", "proje", "masraf merkezi", "masraf yeri",
    "gorev yeri", "cost center", "yansitma", "ilgili sirket", "sirket",
)

# Tarih kolonu ararken asla secilmemesi gereken kolonlar.
_TARIH_HARIC = ("dogum tarihi", "dogum")

#: Isim kolonu bulunamayinca aciklama satirdaki dolu hucrelerden kurulur; bu
#: kolonlar ASLA aciklamaya alinmaz (kimlik, dogum tarihi, telefon, pasaport,
#: IBAN, adres, e-posta). Eski surum hepsini ' | ' ile birlestirip Excel'e
#: (Sonuc/Incele/Eslesmedi) dusuruyordu (olculdu).
_ACIKLAMA_HARIC_BASLIK = re.compile(
    r"(?:^|\s)(?:tc|t c|tckn|tcno|kimlik|dogum|tel|telefon|phone|gsm|cep|mobile|"
    r"pasaport|passport|iban|adres|address|e posta|eposta|email|e mail|mail)(?:\s|$)"
)
#: Telefon numarasi: '+7 921 123 45 67', '0 532 123 45 67', '(212) 555-1234'.
_RE_TELEFON = re.compile(r"^\+?\d[\d\s().-]{7,}\d$")


def _metin_tarih_mi(metin: str) -> bool:
    """Metin bir tarih mi ('01.01.1990', '1990-01-01T00:00:00')? Sayi/seri degil."""
    govde = _SAAT_KUYRUGU.sub("", metin).strip()
    for bicim in _TARIH_BICIMLERI + _AY_GUN_BICIMLERI:
        try:
            datetime.strptime(govde, bicim)
            return True
        except ValueError:
            continue
    return False


def _aciklama_hucreleri(satir: Sequence[Any], ters_harita: dict[int, str]) -> list[str]:
    """Aciklamaya alinabilecek hucre metinleri: kimlik/dogum/telefon disarida.

    Baslik yasakli kolonlar atlanir; basligi ne olursa olsun 11 haneli
    kimlik, tarih ve telefon bicimli degerler de atlanir.
    """
    parcalar: list[str] = []
    for i, hucre in enumerate(satir):
        metin = hucre_metni(hucre)
        if metin is None:
            continue
        baslik = ters_harita.get(i, "")
        if baslik and _ACIKLAMA_HARIC_BASLIK.search(baslik):
            continue
        if tckn_normalize(hucre) is not None:
            continue
        if isinstance(hucre, (datetime, date)) or (isinstance(hucre, str) and _metin_tarih_mi(metin)):
            continue
        if len(_SADECE_RAKAM.sub("", metin)) >= 10 and _RE_TELEFON.match(metin):
            continue
        parcalar.append(metin)
    return parcalar

# Sicil kolonu ararken asla secilmemesi gereken kolonlar: 'kod' / 'id' / 'no'
# kelimesi tasiyan ama kisiyi degil baska bir seyi numaralandiran basliklar.
_SICIL_HARIC = (
    "proje kodu", "proje kod", "posta kodu", "masraf merkezi kod", "masraf yeri kod",
    "ulke kodu", "para birimi kod", "doviz kod", "vergi kod", "hesap kod",
    "urun kodu", "firma kod", "sirket kod", "departman kod", "tedarikci kod",
    "musteri kod", "birim kod", "is kodu", "gider kod", "masraf kod", "hizmet kod",
    "santiye kod", "kod adi", "project code", "postal code", "zip code",
    "cost center code", "cost centre code", "currency code", "country code",
    "tax code", "account code", "product code", "vendor code", "customer code",
    "service code", "invoice id", "invoice no", "fatura id", "fatura no", "belge id",
    "belge no", "kayit id", "kayit no", "transaction id", "islem id", "islem no",
    "order id", "siparis id", "siparis no", "booking id", "document id", "document no",
    "record id", "ticket no", "bilet no", "pnr", "batch",
)
# Kisi kolonu ararken asla secilmemesi gereken kolonlar ('name' icerir ama
# kisi adi degildir).
_ISIM_HARIC = (
    "company name", "firma adi", "sirket adi", "hotel name", "otel adi",
    "proje adi", "project name", "dosya adi", "file name", "bank name", "banka adi",
)

#: Sicil degeri bicimi: 2-9 haneli tam sayi ('632481') ya da kisa alfasayisal
#: kod ('RHI-1234', 'A1234'). 11 haneli TCKN, tarih, ondalikli tutar ve serbest
#: metin sicil degildir.
_SICIL_BICIMI = re.compile(r"^(?:\d{2,9}|[A-Z]{1,4}[-/ ]?\d{2,8}|\d{2,8}[-/ ]?[A-Z]{1,3})$")
#: Secilen sicil kolonunda dolu degerlerin en az bu orani sicil bicimine
#: uymali; uymuyorsa kolon sicil sayilmaz.
_SICIL_KOLON_ESIGI = 0.6
_SICIL_ORNEK_SINIRI = 300


def sicil_bicimi_mi(deger: Any) -> bool:
    """Deger bir sicil numarasina benziyor mu? ('632481' evet, '15.07.2026' hayir)"""
    metin = hucre_metni(deger)
    if metin is None:
        return False
    if isinstance(deger, float) and deger.is_integer():
        metin = str(int(deger))
    elif metin.endswith(".0") and metin[:-2].isdigit():
        metin = metin[:-2]
    return bool(_SICIL_BICIMI.match(ascii_katla(metin).upper()))


def _sicil_kolonu_uygun_mu(satirlar: Sequence[Sequence[Any]], baslik_i: int, i_sicil: int) -> bool:
    """Secilen sicil kolonunun degerleri cogunlukla sicil bicimine uyuyor mu?

    'Proje Kodu' ya da 'Kodu' gibi basliklar sicil sanilabilir; degerler
    ('GPP', 'SK-01') sicil degilse kolon reddedilir. Dolu deger yoksa
    yargilanamaz, kolon korunur.
    """
    dolu = uygun = 0
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        if i_sicil >= len(satir):
            continue
        deger = satir[i_sicil]
        if hucre_metni(deger) is None:
            continue
        dolu += 1
        if sicil_bicimi_mi(deger):
            uygun += 1
        if dolu >= _SICIL_ORNEK_SINIRI:
            break
    if dolu == 0:
        return True
    return uygun / dolu >= _SICIL_KOLON_ESIGI


#: Pivot / ozet tablosu isaretleri. Baslik satirinda ya da ustundeki satirlarda
#: bunlardan biri gecen sayfa kisi listesi DEGIL, ozet tablosudur: 'Count of
#: PERSONEL' kolonu isim sanilip 'Grand Total' ve proje adlari kisi olarak
#: uretiliyordu (olculdu: 34 sahte kisi). Kiril basliklar ascii katlanmis
#: haliyle ('Названия строк' -> 'nazvaniya strok') listelenir.
_PIVOT_ISARETLERI: tuple[str, ...] = (
    "count of", "sum of", "average of", "min of", "max of",
    "row labels", "column labels", "satir etiketleri", "sutun etiketleri",
    "nazvaniya strok", "nazvaniya stolbtsov", "kolichestvo po polyu", "summa po polyu",
)
#: Yalnizca baslik satirinda kolon adi olarak gecerse pivot sayilan isaretler.
_PIVOT_BASLIK_ISARETLERI: frozenset[str] = frozenset({
    "grand total", "genel toplam", "obschiy itog", "toplam", "total",
})


def _pivot_sayfasi_mi(satirlar: Sequence[Sequence[Any]], baslik_i: int) -> bool:
    """Sayfa bir pivot/ozet tablosu mu?"""
    for r in range(0, min(baslik_i + 1, len(satirlar))):
        for hucre in satirlar[r]:
            anahtar = kolon_anahtari(hucre)
            if not anahtar:
                continue
            if any(isaret in anahtar for isaret in _PIVOT_ISARETLERI):
                return True
            if r == baslik_i and anahtar in _PIVOT_BASLIK_ISARETLERI:
                return True
    return False


#: Bir sayfanin kisileri baska bir sayfadakilerin bu oraninda tekrariysa
#: 'tekrar eden sayfa' sayilir.
_SAYFA_TEKRAR_ESIGI = 0.8
_SAYFA_TEKRAR_ASGARI = 10


def kisi_anahtari(satir: GiderSatiri) -> str | None:
    """Satirdaki kisinin sayim anahtari: sicil, yoksa TCKN, yoksa normalize isim."""
    if satir.sicil_ham:
        return "S:" + str(satir.sicil_ham)
    if satir.tckn_ham:
        return "T:" + str(satir.tckn_ham)
    if satir.kisi_ham:
        norm = isim_normalize(satir.kisi_ham)
        return ("I:" + norm) if norm else None
    return None


def sayfa_tekrarlarini_isaretle(satirlar: list[GiderSatiri]) -> dict[str, str]:
    """Ayni calisma kitabinda ayni kisileri tekrar eden sayfalari isaretler.

    Kisi anahtari sicil, yoksa TCKN, yoksa normalize isimdir. Kisileri en
    kalabalik sayfa 'asil' sayilir; kisilerinin %80'i asil bir sayfada da
    bulunan daha kucuk sayfa 'tekrar' sayilir ve satirlarinin ek'ine
    ``sayfa_tekrari = <asil sayfa adi>`` yazilir. Sayim/besleme yapan katman
    bu notla kisileri iki kez saymaz.

    Returns:
        {tekrar eden sayfa: asil sayfa}
    """
    kumeler: dict[str, set[str]] = {}
    sira: dict[str, int] = {}
    for s in satirlar:
        sayfa = str((s.ek or {}).get("sayfa") or "")
        anahtar = kisi_anahtari(s)
        if not sayfa or anahtar is None:
            continue
        kumeler.setdefault(sayfa, set()).add(anahtar)
        sira.setdefault(sayfa, len(sira))
    sirali = sorted(kumeler, key=lambda ad: (-len(kumeler[ad]), sira[ad]))
    asillar: list[str] = []
    tekrarlar: dict[str, str] = {}
    for ad in sirali:
        kume = kumeler[ad]
        asil = None
        if len(kume) >= _SAYFA_TEKRAR_ASGARI:
            for aday in asillar:
                if len(kume & kumeler[aday]) / len(kume) >= _SAYFA_TEKRAR_ESIGI:
                    asil = aday
                    break
        if asil is None:
            asillar.append(ad)
        else:
            tekrarlar[ad] = asil
    if tekrarlar:
        for s in satirlar:
            sayfa = str((s.ek or {}).get("sayfa") or "")
            if sayfa in tekrarlar and isinstance(s.ek, dict):
                s.ek["sayfa_tekrari"] = tekrarlar[sayfa]
    return tekrarlar

#: Baslik satiri bu kadar satir icinde aranir. Kesif (dosya tipi tespiti) de
#: AYNI siniri kullanir; iki sinir ayrisirsa dosya 'genel' tanilip 0 satir
#: doner (olculdu: baslik 12. satirdayken).
BASLIK_ARAMA_SINIRI = 15

# Ozet satiri tespiti. 'GENEL', 'IADE', 'OZET' Turkiye'de gercek soyadlaridir;
# bu yuzden onek olarak DEGIL, tam kelime/ifade olarak aranir. 'TOPLAM ...'
# ve 'TOTAL ...' ile baslayan hucre ise her zaman ozettir.
_OZET_TAM: frozenset[str] = frozenset({
    "toplam", "genel toplam", "ara toplam", "odenecek", "odenecek tutar", "iade",
    "total", "grand total", "sub total", "subtotal", "satir etiketleri", "ozet",
    "genel", "toplam tutar", "net toplam", "kdv", "kdv dahil", "kdv haric",
})
_OZET_ONEKLERI: tuple[str, ...] = ("genel toplam", "ara toplam", "grand total")
#: 'TOPLAM USD', 'TOTAL AMOUNT' gibi: onek + ozet kelimesi. 'TOTAL MEHMET' degil.
_OZET_KUYRUKLARI: frozenset[str] = frozenset({
    "usd", "eur", "rub", "try", "tl", "tutar", "amount", "sum", "kdv", "dahil", "haric",
    "net", "brut", "genel", "ara", "bakiye", "borc", "alacak", "fatura", "invoice",
})

# Para birimi kolonu adaylari. Bulunmazsa tutar basligindaki doviz koduna bakilir.
_DOVIZ_ADAYLARI: tuple[str, ...] = (
    "para birimi", "doviz", "doviz cinsi", "currency", "ccy", "cur", "pb", "валюта",
)
_DOVIZ_KODLARI: dict[str, str] = {
    "usd": "USD", "$": "USD", "dolar": "USD", "eur": "EUR", "euro": "EUR", "€": "EUR",
    "rub": "RUB", "rur": "RUB", "ruble": "RUB", "руб": "RUB", "try": "TRY", "tl": "TRY",
    "₺": "TRY", "gbp": "GBP", "kzt": "KZT", "cny": "CNY",
}


def _ozet_satiri_mi(kimlik: str) -> bool:
    k = kimlik.strip()
    if k in _OZET_TAM or any(k.startswith(o) for o in _OZET_ONEKLERI):
        return True
    parcalar = k.split()
    if len(parcalar) >= 2 and parcalar[0] in ("toplam", "total"):
        # 'TOPLAM USD' ozet, 'TOTAL MEHMET' kisi: kalan kelimeler ozet
        # kelimesiyse ozettir.
        return all(pk in _OZET_KUYRUKLARI for pk in parcalar[1:])
    return False


def _doviz_coz(deger: Any, tutar_basligi: str | None = None) -> str | None:
    """Hucredeki ya da tutar basligindaki doviz kodunu ISO'ya cevirir."""
    for aday in (deger, tutar_basligi):
        if aday is None:
            continue
        metin = kolon_anahtari(str(aday))
        if not metin:
            continue
        if metin.upper() in ("USD", "EUR", "RUB", "TRY", "GBP", "KZT", "CNY"):
            return metin.upper()
        for parca in metin.replace("(", " ").replace(")", " ").split():
            if parca in _DOVIZ_KODLARI:
                return _DOVIZ_KODLARI[parca]
    return None

# Dosya adi / aciklama anahtar kelimesinden gider tipi tahmini.
_TIP_IPUCLARI: tuple[tuple[str, str], ...] = (
    ("konaklama", "Otel"),
    ("otel", "Otel"),
    ("hotel", "Otel"),
    ("bilet", "Bilet"),
    ("ucus", "Bilet"),
    ("vize", "Vize"),
    ("bagaj", "Bagaj"),
    ("arabulucu", "Arabuluculuk"),
    ("saglik", "Saglik"),
    ("egitim", "Egitim"),
    ("assessment", "Egitim"),
    ("katilimci", "Egitim"),
)


def _gider_tipi_tahmin(*metinler: str | None) -> str:
    """Dosya adi ve aciklamadan gider tipini tahmin eder; bulunamazsa 'Diger'."""
    birlesik = kolon_anahtari(" ".join(m for m in metinler if m))
    for anahtar, tip in _TIP_IPUCLARI:
        if anahtar in birlesik:
            return tip
    return "Diger"


def genel_oku(yol: str | Path, notlar: list[str] | None = None) -> list[GiderSatiri]:
    """Sablonu taninmayan Excel/CSV dosyasini en iyi cabayla ayristirir.

    Baslik satirini ilk 15 satir icinde en cok dolu hucreye sahip satir
    olarak belirler, ardindan kolon adlarindan isim / sicil / tckn / tutar /
    tarih / masraf merkezi kolonlarini anahtar kelimeyle tahmin eder.
    Bulunamayan alanlar None birakilir.

    Tum sayfalar taranir; her sayfa icin ayri baslik/kolon cozumu yapilir.
    Pivot/ozet sayfalari ('Count of', 'Row Labels', 'Grand Total') atlanir;
    ayni kisileri tekrar eden sayfalarin satirlari ``ek['sayfa_tekrari']``
    ile isaretlenir. Secilen sicil kolonunun degerleri sicil bicimine
    uymuyorsa kolon kullanilmaz (``ek['sicil_kolonu_reddedildi']``).

    Args:
        notlar: verilirse atlanan sayfalar ve reddedilen kolonlar hakkinda
            kullaniciya gosterilebilecek Turkce notlar buraya eklenir.
    """
    p = Path(yol)
    calisma = calisma_oku(p)
    sonuclar: list[GiderSatiri] = []
    if notlar is None:
        notlar = []

    for sayfa_adi, satirlar in calisma.sayfalar.items():
        baslik_i = baslik_satiri_bul(satirlar, sinir=BASLIK_ARAMA_SINIRI)
        if baslik_i < 0:
            continue
        if _pivot_sayfasi_mi(satirlar, baslik_i):
            notlar.append(f"'{sayfa_adi}' sayfasi pivot/ozet tablosu; kisi listesi olarak okunmadi.")
            continue
        harita = kolon_haritasi(satirlar[baslik_i])
        if not harita:
            continue

        # Kullanici sozlugu (veri/kolon_esanlamlilari.csv) yerlesik adaylarin
        # ONUNE eklenir. Boylece yeni bir tedarikci sablonu geldiginde kod
        # degistirmeden, tek satir CSV ile taninabilir hale gelir.
        try:
            from masraf.kolon_sozlugu import genislet as _genislet
        except Exception:  # noqa: BLE001
            def _genislet(alan, varsayilanlar, veri_dizini="veri"):
                return tuple(varsayilanlar)

        i_isim = kolon_ara(harita, *_genislet("kisi", _ISIM_ADAYLARI), haric=_ISIM_HARIC)
        i_sicil = kolon_ara(harita, *_genislet("sicil", _SICIL_ADAYLARI), haric=_SICIL_HARIC)
        i_tckn = kolon_ara(harita, *_genislet("tckn", _TCKN_ADAYLARI))
        i_tutar = kolon_ara(harita, *_genislet("tutar", _TUTAR_ADAYLARI), haric=_TUTAR_HARIC)
        i_tarih = kolon_ara(harita, *_genislet("tarih", _TARIH_ADAYLARI), haric=_TARIH_HARIC)
        i_merkez = kolon_ara(harita, *_genislet("santiye", _MERKEZ_ADAYLARI))
        i_doviz = kolon_ara(harita, *_genislet("doviz", _DOVIZ_ADAYLARI))
        if i_doviz is not None and i_doviz in (i_tutar, i_isim):
            i_doviz = None
        # Tutar basligi 'Tutar (USD)' gibi doviz tasiyorsa satirlara o yazilir.
        tutar_basligi = None
        if i_tutar is not None:
            for ham_ad, indeks in harita.items():
                if indeks == i_tutar:
                    tutar_basligi = ham_ad
                    break
        # Iki ayri tutar kolonu (RUB ve USD gibi) varsa ilki secilir; bunu
        # sessizce yapmak yanlis: satir ek'ine not dusulur, ozet uyari uretir.
        tutar_adaylari = [
            ad for ad in harita
            if kolon_ara({ad: harita[ad]}, *_genislet("tutar", _TUTAR_ADAYLARI), haric=_TUTAR_HARIC) is not None
        ]

        # Ayni kolon hem isim hem masraf merkezi olarak secilmesin.
        if i_merkez is not None and i_merkez == i_isim:
            i_merkez = None
        if i_sicil is not None and i_sicil == i_tckn:
            i_sicil = None

        # Basligi sicil gibi gorunen kolonun degerleri sicil degilse ('Kodu'
        # altinda 'SK-01') kolon kullanilmaz; aksi halde deger tesadufen bir
        # sicile denk gelen satir yanlis personele baglanir.
        sicil_reddedilen = None
        if i_sicil is not None and not _sicil_kolonu_uygun_mu(satirlar, baslik_i, i_sicil):
            sicil_reddedilen = next((ad for ad, k in harita.items() if k == i_sicil), None)
            notlar.append(
                f"'{sayfa_adi}' sayfasinda '{sicil_reddedilen}' kolonu sicil gibi adlandirilmis "
                "ama degerleri sicil bicimine uymuyor; sicil olarak kullanilmadi."
            )
            i_sicil = None

        if i_isim is None and i_sicil is None and i_tckn is None:
            continue  # kisi bilgisi yok, bu sayfa gider satiri uretmez

        from masraf.kayit import sicil_normalize

        ters_harita = {i: ad for ad, i in harita.items()}
        for r in range(baslik_i + 1, len(satirlar)):
            satir = satirlar[r]
            if dolu_hucre_sayisi(satir) == 0:
                continue

            def al(i: int | None) -> Any:
                if i is None or i >= len(satir):
                    return None
                return satir[i]

            isim = hucre_metni(al(i_isim))
            sicil = sicil_normalize(al(i_sicil)) or None
            tckn = tckn_normalize(al(i_tckn))
            if isim is None and sicil is None and tckn is None:
                continue  # bos satir
            if tckn is None:
                # TOPLAM / ODENECEK gibi ozet satirlarini ele. Tam kelime
                # aranir: 'GENEL AHMET' bir kisidir, 'GENEL TOPLAM' degildir.
                # TCKN varsa satir kesin bir kisiye aittir, elenmez.
                kimlikler = [kolon_anahtari(k) for k in (isim, sicil) if k]
                if any(_ozet_satiri_mi(k) for k in kimlikler):
                    continue

            if isim:
                aciklama = isim
            else:
                # Kimlik, dogum tarihi ve telefon aciklamaya girmez; geriye bir
                # sey kalmazsa kimlik maskelenmis yazilir (eslestirici uslubu).
                aciklama = " | ".join(_aciklama_hucreleri(satir, ters_harita))
                if not aciklama:
                    aciklama = f"TCKN {tckn[:3]}******" if tckn else (f"Sicil {sicil}" if sicil else "")
            tutar, tutar_notu = sayi_coz(al(i_tutar))
            sonuclar.append(
                GiderSatiri(
                    kaynak_dosya=p.name,
                    kaynak_tip="genel",
                    satir_no=r + 1,
                    belge_tarihi=hucre_tarihi(al(i_tarih), calisma.datemode),
                    aciklama=aciklama,
                    kisi_ham=isim,
                    sicil_ham=sicil,
                    tckn_ham=tckn,
                    tutar=tutar,
                    para_birimi=_doviz_coz(al(i_doviz), tutar_basligi),
                    masraf_merkezi_kaynak=hucre_metni(al(i_merkez)),
                    gider_tipi=_gider_tipi_tahmin(p.name, sayfa_adi),
                    ek={
                        "sayfa": sayfa_adi,
                        "baslik_satiri": baslik_i + 1,
                        "cozulen_kolonlar": {
                            "isim": i_isim, "sicil": i_sicil, "tckn": i_tckn,
                            "tutar": i_tutar, "tarih": i_tarih, "merkez": i_merkez,
                            "doviz": i_doviz,
                        },
                        **({"tutar_kolonu_secenekleri": tutar_adaylari}
                           if len(tutar_adaylari) > 1 else {}),
                        **({tutar_notu: True} if tutar_notu else {}),
                        **({"sicil_kolonu_reddedildi": sicil_reddedilen}
                           if sicil_reddedilen else {}),
                    },
                )
            )
    tekrarlar = sayfa_tekrarlarini_isaretle(sonuclar)
    for tekrar, asil in tekrarlar.items():
        notlar.append(f"'{tekrar}' sayfasi '{asil}' sayfasindaki kisileri tekrar ediyor.")
    return sonuclar
