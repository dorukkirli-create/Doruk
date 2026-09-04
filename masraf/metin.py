"""Turkce/Kiril metin normalizasyonu, transliterasyon ve kisi adi temizligi.

Bu modul tamamen deterministiktir; internet veya yapay zeka gerektirmez.
Fatura aciklamalarindan kisi adi cikarmak ve personel ana verisindeki
isimlerle karsilastirilabilir hale getirmek icin kullanilir.
"""

from __future__ import annotations

import re
import unicodedata
from functools import lru_cache
from itertools import product

# --------------------------------------------------------------------------
# Karakter haritalari
# --------------------------------------------------------------------------

# Turkce ozel harfler. NFKD 'i' (noktasiz i) harfini ASCII'ye cevirmedigi icin
# bu harita ZORUNLUDUR; aksi halde 'Kirli' -> 'Krl' gibi bozulmalar olur.
_TURKCE_HARITA: dict[str, str] = {
    "ı": "i", "İ": "I", "ş": "s", "Ş": "S", "ğ": "g", "Ğ": "G",
    "ç": "c", "Ç": "C", "ö": "o", "Ö": "O", "ü": "u", "Ü": "U",
    "â": "a", "Â": "A", "î": "i", "Î": "I", "û": "u", "Û": "U",
}

# Kiril -> Latin. NFKD Kiril harflerini ASCII'ye cevirmez, ayri harita sart.
_KIRIL_HARITA: dict[str, str] = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ё": "e",
    "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k", "л": "l", "м": "m",
    "н": "n", "о": "o", "п": "p", "р": "r", "с": "s", "т": "t", "у": "u",
    "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sch",
    "ъ": "", "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
    "А": "A", "Б": "B", "В": "V", "Г": "G", "Д": "D", "Е": "E", "Ё": "E",
    "Ж": "ZH", "З": "Z", "И": "I", "Й": "Y", "К": "K", "Л": "L", "М": "M",
    "Н": "N", "О": "O", "П": "P", "Р": "R", "С": "S", "Т": "T", "У": "U",
    "Ф": "F", "Х": "H", "Ц": "TS", "Ч": "CH", "Ш": "SH", "Щ": "SCH",
    "Ъ": "", "Ы": "Y", "Ь": "", "Э": "E", "Ю": "YU", "Я": "YA",
    # Ukrayna/Kazak varyantlari (nadir ama veride gorulebilir)
    "і": "i", "І": "I", "ї": "yi", "Ї": "YI", "є": "e", "Є": "E",
    "ґ": "g", "Ґ": "G", "ә": "a", "Ә": "A", "ө": "o", "Ө": "O",
    "ұ": "u", "Ұ": "U", "ү": "u", "Ү": "U", "қ": "k", "Қ": "K",
    "ғ": "g", "Ғ": "G", "ң": "n", "Ң": "N", "һ": "h", "Һ": "H",
}

# Turkce guvenli buyuk harf haritasi.
# Not: Python'un str.upper() metodu 'i' -> 'I' verir. Turkce kuralda 'i' -> 'I'
# (noktali) beklenir; ancak biz sonrasinda ASCII katlama yaptigimiz icin her iki
# yol da ayni yere varir. Yine de belirsizlik birakmamak adina acikca ele
# aliyoruz: noktali/noktasiz tum i varyantlari duz ASCII 'I' olur.
_BUYUK_HARITA: dict[str, str] = {
    "ı": "I", "i": "I", "İ": "I", "I": "I",
    "ş": "S", "Ş": "S", "ğ": "G", "Ğ": "G",
    "ç": "C", "Ç": "C", "ö": "O", "Ö": "O", "ü": "U", "Ü": "U",
    "â": "A", "Â": "A", "î": "I", "Î": "I", "û": "U", "Û": "U",
}


# --------------------------------------------------------------------------
# Temel normalizasyon
# --------------------------------------------------------------------------

def tr_buyuk(s: str) -> str:
    """Turkce guvenli buyuk harf.

    Noktali/noktasiz tum 'i' varyantlarini duz ASCII 'I' harfine cevirir,
    diger Turkce harfleri ASCII karsiligina buyutur, kalan karakterler icin
    standart upper() uygular.

    >>> tr_buyuk("Kirli ismail")
    'KIRLI ISMAIL'
    """
    if not s:
        return ""
    parcalar: list[str] = []
    for ch in s:
        esle = _BUYUK_HARITA.get(ch)
        if esle is not None:
            parcalar.append(esle)
        else:
            parcalar.append(ch.upper())
    return "".join(parcalar)


def ascii_katla(s: str) -> str:
    """Turkce ve Kiril karakterleri ASCII'ye katlar.

    Once acik haritalar (Turkce + Kiril) uygulanir, ardindan kalan aksanli
    Latin harfleri icin unicodedata NFKD ayristirmasi + ASCII suzgeci gecirilir.

    >>> ascii_katla("Çıkış")
    'Cikis'
    >>> ascii_katla("Ыдрак")
    'Ydrak'
    """
    if not s:
        return ""
    parcalar: list[str] = []
    for ch in s:
        esle = _TURKCE_HARITA.get(ch)
        if esle is None:
            esle = _KIRIL_HARITA.get(ch)
        parcalar.append(esle if esle is not None else ch)
    ara = "".join(parcalar)
    ayrisik = unicodedata.normalize("NFKD", ara)
    return ayrisik.encode("ascii", "ignore").decode("ascii")


_SADECE_HARF = re.compile(r"[^A-Z ]+")
_COKLU_BOSLUK = re.compile(r"\s+")


def isim_normalize(s: str) -> str:
    """Isim karsilastirmasi icin kanonik bicim uretir.

    Buyuk harf + ASCII katlama + sadece harf ve bosluk + tek bosluk.

    >>> isim_normalize("Özakay, Mustafa Kemal")
    'OZAKAY MUSTAFA KEMAL'
    """
    if not s:
        return ""
    metin = ascii_katla(tr_buyuk(str(s)))
    # Ayirici isaretleri boslukla degistir ki tokenlar birlesmesin.
    metin = metin.replace("/", " ").replace("\\", " ").replace("-", " ")
    metin = _SADECE_HARF.sub(" ", metin)
    return _COKLU_BOSLUK.sub(" ", metin).strip()


def isim_tokenlari(s: str) -> frozenset[str]:
    """Isimdeki benzersiz tokenlari sirasiz kume olarak dondurur.

    Sirasiz olmasi 'SOYAD AD' ile 'AD SOYAD' yazimlarinin ayni kumeye
    dusmesini saglar.

    >>> isim_tokenlari("Ozakay Mustafa Kemal") == frozenset({"OZAKAY", "MUSTAFA", "KEMAL"})
    True
    """
    norm = isim_normalize(s)
    if not norm:
        return frozenset()
    return frozenset(norm.split(" "))


# --------------------------------------------------------------------------
# Transliterasyon varyantlari
# --------------------------------------------------------------------------

# (kaynak, hedef, sadece_bas) uclulerinden olusan kurallar.
# Sira ONEMLIDIR: once "yabanci yazim -> Turkce yazim" yonu (yuksek degerli),
# sonra ters yon gelir. BFS bu sirayi izledigi icin dogru aday erken uretilir.
_TRANSLIT_KURALLARI: tuple[tuple[str, str, bool], ...] = (
    # Yuksek degerli yon: Rusca/pasaport yazimindan Turkce yazima donus
    ("KH", "H", False),      # GEKHAN -> GEHAN
    ("GE", "GO", True),      # GEKHAN -> GOKHAN (ilk hece)
    ("IY", "YI", False),     # IYLMAZ -> YILMAZ (yer degistirme)
    ("YI", "IY", False),
    ("YR", "IR", False),     # YRMAK -> IRMAK
    ("Y", "I", True),        # bas harf Y <-> I
    ("EI", "EY", False),     # VEISI -> VEYSI
    ("TS", "C", False),
    ("SH", "S", False),
    ("CH", "C", False),
    ("ZH", "J", False),
    ("YU", "U", False),
    ("YA", "A", False),
    ("W", "V", False),
    ("Q", "K", False),
    ("X", "KS", False),
    ("PH", "F", False),
    ("OU", "U", False),
    ("EE", "I", False),
    ("II", "I", False),
    ("IY", "I", False),
    ("YI", "I", False),
    # Ters yon: Turkce yazimdan yabanci yazima
    ("GO", "GE", True),
    ("I", "Y", True),
    ("H", "KH", False),
    ("EY", "EI", False),
    ("IR", "YR", False),
    ("C", "TS", False),
    ("S", "SH", False),
    ("C", "CH", False),
    ("J", "ZH", False),
    ("U", "YU", False),
    ("A", "YA", False),
    ("V", "W", False),
    ("K", "Q", False),
    ("KS", "X", False),
    ("F", "PH", False),
    ("U", "OU", False),
    ("I", "EE", False),
    ("I", "II", False),
    ("I", "IY", False),
    ("I", "YI", False),
)

_TOKEN_VARYANT_SINIRI = 12
_TOPLAM_VARYANT_SINIRI = 64


def _token_varyantlari(token: str, sinir: int = _TOKEN_VARYANT_SINIRI) -> list[str]:
    """Tek bir isim tokeni icin transliterasyon varyantlarini BFS ile uretir.

    Varyantlar kural onceligi sirasindadir; ilk siradaki her zaman tokenin
    kendisidir.
    """
    if not token:
        return []
    gorulen: dict[str, None] = {token: None}
    kuyruk: list[str] = [token]
    bas = 0
    while bas < len(kuyruk) and len(gorulen) < sinir:
        mevcut = kuyruk[bas]
        bas += 1
        for kaynak, hedef, sadece_bas in _TRANSLIT_KURALLARI:
            if sadece_bas:
                if not mevcut.startswith(kaynak):
                    continue
                yeni = hedef + mevcut[len(kaynak):]
            else:
                if kaynak not in mevcut:
                    continue
                yeni = mevcut.replace(kaynak, hedef)
            if not yeni or yeni in gorulen:
                continue
            gorulen[yeni] = None
            kuyruk.append(yeni)
            if len(gorulen) >= sinir:
                break
        # Sondaki cift harf sadelestirmesi (MEHMETT -> MEHMET)
        if len(mevcut) >= 3 and mevcut[-1] == mevcut[-2]:
            yeni = mevcut[:-1]
            if yeni not in gorulen and len(gorulen) < sinir:
                gorulen[yeni] = None
                kuyruk.append(yeni)
    return list(gorulen)


@lru_cache(maxsize=8192)
def translit_varyantlari(s: str) -> set[str]:
    """Bir isim icin olasi alternatif yazimlari uretir.

    Rusca/pasaport transliterasyonundan Turkce yazima geri donus kurallarini
    (KH->H, GE->GO, EI->EY, YR->IR, TS->C, SH->S, CH->C, ZH->J, YU->U, YA->A
    ve tersleri) her token icin ayri ayri uygular, sonra tokenlerin kartezyen
    carpimini alir. Kombinatoryal patlamayi onlemek icin en fazla 64 varyant
    dondurur; normalize edilmis orijinal her zaman kumededir.

    >>> "YILMAZ GOKHAN" in translit_varyantlari("IYLMAZ GEKHAN")
    True
    """
    norm = isim_normalize(s)
    if not norm:
        return set()
    tokenlar = norm.split(" ")
    # Token sayisi arttikca token basina varyant sayisini kis.
    token_siniri = max(2, _TOKEN_VARYANT_SINIRI // max(1, len(tokenlar) - 1))
    listeler = [_token_varyantlari(t, token_siniri) for t in tokenlar]
    sonuc: set[str] = {norm}
    for birlesim in product(*listeler):
        sonuc.add(" ".join(birlesim))
        if len(sonuc) >= _TOPLAM_VARYANT_SINIRI:
            break
    return sonuc


# --------------------------------------------------------------------------
# Bitisik yazilmis adlarin acilmasi
# --------------------------------------------------------------------------

def bitisik_ad_ac(token: str, sozluk: set[str]) -> list[str] | None:
    """Bitisik yazilmis bir adi sozluge gore parcalara ayirir.

    'MUSTAFAKEMAL' -> ['MUSTAFA', 'KEMAL'] gibi. Her parca en az 3 harf
    olmali ve sozlukte bulunmalidir. Once 2 parcali bolme denenir, en uzun
    ilk parcayi veren bolme tercih edilir; bulunamazsa 3 parcali bolme
    denenir. Cozulemezse None doner.
    """
    if not token or not sozluk:
        return None
    t = isim_normalize(token).replace(" ", "")
    n = len(t)
    if n < 6:
        return None
    # 2 parcali bolme - en uzun ilk parca tercih edilir
    en_iyi: list[str] | None = None
    for i in range(3, n - 2):
        sol, sag = t[:i], t[i:]
        if len(sag) < 3:
            continue
        if sol in sozluk and sag in sozluk:
            if en_iyi is None or len(sol) > len(en_iyi[0]):
                en_iyi = [sol, sag]
    if en_iyi is not None:
        return en_iyi
    # 3 parcali bolme
    for i in range(3, n - 5):
        sol = t[:i]
        if sol not in sozluk:
            continue
        for j in range(i + 3, n - 2):
            orta, sag = t[i:j], t[j:]
            if len(sag) < 3:
                continue
            if orta in sozluk and sag in sozluk:
                return [sol, orta, sag]
    return None


# --------------------------------------------------------------------------
# Kisi metni temizligi
# --------------------------------------------------------------------------

# Otel adinin basladigini gosteren anahtar kelimeler. Bu kelimeden itibaren
# metnin geri kalani atilir (kisi adi her zaman otel adindan ONCE gelir).
OTEL_ANAHTARLARI: tuple[str, ...] = (
    "HOTEL", "OTEL", "HOSTEL", "HYATT", "SHERATON", "RADISSON", "PARK INN",
    "PLAZA", "REGENCY", "OCCIDENTAL", "RESORT", "INN", "GRAND", "HILTON",
    "MARRIOTT", "NOVOTEL", "IBIS", "RAMADA", "CROWNE", "MERCURE", "SWISSOTEL",
    "RIXOS", "WYNDHAM", "DOUBLETREE", "COURTYARD", "RENAISSANCE",
    "INTERCONTINENTAL", "KEMPINSKI", "MOVENPICK", "PULLMAN", "SOFITEL",
    "HOLIDAY", "APART", "RESIDENCE", "PALACE", "SUITES", "GUEST HOUSE",
)

# Aciklama metnindeki islem/hizmet kaliplari. Uzundan kisaya siralidir.
GURULTU_KALIPLARI: tuple[str, ...] = (
    "KONAKLAMA YURTDISI", "KONAKLAMA YURTICI", "KONAKLAMA",
    "BILET BEDELI", "BILET IADESI", "BILET IPTALI", "UCAK BILETI", "BILET",
    "RUSYA FEDERASYONU", "TURISTIK", "E VIZE", "EVIZE", "VIZE BEDELI",
    "VIZE ISLEMLERI", "VIZE",
    "EKSTRA BAGAJ UCRETI", "EKSTRA BAGAJ", "BAGAJ UCRETI", "BAGAJ",
    "TARAFINDAN TASINDI", "TARAFINDAN",
    "SARJ EDILECEK", "SARJ",
    "OTEL ISLEMLERI", "BILET ISLEM", "DIGER HIZMETLER",
    "HIZMET BEDELI", "ISLEM UCRETI", "SERVIS UCRETI", "SERVIS BEDELI",
    "REZERVASYON", "DEGISIKLIK UCRETI", "DEGISIKLIK",
    "CENAZE CELENGI", "CELENK", "DAVETIYE", "TRANSFER UCRETI",
    "IADE", "IPTAL", "FARK", "KOMISYON",
    "MR", "MRS", "MS", "MSTR", "CHD", "INF",
    "TK", "PC", "VF", "THY", "PEGASUS", "AJET", "A JET",
    "TURKISH AIRLINES", "AEROFLOT",
)

# Guzergah metinlerinde gecebilen sehir adlari (kod olmayan yazimlar icin).
SEHIR_ADLARI: tuple[str, ...] = (
    "ISTANBUL", "ANKARA", "IZMIR", "ANTALYA", "ADANA", "KAYSERI", "TRABZON",
    "MOSKOVA", "MOSCOW", "PETERSBURG", "SANKT PETERSBURG", "ST PETERSBURG",
    "MURMANSK", "NOVOSIBIRSK", "KAZAN", "SOCI", "SOCHI", "BLAGOVESHCHENSK",
    "AMUR", "USTLUGA", "UST LUGA", "HANOI", "PARIS", "LONDRA", "LONDON",
    "BERLIN", "AMSTERDAM", "DUBAI", "DOHA", "PEKIN", "SEUL", "TOKYO",
    "MILANO", "ROMA", "VIYANA", "BUDAPESTE", "BAKU", "ASTANA", "ALMATI",
)

# Bilet numarasi: iki harfli havayolu kodu + uzun rakam dizisi
_RE_BILET_NO = re.compile(r"\b[A-Z]{2}\d{5,}\b")
# Kose parantezli tarih/aciklama bloklari
_RE_KOSE = re.compile(r"\[[^\]]*\]")
# Parantezli sayilar / notlar
_RE_PARANTEZ = re.compile(r"\([^)]*\)")
# IATA guzergah zinciri: IST-CDG, KYA-SAW-LED
_RE_IATA = re.compile(r"\b[A-Z]{3}(?:\s*-\s*[A-Z]{3})+\b")
# Tarihler: 11.07.2026 / 11/07/2026
_RE_TARIH = re.compile(r"\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b")
# Kesir paylasimlari: 1/3, 2/3
_RE_KESIR = re.compile(r"\b\d+\s*/\s*\d+\b")
# Cinsiyet isareti: \M \F (ters bolu + harf)
_RE_CINSIYET = re.compile(r"\\\s*[MF]\b")
# Kalan uzun rakam dizileri
_RE_RAKAM = re.compile(r"\b\d+\b")

_RE_SEHIR_GUZERGAH = re.compile(
    r"\b(?:" + "|".join(SEHIR_ADLARI) + r")\s*-\s*(?:" + "|".join(SEHIR_ADLARI) + r")\b"
)
_RE_GURULTU = re.compile(
    r"\b(?:" + "|".join(re.escape(k) for k in GURULTU_KALIPLARI) + r")\b"
)
_RE_OTEL = re.compile(
    r"\b(?:" + "|".join(re.escape(k) for k in OTEL_ANAHTARLARI) + r")\b"
)


def kisi_metnini_temizle(s: str) -> str:
    """Fatura aciklamasindan kisi adi disindaki gurultuyu temizler.

    Sirasiyla: cinsiyet isareti, kose parantezli tarihler, parantezli notlar,
    bilet numaralari, tarihler, kesirler, gurultu kaliplari, otel adi (ilk otel
    anahtar kelimesinden itibaren kesilir), guzergah kodlari ve kalan rakamlar
    atilir; sonuc buyuk harf + ASCII olarak dondurulur.

    >>> kisi_metnini_temizle("TK4093099626 OZAKAY/MUSTAFAKEMAL MR  IST-CDG BILET BEDELI")
    'OZAKAY MUSTAFAKEMAL'
    """
    if not s:
        return ""
    metin = ascii_katla(tr_buyuk(str(s)))

    metin = _RE_CINSIYET.sub(" ", metin)
    metin = _RE_KOSE.sub(" ", metin)
    metin = _RE_PARANTEZ.sub(" ", metin)
    metin = _RE_BILET_NO.sub(" ", metin)
    metin = _RE_TARIH.sub(" ", metin)
    metin = _RE_KESIR.sub(" ", metin)

    # Ayiricilari boslukla degistir (SOYAD/AD -> SOYAD AD)
    metin = metin.replace("/", " ").replace("\\", " ")
    metin = re.sub(r"[,;:.]+", " ", metin)
    metin = _COKLU_BOSLUK.sub(" ", metin).strip()

    # Guzergah kodlari (tireli haldeyken yakalanmali)
    metin = _RE_IATA.sub(" ", metin)
    metin = _RE_SEHIR_GUZERGAH.sub(" ", metin)
    metin = metin.replace("-", " ")

    # Gurultu kaliplari
    metin = _RE_GURULTU.sub(" ", metin)
    metin = _COKLU_BOSLUK.sub(" ", metin).strip()

    # Otel adi: ilk otel anahtar kelimesinden itibaren kes (basta ise kesme)
    esle = _RE_OTEL.search(metin)
    if esle is not None and esle.start() > 0:
        metin = metin[: esle.start()]

    metin = _RE_RAKAM.sub(" ", metin)
    metin = _SADECE_HARF.sub(" ", metin)

    # Tek harflik artiklari ve bos tokenlari at
    tokenlar = [t for t in metin.split() if len(t) >= 2]
    return " ".join(tokenlar).strip()


# --------------------------------------------------------------------------
# Rusca disi soyad ekleri
# --------------------------------------------------------------------------

# Rusca soyadlar cinsiyete gore cekimlenir: erkek 'NOVOSELOV', kadin
# 'NOVOSELOVA'. Personel ana verisinde calisan ERKEK ise defterde yalnizca
# erkek hali bulunur; esi bir seyahat faturasinda kadin haliyle gecer ve
# soyad indeksinde karsilik bulamaz. Bu tablo kadin halini erkek haline
# cevirerek aile bagini kurulabilir kilar.
#
# Sira ONEMLIDIR: uzun ekler once denenir, aksi halde 'SKAYA' eki 'AYA'
# kuralina takilir.
_DISI_SOYAD_EKLERI: tuple[tuple[str, str], ...] = (
    ("OVSKAYA", "OVSKIY"),
    ("EVSKAYA", "EVSKIY"),
    ("SKAYA", "SKIY"),
    ("TSKAYA", "TSKIY"),
    ("CKAYA", "CKIY"),
    ("OVA", "OV"),
    ("EVA", "EV"),
    ("YOVA", "YOV"),
    ("INA", "IN"),
    ("YNA", "YN"),
    ("AYA", "IY"),
)

#: Bu ekleri tasiyan bir tokenin soyad sayilabilmesi icin asgari uzunlugu.
#: 'IVANOVA' (7) evet, 'EVA' (3) hayir - 'Eva' bir ADDIR, soyad cekimi degil.
_DISI_SOYAD_ASGARI = 6


@lru_cache(maxsize=4096)
def rus_disi_soyad_erkek_hali(token: str) -> str | None:
    """Rusca kadin soyadinin erkek halini uretir; uygun degilse None.

    Yalnizca TEK bir aday dondurur; dogrulama cagiran tarafa aittir (uretilen
    hal personel soyad indeksinde bulunmuyorsa kullanilmamalidir).

    >>> rus_disi_soyad_erkek_hali("NOVOSELOVA")
    'NOVOSELOV'
    >>> rus_disi_soyad_erkek_hali("SHTEYNIKOVA")
    'SHTEYNIKOV'
    >>> rus_disi_soyad_erkek_hali("MUSTAFA") is None
    True
    """
    t = isim_normalize(token)
    if not t or " " in t or len(t) < _DISI_SOYAD_ASGARI:
        return None
    for son, yeni in _DISI_SOYAD_EKLERI:
        if t.endswith(son) and len(t) - len(son) >= 3:
            return t[: -len(son)] + yeni
    return None
