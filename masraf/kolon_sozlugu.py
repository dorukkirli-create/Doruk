"""Kullanicinin genisletebilecegi kolon adi sozlugu.

Neden gerekli: her ay yeni bir tedarikci veya degismis bir sablon gelebilir.
Genel okuyucu kolon adlarini anahtar kelimeyle tahmin eder ve tanidik adlarin
cogunu (Turkce ve Ingilizce) zaten bulur. Ama hic gorulmemis bir ad geldiginde
sessizce bos doner.

Bu modul o bosluk icin bir CSV sozlugu tutar: ``veri/kolon_esanlamlilari.csv``.
Kullanici yeni bir kolon adini bir kez yazar, sonraki tum dosyalarda calisir.
Kod degistirmek gerekmez.

Dosya bicimi (utf-8-sig, noktali virgul ayirici, Excel ile acilabilir)::

    alan;kolon_adi;not
    kisi;Beneficiary;Yeni tedarikci X boyle yaziyor
    tarih;Dt;
    tutar;Val;
    santiye;Site;

Gecerli ``alan`` degerleri ``ALANLAR`` icinde listelidir. Taninmayan alan adi
olan satirlar sessizce atlanir; boylece kullanicinin yazim hatasi is akisini
durdurmaz.

Ayrica ``tanilama()`` bir dosyanin basliklarini okuyup hangi alanlarin
bulundugunu ve hangilerinin eksik oldugunu soyler. Arayuz bunu "yeni format"
ekraninda gosterir.
"""

from __future__ import annotations

import csv
import logging
from datetime import date
from pathlib import Path
from typing import Iterable

_log = logging.getLogger(__name__)

#: Sozlukte tanimlanabilecek alanlar ve okunakli adlari.
ALANLAR: dict[str, str] = {
    "kisi": "Kisi adi",
    "sicil": "Sicil numarasi",
    "tckn": "TC kimlik numarasi",
    "tutar": "Tutar",
    "tarih": "Belge tarihi",
    "santiye": "Santiye / proje",
}

#: Bir dosyanin islenebilmesi icin en azindan bunlardan biri bulunmalidir.
ZORUNLU_ALANLAR: frozenset[str] = frozenset({"kisi", "sicil", "tckn"})

DOSYA_ADI = "kolon_esanlamlilari.csv"
BASLIKLAR = ("alan", "kolon_adi", "not")
AYIRICI = ";"

#: Dosya yoksa olusturulacak ornek satirlar. Kullaniciya bicimi gosterir.
ORNEK_SATIRLAR: tuple[tuple[str, str, str], ...] = (
    ("kisi", "Beneficiary", "ornek: yeni bir tedarikci kisi kolonuna boyle diyor"),
    ("tarih", "Dt", "ornek satir, silebilirsiniz"),
    ("tutar", "Val", "ornek satir, silebilirsiniz"),
)


def varsayilan_yol(veri_dizini: str | Path = "veri") -> Path:
    """Sozluk dosyasinin varsayilan yolu."""
    return Path(veri_dizini) / DOSYA_ADI


def _anahtar(deger) -> str:
    """Kolon adini karsilastirilabilir bicime cevirir (genel.kolon_anahtari ile ayni)."""
    from masraf.okuyucular.genel import kolon_anahtari

    return kolon_anahtari(deger)


def olustur(yol: str | Path) -> Path:
    """Sozluk dosyasini basliklari ve ornek satirlariyla olusturur."""
    hedef = Path(yol)
    hedef.parent.mkdir(parents=True, exist_ok=True)
    with open(hedef, "w", encoding="utf-8-sig", newline="") as f:
        yazici = csv.writer(f, delimiter=AYIRICI)
        yazici.writerow(BASLIKLAR)
        for satir in ORNEK_SATIRLAR:
            yazici.writerow(satir)
    return hedef


def yukle(veri_dizini: str | Path = "veri", olustur_yoksa: bool = True) -> dict[str, list[str]]:
    """Sozlugu okur ve ``{alan: [kolon adlari]}`` dondurur.

    Dosya yoksa ve ``olustur_yoksa`` True ise ornekleriyle olusturulur.
    Bozuk satirlar atlanir; okuma hicbir durumda istisna firlatmaz.
    """
    yol = varsayilan_yol(veri_dizini)
    if not yol.is_file():
        if olustur_yoksa:
            try:
                olustur(yol)
            except OSError as e:
                _log.warning("Kolon sozlugu olusturulamadi: %s", e)
        return {}

    sonuc: dict[str, list[str]] = {}
    try:
        with open(yol, encoding="utf-8-sig", newline="") as f:
            for satir in csv.DictReader(f, delimiter=AYIRICI):
                alan = (satir.get("alan") or "").strip().lower()
                ad = (satir.get("kolon_adi") or "").strip()
                if alan not in ALANLAR or not ad:
                    continue
                # Ornek satirlar kullanici silmediyse de zarar vermez; sadece
                # hicbir dosyada eslesmezler.
                sonuc.setdefault(alan, []).append(ad)
    except (OSError, csv.Error) as e:
        _log.warning("Kolon sozlugu okunamadi (%s): %s", yol, e)
        return {}
    return sonuc


def ekle(alan: str, kolon_adi: str, aciklama: str = "",
         veri_dizini: str | Path = "veri") -> bool:
    """Sozluge yeni bir kolon adi ekler. Zaten varsa hicbir sey yapmaz.

    Returns:
        Yeni kayit eklendiyse True.
    """
    alan = (alan or "").strip().lower()
    kolon_adi = (kolon_adi or "").strip()
    if alan not in ALANLAR or not kolon_adi:
        return False

    mevcut = yukle(veri_dizini)
    if any(_anahtar(a) == _anahtar(kolon_adi) for a in mevcut.get(alan, ())):
        return False

    yol = varsayilan_yol(veri_dizini)
    if not yol.is_file():
        olustur(yol)
    with open(yol, "a", encoding="utf-8-sig", newline="") as f:
        csv.writer(f, delimiter=AYIRICI).writerow(
            [alan, kolon_adi, aciklama or f"eklendi {date.today():%d.%m.%Y}"]
        )
    return True


def genislet(alan: str, varsayilanlar: Iterable[str],
             veri_dizini: str | Path = "veri") -> tuple[str, ...]:
    """Yerlesik aday listesini kullanici sozlugu ile birlestirir.

    Kullanici girdileri ONCE denenir; boylece bir kolon adini bilincli olarak
    baska bir alana yonlendirebilir.
    """
    kullanici = yukle(veri_dizini).get(alan, [])
    birlesik: list[str] = []
    gorulen: set[str] = set()
    for ad in list(kullanici) + list(varsayilanlar):
        a = _anahtar(ad)
        if a and a not in gorulen:
            gorulen.add(a)
            birlesik.append(ad)
    return tuple(birlesik)


def tanilama(yol: str | Path, veri_dizini: str | Path = "veri") -> dict:
    """Bir dosyanin kolonlarini inceler ve neyin bulunup neyin eksik oldugunu soyler.

    Arayuzdeki "yeni format" ekrani bunu gosterir: kullanici dosyadaki gercek
    kolon adlarini gorur ve eksik alanlari sozluge ekler.

    Returns:
        {
          "dosya": str, "sayfa": str | None,
          "basliklar": [str],                # dosyada bulunan kolon adlari
          "bulunan": {alan: kolon_adi},      # tahmin edilebilen alanlar
          "eksik": [alan],                   # bulunamayan alanlar
          "islenebilir": bool,               # zorunlu alanlardan biri var mi
          "mesaj": str,                      # kullaniciya gosterilecek Turkce ozet
        }
    """
    from masraf.okuyucular.genel import (
        baslik_satiri_bul,
        calisma_oku,
        hucre_metni,
        kolon_ara,
    )

    hedef = Path(yol)
    sonuc = {"dosya": hedef.name, "sayfa": None, "basliklar": [],
             "bulunan": {}, "eksik": list(ALANLAR), "islenebilir": False, "mesaj": ""}
    try:
        calisma = calisma_oku(hedef, satir_siniri=40)
    except Exception as e:  # noqa: BLE001
        sonuc["mesaj"] = f"Dosya acilamadi: {e.__class__.__name__}: {e}"
        return sonuc

    en_iyi = None
    for sayfa_adi, satirlar in calisma.sayfalar.items():
        try:
            indeks = baslik_satiri_bul(satirlar)
        except Exception:  # noqa: BLE001
            continue
        if indeks is None or indeks < 0 or indeks >= len(satirlar):
            continue
        adlar = [hucre_metni(h) or "" for h in satirlar[indeks]]
        dolu = [a for a in adlar if a]
        if en_iyi is None or len(dolu) > len(en_iyi[2]):
            en_iyi = (sayfa_adi, adlar, dolu)

    if en_iyi is None:
        sonuc["mesaj"] = ("Dosyada baslik satiri bulunamadi. Dosya bos olabilir "
                          "veya tablo bicimi cok farkli olabilir.")
        return sonuc

    sayfa_adi, adlar, _ = en_iyi
    sonuc["sayfa"] = sayfa_adi
    sonuc["basliklar"] = [a for a in adlar if a]

    from masraf.okuyucular import genel as _genel

    aday_tablolari = {
        "kisi": getattr(_genel, "_ISIM_ADAYLARI", ()),
        "sicil": getattr(_genel, "_SICIL_ADAYLARI", ()),
        "tckn": getattr(_genel, "_TCKN_ADAYLARI", ()),
        "tutar": getattr(_genel, "_TUTAR_ADAYLARI", ()),
        "tarih": getattr(_genel, "_TARIH_ADAYLARI", ()),
        "santiye": getattr(_genel, "_MERKEZ_ADAYLARI", ()),
    }
    bulunan: dict[str, str] = {}
    for alan, varsayilan in aday_tablolari.items():
        adaylar = genislet(alan, varsayilan, veri_dizini)
        try:
            i = kolon_ara(adlar, *adaylar)
        except Exception:  # noqa: BLE001
            i = None
        if i is not None and 0 <= i < len(adlar) and adlar[i]:
            bulunan[alan] = adlar[i]

    sonuc["bulunan"] = bulunan
    sonuc["eksik"] = [a for a in ALANLAR if a not in bulunan]
    sonuc["islenebilir"] = bool(ZORUNLU_ALANLAR & set(bulunan))

    if sonuc["islenebilir"]:
        sonuc["mesaj"] = (
            f"Dosya islenebilir. Bulunan alanlar: "
            + ", ".join(f"{ALANLAR[a]} = '{k}'" for a, k in bulunan.items())
        )
        if sonuc["eksik"]:
            sonuc["mesaj"] += (". Eksik alanlar: "
                              + ", ".join(ALANLAR[a] for a in sonuc["eksik"]))
    else:
        sonuc["mesaj"] = (
            "Dosyada kisi, sicil veya TC kimlik kolonu bulunamadi, bu yuzden "
            "islenemiyor. Dosyadaki kolon adlari: "
            + ", ".join(sonuc["basliklar"][:12])
            + ". Bu adlardan hangisi kisi adini tasiyorsa 'Kolon Sozlugu' "
              "ekranindan ekleyin; bir kez eklemek yeterlidir."
        )
    return sonuc
