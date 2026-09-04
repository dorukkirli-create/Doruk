"""Kaynak dosya tipini otomatik tespit eden kesif modulu.

Kullanici dosyayi surukleyip biraktiginda hangi parser'in calistirilacagini
belirler. Tespit tamamen deterministiktir: dosya acilir, SAYFA ADLARI ve ilk
15 satirdaki hucre metinleri ASCII katlanmis bicimde toplanir, ardindan
oncelik sirali ipucu kurallari uygulanir.

Ipuclari:
    'Cari Hareket Dokumu'                -> antik_cari
    'SANTIYESI' + 'UCUS GUZERGAHI'       -> yuzyil_dagitilmis
    'Katilimci' + 'Paket'                -> energo_assessment
    'ARABULUCU'                          -> energo_arabulucu
    'BORDROLU LISTE' | 'TCKN'+'SAGLIK'   -> energo_saglik
    'ID' + 'Alt Fonksiyon'               -> koc_katilimci
    aksi halde                           -> genel
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable

from masraf.modeller import GiderSatiri
from masraf.okuyucular.antik import antik_cari_oku, yuzyil_dagitilmis_oku
from masraf.okuyucular.energo import (
    arabulucu_oku,
    assessment_oku,
    koc_katilimci_oku,
    saglik_oku,
)
from masraf.okuyucular.genel import calisma_oku, genel_oku, kolon_anahtari

__all__ = ["dosya_tipini_bul", "oku", "oku_tip", "PARSERLAR"]

# kaynak_tip -> parser fonksiyonu
PARSERLAR: dict[str, Callable[[str | Path], list[GiderSatiri]]] = {
    "antik_cari": antik_cari_oku,
    "yuzyil_dagitilmis": yuzyil_dagitilmis_oku,
    "energo_assessment": assessment_oku,
    "energo_arabulucu": arabulucu_oku,
    "energo_saglik": saglik_oku,
    "koc_katilimci": koc_katilimci_oku,
    "referans_liste": genel_oku,
    "genel": genel_oku,
}

# Bir dosyanin gider satiri mi yoksa kisi kutugu mu oldugunu ayirt eden esik.
# Kutuklerde hicbir satirda tutar yoktur; faturalarda neredeyse her satirda vardir.
_KUTUK_SATIR_ESIGI = 200

# Kesif icin okunacak satir sayisi (dosyanin tamami okunmaz, hizli kalir).
_KESIF_SATIR_SINIRI = 15


def _ipuclarini_topla(yol: Path) -> tuple[set[str], str]:
    """Sayfa adlari ve ilk satirlardaki hucre metinlerini toplar.

    Returns:
        (normalize edilmis benzersiz metinler kumesi, hepsinin birlesimi)
    """
    calisma = calisma_oku(yol, satir_siniri=_KESIF_SATIR_SINIRI)
    metinler: set[str] = set()
    for sayfa_adi, satirlar in calisma.sayfalar.items():
        anahtar = kolon_anahtari(sayfa_adi)
        if anahtar:
            metinler.add(anahtar)
        for satir in satirlar:
            for hucre in satir:
                anahtar = kolon_anahtari(hucre)
                if anahtar:
                    metinler.add(anahtar)
    return metinler, " || ".join(sorted(metinler))


def _iceriyor(blob: str, *parcalar: str) -> bool:
    """Birlesik metinde verilen parcalarin HEPSI geciyor mu?"""
    return all(kolon_anahtari(p) in blob for p in parcalar)


def dosya_tipini_bul(yol: str | Path) -> str:
    """Dosyanin hangi kaynak ailesine ait oldugunu belirler.

    Donen deger modeller.KAYNAK_TIPLERI kumesindendir. Dosya acilamiyorsa
    veya hicbir ipucu eslesmiyorsa 'genel' doner (istisna firlatmaz).
    """
    p = Path(yol)

    # 0) Outlook mesaji: icerik degil uzanti belirler. Mesaj bir kapsayicidir,
    #    icindeki tablo dosyalari ayri ayri tespit edilir.
    if p.suffix.lower() == ".msg":
        return "outlook_msg"

    try:
        metinler, blob = _ipuclarini_topla(p)
    except Exception:
        return "genel"

    # 1) Antik ham cari hareket dokumu
    if _iceriyor(blob, "cari hareket"):
        return "antik_cari"

    # 2) Yuzyil elle dagitilmis (santiye + ucus kolonlari birlikte)
    if _iceriyor(blob, "santiyesi") and _iceriyor(blob, "ucus guzergahi"):
        return "yuzyil_dagitilmis"
    if _iceriyor(blob, "ucus guzergahi") and _iceriyor(blob, "ucus bilgisi"):
        return "yuzyil_dagitilmis"

    # 3) Energo assessment
    if "katilimci" in metinler and "paket" in metinler:
        return "energo_assessment"

    # 4) Energo arabuluculuk
    if "arabulucu" in metinler or _iceriyor(blob, "arabulucu"):
        return "energo_arabulucu"

    # 5) Saglik kontrol listesi
    if _iceriyor(blob, "bordrolu liste") or _iceriyor(blob, "bordrosuz liste"):
        return "energo_saglik"
    if _iceriyor(blob, "tckn") and _iceriyor(blob, "saglik kontrol"):
        return "energo_saglik"

    # 6) Koc Universitesi katilimci listesi
    if "id" in metinler and _iceriyor(blob, "alt fonksiyon"):
        return "koc_katilimci"

    # 7) Ferdi kaza sigorta listesi ve benzeri PERSONEL KUTUKLERI.
    #    Bunlar fatura degil, kisi kutugudur: sicil ve masraf merkezi tasir
    #    ama tutar tasimaz. Gider satiri olarak islenirlerse binlerce sahte
    #    satir uretirler; ayri tip olarak isaretlenip defter beslemesine
    #    yonlendirilirler.
    if _iceriyor(blob, "sicil no") and _iceriyor(blob, "masraf merkezi"):
        return "referans_liste"

    return "genel"


def oku_tip(yol: str | Path, tip: str) -> list[GiderSatiri]:
    """Verilen kaynak tipinin parser'ini calistirir.

    Raises:
        KeyError: tip taninmiyorsa.
    """
    return PARSERLAR[tip](yol)


class MesajOkunamadi(Exception):
    """Outlook mesajindan hic gider satiri cikarilamadi ve sebebi biliniyor.

    Bu istisna KASITLIDIR. Onceki surumde her ek hatasi sessizce yutuluyordu
    (``except Exception: continue``) ve kullaniciya yalnizca '0 satir' deniyordu.
    Kullanici bu mesajla ne yapacagini bilemez: ek mi yok, ek var da okunamadi
    mi, hangi ek, neden? Artik hepsi yaziliyor.
    """


def _msg_oku(yol: Path, cikarma_dizini: str | Path | None = None) -> list[GiderSatiri]:
    """Outlook mesajindaki tum tablo eklerini cikarir ve tek tek okur.

    Mesaj bir kapsayicidir: icinde baska mesajlar, zip arsivleri ve Excel
    dosyalari olabilir. Cikarilan her dosyanin tipi ayrica tespit edilir.
    Her satira hangi mailden geldigi `ek['mail_konusu']` icinde yazilir.

    Raises:
        MesajOkunamadi: Hicbir ekten satir cikmadiysa, sebebiyle birlikte.
    """
    from tempfile import mkdtemp

    from masraf.okuyucular.posta import msg_aciklarini_cikar

    hedef = Path(cikarma_dizini) if cikarma_dizini else Path(mkdtemp(prefix="masraf_msg_"))
    satirlar: list[GiderSatiri] = []
    ekler = msg_aciklarini_cikar(yol, hedef)

    if not ekler:
        raise MesajOkunamadi(
            "mesajin icinde okunabilir tablo eki bulunamadi. Aranan uzantilar: "
            ".xlsx .xls .xlsm .csv .tsv. Mail yalnizca metin/gorsel tasiyor "
            "olabilir, ya da ekler mailin govdesine gomulu olabilir. Ekleri "
            "Outlook'ta kaydedip dogrudan 1_FATURALAR klasorune atmayi deneyin."
        )

    # Her ek icin ne oldugunu ayri ayri tut; hepsi basarisiz olursa raporla.
    bos_kalanlar: list[str] = []
    hatalilar: list[str] = []
    for ek in ekler:
        try:
            ic_satirlar = oku(ek.yol)
        except Exception as hata:  # noqa: BLE001 - kullaniciya gosterilecek
            hatalilar.append(f"{ek.ad}: {hata.__class__.__name__}: {hata}")
            continue
        if not ic_satirlar:
            bos_kalanlar.append(ek.ad)
            continue
        for s in ic_satirlar:
            # Kaynak dosya adini mesaj + ek olarak yaz, izlenebilirlik icin.
            s.kaynak_dosya = f"{yol.name} > {ek.ad}"
            if isinstance(s.ek, dict):
                s.ek.setdefault("mail_konusu", ek.mail_konusu)
                s.ek.setdefault("mail_gonderen", ek.mail_gonderen)
                s.ek.setdefault("mail_tarihi", ek.mail_tarihi)
                s.ek.setdefault("mail_zinciri", ek.kaynak_aciklamasi)
        satirlar.extend(ic_satirlar)

    if satirlar:
        return satirlar

    parcalar = [f"mesajdan {len(ekler)} tablo eki cikarildi ama hicbirinden "
                "gider satiri okunamadi."]
    if bos_kalanlar:
        parcalar.append(
            "Acildi ama bos donenler: " + ", ".join(bos_kalanlar[:8])
            + (f" (+{len(bos_kalanlar) - 8} tane daha)" if len(bos_kalanlar) > 8 else "")
            + ". Bu dosyalarin kolon adlari taninmamis olabilir; "
              "veri/kolon_esanlamlilari.csv dosyasina ekleyin."
        )
    if hatalilar:
        parcalar.append("Hata verenler: " + " | ".join(hatalilar[:5]))
    raise MesajOkunamadi(" ".join(parcalar))


def oku(yol: str | Path, cikarma_dizini: str | Path | None = None) -> list[GiderSatiri]:
    """Dosya tipini bulur ve dogru parser'i calistirir.

    Outlook mesajlari kapsayici olarak ele alinir: icindeki tum tablo ekleri
    cikarilip ayri ayri okunur ve tek listede birlestirilir.

    Ozel parser hic satir uretmezse (sablon beklenenden farkliysa) genel
    parser'a duser; boylece bilinmeyen bir surum sessizce bos sonuc vermez.
    """
    p = Path(yol)
    tip = dosya_tipini_bul(p)
    if tip == "outlook_msg":
        return _msg_oku(p, cikarma_dizini)
    satirlar = oku_tip(p, tip)
    if not satirlar and tip != "genel":
        satirlar = genel_oku(p)

    # Guvenlik agi: cok satirli ve hicbir satirinda tutar olmayan bir dosya
    # fatura degil kisi kutugudur. Gider olarak islenirse sahte satir uretir.
    if (tip == "genel" and len(satirlar) >= _KUTUK_SATIR_ESIGI
            and not any(s.tutar is not None for s in satirlar)):
        tip = "referans_liste"

    if tip == "referans_liste":
        for s in satirlar:
            s.kaynak_tip = "referans_liste"
            if isinstance(s.ek, dict):
                s.ek["referans_liste"] = True
    return satirlar
