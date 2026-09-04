"""Veri siniflari: gider satiri, eslesme ve nihai sonuc kayitlari.

Bu modul hicbir agir bagimliliga (pandas vb.) sahip degildir; tum parser,
eslestirici ve arayuz modulleri bu sozlesmeyi import eder.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

# Kaynak dosya aileleri. Parser modulleri kaynak_tip alanini bu kumeden secer.
KAYNAK_TIPLERI: frozenset[str] = frozenset({
    "antik_cari",
    "yuzyil_dagitilmis",
    "energo_assessment",
    "energo_arabulucu",
    "energo_saglik",
    "koc_katilimci",
    "genel",
})

# Gider tipleri.
GIDER_TIPLERI: frozenset[str] = frozenset({
    "Bilet",
    "Otel",
    "Vize",
    "Egitim",
    "Saglik",
    "Arabuluculuk",
    "Bagaj",
    "Diger",
})

# Eslestirme yontemleri, guvenilirlik sirasina yakin bicimde.
YONTEMLER: frozenset[str] = frozenset({
    "sicil",
    "tckn",
    "tam_isim",
    "alias",
    "alt_kume",
    "transliterasyon",
    "prefix",
    "bulanik",
    "aile",
    "ek_defter",
    "yok",
})

# Sonuc durumlari.
DURUM_OTOMATIK = "OTOMATIK"
DURUM_INCELE = "INCELE"
DURUM_ESLESMEDI = "ESLESMEDI"


@dataclass
class GiderSatiri:
    """Kaynak fatura dosyasindan cikarilmis tek bir gider satiri.

    Parser modulleri ham dosyayi okuyup bu nesnelerin listesini uretir.
    Eslestirici bu nesneleri girdi olarak alir.
    """

    kaynak_dosya: str
    kaynak_tip: str            # 'antik_cari' | 'yuzyil_dagitilmis' | 'energo_assessment' | 'energo_arabulucu' | 'energo_saglik' | 'koc_katilimci' | 'genel'
    satir_no: int              # kaynak dosyadaki 1 tabanli satir
    belge_tarihi: date | None
    aciklama: str
    kisi_ham: str | None       # ham metinden cikarilan kisi adi
    sicil_ham: str | None
    tckn_ham: str | None
    tutar: float | None
    para_birimi: str | None
    masraf_merkezi_kaynak: str | None   # kaynak dosyada belirtilmis santiye/proje (varsa)
    gider_tipi: str | None     # 'Bilet' | 'Otel' | 'Vize' | 'Egitim' | 'Saglik' | 'Arabuluculuk' | 'Bagaj' | 'Diger'
    ek: dict = field(default_factory=dict)


@dataclass
class Eslesme:
    """Bir gider satirindaki kisinin personel defterindeki karsiligi.

    'guven' 0.0 - 1.0 arasindadir; 'aciklama' kullaniciya gosterilecek
    Turkce gerekcedir, kullanici neden bu sonuca varildigini gormelidir.
    """

    sicil: str | None
    ad_soyad: str | None
    yontem: str                # 'sicil' | 'tckn' | 'tam_isim' | 'alias' | 'alt_kume' | 'transliterasyon' | 'prefix' | 'bulanik' | 'aile' | 'ek_defter' | 'yok'
    guven: float               # 0.0 - 1.0
    aday_sayisi: int
    aciklama: str              # kullaniciya gosterilecek Turkce gerekce
    aday_siciller: list = field(default_factory=list)


@dataclass
class Sonuc:
    """Bir gider satirinin nihai mahsuplastirma sonucu.

    'durum' alani ciktidaki sayfayi belirler:
    OTOMATIK -> 'Sonuc', INCELE -> 'Incele', ESLESMEDI -> 'Eslesmedi'.
    """

    satir: GiderSatiri
    eslesme: Eslesme
    donem: date | None
    gorev_yeri: str | None
    masraf_merkezi: str | None
    sirket: str | None
    sirket2: str | None
    statu: str | None
    kategori: str | None       # Aktif / Cikis
    cikis_tarihi: date | None
    durum: str                 # 'OTOMATIK' | 'INCELE' | 'ESLESMEDI'
    uyarilar: list = field(default_factory=list)
    # Gider ayi ile personel kaydinin donemi ortusuyor mu:
    # 'tam' | 'onceki_donem' | 'ilk_donem_oncesi' | 'tarihsiz' | 'yok'
    donem_eslesme: str = "yok"


def bos_eslesme(aciklama: str = "Kisi bulunamadi") -> Eslesme:
    """Eslesme kurulamadigi durumlar icin standart bos eslesme uretir."""
    return Eslesme(
        sicil=None,
        ad_soyad=None,
        yontem="yok",
        guven=0.0,
        aday_sayisi=0,
        aciklama=aciklama,
        aday_siciller=[],
    )
