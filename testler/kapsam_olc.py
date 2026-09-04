"""Kapsam ve eslesme orani olcumu: otomasyonun GERCEK performansini olcer.

Bu betik bir test degil, bir OLCUM ARACIDIR. Ornek veri dizinindeki tum kaynak
dosyalari uctan uca boru hattindan gecirir ve iki soruyu cevaplar:

1. Dosya bazinda otomasyon orani nedir?
   (toplam satir | kisi cikarilan | sicil bulunan | OTOMATIK | INCELE |
    ESLESMEDI | otomasyon orani %)

2. ESLESMEYEN her satir NEDEN eslesmedi? Uc kategoriye ayrilir:

   ``PARSER``
       Satirdan kisi adi hic cikarilamadi. Okuyucu zayif; DUZELTILEBILIR.
   ``VERI_KAPSAMI``
       Kisi adi cikarildi ama personel ana verisinde gercekten yok
       (taseron, grup sirketi calisani, dis danisman, aile bireyi).
       DUZELTILEMEZ - veri kapsaminin disinda.
   ``ALGORITMA``
       Kisi adi cikarildi, personel verisinde DE var, ama eslestirici
       bulamadi. DUZELTILEBILIR - en degerli kategori.

Ucuncu kategoriyi bulmak icin her eslesmeyen isim personel defterinde ELLE
aranir: once soyad indeksi, sonra tum isim sozlugu uzerinde rapidfuzz ile en
yakin 3 aday. Aday skoru ``KESIN_ESIK`` ustundeyse ALGORITMA sorunu sayilir.

Kullanim::

    python3 testler/kapsam_olc.py
    python3 testler/kapsam_olc.py --json rapor.json
    python3 testler/kapsam_olc.py --veri veri --ayri

Varsayilan olarak ogrenen defterlerin GECICI BIR KOPYASI kullanilir; olcum
kullanicinin ``veri/`` dizinini kirletmez ve her kosuda ayni noktadan baslar.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

KOK = Path(__file__).resolve().parent.parent
if str(KOK) not in sys.path:
    sys.path.insert(0, str(KOK))

from rapidfuzz import fuzz, process  # noqa: E402

from masraf.boru import Boru, CalismaAyarlari  # noqa: E402
import masraf.eslestirici as _eslestirici  # noqa: E402
from masraf.kayit import PersonelDefteri  # noqa: E402

#: Eslestiricinin "bu token gercekten soyad" esigi. Olcum, motorun kullandigi
#: ESIGIN AYNISINI kullanmalidir; aksi halde motorun bilerek reddettigi adlar
#: (GOKHAN, CAN, POLINA) "algoritma sorunu" olarak yanlis raporlanir. Esik
#: eslestiricide tanimli degilse (eski surum) yerel varsayilan kullanilir.
AILE_SOYAD_BELIRGIN: float = getattr(_eslestirici, "AILE_SOYAD_BELIRGIN", 0.50)
from masraf.metin import (  # noqa: E402
    isim_normalize,
    isim_tokenlari,
    rus_disi_soyad_erkek_hali,
    translit_varyantlari,
)
from masraf.modeller import (  # noqa: E402
    DURUM_ESLESMEDI,
    DURUM_INCELE,
    DURUM_OTOMATIK,
    Sonuc,
)

__all__ = [
    "KAPSAM_DOSYALARI",
    "VARSAYILAN_PERSONEL",
    "DosyaOlcumu",
    "EslesmeyenVaka",
    "Olcum",
    "olc",
    "elle_ara",
    "rapor_metni",
    "karsilastir",
]

#: Olcume girecek ornek dosyalar (personel ana verisi haric).
KAPSAM_DOSYALARI: tuple[str, ...] = (
    "ornek_veri/antik_travel/ANTIK_CARI_TEMMUZ_2026.xls",
    "ornek_veri/antik_travel/YUZYIL_TEMMUZ_2026_ELLE_DAGITILMIS.xlsx",
    "ornek_veri/energo/ASSESSMENT_YANSITMA_2026_05_06.xlsx",
    "ornek_veri/energo/ARABULUCULUK_2026_06_07.xlsx",
    "ornek_veri/energo/SAGLIK_KONTROL_LISTE.xlsx",
    "ornek_veri/energo/KOC_UNI_KATILIMCI_LISTESI.xlsx",
    "ornek_veri/posta/ornek_mail.msg",
)

VARSAYILAN_PERSONEL = "ornek_veri/personel/2025_2026_giris_cikis.xlsx"

# Elle arama esikleri (rapidfuzz token_set_ratio, 0-100).
#: Bu skorun ustundeki aday "kisi veride GERCEKTEN var" sayilir -> ALGORITMA.
KESIN_ESIK = 90.0
#: Bu skorun ustundeki adaylar rapora bilgi olarak yazilir.
ADAY_ESIGI = 60.0
#: Her eslesmeyen isim icin raporlanacak aday sayisi.
ADAY_SAYISI = 3

KATEGORI_PARSER = "PARSER"
KATEGORI_VERI = "VERI_KAPSAMI"
KATEGORI_ALGORITMA = "ALGORITMA"
KATEGORI_KISISIZ = "KISISIZ"

#: Kategori sirasi: en degerliden (duzeltilebilir) en degersize (kusur degil).
KATEGORILER: tuple[str, ...] = (
    KATEGORI_ALGORITMA, KATEGORI_PARSER, KATEGORI_VERI, KATEGORI_KISISIZ,
)

KATEGORI_ACIKLAMA: dict[str, str] = {
    KATEGORI_ALGORITMA: "Kisi veride var, eslestirici bulamadi - DUZELTILEBILIR",
    KATEGORI_PARSER: "Satirda kisi var ama cikarilamadi - DUZELTILEBILIR",
    KATEGORI_VERI: "Kisi personel ana verisinde yok - DUZELTILEMEZ",
    KATEGORI_KISISIZ: "Satirda kisi YOK (kurumsal gider) - KUSUR DEGIL",
}


# ----------------------------------------------------------------------
# Olcum veri yapilari
# ----------------------------------------------------------------------


@dataclass
class DosyaOlcumu:
    """Tek bir kaynak dosyanin olcum satiri."""

    dosya: str
    tip: str
    #: Bu dosya bir Outlook mesajinin EKI ise mesajin adi; degilse None.
    #: Ekler toplamlara KATILMAZ: ayni faturayi ikinci kez saymamak icin.
    kapsayici: str | None = None
    toplam: int = 0
    kisi_cikarilan: int = 0
    sicil_bulunan: int = 0
    otomatik: int = 0
    incele: int = 0
    eslesmedi: int = 0

    @property
    def otomasyon_orani(self) -> float:
        """OTOMATIK satirlarin toplam icindeki yuzdesi."""
        return round(self.otomatik / self.toplam * 100, 1) if self.toplam else 0.0

    @property
    def kisi_cikarma_orani(self) -> float:
        """Kisi adi cikarilabilen satirlarin yuzdesi (parser gucu)."""
        return round(self.kisi_cikarilan / self.toplam * 100, 1) if self.toplam else 0.0

    @property
    def cozulen_orani(self) -> float:
        """OTOMATIK + INCELE, yani sicile baglanabilen satirlarin yuzdesi."""
        if not self.toplam:
            return 0.0
        return round((self.otomatik + self.incele) / self.toplam * 100, 1)

    def sozluk(self) -> dict[str, Any]:
        return {
            "dosya": self.dosya,
            "tip": self.tip,
            "kapsayici": self.kapsayici,
            "toplam": self.toplam,
            "kisi_cikarilan": self.kisi_cikarilan,
            "sicil_bulunan": self.sicil_bulunan,
            "otomatik": self.otomatik,
            "incele": self.incele,
            "eslesmedi": self.eslesmedi,
            "otomasyon_orani": self.otomasyon_orani,
            "kisi_cikarma_orani": self.kisi_cikarma_orani,
            "cozulen_orani": self.cozulen_orani,
        }


@dataclass
class EslesmeyenVaka:
    """ESLESMEDI durumundaki tek bir satirin teshisi."""

    dosya: str
    satir_no: int
    kisi_ham: str
    aciklama: str
    kategori: str
    gerekce: str
    #: Outlook mesajinin eki ise mesaj adi; kategori dagilimina GIRMEZ.
    kapsayici: str | None = None
    adaylar: list[tuple[str, str, float]] = field(default_factory=list)

    def sozluk(self) -> dict[str, Any]:
        return {
            "dosya": self.dosya,
            "satir_no": self.satir_no,
            "kisi_ham": self.kisi_ham,
            "aciklama": self.aciklama[:160],
            "kategori": self.kategori,
            "gerekce": self.gerekce,
            "kapsayici": self.kapsayici,
            "adaylar": [
                {"sicil": sicil, "ad_soyad": ad, "skor": round(skor, 1)}
                for sicil, ad, skor in self.adaylar
            ],
        }


@dataclass
class Olcum:
    """Bir olcum kosusunun tum sonucu."""

    etiket: str
    dosyalar: list[DosyaOlcumu] = field(default_factory=list)
    vakalar: list[EslesmeyenVaka] = field(default_factory=list)
    yontem_dagilimi: dict[str, int] = field(default_factory=dict)
    hatalar: list[str] = field(default_factory=list)

    @property
    def ana_dosyalar(self) -> list[DosyaOlcumu]:
        """Toplamlara giren dosyalar: mesaj ekleri haric.

        ``ornek_mail.msg`` ayni seyahat ve saglik faturalarini EK olarak
        yeniden tasir. Ekler toplama katilirsa her satir iki kez sayilir ve
        oran yaniltici olur; bu yuzden ekler ayri raporlanir.
        """
        return [d for d in self.dosyalar if d.kapsayici is None]

    @property
    def ek_dosyalar(self) -> list[DosyaOlcumu]:
        """Outlook mesajindan cikan ek dosyalar (ayri raporlanir)."""
        return [d for d in self.dosyalar if d.kapsayici is not None]

    @property
    def toplam(self) -> int:
        return sum(d.toplam for d in self.ana_dosyalar)

    @property
    def otomatik(self) -> int:
        return sum(d.otomatik for d in self.ana_dosyalar)

    @property
    def incele(self) -> int:
        return sum(d.incele for d in self.ana_dosyalar)

    @property
    def eslesmedi(self) -> int:
        return sum(d.eslesmedi for d in self.ana_dosyalar)

    @property
    def kisi_cikarilan(self) -> int:
        return sum(d.kisi_cikarilan for d in self.ana_dosyalar)

    @property
    def otomasyon_orani(self) -> float:
        return round(self.otomatik / self.toplam * 100, 1) if self.toplam else 0.0

    @property
    def cozulen_orani(self) -> float:
        if not self.toplam:
            return 0.0
        return round((self.otomatik + self.incele) / self.toplam * 100, 1)

    @property
    def ana_vakalar(self) -> list[EslesmeyenVaka]:
        """Kategori dagilimina giren vakalar: mesaj eklerininki haric."""
        return [v for v in self.vakalar if v.kapsayici is None]

    @property
    def kategori_dagilimi(self) -> dict[str, int]:
        sayac = Counter(v.kategori for v in self.ana_vakalar)
        return {ad: sayac.get(ad, 0) for ad in KATEGORILER}

    @classmethod
    def sozlukten(cls, veri: dict[str, Any]) -> "Olcum":
        """``sozluk()`` ciktisindan olcumu geri kurar (onceki/sonraki icin).

        Yalnizca karsilastirmada kullanilan alanlar geri yuklenir; adaylar gibi
        detaylar JSON'da kalir.
        """
        olcum = cls(etiket=str(veri.get("etiket") or "olcum"),
                    yontem_dagilimi=dict(veri.get("yontem_dagilimi") or {}),
                    hatalar=list(veri.get("hatalar") or []))
        for ham in veri.get("dosyalar") or []:
            olcum.dosyalar.append(DosyaOlcumu(
                dosya=ham["dosya"], tip=ham.get("tip", "-"),
                kapsayici=ham.get("kapsayici"),
                toplam=int(ham.get("toplam", 0)),
                kisi_cikarilan=int(ham.get("kisi_cikarilan", 0)),
                sicil_bulunan=int(ham.get("sicil_bulunan", 0)),
                otomatik=int(ham.get("otomatik", 0)),
                incele=int(ham.get("incele", 0)),
                eslesmedi=int(ham.get("eslesmedi", 0)),
            ))
        for ham in veri.get("vakalar") or []:
            olcum.vakalar.append(EslesmeyenVaka(
                dosya=ham["dosya"], satir_no=int(ham.get("satir_no", 0)),
                kisi_ham=ham.get("kisi_ham", ""), aciklama=ham.get("aciklama", ""),
                kategori=ham.get("kategori", KATEGORI_VERI),
                gerekce=ham.get("gerekce", ""), kapsayici=ham.get("kapsayici"),
            ))
        return olcum

    def sozluk(self) -> dict[str, Any]:
        return {
            "etiket": self.etiket,
            "toplam": self.toplam,
            "otomatik": self.otomatik,
            "incele": self.incele,
            "eslesmedi": self.eslesmedi,
            "kisi_cikarilan": self.kisi_cikarilan,
            "otomasyon_orani": self.otomasyon_orani,
            "cozulen_orani": self.cozulen_orani,
            "kategori_dagilimi": self.kategori_dagilimi,
            "yontem_dagilimi": self.yontem_dagilimi,
            "dosyalar": [d.sozluk() for d in self.dosyalar],
            "vakalar": [v.sozluk() for v in self.vakalar],
            "hatalar": self.hatalar,
        }


# ----------------------------------------------------------------------
# Elle arama: kisi personel verisinde GERCEKTEN var mi?
# ----------------------------------------------------------------------


class ElleArayici:
    """Personel defterinde bir ismi elle arar (soyad + rapidfuzz).

    Eslestiricinin kademelerinden BAGIMSIZ calisir. Amaci "eslestirici bunu
    bulmali miydi?" sorusunu tarafsiz cevaplamaktir; bu yuzden eslestiricinin
    kendi indekslerini degil, dogrudan defterin isim sozlugunu tarar.
    """

    def __init__(self, defter: PersonelDefteri) -> None:
        self.defter = defter
        # Normalize isim -> sicil listesi. Bordrosuz taseron satirlari isimsiz
        # oldugu icin bu indekste zaten yoktur.
        self._isimler: dict[str, list[str]] = dict(defter._isim_index)
        self._havuz: list[str] = list(self._isimler.keys())
        # Token -> tokeni iceren isim sayisi ve ILK konumda (soyad) gecen isim
        # sayisi. Ikisinin orani bir tokenin soyad olma olasiligidir; eslestirici
        # ile AYNI olcuyu kullanmak sart, aksi halde motorun bilerek reddettigi
        # 'GOKHAN' / 'CAN' gibi ADLAR "algoritma sorunu" sanilir.
        self._token_sayisi: dict[str, int] = {}
        self._soyad_sayisi: dict[str, int] = {}
        for isim in self._havuz:
            parcalar = isim.split(" ")
            for token in set(parcalar):
                self._token_sayisi[token] = self._token_sayisi.get(token, 0) + 1
            self._soyad_sayisi[parcalar[0]] = self._soyad_sayisi.get(parcalar[0], 0) + 1

    def soyad_olasiligi(self, token: str) -> float:
        """Tokenin soyad olma olasiligi (bkz. Eslestirici._soyad_olasiligi)."""
        toplam = self._token_sayisi.get(token, 0)
        if not toplam:
            return 0.0
        return self._soyad_sayisi.get(token, 0) / toplam

    def _sicil_ad(self, isim_norm: str) -> tuple[str, str]:
        siciller = self._isimler.get(isim_norm, [])
        sicil = siciller[0] if siciller else ""
        kayit = self.defter.sicil_ile(sicil) if sicil else None
        return sicil, (kayit or {}).get("ad_soyad") or isim_norm

    def isim_izi(self, aciklama: str) -> str | None:
        """Aciklamada personel isim sozlugunden bir token geciyor mu?

        Kisi adi cikarilamayan satirlar icin ayirt edicidir: aciklamada bir
        isim izi VARSA okuyucu onu kacirmistir (PARSER sorunu); hicbir iz
        yoksa satir zaten kisi barindirmayan kurumsal bir giderdir.
        """
        sozluk = self.defter.isim_sozlugu
        for token in isim_normalize(aciklama).split(" "):
            if len(token) >= 4 and token in sozluk:
                return token
        return None

    def soyad_var_mi(self, ham_isim: str) -> str | None:
        """Isimdeki tokenlardan biri personel SOYAD indeksinde geciyor mu?

        Skor esigini gecemeyen ama soyadi veride bulunan isimler (aile
        bireyleri) icin kullanilir: kisinin kendisi veride olmasa da soyadi
        uzerinden dogru masraf merkezine baglanabilir, yani eslestiricinin
        en azindan ADAY uretmesi beklenir.
        """
        for token in isim_tokenlari(isim_normalize(ham_isim)):
            for aday in (token, rus_disi_soyad_erkek_hali(token)):
                if not aday or len(aday) < 3:
                    continue
                if not self.defter.soyad_ile_adaylar(aday):
                    continue
                # Yalnizca GERCEKTEN soyad olan tokenlar sayilir. 'GOKHAN',
                # 'CAN', 'POLINA' birer AD'dir; birkac kisinin soyadi olmalari
                # aile bagi kanit degildir ve eslestirici bunlari hakli olarak
                # reddeder.
                if self.soyad_olasiligi(aday) < AILE_SOYAD_BELIRGIN:
                    continue
                return aday
        return None

    def ara(self, ham_isim: str, sayi: int = ADAY_SAYISI) -> list[tuple[str, str, float]]:
        """En yakin adaylari (sicil, ad_soyad, skor) olarak dondurur."""
        norm = isim_normalize(ham_isim)
        if not norm:
            return []

        skorlar: dict[str, float] = {}

        def _ekle(isim_norm: str, skor: float) -> None:
            if skor > skorlar.get(isim_norm, 0.0):
                skorlar[isim_norm] = skor

        # 1) Soyad indeksi uzerinden daraltilmis arama. Fatura tarafinda ad
        #    sirasi belirsiz oldugu icin HER token soyad adayi sayilir; Rusca
        #    kadin soyadi eki ('NOVOSELOVA' -> 'NOVOSELOV') de cozulur, aksi
        #    halde defterdeki erkek hali gozden kacar.
        aranacak: set[str] = set()
        for token in isim_tokenlari(norm):
            aranacak.add(token)
            erkek = rus_disi_soyad_erkek_hali(token)
            if erkek:
                aranacak.add(erkek)
        for token in aranacak:
            for sicil in self.defter.soyad_ile_adaylar(token):
                kayit = self.defter.sicil_ile(sicil)
                if not kayit:
                    continue
                aday_norm = kayit.get("ad_soyad_norm") or isim_normalize(
                    kayit.get("ad_soyad") or "")
                if aday_norm:
                    _ekle(aday_norm, fuzz.token_set_ratio(norm, aday_norm))

        # 2) Tum sozluk uzerinde bulanik arama (soyadi da bozulmus olabilir).
        for aday_norm, skor, _ in process.extract(
            norm, self._havuz, scorer=fuzz.token_set_ratio,
            limit=sayi * 4, score_cutoff=ADAY_ESIGI,
        ):
            _ekle(aday_norm, float(skor))

        # 3) Rusca transliterasyon varyantlari uzerinden arama. 'YRMAK MEKHMET'
        #    ham haliyle dusuk skor alir, varyanti tam eslesir.
        for varyant in list(translit_varyantlari(norm))[:12]:
            if varyant == norm:
                continue
            for aday_norm, skor, _ in process.extract(
                varyant, self._havuz, scorer=fuzz.token_set_ratio,
                limit=sayi, score_cutoff=ADAY_ESIGI,
            ):
                _ekle(aday_norm, float(skor))

        sirali = sorted(skorlar.items(), key=lambda p: (-p[1], p[0]))[:sayi]
        sonuc: list[tuple[str, str, float]] = []
        for aday_norm, skor in sirali:
            sicil, ad = self._sicil_ad(aday_norm)
            sonuc.append((sicil, ad, skor))
        return sonuc


def elle_ara(defter: PersonelDefteri, ham_isim: str,
             sayi: int = ADAY_SAYISI) -> list[tuple[str, str, float]]:
    """Tek seferlik elle arama (test ve hata ayiklama kolayligi icin)."""
    return ElleArayici(defter).ara(ham_isim, sayi)


def _kategorize_et(sonuc: Sonuc, arayici: ElleArayici) -> EslesmeyenVaka:
    """Bir ESLESMEDI satirini kategorilerden birine yerlestirir."""
    satir = sonuc.satir
    ham = (satir.kisi_ham or "").strip()
    dosya = Path(satir.kaynak_dosya).name
    ortak = dict(dosya=dosya, satir_no=satir.satir_no,
                 aciklama=satir.aciklama or "",
                 kapsayici=dosya.split(EK_AYIRICI, 1)[0] if EK_AYIRICI in dosya else None)

    if not ham:
        # Kisi cikarilamamis olabilir; ama satirda GERCEKTEN kisi olmayabilir
        # de ('CENAZE CELENK GONDERIMI', 'TOPLANTI ORGANIZASYONU'). Ikisini
        # ayirmak sart: birincisi duzeltilecek bir kusur, ikincisi degil.
        ipucu = arayici.isim_izi(satir.aciklama or "")
        if ipucu is None:
            return EslesmeyenVaka(
                **ortak, kisi_ham="", kategori=KATEGORI_KISISIZ,
                gerekce=("Aciklamada personel isim sozlugunden hicbir token yok; "
                         "satir kurumsal bir gider, kisiye mahsuplasmaz."),
            )
        return EslesmeyenVaka(
            **ortak, kisi_ham="", kategori=KATEGORI_PARSER,
            gerekce=(f"Kisi adi cikarilamadi ama aciklamada isim izi var: '{ipucu}'."),
        )

    adaylar = arayici.ara(ham)
    en_iyi = adaylar[0][2] if adaylar else 0.0
    if en_iyi >= KESIN_ESIK:
        sicil, ad, skor = adaylar[0]
        return EslesmeyenVaka(
            **ortak, kisi_ham=ham, kategori=KATEGORI_ALGORITMA, adaylar=adaylar,
            gerekce=f"Personel verisinde '{ad}' (sicil {sicil}) skor {skor:.0f} ile bulundu.",
        )

    # Kisinin kendisi veride olmasa bile SOYADI veride olabilir (aile bireyi).
    # Bu durumda masraf merkezi soyadastan devralinabilir, yani eslestiricinin
    # en azindan aday uretmesi beklenir: aday uretmemisse ALGORITMA sorunudur.
    soyad = arayici.soyad_var_mi(ham)
    if soyad and not sonuc.eslesme.aday_siciller and sonuc.eslesme.yontem == "yok":
        return EslesmeyenVaka(
            **ortak, kisi_ham=ham, kategori=KATEGORI_ALGORITMA, adaylar=adaylar,
            gerekce=(f"Kisinin kendisi veride yok ama '{soyad}' soyadi veride var; "
                     "aile kurali aday uretmeliydi, hic aday uretmedi."),
        )

    return EslesmeyenVaka(
        **ortak, kisi_ham=ham, kategori=KATEGORI_VERI, adaylar=adaylar,
        gerekce=(
            (f"Personel verisinde karsilik yok (en yakin skor {en_iyi:.0f})"
             + (f"; '{soyad}' soyadli calisanlar aday olarak sunuldu." if soyad
                else "."))
            if adaylar else "Personel verisinde hicbir yakin aday yok."
        ),
    )


# ----------------------------------------------------------------------
# Olcum kosusu
# ----------------------------------------------------------------------


def _gecici_veri_dizini(kaynak: Path) -> Path:
    """Ogrenen defterlerin gecici bir kopyasini olusturur.

    Olcum kosusu defterlere yazar (yardimci kaynaklardan besleme, TCKN koprusu
    turetme). Kullanicinin ``veri/`` dizini kirlenmesin diye her kosu kendi
    kopyasi uzerinde calisir; boylece olcum TEKRARLANABILIRDIR.
    """
    hedef = Path(tempfile.mkdtemp(prefix="kapsam_veri_"))
    if kaynak.is_dir():
        for dosya in kaynak.glob("*.csv"):
            shutil.copy2(dosya, hedef / dosya.name)
    return hedef


def olc(
    dosyalar: Sequence[str | Path] = KAPSAM_DOSYALARI,
    personel_yolu: str | Path = VARSAYILAN_PERSONEL,
    veri_dizini: str | Path = "veri",
    etiket: str = "olcum",
    ayri_kosu: bool = False,
    gecici_veri: bool = True,
    kok: Path = KOK,
) -> Olcum:
    """Tum dosyalari boru hattindan gecirir ve olcumu dondurur.

    Args:
        dosyalar: Olculecek kaynak dosyalar (kok'e gore veya mutlak).
        personel_yolu: Personel ana verisi.
        veri_dizini: Ogrenen defterlerin dizini.
        etiket: Rapor basligi ('ONCESI' / 'SONRASI').
        ayri_kosu: True ise her dosya AYRI boru hattinda islenir. Boylece
            dosyalar arasi ogrenmenin (aile kurali, ek kisi defteri) katkisi
            olculebilir. Varsayilan False: gercek is akisi tum dosyalari
            birlikte verir.
        gecici_veri: Defterlerin gecici kopyasi uzerinde calis.
    """
    kok = Path(kok)
    yollar = [kok / d if not Path(d).is_absolute() else Path(d) for d in dosyalar]
    mevcut = [y for y in yollar if y.exists()]
    eksik = [str(y) for y in yollar if not y.exists()]

    kaynak_veri = kok / veri_dizini if not Path(veri_dizini).is_absolute() else Path(veri_dizini)
    calisma_veri = _gecici_veri_dizini(kaynak_veri) if gecici_veri else kaynak_veri

    try:
        ayarlar = CalismaAyarlari(
            personel_yolu=(kok / personel_yolu) if not Path(personel_yolu).is_absolute()
            else Path(personel_yolu),
            veri_dizini=calisma_veri,
            cikti_dizini=kok / "cikti",
            harita_yolu=kaynak_veri / "masraf_merkezi_haritasi.csv",
            ogrenmeyi_kaydet=gecici_veri,
        )
        boru = Boru(ayarlar)
        boru.hazirla()

        sonuclar: list[Sonuc] = []
        if ayri_kosu:
            for yol in mevcut:
                sonuclar.extend(boru.isle([yol]))
        else:
            sonuclar = boru.isle(mevcut)

        arayici = ElleArayici(boru.defter)

        # Dosya bazinda toplama. Kaynak tipi dosyadan degil satirdan alinir;
        # Outlook mesajindan cikan ekler farkli tipte olabilir.
        gruplar: dict[str, list[Sonuc]] = defaultdict(list)
        for sonuc in sonuclar:
            gruplar[Path(sonuc.satir.kaynak_dosya).name].append(sonuc)

        olcum = Olcum(etiket=etiket, hatalar=list(boru.hatalar))
        for yol in mevcut:
            ad = yol.name
            grup = gruplar.pop(ad, [])
            if not grup and any(a.startswith(ad + EK_AYIRICI) for a in gruplar):
                # Outlook mesaji bir KAPSAYICIDIR: kendisi gider satiri
                # uretmez, ekleri uretir. Bos bir satir olarak gosterilmez.
                continue
            olcum.dosyalar.append(_dosya_olcumu(ad, grup))
        # Boru hattinin urettigi ama listede olmayan kaynaklar (mesaj ekleri).
        for ad, grup in sorted(gruplar.items()):
            olcum.dosyalar.append(_dosya_olcumu(ad, grup))

        for yol in eksik:
            olcum.hatalar.append(f"Dosya bulunamadi: {yol}")

        for sonuc in sonuclar:
            if sonuc.durum == DURUM_ESLESMEDI:
                olcum.vakalar.append(_kategorize_et(sonuc, arayici))

        olcum.yontem_dagilimi = dict(
            sorted(Counter(s.eslesme.yontem or "yok" for s in sonuclar).items(),
                   key=lambda p: -p[1])
        )
        return olcum
    finally:
        if gecici_veri and calisma_veri.exists():
            shutil.rmtree(calisma_veri, ignore_errors=True)


#: Outlook okuyucusu ek dosyalari 'mesaj.msg > ek.xlsx' bicimiyle etiketler.
EK_AYIRICI = " > "


def _dosya_olcumu(ad: str, grup: list[Sonuc]) -> DosyaOlcumu:
    """Bir dosyaya ait sonuclardan olcum satirini uretir."""
    tip = grup[0].satir.kaynak_tip if grup else "-"
    kapsayici = ad.split(EK_AYIRICI, 1)[0] if EK_AYIRICI in ad else None
    olcum = DosyaOlcumu(dosya=ad, tip=tip, kapsayici=kapsayici, toplam=len(grup))
    for sonuc in grup:
        if (sonuc.satir.kisi_ham or "").strip():
            olcum.kisi_cikarilan += 1
        if sonuc.eslesme.sicil:
            olcum.sicil_bulunan += 1
        if sonuc.durum == DURUM_OTOMATIK:
            olcum.otomatik += 1
        elif sonuc.durum == DURUM_INCELE:
            olcum.incele += 1
        else:
            olcum.eslesmedi += 1
    return olcum


# ----------------------------------------------------------------------
# Raporlama
# ----------------------------------------------------------------------


def _tablo(basliklar: Sequence[str], satirlar: Sequence[Sequence[Any]]) -> str:
    """Sabit genislikli metin tablosu uretir."""
    hucreler = [[str(h) for h in basliklar]] + [[str(h) for h in s] for s in satirlar]
    genislik = [max(len(satir[i]) for satir in hucreler) for i in range(len(basliklar))]
    parcalar = []
    for sira, satir in enumerate(hucreler):
        hizali = [
            satir[i].ljust(genislik[i]) if i == 0 else satir[i].rjust(genislik[i])
            for i in range(len(basliklar))
        ]
        parcalar.append(" | ".join(hizali))
        if sira == 0:
            parcalar.append("-+-".join("-" * g for g in genislik))
    return "\n".join(parcalar)


def rapor_metni(olcum: Olcum, vaka_detayi: bool = True) -> str:
    """Olcumu okunabilir metne cevirir."""
    satirlar: list[str] = []
    satirlar.append("=" * 100)
    satirlar.append(f"KAPSAM VE ESLESME ORANI OLCUMU - {olcum.etiket}")
    satirlar.append("=" * 100)
    satirlar.append("")

    basliklar = ["dosya", "tip", "toplam", "kisi", "sicil",
                 "OTOM", "INCE", "ESLESMEDI", "otom %"]

    def _satir(d: DosyaOlcumu) -> list[Any]:
        return [d.dosya, d.tip, d.toplam, d.kisi_cikarilan, d.sicil_bulunan,
                d.otomatik, d.incele, d.eslesmedi, f"{d.otomasyon_orani:.1f}"]

    satirlar.append("DOSYA BAZINDA")
    satirlar.append(_tablo(
        basliklar,
        [_satir(d) for d in olcum.ana_dosyalar] + [
            ["TOPLAM", "-", olcum.toplam, olcum.kisi_cikarilan,
             sum(d.sicil_bulunan for d in olcum.ana_dosyalar),
             olcum.otomatik, olcum.incele, olcum.eslesmedi,
             f"{olcum.otomasyon_orani:.1f}"]
        ],
    ))
    satirlar.append("")
    satirlar.append(f"Cozulen (OTOMATIK + INCELE): %{olcum.cozulen_orani}")
    satirlar.append("")

    ekler = olcum.ek_dosyalar
    if ekler:
        satirlar.append(
            "OUTLOOK MESAJI EKLERI (ayni faturalari yeniden tasir; TOPLAMA GIRMEZ)")
        satirlar.append(_tablo(basliklar, [_satir(d) for d in ekler]))
        satirlar.append("")

    satirlar.append("ESLESMEYEN SATIRLARIN KATEGORI DAGILIMI")
    dagilim = olcum.kategori_dagilimi
    toplam_vaka = sum(dagilim.values()) or 1
    satirlar.append(_tablo(
        ["kategori", "adet", "%", "aciklama"],
        [
            [ad, adet, f"{adet / toplam_vaka * 100:.1f}", KATEGORI_ACIKLAMA[ad]]
            for ad, adet in dagilim.items()
        ],
    ))
    satirlar.append("")

    satirlar.append("ESLESTIRME YONTEMI DAGILIMI")
    satirlar.append(_tablo(
        ["yontem", "adet"],
        [[ad, adet] for ad, adet in olcum.yontem_dagilimi.items()],
    ))
    satirlar.append("")

    if vaka_detayi and olcum.ana_vakalar:
        satirlar.append("ESLESMEYEN VAKALAR (mesaj ekleri haric)")
        for kategori in KATEGORILER:
            grup = [v for v in olcum.ana_vakalar if v.kategori == kategori]
            if not grup:
                continue
            satirlar.append("")
            satirlar.append(f"--- {kategori} ({len(grup)} satir) ---")
            for vaka in grup:
                ad = vaka.kisi_ham or "(kisi cikarilamadi)"
                satirlar.append(f"  [{vaka.dosya}:{vaka.satir_no}] {ad}")
                satirlar.append(f"      gerekce : {vaka.gerekce}")
                if vaka.kategori in (KATEGORI_PARSER, KATEGORI_KISISIZ):
                    satirlar.append(f"      aciklama: {vaka.aciklama[:110]}")
                for sicil, aday_ad, skor in vaka.adaylar:
                    satirlar.append(f"      aday    : {skor:5.1f}  {sicil:>8}  {aday_ad}")
        satirlar.append("")

    if olcum.hatalar:
        satirlar.append("BORU HATTI NOTLARI")
        for hata in olcum.hatalar:
            satirlar.append(f"  - {hata}")
        satirlar.append("")

    return "\n".join(satirlar)


def karsilastir(oncesi: Olcum, sonrasi: Olcum) -> str:
    """Iki olcumu yan yana koyar (iyilestirme oncesi / sonrasi)."""
    satirlar = ["=" * 100, "ONCESI / SONRASI KARSILASTIRMA", "=" * 100, ""]

    onceki = {d.dosya: d for d in oncesi.dosyalar}
    veri: list[list[Any]] = []
    for sonra in sonrasi.ana_dosyalar:
        once = onceki.get(sonra.dosya)
        if once is None:
            continue
        veri.append([
            sonra.dosya,
            f"{once.otomatik}/{once.incele}/{once.eslesmedi}",
            f"{sonra.otomatik}/{sonra.incele}/{sonra.eslesmedi}",
            f"{once.otomasyon_orani:.1f}",
            f"{sonra.otomasyon_orani:.1f}",
            f"{sonra.otomasyon_orani - once.otomasyon_orani:+.1f}",
        ])
    veri.append([
        "TOPLAM",
        f"{oncesi.otomatik}/{oncesi.incele}/{oncesi.eslesmedi}",
        f"{sonrasi.otomatik}/{sonrasi.incele}/{sonrasi.eslesmedi}",
        f"{oncesi.otomasyon_orani:.1f}",
        f"{sonrasi.otomasyon_orani:.1f}",
        f"{sonrasi.otomasyon_orani - oncesi.otomasyon_orani:+.1f}",
    ])
    satirlar.append(_tablo(
        ["dosya", "once O/I/E", "sonra O/I/E", "once %", "sonra %", "fark"], veri))
    satirlar.append("")

    once_kat = oncesi.kategori_dagilimi
    sonra_kat = sonrasi.kategori_dagilimi
    satirlar.append("KATEGORI DAGILIMI")
    satirlar.append(_tablo(
        ["kategori", "once", "sonra", "fark"],
        [[ad, once_kat[ad], sonra_kat[ad], f"{sonra_kat[ad] - once_kat[ad]:+d}"]
         for ad in once_kat],
    ))
    satirlar.append("")
    return "\n".join(satirlar)


def main(argv: Sequence[str] | None = None) -> int:
    """Komut satiri girisi."""
    ayrist = argparse.ArgumentParser(
        description="Masraf mahsuplasma otomasyonunun kapsam ve eslesme oranini olcer.")
    ayrist.add_argument("--personel", default=VARSAYILAN_PERSONEL,
                        help="Personel ana verisi yolu.")
    ayrist.add_argument("--veri", default="veri", help="Ogrenen defterlerin dizini.")
    ayrist.add_argument("--dosya", action="append", default=None,
                        help="Olculecek dosya (birden fazla verilebilir).")
    ayrist.add_argument("--ayri", action="store_true",
                        help="Her dosyayi ayri kosuda isle (dosyalar arasi ogrenme kapali).")
    ayrist.add_argument("--gercek-veri", action="store_true",
                        help="Gecici kopya yerine gercek veri dizinine yaz.")
    ayrist.add_argument("--kisa", action="store_true", help="Vaka detayini bastirma.")
    ayrist.add_argument("--json", default=None, help="Olcumu JSON olarak bu dosyaya yaz.")
    ayrist.add_argument("--karsilastir", default=None,
                        help="Onceki bir olcum JSON'u; oncesi/sonrasi tablosu basilir.")
    ayrist.add_argument("--etiket", default="OLCUM", help="Rapor basligi.")
    secenekler = ayrist.parse_args(argv)

    olcum = olc(
        dosyalar=secenekler.dosya or KAPSAM_DOSYALARI,
        personel_yolu=secenekler.personel,
        veri_dizini=secenekler.veri,
        etiket=secenekler.etiket,
        ayri_kosu=secenekler.ayri,
        gecici_veri=not secenekler.gercek_veri,
    )
    print(rapor_metni(olcum, vaka_detayi=not secenekler.kisa))

    if secenekler.karsilastir:
        onceki_yol = Path(secenekler.karsilastir)
        if not onceki_yol.exists():
            print(f"Karsilastirilacak olcum bulunamadi: {onceki_yol}")
        else:
            oncesi = Olcum.sozlukten(
                json.loads(onceki_yol.read_text(encoding="utf-8")))
            print(karsilastir(oncesi, olcum))

    if secenekler.json:
        hedef = Path(secenekler.json)
        hedef.parent.mkdir(parents=True, exist_ok=True)
        hedef.write_text(
            json.dumps(olcum.sozluk(), ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON yazildi: {hedef}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
