"""Masraf merkezi mahsuplastirma otomasyonu - masaustu Streamlit arayuzu.

Finans ekibi icin tek dosyalik arayuz. Calisma zamaninda YAPAY ZEKA VEYA
INTERNET KULLANILMAZ; tum eslestirme ``masraf`` paketindeki deterministik
modullerle yapilir.

Sekmeler:
    Ayarlar                 personel ana verisini yukle, guven esigi, defter durumu
    Fatura Isle             dosya yukle / klasor sec, isle, sonuc ve Excel indir
    Inceleme                dusuk guvenli satirlari tek tek coz ve SISTEME OGRET
    Masraf Merkezi Haritasi gorev yeri -> masraf merkezi kodu tablosunu duzenle
    Yardim                  kisa Turkce kullanim kilavuzu

Calistirma:
    python -m streamlit run app.py
    (veya Windows'ta baslat.bat, Linux/macOS'ta ./baslat.sh)

Tasarim notu: personel ana verisi 24 MB'lik bir xlsx'tir ve ilk okunusu ~25 sn
surer. Bu yuzden defter ``st.cache_resource`` ile onbelleklenir ve arayuzun
her etkilesiminde YENIDEN YUKLENMEZ.
"""

from __future__ import annotations

import csv
import tempfile
import traceback
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Cekirdek modul importlari (savunmali)
#
# Diger moduller ayni anda gelistiriliyor olabilir. Cekirdek eksikse uygulama
# cokmemeli, arayuzde anlasilir bir uyari gostermelidir.
# --------------------------------------------------------------------------

CEKIRDEK_HAZIR = True
CEKIRDEK_HATA = ""

try:
    from masraf.defter import Defterler
    from masraf.eslestirici import INCELE_ESIGI, Eslestirici, durum_belirle
    from masraf.kayit import PersonelDefteri, sicil_normalize
    from masraf.metin import ascii_katla, isim_normalize, kisi_metnini_temizle
    from masraf.modeller import (
        DURUM_ESLESMEDI,
        DURUM_INCELE,
        DURUM_OTOMATIK,
        Eslesme,
        GiderSatiri,
        Sonuc,
    )
    from masraf.okuyucular.kesif import dosya_tipini_bul, oku
except Exception:  # pragma: no cover - sadece eksik kurulumda calisir
    CEKIRDEK_HAZIR = False
    CEKIRDEK_HATA = traceback.format_exc()
    DURUM_OTOMATIK, DURUM_INCELE, DURUM_ESLESMEDI = "OTOMATIK", "INCELE", "ESLESMEDI"
    INCELE_ESIGI = 0.80

# Boru hatti ve cikti modulleri. Bunlar olmadan da calisir (yerel esdegerler
# asagida tanimlidir) ama VARSA TERCIH EDILIR: isleme sirasi, masraf merkezi
# cozumu ve Excel bicimi orada tanimlidir.
BORU_SINIFI: Any = None
AYAR_SINIFI: Any = None
try:
    from masraf.boru import Boru as BORU_SINIFI  # type: ignore
    from masraf.boru import CalismaAyarlari as AYAR_SINIFI  # type: ignore
except Exception:
    BORU_SINIFI = None
    AYAR_SINIFI = None

CIKTI_MODULU: Any = None
try:
    from masraf import cikti as CIKTI_MODULU  # type: ignore
except Exception:
    CIKTI_MODULU = None

MAHSUP_MODULU: Any = None
try:
    from masraf import mahsuplasma as MAHSUP_MODULU  # type: ignore
except Exception:
    MAHSUP_MODULU = None


# --------------------------------------------------------------------------
# Sabitler
# --------------------------------------------------------------------------

UYGULAMA_ADI = "Masraf Merkezi Otomasyonu"

# Kurumsal renkler: koyu lacivert basliklar, mavi vurgu.
RENK_LACIVERT = "#1F3864"
RENK_MAVI = "#2E75B6"
RENK_ACIK = "#F4F7FB"

RENK_OTOMATIK = "#E2EFDA"
RENK_INCELE = "#FFF2CC"
RENK_ESLESMEDI = "#FBE5E5"

PROJE_KOK = Path(__file__).resolve().parent
VARSAYILAN_YARDIMCI_ADI = "1C_Personnel_List_31082026.xlsx"
VARSAYILAN_PERSONEL = PROJE_KOK / "ornek_veri" / "personel" / "2025_2026_giris_cikis.xlsx"
# 1C personel listesi: grup sirketlerini (Renservis, Renstroydetal, RC, One
# Tower, Top Tower) kapsayan ikincil defter. Ana veride bulunamayan kisiler
# buradan aranir. Zorunlu degildir.
VARSAYILAN_YARDIMCI = PROJE_KOK / "ornek_veri" / "personel" / VARSAYILAN_YARDIMCI_ADI
VERI_KOK = PROJE_KOK / "veri"
CIKTI_KOK = PROJE_KOK / "cikti"
HARITA_DOSYASI = "masraf_merkezi_haritasi.csv"

DESTEKLENEN_UZANTILAR = (".xls", ".xlsx", ".xlsm", ".csv", ".msg")

# kaynak_tip -> kullaniciya gosterilecek Turkce ad.
KAYNAK_TIP_ADLARI: dict[str, str] = {
    "antik_cari": "Antik/Yuzyil seyahat - ham cari hareket dokumu",
    "yuzyil_dagitilmis": "Yuzyil seyahat - elle dagitilmis (referans)",
    "energo_assessment": "Energo - assessment yansitma",
    "energo_arabulucu": "Energo - arabuluculuk",
    "energo_saglik": "Energo - saglik kontrol listesi",
    "koc_katilimci": "Koc Universitesi - egitim katilimci listesi",
    "outlook_msg": "Outlook e-postasi (ekleri ayri ayri okunur)",
    "genel": "Taninmayan sablon (genel okuyucu)",
}

# yontem -> kullaniciya gosterilecek kisa Turkce ad.
YONTEM_ADLARI: dict[str, str] = {
    "sicil": "Sicil numarasi",
    "tckn": "TC kimlik no",
    "alias": "Ogrenilmis eslesme",
    "harici": "Bilinen dis kisi",
    "tam_isim": "Tam isim",
    "alt_kume": "Isim alt kumesi",
    "transliterasyon": "Transliterasyon",
    "prefix": "Kesilmis isim (onek)",
    "ek_defter": "Ek kisi defteri",
    "bulanik": "Bulanik benzerlik",
    "aile": "Aile bireyi (tahmin)",
    "yok": "Eslesme yok",
}

DURUM_SIRASI = {DURUM_ESLESMEDI: 0, DURUM_INCELE: 1, DURUM_OTOMATIK: 2}

TABLO_KOLONLARI = [
    "Durum",
    "Kaynak Dosya",
    "Satir",
    "Belge Tarihi",
    "Gider Tipi",
    "Aciklama",
    "Cikarilan Kisi",
    "Sicil",
    "Ad Soyad",
    "Gorev Yeri",
    "Masraf Merkezi",
    "Sirket",
    "Sirket 2",
    "Statu",
    "Kategori",
    "Tutar",
    "Para Birimi",
    "Yontem",
    "Guven",
    "Aday Sayisi",
    "Donem",
    "Gerekce",
    "Uyarilar",
    "Kaynak Masraf Merkezi",
]


# --------------------------------------------------------------------------
# Genel yardimcilar
# --------------------------------------------------------------------------


def _metin(deger: Any) -> str:
    """Herhangi bir degeri gosterime uygun duz metne cevirir."""
    if deger is None:
        return ""
    try:
        if pd.isna(deger):
            return ""
    except (TypeError, ValueError):
        pass
    metin = str(deger).strip()
    if metin.lower() in {"nan", "nat", "none"}:
        return ""
    return metin


def _tarih_metni(deger: Any) -> str:
    """Tarihi gun.ay.yil bicimine cevirir; cozulemezse duz metin dondurur."""
    if isinstance(deger, datetime):
        deger = deger.date()
    if isinstance(deger, date):
        return deger.strftime("%d.%m.%Y")
    return _metin(deger)


def _dosya_imzasi(yol: Path) -> tuple[int, int]:
    """Dosyanin (mtime, boyut) imzasi; onbellek anahtari olarak kullanilir."""
    try:
        bilgi = yol.stat()
        return int(bilgi.st_mtime), int(bilgi.st_size)
    except OSError:
        return (0, 0)


def _hata_goster(baslik: str, hata: BaseException | str) -> None:
    """Kullaniciya anlasilir Turkce hata, teknik ayrinti gizli panelde."""
    st.error(baslik)
    ayrinti = hata if isinstance(hata, str) else "".join(
        traceback.format_exception(type(hata), hata, hata.__traceback__)
    )
    with st.expander("Teknik ayrinti (destek ekibi icin)"):
        st.code(ayrinti, language="text")


# --------------------------------------------------------------------------
# Personel defteri ve ogrenen defterler
# --------------------------------------------------------------------------


@st.cache_resource(show_spinner=False)
def personel_defteri_yukle(yol: str, imza: tuple[int, int]) -> "PersonelDefteri":
    """Personel ana verisini yukler ve surec boyunca bellekte tutar.

    ``imza`` parametresi kullanilmaz ama onbellek anahtarinin parcasidir:
    kaynak dosya degisirse defter otomatik olarak yeniden okunur.
    """
    del imza  # sadece onbellek anahtari
    return PersonelDefteri.yukle(yol, onbellek=True)


@st.cache_resource(show_spinner=False)
def boru_kur(yol: str, imza: tuple[int, int], yardimci_yol: str = "") -> Any:
    """Boru hattini kurar (personel defteri + defterler + harita bir kez yuklenir).

    Ayni Boru ornegi tum oturum boyunca kullanilir; ``hazirla()`` agir
    yuklemeleri yalnizca ilk cagrida yapar. Boru modulu yoksa None doner ve
    arayuz yerel isleme yoluna duser.
    """
    del imza  # sadece onbellek anahtari
    if BORU_SINIFI is None or AYAR_SINIFI is None:
        return None
    ayarlar = AYAR_SINIFI(
        personel_yolu=yol,
        yardimci_personel_yolu=(yardimci_yol or None),
        veri_dizini=str(VERI_KOK),
        cikti_dizini=str(CIKTI_KOK),
    )
    boru = BORU_SINIFI(ayarlar)
    boru.hazirla()
    return boru


def defterleri_al(kok: Path | None = None) -> "Defterler":
    """Ogrenen defterleri (alias / harici / ek kisi / TCKN) dondurur.

    Boru hatti kuruluysa ONUN defter ornegi kullanilir; boylece inceleme
    ekraninda ogrenilen bir kayit ayni anda boru hattinda da gecerli olur.
    """
    boru = st.session_state.get("boru")
    if boru is not None and getattr(boru, "defterler", None) is not None:
        return boru.defterler
    hedef = Path(kok or VERI_KOK)
    mevcut = st.session_state.get("defterler")
    if mevcut is not None and Path(getattr(mevcut, "kok", "")) == hedef:
        return mevcut
    defterler = Defterler(hedef)
    st.session_state["defterler"] = defterler
    return defterler


# --------------------------------------------------------------------------
# Masraf merkezi haritasi
# --------------------------------------------------------------------------

HARITA_KOLONLARI = ("gorev_yeri", "masraf_merkezi_kodu", "masraf_merkezi_adi", "sirket", "aktif")


def harita_yolu(kok: Path | None = None) -> Path:
    """Masraf merkezi haritasi CSV dosyasinin tam yolu."""
    return Path(kok or VERI_KOK) / HARITA_DOSYASI


def _ayirici_bul(yol: Path) -> str:
    """CSV ayiricisini ilk satirdan tahmin eder (';' veya ',')."""
    try:
        ilk = yol.read_text(encoding="utf-8-sig").splitlines()[0]
    except (OSError, IndexError, UnicodeDecodeError):
        return ","
    return ";" if ilk.count(";") > ilk.count(",") else ","


def _kodlama_bul(yol: Path) -> str:
    """Mevcut dosyanin kodlamasini korur (BOM varsa utf-8-sig, yoksa utf-8).

    Excel Turkce karakterleri BOM'lu dosyalarda dogru gosterir; ancak var olan
    bir dosyanin bicimini gereksiz yere degistirmemek icin mevcut hali korunur.
    Yeni dosyalar BOM ile yazilir (defterlerle ayni bicim).
    """
    if not yol.exists():
        return "utf-8-sig"
    try:
        with yol.open("rb") as akis:
            return "utf-8-sig" if akis.read(3) == b"\xef\xbb\xbf" else "utf-8"
    except OSError:
        return "utf-8-sig"


def harita_oku(kok: Path | None = None) -> pd.DataFrame:
    """Masraf merkezi haritasini DataFrame olarak okur; yoksa bos tablo doner."""
    yol = harita_yolu(kok)
    if not yol.exists():
        return pd.DataFrame(columns=list(HARITA_KOLONLARI))
    try:
        df = pd.read_csv(yol, sep=_ayirici_bul(yol), encoding="utf-8-sig", dtype=str)
    except Exception:
        return pd.DataFrame(columns=list(HARITA_KOLONLARI))
    df = df.fillna("")
    for kolon in HARITA_KOLONLARI:
        if kolon not in df.columns:
            df[kolon] = ""
    return df[list(HARITA_KOLONLARI)]


def harita_yaz(df: pd.DataFrame, kok: Path | None = None) -> Path:
    """Masraf merkezi haritasini diske yazar (atomik, gecici dosya uzerinden)."""
    yol = harita_yolu(kok)
    yol.parent.mkdir(parents=True, exist_ok=True)
    ayirici = _ayirici_bul(yol) if yol.exists() else ","
    kodlama = _kodlama_bul(yol)
    temiz = df.fillna("").astype(str)
    for kolon in HARITA_KOLONLARI:
        if kolon not in temiz.columns:
            temiz[kolon] = ""
    temiz = temiz[list(HARITA_KOLONLARI)]
    temiz = temiz[temiz["gorev_yeri"].str.strip() != ""]
    gecici = yol.with_suffix(".csv.tmp")
    temiz.to_csv(gecici, sep=ayirici, index=False, encoding=kodlama,
                 quoting=csv.QUOTE_MINIMAL, lineterminator="\n")
    gecici.replace(yol)
    return yol


def haritayi_tazele() -> None:
    """Boru hattindaki masraf merkezi haritasini diskten yeniden okur.

    Harita duzenlendikten sonra cagrilmalidir; boru hatti haritayi kurulusta
    bir kez yukler, bu cagri olmadan yeni tanimlar devreye girmez.
    """
    boru = st.session_state.get("boru")
    mevcut = getattr(boru, "harita", None) if boru is not None else None
    if mevcut is None:
        return
    try:
        boru.harita = type(mevcut).yukle(harita_yolu())
    except Exception:
        pass  # Harita tazelenemezse eski harita ile calismaya devam et.


def harita_sozlugu(df: pd.DataFrame) -> dict[str, dict[str, str]]:
    """Gorev yeri (normalize) -> masraf merkezi kaydi sozlugu uretir."""
    sozluk: dict[str, dict[str, str]] = {}
    if df is None or df.empty:
        return sozluk
    for _, satir in df.iterrows():
        gorev = _metin(satir.get("gorev_yeri"))
        if not gorev:
            continue
        anahtar = _harita_anahtari(gorev)
        sozluk[anahtar] = {
            "gorev_yeri": gorev,
            "kod": _metin(satir.get("masraf_merkezi_kodu")) or gorev,
            "ad": _metin(satir.get("masraf_merkezi_adi")) or gorev,
            "sirket": _metin(satir.get("sirket")),
            "aktif": _metin(satir.get("aktif")),
        }
    return sozluk


def _harita_anahtari(deger: str) -> str:
    """Gorev yeri adini karsilastirilabilir bicime indirger."""
    try:
        katlanmis = ascii_katla(str(deger))
    except Exception:
        katlanmis = str(deger)
    return " ".join(katlanmis.lower().replace("-", " ").replace("_", " ").split())


# --------------------------------------------------------------------------
# Boru hatti: dosya okuma -> eslestirme -> masraf merkezi atama
# --------------------------------------------------------------------------


@dataclass
class IslemeOzeti:
    """Bir isleme kosusunun ozeti (arayuzde gosterilir)."""

    motor: str = "yerel"
    dosya_sayisi: int = 0
    satir_sayisi: int = 0
    otomatik: int = 0
    incele: int = 0
    eslesmedi: int = 0
    dosya_bilgileri: list[dict] = field(default_factory=list)
    hatalar: list[str] = field(default_factory=list)
    notlar: list[str] = field(default_factory=list)
    boru_ozeti: dict = field(default_factory=dict)

    @property
    def otomatik_orani(self) -> float:
        return (self.otomatik / self.satir_sayisi * 100.0) if self.satir_sayisi else 0.0


@st.cache_data(show_spinner=False)
def _dosya_oku_onbellekli(yol: str, imza: tuple[int, int]) -> list:
    """Bir kaynak dosyayi okur (ayni dosya tekrar okunmaz)."""
    del imza  # sadece onbellek anahtari
    return oku(yol)


@st.cache_data(show_spinner=False)
def dosya_tipi_onbellekli(yol: str, imza: tuple[int, int]) -> str:
    """Dosya tipini tespit eder (kesif modulu)."""
    del imza
    try:
        return dosya_tipini_bul(yol)
    except Exception:
        return "genel"


def _satir_isim_norm(satir: "GiderSatiri") -> str:
    """Bir gider satirindaki kisi adinin normalize halini dondurur."""
    ham = _metin(getattr(satir, "kisi_ham", "") or "")
    norm = isim_normalize(ham) if ham else ""
    if not norm:
        try:
            norm = isim_normalize(kisi_metnini_temizle(satir.aciklama or ""))
        except Exception:
            norm = ""
    return norm


def _masraf_merkezi_coz(
    gorev_yeri: str | None, harita: dict[str, dict[str, str]]
) -> tuple[str | None, bool]:
    """Gorev yerini masraf merkezi koduna cevirir.

    Returns:
        (masraf merkezi kodu, haritada bulundu mu)
    """
    if not gorev_yeri:
        return None, False
    kayit = harita.get(_harita_anahtari(gorev_yeri))
    if kayit:
        return kayit["kod"], True
    return gorev_yeri, False


def sonuc_kur(
    satir: "GiderSatiri",
    eslesme: "Eslesme",
    defter: "PersonelDefteri",
    defterler: "Defterler",
    harita: dict[str, dict[str, str]],
    esik: float,
) -> "Sonuc":
    """Bir gider satiri + eslesmeden nihai ``Sonuc`` kaydi uretir.

    Personel kaydi HER ZAMAN gider tarihine karsilik gelen donemden alinir;
    kisilerin gorev yeri donemler arasinda degisebildigi icin bu sarttir.
    """
    uyarilar: list[str] = []
    donem = gorev_yeri = masraf_merkezi = sirket = sirket2 = statu = kategori = None
    cikis_tarihi = None

    if eslesme.sicil:
        kayit = defter.donem_kaydi(eslesme.sicil, satir.belge_tarihi)
        if kayit is None:
            uyarilar.append("Sicil personel ana verisinde bulunamadi.")
        else:
            donem = kayit.get("donem")
            gorev_yeri = kayit.get("gorev_yeri")
            sirket = kayit.get("sirket")
            sirket2 = kayit.get("sirket2")
            statu = kayit.get("statu")
            kategori = kayit.get("kategori")
            cikis_tarihi = kayit.get("cikis_tarihi")
            if kayit.get("_donem_tahmini"):
                uyarilar.append(
                    "Gider tarihi personel verisindeki donem araliginin disinda; "
                    "en yakin donem kullanildi."
                )
            if kategori == "Cikis" and cikis_tarihi and satir.belge_tarihi:
                if satir.belge_tarihi > cikis_tarihi:
                    uyarilar.append(
                        f"Kisi {_tarih_metni(cikis_tarihi)} tarihinde isten ayrilmis, "
                        "gider bu tarihten sonra."
                    )
            masraf_merkezi, bulundu = _masraf_merkezi_coz(gorev_yeri, harita)
            if gorev_yeri and not bulundu:
                uyarilar.append(
                    f"'{gorev_yeri}' masraf merkezi haritasinda yok; gorev yeri adi kullanildi."
                )
    elif eslesme.yontem == "harici":
        anahtar = isim_normalize(eslesme.ad_soyad or "") or _satir_isim_norm(satir)
        kayit = defterler.harici.get(anahtar) or {}
        sirket = kayit.get("kurum") or None
        masraf_merkezi = kayit.get("masraf_merkezi") or None
        gorev_yeri = kayit.get("kurum") or None
        if not masraf_merkezi:
            uyarilar.append("Dis kisi icin masraf merkezi tanimlanmamis.")

    if not masraf_merkezi and satir.masraf_merkezi_kaynak:
        masraf_merkezi = _metin(satir.masraf_merkezi_kaynak) or None
        if masraf_merkezi:
            uyarilar.append("Masraf merkezi kaynak dosyadan alindi (personelden cozulemedi).")
    elif masraf_merkezi and satir.masraf_merkezi_kaynak:
        kaynak_mm = _harita_anahtari(_metin(satir.masraf_merkezi_kaynak))
        bulunan = _harita_anahtari(masraf_merkezi)
        bulunan_gorev = _harita_anahtari(gorev_yeri or "")
        if kaynak_mm and kaynak_mm not in (bulunan, bulunan_gorev):
            if kaynak_mm not in bulunan and bulunan not in kaynak_mm:
                uyarilar.append(
                    f"Kaynak dosyada '{_metin(satir.masraf_merkezi_kaynak)}' yaziyor, "
                    f"personelden '{masraf_merkezi}' cikti."
                )

    if eslesme.aday_sayisi > 1:
        uyarilar.append(f"{eslesme.aday_sayisi} aday bulundu, tekil kisi secilemedi.")

    durum = durum_belirle(eslesme)
    if durum == DURUM_OTOMATIK and eslesme.guven < esik:
        durum = DURUM_INCELE
        uyarilar.append(f"Guven skoru esigin ({esik:.2f}) altinda.")
    if durum == DURUM_OTOMATIK and not masraf_merkezi:
        durum = DURUM_INCELE
        uyarilar.append("Masraf merkezi belirlenemedi.")

    return Sonuc(
        satir=satir,
        eslesme=eslesme,
        donem=donem,
        gorev_yeri=gorev_yeri,
        masraf_merkezi=masraf_merkezi,
        sirket=sirket,
        sirket2=sirket2,
        statu=statu,
        kategori=kategori,
        cikis_tarihi=cikis_tarihi,
        durum=durum,
        uyarilar=uyarilar,
    )


def sonuclari_uret(
    satirlar: Sequence["GiderSatiri"],
    defter: "PersonelDefteri",
    defterler: "Defterler",
    harita: dict[str, dict[str, str]],
    esik: float,
) -> list["Sonuc"]:
    """Gider satirlarini eslestirip nihai sonuc listesini uretir.

    ``esle_toplu`` kullanilir: sonuc satirlarin dosyadaki sirasindan bagimsizdir
    ve ayni dosyadaki kesin eslesmeler aile bireyi tespitinde kanit olur.
    """
    if not satirlar:
        return []
    eslestirici = Eslestirici(defter, defterler)
    eslesmeler = eslestirici.esle_toplu(list(satirlar))
    return [
        sonuc_kur(satir, eslesme, defter, defterler, harita, esik)
        for satir, eslesme in zip(satirlar, eslesmeler)
    ]


def dosyalari_oku(
    yollar: Sequence[Path], ilerleme: Callable[[float, str], None] | None = None
) -> tuple[list["GiderSatiri"], list[dict], list[str]]:
    """Kaynak dosyalari okur ve gider satirlarini toplar.

    Returns:
        (satirlar, dosya bilgileri, hata mesajlari)
    """
    satirlar: list["GiderSatiri"] = []
    bilgiler: list[dict] = []
    hatalar: list[str] = []
    toplam = max(len(yollar), 1)
    for sira, yol in enumerate(yollar, start=1):
        p = Path(yol)
        if ilerleme:
            ilerleme(sira / (toplam + 1), f"Okunuyor: {p.name}")
        imza = _dosya_imzasi(p)
        tip = dosya_tipi_onbellekli(str(p), imza)
        try:
            dosya_satirlari = _dosya_oku_onbellekli(str(p), imza)
        except Exception as hata:
            hatalar.append(f"{p.name}: okunamadi ({hata})")
            bilgiler.append({"dosya": p.name, "tip": tip, "satir": 0, "durum": "HATA"})
            continue
        satirlar.extend(dosya_satirlari)
        bilgiler.append(
            {"dosya": p.name, "tip": tip, "satir": len(dosya_satirlari), "durum": "OK"}
        )
    return satirlar, bilgiler, hatalar


def _dosya_bilgileri(yollar: Sequence[Path], sonuclar: Sequence["Sonuc"]) -> list[dict]:
    """Her kaynak dosya icin tip ve uretilen satir sayisini ozetler."""
    sayaclar: dict[str, int] = {}
    for sonuc in sonuclar:
        ad = Path(str(sonuc.satir.kaynak_dosya or "")).name
        sayaclar[ad] = sayaclar.get(ad, 0) + 1
    bilgiler: list[dict] = []
    for yol in yollar:
        p = Path(yol)
        adet = sayaclar.get(p.name, 0)
        bilgiler.append(
            {
                "dosya": p.name,
                "tip": dosya_tipi_onbellekli(str(p), _dosya_imzasi(p)),
                "satir": adet,
                "durum": "OK" if adet else "SATIR YOK",
            }
        )
    return bilgiler


def isle(
    yollar: Sequence[Path],
    defter: "PersonelDefteri",
    defterler: "Defterler",
    harita: dict[str, dict[str, str]],
    esik: float,
    ilerleme: Callable[[float, str], None] | None = None,
    boru: Any = None,
) -> tuple[list["Sonuc"], IslemeOzeti]:
    """Secilen dosyalari bastan sona isler ve sonuc listesini dondurur.

    Boru hatti verilmisse (``masraf.boru.Boru``) o kullanilir: tum dosyalar
    once okunur, yardimci listelerden kisi defteri beslenir, eslestirme tek
    seferde yapilir. Boru yoksa ayni adimlarin yerel esdegeri calisir.
    """
    ozet = IslemeOzeti(dosya_sayisi=len(yollar))
    sonuclar: list["Sonuc"] = []

    if boru is not None:
        try:
            # Kullanicinin sectigi esik boru hattina aktarilir.
            try:
                boru.ayarlar.guven_esigi = float(esik)
            except Exception:
                pass

            def _ilerleme_koprusu(yuzde: float, mesaj: str) -> None:
                if ilerleme:
                    ilerleme(float(yuzde) / 100.0, str(mesaj))

            sonuclar = list(boru.isle([str(y) for y in yollar], _ilerleme_koprusu))
            ozet.motor = "boru"
            ozet.hatalar = list(getattr(boru, "hatalar", []) or [])
            ozet.notlar = list(getattr(boru, "uyarilar", []) or [])
            try:
                ozet.boru_ozeti = boru.ozet(sonuclar) if sonuclar else {}
            except Exception:
                ozet.boru_ozeti = {}
            ozet.dosya_bilgileri = _dosya_bilgileri(yollar, sonuclar)
        except Exception as hata:
            sonuclar = []
            ozet.notlar.append(
                f"Boru hatti calistirilamadi, yerel isleme yapildi ({hata})."
            )

    if not sonuclar and ozet.motor != "boru":
        satirlar, bilgiler, hatalar = dosyalari_oku(yollar, ilerleme)
        ozet.dosya_bilgileri = bilgiler
        ozet.hatalar.extend(hatalar)
        if ilerleme:
            ilerleme(0.7, "Kisiler personel verisiyle eslestiriliyor...")
        sonuclar = sonuclari_uret(satirlar, defter, defterler, harita, esik)

    if ilerleme:
        ilerleme(1.0, "Tamamlandi")

    ozet.satir_sayisi = len(sonuclar)
    ozet.otomatik = sum(1 for s in sonuclar if s.durum == DURUM_OTOMATIK)
    ozet.incele = sum(1 for s in sonuclar if s.durum == DURUM_INCELE)
    ozet.eslesmedi = sum(1 for s in sonuclar if s.durum == DURUM_ESLESMEDI)
    return sonuclar, ozet


# --------------------------------------------------------------------------
# Tablolar ve Excel ciktisi
# --------------------------------------------------------------------------


def sonuc_tablosu(sonuclar: Sequence["Sonuc"]) -> pd.DataFrame:
    """Sonuc listesini gosterim ve Excel icin tabloya cevirir."""
    kayitlar: list[dict[str, Any]] = []
    for sonuc in sonuclar:
        satir = sonuc.satir
        eslesme = sonuc.eslesme
        kayitlar.append(
            {
                "Durum": sonuc.durum,
                "Kaynak Dosya": Path(str(satir.kaynak_dosya or "")).name,
                "Satir": satir.satir_no,
                "Belge Tarihi": _tarih_metni(satir.belge_tarihi),
                "Gider Tipi": _metin(satir.gider_tipi),
                "Aciklama": _metin(satir.aciklama),
                "Cikarilan Kisi": _metin(satir.kisi_ham),
                "Sicil": _metin(eslesme.sicil),
                "Ad Soyad": _metin(eslesme.ad_soyad),
                "Gorev Yeri": _metin(sonuc.gorev_yeri),
                "Masraf Merkezi": _metin(sonuc.masraf_merkezi),
                "Sirket": _metin(sonuc.sirket),
                "Sirket 2": _metin(sonuc.sirket2),
                "Statu": _metin(sonuc.statu),
                "Kategori": _metin(sonuc.kategori),
                "Tutar": satir.tutar,
                "Para Birimi": _metin(satir.para_birimi),
                "Yontem": YONTEM_ADLARI.get(eslesme.yontem, eslesme.yontem),
                "Guven": round(float(eslesme.guven or 0.0), 2),
                "Aday Sayisi": int(eslesme.aday_sayisi or 0),
                "Donem": _tarih_metni(sonuc.donem),
                "Gerekce": _metin(eslesme.aciklama),
                "Uyarilar": " | ".join(sonuc.uyarilar or []),
                "Kaynak Masraf Merkezi": _metin(satir.masraf_merkezi_kaynak),
            }
        )
    if not kayitlar:
        return pd.DataFrame(columns=TABLO_KOLONLARI)
    return pd.DataFrame(kayitlar)[TABLO_KOLONLARI]


def ozet_tablosu(sonuclar: Sequence["Sonuc"]) -> pd.DataFrame:
    """Masraf merkezi bazinda ozet (satir sayisi ve tutar toplami)."""
    kayitlar: list[dict[str, Any]] = []
    for sonuc in sonuclar:
        kayitlar.append(
            {
                "Masraf Merkezi": _metin(sonuc.masraf_merkezi) or "(belirlenemedi)",
                "Gorev Yeri": _metin(sonuc.gorev_yeri) or "(yok)",
                "Durum": sonuc.durum,
                "Para Birimi": _metin(sonuc.satir.para_birimi) or "-",
                "Tutar": float(sonuc.satir.tutar or 0.0),
            }
        )
    if not kayitlar:
        return pd.DataFrame(columns=["Masraf Merkezi", "Para Birimi", "Satir", "Tutar"])
    df = pd.DataFrame(kayitlar)
    ozet = (
        df.groupby(["Masraf Merkezi", "Para Birimi"], dropna=False)
        .agg(Satir=("Tutar", "size"), Tutar=("Tutar", "sum"))
        .reset_index()
        .sort_values(["Tutar"], ascending=False)
    )
    return ozet


def _durum_boya(deger: Any) -> str:
    """Durum hucresi icin arka plan rengi (pandas Styler)."""
    renk = {
        DURUM_OTOMATIK: RENK_OTOMATIK,
        DURUM_INCELE: RENK_INCELE,
        DURUM_ESLESMEDI: RENK_ESLESMEDI,
    }.get(str(deger), "")
    return f"background-color: {renk}" if renk else ""


def tabloyu_boya(df: pd.DataFrame):
    """Durum kolonuna gore renklendirilmis tablo dondurur; olmazsa duz tablo."""
    if df.empty or "Durum" not in df.columns:
        return df
    try:
        stil = df.style
        boyayici = getattr(stil, "map", None) or getattr(stil, "applymap")
        return boyayici(_durum_boya, subset=["Durum"])
    except Exception:
        return df


def mahsup_tablosu(sonuclar: Sequence["Sonuc"]) -> Any:
    """Sonuclardan mahsuplasma (dagitim) tablosunu uretir; modul yoksa None.

    Masraf merkezi haritasi ``Boru`` ornegi uzerinden verilir; boylece
    'RHI 1/3 - RENSTROYDETAL 2/3' gibi paylasim etiketlerinin proje mi sirket
    mi oldugu ayirt edilebilir.
    """
    if MAHSUP_MODULU is None or not sonuclar:
        return None
    harita = None
    try:
        boru = st.session_state.get("boru")
        if boru is not None:
            harita = getattr(boru, "harita", None)
    except Exception:
        harita = None
    try:
        return MAHSUP_MODULU.mahsuplasma_uret(list(sonuclar), harita)
    except Exception:
        return None


def excel_uret(
    sonuclar: Sequence["Sonuc"], hedef: Path, boru_ozeti: dict | None = None
) -> Path:
    """Sonuclari cok sayfali Excel dosyasina yazar.

    Once ``masraf.cikti.excel_yaz`` denenir (renkli, bicimli, gerekce
    kolonlariyla birlikte); modul yoksa buradaki yerel yazici kullanilir.
    Cekirdek yazici varsa sayfalar: Mahsuplasma / Kontrol / Sonuc / Incele /
    Eslesmedi / Ozet. Yerel yedek yazici yalnizca Sonuc / Incele / Eslesmedi
    uretir.
    """
    hedef = Path(hedef)
    hedef.parent.mkdir(parents=True, exist_ok=True)

    yazici = getattr(CIKTI_MODULU, "excel_yaz", None) if CIKTI_MODULU else None
    if callable(yazici):
        try:
            yazici(list(sonuclar), str(hedef), dict(boru_ozeti or {}),
                   mahsup_tablosu(sonuclar))
            if hedef.exists():
                return hedef
        except Exception:
            pass  # Yerel yaziciya dus.

    tablo = sonuc_tablosu(sonuclar)
    sayfalar = {
        "Sonuc": tablo[tablo["Durum"] == DURUM_OTOMATIK] if not tablo.empty else tablo,
        "Incele": tablo[tablo["Durum"] == DURUM_INCELE] if not tablo.empty else tablo,
        "Eslesmedi": tablo[tablo["Durum"] == DURUM_ESLESMEDI] if not tablo.empty else tablo,
    }
    with pd.ExcelWriter(hedef, engine="xlsxwriter") as yazici:
        kitap = yazici.book
        baslik_bicimi = kitap.add_format(
            {
                "bold": True,
                "bg_color": RENK_LACIVERT,
                "font_color": "#FFFFFF",
                "border": 1,
                "align": "left",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        for ad, alt in sayfalar.items():
            alt = alt if alt is not None else pd.DataFrame(columns=TABLO_KOLONLARI)
            alt.to_excel(yazici, sheet_name=ad, index=False, startrow=1, header=False)
            sayfa = yazici.sheets[ad]
            kolonlar = list(alt.columns) if len(alt.columns) else TABLO_KOLONLARI
            for sutun, kolon_adi in enumerate(kolonlar):
                sayfa.write(0, sutun, kolon_adi, baslik_bicimi)
                genislik = 14
                if kolon_adi in ("Aciklama", "Gerekce", "Uyarilar"):
                    genislik = 52
                elif kolon_adi in ("Ad Soyad", "Cikarilan Kisi", "Gorev Yeri", "Masraf Merkezi"):
                    genislik = 26
                sayfa.set_column(sutun, sutun, genislik)
            if len(alt):
                sayfa.autofilter(0, 0, len(alt), max(len(kolonlar) - 1, 0))
            sayfa.freeze_panes(1, 0)

        ozet = ozet_tablosu(sonuclar)
        ozet.to_excel(yazici, sheet_name="Ozet", index=False, startrow=1, header=False)
        sayfa = yazici.sheets["Ozet"]
        for sutun, kolon_adi in enumerate(list(ozet.columns) or ["Masraf Merkezi"]):
            sayfa.write(0, sutun, kolon_adi, baslik_bicimi)
            sayfa.set_column(sutun, sutun, 28 if sutun == 0 else 14)
        satir_no = len(ozet) + 3
        toplam = len(sonuclar)
        otomatik = sum(1 for s in sonuclar if s.durum == DURUM_OTOMATIK)
        incele = sum(1 for s in sonuclar if s.durum == DURUM_INCELE)
        eslesmedi = sum(1 for s in sonuclar if s.durum == DURUM_ESLESMEDI)
        oran = (otomatik / toplam * 100.0) if toplam else 0.0
        for etiket, deger in (
            ("Toplam satir", toplam),
            ("Otomatik eslesen", otomatik),
            ("Incelenecek", incele),
            ("Eslesmeyen", eslesmedi),
            ("Otomatik oran (%)", round(oran, 1)),
            ("Uretim tarihi", datetime.now().strftime("%d.%m.%Y %H:%M")),
        ):
            sayfa.write(satir_no, 0, etiket)
            sayfa.write(satir_no, 1, deger)
            satir_no += 1
    return hedef


# --------------------------------------------------------------------------
# Personel arama (inceleme ekrani icin)
# --------------------------------------------------------------------------


@st.cache_resource(show_spinner=False)
def _arama_indeksi(_defter: "PersonelDefteri", imza: str) -> dict[str, list[str]]:
    """Normalize isim -> sicil listesi haritasi (arama icin)."""
    del imza  # sadece onbellek anahtari
    try:
        return Eslestirici._isim_haritasi(_defter)
    except Exception:
        return dict(getattr(_defter, "_isim_index", {}) or {})


def personel_ara(defter: "PersonelDefteri", sorgu: str, azami: int = 25) -> list[str]:
    """Serbest metinle personel arar; sicil listesi dondurur.

    Once sicil numarasi, sonra birebir isim, sonra token kesisimi ve bulanik
    benzerlik denenir. Tamamen yerel calisir (rapidfuzz).
    """
    sorgu = _metin(sorgu)
    if not sorgu:
        return []

    # 1) Dogrudan sicil
    kanonik = sicil_normalize(sorgu)
    if kanonik and defter.sicil_ile(kanonik):
        return [kanonik]

    norm = isim_normalize(sorgu)
    if not norm:
        return []

    bulunan: list[str] = []
    gorulen: set[str] = set()

    def ekle(siciller: Iterable[str]) -> None:
        for sicil in siciller:
            if sicil and sicil not in gorulen:
                gorulen.add(sicil)
                bulunan.append(sicil)

    ekle(defter.isimle_adaylar(norm))
    ekle(defter.token_ile_adaylar(frozenset(norm.split(" "))))

    harita = _arama_indeksi(defter, str(getattr(defter, "_kaynak_yol", "")))
    tokenlar = set(norm.split(" "))
    icerenler = [isim for isim in harita if tokenlar & set(isim.split(" "))]
    if not icerenler:
        icerenler = [isim for isim in harita if norm in isim]
    for isim in sorted(icerenler, key=len)[: azami * 4]:
        ekle(harita[isim])

    if len(bulunan) < azami:
        try:
            from rapidfuzz import process

            adaylar = process.extract(norm, list(harita.keys()), limit=azami, score_cutoff=70)
            for isim, _puan, _ in adaylar:
                ekle(harita[isim])
        except Exception:
            pass

    return bulunan[:azami]


def sicil_etiketi(defter: "PersonelDefteri", sicil: str) -> str:
    """Bir sicil icin acilir listede gosterilecek aciklayici etiket."""
    kayit = defter.sicil_ile(sicil) or {}
    parcalar = [f"{sicil}"]
    ad = _metin(kayit.get("ad_soyad"))
    if ad:
        parcalar.append(ad)
    gorev = _metin(kayit.get("gorev_yeri"))
    if gorev:
        parcalar.append(gorev)
    sirket = _metin(kayit.get("sirket"))
    if sirket:
        parcalar.append(sirket)
    kategori = _metin(kayit.get("kategori"))
    if kategori:
        parcalar.append(kategori)
    return " - ".join(parcalar)


# --------------------------------------------------------------------------
# Arayuz: ortak parcalar
# --------------------------------------------------------------------------

STIL = f"""
<style>
  .blok-baslik {{
      color: {RENK_LACIVERT};
      font-size: 1.35rem;
      font-weight: 700;
      border-bottom: 3px solid {RENK_MAVI};
      padding-bottom: 0.3rem;
      margin: 0.2rem 0 0.9rem 0;
  }}
  .ust-serit {{
      background: linear-gradient(90deg, {RENK_LACIVERT} 0%, {RENK_MAVI} 100%);
      color: #FFFFFF;
      padding: 0.9rem 1.2rem;
      border-radius: 8px;
      margin-bottom: 1.1rem;
  }}
  .ust-serit h1 {{ margin: 0; font-size: 1.5rem; color: #FFFFFF; }}
  .ust-serit p  {{ margin: 0.25rem 0 0 0; font-size: 0.9rem; opacity: 0.9; }}
  .kart {{
      background: {RENK_ACIK};
      border-left: 5px solid {RENK_MAVI};
      padding: 0.7rem 1rem;
      border-radius: 6px;
      margin-bottom: 0.6rem;
  }}
  .rozet {{
      display: inline-block; padding: 2px 10px; border-radius: 10px;
      font-size: 0.8rem; font-weight: 600; color: #1a1a1a;
  }}
  div[data-testid="stMetricValue"] {{ color: {RENK_LACIVERT}; }}
  .stTabs [data-baseweb="tab"] {{ font-weight: 600; }}
</style>
"""


def blok_baslik(metin: str) -> None:
    """Sekme icinde bolum basligi cizer."""
    st.markdown(f'<div class="blok-baslik">{metin}</div>', unsafe_allow_html=True)


def durum_rozeti(durum: str) -> str:
    """Durum icin renkli HTML rozeti dondurur."""
    renk = {
        DURUM_OTOMATIK: RENK_OTOMATIK,
        DURUM_INCELE: RENK_INCELE,
        DURUM_ESLESMEDI: RENK_ESLESMEDI,
    }.get(durum, "#EEEEEE")
    etiket = {
        DURUM_OTOMATIK: "OTOMATİK",
        DURUM_INCELE: "İNCELE",
        DURUM_ESLESMEDI: "EŞLEŞMEDİ",
    }.get(durum, durum)
    return f'<span class="rozet" style="background:{renk}">{etiket}</span>'


def _oturum_hazirla() -> None:
    """Session state varsayilanlarini kurar."""
    varsayilanlar = {
        "personel_yolu": str(VARSAYILAN_PERSONEL),
        "yardimci_personel_yolu": str(VARSAYILAN_YARDIMCI) if VARSAYILAN_YARDIMCI.exists() else "",
        "esik": 0.90,
        "defter": None,
        "boru": None,
        "sonuclar": [],
        "ozet": None,
        "yollar": [],
        "inceleme_sirasi": 0,
        "son_excel": None,
        "ogrenilen": 0,
    }
    for anahtar, deger in varsayilanlar.items():
        st.session_state.setdefault(anahtar, deger)
    if "yukleme_dizini" not in st.session_state:
        st.session_state["yukleme_dizini"] = tempfile.mkdtemp(prefix="masraf_yukleme_")


def _defter_var_mi() -> bool:
    """Personel defteri yuklu mu? Degilse kullaniciyi yonlendirir."""
    if st.session_state.get("defter") is None:
        st.warning(
            "Once **Ayarlar** sekmesinden personel ana verisini yükleyin. "
            "Eşleştirme bu veri olmadan yapılamaz."
        )
        return False
    return True


def _yeniden_isle() -> None:
    """Kayitli dosya listesini mevcut defterlerle yeniden isler."""
    yollar = st.session_state.get("yollar") or []
    defter = st.session_state.get("defter")
    if not yollar or defter is None:
        return
    defterler = defterleri_al()
    harita = harita_sozlugu(harita_oku())
    sonuclar, ozet = isle(
        [Path(y) for y in yollar],
        defter,
        defterler,
        harita,
        float(st.session_state["esik"]),
        boru=st.session_state.get("boru"),
    )
    st.session_state["sonuclar"] = sonuclar
    st.session_state["ozet"] = ozet


# --------------------------------------------------------------------------
# Sekme 1: Ayarlar
# --------------------------------------------------------------------------


def sekme_ayarlar() -> None:
    """Personel ana verisi, guven esigi ve defter durumu."""
    blok_baslik("1. Personel ana verisi")
    st.caption(
        "Aylık personel snapshot dosyası (giriş/çıkış). Her kişinin her dönem için bir satırı vardır; "
        "masraf merkezi, giderin tarihine karşılık gelen dönemden okunur."
    )

    sutun_yol, sutun_buton = st.columns([4, 1])
    with sutun_yol:
        yol_metni = st.text_input(
            "Dosya yolu",
            value=st.session_state["personel_yolu"],
            help="Örnek: ornek_veri/personel/2025_2026_giris_cikis.xlsx",
            key="personel_yolu_girdi",
        )
    with sutun_buton:
        st.write("")
        st.write("")
        yukle_tiklandi = st.button("Personel verisini yükle", type="primary", width="stretch")

    aday = Path(yol_metni).expanduser()
    if not aday.is_absolute():
        aday = (PROJE_KOK / aday).resolve()

    if aday.exists():
        boyut = aday.stat().st_size / (1024 * 1024)
        st.caption(f"Dosya bulundu: `{aday}` ({boyut:.1f} MB)")
    else:
        st.caption(f"Dosya bulunamadı: `{aday}`")

    st.markdown("**Grup şirketleri listesi (isteğe bağlı)**")
    st.caption(
        "Ana personel dosyası sadece RHI ve UST LUGA tüzel kişilerini kapsar. "
        "Renservis, Renstroydetal, RC, One Tower ve Top Tower personeli ancak "
        "1C personel listesinde bulunur. Bu dosyayı da verirseniz masraf merkezi "
        "çözülen satır oranı ölçülen örnekte yüzde 89'dan yüzde 98'e çıkıyor."
    )
    yardimci_metni = st.text_input(
        "1C personel listesi yolu",
        value=st.session_state.get("yardimci_personel_yolu", ""),
        help="Örnek: ornek_veri/personel/1C_Personnel_List_31082026.xlsx. Boş bırakabilirsiniz.",
        key="yardimci_personel_yolu_girdi",
    )
    st.session_state["yardimci_personel_yolu"] = yardimci_metni
    if yardimci_metni:
        _y = Path(yardimci_metni).expanduser()
        if not _y.is_absolute():
            _y = (PROJE_KOK / _y).resolve()
        st.caption(
            f"1C listesi bulundu: `{_y}`" if _y.exists()
            else f"1C listesi bulunamadı: `{_y}` (bu dosya olmadan da çalışır)"
        )

    if yukle_tiklandi:
        if not aday.exists():
            st.error(
                "Belirtilen dosya bulunamadı. Yolu kontrol edin veya dosyayı "
                "`ornek_veri/personel/` klasörüne kopyalayın."
            )
        else:
            try:
                with st.spinner("Personel verisi okunuyor (ilk okumada ~30 saniye sürebilir)..."):
                    imza = _dosya_imzasi(aday)
                    yardimci_metni = (st.session_state.get("yardimci_personel_yolu") or "").strip()
                    yardimci_aday = ""
                    if yardimci_metni:
                        y = Path(yardimci_metni).expanduser()
                        if not y.is_absolute():
                            y = (PROJE_KOK / y).resolve()
                        yardimci_aday = str(y) if y.exists() else ""
                    boru = boru_kur(str(aday), imza, yardimci_aday)
                    defter = (
                        boru.defter
                        if boru is not None and getattr(boru, "defter", None) is not None
                        else personel_defteri_yukle(str(aday), imza)
                    )
                st.session_state["boru"] = boru
                st.session_state["defter"] = defter
                st.session_state["personel_yolu"] = str(aday)
                st.success("Personel verisi yüklendi. Bu oturumda tekrar okunmayacak.")
                for uyari in list(getattr(boru, "uyarilar", []) or [])[:5]:
                    st.warning(uyari)
            except Exception as hata:
                _hata_goster(
                    "Personel verisi okunamadı. Dosyanın Excel formatında ve açık "
                    "olmadığından emin olun.",
                    hata,
                )

    defter = st.session_state.get("defter")
    if defter is not None:
        try:
            bilgi = defter.istatistik()
        except Exception as hata:
            _hata_goster("Personel verisi özeti çıkarılamadı.", hata)
            bilgi = {}
        if bilgi:
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Toplam satır", f"{bilgi.get('satir_sayisi', 0):,}".replace(",", "."))
            k2.metric("İsimli personel", f"{bilgi.get('isimli_sicil', 0):,}".replace(",", "."))
            k3.metric("Dönem sayısı", bilgi.get("donem_sayisi", 0))
            k4.metric("Görev yeri", bilgi.get("gorev_yeri_sayisi", 0))

            ilk, son = bilgi.get("ilk_donem"), bilgi.get("son_donem")
            st.markdown(
                f'<div class="kart">Dönem aralığı: <b>{_tarih_metni(ilk)} - {_tarih_metni(son)}</b>'
                f' &nbsp;|&nbsp; İsimsiz (bordrosuz taşeron) kayıt: '
                f'<b>{bilgi.get("isimsiz_satir", 0):,}</b>'.replace(",", ".")
                + f' &nbsp;|&nbsp; Çakışan isim oranı: <b>%{bilgi.get("cakisma_orani", 0)}</b></div>',
                unsafe_allow_html=True,
            )
            st.caption(
                "İsimsiz kayıtlar 'Bordrosuz Taşeron' satırlarıdır; isim eşleştirmede kullanılmaz. "
                "Çakışan isim oranı, aynı ada sahip birden fazla personelin bulunma oranıdır."
            )

    st.divider()
    blok_baslik("2. Güven eşiği")
    esik = st.slider(
        "Otomatik kabul için asgari güven skoru",
        min_value=0.50,
        max_value=1.00,
        value=float(st.session_state["esik"]),
        step=0.01,
        help="Bu skorun altındaki eşleşmeler otomatik kabul edilmez, İnceleme sekmesine düşer.",
    )
    if abs(esik - float(st.session_state["esik"])) > 1e-9:
        st.session_state["esik"] = esik
        if st.session_state.get("sonuclar"):
            _yeniden_isle()
            st.info("Eşik değişti, mevcut sonuçlar yeniden değerlendirildi.")
    st.caption(
        f"Motorun kendi alt sınırı %{INCELE_ESIGI * 100:.0f}'dir. Buradaki eşik daha da sıkı "
        "olabilir: yüksek eşik = daha az otomatik, daha çok kontrol."
    )

    st.divider()
    blok_baslik("3. Öğrenen defterler")
    st.caption(
        "Sistem çalışma zamanında yapay zeka kullanmaz. Öğrenme, İnceleme sekmesinde "
        "verdiğiniz kararların bu dosyalara yazılmasıyla olur."
    )
    try:
        defterler = defterleri_al()
        d = defterler.istatistik()
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Öğrenilmiş eşleşme", d.get("alias", 0))
        d2.metric("Dış (harici) kişi", d.get("harici", 0))
        d3.metric("Ek kişi kaydı", d.get("ek_kisi", 0))
        d4.metric("TCKN köprüsü", d.get("tckn_kopru", 0))
        st.caption(f"Defter klasörü: `{d.get('kok', '')}`")
        if d.get("kaydedilmemis"):
            st.warning("Kaydedilmemiş değişiklik var: " + ", ".join(d["kaydedilmemis"]))
        if st.button("Defterleri diskten yeniden oku"):
            defterler.yeniden_yukle()
            st.success("Defterler yeniden okundu.")
    except Exception as hata:
        _hata_goster("Öğrenen defterler okunamadı.", hata)

    st.divider()
    blok_baslik("4. Sistem durumu")
    s1, s2 = st.columns(2)
    with s1:
        st.write("**Çekirdek modüller**")
        st.write("Eşleştirme motoru: " + ("hazır" if CEKIRDEK_HAZIR else "EKSİK"))
        st.write("Boru hattı modülü: " + ("var" if BORU_SINIFI is not None else "yok (yerel işleme)"))
        st.write("Çıktı modülü: " + ("var" if CIKTI_MODULU is not None else "yok (yerel Excel)"))
    with s2:
        st.write("**Klasörler**")
        st.write(f"Veri: `{VERI_KOK}`")
        st.write(f"Çıktı: `{CIKTI_KOK}`")
        st.write(f"Geçici yüklemeler: `{st.session_state['yukleme_dizini']}`")


# --------------------------------------------------------------------------
# Sekme 2: Fatura Isle
# --------------------------------------------------------------------------


def _yuklenenleri_kaydet(dosyalar: Sequence[Any]) -> list[Path]:
    """Tarayicidan yuklenen dosyalari gecici klasore yazar ve yollarini dondurur."""
    hedef_dizin = Path(st.session_state["yukleme_dizini"])
    hedef_dizin.mkdir(parents=True, exist_ok=True)
    yollar: list[Path] = []
    for dosya in dosyalar:
        hedef = hedef_dizin / dosya.name
        try:
            hedef.write_bytes(dosya.getbuffer())
            yollar.append(hedef)
        except OSError as hata:
            st.error(f"'{dosya.name}' geçici klasöre yazılamadı: {hata}")
    return yollar


def _klasorden_topla(klasor: str) -> list[Path]:
    """Bir klasordeki desteklenen dosyalari listeler."""
    kok = Path(klasor).expanduser()
    if not kok.is_absolute():
        kok = (PROJE_KOK / kok).resolve()
    if not kok.is_dir():
        return []
    return sorted(
        p
        for p in kok.iterdir()
        if p.is_file()
        and p.suffix.lower() in DESTEKLENEN_UZANTILAR
        and not p.name.startswith("~$")
    )


def sekme_fatura() -> None:
    """Fatura dosyalarini yukle, tipini gor, isle ve ciktiyi indir."""
    blok_baslik("Fatura dosyalarını seçin")
    if not _defter_var_mi():
        return

    kaynak_secim = st.radio(
        "Dosyaları nasıl vereceksiniz?",
        ["Dosya yükle", "Klasör yolu"],
        horizontal=True,
        key="kaynak_secim",
    )

    yollar: list[Path] = []
    if kaynak_secim == "Dosya yükle":
        yuklenenler = st.file_uploader(
            "Fatura / liste dosyaları (xls, xlsx, csv, Outlook .msg)",
            type=["xls", "xlsx", "xlsm", "csv", "msg"],
            accept_multiple_files=True,
            help="Birden fazla dosya seçebilirsiniz. Dosyalar sadece bu bilgisayarda işlenir.",
        )
        if yuklenenler:
            yollar = _yuklenenleri_kaydet(yuklenenler)
    else:
        klasor = st.text_input(
            "Klasör yolu",
            value=str(PROJE_KOK / "ornek_veri" / "antik_travel"),
            help="Klasördeki tüm xls/xlsx/csv/msg dosyaları işlenir.",
        )
        yollar = _klasorden_topla(klasor)
        if klasor and not yollar:
            st.info("Bu klasörde işlenebilir dosya bulunamadı.")

    if yollar:
        blok_baslik("Tespit edilen dosya tipleri")
        satirlar = []
        for yol in yollar:
            tip = dosya_tipi_onbellekli(str(yol), _dosya_imzasi(yol))
            satirlar.append(
                {
                    "Dosya": yol.name,
                    "Tespit edilen tip": KAYNAK_TIP_ADLARI.get(tip, tip),
                    "Kod": tip,
                    "Boyut (KB)": round(yol.stat().st_size / 1024, 1),
                }
            )
        st.dataframe(pd.DataFrame(satirlar), hide_index=True, width="stretch")
        if any(s["Kod"] == "genel" for s in satirlar):
            st.info(
                "'Tanınmayan şablon' olarak işaretlenen dosyalar genel okuyucuyla işlenir: "
                "isim, tarih ve tutar kolonları otomatik aranır. Sonuçları İnceleme "
                "sekmesinden kontrol edin."
            )

    islet = st.button(
        "İşle", type="primary", disabled=not yollar, width="stretch"
    )

    if islet and yollar:
        cubuk = st.progress(0.0, text="Başlıyor...")

        def ilerleme(oran: float, metin: str) -> None:
            cubuk.progress(min(max(oran, 0.0), 1.0), text=metin)

        try:
            defterler = defterleri_al()
            harita = harita_sozlugu(harita_oku())
            sonuclar, ozet = isle(
                yollar,
                st.session_state["defter"],
                defterler,
                harita,
                float(st.session_state["esik"]),
                ilerleme,
                boru=st.session_state.get("boru"),
            )
            st.session_state["sonuclar"] = sonuclar
            st.session_state["ozet"] = ozet
            st.session_state["yollar"] = [str(y) for y in yollar]
            st.session_state["inceleme_sirasi"] = 0
            st.session_state["son_excel"] = None
            cubuk.empty()
        except Exception as hata:
            cubuk.empty()
            _hata_goster(
                "Dosyalar işlenirken bir hata oluştu. Dosya biçimi beklenenden farklı olabilir.",
                hata,
            )
            return

    sonuclar = st.session_state.get("sonuclar") or []
    ozet: IslemeOzeti | None = st.session_state.get("ozet")
    if not sonuclar:
        st.caption("Henüz işlenmiş satır yok. Dosya seçip **İşle** düğmesine basın.")
        return

    st.divider()
    blok_baslik("Sonuç")
    toplam = len(sonuclar)
    otomatik = sum(1 for s in sonuclar if s.durum == DURUM_OTOMATIK)
    incele = sum(1 for s in sonuclar if s.durum == DURUM_INCELE)
    eslesmedi = sum(1 for s in sonuclar if s.durum == DURUM_ESLESMEDI)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Toplam satır", toplam)
    m2.metric("Otomatik", otomatik, f"%{otomatik / toplam * 100:.1f}" if toplam else "")
    m3.metric("İncelenecek", incele, f"%{incele / toplam * 100:.1f}" if toplam else "")
    m4.metric("Eşleşmedi", eslesmedi, f"%{eslesmedi / toplam * 100:.1f}" if toplam else "")

    if ozet is not None:
        if ozet.hatalar:
            for mesaj in ozet.hatalar:
                st.warning(mesaj)
        if ozet.notlar:
            for mesaj in ozet.notlar:
                st.caption(mesaj)

    if incele or eslesmedi:
        st.info(
            f"**{incele + eslesmedi}** satır kontrol bekliyor. **İnceleme** sekmesinden "
            "tek tek çözebilirsiniz; verdiğiniz her karar sisteme kalıcı olarak öğretilir."
        )

    tablo = sonuc_tablosu(sonuclar)

    f1, f2, f3 = st.columns([2, 2, 3])
    with f1:
        durum_filtre = st.multiselect(
            "Durum filtresi",
            [DURUM_OTOMATIK, DURUM_INCELE, DURUM_ESLESMEDI],
            default=[DURUM_OTOMATIK, DURUM_INCELE, DURUM_ESLESMEDI],
        )
    with f2:
        dosya_filtre = st.multiselect(
            "Dosya filtresi",
            sorted(tablo["Kaynak Dosya"].unique()),
            default=sorted(tablo["Kaynak Dosya"].unique()),
        )
    with f3:
        arama = st.text_input("Tabloda ara (isim, açıklama, masraf merkezi)")

    gorunum = tablo[tablo["Durum"].isin(durum_filtre) & tablo["Kaynak Dosya"].isin(dosya_filtre)]
    if arama.strip():
        desen = arama.strip()
        maske = (
            gorunum["Aciklama"].str.contains(desen, case=False, na=False)
            | gorunum["Cikarilan Kisi"].str.contains(desen, case=False, na=False)
            | gorunum["Ad Soyad"].str.contains(desen, case=False, na=False)
            | gorunum["Masraf Merkezi"].str.contains(desen, case=False, na=False)
        )
        gorunum = gorunum[maske]

    st.dataframe(tabloyu_boya(gorunum), hide_index=True, width="stretch", height=420)
    st.caption(f"{len(gorunum)} / {toplam} satır gösteriliyor.")

    with st.expander("Masraf merkezi özeti"):
        st.dataframe(ozet_tablosu(sonuclar), hide_index=True, width="stretch")

    boru_ozeti = (ozet.boru_ozeti if ozet is not None else {}) or {}
    if boru_ozeti:
        with st.expander("İşleme özeti (yöntem dağılımı, eksik tanımlar)"):
            y1, y2 = st.columns(2)
            with y1:
                st.write("**Eşleştirme yöntemi dağılımı**")
                dagilim = boru_ozeti.get("yontem_dagilimi") or {}
                if dagilim:
                    st.dataframe(
                        pd.DataFrame(
                            [
                                {"Yöntem": YONTEM_ADLARI.get(k, k), "Satır": v}
                                for k, v in sorted(
                                    dagilim.items(), key=lambda x: -x[1]
                                )
                            ]
                        ),
                        hide_index=True,
                        width="stretch",
                    )
                paralar = boru_ozeti.get("para_birimi_toplamlari") or {}
                if paralar:
                    st.write("**Para birimi toplamları**")
                    for para, tutar in paralar.items():
                        st.write(f"{para}: {float(tutar):,.2f}".replace(",", " "))
            with y2:
                eksik = boru_ozeti.get("eksik_masraf_merkezleri") or []
                st.write("**Haritada tanımsız masraf merkezleri**")
                if eksik:
                    st.warning(", ".join(str(e) for e in eksik[:25]))
                    st.caption(
                        "Bunları **Masraf Merkezi Haritası** sekmesinden tanımlayın."
                    )
                else:
                    st.success("Tüm görev yerleri haritada tanımlı.")
                eslesmeyen = boru_ozeti.get("eslesmeyen_kisiler") or {}
                if eslesmeyen:
                    st.write("**En sık eşleşmeyen kişiler**")
                    st.dataframe(
                        pd.DataFrame(
                            [
                                {"Kişi": k, "Satır": v}
                                for k, v in sorted(
                                    eslesmeyen.items(), key=lambda x: -x[1]
                                )[:15]
                            ]
                        ),
                        hide_index=True,
                        width="stretch",
                    )

    st.divider()
    blok_baslik("Excel çıktısı")
    sutun_a, sutun_b = st.columns([1, 3])
    with sutun_a:
        uret = st.button("Excel oluştur", width="stretch")
    if uret:
        try:
            CIKTI_KOK.mkdir(parents=True, exist_ok=True)
            damga = datetime.now().strftime("%Y%m%d_%H%M%S")
            hedef = CIKTI_KOK / f"masraf_merkezi_{damga}.xlsx"
            boru_ozeti = ozet.boru_ozeti if ozet is not None else {}
            with st.spinner("Excel dosyası hazırlanıyor..."):
                excel_uret(sonuclar, hedef, boru_ozeti)
            st.session_state["son_excel"] = str(hedef)
        except Exception as hata:
            _hata_goster("Excel dosyası oluşturulamadı.", hata)

    son_excel = st.session_state.get("son_excel")
    if son_excel and Path(son_excel).exists():
        veri = Path(son_excel).read_bytes()
        with sutun_b:
            st.download_button(
                "Excel indir",
                data=veri,
                file_name=Path(son_excel).name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
            )
        st.caption(f"Dosya ayrıca şuraya kaydedildi: `{son_excel}`")


# --------------------------------------------------------------------------
# Sekme 3: Mahsuplasma (muhasebeye giden tablo)
# --------------------------------------------------------------------------


def _mahsup_dataframe(mahsup: Any) -> "pd.DataFrame":
    """Mahsup satirlarini ekranda gosterilecek tabloya cevirir."""
    if CIKTI_MODULU is None:
        return pd.DataFrame()
    fatura_toplami: dict[tuple[str, str], float] = {}
    for m in mahsup.satirlar:
        anahtar = (m.kaynak, m.para_birimi)
        fatura_toplami[anahtar] = fatura_toplami.get(anahtar, 0.0) + m.tutar
    basliklar = [baslik for baslik, _t, _g in CIKTI_MODULU.MAHSUP_KOLONLARI]
    satirlar = [
        CIKTI_MODULU.mahsup_satir_degerleri(m, fatura_toplami[(m.kaynak, m.para_birimi)])
        for m in mahsup.satirlar
    ]
    return pd.DataFrame(satirlar, columns=basliklar)


def _kontrol_dataframe(mahsup: Any) -> "pd.DataFrame":
    if CIKTI_MODULU is None:
        return pd.DataFrame()
    basliklar = [baslik for baslik, _t, _g in CIKTI_MODULU.KONTROL_KOLONLARI]
    satirlar = [CIKTI_MODULU.kontrol_satir_degerleri(k) for k in mahsup.kontrol]
    return pd.DataFrame(satirlar, columns=basliklar)


def sekme_mahsuplasma() -> None:
    """Her fatura icin hangi projeye ne kadar yazilacagini gosterir.

    Bu sekme finansin muhasebeye gonderecegi tablodur. Ustte mutabakat
    vardir: para kaybolmadigi burada gorulur. Altta dagitim tablosu.
    """
    st.subheader("Mahsuplaşma")
    st.caption(
        "Her fatura için hangi projeye ne kadar yazılacağı. Muhasebeye giden "
        "tablo budur; alttaki satır dökümü bunun dayanağıdır."
    )

    sonuclar = st.session_state.get("sonuclar") or []
    if not sonuclar:
        st.info("Önce **Fatura İşle** sekmesinden dosyaları işleyin.")
        return
    if MAHSUP_MODULU is None:
        st.error(
            "Mahsuplaşma modülü yüklenemedi (`masraf/mahsuplasma.py`). "
            "Kurulum eksik olabilir."
        )
        return

    mahsup = mahsup_tablosu(sonuclar)
    if mahsup is None:
        st.error("Mahsuplaşma tablosu üretilemedi.")
        return

    # --- Mutabakat: para kayboldu mu? ---
    blok_baslik("Mutabakat")
    if mahsup.kapali_mi:
        st.success(
            "Bütün faturalar kapandı. Okunan tutar = yinelenen + dağıtılan + "
            "dağıtılamayan. Para kaybolmadı."
        )
    else:
        st.error(
            "**Mutabakat açık.** Aşağıdaki faturalarda dağıtım toplamı okunan "
            "tutara eşit değil. Bu tablo muhasebeye gönderilmemeli: "
            + ", ".join(f"{k.kaynak} ({k.fark:+.2f})" for k in mahsup.acik_kontroller)
        )

    toplamlar = mahsup.toplamlar()
    for para, deger in toplamlar.items():
        sutunlar = st.columns(5)
        sutunlar[0].metric(f"Okunan ({para})", f"{deger['gelen']:,.2f}")
        sutunlar[1].metric(
            "Yinelenen", f"{deger['yinelenen']:,.2f}",
            help="Aynı işlem başka bir dosyada zaten sayıldı; çift saymamak "
                 "için düşüldü.",
        )
        sutunlar[2].metric(
            "Net", f"{deger['net']:,.2f}",
            help="Okunan eksi yinelenen. Gerçekten dağıtılacak tutar.",
        )
        sutunlar[3].metric("Dağıtılan", f"{deger['dagitilan']:,.2f}")
        sutunlar[4].metric(
            "Dağıtılamayan", f"{deger['dagitilamayan']:,.2f}",
            help="Kişi veya masraf merkezi bulunamadı. Silinmedi, tabloda "
                 "'(DAGITILAMAYAN)' satırı olarak duruyor.",
        )

    if mahsup.yinelenen_sayisi:
        st.info(
            f"**{mahsup.yinelenen_sayisi} satır yinelenen olarak düşüldü.** "
            "Aynı mailde hem acentenin ham dökümü hem de o işlemlerin elle "
            "dağıtılmış hali gelirse para iki kez sayılır; bu eleme onu "
            "engeller. Ayrıntı aşağıdaki kontrol tablosunda."
        )
    for celiski in mahsup.isaret_celiskileri:
        st.warning(celiski.aciklama())

    kontrol_df = _kontrol_dataframe(mahsup)
    if not kontrol_df.empty:
        with st.expander("Fatura bazında kontrol tablosu", expanded=not mahsup.kapali_mi):
            st.dataframe(kontrol_df, hide_index=True, width="stretch")
            st.caption(
                "**Fark** sütunu sıfır olmak zorundadır. Sıfır değilse "
                "dağıtımda kayıp var demektir."
            )

    if mahsup.kutuk_satir_sayisi or mahsup.tutarsiz_satir_sayisi:
        notlar = []
        if mahsup.kutuk_satir_sayisi:
            notlar.append(
                f"{mahsup.kutuk_satir_sayisi} satır kişi kütüğünden geldi "
                "(katılımcı listesi, sağlık kontrol listesi). Bunlar fatura "
                "değil, tutar taşımıyorlar."
            )
        if mahsup.tutarsiz_satir_sayisi:
            notlar.append(
                f"{mahsup.tutarsiz_satir_sayisi} satırda tutar okunamadı. "
                "Kaynak dosyada tutar kolonu boş olabilir ya da kolon adı "
                "tanınmamış olabilir; **Ayarlar → Kolon Sözlüğü**'nden ekleyin."
            )
        st.caption(" ".join(notlar))

    # --- Dagitim tablosu ---
    st.divider()
    blok_baslik("Dağıtım")

    df = _mahsup_dataframe(mahsup)
    if df.empty:
        st.info("Dağıtılacak tutarlı satır bulunamadı.")
        return

    faturalar = ["(hepsi)"] + sorted({m.kaynak for m in mahsup.satirlar})
    sutun_a, sutun_b = st.columns([2, 1])
    with sutun_a:
        secilen = st.selectbox("Fatura", faturalar, key="mahsup_fatura")
    with sutun_b:
        yalniz_sorunlu = st.checkbox(
            "Yalnızca kontrol gerekenler", key="mahsup_sorunlu",
            help="Masraf merkezi bulunamayan, haritada tanımlı olmayan veya "
                 "içinde incelenecek satır bulunan dağıtım satırları.",
        )

    gosterilecek = df
    if secilen != "(hepsi)":
        gosterilecek = gosterilecek[gosterilecek["Fatura / Kaynak Dosya"] == secilen]
    if yalniz_sorunlu:
        gosterilecek = gosterilecek[gosterilecek["Durum"].astype(str) != ""]

    st.dataframe(gosterilecek, hide_index=True, width="stretch")
    if not gosterilecek.empty:
        st.caption(
            f"{len(gosterilecek)} dağıtım satırı, toplam "
            f"{gosterilecek['Tutar'].sum():,.2f}."
        )

    # --- Proje bazinda ozet ---
    merkezler = mahsup.merkez_ozeti()
    if merkezler:
        st.divider()
        blok_baslik("Proje bazında toplam")
        ozet_df = pd.DataFrame([
            {
                "Masraf Merkezi": k["masraf_merkezi"],
                "Adı": k["masraf_merkezi_adi"] or "",
                "Şirket": k["sirket"] or "",
                "Tutar": k["tutar"],
                "Para Birimi": k["para_birimi"],
                "Pay %": k["pay_yuzde"],
                "Satır": k["satir_sayisi"],
                "Kişi": k["kisi_sayisi"],
                "Haritada": "evet" if k["haritada_var"] else "HAYIR",
            }
            for k in merkezler
        ])
        st.dataframe(ozet_df, hide_index=True, width="stretch")
        haritasizlar = [k for k in merkezler if not k["haritada_var"]]
        if haritasizlar:
            st.warning(
                f"{len(haritasizlar)} masraf merkezi haritada tanımlı değil; "
                "görev yeri metni olduğu gibi kullanıldı. Finans kodlarına "
                "çevrilmesi için **Masraf Merkezi Haritası** sekmesine ekleyin: "
                + ", ".join(str(k["masraf_merkezi"]) for k in haritasizlar[:10])
            )

    # --- CSV indirme ---
    st.divider()
    blok_baslik("Dışa aktar")
    st.caption(
        "Tam Excel dosyası (Mahsuplaşma, Kontrol ve satır dökümü sayfalarıyla) "
        "**Fatura İşle** sekmesindeki 'Excel oluştur' düğmesinden alınır. "
        "Aşağıdaki CSV yalnızca dağıtım tablosunu içerir."
    )
    st.download_button(
        "Mahsuplaşma CSV indir",
        data=df.to_csv(index=False, sep=";").encode("utf-8-sig"),
        file_name=f"mahsuplasma_{datetime.now():%Y%m%d_%H%M%S}.csv",
        mime="text/csv",
    )


# --------------------------------------------------------------------------
# Sekme 4: Inceleme (en onemli sekme)
# --------------------------------------------------------------------------


def _bekleyenler(sonuclar: Sequence["Sonuc"]) -> list[int]:
    """Inceleme bekleyen satirlarin indekslerini dondurur (once eslesmeyenler)."""
    indeksler = [
        i for i, s in enumerate(sonuclar) if s.durum in (DURUM_INCELE, DURUM_ESLESMEDI)
    ]
    return sorted(indeksler, key=lambda i: (DURUM_SIRASI.get(sonuclar[i].durum, 9), i))


def _ogrenme_sonrasi(mesaj: str) -> None:
    """Ogrenme kaydedildikten sonra sonuclari tazeler ve ekrani yeniler."""
    st.session_state["son_ogrenme"] = mesaj
    st.session_state["ogrenilen"] = int(st.session_state.get("ogrenilen", 0)) + 1
    _yeniden_isle()
    st.rerun()


def sekme_inceleme() -> None:
    """Dusuk guvenli satirlari tek tek coz ve sisteme ogret."""
    blok_baslik("İnceleme ve öğretme")
    if not _defter_var_mi():
        return

    mesaj = st.session_state.pop("son_ogrenme", None)
    if mesaj:
        st.success(mesaj)

    sonuclar: list["Sonuc"] = st.session_state.get("sonuclar") or []
    if not sonuclar:
        st.caption("Önce **Fatura İşle** sekmesinden bir dosya işleyin.")
        return

    bekleyen = _bekleyenler(sonuclar)
    if not bekleyen:
        st.success("Tebrikler, kontrol bekleyen satır kalmadı. Tüm satırlar otomatik eşleşti.")
        return

    defter = st.session_state["defter"]
    defterler = defterleri_al()
    harita_df = harita_oku()
    harita_kodlari = sorted({_metin(k) for k in harita_df.get("masraf_merkezi_kodu", []) if _metin(k)})

    st.caption(
        f"{len(bekleyen)} satır kontrol bekliyor. Verdiğiniz her karar `veri/` klasöründeki "
        "defterlere yazılır ve bir dahaki sefere otomatik uygulanır."
    )

    sira = int(st.session_state.get("inceleme_sirasi", 0))
    sira = max(0, min(sira, len(bekleyen) - 1))

    etiketler = []
    for yer, indeks in enumerate(bekleyen, start=1):
        s = sonuclar[indeks]
        kisi = _metin(s.satir.kisi_ham) or _metin(s.satir.aciklama)[:40] or "(kişi yok)"
        etiketler.append(f"{yer}/{len(bekleyen)} - [{s.durum}] {kisi}")

    ust1, ust2, ust3 = st.columns([1, 6, 1])
    with ust1:
        if st.button("Önceki", disabled=sira == 0, width="stretch"):
            st.session_state["inceleme_sirasi"] = sira - 1
            st.rerun()
    with ust2:
        secilen_etiket = st.selectbox(
            "İncelenecek satır", etiketler, index=sira, label_visibility="collapsed"
        )
        yeni_sira = etiketler.index(secilen_etiket)
        if yeni_sira != sira:
            st.session_state["inceleme_sirasi"] = yeni_sira
            st.rerun()
    with ust3:
        if st.button("Sonraki", disabled=sira >= len(bekleyen) - 1, width="stretch"):
            st.session_state["inceleme_sirasi"] = sira + 1
            st.rerun()

    indeks = bekleyen[sira]
    sonuc = sonuclar[indeks]
    satir = sonuc.satir
    eslesme = sonuc.eslesme
    isim_norm = _satir_isim_norm(satir)

    st.markdown(
        f'<div class="kart">{durum_rozeti(sonuc.durum)} &nbsp; '
        f'<b>{Path(str(satir.kaynak_dosya or "")).name}</b> satır {satir.satir_no} '
        f'&nbsp;|&nbsp; {_tarih_metni(satir.belge_tarihi)} '
        f'&nbsp;|&nbsp; {_metin(satir.gider_tipi) or "-"} '
        f'&nbsp;|&nbsp; {_metin(satir.tutar)} {_metin(satir.para_birimi)}</div>',
        unsafe_allow_html=True,
    )

    sol, sag = st.columns([3, 2])
    with sol:
        st.write("**Kaynak dosyadaki ham açıklama**")
        st.code(_metin(satir.aciklama) or "(boş)", language="text")
        st.write("**Metinden çıkarılan kişi**")
        st.write(f"`{_metin(satir.kisi_ham) or '(çıkarılamadı)'}`")
        if isim_norm:
            st.caption(f"Normalize hali (öğrenme anahtarı): `{isim_norm}`")
    with sag:
        st.write("**Sistemin gerekçesi**")
        st.info(_metin(eslesme.aciklama) or "-")
        st.write(
            f"Yöntem: **{YONTEM_ADLARI.get(eslesme.yontem, eslesme.yontem)}** &nbsp;|&nbsp; "
            f"Güven: **{eslesme.guven:.2f}** &nbsp;|&nbsp; Aday: **{eslesme.aday_sayisi}**"
        )
        if sonuc.uyarilar:
            for uyari in sonuc.uyarilar:
                st.warning(uyari)
        if _metin(satir.masraf_merkezi_kaynak):
            st.caption(f"Kaynak dosyada yazan masraf merkezi: {satir.masraf_merkezi_kaynak}")

    st.divider()
    st.write("**Doğru personeli seçin**")

    arama = st.text_input(
        "Personel ara (isim ya da sicil)",
        value="",
        key=f"ara_{indeks}",
        placeholder="Örn: Gunal Emre  veya  102084",
        help="Personel ana verisinde arar. Aday listesi arama sonuçlarıyla genişler.",
    )

    adaylar: list[str] = [
        sicil_normalize(a) for a in (eslesme.aday_siciller or []) if _metin(a)
    ]
    if eslesme.sicil:
        adaylar.insert(0, sicil_normalize(eslesme.sicil))
    if arama.strip():
        try:
            for bulunan in personel_ara(defter, arama):
                if bulunan not in adaylar:
                    adaylar.append(bulunan)
        except Exception as hata:
            _hata_goster("Personel araması başarısız oldu.", hata)

    gorulen: set[str] = set()
    benzersiz = [a for a in adaylar if a and not (a in gorulen or gorulen.add(a))]

    secenekler = ["(seçilmedi)"] + [sicil_etiketi(defter, a) for a in benzersiz]
    secim = st.selectbox("Aday personel", secenekler, index=0, key=f"aday_{indeks}")
    secili_sicil = benzersiz[secenekler.index(secim) - 1] if secim != "(seçilmedi)" else None

    if not benzersiz:
        st.caption(
            "Aday bulunamadı. Yukarıdaki arama kutusuna ismin bir parçasını yazın "
            "(örn. sadece soyadı) veya kişi çalışan değilse 'Çalışan değil' seçeneğini kullanın."
        )

    if secili_sicil:
        kayit = defter.donem_kaydi(secili_sicil, satir.belge_tarihi) or {}
        gorev = _metin(kayit.get("gorev_yeri"))
        kod, bulundu = _masraf_merkezi_coz(gorev, harita_sozlugu(harita_df))
        st.markdown(
            f'<div class="kart">Seçilen kişi <b>{_metin(kayit.get("ad_soyad"))}</b> '
            f'({secili_sicil}) &nbsp;|&nbsp; Dönem: {_tarih_metni(kayit.get("donem"))} '
            f'&nbsp;|&nbsp; Görev yeri: <b>{gorev or "-"}</b> '
            f'&nbsp;|&nbsp; Masraf merkezi: <b>{kod or "-"}</b>'
            f'{"" if bulundu else " (haritada tanımlı değil)"}</div>',
            unsafe_allow_html=True,
        )

    st.divider()
    a1, a2, a3 = st.columns(3)

    with a1:
        st.write("**1. Personele bağla**")
        bagla = st.button(
            "Bu sicile bağla ve öğren",
            type="primary",
            width="stretch",
            disabled=not (secili_sicil and isim_norm),
            key=f"bagla_{indeks}",
        )
        if not isim_norm:
            st.caption("Bu satırda kişi adı yok, bağlama yapılamaz.")
        if bagla and secili_sicil and isim_norm:
            try:
                kayit = defter.sicil_ile(secili_sicil) or {}
                defterler.alias_ekle(
                    isim_norm,
                    secili_sicil,
                    kaynak="inceleme",
                    ad_soyad=_metin(kayit.get("ad_soyad")),
                )
                defterler.kaydet()
                _ogrenme_sonrasi(
                    f"'{isim_norm}' → {secili_sicil} olarak öğrenildi. "
                    "Bir dahaki sefere otomatik eşleşecek."
                )
            except Exception as hata:
                _hata_goster("Öğrenme kaydedilemedi.", hata)

    with a2:
        st.write("**2. Çalışan değil**")
        kurum = st.text_input(
            "Kurum / şirket",
            key=f"kurum_{indeks}",
            placeholder="Örn: ONE TOWER, RENSERVIS, Dış danışman",
        )
        mm_secim = st.selectbox(
            "Masraf merkezi",
            ["(elle yaz)"] + harita_kodlari,
            key=f"mm_{indeks}",
        )
        mm_elle = st.text_input(
            "Masraf merkezi (elle)",
            key=f"mmelle_{indeks}",
            disabled=mm_secim != "(elle yaz)",
            placeholder="Örn: ONE-TOWER",
        )
        masraf_merkezi = mm_elle.strip() if mm_secim == "(elle yaz)" else mm_secim
        harici = st.button(
            "Çalışan değil (harici) olarak kaydet",
            width="stretch",
            disabled=not isim_norm,
            key=f"harici_{indeks}",
        )
        if harici and isim_norm:
            if not kurum.strip() and not masraf_merkezi:
                st.error("En az kurum veya masraf merkezi girin.")
            else:
                try:
                    defterler.harici_ekle(
                        isim_norm,
                        _metin(satir.kisi_ham) or isim_norm,
                        kurum.strip(),
                        masraf_merkezi,
                        aciklama=f"{Path(str(satir.kaynak_dosya or '')).name} satır {satir.satir_no}",
                        kaynak="inceleme",
                    )
                    defterler.kaydet()
                    _ogrenme_sonrasi(
                        f"'{isim_norm}' harici kişi olarak kaydedildi "
                        f"({kurum.strip() or masraf_merkezi}). "
                        "Bir dahaki sefere otomatik eşleşecek."
                    )
                except Exception as hata:
                    _hata_goster("Harici kişi kaydedilemedi.", hata)

    with a3:
        st.write("**3. Şimdilik geç**")
        st.caption("Karar vermeden bir sonraki satıra geçer, hiçbir şey kaydedilmez.")
        if st.button("Atla", width="stretch", key=f"atla_{indeks}"):
            st.session_state["inceleme_sirasi"] = min(sira + 1, len(bekleyen) - 1)
            st.rerun()

    with st.expander("Bekleyen satırların tamamı"):
        alt_tablo = sonuc_tablosu([sonuclar[i] for i in bekleyen])
        st.dataframe(tabloyu_boya(alt_tablo), hide_index=True, width="stretch", height=320)


# --------------------------------------------------------------------------
# Sekme 4: Masraf Merkezi Haritasi
# --------------------------------------------------------------------------


def sekme_harita() -> None:
    """Gorev yeri -> masraf merkezi kodu tablosunu duzenle."""
    blok_baslik("Masraf merkezi haritası")
    st.caption(
        "Personel dosyasındaki **Görev Yeri** değerini finans sistemindeki masraf merkezi "
        "koduna çevirir. Yeni bir proje/şantiye açıldığında buraya bir satır ekleyin."
    )

    df = harita_oku()
    if df.empty:
        st.info("Harita boş. Aşağıya satır ekleyip kaydedebilirsiniz.")

    duzenlenen = st.data_editor(
        df,
        num_rows="dynamic",
        width="stretch",
        height=420,
        key="harita_editor",
        column_config={
            "gorev_yeri": st.column_config.TextColumn(
                "Görev Yeri (personel dosyasındaki ad)", width="large", required=True
            ),
            "masraf_merkezi_kodu": st.column_config.TextColumn(
                "Masraf Merkezi Kodu", width="medium"
            ),
            "masraf_merkezi_adi": st.column_config.TextColumn(
                "Masraf Merkezi Adı", width="large"
            ),
            "sirket": st.column_config.TextColumn("Şirket", width="small"),
            "aktif": st.column_config.TextColumn("Aktif (E/H)", width="small"),
        },
    )

    k1, k2 = st.columns([1, 4])
    with k1:
        if st.button("Kaydet", type="primary", width="stretch"):
            try:
                yol = harita_yaz(duzenlenen)
                haritayi_tazele()
                st.success(f"Harita kaydedildi: {yol.name}")
                if st.session_state.get("sonuclar"):
                    _yeniden_isle()
                    st.info("Mevcut sonuçlar yeni haritaya göre yeniden hesaplandı.")
            except Exception as hata:
                _hata_goster("Harita kaydedilemedi. Dosya Excel'de açık olabilir.", hata)
    with k2:
        st.caption(f"Dosya: `{harita_yolu()}`")

    st.divider()
    blok_baslik("Haritada olmayan görev yerleri")
    defter = st.session_state.get("defter")
    if defter is None:
        st.caption("Personel verisi yüklenince eksik görev yerleri burada listelenir.")
        return

    try:
        mevcut = {_harita_anahtari(g) for g in duzenlenen.get("gorev_yeri", []) if _metin(g)}
        eksikler = [g for g in defter.gorev_yerleri if _harita_anahtari(g) not in mevcut]
    except Exception as hata:
        _hata_goster("Görev yerleri karşılaştırılamadı.", hata)
        return

    if not eksikler:
        st.success("Personel verisindeki tüm görev yerleri haritada tanımlı.")
        return

    st.warning(
        f"{len(eksikler)} görev yeri haritada tanımlı değil. Bu görev yerlerindeki kişiler "
        "için masraf merkezi kodu yerine görev yeri adı kullanılır."
    )
    st.dataframe(
        pd.DataFrame({"Haritada olmayan görev yeri": eksikler}),
        hide_index=True,
        width="stretch",
    )
    if st.button("Eksik görev yerlerini haritaya ekle"):
        try:
            yeni = pd.DataFrame(
                [
                    {
                        "gorev_yeri": g,
                        "masraf_merkezi_kodu": "",
                        "masraf_merkezi_adi": g,
                        "sirket": "",
                        "aktif": "E",
                    }
                    for g in eksikler
                ]
            )
            harita_yaz(pd.concat([duzenlenen, yeni], ignore_index=True))
            haritayi_tazele()
            st.success(
                "Eksik görev yerleri eklendi. Masraf merkezi kodlarını yukarıdaki tablodan "
                "doldurup tekrar kaydedin."
            )
            st.rerun()
        except Exception as hata:
            _hata_goster("Eksik görev yerleri eklenemedi.", hata)


# --------------------------------------------------------------------------
# Sekme 5: Yardim
# --------------------------------------------------------------------------


def sekme_yardim() -> None:
    """Kisa Turkce kullanim kilavuzu."""
    blok_baslik("Nasıl kullanılır?")
    st.markdown(
        """
1. **Ayarlar** sekmesinde personel ana verisi dosyasını seçip *Personel verisini yükle* deyin.
   İlk yükleme yaklaşık yarım dakika sürer, sonrasında anında açılır.
2. **Fatura İşle** sekmesinde fatura/liste dosyalarınızı yükleyin ya da bir klasör yolu verin.
   Sistem her dosyanın tipini kendisi tanır.
3. *İşle* düğmesine basın. Sonuçlar üç gruba ayrılır: **Otomatik**, **İncelenecek**, **Eşleşmedi**.
4. **İnceleme** sekmesinde bekleyen satırları tek tek çözün. Verdiğiniz her karar kaydedilir.
5. **Excel oluştur** ve **Excel indir** ile sonucu finans sistemine aktarın.
        """
    )

    blok_baslik("Hangi dosya tipleri destekleniyor?")
    st.markdown(
        """
| Dosya ailesi | Tanınan içerik |
| --- | --- |
| Antik / Yüzyıl seyahat - ham cari hareket dökümü (.xls) | Bilet, otel, vize, bagaj satırları; kişi adı açıklama metninden çıkarılır |
| Yüzyıl - elle dağıtılmış liste (.xlsx) | Şantiye kolonu dolu referans dosyası; doğruluk karşılaştırması için |
| Energo - assessment yansıtma | *Kişi Listesi* sayfasındaki katılımcılar |
| Energo - arabuluculuk | Personel + TC kimlik no + proje |
| Energo - sağlık kontrol listesi | Adı soyadı, TCKN, doğum tarihi, şantiye |
| Koç Üniversitesi katılımcı listesi | ID kolonu doğrudan sicil numarasıdır (en güvenilir eşleşme) |
| Diğer (tanınmayan) | Genel okuyucu: isim, tarih ve tutar kolonlarını otomatik arar |

Excel (.xls, .xlsx, .xlsm) ve .csv dosyaları kabul edilir.
        """
    )

    blok_baslik("Güven skoru ne demek?")
    st.markdown(
        f"""
Her satır için sistem **0,00 - 1,00** arasında bir güven skoru üretir. Ayarlar sekmesindeki
eşiğin (**şu an {float(st.session_state.get('esik', 0.90)):.2f}**) üzerindeki satırlar otomatik kabul edilir.

| Skor | Yöntem | Anlamı |
| --- | --- | --- |
| 1,00 | Sicil numarası | Kaynak dosyada sicil doğrudan yazıyor |
| 0,99 | TC kimlik no | TCKN köprüsünden bulundu |
| 0,98 | Öğrenilmiş eşleşme | Daha önce siz öğrettiniz |
| 0,95 | Tam isim | İsim personel verisiyle birebir aynı |
| 0,90 | İsim alt kümesi | Rus ad-baba adı-soyadı varyantı |
| 0,88 | Transliterasyon | *IYLMAZ GEKHAN* → *Yılmaz Gökhan* gibi |
| 0,85 | Kesilmiş isim | Bilet sisteminde 20 karakterde kesilmiş ad |
| 0,80 civarı | Bulanık benzerlik | Yazım hatası toleranslı eşleşme |
| 0,50 | Aile bireyi | Soyadı eşleşen çalışanın eşi/çocuğu olabilir - **mutlaka kontrol edin** |
| 0,00 | Eşleşme yok | Kişi personel verisinde bulunamadı |

Birden fazla aday bulunduğunda güven skoru bilinçli olarak düşürülür ve satır incelemeye gönderilir.
        """
    )

    blok_baslik("Öğrenme nasıl çalışıyor?")
    st.markdown(
        """
- Sistem **yapay zeka kullanmaz**, internet gerektirmez. Ofis dışında, uçakta bile çalışır.
- Öğrenme, sizin İnceleme sekmesinde verdiğiniz kararların `veri/` klasöründeki CSV dosyalarına
  yazılmasıyla olur:
  - `aliases.csv` - bir ismin hangi sicile ait olduğu
  - `harici_kisiler.csv` - çalışan olmayan kişiler (grup şirketi, dış danışman)
  - `ek_kisiler.csv` - sağlık/katılımcı listelerinden gelen, henüz personel verisinde olmayan kişiler
  - `tckn_sicil.csv` - TC kimlik no ↔ sicil köprüsü (ana veride TCKN yoktur)
- Bir karar verdikten sonra aynı isim **bir daha hiç sorulmaz**.
- Bu dosyaları Excel'de açıp elle de düzenleyebilirsiniz; noktalı virgülle ayrılmıştır.
        """
    )

    blok_baslik("Sık karşılaşılan durumlar")
    st.markdown(
        """
- **Aile bireyi:** *GUNAL DARIA* gibi bir isim çalışan listesinde yoktur ama *Gunal Emre*
  vardır. Sistem bunu aile bireyi olarak işaretler ve masraf merkezini çalışandan devralır,
  ama güveni düşürüp incelemeye gönderir.
- **Kişi hiç yok:** Grup şirketi (RENSERVIS, ONE TOWER…), taşeron veya yeni giren olabilir.
  *Çalışan değil (harici)* ile kurumu ve masraf merkezini bir kez tanımlayın.
- **Kişiye bağlı olmayan satır:** Cenaze çelengi, genel hizmet gibi satırlarda kişi yoktur;
  bunları *Atla* ile geçip Excel'de elle dağıtın.
- **Çıkış yapmış personel:** Gider tarihi çıkış tarihinden sonraysa uyarı verilir.
- **Görev yeri haritada yok:** *Masraf Merkezi Haritası* sekmesinden kodu tanımlayın.
        """
    )

    blok_baslik("Gizlilik")
    st.markdown(
        """
Tüm işlemler bu bilgisayarda yapılır. Hiçbir veri internete gönderilmez.
Kişisel veri içeren `ornek_veri/`, `cikti/` ve `veri/` klasörleri sürüm kontrolüne dahil edilmez.
        """
    )


# --------------------------------------------------------------------------
# Ana akis
# --------------------------------------------------------------------------


def main() -> None:
    """Uygulamayi baslatir ve sekmeleri cizer."""
    st.set_page_config(
        page_title=UYGULAMA_ADI,
        page_icon=":bar_chart:",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
    st.markdown(STIL, unsafe_allow_html=True)
    st.markdown(
        '<div class="ust-serit"><h1>Masraf Merkezi Otomasyonu</h1>'
        "<p>Tedarikçi faturalarındaki her satırı, satırdaki kişiye göre doğru "
        "masraf merkezine mahsuplaştırır. Tamamen bu bilgisayarda çalışır.</p></div>",
        unsafe_allow_html=True,
    )

    if not CEKIRDEK_HAZIR:
        st.error(
            "Çekirdek modüller henüz hazır değil. `masraf/` klasöründeki modüller "
            "yüklenemedi, bu yüzden eşleştirme yapılamıyor. Kurulumun tamamlanmasını "
            "bekleyin veya `pip install -r requirements.txt` komutunu çalıştırın."
        )
        with st.expander("Teknik ayrıntı"):
            st.code(CEKIRDEK_HATA, language="text")
        return

    _oturum_hazirla()

    sekmeler = st.tabs(
        [
            "Ayarlar",
            "Fatura İşle",
            "Mahsuplaşma",
            "İnceleme",
            "Masraf Merkezi Haritası",
            "Yardım",
        ]
    )
    ciziciler = (sekme_ayarlar, sekme_fatura, sekme_mahsuplasma,
                 sekme_inceleme, sekme_harita, sekme_yardim)
    for sekme, cizici in zip(sekmeler, ciziciler):
        with sekme:
            try:
                cizici()
            except Exception as hata:  # pragma: no cover - arayuz koruma katmani
                _hata_goster(
                    "Bu sekme çizilirken beklenmeyen bir hata oluştu. "
                    "Sayfayı yenileyip tekrar deneyin.",
                    hata,
                )


if __name__ == "__main__":
    main()
