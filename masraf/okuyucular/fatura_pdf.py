"""PDF fatura basligini okur: fatura no, tarih, matrah, KDV, toplam, VKN.

Bir PDF fatura TUTARI verir, KISI KIRILIMI vermez. Bu yuzden okuyucu tek bir
TASIYICI satir dondurur ve o satirin ``tutar`` alani BILEREK None birakilir:
tutar dolu birakilsaydi 1.043.126,59 TRY dogrudan dagitima girer ve USD
tablosunda anlamsiz bir kalem olusurdu. Tutarin kisilere nasil bolunecegi
masraf.dagitim'in isidir; oraya kadar satir 'tutari okunamadi' sayilir ve
mutabakati ACIK tutar. Sessiz kapanma bu sekilde imkansiz olur.

Dort satici sablonu olculdu (Energo Mayis-Haziran yansitmasi, 10 metinli PDF):

* A PLUS (Arabuluculuk): etiket ve deger AYRI SATIRDA, 'Fatura Numarasi',
  'Hesaplanan KDV GERCEK (%20.0)'.
* AS Olcme (Assessment) ve Kerem Kockesen: etiket ve deger AYNI SATIRDA,
  'Fatura No:', 'Hesaplanan KDV(%20)'.
* Unvest (Koc Universitesi): noktasiz 'Tutari', para birimi tutara BITISIK
  ('869.272,16TRY'), tarih ayracsiz ('09062026'), ayrica yumusak tire tasir.

Tek regex kumesi hepsini kapsar; her alanin 2-3 alternatifi vardir ve metin
once masraf.okuyucular.pdf_metin.metni_normalize'dan gecer.
"""

from __future__ import annotations

import logging
import re
from datetime import date, datetime
from pathlib import Path

from masraf.metin import ascii_katla
from masraf.modeller import GiderSatiri
from masraf.okuyucular.pdf_metin import PYPDF_VAR, pdf_metni

_log = logging.getLogger(__name__)

__all__ = ["fatura_pdf_oku", "FaturaBasligi", "fatura_basligi_coz"]

#: Bu okuyucunun urettigi satirlarin kaynak_tip'i.
KAYNAK_TIP = "fatura_pdf"

#: Para birimi etiketi -> ISO kodu. Fatura ustundeki yazim bicimleri.
_DOVIZ = {"TL": "TRY", "TRY": "TRY", "TRL": "TRY", "₺": "TRY",
          "USD": "USD", "$": "USD", "EUR": "EUR", "€": "EUR", "RUB": "RUB", "₽": "RUB"}

#: Tutar: 1.043.126,59 ya da 41.250,00; ardindan bitisik ya da bosluklu para birimi.
_TUTAR = r"(?P<tutar>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s*(?P<pb>TRY|TRL|TL|USD|EUR|RUB|₺|\$|€|₽)?"

#: Etiket ile deger arasi: iki nokta ve/veya bosluk, EN FAZLA bir satir sonu.
#: A PLUS sablonunda deger bir alt satirda duruyor; bu yuzden \n serbest.
_ARA = r"\s*:?\s*\n?\s*"


def _kur(*etiketler: str) -> str:
    return "(?:" + "|".join(etiketler) + ")"


_DESENLER: dict[str, re.Pattern[str]] = {
    "fatura_no": re.compile(
        _kur(r"Fatura\s+No", r"Fatura\s+Numarasi") + _ARA + r"(?P<deger>[A-Z0-9][A-Z0-9\-/]{5,32})",
        re.IGNORECASE),
    "fatura_tarihi": re.compile(
        r"Fatura\s+Tarihi" + _ARA + r"(?P<deger>\d{2}[.\-/]?\d{2}[.\-/]?\d{4})",
        re.IGNORECASE),
    "matrah": re.compile(
        _kur(r"Mal\s+Hizmet\s+Toplam\s+Tutari", r"Matrah") + _ARA + _TUTAR,
        re.IGNORECASE),
    # 'Hesaplanan KDV(%20)' ve 'Hesaplanan KDV GERCEK (%20.0)' ayni desene
    # girer; parantezli oran istege baglidir.
    "kdv": re.compile(
        r"Hesaplanan\s+KDV(?:\s+GERCEK)?(?:\s*\([^)\n]{0,20}\))?" + _ARA + _TUTAR,
        re.IGNORECASE),
    "toplam": re.compile(
        r"Vergiler\s+Dahil\s+Toplam\s+Tutar" + _ARA + _TUTAR,
        re.IGNORECASE),
    "odenecek": re.compile(
        r"Odenecek\s+Tutar" + _ARA + _TUTAR,
        re.IGNORECASE),
    "ettn": re.compile(r"ETTN" + _ARA + r"(?P<deger>[0-9A-Fa-f\-]{30,40})"),
}

#: Vergi kimlik numaralari; ilk gecen alici, ikinci satici olmak zorunda DEGIL,
#: bu yuzden hepsi toplanir ve ayri alanda tasinir.
_VKN = re.compile(r"\b(?:VKN|TCKN|Vergi\s+No)\s*:?\s*\n?\s*(\d{10,11})\b", re.IGNORECASE)


def _tutar_coz(metin: str) -> float | None:
    """'1.043.126,59' -> 1043126.59. Turkce bicim: nokta binlik, virgul kurus."""
    try:
        return round(float(metin.replace(".", "").replace(",", ".")), 2)
    except (TypeError, ValueError):  # pragma: no cover - regex zaten suzuyor
        return None


def _tarih_coz(metin: str) -> date | None:
    """'22-05-2026', '02.06.2026', '09062026' -> date. Cozemezse None."""
    ham = re.sub(r"[.\-/]", "", metin or "")
    if len(ham) != 8 or not ham.isdigit():
        return None
    try:
        return datetime.strptime(ham, "%d%m%Y").date()
    except ValueError:
        return None


class FaturaBasligi:
    """Bir PDF faturanin basligindan cikarilan alanlar ve oz-denetim sonucu."""

    __slots__ = ("fatura_no", "tarih", "matrah", "kdv", "toplam", "odenecek",
                 "para_birimi", "vkn_listesi", "ettn", "taranmis", "hata",
                 "bulunamayan", "aritmetik")

    def __init__(self) -> None:
        self.fatura_no: str | None = None
        self.tarih: date | None = None
        self.matrah: float | None = None
        self.kdv: float | None = None
        self.toplam: float | None = None
        self.odenecek: float | None = None
        self.para_birimi: str | None = None
        self.vkn_listesi: list[str] = []
        self.ettn: str | None = None
        self.taranmis: bool = False
        self.hata: str = ""
        self.bulunamayan: list[str] = []
        self.aritmetik: str = ""

    @property
    def okundu_mu(self) -> bool:
        return bool(self.fatura_no and self.toplam is not None)

    def sozluk(self) -> dict:
        return {
            "fatura_no": self.fatura_no,
            "fatura_tarihi": self.tarih.isoformat() if self.tarih else None,
            "fatura_matrah": self.matrah,
            "fatura_kdv": self.kdv,
            "fatura_toplam_yerel": self.toplam,
            "fatura_odenecek_yerel": self.odenecek,
            "para_birimi_yerel": self.para_birimi,
            "fatura_vkn": list(self.vkn_listesi),
            "fatura_ettn": self.ettn,
            "pdf_taranmis": self.taranmis,
            "pdf_bulunamayan_alanlar": list(self.bulunamayan),
            "pdf_aritmetik": self.aritmetik,
        }


def fatura_basligi_coz(metin: str) -> FaturaBasligi:
    """Normalize edilmis PDF metninden fatura alanlarini cikarir.

    Metin ASCII'ye katlanarak aranir: 'Tutari' ile 'Tutarı', 'GERCEK' ile
    'GERÇEK' ayni desenle yakalansin diye. Katlama harf sayisini degistirmez,
    tutar ve fatura numarasi zaten ASCII'dir.
    """
    basligi = FaturaBasligi()
    if not metin:
        return basligi
    duz = ascii_katla(metin)

    bulunan: dict[str, re.Match[str]] = {}
    for ad, desen in _DESENLER.items():
        esle = desen.search(duz)
        if esle:
            bulunan[ad] = esle
        elif ad != "ettn":     # ETTN zorunlu degil
            basligi.bulunamayan.append(ad)

    if "fatura_no" in bulunan:
        basligi.fatura_no = bulunan["fatura_no"].group("deger").strip()
    if "fatura_tarihi" in bulunan:
        basligi.tarih = _tarih_coz(bulunan["fatura_tarihi"].group("deger"))
    if "ettn" in bulunan:
        basligi.ettn = bulunan["ettn"].group("deger").strip()

    pblar: list[str] = []
    for ad in ("matrah", "kdv", "toplam", "odenecek"):
        esle = bulunan.get(ad)
        if not esle:
            continue
        setattr(basligi, ad, _tutar_coz(esle.group("tutar")))
        pb = _DOVIZ.get((esle.group("pb") or "").upper())
        if pb:
            pblar.append(pb)
    if pblar:
        basligi.para_birimi = max(set(pblar), key=pblar.count)

    basligi.vkn_listesi = list(dict.fromkeys(_VKN.findall(duz)))

    # Oz-denetim: matrah + KDV, vergiler dahil toplama esit olmali. Olculdu:
    # gercek on faturanin onunda da tutuyor. Tutmuyorsa okuma suphelidir ve
    # bu satir elle kontrole dusmelidir - sessizce dogru sayilmaz.
    if basligi.matrah is not None and basligi.kdv is not None and basligi.toplam is not None:
        fark = round(basligi.matrah + basligi.kdv - basligi.toplam, 2)
        basligi.aritmetik = "tamam" if abs(fark) < 0.01 else f"matrah+KDV toplamdan {fark:+.2f} farkli"
    else:
        basligi.aritmetik = "eksik alan"
    return basligi


def fatura_pdf_oku(yol: str | Path) -> list[GiderSatiri]:
    """PDF faturayi okur; tek bir tasiyici GiderSatiri dondurur.

    Sozlesme (bkz. kesif.PARSERLAR): tek argument alir, GiderSatiri listesi
    doner ve HICBIR KOSULDA istisna firlatmaz - firlatirsa bir ek yuzunden
    butun mail 'okunamadi' hatasina duser.

    Taranmis (metin katmani olmayan) PDF'te de BOS LISTE degil, taranmis
    isaretli bir tasiyici doner. Bos liste donseydi envanter bu dosyaya
    'kolon adlari taninmamis olabilir' derdi; oysa dogru sebep 'saf goruntu,
    OCR gerekir'dir ve bu bilgi kapaktan kaybolmamalidir.
    """
    p = Path(yol)
    sonuc = pdf_metni(p)
    basligi = fatura_basligi_coz(sonuc.metin)
    basligi.taranmis = sonuc.taranmis
    basligi.hata = sonuc.hata

    if sonuc.hata:
        aciklama = f"PDF okunamadi: {sonuc.hata}"
    elif sonuc.taranmis:
        aciklama = "PDF taranmis goruntu; metin katmani yok (OCR gerekir)"
    elif basligi.okundu_mu:
        aciklama = f"Fatura {basligi.fatura_no}"
        if basligi.toplam is not None:
            aciklama += f"; belge tutari {basligi.toplam:,.2f} {basligi.para_birimi or ''}".rstrip()
    else:
        eksik = ", ".join(basligi.bulunamayan) or "alanlar"
        aciklama = f"PDF acildi ama fatura basligi okunamadi (eksik: {eksik})"

    ek = basligi.sozluk()
    ek["pdf_sayfa_sayisi"] = sonuc.sayfa_sayisi
    ek["pdf_kutuphane"] = "pypdf" if PYPDF_VAR else "yok"
    ek["tutar_yontemi"] = (
        "PDF fatura basligi okundu; tutar kisilere bolunmedi cunku belge kisi "
        "kirilimi tasimiyor. Dagitim kurali bekleniyor."
    )
    if basligi.aritmetik and basligi.aritmetik != "tamam":
        ek["okuyucu_uyarisi"] = (
            f"PDF fatura oz-denetimi tutmadi ({basligi.aritmetik}); tutar elle dogrulanmali"
        )

    return [GiderSatiri(
        kaynak_dosya=p.name,
        kaynak_tip=KAYNAK_TIP,
        satir_no=1,
        belge_tarihi=basligi.tarih,
        aciklama=aciklama,
        kisi_ham=None,
        sicil_ham=None,
        tckn_ham=None,
        tutar=None,                 # BILEREK: tutar yerel para biriminde, ek'te durur
        para_birimi=None,
        masraf_merkezi_kaynak=None,
        gider_tipi="Diger",
        ek=ek,
    )]
