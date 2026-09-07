"""Mail GOVDESINI okur: yansitma ozet tablosu ve egitim katilim isaretleri.

Neden gerekli: Energo yansitma mailinde sekiz kalemin besi (62.688,29 USD,
toplamin %78,8'i) hicbir ekte yok, yalnizca mailin govdesindeki iki kolonlu
'kalem | $ tutar' tablosunda duruyor. Koc egitimine kimin katildigi da ekte
degil: ekteki liste DAVET listesi (25 kisi x 2 gun, hicbir isaret yok),
gercek katilim ic mailin govdesinde yesil boyali satirlarda (17 kisi).
Mailin kendi cumlesi: "yesil highlight olanlar egitime katilanlar".

Bu modul iki sey uretir, ikisi de TASIYICI satirdir (tutar alani bos,
mahsuba girmez; dagitim kurali bekler):

* ``govde_kalemi``: ozet tablonun her satiri icin bir tasiyici. Etiket ve
  tutar ek sozlugunde.
* ``govde_katilim``: yesil isaretli tablolar tasiyan her mail icin bir
  tasiyici. Katilimci kimlikleri ve katildigi gunler ek sozlugunde.

Ozet tablo kurali (164 gercek tabloda olculdu: 1 eslesme, 0 yanlis pozitif;
gevsek surum 7 tablo kabul ediyordu, 6'si sigorta tarife karti):

* Basliktan sonra TAM iki kolon.
* Her deger hucresi PARA BIRIMI ISARETLI bir tutar ($ 5.211,77 / 3.000,00 TL).
  Cozulemeyen tek bir hucre bile tabloyu REDDETTIRIR: kurussuz '3.000' gibi
  bir hucre sessizce dusseydi 3.000 USD kaybolurdu (olculdu).
* Etiketi bos satir kabul edilmez.
* 'TOPLAM' / 'Total' / 'Итого' satiri kalem DEGIL, kontrol toplamidir.
* Karisik para birimi RED.
* Sayi bicimi Turkce: nokta ya da bosluk binlik, virgul kurus. ABD bicimi
  (1,234.56) kabul EDILMEZ; iki turlu okunabilen sayi belirsizdir.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any

from masraf.metin import ascii_katla
from masraf.modeller import GiderSatiri

_log = logging.getLogger(__name__)

__all__ = [
    "OzetTablo", "OzetKalemi", "Katilimci",
    "para_coz", "ozet_tablosu", "yesil_katilimcilar", "govde_satirlari",
    "KAYNAK_TIP_KALEM", "KAYNAK_TIP_KATILIM",
]

KAYNAK_TIP_KALEM = "govde_kalemi"
KAYNAK_TIP_KATILIM = "govde_katilim"

#: Para birimi isareti/etiketi -> ISO kodu.
_PARA = {
    "$": "USD", "usd": "USD", "us$": "USD",
    "₺": "TRY", "tl": "TRY", "try": "TRY", "trl": "TRY",
    "€": "EUR", "eur": "EUR",
    "₽": "RUB", "rub": "RUB", "руб": "RUB", "руб.": "RUB",
}
#: Kontrol toplami satirlari; kalem sayilmaz.
_TOPLAM_ETIKETLERI = ("toplam", "total", "итого", "genel toplam", "grand total", "sum")

#: '$ 5.211,77' / '3.000,00 TL' / '21 997,00' ... Turkce bicim, 2 hane kurus.
_PARA_DESENI = re.compile(
    r"^\s*(?P<on>[$€₺₽]|usd|try|tl|trl|eur|rub|руб\.?)?\s*"
    r"(?P<sayi>-?\d{1,3}(?:[.  ]\d{3})*,\d{2}|-?\d+,\d{2})\s*"
    r"(?P<arka>[$€₺₽]|usd|try|tl|trl|eur|rub|руб\.?)?\s*$",
    re.IGNORECASE,
)


def para_coz(metin: str) -> tuple[float, str] | None:
    """'$ 5.211,77' -> (5211.77, 'USD'). Para birimi isareti ZORUNLU; yoksa None."""
    if not metin:
        return None
    m = _PARA_DESENI.match(metin.replace(" ", " ").strip())
    if not m:
        return None
    etiket = (m.group("on") or m.group("arka") or "").lower()
    pb = _PARA.get(etiket)
    if not pb:
        return None
    sayi = m.group("sayi").replace(".", "").replace(" ", "").replace(",", ".")
    try:
        return round(float(sayi), 2), pb
    except ValueError:  # pragma: no cover - desen zaten suzuyor
        return None


@dataclass
class OzetKalemi:
    etiket: str
    tutar: float
    para_birimi: str


@dataclass
class OzetTablo:
    """Mail govdesindeki 'kalem | tutar' ozet tablosu."""

    kalemler: list[OzetKalemi] = field(default_factory=list)
    para_birimi: str = ""
    baslik: str = ""
    #: Tabloda TOPLAM satiri varsa onun degeri; kendi toplamimizla karsilastirilir.
    beyan_toplam: float | None = None

    @property
    def toplam(self) -> float:
        return round(sum(k.tutar for k in self.kalemler), 2)

    @property
    def beyan_uyusuyor(self) -> bool:
        return self.beyan_toplam is None or abs(self.beyan_toplam - self.toplam) < 0.01


def _hucreler(tr) -> list[str]:
    return [h.get_text(" ", strip=True) for h in tr.find_all(["td", "th"])]


def _tablo_oku(tablo) -> OzetTablo | None:
    """Tek bir <table> icin ozet tablo kuralini uygular; uymuyorsa None."""
    satirlar = [_hucreler(tr) for tr in tablo.find_all("tr")]
    satirlar = [s for s in satirlar if any(h.strip() for h in s)]
    if len(satirlar) < 3:
        return None
    if any(len(s) != 2 for s in satirlar):
        return None
    baslik = ""
    kalemler: list[OzetKalemi] = []
    beyan: float | None = None
    pbler: set[str] = set()
    for i, (etiket, deger) in enumerate(satirlar):
        cozum = para_coz(deger)
        if cozum is None:
            if i == 0 and not kalemler:
                baslik = deger or etiket      # ilk satir baslik olabilir
                continue
            return None                         # cozulemeyen deger: tabloyu reddet
        tutar, pb = cozum
        pbler.add(pb)
        if not etiket.strip():
            return None                         # etiketsiz satir: yapi belirsiz
        if ascii_katla(etiket).strip().lower() in _TOPLAM_ETIKETLERI:
            beyan = tutar
            continue
        kalemler.append(OzetKalemi(etiket=etiket.strip(), tutar=tutar, para_birimi=pb))
    if len(kalemler) < 2 or len(pbler) != 1:
        return None
    # Etiketler benzersiz olmali: ayni etiketi tekrarlayan iki kolonlu tablo
    # bir tarife kartidir ('premium per person per month' 2,75 / 1,95 -
    # olculdu), yansitma ozeti degil.
    if len({ascii_katla(k.etiket).lower() for k in kalemler}) != len(kalemler):
        return None
    return OzetTablo(kalemler=kalemler, para_birimi=next(iter(pbler)), baslik=baslik,
                     beyan_toplam=beyan)


def ozet_tablosu(html: str, duz: str = "") -> OzetTablo | None:
    """Govdedeki ozet tabloyu bulur. Birden fazla uyan varsa en cok kalemlisi.

    HTML yoksa duz metinden 'etiket / $ tutar' ciftleri denenir (ikinci,
    bagimsiz yol; olculdu, ayni sekiz cifti veriyor).
    """
    adaylar: list[OzetTablo] = []
    if html:
        try:
            from bs4 import BeautifulSoup
            corba = BeautifulSoup(html, "html.parser")
            for t in corba.find_all("table"):
                if t.find("table"):
                    continue                     # ic ice tablo: yerlesim, veri degil
                o = _tablo_oku(t)
                if o:
                    adaylar.append(o)
        except Exception as hata:  # noqa: BLE001 - govde bozuksa sessiz gecme
            _log.debug("govde HTML ayristirilamadi: %s", hata)
    if adaylar:
        return max(adaylar, key=lambda o: len(o.kalemler))
    if duz:
        return _duz_metinden(duz)
    return None


_DUZ_CIFT = re.compile(r"^(?P<etiket>[^\n$₺€]{2,60}?)\s*\n+\s*(?P<deger>[$₺€]\s*[\d.  ]+,\d{2})\s*$",
                       re.MULTILINE)


def _duz_metinden(duz: str) -> OzetTablo | None:
    kalemler = []
    pbler = set()
    for m in _DUZ_CIFT.finditer(duz):
        cozum = para_coz(m.group("deger"))
        if not cozum:
            continue
        tutar, pb = cozum
        etiket = m.group("etiket").strip()
        if ascii_katla(etiket).lower() in _TOPLAM_ETIKETLERI:
            continue
        pbler.add(pb)
        kalemler.append(OzetKalemi(etiket=etiket, tutar=tutar, para_birimi=pb))
    if len(kalemler) < 2 or len(pbler) != 1:
        return None
    if len({ascii_katla(k.etiket).lower() for k in kalemler}) != len(kalemler):
        return None                         # tekrarlayan etiket: tarife karti
    return OzetTablo(kalemler=kalemler, para_birimi=next(iter(pbler)))


# ---------------------------------------------------------------------------
# Yesil isaretli katilimcilar
# ---------------------------------------------------------------------------

@dataclass
class Katilimci:
    kimlik: str
    ad: str
    gunler: list[str] = field(default_factory=list)   # 'YYYY-MM-DD' ya da 'gun-1'

    @property
    def gun_sayisi(self) -> int:
        return len(self.gunler)


_KIMLIK = re.compile(r"^\d{4,9}$")
_YESIL = re.compile(r"background(?:-color)?\s*:\s*(?:lime|#00ff00|#0f0|green|#92d050|#00b050)", re.I)
_TARIH = re.compile(r"(\d{2})[./-](\d{2})[./-](\d{4})")


def _tablo_tarihi(tablo, sira: int) -> str:
    """Tablodan hemen once gecen 'dd.mm.yyyy' metnini bulur; yoksa 'gun-N'."""
    try:
        onceki = tablo.find_previous(string=_TARIH)
        if onceki:
            m = _TARIH.search(str(onceki))
            if m:
                g, a, y = m.groups()
                return datetime(int(y), int(a), int(g)).date().isoformat()
    except Exception:  # noqa: BLE001
        pass
    return f"gun-{sira}"


def yesil_katilimcilar(html: str) -> list[Katilimci]:
    """Yesil boyali 'kimlik | ad' satirlarini tum tablolardan toplar.

    Ayni kisi birden fazla tabloda (birden fazla gun) isaretliyse gunleri
    birikir. Hicbir yesil satir yoksa bos liste doner.
    """
    if not html or not _YESIL.search(html):
        return []
    try:
        from bs4 import BeautifulSoup
        corba = BeautifulSoup(html, "html.parser")
    except Exception:  # noqa: BLE001
        return []
    kisiler: dict[str, Katilimci] = {}
    sira = 0
    for tablo in corba.find_all("table"):
        if tablo.find("table"):
            continue
        satirlar = tablo.find_all("tr")
        yesiller = []
        for tr in satirlar:
            h = _hucreler(tr)
            if len(h) < 2 or not _KIMLIK.match(h[0].strip()):
                continue
            if _YESIL.search(str(tr)):
                yesiller.append((h[0].strip(), h[1].strip()))
        if not yesiller:
            continue
        sira += 1
        gun = _tablo_tarihi(tablo, sira)
        for kimlik, ad in yesiller:
            k = kisiler.setdefault(kimlik, Katilimci(kimlik=kimlik, ad=ad))
            if gun not in k.gunler:
                k.gunler.append(gun)
    return list(kisiler.values())


# ---------------------------------------------------------------------------
# Tasiyici satirlar
# ---------------------------------------------------------------------------

def govde_satirlari(govde: Any, mail_adi: str) -> list[GiderSatiri]:
    """Bir mail govdesinden tasiyici GiderSatiri listesi uretir.

    ``govde``: posta.MailGovdesi (konu, gonderen, tarih, zincir, html, duz).
    Istisna firlatmaz; okunacak bir sey yoksa bos liste doner.
    """
    sonuc: list[GiderSatiri] = []
    html = getattr(govde, "html", "") or ""
    duz = getattr(govde, "duz", "") or ""
    konu = getattr(govde, "konu", "") or ""
    tarih = getattr(govde, "tarih", None)
    if not isinstance(tarih, date):
        tarih = None
    ortak = {
        "govde_konu": konu,
        "mail_konusu": konu,
        "mail_gonderen": getattr(govde, "gonderen", None),
        "mail_tarihi": tarih,
        "mail_zinciri": getattr(govde, "kaynak_aciklamasi", konu),
    }

    ozet = ozet_tablosu(html, duz)
    if ozet:
        for i, k in enumerate(ozet.kalemler, start=1):
            sonuc.append(GiderSatiri(
                kaynak_dosya=mail_adi, kaynak_tip=KAYNAK_TIP_KALEM, satir_no=i,
                belge_tarihi=tarih, aciklama=f"Mail ozeti: {k.etiket}",
                kisi_ham=None, sicil_ham=None, tckn_ham=None,
                tutar=None, para_birimi=None, masraf_merkezi_kaynak=None, gider_tipi="Diger",
                ek={**ortak, "govde_kalem": k.etiket, "govde_tutar": k.tutar,
                    "govde_para_birimi": k.para_birimi, "govde_baslik": ozet.baslik,
                    "govde_kalem_sayisi": len(ozet.kalemler), "govde_toplam": ozet.toplam,
                    "govde_beyan_toplam": ozet.beyan_toplam,
                    "tutar_yontemi": "mail govdesindeki ozet tablodan okundu; belge degil, "
                                     "dagitim kurali bekleniyor"},
            ))

    katilim = yesil_katilimcilar(html)
    if katilim:
        sonuc.append(GiderSatiri(
            kaynak_dosya=mail_adi, kaynak_tip=KAYNAK_TIP_KATILIM, satir_no=len(sonuc) + 1,
            belge_tarihi=tarih, aciklama=f"Katilim isaretleri: {len(katilim)} kisi, "
                                          f"{sum(k.gun_sayisi for k in katilim)} kisi-gun",
            kisi_ham=None, sicil_ham=None, tckn_ham=None,
            tutar=None, para_birimi=None, masraf_merkezi_kaynak=None, gider_tipi="Egitim",
            ek={**ortak,
                "katilimcilar": [{"kimlik": k.kimlik, "ad": k.ad, "gunler": list(k.gunler)}
                                 for k in katilim],
                "tutar_yontemi": "mail govdesindeki yesil isaretli satirlar: egitime katilanlar"},
        ))
    return sonuc
