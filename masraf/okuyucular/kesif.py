"""Kaynak dosya tipini otomatik tespit eden kesif modulu.

Kullanici dosyayi surukleyip biraktiginda hangi parser'in calistirilacagini
belirler. Tespit tamamen deterministiktir: dosya acilir, SAYFA ADLARI ile her
sayfanin BASLIK satiri ve ustundeki dosya basligi satirlarindaki hucre
metinleri ASCII katlanmis bicimde toplanir, ardindan oncelik sirali ipucu
kurallari uygulanir. Veri hucreleri ipucu SAYILMAZ: saglik listesinde bir
kisinin gorevi 'ARABULUCU' diye tum dosya arabuluculuk sanilmaz (olculdu:
80 tutarsiz gider satiri uretiyor, kutuk beslemesine girmiyordu).

Ipuclari (baslik / sayfa adi uzerinden, iceren eslesme):
    'Cari Hareket Dokumu'                        -> antik_cari
    'SANTIYESI' + 'UCUS GUZERGAHI'               -> yuzyil_dagitilmis
    'Katilimci' + ('Paket'|'Energo Payi'|...)    -> energo_assessment
    'ARABULUCU'                                  -> energo_arabulucu
    'BORDROLU LISTE' | 'TCKN'+'SAGLIK'           -> energo_saglik
    'ID' + 'Alt Fonksiyon'                       -> koc_katilimci
    'Sicil No' + 'Masraf Merkezi', tutar yok     -> referans_liste
    aksi halde                                   -> genel

Birden fazla kural eslesirse hepsi oncelik sirasiyla ADAY olur: ozel okuyucu
sablonu tanimayip bos donerse sonraki aday denenir, en son genel okuyucuya
dusulur ve bu dusus satirlara (ek['okuyucu_uyarisi']) ve envantere yazilir.
"""

from __future__ import annotations

import logging

from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable

_log = logging.getLogger(__name__)

from masraf.modeller import GiderSatiri
from masraf.okuyucular.antik import antik_cari_oku, yuzyil_dagitilmis_oku
from masraf.okuyucular.energo import (
    arabulucu_oku,
    assessment_oku,
    koc_katilimci_oku,
    saglik_oku,
)
from masraf.okuyucular.fatura_pdf import fatura_pdf_oku
from masraf.okuyucular.genel import (
    baslik_satiri_bul,
    calisma_oku,
    genel_oku,
    hucre_sayisi,
    kolon_anahtari,
    kolon_ara,
    kolon_haritasi,
)

__all__ = [
    "dosya_tipini_bul", "dosya_tip_adaylari", "oku", "oku_tip", "okuyucu_turu",
    "PARSERLAR", "OKUYUCU_UYARISI",
]

#: Satir ek sozlugundeki uyari anahtari. Ozel okuyucu dosyayi tanidi ama
#: sablonu okuyamadi ve genel okuyucuya dusuldu; tutar/para birimi elle
#: dogrulanmali. Boru tarafinda bu anahtar satiri INCELE'ye dusurmelidir.
OKUYUCU_UYARISI = "okuyucu_uyarisi"

#: Genel okuyucunun tutar kolonu adaylari (kural 7 icin). genel.py'deki
#: listeyle ayni; o modul degisirse buradaki yedek liste kullanilir.
try:  # pragma: no cover - savunma
    from masraf.okuyucular.genel import _TUTAR_ADAYLARI as _KESIF_TUTAR_ADAYLARI
except ImportError:  # pragma: no cover
    _KESIF_TUTAR_ADAYLARI = (
        "tutar", "satis", "borc", "amount", "toplam", "bedel", "fiyat",
        "energo payi", "usd", "total",
    )

# kaynak_tip -> parser fonksiyonu
PARSERLAR: dict[str, Callable[[str | Path], list[GiderSatiri]]] = {
    "antik_cari": antik_cari_oku,
    "yuzyil_dagitilmis": yuzyil_dagitilmis_oku,
    "energo_assessment": assessment_oku,
    "energo_arabulucu": arabulucu_oku,
    "energo_saglik": saglik_oku,
    "koc_katilimci": koc_katilimci_oku,
    "referans_liste": genel_oku,
    "fatura_pdf": fatura_pdf_oku,
    "genel": genel_oku,
}

# Bir dosyanin gider satiri mi yoksa kisi kutugu mu oldugunu ayirt eden esik.
# Kutuklerde hicbir satirda tutar yoktur; faturalarda neredeyse her satirda vardir.
_KUTUK_SATIR_ESIGI = 200

# Kesif icin okunacak satir sayisi (dosyanin tamami okunmaz, hizli kalir).
# genel_oku'nun baslik arama siniriyla AYNI olmali.
from masraf.okuyucular.genel import BASLIK_ARAMA_SINIRI as _KESIF_SATIR_SINIRI  # noqa: E402


@dataclass
class _SayfaIpucu:
    """Kesif icin bir sayfanin baslik bilgisi (ilk satirlar)."""

    ad: str
    baslik_i: int                                   # -1: baslik yok
    baslik: dict[str, int] = field(default_factory=dict)  # kolon_haritasi
    satirlar: list[list[Any]] = field(default_factory=list)


def _ipuclarini_topla(yol: Path) -> tuple[set[str], str, list[_SayfaIpucu]]:
    """Sayfa adlari ile baslik ve baslik-ustu satirlarin hucre metinlerini toplar.

    Her sayfada baslik satiri (en cok metin hucresi tasiyan satir) bulunur;
    yalnizca o satir ve USTUNDEKI dosya basligi satirlari ipucu sayilir.
    Veri hucreleri sayilmaz: bir kisinin gorevi, pozisyonu ya da aciklamasi
    dosyanin tipini degistirmemelidir.

    Returns:
        (normalize edilmis benzersiz metinler kumesi, hepsinin birlesimi,
        sayfa basina baslik bilgisi)
    """
    calisma = calisma_oku(yol, satir_siniri=_KESIF_SATIR_SINIRI)
    metinler: set[str] = set()
    sayfalar: list[_SayfaIpucu] = []
    for sayfa_adi, satirlar in calisma.sayfalar.items():
        anahtar = kolon_anahtari(sayfa_adi)
        if anahtar:
            metinler.add(anahtar)
        baslik_i = baslik_satiri_bul(satirlar, sinir=_KESIF_SATIR_SINIRI)
        for satir in satirlar[: baslik_i + 1]:
            for hucre in satir:
                anahtar = kolon_anahtari(hucre)
                if anahtar:
                    metinler.add(anahtar)
        sayfalar.append(_SayfaIpucu(
            ad=sayfa_adi, baslik_i=baslik_i,
            baslik=kolon_haritasi(satirlar[baslik_i]) if baslik_i >= 0 else {},
            satirlar=satirlar,
        ))
    return metinler, " || ".join(sorted(metinler)), sayfalar


def _tutar_kolonu(sayfalar: Iterable[_SayfaIpucu]) -> str | None:
    """Taninan bir tutar kolonu ve altinda sayisal deger var mi?

    Kural 7 icin: 'Sicil No' + 'Masraf Merkezi' tasiyan dosya ancak tutar
    kolonu YOKSA kisi kutugudur. Tutar kolonu (yerlesik adaylar + kullanici
    sozlugu) ve ilk satirlarda en az bir sayisal degeri varsa gider
    dosyasidir. Bulunursa kolonun adi doner, yoksa None.
    """
    adaylar: tuple[str, ...] = tuple(_KESIF_TUTAR_ADAYLARI)
    try:
        from masraf.kolon_sozlugu import genislet as _genislet
        adaylar = tuple(_genislet("tutar", adaylar))
    except Exception:  # noqa: BLE001 - sozluk okunamazsa yerlesik liste yeter
        pass
    for sayfa in sayfalar:
        if sayfa.baslik_i < 0:
            continue
        i = kolon_ara(sayfa.baslik, *adaylar)
        if i is None:
            continue
        for satir in sayfa.satirlar[sayfa.baslik_i + 1:]:
            if i < len(satir) and hucre_sayisi(satir[i]) is not None:
                return next((ad for ad, k in sayfa.baslik.items() if k == i), None) or "tutar"
    return None


def _iceriyor(blob: str, *parcalar: str) -> bool:
    """Birlesik metinde verilen parcalarin HEPSI geciyor mu?"""
    return all(kolon_anahtari(p) in blob for p in parcalar)


def dosya_tip_adaylari(yol: str | Path) -> list[str]:
    """Dosyaya uyan kaynak tiplerini ONCELIK SIRASIYLA listeler.

    Ilk eleman ``dosya_tipini_bul`` sonucudur; sonrakiler, ozel okuyucu
    sablonu tanimayip bos donerse sirayla denenecek adaylardir. Liste her
    zaman 'genel' ile biter (Outlook mesaji haric). Dosya acilamiyorsa
    ['genel'] doner (istisna firlatmaz).
    """
    p = Path(yol)

    # 0) Outlook mesaji: icerik degil uzanti belirler. Mesaj bir kapsayicidir,
    #    icindeki tablo dosyalari ayri ayri tespit edilir.
    if p.suffix.lower() == ".msg":
        return ["outlook_msg"]

    # 0b) PDF fatura: yine icerik degil uzanti belirler. Kisa devre ZORUNLU;
    #     olmazsa _ipuclarini_topla calisma_oku'yu cagirir, o .pdf icin
    #     ValueError firlatir, asagidaki except onu ['genel']'e cevirir ve
    #     genel_oku ayni hatayi bir daha firlatir - dosya OKUNAMADI olur.
    if p.suffix.lower() == ".pdf":
        return ["fatura_pdf"]

    try:
        metinler, blob, sayfalar = _ipuclarini_topla(p)
    except Exception:
        return ["genel"]

    adaylar: list[str] = []

    # 1) Antik ham cari hareket dokumu
    if _iceriyor(blob, "cari hareket"):
        adaylar.append("antik_cari")

    # 2) Yuzyil elle dagitilmis (santiye + ucus kolonlari birlikte)
    if (_iceriyor(blob, "santiyesi") and _iceriyor(blob, "ucus guzergahi")) or (
            _iceriyor(blob, "ucus guzergahi") and _iceriyor(blob, "ucus bilgisi")):
        adaylar.append("yuzyil_dagitilmis")

    # 3) Energo assessment: kisi kolonu ('Katilimci', 'Katilimci Adi'...) ve
    #    sablona ozgu kolonlardan biri. Tam hucre esitligi degil, iceren
    #    eslesme: 'Katilimci Adi' de 'Katilimci'dir, 'Paket' 'Program' olsa
    #    da 'Energo Payi' / 'Uygulama Turu' sablonu ele verir.
    if _iceriyor(blob, "katilimci") and any(
            _iceriyor(blob, k) for k in ("paket", "energo payi", "uygulama turu")):
        adaylar.append("energo_assessment")

    # 4) Energo arabuluculuk (baslik ya da sayfa adinda; veri hucresinde degil)
    if _iceriyor(blob, "arabulucu"):
        adaylar.append("energo_arabulucu")

    # 5) Saglik kontrol listesi
    if (_iceriyor(blob, "bordrolu liste") or _iceriyor(blob, "bordrosuz liste")
            or (_iceriyor(blob, "tckn") and _iceriyor(blob, "saglik kontrol"))):
        adaylar.append("energo_saglik")

    # 6) Koc Universitesi katilimci listesi
    if "id" in metinler and _iceriyor(blob, "alt fonksiyon"):
        adaylar.append("koc_katilimci")

    # 7) Ferdi kaza sigorta listesi ve benzeri PERSONEL KUTUKLERI.
    #    Bunlar fatura degil, kisi kutugudur: sicil ve masraf merkezi tasir
    #    ama tutar tasimaz. Gider satiri olarak islenirlerse binlerce sahte
    #    satir uretirler; ayri tip olarak isaretlenip defter beslemesine
    #    yonlendirilirler. Taninan bir tutar kolonu ve altinda sayisal deger
    #    VARSA kutuk degil gider dosyasidir (olculdu: 5 satir 510 USD
    #    dagitilmiyor, uyari da yanlis yere 'kolon_esanlamlilari.csv'ye
    #    ekleyin' diyordu).
    if (_iceriyor(blob, "sicil no") and _iceriyor(blob, "masraf merkezi")
            and _tutar_kolonu(sayfalar) is None):
        adaylar.append("referans_liste")

    adaylar.append("genel")
    return adaylar


def dosya_tipini_bul(yol: str | Path) -> str:
    """Dosyanin hangi kaynak ailesine ait oldugunu belirler.

    Donen deger modeller.KAYNAK_TIPLERI kumesindendir. Dosya acilamiyorsa
    veya hicbir ipucu eslesmiyorsa 'genel' doner (istisna firlatmaz).
    """
    return dosya_tip_adaylari(yol)[0]


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


def _msg_oku(
    yol: Path,
    cikarma_dizini: str | Path | None = None,
    gorulen_ozetler: set | None = None,
    atlanan_ekler: list | None = None,
    envanter: list | None = None,
) -> list[GiderSatiri]:
    """Outlook mesajindaki tum tablo eklerini cikarir ve tek tek okur.

    Mesaj bir kapsayicidir: icinde baska mesajlar, zip arsivleri ve Excel
    dosyalari olabilir. Cikarilan her dosyanin tipi ayrica tespit edilir.
    Her satira hangi mailden geldigi `ek['mail_konusu']` icinde yazilir.

    Raises:
        MesajOkunamadi: Hicbir ekten satir cikmadiysa, sebebiyle birlikte.
    """
    import shutil
    from tempfile import mkdtemp


    # Cikarilan ekler KISISEL VERI tasir (ad soyad, TC, tutar). Gecici dizini
    # biz actiysak is bitince SILMEK zorundayiz; aksi halde her calistirmada
    # faturalar Windows'ta %TEMP% altinda birikir ve kimse fark etmez.
    bizim_dizin = cikarma_dizini is None
    hedef = Path(cikarma_dizini) if cikarma_dizini else Path(mkdtemp(prefix="mm_"))
    try:
        return _msg_oku_icerik(yol, hedef, gorulen_ozetler, atlanan_ekler, envanter)
    finally:
        if bizim_dizin:
            shutil.rmtree(hedef, ignore_errors=True)


def _dosya_ozeti(yol: Path) -> str:
    """Dosya iceriginin SHA-256 ozeti; ayni ek iki mailde gelirse ayni ozet."""
    import hashlib
    h = hashlib.sha256()
    with open(yol, "rb") as f:
        for parca in iter(lambda: f.read(1 << 20), b""):
            h.update(parca)
    return h.hexdigest()


def _msg_oku_icerik(yol: Path, hedef: Path, gorulen_ozetler: set | None = None,
                    atlanan_ekler: list | None = None,
                    envanter: list | None = None) -> list[GiderSatiri]:
    """``_msg_oku``'nun govdesi; gecici dizin yonetimi disarida tutulur.

    ``envanter`` verilirse her ek icin bir ``DosyaKaydi`` eklenir: okunan,
    kutuk sayilan, bos donen, hata veren, ayni icerik oldugu icin atlanan ve
    tablo olmayan (PDF, sifreli arsiv, bozuk ek) ekler. Ic mailler ve zip
    arsivleri de 'MAIL' / 'ARSIV' bilgi satiri olarak girer. Sessiz atlama
    yoktur.

    Envanter, uyari ve kontrol satirlarinda ekin MAILDEKI ORIJINAL ADI
    (``gosterim_adi``) kullanilir; diske yazilan kisaltilmis/'_2' ekli ad
    kullaniciya gosterilmez. Boylece boru'daki ad tabanli 'ayni adli dosya'
    kontrolu ve mahsuplasmadaki fatura anahtari maildeki adla calisir.
    """
    from masraf.envanter import (
        ARSIV, ATLANDI, AYNI_ICERIK, GOVDE, MAIL, OKUNAMADI, DosyaKaydi, _boyut, satirlardan_kayit,
    )
    from masraf.okuyucular.govde import govde_satirlari
    from masraf.okuyucular.posta import msg_aciklarini_cikar

    def _kaydet(k: DosyaKaydi) -> None:
        if envanter is not None:
            envanter.append(k)

    def _kaynak(zincir) -> str:
        """'mail.msg > konu > ic mail > arsiv.zip' bicimindeki kaynak zinciri."""
        return " > ".join([yol.name] + [str(z) for z in (zincir or [])])

    satirlar: list[GiderSatiri] = []
    atlananlar_yerel: list = atlanan_ekler if atlanan_ekler is not None else []
    onceki_atlanan = len(atlananlar_yerel)
    kapsayicilar: list = []
    govdeler: list = []
    ekler = msg_aciklarini_cikar(yol, hedef, atlananlar=atlananlar_yerel,
                                 kapsayicilar=kapsayicilar, govdeler=govdeler)
    yeni_atlananlar = atlananlar_yerel[onceki_atlanan:]
    # Tablo olmayan ekler (PDF) mailler ARASINDA da tekillenir: ayni PDF iki
    # ayri mailde gelirse kapak/Kontrol 'okunmayan ek' sayisi farkli dosya
    # sayisini versin. (Mail ICINDEKI tekrarlari posta zaten isaretledi.)
    if gorulen_ozetler is not None:
        for a in yeni_atlananlar:
            ozet = getattr(a, "ozet", None)
            if not ozet or getattr(a, "tekrar", False):
                continue
            if ozet in gorulen_ozetler:
                a.tekrar = True
                a.sebep = "icerigi baska bir mailde daha once gorulen ekle birebir ayni; tekrar sayilmadi"
            else:
                gorulen_ozetler.add(ozet)
    tekrar_sayisi = sum(1 for a in yeni_atlananlar if getattr(a, "tekrar", False))
    ic_mail_sayisi = sum(1 for k in kapsayicilar if getattr(k, "tur", "") == "mail")
    arsiv_sayisi = sum(1 for k in kapsayicilar if getattr(k, "tur", "") == "arsiv")
    ozet_metni = f"{len(ekler)} tablo eki, {len(yeni_atlananlar) - tekrar_sayisi} okunmayan ek"
    if tekrar_sayisi:
        ozet_metni += f", {tekrar_sayisi} tekrar eden ek"
    if ic_mail_sayisi:
        ozet_metni += f", {ic_mail_sayisi} ekli mail"
    if arsiv_sayisi:
        ozet_metni += f", {arsiv_sayisi} arsiv"
    _kaydet(DosyaKaydi(ad=yol.name, kaynak="", tur="outlook_msg", durum=MAIL,
                       sebep=ozet_metni, boyut=_boyut(yol)))
    # Kapsayicilar (ic mail, zip): tablo degiller ama envanterde gorunmeliler;
    # aksi halde 'm1.msg 3 ek tasiyordu, 1'i sifreliydi' bilgisi kaybolur.
    for k in kapsayicilar:
        mail_mi = getattr(k, "tur", "") == "mail"
        _kaydet(DosyaKaydi(
            ad=getattr(k, "ad", str(k)), kaynak=_kaynak(getattr(k, "zincir", [])),
            tur="outlook_msg" if mail_mi else "zip", durum=MAIL if mail_mi else ARSIV,
            sebep=getattr(k, "aciklama", ""), boyut=getattr(k, "boyut", None)))
    for a in yeni_atlananlar:
        ad = getattr(a, "ad", str(a))
        _kaydet(DosyaKaydi(
            ad=ad, kaynak=_kaynak(getattr(a, "zincir", [])),
            tur=Path(ad).suffix.lower().lstrip(".") or "?",
            durum=AYNI_ICERIK if getattr(a, "tekrar", False) else ATLANDI,
            sebep=getattr(a, "sebep", "tablo degil; acilmadi"),
            boyut=getattr(a, "boyut", None),
            ozet=(getattr(a, "ozet", None) or None)))

    # Ayni ek iki farkli mailde (iletilmis, tekrar gonderilmis) gelirse
    # icerigi birebir aynidir. Ikisini de okumak parayi cift sayar; yineleme
    # tespiti de yakalayamaz cunku ayni dosya adi 'ayni dosyadaki tekrar'
    # sayilir. Cozum: icerik ozeti. Ilk goruleni oku, sonrakini atla ve soyle.
    if gorulen_ozetler is not None:
        kalan = []
        for ek in ekler:
            ozet = getattr(ek, "ozet", None)
            if not ozet:
                try:
                    ozet = _dosya_ozeti(ek.yol)
                except OSError:
                    kalan.append(ek); continue
            if ozet in gorulen_ozetler:
                _log.warning("Ayni ek daha once okundu, atlandi: %s (%s)", ek.gosterim_adi, yol.name)
                _kaydet(DosyaKaydi(
                    ad=ek.gosterim_adi, kaynak=_kaynak(ek.zincir),
                    tur=Path(ek.gosterim_adi).suffix.lower().lstrip("."), durum=AYNI_ICERIK,
                    sebep="icerigi daha once okunan bir ekle birebir ayni; cift sayim olmasin diye atlandi",
                    boyut=_boyut(ek.yol), ozet=ozet))
                continue
            gorulen_ozetler.add(ozet)
            kalan.append(ek)
        atlanan = len(ekler) - len(kalan)
        ekler = kalan
        if not ekler and atlanan:
            raise MesajOkunamadi(
                f"mesajdaki {atlanan} ekin tamami baska bir mailden zaten okunmustu "
                "(ayni icerik). Cift sayim olmasin diye atlandi."
            )

    if not ekler:
        okunmayan = len(yeni_atlananlar) - tekrar_sayisi
        ipucu = ""
        if okunmayan:
            ornekler = [getattr(a, "ad", str(a)) for a in yeni_atlananlar if not getattr(a, "tekrar", False)][:5]
            ipucu = (f" Mailde {okunmayan} okunmayan ek var (orn. {', '.join(ornekler)}); "
                     "sebepleri Dosyalar sayfasinda.")
        raise MesajOkunamadi(
            "mesajin icinde okunabilir tablo eki bulunamadi. Aranan uzantilar: "
            ".xlsx .xls .xlsm .csv .tsv. Mail yalnizca metin/gorsel tasiyor "
            "olabilir, ya da ekler mailin govdesine gomulu olabilir. Ekleri "
            "Outlook'ta kaydedip dogrudan 1_FATURALAR klasorune atmayi deneyin." + ipucu
        )

    # Her ek icin ne oldugunu ayri ayri tut; hepsi basarisiz olursa raporla.
    bos_kalanlar: list[str] = []
    hatalilar: list[str] = []
    for ek in ekler:
        ek_adi = ek.gosterim_adi
        ek_kaynak = _kaynak(ek.zincir)
        try:
            ek_tip = dosya_tipini_bul(ek.yol)
        except Exception:  # noqa: BLE001
            ek_tip = ""
        try:
            ic_satirlar = oku(ek.yol)
        except Exception as hata:  # noqa: BLE001 - kullaniciya gosterilecek
            hatalilar.append(f"{ek_adi}: {hata.__class__.__name__}: {hata}")
            _kaydet(DosyaKaydi(ad=ek_adi, kaynak=ek_kaynak, tur=ek_tip, durum=OKUNAMADI,
                               sebep=f"{hata.__class__.__name__}: {hata}", boyut=_boyut(ek.yol)))
            continue
        # Envanter turu fiilen kullanilan okuyucudur ('energo_assessment -> genel'
        # gibi); ozel okuyucu sablonu tanimayip genel'e dustuyse sebebi de yazilir.
        kayit = satirlardan_kayit(ek_adi, ek_kaynak, okuyucu_turu(ic_satirlar, ek_tip), ic_satirlar,
                                  boyut=_boyut(ek.yol), ozet=getattr(ek, "ozet", None))
        uyari = next((str(s.ek.get(OKUYUCU_UYARISI)) for s in ic_satirlar
                      if isinstance(getattr(s, "ek", None), dict) and s.ek.get(OKUYUCU_UYARISI)), "")
        if uyari:
            kayit.sebep = f"{kayit.sebep}; {uyari}" if kayit.sebep else uyari
        _kaydet(kayit)
        if not ic_satirlar:
            bos_kalanlar.append(ek_adi)
            continue
        for s in ic_satirlar:
            # Kaynak dosya adini mesaj + ORIJINAL ek adi olarak yaz; izlenebilirlik
            # ve ad tabanli kontroller (ayni adli dosya, fatura anahtari) icin.
            s.kaynak_dosya = f"{yol.name} > {ek_adi}"
            if isinstance(s.ek, dict):
                s.ek.setdefault("mail_konusu", ek.mail_konusu)
                s.ek.setdefault("mail_gonderen", ek.mail_gonderen)
                s.ek.setdefault("mail_tarihi", ek.mail_tarihi)
                s.ek.setdefault("mail_zinciri", ek.kaynak_aciklamasi)
        satirlar.extend(ic_satirlar)

    # Mail GOVDELERI: ozet tablo ve yesil katilim isaretleri. Ek degil,
    # belge de degil; tasiyici satir olarak yukari tasinir ve envanterde
    # 'MAIL GOVDESI' olarak gorunur. Sekiz kalemin besi yalnizca burada.
    for govde in govdeler:
        try:
            govde_satirlar = govde_satirlari(govde, yol.name)
        except Exception as hata:  # noqa: BLE001 - govde bozuksa mail dusmesin
            govde_satirlar = []
            _kaydet(DosyaKaydi(ad=f"(govde) {govde.konu}", kaynak=_kaynak(govde.zincir),
                               tur="govde", durum=OKUNAMADI,
                               sebep=f"govde okunamadi: {hata.__class__.__name__}: {hata}"))
        if not govde_satirlar:
            continue
        kalemler = [g for g in govde_satirlar if g.kaynak_tip == "govde_kalemi"]
        katilim = [g for g in govde_satirlar if g.kaynak_tip == "govde_katilim"]
        parcalar_g = []
        if kalemler:
            ek0 = kalemler[0].ek
            parcalar_g.append(f"ozet tablo: {len(kalemler)} kalem, toplam "
                              f"{ek0.get('govde_toplam', 0):,.2f} {ek0.get('govde_para_birimi', '')}")
        for g in katilim:
            kisi = g.ek.get("katilimcilar") or []
            parcalar_g.append(f"yesil isaretli katilim: {len(kisi)} kisi, "
                              f"{sum(len(k.get('gunler') or []) for k in kisi)} kisi-gun")
        _kaydet(DosyaKaydi(ad=f"(govde) {govde.konu}", kaynak=_kaynak(govde.zincir),
                           tur="govde", durum=GOVDE,
                           sebep="; ".join(parcalar_g) + "; belge degil, dagitima girmedi",
                           satir=len(govde_satirlar)))
        for g in govde_satirlar:
            g.kaynak_dosya = f"{yol.name} > (govde) {govde.konu}"
        satirlar.extend(govde_satirlar)

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


def okuyucu_turu(satirlar: Iterable[Any], tespit_tipi: str,
                 kullanilan_tip: str | None = None) -> str:
    """Envanterde gosterilecek 'Tur': fiilen kullanilan okuyucunun tipi.

    Satirlarin ``kaynak_tip`` alanindan turetilir (assessment okuyucusu
    tutarsiz listeyi 'energo_assessment_detay' olarak etiketler; envanter de
    oyle gostermeli). Ozel okuyucu bos donup baska bir okuyucuya dusulduyse
    'energo_assessment -> genel' bicimindedir; satirlarin ek sozlugundeki
    'okuyucu_turu' notu (oku() yazar) varsa dogrudan o kullanilir, boylece
    mail eklerinde de dogru gorunur.
    """
    satirlar = list(satirlar)
    for s in satirlar:
        ek = getattr(s, "ek", None)
        if isinstance(ek, dict) and ek.get("okuyucu_turu"):
            return str(ek["okuyucu_turu"])
    tipler = Counter(getattr(s, "kaynak_tip", "") for s in satirlar)
    tipler.pop("", None)
    fiili = tipler.most_common(1)[0][0] if tipler else (kullanilan_tip or tespit_tipi)
    if kullanilan_tip is None or kullanilan_tip == tespit_tipi or fiili == tespit_tipi:
        return fiili
    return f"{tespit_tipi} -> {fiili}"


def oku(
    yol: str | Path,
    cikarma_dizini: str | Path | None = None,
    gorulen_ozetler: set | None = None,
    atlanan_ekler: list | None = None,
    envanter: list | None = None,
) -> list[GiderSatiri]:
    """Dosya tipini bulur ve dogru parser'i calistirir.

    Outlook mesajlari kapsayici olarak ele alinir: icindeki tum tablo ekleri
    cikarilip ayri ayri okunur ve tek listede birlestirilir.

    Ozel parser hic satir uretmezse (sablon beklenenden farkliysa) oncelik
    sirasindaki sonraki aday, en son genel parser denenir; boylece bilinmeyen
    bir surum sessizce bos sonuc vermez. Ama SESSIZCE de dusulmez: genel
    okuyucu kolonlari tahminle sectigi icin (TL 'Toplam' kolonunu USD gibi
    dagitabilir, olculdu) her satirin ek['okuyucu_uyarisi'] alanina ve
    envanter kaydinin sebebine uyari yazilir; bu satirlar elle dogrulanmali.
    """
    p = Path(yol)
    adaylar = dosya_tip_adaylari(p)
    tip = adaylar[0]
    if tip == "outlook_msg":
        return _msg_oku(p, cikarma_dizini, gorulen_ozetler, atlanan_ekler, envanter)

    satirlar: list[GiderSatiri] = []
    kullanilan = tip
    bos_donenler: list[str] = []
    for aday in adaylar:
        kullanilan = aday
        satirlar = oku_tip(p, aday)
        if satirlar:
            break
        bos_donenler.append(aday)

    # Ozel okuyucu dosyayi tanidi ama sablonu okuyamadi; genel okuyucuyla
    # (ya da onun kutuk haliyle) okundu. Kolon secimi tahmindir.
    sablon_taninmadi = kullanilan != tip and PARSERLAR[kullanilan] is genel_oku
    uyari: str | None = None
    if sablon_taninmadi:
        uyari = (
            f"ozel okuyucu ({tip}) sablonu tanimadi, genel okuyucuyla okundu; "
            "tutar ve para birimi elle dogrulanmali"
        )
        _log.warning("%s: %s", p.name, uyari)

    # Guvenlik agi: cok satirli ve hicbir satirinda tutar olmayan bir dosya
    # fatura degil kisi kutugudur. Gider olarak islenirse sahte satir uretir.
    kutuk_sebebi: str | None = None
    if kullanilan == "referans_liste":
        kutuk_sebebi = "sicil ve masraf merkezi kolonlari var, tutar kolonu yok"
    elif (kullanilan == "genel" and len(satirlar) >= _KUTUK_SATIR_ESIGI
            and not any(s.tutar is not None for s in satirlar)):
        kullanilan = "referans_liste"
        kutuk_sebebi = f"{len(satirlar)} satirin hicbirinde tutar yok"

    if kullanilan == "referans_liste":
        for s in satirlar:
            s.kaynak_tip = "referans_liste"
            if isinstance(s.ek, dict):
                s.ek["referans_liste"] = True
                s.ek["kutuk_sebebi"] = kutuk_sebebi

    tur = okuyucu_turu(satirlar, tip, kullanilan)
    for s in satirlar:
        if not isinstance(s.ek, dict):
            continue
        if kullanilan != tip:
            s.ek["okuyucu_turu"] = tur
        if uyari:
            s.ek[OKUYUCU_UYARISI] = uyari

    if envanter is not None:
        from masraf.envanter import _boyut, satirlardan_kayit
        kayit = satirlardan_kayit(p.name, "", tur, satirlar, boyut=_boyut(p))
        notlar = [n for n in (uyari, kutuk_sebebi) if n]
        if bos_donenler and not satirlar:
            notlar.append("denenen okuyucular: " + ", ".join(bos_donenler))
        if notlar:
            kayit.sebep = "; ".join(([kayit.sebep] if kayit.sebep else []) + notlar)
        envanter.append(kayit)
    return satirlar
