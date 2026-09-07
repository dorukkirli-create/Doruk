"""Outlook .msg dosyalarindan tablo eklerini cikarir.

Finans ekibine gelen dosyalar Outlook mesaji olarak geliyor. Bir mesajin icinde
baska mesajlar, onlarin icinde zip arsivleri, onlarin icinde de Excel dosyalari
olabiliyor. Bu modul bu agaci sonuna kadar yuruyup butun tablo dosyalarini
duz bir listeye cikarir ve her dosyanin hangi mail zincirinden geldigini kaydeder.

Kullanim:
    from masraf.okuyucular.posta import msg_aciklarini_cikar
    ekler = msg_aciklarini_cikar("gelen.msg", "cikti/acilan")
    for ek in ekler:
        print(ek.yol, ek.gosterim_adi, ek.mail_konusu, ek.mail_tarihi)

Ilkeler:
* Sessiz atlama yok. Okunmayan her ek (PDF, sifreli arsiv, bozuk ek, bulut
  baglantisi, cok derin ic ice mail...) ``atlananlar`` listesine ``AtlananEk``
  olarak yazilir ve kullanici Excel'de gorur.
* Tek bir bozuk ya da sifreli ek mailin TAMAMINI dusurmez; o ek atlanir,
  digerleri okunmaya devam eder.
* Tekrar tespiti ICERIK ozetiyle (sha256) yapilir; tablo olmayan ekler (PDF)
  icin de. Ad + boyut yetmez: tedarikcinin fatura basina gonderdigi tek
  satirlik sablon dosyalari ('Fatura Detayi.xlsx') ayni adi ve ayni boyutu
  tasir ama farkli para tasir.
* Diske yazilan yol KISA tutulur: klasorler 'm01_a1b2c3' / 'z02_d4e5f6', dosya
  adi en cok 40 karakter, zip girdileri duzlestirilir. Windows MAX_PATH (260)
  mail konularindan uretilen klasor adlariyla kolayca asiliyordu. Kullaniciya
  gosterilen zincir (``zincir``, ``gosterim_adi``) yine orijinal adlardir.
* Ic mailler ve arsivler de birer kayittir (``Kapsayici``); envanterde
  'MAIL' / 'ARSIV' bilgi satiri olarak gorunurler.

Bagimlilik: extract-msg. Windows uzerinde `pip install extract-msg` yeterlidir.
"""

from __future__ import annotations

import hashlib
import logging
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from urllib.parse import unquote

_log = logging.getLogger(__name__)

# Cikarilacak tablo uzantilari
TABLO_UZANTILARI = {".xlsx", ".xls", ".xlsm", ".csv", ".tsv"}
# PDF faturalar. TABLO_UZANTILARI'na KARISTIRILMAZ: o kume 'bu ek bir tablo
# okuyucusuna gidecek' anlamini tasir, PDF ise fatura basligi cikarimina gider.
PDF_UZANTILARI = {".pdf"}
# Diske cikarilip okuyucuya verilecek uzantilarin tamami.
CIKARILAN_UZANTILAR = TABLO_UZANTILARI | PDF_UZANTILARI
# Isimize yaramayan, atlanacak uzantilar (mail imzasindaki logolar vb)
GORSEL_UZANTILARI = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emf", ".wmf", ".ico"}

# Ic ice inme siniri. Sonsuz donguye karsi guvenlik.
AZAMI_DERINLIK = 6

# Zip bombasi korumasi. Sinirlar tek arsiv icin DEGIL, bir mailde acilan
# butun arsivlerin toplami icin uygulanir (ic ice arsivler de sayilir).
AZAMI_ACILMIS_BOYUT = 500 * 1024 * 1024  # 500 MB
AZAMI_DOSYA_SAYISI = 2000

#: Govdesi bu boyuttan buyuk (RTF olarak) mailler icin HTML cozumu YAPILMAZ,
#: yalnizca duz metin okunur. Olculdu: extract_msg'in htmlBody'si RTF'ten
#: uretiliyor ve 2,6 MB'lik bir tablo dokumu 35,7 saniye suruyor; ozet tablo
#: ve katilim isaretleri tasiyan mailler 40-174 KB. Sinirin ustundeki govde
#: 'html_atlandi' isaretiyle gecer, sessizce degil.
GOVDE_AZAMI_RTF = 512 * 1024

#: Diske yazilan dosya adinin azami uzunlugu (uzanti dahil). Klasor adlari
#: 10 karakter ('m01_a1b2c3'); 7 seviye + 40 karakter dosya adi %TEMP% ile
#: birlikte 260'in altinda kalir.
DISK_AD_AZAMI = 40

#: Kullaniciya gosterilen sebep metinleri; finanscinin ne yapacagini soyler.
SEBEP_SIFRELI = "sifreli arsiv; parola gerekli. Arsivi parolasiz kaydedip yeniden gonderin"
SEBEP_BULUT = ("bulut baglantisi (OneDrive/SharePoint); icerik mailde yok, "
               "ekli dosyayi elle indirin ve 1_FATURALAR klasorune atin")


def _derinlik_sebebi() -> str:
    return (f"ic ice {AZAMI_DERINLIK} seviyeden derin mail/arsiv; okunmadi. "
            "Maili Outlook'ta acip ekleri dogrudan gonderin")


def _yazma_sebebi(e: BaseException) -> str:
    return (f"dosya gecici dizine yazilamadi ({e.__class__.__name__}: {e}); "
            "yol cok uzun ya da izin yok olabilir")


@dataclass
class AtlananEk:
    """Mail agacinda gorulen ama okunmayan ek. Envanter ve Kontrol sayfasi icin.

    ``ad`` maildeki ORIJINAL ek adidir (diske yazilan kisaltilmis ad degil).
    """

    ad: str
    zincir: list[str] = field(default_factory=list)
    sebep: str = "tablo degil"
    boyut: int | None = None
    #: True ise icerigi (sha256) daha once gorulen bir ekle birebir ayni;
    #: envanterde 'AYNI ICERIK' olarak gorunur, 'okunmayan ek' sayilmaz.
    #: Tablo olmayan ekler (ayni PDF hem zip'te hem ic mailde) icin de gecerli.
    tekrar: bool = False
    #: Icerik ozeti (sha256). Icerigi okunabilen ekler icin dolu.
    ozet: str | None = None

    @property
    def kaynak_aciklamasi(self) -> str:
        return " > ".join(self.zincir)

    def __str__(self) -> str:  # eski metin bicimiyle uyumlu
        return f"{self.ad}  [{self.kaynak_aciklamasi}]" if self.zincir else self.ad


@dataclass
class CikarilanEk:
    """Mail agacindan cikarilmis tek bir dosya ve nereden geldigi."""

    yol: Path
    #: Diskteki dosya adi. Guvenli/kisaltilmis olabilir, ayni klasore dusen
    #: es adli ekler '_2' eki alir. Raporlarda BUNU DEGIL ``gosterim_adi``ni
    #: kullanin.
    ad: str
    mail_konusu: str | None = None
    mail_gonderen: str | None = None
    mail_tarihi: date | None = None
    # Kokten bu dosyaya kadar olan yol: ["ana mail", "ekli mail", "arsiv.zip"]
    zincir: list[str] = field(default_factory=list)
    derinlik: int = 0
    #: Maildeki / arsivdeki orijinal ek adi (kirpilmamis, '_2' eksiz).
    orijinal_ad: str | None = None
    #: Icerik ozeti (sha256, tam). Tekrar tespiti icin.
    ozet: str | None = None

    @property
    def gosterim_adi(self) -> str:
        """Kullaniciya gosterilecek ad: maildeki orijinal ek adi."""
        return self.orijinal_ad or self.ad

    @property
    def kaynak_aciklamasi(self) -> str:
        """Kullaniciya gosterilecek okunakli kaynak zinciri."""
        if not self.zincir:
            return self.gosterim_adi
        return " > ".join(self.zincir + [self.gosterim_adi])


@dataclass
class Kapsayici:
    """Mail agacindaki ic mail ya da arsiv: kendisi tablo degil, ek tasir.

    Envanterde bilgi satiri olarak gorunur ('MAIL' / 'ARSIV'); boylece
    'm1.msg' ya da 'Faturalar.zip' kac ek tasiyordu, kaci okunamadi
    izlenebilir. Kok mail bu listeye girmez; onu cagiran kaydeder.
    """

    ad: str
    tur: str                       # 'mail' | 'arsiv'
    zincir: list[str] = field(default_factory=list)
    boyut: int | None = None
    derinlik: int = 0
    konu: str | None = None        # mail icin: kendi konusu; arsiv icin: bulundugu mailin konusu
    gonderen: str | None = None
    tarih: date | None = None
    ek_sayisi: int = 0             # dogrudan ek / girdi sayisi
    aciklama: str = ""             # kullaniciya gosterilecek ozet

    @property
    def kaynak_aciklamasi(self) -> str:
        return " > ".join(self.zincir)


@dataclass
class MailGovdesi:
    """Bir mailin GOVDESI: ek degil, mesajin kendi metni.

    Sekiz kalemlik yansitma ozeti ve egitim katilim isaretleri hicbir ekte
    degil, mailin govdesindedir (olculdu: toplamin %78,8'i yalnizca orada).
    Bu yuzden govde de yukari tasinir; masraf.okuyucular.govde onu okur.
    """

    konu: str
    gonderen: str | None
    tarih: object
    zincir: list           # kok mailden bu mesaja kadar konu/arsiv adlari
    derinlik: int
    html: str = ""         # htmlBody (UTF-8'e cozulmus)
    duz: str = ""          # body (duz metin)
    html_atlandi: bool = False   # govde RTF cok buyuk, HTML cozulmedi

    @property
    def kaynak_aciklamasi(self) -> str:
        return " > ".join(str(z) for z in (self.zincir + [self.konu]))


@dataclass
class _Yuruyus:
    """Bir mailin agaci yurunurken tasinan ortak durum.

    Sonuc listeleri, kisa klasor sayaci ve zip bombasi butcesi (bu mailde
    toplam acilan bayt ve dosya sayisi) burada durur; ic ice cagrilar ayni
    nesneyi paylasir.
    """

    sonuc: list[CikarilanEk]
    atlananlar: list | None = None
    kapsayicilar: list | None = None
    govdeler: list | None = None
    acilan_bayt: int = 0
    acilan_dosya: int = 0
    sayac: int = 0
    #: Bu yuruyuste uretilen AtlananEk'ler (cagiranin listesindeki eski
    #: kayitlara dokunmadan tekrar isaretlemek icin).
    atlanan_kayitlari: list[AtlananEk] = field(default_factory=list)

    def atla(self, ek: AtlananEk, uyari: bool = False) -> None:
        (_log.warning if uyari else _log.info)("Ek atlandi %s: %s", ek.ad, ek.sebep)
        self.atlanan_kayitlari.append(ek)
        if self.atlananlar is not None:
            self.atlananlar.append(ek)

    def kapsa(self, k: Kapsayici) -> None:
        if self.kapsayicilar is not None:
            self.kapsayicilar.append(k)

    def kisa_dizin(self, ust: Path, tur: str, etiket: str) -> Path:
        """Kisa, benzersiz alt klasor: 'm01_a1b2c3' (mail) / 'z02_d4e5f6' (arsiv).

        Windows'ta toplam yol 260 karakteri asinca dosya yazimi patlar; mail
        konulari uzun olur ve ic ice mesajlarda bu siniri kolayca gecer.
        Konunun tamami zaten `zincir` icinde saklaniyor, klasor adinin
        okunakli olmasina gerek yok. Sayac benzersizligi, kisa ozet ise
        klasoru elle incelerken ipucunu saglar.
        """
        self.sayac += 1
        ozet = hashlib.sha1(str(etiket).encode("utf-8", "replace")).hexdigest()[:6]
        return ust / f"{tur}{self.sayac:02d}_{ozet}"


def _guvenli_ad(ad: str, varsayilan: str = "adsiz", azami: int = 60) -> str:
    """Dosya adini isletim sistemi icin guvenli hale getirir (DISK adi).

    Outlook ekleri '>>: Konu' gibi adlar tasiyabiliyor, zip icinde de
    '#U0131' seklinde kacislanmis Turkce karakterler cikiyor.

    Kisaltma UZANTIYI KORUR. Onceki surum adi duz kesiyordu; 150 karakterden
    uzun bir ek adinda '.xlsx' kirpiliyor, dosya uzantisiz kaliyor ve tablo
    olarak taninmiyordu. Windows'ta uzun mail konulari bunu kolayca tetikler.

    Raporlarda bu ad DEGIL ``_gosterim_adi`` kullanilir.
    """
    if not ad:
        return varsayilan
    # Zip araclarinin urettigi #Uxxxx kacislarini geri cevir
    ad = re.sub(r"#U([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), ad)
    # Yol ayiricilarini ve yasak karakterleri temizle
    ad = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", ad)
    ad = ad.strip(" .")
    if not ad:
        return varsayilan
    if len(ad) <= azami:
        return ad
    kok, nokta, uzanti = ad.rpartition(".")
    if nokta and 0 < len(uzanti) <= 5:
        pay = max(1, azami - len(uzanti) - 1)
        return f"{kok[:pay]}.{uzanti}"
    return ad[:azami]


def _gosterim_adi(ham: str | None, varsayilan: str = "ek") -> str:
    """Kullaniciya gosterilecek ad: maildeki orijinal ek adi, kirpilmadan.

    Yalnizca zip araclarinin #Uxxxx kacislari cozulur ve bastaki/sondaki
    bosluk atilir; uzunluk kirpma, '_2' eki ya da karakter degisimi YOKTUR.
    Boylece rapordaki ad maildekiyle birebir ayni olur ve ad tabanli
    kontroller (ayni adli dosya, fatura anahtari) dogru calisir.
    """
    if not ham:
        return varsayilan
    ad = re.sub(r"#U([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), str(ham)).strip()
    return ad or varsayilan


def _govde_yakala(msg, konu, gonderen, tarih, zincir, derinlik) -> "MailGovdesi":
    """Mesajin govdesini okur; RTF cok buyukse HTML'i atlar, duz metni alir."""
    duz = ""
    try:
        duz = _govde_metni(getattr(msg, "body", None))
    except Exception as hata:  # noqa: BLE001 - govde okunamadi diye mail dusmesin
        _log.debug("govde duz metni okunamadi (%s): %s", konu, hata)
    html, atlandi = "", False
    try:
        rtf = getattr(msg, "rtfBody", None)
        if rtf is not None and len(rtf) > GOVDE_AZAMI_RTF:
            atlandi = True
        else:
            html = _govde_metni(getattr(msg, "htmlBody", None))
    except Exception as hata:  # noqa: BLE001
        _log.debug("govde HTML okunamadi (%s): %s", konu, hata)
    return MailGovdesi(konu=konu, gonderen=gonderen, tarih=tarih, zincir=list(zincir),
                       derinlik=derinlik, html=html, duz=duz, html_atlandi=atlandi)


def _govde_metni(v) -> str:
    """extract_msg govdesini (bytes ya da str) UTF-8 metne cevirir; yoksa bos."""
    if v is None:
        return ""
    if isinstance(v, bytes):
        for kodlama in ("utf-8", "cp1254", "latin-1"):
            try:
                return v.decode(kodlama)
            except UnicodeDecodeError:
                continue
        return v.decode("utf-8", "replace")
    return str(v)


def _tarihe_cevir(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _ozet_bayt(veri: bytes | bytearray) -> str:
    return hashlib.sha256(bytes(veri)).hexdigest()


def _ozet_dosya(yol: Path) -> str | None:
    """Dosya iceriginin sha256 ozeti; okunamazsa None."""
    try:
        h = hashlib.sha256()
        with open(yol, "rb") as f:
            for parca in iter(lambda: f.read(1 << 20), b""):
                h.update(parca)
        return h.hexdigest()
    except OSError:
        return None


def _boyut(yol: Path) -> int | None:
    try:
        return yol.stat().st_size
    except OSError:
        return None


def _benzersiz_yol(dizin: Path, ad: str) -> Path:
    """Ayni adda dosya varsa sonuna sayi ekler (yalnizca DISK adi degisir)."""
    hedef = dizin / ad
    if not hedef.exists():
        return hedef
    kok, uzanti = hedef.stem, hedef.suffix
    for i in range(2, 500):
        aday = dizin / f"{kok}_{i}{uzanti}"
        if not aday.exists():
            return aday
    raise RuntimeError(f"Benzersiz dosya adi uretilemedi: {ad}")


def _ek_adi_oku(ek) -> tuple[str, BaseException | None]:
    """Ekin adini okur. Outlook 365 bulut ekleri (WebAttachment) ad icin
    NotImplementedError firlatir; getattr bunu yakalamaz."""
    try:
        return (getattr(ek, "longFilename", None) or getattr(ek, "shortFilename", None) or ""), None
    except Exception as e:  # noqa: BLE001
        return "", e


def _bulut_adi(ek, ham_ad: str) -> str:
    """Bulut eki icin gosterilecek ad: varsa dosya adi, yoksa baglantidan, yoksa 'bulut eki'."""
    if ham_ad:
        return _gosterim_adi(ham_ad, "bulut eki")
    for ozellik in ("name", "url"):
        try:
            deger = getattr(ek, ozellik, None)
        except Exception:  # noqa: BLE001
            continue
        if deger:
            return _gosterim_adi(unquote(str(deger).rstrip("/").rsplit("/", 1)[-1].split("?")[0]), "bulut eki")
    return "bulut eki"


# ---------------------------------------------------------------------------
# Zip arsivleri
# ---------------------------------------------------------------------------

@dataclass
class _ZipGirdisi:
    yol: Path            # diske yazilan dosya
    orijinal_ad: str     # arsivdeki ad (klasor yolu haric, #U kacislari cozulmus)
    boyut: int


@dataclass
class ZipSonucu:
    """Bir arsivin acilma sonucu: cikan dosyalar, cikarilamayan girdiler, genel hata."""

    cikanlar: list[_ZipGirdisi] = field(default_factory=list)
    #: (girdi adi, sebep, boyut) -- sifreli ya da bozuk girdiler
    atlananlar: list[tuple[str, str, int | None]] = field(default_factory=list)
    #: Arsivin tamami acilamadiysa sebebi (bozuk, zip degil, guvenlik siniri)
    hata: str | None = None
    girdi_sayisi: int = 0


def _sifreli_hatasi_mi(e: BaseException) -> bool:
    m = str(e).lower()
    return "encrypted" in m or "password" in m


def _guvenlik_siniri(girdi_sayisi: int, toplam: int, yuruyus: _Yuruyus | None) -> str | None:
    """Zip bombasi kontrolu; sinir asiliyorsa kullaniciya gosterilecek sebep."""
    mb = AZAMI_ACILMIS_BOYUT // (1024 * 1024)
    if girdi_sayisi > AZAMI_DOSYA_SAYISI:
        return (f"arsivde {girdi_sayisi} dosya var; guvenlik siniri "
                f"({AZAMI_DOSYA_SAYISI} dosya) asildi, acilmadi")
    if toplam > AZAMI_ACILMIS_BOYUT:
        return (f"arsiv acilinca {toplam / (1024 * 1024):.0f} MB olacak; guvenlik siniri "
                f"({mb} MB) asildi, acilmadi")
    if yuruyus is not None and (yuruyus.acilan_dosya + girdi_sayisi > AZAMI_DOSYA_SAYISI
                                or yuruyus.acilan_bayt + toplam > AZAMI_ACILMIS_BOYUT):
        return ("bu mailde acilan arsivlerin toplami guvenlik sinirini asti "
                f"({AZAMI_DOSYA_SAYISI} dosya / {mb} MB); bu arsiv acilmadi")
    return None


def _zip_ac_ayrintili(zip_yolu: Path, hedef: Path, yuruyus: _Yuruyus | None = None) -> ZipSonucu:
    """Zip arsivini guvenli sekilde acar; her girdinin akibetini dondurur.

    Girdiler ``hedef`` altina DUZ yazilir (arsiv ici klasor yapisi korunmaz;
    Windows yol siniri icin). Ayni adli girdiler '_2' eki alir; gosterilen ad
    yine arsivdeki addir. Yol gecisi (path traversal) ve zip bombasina karsi
    korumali. Sifreli ya da bozuk bir girdi arsivin geri kalanini engellemez:
    o girdi ``atlananlar``a yazilir, digerleri cikarilir. Hicbir istisna
    disari sizmaz; arsivin tamami acilamiyorsa ``hata`` doludur.
    """
    sonuc = ZipSonucu()
    try:
        hedef.mkdir(parents=True, exist_ok=True)
        kok = hedef.resolve()
        with zipfile.ZipFile(zip_yolu) as zf:
            girdiler = [g for g in zf.infolist() if not g.is_dir()]
            sonuc.girdi_sayisi = len(girdiler)
            toplam = sum(g.file_size for g in girdiler)
            sinir = _guvenlik_siniri(len(girdiler), toplam, yuruyus)
            if sinir:
                sonuc.hata = sinir
                return sonuc
            if yuruyus is not None:
                yuruyus.acilan_dosya += len(girdiler)
                yuruyus.acilan_bayt += toplam
            for g in girdiler:
                taban = Path(g.filename.replace("\\", "/")).name
                orijinal = _gosterim_adi(taban, "girdi")
                if g.flag_bits & 0x1:
                    # zipfile sifreli girdi icin RuntimeError firlatir; onceden yakala
                    sonuc.atlananlar.append((orijinal, SEBEP_SIFRELI, g.file_size))
                    continue
                cikti = _benzersiz_yol(hedef, _guvenli_ad(taban, "girdi", DISK_AD_AZAMI))
                # Yol gecisi kontrolu (duzlestirme sonrasi da kalsin)
                if not str(cikti.resolve()).startswith(str(kok)):
                    sonuc.atlananlar.append(
                        (orijinal, "arsiv disina cikan yol (guvenlik); acilmadi", g.file_size))
                    continue
                try:
                    with zf.open(g) as kaynak, open(cikti, "wb") as f:
                        shutil.copyfileobj(kaynak, f)
                except OSError as e:  # yol cok uzun, disk dolu, izin
                    cikti.unlink(missing_ok=True)
                    sonuc.atlananlar.append((orijinal, _yazma_sebebi(e), g.file_size))
                    continue
                except Exception as e:  # sifreli (RuntimeError), bozuk (BadZipFile, zlib.error),
                    #                     desteklenmeyen sikistirma (NotImplementedError)
                    cikti.unlink(missing_ok=True)
                    sebep = (SEBEP_SIFRELI if _sifreli_hatasi_mi(e)
                             else f"arsivden cikarilamadi; girdi bozuk olabilir ({e.__class__.__name__}: {e})")
                    sonuc.atlananlar.append((orijinal, sebep, g.file_size))
                    continue
                sonuc.cikanlar.append(_ZipGirdisi(cikti, orijinal, g.file_size))
    except zipfile.BadZipFile as e:
        sonuc.hata = f"arsiv acilamadi; bozuk ya da zip degil ({e})"
    except OSError as e:
        sonuc.hata = f"arsiv gecici dizine acilamadi ({e.__class__.__name__}: {e}); yol cok uzun ya da izin yok olabilir"
    except Exception as e:  # noqa: BLE001 - arsiv yuzunden mail dusmesin
        sonuc.hata = f"arsiv acilamadi ({e.__class__.__name__}: {e})"
    if sonuc.hata:
        _log.warning("Arsiv acilamadi %s: %s", zip_yolu.name, sonuc.hata)
    for ad, sebep, _ in sonuc.atlananlar:
        _log.warning("Arsiv girdisi cikarilamadi %s / %s: %s", zip_yolu.name, ad, sebep)
    return sonuc


def _zip_ac(zip_yolu: Path, hedef: Path) -> list[Path]:
    """Zip arsivini acar, cikan dosya yollarini dondurur (geriye uyumlu sarmal)."""
    return [g.yol for g in _zip_ac_ayrintili(zip_yolu, hedef).cikanlar]


def _zip_isle(
    zip_yolu: Path, zip_adi: str, hedef: Path, zincir: list[str], derinlik: int, y: _Yuruyus,
    konu: str | None, gonderen: str | None, tarih: date | None, boyut: int | None,
) -> None:
    """Bir zip ekini acar; icindeki tablolari, mailleri ve IC ARSIVLERI isler.

    ``derinlik`` arsivin bulundugu seviyedir; icindekiler bir seviye derindedir.
    ``hedef`` ic maillerin acilacagi klasor (bulundugu mailin dali).
    """
    if derinlik > AZAMI_DERINLIK:
        y.atla(AtlananEk(ad=zip_adi, zincir=list(zincir), sebep=_derinlik_sebebi(), boyut=boyut), uyari=True)
        return
    kapsayici = Kapsayici(ad=zip_adi, tur="arsiv", zincir=list(zincir), boyut=boyut, derinlik=derinlik,
                          konu=konu, gonderen=gonderen, tarih=tarih)
    y.kapsa(kapsayici)
    acilan = _zip_ac_ayrintili(zip_yolu, y.kisa_dizin(zip_yolu.parent, "z", zip_adi), y)
    kapsayici.ek_sayisi = acilan.girdi_sayisi
    if acilan.hata:
        kapsayici.aciklama = f"arsiv; {acilan.hata}"
        y.atla(AtlananEk(ad=zip_adi, zincir=list(zincir), sebep=acilan.hata, boyut=boyut), uyari=True)
        return
    ic_zincir = zincir + [zip_adi]
    for girdi_adi, sebep, g_boyut in acilan.atlananlar:
        y.atla(AtlananEk(ad=girdi_adi, zincir=ic_zincir, sebep=sebep, boyut=g_boyut), uyari=True)
    for g in acilan.cikanlar:
        uz = g.yol.suffix.lower()
        if uz in CIKARILAN_UZANTILAR:
            y.sonuc.append(CikarilanEk(
                yol=g.yol, ad=g.yol.name, mail_konusu=konu, mail_gonderen=gonderen,
                mail_tarihi=tarih, zincir=ic_zincir, derinlik=derinlik + 1,
                orijinal_ad=g.orijinal_ad, ozet=_ozet_dosya(g.yol)))
        elif uz == ".msg":
            _msg_ac_ve_yuru(g.yol, hedef, ic_zincir, derinlik + 1, y.sonuc, y.atlananlar,
                            yuruyus=y, ek_adi=g.orijinal_ad, boyut=g.boyut)
        elif uz == ".zip":
            _zip_isle(g.yol, g.orijinal_ad, hedef, ic_zincir, derinlik + 1, y,
                      konu, gonderen, tarih, g.boyut)
        elif uz not in GORSEL_UZANTILARI:
            y.atla(AtlananEk(ad=g.orijinal_ad, zincir=ic_zincir,
                             sebep=f"tablo degil ({uz or 'uzantisiz'}); acilmadi", boyut=g.boyut,
                             ozet=_ozet_dosya(g.yol)))
    sifreli = sum(1 for _, s, _ in acilan.atlananlar if s == SEBEP_SIFRELI)
    diger = len(acilan.atlananlar) - sifreli
    kapsayici.aciklama = f"arsiv; {acilan.girdi_sayisi} girdi"
    if sifreli:
        kapsayici.aciklama += f", {sifreli} sifreli (parola gerekli)"
    if diger:
        kapsayici.aciklama += f", {diger} girdi cikarilamadi"


# ---------------------------------------------------------------------------
# Mesajlar
# ---------------------------------------------------------------------------

def _msg_yuru(
    msg,
    hedef: Path,
    zincir: list[str],
    derinlik: int,
    sonuc: list[CikarilanEk],
    atlananlar: list | None = None,
    *,
    yuruyus: _Yuruyus | None = None,
    ek_adi: str | None = None,
    boyut: int | None = None,
) -> None:
    """Bir mesaj nesnesinin eklerini yurur, tablo dosyalarini sonuc listesine ekler.

    Tablo olmayan ekler (PDF fatura, docx...) okunmaz ama ``atlananlar``
    listesine ``AtlananEk`` olarak yazilir; kullanici neyin okunmadigini
    Excel'de gorur. Sessiz atlama yok. Tek bir ekin hatasi (bozuk ek, sifreli
    arsiv, bulut baglantisi, yazilamayan dosya, cok derin ic mail) yalnizca o
    eki dusurur; mailin geri kalani okunmaya devam eder.

    ``ek_adi`` verilirse bu mesaj bir EK olarak geldi (ic mail); ``Kapsayici``
    kaydi uretilir. Kok mail icin None.
    """
    y = yuruyus if yuruyus is not None else _Yuruyus(sonuc=sonuc, atlananlar=atlananlar)

    if derinlik > AZAMI_DERINLIK:
        y.atla(AtlananEk(
            ad=ek_adi or _gosterim_adi(getattr(msg, "subject", None), "konusuz"),
            zincir=list(zincir), sebep=_derinlik_sebebi(), boyut=boyut), uyari=True)
        return

    konu = _guvenli_ad(getattr(msg, "subject", None) or "konusuz", "konusuz")
    gonderen = getattr(msg, "sender", None)
    tarih = _tarihe_cevir(getattr(msg, "date", None))
    yeni_zincir = zincir + [konu]
    ekler = list(getattr(msg, "attachments", None) or [])

    if y.govdeler is not None:
        y.govdeler.append(_govde_yakala(msg, konu, gonderen, tarih, zincir, derinlik))

    if ek_adi is not None:
        y.kapsa(Kapsayici(
            ad=ek_adi, tur="mail", zincir=list(zincir), boyut=boyut, derinlik=derinlik,
            konu=konu, gonderen=gonderen, tarih=tarih, ek_sayisi=len(ekler),
            aciklama=f"ekli mail; {len(ekler)} ek; konu: {konu}"))

    # Klasor adi KISA ve konudan bagimsiz ('m01_a1b2c3'); bkz. _Yuruyus.kisa_dizin.
    dal = y.kisa_dizin(hedef, "m", konu)
    dal.mkdir(parents=True, exist_ok=True)

    for ek in ekler:
        ham_ad, ad_hatasi = _ek_adi_oku(ek)
        ad = _guvenli_ad(ham_ad, "ek", DISK_AD_AZAMI)   # disk adi (kisa)
        orijinal = _gosterim_adi(ham_ad, ad)            # rapor adi (orijinal)
        uzanti = Path(ad).suffix.lower()

        # Mail imzasindaki logolar isimize yaramaz
        if uzanti in GORSEL_UZANTILARI:
            continue

        # Ek verisine erisim hata verebilir (bozuk mail, bulut eki). Yalnizca bu ek dusuyor.
        try:
            veri = ek.data
            # hasattr yalnizca AttributeError yutar; bozuk bir ic mailin
            # 'attachments' ozelligi baska hata verirse o da burada yakalansin.
            gomulu_mesaj = hasattr(veri, "attachments")
        except NotImplementedError:
            # extract_msg WebAttachment: Outlook 365 bulut eki / OneDrive baglantisi.
            # Icerik mailde degil; ad da okunamayabilir.
            y.atla(AtlananEk(ad=_bulut_adi(ek, ham_ad), zincir=list(yeni_zincir), sebep=SEBEP_BULUT))
            continue
        except Exception as e:  # noqa: BLE001
            y.atla(AtlananEk(
                ad=orijinal, zincir=list(yeni_zincir),
                sebep=(f"ek okunamadi; mail bozuk olabilir ({e.__class__.__name__}: {e}). "
                       "Eki Outlook'ta ayrica kaydedip gonderin")), uyari=True)
            continue

        if ad_hatasi is not None:
            # Adi okunamayan ama verisi gelen ek: yine de okumaya calis, ad yedek olsun.
            _log.warning("Ek adi okunamadi (%s); '%s' olarak devam", ad_hatasi, orijinal)

        # Ekli mesaj: ic ice in. Icerideki bir hata yalnizca o ic maili dusurur.
        if gomulu_mesaj:
            ic_ad = orijinal if ham_ad else _gosterim_adi(getattr(veri, "subject", None), "ekli mail")
            try:
                _msg_yuru(veri, dal, yeni_zincir, derinlik + 1, y.sonuc, y.atlananlar,
                          yuruyus=y, ek_adi=ic_ad)
            except Exception as e:  # noqa: BLE001
                y.atla(AtlananEk(
                    ad=ic_ad, zincir=list(yeni_zincir),
                    sebep=(f"ekli mail okunamadi ({e.__class__.__name__}: {e}). "
                           "Maili Outlook'ta acip ekleri ayrica gonderin")), uyari=True)
            continue

        if not isinstance(veri, (bytes, bytearray)):
            y.atla(AtlananEk(
                ad=orijinal, zincir=list(yeni_zincir),
                sebep="ek icerigi mailde yok (bulut baglantisi ya da bos ek olabilir); acilmadi"))
            continue

        # Diske yaz. Yol cok uzun / izin yok / disk dolu: yalnizca bu ek dusuyor.
        try:
            yol = _benzersiz_yol(dal, ad if uzanti else ad + ".bin")
            yol.write_bytes(veri)
        except OSError as e:
            y.atla(AtlananEk(ad=orijinal, zincir=list(yeni_zincir), sebep=_yazma_sebebi(e),
                             boyut=len(veri), ozet=_ozet_bayt(veri)), uyari=True)
            continue

        if uzanti == ".zip":
            _zip_isle(yol, orijinal, dal, yeni_zincir, derinlik, y, konu, gonderen, tarih, len(veri))
            continue

        if uzanti == ".msg":
            _msg_ac_ve_yuru(yol, dal, yeni_zincir, derinlik + 1, y.sonuc, y.atlananlar,
                            yuruyus=y, ek_adi=orijinal, boyut=len(veri))
            continue

        if uzanti in CIKARILAN_UZANTILAR:
            y.sonuc.append(CikarilanEk(
                yol=yol, ad=yol.name, mail_konusu=konu, mail_gonderen=gonderen,
                mail_tarihi=tarih, zincir=list(yeni_zincir), derinlik=derinlik,
                orijinal_ad=orijinal, ozet=_ozet_bayt(veri)))
        else:
            y.atla(AtlananEk(
                ad=orijinal, zincir=list(yeni_zincir),
                sebep=f"tablo degil ({uzanti or 'uzantisiz'}); acilmadi", boyut=len(veri),
                ozet=_ozet_bayt(veri)))


class MesajAcilamadi(Exception):
    """Kullanicinin verdigi .msg dosyasi Outlook mesaji olarak acilamadi."""


def _msg_ac_ve_yuru(
    yol: Path, hedef: Path, zincir: list[str], derinlik: int, sonuc: list[CikarilanEk],
    atlananlar: list | None = None,
    *,
    yuruyus: _Yuruyus | None = None,
    ek_adi: str | None = None,
    boyut: int | None = None,
) -> None:
    import extract_msg

    y = yuruyus if yuruyus is not None else _Yuruyus(sonuc=sonuc, atlananlar=atlananlar)
    try:
        m = extract_msg.openMsg(str(yol))
    except Exception as e:
        if derinlik == 0:
            # Kullanicinin verdigi dosyanin kendisi acilmiyor: bos, bozuk ya da
            # .msg degil. 'Tablo eki bulunamadi' demek yaniltir; nedeni soyle.
            raise MesajAcilamadi(
                f"dosya Outlook mesaji olarak acilamadi ({e.__class__.__name__}). "
                "Dosya bos, bozuk ya da .msg uzantili baska bir dosya olabilir; "
                "Outlook'ta acip 'Farkli Kaydet' ile yeniden kaydedin."
            ) from e
        y.atla(AtlananEk(
            ad=ek_adi or yol.name, zincir=list(zincir), boyut=boyut,
            sebep=(f"ekli mail acilamadi; bozuk olabilir ({e.__class__.__name__}). "
                   "Maili Outlook'ta acip ekleri ayrica kaydedin")), uyari=True)
        return
    try:
        _msg_yuru(m, hedef, zincir, derinlik, y.sonuc, y.atlananlar,
                  yuruyus=y, ek_adi=ek_adi, boyut=boyut)
    except Exception as e:  # noqa: BLE001
        if derinlik == 0:
            raise
        # Ic mailin icindeki bir hata ana mailin diger eklerini dusurmesin.
        y.atla(AtlananEk(
            ad=ek_adi or yol.name, zincir=list(zincir), boyut=boyut,
            sebep=(f"ekli mail okunamadi ({e.__class__.__name__}: {e}). "
                   "Maili Outlook'ta acip ekleri ayrica gonderin")), uyari=True)
    finally:
        try:
            m.close()
        except Exception:
            pass


def _tekrar_sebebi(ilk_ad: str, ad: str, tablo: bool) -> str:
    sebep = ("icerigi daha once cikan bir ekle birebir ayni "
             "(mail icinde iletilmis / tekrar eklenmis); "
             + ("ilk kopya okundu" if tablo else "tekrar sayilmadi"))
    if ilk_ad != ad:
        sebep += f" (ilk kopya: {ilk_ad})"
    return sebep


def msg_aciklarini_cikar(msg_yolu: str | Path, hedef_dizin: str | Path,
                         atlananlar: list | None = None,
                         kapsayicilar: list | None = None,
                         govdeler: list | None = None) -> list[CikarilanEk]:
    """Bir .msg dosyasindaki butun tablo eklerini ic ice arsiv ve mesajlarla birlikte cikarir.

    Args:
        msg_yolu: Outlook mesaj dosyasi.
        hedef_dizin: Cikarilan dosyalarin yazilacagi dizin. Yoksa olusturulur.
            Altindaki klasor ve dosya adlari KISA tutulur (Windows yol siniri);
            orijinal adlar donen nesnelerdedir.
        atlananlar: Verilirse okunmayan ekler (PDF, sifreli arsiv, bozuk ek,
            bulut baglantisi, tekrar eden ek...) bu listeye ``AtlananEk``
            olarak yazilir.
        kapsayicilar: Verilirse ic mailler ve arsivler bu listeye ``Kapsayici``
            olarak yazilir (envanterde 'MAIL' / 'ARSIV' bilgi satiri icin).

    Returns:
        Bulunan tablo dosyalarinin listesi. Her biri hangi mailden geldigini
        ve maildeki orijinal adini (``gosterim_adi``) tasir. Icerigi birebir
        ayni (sha256) olan ikinci kopyalar listeye girmez; ``atlananlar``a
        ``tekrar=True`` ile yazilir. Tablo olmayan eklerin (PDF) tekrarlari da
        ``tekrar=True`` isaretlenir; boylece 'okunmayan ek' sayisi FARKLI
        dosya sayisini verir.

    Raises:
        ImportError: extract-msg kurulu degilse.
        FileNotFoundError: msg_yolu yoksa.
        MesajAcilamadi: dosyanin kendisi Outlook mesaji olarak acilamiyorsa.
    """
    try:
        import extract_msg  # noqa: F401
    except ImportError as e:
        raise ImportError(
            "Outlook mesajlarini okumak icin extract-msg gerekli. "
            "Kurmak icin: pip install extract-msg"
        ) from e

    msg_yolu = Path(msg_yolu)
    if not msg_yolu.is_file():
        raise FileNotFoundError(f"Mesaj dosyasi bulunamadi: {msg_yolu}")

    hedef = Path(hedef_dizin)
    hedef.mkdir(parents=True, exist_ok=True)

    y = _Yuruyus(sonuc=[], atlananlar=atlananlar, kapsayicilar=kapsayicilar, govdeler=govdeler)
    _msg_ac_ve_yuru(msg_yolu, hedef, [], 0, y.sonuc, atlananlar, yuruyus=y)

    # Ayni dosyanin iki kez cikmasini onle (ayni mail iki kez forward edilmis
    # olabilir). Olcut ICERIK ozeti; ad + boyut YETMEZ: tedarikcinin fatura
    # basina gonderdigi tek satirlik 'Fatura Detayi.xlsx' sablonlari ayni adi
    # ve ayni boyutu tasir ama farkli para tasir; ikincisini atmak para kaybidir.
    gorulen: dict[str, str] = {}      # ozet -> ilk gorulen ekin gosterim adi
    benzersiz: list[CikarilanEk] = []
    for ek in y.sonuc:
        ozet = ek.ozet or _ozet_dosya(ek.yol)
        if ozet is None:
            y.atla(AtlananEk(ad=ek.gosterim_adi, zincir=list(ek.zincir),
                             sebep="cikarilan dosya diskten geri okunamadi"), uyari=True)
            continue
        ek.ozet = ozet
        if ozet in gorulen:
            # Sessiz atlama yok: envanterde 'AYNI ICERIK' olarak gorunsun.
            y.atla(AtlananEk(ad=ek.gosterim_adi, zincir=list(ek.zincir), boyut=_boyut(ek.yol),
                             tekrar=True, ozet=ozet,
                             sebep=_tekrar_sebebi(gorulen[ozet], ek.gosterim_adi, tablo=True)))
            continue
        gorulen[ozet] = ek.gosterim_adi
        benzersiz.append(ek)

    # Tablo olmayan ekler de tekillenir: ayni PDF hem zip'te hem ic mailde
    # gelince kapak '17 ek okunmadi' derken aslinda 13 farkli dosya vardir.
    for a in y.atlanan_kayitlari:
        if a.tekrar or not a.ozet:
            continue
        if a.ozet in gorulen:
            a.tekrar = True
            a.sebep = _tekrar_sebebi(gorulen[a.ozet], a.ad, tablo=False)
            continue
        gorulen[a.ozet] = a.ad
    return benzersiz


def msg_mi(yol: str | Path) -> bool:
    """Dosya bir Outlook mesaji mi."""
    return Path(yol).suffix.lower() == ".msg"
