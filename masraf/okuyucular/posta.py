"""Outlook .msg dosyalarindan tablo eklerini cikarir.

Finans ekibine gelen dosyalar Outlook mesaji olarak geliyor. Bir mesajin icinde
baska mesajlar, onlarin icinde zip arsivleri, onlarin icinde de Excel dosyalari
olabiliyor. Bu modul bu agaci sonuna kadar yuruyup butun tablo dosyalarini
duz bir listeye cikarir ve her dosyanin hangi mail zincirinden geldigini kaydeder.

Kullanim:
    from masraf.okuyucular.posta import msg_aciklarini_cikar
    ekler = msg_aciklarini_cikar("gelen.msg", "cikti/acilan")
    for ek in ekler:
        print(ek.yol, ek.mail_konusu, ek.mail_tarihi)

Bagimlilik: extract-msg. Windows uzerinde `pip install extract-msg` yeterlidir.
"""

from __future__ import annotations

import logging
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

_log = logging.getLogger(__name__)

# Cikarilacak tablo uzantilari
TABLO_UZANTILARI = {".xlsx", ".xls", ".xlsm", ".csv", ".tsv"}
# Isimize yaramayan, atlanacak uzantilar (mail imzasindaki logolar vb)
GORSEL_UZANTILARI = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emf", ".wmf", ".ico"}

# Ic ice inme siniri. Sonsuz donguye karsi guvenlik.
AZAMI_DERINLIK = 6

# Zip bombasi korumasi
AZAMI_ACILMIS_BOYUT = 500 * 1024 * 1024  # 500 MB
AZAMI_DOSYA_SAYISI = 2000


@dataclass
class CikarilanEk:
    """Mail agacindan cikarilmis tek bir dosya ve nereden geldigi."""

    yol: Path
    ad: str
    mail_konusu: str | None = None
    mail_gonderen: str | None = None
    mail_tarihi: date | None = None
    # Kokten bu dosyaya kadar olan yol: ["ana mail", "ekli mail", "arsiv.zip"]
    zincir: list[str] = field(default_factory=list)
    derinlik: int = 0

    @property
    def kaynak_aciklamasi(self) -> str:
        """Kullaniciya gosterilecek okunakli kaynak zinciri."""
        if not self.zincir:
            return self.ad
        return " > ".join(self.zincir + [self.ad])


def _guvenli_ad(ad: str, varsayilan: str = "adsiz") -> str:
    """Dosya adini isletim sistemi icin guvenli hale getirir.

    Outlook ekleri '>>: Konu' gibi adlar tasiyabiliyor, zip icinde de
    '#U0131' seklinde kacislanmis Turkce karakterler cikiyor.
    """
    if not ad:
        return varsayilan
    # Zip araclarinin urettigi #Uxxxx kacislarini geri cevir
    ad = re.sub(r"#U([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), ad)
    # Yol ayiricilarini ve yasak karakterleri temizle
    ad = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", ad)
    ad = ad.strip(" .")
    return ad[:150] or varsayilan


def _tarihe_cevir(v) -> date | None:
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def _benzersiz_yol(dizin: Path, ad: str) -> Path:
    """Ayni adda dosya varsa sonuna sayi ekler."""
    hedef = dizin / ad
    if not hedef.exists():
        return hedef
    kok, uzanti = hedef.stem, hedef.suffix
    for i in range(2, 500):
        aday = dizin / f"{kok}_{i}{uzanti}"
        if not aday.exists():
            return aday
    raise RuntimeError(f"Benzersiz dosya adi uretilemedi: {ad}")


def _zip_ac(zip_yolu: Path, hedef: Path) -> list[Path]:
    """Zip arsivini guvenli sekilde acar, cikan dosya yollarini dondurur.

    Yol gecisi (path traversal) ve zip bombasina karsi korumali.
    """
    cikanlar: list[Path] = []
    hedef.mkdir(parents=True, exist_ok=True)
    kok = hedef.resolve()
    try:
        with zipfile.ZipFile(zip_yolu) as zf:
            girdiler = zf.infolist()
            if len(girdiler) > AZAMI_DOSYA_SAYISI:
                _log.warning("%s icinde %d dosya var, atlaniyor", zip_yolu.name, len(girdiler))
                return cikanlar
            toplam = sum(g.file_size for g in girdiler)
            if toplam > AZAMI_ACILMIS_BOYUT:
                _log.warning("%s acilinca %d bayt olacak, atlaniyor", zip_yolu.name, toplam)
                return cikanlar
            for g in girdiler:
                if g.is_dir():
                    continue
                parcalar = [_guvenli_ad(p) for p in Path(g.filename).parts if p not in ("..", "/", "\\")]
                if not parcalar:
                    continue
                cikti = hedef.joinpath(*parcalar)
                # Yol gecisi kontrolu
                if not str(cikti.resolve()).startswith(str(kok)):
                    _log.warning("Arsiv disina cikan yol atlandi: %s", g.filename)
                    continue
                cikti.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(g) as kaynak, open(cikti, "wb") as f:
                    shutil.copyfileobj(kaynak, f)
                cikanlar.append(cikti)
    except (zipfile.BadZipFile, OSError) as e:
        _log.warning("Arsiv acilamadi %s: %s", zip_yolu.name, e)
    return cikanlar


def _msg_yuru(
    msg,
    hedef: Path,
    zincir: list[str],
    derinlik: int,
    sonuc: list[CikarilanEk],
) -> None:
    """Bir mesaj nesnesinin eklerini yurur, tablo dosyalarini sonuc listesine ekler."""
    if derinlik > AZAMI_DERINLIK:
        _log.warning("Azami derinlik asildi, dal atlandi: %s", " > ".join(zincir))
        return

    konu = _guvenli_ad(getattr(msg, "subject", None) or "konusuz", "konusuz")
    gonderen = getattr(msg, "sender", None)
    tarih = _tarihe_cevir(getattr(msg, "date", None))
    yeni_zincir = zincir + [konu]

    dal = hedef / f"{derinlik:02d}_{konu}"
    dal.mkdir(parents=True, exist_ok=True)

    for ek in getattr(msg, "attachments", []) or []:
        ham_ad = getattr(ek, "longFilename", None) or getattr(ek, "shortFilename", None) or ""
        ad = _guvenli_ad(ham_ad, "ek")
        uzanti = Path(ad).suffix.lower()

        # Mail imzasindaki logolar isimize yaramaz
        if uzanti in GORSEL_UZANTILARI:
            continue

        # Ekli mesaj: ic ice in
        gomulu = getattr(ek, "data", None)
        if hasattr(gomulu, "attachments"):
            _msg_yuru(gomulu, dal, yeni_zincir, derinlik + 1, sonuc)
            continue

        # Diske yaz
        try:
            veri = ek.data
        except Exception as e:  # bozuk ek
            _log.warning("Ek okunamadi %s: %s", ad, e)
            continue
        if not isinstance(veri, (bytes, bytearray)):
            continue

        yol = _benzersiz_yol(dal, ad if uzanti else ad + ".bin")
        yol.write_bytes(veri)

        if uzanti == ".zip":
            for icerik in _zip_ac(yol, dal / (yol.stem + "_acilmis")):
                ic_uzanti = icerik.suffix.lower()
                if ic_uzanti in TABLO_UZANTILARI:
                    sonuc.append(CikarilanEk(
                        yol=icerik, ad=icerik.name, mail_konusu=konu,
                        mail_gonderen=gonderen, mail_tarihi=tarih,
                        zincir=yeni_zincir + [yol.name], derinlik=derinlik + 1))
                elif ic_uzanti == ".msg":
                    _msg_ac_ve_yuru(icerik, dal, yeni_zincir + [yol.name], derinlik + 1, sonuc)
            continue

        if uzanti == ".msg":
            _msg_ac_ve_yuru(yol, dal, yeni_zincir, derinlik + 1, sonuc)
            continue

        if uzanti in TABLO_UZANTILARI:
            sonuc.append(CikarilanEk(
                yol=yol, ad=yol.name, mail_konusu=konu, mail_gonderen=gonderen,
                mail_tarihi=tarih, zincir=yeni_zincir, derinlik=derinlik))


def _msg_ac_ve_yuru(
    yol: Path, hedef: Path, zincir: list[str], derinlik: int, sonuc: list[CikarilanEk]
) -> None:
    import extract_msg

    try:
        m = extract_msg.openMsg(str(yol))
    except Exception as e:
        _log.warning("Mesaj acilamadi %s: %s", yol.name, e)
        return
    try:
        _msg_yuru(m, hedef, zincir, derinlik, sonuc)
    finally:
        try:
            m.close()
        except Exception:
            pass


def msg_aciklarini_cikar(msg_yolu: str | Path, hedef_dizin: str | Path) -> list[CikarilanEk]:
    """Bir .msg dosyasindaki butun tablo eklerini ic ice arsiv ve mesajlarla birlikte cikarir.

    Args:
        msg_yolu: Outlook mesaj dosyasi.
        hedef_dizin: Cikarilan dosyalarin yazilacagi dizin. Yoksa olusturulur.

    Returns:
        Bulunan tablo dosyalarinin listesi. Her biri hangi mailden geldigini tasir.

    Raises:
        ImportError: extract-msg kurulu degilse.
        FileNotFoundError: msg_yolu yoksa.
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

    sonuc: list[CikarilanEk] = []
    _msg_ac_ve_yuru(msg_yolu, hedef, [], 0, sonuc)

    # Ayni dosyanin iki kez cikmasini onle (ayni mail iki kez forward edilmis olabilir)
    gorulen: set[tuple[str, int]] = set()
    benzersiz: list[CikarilanEk] = []
    for ek in sonuc:
        try:
            anahtar = (ek.ad, ek.yol.stat().st_size)
        except OSError:
            continue
        if anahtar in gorulen:
            continue
        gorulen.add(anahtar)
        benzersiz.append(ek)
    return benzersiz


def msg_mi(yol: str | Path) -> bool:
    """Dosya bir Outlook mesaji mi."""
    return Path(yol).suffix.lower() == ".msg"
