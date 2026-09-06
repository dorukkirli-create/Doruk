"""Ogrenen yardimci defterler: alias, harici kisi, ek kisi ve TCKN koprusu.

Sistem CALISMA ZAMANINDA YAPAY ZEKA KULLANMAZ; ogrenme tamamen kullanicinin
inceleme ekraninda verdigi kararlarin CSV dosyalarina yazilmasiyla olur.
Bir sonraki calistirmada ayni isim otomatik olarak dogru sicile duser.

Yonetilen dosyalar (hepsi ``veri/`` altinda, utf-8-sig + ';' ayirici, Excel'de
cift tiklayarak acilabilir):

======================== ===================================================
Dosya                    Icerik
======================== ===================================================
``aliases.csv``          normalize isim -> sicil (kullanici duzeltmesi)
``harici_kisiler.csv``   calisan OLMAYAN kisiler (dis danisman, grup sirketi)
``ek_kisiler.csv``       personel ana verisinde bulunmayan ama yardimci
                         kaynaklarda (saglik listesi vb) gecen kisiler
``tckn_sicil.csv``       TCKN -> sicil koprusu (ana veride TCKN YOKTUR)
======================== ===================================================

Dosyalar yoksa basliklariyla olusturulur. Her satirda ``kaynak`` ve
``eklenme_tarihi`` kolonlari bulunur, boylece bir kaydin nereden geldigi ve ne
zaman ogrenildigi izlenebilir.

GIZLILIK UYARISI: ``ek_kisiler.csv`` ve ``tckn_sicil.csv`` TC kimlik numarasi
icerir. Bu dosyalarin surum kontrolune girmemesi gerekir; ``veri/`` dizini
.gitignore icinde tutulmalidir.
"""

from __future__ import annotations

import csv
from datetime import date
from pathlib import Path
from typing import Any, Iterable

from masraf.metin import isim_normalize
from masraf.modeller import GiderSatiri

# CSV bicimi. Excel'in Turkce yerel ayarinda liste ayirici ';' oldugu icin
# yazarken bu kullanilir; okurken her iki ayirici da kabul edilir.
AYIRICI = ";"
KODLAMA = "utf-8-sig"

DOSYA_ALIAS = "aliases.csv"
DOSYA_HARICI = "harici_kisiler.csv"
DOSYA_EK_KISI = "ek_kisiler.csv"
DOSYA_TCKN = "tckn_sicil.csv"

# Her dosyanin kolon duzeni. Kolon eklemek gerekirse SONA eklenmelidir;
# eski dosyalar eksik kolonlari bos kabul edilerek okunur.
BASLIKLAR: dict[str, tuple[str, ...]] = {
    DOSYA_ALIAS: ("isim_norm", "sicil", "ad_soyad", "kaynak", "eklenme_tarihi"),
    DOSYA_HARICI: (
        "isim_norm", "ad_soyad", "kurum", "masraf_merkezi", "aciklama",
        "kaynak", "eklenme_tarihi",
    ),
    DOSYA_EK_KISI: (
        "anahtar", "tckn", "ad_soyad", "santiye", "kaynak", "eklenme_tarihi",
    ),
    DOSYA_TCKN: ("tckn", "sicil", "ad_soyad", "kaynak", "eklenme_tarihi"),
}

# yardimci_kaynaktan_besle() icin: hangi kaynak tipleri ek kisi defterini
# besler. Seyahat/fatura dosyalari (antik_cari) BESLEMEZ, cunku oradaki isimler
# dogrulanmis kimlik degil, ham metinden cikarilmis tahminlerdir.
BESLEYEN_KAYNAKLAR: frozenset[str] = frozenset({
    "energo_saglik",
    "energo_arabulucu",
    "energo_assessment",
    "energo_assessment_detay",
    "koc_katilimci",
})

#: Besleme listesinde OLMAYAN kaynak tipleri icin kullaniciya soylenecek neden.
#: Kutuk uyarisi 'defter beslemesinde kullanildi' derken fiilen hicbir satir
#: beslemiyordu (olculdu: ek_kisi 0, atlanan 20919); artik neden yazilir.
BESLEMEYEN_NEDENLER: dict[str, str] = {
    "referans_liste": (
        "sicil tasiyan personel kutugu; kisileri ana veri ve 1C listesiyle zaten "
        "eslesir, ek kisi defteri TCKN'li kucuk listeler icindir"
    ),
    "antik_cari": "kisi adi serbest metinden tahmin ediliyor, dogrulanmis kimlik degil",
    "yuzyil_dagitilmis": "kisi adi serbest metinden tahmin ediliyor, dogrulanmis kimlik degil",
    "genel": "sablonu taninmayan dosya; kolonlar tahminle secildi, dogrulanmis kimlik degil",
}


def besleme_nedeni(kaynak_tip: str) -> str:
    """Bir kaynak tipinin ek kisi defterini neden beslemedigini soyler."""
    if kaynak_tip in BESLEYEN_KAYNAKLAR:
        return "besleme kaynagi"
    return BESLEMEYEN_NEDENLER.get(kaynak_tip, "bu kaynak tipi besleme listesinde degil")


def besleme_aciklamasi(ozet: dict | None, kaynak_tip: str, besleme_acik: bool = True) -> str:
    """Bir kaynak tipinin defteri besleyip beslemedigini finansciya anlatir.

    ``ozet`` ``yardimci_kaynaktan_besle`` ciktisidir; None ise besleme
    henuz yapilmamis ya da yapilamamistir. Donen metin 'defteri besledi: ...'
    ya da 'defteri beslemedi: ...' ile baslar.
    """
    if not besleme_acik:
        return "defteri beslemedi: defter beslemesi bu calistirmada kapali (ayar)"
    if kaynak_tip not in BESLEYEN_KAYNAKLAR:
        return f"defteri beslemedi: {besleme_nedeni(kaynak_tip)}"
    if not ozet:
        return "defteri beslemedi: besleme yapilmadi"
    tip = (ozet.get("tip_ozeti") or {}).get(kaynak_tip)
    if not tip:
        return "defteri beslemedi: bu tipte satir islenmedi"
    parcalar: list[str] = []
    if tip["ek_kisi"]:
        parcalar.append(f"{tip['ek_kisi']} kisi ek kisi defterine yazildi")
    if tip["tckn_kopru"]:
        parcalar.append(f"{tip['tckn_kopru']} TCKN-sicil koprusu kuruldu")
    if tip["calisan"]:
        parcalar.append(f"{tip['calisan']} satir zaten ana verideki calisan (yazilmadi)")
    if tip["kimliksiz"]:
        parcalar.append(f"{tip['kimliksiz']} satirda ad/TCKN yok")
    mevcut = tip["satir"] - tip["ek_kisi"] - tip["calisan"] - tip["kimliksiz"]
    if mevcut > 0:
        parcalar.append(f"{mevcut} satir defterde zaten kayitli")
    if not tip["ek_kisi"] and not tip["tckn_kopru"]:
        return "defteri beslemedi: " + ("; ".join(parcalar) or "yeni bilgi yok")
    return "defteri besledi: " + "; ".join(parcalar)


def yedekle(hedef: Path, azami: int = 30) -> Path | None:
    """Ustune yazmadan once mevcut dosyayi veri/gecmis/ altina kopyalar.

    Harita ve ogrenen defterler surumsuz ustune yaziliyordu; hangi harita
    surumuyle hangi Excel'in uretildigi izlenemiyordu. Her yazimdan once
    '<ad>_<YYYYAAGG_SSDDss>.csv' kopyasi alinir; en eski kopyalar silinir.
    """
    import datetime as _dt
    import shutil

    try:
        if not hedef.is_file() or hedef.stat().st_size == 0:
            return None
        gecmis = hedef.parent / "gecmis"
        gecmis.mkdir(parents=True, exist_ok=True)
        damga = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        kopya = gecmis / f"{hedef.stem}_{damga}{hedef.suffix}"
        shutil.copy2(hedef, kopya)
        eskiler = sorted(gecmis.glob(f"{hedef.stem}_*{hedef.suffix}"))
        for e in eskiler[:-azami]:
            try:
                e.unlink()
            except OSError:
                pass
        return kopya
    except OSError:
        return None


def tckn_normalize(deger: Any) -> str:
    """TCKN / kimlik numarasini kanonik bicime cevirir.

    Sadece rakamlari alir. Turkiye TCKN 11 hanelidir; farkli uzunluktaki
    degerler (pasaport vb) yine de anahtar olarak kullanilabilir, bu yuzden
    uzunluk zorlanmaz. Cozulemeyen deger icin bos metin doner.

    >>> tckn_normalize(" 123 456 789 01 ")
    '12345678901'
    """
    if deger is None:
        return ""
    metin = str(deger).strip()
    if not metin or metin.lower() in {"nan", "nat", "none"}:
        return ""
    if metin.endswith(".0") and metin[:-2].isdigit():
        metin = metin[:-2]
    rakamlar = "".join(ch for ch in metin if ch.isdigit())
    return rakamlar


def _metin(deger: Any) -> str:
    """Hucre/alan degerini duz metne cevirir; bos degerler icin bos metin."""
    if deger is None:
        return ""
    metin = str(deger).strip()
    if metin.lower() in {"nan", "nat", "none"}:
        return ""
    return metin


def _sicil_metni(deger: Any) -> str:
    """Sicili kanonik metne cevirir ('.0' ekini temizler).

    ``masraf.kayit.sicil_normalize`` ile ayni davranisi gosterir; bu modulun
    kayit.py'ye (ve dolayisiyla pandas'a) bagimli olmamasi icin burada
    tekrarlanmistir.
    """
    metin = _metin(deger)
    if metin.endswith(".0") and metin[:-2].isdigit():
        metin = metin[:-2]
    return metin


class Defterler:
    """Kullanici duzeltmelerinden ogrenen yardimci defterler kumesi.

    Kullanim::

        defterler = Defterler(Path("veri"))
        defterler.alias_ekle("KARADUMAN HALIT CAN", "123456")
        defterler.kaydet()

    Butun sozlukler bellekte tutulur; ``kaydet()`` cagrilana kadar diske
    yazilmaz. ``kaydet()`` yalnizca degisen dosyalari yeniden yazar.
    """

    def __init__(self, kok: str | Path, olustur: bool = True) -> None:
        self.kok = Path(kok)
        self.aliases: dict[str, str] = {}
        self.harici: dict[str, dict] = {}
        self.ek_kisiler: dict[str, dict] = {}
        self.tckn_sicil: dict[str, str] = {}

        # Ham satirlar korunur ki tekrar yazarken ad_soyad/kaynak/tarih gibi
        # kolonlar kaybolmasin.
        self._alias_satirlari: dict[str, dict[str, str]] = {}
        self._harici_satirlari: dict[str, dict[str, str]] = {}
        self._ek_satirlari: dict[str, dict[str, str]] = {}
        self._tckn_satirlari: dict[str, dict[str, str]] = {}
        self._kirli: set[str] = set()
        #: Yukleme sirasinda saptanan sorunlar (bozuk dosya, yanlis ayirici,
        #: bilinmeyen kodlama, ayristirilamayan satirlar). Bos liste = sorun yok.
        #: Boru hatti bunlari calistirma uyarilarina tasir; eski surum bozuk bir
        #: defteri sessizce bos okuyor, ogretilen kayitlar kayboluyordu.
        self.uyarilar: list[str] = []
        #: Dosya bazinda okunan (CSV satiri) ve yuklenen (gecerli kayit) sayilari.
        self.okunan: dict[str, int] = {}
        self.yuklenen: dict[str, int] = {}

        if olustur:
            self._dosyalari_hazirla()
        self.yeniden_yukle()

    # ------------------------------------------------------------------
    # Dosya duzeyi yardimcilar
    # ------------------------------------------------------------------

    def yol(self, dosya: str) -> Path:
        """Defter dosyasinin tam yolunu dondurur."""
        return self.kok / dosya

    def _dosyalari_hazirla(self) -> None:
        """Eksik CSV dosyalarini basliklariyla olusturur."""
        try:
            self.kok.mkdir(parents=True, exist_ok=True)
        except OSError:
            return
        for dosya, basliklar in BASLIKLAR.items():
            hedef = self.yol(dosya)
            if hedef.exists():
                continue
            try:
                with hedef.open("w", encoding=KODLAMA, newline="") as akis:
                    csv.writer(akis, delimiter=AYIRICI).writerow(basliklar)
            except OSError:
                pass  # Salt okunur dizin: bellekte calismaya devam et.

    @staticmethod
    def _csv_oku(hedef: Path, beklenen: tuple[str, ...]) -> tuple[list[dict[str, str]], str | None]:
        """CSV dosyasini sozluk listesi olarak okur; ``(satirlar, sorun)`` dondurur.

        Kodlama (UTF-16 BOM, utf-8-sig, cp1254) ve ayirici (';' ',' sekme '|')
        dosyadan cikarilir. Beklenen ilk baslik hicbir ayiriciyla bulunamazsa
        ya da icerik metin degilse ``sorun`` dolu doner ve liste bos kalir.
        """
        if not hedef.exists():
            return [], None
        try:
            ham = hedef.read_bytes()
        except OSError as hata:
            return [], f"okunamadi ({hata.__class__.__name__}: {hata})"
        if not ham.strip():
            return [], None
        metin: str | None = None
        if ham[:2] in (b"\xff\xfe", b"\xfe\xff"):
            try:
                metin = ham.decode("utf-16")
            except UnicodeDecodeError:
                metin = None
        if metin is None:
            for kodlama in (KODLAMA, "cp1254"):
                try:
                    metin = ham.decode(kodlama)
                    break
                except UnicodeDecodeError:
                    continue
        if metin is None or "\x00" in metin:
            return [], "icerik metin dosyasi degil ya da kodlamasi taninmiyor"
        satir_metinleri = metin.splitlines()
        ilk_satir = satir_metinleri[0] if satir_metinleri else ""
        ayirici: str | None = None
        for aday in (AYIRICI, ",", "\t", "|"):
            basliklar = [b.strip().lstrip("\ufeff").strip() for b in ilk_satir.split(aday)]
            if beklenen[0] in basliklar:
                ayirici = aday
                break
        if ayirici is None:
            return [], (
                f"beklenen basliklar ({', '.join(beklenen[:3])}...) bulunamadi; "
                f"ilk satir: '{ilk_satir[:60]}'"
            )
        try:
            okuyucu = csv.DictReader(satir_metinleri, delimiter=ayirici)
            satirlar: list[dict[str, str]] = []
            for ham_satir in okuyucu:
                satir = {
                    (anahtar or "").strip(): _metin(deger)
                    for anahtar, deger in ham_satir.items()
                    if anahtar
                }
                if any(satir.values()):
                    satirlar.append(satir)
        except csv.Error as hata:
            return [], f"CSV ayristirilamadi ({hata})"
        return satirlar, None

    @staticmethod
    def _satirlari_oku(hedef: Path) -> list[dict[str, str]]:
        """CSV dosyasini sozluk listesi olarak okur (sorunlar yutulur; geriye uyum)."""
        dosya = hedef.name if hedef.name in BASLIKLAR else DOSYA_ALIAS
        return Defterler._csv_oku(hedef, BASLIKLAR[dosya])[0]

    def _dosya_oku(self, dosya: str) -> list[dict[str, str]]:
        """Bir defter dosyasini okur; sorun varsa ``uyarilar``'a yazar."""
        satirlar, sorun = self._csv_oku(self.yol(dosya), BASLIKLAR[dosya])
        self.okunan[dosya] = len(satirlar)
        if sorun:
            self.uyarilar.append(
                f"{dosya} okunamadi/bozuk: {sorun} (0 kayit yuklendi). Dosyayi Excel'de "
                "acip ';' ayiricili CSV (UTF-8) olarak yeniden kaydedin; ogretilen "
                "kayitlar bu calistirmada kullanilamadi."
            )
        return satirlar

    def _yukleme_kontrolu(self, dosya: str, yuklenen: int) -> None:
        """Okunan satirlarin cogu yuklenemediyse (zorunlu alanlar bos) uyarir."""
        self.yuklenen[dosya] = yuklenen
        okunan = self.okunan.get(dosya, 0)
        if okunan and yuklenen * 2 < okunan:
            self.uyarilar.append(
                f"{dosya} bozuk olabilir: {okunan} satirdan yalnizca {yuklenen} kayit "
                "yuklendi (zorunlu alanlar bos ya da ayirici yanlis). Dosyayi kontrol edin."
            )

    def _yaz(self, dosya: str, satirlar: Iterable[dict[str, str]]) -> bool:
        """Bir defter dosyasini basliklariyla birlikte yeniden yazar."""
        basliklar = BASLIKLAR[dosya]
        hedef = self.yol(dosya)
        try:
            hedef.parent.mkdir(parents=True, exist_ok=True)
            yedekle(hedef)
            gecici = hedef.with_suffix(".csv.tmp")
            with gecici.open("w", encoding=KODLAMA, newline="") as akis:
                yazici = csv.DictWriter(
                    akis, fieldnames=list(basliklar), delimiter=AYIRICI,
                    extrasaction="ignore",
                )
                yazici.writeheader()
                for satir in satirlar:
                    yazici.writerow({ad: satir.get(ad, "") for ad in basliklar})
            gecici.replace(hedef)
            return True
        except OSError:
            return False

    # ------------------------------------------------------------------
    # Yukleme
    # ------------------------------------------------------------------

    def yeniden_yukle(self) -> None:
        """Tum defterleri diskten yeniden okur (bellekteki degisiklikler gider)."""
        self.aliases.clear()
        self.harici.clear()
        self.ek_kisiler.clear()
        self.tckn_sicil.clear()
        self._alias_satirlari.clear()
        self._harici_satirlari.clear()
        self._ek_satirlari.clear()
        self._tckn_satirlari.clear()
        self._kirli.clear()
        self.uyarilar.clear()
        self.okunan.clear()
        self.yuklenen.clear()

        for satir in self._dosya_oku(DOSYA_ALIAS):
            isim = isim_normalize(satir.get("isim_norm", ""))
            sicil = _sicil_metni(satir.get("sicil"))
            if not isim or not sicil:
                continue
            satir["isim_norm"] = isim
            satir["sicil"] = sicil
            self.aliases[isim] = sicil
            self._alias_satirlari[isim] = satir
        self._yukleme_kontrolu(DOSYA_ALIAS, len(self._alias_satirlari))

        for satir in self._dosya_oku(DOSYA_HARICI):
            isim = isim_normalize(satir.get("isim_norm", "")) or isim_normalize(
                satir.get("ad_soyad", "")
            )
            if not isim:
                continue
            satir["isim_norm"] = isim
            self.harici[isim] = {
                "ad_soyad": satir.get("ad_soyad", "") or isim,
                "kurum": satir.get("kurum", ""),
                "masraf_merkezi": satir.get("masraf_merkezi", ""),
                "aciklama": satir.get("aciklama", ""),
                "kaynak": satir.get("kaynak", ""),
            }
            self._harici_satirlari[isim] = satir
        self._yukleme_kontrolu(DOSYA_HARICI, len(self._harici_satirlari))

        for satir in self._dosya_oku(DOSYA_EK_KISI):
            tckn = tckn_normalize(satir.get("tckn"))
            ad_soyad = satir.get("ad_soyad", "")
            isim = isim_normalize(satir.get("anahtar", "")) or isim_normalize(ad_soyad)
            anahtar = tckn or isim
            if not anahtar:
                continue
            satir["anahtar"] = anahtar
            satir["tckn"] = tckn
            kayit = {
                "ad_soyad": ad_soyad or isim,
                "santiye": satir.get("santiye", ""),
                "tckn": tckn,
                "kaynak": satir.get("kaynak", ""),
            }
            self._ek_satirlari[anahtar] = satir
            # Hem TCKN hem normalize isim ile aranabilir olsun.
            self.ek_kisiler[anahtar] = kayit
            if isim:
                self.ek_kisiler.setdefault(isim, kayit)
        self._yukleme_kontrolu(DOSYA_EK_KISI, len(self._ek_satirlari))

        for satir in self._dosya_oku(DOSYA_TCKN):
            tckn = tckn_normalize(satir.get("tckn"))
            sicil = _sicil_metni(satir.get("sicil"))
            if not tckn or not sicil:
                continue
            satir["tckn"] = tckn
            satir["sicil"] = sicil
            self.tckn_sicil[tckn] = sicil
            self._tckn_satirlari[tckn] = satir
        self._yukleme_kontrolu(DOSYA_TCKN, len(self._tckn_satirlari))

    # ------------------------------------------------------------------
    # Ekleme (ogrenme)
    # ------------------------------------------------------------------

    def alias_ekle(
        self,
        isim_norm: str,
        sicil: str,
        kaynak: str = "inceleme",
        ad_soyad: str = "",
    ) -> bool:
        """Bir ismi kalici olarak bir sicile baglar (kullanici duzeltmesi).

        Ayni isim daha once baska bir sicile baglanmissa kayit GUNCELLENIR;
        kullanicinin son karari gecerlidir.
        """
        isim = isim_normalize(isim_norm)
        kanonik = _sicil_metni(sicil)
        if not isim or not kanonik:
            return False
        if self.aliases.get(isim) == kanonik:
            return False
        self.aliases[isim] = kanonik
        self._alias_satirlari[isim] = {
            "isim_norm": isim,
            "sicil": kanonik,
            "ad_soyad": _metin(ad_soyad),
            "kaynak": _metin(kaynak) or "inceleme",
            "eklenme_tarihi": date.today().isoformat(),
        }
        self._kirli.add(DOSYA_ALIAS)
        return True

    def harici_ekle(
        self,
        isim_norm: str,
        ad_soyad: str,
        kurum: str,
        masraf_merkezi: str,
        aciklama: str = "",
        kaynak: str = "inceleme",
    ) -> bool:
        """Calisan olmayan bir kisiyi (dis danisman, grup sirketi) kaydeder."""
        isim = isim_normalize(isim_norm) or isim_normalize(ad_soyad)
        if not isim:
            return False
        kayit = {
            "ad_soyad": _metin(ad_soyad) or isim,
            "kurum": _metin(kurum),
            "masraf_merkezi": _metin(masraf_merkezi),
            "aciklama": _metin(aciklama),
            "kaynak": _metin(kaynak) or "inceleme",
        }
        if self.harici.get(isim) == kayit:
            return False
        self.harici[isim] = kayit
        self._harici_satirlari[isim] = {
            "isim_norm": isim,
            **kayit,
            "eklenme_tarihi": date.today().isoformat(),
        }
        self._kirli.add(DOSYA_HARICI)
        return True

    def ek_kisi_ekle(
        self,
        ad_soyad: str,
        tckn: str = "",
        santiye: str = "",
        kaynak: str = "yardimci",
    ) -> bool:
        """Personel ana verisinde olmayan bir kisiyi ek deftere yazar.

        Kaynak tipik olarak saglik kontrol listesi, arabuluculuk listesi veya
        egitim katilimci listesidir. Anahtar olarak once TCKN, yoksa normalize
        isim kullanilir.
        """
        temiz_tckn = tckn_normalize(tckn)
        isim = isim_normalize(ad_soyad)
        anahtar = temiz_tckn or isim
        if not anahtar:
            return False
        kayit = {
            "ad_soyad": _metin(ad_soyad) or isim,
            "santiye": _metin(santiye),
            "tckn": temiz_tckn,
            "kaynak": _metin(kaynak) or "yardimci",
        }
        mevcut = self.ek_kisiler.get(anahtar)
        if mevcut == kayit:
            return False
        # Mevcut kaydin dolu alanlarini bos yeni degerlerle EZME.
        if mevcut:
            for alan in ("santiye", "tckn"):
                if not kayit[alan] and mevcut.get(alan):
                    kayit[alan] = mevcut[alan]
            if mevcut == kayit:
                return False
        self.ek_kisiler[anahtar] = kayit
        if isim:
            self.ek_kisiler[isim] = kayit
        self._ek_satirlari[anahtar] = {
            "anahtar": anahtar,
            "tckn": kayit["tckn"],
            "ad_soyad": kayit["ad_soyad"],
            "santiye": kayit["santiye"],
            "kaynak": kayit["kaynak"],
            "eklenme_tarihi": date.today().isoformat(),
        }
        self._kirli.add(DOSYA_EK_KISI)
        return True

    def tckn_kopru_ekle(
        self,
        tckn: str,
        sicil: str,
        ad_soyad: str = "",
        kaynak: str = "yardimci",
    ) -> bool:
        """TCKN -> sicil koprusune kayit ekler.

        Personel ana verisinde TCKN YOKTUR; bu kopru arabuluculuk listesi gibi
        hem TCKN hem kimlik tasiyan yardimci kaynaklardan ya da kullanicinin
        inceleme ekranindaki kararindan beslenir.
        """
        temiz_tckn = tckn_normalize(tckn)
        kanonik = _sicil_metni(sicil)
        if not temiz_tckn or not kanonik:
            return False
        if self.tckn_sicil.get(temiz_tckn) == kanonik:
            return False
        self.tckn_sicil[temiz_tckn] = kanonik
        self._tckn_satirlari[temiz_tckn] = {
            "tckn": temiz_tckn,
            "sicil": kanonik,
            "ad_soyad": _metin(ad_soyad),
            "kaynak": _metin(kaynak) or "yardimci",
            "eklenme_tarihi": date.today().isoformat(),
        }
        self._kirli.add(DOSYA_TCKN)
        return True

    # ------------------------------------------------------------------
    # Yardimci kaynaklardan besleme
    # ------------------------------------------------------------------

    def yardimci_kaynaktan_besle(self, satirlar: list[GiderSatiri], defter: Any = None) -> dict[str, Any]:
        """Yardimci kaynak satirlarindan ek kisi defterini ve TCKN koprusunu doldurur.

        Saglik kontrol listesi, arabuluculuk listesi ve egitim katilimci
        listeleri gibi dosyalar TCKN, santiye ve bazen dogrudan sicil icerir.
        Bu bilgiler ek deftere yazilir, boylece personel ana verisinde YER
        ALMAYAN kisiler (taseron, yeni giren, aday) sonraki calistirmalarda
        eslesebilir.

        Seyahat faturasi (``antik_cari``) gibi kisi adinin serbest metinden
        tahmin edildigi kaynaklar defteri BESLEMEZ; oradaki isimler dogrulanmis
        kimlik degildir. Sicil tasiyan personel kutukleri (``referans_liste``,
        ferdi kaza sigorta listesi gibi) de beslemez: kisileri ana veri ve 1C
        listesiyle zaten eslesir, ek kisi defteri sicil tutmaz ve on binlerce
        satirlik kutuk defteri sisirir. Nedenler ``BESLEMEYEN_NEDENLER``'de.

        Returns:
            {'ek_kisi': eklenen ek kisi, 'tckn_kopru': eklenen kopru,
             'atlanan': islenmeyen satir (kaynak disi + kimliksiz),
             'islenen': besleme kaynagi olup ad/TCKN tasiyan satir,
             'calisan': ana veride zaten olan (yazilmayan) satir,
             'kimliksiz': ad ve TCKN'si olmayan satir,
             'kaynak_disi': {kaynak_tip: satir}  # besleme listesinde olmayanlar
             'tip_ozeti': {kaynak_tip: {'satir', 'ek_kisi', 'tckn_kopru',
                                        'calisan', 'kimliksiz', 'atlanan', 'neden'}}}
        """
        ozet: dict[str, Any] = {
            "ek_kisi": 0, "tckn_kopru": 0, "atlanan": 0, "islenen": 0,
            "calisan": 0, "kimliksiz": 0, "kaynak_disi": {}, "tip_ozeti": {},
        }

        def tip_kaydi(kaynak_tip: str) -> dict[str, Any]:
            return ozet["tip_ozeti"].setdefault(kaynak_tip, {
                "satir": 0, "ek_kisi": 0, "tckn_kopru": 0, "calisan": 0,
                "kimliksiz": 0, "atlanan": 0, "neden": besleme_nedeni(kaynak_tip),
            })

        for satir in satirlar:
            tip = tip_kaydi(satir.kaynak_tip)
            tip["satir"] += 1
            if satir.kaynak_tip not in BESLEYEN_KAYNAKLAR:
                ozet["atlanan"] += 1
                tip["atlanan"] += 1
                ozet["kaynak_disi"][satir.kaynak_tip] = ozet["kaynak_disi"].get(satir.kaynak_tip, 0) + 1
                continue
            ad = _metin(satir.kisi_ham)
            tckn = tckn_normalize(satir.tckn_ham)
            sicil = _sicil_metni(satir.sicil_ham)
            if not ad and not tckn:
                ozet["atlanan"] += 1
                ozet["kimliksiz"] += 1
                tip["atlanan"] += 1
                tip["kimliksiz"] += 1
                continue
            ozet["islenen"] += 1
            santiye = _metin(satir.masraf_merkezi_kaynak)
            # Ana veride zaten olan calisan ek kisi defterine yazilmaz; defter
            # sicili olmayanlar icindir (olculdu: 106 satirin 81'i calisandi).
            calisan_mi = False
            if defter is not None and ad:
                try:
                    from masraf.metin import isim_normalize
                    norm = isim_normalize(ad)
                    calisan_mi = bool(norm) and bool(
                        defter.isimle_adaylar(norm)
                        or defter.token_ile_adaylar(frozenset(norm.split(" ")))
                    )
                except Exception:  # noqa: BLE001
                    calisan_mi = False
            if calisan_mi:
                ozet["calisan"] += 1
                tip["calisan"] += 1
            elif self.ek_kisi_ekle(ad, tckn=tckn, santiye=santiye, kaynak=satir.kaynak_tip):
                ozet["ek_kisi"] += 1
                tip["ek_kisi"] += 1
            if tckn and sicil:
                if self.tckn_kopru_ekle(tckn, sicil, ad_soyad=ad, kaynak=satir.kaynak_tip):
                    ozet["tckn_kopru"] += 1
                    tip["tckn_kopru"] += 1
        return ozet

    # ------------------------------------------------------------------
    # Kaydetme ve ozet
    # ------------------------------------------------------------------

    def kaydet(self, tumu: bool = False) -> list[str]:
        """Degisen defterleri diske yazar; yazilan dosya adlarini dondurur.

        ``tumu=True`` verilirse degismemis dosyalar da yeniden yazilir
        (bicim tazeleme / onarim icin).
        """
        yazilanlar: list[str] = []
        eslesme = {
            DOSYA_ALIAS: self._alias_satirlari,
            DOSYA_HARICI: self._harici_satirlari,
            DOSYA_EK_KISI: self._ek_satirlari,
            DOSYA_TCKN: self._tckn_satirlari,
        }
        for dosya, satirlar in eslesme.items():
            if not tumu and dosya not in self._kirli:
                continue
            sirali = sorted(satirlar.values(), key=lambda s: str(s.get(BASLIKLAR[dosya][0], "")))
            if self._yaz(dosya, sirali):
                yazilanlar.append(dosya)
                self._kirli.discard(dosya)
        return yazilanlar

    def istatistik(self) -> dict:
        """Defterlerin ozetini dondurur (arayuzde gosterilir)."""
        return {
            "alias": len(self.aliases),
            "harici": len(self.harici),
            "ek_kisi": len(self._ek_satirlari),
            "tckn_kopru": len(self.tckn_sicil),
            "kok": str(self.kok),
            "kaydedilmemis": sorted(self._kirli),
            # Dosya bazinda: CSV'den okunan satir ve gecerli yuklenen kayit sayisi.
            "okunan": dict(self.okunan),
            "yuklenen": dict(self.yuklenen),
            "uyarilar": list(self.uyarilar),
        }
