"""Masraf Merkezi Otomasyonu - klasor tabanli calistirici.

Kullanim (Windows'ta CALISTIR.bat bunu cagirir):

* Faturalari ``1_FATURALAR`` klasorune atip CALISTIR.bat'a cift tiklayin, ya da
* Dosyalari dogrudan CALISTIR.bat uzerine surukleyip birakin.

Personel dosyalari ``PERSONEL`` klasorunden otomatik bulunur. Cikti
``2_EXCEL_CIKTI`` klasorune zaman damgali olarak yazilir ve otomatik acilir.

Bu dosya arayuzsuz calisir; hicbir sey kurmaya gerek yoktur ve internet
gerekmez. Butun veri bilgisayarda kalir.
"""

from __future__ import annotations

import os
import sys
import traceback
from datetime import datetime
from pathlib import Path

#: Kullanicinin dokunacagi klasorler. Yoksa olusturulur.
FATURA_DIZINI = "1_FATURALAR"
CIKTI_DIZINI = "2_EXCEL_CIKTI"
ARSIV_DIZINI = "3_ISLENENLER"
PERSONEL_DIZINI = "PERSONEL"
AYAR_DIZINI = "veri"

#: 1_FATURALAR icindeki isaret dosyasi; arsive tasinmaz.
ISARET_DOSYALARI = frozenset({"BURAYA_FATURA_ATIN.txt"})

#: Okumaya calisacagimiz uzantilar. Digerleri sessizce atlanir.
UZANTILAR = frozenset({".msg", ".xlsx", ".xls", ".xlsm", ".csv"})

#: Bu adlari tasiyan dosyalar fatura degil, yardimci personel listesidir.
YARDIMCI_IPUCLARI = ("1c", "personnel_list", "personnel list", "personel_list")

CIZGI = "=" * 62


def yaz(*parcalar) -> None:
    """Konsola yazar ve hemen bosaltir; kullanici ilerlemeyi canli gorsun."""
    print(*parcalar)
    sys.stdout.flush()


def kok_dizin() -> Path:
    """Otomasyon klasorunun kokunu bulur (program/kod/calistir.py -> ../..).

    Dosya beklenen yerlesimde degilse (ornegin depodan `paketle/calistir.py`
    olarak calistirildiysa) iki ust dizin ev klasoru olabilir ve kisisel veri
    iceren klasorler oraya acilirdi. O durumda calisma dizininin altinda
    ayri bir klasor kullanilir ve bu acikca soylenir.
    """
    burasi = Path(__file__).resolve()
    if burasi.parent.name == "kod" and burasi.parent.parent.name == "program":
        return burasi.parent.parent.parent
    return Path.cwd() / "OTOMASYON_CALISMA"


def dizinleri_hazirla(kok: Path) -> None:
    for ad in (FATURA_DIZINI, CIKTI_DIZINI, ARSIV_DIZINI, PERSONEL_DIZINI, AYAR_DIZINI):
        (kok / ad).mkdir(parents=True, exist_ok=True)


def _tablo_dosyasi_mi(yol: Path) -> bool:
    if yol.suffix.lower() not in UZANTILAR:
        return False
    # Excel gecici dosyalari (~$ ile baslar) ve onbellekler
    if yol.name.startswith("~$") or yol.name.startswith("."):
        return False
    if yol.suffix.lower() == ".pkl":
        return False
    return True


def personel_dosyalarini_bul(kok: Path) -> tuple[Path | None, Path | None, list[str]]:
    """PERSONEL klasorunden ana veri ve 1C listesini ayirir.

    Ayirma kurali: adinda '1C' ya da 'personnel list' gecen dosya yardimci
    listedir; kalanlarin EN BUYUGU ana veridir (ana veri aylik snapshot
    tasidigi icin her zaman daha buyuktur).

    Returns:
        (ana veri yolu, yardimci liste yolu, kullaniciya gosterilecek notlar)
    """
    dizin = kok / PERSONEL_DIZINI
    adaylar = [
        y for y in sorted(dizin.glob("*"))
        if y.is_file() and y.suffix.lower() in (".xlsx", ".xls", ".xlsm")
        and not y.name.startswith("~$")
    ]
    notlar: list[str] = []
    if not adaylar:
        return None, None, notlar

    yardimcilar, anadaylar = [], []
    for y in adaylar:
        ad = y.name.lower()
        (yardimcilar if any(i in ad for i in YARDIMCI_IPUCLARI) else anadaylar).append(y)

    ana = max(anadaylar, key=lambda y: y.stat().st_size) if anadaylar else None
    # 1C listesi aylik gelir; iki ay yan yana durursa EN YENISI (degistirilme
    # tarihi) gecerlidir, boyutu degil. Atlanan soylenir.
    yardimci = max(yardimcilar, key=lambda y: y.stat().st_mtime) if yardimcilar else None

    if ana is None and yardimci is not None:
        notlar.append(
            "PERSONEL klasorunde yalnizca 1C listesi var, ana personel verisi yok. "
            "Ana veri ZORUNLUDUR: kisi eslestirme ve donem (gider ayi) kontrolu onunla yapilir."
        )
    for atlanan in anadaylar:
        if atlanan is not ana:
            notlar.append(f"Atlandi (ana veri olarak en buyugu secildi): {atlanan.name}")
    for atlanan in yardimcilar:
        if atlanan is not yardimci:
            notlar.append(f"Atlandi (1C listesi olarak en yenisi secildi): {atlanan.name}")
    return ana, yardimci, notlar


#: Personel ana verisi yanlislikla 1_FATURALAR'a atilirsa 150 bin satir okunup
#: 'kutuk' sayilir ve arsive tasinir; kullanici dosyasini kaybolmus gorur.
_PERSONEL_IPUCLARI = ("giris_cikis", "giris cikis", "personnel", "personel", "1c_")


def _personel_dosyasi_gibi(yol: Path) -> bool:
    ad = yol.name.lower()
    return any(ip in ad for ip in _PERSONEL_IPUCLARI)


def fatura_dosyalarini_topla(kok: Path, argumanlar: list[str]) -> tuple[list[Path], list[str]]:
    """Surukle-birak ile gelenler + 1_FATURALAR klasorundekiler.

    Returns:
        (dosyalar, notlar). Notlar: atlanan dosyalar ve sebebi.
    """
    bulunan: list[Path] = []
    notlar: list[str] = []
    gorulen: set[str] = set()

    def ekle(yol: Path) -> None:
        try:
            anahtar = str(yol.resolve()).lower()
        except OSError:
            anahtar = str(yol).lower()
        if anahtar in gorulen or not _tablo_dosyasi_mi(yol):
            return
        gorulen.add(anahtar)
        if _personel_dosyasi_gibi(yol):
            notlar.append(
                f"{yol.name}: adi personel verisine benziyor, fatura olarak islenmedi "
                f"ve yerinden oynatilmadi. Personel dosyalari {PERSONEL_DIZINI} klasorune konur."
            )
            return
        bulunan.append(yol)

    for ham in argumanlar:
        yol = Path(ham)
        if yol.is_dir():
            for alt in sorted(yol.rglob("*")):
                if alt.is_file():
                    ekle(alt)
        elif yol.is_file():
            ekle(yol)

    for alt in sorted((kok / FATURA_DIZINI).rglob("*")):
        if alt.is_file():
            ekle(alt)
    return bulunan, notlar


# --------------------------------------------------------------------------
# Calistirmalar arasi tekrar korumasi
# --------------------------------------------------------------------------

ISLENEN_OZETLER_DOSYASI = "ISLENEN_DOSYALAR.txt"


def dosya_ozeti(yol: Path) -> str | None:
    """Dosya iceriginin SHA-256'si (ad degisse de ayni kalir)."""
    import hashlib

    try:
        h = hashlib.sha256()
        with open(yol, "rb") as f:
            for parca in iter(lambda: f.read(1 << 20), b""):
                h.update(parca)
        return h.hexdigest()
    except OSError:
        return None


def islenen_ozetleri_oku(kok: Path) -> dict[str, tuple[str, str]]:
    """3_ISLENENLER/ISLENEN_DOSYALAR.txt -> {ozet: (damga, dosya adi)}."""
    yol = kok / ARSIV_DIZINI / ISLENEN_OZETLER_DOSYASI
    kayit: dict[str, tuple[str, str]] = {}
    try:
        for satir in yol.read_text(encoding="utf-8").splitlines():
            parcalar = satir.split("\t")
            if len(parcalar) >= 3 and len(parcalar[0]) == 64:
                kayit.setdefault(parcalar[0], (parcalar[1], parcalar[2]))
    except OSError:
        pass
    return kayit


def islenen_ozetleri_yaz(kok: Path, damga: str, dosyalar: list[Path],
                         ozetler: dict[str, str | None] | None = None,
                         envanter: list | None = None) -> None:
    """Bu calistirmada islenen dosyalarin ozetlerini kalici listeye ekler.

    Ust duzey dosyalarin yaninda mail/arsiv iclerinden okunan eklerin ozetleri
    de yazilir: ayni fatura bir ay mail eki, sonraki ay tedarikciden gelen
    xlsx olarak gelirse 'DAHA ONCE ISLENDI' uyarisi yine cikar.
    """
    yol = kok / ARSIV_DIZINI / ISLENEN_OZETLER_DOSYASI
    try:
        yol.parent.mkdir(parents=True, exist_ok=True)
        yeni = not yol.exists()
        with open(yol, "a", encoding="utf-8") as f:
            if yeni:
                f.write("# sha256\tcalistirma\tdosya   (ayni icerik tekrar gelirse uyarilir)\n")
            yazilan: set[str] = set()
            for d in dosyalar:
                oz = (ozetler or {}).get(str(d)) or (ozetler or {}).get(d.name) or dosya_ozeti(d)
                if oz and oz not in yazilan:
                    yazilan.add(oz)
                    f.write(f"{oz}\t{damga}\t{d.name}\n")
            for k in envanter or []:
                oz = str(k.get("ozet") or "")
                if len(oz) != 64 or oz in yazilan or not k.get("kaynak"):
                    continue
                if str(k.get("durum")) not in ("OKUNDU", "KUTUK", "DETAY LISTESI", "AYNI ICERIK"):
                    continue
                yazilan.add(oz)
                f.write(f"{oz}\t{damga}\t{k.get('kaynak')} > {k.get('ad')}\n")
    except OSError:
        pass


def daha_once_islenenler(kok: Path, faturalar: list[Path]) -> list[str]:
    """Icerigi daha onceki bir calistirmada islenmis dosyalari uyari metni olarak doner."""
    kayit = islenen_ozetleri_oku(kok)
    if not kayit:
        return []
    uyarilar: list[str] = []
    for f in faturalar:
        oz = dosya_ozeti(f)
        if oz and oz in kayit:
            damga, eski_ad = kayit[oz]
            uyarilar.append(
                f"DAHA ONCE ISLENDI: {f.name} icerigi {damga} calistirmasinda "
                f"('{eski_ad}') zaten islenmisti. Ayni ay iki kez muhasebeye gitmesin; "
                f"bu Excel'i kullanmadan once {ARSIV_DIZINI}/{damga} ile karsilastirin."
            )
    return uyarilar


def _sayi(deger: float) -> str:
    """Turkce bicimde sayi: 48.946,59"""
    return f"{deger:,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")


def ozet_bas(sonuclar, mahsup, uyarilar: list[str]) -> None:
    """Konsola okunakli Turkce ozet basar."""
    from masraf.modeller import DURUM_ESLESMEDI, DURUM_INCELE, DURUM_OTOMATIK

    otomatik = sum(1 for s in sonuclar if s.durum == DURUM_OTOMATIK)
    incele = sum(1 for s in sonuclar if s.durum == DURUM_INCELE)
    eslesmedi = sum(1 for s in sonuclar if s.durum == DURUM_ESLESMEDI)

    yaz()
    yaz(CIZGI)
    yaz("  SONUC")
    yaz(CIZGI)
    yaz(f"  Okunan satir      : {len(sonuclar)}")
    yaz(f"  Otomatik dagitildi: {otomatik}")
    yaz(f"  Incelenecek       : {incele}")
    yaz(f"  Kisi bulunamadi   : {eslesmedi}")

    if mahsup is not None:
        yaz()
        yaz("  MUTABAKAT")
        for para, d in mahsup.toplamlar().items():
            yaz(f"    Okunan       : {_sayi(d['gelen']):>14s} {para}")
            if d["yinelenen"]:
                yaz(f"    Yinelenen    : {_sayi(d['yinelenen']):>14s} {para}"
                    "   (ayni islem baska dosyada da vardi)")
            yaz(f"    Dagitilan    : {_sayi(d['dagitilan']):>14s} {para}")
            yaz(f"    Dagitilamayan: {_sayi(d['dagitilamayan']):>14s} {para}")
        yaz()
        if not mahsup.kapali_mi:
            yaz("  [DIKKAT] MUTABAKAT ACIK. Bu tablo muhasebeye gonderilmemeli:")
            for k in mahsup.acik_kontroller:
                sebep = getattr(k, "acik_sebebi", "") or f"fark {k.fark:+.2f}"
                yaz(f"     {k.kaynak} ({k.para_birimi}): {sebep}")
        elif incele or eslesmedi:
            # Excel kapagiyla ayni dil: para kaybolmadi ama tablo TASLAKTIR.
            yaz("  [TASLAK] Para kaybolmadi; mutabakat kapandi. Ama "
                f"{incele} satir inceleme, {eslesmedi} satir kisi bulunamadi.")
            yaz("  Excel'deki 'Incele' ve 'Eslesmedi' sayfalari gorulmeden tablo onaya sunulmamali.")
        else:
            yaz("  [TAMAM] Butun faturalar kapandi, inceleme bekleyen satir yok. Tablo onaya hazir.")

        merkezler = mahsup.merkez_ozeti()
        if merkezler:
            yaz()
            yaz("  PROJE BAZINDA DAGILIM")
            for m in merkezler[:12]:
                if str(m["masraf_merkezi"]) == "(DAGITILAMAYAN)":
                    isaret = "  <- kisi / merkez bulunamadi"
                else:
                    isaret = "" if m["haritada_var"] else "  <- haritada tanimli degil"
                yaz(f"    {m['masraf_merkezi'][:32]:32s} "
                    f"{_sayi(m['tutar']):>13s} {m['para_birimi']}"
                    f"  %{m['pay_yuzde']:.1f}{isaret}")
            if len(merkezler) > 12:
                yaz(f"    ... ve {len(merkezler) - 12} tane daha (Excel'de tamami var)")

    if uyarilar:
        yaz()
        yaz("  UYARILAR")
        for u in uyarilar[:10]:
            yaz(f"    - {u[:150]}")
        if len(uyarilar) > 10:
            yaz(f"    ... ve {len(uyarilar) - 10} uyari daha")


def islenenleri_arsivle(kok: Path, faturalar: list[Path], damga: str) -> tuple[list[Path], list[str]]:
    """Basariyla islenen faturalari 1_FATURALAR'dan 3_ISLENENLER/<damga>/ altina tasir.

    Neden: 1_FATURALAR bosaltilmazsa gelecek ay ayni dosyalar TEKRAR islenir ve
    tutarlar iki kez sayilir. Bunu kullanicinin hatirlamasina birakmak muhasebe
    riskidir. Tasima geri alinabilir; dosya silinmez, sadece yer degistirir.

    Surukle-birak ile disaridan verilen dosyalara DOKUNULMAZ; onlar kullanicinin
    kendi klasorundedir ve yerinden oynatmak surpriz olur.

    Returns:
        (tasinanlar, hatalar)
    """
    import shutil

    kaynak_kok = (kok / FATURA_DIZINI).resolve()
    hedef = kok / ARSIV_DIZINI / damga
    tasinan: list[Path] = []
    hatalar: list[str] = []
    for f in faturalar:
        try:
            gercek = f.resolve()
        except OSError:
            continue
        if kaynak_kok not in gercek.parents:
            continue  # disaridan gelen dosya, dokunma
        if f.name in ISARET_DOSYALARI:
            continue
        try:
            hedef.mkdir(parents=True, exist_ok=True)
            # Alt klasor yapisini koru: 1_FATURALAR/temmuz/a.msg -> arsiv/temmuz/a.msg
            gorece = gercek.relative_to(kaynak_kok)
            varis = hedef / gorece
            varis.parent.mkdir(parents=True, exist_ok=True)
            if varis.exists():
                varis = varis.with_name(f"{varis.stem}_{damga}{varis.suffix}")
            shutil.move(str(gercek), str(varis))
            tasinan.append(varis)
        except Exception as hata:  # noqa: BLE001 - dosya kilitli olabilir (Outlook acik)
            hatalar.append(f"{f.name}: tasinamadi ({hata.__class__.__name__}: {hata})")
    # Bosalan alt klasorleri temizle
    for alt in sorted((p for p in kaynak_kok.rglob("*") if p.is_dir()), reverse=True):
        try:
            if not any(alt.iterdir()):
                alt.rmdir()
        except OSError:
            pass
    return tasinan, hatalar


def calistirma_kaydi_yaz(kok: Path, damga: str, faturalar: list[Path], ana: Path | None,
                         yardimci: Path | None, sonuclar, mahsup, excel_yolu: str,
                         tasinan: list[Path], okunamayanlar: list[str] | None = None,
                         uyarilar: list[str] | None = None,
                         ozetler: dict[str, str | None] | None = None,
                         envanter: list[dict] | None = None) -> Path | None:
    """Her calistirma icin kisa bir denetim kaydi birakir.

    Finans 'bu Excel hangi dosyalardan, hangi personel verisiyle uretildi'
    sorusunu aylar sonra da cevaplayabilmeli. Kayit hem arsiv klasorune hem
    de kok dizindeki CALISTIRMA_GECMISI.txt dosyasina eklenir.
    """
    from masraf.modeller import DURUM_ESLESMEDI, DURUM_INCELE, DURUM_OTOMATIK

    def _boyut(y: Path | None) -> str:
        try:
            return f"{y.stat().st_size:,} bayt, degistirilme {datetime.fromtimestamp(y.stat().st_mtime):%d.%m.%Y %H:%M}" if y else "-"
        except OSError:
            return "-"

    satirlar = [
        f"CALISTIRMA: {damga}",
        f"Tarih          : {datetime.now():%d.%m.%Y %H:%M:%S}",
        f"Excel          : {excel_yolu or '(uretilemedi)'}",
        f"Ana personel   : {ana.name if ana else '-'}  ({_boyut(ana)})",
        f"1C listesi     : {yardimci.name if yardimci else '-'}  ({_boyut(yardimci)})",
        f"Islenen dosya  : {len(faturalar) - len(okunamayanlar or [])}",
    ]
    for f in faturalar:
        if f.name in (okunamayanlar or []):
            continue
        oz = (ozetler or {}).get(f.name) or dosya_ozeti(f)
        satirlar.append(f"    - {f.name}  sha256={oz or '?'}")
    for ad in (okunamayanlar or []):
        satirlar.append(f"    - {ad}  OKUNAMADI (1_FATURALAR'da birakildi)")
    # Hangi harita ve defter surumuyle uretildi? Ay sonra 'neden boyle
    # dagitilmis' sorusu bunlarla cevaplanir.
    veri = kok / AYAR_DIZINI
    for ad in ("masraf_merkezi_haritasi.csv", "aliases.csv", "harici_kisiler.csv",
               "ek_kisiler.csv", "tckn_sicil.csv", "kolon_esanlamlilari.csv"):
        y = veri / ad
        if y.is_file():
            satirlar.append(f"Veri dosyasi   : {ad}  ({_boyut(y)}, sha256={dosya_ozeti(y) or '?'})")
    if envanter:
        satirlar.append("Dosya envanteri: (durum | satir | okunan tutar | dosya  [nereden]  sebep)")
        for k in envanter:
            tutar = k.get("tutar")
            tutar_m = f"{tutar:,.2f} {k.get('para_birimi') or ''}".strip() if isinstance(tutar, (int, float)) else "-"
            satirlar.append(
                f"    {str(k.get('durum')):13} {int(k.get('satir') or 0):5d}  {tutar_m:>16}  "
                f"{k.get('ad')}  [{k.get('kaynak') or 'dogrudan'}]  {k.get('sebep') or ''}".rstrip()
            )
    tarihler = [s.satir.belge_tarihi for s in sonuclar
                if getattr(getattr(s, "satir", None), "belge_tarihi", None)]
    if tarihler:
        satirlar.append(f"Gider donemi   : {min(tarihler):%d.%m.%Y} - {max(tarihler):%d.%m.%Y}")
    satirlar.append(f"Okunan satir   : {len(sonuclar)}")
    satirlar.append(f"  otomatik     : {sum(1 for s in sonuclar if s.durum == DURUM_OTOMATIK)}")
    satirlar.append(f"  incele       : {sum(1 for s in sonuclar if s.durum == DURUM_INCELE)}")
    satirlar.append(f"  eslesmedi    : {sum(1 for s in sonuclar if s.durum == DURUM_ESLESMEDI)}")
    if mahsup is not None:
        for para, d in mahsup.toplamlar().items():
            satirlar.append(
                f"Mutabakat {para}: okunan {d['gelen']:,.2f} | yinelenen {d['yinelenen']:,.2f} | "
                f"dagitilan {d['dagitilan']:,.2f} | dagitilamayan {d['dagitilamayan']:,.2f} | "
                f"{'KAPALI' if mahsup.kapali_mi else 'ACIK'}"
            )
    if tasinan:
        satirlar.append(f"Arsive tasinan : {len(tasinan)} dosya -> {ARSIV_DIZINI}/{damga}/")
    for u in (uyarilar or []):
        satirlar.append(f"Uyari          : {u}")
    metin = "\n".join(satirlar) + "\n"
    try:
        arsiv = kok / ARSIV_DIZINI / damga
        arsiv.mkdir(parents=True, exist_ok=True)
        (arsiv / "OZET.txt").write_text(metin, encoding="utf-8")
        gecmis = kok / "CALISTIRMA_GECMISI.txt"
        with open(gecmis, "a", encoding="utf-8") as f:
            f.write(metin + "-" * 60 + "\n")
        return gecmis
    except OSError:
        return None


def excel_ac(yol: Path) -> None:
    """Uretilen Excel'i varsayilan programda acar. Basarisiz olursa sessiz gecer."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(yol))  # type: ignore[attr-defined]
    except Exception:
        pass


def calistir() -> int:
    kok = kok_dizin()
    dizinleri_hazirla(kok)

    yaz(CIZGI)
    yaz("  MASRAF MERKEZI OTOMASYONU")
    yaz(CIZGI)
    yaz(f"  Klasor: {kok}")
    yaz()

    # --- Personel dosyalari ---
    ana, yardimci, notlar = personel_dosyalarini_bul(kok)
    if ana is None and yardimci is None:
        yaz("HATA: Personel dosyasi bulunamadi.")
        yaz()
        yaz("  Su klasore personel dosyalarini koyun:")
        yaz(f"     {kok / PERSONEL_DIZINI}")
        yaz()
        yaz("  Gereken dosyalar:")
        yaz("    1) Ana personel verisi (ornek: 2025_2026_giris_cikis.xlsx)  ZORUNLU")
        yaz("    2) 1C personel listesi (ornek: 1C_Personnel_List_...xlsx)   ONERILIR")
        yaz()
        yaz("  Bu dosyalari koyduktan sonra bu programi tekrar calistirin.")
        return 1

    if ana is None:
        yaz("HATA: Ana personel verisi bulunamadi; yalnizca 1C listesi var.")
        yaz()
        yaz("  Ana veri ZORUNLUDUR: kisi eslestirme ve gider ayi kontrolu onunla yapilir.")
        yaz(f"  Ornek: 2025_2026_giris_cikis.xlsx  ->  {kok / PERSONEL_DIZINI}")
        yaz("  Not: adinda 'personnel list' ya da '1C' gecen dosya 1C listesi sayilir;")
        yaz("  ana verinin adinda bu ifadeler gecmemeli.")
        return 1
    if ana is not None:
        yaz(f"  Ana personel verisi : {ana.name}")
    if yardimci is not None:
        yaz(f"  1C personel listesi : {yardimci.name}")
    else:
        yaz("  1C personel listesi : YOK")
        yaz("    (grup sirketi personeli - Renservis, Renstroydetal, RC, One Tower,")
        yaz("     Top Tower - bu liste olmadan 'eslesmedi' olarak kalir)")
    for n in notlar:
        yaz(f"  NOT: {n}")

    # --- Fatura dosyalari ---
    faturalar, topla_notlari = fatura_dosyalarini_topla(kok, sys.argv[1:])
    for n in topla_notlari:
        yaz(f"  NOT: {n}")
    if not faturalar:
        yaz()
        yaz("HATA: Islenecek fatura bulunamadi.")
        yaz()
        yaz("  Faturalari su klasore atin:")
        yaz(f"     {kok / FATURA_DIZINI}")
        yaz()
        yaz("  Ya da dosyalari dogrudan CALISTIR.bat uzerine surukleyip birakin.")
        yaz("  Kabul edilen tipler: .msg (Outlook), .xlsx, .xls, .xlsm, .csv")
        return 1

    yaz()
    yaz(f"  Islenecek dosya: {len(faturalar)}")
    for f in faturalar[:15]:
        yaz(f"    - {f.name}")
    if len(faturalar) > 15:
        yaz(f"    ... ve {len(faturalar) - 15} dosya daha")

    # --- Isle ---
    from masraf.boru import Boru, CalismaAyarlari

    yaz()
    yaz("  Personel verisi okunuyor. Ilk seferde 1-2 dakika surebilir (24 bin kayit),")
    yaz("  sonraki calistirmalar onbellek sayesinde 10 saniyenin altina iner.")
    yaz()

    ayarlar = CalismaAyarlari(
        personel_yolu=ana,
        yardimci_personel_yolu=yardimci,
        veri_dizini=str(kok / AYAR_DIZINI),
        cikti_dizini=str(kok / CIKTI_DIZINI),
    )
    boru = Boru(ayarlar)
    try:
        boru.hazirla()
    except Exception as hata:  # noqa: BLE001 - kullaniciya Turkce soylenir
        from masraf.boru import _hata_metni

        yaz()
        yaz("HATA: Personel dosyasi okunamadi.")
        yaz(f"  {_hata_metni(hata)}")
        yaz()
        yaz("  Olasi sebepler: dosya parola korumali (Excel'de acip parolasiz kaydedin),")
        yaz("  dosya bozuk ya da bos, zorunlu kolonlar yok (Sicil, Adi Soyadi, Gorev Yeri, Donem).")
        yaz(f"  Dosya: {ana}")
        return 1

    son_yuzde = [-10.0]

    def ilerleme(yuzde: float, mesaj: str) -> None:
        """Boru hatti yuzdeyi 0-100 olcuginde bildirir.

        Her adimi basmiyoruz; 405 satirlik bir dosyada onlarca satir akiyor
        ve kullanici konsolda ne oldugunu takip edemiyor. Yuzde 5'ten az
        ilerleyen adimlar atlanir, sonuncusu her zaman basilir.
        """
        # Her dosyanin 'Okunuyor' satiri gorunsun; kullanici hangi dosyada
        # beklendigini bilsin (ilk dosya %5 kuralina takilip kayboluyordu).
        if yuzde - son_yuzde[0] < 5 and yuzde < 100 and not str(mesaj).startswith("Okunuyor"):
            return
        son_yuzde[0] = yuzde
        yaz(f"    [%{yuzde:3.0f}] {mesaj}")

    damga = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Ayni dosya (icerik olarak) daha once islendiyse soyle: konsola, Excel
    # kapagina ve calistirma kaydina. Islemeyi durdurmaz; karar finansindir.
    tekrar_uyarilari = daha_once_islenenler(kok, faturalar)
    for u in tekrar_uyarilari:
        yaz(f"  UYARI: {u}")
    # Adi personel verisine benzeyen dosyalar islenmedi; bu yalnizca konsolda
    # kalmasin, Excel kapagina ve OZET.txt'e de gecsin.
    boru.on_uyarilar = list(tekrar_uyarilari) + [
        f"ISLENMEDI: {n}" for n in topla_notlari if "personel verisine benziyor" in n
    ]
    try:
        sonuc = boru.calistir(faturalar, cikti_adi=f"Masraf_Dagitimi_{damga}.xlsx",
                              ilerleme=ilerleme)
    except (PermissionError, OSError) as hata:
        yaz()
        yaz("HATA: Excel ciktisi yazilamadi.")
        yaz(f"  {hata.__class__.__name__}: {hata}")
        yaz(f"  Klasor: {kok / CIKTI_DIZINI}")
        yaz("  Klasor salt okunur olabilir, ag surucusu kopmus olabilir ya da ayni adli")
        yaz("  bir Excel baska programda acik olabilir. Faturalar yerinde birakildi.")
        return 1
    except Exception as hata:  # noqa: BLE001
        if hata.__class__.__name__ == "FileCreateError":
            yaz()
            yaz("HATA: Excel ciktisi yazilamadi (dosya kilitli ya da klasor yazilamiyor).")
            yaz(f"  {hata}")
            yaz(f"  Klasor: {kok / CIKTI_DIZINI}")
            yaz("  Faturalar yerinde birakildi; sorunu giderip tekrar calistirin.")
            return 1
        raise

    sonuclar = sonuc.get("sonuclar") or []
    if not sonuclar:
        yaz()
        yaz("HATA: Dosyalardan hicbir satir okunamadi.")
        yaz()
        yaz("  Olasi sebepler:")
        yaz("   - Dosya bicimi taninmiyor. veri/kolon_esanlamlilari.csv dosyasina")
        yaz("     o dosyadaki kisi/tutar/tarih kolon adlarini ekleyin.")
        yaz("   - Dosya bos ya da parola korumali.")
        for h in (boru.hatalar or [])[:5]:
            yaz(f"   - {h}")
        return 1

    ozet_bas(sonuclar, sonuc.get("mahsup"), list(boru.uyarilar or []))

    envanter = list((sonuc.get("ozet") or {}).get("dosya_envanteri") or [])
    if envanter:
        from collections import Counter
        sayim = Counter(str(k.get("durum")) for k in envanter)
        yaz()
        yaz("  DOSYA ENVANTERI  (tam liste Excel'in 'Dosyalar' sayfasinda)")
        yaz("    " + "   ".join(f"{d}: {n}" for d, n in sayim.most_common()))
        for k in envanter:
            if str(k.get("durum")) in ("OKUNAMADI", "SATIR YOK", "AYNI ICERIK"):
                yaz(f"    - {k.get('durum'):12} {k.get('ad')}  ({str(k.get('sebep'))[:90]})")
        atlanan = [k for k in envanter if str(k.get("durum")) == "ATLANDI"]
        if atlanan:
            yaz(f"    - ATLANDI      {len(atlanan)} ek tablo degil (PDF vb.), acilmadi; adlari Excel'de")

    if boru.hatalar:
        yaz()
        yaz("  HATALAR (bu dosyalar okunamadi)")
        for h in boru.hatalar[:10]:
            yaz(f"    - {h}")

    excel_yolu = sonuc.get("excel_yolu") or ""
    yaz()
    yaz(CIZGI)
    tasinan: list[Path] = []
    if excel_yolu and Path(excel_yolu).exists():
        yaz("  EXCEL HAZIR")
        yaz(f"    {excel_yolu}")
        yaz()
        yaz("  Excel'de once 'Ozet' ve 'Mahsuplasma' sayfalarina bakin: muhasebeye")
        yaz("  gidecek tablo odur. 'Kontrol' sayfasi paranin kaybolmadigini gosterir.")

        # Basarili calistirmadan sonra islenen faturalari arsive tasi. Gelecek ay
        # ayni dosyalarin tekrar islenip cift sayilmasini onler. Iki istisna:
        #  - okunamayan dosyalar yerinde kalir (duzeltilip tekrar denenecek),
        #  - mutabakat ACIKSA hicbiri tasinmaz; Excel 'gonderilmemeli' diyor,
        #    kullanici duzeltip tekrar calistiracak, dosyalari aramasin.
        mahsup = sonuc.get("mahsup")
        # 'Kisi defteri beslenemedi: ...' gibi dosya disi hata metinleri
        # okunamayan dosya sayilmasin; yalnizca fatura adiyla baslayanlar.
        fatura_adlari = {f.name for f in faturalar}
        okunamayanlar = {
            h.split(":", 1)[0].strip() for h in boru.hatalar
            if h.split(":", 1)[0].strip() in fatura_adlari
        }
        arsivlenecek = [f for f in faturalar if f.name not in okunamayanlar]
        # Ozetler tasimadan ONCE alinir; tasinan dosyanin eski yolu kalmaz.
        # Anahtar tam yol: farkli alt klasorlerdeki ayni adli dosyalar karismasin.
        ozetler = {str(f): dosya_ozeti(f) for f in faturalar}
        tasinan, tasima_hatalari = [], []
        if mahsup is not None and not mahsup.kapali_mi:
            yaz()
            yaz("  MUTABAKAT ACIK: faturalar 1_FATURALAR klasorunde BIRAKILDI.")
            yaz("  Excel'deki 'Kontrol' sayfasina bakip sorunu giderin ve tekrar")
            yaz("  calistirin; dosyalar yerinde oldugu icin geri almaniz gerekmez.")
        else:
            tasinan, tasima_hatalari = islenenleri_arsivle(kok, arsivlenecek, damga)
            if tasinan:
                yaz()
                yaz(f"  Islenen {len(tasinan)} dosya arsive tasindi:")
                yaz(f"    {kok / ARSIV_DIZINI / damga}")
                yaz("  1_FATURALAR klasoru gelecek ay icin bos. Dosyalar silinmedi,")
                yaz("  gerekirse arsivden geri alabilirsiniz.")
            islenen_ozetleri_yaz(kok, damga, arsivlenecek, ozetler, envanter=envanter)
        if okunamayanlar:
            yaz(f"  Okunamayan {len(okunamayanlar)} dosya 1_FATURALAR'da birakildi:")
            for ad in sorted(okunamayanlar):
                yaz(f"    - {ad}")
        for h in tasima_hatalari:
            yaz(f"  UYARI: {h}")
        kayit = calistirma_kaydi_yaz(kok, damga, faturalar, ana, yardimci,
                                     sonuclar, mahsup, excel_yolu, tasinan,
                                     okunamayanlar=sorted(okunamayanlar),
                                     uyarilar=list(boru.uyarilar), ozetler=ozetler,
                                     envanter=envanter)
        if kayit:
            yaz(f"  Calistirma kaydi: {kayit.name}")
        excel_ac(Path(excel_yolu))
    else:
        yaz("  UYARI: Excel dosyasi olusturulamadi. Faturalar yerinde birakildi.")
    yaz(CIZGI)
    return 0


def main() -> int:
    try:
        return calistir()
    except KeyboardInterrupt:
        yaz()
        yaz("Islem kullanici tarafindan durduruldu.")
        return 130
    except Exception:
        yaz()
        yaz(CIZGI)
        yaz("  BEKLENMEYEN HATA")
        yaz(CIZGI)
        yaz(traceback.format_exc())
        yaz()
        yaz("  Bu metnin tamamini kopyalayip iletirseniz sorun cozulebilir.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
