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
    """Otomasyon klasorunun kokunu bulur (program/kod/calistir.py -> ../..)."""
    return Path(__file__).resolve().parent.parent.parent


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
    yardimci = max(yardimcilar, key=lambda y: y.stat().st_size) if yardimcilar else None

    if ana is None and yardimci is not None:
        notlar.append(
            "PERSONEL klasorunde yalnizca 1C listesi var, ana personel verisi yok. "
            "Ana veri olmadan donem (gider ayi) kontrolu yapilamaz."
        )
    for atlanan in anadaylar:
        if atlanan is not ana:
            notlar.append(f"Atlandi (ana veri olarak en buyugu secildi): {atlanan.name}")
    return ana, yardimci, notlar


def fatura_dosyalarini_topla(kok: Path, argumanlar: list[str]) -> list[Path]:
    """Surukle-birak ile gelenler + 1_FATURALAR klasorundekiler."""
    bulunan: list[Path] = []
    gorulen: set[str] = set()

    def ekle(yol: Path) -> None:
        try:
            anahtar = str(yol.resolve()).lower()
        except OSError:
            anahtar = str(yol).lower()
        if anahtar not in gorulen and _tablo_dosyasi_mi(yol):
            gorulen.add(anahtar)
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
    return bulunan


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
        if mahsup.kapali_mi:
            yaz("  [TAMAM] Butun faturalar kapandi. Para kaybolmadi.")
        else:
            yaz("  [DIKKAT] MUTABAKAT ACIK. Bu tablo muhasebeye gonderilmemeli:")
            for k in mahsup.acik_kontroller:
                yaz(f"     {k.kaynak} ({k.para_birimi}) fark {k.fark:+.2f}")

        merkezler = mahsup.merkez_ozeti()
        if merkezler:
            yaz()
            yaz("  PROJE BAZINDA DAGILIM")
            for m in merkezler[:12]:
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
                         tasinan: list[Path]) -> Path | None:
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
        f"Islenen dosya  : {len(faturalar)}",
    ]
    for f in faturalar:
        satirlar.append(f"    - {f.name}")
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
        yaz(f"  Su klasore personel dosyalarini koyun:")
        yaz(f"     {kok / PERSONEL_DIZINI}")
        yaz()
        yaz("  Gereken dosyalar:")
        yaz("    1) Ana personel verisi (ornek: 2025_2026_giris_cikis.xlsx)  ZORUNLU")
        yaz("    2) 1C personel listesi (ornek: 1C_Personnel_List_...xlsx)   ONERILIR")
        yaz()
        yaz("  Bu dosyalari koyduktan sonra bu programi tekrar calistirin.")
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
    faturalar = fatura_dosyalarini_topla(kok, sys.argv[1:])
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
    yaz("  Personel verisi okunuyor. Ilk seferde 30 saniye kadar surebilir,")
    yaz("  sonraki calistirmalar cok daha hizli olur.")
    yaz()

    ayarlar = CalismaAyarlari(
        personel_yolu=ana,
        yardimci_personel_yolu=yardimci,
        veri_dizini=str(kok / AYAR_DIZINI),
        cikti_dizini=str(kok / CIKTI_DIZINI),
    )
    boru = Boru(ayarlar)

    son_yuzde = [-10.0]

    def ilerleme(yuzde: float, mesaj: str) -> None:
        """Boru hatti yuzdeyi 0-100 olcuginde bildirir.

        Her adimi basmiyoruz; 405 satirlik bir dosyada onlarca satir akiyor
        ve kullanici konsolda ne oldugunu takip edemiyor. Yuzde 5'ten az
        ilerleyen adimlar atlanir, sonuncusu her zaman basilir.
        """
        if yuzde - son_yuzde[0] < 5 and yuzde < 100:
            return
        son_yuzde[0] = yuzde
        yaz(f"    [%{yuzde:3.0f}] {mesaj}")

    damga = datetime.now().strftime("%Y%m%d_%H%M")
    sonuc = boru.calistir(faturalar, cikti_adi=f"Masraf_Dagitimi_{damga}.xlsx",
                          ilerleme=ilerleme)

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
        # ayni dosyalarin tekrar islenip cift sayilmasini onler.
        tasinan, tasima_hatalari = islenenleri_arsivle(kok, faturalar, damga)
        if tasinan:
            yaz()
            yaz(f"  Islenen {len(tasinan)} dosya arsive tasindi:")
            yaz(f"    {kok / ARSIV_DIZINI / damga}")
            yaz("  1_FATURALAR klasoru gelecek ay icin bos. Dosyalar silinmedi,")
            yaz("  gerekirse arsivden geri alabilirsiniz.")
        for h in tasima_hatalari:
            yaz(f"  UYARI: {h}")
        kayit = calistirma_kaydi_yaz(kok, damga, faturalar, ana, yardimci,
                                     sonuclar, sonuc.get("mahsup"), excel_yolu, tasinan)
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
