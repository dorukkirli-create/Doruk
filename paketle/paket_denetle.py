"""Windows paketini gonderMEDEN once denetler.

Neden var: paket, Windows tekerlekleri Linux'ta indirilerek kuruluyor. Bu
islemin sessiz bir tuzagi vardir ve gercek bir hataya yol acti.

``pip download --platform win_amd64`` yalnizca TEKERLEK UYUMLULUK ETIKETLERINI
degistirir. Ortam isaretlerini (environment markers) degistirmez; onlar
CALISAN yorumlayiciya gore degerlendirilir. Yani Linux'ta::

    tzdata; platform_system == "Windows"

isareti False cikar ve pip ``tzdata``yi sessizce ATLAR. Paket Linux'ta
kusursuz test edilir, Windows'ta ilk mailde patlar::

    Error occured using tzlocal. If you are seeing this, this is likely a
    problem with your installation of tzlocal or tzdata.

Sebep basit: Windows'ta IANA saat dilimi veritabani isletim sisteminde
yoktur, ``tzdata`` paketinden gelir. Linux'ta ``/usr/share/zoneinfo`` zaten
vardir, o yuzden eksiklik hicbir testte gorunmez.

Bu betik paketi tarar ve Windows'ta gerekip de kurulmamis bagimliliklari
listeler. Paketi gondermeden once calistirin::

    python paketle/paket_denetle.py /yol/Otomasyon

Cikis kodu 0 ise paket gonderilebilir.
"""

from __future__ import annotations

import email
import re
import sys
import zipfile
from pathlib import Path

#: Bu isaretler Windows'ta DOGRU, Linux'ta YANLIS olur. pip Linux'ta
#: degerlendirdigi icin bunlarla isaretli bagimliliklar atlanir.
WINDOWS_ISARETLERI = (
    'platform_system == "Windows"',
    "platform_system == 'Windows'",
    'sys_platform == "win32"',
    "sys_platform == 'win32'",
)

#: Calisma aninda hicbir zaman ice aktarilmadigi OLCULEN bagimliliklar.
#: Eksik olmalari sorun degildir; gurultu yapmasinlar.
#: (paket adi, neden guvenli)
BEKLENEN_EKSIKLER: dict[str, str] = {
    "win-unicode-console": (
        "pcodedmp'in bagimliligi. pcodedmp, extract_msg ice aktarilirken "
        "yuklenmiyor (olculdu), bu yuzden gerekmiyor. Tekerlegi de yok."
    ),
}

#: Paketin calismasi icin ice aktarilabilmesi ZORUNLU olan moduller.
ZORUNLU_MODULLER = (
    "pandas", "numpy", "openpyxl", "xlrd", "rapidfuzz", "xlsxwriter",
    "extract_msg", "olefile", "oletools", "bs4", "RTFDE", "compressed_rtf",
    "ebcdic", "tzlocal", "tzdata", "dateutil", "red_black_dict_mod",
)


def _paket_adi(ad: str) -> str:
    return re.split(r"[<>=!\[ ;]", ad.strip())[0].lower().replace("_", "-")


def site_packages_bul(kok: Path) -> Path | None:
    for aday in (kok / "program" / "Lib" / "site-packages",
                 kok / "Lib" / "site-packages", kok):
        if aday.is_dir() and any(aday.glob("*.dist-info")):
            return aday
    return None


def eksik_windows_bagimliliklari(sp: Path) -> list[tuple[str, str]]:
    """Windows'ta gerekip pakette olmayan bagimliliklari dondurur."""
    kurulu = {
        m.parent.name.split("-")[0].lower().replace("_", "-")
        for m in sp.glob("*.dist-info/METADATA")
    }
    eksik: list[tuple[str, str]] = []
    for meta_yolu in sorted(sp.glob("*.dist-info/METADATA")):
        sahip = meta_yolu.parent.name.split("-")[0]
        meta = email.message_from_string(
            meta_yolu.read_text(encoding="utf-8", errors="replace"))
        for satir in meta.get_all("Requires-Dist") or []:
            if ";" not in satir:
                continue
            ad, _, isaret = satir.partition(";")
            if "extra ==" in isaret:
                continue
            if not any(i in isaret for i in WINDOWS_ISARETLERI):
                continue
            gereken = _paket_adi(ad)
            if gereken not in kurulu:
                eksik.append((sahip, gereken))
    return eksik


def eksik_moduller(sp: Path) -> list[str]:
    """Zorunlu modullerden pakette bulunmayanlar."""
    eksik = []
    for modul in ZORUNLU_MODULLER:
        varmi = (
            (sp / modul).is_dir()
            or (sp / f"{modul}.py").is_file()
            or any(sp.glob(f"{modul}-*.dist-info"))
            or any(sp.glob(f"{modul.replace('_', '?')}-*.dist-info"))
        )
        if not varmi:
            eksik.append(modul)
    return eksik


def tzdata_saglam_mi(sp: Path) -> bool:
    """tzdata gercekten saat dilimi verisi tasiyor mu."""
    return (sp / "tzdata" / "zoneinfo" / "Europe" / "Moscow").is_file()


def denetle(kok: Path) -> int:
    print(f"Paket denetleniyor: {kok}\n")

    if kok.is_file() and kok.suffix.lower() == ".zip":
        print("HATA: ZIP degil, acilmis klasor verin.")
        return 2

    sp = site_packages_bul(kok)
    if sp is None:
        print("HATA: site-packages bulunamadi. Yol yanlis olabilir.")
        return 2

    sorun = 0

    yok = eksik_moduller(sp)
    if yok:
        print("[HATA] Zorunlu moduller eksik:")
        for m in yok:
            print(f"   - {m}")
        sorun += 1
    else:
        print(f"[OK] {len(ZORUNLU_MODULLER)} zorunlu modulun hepsi pakette.")

    eksik = eksik_windows_bagimliliklari(sp)
    beklenmeyen = [(s, g) for s, g in eksik if g not in BEKLENEN_EKSIKLER]
    beklenen = [(s, g) for s, g in eksik if g in BEKLENEN_EKSIKLER]
    if beklenmeyen:
        print("\n[HATA] Windows'ta gerekip pakette OLMAYAN bagimliliklar:")
        for sahip, gereken in beklenmeyen:
            print(f"   - {gereken}  ({sahip} istiyor)")
        print("\n   Bunlar Linux'ta test ederken GORUNMEZ. Indirin:")
        adlar = " ".join(sorted({g for _, g in beklenmeyen}))
        print(f"     pip download --dest wheels --platform win_amd64 \\")
        print(f"       --only-binary=:all: --python-version 3.11 {adlar}")
        sorun += 1
    else:
        print("[OK] Windows'a ozgu bagimliliklarin hepsi pakette.")
    for sahip, gereken in beklenen:
        print(f"     (bilerek atlandi: {gereken} - {BEKLENEN_EKSIKLER[gereken]})")

    if not tzdata_saglam_mi(sp):
        print("\n[HATA] tzdata saat dilimi verisi eksik veya bozuk.")
        print("   Windows'ta IANA saat dilimi veritabani isletim sisteminde YOKTUR.")
        print("   Bu olmadan extract_msg ilk mailde su hatayi verir:")
        print("     'Error occured using tzlocal ... problem with tzlocal or tzdata'")
        sorun += 1
    else:
        print("[OK] tzdata saat dilimi verisi yerinde.")

    for gereken in ("python.exe", "kod/calistir.py", "kod/masraf/boru.py"):
        if not (kok / "program" / gereken).exists():
            print(f"\n[HATA] Eksik: program/{gereken}")
            sorun += 1

    print()
    if sorun:
        print(f"SONUC: {sorun} sorun bulundu. PAKET GONDERILMEMELI.")
        return 1
    print("SONUC: Paket temiz.")
    return 0


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        print("Kullanim: python paketle/paket_denetle.py /yol/Otomasyon")
        return 2
    return denetle(Path(sys.argv[1]))


if __name__ == "__main__":
    sys.exit(main())
