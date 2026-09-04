"""Otomasyonun dogrulugunu ELLE DAGITILMIS dosyaya karsi olcer.

Ground truth: ``YUZYIL_TEMMUZ_2026_ELLE_DAGITILMIS.xlsx`` dosyasindaki
'SANTIYESI' kolonu. Bu kolonu bir insan doldurmustur.

Olcum akisi::

    1. Ham cari hareket dokumu (ANTIK_CARI_TEMMUZ_2026.xls) boru hattindan
       gecirilir; her satir icin kisi eslestirilir ve masraf merkezi bulunur.
    2. Elle dagitilmis dosya okunur.
    3. Iki dosyanin satirlari HIZALANIR. Dosyalar ayni islemleri ayni sirada
       icerir; hizalama yine de |tutar| + belge tarihi ile DOGRULANIR ve
       kayma olursa ileri bakis penceresiyle duzeltilir (bkz. ``hizala``).
    4. Otomasyonun masraf merkezi kodu ile elle yazilan santiye etiketi
       KARSILASTIRILIR. Iki taksonomi farklidir, bu yuzden otomasyonun
       sonucundan 'elle dosyada hangi etiket beklenirdi' kumesi turetilir
       (bkz. ``otomasyon_etiketleri``).

Taksonomi notu: elle dosyadaki etiketlerin cogu PROJE degil TUZEL KISI
adidir (RHI, RENSERVIS, ONE TOWER...). Sadece 'UST LUGA GPP' ve 'AMUR'
proje adidir. Otomasyon ise her zaman proje (gorev yeri) bazli calisir.
Bu yuzden karsilastirma otomasyonun bulduğu gorev yerinin BAGLI OLDUGU
tuzel kisi uzerinden yapilir.

Calistirma::

    python3 -m testler.dogruluk_olc
    python3 -m testler.dogruluk_olc --ogrenmeyi-kaydet   # veri/ dizinine yazar

Varsayilan olarak veri/ dizinine HICBIR SEY YAZILMAZ; olcum ogrenen
defterleri kirletmemelidir.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence

KOK = Path(__file__).resolve().parent.parent
if str(KOK) not in sys.path:
    sys.path.insert(0, str(KOK))

from masraf.boru import Boru, CalismaAyarlari  # noqa: E402
from masraf.modeller import (  # noqa: E402
    DURUM_ESLESMEDI,
    GiderSatiri,
    Sonuc,
)
from masraf.okuyucular.antik import yuzyil_dagitilmis_oku  # noqa: E402

__all__ = [
    "Karsilastirma",
    "DogrulukRaporu",
    "hizala",
    "elle_etiketler",
    "otomasyon_etiketleri",
    "olc",
]

# --------------------------------------------------------------------------
# Varsayilan dosya yollari
# --------------------------------------------------------------------------

HAM_DOSYA = KOK / "ornek_veri" / "antik_travel" / "ANTIK_CARI_TEMMUZ_2026.xls"
ELLE_DOSYA = KOK / "ornek_veri" / "antik_travel" / "YUZYIL_TEMMUZ_2026_ELLE_DAGITILMIS.xlsx"
PERSONEL_DOSYA = KOK / "ornek_veri" / "personel" / "2025_2026_giris_cikis.xlsx"

# --------------------------------------------------------------------------
# Taksonomi koprusu
# --------------------------------------------------------------------------

#: Elle dosyada gecen ve personel ana verisinde KARSILIGI OLMAYAN tuzel
#: kisiler. Bu satirlar otomasyonun kapsami disindadir: personel dosyasi
#: sadece RHI ve UST LUGA tuzel kisilerini icerir.
KAPSAM_DISI_ETIKETLER: frozenset[str] = frozenset({
    "RENSERVIS",
    "RENSTROYDETAL",
    "RSD",
    "RC PETER",
    "RC PETERSBURG",
    "RC MOSKOVA",
    "RC MOSCOW",
    "ONE TOWER",
    "TOP TOWER",
    "SAREN",
    "YAKA",
    "YAKA LLC",
})

#: Otomasyonun masraf merkezi KODUNDAN elle dosyada beklenecek etiketler.
#: Elle dosya iki projeyi ad ile ayirir; geri kalan her sey tuzel kisi
#: adiyla yazilir.
KOD_ETIKETLERI: dict[str, frozenset[str]] = {
    "GPP": frozenset({"UST LUGA GPP", "UST LUGA", "GPP"}),
    "AGPP": frozenset({"AMUR", "AMURSKY", "AGPZ"}),
}

#: Harita 'sirket' kolonundan turetilen genel etiketler.
SIRKET_ETIKETLERI: dict[str, frozenset[str]] = {
    "RHI": frozenset({"RHI", "RHI RUSSIA"}),
    "UST LUGA": frozenset({"UST LUGA", "UST LUGA GPP", "USTLUGA"}),
}


def _etiket_normalize(deger: Any) -> str:
    """Etiketi karsilastirilabilir bicime getirir (buyuk harf, tek bosluk)."""
    if deger is None:
        return ""
    metin = str(deger).strip().upper()
    for eski, yeni in (("İ", "I"), ("Ş", "S"), ("Ğ", "G"), ("Ü", "U"),
                       ("Ö", "O"), ("Ç", "C"), ("I", "I")):
        metin = metin.replace(eski, yeni)
    metin = metin.replace("-", " ").replace(".", " ")
    return " ".join(metin.split())


def elle_etiketler(satir: GiderSatiri) -> list[str]:
    """Elle dosyadaki satirin santiye etiket(ler)ini dondurur.

    Paylasimli satirlarda ('RHI 1/3- RENSTROYDETAL 2/3') birden fazla etiket
    doner; otomasyon bunlardan HERHANGI birini bulduysa eslesmis sayilir
    (paylasim orani otomasyonun hedefi degildir).
    """
    paylasim = satir.ek.get("paylasim") or []
    if paylasim:
        return [_etiket_normalize(p.get("masraf_merkezi")) for p in paylasim
                if _etiket_normalize(p.get("masraf_merkezi"))]
    tek = _etiket_normalize(satir.masraf_merkezi_kaynak)
    return [tek] if tek else []


def otomasyon_etiketleri(sonuc: Sonuc) -> frozenset[str]:
    """Otomasyonun sonucundan elle dosyada BEKLENEN etiket kumesini turetir.

    Once masraf merkezi koduna ozel bir esleme aranir (GPP, AGPP); yoksa
    gorev yerinin bagli oldugu tuzel kisi (harita 'sirket' kolonu, yoksa
    personel kaydindaki 'Sirket 2') kullanilir.
    """
    kod = _etiket_normalize(sonuc.masraf_merkezi)
    if kod in KOD_ETIKETLERI:
        return KOD_ETIKETLERI[kod]
    sirket = _etiket_normalize(sonuc.sirket) or _etiket_normalize(sonuc.sirket2)
    if sirket in SIRKET_ETIKETLERI:
        return SIRKET_ETIKETLERI[sirket]
    if sirket:
        return frozenset({sirket})
    if kod:
        return frozenset({kod})
    return frozenset()


# --------------------------------------------------------------------------
# Hizalama
# --------------------------------------------------------------------------


def _tutar_yakin(a: float | None, b: float | None, tolerans: float = 0.02) -> bool:
    """Iki tutari MUTLAK DEGER uzerinden karsilastirir.

    Iade satirlarinda ham dosya negatif (alacak), elle dosya pozitif yazar;
    bu bir hizalama farki degildir.
    """
    if a is None or b is None:
        return a is None and b is None
    return abs(abs(a) - abs(b)) <= tolerans


def _tarih_yakin(a: date | None, b: date | None) -> bool:
    """Belge tarihlerini karsilastirir; biri bossa engel sayilmaz."""
    if a is None or b is None:
        return True
    return a == b


def _puan(ham: GiderSatiri, elle: GiderSatiri) -> int:
    """Iki satirin ayni islem olma gucu (2 = tutar+tarih, 1 = tutar, 0 = hayir)."""
    if not _tutar_yakin(ham.tutar, elle.tutar):
        return 0
    return 2 if _tarih_yakin(ham.belge_tarihi, elle.belge_tarihi) else 1


def hizala(
    ham_satirlar: Sequence[GiderSatiri],
    elle_satirlar: Sequence[GiderSatiri],
    pencere: int = 5,
) -> list[tuple[int | None, int | None, str]]:
    """Iki dosyanin satirlarini sirali olarak hizalar.

    Dosyalar ayni islemleri ayni sirada icerir; bu yuzden iki imlecli sirali
    yurume kullanilir. Her adimda once dogrudan karsilik (tutar + tarih)
    denenir; tutmazsa ileri bakis penceresi icinde ilk uyan cift aranir ve
    atlanan satirlar eslesmemis olarak isaretlenir.

    Returns:
        (ham_indeks, elle_indeks, yontem) uclulerinin listesi. Indekslerden
        biri None ise o satirin karsiligi bulunamamistir.
    """
    ciftler: list[tuple[int | None, int | None, str]] = []
    i = j = 0
    while i < len(ham_satirlar) and j < len(elle_satirlar):
        puan = _puan(ham_satirlar[i], elle_satirlar[j])
        if puan == 2:
            ciftler.append((i, j, "tutar+tarih"))
            i += 1
            j += 1
            continue
        if puan == 1:
            ciftler.append((i, j, "tutar"))
            i += 1
            j += 1
            continue
        # Kayma var: pencere icinde ilk saglam cifti ara.
        bulundu = False
        for adim in range(1, pencere + 1):
            for di, dj in ((adim, 0), (0, adim), (adim, adim)):
                ni, nj = i + di, j + dj
                if ni >= len(ham_satirlar) or nj >= len(elle_satirlar):
                    continue
                if _puan(ham_satirlar[ni], elle_satirlar[nj]) == 2:
                    for k in range(i, ni):
                        ciftler.append((k, None, "hizalanamadi"))
                    for k in range(j, nj):
                        ciftler.append((None, k, "hizalanamadi"))
                    ciftler.append((ni, nj, "tutar+tarih"))
                    i, j = ni + 1, nj + 1
                    bulundu = True
                    break
            if bulundu:
                break
        if not bulundu:
            # Pencerede karsilik yok; sira korunuyor varsayilir ama isaretlenir.
            ciftler.append((i, j, "sira (dogrulanamadi)"))
            i += 1
            j += 1
    for k in range(i, len(ham_satirlar)):
        ciftler.append((k, None, "hizalanamadi"))
    for k in range(j, len(elle_satirlar)):
        ciftler.append((None, k, "hizalanamadi"))
    return ciftler


# --------------------------------------------------------------------------
# Karsilastirma kayitlari
# --------------------------------------------------------------------------

# Karsilastirma durumlari
AYNI = "AYNI"                      # otomasyon ve elle ayni sonuca varmis
FARKLI = "FARKLI"                  # ikisi de bir sey soylemis, sonuc celisiyor
KAPSAM_DISI = "KAPSAM DISI"        # elle grup sirketi yazmis, otomasyon da kisi bulamamis
CELISKI_GRUP = "CELISKI (GRUP SIRKETI)"  # elle grup sirketi ama otomasyon RHI/UL calisani buldu
ELLE_BOS = "ELLE BOS"              # elle dosyada santiye yazilmamis
OTOMASYON_YOK = "OTOMASYON BULAMADI"  # elle yazmis, otomasyon eslestiremmis
HIZALANAMADI = "HIZALANAMADI"


@dataclass
class Karsilastirma:
    """Tek bir islem satirinin otomasyon/elle karsilastirmasi."""

    sira: int
    s_no: str | None
    tarih: date | None
    tutar: float | None
    aciklama: str
    hizalama: str
    kisi_ham: str | None
    elle_kisi: str | None
    elle_etiket: list[str]
    otomasyon_kodu: str | None
    otomasyon_gorev_yeri: str | None
    otomasyon_sirket: str | None
    otomasyon_etiket: frozenset[str]
    sicil: str | None
    personel_adi: str | None
    yontem: str
    guven: float
    otomasyon_durumu: str
    uyarilar: list[str] = field(default_factory=list)
    durum: str = FARKLI

    @property
    def elle_metni(self) -> str:
        return " + ".join(self.elle_etiket) if self.elle_etiket else "(bos)"

    @property
    def otomasyon_metni(self) -> str:
        if self.otomasyon_kodu:
            return f"{self.otomasyon_kodu} ({self.otomasyon_gorev_yeri})"
        return "(bulunamadi)"


@dataclass
class DogrulukRaporu:
    """Tum olcumun sonucu."""

    karsilastirmalar: list[Karsilastirma]
    ozet: dict
    boru_ozeti: dict

    def sayim(self, durum: str) -> int:
        return sum(1 for k in self.karsilastirmalar if k.durum == durum)

    @property
    def karsilastirilabilir(self) -> list[Karsilastirma]:
        """Dogruluk yuzdesinin paydasi: iki tarafin da sonuc urettigi satirlar."""
        return [k for k in self.karsilastirmalar if k.durum in (AYNI, FARKLI)]

    @property
    def dogruluk(self) -> float:
        payda = len(self.karsilastirilabilir)
        if not payda:
            return 0.0
        return round(self.sayim(AYNI) / payda * 100, 1)


def _durum_belirle(k: Karsilastirma) -> str:
    """Bir karsilastirmanin durumunu belirler."""
    if k.hizalama == "hizalanamadi":
        return HIZALANAMADI
    if not k.elle_etiket:
        return ELLE_BOS
    grup_sirketi = all(e in KAPSAM_DISI_ETIKETLER for e in k.elle_etiket)
    otomasyon_var = bool(k.otomasyon_etiket) and k.otomasyon_durumu != DURUM_ESLESMEDI
    if grup_sirketi:
        # Elle dosya grup sirketine yazmis. Otomasyon personel ana verisinde
        # (RHI / UST LUGA) bir calisan bulduysa bu gercek bir celiskidir:
        # ya elle etiket yanlis, ya kisi baska sirkete odunc calisiyor.
        return CELISKI_GRUP if otomasyon_var else KAPSAM_DISI
    if not otomasyon_var:
        return OTOMASYON_YOK
    return AYNI if (set(k.elle_etiket) & set(k.otomasyon_etiket)) else FARKLI


def karsilastir(
    ham_satirlar: Sequence[GiderSatiri],
    sonuclar: Sequence[Sonuc],
    elle_satirlar: Sequence[GiderSatiri],
) -> list[Karsilastirma]:
    """Hizalanmis satirlari tek tek karsilastirir."""
    # Boru hatti satirlari filtreleyebilir; kimlik uzerinden geri esle.
    sonuc_haritasi: dict[tuple[str, int], Sonuc] = {
        (s.satir.kaynak_dosya, s.satir.satir_no): s for s in sonuclar
    }
    ciktilar: list[Karsilastirma] = []
    for sira, (hi, ei, yontem) in enumerate(
        hizala(ham_satirlar, elle_satirlar), start=1
    ):
        ham = ham_satirlar[hi] if hi is not None else None
        elle = elle_satirlar[ei] if ei is not None else None
        sonuc = sonuc_haritasi.get((ham.kaynak_dosya, ham.satir_no)) if ham else None

        k = Karsilastirma(
            sira=sira,
            s_no=(elle.ek.get("s_no") if elle else None),
            tarih=(ham.belge_tarihi if ham else (elle.belge_tarihi if elle else None)),
            tutar=(ham.tutar if ham else (elle.tutar if elle else None)),
            aciklama=(ham.aciklama if ham else (elle.aciklama if elle else "")),
            hizalama=yontem,
            kisi_ham=(ham.kisi_ham if ham else None),
            elle_kisi=(elle.kisi_ham if elle else None),
            elle_etiket=(elle_etiketler(elle) if elle else []),
            otomasyon_kodu=(sonuc.masraf_merkezi if sonuc else None),
            otomasyon_gorev_yeri=(sonuc.gorev_yeri if sonuc else None),
            otomasyon_sirket=(sonuc.sirket or sonuc.sirket2 if sonuc else None),
            otomasyon_etiket=(otomasyon_etiketleri(sonuc) if sonuc else frozenset()),
            sicil=(sonuc.eslesme.sicil if sonuc else None),
            personel_adi=(sonuc.eslesme.ad_soyad if sonuc else None),
            yontem=(sonuc.eslesme.yontem if sonuc else "yok"),
            guven=(sonuc.eslesme.guven if sonuc else 0.0),
            otomasyon_durumu=(sonuc.durum if sonuc else DURUM_ESLESMEDI),
            uyarilar=(list(sonuc.uyarilar) if sonuc else []),
        )
        k.durum = _durum_belirle(k)
        ciktilar.append(k)
    return ciktilar


# --------------------------------------------------------------------------
# Olcum
# --------------------------------------------------------------------------


def olc(
    ham_yolu: str | Path = HAM_DOSYA,
    elle_yolu: str | Path = ELLE_DOSYA,
    personel_yolu: str | Path = PERSONEL_DOSYA,
    veri_dizini: str | Path = KOK / "veri",
    ogrenmeyi_kaydet: bool = False,
) -> DogrulukRaporu:
    """Ham dosyayi boru hattindan gecirir ve elle dosyaya karsi olcer."""
    from masraf.okuyucular.antik import antik_cari_oku

    boru = Boru(
        CalismaAyarlari(
            personel_yolu=personel_yolu,
            veri_dizini=veri_dizini,
            ogrenmeyi_kaydet=ogrenmeyi_kaydet,
        )
    )
    sonuclar = boru.isle([ham_yolu])
    boru_ozeti = boru.ozet(sonuclar)

    ham_satirlar = antik_cari_oku(ham_yolu)
    elle_satirlar = yuzyil_dagitilmis_oku(elle_yolu)
    karsilastirmalar = karsilastir(ham_satirlar, sonuclar, elle_satirlar)

    ozet = {
        "ham_satir": len(ham_satirlar),
        "elle_satir": len(elle_satirlar),
        "hizalanan": sum(1 for k in karsilastirmalar if k.durum != HIZALANAMADI),
    }
    rapor = DogrulukRaporu(karsilastirmalar, ozet, boru_ozeti)
    return rapor, boru


# --------------------------------------------------------------------------
# Raporlama
# --------------------------------------------------------------------------


def _kanit(k: Karsilastirma, boru: Boru) -> list[str]:
    """Bir uyusmazlik icin personel verisinden kanit satirlari uretir."""
    satirlar: list[str] = []
    if not k.sicil:
        satirlar.append("      personel verisinde karsilik YOK")
        return satirlar
    kayit = boru.defter.donem_kaydi(k.sicil, k.tarih) or boru.defter.sicil_ile(k.sicil)
    if not kayit:
        satirlar.append(f"      sicil {k.sicil} personel verisinde bulunamadi")
        return satirlar
    satirlar.append(
        f"      personel: {k.sicil} / {kayit.get('ad_soyad')} / "
        f"{kayit.get('gorev_yeri')} / {kayit.get('sirket2')} / "
        f"{kayit.get('statu')} / donem {kayit.get('donem')} / {kayit.get('kategori')}"
    )
    satirlar.append(f"      eslestirme: {k.yontem} (guven {k.guven:.2f})")
    for uyari in k.uyarilar[:3]:
        satirlar.append(f"      uyari: {uyari}")
    return satirlar


def rapor_yaz(rapor: DogrulukRaporu, boru: Boru) -> None:
    """Olcum sonucunu konsola basar."""
    print("=" * 78)
    print("DOGRULUK OLCUMU - Temmuz 2026 seyahat faturasi")
    print("=" * 78)
    print(f"Ham dosya satiri      : {rapor.ozet['ham_satir']}")
    print(f"Elle dosya satiri     : {rapor.ozet['elle_satir']}")
    print(f"Hizalanan satir       : {rapor.ozet['hizalanan']}")
    print()
    print("BORU HATTI DURUMU")
    for ad, adet in rapor.boru_ozeti["durum_dagilimi"].items():
        print(f"  {ad:<10} {adet:>4}")
    print(f"  otomatik orani: %{rapor.boru_ozeti['otomatik_orani']}")
    print()
    print("KARSILASTIRMA")
    for durum in (AYNI, FARKLI, CELISKI_GRUP, KAPSAM_DISI,
                  OTOMASYON_YOK, ELLE_BOS, HIZALANAMADI):
        print(f"  {durum:<24} {rapor.sayim(durum):>4}")
    print()
    print(f"DOGRULUK: %{rapor.dogruluk} "
          f"({rapor.sayim(AYNI)}/{len(rapor.karsilastirilabilir)} "
          f"karsilastirilabilir satirda ayni)")
    print()

    for baslik, durum in (
        ("UYUSMAZLIKLAR (ikisi de sonuc uretti, sonuc farkli)", FARKLI),
        ("CELISKI: elle grup sirketi yazmis, otomasyon RHI/UL calisani buldu", CELISKI_GRUP),
        ("OTOMASYON BULAMADI (elle yazmis, otomasyon eslestiremedi)", OTOMASYON_YOK),
        ("KAPSAM DISI (grup sirketi; personel verisinde yok)", KAPSAM_DISI),
        ("ELLE BOS (insan santiye yazmamis)", ELLE_BOS),
        ("HIZALANAMADI", HIZALANAMADI),
    ):
        secilenler = [k for k in rapor.karsilastirmalar if k.durum == durum]
        if not secilenler:
            continue
        print("-" * 78)
        print(f"{baslik}  [{len(secilenler)}]")
        print("-" * 78)
        for k in secilenler:
            print(
                f"  #{k.s_no or k.sira} {k.tarih} {k.tutar:>10.2f} "
                f"{(k.kisi_ham or k.elle_kisi or '?')[:30]:<30}"
            )
            print(f"      elle      : {k.elle_metni}")
            print(f"      otomasyon : {k.otomasyon_metni} [{k.otomasyon_durumu}]")
            for satir in _kanit(k, boru):
                print(satir)
            print(f"      aciklama  : {k.aciklama[:90]}")
            print()


def main(argv: Sequence[str] | None = None) -> int:
    """Komut satiri girisi."""
    ayrist = argparse.ArgumentParser(description="Otomasyon dogrulugunu olcer.")
    ayrist.add_argument("--ham", default=str(HAM_DOSYA))
    ayrist.add_argument("--elle", default=str(ELLE_DOSYA))
    ayrist.add_argument("--personel", default=str(PERSONEL_DOSYA))
    ayrist.add_argument("--veri", default=str(KOK / "veri"))
    ayrist.add_argument(
        "--ogrenmeyi-kaydet",
        action="store_true",
        help="Beslenen defterleri veri/ dizinine yaz (varsayilan: yazma).",
    )
    secim = ayrist.parse_args(argv)

    rapor, boru = olc(
        secim.ham, secim.elle, secim.personel, secim.veri, secim.ogrenmeyi_kaydet
    )
    rapor_yaz(rapor, boru)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
