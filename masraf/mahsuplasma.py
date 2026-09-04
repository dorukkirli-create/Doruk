"""Mahsuplasma tablosu: her fatura kaleminin projelere dagilimi.

Satir seviyesindeki sonuc (``Sonuc`` listesi) "bu satir kime ve hangi projeye
ait" sorusunu cevaplar. Finansin muhasebeye verecegi sey ise farkli bir sey:
her fatura icin ALTINDA hangi projelere ne kadar yazilacagi.

Bu modul o donusumu yapar ve dort kurala uyar:

1. **Her fatura kendi icinde kapanir.** Okunan tutar = yinelenen + dagitilan +
   dagitilamayan. Kurus farki bile birakilmaz: yuvarlama artigi faturanin en
   buyuk satirina eklenir. Para asla kaybolmaz.
2. **Dagitilamayan tutar gorunur.** Sessizce dusurulmez; kendi satiri olur ve
   kontrol tablosunda raporlanir.
3. **Yinelenen islemler bir kez sayilir.** Ayni mail iki dosya tasiyabilir:
   acentenin ham cari dokumu ve ayni islemlerin elle dagitilmis hali. Ikisi de
   okunursa para CIFT sayilir.
4. **Yinelenmeler silinmez, raporlanir.** Elenen dosya kontrol tablosunda
   'yinelenen' sutunuyla gorunur; boylece finans neyin neden dusuruldugunu
   gorur.

Olculen ornek (Temmuz 2026, tek Outlook mesaji)::

    ENERGO TEMMUZ.xls    134 islem   48.946,59 USD   (acentenin ham dokumu)
    YUZYIL TEMMUZ.xlsx   134 islem   48.978,59 USD   (ayni islemler, elle dagitilmis)

Bu ikisi AYNI islemlerdir. Naif toplama 97.925,18 USD verir; gercek rakam
yarisidir. Tek fark 14.07.2026 tarihli bir kalemdir: ham dokumde -16,00 USD
(iade), elle dagitilmis halde +16,00 USD. Aradaki 32,00 USD tam olarak budur.
Isaret celiskisi ``isaret_celiskileri`` listesinde raporlanir.

Yineleme anahtari
-----------------
(belge tarihi, mutlak tutar, isim harflerinin siralanmis hali) uclusudur.
Isim SIRADAN BAGIMSIZ karsilastirilir cunku iki dosya ayni kisiyi farkli
yazar: ham dokum 'OZAKAY MUSTAFAKEMAL', elle dagitilmis hal
'MUSTAFA KEMAL OZAKAY'. Harfler siralandiginda ikisi ayni anahtari verir.
Tarih ve tutar zaten esitken iki farkli kisinin isminin harf harf ayni
olmasi pratikte imkansizdir.

Diger kurallar
--------------
* Kisi listeleri (katilimci listesi, saglik kontrol listesi) tutar tasimaz;
  bunlar fatura degil kutuktur ve dagilima girmez.
* Bir satir birden fazla tuzel kisiye paylastirilabilir ('RHI 1/3 -
  RENSTROYDETAL 2/3'). Bu etiketler PROJE degil SIRKET adidir; proje ayni
  kalir, tutar sirketler arasinda bolunur. Etiket masraf merkezi haritasinda
  bir projeye karsilik geliyorsa o zaman proje de bolunur.
* Farkli para birimleri ASLA toplanmaz; her doviz kendi icinde kapanir.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Sequence

from masraf.metin import isim_normalize

__all__ = [
    "MahsupSatiri",
    "KontrolSatiri",
    "IsaretCeliskisi",
    "MahsupTablosu",
    "mahsuplasma_uret",
    "DAGITILAMAYAN",
]

#: Masraf merkezi cozulemeyen tutarlar bu etiket altinda toplanir.
DAGITILAMAYAN = "(DAGITILAMAYAN)"

#: Yinelenen islem tespitinde hangi kaynak tipi tercih edilir (kucuk = oncelikli).
#: Ham cari dokum faturanin kendisidir; elle dagitilmis hali ondan turetilmistir,
#: dolayisiyla insan hatasi tasiyabilir.
_KAYNAK_ONCELIGI: dict[str, int] = {
    "antik_cari": 0,
    "energo_arabulucu": 1,
    "energo_assessment": 1,
    "yuzyil_dagitilmis": 2,
    "genel": 3,
}

#: Dagilima girmeyen kaynak tipleri: bunlar fatura degil kisi kutugudur.
_KUTUK_TIPLERI = frozenset({"referans_liste", "energo_saglik", "koc_katilimci"})

#: Kurus altinda kalan farklar kapali sayilir.
_TOLERANS = 0.005


@dataclass
class MahsupSatiri:
    """Bir faturanin tek bir masraf merkezine (ve sirkete) dusen payi."""

    kaynak: str                    # fatura / kaynak dosya adi
    gider_tipi: str
    masraf_merkezi: str            # finans kodu, gorev yeri metni, ya da DAGITILAMAYAN
    masraf_merkezi_adi: str | None
    sirket: str | None
    para_birimi: str
    tutar: float
    satir_sayisi: int
    kisi_sayisi: int
    otomatik: int = 0
    incele: int = 0
    eslesmedi: int = 0
    #: Bu satirdaki giderlerin belge tarihi araligi. Tek bir tarih degil,
    #: cunku bir mahsup satiri birden fazla gunun kalemini toplayabilir.
    ilk_tarih: date | None = None
    son_tarih: date | None = None
    #: Masraf merkezi haritasinda tanimli bir kod mu, yoksa ham gorev yeri metni mi.
    haritada_var: bool = True
    #: 'RHI 1/3' gibi paylasim notu; paylastirilmamis satirlarda None.
    pay_notu: str | None = None

    @property
    def dagitildi_mi(self) -> bool:
        return self.masraf_merkezi != DAGITILAMAYAN

    @property
    def gider_donemi(self) -> str:
        """Muhasebenin hangi doneme yazacagini soyleyen okunakli etiket.

        Tek ay ise '07.2026', birden fazla aya yayiliyorsa '04.2026 - 07.2026'.
        """
        if self.ilk_tarih is None:
            return ""
        bas = f"{self.ilk_tarih:%m.%Y}"
        if self.son_tarih is None or (self.son_tarih.year, self.son_tarih.month) == (
            self.ilk_tarih.year, self.ilk_tarih.month
        ):
            return bas
        return f"{bas} - {self.son_tarih:%m.%Y}"

    @property
    def kontrol_gerek(self) -> bool:
        """Muhasebeye gitmeden once insan gozu gerekiyor mu."""
        return (not self.dagitildi_mi) or (not self.haritada_var) \
            or self.incele > 0 or self.eslesmedi > 0


@dataclass
class IsaretCeliskisi:
    """Ayni islem iki dosyada zit isaretle gorunuyor."""

    belge_tarihi: date | None
    kisi: str | None
    para_birimi: str
    kaynaklar: tuple[str, ...]
    tutarlar: tuple[float, ...]
    kullanilan: float

    def aciklama(self) -> str:
        parcalar = ", ".join(
            f"{k}: {t:+,.2f}" for k, t in zip(self.kaynaklar, self.tutarlar)
        )
        return (
            f"{self.kisi or '(kisi yok)'} / "
            f"{self.belge_tarihi:%d.%m.%Y} tarihli kalem iki dosyada zit isaretli "
            f"({parcalar}). Ham dokumdeki deger kullanildi: "
            f"{self.kullanilan:+,.2f} {self.para_birimi}."
        )


@dataclass
class KontrolSatiri:
    """Bir kaynak dosya icin gelen / yinelenen / dagitilan mutabakati.

    ``gelen`` dosyadan OKUNAN her seydir. Yinelenen tutar baska bir dosyada
    zaten sayildigi icin dagilima girmez ama burada gorunur; boylece toplam
    daima kapanir.
    """

    kaynak: str
    para_birimi: str
    gelen: float
    dagitilan: float
    dagitilamayan: float
    satir_sayisi: int
    yinelenen_tutar: float = 0.0
    yinelenen_satir: int = 0

    @property
    def fark(self) -> float:
        """Sifir olmalidir. Degilse dagitimda kayip var demektir."""
        return round(
            self.gelen - (self.dagitilan + self.dagitilamayan + self.yinelenen_tutar), 2
        )

    @property
    def kapali_mi(self) -> bool:
        return abs(self.fark) < 0.01

    @property
    def net(self) -> float:
        """Bu dosyanin gercekten kattigi tutar (yinelenenler dusulmus)."""
        return round(self.dagitilan + self.dagitilamayan, 2)

    @property
    def dagitim_orani(self) -> float:
        return (self.dagitilan / self.net * 100.0) if self.net else 0.0


@dataclass
class MahsupTablosu:
    """Mahsuplasma ciktisinin tamami."""

    satirlar: list[MahsupSatiri] = field(default_factory=list)
    kontrol: list[KontrolSatiri] = field(default_factory=list)
    isaret_celiskileri: list[IsaretCeliskisi] = field(default_factory=list)
    yinelenen_sayisi: int = 0
    kutuk_satir_sayisi: int = 0
    tutarsiz_satir_sayisi: int = 0

    def merkez_ozeti(self) -> list[dict]:
        """Masraf merkezi bazinda toplam (butun faturalar birlestirilmis)."""
        birikim: dict[tuple[str, str], dict] = {}
        for s in self.satirlar:
            anahtar = (s.masraf_merkezi, s.para_birimi)
            kayit = birikim.setdefault(anahtar, {
                "masraf_merkezi": s.masraf_merkezi,
                "masraf_merkezi_adi": s.masraf_merkezi_adi,
                "sirket": s.sirket,
                "para_birimi": s.para_birimi,
                "haritada_var": s.haritada_var,
                "tutar": 0.0, "satir_sayisi": 0,
                "otomatik": 0, "incele": 0, "eslesmedi": 0,
                "_kisiler": set(),
            })
            kayit["tutar"] += s.tutar
            kayit["satir_sayisi"] += s.satir_sayisi
            kayit["otomatik"] += s.otomatik
            kayit["incele"] += s.incele
            kayit["eslesmedi"] += s.eslesmedi
            kayit["haritada_var"] = kayit["haritada_var"] and s.haritada_var
            # Kisi sayisi merkez bazinda BENZERSIZ olmali: ayni kisi hem otel
            # hem bilet satirinda gorunur, iki kez sayilmamali.
            kayit["_kisiler"] |= getattr(s, "_kimlikler", set())

        toplamlar: dict[str, float] = defaultdict(float)
        for kayit in birikim.values():
            toplamlar[kayit["para_birimi"]] += kayit["tutar"]
        cikti = []
        for kayit in birikim.values():
            genel = toplamlar[kayit["para_birimi"]]
            kayit["tutar"] = round(kayit["tutar"], 2)
            kayit["kisi_sayisi"] = len(kayit.pop("_kisiler"))
            kayit["pay_yuzde"] = round(kayit["tutar"] / genel * 100, 2) if genel else 0.0
            cikti.append(kayit)
        return sorted(cikti, key=lambda k: (-k["tutar"], k["masraf_merkezi"]))

    def toplamlar(self) -> dict[str, dict[str, float]]:
        """Para birimi bazinda gelen / yinelenen / dagitilan / dagitilamayan."""
        sonuc: dict[str, dict[str, float]] = {}
        for k in self.kontrol:
            d = sonuc.setdefault(k.para_birimi, {
                "gelen": 0.0, "yinelenen": 0.0, "dagitilan": 0.0, "dagitilamayan": 0.0,
            })
            d["gelen"] += k.gelen
            d["yinelenen"] += k.yinelenen_tutar
            d["dagitilan"] += k.dagitilan
            d["dagitilamayan"] += k.dagitilamayan
        for d in sonuc.values():
            d["net"] = round(d["dagitilan"] + d["dagitilamayan"], 2)
            for anahtar in ("gelen", "yinelenen", "dagitilan", "dagitilamayan"):
                d[anahtar] = round(d[anahtar], 2)
            d["oran"] = (d["dagitilan"] / d["net"] * 100.0) if d["net"] else 0.0
        return sonuc

    @property
    def acik_kontroller(self) -> list[KontrolSatiri]:
        """Kapanmayan mutabakat satirlari. Bos olmalidir."""
        return [k for k in self.kontrol if not k.kapali_mi]

    @property
    def kapali_mi(self) -> bool:
        return not self.acik_kontroller


def _kaynak_adi(sonuc: Any) -> str:
    """Satirin ait oldugu faturayi adlandirir.

    Outlook mesajindan gelen satirlarda kaynak 'mesaj.msg > ek.xlsx' bicimindedir;
    fatura o ekin kendisidir.
    """
    ham = str(getattr(sonuc.satir, "kaynak_dosya", "") or "")
    return ham.split("> ")[-1].strip() or "(bilinmeyen)"


def _isim_imzasi(ham: str | None) -> str:
    """Isim sirasindan ve bosluklardan bagimsiz karsilastirma imzasi.

    'OZAKAY MUSTAFAKEMAL' ve 'MUSTAFA KEMAL OZAKAY' ayni imzayi verir; iki
    dosya ayni kisiyi farkli sirayla ve farkli bitisiklikte yazdigi icin
    gereklidir.
    """
    normal = isim_normalize(ham or "")
    return "".join(sorted(ch for ch in normal if ch.isalnum()))


def _yineleme_anahtari(sonuc: Any) -> tuple | None:
    """Ayni islemi iki farkli dosyada tanimak icin KABA anahtar.

    (belge tarihi, mutlak tutar, para birimi). Isim BILEREK disarida birakilir:
    iki dosya ayni kisiyi farkli yazar, hatta kirpar ('OZAKAY MUSTAFAKEMA'),
    hatta yanlis yazar ('YALCINKAYA ANIL' / 'ALI YALCINKAYA'). Isim bu kaba
    kova icinde ESLESTIRICI olarak kullanilir, anahtar olarak degil.

    Mutlak tutar kullanilir: bir dosya iadeyi eksi, digeri arti yazabilir.
    Tarih veya tutar yoksa yineleme tespiti yapilmaz (None doner) ve satir
    oldugu gibi korunur.
    """
    satir = sonuc.satir
    if satir.tutar is None or not satir.belge_tarihi:
        return None
    return (satir.belge_tarihi, round(abs(float(satir.tutar)), 2),
            satir.para_birimi or "?")


def _eslesme_puani(a: Any, b: Any) -> int:
    """Ayni kovadaki iki kaydin ayni kisi olma gucu. Buyuk = daha guclu.

    3: isim imzalari birebir ayni.
    2: biri digerinin kirpilmis hali ('OZAKAY MUSTAFAKEMA' <- 'OZAKAY MUSTAFAKEMAL').
    1: ortak token var (bir kelime yanlis yazilmis olabilir).
    0: isim yok veya hicbir benzerlik yok.
    """
    ia, ib = _isim_imzasi(a.satir.kisi_ham), _isim_imzasi(b.satir.kisi_ham)
    if ia and ib:
        if ia == ib:
            return 3
        kisa, uzun = (ia, ib) if len(ia) <= len(ib) else (ib, ia)
        # Kirpilmis isim: kisa olanin butun harfleri uzunda var ve fark kucuk.
        if len(uzun) - len(kisa) <= 4 and _harf_alt_kumesi(kisa, uzun):
            return 2
        ta = set(isim_normalize(a.satir.kisi_ham or "").split())
        tb = set(isim_normalize(b.satir.kisi_ham or "").split())
        if ta & tb:
            return 1
    return 0


def _harf_alt_kumesi(kisa: str, uzun: str) -> bool:
    """Siralanmis harf dizisi ``kisa``, ``uzun`` icinde alt dizi mi."""
    i = 0
    for ch in uzun:
        if i < len(kisa) and kisa[i] == ch:
            i += 1
    return i == len(kisa)


def _kovayi_esle(tutulanlar: list[Any], digerleri: list[Any]) -> list[tuple[Any, Any]]:
    """Ayni kovadaki kayitlari 1:1 eslestirir; eslesemeyenler disarida kalir.

    Once en guclu isim eslesmeleri baglanir, sonra kalanlar sirayla. Kova
    zaten (tarih, tutar, doviz) esitligini garanti ettigi icin isimsiz veya
    yanlis yazilmis kayitlar da dogru sekilde eslesir.

    Returns:
        [(tutulan, elenen)] ciftleri. ``digerleri`` icinde ciftlenmeyenler
        gercekten yeni islemlerdir ve korunur.
    """
    ciftler: list[tuple[Any, Any]] = []
    bos_tutulan = list(tutulanlar)
    bekleyen = list(digerleri)
    for esik in (3, 2, 1):
        kalan_bekleyen: list[Any] = []
        for d in bekleyen:
            en_iyi = None
            for t in bos_tutulan:
                if _eslesme_puani(t, d) >= esik:
                    en_iyi = t
                    break
            if en_iyi is None:
                kalan_bekleyen.append(d)
            else:
                bos_tutulan.remove(en_iyi)
                ciftler.append((en_iyi, d))
        bekleyen = kalan_bekleyen
    # Kalanlar: isim eslesmedi ama kova ayni. Sirayla baglanir.
    for d in bekleyen:
        if not bos_tutulan:
            break
        ciftler.append((bos_tutulan.pop(0), d))
    return ciftler


def _talimati_devral(tutulan: Any, elenen: Any) -> None:
    """Elenen kaydin tasidigi DAGITIM TALIMATINI tutulan kayda aktarir.

    Iki dosya birbirini tamamlar: ham cari dokum TUTARI dogru tasir, elle
    dagitilmis hal ise insanin ALDIGI KARARI tasir ('RHI 1/3 - RENSTROYDETAL
    2/3' gibi). Elenen kayit atilirken bu karar da atilirsa veri kaybolur.
    Bu yuzden tutulan kayitta yoksa devralinir; varsa dokunulmaz.
    """
    if not isinstance(tutulan.satir.ek, dict) or not isinstance(elenen.satir.ek, dict):
        return
    if not tutulan.satir.ek.get("paylasim") and elenen.satir.ek.get("paylasim"):
        tutulan.satir.ek["paylasim"] = elenen.satir.ek["paylasim"]
        tutulan.satir.ek["paylasim_kaynagi"] = _kaynak_adi(elenen)
    if not tutulan.satir.masraf_merkezi_kaynak and elenen.satir.masraf_merkezi_kaynak:
        tutulan.satir.ek["kaynak_santiye_devralindi"] = elenen.satir.masraf_merkezi_kaynak
    if tutulan.masraf_merkezi is None and elenen.masraf_merkezi:
        # Kisi ham dokumde cozulemedi ama elle dagitilmis halde cozulmus.
        # Sessizce almiyoruz; sadece isaretliyoruz ki operator gorsun.
        tutulan.satir.ek["yinelenen_kayittan_oneri"] = elenen.masraf_merkezi


def _oncelik(sonuc: Any) -> tuple:
    """Yinelenen grupta hangi kaydin tutulacagini belirler."""
    return (
        _KAYNAK_ONCELIGI.get(getattr(sonuc.satir, "kaynak_tip", ""), 9),
        0 if sonuc.masraf_merkezi else 1,
        0 if sonuc.durum == "OTOMATIK" else 1,
    )


def _paylar(sonuc: Any, harita: Any = None) -> list[tuple]:
    """Bir satirin paylarini dondurur.

    Her pay: ``(masraf_merkezi, merkez_adi, sirket, haritada_var, pay_notu, oran)``.

    Kaynak dosyada 'RHI 1/3 - RENSTROYDETAL 2/3' gibi bir paylasim yaziliysa
    tutar bolunur. Bu etiketler pratikte TUZEL KISI adidir, proje degil:
    dolayisiyla proje (masraf merkezi) ayni kalir, sadece sirket degisir.
    Etiket masraf merkezi haritasinda gercekten bir projeye karsilik geliyorsa
    o zaman proje de bolunur.
    """
    ek = sonuc.satir.ek if isinstance(sonuc.satir.ek, dict) else {}
    merkez = sonuc.masraf_merkezi or DAGITILAMAYAN
    ad = ek.get("masraf_merkezi_adi") or None
    sirket = sonuc.sirket or sonuc.sirket2
    haritada = bool(ek.get("masraf_merkezi_haritada", True)) and merkez != DAGITILAMAYAN

    paylasim = ek.get("paylasim") or []
    toplam_oran = sum(float(p.get("oran") or 0) for p in paylasim)
    if not paylasim or toplam_oran <= 0:
        return [(merkez, ad, sirket, haritada, None, 1.0)]

    paylar: list[tuple] = []
    for p in paylasim:
        oran = float(p.get("oran") or 0) / toplam_oran
        etiket = str(p.get("masraf_merkezi") or "").strip()
        not_metni = f"{etiket} {p.get('pay')}/{p.get('bolen')}".strip()
        cozum = harita.coz(etiket) if (harita is not None and etiket) else None
        if cozum:
            # Etiket gercek bir proje: masraf merkezi de bolunur.
            paylar.append((
                cozum["masraf_merkezi_kodu"], cozum["masraf_merkezi_adi"],
                cozum["sirket"] or sirket, True, not_metni, oran,
            ))
        else:
            # Etiket tuzel kisi adi: proje ayni, sirket degisiyor.
            paylar.append((merkez, ad, etiket or sirket, haritada, not_metni, oran))
    return paylar


def _artigi_dagit(satirlar: list[MahsupSatiri], hedef: float) -> None:
    """Yuvarlama artigini en buyuk satira ekleyerek toplami hedefe esitler.

    Oransal bolme kurus altinda kalan farklar birakir. Muhasebe bu farki kabul
    etmez: fatura kurusuna kadar kapanmalidir.
    """
    if not satirlar:
        return
    mevcut = round(sum(s.tutar for s in satirlar), 2)
    fark = round(hedef - mevcut, 2)
    if abs(fark) < _TOLERANS:
        return
    en_buyuk = max(satirlar, key=lambda s: abs(s.tutar))
    en_buyuk.tutar = round(en_buyuk.tutar + fark, 2)


def mahsuplasma_uret(
    sonuclar: Sequence[Any],
    harita: Any = None,
    yinelenenleri_ele: bool = True,
) -> MahsupTablosu:
    """Satir sonuclarindan mahsuplasma tablosunu uretir.

    Args:
        sonuclar: ``masraf_merkezi_coz`` ciktisi olan ``Sonuc`` listesi.
        harita: ``MasrafMerkeziHaritasi``. Paylasim etiketlerini proje mi sirket
            mi oldugunu anlamak icin kullanilir. Verilmezse paylasim etiketleri
            sirket kabul edilir.
        yinelenenleri_ele: Ayni islemi tasiyan ikinci dosyayi dagilima sokma.
            Ham cari dokum ile onun elle dagitilmis hali ayni mailde gelirse
            para cift sayilmasin diye varsayilan olarak aciktir.

    Returns:
        Mahsup satirlari, kaynak bazinda mutabakat, isaret celiskileri ve
        sayaclar. Her kontrol satiri kurusuna kadar kapanir.
    """
    tablo = MahsupTablosu()

    # 1) Kutuk satirlarini ve tutarsizlari ayikla.
    aday: list[Any] = []
    for s in sonuclar:
        if getattr(s.satir, "kaynak_tip", "") in _KUTUK_TIPLERI:
            tablo.kutuk_satir_sayisi += 1
            continue
        if s.satir.tutar is None:
            tablo.tutarsiz_satir_sayisi += 1
            continue
        aday.append(s)

    # 2) Okunan her sey once kontrol tablosuna yazilir. Eleme sonrasi degil
    #    ONCESI kaydedilir; boylece 'gelen' dosyada gercekten ne varsa odur.
    kontrol: dict[tuple[str, str], KontrolSatiri] = {}

    def _kontrol(sonuc: Any) -> KontrolSatiri:
        pb = sonuc.satir.para_birimi or "?"
        anahtar = (_kaynak_adi(sonuc), pb)
        k = kontrol.get(anahtar)
        if k is None:
            k = KontrolSatiri(kaynak=anahtar[0], para_birimi=pb,
                              gelen=0.0, dagitilan=0.0, dagitilamayan=0.0,
                              satir_sayisi=0)
            kontrol[anahtar] = k
        return k

    for s in aday:
        k = _kontrol(s)
        k.gelen += float(s.satir.tutar)
        k.satir_sayisi += 1

    # 3) Yinelenen islemleri ele. Kaba kova (tarih, tutar, doviz) icinde
    #    dosyalar 1:1 eslestirilir; kaynak onceligi dusuk olan (ham dokum)
    #    tutulur, esi kendi dosyasinin 'yinelenen' sutununa yazilir.
    if yinelenenleri_ele:
        gruplar: dict[tuple, list[Any]] = defaultdict(list)
        secilen: list[Any] = []
        for s_ in aday:
            anahtar = _yineleme_anahtari(s_)
            if anahtar is None:
                secilen.append(s_)
            else:
                gruplar[anahtar].append(s_)

        for anahtar, grup in gruplar.items():
            dosyalara: dict[str, list[Any]] = defaultdict(list)
            for g in grup:
                dosyalara[_kaynak_adi(g)].append(g)
            if len(dosyalara) == 1:
                # Ayni dosya icindeki tekrarlar gercek tekrardir: ayni ucus,
                # ayni ucret, farkli kisiler. Hepsi korunur.
                secilen.extend(grup)
                continue

            sirali = sorted(
                dosyalara.items(),
                key=lambda kv: (_oncelik(min(kv[1], key=_oncelik)), kv[0]),
            )
            tutulanlar = sorted(sirali[0][1], key=_oncelik)
            secilen.extend(tutulanlar)

            elenenler: list[Any] = []
            for _, digerleri in sirali[1:]:
                ciftler = _kovayi_esle(tutulanlar, digerleri)
                eslenen = {id(d) for _, d in ciftler}
                # Eslesemeyenler gercekten yeni islemdir, korunur.
                secilen.extend(d for d in digerleri if id(d) not in eslenen)
                for tutulan, elenen in ciftler:
                    _talimati_devral(tutulan, elenen)
                    k = _kontrol(elenen)
                    k.yinelenen_tutar += float(elenen.satir.tutar)
                    k.yinelenen_satir += 1
                    tablo.yinelenen_sayisi += 1
                    elenenler.append((tutulan, elenen))

            # Isaret celiskisi: ayni islem bir dosyada eksi, digerinde arti.
            for tutulan, elenen in elenenler:
                if (float(tutulan.satir.tutar) >= 0) != (float(elenen.satir.tutar) >= 0):
                    tablo.isaret_celiskileri.append(IsaretCeliskisi(
                        belge_tarihi=anahtar[0],
                        kisi=tutulan.satir.kisi_ham or elenen.satir.kisi_ham,
                        para_birimi=anahtar[2],
                        kaynaklar=(_kaynak_adi(tutulan), _kaynak_adi(elenen)),
                        tutarlar=(round(float(tutulan.satir.tutar), 2),
                                  round(float(elenen.satir.tutar), 2)),
                        kullanilan=round(float(tutulan.satir.tutar), 2),
                    ))
        aday = secilen

    # 4) Mahsup satirlarini biriktir.
    birikim: dict[tuple, MahsupSatiri] = {}
    for s in aday:
        kaynak = _kaynak_adi(s)
        pb = s.satir.para_birimi or "?"
        tip = s.satir.gider_tipi or "Diger"
        kimlik = s.eslesme.sicil or isim_normalize(s.satir.kisi_ham or "")
        for merkez, ad, sirket, haritada, pay_notu, oran in _paylar(s, harita):
            anahtar = (kaynak, tip, merkez, sirket, pb, pay_notu)
            kayit = birikim.get(anahtar)
            if kayit is None:
                kayit = MahsupSatiri(
                    kaynak=kaynak, gider_tipi=tip, masraf_merkezi=merkez,
                    masraf_merkezi_adi=ad, sirket=sirket, para_birimi=pb,
                    tutar=0.0, satir_sayisi=0, kisi_sayisi=0,
                    haritada_var=haritada, pay_notu=pay_notu,
                )
                kayit._kimlikler = set()   # type: ignore[attr-defined]
                birikim[anahtar] = kayit
            kayit.tutar += float(s.satir.tutar) * oran
            kayit.satir_sayisi += 1
            if s.durum == "OTOMATIK":
                kayit.otomatik += 1
            elif s.durum == "INCELE":
                kayit.incele += 1
            else:
                kayit.eslesmedi += 1
            if kimlik:
                kayit._kimlikler.add(kimlik)   # type: ignore[attr-defined]
            tarih = s.satir.belge_tarihi
            if tarih is not None:
                if kayit.ilk_tarih is None or tarih < kayit.ilk_tarih:
                    kayit.ilk_tarih = tarih
                if kayit.son_tarih is None or tarih > kayit.son_tarih:
                    kayit.son_tarih = tarih

    for kayit in birikim.values():
        kayit.kisi_sayisi = len(kayit._kimlikler)   # type: ignore[attr-defined]
        kayit.tutar = round(kayit.tutar, 2)

    # 5) Yuvarlama artigini dagit. Hedef, kontrol satirinin kendi olculerinden
    #    turetilir: okunan tutar eksi yinelenen tutar. Boylece mutabakat
    #    TANIM GEREGI kapanir; gercek bir kayip olsa bile yuvarlama farki
    #    olarak gizlenmez, cunku 'gelen' ve 'yinelenen' bagimsiz olculur.
    for k in kontrol.values():
        k.gelen = round(k.gelen, 2)
        k.yinelenen_tutar = round(k.yinelenen_tutar, 2)
    gruplu: dict[tuple[str, str], list[MahsupSatiri]] = defaultdict(list)
    for kayit in birikim.values():
        gruplu[(kayit.kaynak, kayit.para_birimi)].append(kayit)
    for anahtar, satirlar in gruplu.items():
        k = kontrol.get(anahtar)
        if k is None:
            continue
        _artigi_dagit(satirlar, round(k.gelen - k.yinelenen_tutar, 2))

    tablo.satirlar = sorted(
        birikim.values(),
        key=lambda m: (m.kaynak, -m.tutar, m.masraf_merkezi, m.sirket or ""),
    )

    # 6) Mutabakati kapat.
    for m in tablo.satirlar:
        k = kontrol.get((m.kaynak, m.para_birimi))
        if k is None:
            continue
        if m.dagitildi_mi:
            k.dagitilan += m.tutar
        else:
            k.dagitilamayan += m.tutar
    for k in kontrol.values():
        k.dagitilan = round(k.dagitilan, 2)
        k.dagitilamayan = round(k.dagitilamayan, 2)
    tablo.kontrol = sorted(kontrol.values(), key=lambda k: (-k.gelen, k.kaynak))
    return tablo
