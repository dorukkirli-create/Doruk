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
yazar: ham dokum 'ORNEKSOY AHMETCAN', elle dagitilmis hal
'AHMET CAN ORNEKSOY'. Harfler siralandiginda ikisi ayni anahtari verir.
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

from masraf.metin import isim_imzasi, isim_normalize

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
from masraf.envanter import DETAY_TIPLERI as _DETAY_TIPLERI, KUTUK_TIPLERI as _ENV_KUTUK
#: Dagilima girmeyen satir tipleri: kisi kutukleri + fatura detay listeleri
#: (tutar yansitma dosyasindadir; capraz kontrol edilir). Tek kaynak: envanter.
_KUTUK_TIPLERI = _ENV_KUTUK | _DETAY_TIPLERI
#: Detay listesinin karsiligi olan, tutar tasiyan kaynak tipi.
_DETAY_KARSILIGI = "energo_assessment"

#: Kurus altinda kalan farklar kapali sayilir.
_TOLERANS = 0.005
#: Bundan buyuk bir fark yuvarlama artigi olamaz; gomulmez, acik birakilir.
_AZAMI_YUVARLAMA = 0.50


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
    #: Elle dagitilmis dosyada insanin yazdigi sirket/santiye etiketleri
    #: (birden fazlaysa '; ' ile). Muhasebe insan kararini gorsun diye.
    elle_etiket: str | None = None
    #: Bu satira giren kalemlerin evrak / fatura numaralari (ilk birkaci).
    evrak_no: str | None = None

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

    def kisa_aciklama(self) -> str:
        """Kisi adi tasimayan ozet; kapak sayfasi gibi yonetici gorunumleri icin."""
        tarih = f"{self.belge_tarihi:%d.%m.%Y}" if self.belge_tarihi else "tarihsiz"
        tutar = max(abs(t) for t in self.tutarlar) if self.tutarlar else 0.0
        return (
            f"{tarih} tarihli {tutar:,.2f} {self.para_birimi} kalem iki dosyada zit "
            f"isaretli ({' / '.join(self.kaynaklar)}); ham dokumdeki deger kullanildi. "
            "Ayrinti Kontrol sayfasinda."
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
    #: Tutari okunamayan satirlar. Bunlar 'gelen'e giremez ama YOK sayilamaz:
    #: fatura toplami onlari icerir. Sifirdan buyukse mutabakat acik kalir.
    tutarsiz_satir: int = 0
    #: Kaynak dosyanin KENDI beyan ettigi toplam (varsa). Okunanla farki,
    #: okuyucunun kacirdigi tutari ele verir.
    beyan_toplam: float | None = None
    #: Yineleme suphesi: elenmedi ama baska dosyadaki bir kalemle ayni kisi/
    #: tarih ve farkli tutar, ya da tek ortak kelimeyle baglanmis cift.
    suphe_satir: int = 0
    suphe_tutar: float = 0.0
    #: Belge tarihi okunamayan satirlar: yineleme kontrolune giremezler.
    tarihsiz_satir: int = 0

    @property
    def para_birimi_yok(self) -> bool:
        return self.para_birimi == "?"

    @property
    def beyan_farki(self) -> float | None:
        """Faturanin beyan ettigi toplam ile okunan arasindaki fark."""
        if self.beyan_toplam is None:
            return None
        return round(self.beyan_toplam - self.gelen, 2)

    @property
    def beyan_toleransi(self) -> float:
        """Tedarikci yuvarlamasindan gelebilecek en buyuk fark: satir basina 1 kurus, en cok 50."""
        return min(_AZAMI_YUVARLAMA, 0.01 * max(1, self.satir_sayisi))

    @property
    def beyan_uyusuyor(self) -> bool:
        bf = self.beyan_farki
        return bf is None or abs(bf) <= self.beyan_toleransi + 1e-9

    @property
    def fark(self) -> float:
        """Sifir olmalidir. Degilse dagitimda kayip var demektir."""
        return round(
            self.gelen - (self.dagitilan + self.dagitilamayan + self.yinelenen_tutar), 2
        )

    @property
    def kapali_mi(self) -> bool:
        """Kapali: dagitim toplami tutuyor, tutarsiz satir yok, beyanla fark yok."""
        if abs(self.fark) >= 0.01:
            return False
        if self.tutarsiz_satir or self.para_birimi_yok:
            return False
        return self.beyan_uyusuyor

    @property
    def acik_sebebi(self) -> str:
        """Kapanmadiysa neden; kapaliysa bos."""
        sebepler = []
        if abs(self.fark) >= 0.01:
            sebepler.append(f"dagitim farki {self.fark:+.2f}")
        if self.tutarsiz_satir:
            sebepler.append(f"{self.tutarsiz_satir} satirda tutar okunamadi")
        if self.para_birimi_yok:
            sebepler.append("para birimi okunamadi")
        bf = self.beyan_farki
        if not self.beyan_uyusuyor:
            sebepler.append(f"faturanin beyan ettigi toplamdan {bf:+.2f} farkli")
        return "; ".join(sebepler)

    @property
    def suphe_notu(self) -> str:
        """Kapali ama dikkat: yineleme suphesi ya da tarihsiz satir varsa."""
        parcalar = []
        if self.suphe_satir:
            parcalar.append(f"YINELEME SUPHESI: {self.suphe_satir} satir, {self.suphe_tutar:,.2f}")
        if self.tarihsiz_satir:
            parcalar.append(f"{self.tarihsiz_satir} satir tarihsiz, yineleme kontrolune girmedi")
        return "; ".join(parcalar)

    @property
    def net(self) -> float:
        """Bu dosyanin gercekten kattigi tutar (yinelenenler dusulmus)."""
        return round(self.dagitilan + self.dagitilamayan, 2)

    @property
    def dagitim_orani(self) -> float:
        return (self.dagitilan / self.net * 100.0) if self.net else 0.0


@dataclass
class YinelemeSuphesi:
    """Elenmeyen ama yinelenen olabilecek kalem; insan bakmali."""

    kaynak: str
    kisi: str | None
    belge_tarihi: date | None
    tutar: float
    para_birimi: str
    karsi_kaynak: str
    karsi_tutar: float | None
    sebep: str

    def aciklama(self) -> str:
        tarih = f"{self.belge_tarihi:%d.%m.%Y}" if self.belge_tarihi else "tarihsiz"
        karsi = f"{self.karsi_tutar:,.2f}" if self.karsi_tutar is not None else "-"
        return (f"{self.kaynak}: {self.kisi or '(kisi yok)'} / {tarih} / {self.tutar:,.2f} {self.para_birimi} "
                f"<-> {self.karsi_kaynak}: {karsi}. {self.sebep}")


@dataclass
class DetayKontrolu:
    """Fatura detay listesindeki kisilerin yansitma satirlariyla capraz kontrolu.

    Tedarikci her fatura icin tutarsiz bir katilimci listesi gonderir; tutar
    yansitma dosyasindadir. Iki liste ayni kisileri tasimali. Tasimiyorsa ya
    yansitmada bir kisi eksik (para eksik dagitilir) ya da fazla (fatura
    kapsaminda olmayan kisiye pay yazilir). Ikisi de muhasebe hatasidir.
    """

    fatura_no: str
    kaynak: str
    detay_kisi: int
    yansitma_kisi: int
    eksik: list[str] = field(default_factory=list)   # detayda var, yansitmada yok
    fazla: list[str] = field(default_factory=list)   # yansitmada var, detayda yok

    @property
    def tutarli_mi(self) -> bool:
        return not self.eksik and not self.fazla

    @property
    def eslesen(self) -> int:
        return self.detay_kisi - len(self.eksik)

    def aciklama(self) -> str:
        if self.tutarli_mi:
            return f"{self.fatura_no}: detay listesindeki {self.detay_kisi} kisi yansitmayla birebir"
        parcalar = [f"{self.fatura_no}: fatura detay listesi ile yansitma UYUSMUYOR"]
        if self.yansitma_kisi == 0:
            parcalar.append("yansitma dosyasinda bu fatura numarasi hic yok")
        if self.eksik:
            parcalar.append("yansitmada eksik: " + ", ".join(self.eksik))
        if self.fazla:
            parcalar.append("detay listesinde olmayan: " + ", ".join(self.fazla))
        return "; ".join(parcalar)


@dataclass
class MahsupTablosu:
    """Mahsuplasma ciktisinin tamami."""

    satirlar: list[MahsupSatiri] = field(default_factory=list)
    kontrol: list[KontrolSatiri] = field(default_factory=list)
    isaret_celiskileri: list[IsaretCeliskisi] = field(default_factory=list)
    detay_kontrolleri: list[DetayKontrolu] = field(default_factory=list)
    supheler: list[YinelemeSuphesi] = field(default_factory=list)
    #: Elle dagitilmis dosyada insanin yazdigi sirket ile tablonun sirketi
    #: farkli olan satir sayisi ve tutari (para birimi bazinda).
    sirket_uyusmazligi: dict = field(default_factory=dict)
    uyarilar: list[str] = field(default_factory=list)
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
            kayit.setdefault("_sirketler", set())
            if s.sirket:
                kayit["_sirketler"].add(s.sirket)
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
            toplamlar[kayit["para_birimi"]] += abs(kayit["tutar"])
        cikti = []
        for kayit in birikim.values():
            genel = toplamlar[kayit["para_birimi"]]
            kayit["tutar"] = round(kayit["tutar"], 2)
            kayit["kisi_sayisi"] = len(kayit.pop("_kisiler"))
            sirketler = sorted(kayit.pop("_sirketler", set()))
            if len(sirketler) > 1:
                kayit["sirket"] = "; ".join(sirketler)
            elif sirketler:
                kayit["sirket"] = sirketler[0]
            # Pay mutlak tutarlara gore: iade agirlikli merkezde %150 / %-50
            # gibi yaniltici yuzdeler cikmasin.
            kayit["pay_yuzde"] = round(abs(kayit["tutar"]) / genel * 100, 2) if genel else 0.0
            cikti.append(kayit)
        return sorted(cikti, key=lambda k: (-k["tutar"], k["masraf_merkezi"]))

    def sirket_ozeti(self) -> list[dict]:
        """Tuzel kisi (1C 'Firm 2') bazinda toplam, altinda projeleriyle.

        Muhasebenin okuma sirasi budur: once hangi sirkete, sonra o sirketin
        hangi projesine. Donen her oge bir sirkettir ve ``projeler`` anahtari
        altinda kendi masraf merkezlerini tasir.
        """
        gruplar: dict[tuple[str, str], dict] = {}
        for s in self.satirlar:
            # Masraf merkezi cozulemeyen satirlarin sirketi de yoktur; onlari
            # "(sirket yok)" yerine ne olduklariyla adlandirmak daha durust.
            sirket = s.sirket or ("(dagitilamayan)" if s.masraf_merkezi == DAGITILAMAYAN else "(sirket yok)")
            anahtar = (sirket, s.para_birimi)
            grup = gruplar.setdefault(anahtar, {
                "sirket": sirket, "para_birimi": s.para_birimi,
                "tutar": 0.0, "satir_sayisi": 0,
                "_kisiler": set(), "_projeler": {},
            })
            grup["tutar"] += s.tutar
            grup["satir_sayisi"] += s.satir_sayisi
            grup["_kisiler"] |= getattr(s, "_kimlikler", set())
            proje = grup["_projeler"].setdefault(s.masraf_merkezi, {
                "masraf_merkezi": s.masraf_merkezi,
                "masraf_merkezi_adi": s.masraf_merkezi_adi,
                "haritada_var": s.haritada_var,
                "tutar": 0.0, "satir_sayisi": 0, "_kisiler": set(),
            })
            proje["tutar"] += s.tutar
            proje["satir_sayisi"] += s.satir_sayisi
            proje["_kisiler"] |= getattr(s, "_kimlikler", set())

        toplam_pb: dict[str, float] = defaultdict(float)
        for grup in gruplar.values():
            toplam_pb[grup["para_birimi"]] += grup["tutar"]

        cikti: list[dict] = []
        for grup in gruplar.values():
            genel = toplam_pb[grup["para_birimi"]]
            grup["tutar"] = round(grup["tutar"], 2)
            grup["kisi_sayisi"] = len(grup.pop("_kisiler"))
            grup["pay_yuzde"] = round(grup["tutar"] / genel * 100, 2) if genel else 0.0
            projeler = []
            for proje in grup.pop("_projeler").values():
                proje["tutar"] = round(proje["tutar"], 2)
                proje["kisi_sayisi"] = len(proje.pop("_kisiler"))
                proje["pay_yuzde"] = (
                    round(proje["tutar"] / grup["tutar"] * 100, 2) if grup["tutar"] else 0.0
                )
                projeler.append(proje)
            grup["projeler"] = sorted(projeler, key=lambda k: -k["tutar"])
            cikti.append(grup)
        return sorted(cikti, key=lambda k: (-k["tutar"], k["sirket"]))

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


def _detay_kontrolu(detaylar: list[Any], aday: list[Any]) -> list[DetayKontrolu]:
    """Detay listesi kisilerini ayni fatura numarali yansitma kisileriyle kiyaslar."""
    if not detaylar:
        return []
    yansitma: dict[str, dict[str, str]] = defaultdict(dict)   # fatura_no -> imza -> ad
    for s in aday:
        if getattr(s.satir, "kaynak_tip", "") != _DETAY_KARSILIGI:
            continue
        ek = s.satir.ek if isinstance(s.satir.ek, dict) else {}
        no = str(ek.get("fatura_no") or "").strip()
        imza = _isim_imzasi(s.satir.kisi_ham)
        if no and imza:
            yansitma[no][imza] = s.satir.kisi_ham or ""
    gruplar: dict[tuple[str, str], dict[str, str]] = defaultdict(dict)
    for s in detaylar:
        ek = s.satir.ek if isinstance(s.satir.ek, dict) else {}
        no = str(ek.get("fatura_no") or "").strip() or "(fatura no yok)"
        imza = _isim_imzasi(s.satir.kisi_ham)
        if imza:
            gruplar[(no, _kaynak_adi(s))][imza] = s.satir.kisi_ham or ""
    cikti: list[DetayKontrolu] = []
    for (no, kaynak), kisiler in sorted(gruplar.items()):
        karsi = yansitma.get(no, {})
        cikti.append(DetayKontrolu(
            fatura_no=no, kaynak=kaynak,
            detay_kisi=len(kisiler), yansitma_kisi=len(karsi),
            eksik=sorted(ad for imza, ad in kisiler.items() if imza not in karsi),
            fazla=sorted(ad for imza, ad in karsi.items() if imza not in kisiler),
        ))
    return cikti


def _kismi_yineleme_suphesi(aday: list[Any], tablo: Any, kontrol: dict) -> None:
    """Elenmeyen ama yinelenen olabilecek kalemleri bulur.

    Iki dosya arasinda en az bir eleme olduysa, elenmeyen kalanlardan ayni
    kisi + ayni tarih ile karsi dosyada FARKLI tutarli bir kayit bulunanlar
    'yineleme suphesi'dir (elle dagitilmis dosyada tutar duzeltilmis olabilir;
    olculdu: 1.982,11 USD cift sayiliyordu, uyari yoktu). Tarihsiz kalanlar
    icin (|tutar|, doviz, isim imzasi) zayif kovasi denenir. Hicbir sey
    ELENMEZ; yalnizca raporlanir ve kontrol satirina yazilir.
    """
    yinelenen_var = {k.kaynak for k in kontrol.values() if k.yinelenen_satir}
    if not yinelenen_var:
        return
    # (kaynak, imza, tarih) -> kayitlar
    indeks: dict[tuple[str, str, Any], list[Any]] = defaultdict(list)
    for s in aday:
        imza = _isim_imzasi(s.satir.kisi_ham)
        if imza and s.satir.tutar is not None:
            indeks[(_kaynak_adi(s), imza, s.satir.belge_tarihi)].append(s)
    kaynaklar = sorted({k[0] for k in indeks})
    gorulen: set[int] = set()
    for s in aday:
        kaynak = _kaynak_adi(s)
        if kaynak not in yinelenen_var or s.satir.tutar is None:
            continue
        imza = _isim_imzasi(s.satir.kisi_ham)
        if not imza:
            continue
        for karsi in kaynaklar:
            if karsi == kaynak:
                continue
            for r in indeks.get((karsi, imza, s.satir.belge_tarihi), []):
                if round(abs(float(r.satir.tutar)), 2) == round(abs(float(s.satir.tutar)), 2):
                    continue   # ayni tutar: zaten kova elemesinde degerlendirildi
                if id(s) in gorulen:
                    break
                gorulen.add(id(s))
                tablo.supheler.append(YinelemeSuphesi(
                    kaynak=kaynak, kisi=s.satir.kisi_ham, belge_tarihi=s.satir.belge_tarihi,
                    tutar=round(float(s.satir.tutar), 2), para_birimi=_pb(s.satir),
                    karsi_kaynak=karsi, karsi_tutar=round(float(r.satir.tutar), 2),
                    sebep="Ayni kisi ve tarih, tutar farkli; iki dosyada da dagitima girdi. "
                          "Tutari duzeltilmis ayni kalem olabilir."))
                k = kontrol.get((kaynak, _pb(s.satir)))
                if k is not None:
                    k.suphe_satir += 1
                    k.suphe_tutar = round(k.suphe_tutar + float(s.satir.tutar), 2)
    # Tarihsiz kalanlar: zayif kova.
    tarihsiz = [s for s in aday if not s.satir.belge_tarihi and s.satir.tutar is not None]
    if tarihsiz:
        zayif: dict[tuple, list[Any]] = defaultdict(list)
        for s in tarihsiz:
            imza = _isim_imzasi(s.satir.kisi_ham)
            if imza:
                zayif[(round(abs(float(s.satir.tutar)), 2), _pb(s.satir), imza)].append(s)
        for grup in zayif.values():
            dosyalar = {_kaynak_adi(g) for g in grup}
            if len(dosyalar) < 2:
                continue
            for g in grup[1:]:
                tablo.supheler.append(YinelemeSuphesi(
                    kaynak=_kaynak_adi(g), kisi=g.satir.kisi_ham, belge_tarihi=None,
                    tutar=round(float(g.satir.tutar), 2), para_birimi=_pb(g.satir),
                    karsi_kaynak=_kaynak_adi(grup[0]), karsi_tutar=round(float(grup[0].satir.tutar), 2),
                    sebep="Tarihsiz; ayni kisi ve tutar baska dosyada da var, yineleme olabilir."))
                k = kontrol.get((_kaynak_adi(g), _pb(g.satir)))
                if k is not None:
                    k.suphe_satir += 1
                    k.suphe_tutar = round(k.suphe_tutar + float(g.satir.tutar), 2)
    # Dosya bazinda yinelenen orani %50-%99: kalanlar cift sayilmis olabilir.
    for k in kontrol.values():
        if k.satir_sayisi and 0.5 <= k.yinelenen_satir / k.satir_sayisi < 1.0:
            kalan = k.satir_sayisi - k.yinelenen_satir
            tablo.uyarilar.append(
                f"YINELEME SUPHESI: {k.kaynak} dosyasinin {k.yinelenen_satir}/{k.satir_sayisi} satiri "
                f"baska dosyayla yinelenen cikti, {kalan} satir tek basina kaldi. Bu satirlar iki "
                "dosyada da dagitima girdi; ayni kalemin duzeltilmis hali olabilir. Kontrol sayfasindaki "
                "'Yineleme suphesi' listesine bakin.")
    for sp in tablo.supheler:
        if sp.sebep.startswith("Ayni gun, ayni tutar ama"):
            k = kontrol.get((sp.kaynak, sp.para_birimi))
            if k is not None:
                k.suphe_satir += 1
                k.suphe_tutar = round(k.suphe_tutar + sp.tutar, 2)


def _kaynak_adi(sonuc: Any) -> str:
    """Satirin ait oldugu faturayi adlandirir.

    Outlook mesajindan gelen satirlarda kaynak 'mesaj.msg > ek.xlsx' bicimindedir;
    fatura o ekin kendisidir.
    """
    ek = sonuc.satir.ek if isinstance(getattr(sonuc.satir, "ek", None), dict) else {}
    etiket = ek.get("kaynak_etiketi")
    if etiket:
        # Ayni adli ama farkli icerikli iki dosya geldiginde boru bu etiketi
        # koyar; iki dosya kontrol tablosunda ayri satir olur, birlesmez.
        return str(etiket)
    ham = str(getattr(sonuc.satir, "kaynak_dosya", "") or "")
    return ham.split("> ")[-1].strip() or "(bilinmeyen)"


def _isim_imzasi(ham: str | None) -> str:
    """Isim sirasindan ve bosluklardan bagimsiz karsilastirma imzasi.

    'ORNEKSOY AHMETCAN' ve 'AHMET CAN ORNEKSOY' ayni imzayi verir; iki
    dosya ayni kisiyi farkli sirayla ve farkli bitisiklikte yazdigi icin
    gereklidir. Tanim ``masraf.metin`` icindedir; harici kisiler defteri de
    ayni imzayi kullanir, ikisi ayrisamasin diye tek kaynaktan gelir.
    """
    return isim_imzasi(ham or "")


def _yineleme_anahtari(sonuc: Any) -> tuple | None:
    """Ayni islemi iki farkli dosyada tanimak icin KABA anahtar.

    (belge tarihi, mutlak tutar, para birimi). Isim BILEREK disarida birakilir:
    iki dosya ayni kisiyi farkli yazar, hatta kirpar ('ORNEKSOY AHMETCA'),
    hatta yanlis yazar ('ORNEKKAYA ANIL' / 'ALI YALCINKAYA'). Isim bu kaba
    kova icinde ESLESTIRICI olarak kullanilir, anahtar olarak degil.

    Mutlak tutar kullanilir: bir dosya iadeyi eksi, digeri arti yazabilir.
    Tarih veya tutar yoksa yineleme tespiti yapilmaz (None doner) ve satir
    oldugu gibi korunur.
    """
    satir = sonuc.satir
    if satir.tutar is None or not satir.belge_tarihi:
        return None
    return (satir.belge_tarihi, round(abs(float(satir.tutar)), 2), _pb(satir))


def _pb(satir: Any) -> str:
    """Para birimini normalize eder: 'usd ' ve 'USD' ayni kovadir; yoksa '?'."""
    return (str(getattr(satir, "para_birimi", "") or "")).strip().upper() or "?"


#: Bilinen 'ham dokum <-> elle dagitilmis' ciftleri. Bu ciftlerde tek ortak
#: kelime (puan 1) bile ayni islemdir; cunku iki dosya AYNI islemleri tasir.
_GEVSEK_CIFTLER: frozenset[frozenset[str]] = frozenset({
    frozenset({"antik_cari", "yuzyil_dagitilmis"}),
})


def _gevsek_cift(a: Any, b: Any) -> bool:
    """Iki kayit gevsek (puan 1) eslemeye izin verilen bir dosya ciftinden mi?

    Ham dokum ile onun elle dagitilmis hali ya da AYNI mailden gelen iki ek.
    Bagimsiz iki tedarikci dosyasinda tek ortak kelime kanit degildir: ayni
    gun ayni sabit ucretli (vize, saglik) baska kisiler elenirdi (olculdu).
    """
    ta = getattr(a.satir, "kaynak_tip", "") or ""
    tb = getattr(b.satir, "kaynak_tip", "") or ""
    if frozenset({ta, tb}) in _GEVSEK_CIFTLER:
        return True
    ka = str(getattr(a.satir, "kaynak_dosya", "") or "")
    kb = str(getattr(b.satir, "kaynak_dosya", "") or "")
    if "> " in ka and "> " in kb and ka.split("> ")[0] == kb.split("> ")[0]:
        return True
    return False


def _eslesme_puani(a: Any, b: Any) -> int:
    """Ayni kovadaki iki kaydin ayni kisi olma gucu. Buyuk = daha guclu.

    3: isim imzalari birebir ayni.
    2: biri digerinin kirpilmis hali ('ORNEKSOY AHMETCA' <- 'ORNEKSOY AHMETCAN').
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


def _kovayi_esle(tutulanlar: list[Any], digerleri: list[Any],
                 zayif_ciftler: list[tuple[Any, Any]] | None = None) -> list[tuple[Any, Any]]:
    """Ayni kovadaki kayitlari 1:1 eslestirir; eslesemeyenler disarida kalir.

    Once en guclu isim eslesmeleri baglanir, sonra kalanlar sirayla. Kova
    zaten (tarih, tutar, doviz) esitligini garanti ettigi icin isimsiz veya
    yanlis yazilmis kayitlar da dogru sekilde eslesir.

    Puan 1 (tek ortak kelime) yalnizca gevsek dosya ciftlerinde (ham dokum
    <-> elle dagitilmis, ayni mailin ekleri) kabul edilir ve ``zayif_ciftler``
    listesine yazilir; finans bunlari 'yineleme suphesi' olarak gorur.

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
                if _eslesme_puani(t, d) >= esik and (esik > 1 or _gevsek_cift(t, d)):
                    en_iyi = t
                    break
            if en_iyi is None:
                kalan_bekleyen.append(d)
            else:
                bos_tutulan.remove(en_iyi)
                ciftler.append((en_iyi, d))
                if esik == 1 and zayif_ciftler is not None:
                    zayif_ciftler.append((en_iyi, d))
        bekleyen = kalan_bekleyen
    # Kalanlar: isim eslesmedi ama kova (tarih, tutar, doviz) ayni. Yalnizca
    # iki taraf da isimsiz kurumsal kalemse (cenaze celengi, organizasyon)
    # sirayla baglanir. Isimli kayitlar isim benzerligi olmadan BAGLANMAZ:
    # iki faturada ayni gun ayni sabit ucretle (vize, saglik, bagaj) gecen
    # bambaska kisiler 'yinelenen' diye elenirse para dagilimdan duser ve
    # mutabakat yine kapali gorunur (olculdu). Sayilarin esit olmasi kanit
    # degildir; bu kayitlar korunur ve 'yineleme suphesi' olarak raporlanir.
    if bekleyen and bos_tutulan:
        # Isimsiz kalemler kendi aralarinda sirayla baglanir; isimli kalanlar
        # isimsizlere ya da birbirine BAGLANMAZ. (Onceki kural iki tarafin
        # TAMAMI isimsiz olmadikca hic baglamiyordu; isimsiz cift bile
        # elenmiyor, 50 USD iki kez dagitiliyordu, olculdu.)
        isimsiz_bekleyen = [d for d in bekleyen if not _isim_imzasi(d.satir.kisi_ham)]
        isimsiz_tutulan = [t for t in bos_tutulan if not _isim_imzasi(t.satir.kisi_ham)]
        for d, t in zip(isimsiz_bekleyen, isimsiz_tutulan):
            ciftler.append((t, d))
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
    # Devralinan talimat AYRI anahtarda tutulur; satirin kendi 'paylasim'i
    # degismez. Boylece ayni sonuclarla yinelenenleri_ele=False calistirmak
    # ham satiri paylasimli gostermez (olculdu).
    if not tutulan.satir.ek.get("paylasim") and elenen.satir.ek.get("paylasim"):
        tutulan.satir.ek["paylasim_devralindi"] = elenen.satir.ek["paylasim"]
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


def _paylar(sonuc: Any, harita: Any = None, devralinan_kullan: bool = True,
            uyarilar: list[str] | None = None) -> list[tuple]:
    """Bir satirin paylarini dondurur.

    Her pay: ``(masraf_merkezi, merkez_adi, sirket, haritada_var, pay_notu, oran)``.

    Kurallar (olculdu, hepsi para kaybi ya da yanlis sirket uretiyordu):
    * (DAGITILAMAYAN) satir bolunmez; sirket tasimaz.
    * Oranlar toplami 1'den kucukse kalan pay satirin KENDI sirketine
      '(kalan)' notuyla yazilir; 'RHI 1/2' tutarin tamamini RHI'ye vermez.
    * Toplam 1'den buyukse paylasim uygulanmaz ve uyari yazilir.
    * Ne tuzel kisi ne haritada proje olan etiket ('BILET 2/2 KISI') pay
      sayilmaz.

    Kaynak dosyada 'RHI 1/3 - RENSTROYDETAL 2/3' gibi bir paylasim yaziliysa
    tutar bolunur. Bu etiketler pratikte TUZEL KISI adidir, proje degil:
    dolayisiyla proje (masraf merkezi) ayni kalir, sadece sirket degisir.
    Etiket masraf merkezi haritasinda gercekten bir projeye karsilik geliyorsa
    o zaman proje de bolunur.
    """
    ek = sonuc.satir.ek if isinstance(sonuc.satir.ek, dict) else {}
    # ESLESMEDI: kimlik kabul esiginin altinda. Zayif bir tahmine dayanarak
    # tutari bir projeye yazmak yanlis mahsuplasmadir; tutar gorunur bicimde
    # dagitilamayan kalir, insan karar verir.
    if getattr(sonuc, "durum", None) == "ESLESMEDI":
        merkez = DAGITILAMAYAN
    else:
        merkez = sonuc.masraf_merkezi or DAGITILAMAYAN
    if merkez == DAGITILAMAYAN:
        # Reddedilen zayif adayin sirketi/projesi tasinmaz; aksi halde
        # 'kisi bulunamadi' denen tutar Sirket Kirilimi'nde o adayin tuzel
        # kisisine yazilir (olculdu: 527,83 USD UST LUGA/RHI altina gitmisti).
        ad, sirket = None, None
    else:
        ad = ek.get("masraf_merkezi_adi") or None
        sirket = sonuc.sirket or sonuc.sirket2
    haritada = bool(ek.get("masraf_merkezi_haritada", True)) and merkez != DAGITILAMAYAN

    paylasim = ek.get("paylasim") or (ek.get("paylasim_devralindi") if devralinan_kullan else None) or []
    if merkez == DAGITILAMAYAN or not paylasim:
        return [(merkez, ad, sirket, haritada, None, 1.0)]

    from masraf.masraf_merkezi import sirket_kanonik
    kaynak_adi = _kaynak_adi(sonuc)
    gecerli: list[tuple] = []     # (oran, etiket, not, cozum, kanonik)
    for p in paylasim:
        oran = float(p.get("oran") or 0)
        etiket = str(p.get("masraf_merkezi") or "").strip()
        not_metni = f"{etiket} {p.get('pay')}/{p.get('bolen')}".strip()
        cozum = harita.coz(etiket) if (harita is not None and etiket) else None
        kanonik = sirket_kanonik(etiket) if etiket else None
        if harita is None:
            # Harita yoksa etiketin proje mi tuzel kisi mi oldugu bilinemez;
            # eski davranis korunur: etiket sirket adi sayilir.
            tuzel = bool(etiket)
        else:
            tuzel_mi = getattr(harita, "tuzel_kisi_mi", None)
            tuzel = bool(kanonik) or bool(etiket and (tuzel_mi(etiket) if callable(tuzel_mi) else True))
        if oran <= 0 or (not cozum and not tuzel):
            if uyarilar is not None and etiket:
                uyarilar.append(
                    f"{kaynak_adi} satir {sonuc.satir.satir_no}: paylasim etiketi '{etiket}' ne tuzel "
                    "kisi ne haritada proje; pay sayilmadi.")
            continue
        gecerli.append((oran, etiket, not_metni, cozum, kanonik))
    if not gecerli:
        return [(merkez, ad, sirket, haritada, None, 1.0)]
    toplam_oran = sum(g[0] for g in gecerli)
    if toplam_oran > 1.0 + 1e-6:
        if uyarilar is not None:
            uyarilar.append(
                f"{kaynak_adi} satir {sonuc.satir.satir_no}: paylasim oranlari toplami "
                f"{toplam_oran:.2f} > 1 ({', '.join(g[2] for g in gecerli)}); paylasim uygulanmadi, "
                "satir kendi sirketinde kaldi. Kaynak dosyadaki notu duzeltin.")
        return [(merkez, ad, sirket, haritada, "PAYLASIM HATALI: " + ", ".join(g[2] for g in gecerli), 1.0)]

    paylar: list[tuple] = []
    for oran, etiket, not_metni, cozum, kanonik in gecerli:
        if cozum:
            # Etiket gercek bir proje: masraf merkezi de bolunur.
            paylar.append((
                cozum["masraf_merkezi_kodu"], cozum["masraf_merkezi_adi"],
                cozum["sirket"] or sirket, True, not_metni, oran,
            ))
        else:
            # Etiket tuzel kisi adi: proje ayni, sirket degisiyor. Etiket
            # haritanin sirket koduna cevrilir (RENSTROYDETAL -> RSS); aksi
            # halde Sirket Kirilimi'nde ayni tuzel kisi iki satir olur.
            paylar.append((merkez, ad, kanonik or etiket or sirket, haritada, not_metni, oran))
    kalan = round(1.0 - toplam_oran, 6)
    if kalan > 1e-6:
        # 'RHI 1/2' yalnizca yarisini RHI'ye verir; kalan yari satirin kendi
        # sirketinde kalir ve bu acikca yazilir.
        paylar.append((merkez, ad, sirket, haritada, f"(kalan {kalan:.2f})", kalan))
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
    # Yuvarlama artigi kurus mertebesindedir. Daha buyuk bir fark yuvarlama
    # degil hatadir; onu en buyuk satira gomersek mutabakat sahte kapanir.
    # Dokunma, kontrol satiri acik kalsin ve gorunsun.
    if abs(fark) > _AZAMI_YUVARLAMA:
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
    tutarsizlar: list[Any] = []
    detaylar: list[Any] = []
    for s in sonuclar:
        tip = getattr(s.satir, "kaynak_tip", "")
        if tip in _KUTUK_TIPLERI:
            tablo.kutuk_satir_sayisi += 1
            if tip == "energo_assessment_detay":
                detaylar.append(s)
            continue
        if s.satir.tutar is None:
            tablo.tutarsiz_satir_sayisi += 1
            tutarsizlar.append(s)
            continue
        aday.append(s)

    # 1b) Fatura detay listeleri (tutarsiz) yansitma satirlariyla capraz
    #     kontrol edilir. Kisi kumesi ayni degilse uyari uretilir.
    tablo.detay_kontrolleri = _detay_kontrolu(detaylar, list(aday) + tutarsizlar)
    for dk in tablo.detay_kontrolleri:
        if not dk.tutarli_mi:
            tablo.uyarilar.append("FATURA DETAYI: " + dk.aciklama())

    # 2) Okunan her sey once kontrol tablosuna yazilir. Eleme sonrasi degil
    #    ONCESI kaydedilir; boylece 'gelen' dosyada gercekten ne varsa odur.
    kontrol: dict[tuple[str, str], KontrolSatiri] = {}

    def _kontrol(sonuc: Any) -> KontrolSatiri:
        pb = _pb(sonuc.satir)
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
        if not s.satir.belge_tarihi:
            k.tarihsiz_satir += 1
    # Tutari okunamayan satirlar kendi dosyalarinin kontrol satirinda
    # SAYILIR; boylece 'mutabakat kapali' derken 76 dolar sessizce kaybolmaz.
    for s in tutarsizlar:
        _kontrol(s).tutarsiz_satir += 1
    # Kaynak dosya kendi toplamini beyan ediyorsa (energo okuyucusu fatura
    # detayindaki 'Genel Toplam'i ek['fatura_ozeti'] icinde tasir) onu al.
    for s in list(aday) + tutarsizlar:
        ek = s.satir.ek if isinstance(s.satir.ek, dict) else {}
        ozet = ek.get("fatura_ozeti")
        if isinstance(ozet, dict) and ozet:
            k = _kontrol(s)
            if k.beyan_toplam is None:
                try:
                    k.beyan_toplam = round(sum(float(v) for v in ozet.values()), 2)
                except (TypeError, ValueError):
                    pass

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
                zayif: list[tuple[Any, Any]] = []
                ciftler = _kovayi_esle(tutulanlar, digerleri, zayif)
                for t, d in zayif:
                    tablo.supheler.append(YinelemeSuphesi(
                        kaynak=_kaynak_adi(d), kisi=d.satir.kisi_ham, belge_tarihi=d.satir.belge_tarihi,
                        tutar=round(float(d.satir.tutar), 2), para_birimi=_pb(d.satir),
                        karsi_kaynak=_kaynak_adi(t), karsi_tutar=round(float(t.satir.tutar), 2),
                        sebep="Ayni gun, ayni tutar ama isimlerin yalnizca bir kelimesi ortak; yinelenen "
                              "sayilip elendi. Farkli kisiyse elle geri ekleyin."))
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
        _kismi_yineleme_suphesi(aday, tablo, kontrol)

    # 4) Mahsup satirlarini biriktir.
    birikim: dict[tuple, MahsupSatiri] = {}
    from masraf.masraf_merkezi import sirket_kanonik
    uyusmazlik: dict[str, dict] = defaultdict(
        lambda: {"satir": 0, "tutar": 0.0, "ciftler": defaultdict(lambda: {"satir": 0, "tutar": 0.0})})
    for s in aday:
        kaynak = _kaynak_adi(s)
        pb = _pb(s.satir)
        tip = s.satir.gider_tipi or "Diger"
        kimlik = s.eslesme.sicil or isim_normalize(s.satir.kisi_ham or "")
        ek = s.satir.ek if isinstance(s.satir.ek, dict) else {}
        elle = str(s.satir.masraf_merkezi_kaynak or ek.get("kaynak_santiye_devralindi") or "").strip()
        evrak = str(ek.get("evrak_no") or ek.get("fatura_no") or "").strip()
        paylar = _paylar(s, harita, devralinan_kullan=yinelenenleri_ele, uyarilar=tablo.uyarilar)
        # Kaynak dosyadaki sirket etiketi tablonun sirketiyle farkli mi?
        # Acente cogu zaman faturanin KESILDIGI tarafi yazar (RHI); tablo ise
        # kisinin personel kaydindaki tuzel kisiyi kullanir (ULTK, RSS...).
        # Fark, sirketler arasi yansitmanin ta kendisidir; cift bazinda
        # sayilir ki muhasebe 'RHI -> UST LUGA 96 satir' diye tek bakista gorsun.
        if elle and s.durum != "ESLESMEDI":
            elle_sirket = sirket_kanonik(elle)
            tablo_sirketleri = {p[2] for p in paylar if p[2]}
            if elle_sirket and tablo_sirketleri and elle_sirket not in tablo_sirketleri:
                u = uyusmazlik[pb]
                u["satir"] += 1
                u["tutar"] += float(s.satir.tutar)
                for hedef in sorted(tablo_sirketleri):
                    c = u["ciftler"][f"{elle_sirket} -> {hedef}"]
                    c["satir"] += 1
                    c["tutar"] += float(s.satir.tutar)
        for merkez, ad, sirket, haritada, pay_notu, oran in paylar:
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
                kayit._elle = set()        # type: ignore[attr-defined]
                kayit._evrak = []          # type: ignore[attr-defined]
                birikim[anahtar] = kayit
            if elle:
                kayit._elle.add(elle)      # type: ignore[attr-defined]
            if evrak and evrak not in kayit._evrak:   # type: ignore[attr-defined]
                kayit._evrak.append(evrak)            # type: ignore[attr-defined]
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
        elle = sorted(kayit._elle)                  # type: ignore[attr-defined]
        kayit.elle_etiket = "; ".join(elle) if elle else None
        evrak = kayit._evrak                        # type: ignore[attr-defined]
        if evrak:
            kayit.evrak_no = ", ".join(evrak[:5]) + (f" (+{len(evrak) - 5})" if len(evrak) > 5 else "")
    for pb, u in uyusmazlik.items():
        # Uyari listesine GIRMEZ: bu bir hata degil, yansitma bilgisidir.
        # Kontrol sayfasinda cift bazinda listelenir; kapak 'Dikkat'te tek satir.
        tablo.sirket_uyusmazligi[pb] = {
            "satir": u["satir"], "tutar": round(u["tutar"], 2),
            "ciftler": {
                ad: {"satir": c["satir"], "tutar": round(c["tutar"], 2)}
                for ad, c in sorted(u["ciftler"].items(), key=lambda kv: -kv[1]["tutar"])
            },
        }

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
