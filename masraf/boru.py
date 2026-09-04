"""Uctan uca boru hatti: dosya oku -> kisi esle -> masraf merkezine bagla.

Kullanim (masaustu arayuzu ve komut satiri ayni yolu kullanir)::

    from masraf.boru import Boru, CalismaAyarlari

    boru = Boru(CalismaAyarlari(personel_yolu="ornek_veri/personel/2025_2026_giris_cikis.xlsx"))
    boru.hazirla()
    sonuclar = boru.isle(["gelen/ANTIK_CARI_TEMMUZ_2026.xls"])
    ozet = boru.ozet(sonuclar)

Islem sirasi ve nedenleri:

1. TUM dosyalar once okunur. Boylece ayni calistirmada verilen yardimci
   listeler (saglik kontrol listesi, egitim katilimci listesi, arabuluculuk
   listesi) seyahat faturasindan ONCE ek kisi defterine islenir ve personel ana
   verisinde sicili olmayan kisiler o calistirmada eslesebilir.
2. Yardimci kaynaklardan defterler beslenir, sonra eslestirici kurulur
   (eslestirici indekslerini kurulusta olusturur, bu yuzden sira onemlidir).
3. Eslestirme TEK seferde, tum satirlar uzerinde yapilir. Aile bireyi kurali
   dosya sinirini asar: 'GUNAL EMRE' hangi dosyada eslesirse eslessin,
   'GUNAL DARIA' onun uzerinden cozulur.
4. Her satir masraf merkezine baglanir ve durumu belirlenir.

Calisma zamaninda yapay zeka veya internet KULLANILMAZ.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

from masraf.masraf_merkezi import (
    ALT_ESIK,
    GUVEN_ESIGI,
    MasrafMerkeziHaritasi,
    masraf_merkezi_coz,
    varsayilan_harita_yolu,
)
from masraf.metin import isim_normalize
from masraf.modeller import (
    DURUM_ESLESMEDI,
    DURUM_INCELE,
    DURUM_OTOMATIK,
    GiderSatiri,
    Sonuc,
)

__all__ = ["CalismaAyarlari", "Boru", "Ilerleme"]

#: Ilerleme geri cagrisi: (yuzde 0-100, mesaj)
Ilerleme = Callable[[float, str], None]


@dataclass
class CalismaAyarlari:
    """Bir calistirmanin tum ayarlari.

    Attributes:
        personel_yolu: Personel ana verisi (2025_2026_giris_cikis.xlsx).
        veri_dizini: Ogrenen defterlerin ve masraf merkezi haritasinin dizini.
        cikti_dizini: Excel/CSV ciktilarinin yazilacagi dizin.
        guven_esigi: Bu ve ustu guven + uyarisiz satir OTOMATIK sayilir.
        alt_esik: Bu esigin altindaki guven ESLESMEDI sayilir.
        harita_yolu: Masraf merkezi haritasi; None ise veri_dizini altindaki
            varsayilan dosya kullanilir.
        defterleri_besle: Yardimci listelerden ek kisi defterini doldur.
        ogrenmeyi_kaydet: Beslenen defterleri diske yaz.
        tuzel_kisi_uyar: Kaynak dosyadaki tuzel kisi etiketi personel
            kaydindaki 'Sirket 2' ile celisince uyari uret. Varsayilan KAPALI;
            elle etiketleme kendi icinde tutarsiz oldugu icin acildiginda
            satirlarin buyuk kismi incelemeye duser.
        onbellek: Personel dosyasi icin pickle onbellegini kullan.
    """

    personel_yolu: str | Path
    veri_dizini: str | Path = "veri"
    cikti_dizini: str | Path = "cikti"
    guven_esigi: float = GUVEN_ESIGI
    alt_esik: float = ALT_ESIK
    harita_yolu: str | Path | None = None
    defterleri_besle: bool = True
    ogrenmeyi_kaydet: bool = True
    tuzel_kisi_uyar: bool = False
    onbellek: bool = True

    def cozulmus_harita_yolu(self) -> Path:
        """Kullanilacak masraf merkezi haritasinin yolu."""
        if self.harita_yolu:
            return Path(self.harita_yolu)
        return varsayilan_harita_yolu(self.veri_dizini)


def _bildir(ilerleme: Ilerleme | None, yuzde: float, mesaj: str) -> None:
    """Ilerleme geri cagrisini guvenli bicimde tetikler."""
    if ilerleme is None:
        return
    try:
        ilerleme(max(0.0, min(100.0, float(yuzde))), mesaj)
    except Exception:
        pass  # Arayuz hatasi is akisini durdurmamali.


class Boru:
    """Personel defteri, ogrenen defterler, eslestirici ve masraf merkezi
    haritasini tek bir is akisinda birlestirir.

    ``hazirla()`` agir yuklemeleri (24 MB xlsx) bir kez yapar; ayni Boru
    ornegi ile birden fazla ``isle()`` cagrilabilir.
    """

    def __init__(self, ayarlar: CalismaAyarlari) -> None:
        self.ayarlar = ayarlar
        self.defter: Any = None          # kayit.PersonelDefteri
        self.defterler: Any = None       # defter.Defterler
        self.harita: MasrafMerkeziHaritasi | None = None
        self.eslestirici: Any = None
        self.hatalar: list[str] = []
        self.uyarilar: list[str] = []
        self._son_donem: date | None = None
        self._hazir = False

    # ------------------------------------------------------------------
    # Kurulum
    # ------------------------------------------------------------------

    def hazirla(self) -> None:
        """Personel defterini, ogrenen defterleri ve haritayi yukler.

        Birden fazla cagrilirsa ikinci ve sonraki cagrilar hicbir sey yapmaz.
        """
        if self._hazir:
            return

        from masraf.defter import Defterler
        from masraf.kayit import PersonelDefteri

        self.defter = PersonelDefteri.yukle(
            self.ayarlar.personel_yolu, onbellek=self.ayarlar.onbellek
        )
        donemler = self.defter.donemler
        self._son_donem = donemler[-1] if donemler else None

        self.defterler = Defterler(self.ayarlar.veri_dizini)

        harita_yolu = self.ayarlar.cozulmus_harita_yolu()
        self.harita = MasrafMerkeziHaritasi.yukle(harita_yolu)
        if not self.harita.kaynak_var:
            self.uyarilar.append(
                f"Masraf merkezi haritasi bulunamadi ({harita_yolu}); gorev yerleri "
                "oldugu gibi kullanilacak."
            )
        else:
            eksik = self.harita.eksikleri_bildir(set(self.defter.gorev_yerleri))
            if eksik:
                self.uyarilar.append(
                    "Masraf merkezi haritasinda tanimsiz gorev yeri: " + ", ".join(eksik)
                )

        self._hazir = True

    def _eslestirici_kur(self) -> None:
        """Eslestiriciyi (yeniden) kurar.

        Defterler beslendikten SONRA cagrilmalidir; eslestirici indekslerini
        kurulusta olusturur.
        """
        from masraf.eslestirici import Eslestirici

        self.eslestirici = Eslestirici(self.defter, self.defterler)

    # ------------------------------------------------------------------
    # Okuma
    # ------------------------------------------------------------------

    def oku(self, dosya_yollari: Sequence[str | Path],
            ilerleme: Ilerleme | None = None) -> list[GiderSatiri]:
        """Verilen dosyalari okur ve tum gider satirlarini dondurur.

        Okunamayan dosya is akisini durdurmaz; hata ``self.hatalar`` listesine
        yazilir ve digerleri islenmeye devam eder.
        """
        from masraf.okuyucular.kesif import dosya_tipini_bul, oku_tip
        from masraf.okuyucular.genel import genel_oku

        satirlar: list[GiderSatiri] = []
        toplam = max(1, len(dosya_yollari))
        for sira, yol in enumerate(dosya_yollari, start=1):
            hedef = Path(yol)
            _bildir(ilerleme, 5 + 25 * (sira - 1) / toplam, f"Okunuyor: {hedef.name}")
            try:
                tip = dosya_tipini_bul(hedef)
                dosya_satirlari = oku_tip(hedef, tip)
                if not dosya_satirlari and tip != "genel":
                    dosya_satirlari = genel_oku(hedef)
                if not dosya_satirlari:
                    self.hatalar.append(
                        f"{hedef.name}: dosyadan hic gider satiri cikarilamadi "
                        f"(tespit edilen tip: {tip})."
                    )
                satirlar.extend(dosya_satirlari)
            except Exception as hata:  # noqa: BLE001 - kullaniciya gosterilecek
                self.hatalar.append(f"{hedef.name}: okunamadi ({hata.__class__.__name__}: {hata})")
        return satirlar

    # ------------------------------------------------------------------
    # Ana is akisi
    # ------------------------------------------------------------------

    def isle(self, dosya_yollari: list[str] | Sequence[str | Path],
             ilerleme: Ilerleme | None = None) -> list[Sonuc]:
        """Dosyalari uctan uca isler ve sonuc listesini dondurur."""
        self.hatalar = []
        _bildir(ilerleme, 1, "Personel verisi yukleniyor")
        self.hazirla()

        satirlar = self.oku(dosya_yollari, ilerleme)
        if not satirlar:
            _bildir(ilerleme, 100, "Islenecek satir bulunamadi")
            return []

        # Yardimci listelerden ek kisi defterini besle (sicili olmayan kisiler).
        if self.ayarlar.defterleri_besle:
            _bildir(ilerleme, 32, "Yardimci listelerden kisi defteri besleniyor")
            try:
                besleme = self.defterler.yardimci_kaynaktan_besle(list(satirlar))
                if self.ayarlar.ogrenmeyi_kaydet and (
                    besleme.get("ek_kisi") or besleme.get("tckn_kopru")
                ):
                    self.defterler.kaydet()
            except Exception as hata:  # noqa: BLE001
                self.hatalar.append(f"Kisi defteri beslenemedi: {hata}")

        _bildir(ilerleme, 38, "Eslestirme motoru kuruluyor")
        self._eslestirici_kur()

        _bildir(ilerleme, 45, f"{len(satirlar)} satir eslestiriliyor")
        eslesmeler = self.eslestirici.esle_toplu(satirlar)

        sonuclar: list[Sonuc] = []
        toplam = max(1, len(satirlar))
        for sira, (satir, eslesme) in enumerate(zip(satirlar, eslesmeler), start=1):
            if sira % 25 == 0 or sira == toplam:
                _bildir(
                    ilerleme,
                    75 + 24 * sira / toplam,
                    f"Masraf merkezi cozumleniyor ({sira}/{toplam})",
                )
            sonuclar.append(
                masraf_merkezi_coz(
                    satir,
                    eslesme,
                    self.defter,
                    self.harita,
                    guven_esigi=self.ayarlar.guven_esigi,
                    alt_esik=self.ayarlar.alt_esik,
                    ek_masraf_merkezi=self._ek_merkez(satir, eslesme),
                    tuzel_kisi_uyar=self.ayarlar.tuzel_kisi_uyar,
                    son_donem=self._son_donem,
                )
            )

        _bildir(ilerleme, 100, "Tamamlandi")
        return sonuclar

    def _ek_merkez(self, satir: GiderSatiri, eslesme: Any) -> str | None:
        """Sicili olmayan kisiler icin defterlerden masraf merkezi onerisi.

        Harici kisiler defteri kullanicinin ogretttigi kesin bilgidir; ek kisi
        defteri (saglik/egitim listeleri) ise santiye bilgisi tasir.
        """
        if eslesme.sicil or self.defterler is None:
            return None
        adaylar = [eslesme.ad_soyad, satir.kisi_ham]
        for aday in adaylar:
            norm = isim_normalize(aday or "")
            if not norm:
                continue
            kayit = self.defterler.harici.get(norm)
            if kayit and kayit.get("masraf_merkezi"):
                return str(kayit["masraf_merkezi"])
            kayit = self.defterler.ek_kisiler.get(norm)
            if kayit and kayit.get("santiye"):
                return str(kayit["santiye"])
        tckn = str(satir.tckn_ham or "").strip()
        if tckn:
            kayit = self.defterler.ek_kisiler.get(tckn)
            if kayit and kayit.get("santiye"):
                return str(kayit["santiye"])
        return None

    # ------------------------------------------------------------------
    # Ozet
    # ------------------------------------------------------------------

    def ozet(self, sonuclar: Iterable[Sonuc]) -> dict:
        """Calistirmanin ozetini uretir (arayuz ve Excel 'Ozet' sayfasi icin).

        Tutarlar PARA BIRIMI BAZINDA toplanir; farkli dovizler asla toplanmaz.
        """
        sonuclar = list(sonuclar)
        durum_dagilimi: dict[str, int] = {
            DURUM_OTOMATIK: 0,
            DURUM_INCELE: 0,
            DURUM_ESLESMEDI: 0,
        }
        yontem_dagilimi: dict[str, int] = {}
        gider_tipi_dagilimi: dict[str, int] = {}
        kaynak_dagilimi: dict[str, int] = {}
        merkez_ozeti: dict[str, dict] = {}
        para_toplamlari: dict[str, float] = {}
        eslesmeyenler: dict[str, int] = {}
        gorev_yerleri: set[str] = set()

        for sonuc in sonuclar:
            durum_dagilimi[sonuc.durum] = durum_dagilimi.get(sonuc.durum, 0) + 1
            yontem = sonuc.eslesme.yontem or "yok"
            yontem_dagilimi[yontem] = yontem_dagilimi.get(yontem, 0) + 1
            tip = sonuc.satir.gider_tipi or "Diger"
            gider_tipi_dagilimi[tip] = gider_tipi_dagilimi.get(tip, 0) + 1
            kaynak = sonuc.satir.kaynak_tip or "genel"
            kaynak_dagilimi[kaynak] = kaynak_dagilimi.get(kaynak, 0) + 1

            if sonuc.gorev_yeri:
                gorev_yerleri.add(sonuc.gorev_yeri)
            # Kaynak dosyadan gelen ve haritada karsiligi olmadigi icin oldugu
            # gibi kullanilan etiketler de kullaniciya bildirilmelidir.
            if sonuc.masraf_merkezi and sonuc.satir.ek.get("cozum_kaynagi") != "personel":
                gorev_yerleri.add(sonuc.masraf_merkezi)

            tutar = sonuc.satir.tutar
            para = (sonuc.satir.para_birimi or "").strip() or "?"
            if tutar is not None:
                para_toplamlari[para] = para_toplamlari.get(para, 0.0) + float(tutar)

            kod = sonuc.masraf_merkezi or "(atanmadi)"
            kayit = merkez_ozeti.setdefault(
                kod,
                {
                    "kod": kod,
                    "ad": sonuc.satir.ek.get("masraf_merkezi_adi") or kod,
                    "adet": 0,
                    "tutarlar": {},
                    "otomatik": 0,
                    "incele": 0,
                    "eslesmedi": 0,
                },
            )
            kayit["adet"] += 1
            if sonuc.durum == DURUM_OTOMATIK:
                kayit["otomatik"] += 1
            elif sonuc.durum == DURUM_INCELE:
                kayit["incele"] += 1
            else:
                kayit["eslesmedi"] += 1
            if tutar is not None:
                kayit["tutarlar"][para] = kayit["tutarlar"].get(para, 0.0) + float(tutar)

            if sonuc.durum == DURUM_ESLESMEDI:
                ad = (sonuc.satir.kisi_ham or "").strip()
                if ad:
                    eslesmeyenler[ad] = eslesmeyenler.get(ad, 0) + 1

        toplam = len(sonuclar)
        dosyalar = sorted({s.satir.kaynak_dosya for s in sonuclar})
        eksik_merkezler = (
            self.harita.eksikleri_bildir(gorev_yerleri) if self.harita else sorted(gorev_yerleri)
        )

        return {
            "satir_sayisi": toplam,
            "dosya_sayisi": len(dosyalar),
            "dosyalar": dosyalar,
            "durum_dagilimi": durum_dagilimi,
            "durum_orani": {
                ad: (round(adet / toplam * 100, 1) if toplam else 0.0)
                for ad, adet in durum_dagilimi.items()
            },
            "otomatik_orani": (
                round(durum_dagilimi[DURUM_OTOMATIK] / toplam * 100, 1) if toplam else 0.0
            ),
            "yontem_dagilimi": dict(
                sorted(yontem_dagilimi.items(), key=lambda p: -p[1])
            ),
            "gider_tipi_dagilimi": dict(
                sorted(gider_tipi_dagilimi.items(), key=lambda p: -p[1])
            ),
            "kaynak_dagilimi": dict(sorted(kaynak_dagilimi.items(), key=lambda p: -p[1])),
            "masraf_merkezi_ozeti": sorted(
                merkez_ozeti.values(), key=lambda k: (-k["adet"], k["kod"])
            ),
            "para_birimi_toplamlari": para_toplamlari,
            "eslesmeyen_kisiler": dict(
                sorted(eslesmeyenler.items(), key=lambda p: (-p[1], p[0]))
            ),
            "eksik_masraf_merkezleri": eksik_merkezler,
            "uyarilar": list(self.uyarilar),
            "hatalar": list(self.hatalar),
            "guven_esigi": self.ayarlar.guven_esigi,
            "personel_dosyasi": str(self.ayarlar.personel_yolu),
            "son_donem": self._son_donem,
        }

    # ------------------------------------------------------------------
    # Kolaylik
    # ------------------------------------------------------------------

    def calistir(
        self,
        dosya_yollari: list[str] | Sequence[str | Path],
        cikti_adi: str | None = None,
        ilerleme: Ilerleme | None = None,
    ) -> dict:
        """Isle + ozet + Excel ciktisi; arayuzun cagirdigi tek fonksiyon.

        Returns:
            {'sonuclar': [...], 'ozet': {...}, 'excel_yolu': '...'}
        """
        from masraf.cikti import excel_yaz, varsayilan_cikti_adi

        sonuclar = self.isle(dosya_yollari, ilerleme)
        ozet = self.ozet(sonuclar)
        cikti_dizini = Path(self.ayarlar.cikti_dizini)
        cikti_dizini.mkdir(parents=True, exist_ok=True)
        yol = cikti_dizini / (cikti_adi or varsayilan_cikti_adi())
        excel_yolu = excel_yaz(sonuclar, str(yol), ozet) if sonuclar else ""
        return {"sonuclar": sonuclar, "ozet": ozet, "excel_yolu": excel_yolu}

    def istatistik(self) -> dict:
        """Yuklu bilesenlerin ozeti (arayuzde 'sistem durumu' panelinde)."""
        return {
            "hazir": self._hazir,
            "personel": self.defter.istatistik() if self.defter else None,
            "defterler": self.defterler.istatistik() if self.defterler else None,
            "harita": self.harita.istatistik() if self.harita else None,
            "eslestirici": self.eslestirici.istatistik() if self.eslestirici else None,
        }
