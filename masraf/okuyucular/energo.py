"""Energo yansitma dosyalarinin okuyuculari.

Dort ayri sablon, dort ayri fonksiyon:

    assessment_oku()      Assessment/degerlendirme yansitmasi. 'Kisi Listesi'
                          sayfasindaki 'Katilimci' kolonu, kisi basina
                          'Energo Payi' tutari.
    arabulucu_oku()       Arabuluculuk yansitmasi. 'Kisi Listesi' sayfasinda
                          PERSONEL + PERSONEL T.C. + PROJE bulunur; tutar
                          'Fatura Detay' sayfasindan sirket bazinda dagitilir.
    saglik_oku()          Saglik kontrol listesi. ADI SOYADI + TCKN + SANTIYE.
    koc_katilimci_oku()   Koc Universitesi katilimci listesi. ID kolonu
                          dogrudan SICIL NUMARASIDIR (en guvenilir eslesme).

Ortak kurallar: sayfa ve kolon adlari ASCII katlanmis karsilastirma ile
esnek cozulur, kisi adi bos olan ozet/toplam satirlari atlanir, TCKN 11
haneli rakam olarak dogrulanir, sicil metne cevrilip '.0' eki atilir.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from masraf.kayit import sicil_normalize
from masraf.modeller import GiderSatiri
from masraf.okuyucular.genel import (
    baslik_satiri_bul,
    calisma_oku,
    dolu_hucre_sayisi,
    hucre_metni,
    hucre_sayisi,
    hucre_tarihi,
    kolon_anahtari,
    kolon_ara,
    kolon_haritasi,
    sayfa_sec,
    tckn_normalize,
)

__all__ = [
    "assessment_oku",
    "arabulucu_oku",
    "saglik_oku",
    "koc_katilimci_oku",
]

# Sirket kisaltmalarinin kanonik karsiliklari. 'Fatura Detay' sayfasindaki
# 'Masraf yeri' kodlari ile 'Kisi Listesi' sayfasindaki 'Ilgili Sirket'
# degerlerini ayni anahtara indirger.
_SIRKET_ESLERI: dict[str, str] = {
    "rhi": "RHI",
    "rc": "RC",
    "rc peter": "RC",
    "rc moskova": "RC",
    "rsd": "RSD",
    "renstroydetal": "RSD",
    "renservis": "RENSERVIS",
    "rs": "RENSERVIS",
}


def _sirket_anahtari(deger: Any) -> str | None:
    """Sirket kodunu kanonik anahtara cevirir ('renstroydetal' -> 'RSD')."""
    metin = hucre_metni(deger)
    if metin is None:
        return None
    anahtar = kolon_anahtari(metin)
    return _SIRKET_ESLERI.get(anahtar, anahtar.upper() or None)


def _sayfa_satirlari(
    calisma: Any, *aday_adlar: str
) -> tuple[str | None, list[list[Any]]]:
    """Aday adlarla eslesen sayfayi bulur; bulamazsa ilk sayfayi dondurur."""
    ad = sayfa_sec(calisma.sayfa_adlari, *aday_adlar)
    if ad is None:
        ad = calisma.sayfa_adlari[0] if calisma.sayfa_adlari else None
    return ad, (calisma.satirlar(ad) if ad else [])


def _hucre_alici(satir: list[Any]):
    """Satir icin guvenli indeksli hucre okuyucu uretir."""

    def al(i: int | None) -> Any:
        if i is None or i >= len(satir):
            return None
        return satir[i]

    return al


# --------------------------------------------------------------------------
# 1) Assessment yansitma
# --------------------------------------------------------------------------

def assessment_oku(yol: str | Path) -> list[GiderSatiri]:
    """Assessment yansitma dosyasinin 'Kisi Listesi' sayfasini okur.

    Kisi adi 'Katilimci' kolonunda AD SOYAD sirasiyla bulunur. Tutar olarak
    kisi satirindaki 'Energo Payi' (USD) alinir; fatura toplami ve USD
    tutari ek sozlugunde saklanir. Masraf merkezi kaynagi 'Yansitma'
    kolonudur (RHI / RSD).
    """
    p = Path(yol)
    calisma = calisma_oku(p)
    sayfa_adi, satirlar = _sayfa_satirlari(calisma, "Kisi Listesi", "Katilimci Listesi")
    if not satirlar:
        return []

    baslik_i = baslik_satiri_bul(satirlar, aranan=("Katilimci", "Katılımcı"))
    if baslik_i < 0:
        return []
    harita = kolon_haritasi(satirlar[baslik_i])

    i_katilimci = kolon_ara(harita, "katilimci", "personel", "ad soyad")
    if i_katilimci is None:
        return []
    i_tarih = kolon_ara(harita, "tarih", icerir=False)
    i_pozisyon = kolon_ara(harita, "pozisyon")
    i_yetkili = kolon_ara(harita, "firma yetkilisi", "yetkili")
    i_tur = kolon_ara(harita, "uygulama turu")
    i_yer = kolon_ara(harita, "uygulama yeri")
    i_paket = kolon_ara(harita, "paket")
    i_fatura_no = kolon_ara(harita, "fatura numarasi", "fatura no")
    i_fatura_tarihi = kolon_ara(harita, "fatura tarihi")
    i_toplam = kolon_ara(harita, "toplam")
    i_usd = kolon_ara(harita, "usd")
    i_pay = kolon_ara(harita, "energo payi", "pay")
    i_yansitma = kolon_ara(harita, "yansitma", "masraf yeri", "sirket")

    sonuclar: list[GiderSatiri] = []
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        al = _hucre_alici(satir)
        katilimci = hucre_metni(al(i_katilimci))
        if katilimci is None:
            continue  # toplam satiri / pivot artigi

        tutar = hucre_sayisi(al(i_pay))
        if tutar is None:
            tutar = hucre_sayisi(al(i_usd))

        fatura_tarihi = hucre_tarihi(al(i_fatura_tarihi), calisma.datemode)
        katilim_tarihi = hucre_tarihi(al(i_tarih), calisma.datemode)

        sonuclar.append(
            GiderSatiri(
                kaynak_dosya=p.name,
                kaynak_tip="energo_assessment",
                satir_no=r + 1,
                belge_tarihi=fatura_tarihi or katilim_tarihi,
                aciklama=" | ".join(
                    m
                    for m in (
                        katilimci,
                        hucre_metni(al(i_paket)),
                        hucre_metni(al(i_tur)),
                        hucre_metni(al(i_fatura_no)),
                    )
                    if m
                ),
                kisi_ham=katilimci,
                sicil_ham=None,
                tckn_ham=None,
                tutar=tutar,
                para_birimi="USD",
                masraf_merkezi_kaynak=hucre_metni(al(i_yansitma)),
                gider_tipi="Egitim",
                ek={
                    "sayfa": sayfa_adi,
                    "katilim_tarihi": katilim_tarihi,
                    "pozisyon": hucre_metni(al(i_pozisyon)),
                    "firma_yetkilisi": hucre_metni(al(i_yetkili)),
                    "uygulama_turu": hucre_metni(al(i_tur)),
                    "uygulama_yeri": hucre_metni(al(i_yer)),
                    "paket": hucre_metni(al(i_paket)),
                    "fatura_no": hucre_metni(al(i_fatura_no)),
                    "fatura_toplam": hucre_sayisi(al(i_toplam)),
                    "fatura_usd": hucre_sayisi(al(i_usd)),
                    "tutar_yontemi": "kisi satirindaki Energo Payi",
                },
            )
        )
    return sonuclar


# --------------------------------------------------------------------------
# 2) Arabuluculuk yansitma
# --------------------------------------------------------------------------

def _arabulucu_fatura_ozeti(calisma: Any) -> dict[str, float]:
    """'Fatura Detay' sayfasindan sirket bazinda Energo Payi toplamini cikarir.

    Ozet/pivot satirlari 'Masraf yeri' bos oldugu icin dogal olarak atlanir.
    """
    ad = sayfa_sec(calisma.sayfa_adlari, "Fatura Detay", "Fatura Detayi")
    if ad is None:
        return {}
    satirlar = calisma.satirlar(ad)
    baslik_i = baslik_satiri_bul(satirlar, aranan=("Masraf yeri", "Energo Payi"))
    if baslik_i < 0:
        return {}
    harita = kolon_haritasi(satirlar[baslik_i])
    i_yer = kolon_ara(harita, "masraf yeri", "masraf merkezi", "sirket")
    i_pay = kolon_ara(harita, "energo payi", "fatura tutari usd", "pay")
    if i_yer is None or i_pay is None:
        return {}

    ozet: dict[str, float] = {}
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        al = _hucre_alici(satir)
        anahtar = _sirket_anahtari(al(i_yer))
        tutar = hucre_sayisi(al(i_pay))
        if anahtar is None or tutar is None:
            continue
        ozet[anahtar] = ozet.get(anahtar, 0.0) + tutar
    return ozet


def arabulucu_oku(yol: str | Path) -> list[GiderSatiri]:
    """Arabuluculuk yansitma dosyasinin 'Kisi Listesi' sayfasini okur.

    Bu sablonda PERSONEL T.C. (TCKN) ve PROJE kolonlari vardir; PROJE
    degerleri personel ana verisindeki 'Gorev Yeri' degerleriyle birebir
    ayni DEGILDIR, esleme tablosu gerektirir (veri/masraf_merkezi_haritasi.csv).

    Kisi Listesi sayfasinda tutar kolonu bulunmadigi icin tutar, 'Fatura
    Detay' sayfasindaki sirket bazli Energo Payi toplaminin o sirkete ait
    kisi sayisina esit bolunmesiyle hesaplanir; yontem ek['tutar_yontemi']
    icinde acikca belirtilir. Eslesen masraf yeri yoksa tutar None kalir.
    """
    p = Path(yol)
    calisma = calisma_oku(p)
    sayfa_adi, satirlar = _sayfa_satirlari(calisma, "Kisi Listesi", "Personel Listesi")
    if not satirlar:
        return []

    baslik_i = baslik_satiri_bul(satirlar, aranan=("PERSONEL", "ARABULUCU"))
    if baslik_i < 0:
        return []
    harita = kolon_haritasi(satirlar[baslik_i])

    i_tckn = kolon_ara(harita, "personel t c", "tckn", "tc kimlik no", "kimlik no")
    i_personel = kolon_ara(harita, "personel", "ad soyad", "adi soyadi")
    if i_personel is not None and i_personel == i_tckn:
        # 'personel' anahtari 'personel t c' kolonuna dusmus olabilir
        i_personel = harita.get("personel")
    if i_personel is None:
        return []
    i_tarih = kolon_ara(harita, "tarih", icerir=False)
    i_yetkili = kolon_ara(harita, "yetkili")
    i_proje = kolon_ara(harita, "proje", "santiye", "masraf yeri")
    i_sirket = kolon_ara(harita, "ilgili sirket", "sirket")
    i_arabulucu = kolon_ara(harita, "arabulucu")
    i_fatura_no = kolon_ara(harita, "fatura no", "fatura numarasi")
    i_fatura_tarihi = kolon_ara(harita, "fatura tarihi")

    # Once kisileri topla, sonra sirket basina esit paylastir.
    ham_satirlar: list[tuple[int, list[Any]]] = []
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        if dolu_hucre_sayisi(satir) == 0:
            continue
        al = _hucre_alici(satir)
        if hucre_metni(al(i_personel)) is None:
            continue
        ham_satirlar.append((r, satir))

    fatura_ozeti = _arabulucu_fatura_ozeti(calisma)
    sayimlar: dict[str, int] = {}
    for _, satir in ham_satirlar:
        anahtar = _sirket_anahtari(_hucre_alici(satir)(i_sirket))
        if anahtar:
            sayimlar[anahtar] = sayimlar.get(anahtar, 0) + 1

    sonuclar: list[GiderSatiri] = []
    for r, satir in ham_satirlar:
        al = _hucre_alici(satir)
        personel = hucre_metni(al(i_personel))
        sirket = _sirket_anahtari(al(i_sirket))

        tutar: float | None = None
        yontem = "fatura detayinda eslesen masraf yeri yok"
        if sirket and sirket in fatura_ozeti and sayimlar.get(sirket):
            tutar = round(fatura_ozeti[sirket] / sayimlar[sirket], 2)
            yontem = (
                f"'{sirket}' masraf yeri toplami "
                f"({fatura_ozeti[sirket]:.2f}) / {sayimlar[sirket]} kisi"
            )

        olay_tarihi = hucre_tarihi(al(i_tarih), calisma.datemode)
        fatura_tarihi = hucre_tarihi(al(i_fatura_tarihi), calisma.datemode)

        sonuclar.append(
            GiderSatiri(
                kaynak_dosya=p.name,
                kaynak_tip="energo_arabulucu",
                satir_no=r + 1,
                belge_tarihi=fatura_tarihi or olay_tarihi,
                aciklama=" | ".join(
                    m
                    for m in (
                        personel,
                        hucre_metni(al(i_proje)),
                        hucre_metni(al(i_arabulucu)),
                        hucre_metni(al(i_fatura_no)),
                    )
                    if m
                ),
                kisi_ham=personel,
                sicil_ham=None,
                tckn_ham=tckn_normalize(al(i_tckn)),
                tutar=tutar,
                para_birimi="USD",
                masraf_merkezi_kaynak=hucre_metni(al(i_proje)),
                gider_tipi="Arabuluculuk",
                ek={
                    "sayfa": sayfa_adi,
                    "olay_tarihi": olay_tarihi,
                    "yetkili": hucre_metni(al(i_yetkili)),
                    "ilgili_sirket": hucre_metni(al(i_sirket)),
                    "sirket_kodu": sirket,
                    "arabulucu": hucre_metni(al(i_arabulucu)),
                    "fatura_no": hucre_metni(al(i_fatura_no)),
                    "tutar_yontemi": yontem,
                    "fatura_ozeti": dict(fatura_ozeti),
                },
            )
        )
    return sonuclar


# --------------------------------------------------------------------------
# 3) Saglik kontrol listesi
# --------------------------------------------------------------------------

def saglik_oku(yol: str | Path, bordrosuz_dahil: bool = False) -> list[GiderSatiri]:
    """Saglik kontrol listesini okur ('BORDROLU LISTE' sayfasi).

    ADI SOYADI (AD SOYAD sirasiyla, Turkce karakterli), TCKN, DOGUM TARIHI ve
    SANTIYE kolonlari bulunur. TCKN sayesinde bu dosya, personel ana
    verisinde bulunmayan kisiler icin EK KISI DEFTERI kaynagidir.

    Args:
        bordrosuz_dahil: True ise 'BORDROSUZ LISTE' sayfasi da okunur.
            Varsayilan False - bordrosuz taseronlar masraf mahsuplastirma
            akisinda ayri ele alinir.
    """
    p = Path(yol)
    calisma = calisma_oku(p)

    hedef_sayfalar: list[str] = []
    bordrolu = sayfa_sec(calisma.sayfa_adlari, "BORDROLU LISTE", "BORDROLU")
    if bordrolu:
        hedef_sayfalar.append(bordrolu)
    if bordrosuz_dahil:
        bordrosuz = sayfa_sec(calisma.sayfa_adlari, "BORDROSUZ LISTE", "BORDROSUZ")
        if bordrosuz and bordrosuz not in hedef_sayfalar:
            hedef_sayfalar.append(bordrosuz)
    if not hedef_sayfalar and calisma.sayfa_adlari:
        hedef_sayfalar.append(calisma.sayfa_adlari[0])

    sonuclar: list[GiderSatiri] = []
    for sayfa_adi in hedef_sayfalar:
        satirlar = calisma.satirlar(sayfa_adi)
        baslik_i = baslik_satiri_bul(satirlar, aranan=("ADI SOYADI", "TCKN", "S.NO"))
        if baslik_i < 0:
            continue
        harita = kolon_haritasi(satirlar[baslik_i])

        i_sno = kolon_ara(harita, "s no", "sno", "sira no")
        i_isim = kolon_ara(harita, "adi soyadi", "ad soyad", "isim")
        if i_isim is None:
            continue
        i_tckn = kolon_ara(harita, "tckn", "tc kimlik no", "tc kimlik", "kimlik no")
        i_dogum = kolon_ara(harita, "dogum tarihi")
        i_ulke = kolon_ara(harita, "ulke")
        i_gorev = kolon_ara(harita, "gorevi", "gorev")
        i_santiye = kolon_ara(harita, "santiye", "proje")
        i_firma = kolon_ara(harita, "firma ekip formen", "firma")
        i_iletisim = kolon_ara(harita, "iletisim bilgileri 1", "iletisim")
        i_talep = kolon_ara(harita, "talep tarihi")
        i_kontrol = kolon_ara(
            harita, "saglik kontrol tarihi", "kontrol tarihi dr", "kontrol tarihi"
        )
        bordrolu_mu = "bordrosuz" not in kolon_anahtari(sayfa_adi)

        for r in range(baslik_i + 1, len(satirlar)):
            satir = satirlar[r]
            al = _hucre_alici(satir)
            isim = hucre_metni(al(i_isim))
            if isim is None:
                continue

            kontrol_tarihi = hucre_tarihi(al(i_kontrol), calisma.datemode)
            talep_tarihi = hucre_tarihi(al(i_talep), calisma.datemode)

            sonuclar.append(
                GiderSatiri(
                    kaynak_dosya=p.name,
                    kaynak_tip="energo_saglik",
                    satir_no=r + 1,
                    belge_tarihi=kontrol_tarihi or talep_tarihi,
                    aciklama=" | ".join(
                        m
                        for m in (
                            isim,
                            hucre_metni(al(i_gorev)),
                            hucre_metni(al(i_santiye)),
                        )
                        if m
                    ),
                    kisi_ham=isim,
                    sicil_ham=None,
                    tckn_ham=tckn_normalize(al(i_tckn)),
                    tutar=None,
                    para_birimi=None,
                    masraf_merkezi_kaynak=hucre_metni(al(i_santiye)),
                    gider_tipi="Saglik",
                    ek={
                        "sayfa": sayfa_adi,
                        "bordrolu": bordrolu_mu,
                        "s_no": hucre_metni(al(i_sno)),
                        "dogum_tarihi": hucre_tarihi(al(i_dogum), calisma.datemode),
                        "ulke": hucre_metni(al(i_ulke)),
                        "gorevi": hucre_metni(al(i_gorev)),
                        "firma": hucre_metni(al(i_firma)),
                        "iletisim": hucre_metni(al(i_iletisim)),
                        "talep_tarihi": talep_tarihi,
                        "kontrol_tarihi": kontrol_tarihi,
                    },
                )
            )
    return sonuclar


# --------------------------------------------------------------------------
# 4) Koc Universitesi katilimci listesi
# --------------------------------------------------------------------------

def koc_katilimci_oku(yol: str | Path) -> list[GiderSatiri]:
    """Koc Universitesi egitim katilimci listesini okur.

    'ID' kolonu dogrudan SICIL NUMARASIDIR; bu dosya en guvenilir eslesmeyi
    saglar. 'Ad Soyad' degerleri SOYAD AD sirasindadir ('Birladean Alexandr').

    Ayni kisi birden fazla katilim tarihi icin tekrar eder; her satir AYRI
    bir GiderSatiri olarak dondurulur (tekillestirme yapilmaz).
    """
    p = Path(yol)
    calisma = calisma_oku(p)
    sayfa_adi, satirlar = _sayfa_satirlari(
        calisma, "Sheet1", "Katilimci Listesi", "Kisi Listesi"
    )
    if not satirlar:
        return []

    baslik_i = baslik_satiri_bul(satirlar, aranan=("Ad Soyad", "Alt Fonksiyon", "ID"))
    if baslik_i < 0:
        return []
    harita = kolon_haritasi(satirlar[baslik_i])

    i_id = kolon_ara(harita, "id", "sicil", "sicil no", "personel no")
    i_isim = kolon_ara(harita, "ad soyad", "adi soyadi", "katilimci", "isim")
    if i_isim is None:
        return []
    i_pozisyon = kolon_ara(harita, "pozisyon", "gorev")
    i_fonksiyon = kolon_ara(harita, "alt fonksiyon", "fonksiyon")
    i_tarih = kolon_ara(harita, "katilim tarihi", "tarih")

    sonuclar: list[GiderSatiri] = []
    for r in range(baslik_i + 1, len(satirlar)):
        satir = satirlar[r]
        al = _hucre_alici(satir)
        isim = hucre_metni(al(i_isim))
        sicil = sicil_normalize(al(i_id)) or None
        if isim is None and sicil is None:
            continue

        katilim = hucre_tarihi(al(i_tarih), calisma.datemode)
        sonuclar.append(
            GiderSatiri(
                kaynak_dosya=p.name,
                kaynak_tip="koc_katilimci",
                satir_no=r + 1,
                belge_tarihi=katilim,
                aciklama=" | ".join(
                    m
                    for m in (
                        isim,
                        hucre_metni(al(i_pozisyon)),
                        hucre_metni(al(i_fonksiyon)),
                    )
                    if m
                ),
                kisi_ham=isim,
                sicil_ham=sicil,
                tckn_ham=None,
                tutar=None,
                para_birimi=None,
                masraf_merkezi_kaynak=None,
                gider_tipi="Egitim",
                ek={
                    "sayfa": sayfa_adi,
                    "pozisyon": hucre_metni(al(i_pozisyon)),
                    "alt_fonksiyon": hucre_metni(al(i_fonksiyon)),
                    "katilim_tarihi": katilim,
                },
            )
        )
    return sonuclar
