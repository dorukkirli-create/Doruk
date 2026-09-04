"""Kimlik eslestirme motoru: fatura satirindaki kisiyi personel sicili ile eslestirir.

Bu modul projenin en kritik parcasidir. Tamamen DETERMINISTIKTIR: internet,
yapay zeka veya harici servis kullanmaz. Ayni girdi her zaman ayni sonucu verir.

Eslestirme KADEMELIDIR. Kademeler guvenilirlik sirasina gore denenir, ilk
basarili olan kazanir:

===== ============================================================= ======
Adim  Yontem                                                        Guven
===== ============================================================= ======
1     ``sicil``           kaynak dosyada sicil verilmis               1.00
2     ``tckn``            TCKN -> sicil koprusu                       0.99
3     ``alias``           kullanicinin daha once ogrettigi eslesme    0.98
4     ``harici``          calisan degil, bilinen dis kisi             0.95
5     ``tam_isim``        normalize isim / token kumesi birebir       0.95
6     ``alt_kume``        fatura tokenlari personel tokenlarinin      0.90
                          alt kumesi (Rus patronimikleri)
7     ``tam_isim``        bitisik ad acilarak birebir                 0.92
8     ``transliterasyon`` Rusca yazim varyanti ile birebir            0.88
9     ``prefix``          kesilmis (truncate) PNR ismi                0.85
10    ``ek_defter``       yardimci kaynaklardaki ek kisi defteri      0.70
11    ``bulanik``         rapidfuzz token_set_ratio >= 88             ~0.80
12    ``aile``            soyadi eslesen calisanin aile bireyi        0.50
13    ``yok``             hicbiri                                     0.00
===== ============================================================= ======

SOZLESMEDEN BILINCLI SAPMALAR (hepsi dogruluk lehinedir):

* ``harici`` ve ``ek_defter`` kademeleri ``bulanik`` ve ``aile`` kademelerinden
  ONCE denenir. BILINEN kimlik, TAHMINDEN once gelir. Aksi halde "TALIP KEREM
  KOCKESEN" gibi calisan olmadigi kesin bilinen bir kisi, soyadi tesadufen
  eslesen bir calisanin aile bireyi sanilabilir; ya da saglik listesinden gelen
  "AHMET CELER", %81 benzerlikteki "Bicer Ahmet" ile karistirilabilirdi.
  ``harici`` kullanicinin acik karari oldugu icin ``tam_isim``den de oncedir.
* Aday sayisi birden fazlaysa ``sicil`` alani DOLDURULMAZ (``None`` kalir) ve
  adaylar ``aday_siciller`` icinde listelenir. Tek istisna: butun adaylar ayni
  gorev yerinde calisiyorsa masraf merkezi zaten tektir, bu durumda ilk aday
  onerilir ama guven yine 0,6'nin altinda tutulur.

BATCH IPUCU: ayni dosyadaki baska bir satirda kesin eslesen bir calisanin
soyadi, aile bireyi tespitinde kanit olarak kullanilir (ornegin "GUNAL EMRE"
kesin eslestiyse "GUNAL DARIA" onun aile bireyi sayilir). Sirali ``esle()``
cagrilarinda bu kanit ONCEKI satirlardan toplanir; sira bagimsiz ve tam sonuc
icin ``esle_toplu()`` kullanilmalidir.
"""

from __future__ import annotations

import bisect
from typing import Any, Iterable, Sequence

from rapidfuzz import fuzz, process

from masraf.defter import Defterler, tckn_normalize
from masraf.kayit import PersonelDefteri, sicil_normalize
from masraf.metin import (
    isim_imzasi,
    bitisik_ad_ac,
    isim_normalize,
    kisi_metnini_temizle,
    rus_disi_soyad_erkek_hali,
    translit_varyantlari,
)
from masraf.modeller import (
    DURUM_ESLESMEDI,
    DURUM_INCELE,
    DURUM_OTOMATIK,
    Eslesme,
    GiderSatiri,
    bos_eslesme,
)

# --------------------------------------------------------------------------
# Ayar sabitleri
# --------------------------------------------------------------------------

#: Bu guvenin altindaki eslesmeler kullanici incelemesine gonderilir.
INCELE_ESIGI = 0.80

#: Birden fazla aday bulundugunda guven bu degerin ustune CIKAMAZ.
COKLU_ADAY_TAVANI = 0.58

#: Bir eslesmenin "kesin" sayilip aile kanitina eklenmesi icin gereken guven.
OGRENME_ESIGI = 0.85

#: Bulanik eslesme icin asgari rapidfuzz token_set_ratio puani.
BULANIK_ESIK = 88.0

#: Bulanik eslesmede birinci adayin ikinciden onde olmasi gereken puan farki.
BULANIK_FARK = 6.0

#: Bulanik eslesmede bir tokenin aday havuzuna alinmasi icin ust siklik siniri.
#: Cok yaygin tokenlar (MEHMET, ALI) havuzu sisirir ve bir sey kazandirmaz.
TOKEN_SIKLIK_SINIRI = 600

#: Bulanik aday havuzunun mutlak ust siniri (performans korumasi).
HAVUZ_SINIRI = 6000

#: Ters alt kume (personel adi fatura metninin icinde) icin siklik siniri.
TERS_SIKLIK_SINIRI = 2000

#: Onek (prefix) eslesmesinde bir tokenin en az uzunlugu.
ONEK_ASGARI = 4

#: Aile kuralinda listelenecek azami aday sayisi.
AZAMI_ADAY = 10

#: Aile kuralinin "ayni dosyada kesin eslesen soyadas" ve "hepsi ayni gorev
#: yerinde" kanitlari yalnizca NADIR soyadlar icin gecerlidir. 28 bin kisilik
#: bir sirkette 'Ozturk' soyadli 12 calisan varken bunlardan birinin ayni ayda
#: seyahat etmesi TESADUF olabilir; 'Celenligil' veya 'Gunal' gibi 1-2 kisilik
#: soyadlarda ise aile bagi cok daha olasidir. Bu sinirin ustundeki soyadlarda
#: satir eslesmemis kabul edilip adaylariyla birlikte incelemeye gonderilir.
NADIR_SOYAD_SINIRI = 8

#: Bilet satirlarinda isim "SOYAD AD" sirasindadir (PNR bicimi); otel, vize ve
#: diger satirlarda "AD SOYAD" sirasindadir. Aile kuralinda soyadi dogru
#: konumdan almak icin kullanilir.
SOYAD_ONDE_TIPLERI: frozenset[str] = frozenset({"Bilet"})

#: Aile kuralinda bir tokenin "belirgin bicimde soyad" sayilmasi icin gereken
#: soyad olasiligi. Personel verisinde isimler 'SOYAD AD' sirali oldugundan bu
#: oran tokenin ilk konumda gecme sikligidir: 'GOZUKARA' 1,00 iken 'HASAN'
#: 0,02 civarindadir.
AILE_SOYAD_BELIRGIN = 0.50

#: Ote uc belirgin bicimde soyad iken bu esigin altinda kalan uc tamamen
#: elenir. 'HASAN HUSEYIN GOZUKARA' satirinda 'HASAN' bu kurala takilir ve
#: soyadi HASAN olan alakasiz calisanlara baglanma hatasi onlenir.
AILE_SOYAD_ASGARI = 0.10


def durum_belirle(eslesme: Eslesme) -> str:
    """Bir eslesmenin hangi cikti sayfasina gidecegini soyler.

    ``OTOMATIK`` (guven >= 0,80 ve sicil bulundu), ``INCELE`` (bir aday veya
    dis kisi bulundu ama guven dusuk) ya da ``ESLESMEDI``.
    """
    if eslesme.yontem == "yok":
        return DURUM_ESLESMEDI
    if eslesme.sicil is None and eslesme.yontem not in ("harici",):
        return DURUM_INCELE if eslesme.aday_siciller or eslesme.yontem == "ek_defter" else DURUM_ESLESMEDI
    if eslesme.guven >= INCELE_ESIGI:
        return DURUM_OTOMATIK
    return DURUM_INCELE


def _tekil(siciller: Iterable[str]) -> list[str]:
    """Sirayi bozmadan tekrarlari atar."""
    gorulen: set[str] = set()
    sonuc: list[str] = []
    for sicil in siciller:
        if sicil and sicil not in gorulen:
            gorulen.add(sicil)
            sonuc.append(sicil)
    return sonuc


class Eslestirici:
    """Gider satirlarindaki kisileri personel sicilleri ile eslestirir.

    Kurulumda personel defterinden ters indeksler (token -> isim) uretilir;
    bu sayede bulanik ve onek aramalarinda 24 bin ismin tamami taranmaz,
    once aday havuzu daraltilir (blocking). 130 satirlik bir dosya tipik
    olarak 1 saniyenin altinda eslesir.
    """

    def __init__(self, defter: PersonelDefteri, defterler: Defterler,
                 yardimci: Any = None) -> None:
        self._defter = defter
        self._defterler = defterler
        # 1C personel listesi: grup sirketlerini (Renservis, Renstroydetal,
        # RC, One Tower, Top Tower) kapsayan ikincil defter. Ana veride
        # bulunamayan kisiler BURADAN aranir. Opsiyoneldir.
        self._yardimci = yardimci

        # isim (normalize) -> sicil listesi
        self._isim_siciller: dict[str, list[str]] = self._isim_haritasi(defter)
        # isim -> token kumesi
        self._isim_tokenlar: dict[str, frozenset[str]] = {
            isim: frozenset(isim.split(" ")) for isim in self._isim_siciller
        }
        # token -> o tokeni iceren isimler (ters indeks / blocking)
        self._token_isimler: dict[str, list[str]] = {}
        for isim, tokenlar in self._isim_tokenlar.items():
            for token in tokenlar:
                self._token_isimler.setdefault(token, []).append(isim)
        # soyad (ilk token) -> o soyadi tasiyan isimler. Personel verisinde
        # isimler 'SOYAD AD' sirali oldugu icin bu indeks bir tokenin soyad
        # olma olasiligini olcmeye yarar (bkz. _soyad_olasiligi).
        self._soyad_isimler: dict[str, list[str]] = {}
        for isim in self._isim_siciller:
            self._soyad_isimler.setdefault(isim.split(" ", 1)[0], []).append(isim)
        # onek aramasi icin sirali token listesi
        self._tokenlar_sirali: list[str] = sorted(self._token_isimler)
        self._sozluk: set[str] = set(self._tokenlar_sirali)

        # Defterleri sirasiz (token kumesi) anahtarla da aranabilir yap; boylece
        # 'AD SOYAD' ile 'SOYAD AD' yazimlari ayni kayda ulasir ve her satirda
        # defterin tamamini taramak gerekmez.
        self._alias_token: dict[frozenset[str], str] = {}
        for isim, sicil in self._defterler.aliases.items():
            self._alias_token.setdefault(frozenset(isim.split(" ")), sicil)
        self._harici_token: dict[frozenset[str], dict] = {}
        # Isim sirasi ve bitisiklikten bagimsiz indeks. Fatura metinlerinde
        # ayni kisi 'KOCKESEN TALIPKEREM' ve 'TALIP KEREM KOCKESEN' olarak
        # iki turlu geciyor; token kumesi bunlari ayni saymiyor cunku
        # 'TALIPKEREM' tek token. Harf imzasi ikisini de yakalar.
        self._harici_imza: dict[str, dict] = {}
        for isim, kayit in self._defterler.harici.items():
            self._harici_token.setdefault(frozenset(isim.split(" ")), kayit)
            imza = isim_imzasi(isim)
            if imza:
                self._harici_imza.setdefault(imza, kayit)
            # Defterdeki 'ad_soyad' yazimi da indekslensin; kullanici
            # isim_norm kolonunu bos birakabilir ya da farkli yazabilir.
            ad_soyad = kayit.get("ad_soyad") or ""
            if ad_soyad:
                imza2 = isim_imzasi(ad_soyad)
                if imza2:
                    self._harici_imza.setdefault(imza2, kayit)
        self._ek_token: dict[frozenset[str], dict] = {}
        for anahtar, kayit in self._defterler.ek_kisiler.items():
            if " " in anahtar:
                self._ek_token.setdefault(frozenset(anahtar.split(" ")), kayit)

        # Ayni dosyada kesin eslesen calisanlarin soyadlari (aile kaniti).
        self._kesin_soyadlar: dict[str, set[str]] = {}

    # ------------------------------------------------------------------
    # Kurulum yardimcilari
    # ------------------------------------------------------------------

    @staticmethod
    def _isim_haritasi(defter: PersonelDefteri) -> dict[str, list[str]]:
        """Personel defterinden 'normalize isim -> sicil listesi' haritasi cikarir.

        Once defterin hazir isim indeksi kullanilir; yapisi degismisse harita
        ham kayitlardan yeniden kurulur. Bordrosuz taseron kayitlarinin adi bos
        oldugu icin bu haritaya girmezler.
        """
        hazir = getattr(defter, "_isim_index", None)
        if isinstance(hazir, dict) and hazir:
            return {isim: list(siciller) for isim, siciller in hazir.items()}

        harita: dict[str, list[str]] = {}
        for kayit in getattr(defter, "_kayitlar", []):
            isim = kayit.get("ad_soyad_norm") or isim_normalize(kayit.get("ad_soyad") or "")
            sicil = kayit.get("sicil")
            if not isim or not sicil:
                continue
            liste = harita.setdefault(isim, [])
            if sicil not in liste:
                liste.append(sicil)
        return harita

    def _ad(self, sicil: str) -> str | None:
        """Sicile ait en guncel adi dondurur."""
        kayit = self._defter.sicil_ile(sicil)
        if kayit is None:
            return None
        return kayit.get("ad_soyad")

    def _etiket(self, sicil: str) -> str:
        """Kullaniciya gosterilecek 'Ad Soyad / sicil' etiketi."""
        ad = self._ad(sicil)
        return f"{ad} / {sicil}" if ad else f"(isimsiz kayit) / {sicil}"

    def _gorev_yeri(self, sicil: str) -> str | None:
        kayit = self._defter.sicil_ile(sicil)
        return kayit.get("gorev_yeri") if kayit else None

    def _ortak_gorev_yeri(self, siciller: Sequence[str]) -> str | None:
        """Butun adaylar ayni gorev yerindeyse o gorev yerini dondurur."""
        yerler = {self._gorev_yeri(sicil) for sicil in siciller}
        if len(yerler) == 1:
            tek = yerler.pop()
            return tek if tek else None
        return None

    # ------------------------------------------------------------------
    # Aday arama (blocking + indeks)
    # ------------------------------------------------------------------

    def _siklik(self, token: str) -> int:
        return len(self._token_isimler.get(token, ()))

    def _onekli_tokenlar(self, onek: str, azami: int = 60) -> list[str]:
        """Verilen onekle baslayan personel tokenlarini dondurur."""
        if not onek:
            return []
        bas = bisect.bisect_left(self._tokenlar_sirali, onek)
        sonuc: list[str] = []
        for token in self._tokenlar_sirali[bas: bas + azami]:
            if not token.startswith(onek):
                break
            sonuc.append(token)
        return sonuc

    def _tam_isimler(self, norm: str, tokenlar: frozenset[str]) -> list[str]:
        """Birebir isim ve birebir token kumesi eslesmelerinin sicilleri.

        Token kumesi karsilastirmasi sirasizdir; boylece faturadaki
        'AD SOYAD' yazimi personel verisindeki 'SOYAD AD' yazimina eslesir.
        """
        adaylar = list(self._defter.isimle_adaylar(norm))
        adaylar.extend(self._defter.token_ile_adaylar(tokenlar))
        return _tekil(adaylar)

    def _alt_kume_isimleri(self, tokenlar: frozenset[str]) -> list[str]:
        """Fatura tokenlarini TAMAMEN iceren personel isimleri.

        'TKACHEVA NATALIA' -> 'Tkacheva Natalia Vitalievna' gibi Rus
        patronimikleri bu adimda cozulur.
        """
        if len(tokenlar) < 2:
            return []
        sikliklar = []
        for token in tokenlar:
            adet = self._siklik(token)
            if adet == 0:
                return []  # Tokenlardan biri veride hic gecmiyor: kapsam imkansiz.
            sikliklar.append((adet, token))
        _, en_nadir = min(sikliklar)
        return [
            isim
            for isim in self._token_isimler[en_nadir]
            if tokenlar < self._isim_tokenlar[isim]
        ]

    def _ters_alt_kume_isimleri(self, tokenlar: frozenset[str]) -> list[str]:
        """Personel isminin TAMAMI fatura metninin icinde geciyorsa.

        Ayiklama artigi kalan satirlar icindir:
        'OMER CAN CETIR KONSOLOSLUK UCRETI' gibi. Riskli oldugu icin dusuk
        guvenle ve yalnizca en az iki tokenli personel isimleri icin kullanilir.
        """
        if len(tokenlar) < 3:
            return []
        havuz: set[str] = set()
        for token in tokenlar:
            isimler = self._token_isimler.get(token)
            if isimler and len(isimler) <= TERS_SIKLIK_SINIRI:
                havuz.update(isimler)
            if len(havuz) > HAVUZ_SINIRI:
                break
        return [
            isim
            for isim in havuz
            if len(self._isim_tokenlar[isim]) >= 2 and self._isim_tokenlar[isim] < tokenlar
        ]

    def _onek_isimleri(self, tokenlar: frozenset[str]) -> list[str]:
        """Kesilmis (truncate edilmis) fatura ismine uyan personel isimleri.

        PNR alanlari sabit uzunlukta oldugu icin uzun isimler kesilir:
        'OZAKAY/MUSTAFAKEMA', 'ALLANAZAROV/ALLANAZA'. Kural: her fatura tokeni
        ya birebir eslesir ya da (en az 4 harfliyse) bir personel tokeninin
        onekidir; en az bir token birebir, en az bir token onek olmalidir.
        """
        if len(tokenlar) < 2:
            return []
        tam_eslesenler = [t for t in tokenlar if t in self._token_isimler]
        if not tam_eslesenler:
            return []
        en_nadir = min(tam_eslesenler, key=self._siklik)
        havuz = self._token_isimler[en_nadir]
        if len(havuz) > HAVUZ_SINIRI:
            return []
        sirali_tokenlar = sorted(tokenlar, key=len, reverse=True)
        sonuc: list[str] = []
        for isim in havuz:
            personel_tokenlar = self._isim_tokenlar[isim]
            if len(personel_tokenlar) < len(tokenlar):
                continue
            kalan = set(personel_tokenlar)
            onek_kullanildi = False
            uygun = True
            for token in sirali_tokenlar:
                if token in kalan:
                    kalan.discard(token)
                    continue
                if len(token) < ONEK_ASGARI:
                    uygun = False
                    break
                adaylar = sorted(
                    (k for k in kalan if len(k) > len(token) and k.startswith(token)),
                    key=len,
                )
                if not adaylar:
                    uygun = False
                    break
                kalan.discard(adaylar[0])
                onek_kullanildi = True
            if uygun and onek_kullanildi:
                sonuc.append(isim)
        return sonuc

    def _bulanik_havuz(self, tokenlar: frozenset[str]) -> list[str]:
        """Bulanik eslesme icin aday isim havuzunu daraltir (blocking).

        Her fatura tokeni icin hem tokenin kendisi hem de ilk dort harfini
        paylasan personel tokenlari kullanilir; boylece 'NERGIS' -> 'NERGIZ'
        gibi son harf hatalari da havuza girer. Cok yaygin tokenlar atlanir.
        """
        havuz: set[str] = set()
        for token in tokenlar:
            if len(token) < 3:
                continue
            yakinlar = {token}
            yakinlar.update(self._onekli_tokenlar(token[:ONEK_ASGARI]))
            for yakin in yakinlar:
                isimler = self._token_isimler.get(yakin)
                if isimler and len(isimler) <= TOKEN_SIKLIK_SINIRI:
                    havuz.update(isimler)
            if len(havuz) >= HAVUZ_SINIRI:
                break
        return list(havuz)[:HAVUZ_SINIRI]

    # ------------------------------------------------------------------
    # Eslesme uretimi
    # ------------------------------------------------------------------

    def _isimlerden_siciller(self, isimler: Iterable[str]) -> list[str]:
        siciller: list[str] = []
        for isim in isimler:
            siciller.extend(self._isim_siciller.get(isim, ()))
        return _tekil(siciller)

    def _tekil_eslesme(
        self, siciller: Sequence[str], yontem: str, guven: float, aciklama: str
    ) -> Eslesme | None:
        """Aday listesinden Eslesme uretir; liste bossa None.

        Aday sayisi birden fazlaysa guven 0,6'nin altina cekilir ve sicil
        DOLDURULMAZ. Tek istisna: butun adaylar ayni gorev yerindeyse masraf
        merkezi zaten tektir, ilk aday oneri olarak verilir.
        """
        adaylar = _tekil(siciller)
        if not adaylar:
            return None
        if len(adaylar) == 1:
            sicil = adaylar[0]
            return Eslesme(
                sicil=sicil,
                ad_soyad=self._ad(sicil),
                yontem=yontem,
                guven=guven,
                aday_sayisi=1,
                aciklama=f"{aciklama}: {self._etiket(sicil)}",
                aday_siciller=[sicil],
            )

        kisa_liste = adaylar[:AZAMI_ADAY]
        ortak = self._ortak_gorev_yeri(adaylar)
        if ortak:
            sicil = adaylar[0]
            return Eslesme(
                sicil=sicil,
                ad_soyad=self._ad(sicil),
                yontem=yontem,
                guven=COKLU_ADAY_TAVANI,
                aday_sayisi=len(adaylar),
                aciklama=(
                    f"{aciklama}: {len(adaylar)} aday var, ancak hepsi ayni gorev "
                    f"yerinde ({ortak}); masraf merkezi tek. Oneri: {self._etiket(sicil)}. "
                    "Kisi dogrulanmali."
                ),
                aday_siciller=kisa_liste,
            )
        ornekler = ", ".join(self._etiket(s) for s in kisa_liste[:4])
        return Eslesme(
            sicil=None,
            ad_soyad=None,
            yontem=yontem,
            guven=0.55,
            aday_sayisi=len(adaylar),
            aciklama=(
                f"{aciklama}: {len(adaylar)} farkli calisan ayni isme sahip ve gorev "
                f"yerleri farkli, otomatik secim yapilmadi. Adaylar: {ornekler}"
                + (" ..." if len(adaylar) > 4 else "")
            ),
            aday_siciller=kisa_liste,
        )

    # ------------------------------------------------------------------
    # Kademeler
    # ------------------------------------------------------------------

    def _adim_sicil(self, satir: GiderSatiri, notlar: list[str]) -> Eslesme | None:
        ham = sicil_normalize(satir.sicil_ham or "")
        if not ham:
            return None
        kayit = self._defter.sicil_ile(ham)
        if kayit is None:
            notlar.append(
                f"Kaynak dosyadaki {ham} sicili personel ana verisinde bulunamadi"
            )
            return None
        return Eslesme(
            sicil=ham,
            ad_soyad=kayit.get("ad_soyad"),
            yontem="sicil",
            guven=1.00,
            aday_sayisi=1,
            aciklama=f"Kaynak dosyada sicil dogrudan verilmis: {self._etiket(ham)}",
            aday_siciller=[ham],
        )

    def _adim_tckn(self, satir: GiderSatiri, notlar: list[str]) -> Eslesme | None:
        tckn = tckn_normalize(satir.tckn_ham)
        if not tckn:
            return None
        sicil = self._defterler.tckn_sicil.get(tckn)
        if not sicil:
            notlar.append(f"TCKN {tckn[:3]}****** icin sicil koprusu tanimli degil")
            return None
        if self._defter.sicil_ile(sicil) is None:
            notlar.append(f"TCKN koprusundeki {sicil} sicili personel verisinde yok")
            return None
        return Eslesme(
            sicil=sicil,
            ad_soyad=self._ad(sicil),
            yontem="tckn",
            guven=0.99,
            aday_sayisi=1,
            aciklama=f"TCKN -> sicil koprusu ile eslesti: {self._etiket(sicil)}",
            aday_siciller=[sicil],
        )

    def _adim_alias(self, norm: str, tokenlar: frozenset[str]) -> Eslesme | None:
        sicil = self._defterler.aliases.get(norm)
        gerekce = "Kullanicinin daha once ogrettigi eslesme (alias)"
        if not sicil:
            sicil = self._alias_token.get(tokenlar)
            gerekce = "Kullanicinin ogrettigi eslesme (alias, kelime sirasi farkli)"
        if not sicil:
            return None
        if self._defter.sicil_ile(sicil) is None:
            return Eslesme(
                sicil=sicil,
                ad_soyad=None,
                yontem="alias",
                guven=0.60,
                aday_sayisi=1,
                aciklama=(
                    f"Alias tablosunda {sicil} sicili yazili ancak bu sicil personel "
                    "ana verisinde yok; alias guncel olmayabilir."
                ),
                aday_siciller=[sicil],
            )
        return Eslesme(
            sicil=sicil,
            ad_soyad=self._ad(sicil),
            yontem="alias",
            guven=0.98,
            aday_sayisi=1,
            aciklama=f"{gerekce}: {self._etiket(sicil)}",
            aday_siciller=[sicil],
        )

    def _adim_tam_isim(self, norm: str, tokenlar: frozenset[str]) -> Eslesme | None:
        siciller = self._tam_isimler(norm, tokenlar)
        return self._tekil_eslesme(siciller, "tam_isim", 0.95, "Tam isim eslesmesi")

    def _adim_alt_kume(self, tokenlar: frozenset[str]) -> Eslesme | None:
        isimler = self._alt_kume_isimleri(tokenlar)
        if isimler:
            eslesme = self._tekil_eslesme(
                self._isimlerden_siciller(isimler),
                "alt_kume",
                0.90,
                "Fatura ismi personel isminin icinde geciyor (ornegin baba adi/patronimik eksik)",
            )
            if eslesme is not None:
                return eslesme
        # Ters yon: personel ismi, ayiklama artigi kalmis fatura metninin icinde.
        ters = self._ters_alt_kume_isimleri(tokenlar)
        if ters:
            return self._tekil_eslesme(
                self._isimlerden_siciller(ters),
                "alt_kume",
                0.72,
                "Personel adinin tamami fatura metninde geciyor (metinde fazladan kelimeler var)",
            )
        return None

    def _adim_bitisik_ad(self, tokenlar: frozenset[str]) -> Eslesme | None:
        """Bitisik yazilmis adi acar ('MUSTAFAKEMAL' -> 'MUSTAFA KEMAL')."""
        acilanlar: list[tuple[str, list[str]]] = []
        for token in tokenlar:
            if len(token) < 6:
                continue
            if self._siklik(token) > 0:
                continue  # Token zaten gecerli bir isim, acmaya gerek yok.
            parcalar = bitisik_ad_ac(token, self._sozluk)
            if parcalar:
                acilanlar.append((token, parcalar))
        if not acilanlar:
            return None

        yeni = set(tokenlar)
        aciklamalar: list[str] = []
        for token, parcalar in acilanlar:
            yeni.discard(token)
            yeni.update(parcalar)
            aciklamalar.append(f"{token} -> {' '.join(parcalar)}")
        yeni_tokenlar = frozenset(yeni)
        yeni_norm = " ".join(sorted(yeni_tokenlar))

        gerekce = "Bitisik yazilmis ad acildi (" + "; ".join(aciklamalar) + ")"
        eslesme = self._tekil_eslesme(
            self._tam_isimler(yeni_norm, yeni_tokenlar), "tam_isim", 0.92, gerekce
        )
        if eslesme is not None:
            return eslesme
        isimler = self._alt_kume_isimleri(yeni_tokenlar)
        if isimler:
            return self._tekil_eslesme(
                self._isimlerden_siciller(isimler), "alt_kume", 0.86, gerekce
            )
        return None

    def _adim_transliterasyon(self, norm: str, tokenlar: frozenset[str]) -> Eslesme | None:
        """Rusca/pasaport transliterasyonunu geri cevirerek eslestirir."""
        for varyant in sorted(translit_varyantlari(norm)):
            if varyant == norm:
                continue
            varyant_tokenlar = frozenset(varyant.split(" "))
            gerekce = f"Transliterasyon varyanti '{norm}' -> '{varyant}' ile eslesti"
            eslesme = self._tekil_eslesme(
                self._tam_isimler(varyant, varyant_tokenlar),
                "transliterasyon",
                0.88,
                gerekce,
            )
            if eslesme is not None:
                return eslesme
        for varyant in sorted(translit_varyantlari(norm)):
            if varyant == norm:
                continue
            varyant_tokenlar = frozenset(varyant.split(" "))
            isimler = self._alt_kume_isimleri(varyant_tokenlar)
            if isimler:
                eslesme = self._tekil_eslesme(
                    self._isimlerden_siciller(isimler),
                    "transliterasyon",
                    0.85,
                    f"Transliterasyon varyanti '{norm}' -> '{varyant}', personel isminin "
                    "icinde geciyor",
                )
                if eslesme is not None:
                    return eslesme
        return None

    def _adim_onek(self, tokenlar: frozenset[str]) -> Eslesme | None:
        isimler = self._onek_isimleri(tokenlar)
        if not isimler:
            return None
        return self._tekil_eslesme(
            self._isimlerden_siciller(isimler),
            "prefix",
            0.85,
            "Fatura ismi kesilmis (PNR alan sinirlamasi), onek olarak eslesti",
        )

    def _adim_bulanik(self, norm: str, tokenlar: frozenset[str]) -> Eslesme | None:
        if len(tokenlar) < 2:
            # Tek kelimelik bir metin ('MEHMET') binlerce kisiye benzer; bulanik
            # eslesme burada bilgi degil gurultu uretir.
            return None
        havuz = self._bulanik_havuz(tokenlar)
        if not havuz:
            return None
        sonuclar = process.extract(
            norm,
            havuz,
            scorer=fuzz.token_set_ratio,
            limit=12,
            score_cutoff=BULANIK_ESIK - BULANIK_FARK,
        )
        if not sonuclar:
            return None
        # Ayni kisi birden fazla isim yazimiyla gelebilir; sicil bazinda en iyi puan.
        sicil_puan: dict[str, float] = {}
        sicil_isim: dict[str, str] = {}
        for isim, puan, _ in sonuclar:
            for sicil in self._isim_siciller.get(isim, ()):
                if puan > sicil_puan.get(sicil, -1.0):
                    sicil_puan[sicil] = puan
                    sicil_isim[sicil] = isim
        if not sicil_puan:
            return None
        sirali = sorted(sicil_puan.items(), key=lambda ikili: (-ikili[1], ikili[0]))
        en_iyi_sicil, en_iyi_puan = sirali[0]
        if en_iyi_puan < BULANIK_ESIK:
            return None
        if len(sirali) > 1 and en_iyi_puan - sirali[1][1] < BULANIK_FARK:
            kisa_liste = [s for s, p in sirali if en_iyi_puan - p < BULANIK_FARK][:AZAMI_ADAY]
            ortak = self._ortak_gorev_yeri(kisa_liste)
            aciklama = (
                f"Bulanik benzerlik %{en_iyi_puan:.0f}, ancak {len(kisa_liste)} aday "
                "birbirine cok yakin puanda"
            )
            if ortak:
                return Eslesme(
                    sicil=kisa_liste[0],
                    ad_soyad=self._ad(kisa_liste[0]),
                    yontem="bulanik",
                    guven=COKLU_ADAY_TAVANI,
                    aday_sayisi=len(kisa_liste),
                    aciklama=(
                        f"{aciklama}; hepsi ayni gorev yerinde ({ortak}), masraf merkezi "
                        f"tek. Oneri: {self._etiket(kisa_liste[0])}"
                    ),
                    aday_siciller=kisa_liste,
                )
            return Eslesme(
                sicil=None,
                ad_soyad=None,
                yontem="bulanik",
                guven=0.50,
                aday_sayisi=len(kisa_liste),
                aciklama=(
                    f"{aciklama}: "
                    + ", ".join(self._etiket(s) for s in kisa_liste[:4])
                ),
                aday_siciller=kisa_liste,
            )
        return Eslesme(
            sicil=en_iyi_sicil,
            ad_soyad=self._ad(en_iyi_sicil),
            yontem="bulanik",
            guven=round(en_iyi_puan / 100.0 * 0.9, 4),
            aday_sayisi=1,
            aciklama=(
                f"Bulanik benzerlik %{en_iyi_puan:.0f} ('{norm}' ~ "
                f"'{sicil_isim[en_iyi_sicil]}'): {self._etiket(en_iyi_sicil)}"
            ),
            aday_siciller=[en_iyi_sicil],
        )

    def _adim_harici(self, norm: str, tokenlar: frozenset[str]) -> Eslesme | None:
        kayit = (
            self._defterler.harici.get(norm)
            or self._harici_token.get(tokenlar)
            or self._harici_imza.get(isim_imzasi(norm))
        )
        if kayit is None:
            return None
        kurum = kayit.get("kurum") or "bilinmeyen kurum"
        merkez = kayit.get("masraf_merkezi") or ""
        ek = f", masraf merkezi: {merkez}" if merkez else ""
        return Eslesme(
            sicil=None,
            ad_soyad=kayit.get("ad_soyad") or norm,
            yontem="harici",
            guven=0.95,
            aday_sayisi=0,
            aciklama=(
                f"Calisan degil, harici kisiler defterinde kayitli: {kayit.get('ad_soyad') or norm}"
                f" ({kurum}{ek})"
            ),
            aday_siciller=[],
        )

    def _adim_yardimci_defter(
        self, norm: str, tokenlar: frozenset[str]
    ) -> Eslesme | None:
        """1C personel listesinde arar (grup sirketleri dahil).

        Ana veri sadece RHI ve UST LUGA tuzel kisilerini kapsar. Renservis,
        Renstroydetal, RC, One Tower, Top Tower personeli ancak burada bulunur.
        Olculdu: 1C listesindeki 17.517 isimli kaydin 5.234'u ana veride yok.
        """
        if self._yardimci is None:
            return None
        adaylar = self._yardimci.isimle_adaylar(norm)
        yontem_notu = "ad soyad"
        if not adaylar:
            adaylar = self._yardimci.token_ile_adaylar(tokenlar)
            yontem_notu = "ad soyad (kelime sirasi farkli)"
        if not adaylar:
            return None
        if len(adaylar) > 1:
            kayitlar = [self._yardimci.sicil_ile(s) or {} for s in adaylar[:6]]
            return Eslesme(
                sicil=None,
                ad_soyad=(kayitlar[0].get("ad_soyad") if kayitlar else None),
                yontem="yardimci_defter",
                guven=0.55,
                aday_sayisi=len(adaylar),
                aciklama=(
                    f"1C personel listesinde '{norm}' icin {len(adaylar)} aday var: "
                    + ", ".join(
                        f"{k.get('ad_soyad')} / {k.get('sirket2')} / {k.get('gorev_yeri')}"
                        for k in kayitlar
                    )
                    + ". Dogru kisiyi secin."
                ),
                aday_siciller=list(adaylar),
            )
        sicil = adaylar[0]
        kayit = self._yardimci.sicil_ile(sicil) or {}
        return Eslesme(
            sicil=sicil,
            ad_soyad=kayit.get("ad_soyad"),
            yontem="yardimci_defter",
            guven=0.90,
            aday_sayisi=1,
            aciklama=(
                f"1C personel listesinde {yontem_notu} ile bulundu: "
                f"{kayit.get('ad_soyad')} / {sicil}, sirket {kayit.get('sirket')} "
                f"({kayit.get('sirket2')}), proje {kayit.get('gorev_yeri')}. "
                "Ana personel verisinde yok, bu yuzden donem dogrulanamaz."
            ),
            aday_siciller=[sicil],
        )

    def _adim_ek_defter(
        self, satir: GiderSatiri, norm: str, tokenlar: frozenset[str]
    ) -> Eslesme | None:
        tckn = tckn_normalize(satir.tckn_ham)
        kayit = self._defterler.ek_kisiler.get(tckn) if tckn else None
        gerekce = "TCKN ile ek kisi defterinde bulundu"
        if kayit is None:
            kayit = self._defterler.ek_kisiler.get(norm)
            gerekce = "Ek kisi defterinde (saglik/egitim/arabuluculuk listeleri) bulundu"
        if kayit is None:
            kayit = self._ek_token.get(tokenlar)
            if kayit is not None:
                gerekce = "Ek kisi defterinde bulundu (kelime sirasi farkli)"
        if kayit is None:
            return None
        santiye = kayit.get("santiye") or ""
        ek = f", santiye: {santiye}" if santiye else ", santiye bilgisi yok"
        return Eslesme(
            sicil=None,
            ad_soyad=kayit.get("ad_soyad") or norm,
            yontem="ek_defter",
            guven=0.70,
            aday_sayisi=0,
            aciklama=(
                f"{gerekce}: {kayit.get('ad_soyad') or norm}{ek}. Personel ana verisinde "
                "sicili yok (taseron / yeni giren / aday olabilir), dogrulanmali."
            ),
            aday_siciller=[],
        )

    def _soyad_olasiligi(self, token: str) -> float:
        """Bir tokenin SOYAD olma olasiligi (0.0 - 1.0).

        Personel ana verisinde isimler 'SOYAD AD' sirasindadir, yani bir
        tokenin ilk konumda gecme orani onun ne kadar soyad oldugunu soyler.
        'GOZUKARA' yalnizca soyad olarak gecer (1,00); 'HASAN' yuzlerce isimde
        ad olarak gecer, soyad olarak yalnizca birkacinda (~0,02).

        Bu olcu, gider tipinden gelen "bilet satirlarinda soyad ONDEDIR"
        varsayimi yanildiginda dogru tokeni secmeyi saglar.
        """
        toplam = self._siklik(token)
        if not toplam:
            return 0.0
        return len(self._soyad_isimler.get(token, ())) / toplam

    def _aile_adaylari(self, soyad: str, tokenlar: frozenset[str]) -> tuple[list[str], str]:
        """Bir soyad tokeni icin aile adayi sicilleri toplar.

        Dogrudan arama sonuc vermezse Rusca kadin soyadi eki geri cevrilir
        ('NOVOSELOVA' -> 'NOVOSELOV'); uretilen erkek hali personel soyad
        indeksinde GERCEKTEN varsa kullanilir, yoksa atilir.

        Returns:
            (aday siciller, kullanilan soyad)
        """
        kullanilan = soyad
        adaylar = _tekil(self._defter.soyad_ile_adaylar(soyad))
        if not adaylar:
            erkek = rus_disi_soyad_erkek_hali(soyad)
            if erkek:
                bulunan = _tekil(self._defter.soyad_ile_adaylar(erkek))
                if bulunan:
                    adaylar, kullanilan = bulunan, erkek
        # Ismi faturadakiyle birebir ayni olanlar zaten onceki adimlarda
        # denendi; burada sadece FARKLI kisiler kalmali.
        adaylar = [s for s in adaylar if self._kayit_tokenlari(s) != tokenlar]
        return adaylar, kullanilan

    def _aile_kaniti(
        self, soyad: str, tokenlar: frozenset[str]
    ) -> tuple[int, Eslesme] | None:
        """Tek bir soyad adayi icin kanit kademesini ve eslesmeyi uretir.

        Returns:
            (kanit kademesi, Eslesme) veya aday yoksa None. Kanit kademesi
            buyudukce kimlik tespiti guclenir:
            3 = ayni dosyada kesin eslesen soyadas, 2 = tek calisan,
            1 = coklu aday ama hepsi ayni gorev yerinde, 0 = belirsiz.
        """
        adaylar, kullanilan = self._aile_adaylari(soyad, tokenlar)
        if not adaylar:
            return None
        cevrildi = kullanilan != soyad
        ek_not = (
            f" ('{soyad}' Rusca kadin soyadi eki cozulerek '{kullanilan}' okundu)"
            if cevrildi else ""
        )

        # Soyadi tasiyan calisan sayisi: kanit kurallarinin gecerlilik olcusu.
        nadir = len(adaylar) <= NADIR_SOYAD_SINIRI

        # Ayni dosyada kesin eslesen bir calisan varsa aile bagi kanitlanir.
        # Yaygin soyadlarda bu kanit gecersizdir (tesadufen ayni soyadli
        # baska bir calisan da seyahat etmis olabilir).
        kesinler = self._kesin_soyadlar.get(kullanilan, set()) if nadir else set()
        ortak_kesin = [s for s in adaylar if s in kesinler]
        if len(ortak_kesin) == 1:
            sicil = ortak_kesin[0]
            return 3, Eslesme(
                sicil=sicil,
                ad_soyad=self._ad(sicil),
                yontem="aile",
                guven=0.60,
                aday_sayisi=len(adaylar),
                aciklama=(
                    f"'{kullanilan}' soyadli calisan {self._etiket(sicil)} ayni dosyada "
                    f"kesin eslesti; bu satir onun aile bireyi olabilir{ek_not}. Masraf "
                    "merkezi o calisandan devralindi, dogrulanmali."
                ),
                aday_siciller=adaylar[:AZAMI_ADAY],
            )

        if len(adaylar) == 1:
            sicil = adaylar[0]
            return 2, Eslesme(
                sicil=sicil,
                ad_soyad=self._ad(sicil),
                yontem="aile",
                guven=0.50,
                aday_sayisi=1,
                aciklama=(
                    f"Soyadi '{kullanilan}' olan tek calisan {self._etiket(sicil)} ile "
                    f"eslesiyor, ad tokenlari tutmuyor: aile bireyi olabilir{ek_not}. "
                    "Masraf merkezi o calisandan devralindi, dogrulanmali."
                ),
                aday_siciller=[sicil],
            )

        ortak = self._ortak_gorev_yeri(adaylar)
        if ortak and nadir:
            sicil = adaylar[0]
            return 1, Eslesme(
                sicil=sicil,
                ad_soyad=self._ad(sicil),
                yontem="aile",
                guven=0.45,
                aday_sayisi=len(adaylar),
                aciklama=(
                    f"Soyadi '{kullanilan}' olan {len(adaylar)} calisan var; hepsi ayni "
                    f"gorev yerinde ({ortak}), bu yuzden masraf merkezi kesin{ek_not}. "
                    "Aile bireyi olabilir, kisi dogrulanmali."
                ),
                aday_siciller=adaylar[:AZAMI_ADAY],
            )

        return 0, Eslesme(
            sicil=None,
            ad_soyad=None,
            yontem="aile",
            guven=0.30,
            aday_sayisi=len(adaylar),
            aciklama=(
                f"Soyadi '{kullanilan}' olan {len(adaylar)} farkli calisan var ve gorev "
                f"yerleri farkli; aile bireyi olabilir ama hangi calisana ait oldugu "
                f"belirlenemedi{ek_not}. Adaylar: "
                + ", ".join(self._etiket(s) for s in adaylar[:4])
            ),
            aday_siciller=adaylar[:AZAMI_ADAY],
        )

    def _adim_aile(
        self, satir: GiderSatiri, norm: str, tokenlar: frozenset[str]
    ) -> Eslesme | None:
        """Soyadi bir calisanla eslesen aile bireylerini isaretler.

        Soyadin hangi tokende oldugunu iki kaynak birlikte belirler:

        1. GIDER TIPI. Bilet satirlari PNR bicimindedir ('SOYAD AD'), otel /
           vize satirlari 'AD SOYAD' bicimindedir. Bu yalnizca bir TERCIHTIR.
        2. VERININ KENDISI. Her iki uctaki token icin ``_soyad_olasiligi``
           hesaplanir. Gider tipi yaniltici olabilir: elle dagitilmis seyahat
           dosyasinda bilet satirlari PNR degil duz 'AD SOYAD' yazilidir, bu
           yuzden 'HASAN HUSEYIN GOZUKARA' satirinda tercih edilen ilk token
           ('HASAN') soyad DEGILDIR.

        Iki uc de degerlendirilir, en guclu kanit kazanir. Bir uctaki token
        soyad olma olasiligi cok dusukken ote uc belirgin bicimde soyad ise
        zayif uc tamamen elenir; boylece 'HASAN' soyadli alakasiz bir calisana
        baglanma hatasi olusmaz.
        """
        parcalar = norm.split(" ")
        if len(parcalar) < 2:
            return None

        if satir.gider_tipi in SOYAD_ONDE_TIPLERI:
            tercih = "ilk"
        elif satir.gider_tipi:
            tercih = "son"
        else:
            tercih = ""

        uclar = [(parcalar[0], "ilk"), (parcalar[-1], "son")]
        uclar = [(t, k) for t, k in uclar if len(t) >= 3]
        if not uclar:
            return None

        olasiliklar = {konum: self._soyad_olasiligi(token) for token, konum in uclar}
        en_yuksek = max(olasiliklar.values(), default=0.0)

        adaylar: list[tuple[int, float, int, Eslesme]] = []
        for token, konum in uclar:
            olasilik = olasiliklar[konum]
            tercihli = not tercih or konum == tercih
            # Gider tipinin gosterdigi ucun DISINDAKI uc ancak belirgin bicimde
            # soyad ise degerlendirilir. Aksi halde 'TRAPEZNIKOVA POLINA'
            # satirinda Rusca bir AD olan 'POLINA' soyad sanilir ve alakasiz
            # bir calisana baglanir.
            if not tercihli and olasilik < AILE_SOYAD_BELIRGIN:
                continue
            # Ote uc belirgin bicimde soyad iken bu uc soyad degilse ele.
            if (
                en_yuksek >= AILE_SOYAD_BELIRGIN
                and olasilik < AILE_SOYAD_ASGARI
                and olasilik < en_yuksek
            ):
                continue
            kanit = self._aile_kaniti(token, tokenlar)
            if kanit is None:
                continue
            kademe, eslesme = kanit
            adaylar.append((kademe, olasilik, 1 if konum == tercih else 0, eslesme))

        if not adaylar:
            return None
        adaylar.sort(key=lambda k: (k[0], k[1], k[2]), reverse=True)
        return adaylar[0][3]

    def _kayit_tokenlari(self, sicil: str) -> frozenset[str]:
        kayit = self._defter.sicil_ile(sicil)
        if not kayit:
            return frozenset()
        return frozenset((kayit.get("ad_soyad_norm") or "").split(" ")) - {""}

    # ------------------------------------------------------------------
    # Ana giris noktasi
    # ------------------------------------------------------------------

    def esle(self, satir: GiderSatiri) -> Eslesme:
        """Bir gider satirindaki kisiyi personel sicili ile eslestirir.

        Kademeler modul docstring'indeki sirayla denenir; ilk basarili olan
        kazanir. Donen ``Eslesme.aciklama`` her zaman Turkce ve gerekcelidir.
        """
        notlar: list[str] = []

        eslesme = self._adim_sicil(satir, notlar)
        if eslesme is None:
            eslesme = self._adim_tckn(satir, notlar)
        if eslesme is not None:
            return self._bitir(eslesme, notlar)

        ham = satir.kisi_ham or ""
        norm = isim_normalize(ham)
        if not norm:
            # Kisi adi ayiklanamadiysa aciklamanin kendisinden bir kez daha dene.
            norm = isim_normalize(kisi_metnini_temizle(satir.aciklama or ""))
        if not norm:
            return self._bitir(
                bos_eslesme(
                    "Satirda kisi adi bulunamadi; masraf bir kisiye bagli degil "
                    "(organizasyon, celenk, genel hizmet vb)."
                ),
                notlar,
            )

        tokenlar = frozenset(norm.split(" "))

        for uretici in (
            lambda: self._adim_alias(norm, tokenlar),
            lambda: self._adim_harici(norm, tokenlar),
            lambda: self._adim_tam_isim(norm, tokenlar),
            lambda: self._adim_alt_kume(tokenlar),
            lambda: self._adim_bitisik_ad(tokenlar),
            lambda: self._adim_transliterasyon(norm, tokenlar),
            lambda: self._adim_onek(tokenlar),
            lambda: self._adim_yardimci_defter(norm, tokenlar),
            lambda: self._adim_ek_defter(satir, norm, tokenlar),
            lambda: self._adim_bulanik(norm, tokenlar),
            lambda: self._adim_aile(satir, norm, tokenlar),
        ):
            eslesme = uretici()
            if eslesme is not None:
                return self._bitir(eslesme, notlar)

        return self._bitir(
            bos_eslesme(
                f"'{norm}' icin personel ana verisinde, alias ve ek kisi defterlerinde "
                "karsilik bulunamadi. Kisi grup sirketi calisani, taseron, dis danisman "
                "veya yeni giren olabilir; inceleyip deftere ekleyin."
            ),
            notlar,
        )

    def esle_toplu(self, satirlar: Sequence[GiderSatiri]) -> list[Eslesme]:
        """Bir dosyanin tamamini iki gecisli olarak eslestirir (SIRA BAGIMSIZ).

        Birinci gecis kesin eslesmeleri bulur ve bunlarin soyadlarini kanit
        olarak toplar; ikinci gecis bu kanitla aile bireylerini cozer. Ornegin
        'GUNAL EMRE' dosyanin herhangi bir yerinde kesin eslestiyse
        'GUNAL DARIA' onun aile bireyi olarak isaretlenir.

        Tekil ``esle()`` cagrilarindan farki: sonuc, satirlarin dosyadaki
        sirasindan BAGIMSIZDIR.
        """
        self.sifirla()
        ilk = [self.esle(satir) for satir in satirlar]
        ikinci: list[Eslesme] = []
        for satir, eslesme in zip(satirlar, ilk):
            if eslesme.yontem in ("aile", "yok") or (
                eslesme.sicil is None and eslesme.guven < OGRENME_ESIGI
            ):
                ikinci.append(self.esle(satir))
            else:
                ikinci.append(eslesme)
        return ikinci

    def sifirla(self) -> None:
        """Dosya ici aile kanitlarini temizler (yeni bir dosyaya gecerken)."""
        self._kesin_soyadlar.clear()

    # ------------------------------------------------------------------
    # Ic yardimcilar
    # ------------------------------------------------------------------

    def _bitir(self, eslesme: Eslesme, notlar: list[str]) -> Eslesme:
        """Uyari notlarini aciklamaya ekler ve aile kanitini gunceller."""
        if notlar:
            eslesme.aciklama = f"{eslesme.aciklama} | Uyari: {'; '.join(notlar)}"
        self._ogren(eslesme)
        return eslesme

    def _ogren(self, eslesme: Eslesme) -> None:
        """Kesin eslesmeleri dosya ici aile kaniti olarak kaydeder."""
        if not eslesme.sicil or eslesme.guven < OGRENME_ESIGI:
            return
        if eslesme.yontem == "aile":
            return
        kayit = self._defter.sicil_ile(eslesme.sicil)
        if not kayit:
            return
        norm = kayit.get("ad_soyad_norm") or ""
        if not norm:
            return
        soyad = norm.split(" ", 1)[0]
        self._kesin_soyadlar.setdefault(soyad, set()).add(eslesme.sicil)

    def istatistik(self) -> dict[str, Any]:
        """Motorun kurulum ozeti (arayuzde ve testte gosterilir)."""
        return {
            "isim_sayisi": len(self._isim_siciller),
            "token_sayisi": len(self._token_isimler),
            "alias": len(self._defterler.aliases),
            "harici": len(self._defterler.harici),
            "ek_kisi": len(self._defterler.ek_kisiler),
            "tckn_kopru": len(self._defterler.tckn_sicil),
        }
