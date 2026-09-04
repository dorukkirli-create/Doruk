# Windows paketi nasil uretilir

Bu klasor, finans ekibinin bilgisayarina **Python kurmadan** kurulabilen
tasinabilir Windows paketini uretir.

## Mantik

Paket uc parcadan olusur:

1. **Tasinabilir CPython** (`python-build-standalone`, `install_only` surumu).
   Kayit defterine dokunmaz, yonetici hakki istemez, klasor tasinabilir.
2. **Windows tekerlekleri** (`.whl`). PyPI'dan `--platform win_amd64` ile
   indirilir; bu makinede Linux olsa bile Windows ikilileri gelir.
3. **Uygulama kodu** (`masraf/` paketi ve `calistir.py`).

Kullanicinin gordugu klasor:

```
Otomasyon\
  CALISTIR.bat        <- cift tiklanan dosya
  OKU_BENI.txt
  1_FATURALAR\        <- faturalar buraya
  2_EXCEL_CIKTI\      <- Excel buraya cikar
  PERSONEL\           <- personel ana verisi + 1C listesi
  veri\               <- harita ve ogrenilen defterler
  program\            <- python.exe, kutuphaneler, kod (kullanici dokunmaz)
```

`calistir.py` arayuzsuzdur; `app.py` (Streamlit) ayri bir kullanim yoludur ve
bu pakete girmez.

## Kirpilan bilesenler

Paketi 30 MB altina indirmek icin su parcalar cikarilir. Her biri once
"gercekten yuklenmiyor mu" diye olculmustur:

| Cikarilan | Neden guvenli |
|---|---|
| `tcl/`, `Lib/tkinter`, `tk*.dll` | Grafik arayuz; bu akista hic kullanilmiyor |
| `Lib/test`, `Lib/idlelib`, `Lib/ensurepip`, `Lib/lib2to3`, `Lib/pydoc_data` | CPython'un kendi test/gelistirme araclari |
| `include/`, `libs/`, `Scripts/` | C eklentisi derlemek icin; calisma aninda gereksiz |
| `pandas/tests`, `numpy/**/tests`, `numpy/f2py`, `numpy/distutils` | Paketlerin kendi test paketleri; `import` icin gerekmiyor |
| `cryptography`, `msoffcrypto` | Yalnizca PAROLA KORUMALI Office dosyalari icin; olculdu, normal akista yuklenmiyor |
| `easygui` | `extract_msg`'in istege bagli arayuzu; olculdu, yuklenmiyor |
| `*.pyi` | Tip ipuclari; calisma aninda okunmaz |

**Kirpmanin bedeli:** parola korumali Excel/Outlook dosyalari acilamaz.
`OKU_BENI.txt` bunu kullaniciya soyler.

Kirpmadan sonra dogrulama, bu modulleri `sys.meta_path` ile engelleyip
boru hattini gercek ornek mesajla ucdan uca calistirarak yapilir. Beklenen
sonuc: 405 satir okunur, 34 mahsup satiri uretilir, mutabakat kapanir.

## Boyut

Sikistirilmis paket ~30 MB. En buyuk parcalar `numpy.libs` (OpenBLAS, 21 MB)
ve `pandas` (22 MB). Bunlari da cikarmanin tek yolu `kayit.py` ile
`yardimci_defter.py` icindeki Excel okumayi pandas yerine dogrudan
openpyxl'e cevirmektir; o zaman paket ~12 MB'a duser.

## Uretme adimlari

Internet erisimi olan bir makinede:

```bash
# 1) Tasinabilir Windows Python
curl -sSL -o pywin.tar.gz \
  "https://github.com/astral-sh/python-build-standalone/releases/download/20250115/cpython-3.11.11+20250115-x86_64-pc-windows-msvc-install_only.tar.gz"

# 2) Windows tekerlekleri (surumler test edilenlerle AYNI olmali)
pip download --dest wheels --platform win_amd64 --only-binary=:all: --python-version 3.11 \
  "pandas==3.0.5" "numpy==2.4.6" "openpyxl==3.1.5" "xlrd==2.0.2" "rapidfuzz==3.14.6" \
  "xlsxwriter==3.2.9" "olefile==0.47" "tzlocal<6,>=4.2" "compressed-rtf<2,>=1.0.6" \
  "ebcdic<3,>=1.1.1" "beautifulsoup4<5,>=4.11.1" "RTFDE<0.2,>=0.1.1"
pip download --dest wheels --no-deps --platform win_amd64 --only-binary=:all: \
  --python-version 3.11 "extract-msg==0.56.1"
```

`red-black-tree-mod` icin tekerlek yoktur ve kaynaktan derlemesi bozuktur.
Paket saf Python'dur; kaynak arsivinden `red_black_dict_mod.py` ve
`red_black_set_mod.py` dosyalarini dogrudan `site-packages` icine kopyalayin.

Sonra tekerlekleri `program/Lib/site-packages` icine acin, `masraf/` ve
`calistir.py` dosyalarini `program/kod/` altina koyun, klasor yapisini
olusturun ve zipleyin.

## Surum uyumu

Tekerlek surumleri, gelistirme makinesinde test edilen surumlerle AYNI
olmalidir. Farkli surum, test edilmemis bir bilesim demektir. Ozellikle
`extract-msg` surumler arasi API degistirir.
