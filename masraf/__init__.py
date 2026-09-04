"""Masraf merkezi mahsuplastirma otomasyonu - cekirdek paket.

Gelen tedarikci faturalarindaki her satiri, o satirdaki kisiye gore dogru
masraf merkezine (proje / santiye) mahsuplastirmak icin kullanilan
deterministik (yapay zekasiz) modullerin paketi.

Moduller:
    modeller : Veri siniflari (GiderSatiri, Eslesme, Sonuc)
    metin    : Turkce/Kiril metin normalizasyonu, transliterasyon, temizlik
    kayit    : Personel ana verisi defteri ve indeksleri
"""

__version__ = "1.0.0"

__all__ = ["modeller", "metin", "kayit"]
