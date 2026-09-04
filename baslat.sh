#!/usr/bin/env bash
# Masraf Merkezi Otomasyonu - Linux / macOS baslatici
# Kullanim: ./baslat.sh   (ilk seferde: chmod +x baslat.sh)
set -euo pipefail

cd "$(dirname "$0")"

echo "============================================"
echo "  Masraf Merkezi Otomasyonu baslatiliyor..."
echo "============================================"
echo

# Python yorumlayicisini bul (python3 tercih edilir).
PY=""
for aday in python3 python; do
    if command -v "$aday" >/dev/null 2>&1; then
        PY="$aday"
        break
    fi
done

if [ -z "$PY" ]; then
    echo "HATA: Python bulunamadi."
    echo "Python 3.11 kurun: https://www.python.org/downloads/"
    exit 1
fi

if [ ! -d .venv ]; then
    echo "Sanal ortam olusturuluyor, bu islem bir kez yapilir..."
    "$PY" -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate

echo "Gerekli paketler kontrol ediliyor..."
python -m pip install --quiet --upgrade pip || true
python -m pip install --quiet -r requirements.txt || {
    echo "UYARI: Paketler kurulamadi (internet olmayabilir)."
    echo "Paketler daha once kurulduysa uygulama yine de calisir."
}

echo
echo "Arayuz aciliyor. Tarayici otomatik acilmazsa: http://localhost:8501"
echo "Kapatmak icin bu pencerede Ctrl+C tuslarina basin."
echo
exec python -m streamlit run app.py
