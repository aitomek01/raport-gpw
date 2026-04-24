# 📊 raport-gpw

Skrypt Python do automatycznej analizy portfela akcji GPW.  
Pobiera dane historyczne, oblicza statystyki i generuje sformatowany raport Excel z wykresem.

## Co robi

- Pobiera dane 5 spółek GPW za ostatnie 3 miesiące (yfinance)
- Oblicza: zwrot, zmienność, min/max, średnią cenę zamknięcia
- Generuje plik `.xlsx` z formatowaniem i wykresem kursów

## Spółki w portfelu

| Spółka | Ticker |
|---|---|
| PKO BP | PKO.WA |
| CD Projekt | CDR.WA |
| ORLEN | PKN.WA |
| Allegro | ALE.WA |
| Dino | DNP.WA |

## Uruchomienie

```bash
git clone https://github.com/aitomek01/raport-gpw.git
cd raport-gpw
pip install -r requirements.txt
python raport.py
```

## Wymagania

Python 3.10+ — szczegóły w `requirements.txt`

## Stack

`Python` · `yfinance` · `pandas` · `openpyxl` · `xlsxwriter`
