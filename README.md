# Harmonogram odbioru odpadów dla gminy Janowice Wielkie

Skrypt generujący kalendarz w Google Calendar z harmonogramem
odbioru odpadów dla gminy Janowice Wielkie.

Harmonogram w formacie arkusza Excel otrzymany z gminy lub
bezpośrednio od firmy COM-D, która realizuje odbiór odpadów.

## Instalacja

```console
python3 -m venv .venv
. .\venv\Scripts\activate.ps1
pip install --upgrade pip
pip install openpyxl
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
pip install tzdata
```

## Użycie

1. https://developers.google.com/calendar/api/quickstart/python
2. `python3 .\generate_schedule.py schedule.xlsx`