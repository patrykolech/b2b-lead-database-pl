"""
Uruchamia caly proces jednym poleceniem:

    python uruchom_wszystko.py [plik_z_linkami.txt]

Etapy:
1. Pobranie wizytowek Google z linkow (linki.txt)      -> wizytowki.json / .csv
2. Wyszukanie maili dla firm, ktore maja strone WWW    -> wyniki_email.csv
3. Scalenie jednego z drugim                           -> wizytowki_z_mailami.json
4. Arkusz Excela gotowy do przegladania                -> wizytowki_z_mailami.xlsx

Kazdy etap otwiera i zamyka wlasna przegladarke, jeden po drugim - dwa
konteksty Playwrighta nie moga naraz korzystac z tego samego profilu.

Oba skrypty skladowe mozna nadal uruchamiac osobno:
    python scraper_wizytowki.py linki.txt
    python email_finder.py domeny.txt maile.csv
"""

import json
import shutil
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

import scraper_wizytowki as scraper
from email_finder import (
    czy_agregator,
    domena_z_url,
    szukaj_maili,
    zapisz_csv,
    zapytanie_dla,
)
from wspolne import otworz_przegladarke

# Popularne darmowe skrzynki - firmy naprawde ich uzywaja jako kontaktu,
# wiec adres na wp.pl czy gmailu nie jest podejrzany sam w sobie.
POCZTA_PUBLICZNA = {
    "wp.pl", "gmail.com", "o2.pl", "interia.pl", "onet.pl", "op.pl",
    "poczta.onet.pl", "vp.pl", "gazeta.pl", "tlen.pl", "outlook.com",
    "hotmail.com", "yahoo.com", "icloud.com",
}

KAT = Path(__file__).parent
PLIK_MAILE = KAT / "wyniki_email.csv"
PLIK_SCALONY = KAT / "wizytowki_z_mailami.json"
PLIK_SCALONY_CSV = KAT / "wizytowki_z_mailami.csv"
PLIK_EXCEL = KAT / "wizytowki_z_mailami.xlsx"

# Pliki, ktore nadpisujemy - przed startem robimy z nich kopie "_poprzednie".
DO_KOPII = [
    KAT / "wizytowki.json",
    KAT / "wizytowki.csv",
    PLIK_MAILE,
    PLIK_SCALONY,
    PLIK_SCALONY_CSV,
    PLIK_EXCEL,
]

NAGLOWKI = {
    "nazwa": "Nazwa firmy",
    "kategoria": "Kategoria",
    "ocena": "Ocena",
    "liczba_opinii": "Liczba opinii",
    "adres": "Adres",
    "telefon": "Telefon",
    "strona_www": "Strona WWW",
    "email": "E-mail",
    "email_pewnosc": "Pewnosc e-maila",
    "status_otwarcia": "Status otwarcia",
    "godziny_otwarcia": "Godziny otwarcia",
    "opis": "Opis",
    "profile_spolecznosciowe": "Profile spolecznosciowe",
    "place_id": "Place ID",
    "cid_hex": "CID (hex)",
    "link": "Link do wizytowki",
    "blad": "Blad",
}

SZEROKOSCI = {
    "nazwa": 45, "kategoria": 28, "ocena": 8, "liczba_opinii": 12, "adres": 38,
    "telefon": 16, "strona_www": 40, "email": 34, "email_pewnosc": 20,
    "status_otwarcia": 26,
    "godziny_otwarcia": 55, "opis": 60, "profile_spolecznosciowe": 40,
    "place_id": 22, "cid_hex": 38, "link": 50, "blad": 20,
}

ZAWIJANE = ("opis", "godziny_otwarcia", "adres", "nazwa")


def zrob_kopie_zapasowe():
    """Zabezpiecza poprzednie wyniki - nieudany przebieg nie skasuje
    tego, co juz bylo zebrane."""
    for plik in DO_KOPII:
        if plik.exists():
            kopia = plik.with_name(f"{plik.stem}_poprzednie{plik.suffix}")
            shutil.copy2(plik, kopia)
    print("Kopie zapasowe poprzednich wynikow: *_poprzednie.*\n")


def etap_wizytowki(linki: list[str]) -> list[dict]:
    """Etap 1 - pobranie wizytowek. Korzysta z petli scrapera."""
    print("=" * 64)
    print(f"ETAP 1/2: pobieram {len(linki)} wizytowek Google")
    print("=" * 64)
    wyniki = scraper.przetworz_linki(linki)
    scraper.zapisz_wyniki(wyniki)
    udane = sum(1 for w in wyniki if w.get("nazwa"))
    print(f"\nEtap 1 gotowy: {udane}/{len(wyniki)} wizytowek pobranych.\n")
    return wyniki


def etap_maile(wizytowki: list[dict]) -> dict[str, str]:
    """Etap 2 - maile tylko dla firm, ktore maja adres WWW.

    Zwraca mape place_id -> email."""
    do_szukania = []  # (place_id, fraza, domena)
    for w in wizytowki:
        if not w.get("strona_www"):
            continue
        fraza, domena = zapytanie_dla(w.get("strona_www"), w.get("nazwa"))
        if fraza:
            do_szukania.append((w.get("place_id"), fraza, domena))

    print("=" * 64)
    print(f"ETAP 2/2: szukam maili dla {len(do_szukania)} firm ze strona WWW")
    print("=" * 64)

    if not do_szukania:
        print("Zadna wizytowka nie ma adresu WWW - pomijam etap maili.\n")
        return {}

    pozycje = [f for _, f, _ in do_szukania]
    domeny = {f: d for _, f, d in do_szukania}

    znalezione: dict[str, str] = {}
    with sync_playwright() as p:
        kontekst, strona, _ = otworz_przegladarke(p)
        try:
            znalezione = szukaj_maili(strona, pozycje, domeny)
        finally:
            zapisz_csv(znalezione, PLIK_MAILE)
            kontekst.close()

    # przepisujemy z fraz na place_id - to jednoznaczny klucz wizytowki
    wg_place_id = {}
    for place_id, fraza, _ in do_szukania:
        mail = znalezione.get(fraza)
        if mail:
            wg_place_id[place_id] = mail

    print(f"\nEtap 2 gotowy: {len(wg_place_id)}/{len(do_szukania)} adresow znalezionych.\n")
    return wg_place_id


def pewnosc_maila(email: str | None, strona_www: str | None) -> str | None:
    """Ocenia, na ile mozna ufac znalezionemu adresowi.

    Adres w domenie firmy jest pewny. Adres w innej domenie moze byc rownie
    dobry (firmy czesto uzywaja wp.pl czy gmail.com), ale moze tez pochodzic
    z sasiedniego wyniku wyszukiwania - takie warto zerknac okiem."""
    if not email:
        return None

    domena = domena_z_url(strona_www)
    po_malpie = email.lower().split("@")[-1]

    if domena and (po_malpie == domena or po_malpie.endswith("." + domena)):
        return "domena firmy"

    # ta sama nazwa, inna koncowka: promed24.pl vs promed24.eu
    if domena and po_malpie.split(".")[0] == domena.split(".")[0]:
        return "domena firmy"

    if po_malpie in POCZTA_PUBLICZNA:
        return "poczta publiczna"

    # firma nie ma wlasnej strony (Facebook, ZnanyLekarz) - szukalismy po
    # nazwie, wiec nie ma domeny, z ktora mozna by porownac adres
    if czy_agregator(domena):
        return "szukane po nazwie"

    return "inna domena - sprawdz"


def scal(wizytowki: list[dict], maile: dict[str, str]) -> list[dict]:
    """Dokłada pola email i email_pewnosc zaraz za strona_www,
    dopasowujac po place_id."""
    scalone = []
    for w in wizytowki:
        rekord = {}
        for pole, wartosc in w.items():
            rekord[pole] = wartosc
            if pole == "strona_www":
                mail = maile.get(w.get("place_id")) or None
                rekord["email"] = mail
                rekord["email_pewnosc"] = pewnosc_maila(mail, wartosc)
        if "email" not in rekord:  # gdyby kolumny kiedys zmienily kolejnosc
            mail = maile.get(w.get("place_id")) or None
            rekord["email"] = mail
            rekord["email_pewnosc"] = pewnosc_maila(mail, w.get("strona_www"))
        scalone.append(rekord)
    return scalone


def zapisz_scalony_csv(rekordy: list[dict], sciezka: Path):
    """Te same dane co w Excelu, ale w CSV - do importu gdzie indziej."""
    if not rekordy:
        return
    import csv as _csv

    with sciezka.open("w", newline="", encoding="utf-8-sig") as f:
        writer = _csv.DictWriter(f, fieldnames=list(rekordy[0].keys()), extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rekordy)


def zapisz_excel(rekordy: list[dict], sciezka: Path):
    if not rekordy:
        print("Brak danych do zapisania w Excelu.")
        return

    kolumny = list(rekordy[0].keys())
    wb = Workbook()
    ws = wb.active
    ws.title = "Wizytowki"

    czcionka = Font(name="Arial", size=10)
    naglowek_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    naglowek_tlo = PatternFill("solid", fgColor="1F4E79")

    for nr, pole in enumerate(kolumny, start=1):
        kom = ws.cell(row=1, column=nr, value=NAGLOWKI.get(pole, pole))
        kom.font = naglowek_font
        kom.fill = naglowek_tlo
        kom.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(nr)].width = SZEROKOSCI.get(pole, 20)

    for wiersz, rekord in enumerate(rekordy, start=2):
        for nr, pole in enumerate(kolumny, start=1):
            wartosc = rekord.get(pole)
            # ocena i liczba opinii jako liczby - zeby dalo sie sortowac
            if pole == "ocena" and wartosc:
                try:
                    wartosc = float(wartosc)
                except (TypeError, ValueError):
                    pass
            elif pole == "liczba_opinii" and wartosc:
                try:
                    wartosc = int(wartosc)
                except (TypeError, ValueError):
                    pass
            kom = ws.cell(row=wiersz, column=nr, value=wartosc)
            kom.font = czcionka
            kom.alignment = Alignment(vertical="top", wrap_text=pole in ZAWIJANE)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 28
    wb.save(sciezka)


def uruchom(linki: list[str] | None = None, szukaj_takze_maili: bool = True) -> dict:
    """Caly proces jako jedna funkcja - uzywa jej i wiersz polecen, i dashboard.

    linki=None oznacza tryb "tylko maile": wizytowki brane sa z gotowego
    pliku, bez ponownego odpytywania Google.

    Zwraca statystyki i sciezki do plikow wynikowych."""
    zrob_kopie_zapasowe()

    if linki is None:
        if not scraper.PLIK_JSON.exists():
            raise FileNotFoundError(
                f"Brak {scraper.PLIK_JSON.name} - najpierw uruchom pelny przebieg."
            )
        wizytowki = json.loads(scraper.PLIK_JSON.read_text(encoding="utf-8"))
        print(f"Tryb tylko-maile: wczytano {len(wizytowki)} wizytowek z pliku.\n")
    else:
        wizytowki = etap_wizytowki(linki)

    maile = etap_maile(wizytowki) if szukaj_takze_maili else {}
    if not szukaj_takze_maili:
        print("Etap maili pominiety na zyczenie.\n")

    scalone = scal(wizytowki, maile)
    PLIK_SCALONY.write_text(
        json.dumps(scalone, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    zapisz_scalony_csv(scalone, PLIK_SCALONY_CSV)
    zapisz_excel(scalone, PLIK_EXCEL)

    staty = {
        "wizytowek": len(scalone),
        "pobranych": sum(1 for w in scalone if w.get("nazwa")),
        "ze_strona": sum(1 for w in scalone if w.get("strona_www")),
        "z_mailem": sum(1 for w in scalone if w.get("email")),
        "do_sprawdzenia": sum(
            1 for w in scalone if (w.get("email_pewnosc") or "").startswith("inna domena")
        ),
        "pliki": [PLIK_EXCEL, PLIK_SCALONY_CSV, PLIK_SCALONY, PLIK_MAILE],
    }

    print("=" * 64)
    print("GOTOWE")
    print("=" * 64)
    print(f"  Wizytowek pobranych : {staty['pobranych']}/{staty['wizytowek']}")
    print(f"  Z adresem WWW       : {staty['ze_strona']}")
    print(f"  Z adresem e-mail    : {staty['z_mailem']}")
    if staty["do_sprawdzenia"]:
        print(f"  Maili do sprawdzenia: {staty['do_sprawdzenia']}")
    print()
    for p in staty["pliki"]:
        print(f"  {p.name}")
    return staty


def main():
    # --tylko-maile: pomija etap 1 i bierze wizytowki z gotowego pliku.
    tylko_maile = "--tylko-maile" in sys.argv
    if tylko_maile:
        sys.argv = [a for a in sys.argv if a != "--tylko-maile"]

    try:
        uruchom(None if tylko_maile else scraper.wczytaj_linki())
    except FileNotFoundError as e:
        print(e)
        sys.exit(1)


if __name__ == "__main__":
    main()
