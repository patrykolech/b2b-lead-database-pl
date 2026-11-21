"""
Finder e-maili z wynikow wyszukiwania Google.

Jak dziala:
1. Wczytuje liste stron/domen z pliku wejsciowego (jedna linia = jedna pozycja).
2. Dla kazdej wpisuje w Google zapytanie: "<pozycja> email".
3. Wyciaga adres e-mail widoczny bezposrednio we fragmentach wynikow
   wyszukiwania - bez wchodzenia na strone firmy.
4. Gdy Google pokaze captche, probuje rozwiazac ja automatycznie przez
   2captcha, a gdy sie nie uda - czeka na reczne rozwiazanie w oknie.
5. Zapisuje wyniki (strona -> e-mail) do pliku CSV.

Uzycie:
    python email_finder.py input.txt output.csv

Plik wejsciowy - jedna strona/domena w linii, np.:
    firma-przyklad.pl
    inna-firma.pl

Ten skrypt jest tez uzywany przez uruchom_wszystko.py, ktore samo podaje
liste pozycji (funkcja szukaj_maili) - wtedy plik wejsciowy nie jest potrzebny.

Ograniczenie: pomyslane do malej skali (kilkanascie zapytan), nie do
masowego scrapowania Google.
"""

import random
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote_plus, urlparse

from playwright.sync_api import sync_playwright

from wspolne import ostatnia_deska_ratunku, otworz_przegladarke, wymaga_reki

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")

# Regex powyzej matchuje tez nazwy plikow w stylu "logo@2x.png" albo
# "ikona@3x.jpg" - to nie sa adresy, tylko grafiki z retiny. Odrzucamy je
# po rozszerzeniu.
ROZSZERZENIA_PLIKOW = (
    ".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp",
    ".css", ".js", ".ico", ".bmp", ".woff", ".woff2",
)

# Domeny techniczne - pojawiaja sie w kodzie stron i w stopkach Google,
# ale nigdy nie sa adresem kontaktowym firmy.
ODRZUCANE_DOMENY = (
    "sentry.io", "example.com", "wixpress.com", "google.com", "gstatic.com",
    "googleapis.com", "schema.org", "w3.org", "youtube.com", "facebook.com",
    "sentry-cdn.com", "cloudflare.com", "jquery.com",
)

# Agencje i hostingi, ktore robia strony firmom - ich adres jest w stopce
# cudzej strony i latwo wziac go za adres firmy. Realny przypadek z testu:
# dla salon-przyklad.pl skrypt wyciagal kontakt@hosting-przyklad.pl.
ODRZUCANE_AGENCJE = (
    "wenet.pl", "home.pl", "nazwa.pl", "ovh.pl", "cyberfolks.pl", "dhosting.pl",
    "wix.com", "squarespace.com", "godaddy.com", "wordpress.com", "shoper.pl",
    "netadmin.pl", "seo.pl", "webd.pl", "zenbox.pl",
)

# Serwisy, ktore nie mowia nic o firmie - jesli wizytowka ma taki "adres www",
# lepiej szukac po nazwie firmy niz po tej domenie.
AGREGATORY = (
    "facebook.com", "instagram.com", "znanylekarz.pl", "gowork.pl",
    "swiatprzychodni.pl", "booksy.com", "targeo.pl", "panoramafirm.pl",
    "pkt.pl", "aleo.com", "linkedin.com", "twitter.com", "x.com",
    "mapy.cz", "yelp.com", "oferteo.pl",
)

PRZERWA_MIEDZY_ZAPYTANIAMI = (4, 8)  # sekundy


def domena_z_url(adres: str) -> str | None:
    """Wyciaga sama domene: https://www.firma-przyklad.pl/kontakt -> firma-przyklad.pl"""
    if not adres:
        return None
    adres = adres.strip()
    if "://" not in adres:
        adres = "http://" + adres
    host = (urlparse(adres).hostname or "").lower()
    return host[4:] if host.startswith("www.") else host or None


def czy_agregator(domena: str | None) -> bool:
    return bool(domena) and any(a in domena for a in AGREGATORY)


def zapytanie_dla(strona_www: str | None, nazwa: str | None) -> tuple[str, str | None]:
    """Zwraca (fraza do wyszukania, domena firmy albo None).

    Dla zwyklej strony firmowej szukamy po domenie - to najcelniejsze.
    Dla Facebooka czy ZnanyLekarz domena nic nie mowi o firmie, wiec
    szukamy po nazwie."""
    domena = domena_z_url(strona_www)
    if domena and not czy_agregator(domena):
        return domena, domena
    if nazwa:
        return nazwa.strip(), None
    return (domena or ""), domena


def wybierz_email(tekst: str, domena: str | None, obce_domeny: set[str] | None = None) -> str | None:
    """Wybiera najlepszy adres z tekstu strony wynikow.

    Najpierw szuka adresu w domenie firmy (kontakt@firma-przyklad.pl dla
    firma-przyklad.pl). Gdy takiego nie ma, bierze pierwszy sensowny - ale
    z dwoma zabezpieczeniami, bo bez nich lapal cudze adresy:

    - obce_domeny: adresy INNYCH firm z tej samej listy. Realny przypadek:
      dla gabinet-przyklad.pl wychodzil kontakt@portal-przyklad.pl, czyli
      mail sasiedniego gabinetu z wynikow wyszukiwania.
    - ODRZUCANE_AGENCJE: adresy firm robiacych strony (np. wenet.pl),
      ktore siedza w stopce cudzej witryny."""
    obce_domeny = obce_domeny or set()
    znalezione, widziane = [], set()

    for adres in EMAIL_RE.findall(tekst):
        maly = adres.lower()
        if maly in widziane:
            continue
        widziane.add(maly)
        if maly.endswith(ROZSZERZENIA_PLIKOW):
            continue
        if any(maly.endswith(d) or f"@{d}" in maly for d in ODRZUCANE_DOMENY):
            continue
        znalezione.append(adres)

    if not znalezione:
        return None

    # 1. najlepiej: adres w domenie szukanej firmy
    if domena:
        for adres in znalezione:
            maly = adres.lower()
            if maly.endswith("@" + domena) or maly.endswith("." + domena):
                return adres

    # 2. zapasowo: pierwszy, ktory nie nalezy do innej firmy z listy
    #    ani do agencji od stron internetowych
    for adres in znalezione:
        maly = adres.lower()
        po_malpie = maly.split("@")[-1]
        if any(po_malpie == o or po_malpie.endswith("." + o) for o in obce_domeny):
            continue
        if any(po_malpie == a or po_malpie.endswith("." + a) for a in ODRZUCANE_AGENCJE):
            continue
        return adres

    return None


def szukaj_maili(strona, pozycje: list[str], domeny: dict[str, str | None] | None = None) -> dict[str, str]:
    """Dla kazdej pozycji wyszukuje w Google mail. Zwraca {pozycja: email}.

    pozycje - frazy do wpisania w Google (domena albo nazwa firmy),
    domeny  - opcjonalna mapa fraza -> domena firmy, uzywana do wyboru
              najtrafniejszego adresu."""
    domeny = domeny or {}
    wyniki: dict[str, str] = {}

    # domeny wszystkich firm z listy - posluza do odrzucania cudzych adresow
    wszystkie_domeny = {d for d in domeny.values() if d}

    for nr, pozycja in enumerate(pozycje, start=1):
        fraza = f"{pozycja} email"
        url = f"https://www.google.pl/search?q={quote_plus(fraza)}&hl=pl"
        print(f"[{nr}/{len(pozycje)}] Szukam: {fraza}")

        try:
            strona.goto(url, timeout=30000, wait_until="domcontentloaded")
            strona.wait_for_timeout(2000)
        except Exception as e:
            print(f"    -> blad wczytywania: {e}")
            wyniki[pozycja] = ""
            continue

        if wymaga_reki(strona):
            if not ostatnia_deska_ratunku(strona, url):
                print("    -> captcha nierozwiazana, pomijam")
                wyniki[pozycja] = ""
                continue

        try:
            tekst = strona.inner_text("body")
        except Exception:
            tekst = ""

        moja = domeny.get(pozycja)
        obce = wszystkie_domeny - {moja} if moja else wszystkie_domeny
        email = wybierz_email(tekst, moja, obce)
        wyniki[pozycja] = email or ""
        print(f"    -> {email if email else 'nie znaleziono adresu'}")

        if nr < len(pozycje):
            time.sleep(random.uniform(*PRZERWA_MIEDZY_ZAPYTANIAMI))

    return wyniki


def zapisz_csv(wyniki: dict[str, str], sciezka: Path):
    with sciezka.open("w", encoding="utf-8", newline="") as f:
        f.write("strona,email\n")
        for pozycja, email in wyniki.items():
            bezpieczna = pozycja.replace('"', '""')
            f.write(f'"{bezpieczna}","{email}"\n')
    print(f"\nZapisano wyniki do: {sciezka}")


def main():
    if len(sys.argv) < 3:
        print("Uzycie: python email_finder.py <plik_wejsciowy.txt> <plik_wyjsciowy.csv>")
        sys.exit(1)

    plik_we, plik_wy = Path(sys.argv[1]), Path(sys.argv[2])

    if not plik_we.exists():
        print(f"Nie znaleziono pliku wejsciowego: {plik_we}")
        sys.exit(1)

    pozycje = [
        l.strip()
        for l in plik_we.read_text(encoding="utf-8").splitlines()
        if l.strip() and not l.strip().startswith("#")
    ]
    if not pozycje:
        print("Plik wejsciowy jest pusty.")
        sys.exit(1)

    print(f"Wczytano {len(pozycje)} pozycji do wyszukania.\n")

    # w trybie plikowym pozycje to zwykle domeny - uzywamy ich do wyboru
    # najtrafniejszego adresu
    domeny = {p: domena_z_url(p) for p in pozycje}

    # ustawiamy z gory, zeby finally mialo co zapisac nawet po wyjatku
    wyniki: dict[str, str] = {}

    with sync_playwright() as p:
        kontekst, strona, _ = otworz_przegladarke(p)
        try:
            wyniki = szukaj_maili(strona, pozycje, domeny)
        finally:
            zapisz_csv(wyniki, plik_wy)
            kontekst.close()

    znalezione = sum(1 for e in wyniki.values() if e)
    print(f"Znaleziono {znalezione}/{len(pozycje)} adresow.")


if __name__ == "__main__":
    main()
