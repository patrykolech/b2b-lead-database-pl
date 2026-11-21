"""
Skrypt pobiera dane z wizytówki Google (panel firmy) dla linków typu:
    https://local.google.com/place?id=10043757187928880068&use=srp

Taki link przekierowuje na wyniki Google z rozwiniętym panelem wizytówki,
z którego wyciągamy komplet danych.

Linki wklejasz do pliku linki.txt (jeden na linię, # to komentarz).
Wyniki zapisywane są do wizytowki.csv oraz wizytowki.json.

Captcha: skrypt otwiera WIDOCZNE okno przeglądarki. Jeśli Google pokaże
captchę, rozwiązujesz ją ręcznie w tym oknie - skrypt czeka i sam rusza dalej.

Wymagania:
    pip install playwright
    playwright install chromium

Projekt do nauki - celowo praca po kolei w jednym oknie, żeby nie
generować ruchu, który Google od razu zablokuje.
"""

import csv
import json
import random
import re
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from wspolne import (
    CAPTCHA,
    MAX_ROTACJI_IP,
    PROXY_HOST,
    PRZERWA_MIEDZY_LINKAMI,
    UZYWAJ_PROXY,
    ostatnia_deska_ratunku,
    otworz_przegladarke,
    wymaga_reki,
)

PLIK_Z_LINKAMI = Path(__file__).with_name("linki.txt")
PLIK_CSV = Path(__file__).with_name("wizytowki.csv")
PLIK_JSON = Path(__file__).with_name("wizytowki.json")


KOLUMNY = [
    "nazwa",
    "kategoria",
    "ocena",
    "liczba_opinii",
    "adres",
    "telefon",
    "strona_www",
    "status_otwarcia",
    "godziny_otwarcia",
    "opis",
    "profile_spolecznosciowe",
    "place_id",
    "cid_hex",
    "link",
    "blad",
]

# Kod JS czytający panel wizytówki bezpośrednio z DOM.
# Google używa losowych nazw klas (Aq14fc, E5BaQ...), ale atrybuty data-attrid
# są stabilne - dlatego opieramy się głównie na nich.
JS_WYCIAGNIJ = r"""
() => {
    const tekst = (el) => (el ? el.innerText.trim() : null);
    const attrid = (nazwa) => document.querySelector(`[data-attrid="${nazwa}"]`);

    // Pola typu "Adres: Katowicka 11/U12, 61-131 Poznań" - obcinamy etykietę
    const poDwukropku = (el) => {
        const t = tekst(el);
        if (!t) return null;
        const i = t.indexOf(":");
        return i === -1 ? t : t.slice(i + 1).trim();
    };

    const dane = {};

    dane.nazwa = tekst(attrid("title"));

    // Podtytuł: ocena + link "46 opinii" + kategoria
    const podtytul = attrid("subtitle");
    if (podtytul) {
        const ocenaEl = podtytul.querySelector("span.Aq14fc");
        if (ocenaEl) {
            dane.ocena = ocenaEl.innerText.trim().replace(",", ".");
        } else {
            // zapasowo: aria-label typu "Ocena: 4,8 na 5,"
            const aria = podtytul.querySelector('[aria-label*="cena"]');
            const m = aria && aria.getAttribute("aria-label").match(/(\d+[.,]\d+)/);
            if (m) dane.ocena = m[1].replace(",", ".");
        }

        const linkOpinii = [...podtytul.querySelectorAll("a")].find((a) =>
            /opini|recenz|review/i.test(a.innerText)
        );
        if (linkOpinii) {
            const m = linkOpinii.innerText.match(/([\d\s ]+)/);
            if (m) dane.liczba_opinii = m[1].replace(/[\s ]/g, "");
        }

        // kategoria to ostatni span z tekstem (po znaczniku <br>)
        const spany = [...podtytul.children].filter(
            (c) => c.tagName === "SPAN" && c.innerText.trim() && c !== ocenaEl
        );
        if (spany.length) dane.kategoria = spany[spany.length - 1].innerText.trim();
    }

    dane.adres = poDwukropku(attrid("kc:/location/location:address"));
    dane.telefon = poDwukropku(
        attrid("kc:/local:alt phone") || attrid("kc:/collection/knowledge_panels/has_phone:phone")
    );
    dane.status_otwarcia = poDwukropku(attrid("kc:/location/location:hours"));

    // Strona WWW - w pasku akcji panelu (Strona / Trasa / Zadzwoń...)
    const akcje = attrid("kc:/local:unified_actions");
    if (akcje) {
        const www = [...akcje.querySelectorAll("a")].find(
            (a) => /^https?:/.test(a.href) && !/google\./.test(new URL(a.href).hostname)
        );
        if (www) dane.strona_www = www.href;

        const tel = akcje.querySelector('a[href^="tel:"]');
        if (tel && !dane.telefon) dane.telefon = tel.href.replace("tel:", "");
    }

    const opis = attrid("kc:/local:merchant_description");
    if (opis) {
        // pierwsza linia to "Dodane przez: <nazwa>", właściwy opis jest niżej
        const linie = opis.innerText.split("\n").map((l) => l.trim()).filter(Boolean);
        dane.opis = (linie.length > 1 ? linie.slice(1).join(" ") : linie[0] || "")
            .replace(/^"|"$/g, "")
            .trim() || null;
    }

    const social = document.querySelector('[data-attrid*="social media presence"]');
    if (social) {
        const profile = [...social.querySelectorAll("a")]
            .map((a) => a.href)
            .filter((h) => /^https?:/.test(h));
        if (profile.length) dane.profile_spolecznosciowe = [...new Set(profile)].join(" | ");
    }

    // CID wizytówki w formacie szesnastkowym (przydatny identyfikator miejsca)
    const feature = document.querySelector("[data-feature-id]");
    if (feature) dane.cid_hex = feature.getAttribute("data-feature-id");

    return dane;
}
"""


def wczytaj_linki() -> list[str]:
    # można podać inny plik z linkami: python scraper_wizytowki.py moje_linki.txt
    plik = Path(sys.argv[1]) if len(sys.argv) > 1 else PLIK_Z_LINKAMI
    if not plik.exists():
        print(f"Brak pliku {plik.name} - utwórz go i wklej linki (jeden na linię).")
        sys.exit(1)

    linki = []
    for linia in plik.read_text(encoding="utf-8").splitlines():
        linia = linia.strip()
        if linia and not linia.startswith("#"):
            linki.append(linia)

    if not linki:
        print(f"Plik {plik.name} nie zawiera żadnych linków.")
        sys.exit(1)

    return linki


def wyciagnij_godziny(html: str) -> str | None:
    """Pełny tydzień godzin otwarcia siedzi w tabeli table.WgFkxc, którą Google
    dokleja do strony dopiero po rozwinięciu sekcji "Godziny". Sam kod HTML
    tabeli jest jednak od razu w źródle strony (w bloku jsl.dh), więc czytamy
    go regexem - bez klikania w rozwijacz."""
    tabela = re.search(r"<table class=\\?\"?WgFkxc.*?</table>", html, re.S)
    if not tabela:
        return None

    fragment = tabela.group(0).replace("\\x3c", "<").replace("\\x3e", ">").replace('\\"', '"')
    wiersze = re.findall(r"<tr[^>]*>(.*?)</tr>", fragment, re.S)

    dni = []
    for wiersz in wiersze:
        komorki = [
            re.sub(r"<[^>]+>", "", k).strip()
            for k in re.findall(r"<td[^>]*>(.*?)</td>", wiersz, re.S)
        ]
        if len(komorki) >= 2:
            dni.append(f"{komorki[0]}: {komorki[1]}")

    return "; ".join(dni) if dni else None


def place_id_z_linku(link: str) -> str | None:
    m = re.search(r"[?&]id=(\d+)", link)
    return m.group(1) if m else None


def pusty_rekord(link: str) -> dict:
    rekord = {kol: None for kol in KOLUMNY}
    rekord["link"] = link
    rekord["place_id"] = place_id_z_linku(link)
    return rekord


def pobierz_wizytowke(strona, link: str) -> dict | str:
    """Zwraca komplet danych albo stałą CAPTCHA, gdy trzeba zmienić IP."""
    rekord = pusty_rekord(link)

    # Jedna ponowna proba - lacza rezydencjalne bywaja kapryśne i pojedynczy
    # timeout nie znaczy jeszcze, ze link jest nie do pobrania.
    ostatni_blad = None
    for podejscie in range(2):
        try:
            strona.goto(link, timeout=30000, wait_until="domcontentloaded")
            ostatni_blad = None
            break
        except Exception as e:
            ostatni_blad = e
            if podejscie == 0:
                time.sleep(2)

    if ostatni_blad is not None:
        # Wazne: NIE zwracamy tu CAPTCHA. Wczesniej kazdy timeout udawal
        # captche, przez co w logach widac bylo "captcha" tam, gdzie
        # naprawde strona w ogole sie nie wczytala - i szukalo sie nie tam,
        # gdzie trzeba.
        rekord["blad"] = f"nie udało się wczytać strony: {type(ostatni_blad).__name__}"
        return rekord

    strona.wait_for_timeout(2000)

    if wymaga_reki(strona):
        return CAPTCHA

    # panel wizytówki ładuje się chwilę po samej stronie
    try:
        strona.wait_for_selector('[data-attrid="title"]', timeout=15000)
    except Exception:
        rekord["blad"] = "nie znaleziono panelu wizytówki"
        return rekord

    dane = strona.evaluate(JS_WYCIAGNIJ)
    rekord.update({k: v for k, v in dane.items() if v})
    rekord["godziny_otwarcia"] = wyciagnij_godziny(strona.content())
    return rekord


def zapisz_wyniki(wyniki: list[dict]):
    with open(PLIK_CSV, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=KOLUMNY, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(wyniki)

    PLIK_JSON.write_text(
        json.dumps(wyniki, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def przetworz_linki(linki: list[str]) -> list[dict]:
    """Pobiera wizytówki dla listy linków i zwraca gotowe rekordy.

    Wydzielone z main(), żeby uruchom_wszystko.py mogło wywołać sam etap
    pobierania, bez zapisu do plików i bez własnego wczytywania linków."""
    print(f"Do pobrania: {len(linki)} wizytówek.")
    if UZYWAJ_PROXY:
        print(f"Proxy: {PROXY_HOST} (przy captchy zmieniam IP, max {MAX_ROTACJI_IP} razy na link).")
    else:
        print("Proxy: wyłączone - łączę się z własnego IP (tak działa najlepiej).")
    print()

    wyniki = []
    zmiany_ip = 0
    # bez proxy zmiana sesji nic nie daje - IP zostaje to samo
    podejscia = MAX_ROTACJI_IP if UZYWAJ_PROXY else 1

    with sync_playwright() as p:
        kontekst, strona, sesja = otworz_przegladarke(p)
        if UZYWAJ_PROXY:
            print(f"IP sesji: {sesja}")
        print()

        try:
            for nr, link in enumerate(linki, start=1):
                print(f"[{nr}/{len(linki)}] {link}")
                dane = None

                for podejscie in range(1, podejscia + 1):
                    try:
                        wynik = pobierz_wizytowke(strona, link)
                    except Exception as e:
                        wynik = pusty_rekord(link)
                        wynik["blad"] = str(e)

                    if wynik != CAPTCHA:
                        dane = wynik
                        break

                    # spalone IP - zamykamy przeglądarkę i wracamy na nowym adresie
                    if podejscie < podejscia:
                        print(f"    captcha - zmieniam IP (podejście {podejscie}/{podejscia})")
                        kontekst.close()
                        kontekst, strona, sesja = otworz_przegladarke(p)
                        zmiany_ip += 1
                        time.sleep(random.uniform(1, 2))

                if dane is None:
                    # rotacje wyczerpane - zostaje 2captcha albo ręczne kliknięcie
                    print("    Captcha - trzeba ją rozwiązać.")
                    if ostatnia_deska_ratunku(strona, link):
                        try:
                            wynik = pobierz_wizytowke(strona, link)
                            dane = wynik if wynik != CAPTCHA else None
                        except Exception as e:
                            dane = pusty_rekord(link)
                            dane["blad"] = str(e)
                    if dane is None:
                        dane = pusty_rekord(link)
                        dane["blad"] = "captcha - nie udało się obejść"

                wyniki.append(dane)
                print(f"    -> {dane.get('nazwa') or 'BRAK DANYCH'}  |  {dane.get('adres') or '-'}")
                if dane.get("blad"):
                    print(f"    -> błąd: {dane['blad']}")

                if nr < len(linki):
                    time.sleep(random.uniform(*PRZERWA_MIEDZY_LINKAMI))
        finally:
            try:
                kontekst.close()
            except Exception:
                pass

    if UZYWAJ_PROXY:
        print(f"\n(zmian IP: {zmiany_ip})")
    return wyniki


def main():
    wyniki = przetworz_linki(wczytaj_linki())
    zapisz_wyniki(wyniki)
    udane = sum(1 for w in wyniki if w.get("nazwa"))
    print(f"\nGotowe: {udane}/{len(wyniki)} wizytówek pobranych.")
    print(f"Wyniki: {PLIK_CSV.name} oraz {PLIK_JSON.name}")


if __name__ == "__main__":
    main()
