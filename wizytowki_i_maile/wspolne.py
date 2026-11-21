"""
Wspolna warstwa dla obu skryptow: przegladarka, proxy i obsluga captchy.

Trzyma wszystko, co scraper wizytowek i finder maili robia tak samo:
- otwieranie przegladarki na trwalym profilu (z ciasteczkami Google),
- wykrywanie captchy / ekranu zgod,
- automatyczne rozwiazywanie captchy przez 2captcha,
- reczne rozwiazanie jako ostatnia deska ratunku.

WAZNE: automatyczne rozwiazywanie captchy Google dziala TYLKO dzieki
ciasteczkom (NID, SOCS, AEC...) z katalogu profil_przegladarki. Bez nich
Google odrzuca token, nawet poprawnie rozwiazany. Nie kasuj tego katalogu -
a jesli juz skasujesz, przejdz raz captche recznie, zeby profil je zapamietal.
"""

import os
import random
import time
from pathlib import Path

import requests

# Czy pokazywac okno przegladarki.
# Lokalnie: TAK - inaczej nie da sie recznie kliknac captchy.
# Na serwerze bez ekranu: ustaw zmienna srodowiskowa UKRYJ_OKNO=1,
# wtedy captcha moze byc rozwiazana wylacznie przez 2captcha.
POKAZUJ_OKNO = os.environ.get("UKRYJ_OKNO") != "1"

# Ile sekund czekac, az czlowiek rozwiaze captche w otwartym oknie, zanim
# uznamy link za stracony. Skrypt sam wykrywa moment rozwiazania.
MAX_CZEKANIA_NA_RECZNA = int(os.environ.get("CZEKAJ_NA_CAPTCHE", "300"))

# Trwały profil przeglądarki. Dzięki niemu ciasteczka (zgody RODO, rozwiązana
# captcha) zostają zapamiętane i przy kolejnych uruchomieniach Google zwykle
# nie pyta o nic ponownie.
PROFIL_PRZEGLADARKI = Path(__file__).with_name("profil_przegladarki")

PRZERWA_MIEDZY_LINKAMI = (4, 8)  # losowa pauza w sekundach

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# --- Proxy rezydencjalne 2captcha ---------------------------------------
# DOMYŚLNIE WYŁĄCZONE - i to nie przez lenistwo, tylko na podstawie testów:
#   * z własnego IP:  32/32 wizytówki pobrane bez problemu,
#   * przez to proxy: captcha przy KAŻDYM podejściu (5/5, potem 2/2).
# IP w puli rezydencjalnej są współdzielone i Google ma je już oflagowane,
# więc proxy tutaj szkodzi zamiast pomagać. Włącz tylko, jeśli chcesz
# eksperymentować albo Twoje własne IP zostanie zablokowane.
#
# Uwaga na przyszłość: to proxy domyślnie ROTUJE (każde żądanie z innego IP),
# co dla przeglądarki jest zabójcze - jedno wejście na stronę to kilkadziesiąt
# żądań. Dlatego login ma doklejone -session-<id>-sessTime-<minuty>, które
# przypina jedno IP. Składnia sprawdzona empirycznie i działa.
UZYWAJ_PROXY = bool(os.environ.get("PROXY_HOST"))

# Stale proxy rezydencjalne (jeden dedykowany adres, nikt inny go nie uzywa).
# Zmiana adresu proxy = wyczysc profil przegladarki (ciasteczka sa
# zwiazane z IP).
PROXY_HOST = os.environ.get("PROXY_HOST", "")  # np. host:port
PROXY_LOGIN_BAZOWY = os.environ.get("PROXY_LOGIN", "")
PROXY_HASLO = os.environ.get("PROXY_PASSWORD", "")

# Przy stalym IP login idzie goly - bez dopisku -session-...-sessTime-...
# Tamten sufiks to skladnia puli rotujacej 2captcha; inny dostawca go nie
# rozumie i odrzuca logowanie. Ustaw False tylko wracajac na pule rotujaca.
PROXY_STALE_IP = True
PROXY_MINUTY_SESJI = 30  # dotyczy wylacznie puli rotujacej (max 120)

# Ile razy zmienic IP, zanim siegniemy po 2captcha. Przy stalym IP nie ma
# czego zmieniac - kazde "nowe" polaczenie i tak wyjdzie z tego samego
# adresu - wiec rotacja jest bezcelowa i pomijamy ja.
MAX_ROTACJI_IP = 1 if PROXY_STALE_IP else 3

# --- 2captcha ------------------------------------------------------------
UZYWAJ_2CAPTCHA = True
TWOCAPTCHA_API_KEY = os.environ.get("TWOCAPTCHA_API_KEY", "")
MAX_CZEKANIA_NA_2CAPTCHA = 180  # sekund na rozwiązanie jednego zadania

# Adres, który podajemy 2captcha jako stronę z captchą. Dla captchy z
# wyszukiwarki Google podaje się adres wyszukiwarki, a nie /sorry/index.
URL_STRONY_CAPTCHY = "https://www.google.com/search"

# Ile razy zlecić captchę, zanim odpuścimy. Przy każdej kolejnej próbie
# strona jest pobierana od nowa, żeby data-s było świeże.
MAX_PROB_2CAPTCHA = 2

# Ile ms czekać po wstawieniu tokenu - Google potrzebuje czasu na
# przetworzenie odpowiedzi i przekierowanie.
PAUZA_PO_TOKENIE = 10000

# Strona /sorry/ ładuje recaptcha/enterprise.js, więc kod domyślnie zleca
# zadanie jako Enterprise. Ustaw True, żeby wymusić zwykłe v2 (tak robi
# przykład z poradnika 2captcha dla wyszukiwarki Google).
WYMUS_ZWYKLE_V2 = False

# Zlecając zadanie podajemy 2captcha nasze proxy, żeby pracownik rozwiązywał
# captchę z tego samego IP, z którego potem wysyłamy token.
PRZEKAZUJ_PROXY_DO_2CAPTCHA = True

# Aktualny login proxy - podmieniany przy każdej rotacji IP
PROXY_LOGIN = ""


def nowa_sesja_proxy() -> str:
    """Ustawia login proxy na kolejne polaczenie.

    Przy stalym IP login jest zawsze ten sam. Przy puli rotujacej doklejamy
    losowy identyfikator sesji, ktory przypina jedno IP na kilkadziesiat minut."""
    global PROXY_LOGIN

    if PROXY_STALE_IP:
        PROXY_LOGIN = PROXY_LOGIN_BAZOWY
        return "stale IP"

    sesja = f"s{random.randint(10**7, 10**8 - 1)}"
    PROXY_LOGIN = f"{PROXY_LOGIN_BAZOWY}-session-{sesja}-sessTime-{PROXY_MINUTY_SESJI}"
    return sesja


def proxy_dla_playwright() -> dict | None:
    if not UZYWAJ_PROXY:
        return None
    return {
        "server": f"http://{PROXY_HOST}",
        "username": PROXY_LOGIN,
        "password": PROXY_HASLO,
    }


def proxy_dla_2captcha() -> str | None:
    """Format wymagany przez 2captcha: login:haslo@host:port"""
    if not (UZYWAJ_PROXY and PRZEKAZUJ_PROXY_DO_2CAPTCHA):
        return None
    return f"{PROXY_LOGIN}:{PROXY_HASLO}@{PROXY_HOST}"

def wymaga_reki(strona) -> bool:
    """Czy strona to captcha albo ekran zgód RODO - czyli coś, co człowiek
    musi kliknąć sam, zanim skrypt zobaczy wizytówkę."""
    try:
        if "/sorry/" in strona.url or "consent.google" in strona.url:
            return True
        tresc = strona.content()
        return "g-recaptcha" in tresc or "recaptcha/api" in tresc
    except Exception:
        return False


def parametry_captchy(strona) -> dict | None:
    """Wyciąga ze strony wszystko, czego 2captcha potrzebuje.

    Szukamy konkretnie <div id="recaptcha">, a nie pierwszego lepszego
    elementu z data-sitekey - na stronie potrafi być więcej takich elementów
    i łatwo złapać nie ten, co trzeba (objaw: sitekey inny niż w źródle).

    data-s to token generowany osobno dla KAŻDEJ captchy. Musi być świeży -
    przy ponownej próbie trzeba pobrać nową stronę, a nie użyć starego."""
    try:
        return strona.evaluate(
            """() => {
                const html = document.documentElement.innerHTML;
                const el = document.querySelector('div[id="recaptcha"]')
                    || document.querySelector('.g-recaptcha[data-sitekey]')
                    || document.querySelector('[data-sitekey]');
                const sk = html.match(/data-sitekey="([^"]+)"/);
                const ds = html.match(/data-s="([^"]+)"/);
                const sitekey = (el && el.getAttribute('data-sitekey')) || (sk ? sk[1] : null);
                if (!sitekey) return null;
                return {
                    sitekey: sitekey,
                    data_s: (el && el.getAttribute('data-s')) || (ds ? ds[1] : null),
                    enterprise: html.includes('recaptcha/enterprise'),
                    invisible: (el && el.getAttribute('data-size') === 'invisible') || false,
                };
            }"""
        )
    except Exception:
        return None


def ciasteczka_google(strona) -> str:
    """Zbiera ciasteczka domeny google.* i skleja w format wymagany przez
    2captcha: "NID=...; ANID=...".

    To najważniejszy parametr przy captchy Google: bez ciasteczek Google widzi
    zapytanie bez sesji i odrzuca token, nawet jeśli captcha została poprawnie
    rozwiązana. Ciasteczka bierzemy z trwałego profilu przeglądarki, więc
    warto raz przejść captchę ręcznie - profil je zapamięta."""
    try:
        wszystkie = strona.context.cookies()
    except Exception:
        return ""

    google = [c for c in wszystkie if "google." in (c.get("domain") or "")]
    if not google:
        return ""

    # NID i ANID na początek - to one potwierdzają sesję
    wazne = [c for c in google if c["name"] in ("NID", "ANID", "AEC", "SOCS", "CONSENT")]
    reszta = [c for c in google if c not in wazne]
    return "; ".join(f"{c['name']}={c['value']}" for c in wazne + reszta)


def zlec_2captcha(parametry: dict, cookies: str) -> tuple[str | None, int | None]:
    """Zleca rozwiązanie captchy przez API v2. Zwraca (token, id_zadania).

    id_zadania jest potrzebne, żeby zgłosić zły token i odzyskać środki."""
    if not TWOCAPTCHA_API_KEY:
        return None, None

    enterprise = parametry.get("enterprise") and not WYMUS_ZWYKLE_V2
    proxy = proxy_dla_2captcha()

    zadanie = {
        "type": (
            ("RecaptchaV2EnterpriseTask" if enterprise else "RecaptchaV2Task")
            if proxy
            else ("RecaptchaV2EnterpriseTaskProxyless" if enterprise else "RecaptchaV2TaskProxyless")
        ),
        # Adres strony z captchą. Dla captchy z wyszukiwarki Google podaje się
        # adres wyszukiwarki, a nie /sorry/index - tak wymaga 2captcha.
        "websiteURL": URL_STRONY_CAPTCHY,
        "websiteKey": parametry["sitekey"],
        "isInvisible": bool(parametry.get("invisible")),
        "userAgent": USER_AGENT,
    }

    # data-s trafia w inne miejsce zależnie od wariantu captchy
    if parametry.get("data_s"):
        if enterprise:
            zadanie["enterprisePayload"] = {"s": parametry["data_s"]}
        else:
            zadanie["recaptchaDataSValue"] = parametry["data_s"]

    # bez ciasteczek Google odrzuca token, bo nie widzi sesji
    if cookies:
        zadanie["cookies"] = cookies

    if proxy:
        zadanie.update(
            {
                "proxyType": "http",
                "proxyAddress": PROXY_HOST.split(":")[0],
                "proxyPort": int(PROXY_HOST.split(":")[1]),
                "proxyLogin": PROXY_LOGIN,
                "proxyPassword": PROXY_HASLO,
            }
        )

    print(f"  2captcha: typ zadania {zadanie['type']}")

    try:
        odp = requests.post(
            "https://api.2captcha.com/createTask",
            json={"clientKey": TWOCAPTCHA_API_KEY, "task": zadanie},
            timeout=30,
        ).json()
    except Exception as e:
        print(f"  2captcha: nie udało się wysłać zadania: {e}")
        return None, None

    if odp.get("errorId"):
        print(f"  2captcha: zadanie odrzucone: {odp.get('errorCode')} {odp.get('errorDescription')}")
        return None, None

    id_zadania = odp.get("taskId")
    print(f"  2captcha: zadanie {id_zadania} przyjęte, czekam na rozwiązanie...")

    start = time.time()
    time.sleep(15)  # 2captcha prosi, żeby pierwszy raz pytać po ~15 s
    while time.time() - start < MAX_CZEKANIA_NA_2CAPTCHA:
        try:
            wynik = requests.post(
                "https://api.2captcha.com/getTaskResult",
                json={"clientKey": TWOCAPTCHA_API_KEY, "taskId": id_zadania},
                timeout=30,
            ).json()
        except Exception as e:
            print(f"  2captcha: błąd przy sprawdzaniu statusu: {e}")
            return None, id_zadania

        if wynik.get("errorId"):
            print(f"  2captcha: błąd: {wynik.get('errorCode')} {wynik.get('errorDescription')}")
            return None, id_zadania

        if wynik.get("status") == "ready":
            print(f"  2captcha: rozwiązane w {int(time.time() - start)} s (koszt {wynik.get('cost')}).")
            return wynik["solution"]["gRecaptchaResponse"], id_zadania

        time.sleep(5)

    print("  2captcha: przekroczono czas oczekiwania.")
    return None, id_zadania


def zglos_zly_token(id_zadania: int | None):
    """Zgłasza nieudany token - 2captcha zwraca wtedy środki za to zadanie."""
    if not id_zadania:
        return
    try:
        requests.post(
            "https://api.2captcha.com/reportIncorrect",
            json={"clientKey": TWOCAPTCHA_API_KEY, "taskId": id_zadania},
            timeout=20,
        )
        print(f"  2captcha: zgłoszono zły token (zadanie {id_zadania}) - środki wrócą.")
    except Exception:
        pass


def wstaw_token_i_wyslij(strona, token: str) -> bool:
    """Wkleja token do pola g-recaptcha-response i wysyła formularz captchy.

    Strona /sorry/ Google definiuje własny callback (submitCallback), który
    sam wysyła formularz - jeśli istnieje, wołamy właśnie jego, bo tak samo
    zachowałaby się prawdziwa przeglądarka po rozwiązaniu captchy."""
    try:
        return strona.evaluate(
            """(token) => {
                let pole = document.getElementById('g-recaptcha-response')
                    || document.querySelector('textarea[name="g-recaptcha-response"]');
                const formularz = document.querySelector('#captcha-form')
                    || (pole && pole.closest('form'))
                    || document.querySelector('form');
                if (!formularz) return false;

                // gdy reCAPTCHA nie zdążyła się załadować, pola może nie być
                if (!pole) {
                    pole = document.createElement('textarea');
                    pole.id = 'g-recaptcha-response';
                    pole.name = 'g-recaptcha-response';
                    formularz.appendChild(pole);
                }
                pole.style.display = 'block';
                pole.value = token;
                pole.innerHTML = token;

                const el = document.querySelector('[data-callback]');
                const nazwa = el && el.getAttribute('data-callback');
                if (nazwa && typeof window[nazwa] === 'function') {
                    window[nazwa](token);
                    return true;
                }
                formularz.submit();
                return true;
            }""",
            token,
        )
    except Exception as e:
        print(f"  Nie udało się wysłać formularza captchy: {e}")
        return False


def sprobuj_2captcha(strona, link: str) -> bool:
    """Rozwiązuje captchę przez 2captcha. True = captcha zniknęła.

    Gdy token zostanie odrzucony, NIE ponawiamy na tej samej stronie -
    trzeba pobrać ją od nowa, żeby dostać świeże data-s. Ponowne zlecanie
    ze starym data-s to pewne odrzucenie i ryzyko blokady IP."""
    for proba in range(1, MAX_PROB_2CAPTCHA + 1):
        parametry = parametry_captchy(strona)
        if not parametry or not parametry.get("sitekey"):
            print("  Nie znalazłem sitekey - automat odpada.")
            return False

        cookies = ciasteczka_google(strona)
        print(f"  Próba {proba}/{MAX_PROB_2CAPTCHA}, sitekey: {parametry['sitekey']}")
        print(
            f"  Wariant: {'Enterprise' if parametry.get('enterprise') and not WYMUS_ZWYKLE_V2 else 'zwykłe v2'}"
            f"{', invisible' if parametry.get('invisible') else ''}"
            f"{', z data-s' if parametry.get('data_s') else ', BEZ data-s'}"
        )
        if cookies:
            nazwy = [c.split("=")[0] for c in cookies.split("; ")]
            print(f"  Ciasteczka Google: {len(nazwy)} szt. ({', '.join(nazwy[:5])})")
        else:
            print("  UWAGA: brak ciasteczek google.com - Google prawie na pewno odrzuci token.")

        token, id_zadania = zlec_2captcha(parametry, cookies)
        if not token:
            return False

        if not wstaw_token_i_wyslij(strona, token):
            return False

        # Google potrzebuje chwili na przetworzenie odpowiedzi
        strona.wait_for_timeout(PAUZA_PO_TOKENIE)
        if not wymaga_reki(strona):
            return True

        print("  Token wysłany, ale captcha nadal widoczna - zgłaszam zły token.")
        zglos_zly_token(id_zadania)

        if proba < MAX_PROB_2CAPTCHA:
            # świeża strona = świeże data-s
            print("  Pobieram stronę od nowa, żeby dostać świeże data-s...")
            try:
                strona.goto(link, timeout=30000, wait_until="domcontentloaded")
                strona.wait_for_timeout(2500)
            except Exception:
                return False
            if not wymaga_reki(strona):
                return True

    return False


def ostatnia_deska_ratunku(strona, link: str) -> bool:
    """Najpierw 2captcha, a gdy nie pomoże - ręczne rozwiązanie w oknie."""
    if not wymaga_reki(strona):
        return True

    if UZYWAJ_2CAPTCHA and TWOCAPTCHA_API_KEY:
        print("  Próbuję rozwiązać przez 2captcha...")
        if sprobuj_2captcha(strona, link):
            print("  Captcha rozwiązana przez 2captcha.")
            strona.wait_for_timeout(1500)
            return True
        print("  2captcha nie pomogło.")

    # Bez widocznego okna (serwer, dashboard w trybie ukrytym) nie ma czego
    # kliknac - nie ma sensu blokowac przebiegu czekaniem na cos niewidocznego.
    if not POKAZUJ_OKNO:
        print("  Okno przegladarki jest ukryte - pomijam ten link.")
        return False

    print(f"  >>> Rozwiąż captchę w otwartym oknie przeglądarki.")
    print(f"  >>> Nie musisz nic klikać w konsoli - ruszę dalej sam.")
    print(f"  >>> Czekam maksymalnie {MAX_CZEKANIA_NA_RECZNA} s...")

    # Sami wykrywamy, kiedy captcha znika. Wczesniej bylo tu input() czekajace
    # na Enter, ale to wymaga terminala - a uruchomiony z dashboardu skrypt
    # zadnego terminala nie ma i przebieg wisial w nieskonczonosc.
    #
    # WAZNE: strona.url to wlasciwosc buforowana lokalnie przez Playwright -
    # nawet gdy okno zostanie zamkniete (przez uzytkownika albo awarie),
    # wymaga_reki() nadal widzi stary adres "/sorry/" i zwraca True, wiec
    # sama petla "while wymaga_reki(strona)" NIGDY by tego nie wykryla i
    # czekalaby bez sensu pelne MAX_CZEKANIA_NA_RECZNA sekund, konczac sie
    # mylacym komunikatem "nikt nie rozwiazal captchy na czas". Dlatego
    # zamkniecie okna sprawdzamy osobno, wprost, na poczatku kazdej iteracji.
    start = time.time()
    while wymaga_reki(strona):
        if strona.is_closed():
            print("  Okno przeglądarki zostało zamknięte - nie mogę kontynuować.")
            return False
        if time.time() - start > MAX_CZEKANIA_NA_RECZNA:
            print("  Nikt nie rozwiązał captchy na czas - pomijam ten link.")
            return False
        time.sleep(2)

    if strona.is_closed():
        print("  Okno przeglądarki zostało zamknięte - nie mogę kontynuować.")
        return False

    print(f"  Captcha rozwiązana po {int(time.time() - start)} s, kontynuuję.")
    strona.wait_for_timeout(1500)
    return True

# Sygnał dla pętli głównej: "to IP jest spalone, zmień je i spróbuj ponownie"
CAPTCHA = "__captcha__"

def otworz_przegladarke(p):
    """Otwiera przeglądarkę na świeżej sesji proxy (czyli nowym IP)."""
    sesja = nowa_sesja_proxy()
    kontekst = p.chromium.launch_persistent_context(
        user_data_dir=str(PROFIL_PRZEGLADARKI),
        headless=not POKAZUJ_OKNO,
        args=[
            "--window-size=1280,900",
            "--disable-gpu",
            # Bez tych dwoch flag Google NIE LADUJE SIE przez proxy - zwykly
            # timeout po 40 s, bez zadnego bledu. Zmierzone: z HTTP/2 strona
            # nie wstaje wcale, z wymuszonym HTTP/1.1 wstaje w 3,6 s.
            # Chromium dogaduje sie z Google po HTTP/2 (i probuje QUIC),
            # a proxy tego nie przepuszcza.
            "--disable-http2",
            "--disable-quic",
        ],
        user_agent=USER_AGENT,
        locale="pl-PL",
        viewport={"width": 1280, "height": 820},
        proxy=proxy_dla_playwright(),
    )
    strona = kontekst.pages[0] if kontekst.pages else kontekst.new_page()
    return kontekst, strona, sesja
