"""
Laczy zakupione dane z Google Maps (Google Maps Leads Scraper) z oficjalnymi
danymi firm z CEIDG i porownuje je ze soba (weryfikacja poprawnosci).

Klucz laczenia: kolumna "keyword" w pliku Google zawiera dokladnie ten sam
ciag "NazwaFirmy + Adres", ktory posluzyl do wyszukania - dzieki temu
dopasowanie jest pewne, a nie tylko przyblizone po adresie.

Gdy dla jednej firmy Google zwrocil kilka wynikow (ten sam keyword), wybierany
jest najlepszy kandydat wg punktacji:
  - kategoria Google (jubilerskie > zegarmistrzowskie > ogolne > pozostale),
  - zgodnosc adresu (kod pocztowy, numer budynku, nazwa ulicy),
  - podobienstwo nazwy do nazwy z CEIDG.

Wynik: Jubilerzy_CEIDG_x_GoogleMaps.xlsx
"""

import csv
import os
import re
import unicodedata
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

csv.field_size_limit(10_000_000)

SRC_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data"))
CEIDG_CSV = os.path.join(SRC_DIR, "Zarejestrowane dzialalnosci - 2026-07-10 - POLACZONE - aktywne.csv")
GMAPS_XLSX = os.path.join(SRC_DIR, "google_maps_leads.xlsx")
KONTAKT_XLSX = os.path.join(SRC_DIR, "kontakty_ze_stron.xlsx")
OUT_XLSX = os.path.join(SRC_DIR, "Jubilerzy_CEIDG_x_GoogleMaps.xlsx")

PKD_CODES = {"3212Z", "3213Z", "4777Z", "9525Z", "4648Z"}

# Kategorie Google wg priorytetu - dokladnie te, ktore interesuja uzytkownika,
# maja najwyzsza wage przy wyborze wlasciwej wizytowki.
KAT_JUBILERSKIE = {
    "jeweler", "jewelry store", "jewelry designer", "jewelry manufacturer",
    "goldsmith", "jewelry", "jewellery store", "jeweller",
}
KAT_POKREWNE = {
    "watch repair service", "watch store", "wholesale jeweler", "costume jewelry shop",
    "clock repair service", "jewelry appraiser", "jewelry buyer", "gold dealer",
    "diamond dealer", "engraver", "pawn shop", "watchmaker", "silversmith",
    "jewelry equipment supplier", "bead store", "gemologist",
}
KAT_OGOLNE = {
    "store", "manufacturer", "wholesaler", "craft store", "fashion accessories store",
    "internet shop", "corporate office", "souvenir store", "gift shop", "art gallery",
    "shopping mall", "boutique",
}
# Kategorie oznaczajace, ze Google zwrocil nie firme, tylko sam budynek/adres.
KAT_PUSTE = {"building", "apartment building", "housing complex", "null", "none", ""}


def norm(s):
    """Normalizuje tekst: bez polskich znakow, malymi literami, bez interpunkcji."""
    if s is None:
        return ""
    s = str(s).lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return " ".join(s.split())


def norm_phone(s):
    """Zostawia same cyfry i obcina prefiks kraju - do porownywania numerow."""
    if not s or str(s).lower() == "null":
        return ""
    d = re.sub(r"\D", "", str(s))
    if d.startswith("0048"):
        d = d[4:]
    elif d.startswith("48") and len(d) > 9:
        d = d[2:]
    return d[-9:] if len(d) >= 9 else d


def kod_pocztowy(s):
    m = re.search(r"\d{2}-\d{3}", str(s or ""))
    return m.group(0) if m else ""


def numer_budynku(s):
    """Wyciaga numer budynku (pierwsza liczba wystepujaca po nazwie ulicy)."""
    s = re.sub(r"\d{2}-\d{3}", " ", str(s or ""))  # usun kod pocztowy
    nums = re.findall(r"\b(\d+[a-zA-Z]?)\b", s)
    return nums[0].lower() if nums else ""


def clean(v):
    """Zamienia tekstowe 'null' ze scrapera na pusta wartosc."""
    if v is None:
        return ""
    v = str(v).strip()
    return "" if v.lower() in ("null", "none", "nan") else v


def zbuduj_adres(row, idx):
    ulica = row[idx["Ulica"]].strip()
    nr_bud = row[idx["NrBudynku"]].strip()
    nr_lok = row[idx["NrLokalu"]].strip()
    kod = row[idx["KodPocztowy"]].strip()
    miasto = row[idx["Miejscowosc"]].strip()
    if not (ulica and nr_bud and kod and miasto):
        return None
    street = f"{ulica} {nr_bud}"
    if nr_lok:
        street += f"/{nr_lok}"
    return f"{street}, {kod} {miasto}"


def wczytaj_ceidg():
    """Zwraca liste rekordow CEIDG (aktywne, wybrane PKD, pelny adres)."""
    firmy = []
    with open(CEIDG_CSV, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        header = next(reader)
        idx = {h: i for i, h in enumerate(header)}
        for row in reader:
            if row[idx["GlownyKodPkd"]].strip() not in PKD_CODES:
                continue
            nazwa = row[idx["NazwaPodmiotu"]].strip()
            adres = zbuduj_adres(row, idx)
            if not nazwa or nazwa == "-" or not adres:
                continue
            firmy.append({
                "klucz": norm(f"{nazwa} {adres}"),
                "nazwa": nazwa,
                "adres": adres,
                "nip": row[idx["Nip"]].strip(),
                "regon": row[idx["Regon"]].strip(),
                "imie": row[idx["Imie"]].strip(),
                "nazwisko": row[idx["Nazwisko"]].strip(),
                "telefon": row[idx["Telefon"]].strip(),
                "email": row[idx["Email"]].strip(),
                "www": row[idx["AdresWWW"]].strip(),
                "pkd": row[idx["GlownyKodPkd"]].strip(),
                "miasto": row[idx["Miejscowosc"]].strip(),
                "kod": row[idx["KodPocztowy"]].strip(),
                "powiat": row[idx["Powiat"]].strip(),
                "start": row[idx["DataRozpoczeciaDzialalnosci"]].strip(),
            })
    return firmy


def wczytaj_gmaps():
    """Grupuje wiersze z pliku Google wg znormalizowanego keyword."""
    wb = load_workbook(GMAPS_XLSX, read_only=True)
    ws = wb["Sheet1"]
    rows = ws.iter_rows(values_only=True)
    hdr = list(next(rows))
    i = {h: n for n, h in enumerate(hdr)}
    grupy = defaultdict(list)
    for r in rows:
        grupy[norm(r[i["keyword"]])].append({
            "title": clean(r[i["Title"]]),
            "address": clean(r[i["Address"]]),
            "city": clean(r[i["City"]]),
            "phone": clean(r[i["Phone"]]),
            "website": clean(r[i["Website"]]),
            "reviews": clean(r[i["Reviews"]]),
            "points": clean(r[i["Review_points"]]),
            "kat_list": clean(r[i["Category_listing"]]),
            "kat_det": clean(r[i["Category_detail"]]),
            "place_id": clean(r[i["Place_id"]]),
            "url": clean(r[i["Page_URL"]]),
            "lat": clean(r[i["Latitude"]]),
            "lon": clean(r[i["Longitude"]]),
            "status": clean(r[i["Current_Status"]]),
            "claim": clean(r[i["Claim_This_Business"]]),
        })
    wb.close()
    return grupy


def domena(url):
    """Sprowadza adres strony do samej domeny - klucz laczenia z kontaktami."""
    if not url:
        return ""
    u = str(url).strip().lower()
    u = re.sub(r"^https?://", "", u)
    u = re.sub(r"^www\.", "", u)
    return u.split("/")[0].split("?")[0].strip()


# Adresy-wypelniacze z szablonow stron ("mail@domain.com", "jan.kowalski@email.pl")
# oraz techniczne - nie sa prawdziwymi kontaktami do firmy.
SMIECIOWE_DOMENY_MAIL = {
    "domain.com", "email.com", "email.pl", "example.com", "example.pl",
    "yourdomain.com", "test.com", "sentry.io", "sentry.ovhcloud.com",
    "users.noreply.github.com", "wixpress.com", "twojadomena.pl",
}
SMIECIOWE_WZORCE = re.compile(
    r"(noreply|no-reply|^twoj@|^your@|^jan\.kowalski@|^johndoe@|^example@|"
    r"^znajomy@|^email@|\.(png|jpg|jpeg|gif|webp)$)"
)


def czy_mail_prawdziwy(mail):
    if "@" not in mail:
        return False
    if mail.split("@")[-1] in SMIECIOWE_DOMENY_MAIL:
        return False
    return not SMIECIOWE_WZORCE.search(mail)


def wczytaj_kontakty():
    """Agreguje dane kontaktowe ze stron firmowych - jeden wpis na domene.

    Plik ze scrapera ma wiele wierszy na te sama strone (kolejne podstrony),
    wiec maile i telefony trzeba zebrac i odduplikowac per domena.
    """
    wb = load_workbook(KONTAKT_XLSX, read_only=True)
    ws = wb["Sheet1"]
    rows = ws.iter_rows(values_only=True)
    hdr = list(next(rows))
    i = {h: n for n, h in enumerate(hdr)}
    dane = defaultdict(lambda: {"maile": [], "tel": [], "fb": "", "ig": ""})

    def dodaj(lista, wartosc):
        """Dodaje zachowujac kolejnosc wystepowania i bez duplikatow."""
        if wartosc and wartosc not in lista:
            lista.append(wartosc)

    for r in rows:
        d = domena(r[i["Start_URL"]])
        if not d:
            continue
        wpis = dane[d]
        if r[i["Emails"]]:
            for e in str(r[i["Emails"]]).split(";"):
                e = e.strip().lower()
                if e and czy_mail_prawdziwy(e):
                    dodaj(wpis["maile"], e)
        if r[i["Phones"]]:
            for p in str(r[i["Phones"]]).split(";"):
                dodaj(wpis["tel"], p.strip())
        if r[i["Facebook"]] and not wpis["fb"]:
            wpis["fb"] = str(r[i["Facebook"]]).split(";")[0].strip()
        if r[i["Instagram"]] and not wpis["ig"]:
            wpis["ig"] = str(r[i["Instagram"]]).split(";")[0].strip()
    wb.close()
    return dane


# Ogolnokrajowe sieci jubilerskie z wieloma salonami - w tych samych
# okolicach (ten sam kod pocztowy) potrafi wystapic kilka ich placowek,
# wiec samo trafienie adresu bez zgodnosci nazwy jest tu zawodne.
SIECI_WIELOODDZIALOWE = {norm(n) for n in (
    "apart", "w kruk", "wkruk", "yes", "marko", "savicki", "briju",
    "jubiler pl", "diament", "biju", "tous", "swarovski", "pandora",
)}


def punktuj(kandydat, firma):
    """Ocenia, na ile dany wynik z Google pasuje do firmy z CEIDG.

    Zgodnosc nazwy jest warunkiem koniecznym wysokiej pewnosci: bez niej
    wynik moze wskazywac inna firme pod tym samym adresem (typowe przy
    sieciach z wieloma salonami, np. Apart w kilku lokalizacjach o tym
    samym kodzie pocztowym co szukana mala firma).
    """
    score = 0
    kat = norm(kandydat["kat_list"])

    if kat in {norm(k) for k in KAT_JUBILERSKIE}:
        score += 100
    elif kat in {norm(k) for k in KAT_POKREWNE}:
        score += 60
    elif kat in {norm(k) for k in KAT_OGOLNE}:
        score += 20
    elif kat in {norm(k) for k in KAT_PUSTE}:
        score -= 20  # sam budynek, nie firma

    # zgodnosc adresu
    if kod_pocztowy(kandydat["address"]) and kod_pocztowy(kandydat["address"]) == firma["kod"]:
        score += 40
    if numer_budynku(kandydat["address"]) and numer_budynku(kandydat["address"]) == numer_budynku(firma["adres"]):
        score += 25

    # podobienstwo nazwy (wspolne slowa dluzsze niz 3 znaki)
    t_firma = {w for w in norm(firma["nazwa"]).split() if len(w) > 3}
    t_google = {w for w in norm(kandydat["title"]).split() if len(w) > 3}
    zgodnosc_nazwy = len(t_firma & t_google) / len(t_firma) if t_firma and t_google else 0
    score += 30 * zgodnosc_nazwy

    # wizytowka z realna trescia jest bardziej wiarygodna niz pusty pin
    if kandydat["phone"]:
        score += 10
    if kandydat["reviews"]:
        score += 5

    # Bez ZADNEJ zgodnosci nazwy dopasowanie moze byc przypadkowe (inna
    # firma pod tym samym adresem/kodem pocztowym) - ograniczamy pewnosc
    # do poziomu "do sprawdzenia", nigdy "pewne"/"prawdopodobne".
    if zgodnosc_nazwy == 0:
        score = min(score, 99)
        # Znana siec wieloplacowkowa bez potwierdzenia nazwa - jeszcze nizej,
        # bo prawdopodobienstwo pomylki z inna placowka tej samej sieci jest
        # wysokie.
        google_title_norm = norm(kandydat["title"])
        if any(siec in google_title_norm for siec in SIECI_WIELOODDZIALOWE):
            score = min(score, 59)

    return score


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FILL_OK = PatternFill("solid", fgColor="C6EFCE")
FILL_ROZNI = PatternFill("solid", fgColor="FFC7CE")
FILL_BRAK = PatternFill("solid", fgColor="F2F2F2")
FILL_UWAGA = PatternFill("solid", fgColor="FFEB9C")


def main():
    print("Wczytywanie CEIDG...")
    firmy = wczytaj_ceidg()
    print(f"  rekordow CEIDG (PKD jubilerskie, pelny adres): {len(firmy)}")

    print("Wczytywanie danych Google Maps...")
    grupy = wczytaj_gmaps()
    print(f"  unikalnych zapytan w pliku Google: {len(grupy)}")

    print("Wczytywanie kontaktow ze stron firmowych...")
    kontakty = wczytaj_kontakty()
    z_mailem = sum(1 for v in kontakty.values() if v["maile"])
    print(f"  domen w pliku kontaktow: {len(kontakty)} (z mailem: {z_mailem})")

    wb = Workbook()
    ws = wb.active
    ws.title = "CEIDG x Google Maps"
    kolumny = [
        "Nazwa firmy (CEIDG)", "Nazwa w Google Maps", "NIP", "REGON",
        "Imie", "Nazwisko", "Stanowisko",
        "Adres (CEIDG)", "Adres (Google)", "Miasto", "Kod pocztowy", "Powiat",
        "Telefon (CEIDG)", "Telefon (Google)", "Zgodnosc telefonu",
        "Email (CEIDG)", "Email ze strony WWW", "Telefon ze strony WWW",
        "Facebook", "Instagram", "WWW (CEIDG)", "WWW (Google)",
        "PKD", "Kategoria G-Maps", "Kategoria szczegolowa",
        "Liczba opinii", "Srednia ocena", "Status",
        "Zgodnosc kodu pocztowego", "Zgodnosc nr budynku",
        "Wiarygodnosc dopasowania", "Liczba wynikow Google", "Trafnosc dopasowania",
        "Szerokosc geo.", "Dlugosc geo.", "Place ID", "Link do wizytowki Google",
        "Data rozpoczecia dzialalnosci",
    ]
    ws.append(kolumny)

    st = {"z_google": 0, "bez_google": 0, "tel_ok": 0, "tel_rozne": 0,
          "kod_ok": 0, "kod_rozny": 0, "kat_jub": 0,
          "pewne": 0, "prawdopodobne": 0, "do_sprawdzenia": 0, "watpliwe": 0,
          "mail_www": 0, "mail_www_nowy": 0}

    for f in firmy:
        kandydaci = grupy.get(f["klucz"], [])
        if kandydaci:
            oceny = [(punktuj(k, f), k) for k in kandydaci]
            oceny.sort(key=lambda x: -x[0])
            best_score, g = oceny[0]
            st["z_google"] += 1
        else:
            best_score, g = None, None
            st["bez_google"] += 1

        tel_ceidg_n = norm_phone(f["telefon"])
        tel_google_n = norm_phone(g["phone"]) if g else ""
        if tel_ceidg_n and tel_google_n:
            if tel_ceidg_n == tel_google_n:
                zgodnosc_tel = "zgodny"
                st["tel_ok"] += 1
            else:
                zgodnosc_tel = "ROZNI SIE"
                st["tel_rozne"] += 1
        elif tel_google_n and not tel_ceidg_n:
            zgodnosc_tel = "tylko w Google"
        elif tel_ceidg_n and not tel_google_n:
            zgodnosc_tel = "tylko w CEIDG"
        else:
            zgodnosc_tel = "brak w obu"

        if g:
            kod_g = kod_pocztowy(g["address"])
            if kod_g and f["kod"]:
                if kod_g == f["kod"]:
                    zgodnosc_kod = "zgodny"
                    st["kod_ok"] += 1
                else:
                    zgodnosc_kod = "ROZNI SIE"
                    st["kod_rozny"] += 1
            else:
                zgodnosc_kod = "brak danych"
            nr_g = numer_budynku(g["address"])
            nr_c = numer_budynku(f["adres"])
            zgodnosc_nr = "zgodny" if (nr_g and nr_g == nr_c) else ("ROZNI SIE" if nr_g else "brak danych")
            if norm(g["kat_list"]) in {norm(k) for k in KAT_JUBILERSKIE}:
                st["kat_jub"] += 1
        else:
            zgodnosc_kod = zgodnosc_nr = "brak w Google"

        # Slowna ocena, na ile mozna ufac, ze to ta sama firma. Google czesto
        # zwraca sam pin budynku albo sasiednia firme, gdy wizytowki nie ma.
        if best_score is None:
            wiarygodnosc = "brak w Google"
        elif best_score >= 150:
            wiarygodnosc = "pewne"
            st["pewne"] += 1
        elif best_score >= 100:
            wiarygodnosc = "prawdopodobne"
            st["prawdopodobne"] += 1
        elif best_score >= 60:
            wiarygodnosc = "do sprawdzenia"
            st["do_sprawdzenia"] += 1
        else:
            wiarygodnosc = "watpliwe (inny obiekt)"
            st["watpliwe"] += 1

        # Kontakty ze strony firmowej - dopasowanie po domenie. Priorytet ma
        # strona z wizytowki Google (aktualniejsza), z fallbackiem na CEIDG.
        dom_www = domena(g["website"]) if g and g["website"] else ""
        if not dom_www:
            dom_www = domena(f["www"])
        kon = kontakty.get(dom_www) if dom_www else None
        mail_www = "; ".join(kon["maile"]) if kon else ""
        tel_www = "; ".join(kon["tel"][:3]) if kon else ""
        if mail_www:
            st["mail_www"] += 1
            if not f["email"]:
                st["mail_www_nowy"] += 1

        ws.append([
            f["nazwa"], g["title"] if g else "", f["nip"], f["regon"],
            f["imie"], f["nazwisko"], "Właściciel",
            f["adres"], g["address"] if g else "", f["miasto"], f["kod"], f["powiat"],
            f["telefon"], g["phone"] if g else "", zgodnosc_tel,
            f["email"], mail_www, tel_www,
            kon["fb"] if kon else "", kon["ig"] if kon else "",
            f["www"], g["website"] if g else "",
            f["pkd"], g["kat_list"] if g else "", g["kat_det"] if g else "",
            int(g["reviews"]) if g and g["reviews"].isdigit() else "",
            float(g["points"]) if g and re.match(r"^\d+(\.\d+)?$", g["points"] or "") else "",
            g["status"] if g else "",
            zgodnosc_kod, zgodnosc_nr,
            wiarygodnosc, len(kandydaci), round(best_score) if best_score is not None else "",
            g["lat"] if g else "", g["lon"] if g else "",
            g["place_id"] if g else "", g["url"] if g else "",
            f["start"],
        ])

    # formatowanie
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "C2"
    ws.row_dimensions[1].height = 30

    col_tel = kolumny.index("Zgodnosc telefonu") + 1
    col_kod = kolumny.index("Zgodnosc kodu pocztowego") + 1
    col_nr = kolumny.index("Zgodnosc nr budynku") + 1
    col_wiar = kolumny.index("Wiarygodnosc dopasowania") + 1
    for r in range(2, ws.max_row + 1):
        for c in (col_tel, col_kod, col_nr):
            val = ws.cell(r, c).value
            if val == "zgodny":
                ws.cell(r, c).fill = FILL_OK
            elif val == "ROZNI SIE":
                ws.cell(r, c).fill = FILL_ROZNI
            elif val in ("brak w Google", "brak w obu", "brak danych"):
                ws.cell(r, c).fill = FILL_BRAK
        val = ws.cell(r, col_wiar).value
        if val == "pewne":
            ws.cell(r, col_wiar).fill = FILL_OK
        elif val == "prawdopodobne":
            ws.cell(r, col_wiar).fill = FILL_UWAGA
        elif val in ("watpliwe (inny obiekt)", "brak w Google"):
            ws.cell(r, col_wiar).fill = FILL_ROZNI

    szerokosci = {
        "Nazwa firmy (CEIDG)": 45, "Nazwa w Google Maps": 32, "Adres (CEIDG)": 40,
        "Adres (Google)": 40, "Link do wizytowki Google": 45, "Email (CEIDG)": 28,
        "WWW (CEIDG)": 25, "WWW (Google)": 28, "Kategoria G-Maps": 22,
        "Kategoria szczegolowa": 22, "Wiarygodnosc dopasowania": 24,
        "Email ze strony WWW": 34, "Telefon ze strony WWW": 26,
        "Facebook": 30, "Instagram": 30,
    }
    for n, nazwa_kol in enumerate(kolumny, start=1):
        ws.column_dimensions[get_column_letter(n)].width = szerokosci.get(nazwa_kol, 16)

    ws.add_table(Table(
        displayName="TabelaWeryfikacja",
        ref=f"A1:{get_column_letter(len(kolumny))}{ws.max_row}",
        tableStyleInfo=TableStyleInfo(name="TableStyleLight9", showRowStripes=True),
    ))

    # arkusz podsumowania
    ws2 = wb.create_sheet("Podsumowanie")
    ws2.append(["Wskaznik", "Liczba"])
    podsumowanie = [
        ("Firmy CEIDG (PKD jubilerskie, pelny adres)", len(firmy)),
        ("Dopasowane do wizytowki Google", st["z_google"]),
        ("Bez wyniku w pliku Google", st["bez_google"]),
        ("", ""),
        ("Dopasowanie PEWNE", st["pewne"]),
        ("Dopasowanie PRAWDOPODOBNE", st["prawdopodobne"]),
        ("Dopasowanie DO SPRAWDZENIA", st["do_sprawdzenia"]),
        ("Dopasowanie WATPLIWE (Google zwrocil inny obiekt)", st["watpliwe"]),
        ("", ""),
        ("Trafiony profil w kategorii jubilerskiej", st["kat_jub"]),
        ("Telefon zgodny (CEIDG = Google)", st["tel_ok"]),
        ("Telefon rozni sie", st["tel_rozne"]),
        ("Kod pocztowy zgodny", st["kod_ok"]),
        ("Kod pocztowy rozni sie", st["kod_rozny"]),
        ("", ""),
        ("Email pozyskany ze strony WWW", st["mail_www"]),
        ("  w tym firmy bez emaila w CEIDG (nowy kontakt)", st["mail_www_nowy"]),
    ]
    for k, v in podsumowanie:
        ws2.append([k, v])
    for cell in ws2[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws2.column_dimensions["A"].width = 45
    ws2.column_dimensions["B"].width = 14

    # Jesli plik jest otwarty w Excelu, Windows blokuje zapis - zapisujemy
    # wtedy obok, zeby nie stracic wyniku obliczen.
    sciezka = OUT_XLSX
    try:
        wb.save(sciezka)
    except PermissionError:
        baza, roz = os.path.splitext(OUT_XLSX)
        sciezka = f"{baza}_NOWY{roz}"
        wb.save(sciezka)
        print("\nUWAGA: docelowy plik byl otwarty w Excelu - zapisano pod nowa nazwa.")
    print(f"\nZapisano: {sciezka}")
    for k, v in podsumowanie:
        print(f"  {k}: {v}")


if __name__ == "__main__":
    main()
