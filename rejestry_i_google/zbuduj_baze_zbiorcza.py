"""
Buduje JEDNA zbiorcza baze firm jubilerskich: JDG (z CEIDG) + spolki (z KRS),
wzbogacona o dane z Map Google i kontakty ze stron firmowych, wraz z
weryfikacja zgodnosci danych rejestrowych z tym, co pokazuje Google.

Zrodla:
  1. CEIDG      - jednoosobowe dzialalnosci (NIP, REGON, adres, telefon, email)
  2. Rejestr.io - spolki (KRS, NIP, REGON, adres, osoby, forma prawna, kapital)
  3. Google Maps Leads Scraper - 3 pliki (2x JDG + 1x spolki)
  4. Contact Details Scraper   - maile i telefony ze stron firmowych

Oba typy podmiotow trafiaja do jednego arkusza z kolumna "Typ podmiotu",
dzieki czemu caly zbior filtruje sie i sortuje tak samo.

Wynik: Jubilerzy_BAZA_ZBIORCZA.xlsx
"""

import os
import re
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from polacz_gmaps_z_ceidg import (
    FILL_BRAK, FILL_OK, FILL_ROZNI, FILL_UWAGA, HEADER_FILL, HEADER_FONT,
    KAT_JUBILERSKIE, clean, domena, kod_pocztowy, norm, norm_phone,
    numer_budynku, punktuj, wczytaj_ceidg, wczytaj_kontakty,
)

SRC_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data"))
DL = SRC_DIR  # pliki wejsciowe (eksporty narzedzi) w katalogu data/

GMAPS_PLIKI = [
    os.path.join(DL, "Google Maps Leads Scraper _by Keywords_.xlsx"),
    os.path.join(DL, "Google Maps Leads Scraper _by Keywords_(1).xlsx"),
    os.path.join(DL, "Google Maps Leads Scraper _by Keywords_(2).xlsx"),
]
SPOLKI_SZCZEGOLY = os.path.join(DL, "rejestrio_spolki_szczegoly.xlsx")
SPOLKI_LISTA = os.path.join(DL, "rejestrio_spolki_lista.xlsx")
OUT_XLSX = os.path.join(SRC_DIR, "Jubilerzy_BAZA_ZBIORCZA.xlsx")


def wczytaj_gmaps_wiele(sciezki):
    """Laczy wszystkie pliki z Map Google w jedna pule kandydatow.

    Klucz to znormalizowany "keyword" (nazwa + adres uzyty do wyszukania).
    Ten sam podmiot moze wystapic w kilku plikach - duplikaty (to samo
    Place_id) sa pomijane, zeby nie zawyzac liczby wynikow.
    """
    grupy = defaultdict(list)
    for sciezka in sciezki:
        if not os.path.exists(sciezka):
            print(f"  POMINIETO (brak pliku): {os.path.basename(sciezka)}")
            continue
        wb = load_workbook(sciezka, read_only=True)
        ws = wb["Sheet1"]
        rows = ws.iter_rows(values_only=True)
        hdr = list(next(rows))
        i = {h: n for n, h in enumerate(hdr)}
        n = 0
        for r in rows:
            klucz = norm(r[i["keyword"]])
            wpis = {
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
            }
            istniejace = {w["place_id"] for w in grupy[klucz] if w["place_id"]}
            if wpis["place_id"] and wpis["place_id"] in istniejace:
                continue
            grupy[klucz].append(wpis)
            n += 1
        wb.close()
        print(f"  {os.path.basename(sciezka)}: +{n} wynikow")
    return grupy


def wczytaj_spolki():
    """Scala dwa pliki z Rejestr.io w jedna liste spolek (klucz: KRS).

    Plik szczegolowy ma adres, NIP i REGON; plik zbiorczy dokłada forme
    prawna, rok rejestracji, kapital, opis PKD i pozostale osoby.
    """
    spolki = {}

    wb = load_workbook(SPOLKI_SZCZEGOLY, read_only=True)
    ws = wb["Sheet1"]
    rows = ws.iter_rows(values_only=True)
    hdr = list(next(rows))
    j = {h: n for n, h in enumerate(hdr)}
    for r in rows:
        krs = re.sub(r"\D", "", str(r[j["KRS"]] or ""))
        if not krs:
            continue
        adres_raw = clean(r[j["Text1"]])
        # "Polska" bez ulicy oznacza brak adresu w rejestrze
        if adres_raw.strip().lower() in ("polska", "poland"):
            adres_raw = ""
        # Wyszukiwania w Google robiono na adresie w oryginalnej postaci
        # (z koncowka ", Polska"), wiec to ona musi trafic do klucza laczenia,
        # a wersja bez koncowki sluzy tylko do wyswietlania.
        adres = re.sub(r",\s*(Polska|Poland)\s*$", "", adres_raw).strip()
        osoba = clean(r[j["Text2"]])
        if krs in spolki:
            if osoba and osoba not in spolki[krs]["osoby"]:
                spolki[krs]["osoby"].append(osoba)
            continue
        spolki[krs] = {
            "krs": krs,
            "nazwa": clean(r[j["Text"]]),
            "nip": clean(r[j["NIP"]]),
            "regon": clean(r[j["REGON"]]),
            "adres": adres,
            "adres_klucz": adres_raw,
            "osoby": [osoba] if osoba else [],
            "forma": "", "rok": "", "miasto": "", "kapital": "",
            "pkd_opis": "", "link_rejestr": "",
        }
    wb.close()

    if os.path.exists(SPOLKI_LISTA):
        wb = load_workbook(SPOLKI_LISTA, read_only=True)
        ws = wb["Sheet1"]
        rows = ws.iter_rows(values_only=True)
        hdr = list(next(rows))
        j = {h: n for n, h in enumerate(hdr)}
        for r in rows:
            krs = re.sub(r"\D", "", str(r[j["Number"]] or ""))
            s = spolki.get(krs)
            if not s:
                continue
            s["forma"] = s["forma"] or clean(r[j["Info"]])
            s["rok"] = s["rok"] or clean(r[j["Info1"]])
            s["miasto"] = s["miasto"] or clean(r[j["Info2"]])
            s["kapital"] = s["kapital"] or clean(r[j["Info3"]])
            s["pkd_opis"] = s["pkd_opis"] or clean(r[j["val"]])
            s["link_rejestr"] = s["link_rejestr"] or clean(r[j["Title_URL"]])
            osoba = clean(r[j["Info4"]])
            if osoba and osoba not in s["osoby"]:
                s["osoby"].append(osoba)
        wb.close()

    # uzupelnij miasto i kod pocztowy z adresu, jesli sa tylko tam
    for s in spolki.values():
        s["kod"] = kod_pocztowy(s["adres"])
        if not s["miasto"] and s["adres"]:
            m = re.search(r"\d{2}-\d{3}\s+(.+)$", s["adres"])
            s["miasto"] = m.group(1).strip() if m else ""
    return spolki


def dopasuj(grupy, klucz_glowny, klucz_zapasowy=""):
    """Zwraca kandydatow z Google - najpierw po nazwie z adresem,
    a gdy podmiot nie ma adresu w rejestrze, po samej nazwie."""
    if klucz_glowny in grupy:
        return grupy[klucz_glowny]
    if klucz_zapasowy and klucz_zapasowy in grupy:
        return grupy[klucz_zapasowy]
    return []


KOLUMNY = [
    "Typ podmiotu", "Nazwa firmy (rejestr)", "Nazwa w Google Maps",
    "NIP", "KRS", "REGON", "Forma prawna",
    "Osoba decyzyjna", "Stanowisko",
    "Adres (rejestr)", "Adres (Google)", "Miasto", "Kod pocztowy", "Powiat",
    "Telefon (rejestr)", "Telefon (Google)", "Telefon ze strony WWW", "Zgodnosc telefonu",
    "Email (rejestr)", "Email ze strony WWW",
    "WWW (rejestr)", "WWW (Google)", "Facebook", "Instagram",
    "PKD / Branza", "Kategoria G-Maps", "Kategoria szczegolowa",
    "Liczba opinii", "Srednia ocena", "Status",
    "Wiarygodnosc dopasowania", "Zgodnosc kodu pocztowego", "Zgodnosc nr budynku",
    "Liczba wynikow Google", "Trafnosc dopasowania",
    "Rok / data rozpoczecia", "Kapital zakladowy",
    "Szerokosc geo.", "Dlugosc geo.", "Place ID",
    "Link do wizytowki Google", "Link do rejestru",
]


def zbuduj_wiersz(p, grupy, kontakty, st):
    """Sklada jeden wiersz wynikowy dla podmiotu (JDG albo spolki)."""
    kandydaci = dopasuj(grupy, p["klucz"], p.get("klucz_zapas", ""))
    if kandydaci:
        oceny = sorted(((punktuj(k, p), k) for k in kandydaci), key=lambda x: -x[0])
        best_score, g = oceny[0]
        st[p["typ"]]["z_google"] += 1
    else:
        best_score, g = None, None
        st[p["typ"]]["bez_google"] += 1

    tel_rej = norm_phone(p["telefon"])
    tel_goo = norm_phone(g["phone"]) if g else ""
    if tel_rej and tel_goo:
        zgodnosc_tel = "zgodny" if tel_rej == tel_goo else "ROZNI SIE"
        st[p["typ"]]["tel_ok" if tel_rej == tel_goo else "tel_rozne"] += 1
    elif tel_goo:
        zgodnosc_tel = "tylko w Google"
    elif tel_rej:
        zgodnosc_tel = "tylko w rejestrze"
    else:
        zgodnosc_tel = "brak w obu"

    if g:
        kod_g = kod_pocztowy(g["address"])
        if kod_g and p["kod"]:
            zgodny = kod_g == p["kod"]
            zgodnosc_kod = "zgodny" if zgodny else "ROZNI SIE"
            st[p["typ"]]["kod_ok" if zgodny else "kod_rozny"] += 1
        else:
            zgodnosc_kod = "brak danych"
        nr_g, nr_r = numer_budynku(g["address"]), numer_budynku(p["adres"])
        zgodnosc_nr = "zgodny" if (nr_g and nr_g == nr_r) else ("ROZNI SIE" if nr_g and nr_r else "brak danych")
        if norm(g["kat_list"]) in {norm(k) for k in KAT_JUBILERSKIE}:
            st[p["typ"]]["kat_jub"] += 1
    else:
        zgodnosc_kod = zgodnosc_nr = "brak w Google"

    if best_score is None:
        wiarygodnosc = "brak w Google"
    elif best_score >= 150:
        wiarygodnosc = "pewne"
        st[p["typ"]]["pewne"] += 1
    elif best_score >= 100:
        wiarygodnosc = "prawdopodobne"
        st[p["typ"]]["prawdopodobne"] += 1
    elif best_score >= 60:
        wiarygodnosc = "do sprawdzenia"
        st[p["typ"]]["do_sprawdzenia"] += 1
    else:
        wiarygodnosc = "watpliwe (inny obiekt)"
        st[p["typ"]]["watpliwe"] += 1

    dom = domena(g["website"]) if g and g["website"] else domena(p["www"])
    kon = kontakty.get(dom) if dom else None
    mail_www = "; ".join(kon["maile"]) if kon else ""
    if mail_www:
        st[p["typ"]]["mail_www"] += 1
        if not p["email"]:
            st[p["typ"]]["mail_www_nowy"] += 1

    return [
        p["typ"], p["nazwa"], g["title"] if g else "",
        p["nip"], p["krs"], p["regon"], p["forma"],
        p["osoba"], p["stanowisko"],
        p["adres"], g["address"] if g else "", p["miasto"], p["kod"], p["powiat"],
        p["telefon"], g["phone"] if g else "", "; ".join(kon["tel"][:3]) if kon else "", zgodnosc_tel,
        p["email"], mail_www,
        p["www"], g["website"] if g else "",
        kon["fb"] if kon else "", kon["ig"] if kon else "",
        p["pkd"], g["kat_list"] if g else "", g["kat_det"] if g else "",
        int(g["reviews"]) if g and g["reviews"].isdigit() else "",
        float(g["points"]) if g and re.match(r"^\d+(\.\d+)?$", g["points"] or "") else "",
        g["status"] if g else "",
        wiarygodnosc, zgodnosc_kod, zgodnosc_nr,
        len(kandydaci), round(best_score) if best_score is not None else "",
        p["rok"], p["kapital"],
        g["lat"] if g else "", g["lon"] if g else "", g["place_id"] if g else "",
        g["url"] if g else "", p["link_rejestr"],
    ]


def nowe_staty():
    return defaultdict(int)


def main():
    print("Wczytywanie CEIDG (JDG)...")
    jdg = wczytaj_ceidg()
    print(f"  JDG: {len(jdg)}")

    print("Wczytywanie spolek (Rejestr.io)...")
    spolki = wczytaj_spolki()
    z_adresem = sum(1 for s in spolki.values() if s["adres"])
    print(f"  Spolki: {len(spolki)} (z pelnym adresem: {z_adresem})")

    print("Wczytywanie Map Google...")
    grupy = wczytaj_gmaps_wiele(GMAPS_PLIKI)
    print(f"  laczna liczba unikalnych zapytan: {len(grupy)}")

    print("Wczytywanie kontaktow ze stron...")
    kontakty = wczytaj_kontakty()
    print(f"  domen: {len(kontakty)}")

    # ujednolicenie obu typow podmiotow do wspolnej struktury
    podmioty = []
    for f in jdg:
        podmioty.append({
            "typ": "JDG", "klucz": f["klucz"], "nazwa": f["nazwa"],
            "nip": f["nip"], "krs": "", "regon": f["regon"], "forma": "Jednoosobowa dzialalnosc",
            "osoba": f"{f['imie']} {f['nazwisko']}".strip(), "stanowisko": "Właściciel",
            "adres": f["adres"], "miasto": f["miasto"], "kod": f["kod"], "powiat": f["powiat"],
            "telefon": f["telefon"], "email": f["email"], "www": f["www"],
            "pkd": f["pkd"], "rok": f["start"], "kapital": "", "link_rejestr": "",
        })
    for s in spolki.values():
        podmioty.append({
            "typ": "Spolka", "klucz": norm(f"{s['nazwa']} {s['adres_klucz']}"),
            "klucz_zapas": norm(s["nazwa"]),
            "nazwa": s["nazwa"], "nip": s["nip"], "krs": s["krs"], "regon": s["regon"],
            "forma": s["forma"] or "Spolka", "osoba": "; ".join(s["osoby"]),
            "stanowisko": "Reprezentant / zarzad",
            "adres": s["adres"], "miasto": s["miasto"], "kod": s["kod"], "powiat": "",
            "telefon": "", "email": "", "www": "",
            "pkd": s["pkd_opis"], "rok": s["rok"], "kapital": s["kapital"],
            "link_rejestr": s["link_rejestr"],
        })

    st = {"JDG": nowe_staty(), "Spolka": nowe_staty()}

    wb = Workbook()
    ws = wb.active
    ws.title = "Baza zbiorcza"
    ws.append(KOLUMNY)
    for p in podmioty:
        ws.append(zbuduj_wiersz(p, grupy, kontakty, st))

    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "C2"
    ws.row_dimensions[1].height = 32

    kol = {n: i + 1 for i, n in enumerate(KOLUMNY)}
    for r in range(2, ws.max_row + 1):
        for nazwa_kol in ("Zgodnosc telefonu", "Zgodnosc kodu pocztowego", "Zgodnosc nr budynku"):
            c = ws.cell(r, kol[nazwa_kol])
            if c.value == "zgodny":
                c.fill = FILL_OK
            elif c.value == "ROZNI SIE":
                c.fill = FILL_ROZNI
            elif c.value in ("brak w Google", "brak w obu", "brak danych"):
                c.fill = FILL_BRAK
        c = ws.cell(r, kol["Wiarygodnosc dopasowania"])
        if c.value == "pewne":
            c.fill = FILL_OK
        elif c.value == "prawdopodobne":
            c.fill = FILL_UWAGA
        elif c.value in ("watpliwe (inny obiekt)", "brak w Google"):
            c.fill = FILL_ROZNI

    szer = {
        "Nazwa firmy (rejestr)": 45, "Nazwa w Google Maps": 30, "Osoba decyzyjna": 34,
        "Adres (rejestr)": 40, "Adres (Google)": 40, "Email (rejestr)": 26,
        "Email ze strony WWW": 32, "WWW (rejestr)": 24, "WWW (Google)": 28,
        "Facebook": 28, "Instagram": 28, "PKD / Branza": 34,
        "Kategoria G-Maps": 22, "Kategoria szczegolowa": 22,
        "Wiarygodnosc dopasowania": 24, "Link do wizytowki Google": 42,
        "Link do rejestru": 30, "Forma prawna": 22, "Stanowisko": 20,
    }
    for n, nazwa_kol in enumerate(KOLUMNY, start=1):
        ws.column_dimensions[get_column_letter(n)].width = szer.get(nazwa_kol, 15)

    ws.add_table(Table(
        displayName="BazaZbiorcza",
        ref=f"A1:{get_column_letter(len(KOLUMNY))}{ws.max_row}",
        tableStyleInfo=TableStyleInfo(name="TableStyleLight9", showRowStripes=True),
    ))

    # --- podsumowanie ---
    ws2 = wb.create_sheet("Podsumowanie")
    ws2.append(["Wskaznik", "JDG", "Spolki", "RAZEM"])
    liczby = {t: sum(1 for p in podmioty if p["typ"] == t) for t in ("JDG", "Spolka")}
    wiersze = [
        ("Liczba podmiotow", liczby["JDG"], liczby["Spolka"]),
        ("Dopasowane do wizytowki Google", st["JDG"]["z_google"], st["Spolka"]["z_google"]),
        ("Bez wyniku w Google", st["JDG"]["bez_google"], st["Spolka"]["bez_google"]),
        ("", "", ""),
        ("Dopasowanie PEWNE", st["JDG"]["pewne"], st["Spolka"]["pewne"]),
        ("Dopasowanie PRAWDOPODOBNE", st["JDG"]["prawdopodobne"], st["Spolka"]["prawdopodobne"]),
        ("Dopasowanie DO SPRAWDZENIA", st["JDG"]["do_sprawdzenia"], st["Spolka"]["do_sprawdzenia"]),
        ("Dopasowanie WATPLIWE", st["JDG"]["watpliwe"], st["Spolka"]["watpliwe"]),
        ("", "", ""),
        ("Profil w kategorii jubilerskiej", st["JDG"]["kat_jub"], st["Spolka"]["kat_jub"]),
        ("Telefon zgodny (rejestr = Google)", st["JDG"]["tel_ok"], st["Spolka"]["tel_ok"]),
        ("Telefon rozni sie", st["JDG"]["tel_rozne"], st["Spolka"]["tel_rozne"]),
        ("Kod pocztowy zgodny", st["JDG"]["kod_ok"], st["Spolka"]["kod_ok"]),
        ("Kod pocztowy rozni sie", st["JDG"]["kod_rozny"], st["Spolka"]["kod_rozny"]),
        ("", "", ""),
        ("Email ze strony WWW", st["JDG"]["mail_www"], st["Spolka"]["mail_www"]),
        ("  w tym nowy (brak w rejestrze)", st["JDG"]["mail_www_nowy"], st["Spolka"]["mail_www_nowy"]),
    ]
    for opis, a, b in wiersze:
        ws2.append([opis, a, b, (a + b) if isinstance(a, int) else ""])
    for cell in ws2[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws2.column_dimensions["A"].width = 42
    for c in "BCD":
        ws2.column_dimensions[c].width = 12

    sciezka = OUT_XLSX
    try:
        wb.save(sciezka)
    except PermissionError:
        baza, roz = os.path.splitext(OUT_XLSX)
        sciezka = f"{baza}_NOWY{roz}"
        wb.save(sciezka)
        print("\nUWAGA: plik byl otwarty w Excelu - zapisano pod nowa nazwa.")

    print(f"\nZapisano: {sciezka}\n")
    for opis, a, b in wiersze:
        if opis:
            razem = f"{a + b:>6}" if isinstance(a, int) else ""
            print(f"  {opis:<38} JDG:{a:>6}  Spolki:{b:>5}  Razem:{razem}")


if __name__ == "__main__":
    main()
