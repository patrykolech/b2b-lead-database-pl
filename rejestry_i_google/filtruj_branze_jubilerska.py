"""
Wybiera z bazy zbiorczej wylacznie podmioty faktycznie zwiazane z branza
jubilerska i przypisuje im segment (jubiler/zlotnik, sklep, producent,
sprzedaz online, zegarmistrz...).

Kwalifikacja opiera sie na trzech niezaleznych sygnalach:
  1. kategoria z Map Google  - najmocniejszy dowod, bo opisuje realna dzialalnosc,
  2. kod / opis PKD          - deklaracja z rejestru,
  3. nazwa firmy             - slowa typu "jubiler", "zlotnik", "bizuteria", "gold".

Podmiot trafia do wyniku, gdy kategoria Google jest jubilerska, albo gdy
kategoria jest ogolnohandlowa i potwierdza ja nazwa lub PKD. Firmy bez
wizytowki w Google przechodza, jesli branze potwierdza nazwa lub PKD.

Wynik: Jubilerzy_BRANZA.xlsx (arkusze: wybrane / odrzucone / podsumowanie)
"""

import os
import re
from collections import Counter

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from polacz_gmaps_z_ceidg import FILL_OK, FILL_UWAGA, HEADER_FILL, HEADER_FONT, norm

SRC_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data"))
IN_XLSX = os.path.join(SRC_DIR, "Jubilerzy_BAZA_ZBIORCZA.xlsx")
OUT_XLSX = os.path.join(SRC_DIR, "Jubilerzy_BRANZA.xlsx")

# --- kategorie Google -> segment branzy -------------------------------------
# Rdzen branzy: sama kategoria wystarczy, zeby zakwalifikowac podmiot.
KAT_RDZEN = {
    "jeweler": "Jubiler / zlotnik",
    "jewelry store": "Sklep jubilerski",
    "jewelry designer": "Projektant bizuterii",
    "jewelry manufacturer": "Producent bizuterii",
    "goldsmith": "Jubiler / zlotnik",
    "silversmith": "Jubiler / zlotnik",
    "costume jewelry shop": "Sklep - bizuteria sztuczna",
    "wholesale jeweler": "Hurt bizuterii",
    "jewelry repair service": "Naprawa bizuterii",
    "jewelry appraiser": "Rzeczoznawca / wycena",
    "jewelry buyer": "Skup zlota i bizuterii",
    "gold dealer": "Skup zlota i bizuterii",
    "diamond dealer": "Handel kamieniami",
    "jewelry equipment supplier": "Zaopatrzenie jubilerskie",
    "bead store": "Sklep - komponenty",
    "rock shop": "Kamienie i mineraly",
}
# Zegarmistrzostwo - w Polsce czesto laczone z jubilerstwem, ale to osobny
# segment, zeby dalo sie go w razie potrzeby odfiltrowac.
KAT_ZEGARY = {
    "watch repair service": "Zegarmistrz",
    "watch store": "Sklep z zegarkami",
    "clock repair service": "Zegarmistrz",
    "watchmaker": "Zegarmistrz",
}
# Kategorie ogolne - kwalifikuja TYLKO gdy nazwa lub PKD potwierdza bizuterie.
KAT_WARUNKOWE = {
    "fashion accessories store": "Sklep - akcesoria",
    "store": "Sklep",
    "general store": "Sklep",
    "boutique": "Sklep",
    "gift shop": "Sklep - upominki",
    "souvenir store": "Sklep - upominki",
    "internet shop": "Sprzedaz online",
    "e-commerce service": "Sprzedaz online",
    "craft store": "Rekodzielo",
    "handicraft": "Rekodzielo",
    "art studio": "Pracownia artystyczna",
    "artist": "Pracownia artystyczna",
    "art gallery": "Galeria / sztuka",
    "manufacturer": "Producent",
    "wholesaler": "Hurt",
    "engraver": "Grawer",
    "pawn shop": "Lombard / skup",
    "antique store": "Antyki",
    "religious goods store": "Dewocjonalia",
}

# --- sygnaly z nazwy firmy ---------------------------------------------------
# Wyrazenia jednoznacznie jubilerskie (dopasowanie na znormalizowanej nazwie).
NAZWA_MOCNE = re.compile(
    r"jubiler|zlotni|zlotnic|bizuteri|jewel|goldsmith|"
    r"\bgold\b|\bsilver\b|srebr|\bzloto\b|\bzlota\b|\bzlote\b|"
    r"diament|brylant|\bkarat\b|klejnot|bursztyn|\bamber\b|"
    r"obraczk|pierscion|\bperl|kamien szlachetn|\bargent|\baurum\b"
)
# Nazwy kamieni/metali - slabszy sygnal (bywaja nazwami firm spoza branzy).
NAZWA_SLABE = re.compile(
    r"\bszmaragd\b|\brubin\b|\btopaz\b|\bametyst\b|\bopal\b|\bturkus\b|"
    r"\bjantar\b|\bkoral\b|\bagat\b|\bnefryt\b|\bszafir\b|\bonyks\b|\bonyx\b"
)
NAZWA_ZEGARY = re.compile(r"zegar|zegarmistrz|\bwatch|\btime\b|chrono")

# --- sygnaly z PKD -----------------------------------------------------------
PKD_JUBILERSKIE = {"3212Z", "3213Z", "4648Z"}   # produkcja i hurt bizuterii
PKD_HANDEL_NAPRAWA = {"4777Z", "9525Z"}          # detal i naprawa (tez zegarki)
PKD_OPIS = re.compile(r"jubilers|bizuteri|zlotnic|zegar|metali szlachetnych|kamieni szlachetnych")

# Kategorie, ktore same w sobie oznaczaja, ze Google wskazal nie te firme.
KAT_ODRZUC = {
    "building", "corporate office", "shopping mall", "virtual office rental",
    "postal code", "package locker", "multi-unit residential building",
    "apartment building", "housing complex",
}


def sygnal_nazwa(nazwa):
    n = norm(nazwa)
    if NAZWA_MOCNE.search(n):
        return "mocny"
    if NAZWA_SLABE.search(n):
        return "slaby"
    return ""


def sygnal_pkd(pkd):
    """PKD moze byc kodem (JDG) albo opisem slownym (spolki)."""
    p = str(pkd or "").strip()
    if p in PKD_JUBILERSKIE:
        return "mocny"
    if p in PKD_HANDEL_NAPRAWA:
        return "sredni"
    if PKD_OPIS.search(norm(p)):
        return "mocny"
    return ""


# Segment wynikajacy z samego PKD - uzywany, gdy Google nie ma wizytowki firmy.
SEGMENT_WG_PKD = {
    "3212Z": "Pracownia / producent bizuterii",
    "3213Z": "Producent bizuterii sztucznej",
    "4648Z": "Hurt bizuterii i zegarkow",
    "4777Z": "Sklep jubilerski (brak wizytowki)",
    "9525Z": "Naprawa bizuterii i zegarkow",
}


def segment_z_pkd(pkd):
    p = str(pkd or "").strip()
    if p in SEGMENT_WG_PKD:
        return SEGMENT_WG_PKD[p]
    n = norm(p)
    if "produkcja" in n and ("jubilers" in n or "bizuteri" in n):
        return "Pracownia / producent bizuterii"
    if "sprzedaz" in n and "detaliczn" in n:
        return "Sklep jubilerski (brak wizytowki)"
    if "sprzedaz" in n and "hurtow" in n:
        return "Hurt bizuterii i zegarkow"
    if "naprawa" in n:
        return "Naprawa bizuterii i zegarkow"
    return "Branza jubilerska (wg PKD)"


def zaklasyfikuj(nazwa, kat_google, kat_szcz, pkd, ma_google):
    """Zwraca (segment, podstawa, sila) albo (None, powod, 0) gdy odrzucony.

    Sila: 3 = potwierdzone kategoria Google, 2 = zgodne PKD lub nazwa,
    1 = slaba poszlaka (do recznego sprawdzenia).
    """
    kat = norm(kat_google)
    kat_d = norm(kat_szcz)
    s_naz = sygnal_nazwa(nazwa)
    s_pkd = sygnal_pkd(pkd)

    # Kategoria typu "Building" czy "Postal code" nie opisuje zadnej firmy -
    # dla oceny branzy to tyle samo, co brak wizytowki.
    bez_wizytowki = (not ma_google) or kat in KAT_ODRZUC or kat_d in KAT_ODRZUC

    # 1. Kategoria Google z rdzenia branzy - najpewniejszy przypadek.
    for zrodlo in (kat, kat_d):
        if zrodlo in KAT_RDZEN:
            return KAT_RDZEN[zrodlo], "kategoria Google (jubilerska)", 3

    # 2. Zegarmistrzostwo - branza pokrewna.
    for zrodlo in (kat, kat_d):
        if zrodlo in KAT_ZEGARY:
            sila = 3 if (s_naz or s_pkd) else 2
            return KAT_ZEGARY[zrodlo], "kategoria Google (zegarmistrzowska)", sila

    # 3. Kategoria ogolna - potrzebne potwierdzenie z nazwy albo PKD.
    for zrodlo in (kat, kat_d):
        if zrodlo in KAT_WARUNKOWE:
            if s_naz == "mocny" or s_pkd == "mocny":
                return KAT_WARUNKOWE[zrodlo], "kategoria ogolna + nazwa/PKD", 2
            if s_naz or s_pkd:
                return KAT_WARUNKOWE[zrodlo], "kategoria ogolna + slaby sygnal", 1
            return None, "kategoria ogolna bez potwierdzenia", 0

    # 4. Brak wizytowki firmy w Google - o branzy decyduje rejestr i nazwa.
    #    PKD produkcji/hurtu bizuterii to oficjalna deklaracja glownej
    #    dzialalnosci, wiec kwalifikuje samodzielnie.
    if bez_wizytowki:
        if s_pkd == "mocny":
            podstawa = "PKD + nazwa firmy" if s_naz else "PKD (glowna dzialalnosc)"
            return segment_z_pkd(pkd), podstawa, 2
        if s_naz == "mocny":
            return "Bez wizytowki (nazwa wskazuje branze)", "nazwa firmy", 2
        if s_pkd == "sredni":
            podstawa = "PKD + nazwa firmy" if s_naz else "PKD (handel / naprawa)"
            return segment_z_pkd(pkd), podstawa, 2 if s_naz else 1
        if s_naz == "slaby":
            return "Bez wizytowki (slaby sygnal)", "nazwa poboczna", 1
        return None, "brak wizytowki i brak sygnalow", 0

    # 5. Google pokazuje inna, dzialajaca firme pod tym adresem.
    if s_naz == "mocny" or s_pkd == "mocny":
        podstawa = "nazwa firmy" if s_naz == "mocny" else "PKD (glowna dzialalnosc)"
        return "Do weryfikacji (Google wskazal inny obiekt)", podstawa, 1
    return None, "kategoria spoza branzy", 0


def main():
    print("Wczytywanie bazy zbiorczej...")
    wb_in = load_workbook(IN_XLSX, read_only=True)
    ws_in = wb_in["Baza zbiorcza"]
    rows = ws_in.iter_rows(values_only=True)
    hdr = list(next(rows))
    i = {h: n for n, h in enumerate(hdr)}

    nowe_kol = ["Segment branzy", "Podstawa kwalifikacji", "Sila powiazania"]
    kol_wynik = nowe_kol + hdr

    wybrane, odrzucone = [], []
    seg = Counter()
    powody = Counter()
    typy = Counter()

    for r in rows:
        ma_google = bool(r[i["Nazwa w Google Maps"]])
        segment, podstawa, sila = zaklasyfikuj(
            r[i["Nazwa firmy (rejestr)"]], r[i["Kategoria G-Maps"]],
            r[i["Kategoria szczegolowa"]], r[i["PKD / Branza"]], ma_google,
        )
        if segment:
            etykieta = {3: "wysoka", 2: "srednia", 1: "niska"}[sila]
            wybrane.append([segment, podstawa, etykieta] + list(r))
            seg[segment] += 1
            typy[r[i["Typ podmiotu"]]] += 1
        else:
            odrzucone.append([""] * 3 + list(r))
            powody[podstawa] += 1
    wb_in.close()

    # najpierw najpewniejsze, w obrebie tej samej pewnosci - najlepiej oceniane
    kolejnosc = {"wysoka": 0, "srednia": 1, "niska": 2}
    wybrane.sort(key=lambda w: (
        kolejnosc[w[2]],
        -(w[3 + i["Liczba opinii"]] if isinstance(w[3 + i["Liczba opinii"]], int) else 0),
    ))

    wb = Workbook()
    ws = wb.active
    ws.title = "Branza jubilerska"
    ws.append(kol_wynik)
    for w in wybrane:
        ws.append(w)

    ws2 = wb.create_sheet("Odrzucone")
    ws2.append(kol_wynik)
    for w in odrzucone:
        ws2.append(w)

    for arkusz in (ws, ws2):
        for cell in arkusz[1]:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        arkusz.freeze_panes = "D2"
        arkusz.row_dimensions[1].height = 32
        szer = {
            "Segment branzy": 32, "Podstawa kwalifikacji": 30, "Sila powiazania": 14,
            "Nazwa firmy (rejestr)": 45, "Nazwa w Google Maps": 30, "Osoba decyzyjna": 32,
            "Adres (rejestr)": 38, "Adres (Google)": 38, "Email (rejestr)": 26,
            "Email ze strony WWW": 32, "WWW (Google)": 26, "PKD / Branza": 32,
            "Kategoria G-Maps": 22, "Kategoria szczegolowa": 22,
            "Wiarygodnosc dopasowania": 22, "Link do wizytowki Google": 40,
            "Link do rejestru": 28, "Forma prawna": 20,
        }
        for n, nazwa_kol in enumerate(kol_wynik, start=1):
            arkusz.column_dimensions[get_column_letter(n)].width = szer.get(nazwa_kol, 15)

    for r in range(2, ws.max_row + 1):
        c = ws.cell(r, 3)
        if c.value == "wysoka":
            c.fill = FILL_OK
        elif c.value == "srednia":
            c.fill = FILL_UWAGA

    if wybrane:
        ws.add_table(Table(
            displayName="BranzaJubilerska",
            ref=f"A1:{get_column_letter(len(kol_wynik))}{ws.max_row}",
            tableStyleInfo=TableStyleInfo(name="TableStyleLight9", showRowStripes=True),
        ))

    ws3 = wb.create_sheet("Podsumowanie")
    ws3.append(["Segment branzy", "Liczba firm"])
    for k, v in seg.most_common():
        ws3.append([k, v])
    ws3.append(["", ""])
    ws3.append(["RAZEM WYBRANE", len(wybrane)])
    ws3.append(["  w tym JDG", typy["JDG"]])
    ws3.append(["  w tym spolki", typy["Spolka"]])
    ws3.append(["", ""])
    ws3.append(["ODRZUCONE - powod", "Liczba"])
    for k, v in powody.most_common():
        ws3.append([k, v])
    ws3.append(["RAZEM ODRZUCONE", len(odrzucone)])
    for cell in ws3[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    ws3.column_dimensions["A"].width = 46
    ws3.column_dimensions["B"].width = 14

    sciezka = OUT_XLSX
    try:
        wb.save(sciezka)
    except PermissionError:
        baza, roz = os.path.splitext(OUT_XLSX)
        sciezka = f"{baza}_NOWY{roz}"
        wb.save(sciezka)
        print("\nUWAGA: plik byl otwarty w Excelu - zapisano pod nowa nazwa.")

    print(f"\nZapisano: {sciezka}")
    print(f"\nWYBRANE: {len(wybrane)} (JDG: {typy['JDG']}, spolki: {typy['Spolka']})")
    for k, v in seg.most_common():
        print(f"  {v:5}  {k}")
    print(f"\nODRZUCONE: {len(odrzucone)}")
    for k, v in powody.most_common():
        print(f"  {v:5}  {k}")


if __name__ == "__main__":
    main()
