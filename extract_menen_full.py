#!/usr/bin/env python3
"""Haalt alle bedrijven in Menen (BE) op uit OpenStreetMap met adres en telefoonnummer.

Uitvoer: data/menen_bedrijven.xlsx (opgemaakt Excel-bestand) en data/menen_bedrijven.csv.
Vereist: pip install openpyxl   daarna:  python extract_menen_full.py
"""
import csv, json, sys, time, urllib.parse, urllib.request, urllib.error
from pathlib import Path

OUT = Path(__file__).parent / "data" / "menen_bedrijven.csv"
XLSX = OUT.with_suffix(".xlsx")
MIRRORS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.private.coffee/api/interpreter",
]
QUERY = """
[out:json][timeout:180];
area["boundary"="administrative"]["admin_level"="8"]["name"="Menen"]["ISO3166-1"!~"."]->.a;
(
  nwr(area.a)["name"]["shop"];
  nwr(area.a)["name"]["office"];
  nwr(area.a)["name"]["craft"];
  nwr(area.a)["name"]["industrial"];
  nwr(area.a)["name"]["healthcare"];
  nwr(area.a)["name"]["tourism"~"hotel|guest_house|hostel|motel"];
  nwr(area.a)["name"]["amenity"~"restaurant|cafe|bar|pub|fast_food|bank|pharmacy|dentist|doctors|clinic|veterinary|fuel|car_wash|car_rental|driving_school|hairdresser|nightclub|cinema|ice_cream|biergarten|post_office|childcare|kindergarten"];
  nwr(area.a)["name"]["leisure"~"fitness_centre|sports_centre|dance|bowling_alley"];
);
out center tags;
"""
TAGS = ["shop", "office", "craft", "industrial", "healthcare", "amenity", "tourism", "leisure"]
FIELDS = ["naam", "categorie", "straat", "huisnummer", "postcode", "gemeente",
          "telefoon", "gsm", "email", "website", "openingsuren", "btw_nummer",
          "lat", "lon", "osm_url"]


def fetch():
    body = urllib.parse.urlencode({"data": QUERY}).encode()
    last = None
    for attempt in range(3):
        for url in MIRRORS:
            try:
                req = urllib.request.Request(url, body, {"User-Agent": "menen-extract/1.0"})
                with urllib.request.urlopen(req, timeout=240) as r:
                    return json.load(r)["elements"]
            except (urllib.error.URLError, TimeoutError, ValueError, KeyError) as e:
                last = e
                print(f"mislukt {url}: {e}", file=sys.stderr)
        time.sleep(20 * (attempt + 1))
    sys.exit(f"Geen Overpass-server bereikbaar: {last}")


def row(el):
    t = el.get("tags", {})
    c = el.get("center") or el
    cat = next((f"{k}={t[k]}" for k in TAGS if k in t), "")
    g = lambda *ks: next((t[k] for k in ks if t.get(k)), "")
    return {
        "naam": t.get("name", ""), "categorie": cat,
        "straat": t.get("addr:street", ""), "huisnummer": t.get("addr:housenumber", ""),
        "postcode": t.get("addr:postcode", ""), "gemeente": t.get("addr:city", "Menen"),
        "telefoon": g("phone", "contact:phone"), "gsm": g("mobile", "contact:mobile"),
        "email": g("email", "contact:email"), "website": g("website", "contact:website"),
        "openingsuren": t.get("opening_hours", ""),
        "btw_nummer": g("ref:vatin", "ref:VAT", "vat"),
        "lat": c.get("lat", ""), "lon": c.get("lon", ""),
        "osm_url": f"https://www.openstreetmap.org/{el['type']}/{el['id']}",
    }


HEADERS = {"naam": "Naam", "categorie": "Categorie", "straat": "Straat", "huisnummer": "Nr",
           "postcode": "Postcode", "gemeente": "Gemeente", "telefoon": "Telefoon", "gsm": "GSM",
           "email": "E-mail", "website": "Website", "openingsuren": "Openingsuren",
           "btw_nummer": "BTW-nummer", "lat": "Latitude", "lon": "Longitude", "osm_url": "Kaart (OSM)"}
WIDTHS = {"naam": 38, "categorie": 24, "straat": 28, "huisnummer": 8, "postcode": 10, "gemeente": 14,
          "telefoon": 20, "gsm": 20, "email": 34, "website": 40, "openingsuren": 40,
          "btw_nummer": 16, "lat": 11, "lon": 11, "osm_url": 14}


def write_xlsx(rows):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "Bedrijven Menen"
    ws.append([HEADERS[f] for f in FIELDS])
    for r in rows:
        ws.append([r[f] for f in FIELDS])
    head = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 24
    link = Font(color="0563C1", underline="single")
    for i, f in enumerate(FIELDS, 1):
        ws.column_dimensions[get_column_letter(i)].width = WIDTHS[f]
    for row in ws.iter_rows(min_row=2):
        for c, f in zip(row, FIELDS):
            c.alignment = Alignment(vertical="top", wrap_text=f in ("openingsuren", "naam"))
            v = str(c.value or "")
            if f == "website" and v:
                c.hyperlink = v if v.startswith("http") else "http://" + v
                c.font = link
            elif f == "email" and v:
                c.hyperlink = "mailto:" + v.split(";")[0].strip()
                c.font = link
            elif f == "osm_url":
                c.hyperlink, c.value, c.font = v, "Open kaart", link
            elif f in ("telefoon", "gsm"):
                c.number_format = "@"
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.zoomScale = 100
    wb.save(XLSX)


def main():
    rows = {r["osm_url"]: r for r in map(row, fetch())}
    rows = sorted(rows.values(), key=lambda r: (r["naam"].lower(), r["straat"]))
    OUT.parent.mkdir(exist_ok=True)
    with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, FIELDS, delimiter=";")
        w.writeheader()
        w.writerows(rows)
    write_xlsx(rows)
    tel = sum(1 for r in rows if r["telefoon"] or r["gsm"])
    print(f"{len(rows)} bedrijven opgeslagen in {OUT} ({tel} met telefoonnummer)")


if __name__ == "__main__":
    main()
