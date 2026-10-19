#!/usr/bin/env python3
# Writes profile/js/egrid-data.js for the energy view in localsite/info (set=energy), drawn by profile/js/state-energy.js.
#
# Sources in this folder:
#   eGRID2023_vs_eGRID2024.xlsx - the 27 eGRID subregions with 2023 and 2024 emission rates
#   egrid2024_data_metric.xlsx  - ST24 (states, PSTATABB), US24 (national), SRL24 (subregions)
#                                 and PLNT24 (plants, for each state's generation share by subregion)
#
# Uses only the Python standard library (xlsx files are zipped XML).
# Run from the webroot: python3 useeio-json/egrid/egrid_to_js.py
# Online: https://github.com/ModelEarth/useeio-json/tree/main/egrid

import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET

HERE = os.path.dirname(os.path.abspath(__file__))
COMPARE_FILE = os.path.join(HERE, "eGRID2023_vs_eGRID2024.xlsx")
DATA_FILE = os.path.join(HERE, "egrid2024_data_metric.xlsx")
OUT_FILE = os.path.join(HERE, "..", "..", "profile", "js", "egrid-data.js")

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
M = "{%s}" % NS["m"]

# Resource mix columns share these codes: STGENACL (state coal), SRGENACL (subregion coal), USGENACL (US coal)
FUELS = [
    ("CL", "coal", "Coal"),
    ("OL", "oil", "Oil"),
    ("GS", "gas", "Gas"),
    ("NC", "nuclear", "Nuclear"),
    ("HY", "hydro", "Hydro"),
    ("BM", "biomass", "Biomass"),
    ("WI", "wind", "Wind"),
    ("SO", "solar", "Solar"),
    ("GT", "geothermal", "Geothermal"),
    ("OF", "otherFossil", "Other fossil"),
    ("OP", "other", "Other / purchased"),
]


def col_index(ref):
    n = 0
    for c in re.match(r"[A-Z]+", ref).group(0):
        n = n * 26 + ord(c) - 64
    return n - 1


def open_book(path):
    z = zipfile.ZipFile(path)
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    sheets = {}
    for s in wb.find("m:sheets", NS):
        t = targets[s.get("{%s}id" % NS["r"])].lstrip("/")
        sheets[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        strings = ["".join(t.text or "" for t in si.iter(M + "t")) for si in root.findall("m:si", NS)]
    except KeyError:
        strings = []
    return z, sheets, strings


def sheet_rows(book, name):
    z, sheets, strings = book
    with z.open(sheets[name]) as f:
        for _, el in ET.iterparse(f):
            if el.tag != M + "row":
                continue
            vals = {}
            for c in el.findall(M + "c"):
                t = c.get("t")
                v = c.find(M + "v")
                if t == "inlineStr":
                    val = "".join(x.text or "" for x in c.iter(M + "t"))
                elif v is None:
                    continue
                elif t == "s":
                    val = strings[int(v.text)]
                else:
                    val = v.text
                vals[col_index(c.get("r"))] = val
            yield [vals.get(i, "") for i in range(max(vals) + 1)] if vals else []
            el.clear()


def coded_rows(book, name):
    # Row 1 holds long labels, row 2 the column codes (PSTATABB, STNGENAN, ...)
    rows = sheet_rows(book, name)
    next(rows)
    codes = next(rows)
    for row in rows:
        if any(row):
            yield {code: (row[i] if i < len(row) else "") for i, code in enumerate(codes)}


def num(value, digits=None):
    try:
        n = float(value)
    except (TypeError, ValueError):
        return None  # eGRID uses "--" for not available
    if digits is None:
        return n
    n = round(n, digits)
    return int(n) if digits == 0 else n


def region_summary(row, p):
    # p is the column prefix: ST (state), SR (subregion) or US
    return {
        "capacity": num(row.get(p + "NAMEPCAP"), 1),  # MW
        "generation": num(row.get(p + "NGENAN"), 0),  # MWh
        "co2e": num(row.get(p + "CO2EQA"), 0),  # metric tons
        "co2eRate": num(row.get(p + "C2ERTA"), 2),  # kg/MWh
        "co2Rate": num(row.get(p + "CO2RTA"), 2),
        "noxRate": num(row.get(p + "NOXRTA"), 4),
        "so2Rate": num(row.get(p + "SO2RTA"), 4),
        "renewablePct": num(row.get(p + "TRPR"), 4),
        "nonhydroRenewablePct": num(row.get(p + "THPR"), 4),
        "mix": [num(row.get(p + "GENA" + code), 0) or 0 for code, _, _ in FUELS],  # MWh, in FUELS order
    }


def main():
    # Subregions with 2023 and 2024 rates. Columns 6-8 are the 2024 development branch,
    # which matches egrid2024_data_metric.xlsx (see metadata.txt).
    compare = open_book(COMPARE_FILE)
    rows = sheet_rows(compare, "Sheet1")
    next(rows), next(rows)
    subregions = []
    for r in rows:
        if not r or not r[0]:
            continue
        subregions.append({
            "id": r[0],
            "name": r[1],
            "capacity": num(r[2], 1),
            "co2Rate2024": num(r[6], 2),
            "ch4Rate2024": num(r[7], 4),
            "n2oRate2024": num(r[8], 4),
            "co2Rate2023": num(r[9], 2),
            "co2eRate2024": num(r[13], 2),
            "co2eRate2023": num(r[14], 2),
            "co2eChange": num(r[17], 4),  # fraction, 2024 development branch vs 2023
        })

    data = open_book(DATA_FILE)

    # Generation and renewables share for each subregion from SRL24
    srl = {row["SUBRGN"]: region_summary(row, "SR") for row in coded_rows(data, "SRL24")}
    for s in subregions:
        extra = srl.get(s["id"], {})
        s["generation"] = extra.get("generation")
        s["renewablePct"] = extra.get("renewablePct")
        s["mix"] = extra.get("mix")

    states = {}
    for row in coded_rows(data, "ST24"):
        summary = region_summary(row, "ST")
        summary["fips"] = row.get("FIPSST")
        states[row["PSTATABB"]] = summary

    us = region_summary(next(coded_rows(data, "US24")), "US")

    # Each state's net generation by subregion, from the plant file
    gen = {}
    cap = {}
    for row in coded_rows(data, "PLNT24"):
        st, sr = row.get("PSTATABB"), row.get("SUBRGN")
        if not st or not sr:
            continue
        key = (st, sr)
        gen[key] = gen.get(key, 0) + max(num(row.get("PLNGENAN")) or 0, 0)
        cap[key] = cap.get(key, 0) + (num(row.get("NAMEPCAP")) or 0)
    state_subregions = {}
    for st in sorted({k[0] for k in cap}):
        keys = [k for k in cap if k[0] == st]
        total_gen = sum(gen[k] for k in keys)
        total_cap = sum(cap[k] for k in keys)
        shares = []
        for k in keys:
            share = gen[k] / total_gen if total_gen > 0 else (cap[k] / total_cap if total_cap > 0 else 0)
            shares.append([k[1], round(share, 4), int(round(gen[k])), round(cap[k], 1)])
        shares.sort(key=lambda s: -s[1])
        state_subregions[st] = shares

    out = {
        "year": 2024,
        "compareYear": 2023,
        "fuels": [{"id": fid, "label": label} for _, fid, label in FUELS],
        "subregions": subregions,
        "us": us,
        "states": states,
        # [subregion, share of the state's net generation, net generation MWh, nameplate capacity MW]
        "stateSubregions": state_subregions,
    }

    with open(OUT_FILE, "w") as f:
        f.write("// eGRID 2024 electricity data for the energy view in localsite/info (set=energy), drawn by profile/js/state-energy.js.\n")
        f.write("// Generated by useeio-json/egrid/egrid_to_js.py from eGRID2023_vs_eGRID2024.xlsx and egrid2024_data_metric.xlsx.\n")
        f.write("// Script and workbooks: https://github.com/ModelEarth/useeio-json/tree/main/egrid\n")
        f.write("// Source: US EPA eGRID, https://www.epa.gov/egrid - regenerate rather than editing by hand.\n")
        f.write("// Units: capacity MW, generation MWh, co2e metric tons, rates kg/MWh, pct and change as fractions.\n")
        f.write("var egridData = ")
        json.dump(out, f, separators=(",", ":"))
        f.write(";\n")
    print("Wrote", os.path.relpath(OUT_FILE), len(subregions), "subregions,", len(states), "states")


if __name__ == "__main__":
    main()
