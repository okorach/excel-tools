#!python3

import sys

months = {"janvier": 1, "février": 2, "mars": 3, "avril": 4, "mai": 5, "juin": 6,
        "juillet": 7, "août": 8, "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12}
jour = r"(0?\d|[12]\d|3[01])"
year = r"202\d"
no_cb, nom_cb = "CB Débit différé Elodie", "CB Débit différé Elodie"
step = 0
outline = ""
with open(sys.argv[1], "r", encoding="utf-8") as fd:
    lines = fd.readlines()
    lines = [line.rstrip() for line in lines]
    started = False
    for line in lines:
        # print(f"------------> Processing {line}")
        if len(line) == 0:
            continue
        splitted = line.split(" ")
        try:
            year = int(splitted[2])
            month = months[splitted[1]]
            day = int(splitted[0])
            step = 1
            current_date = f"{year:04d}-{month:02d}-{day:02d}"
            continue
        except (ValueError, IndexError):
            pass

        if step == 1:
            desc = line
            step += 1
        elif step == 2:
            categ = line
            step += 1
        elif step == 3:
            amount = float(line[:-2].replace(",", ".").replace(" ", "").replace("−", "-"))     # Strip off euro sign, comma and spaces
            step += 1
        if step == 4:
            outline += f'{current_date};{current_date};"{desc}";{categ};{categ};{amount};;{no_cb};{nom_cb}\n'
            print(f'{current_date};{current_date};"{desc}";{categ};{categ};{amount};;{no_cb};{nom_cb}')
            step = 1

with open(f"{sys.argv[1]}.csv", "w", encoding="utf-8") as fp:
    print(outline, file=fp)