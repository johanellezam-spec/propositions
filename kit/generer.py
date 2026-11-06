#!/usr/bin/env python3
"""MediaHub Campus - génération d'une proposition web (page + Excel du réseau).

Usage :
    python3 kit/generer.py brief.json --tarifs regie_tarification_finale.xlsx --out sortie/

Produit dans --out :
    <slug>-<code>.html            la proposition (une seule page, images incluses)
    reseau-<slug>-<code>.xlsx     la liste des établissements
et affiche un résumé JSON (liens, chiffres, alertes) sur la sortie standard.
Dépendance : openpyxl.
"""
import argparse, base64, datetime, html, json, os, random, re, string, sys, unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

KIT = os.path.dirname(os.path.abspath(__file__))
A = os.path.join(KIT, "assets")
BASE_URL = "https://propositions.mediahubcampus.com"
NB, NBS = " ", " "
MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
        "septembre", "octobre", "novembre", "décembre"]
E = html.escape

def n(x, dec=0):
    return f"{x:,.{dec}f}".replace(",", " ").replace(".", ",").replace(" ", NB)
def eur(x, dec=2):
    return n(x, dec) + NBS + "€"
def img(path):
    ext = path.rsplit(".", 1)[-1].lower()
    mime = {"png": "image/png", "svg": "image/svg+xml"}.get(ext, "image/jpeg")
    return f"data:{mime};base64," + base64.b64encode(open(path, "rb").read()).decode()
def slugify(v):
    v = unicodedata.normalize("NFKD", v).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "", v) or "client"

PETITS = {"en", "de", "du", "des", "la", "le", "les", "sur", "sous", "aux", "et"}
def ville_fr(v):
    parts = re.split(r"([- ])", (v or "").strip().lower())
    out = []
    for i, p in enumerate(parts):
        if p in "- " or (i and p in PETITS):
            out.append(p)
        elif p[:2] in ("d'", "l'"):
            out.append(p[:2] + p[2:].capitalize())
        else:
            out.append(p.capitalize())
    return "".join(out)
def casse_fr(v):
    mots = re.split(r"(\s+)", (v or "").strip())
    out = []
    for i, m in enumerate(mots):
        lm = m.lower()
        if i and lm in {"de", "du", "des", "et", "en", "à", "au", "aux", "sur"}:
            out.append(lm)
        elif i and lm[:2] in ("d'", "d’", "l'", "l’"):
            out.append(lm[:2] + m[2:])
        else:
            out.append(m)
    return re.sub(r"-(En|De|Du|Sur|Les|La|Le)-", lambda m: "-" + m.group(1).lower() + "-", "".join(out))
def zone_label(z):
    return " – ".join(ville_fr(p) for p in z.split(" - "))

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("brief"); ap.add_argument("--tarifs", required=True); ap.add_argument("--out", default="sortie")
    a = ap.parse_args()
    B = json.load(open(a.brief, encoding="utf-8"))
    alertes = []

    client = B["client"].strip()
    slug = B.get("slug") or slugify(client)
    code = B.get("code") or "".join(random.choice(string.ascii_lowercase + string.digits) for _ in range(4))
    today = datetime.date.today()
    date_txt = B.get("date") or f"{today.day} {MOIS[today.month - 1]} {today.year}"
    sem = int(B.get("semaines", 4))
    duree_txt = f"{sem} semaine" + ("s" if sem > 1 else "")

    # --- Grille tarifaire ---
    wb = openpyxl.load_workbook(a.tarifs, read_only=True, data_only=True)
    tarifs = [r for r in wb["TARIFS_CAMPUS"].iter_rows(values_only=True)][1:]
    if B.get("zones"):
        voulues = [z.strip().upper() for z in B["zones"]]
        zones = [r for r in tarifs if r[3] and r[3].strip().upper() in voulues]
        manquantes = sorted(set(voulues) - {r[3].strip().upper() for r in zones})
        if manquantes:
            alertes.append("Zones absentes de la grille : " + ", ".join(manquantes))
    elif B.get("region"):
        zones = [r for r in tarifs if r[0] == B["region"]]
    else:
        sys.exit("Le brief doit contenir 'zones' (liste) ou 'region'.")
    if not zones:
        sys.exit("Aucune zone trouvée dans TARIFS_CAMPUS : vérifier l'orthographe (colonne Zone ou Région).")
    keys = {r[3] for r in zones}
    base = [r for r in wb["BASE_CAMPUS"].iter_rows(values_only=True)][1:]
    etabs = [r for r in base if r[13] in keys]

    coef = sem / 4
    if sem != 4:
        alertes.append(f"Durée de {sem} semaines : budget et ODV calculés au prorata de la grille 4 semaines, à valider par Johan.")
    z_budget = {r[3]: r[9] * coef for r in zones}
    sous_total = round(sum(z_budget.values()), 2)
    paliers = [(5000, .05), (10000, .10), (20000, .15), (30000, .20), (float("inf"), .25)]
    taux = next(t for plafond, t in paliers if sous_total <= plafond)
    remise = round(sous_total * taux, 2)
    total_ht = round(sous_total - remise, 2)
    tva = round(total_ht * .20, 2)
    ttc = round(total_ht + tva, 2)
    reach = sum(r[4] for r in zones)
    odv = sum(r[7] for r in zones) * coef
    affiches = sum(r[8] for r in zones)
    cpm_a = total_ht / reach * 1000
    cpm_e = total_ht / odv * 1000
    zs = sorted(zones, key=lambda r: zone_label(r[3]))

    # --- Excel du réseau ---
    cols = ["Zone", "Ville", "Code postal", "Adresse", "Établissement", "Typologie d’étude", "Catégorie", "Discipline"]
    rows = sorted(([zone_label(r[13]), ville_fr(r[2]), str(r[3] or "").strip(), casse_fr(r[4]), casse_fr(r[5]),
                    (r[6] or "").strip(), (r[7] or "").strip(), (r[8] or "").strip()] for r in etabs),
                  key=lambda x: (x[0], x[1], x[4]))
    os.makedirs(a.out, exist_ok=True)
    xlsx = f"reseau-{slug}-{code}.xlsx"
    ox = openpyxl.Workbook(); ws = ox.active; ws.title = f"Réseau {client}"[:31]
    ws.append(cols)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="22234C")
        c.alignment = Alignment(vertical="center")
    for i, w in enumerate([18, 22, 12, 42, 52, 40, 24, 26], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"; ws.auto_filter.ref = ws.dimensions
    ox.properties.creator = ox.properties.lastModifiedBy = "MediaHub Campus"
    ox.properties.title = f"Réseau {client} - {date_txt}"
    ox.save(os.path.join(a.out, xlsx))

    # --- Cas clients et photos ---
    CAS = json.load(open(os.path.join(A, "cas.json"), encoding="utf-8"))
    secteur = B.get("secteur", "marque")
    if B.get("cas"):
        choix = [c for i in B["cas"] for c in CAS if c["id"] == i]
    else:
        meme = sorted([c for c in CAS if c["secteur"] == secteur], key=lambda c: -c.get("note", 0))
        autres = sorted([c for c in CAS if c["secteur"] != secteur], key=lambda c: -c.get("note", 0))
        choix, vus = meme[:2], {c["secteur"] for c in meme[:2]}
        for c in autres:                      # varier les secteurs
            if len(choix) >= 4: break
            if c["secteur"] not in vus: choix.append(c); vus.add(c["secteur"])
        for c in meme[2:] + autres:
            if len(choix) >= 4: break
            if c not in choix: choix.append(c)
    photos, k = [], 0
    while len(photos) < 5 and k < 5:
        for c in choix:
            if k < len(c["photos"]) and len(photos) < 5:
                photos.append((c, c["photos"][k]))
        k += 1
    if len(photos) < 5:
        sys.exit("Pas assez de photos pour la mosaïque (5 nécessaires) : ajouter des cas ou en choisir d'autres.")
    mosaic = "<div class=\"mm\">" + "".join(
        f"<figure class=\"m{i + 1}\"><img loading=\"lazy\" src=\"{img(os.path.join(A, 'photos', p['f']))}\" "
        f"style=\"object-position:{p.get('pos', '50% 50%')}\" alt=\"Affiche {E(c['client'])} sur un campus\"></figure>"
        for i, (c, p) in enumerate(photos)) + "</div>"
    cases = "".join(f"<div class=\"cs\"><span class=\"tag\">{E(c['tag'])}</span><b>{E(c['client'])}</b>"
                    f"<p>{E(c['texte'])}</p></div>" for c in choix[:4])

    # --- Références (logos harmonisés, secteur du client en premier) ---
    LOG = json.load(open(os.path.join(A, "logos.json"), encoding="utf-8"))
    LOG = [l for l in LOG if l["secteur"] == secteur] + [l for l in LOG if l["secteur"] != secteur]
    wall = "".join(f"<div class='tile'><img loading='lazy' src='{img(os.path.join(A, 'logos', l['id'] + '.png'))}' "
                   f"alt='{E(l['nom'])}' style='width:{l['w']}px;height:{l['h']}px'></div>" for l in LOG)

    # --- Indicateurs ---
    masquer = set(B.get("masquer_cpm", []))
    kpis = [f"<div class=\"kpi hl\"><span class=\"k\">Reach – audience unique</span><b>{n(reach)}</b><p>étudiants touchés par votre campagne</p></div>",
            f"<div class=\"kpi\"><span class=\"k\">ODV – occasions de voir</span><b>{n(odv / 1e6, 1)}<small>millions</small></b><p>expositions potentielles, répétitions incluses</p></div>"]
    kpis.append(f"<div class=\"kpi\"><span class=\"k\">CPM audience</span><b>{eur(cpm_a)}</b><p>pour toucher 1{NB}000 étudiants</p></div>"
                if "audience" not in masquer else
                f"<div class=\"kpi\"><span class=\"k\">Établissements</span><b>{n(len(rows))}</b><p>lieux d’affichage sélectionnés</p></div>")
    kpis.append(f"<div class=\"kpi\"><span class=\"k\">CPM exposition</span><b>{eur(cpm_e)}</b><p>pour 1{NB}000 expositions</p></div>"
                if "exposition" not in masquer else
                f"<div class=\"kpi\"><span class=\"k\">Affiches</span><b>{n(affiches)}</b><p>au format A2, impression incluse</p></div>")

    max_aud = max(r[4] for r in zones)
    bars = "".join(f"<div class='bar'><span class='bl'>{E(zone_label(r[3]))}</span><span class='bt'><i style='width:{r[4] / max_aud * 100:.1f}%'></i></span>"
                   f"<span class='bv'>{n(r[4])}</span></div>" for r in sorted(zones, key=lambda r: -r[4]))
    pct = [(69, "révisent leurs examens sur leur campus"), (62, "travaillent en groupe à l’université"),
           (45, "profitent de la cafétéria de leur école"), (21, "font du sport régulièrement"),
           (14, "sont dans une association étudiante")]
    pct_html = "".join(f"<li><div class='ph'><b>{v}{NBS}%</b><span>{E(t)}</span></div><div class='pt'><i style='width:{v}%'></i></div></li>" for v, t in pct)
    zones_rows = "".join(f"<tr><td>{E(zone_label(r[3]))}</td><td class='num'>{n(r[8])}</td><td class='num'>{n(r[4])}</td>"
                         f"<td class='num'>{eur(z_budget[r[3]])}</td></tr>" for r in zs)
    villes = sorted({r[0] for r in rows}); cats = sorted({r[6] for r in rows if r[6]})
    periode = B.get("periode", "")
    zone_txt = B.get("zone_txt") or ("de " + ", ".join(zone_label(r[3]) for r in zs))
    web = f"{slug}-{code}.html"
    spark = ('<svg class="spark" viewBox="0 0 40 40" aria-hidden="true"><path d="M20 0 C22 14 26 18 40 20 '
             'C26 22 22 26 20 40 C18 26 14 22 0 20 C14 18 18 14 20 0Z"/></svg>')
    V = {
        "CLIENT": E(client), "DATE": E(date_txt), "LOGO": img(os.path.join(A, "logo.png")), "SPARK": spark,
        "CHAT": img(os.path.join(A, "chat_cover.png")), "PCT": pct_html, "MOSAIC": mosaic, "CASES": cases,
        "SOUS_TITRE": E(B.get("sous_titre") or f"Vos affiches sur les campus {zone_txt}" + (f", {periode}" if periode else "")),
        "VILLES_ZONE": E(", ".join(zone_label(r[3]) for r in zs)), "AFFICHES": n(affiches),
        "DATES_TXT": E(B.get("dates_txt") or f"{duree_txt}" + (f" {periode}" if periode else "") + ", dates à caler ensemble."),
        "ZONES_ROWS": zones_rows, "REACH": n(reach), "SOUS_TOTAL": eur(sous_total), "TOTAL_HT": eur(total_ht),
        "TAUX": n(taux * 100), "REMISE": eur(remise), "TVA": eur(tva), "TTC": eur(ttc), "DUREE_TXT": duree_txt,
        "KPIS": "\n".join(kpis), "BARS": bars, "PROFIL": E(B.get("profil", "Étudiants de 18 à 27 ans")),
        "FILIERES": E(B.get("filieres", "Toutes filières")), "FREQUENCE": E(B.get("frequence", "5 expositions par jour")),
        "NB_SEM": str(min(sem, 12)), "WEEKS": "".join(f"<span>S{i + 1}</span>" for i in range(min(sem, 12))),
        "PERIODE_COURTE": E(B.get("periode_courte", periode[:1].upper() + periode[1:] if periode else "")),
        "XLSX_URL": f"{BASE_URL}/{xlsx}", "OPT_ZONES": "".join(f"<option>{E(v)}</option>" for v in villes),
        "OPT_CATS": "".join(f"<option>{E(c)}</option>" for c in cats), "NB_ETAB": str(len(rows)),
        "WEB_URL": f"{BASE_URL}/{web}", "WEB_URL_TXT": f"{BASE_URL}/{web}".replace("https://", ""), "WALL": wall,
        "COLS_JSON": json.dumps(cols, ensure_ascii=False), "ROWS_JSON": json.dumps(rows, ensure_ascii=False),
    }
    page = open(os.path.join(KIT, "modele.html"), encoding="utf-8").read()
    for k2, v in V.items():
        page = page.replace(f"[[{k2}]]", v)
    reste = re.findall(r"\[\[[A-Z_]+\]\]", page)
    if reste:
        sys.exit("Champs non remplis dans le modèle : " + ", ".join(sorted(set(reste))))
    if re.search(r"claude|anthropic", page.replace(json.dumps(rows, ensure_ascii=False), ""), re.I):
        alertes.append("Le mot claude/anthropic apparaît dans la page : à vérifier avant envoi.")
    open(os.path.join(a.out, web), "w", encoding="utf-8").write(page)
    if cpm_a > 60:
        alertes.append(f"CPM audience élevé ({eur(cpm_a)}) : voir avec Johan s'il faut le remplacer (\"masquer_cpm\": [\"audience\"]).")

    print(json.dumps({
        "page": f"{BASE_URL}/{web}", "excel": f"{BASE_URL}/{xlsx}", "fichiers": [web, xlsx],
        "zones": [zone_label(r[3]) for r in zs], "etablissements": len(rows), "affiches": affiches, "reach": reach,
        "odv": round(odv), "sous_total_ht": sous_total, "degressif": f"{round(taux * 100)} %", "remise": remise,
        "total_ht": total_ht, "tva": tva, "ttc": ttc, "cpm_audience": round(cpm_a, 2), "cpm_exposition": round(cpm_e, 2),
        "cas_clients": [c["client"] for c in choix[:4]], "poids_ko": round(len(page) / 1024), "alertes": alertes,
    }, ensure_ascii=False, indent=1))

if __name__ == "__main__":
    main()
