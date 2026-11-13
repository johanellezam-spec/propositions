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
ACCENTS = {"ANNEÇY": "Annecy", "BESANCON": "Besançon", "BEZIER": "Béziers", "CHAMBERY": "Chambéry",
           "CHATEAUROUX": "Châteauroux", "EVREUX": "Évreux", "NIMES": "Nîmes", "ORLEANS": "Orléans",
           "SAINT-ETIENNE": "Saint-Étienne", "ÎLE-DE-FRANCE": "Île-de-France"}
def zone_label(z):
    z = (z or "").strip()
    if z.upper() in ACCENTS:
        return ACCENTS[z.upper()]
    return " – ".join(ville_fr(p) for p in z.split(" - "))

RESEAUX = {"campus": "sur les campus", "lycees": "dans les lycées", "mixte": "sur les campus et dans les lycées"}
AUD = {"campus": "étudiants", "lycees": "lycéens", "mixte": "étudiants et lycéens"}
CIBLES = {"campus": "Universités, écoles, résidences et restaurants universitaires, cafétérias, bibliothèques.",
          "lycees": "Lycées généraux, technologiques et professionnels.",
          "mixte": "Universités, écoles, résidences et restaurants universitaires ; lycées généraux, technologiques et professionnels."}
PROFILS = {"campus": "Étudiants de 18 à 27 ans", "lycees": "Lycéens de 15 à 18 ans", "mixte": "Lycéens et étudiants de 15 à 27 ans"}
FREQ = {"campus": "5 expositions par jour", "lycees": "Tous les jours de cours", "mixte": "Tous les jours de cours"}
SECTEURS = {"marque", "ecole", "institution", "culture", "media", "sante"}


def choisir_zones(tarifs, spec, nom, alertes):
    """spec = {"zones": [...]} ou {"region": "..."}"""
    if spec.get("france"):
        zs = [r for r in tarifs if r[3]]
    elif spec.get("zones"):
        voulues = [z.strip().upper() for z in spec["zones"]]
        zs = [r for r in tarifs if r[3] and r[3].strip().upper() in voulues]
        manq = sorted(set(voulues) - {r[3].strip().upper() for r in zs})
        if manq:
            alertes.append(f"Zones {nom} absentes de la grille : " + ", ".join(manq))
    elif spec.get("region"):
        zs = [r for r in tarifs if r[0] == spec["region"]]
    else:
        sys.exit(f"Le brief doit préciser les zones {nom} : \"{nom}\": {{\"france\": true}} ou {{\"zones\": [...]}} ou {{\"region\": \"...\"}}.")
    if not zs:
        sys.exit(f"Aucune zone {nom} trouvée dans la grille : vérifier l'orthographe (colonne Zone ou Région).")
    return zs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("brief"); ap.add_argument("--tarifs", required=True); ap.add_argument("--out", default="sortie")
    a = ap.parse_args()
    B = json.load(open(a.brief, encoding="utf-8"))
    alertes = []

    # --- Brief : champs obligatoires (lus dans la demande du client) ---
    manque = [k for k in ("client", "reseau", "secteur", "semaines") if not B.get(k)]
    if manque:
        sys.exit("Brief incomplet, champs obligatoires manquants : " + ", ".join(manque) +
                 ". Relis la demande : réseau (campus / lycees / mixte), secteur du client, nombre de semaines (4 si la demande ne dit rien).")
    reseau = B["reseau"].strip().lower().replace("é", "e")
    if reseau not in RESEAUX:
        sys.exit("reseau doit valoir campus, lycees ou mixte.")
    secteur = B["secteur"].strip().lower()
    if secteur not in SECTEURS:
        sys.exit("secteur doit valoir : " + ", ".join(sorted(SECTEURS)) + ".")
    client = B["client"].strip()
    slug = B.get("slug") or slugify(client)
    code = B.get("code") or "".join(random.choice(string.ascii_lowercase + string.digits) for _ in range(4))
    today = datetime.date.today()
    date_txt = B.get("date") or f"{today.day} {MOIS[today.month - 1]} {today.year}"
    sem = int(B["semaines"])
    duree_txt = f"{sem} semaine" + ("s" if sem > 1 else "")
    coef = sem / 4
    if sem != 4:
        alertes.append(f"Durée de {sem} semaines : budget et ODV calculés au prorata de la grille 4 semaines, à valider par Johan.")

    # --- Grilles tarifaires ---
    wb = openpyxl.load_workbook(a.tarifs, read_only=True, data_only=True)
    lignes, rows = [], []          # lignes de prix par zone ; établissements
    mix = reseau == "mixte"
    def norm(v):
        return unicodedata.normalize("NFKD", str(v or "")).encode("ascii", "ignore").decode().lower().replace("’", "'").strip()
    def filtre(spec, champs):
        """champs = {nom_du_filtre: index_colonne}. Garde une ligne si chaque filtre demandé correspond (sous-chaîne).
        "texte" / "exclure" : mots cherchés dans le nom de l'établissement et sa typologie (colonnes champs["_txt"]) ;
        la ligne est gardée si un mot de "texte" y figure et aucun mot de "exclure"."""
        txt = champs.get("_txt", ())
        f = {k: [norm(x) for x in spec.get(k, [])] for k in champs if k != "_txt" and spec.get(k)}
        inc = [norm(x) for x in spec.get("texte", [])] if txt else []
        exc = [norm(x) for x in spec.get("exclure", [])] if txt else []
        if inc: f["texte"] = inc
        if exc: f["exclure"] = exc
        ids = {" | ".join(norm(x) for x in k.split("|")) for k in spec.get("lignes", [])} if txt else set()
        if ids: f["lignes"] = sorted(ids)
        def ok(r):
            t = " | ".join(norm(r[i]) for i in txt)
            if ids and " | ".join(norm(r[i]) for i in (2,) + tuple(txt)) in ids:
                return True
            if ids and not (inc or exc or any(k not in ("texte", "exclure", "lignes") for k in f)):
                return False
            return (all(any(x in norm(r[champs[k]]) for x in vals) for k, vals in f.items() if k not in ("texte", "exclure", "lignes"))
                    and (not inc or any(x in t for x in inc)) and not any(x in t for x in exc))
        return ok, f
    ciblages = []
    def ajouter(nom, tarifs_feuille, base_feuille, spec, col_zone, champs, i_aud, i_odv, i_aff, i_bud, ligne_etab, suffixe):
        t = [r for r in wb[tarifs_feuille].iter_rows(values_only=True)][1:]
        zs = choisir_zones(t, spec, nom, alertes)
        keys = {r[3]: r for r in zs}
        base = [r for r in wb[base_feuille].iter_rows(values_only=True)][1:]
        ok, f = filtre(spec, champs)
        if f:
            ciblages.extend(v for k2, vals in spec.items() if isinstance(vals, list) and k2 not in ("texte", "exclure", "lignes")
                            for v in vals if v not in spec.get("zones", []))
        par_zone = {}
        for r in base:
            if r[col_zone] in keys:
                par_zone.setdefault(r[col_zone], [0, []])
                par_zone[r[col_zone]][0] += 1
                if ok(r):
                    par_zone[r[col_zone]][1].append(r)
        cibles = sum(len(v[1]) for v in par_zone.values())
        if f and not cibles:
            sys.exit(f"Ciblage {nom} sans résultat ({f}) : vérifier les valeurs (voir la section Ciblage du LISEZMOI).")
        for z, r in keys.items():
            total, gardes = par_zone.get(z, [0, []])
            ratio = (len(gardes) / total) if (f and total) else (1 if not f else 0)
            if ratio == 0:
                continue
            aff = r[i_aff] or 0
            lignes.append({"zone": zone_label(z), "label": zone_label(z) + (suffixe if mix else ""),
                           "aud": round((r[i_aud] or 0) * ratio), "odv": (r[i_odv] or 0) * coef * ratio if i_odv is not None else None,
                           "aff": max(len(gardes), round(aff * ratio)) if f else aff, "budget": round((r[i_bud] or 0) * coef * ratio, 2)})
            for e in gardes if f else [x for x in base if x[col_zone] == z]:
                rows.append(ligne_etab(e))
        if f:
            alertes.append(f"Ciblage {nom} appliqué ({cibles} établissements) : affiches, audience, budget et indicateurs au prorata des établissements ciblés de chaque zone.")
        return zs

    if reseau in ("campus", "mixte"):
        spec = B.get("campus") or {k: B[k] for k in ("zones", "region", "france") if B.get(k)}
        ajouter("campus", "TARIFS_CAMPUS", "BASE_CAMPUS", spec, 13,
                {"categories": 7, "disciplines": 8, "villes": 2, "typologies": 6, "etablissements": 5, "_txt": (5, 6)},
                4, 7, 8, 9,
                lambda r: [zone_label(r[13]), ville_fr(r[2]), str(r[3] or "").strip(), casse_fr(r[4]), casse_fr(r[5]),
                           (r[6] or "").strip(), (r[7] or "").strip(), (r[8] or "").strip()], " · campus")
    if reseau in ("lycees", "mixte"):
        spec = B.get("lycees") or ({k: B[k] for k in ("zones", "region", "france") if B.get(k)} if reseau == "lycees" else {})
        zl = ajouter("lycees", "TARIFS_LYCEES", "BASE_LYCEES", spec, 13,
                     {"filieres": 15, "villes": 3, "specialites": 6, "etablissements": 0},
                     8, None, 5, 6,
                     lambda r: [zone_label(r[13]), ville_fr(r[3]), str(r[2] or "").strip(), casse_fr(r[1]), casse_fr(r[0]),
                                "Lycée", "Lycée " + (r[15] or "").strip().lower(), (r[6] or "").strip()], " · lycées")
        ng, nr = sum(r[4] or 0 for r in zl), sum(r[7] or 0 for r in zl)
        if ng != nr and not any(spec.get(k) for k in ("filieres", "villes", "specialites", "etablissements")):
            alertes.append(f"Grille lycées : prix et affiches calculés sur {ng} lycées, alors que le réseau en liste {nr} : à valider par Johan avant envoi.")
    rows.sort(key=lambda x: (x[0], x[1], x[4]))
    lignes.sort(key=lambda l: l["label"])

    sous_total = round(sum(l["budget"] for l in lignes), 2)
    paliers = [(5000, .05), (10000, .10), (20000, .15), (30000, .20), (float("inf"), .25)]
    taux = next(t for plafond, t in paliers if sous_total <= plafond)
    remise = round(sous_total * taux, 2)
    total_ht = round(sous_total - remise, 2)
    tva = round(total_ht * .20, 2)
    ttc = round(total_ht + tva, 2)
    reach = sum(l["aud"] for l in lignes)
    affiches = sum(l["aff"] for l in lignes)
    odv = sum(l["odv"] for l in lignes if l["odv"]) if reseau == "campus" else None
    cpm_a = total_ht / reach * 1000 if reach else 0
    cpm_e = total_ht / odv * 1000 if odv else None

    # --- Reprise à l'identique d'une ancienne proposition (copie stricte, consigne de Johan du 09/11/2026) ---
    # "reprise": {"degressif": 20, "odv": 1416000, "affiches": 177, "reach": 23400, "total_ht": 1244}
    # degressif et odv sont imposés ; affiches, reach et total_ht sont des contrôles : tout écart arrête le script.
    R = B.get("reprise") or {}
    if R and R.get("reseau_xlsx"):
        # Réseau choisi à la main dans l'ancienne proposition : on reprend sa liste d'établissements telle quelle
        # et ses totaux (sous_total, affiches, reach) ; le détail par zone devient une seule ligne (zones_txt).
        manq = [k for k in ("sous_total", "affiches", "reach", "zones_txt") if k not in R]
        if manq:
            sys.exit("Reprise avec reseau_xlsx : champs manquants dans reprise : " + ", ".join(manq) + ".")
        old = [r for r in openpyxl.load_workbook(R["reseau_xlsx"], read_only=True).active.iter_rows(values_only=True)]
        h = [norm(x) for x in old[0]]
        def col(*noms):
            for nm in noms:
                for j, x in enumerate(h):
                    if x.startswith(norm(nm)):
                        return j
            return None
        ix = [col("zone", "ville"), col("ville"), col("code postal"), col("adresse"), col("etablissement", "ecole"),
              col("typologie"), col("categorie"), col("discipline")]
        rows = sorted([[str(r[j]).strip() if j is not None and r[j] is not None else "" for j in ix] for r in old[1:] if any(r)],
                      key=lambda x: (x[0], x[1], x[4]))
        sous_total, affiches, reach = round(float(R["sous_total"]), 2), int(R["affiches"]), int(R["reach"])
        lignes = [{"label": R["zones_txt"], "zone": R["zones_txt"], "aff": affiches, "aud": reach, "budget": sous_total, "odv": None}]
        odv = None
        remise = round(sous_total * taux, 2); total_ht = round(sous_total - remise, 2)
        tva = round(total_ht * .20, 2); ttc = round(total_ht + tva, 2)
        alertes.append(f"Reprise du réseau de l'ancienne proposition ({len(rows)} établissements) et de ses totaux, sans recalcul par la grille.")
    if R:
        if "degressif" in R:
            taux = R["degressif"] / 100
            remise = round(sous_total * taux, 2)
            total_ht = round(sous_total - remise, 2)
            tva = round(total_ht * .20, 2)
            ttc = round(total_ht + tva, 2)
        if R.get("odv"):
            odv = R["odv"]
        imp = [k for k in R.get("imposer", []) if k in ("reach", "affiches")]
        if imp:
            # Valeur de l'ancienne proposition différente de la grille : on la recopie telle quelle et, pour rester
            # cohérent, le détail par zone devient une seule ligne (zones_txt).
            if "reach" in imp: reach = int(R["reach"])
            if "affiches" in imp: affiches = int(R["affiches"])
            zt = R.get("zones_txt") or ", ".join(sorted({l["zone"] for l in lignes}))
            lignes = [{"label": zt, "zone": zt, "aff": affiches, "aud": reach, "budget": sous_total, "odv": None}]
            alertes.append(f"Valeurs imposées depuis l'ancienne proposition ({', '.join(imp)}) : différentes de la grille, à signaler à Johan.")
        cpm_a = total_ht / reach * 1000 if reach else 0
        cpm_e = total_ht / odv * 1000 if odv else None
        ecarts = [f"{k} : ancienne proposition {R[k]}, kit {v}" for k, v in
                  (("affiches", affiches), ("reach", reach), ("total_ht", total_ht), ("total_ttc", ttc))
                  if k in R and round(float(R[k]), 2) != round(float(v), 2)]
        if ecarts:
            sys.exit("Reprise impossible à l'identique, rien n'est généré : " + " ; ".join(ecarts) +
                     ". Relancer avec l'ancien lien et prévenir Johan.")
        alertes.append("Reprise d'une ancienne proposition : remise et ODV imposés, affiches, audience et total HT contrôlés.")

    # --- Excel du réseau ---
    cols = ["Zone", "Ville", "Code postal", "Adresse", "Établissement", "Typologie d’étude", "Catégorie", "Discipline"]
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

    # --- Cas clients et photos (secteur du client d'abord, ordre propre à chaque client) ---
    CAS = json.load(open(os.path.join(A, "cas.json"), encoding="utf-8"))
    rng = random.Random(slug)
    cle = lambda c: (0 if c.get("note", 0) >= 8 else 1, rng.random())
    if B.get("cas"):
        choix = [c for i in B["cas"] for c in CAS if c["id"] == i]
    else:
        meme = sorted([c for c in CAS if c["secteur"] == secteur], key=cle)
        autres = sorted([c for c in CAS if c["secteur"] != secteur], key=cle)
        choix, vus = meme[:2], {c["secteur"] for c in meme[:2]}
        for c in autres:
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
    for c in CAS:                       # compléter si les cas choisis ont trop peu de photos
        for p in c["photos"]:
            if len(photos) < 5 and (c, p) not in photos and c in choix[:4]:
                photos.append((c, p))
    if len(photos) < 5:
        extra = [(c, p) for c in sorted(CAS, key=cle) if c not in choix for p in c["photos"]]
        photos += extra[:5 - len(photos)]
    mosaic = "<div class=\"mm\">" + "".join(
        f"<figure class=\"m{i + 1}\"><img loading=\"lazy\" src=\"{img(os.path.join(A, 'photos', p['f']))}\" "
        f"style=\"object-position:{p.get('pos', '50% 50%')}\" alt=\"Affiche {E(c['client'])}\"></figure>"
        for i, (c, p) in enumerate(photos)) + "</div>"
    cases = "".join(f"<div class=\"cs\"><span class=\"tag\">{E(c['tag'])}</span><b>{E(c['client'])}</b>"
                    f"<p>{E(c['texte'])}</p></div>" for c in choix[:4])

    # --- Références (logos harmonisés, secteur du client en premier) ---
    LOG = json.load(open(os.path.join(A, "logos.json"), encoding="utf-8"))
    LOG = [l for l in LOG if l["secteur"] == secteur] + [l for l in LOG if l["secteur"] != secteur]
    wall = "".join(f"<div class='tile'><img loading='lazy' src='{img(os.path.join(A, 'logos', l['id'] + '.png'))}' "
                   f"alt='{E(l['nom'])}' style='width:{l['w']}px;height:{l['h']}px'></div>" for l in LOG)

    # --- Indicateurs ---
    aud = AUD[reseau]
    masquer = set(B.get("masquer_cpm", []))
    garder = set(B.get("garder_cpm", []))
    if cpm_a > 60 and "audience" not in garder and "audience" not in masquer:
        masquer.add("audience")
        alertes.append(f"CPM audience de {eur(cpm_a)} jugé trop élevé : remplacé automatiquement sur la page (règle de Johan). Pour l'afficher quand même : \"garder_cpm\": [\"audience\"].")
    if cpm_e and cpm_e > 1 and "exposition" not in garder and "exposition" not in masquer:
        masquer.add("exposition")
        alertes.append(f"CPM exposition de {eur(cpm_e)} jugé trop élevé : remplacé automatiquement sur la page.")
    kpi = lambda k, v, t, hl=False: f"<div class=\"kpi{' hl' if hl else ''}\"><span class=\"k\">{k}</span><b>{v}</b><p>{t}</p></div>"
    cand = [kpi("Reach – audience unique", n(reach), f"{aud} touchés par votre campagne", True)]
    if odv:
        cand.append(kpi("ODV – occasions de voir", f"{n(odv / 1e6, 1)}<small>millions</small>", "expositions potentielles, répétitions incluses"))
    if "audience" not in masquer:
        cand.append(kpi("CPM audience", eur(cpm_a), f"pour toucher 1{NB}000 {aud}"))
    if odv and "exposition" not in masquer:
        cand.append(kpi("CPM exposition", eur(cpm_e), f"pour 1{NB}000 expositions"))
    cand += [kpi("Établissements", n(len(rows)), "lieux d’affichage sélectionnés"),
             kpi("Affiches", n(affiches), "au format A2, impression incluse"),
             kpi("Villes", n(len({l["zone"] for l in lignes})), "zones couvertes par la campagne")]
    kpis = cand[:4]

    max_aud = max(l["aud"] for l in lignes) or 1
    bars = "".join(f"<div class='bar'><span class='bl'>{E(l['label'])}</span><span class='bt'><i style='width:{l['aud'] / max_aud * 100:.1f}%'></i></span>"
                   f"<span class='bv'>{n(l['aud'])}</span></div>" for l in sorted(lignes, key=lambda l: -l["aud"]))
    pct = [(69, "révisent leurs examens sur leur campus"), (62, "travaillent en groupe à l’université"),
           (45, "profitent de la cafétéria de leur école"), (21, "font du sport régulièrement"),
           (14, "sont dans une association étudiante")]
    pct_html = "".join(f"<li><div class='ph'><b>{v}{NBS}%</b><span>{E(t)}</span></div><div class='pt'><i style='width:{v}%'></i></div></li>" for v, t in pct)
    cible = ("" if reseau == "lycees" else
             f"<section id=\"cible\"><div class=\"wrap cible\"><div><div class=\"eyebrow\">Leurs habitudes</div>"
             f"<h2 class=\"h2\">Notre <em>cible</em></h2><div class=\"src\">Source{NBS}: étude Iligo x DYL, 2023</div></div>"
             f"<ul class=\"pcts\">{pct_html}</ul></div></section>")
    zones_rows = "".join(f"<tr><td>{E(l['label'])}</td><td class='num'>{n(l['aff'])}</td><td class='num'>{n(l['aud'])}</td>"
                         f"<td class='num'>{eur(l['budget'])}</td></tr>" for l in lignes)
    zones_uniques = sorted({l["zone"] for l in lignes})
    villes = sorted({r[0] for r in rows}); cats = sorted({r[6] for r in rows if r[6]})
    periode = (B.get("periode") or "").strip()
    if re.search(r"\bsemaines?\b|caler|définir|a definir|preciser|préciser", periode, re.I):   # durée ou formule vide : pas une période
        periode = ""
    liste = ", ".join(zones_uniques[:-1]) + (" et " if len(zones_uniques) > 1 else "") + zones_uniques[-1]
    national = any((B.get(k) or {}).get("france") for k in ("campus", "lycees")) or B.get("france")
    zone_txt = B.get("zone_txt") or ("de toute la France" if national else (("d’" if liste[:1].lower() in "aeiouyhéèêâîô" else "de ") + liste))
    JOLI = {"école d'ingénieur": "écoles d'ingénieurs", "école de commerce": "écoles de commerce", "pro": "lycées professionnels",
            "général": "lycées généraux", "techno": "lycées technologiques", "sciences & ingénierie": "sciences et ingénierie",
            "iut et but": "IUT", "résidence étudiante": "résidences étudiantes", "restauration étudiante": "restaurants universitaires",
            "bibliothèque universitaire": "bibliothèques universitaires", "commerce & management": "commerce et management",
            "droit & sciences po": "droit et sciences politiques", "santé & paramédical": "santé et paramédical",
            "lettres & sciences humaines": "lettres et sciences humaines", "arts & design": "arts et design",
            "journalisme & communication": "journalisme et communication", "sport (staps)": "sport (STAPS)"}
    ciblage_txt = B.get("ciblage_txt") or ", ".join(dict.fromkeys(JOLI.get(c.lower().replace("’", "'"), c.lower()) for c in ciblages))
    lieux_types = ("écoles", "IUT", "résidences", "restaurants", "bibliothèques", "universités", "lycées")
    if ciblage_txt and reseau != "mixte" and ciblage_txt.startswith(lieux_types):
        lieu = "dans les " + ciblage_txt
    elif ciblage_txt:
        lieu = f"{RESEAUX[reseau]} ({ciblage_txt})"
    else:
        lieu = RESEAUX[reseau]
    if B.get("sous_titre"):
        alertes.append("Le champ sous_titre est ignoré : le sous-titre est toujours généré à partir du réseau et du ciblage réels.")
    web = f"{slug}-{code}.html"
    spark = ('<svg class="spark" viewBox="0 0 40 40" aria-hidden="true"><path d="M20 0 C22 14 26 18 40 20 '
             'C26 22 22 26 20 40 C18 26 14 22 0 20 C14 18 18 14 20 0Z"/></svg>')
    V = {
        "CLIENT": E(client), "DATE": E(date_txt), "LOGO": img(os.path.join(A, "logo.png")), "SPARK": spark,
        "CHAT": img(os.path.join(A, "chat_cover.png")), "CIBLE": cible, "MOSAIC": mosaic, "CASES": cases,
        "SOUS_TITRE": E(f"Vos affiches {lieu} {zone_txt}" + (f", {periode}" if periode else "")),
        "CIBLES_TXT": E(B.get("cibles_txt") or CIBLES[reseau]),
        "VILLES_ZONE": E(", ".join(zones_uniques)), "AFFICHES": n(affiches),
        "MATERIEL_EXTRA": (" " + E(B["materiel_extra"])) if B.get("materiel_extra") else "",
        "DATES_TXT": E(B.get("dates_txt") or f"{duree_txt}" + (f" {periode}" if periode else "") + ", dates à caler ensemble."),
        "ZONES_ROWS": zones_rows, "REACH": n(reach), "SOUS_TOTAL": eur(sous_total), "TOTAL_HT": eur(total_ht),
        "TAUX": n(taux * 100), "REMISE": eur(remise), "TVA": eur(tva), "TTC": eur(ttc), "DUREE_TXT": duree_txt,
        "KPIS": "\n".join(kpis), "BARS": bars, "PROFIL": E(B.get("profil") or PROFILS[reseau]),
        "FILIERES": E(B.get("filieres_txt") or (ciblage_txt[:1].upper() + ciblage_txt[1:] if ciblage_txt else "Toutes filières")), "FREQUENCE": E(B.get("frequence") or FREQ[reseau]),
        "NB_SEM": str(min(sem, 12)), "WEEKS": "".join(f"<span>S{i + 1}</span>" for i in range(min(sem, 12))),
        "PERIODE_COURTE": E(B.get("periode_courte") or (periode[:1].upper() + periode[1:] if periode else "Dates à caler ensemble")),
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
    if reseau == "lycees" or reseau == "mixte":
        alertes.append("Réseau lycées : certains sujets ne peuvent pas être abordés auprès des lycéens ; vérifier le message avec Johan.")
    open(os.path.join(a.out, web), "w", encoding="utf-8").write(page)

    print(json.dumps({
        "page": f"{BASE_URL}/{web}", "excel": f"{BASE_URL}/{xlsx}", "fichiers": [web, xlsx], "reseau": reseau,
        "zones": [l["label"] for l in lignes], "etablissements": len(rows), "affiches": affiches, "reach": reach,
        "odv": round(odv) if odv else None, "sous_total_ht": sous_total, "degressif": f"{round(taux * 100)} %", "remise": remise,
        "total_ht": total_ht, "tva": tva, "ttc": ttc, "cpm_audience": round(cpm_a, 2),
        "cpm_exposition": round(cpm_e, 2) if cpm_e else None, "semaines": sem,
        "cas_clients": [c["client"] for c in choix[:4]], "poids_ko": round(len(page) / 1024), "alertes": alertes,
    }, ensure_ascii=False, indent=1))

if __name__ == "__main__":
    main()
