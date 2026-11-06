# Kit proposition MediaHub Campus

Une proposition = **une page web** (présentation, tarifs, indicateurs, réseau filtrable, bouton Excel, bouton PDF)
publiée sur `https://propositions.mediahubcampus.com/<client>-<code>.html`.
Cette méthode remplace toutes les anciennes (PDF Canva, img2pdf, liens raw.githubusercontent, liens github.io).

## Contenu du kit (branche `main`, dossier `kit/`, jamais publié)

| Fichier | Rôle |
|---|---|
| `generer.py` | Génère la page et l'Excel à partir d'un brief JSON (0,5 s). Dépendance : openpyxl. |
| `modele.html` | Gabarit de la page. Ne pas modifier sans accord de Johan. |
| `assets/cas.json` + `assets/photos/` | Banque de cas clients : photos HD déjà recadrées, texte court, secteur. |
| `assets/logos.json` + `assets/logos/` | Logos clients harmonisés, avec leurs tailles. |
| `exemples/` | Briefs d'exemple. |

## Procédure (5 étapes)

1. **Tarifs** : télécharger la grille `regie_tarification_finale.xlsx` (Google Drive, id `1Az51E5ckKInaU59RTQ0kmt4vmOeXy1Xi`)
   et la décoder en `/tmp/tarifs.xlsx`. C'est la seule source des prix : ne jamais inventer un prix.
2. **Brief** : écrire `/tmp/brief.json` (voir les champs ci-dessous).
3. **Générer** : `python3 kit/generer.py /tmp/brief.json --tarifs /tmp/tarifs.xlsx --out /tmp/sortie`
   Le script affiche un résumé JSON : liens, chiffres, cas clients choisis, **alertes**. Toute alerte est signalée à Johan.
4. **Publier** (Alfred et Aurore ; Rambo ne publie pas) :
   ```bash
   git clone -q --depth 1 --filter=blob:none --no-checkout -b claude/propositions https://github.com/johanellezam-spec/propositions /tmp/pub
   cd /tmp/pub && git reset -q
   cp /tmp/sortie/*.html /tmp/sortie/*.xlsx /tmp/pub/
   git add -- $(cd /tmp/sortie && ls *.html *.xlsx)
   test -z "$(git diff --cached --diff-filter=DM --name-only)" || { echo "STOP : suppression ou modification détectée"; exit 1; }
   git commit -qm "Proposition <Client>" && git push -q origin HEAD:claude/propositions
   ```
   `git reset -q` est OBLIGATOIRE : sans lui, le commit supprimerait toutes les propositions déjà publiées.
   N'utiliser jamais `git add -A`, `git add .` ni `git commit -a`. Si le garde-fou affiche STOP, ne rien pousser et prévenir Johan.
   Si le clone partiel échoue, refaire le clone sans `--filter=blob:none --no-checkout` (et sans `git reset`).
   Ne jamais supprimer ni modifier un fichier déjà publié (les autres propositions sont chez les clients).
5. **Vérifier avant tout envoi** : `curl -s -o /dev/null -w "%{http_code}" <lien page>` puis `<lien excel>`,
   toutes les 20 secondes, 10 minutes au maximum, jusqu'à obtenir 200 pour les deux. Sans 200, on n'envoie pas
   et on prévient Johan. La publication prend en général 1 à 2 minutes.

Puis brouillon du mail : lien de la page seul (pas de pièce jointe), quelques chiffres clés, une invitation à répondre
par mail pour ajuster. Le client peut télécharger le PDF et l'Excel depuis la page.

## Champs du brief

```json
{
 "client": "Nom de l'annonceur",
 "region": "Provence-Alpes-Côte d’Azur",
 "zones": ["LYON", "GRENOBLE"],
 "zone_txt": "de la région PACA",
 "semaines": 4,
 "periode": "à la rentrée de janvier 2027",
 "dates_txt": "4 semaines en janvier 2027, dates à caler ensemble.",
 "periode_courte": "Rentrée de janvier 2027",
 "secteur": "marque",
 "cas": ["icn", "figaro"],
 "masquer_cpm": [],
 "profil": "Étudiants de 18 à 27 ans",
 "filieres": "Toutes filières"
}
```

- `region` **ou** `zones` (noms exacts de la colonne Zone de `TARIFS_CAMPUS`). Seul `client` est obligatoire en plus.
- `semaines` : la grille est en 4 semaines ; une autre durée est calculée au prorata et signalée en alerte (à valider par Johan).
- `secteur` : `marque`, `ecole`, `institution`, `culture`, `media`, `sante`. Sert à choisir les cas clients
  et à mettre les logos du même secteur en premier. `cas` force une sélection précise (ids de `assets/cas.json`).
- `masquer_cpm` : `["audience"]` et/ou `["exposition"]` remplace un CPM trop mauvais par un autre indicateur
  (règle de Johan : seulement pour les deux CPM). Le script alerte si le CPM audience dépasse 60 €.
- `code` et `date` sont automatiques (code aléatoire de 4 caractères, date du jour).

## Règles qui ne changent pas

- Rien dans la page ni dans le mail ne doit laisser penser à une IA. Aucun lien `raw.githubusercontent.com`,
  aucun lien `github.io` : uniquement `propositions.mediahubcampus.com`.
- Photos : jamais floues ni pixélisées, l'affiche du client toujours mise en valeur (voir la banque). On n'ajoute
  une photo à la banque qu'avec l'accord d'Atlas ou de Johan.
- Cas clients : uniquement des campagnes réelles, texte factuel lu sur l'affiche.
