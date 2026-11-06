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
5. **Vérifier avant tout envoi** : depuis le cloud, le domaine propositions.mediahubcampus.com n'est pas joignable
   avec curl. On vérifie donc la publication auprès de GitHub :
   ```bash
   SHA=$(git rev-parse HEAD)
   curl -s "https://api.github.com/repos/johanellezam-spec/propositions/actions/runs?branch=claude/propositions&per_page=5"
   ```
   Chercher le run dont `head_sha` = `$SHA` ; attendre `status` = `completed` et `conclusion` = `success`
   (toutes les 20 secondes, 10 minutes au maximum). Sans succès, on n'envoie pas et on prévient Johan.
   La publication prend en général 1 à 2 minutes.

Puis brouillon du mail : lien de la page seul (pas de pièce jointe), quelques chiffres clés, une invitation à répondre
par mail pour ajuster. Le client peut télécharger le PDF et l'Excel depuis la page.

## Champs du brief

**Obligatoires** (le script refuse un brief incomplet) : lis-les dans la demande du client.

| Champ | Valeurs | Exemple |
|---|---|---|
| `client` | nom de l'annonceur | `"Tartiflette"` |
| `reseau` | `campus`, `lycees` ou `mixte` (campus + lycées) | `"mixte"` |
| `secteur` | `marque`, `ecole`, `institution`, `culture`, `media`, `sante` : le secteur DU CLIENT | `"ecole"` |
| `semaines` | durée demandée (4 si la demande ne dit rien, et le dire dans le mail) | `3` |

**Zones** (noms exacts de la colonne Zone des grilles) :
- `campus` : `{"zones": ["NANTES", "ANGERS"]}` ou `{"region": "Pays de la Loire"}` (grille TARIFS_CAMPUS)
- `lycees` : même format (grille TARIFS_LYCEES)
- mixte : renseigner les deux blocs.

**Facultatifs** : `periode` (tel que le client l'écrit, même flou : « en mars 2027 », « à la rentrée » ; jamais de date inventée, jamais une durée comme « sur 8 semaines » ; sans période, la page dit « dates à caler ensemble »), `zone_txt` (« de la région PACA »), `periode_courte`, `dates_txt`, `profil`, `filieres`, `frequence`,
`cas` (ids de `assets/cas.json` pour forcer les cas clients), `masquer_cpm` / `garder_cpm`, `sous_titre`, `code`, `date`.

CPM (règle de Johan, seulement pour ces deux indicateurs) : le kit remplace automatiquement un CPM audience
au-delà de 60 € et un CPM exposition au-delà de 1 € par un autre indicateur, et le signale en alerte.
`garder_cpm: ["audience"]` force l'affichage, `masquer_cpm` force le retrait.

Exemples complets dans `exemples/` : campus (Tartiflette), lycées, mixte.

Ce que le kit adapte tout seul selon `reseau` : sous-titre (campus / lycées / campus et lycées), cibles, profil
(étudiants / lycéens), section « Notre cible » (retirée pour les lycées seuls), indicateurs (ODV et CPM exposition
seulement pour le campus), détail par zone (« Nantes · lycées »), Excel. Les semaines (S1, S2…) suivent `semaines`.
Les cas clients suivent `secteur` et changent d'un client à l'autre.

Alertes à transmettre à Johan : durée différente de 4 semaines (prorata), CPM remplacé, réseau lycées
(sujets sensibles), écart entre le nombre de lycées de la grille et celui du réseau.

## Règles qui ne changent pas

- Rien dans la page ni dans le mail ne doit laisser penser à une IA. Aucun lien `raw.githubusercontent.com`,
  aucun lien `github.io` : uniquement `propositions.mediahubcampus.com`.
- Photos : jamais floues ni pixélisées, l'affiche du client toujours mise en valeur (voir la banque). On n'ajoute
  une photo à la banque qu'avec l'accord d'Atlas ou de Johan.
- Cas clients : uniquement des campagnes réelles, texte factuel lu sur l'affiche.
