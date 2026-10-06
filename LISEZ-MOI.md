# Résidence Belvédère — Console clients & lettres

Logiciel local pour **Romana Immobilier**. Il lit les données clients
éparpillées dans un dossier de ventes, les rassemble dans un tableau clair,
produit un fichier Excel propre, et génère les **lettres de publipostage**
(titre foncier disponible) au modèle exact de la société.

---

## Démarrer

Aucune compétence technique requise.

- **Windows** : double-cliquez sur **`DEMARRER_Windows.bat`**
- **Mac** : double-cliquez sur **`DEMARRER_Mac.command`**

Le navigateur s'ouvre automatiquement sur l'interface. Laissez la fenêtre noire
ouverte pendant l'utilisation ; fermez-la pour quitter.

> La première fois, le logiciel installe tout seul deux petites bibliothèques
> (lecture Word / Excel). Si l'ordinateur n'a pas internet, elles sont déjà
> fournies dans le dossier `vendor/` et s'installent hors-ligne.
>
> Il faut simplement **Python 3** sur l'ordinateur. S'il manque, le logiciel
> ouvre la page de téléchargement et vous indique quoi faire.

## Copier sur un autre PC

Copiez **tout le dossier `belvedere_tool`** tel quel (clé USB, réseau…).
Il contient tout : le programme, le logo, le modèle de lettre et les
bibliothèques hors-ligne. Lancez le fichier `DEMARRER_…` sur le nouveau poste.
Rien n'est codé en dur à un emplacement : le logiciel fonctionne où qu'il soit.

## Adresse stable / HTTPS (optionnel, avec Herd)

Par défaut l'interface s'ouvre sur `http://127.0.0.1:8700` (port fixe, il ne
change plus). Si vous utilisez **Herd**, vous pouvez obtenir une adresse stable et
sécurisée, par ex. `https://belvedere.test` :

```bash
herd proxy belvedere http://127.0.0.1:8700 --secure
```

Lancez ensuite le logiciel ainsi (le port reste 8700) :

```bash
PUBLIC_URL="https://belvedere.test" python3 app.py
```

## Héberger / accès réseau

Le logiciel est un simple serveur web : il peut tourner sur un poste et être
consulté depuis d'autres machines du réseau. Démarrez-le en écoutant sur toutes
les interfaces :

```bash
HOST=0.0.0.0 PORT=8700 python3 app.py
```

puis ouvrez `http://ADRESSE-IP-DU-POSTE:8700` depuis un autre ordinateur. Le
dossier des ventes doit être accessible depuis la machine qui exécute le serveur.

---

## Où sont stockées les données

Les données sont enregistrées dans une base **SQLite** (`belvedere.db`, à côté du
programme) — c'est elle qui garde les clients, la répartition des co-acquéreurs
et les préférences (taille de page, etc.). Un fichier `data.json` lisible et
l'export `BELVEDERE_clients.xlsx` sont aussi produits. Rien n'est perdu entre
deux ouvertures.

## Utilisation

1. **Dossier des ventes** — à gauche, cliquez *Parcourir*. Vous pouvez&nbsp;:
   - **chercher un dossier par son nom** (champ de recherche en haut) — il est
     retrouvé automatiquement dans le Bureau, Documents, Téléchargements… ;
   - utiliser les **raccourcis** (Bureau, Documents, etc.) ;
   - ou naviguer dossier par dossier.
   On peut aussi coller directement le chemin dans le champ.
2. **Lancer le scan** — une barre de progression montre l'avancement en direct.
3. Le tableau se remplit. Recherchez, filtrez par type ou statut, triez les colonnes.
4. **Lettre** — cliquez *Générer* sur une ligne pour créer sa lettre (puis
   *Voir la lettre*). Cochez *« Générer les lettres automatiquement »* avant le
   scan pour toutes les produire d'un coup.
5. **Re-scan** — si des données existent déjà, une fenêtre demande&nbsp;:
   - **Tout remplacer** : les nouvelles données écrasent les anciennes.
   - **Garder et compléter** : on conserve l'existant, on ne remplit que les vides.
6. **Télécharger l'Excel** — tableau propre `BELVEDERE_clients.xlsx`.

Les lettres sont **mises en cache** : une lettre n'est régénérée que si les
informations du client ont changé depuis sa dernière création.

---

## D'où viennent les données

| Source | Informations extraites |
|---|---|
| Fiches notaire (`.doc`, `.docx`, `.txt`, `.pdf`) | Nom(s), C.I.N, adresse, téléphone, quote-part, type, superficie, prix, avance, reste |
| Chemin du dossier (`ILOT 1/IMM E APPT 13 ETAGE 3`) | Îlot, Immeuble, N°, Étage |
| Tableaux Excel | Ville, Pays, dates, observations (complètent les champs manquants) |

### Recherche approfondie (OCR)

Cochée par défaut, l'option **Recherche approfondie** ouvre *toutes* les pièces de
chaque dossier — pas seulement les deux fiches Word, mais aussi les **scans de
CIN, reçus, chèques, virements** — et en extrait automatiquement les
informations :

- confirmation du **numéro de C.I.N** à partir de la carte d'identité scannée ;
- **date de naissance**, **validité de la carte**, **noms des parents** ;
- **numéro et montant du reçu**.

Chaque bien affiche alors le nombre de **pièces** dans sa colonne *Dossier* ;
cliquez dessus pour voir la fiche complète du client et ouvrir chaque document.

> Le premier scan d'un dossier est plus long (lecture image par image), puis le
> résultat est **mis en cache** : les scans suivants sont instantanés.

**OCR sans installation sur macOS** : le logiciel utilise le moteur de
reconnaissance **intégré à macOS** (Apple Vision) — rien à installer, il lit le
français comme l'arabe.

**Sur Windows / Linux** : l'OCR fonctionne avec **Tesseract** (gratuit et
open-source). S'il n'est pas installé, le logiciel marche quand même : il lit les
fiches Word et les PDF textuels, et affiche simplement les scans sans en extraire
le texte. Pour activer l'OCR sous Windows : installez Tesseract depuis
<https://github.com/UB-Mannheim/tesseract/wiki>, puis relancez. Le logiciel le
détecte automatiquement.

Les pastilles en haut à droite indiquent ce que l'ordinateur sait lire
(Word, Excel, PDF, OCR) et quel moteur OCR est utilisé.

### D'où vient chaque donnée (provenance)

Dans la fiche d'un client (colonne *Dossier*), chaque information porte une petite
icône d'information. Au clic, elle indique **de quel(s) fichier(s)** provient la
donnée — une ou plusieurs sources. Une donnée confirmée par plusieurs documents
(ex. une C.I.N présente sur la fiche, la fiche comptable **et** la carte scannée)
est signalée en vert. Chaque source est **cliquable** (ouvre le document dans le
logiciel) et son **chemin est copiable** d'un clic.

## Lettres produites

Dossier `Letters/NNN_Nom_CIN/` par client :
- `lettre.docx` — document Word modifiable
- `lettre.html` — aperçu à l'écran et impression / enregistrement en PDF

Le document Word produit est **exactement le modèle fourni**
(`Lettre_Belvedere_modele_publipostage.docx`) : rien n'est réécrit ni réinventé.
Seuls les champs de fusion sont remplis — nom, adresse, ville, pays, date, type
de bien, îlot, immeuble, numéro. Le texte, la mise en page et la structure
restent identiques à l'original.

> La fiche de données vient de **deux sources par bien** : la fiche notaire et la
> fiche comptabilité. Si l'une est vide, l'autre complète (nom, C.I.N,
> quote-part, prix). C'est ainsi qu'aucun bien ne reste « à compléter » quand
> l'information existe quelque part.

---

## En ligne de commande (optionnel)

```bash
python3 extract.py  "/chemin/du/dossier"     # -> data.json + Excel
python3 letters.py  "/chemin/du/dossier"     # -> toutes les lettres (--force pour tout refaire)
python3 app.py                               # -> interface web
```
