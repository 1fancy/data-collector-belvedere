# Data Collector — Guide d'utilisation (client)

Application **Romana Immobilier** pour rassembler les données clients de la
Résidence Belvédère et préparer les lettres.

## Installation : il n'y en a pas

Il n'y a **rien à installer**. Vous recevez un dossier **« Data Collector »**.

1. Copiez ce dossier où vous voulez (Bureau, Documents, clé USB…).
2. Ouvrez le dossier et **double-cliquez sur `Data Collector`** (le fichier avec
   le logo Romana).
3. Une petite fenêtre noire s'ouvre, puis l'application s'ouvre **toute seule
   dans votre navigateur** (Chrome, Edge…).

> Au tout premier lancement, Windows peut afficher un avertissement bleu
> (« Windows a protégé votre PC »). C'est normal pour un programme interne.
> Cliquez sur **« Informations complémentaires »** puis **« Exécuter quand
> même »**. Les fois suivantes, il n'y a plus d'avertissement.

> Gardez la **fenêtre noire ouverte** tant que vous utilisez l'application.
> Pour quitter, fermez simplement cette fenêtre.

## Utilisation en 3 étapes

1. **Choisir le dossier** des ventes (bouton *Parcourir…*). C'est le dossier qui
   contient les sous-dossiers ILOT 1, ILOT 2, etc.
2. Cliquer **« Lancer le scan »**. Une barre de progression s'affiche.
3. Le tableau se remplit. Vous pouvez :
   - rechercher, filtrer, trier ;
   - cliquer **« Voir »** / **« Générer »** pour la lettre d'un client ;
   - cliquer sur **« … pièces »** pour voir la fiche détaillée et les documents ;
   - **Télécharger l'Excel** (nommé avec la date du jour).

Tout reste **sur votre ordinateur** : rien n'est envoyé sur internet.

## Questions fréquentes

**Faut-il internet ?** Non, tout fonctionne hors-ligne.

**Où sont mes données ?** Dans le dossier de l'application (fichier
`belvedere.db`). Elles restent d'une fois sur l'autre.

**L'application ne s'ouvre pas dans le navigateur ?** Ouvrez votre navigateur et
tapez l'adresse affichée dans la fenêtre noire (par ex. `http://127.0.0.1:8700`).

**Lecture des documents scannés (CIN, reçus) :** elle fonctionne automatiquement
sur Mac. Sur Windows, elle nécessite l'outil gratuit *Tesseract* (facultatif) ;
sans lui, l'application fonctionne normalement pour tout le reste.
