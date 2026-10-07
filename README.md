# Suivi GC5 2026-2027 : cahier de texte, présences et suivi

Classeur Excel avec macros (`.xlsm`) pour le délégué de la promotion GC5 BTP de l'EPAC, semestre 9.
Il regroupe :

- **le cahier de texte** : séances faites, date, heures réelles, présence du professeur, contenu traité ;
- **le cahier de présence** : appel en direct, avec les statuts P / A / R / E ;
- **le suivi des étudiants** : fiche individuelle, alertes d'absences, rapports, exposés ;
- **le suivi des enseignants** : séances tenues, absences, heures faites par rapport au volume prévu.

Configuration requise : Excel 2016, 2019, 2021 ou 365 sous Windows. Le classeur n'utilise pas les fonctions réservées à Excel 365 (XLOOKUP, FILTER, LET…) : les listes de consultation sont remplies par les macros.

> **Nouveau :** une feuille **MODE D'EMPLOI** pas à pas est incluse dans le classeur (2e onglet, et lien « ? Mode d'emploi » en haut de chaque feuille).

## Mise en route

1. **Débloquer le fichier.** Faites un clic droit sur `SUIVI_GC5_2026-2027.xlsm`, puis *Propriétés*, cochez **Débloquer** et validez avec *OK*. Sans cela, Windows bloque les macros d'un fichier téléchargé.
2. Ouvrez le classeur, puis cliquez sur **Activer le contenu**.
3. Allez dans **PARAMÈTRES** et indiquez votre nom dans « Délégué de classe ». Ce nom apparaît sur les PDF.
4. Vérifiez la liste des étudiants (**ÉTUDIANTS**) et les groupes par défaut.

> Mot de passe des feuilles protégées : `GC5BTP`. Le bouton « Mode administrateur » de l'accueil déverrouille tout.
> Pour changer le mot de passe, modifiez la constante `MDP` du module `modOutils` (Alt+F11).

## Utilisation au quotidien

| Moment | Action |
|---|---|
| Début du cours | **ACCUEIL**, puis **▶ DÉMARRER LE COURS**. La matière est retrouvée dans l'emploi du temps à partir du jour et de l'heure, avec 30 minutes de marge avant le début. La séance est créée dans le cahier de texte. |
| Pendant le cours | Dans **COURS DU JOUR** : indiquez la présence du professeur, le type de séance et le contenu. Faites l'appel : **✓ TOUS PRÉSENTS**, puis double-cliquez sur les exceptions (P → A → R → E). Un retard (R) enregistre l'heure d'arrivée. |
| Après l'appel | **✔ ENREGISTRER**. Les étudiants sans statut peuvent être marqués absents ou présents en un clic. |
| Fin du cours | **■ CLÔTURER**. L'heure de fin est enregistrée et la durée se calcule. |
| Correction | Choisissez le n° de séance dans la case « N° séance » de COURS DU JOUR. La séance et son appel se rechargent ; modifiez-les puis ENREGISTRER. |

La case **« Afficher le groupe »** filtre l'appel sur un seul groupe.

**Listes déroulantes.** Les matières se choisissent sous la forme « code – nom » (ex. « GEC2304 – Ponts »). Les enseignants et les remplaçants se choisissent dans la liste des enseignants (feuille MATIÈRES, bouton « Ajouter un enseignant »). Les autres choix ont aussi leur liste : groupes, étudiants, statuts, types de séance, présence du professeur, types de rapport.

## Les feuilles

| Feuille | Rôle |
|---|---|
| **ACCUEIL** | Cours du moment et prochain cours, cours du jour avec leur état, indicateurs, graphique d'avancement, accès rapide. |
| **COURS DU JOUR** | Saisie en direct de la séance et de l'appel. |
| **CAHIER DE TEXTE** | Toutes les séances, avec le nombre de présents, absents, retards et excusés, et le taux de présence. |
| **PRÉSENCES** | Registre détaillé : une ligne par étudiant et par séance. |
| **EXPOSÉS** | « Une ligne par groupe » pour la matière choisie, puis « Marquer présenté ». La synthèse indique les groupes passés et ceux qui restent, par matière. |
| **RAPPORTS** | Consignes de rapports (individuels ou de groupe), date et heure limite, bilan des remises. |
| **REMISES** | Une ligne par étudiant ou groupe attendu. « Marquer remis maintenant » horodate le dépôt. Le statut se calcule seul : *À temps*, *En retard*, *Non remis* ou *En attente*. |
| **FICHE ÉTUDIANT** | Tapez un nom ou une partie du nom (« akpa »), sans tenir compte des accents. La fiche affiche le nombre de séances et d'absences, le taux, le détail par matière avec les alertes, la liste des cours manqués et suivis, les rapports et les exposés. Elle s'exporte en PDF. |
| **SUIVI ENSEIGNANTS** | Par matière : séances tenues, absences et retards du professeur, heures faites et restantes, avancement, dernier contenu traité. À droite, choisissez un enseignant dans la liste pour voir ses chiffres et toutes ses séances. |
| **RÉCAP ABSENCES** | Tableau croisé étudiants × matières. Une case passe en **rouge au-delà de 3 absences** et en orange à 3. |
| **EXPORT PDF** | Présence par matière et par groupe (un PDF par groupe ou un seul), avec une période facultative. Exporte aussi le cahier de texte. |
| **ÉTUDIANTS / GROUPES** | Liste et groupes par défaut. Les groupes peuvent être différents pour une matière : bouton « Préparer les groupes de cette matière ». |
| **MATIÈRES / EMPLOI DU TEMPS / PARAMÈTRES** | Référentiels : codes, enseignants, volumes horaires, créneaux, seuil d'alerte, palette de couleurs. |

Les PDF sont rangés dans un dossier `Export_PDF\<code matière>\`, à côté du classeur. Le bouton « Sauvegarder » crée en plus une copie datée dans `Sauvegardes\`.

## Palette

| Couleur | Usage |
|---|---|
| Bleu marine `#1F3864` | Titres |
| Bleu `#2E75B6` | En-têtes |
| Bleu pâle `#F3F7FC` | Lignes alternées, cartes |
| Orange `#ED7D31` | Feuilles de saisie |
| Vert `#70AD47` | Feuilles de consultation |
| Gris | Référentiels |
| Rouge `#C00000` | Alertes |
| Jaune clair | Cellules à remplir |

Statuts : **P** vert, **A** rouge, **R** jaune, **E** gris. Chaque matière a sa couleur dans l'emploi du temps.

## Si les macros ne se chargent pas

Le projet VBA est intégré au fichier sous forme de code source, et Excel le compile à la première ouverture. Si Excel signale un problème de contenu ou si les boutons ne réagissent pas :

1. Ouvrez le VBA avec Alt+F11, puis *Fichier* > *Importer un fichier*, et importez les modules `vba/*.bas`.
2. Copiez le contenu des fichiers `vba/*.cls` dans le module de la feuille correspondante :
   - `shSaisie` : COURS DU JOUR
   - `shFiche` : FICHE ÉTUDIANT
   - `shExposes` : EXPOSÉS
   - `shRemises` : REMISES
   - `ThisWorkbook`
3. Enregistrez, fermez, puis rouvrez le classeur.

## Régénérer le classeur

```bash
pip install xlsxwriter openpyxl
python build/build_classeur.py SUIVI_GC5_2026-2027.xlsm --etudiants ANCIEN_CLASSEUR.xlsm
```

`--etudiants` lit la feuille « Liste » de l'ancien classeur. `--reprendre ANCIEN_SUIVI.xlsm` récupère les saisies d'une version précédente : séances, présences, exposés, rapports, remises, groupes, enseignants et paramètres. Sans cette option, la liste des étudiants est vide : c'est la version publiée dans ce dépôt, pour ne pas exposer de données personnelles.

Contenu du dépôt :

- `build/build_classeur.py` : feuilles, styles, formules, listes et mises en forme conditionnelles ;
- `build/vba_project.py` : génère le projet VBA (format MS-OVBA) à partir des sources ;
- `vba/` : code VBA, avec un module par fonctionnalité.
