# ExDForm

Application de formulaire dynamique avec analyse statistique intégrée (PyQt6)

## Fonctionnalités principales
- **Import de variables** depuis des tableaux Word (.docx)
- **Enregistrement** dans une base SQLite locale
- **Export des données** en CSV/Excel
- **Analyse exploratoire** avec visualisations (distributions, tests de normalité)
- **Encodage automatique** des variables catégorielles multiples

**Note spéciale** : Les variables `CATEGORIELLE_MULTIPLE` sont automatiquement encodées en variables binaires (0/1) pour chaque modalité lors de l'export.

## Types de variables supportés

| Type                 | Format                      | Exemple               | Taille champ | Encodage auto |
|----------------------|----------------------------|-----------------------|--------------|---------------|
| `NUM_CONTINUE`       | Nombre décimal              | 12.34                 | Optionnel*   | -             |
| `NUM_DISCRETE`       | Nombre entier               | 42                    | Optionnel*   | -             |
| `TEXTE`              | Chaîne de caractères        | "Commentaire"         | **Requis**   | -             |
| `CATEGORIELLE`       | Modalités prédéfinies       | 1-A ,2-B              | -            | 1 ou  2       |
| `CATEGORIELLE_MULTIPLE` | Choix multiples          |  1-A,2-B              | -            | A(0/1)        |
| `DATE`               | Format JJ/MM/AAAA           | 15/07/2023            | -            | -             |
| `TEMPS`              | Format HH:MM:SS             | 14:30:00              | 8            | -             |

**Notes** :
- *Optionnel* : Limite le nombre de caractères saisis si spécifié
- **Requis** : Doit être renseigné pour les champs texte
- L'encodage auto transforme les modalités en variables numériques

## Exemple de tableau de variables (format Word)

| Nom variable    | Description          | Modalités              | Type                | Taille |
|-----------------|----------------------|------------------------|---------------------|--------|
| patient_id      | ID unique            |                        | NUM_DISCRETE        | 6      |
| nom             | Nom complet          |                        | TEXTE               | 50     |
| age             | Âge en années        |                        | NUM_DISCRETE        | 3      |
| groupe_sanguin  | GS du patient        | 1-A, 2-B, 3-AB, 4-O    | CATEGORIELLE        |        |
| allergies       | Allergies connues    | 1-Acariens, 2-Gluten   | CATEGORIELLE_MULTIPLE |      |
| date_visite     | Date de consultation |                        | DATE                |        |
| duree           | Durée (hh:mm:ss)     |                        | TEMPS               | 8      |

**Conventions** :
- Pour `TEXTE` : La taille correspond au nombre max de caractères
- Pour `NUM_*` : La taille indique le nombre max de chiffres
- `accents` : Non applicable sur les noms des variables

# Workflow typique :

Créer une base de données ou ouvrir une base de données SQLite locale existante 
Importer un tableau de variables depuis Word
Saisir les données via le formulaire généré automatiquement
Exporter en CSV ou analyser les données

utiliser les touches du clavier suivant pour naviguer
- **tab**: pour passer d'une variable a l'autre
- **espace** : pour derouler la liste des modalites puis **entrer** pour valider une modalite
