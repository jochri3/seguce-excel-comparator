# Application de Réconciliation paie

Une application web Node.js permettant de comparer et réconcilier deux fichiers Excel contenant des données APEI (Allocations Pour Enfants et Invalides).

## Fonctionnalités

- Upload de deux fichiers Excel
- Analyse comparative des données
- Identification des différences entre les fichiers
- Affichage détaillé des écarts par colonne et par ligne
- Interface utilisateur intuitive et responsive
- Possibilité d'imprimer ou d'exporter les résultats

## Technologies utilisées

- **Backend** : Node.js avec Express
- **Frontend** : EJS, Bootstrap 5, JavaScript
- **Traitement Excel** : xlsx.js
- **Upload de fichiers** : Multer

## Structure du projet

```
/paie-reconciliation
│
├── /public              # Fichiers statiques
│   ├── /css             # Feuilles de style
│   └── /js              # Scripts JavaScript
│
├── /views               # Templates EJS
│   ├── index.ejs        # Page d'accueil
│   ├── compare.ejs      # Page de résultats
│   ├── error.ejs        # Page d'erreur
│   └── /partials        # Éléments réutilisables
│
├── /uploads             # Dossier pour les fichiers uploadés
│
├── /utils               # Utilitaires
│   └── excel-parser.js  # Module de traitement Excel
│
├── app.js               # Point d'entrée de l'application
└── package.json         # Dépendances et scripts
```

## Installation

1. Cloner le dépôt ou télécharger les fichiers sources
2. Installer les dépendances avec `npm install`
3. Démarrer l'application avec `npm start` ou `npm run dev` (mode développement)
4. Accéder à l'application via `http://localhost:3000`

Pour des instructions détaillées, consultez le [Guide d'installation et d'utilisation](INSTALLATION.md).

## Format des fichiers Excel

L'application attend des fichiers Excel avec les colonnes suivantes :

- `id` : Identifiant unique (texte)
- `nom` : Nom de famille (texte)
- `prenom` : Prénom (texte)
- `departement` : Département/Service (texte)
- `jours_prestes` : Nombre de jours travaillés (nombre entier)
- `taux_journalier` : Taux journalier (nombre décimal)
- `montant_apei` : Montant APEI calculé (nombre décimal)

Des exemples de fichiers sont disponibles dans le dossier `examples`.

## Licence

Ce projet est distribué sous licence MIT. Voir le fichier `LICENSE` pour plus d'informations.

## Auteur

[Votre nom]

## Remerciements

- L'équipe de développement
- Les testeurs
- Les utilisateurs finaux pour leur feedback précieux