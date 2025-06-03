# 📊 Application d'Analyse Financière

Une application web développée avec **Flask** permettant aux utilisateurs de :

- Téléverser un fichier Excel brut (`.xlsx`)
- Générer automatiquement des **états financiers**, des **ratios comptables**, et des **graphiques interactifs**
- Télécharger les fichiers traités

---

## ⚙️ Fonctionnalités principales

### 📁 Téléversement et traitement
- Téléversement d’un fichier **.xlsx** via l’interface
- Nettoyage et préparation automatique des données comptables

### 📦 Fichiers générés
- `Comptes_Cleans.xlsx` : Données comptables nettoyées
- `Financial_Statements.xlsx` : Bilan et compte de résultat
- `Financial_Analysis.xlsx` : Ratios financiers (liquidité, rentabilité, solvabilité, efficacité)
- `Plan_Comptable.xlsx` : Plan comptable extrait
- `Summary.xlsx` : Résumé analytique par catégorie

### 📊 Visualisation des données
- Affichage de ratios financiers clés
- Graphiques dynamiques avec **Chart.js** :
  - Graphiques à secteurs (répartition des charges/produits)
  - Graphiques en barres (indicateurs de performance)
  - Graphiques en lignes (évolution des ratios)

### 📥 Export
- Téléchargement de tous les fichiers générés depuis l’interface

### 🌍 Interface
- Interface utilisateur entièrement en **français**

---
