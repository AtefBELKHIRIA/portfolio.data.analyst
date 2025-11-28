<!-- Force Jekyll to render correctly -->
<!-- Responsive-friendly README -->
# Projet : Excel & VBA Application (Gestion de caisse d'un supermarché)

Application Excel VBA pour gérer les recettes de caisse, détecter les écarts, suivre les titres restaurant, automatiser les rapports et réaliser le rapprochement bancaire. Formulaires, macros, calculs automatiques et tableau de bord intégré.

---

## Objectifs du projet

- Centraliser les données de caisse en un seul endroit.
- Automatiser les tâches manuelles (calculs, rapports, alertes).
- Suivre en détail les titres restaurant par émetteur.
- Détecter automatiquement les écarts de caisse.
- Structurer la traçabilité des justificatifs.
- Faciliter le rapprochement bancaire mensuel.
- Améliorer la fiabilité et la transparence du processus de caisse.

---

## Fonctionnalités principales

---

### 1. Dashboard interactif

- Accueil central de l’application.
- Raccourcis vers les formulaires de Saisie, Recherche, Titres Restaurant, Écarts Caisse.
- Indicateurs clés : TR reçus/versés, écarts, chèques, virements.
- Vue synthétique des performances.

<div align="center">
  <img src="images/Dashboard.png" style="max-width:100%; height:auto;" width="750">
</div>

---

### 2. Base de données des recettes (Database)

- Enregistrement structuré et horodaté des recettes journalières.
- Montants par mode de paiement : espèces, CB, titres restaurant, etc.
- Suivi des écarts et commentaires de caisse.
- Fonctionne comme une base SQL simplifiée.

<div align="center">
  <img src="images/DataBase.png" style="max-width:100%; height:auto;" width="750">
</div>

---

### 3. Titres Restaurant – par émetteur

- Suivi détaillé par prestataire : BIMPLI, Edenred, UP Déjeuner, Pluxee.
- Montants collectés, versements, totaux par période.
- Indicateurs pour prévenir les écarts TR.

<div align="center">
  <img src="images/TitreRestaurant.png" style="max-width:100%; height:auto;" width="750">
</div>

---

### 4. Reçus Caisse

- Gestion des justificatifs : monnaie, virements, chèques, factures.
- Suivi des entrées et sorties non liées aux encaissements.
- Traçabilité complète.

<div align="center">
  <img src="images/ReçusCaisse.png" style="max-width:100%; height:auto;" width="750">
</div>

---

### 5. Écarts de caisse

- Analyse automatique des écarts entre encaissé et versé.
- Alertes en cas d’incohérences.
- Identification des anomalies : jour, terminal, opérateur.

<div align="center">
  <img src="images/EcartsCaisse.png" style="max-width:100%; height:auto;" width="750">
</div>

---

### 6. Rapports de caisse

- Synthèse automatisée :
  - Écarts non justifiés  
  - Justificatifs manquants  
- Export comptable.
- Aide aux audits internes et externes.

<div align="center">
  <img src="images/RapportsCaisse.png" style="max-width:100%; height:auto;" width="750">
</div>

---

### 7. E-mails automatisés

- Envoi automatique aux responsables.
- Alertes sur écarts importants et justificatifs manquants.
- Amélioration du suivi opérationnel.

<div align="center">
  <img src="images/Email.png" style="max-width:100%; height:auto;" width="750">
</div>

---

### 8. Comparaison & rapprochement bancaire

- Comparaison Recettes Caisse vs Comptabilité/Banque.
- Détection des écarts mensuels.
- Fiabilisation comptable.

<div align="center">
  <img src="images/Rapprochement.png" style="max-width:100%; height:auto;" width="750">
</div>

---

## Stack Technique

**Technologies utilisées :**
- Excel (.xlsm)
- VBA – Visual Basic for Applications

**Compétences démontrées :**
- UserForms personnalisés  
- Gestion des événements  
- Modules VBA modulaires  
- Traitement dynamique des données  
- Vérification d’intégrité  
- Automatisation des exports & rapports  
- Interface utilisateur Excel  
- Architecture multi-feuilles  

---

## Architecture du projet

<div align="center">
  <img src="images/Architecture.png" style="max-width:100%; height:auto;" width="750">
</div>

Chaque feuille assure un rôle métier dédié, et le VBA garantit l’automatisation et la cohérence globale.

---

## Ce que ce projet démontre

✔ Automatisation avancée dans Excel  
✔ Développement VBA structuré  
✔ Mini système d'information interne  
✔ Fiabilisation d’un processus métier  
✔ Compréhension du flux financier d’un magasin  
✔ Automatisation d’un processus manuel  

---

> _“Les données racontent toujours une histoire — il suffit de savoir comment les écouter.”_
