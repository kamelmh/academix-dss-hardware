# Academix DSS — Quincaillerie El Bayadh

## Guide d'utilisation / User Guide

**Version:** 14.0  
**Date:** Juillet 2026  
**Auteur:** Mahi Kamel Abdelghani  
**Contact:** kamelmahi71@gmail.com | +213 676 77 38 92

---

## Table des matières

1. [Installation](#1-installation)
2. [Première ouverture](#2-première-ouverture)
3. [Configuration](#3-configuration)
4. [Utilisation quotidienne](#4-utilisation-quotidienne)
5. [Tableau de bord](#5-tableau-de-bord)
6. [Gestion des articles](#6-gestion-des-articles)
7. [Gestion des fournisseurs](#7-gestion-des-fournisseurs)
8. [Mouvements de stock](#8-mouvements-de-stock)
9. [Bons de réception](#9-bons-de-réception)
10. [Facturation](#10-facturation)
11. [Rapports](#11-rapports)
12. [Configuration avancée](#12-configuration-avancée)
13. [Sauvegarde et restauration](#13-sauvegarde-et-restauration)
14. [Dépannage](#14-dépannage)

---

## 1. Installation

### Prérequis

- **Excel 2010** ou supérieur (Excel 2016+ recommandé)
- **Windows 7** ou supérieur
- **Macro activée** : Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de gestion de la confidentialité > Paramètres des macro > « Activer toutes les macros »
- **Accès au projet VBA** : Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de gestion de la confidentialité > Paramètres des macro > Cocher « Faire confiance à l'accès au modèle d'objet du projet VBA »

### Étapes

1. Copiez le fichier `ERP_dss_v13.4_hardware_store.xlsm` sur votre ordinateur
2. Ouvrez le fichier avec Excel
3. Si Excel affiche un avertissement de macro, cliquez sur « Activer le contenu »
4. L'assistant de configuration s'affiche automatiquement (voir section 2)

---

## 2. Première ouverture

Lors de la première ouverture, l'assistant de configuration s'affiche automatiquement. Il comporte deux étapes :

### Étape 1 : Identité de l'établissement

| Champ | Description | Obligatoire |
|-------|-------------|-------------|
| Nom commercial | Le nom de votre magasin | Oui |
| Adresse | Adresse complète | Non |
| Téléphone | Numéro de téléphone | Non |
| NIF | Numéro d'Identification Fiscale (15 chiffres) | Non |
| NIS | Numéro d'Identification Statistique (15 chiffres) | Non |
| RC | Registre de Commerce (00/00-0000000A00) | Non |

> **Note :** Ces informations apparaîtront sur vos factures et bons de livraison.

### Étape 2 : Paramètres de gestion

| Champ | Description | Défaut |
|-------|-------------|--------|
| Jours ouvrés par an | Nombre de jours d'ouverture | 300 |
| Coût d'une commande | Coût fixe par commande (DZD) | 300 |
| Taux de possession | Taux de stockage annuel | 20% |
| Délai de livraison | Délai moyen en jours | 2 |
| Taux de TVA | Taux de TVA applicable | 19% |
| Devise | Devise utilisée | DZD |

> **Note :** Les valeurs par défaut conviennent à une quincaillerie. Vous pouvez les modifier plus tard via le formulaire de configuration.

### Choix final

- **Démarrer à vide** : Système prêt, aucun article
- **Charger données démo** : 40 articles de démonstration pour vous entraîner

---

## 3. Configuration

### Accès

Cliquez sur **Configuration** dans le tableau de bord.

### Paramètres modifiables

| Paramètre | Description | Impact |
|-----------|-------------|--------|
| Jours ouvrés | Influente le calcul EOQ | Calcul Wilson |
| Coût commande | Influente le calcul EOQ | Calcul Wilson |
| Taux possession | Influente le calcul EOQ | Calcul Wilson |
| Délai livraison | Seuil de réapprovisionnement | ROP |
| Taux TVA | Calcul des factures | Facturation |
| Devise | Affichage des montants | Interface |

### Sauvegarde

Cliquez sur **Enregistrer** pour sauvegarder les modifications. Les paramètres sont sauvegardés dans la feuille `CONFIG` du classeur.

---

## 4. Utilisation quotidienne

### Flux de travail typique

```
Matin :
1. Ouvrir le classeur
2. Vérifier le tableau de bord
3. Enregistrer les réceptions (si livraison)

Pendant la journée :
4. Enregistrer les ventes (mouvements de sortie)
5. Imprimer les bons de livraison

Fin de journée :
6. Vérifier les alertes stock
7. Sauvegarder le fichier
```

---

## 5. Tableau de bord

### Accès

Le tableau de bord s'affiche automatiquement à l'ouverture. Vous pouvez y accéder via **Tableau de bord** dans le menu principal.

### Indicateurs

| Indicateur | Description |
|------------|-------------|
| Nombre total d'articles | Articles enregistrés |
| Nombre de fournisseurs | Fournisseurs actifs |
| Stock total (DZD) | Valeur totale du stock |
| Alertes stock | Articles en dessous du seuil |

### Graphiques

- **Répartition ABC** : Distribution des articles par classe
- **Évolution des mouvements** : Mouvements entrée/sortie
- **Top 10 articles** : Articles les plus vendus

---

## 6. Gestion des articles

### Ajouter un article

1. Cliquez sur **Articles** dans le menu
2. Cliquez sur **Ajouter**
3. Remplissez les champs :
   - **Référence** : Code unique de l'article (ex: ART-001)
   - **Désignation** : Nom de l'article
   - **Catégorie** : Quincaillerie, Outillage, etc.
   - **Fournisseur** : Fournisseur principal
   - **Prix unitaire** : Prix d'achat HT
   - **Stock actuel** : Quantité en stock
   - **Stock minimum** : Seuil d'alerte
4. Cliquez sur **Enregistrer**

### Modifier un article

1. Sélectionnez l'article dans la liste
2. Cliquez sur **Modifier**
3. Modifiez les champs souhaités
4. Cliquez sur **Enregistrer**

### Classification ABC

Le système classe automatiquement les articles :

| Classe | Critère | Action |
|--------|---------|--------|
| **A** | 80% de la valeur | Surveillance étroite |
| **B** | 15% de la valeur | Surveillance moyenne |
| **C** | 5% de la valeur | Surveillance minimale |

---

## 7. Gestion des fournisseurs

### Ajouter un fournisseur

1. Cliquez sur **Fournisseurs** dans le menu
2. Cliquez sur **Ajouter**
3. Remplissez les champs :
   - **Code** : Code unique (ex: SIDERAL)
   - **Nom** : Raison sociale
   - **Adresse** : Adresse complète
   - **Téléphone** : Numéro de téléphone
   - **Email** : Adresse email
   - **NIF** : Numéro d'identification fiscale
   - **Contact** : Personne à contacter
   - **Délai livraison** : Délai moyen en jours
4. Cliquez sur **Enregistrer**

---

## 8. Mouvements de stock

### Enregistrer une entrée

1. Cliquez sur **Mouvements** dans le menu
2. Sélectionnez **Entrée**
3. Sélectionnez l'article
4. Entrez la quantité et le prix unitaire
5. Cliquez sur **Enregistrer**

### Enregistrer une sortie

1. Cliquez sur **Mouvements** dans le menu
2. Sélectionnez **Sortie**
3. Sélectionnez l'article
4. Entrez la quantité
5. Cliquez sur **Enregistrer**

### Impact sur le stock

- **Entrée** : Augmente le stock et met à jour le CMUP
- **Sortie** : Diminue le stock, prix basé sur le CMUP

---

## 9. Bons de réception

### Créer un bon de réception

1. Cliquez sur **Réception** dans le menu
2. Sélectionnez le fournisseur
3. Ajoutez les articles reçus
4. Cliquez sur **Enregistrer**
5. Cliquez sur **Imprimer** pour le bon

### Impression

Le bon de réception est imprimé au format A4 avec :
- En-tête avec les informations de l'établissement
- Liste des articles reçus
- Montant total
- Date et signature

---

## 10. Facturation

### Créer une facture

1. Cliquez sur **Facturation** dans le menu
2. Sélectionnez le client
3. Ajoutez les articles vendus
4. Le système calcule automatiquement :
   - Sous-total HT
   - TVA (19%)
   - Total TTC
5. Cliquez sur **Enregistrer**
6. Cliquez sur **Imprimer** pour la facture

### Numérotation

Les factures sont numérotées automatiquement : `FAC-AAAAMMJJ-NNN`

---

## 11. Rapports

### Types de rapports

| Rapport | Description |
|---------|-------------|
| **Classification ABC** | Répartition des articles par valeur |
| **Vieillissement stock** | Ancienneté des articles en stock |
| **Performance fournisseurs** | Analyse des livraisons |
| **Résumé stock** | État général du stock |
| **Historique mouvements** | Liste des mouvements |

### Accès

Cliquez sur **Rapports** dans le menu, puis sélectionnez le type de rapport.

---

## 12. Configuration avancée

### Formules Wilson (EOQ)

Le système calcule automatiquement la quantité optimale de commande :

```
EOQ = √(2DS / H)
```

Où :
- **D** = Demande annuelle
- **S** = Coût de passage de commande (300 DZD)
- **H** = Coût de possession (20% × prix unitaire)

### CMUP (Coût Moyen Pondéré Unitaire)

Le système utilise la méthode du coût moyen pondéré unitaire :
- Calcul chronologique ( FIFO)
- Mise à jour à chaque entrée de stock
- Conforme à l'arrêté du 26/07/2008 (SCF)

### Seuil de réapprovisionnement (ROP)

```
ROP = (D × LT) / Jours ouvrés + Stock de sécurité
```

---

## 13. Sauvegarde et restauration

### Sauvegarde automatique

Le système crée automatiquement une sauvegarde à chaque ouverture :
- Emplacement : `ERP_Backup_YYYYMMDD.xlsm`
- Nombre de sauvegardes conservées : 5

### Sauvegarde manuelle

1. Cliquez sur **Fichier** > **Enregistrer sous**
2. Choisissez l'emplacement
3. Nommez le fichier avec la date

### Restauration

1. Copiez le fichier de sauvegarde
2. Renommez-le avec le nom original
3. Ouvrez le fichier avec Excel

---

## 14. Dépannage

### Problème : Les macros ne s'exécutent pas

**Solution :**
1. Fichier > Options > Centre de gestion de la confidentialité
2. Paramètres des macro > « Activer toutes les macros »
3. Redémarrez Excel

### Problème : « Erreur d'exécution 1004 »

**Solution :**
1. Fermez toutes les instances d'Excel
2. Ouvrez le fichier à nouveau
3. Si le problème persiste, restaurez une sauvegarde

### Problème : Les données ne s'affichent pas

**Solution :**
1. Vérifiez que la feuille n'est pas protégée
2. Cliquez sur **Actualiser** dans le tableau de bord

### Problème : L'impression ne fonctionne pas

**Solution :**
1. Vérifiez que l'imprimante est configurée
2. Vérifiez les paramètres de mise en page

---

## Contact

Pour toute question ou assistance :

- **Email :** kamelmahi71@gmail.com
- **Téléphone :** +213 676 77 38 92
- **GitHub :** github.com/kamelmh/academix-dss-hardware

---

## Licence

Ce logiciel est propriétaire. Toute reproduction non autorisée est interdite.

**Copyright © 2026 Mahi Kamel Abdelghani. Tous droits réservés.**
