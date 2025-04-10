# Convertisseur FEC en PDF - Multi Comptes + Scroll + Tableau unique
## Description
Ce programme Python permet de convertir un fichier FEC (Fichier des Écritures Comptables) en PDF, en générant un tableau unique pour chaque compte sélectionné. L'interface graphique est construite avec Tkinter et permet de sélectionner un fichier FEC, de choisir un ou plusieurs comptes, puis d'exporter ces comptes au format PDF. Le PDF généré comprend un tableau structuré avec les écritures comptables, incluant les informations sur les crédits, débits, et les soldes.

## Fonctionnalités
Sélection de fichier FEC : Le programme permet à l'utilisateur de sélectionner un fichier FEC (au format texte .txt).

Chargement des comptes : Une fois le fichier sélectionné, les comptes présents dans le fichier sont extraits et affichés dans une liste.

Sélection de comptes à exporter : L'utilisateur peut sélectionner un ou plusieurs comptes à exporter dans le PDF.

Exportation en PDF : Lorsque l'utilisateur a sélectionné les comptes, il peut cliquer sur un bouton pour générer un PDF. Le fichier PDF contient un tableau détaillant les écritures de chaque compte sélectionné.

Traitement des écritures : Les écritures sont filtrées par compte, et les totaux sont calculés (crédit, débit et solde).

## Détails techniques
1. Lecture du fichier FEC :
Le fichier FEC est lu avec la fonction detect_separator() qui détecte automatiquement le séparateur (soit un pipe (|) ou un tabulation (\t)), puis il est chargé avec pandas dans un DataFrame.

2. Interface graphique (Tkinter) :
Une entrée de texte permet à l'utilisateur de sélectionner un fichier FEC.

Une liste déroulante (avec une barre de défilement) permet de choisir un ou plusieurs comptes parmi ceux présents dans le fichier.

Un bouton "Exporter en PDF" est activé lorsque l'utilisateur sélectionne au moins un compte et un fichier.

Les erreurs et messages de succès sont affichés via des boîtes de message (messagebox).

3. Exportation PDF :
Le programme utilise la bibliothèque reportlab pour générer le PDF. Voici les étapes de l'exportation :

Un tableau est créé avec les colonnes suivantes :

CompteNum : Numéro de compte.

CompteLib : Libellé du compte.

EcritureDate : Date de l'écriture.

Credit : Montant crédité.

Debit : Montant débité.

Solde : Solde de l'écriture (débit - crédit).

Les lignes sont ajoutées pour chaque écriture, avec des styles personnalisés pour les titres de comptes et les totaux.

Le tableau est formaté et stylisé pour améliorer la lisibilité dans le PDF.

4. Formatage du PDF :
En-têtes : Les en-têtes du tableau sont stylisés avec une police en gras et un fond gris.

Lignes de compte : Les comptes sont indiqués avec une ligne spéciale, et les totaux sont mis en valeur avec un fond bleu clair.

Gestion des textes longs : Si un libellé de compte est trop long, il est converti en paragraphe pour un affichage correct dans le tableau.

5. Calculs des totaux :
Pour chaque compte, le programme calcule le total des crédits, débits, et solde, et les affiche dans le tableau sous forme de ligne de total.

## Prérequis
Avant d'exécuter ce programme, assurez-vous d'avoir les éléments suivants installés sur votre machine :

Python 3.x

Tkinter (généralement installé par défaut avec Python)

pandas (pour la manipulation de données)

reportlab (pour la génération de PDF)

## Utilisation
Lancer l'application : Exécutez le script Python pour ouvrir l'interface graphique.

Sélectionner un fichier FEC : Cliquez sur le bouton "Sélectionner un fichier FEC" pour choisir un fichier .txt contenant les écritures comptables.

Sélectionner les comptes : Sélectionnez un ou plusieurs comptes à partir de la liste affichée.

Exporter en PDF : Cliquez sur le bouton "Exporter en PDF" pour générer le fichier PDF contenant les écritures comptables des comptes sélectionnés.

Enregistrer le PDF : Choisissez l'emplacement et le nom du fichier PDF à enregistrer.

## Limitations
Le fichier FEC doit être formaté de manière correcte (séparateur | ou \t).

Le format de date doit être strictement respecté (au format YYYYMMDD).

Le programme n'accepte que les fichiers .txt pour l'instant.

Le calcul du solde est effectué en soustrayant les crédits des débits.
