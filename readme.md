# Extraction Automatique pour l'Envoi des Mails SCI 

## Présentation

Ce projet automatise le processus d'extraction et d'envoi de mails pour les jobs SCI. Il s'intègre parfaitement avec l'outil IBM Cognos et facilite les tâches répétitives tout en réduisant les erreurs humaines.

Il permet de planifier automatique le lancement des jobs qui vont extraire des rapports automatiquement selon le format demandé (excel, excel data, csv, hmtl ect..)

Celui ci enverra un mail sur notre robot qui va trigger (déclencher) un fluw power automate qui va concatener les données chaque jour du WMS afin de créer une base de données.

## Fonctionnalités Principales

- **Automatisation complète** de l’extraction de données via les rapports SCI [IBM Cognos].
- **Authentification sécurisée** et gestion automatisée de la connexion.
- **Interface intuitive** grâce à l'intégration de Selenium avec Edge WebDriver.
- **Enregistrement des opérations** dans une base de données SQL Server pour le suivi des extractions.

## Historique des Versions

| Version | Description                                                              | Date de Mise à jour |
|---------|--------------------------------------------------------------------------|---------------------|
| 0.0     | Développement initial de l'interface utilisateur.                        | avril 2024                 |
| 1.0     | Adaptation et optimisation du script suite à la mise à jour IBM Cognos.  | février 2025                 |

## Technologies Utilisées

- **Python**
- **Selenium WebDriver (Edge)**
- **WebDriver Manager**
- **PyODBC**
- **ConfigParser**
- **BeautifulSoup** 
- **Pandas**

## Structure du Projet

```bash
├── config
│   └── Identifiant.ini
├── logs
│   └── extraction.log
└── src
    └── extraction_mail_sci.py

```

# Installation et Configuration

## 1. Installation des dépendances

```bash
pip install selenium pyodbc webdriver-manager configparser beautifulsoup4 pandas

```

# 2. Configuration des paramètres

## Modifier le fichier Identifiant.ini avec vos informations :

```bash

[Manhattan]
username = _username
password = _password
link = lien_de_connexion_cognos
server = serveur_sql
database = base_de_données
data_table = table_sql
balise_custom_content = balise_custom_content
balise_new_orga_v2 = balise_orga_
balise_site = site_frxx
balise_frxxl = votre_balise_frxxl
balise_job = votre_balise_job

[Transport_f]
balise_type = _balise_type
balise_name_job = balise_name_job
data_name = data_name

```
# Lancer le script

## Exécuter la commande suivante en passant les arguments requis :

```bash
python extraction_mail_sci.py argument1 "report"

```

- argument1 : premier argument 
- report_SCI : paramètre spécifique 

# Suivi et Logs

Chaque extraction est enregistrée dans la base de données SQL Server afin d'assurer un suivi précis. Les erreurs sont capturées et affichées clairement pour faciliter la résolution.

# Auteur

Développé par votre équipe informatique dédiée à l'automatisation et l'optimisation des processus.
