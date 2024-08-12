# Convertisseur CSV vers Excel

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-5391FE?logo=powershell&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow)

**Convertisseur batch PowerShell** qui transforme tous les fichiers CSV d'un répertoire d'entrée en fichiers Excel (.xlsx) formatés. Les fichiers CSV traités sont automatiquement renommés avec un horodatage pour éviter le retraitement.

## Fonctionnalités

- **Traitement par lot** — Convertit tous les fichiers `.csv` du répertoire d'entrée
- **Dimensionnement auto** — Colonnes Excel automatiquement ajustées pour la lisibilité
- **Journalisation** — Logs horodatés pour chaque exécution
- **Archivage** — Fichiers CSV traités renommés avec un suffixe horodaté
- **Délimiteur point-virgule** — Supporte les fichiers CSV au format européen (`;`)

## Stack technique

| Composant | Technologie |
|-----------|------------|
| Langage | PowerShell 5.1+ |
| Module Excel | ImportExcel |

## Installation

```bash
git clone https://github.com/jmanu1983/csv-to-excel-converter.git
cd csv-to-excel-converter
```

Installer le module PowerShell requis :

```powershell
Install-Module -Name ImportExcel -Force
```

## Utilisation

1. Placer vos fichiers CSV dans le répertoire `input/`
2. Exécuter le script :

```powershell
.\convertCSVtoExcel.ps1
```

3. Récupérer les fichiers Excel générés dans `output/`

## Structure du projet

```
csv-to-excel-converter/
├── convertCSVtoExcel.ps1   # Script de conversion principal
├── input/                  # Déposer les fichiers CSV ici
├── output/                 # Fichiers Excel générés
├── logs/                   # Logs de conversion
└── README.md
```

## Licence

Ce projet est sous licence MIT.
