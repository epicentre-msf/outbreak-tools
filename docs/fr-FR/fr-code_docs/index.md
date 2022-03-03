# Préambule et description du répertoire

Cette Section contient une description des codes et de la structure
du répertoire github associé au développement du projet.

Le répertoire est divisé en quatre principales branches:

- Une branche main, qui contient la dernière version stable du designer
- Une branche dev, branche de développement
- Trois branches users/[`nom-personne`], où `nom-personne` désigne le code
du nom d'un des utilisateurs qui ne développe pas à plein temps sur le projet.

A priori, le développeur principal travaillera sur la branche dev.


La structure du répertoire github est la suivante:

```
docs/
input/
    |-- geobase
    |-- setup
output/
src/
    |-- class
    |-- Form
    |-- Module

xvba_modules/
xvba_unit_tests/
.gitignore
config.json
LICENSE
linelist_designer.xlsb
package.json
README.md
```

- Le dossier `input` contient les inputs utilisés pour la génération de la linelist: il s'agit du fichier setup, et de
la géobase. 
 
- Le dossier `src` contient les codes sources utilisés pour la construction du designer. Les codes sont importés/exportés depuis/vers
le fichier excel linelist_designer.xlsb en utilisant l'extension xvba de vscode studio. Il sont divisés en trois parties: les classes, les modules et les formulaires.

- `linelist_designer.xlsb` est le designer
- Le dossier `output` contient la linelist testée, générée à partir du designer

Les autres fichiers sont des configurations pour xvba ou de la documentation, en général il n'auront pas
besoin d'être changés régulièrement.
