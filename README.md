**Présentation**

Ce script Python utilise pyautocad, win32com, et pythoncom pour automatiser la création des plans sur AutoCAD.
Il permet notamment de :

Importer des fichiers DWG,

Récupérer des points de bornes à partir d’un plan,

Générer un tableau AutoCAD avec les coordonnées des points récupérés.

**Installation**

1- Prérequis
- Windows

- AutoCAD installé et activé

Python 3.8 ou version ultérieure

- Dépendances Python
pyautocad, pywin32, PyQt5, pythoncom

**Utilisation**

Ouvrir AutoCAD manuellement.

Lancer un dessin vierge (ou celui que tu veux utiliser comme destination).

Exécuter le script Python :

Il ouvrira un fichier DWG source,

Récupérera le plan, la legende,

Les insérera dans les places convenables, 

Calculera les coordonnées des bornes,

**Remarques**

Assure-toi que les unités sont cohérentes entre le fichier source et le fichier courant dans AutoCAD.

Si une erreur survient (<unknown>.Activate ou autre), vérifie que tu travailles avec les bons documents actifs.
