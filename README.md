# 📚 Bibliothèque VBA

![Langage](https://img.shields.io/badge/langage-VBA-blue)
![Licence](https://img.shields.io/badge/Licence-MIT-green)

Module VBA regroupant des fonctions utilitaires pour faciliter le développement de projets Excel.  
Il couvre la gestion des feuilles, des validations, des fichiers, ainsi que des outils génériques comme le tri ou la gestion des traitements.

---

## 📄 Licence

Ce projet est distribué sous licence **MIT**.  
Consultez le fichier [`LICENSE`](LICENSE) pour plus de détails.

---

## 🧰 Prérequis

- Environnement : **Microsoft Visual Basic for Applications (VBA)**
- Compatible Excel (Windows)

---

# 🧩 Fonctions et procédures disponibles

Le module `Bibliotheque` propose un ensemble de fonctions récurrentes et utiles pour vos projets VBA.

---

## 📑 Fonctions relatives aux feuilles

| Fonction | Description |
|---------|-------------|
| **DeprotegerFeuille** | Ôte la protection d’une feuille. |
| **ProtegerFeuille** | Protège une feuille. |
| **EstFeuilleExistante** | Vérifie si un onglet existe dans le classeur.<br>Ex : `EstFeuilleExistante(ActiveWorkbook, "Feuil1")` |
| **ValidationExiste** | Vérifie si une cellule contient une liste déroulante.<br>Ex : `ValidationExiste(ActiveSheet, Range("B1"))` |
| **DerniereLigne** | Retourne la dernière ligne renseignée d’une colonne. |
| **DerniereColonne** | Retourne la dernière colonne renseignée d’une ligne. |
| **NumeroColonne** | Convertit une lettre de colonne en numéro.<br>Ex : `"A"` → `1` |
| **LettreColonne** | Convertit un numéro de colonne en lettre.<br>Ex : `1` → `"A"` |
| **AjouterListeDeroulante** | Ajoute une liste déroulante dans une cellule. |

---

## 🔧 Fonctions génériques

| Fonction | Description |
|---------|-------------|
| **ExtensionFichier** | Retourne l’extension d’un fichier. |
| **TriBulles** | Trie un tableau de chaînes (méthode du tri à bulles). |
| **TriRapide** | Trie un tableau de chaînes (méthode du tri rapide).<br>⚠️ Nécessite l’initialisation de sentinelles. |
| **InitialiserTraitement** | Désactive les rafraîchissements et événements pour accélérer un traitement. |
| **TerminerTraitement** | Réactive les options désactivées par `InitialiserTraitement`. |
| **EstNomExistant** | Vérifie si un nom Excel existe dans le classeur. |
| **ConvertirUrlSharePoint** | Convertit une URL SharePoint en chemin compatible Windows. |
| **FichierExiste** | Vérifie si un fichier existe physiquement. |
| **RepertoireExiste** | Vérifie si un répertoire existe physiquement. |
| **ListeLignesSelectionnees** | Déterminer la liste des lignes sélectionnées après un numéro de ligne d'en-tête. |

---

## 📦 Langage

Ce projet est intégralement écrit en **VBA (Visual Basic for Applications)**.

---

