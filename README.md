# Bibliotheque
<h1>Fonctions et procédures VBA</h1>
<p>Le module VBA "Bibliotheque" propose un choix de fonctions et procédures utiles pour gérer vos projets. Ces fonctions permettent de réaliser des actions basiques et souvent redondantes dans les projets</p>
<h2>Fonctions relatives aux onglets</h2>
<ul>
 <li><strong>DeprotegerFeuille</strong> : ôter la protection d'une feuille.</li>
 <li><strong>ProtegerFeuille</strong> : protéger une feuille.</li>
 <li><strong>EstFeuilleExistante</strong> : vérifie si le nom de l'onglet existe dans le classeur. Exemple : EstFeuilleExistante(activeWorkBook,"Feuil1")</li>
 <li><strong>ValidationExiste</strong> : vérifie si la cellule de la feuille est une liste déroulante. Exemple : ValidationExiste(activeSheet, Range("B1")</li>
 <li><strong>DerniereLigne</strong> : retourne le numéro de la dernière ligne renseignée d'une colonne d'une feuille.</li>
 <li><strong>DerniereColonne</strong> : retourne le numéro de la dernière colonne renseignée d'une ligne d'une feuille.</li>
 <li><strong>NumeroColonne</strong> : convertit les lettres d'une colonne au numéro de colonne correspondant. Exemple : NumeroColonne("A") retourne 1.</li>
 <li><strong>LettreColonne</strong> : convertit un numéro de colonne au format Lettre. Exemple : LettreColonne(1) retourne "A".</li>
 <li><strong>CreerLienHypertexte</strong> : crée un lien hypertexte dans une cellule donnée du classeur, avec un nom affiché.</strong></li>
 <li><strong>AjouterListeDeroulante</strong> : ajoute une liste déroulante dans la feuille.</li>
</ul>
<h2>Fonctions génériques</h2>
<ul>
 <li><strong>ExtensionFichier</strong> : retourne l'extension d'un fichier.</li>
 <li><strong>TriBulles</strong> : trie un tableau de chaînes de caractères avec la méthode du tri à bulles.</li>
 <li><strong>TriRapide</strong> : trie un tableau de chaînes de caractères avec la méthode du tri rapide. Cette méthode nécessite d'initialiser des sentinelles avant de trier.</li>
 <li><strong>InitialiserTraitement</strong> : procédure à exécuter au début d'un traitement afin de désactiver le rafraîchissement automatique et les événements. Elle permet d'améliorer les performances en désactivant les rafraîchissements de l'écran en arrière-plan.</li>
 <li><strong>TerminerTraitement</strong> : procédure à exécuter à la fin du traitement afin d'annuler les désactivations réalisées à l'initialisation.</li>
 <li><strong>EstNomExistant</strong> : vérifie si un nom Excel existe dans le classeur.</li>
 <li><strong>ConvertirUrlSharePoint</strong> : convertit les répertoires sous forme d'URL (https://live....) dans un format compatible avec le systèmes de fichiers de Windows.</li> 
 <li><strong>FichierExiste</strong> : vérifie si le fichier en paramètre existe physiquement.</li>
 <li><strong>RepertoireExiste</strong> : vérifie si le répertoire en paramètre existe physiquement.</li>
 <li><strong>EnregistrerClasseurSous</strong> : enregistre le classeur actif sous le nom sélectionné dans la boîte de dialogue et avec le format prédéfini.</li>
</ul>
