Attribute VB_Name = "Bibliotheque"
'
' https://github.com/VBA-Outils/Bibliotheque
'
' Fonctions génériques VBA
'
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Copyright (c) 2026, Vincent ROSSET
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without
' modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGligneCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

'-------------------------------------------------------------------------------------------------------------------------
' Bibliothèque de procédures / fonctions multi-projets
'-------------------------------------------------------------------------------------------------------------------------

' +--------------------------------+-------------------------------------------------------------------------------------+
' | Fonction / Procédure           | Description                                                                         |
' +--------------------------------+-------------------------------------------------------------------------------------+
' | DeprotegerFeuille              | Oter la protection d'une feuille.                                                   |
' | ProtegerFeuille                | Protéger une feuille.                                                               |
' | EstFeuilleExistante            | Vérifie si le nom de l'onglet existe dans le classeur.                              |
' | EstClasseurOuvert              | Vérifie si un classeur est ouvert dans Excel                                        |
' | EstListeDeroulante             | Vérifie si la cellule de la feuille est une liste déroulante.                       |
' | DerniereLigne                  | Retourne le numéro de la dernière ligne renseignée d'une colonne d'une feuille.     |
' | DerniereColonne                | Retourne le numéro de la dernière colonne renseignée d'une ligne d'une feuille.     |
' | NumeroColonne                  | Convertit les lettres d'une colonne au numéro de colonne correspondant.             |
' | LettreColonne                  | Convertit un numéro de colonne au format Lettre.                                    |
' | AjouterListeDeroulante         | Ajoute une liste déroulante dans la feuille.                                        |
' | ExtensionFichier               | Retourne l'extension d'un fichier.                                                  |
' | TriBulles                      | Trie un tableau de chaînes de caractères avec la méthode du tri à bulles.           |
' | TriRapide                      | Trie un tableau de chaînes de caractères avec la méthode du tri rapide.             |
' |                                | Cette méthode nécessite d'initialiser des sentinelles avant de trier.               |
' | InitialiserTraitement          | Procédure à exécuter au début d'un traitement afin de désactiver le rafraîchissement|
' |                                | automatique et les événements. Elle permet d'améliorer les performances en          |
' |                                | désactivant les rafraîchissements de l'écran en arrière-plan.                       |
' | TerminerTraitement             | Procédure à exécuter à la fin du traitement afin d'annuler les désactivations       |
' |                                | réalisées à l'initialisation.                                                       |
' | EstNomExistant                 | Vérifie si un nom Excel existe dans le classeur.                                    |
' | ConvertirUrlSharePoint         | Convertit les répertoires sous forme d'URL (https://live....) dans un format        |
' |                                | compatible avec le systèmes de fichiers de Windows.                                 |
' | FichierExiste                  | Vérifie si le fichier en paramètre existe physiquement.                             |
' | RepertoireExiste               | Vérifie si le répertoire en paramètre existe physiquement.                          |
' | ListeLignesSelectionnees       | Déterminer la liste des lignes sélectionnées après un numéro de ligne d'en-tête     |
' | CreerTS                        | Créer un tableau structuré qui contient une table                                   |
' +--------------------------------+-------------------------------------------------------------------------------------+

Option Explicit
Option Compare Text

'-------------------------------------------------------------------------------------------------------------------------
' Enum pour l'ordre du tri à bulles
'-------------------------------------------------------------------------------------------------------------------------
Public Enum OrderByEnum
    Ascending = 1
    Descending = 2
End Enum

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : ExtensionFichier
' Rôle      : Retourne l'extension d'un fichier
' Paramètre : NomFichier [String] - nom physique d'un fichier avec son extension (répertoire facultatif)
' Résultat  : La fonction retourne l'extension du fichier seule (sans le point)
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call ExtensionFichier("Classeur.xlsx") => retourne "xlsx"
'-------------------------------------------------------------------------------------------------------------------------
Public Function ExtensionFichier(NomFichier As String) As String

    Dim lPosPt As Long
    
    lPosPt = InStrRev(NomFichier, ".")
    If lPosPt > 0 Then
        ExtensionFichier = LCase$(Right$(NomFichier, Len(NomFichier) - lPosPt))
    End If
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Procédure : DeprotegerFeuille
' Rôle      : Oter la protection d'une feuille protégée
' Paramètre : wsFeuille [Worksheet] - Objet feuille d'un classeur Excel
' Résultat  : Si la feuille est protégée alors la protection est désactivée
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call DeprotegerFeuille(worksheets("Feuil1")) => supprime la protection de la feuille "Feuil1"
'-------------------------------------------------------------------------------------------------------------------------
Public Sub DeprotegerFeuille(wsFeuille As Worksheet, Optional Password As String = "")

    wsFeuille.Unprotect Password:=Password

End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Procédure : ProtegerFeuille
' Rôle      : Protéger une feuille d'un classeur
' Paramètre : wsFeuille [Worksheet] -> Objet feuille d'un classeur Excel
'             Password (string) -> Mot de passe de protection. Chaîne de caractères.
'             DrawingObjects (Boolean) -> Protège les objets de dessin : formes, zones de texte, graphiques, SmartArt, etc.
'             Contents (Boolean) -> Protège le contenu des cellules. C’est le paramètre le plus courant.
'             Scenarios (Boolean) -> Protège les scénarios (fonctionnalité Excel peu utilisée aujourd’hui).
'             UserInterfaceOnly (Boolean) -> Protège la feuille pour l’utilisateur, mais autorise le VBA à modifier la feuille.
'                                            Très utile pour les macros qui doivent écrire dans une feuille protégée.
'             AllowFormattingCells (Boolean)
'             AllowFormattingColumns (Boolean)
'             AllowFormattingRows (Boolean)
'             AllowInsertingColumns (Boolean)
'             AllowInsertingRows (Boolean)
'             AllowInsertingHyperlinks (Boolean)
'             AllowDeletingColumns (Boolean)
'             AllowDeletingRows (Boolean)
'             AllowSorting (Boolean)
'             AllowFiltering (Boolean)
'             AllowUsingPivotTables (Boolean)
' Résultat  : Si la feuille n'est pas protégée alors la protection est activée en protégeant l'interface utilisateur mais pas les macros
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call ProtegerFeuille(worksheets("Feuil1")) => protège la feuille "Feuil1"
'-------------------------------------------------------------------------------------------------------------------------
Public Sub ProtegerFeuille(wsFeuille As Worksheet, Optional Password As String = "", Optional DrawingObjects As Boolean = False, Optional Contents As Boolean = False, _
                           Optional Scenarios As Boolean = False, Optional UserInterfaceOnly As Boolean = False, Optional AllowFormattingCells As Boolean = False, _
                           Optional AllowFormattingColumns As Boolean = False, Optional AllowFormattingRows As Boolean = False, Optional AllowInsertingColumns As Boolean = False, _
                           Optional AllowInsertingRows As Boolean = False, Optional AllowInsertingHyperlinks As Boolean = False, Optional AllowDeletingColumns As Boolean = False, _
                           Optional AllowDeletingRows As Boolean = False, Optional AllowSorting As Boolean = False, Optional AllowFiltering As Boolean = False, _
                           Optional AllowUsingPivotTables As Boolean = False)

    wsFeuille.Protect Password:=Password, DrawingObjects:=DrawingObjects, Contents:=Contents, Scenarios:=Scenarios, UserInterfaceOnly:=UserInterfaceOnly, _
        AllowFormattingCells:=AllowFormattingCells, AllowFormattingColumns:=AllowFormattingColumns, AllowFormattingRows:=AllowFormattingRows, _
        AllowInsertingColumns:=AllowInsertingColumns, AllowInsertingRows:=AllowInsertingRows, AllowInsertingHyperlinks:=AllowInsertingHyperlinks, _
        AllowDeletingColumns:=AllowDeletingColumns, AllowDeletingRows:=AllowDeletingRows, AllowSorting:=AllowSorting, _
        AllowFiltering:=AllowFiltering, AllowUsingPivotTables:=AllowUsingPivotTables
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : NumeroColonne
' Rôle      : Convertir une lettre de colonne en son numéro équivalent.
'             Par exemple, la colonne A correspond au numéro 1, Z à 26, AA à 27, etc
' Paramètre : ColonneAlphabet [String] - Lettre(s) de la colonne (entre "A" et "XFD")
' Résultat  : La fonction retourne le numéro de la colonne qui correspond aux lettres communiquées
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call NumeroColonne("AAA") => retourne 703
'-------------------------------------------------------------------------------------------------------------------------
Public Function NumeroColonne(ColonneAlphabet As String) As Long

    Dim IcLettre As Integer, NbreLettres As Integer, Lettre As String
    
    NbreLettres = Len(ColonneAlphabet)
    ' 3 lettres maximun par colonne, et la dernière colonne présente dans Excel est "XFD"
    If NbreLettres > 3 Then Exit Function
    
    NumeroColonne = 0
    For IcLettre = 1 To NbreLettres
        Lettre = UCase$(Mid$(ColonneAlphabet, IcLettre, 1))
        If Lettre < "A" Or Lettre > "Z" Then
            Err.Raise -10, "Numéro d'une colonne", "Lettre de colonne invalide : """ & Lettre & """"
        End If
        NumeroColonne = NumeroColonne * 26 + Asc(Lettre) - 64
    Next IcLettre
    
    ' La dernière colonne est XFD, soit le numéro 16384
    If NumeroColonne > 16384 Then
        Err.Raise -11, "Numéro d'une colonne", "Référence de colonne invalide"
    End If
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : LettreColonne
' Rôle      : Convertir un numéro de colonne en lettre(s)
'             Par exemple, la colonne 1 correspond au numéro A, 26 à Z, 27 à AA, etc
' Paramètre : NumeroColonne [Long] - Numéro de la colonne (entre 1 et 16384)
' Résultat  : La fonction retourne les lettres de la colonne qui correspondent au numéro communiqué
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call LettreColonne(703) => retourne "AAA"
'-------------------------------------------------------------------------------------------------------------------------
Public Function LettreColonne(ByVal NumeroColonne As Long) As String

    Dim Numero1ereLettre As Long, Numero2emeLettre As Long, Numero3emeLettre As Long

    If NumeroColonne > 16384 Or NumeroColonne < 1 Then
        Err.Raise -20, "Lettre(s) d'une colonne", "Numéro de colonne invalide"
    End If
    
    ' Si le numéro de colonne > 702 alors 3 lettres sont nécessaires
    ' Entre chaque 1ère lettre (Axx et Bxx) il existe 26*26=676 combinaisons
    ' On calcule d'abord le nombre de colonnes - 26 premières colonnes (A à Z) module 676 afin d'obtenir le rang de la 1ère lettre (0 = 2 lettres seulement, 1 = Axx, 2 = Bxx)
    Numero1ereLettre = (NumeroColonne - 27) \ 676
    ' Calcul la valeur du numéro de colonne (des 2ème et 3ème lettre) sans la première lettre
    NumeroColonne = NumeroColonne - Numero1ereLettre * 676
    ' Calcul du résultat modulo 26 afin d'obtenir le rang de la 2ème lettre (1 = Ax, 2 = Bx)
    Numero2emeLettre = (NumeroColonne - 1) \ 26
    ' Calcul du rang de la 3ème lettre, c'est-à-dire le reste de la division par 26
    Numero3emeLettre = NumeroColonne - Numero2emeLettre * 26
    ' Concatène les 3 résultats afin d'obtenir les lettres qui correspondent au n° de colonne
    LettreColonne = IIf(Numero1ereLettre = 0, "", Chr(64 + Numero1ereLettre)) & IIf(Numero2emeLettre = 0, "", Chr(64 + Numero2emeLettre)) + Chr(64 + Numero3emeLettre)

End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : EstFeuilleExistante
' Rôle      : Vérifie si un nom de feuille existe déjà dans un classeur
' Paramètre : wbClasseur [Workbook] - Objet classeur qui contiendrait la feuille dont on veut vérifier la présence
'             NomFeuille [String]   - Nom de la feuille (onglet) dont on veut vérifier la présence dans un classeur donné
' Résultat  : La fonction retourne True si la feuille existe dans le classeur sinon False
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' EstFeuilleExistante(ThisWorkBook,"Feuil1") => retourne True si "Feuil1" est présent dans le classeur qui exécute la macro
'-------------------------------------------------------------------------------------------------------------------------
Public Function EstFeuilleExistante(wbClasseur As Workbook, NomFeuille As String) As Boolean

    Dim wsFeuille As Worksheet

    ' Pour chaque feuille présente dans le classeur
    For Each wsFeuille In wbClasseur.Worksheets
        ' Si le nom de la feuille en entrée est identique à celui d'une feuille du classeur (ne pas tenir compte de la casse)
        If StrComp(wsFeuille.Name, NomFeuille, vbTextCompare) = 0 Then
            ' La feuille existe dans le classeur, on retourne le booléen Vrai
            EstFeuilleExistante = True
            Exit Function
        End If
    Next wsFeuille
    EstFeuilleExistante = False
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : EstClasseurOuvert
' Rôle      : Vérifie si le classeur est ouvert dans Excel
' Paramètre : NomClasseur [string] - Nom du classeur à contrôler
' Résultat  : La fonction retourne True si le classeur est ouvert dans Excel sinon False
'-------------------------------------------------------------------------------------------------------------------------
Public Function EstClasseurOuvert(NomClasseur As String) As Boolean

    Dim wbClasseur As Workbook

    ' Pour chaque classeur ouvert dans Excel
    For Each wbClasseur In Workbooks
        If StrComp(NomClasseur, wbClasseur.Name, vbTextCompare) = 0 Then
            EstClasseurOuvert = True
            Exit Function
        End If
    Next wbClasseur
    EstClasseurOuvert = False
    
End Function


'-------------------------------------------------------------------------------------------------------------------------
' Procédure : InitialiserTraitement
' Rôle      : Initialiser des traitements longs en désactivant le rafraichissement automatique de l'écran et les événements, affichant un sablier
' Paramètre : N/A
' Résultat  : Désactive le rafraichissement, affiche un sablier, désactive les événements
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call InitialiserTraitement
'-------------------------------------------------------------------------------------------------------------------------
Public Sub InitialiserTraitement()

    ' Ne plus rafraichir l'écran
    Application.ScreenUpdating = False
    ' Afficher le curseur d'attente (sablier)
    Application.Cursor = xlWait
    ' Annuler le copier/couper d'Excel qui serait encore actif (cela perturbe certaines actions faites par VBA)
    Application.CutCopyMode = False
    ' Pour toute automatisation, on commence par inhiber les événements, afin de ne pas déclencher Worksheet_Change
    Application.EnableEvents = False

End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Procédure : TerminerTraitement
' Rôle      : Terminer des traitements longs en réactivant le rafraichissement automatique de l'écran et les événements, affichant le curseur se souris
' Paramètre : N/A
' Résultat  : Réactive le rafraichissement, affiche le curseur par défaut, active les événements
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call TerminerTraitement
'-------------------------------------------------------------------------------------------------------------------------
Public Sub TerminerTraitement()

    ' Rafraichier de nouveau l'écran
    Application.ScreenUpdating = True
    ' Afficher le curseur de souris par défaut
    Application.Cursor = xlDefault
    ' Réactiver les événements
    Application.EnableEvents = True

End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : EstListeDeroulante
' Rôle      : Vérifie si une cellule est une liste déroulante
' Paramètre : wsFeuille [Worksheet] - Objet feuille qui contient la cellule à inspecter
'             Cellule [Range]       - Objet Cellule dont on veut déterminer si une liste déroulante est présente
' Résultat  : La fonction retourne True si la cellule contient une liste déroulante
'-------------------------------------------------------------------------------------------------------------------------
Public Function EstListeDeroulante(wsFeuille As Worksheet, Cellule As Range, Optional MotDePasse As String = "") As Boolean

    Dim rCible As Range
    Dim ProtectionCI As New ProtectionState
 
    ' Déprotéger la feuille afin de pouvoir insérer une ligne
    ProtectionCI.LoadFromWorksheet wsFeuille
    ProtectionCI.UnprotectWorksheet wsFeuille, MotDePasse
    
    ' Recherche toutes les cellules contenant une liste de validation dans la feuille active et non protégée.
    Set rCible = wsFeuille.Cells.SpecialCells(xlCellTypeAllValidation)
    
    ' Si aucune cellule de validation trouvée dans la feuille
    If rCible Is Nothing Then
        EstListeDeroulante = False
    Else
        If Intersect(rCible, Cellule) Is Nothing Then
            EstListeDeroulante = False
        Else
            EstListeDeroulante = True
        End If
    End If
    
    ' Protéger de nouveau la feuille
    ProtectionCI.ApplyToWorksheet wsFeuille, MotDePasse
    Set ProtectionCI = Nothing
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : EstNomExistant
' Rôle      : Vérifie si un nom donné existe dans un classeur
' Paramètre : wsClasseur [Workbook] - Objet classeur qui contiendrait le nom cherché
'             Nom [String]          - Nom d'une cellule ou plage de cellules (Formules / Gestionnaire de noms)
' Résultat  : La fonction retourne True si le Nom donné existe dans le classeur donné
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' EstNomExistant(ThisWorkBook,"Test") => retourne True si le nom "Test" existe dans le classeur qui exécute la macro
'-------------------------------------------------------------------------------------------------------------------------
Public Function EstNomExistant(wbClasseur As Workbook, Nom As String) As Boolean

    Dim nNom As Name
    
    EstNomExistant = False
    ' Pour chaque nom présent dans le classeur
    For Each nNom In wbClasseur.Names
        ' Si le nom en entrée existe dans le classeur
        If StrComp(nNom.Name, Nom, vbTextCompare) = 0 Then
            EstNomExistant = True
            Exit For
        End If
    Next
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : DerniereLigne
' Rôle      : Retourne le numéro de la dernière ligne qui contient des données dans une colonne donnée
' Paramètre : wsFeuille [Worksheet] - Objet Feuille dans laquelle la recherche sera effectuée
'             NumeroColonne [Long]  - Numéro de la colonne dans laquelle rechercher la dernière donnée présente
' Résultat  : La fonction retourne le numéro de la ligne qui contient la dernière donnée renseignée dans la colonne donnée
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' DerniereLigne(ActiveSheet,1) => retourne le numéro de la dernière ligne renseignée en colonne A
'-------------------------------------------------------------------------------------------------------------------------
Public Function DerniereLigne(wsFeuille As Worksheet, NumeroColonne As Long) As Long

    Dim rCellule As Range
    
    ' Dans la colonne n de la feuille
    With wsFeuille.Columns(NumeroColonne)
        ' Rechercher la ligne précédente qui contient un texte
        Set rCellule = .Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues)
        If rCellule Is Nothing Then
            DerniereLigne = 1
        Else
            DerniereLigne = rCellule.Row
        End If
    End With
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : DerniereColonne
' Rôle      : Retourne le numéro de la dernière colonne qui contient des données dans une ligne donnée
' Paramètre : wsFeuille [Worksheet] - Objet Feuille dans laquelle la recherche sera effectuée
'             NumeroLigne [Long]    - Numéro de la ligne dans laquelle rechercher la dernière donnée présente
' Résultat  : La fonction retourne le numéro de la colonne qui contient la dernière donnée renseignée dans la ligne donnée
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' DerniereColonne(ActiveSheet,1) => retourne le numéro de la dernière colonne renseignée pour la ligne 1
'-------------------------------------------------------------------------------------------------------------------------
Public Function DerniereColonne(wsFeuille As Worksheet, NumeroLigne As Long) As Long

    Dim rCellule As Range
    
    ' Dans la ligne n de la feuille
    With wsFeuille.Rows(NumeroLigne)
        ' Rechercher la colonne précédente qui contient un texte
        Set rCellule = .Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues)
        If rCellule Is Nothing Then
            DerniereColonne = 1
        Else
            DerniereColonne = rCellule.Column
        End If
    End With
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : ConvertirUrlSharePoint
' Rôle      : Convertir un nom de chemin défini par une URL OneDrive ou SharePoint vers un nom de chemin Windows
'             Exemple : https://xxx-my.sharepoint.com/personal/ devient c:\Users\xxxx\OneDrive - xxx
' Paramètre : Chemin [String] - chemin d'accès à un répertoire ou fichier
' Résultat  : La fonction retourne le répertoire pour accéder à l'URL en entrée
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' ConvertirUrlSharePoint("https://d.docs.live.net/5938f6g833d79c7d/Documents") => retourne "c:\users\vince\OneDrive\Documents"
'-------------------------------------------------------------------------------------------------------------------------
Public Function ConvertirUrlSharePoint(Chemin As String) As String

    Dim sListeDossiers() As String, iNbDossiers As Integer, lPosDoc As Long, sRepertoire As String
    
    ' Si le chemin du fichier commence par http
    If StrComp(Left(Chemin, 4), "http", vbTextCompare) = 0 Then
        Select Case True
        ' Espace personnel sur SharePoint (i.e. OneDrive Commercial)
        Case Chemin Like "https://*-my.sharepoint.com/personal/*"
            ' Recherche la chaîne "/Documents/documents" afin d'obtenir le début de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, Chemin, "/Documents/Documents/", vbTextCompare) + Len("/Documents")
            ' Le répertoire local est récupéré à partir du 2ème /Documents
            sRepertoire = Mid(Chemin, lPosDoc, Len(Chemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDriveCommercial") & Replace(sRepertoire, "/", "\")
        ' Espace de travail partagé
        Case Chemin Like "https://weshare*"
            sListeDossiers = Split(Chemin, "/")
            ConvertirUrlSharePoint = "\\" & sListeDossiers(2) & "@SSL\DavWWWRoot"
            For iNbDossiers = 3 To UBound(sListeDossiers)
                ConvertirUrlSharePoint = ConvertirUrlSharePoint & "\" & sListeDossiers(iNbDossiers)
            Next
        Case Chemin Like "https://d.docs.live.net/*"
            ' Recherche la chaîne "/documents" afin d'obtenir le début de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, Chemin, "/Documents/", vbTextCompare)
            ' Le répertoire local est récupéré à partir du 2ème /Documents
            sRepertoire = Mid(Chemin, lPosDoc, Len(Chemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDrive") & Replace(sRepertoire, "/", "\")
        End Select
    Else
        ConvertirUrlSharePoint = Chemin
    End If
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : FichierExiste
' Rôle      : Vérifie si un fichier physique existe
' Paramètre : NomFichier [String] - Nom du fichier dont l'existence doit être vérifiée (inclure le répertoire avant le nom)
' Résultat  : La fonction retourne True si le fichier existe
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' FichierExiste("c:\Windows\Notepad.exe") => retourne True si ce fichier est présent
'-------------------------------------------------------------------------------------------------------------------------
Public Function FichierExiste(NomFichier) As Boolean
    
    FichierExiste = Dir(NomFichier, vbNormal) <> ""
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : RepertoireExiste
' Rôle      : Vérifie si un répertoire existe
' Paramètre : Repertoire [String] - répertoire dont l'existence doit être vérifiée
' Résultat  : La fonction retourne True si le répertoire existe
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' RepertoireExiste("c:\Windows\") => retourne True si ce répertoire est présent
'-------------------------------------------------------------------------------------------------------------------------
Public Function RepertoireExiste(Repertoire As String) As Boolean
    
    RepertoireExiste = Dir(Repertoire, vbDirectory) <> ""
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Procédure : AjouterListeDeroulante
' Rôle      : Créer une liste déroulante dans une cellule donnée
' Paramètre : Cellule [Range]          - Objet Cellule (unique) dans lequel la liste déroulante doit être créée
'             Formula1 [String]        - Renvoie la valeur ou l'expression associée au format conditionnel ou à la validation des données.
'                                        Il peut s’agir d’une valeur constante, d’une valeur de chaîne, d’une référence de cellule ou d’une formule. Type de données String en lecture seule.
'             InCellDropdown [Boolean] - True si la validation de données affiche une liste déroulante qui contient les valeurs autorisées.
'             IgnoreBlank [Boolean]    - Cette propriété a la valeur True si des valeurs nulles sont autorisées par la validation de données de la plage.
'             ShowError [Boolean]      - True si le message d’erreur de validation de données s’affiche lorsque l’utilisateur entre des données non valides.
' Résultat  : La procédure crée une liste déroulante constituée des éléments présents dans le nom donné
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' AjouterListeDeroulante(Range("A1"),"=Pays",True,True,True)
'-------------------------------------------------------------------------------------------------------------------------
Public Sub AjouterListeDeroulante(Cellule As Range, Formula1 As String, IgnoreBlank As Boolean, InCellDropdown As Boolean, ShowError As Boolean)

    ' Création d'une liste déroulante
    With Cellule.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Formula1
        .IgnoreBlank = IgnoreBlank
        .InCellDropdown = InCellDropdown
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = ShowError
    End With
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Fonction  : ListeLignesSelectionnees
' Rôle      : Déterminer la liste des lignes sélectionnées après un numéro de ligne d'en-tête
' Paramètre : NumeroLigneEntete [Long] - Numéro de ligne à partir duquel les lignes sont ajoutées dans la liste
' Résultat  : numéros des lignes sélectionnées triées
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' dim aListe() as long
' aListe = ListeLignesSelectionnees()
'-------------------------------------------------------------------------------------------------------------------------
Public Function ListeLignesSelectionnees(Optional NumeroLigneEntete As Long = 0) As Variant

    Dim rCell As Range, dListeLignes As New Dictionary, aListe() As Variant
    
    ' Pour chaque cellule sélectionnée dans le classeur actif
    For Each rCell In Selection.Cells
        ' Si le numéro de ligne de la cellule est supérieur à celui de l'en-tête alors on ajoute ce numéro à la liste
        If rCell.Row > NumeroLigneEntete Then
            dListeLignes(rCell.Row) = True
        End If
    Next rCell
    
    ' Convertir en tableau pour le tri
    aListe = dListeLignes.Keys
    Call TriBulles(aListe, Ascending)
    
    ListeLignesSelectionnees = aListe
    
    Set dListeLignes = Nothing
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Procédure : TriBulles
' Rôle      : Tri à bulles
' Paramètre : aTableau() [Variant]  - Tableau à trier
'             OrderBy [OrderByEnum] - Ordre du tri
' Résultat  : Tableau trié
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' dim aListe() as long
' aListe(0) = 5
' aListe(1) = 2
' Call TriBulles(aListe)
'-------------------------------------------------------------------------------------------------------------------------
Public Sub TriBulles(aTableau() As Variant, OrderBy As OrderByEnum)

    Dim IcBoucle1 As Long, IcBoucle2 As Long, IcPremOccur As Long, IcDernOccur As Long, vPermutation As Variant, bTableauTrie As Boolean

    ' Indice de la 1ère occurrence du tableau (0 ou 1) en fonction des options VBA
    IcPremOccur = LBound(aTableau)
    ' Indice de la dernière occurrence du tableau
    IcDernOccur = UBound(aTableau)
    ' 1ère boucle de la fin du tableau jusqu'à la 2ème occurrence
    For IcBoucle1 = IcDernOccur To IcPremOccur + 1 Step -1
        ' Le tableau est considéré comme trié tant qu'aucune vPermutation n'a eu lieu
        bTableauTrie = True
        ' 2ème boucle du début du tableau jusqu'à l'occurrence précédente de la 1ère boucle
        For IcBoucle2 = IcPremOccur To IcBoucle1 - 1
            ' Comparaison de 2 occurrences consécutives afin de les permuter si nécessaire
            If OrderBy = Ascending And aTableau(IcBoucle2) > aTableau(IcBoucle2 + 1) Or _
               OrderBy = Descending And aTableau(IcBoucle2) < aTableau(IcBoucle2 + 1) Then
                ' Les 2 occurrences sont permutées
                vPermutation = aTableau(IcBoucle2)
                aTableau(IcBoucle2) = aTableau(IcBoucle2 + 1)
                aTableau(IcBoucle2 + 1) = vPermutation
                ' Le tableau n'est pas trié
                bTableauTrie = False
            End If
        Next IcBoucle2
        ' Si aucune vPermutation n'a été réalisée alors le tableau est trié, on peut sortir de la boucle
        If bTableauTrie Then Exit For
    Next IcBoucle1

End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Procédure : TriRapide
' Rôle      : Tri rapide d'un tableau de chaînes de caractères par ordre croissant
'             Avant appel du tri, les sentinelles doivent être placées en début et fin de tableau.
' Paramètre : aTableau() (Varaint) - Tableau à trier
'             BorneInf [Long]      - numéro de la limite inférieure à trier. Le tri est effectué entre les 2 bornes.
'             BorneSup [Long]      - Numéro de la limite supérieure à trier
' Résultat  : Tableau trié
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Dim aListe(5) As Variant
' ' Sentinelle inférieure
' aListe(0) = -2147483648#
' aListe(1) = 8
' aListe(2) = 3
' aListe(3) = -8
' aListe(4) = 6
' ' Sentinelle supérieure
' aListe(5) = 2147483647
' Call TriRapide(aListe, 1, 4)
' Debug.Print aListe(1), aListe(2), aListe(3), aListe(4)
'-------------------------------------------------------------------------------------------------------------------------
Public Sub TriRapide(aTableau() As Variant, BorneInf As Long, BorneSup As Long)
    
    ' Indice afin de parcourir le tableau depuis le début jusqu'au pivot
    Dim IcDebTab As Long
    ' Indice afin de parcourir le tableau depuis la fin jusqu'au pivot
    Dim IcFinTab As Long
    ' Permutation des valeurs
    Dim vPermutation As Variant
    ' Valeur pivot
    Dim vValPivot As Variant
    ' Booléen de fin de recherche du pivot
    Dim bContinueTrt As Boolean
    
    If BorneSup > BorneInf Then
        vValPivot = aTableau(BorneInf)
        ' Débute la recherche à partir de l'indice suivant
        IcDebTab = BorneInf + 1
        IcFinTab = BorneSup
        bContinueTrt = True
        Do While bContinueTrt
            Do While aTableau(IcDebTab) < vValPivot
                IcDebTab = IcDebTab + 1
            Loop
            Do While aTableau(IcFinTab) >= vValPivot
                IcFinTab = IcFinTab - 1
            Loop
            If IcDebTab >= IcFinTab Then
                bContinueTrt = False
            Else
                vPermutation = aTableau(IcDebTab)
                aTableau(IcDebTab) = aTableau(IcFinTab)
                aTableau(IcFinTab) = vPermutation
            End If
        Loop
        vPermutation = aTableau(IcDebTab - 1)
        aTableau(IcDebTab - 1) = vValPivot
        aTableau(BorneInf) = vPermutation
        Call TriRapide(aTableau, BorneInf, IcDebTab - 2)
        Call TriRapide(aTableau, IcDebTab, BorneSup)
    End If
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Procédure : CreerTS
' Rôle      : Créer un tableau structuré à partir de la ligne qui contient le nom des colonnes
' Paramètre : wsFeuille -> Objet feuille où le tableau structuré doit être créé
'             lLigTs    -> première ligne du tableau structuré
'             LDernCol  -> Dernière colonne du tableau structuré
'             sNomTs    -> Nom du tableau structuré
' Résultat  : Tableau structuré créé
'-------------------------------------------------------------------------------------------------------------------------
Public Sub CreerTS(wsFeuille As Worksheet, lLigTs As Long, lDernCol As Long, sNomTS As String)

    Dim tsTable As ListObject, rCell As Range
    
    With wsFeuille
        .Activate
        .Range("A" & (lLigTs + 1)).Select
    End With
    ActiveWindow.FreezePanes = True
    
    Set rCell = wsFeuille.Range("A" & lLigTs & ":" & LettreColonne(lDernCol) & lLigTs)
    Set tsTable = wsFeuille.ListObjects.Add(SourceType:=xlSrcRange, Source:=rCell, XlListObjectHasHeaders:=xlYes)
    
    With tsTable
       .ShowTableStyleRowStripes = True       ' Lignes sur couleurs de fond alternées
       .ShowTableStyleColumnStripes = True    ' Colonnes sur couleurs de fond alternées
       .ShowTotals = False                    ' Affichage de la ligne de totaux
       .ShowAutoFilterDropDown = True         ' Affichage des boutons de filtres automatiques sur les en-têtes
       .TableStyle = "TableStyleLight9"       ' Style général (parmi la liste des styles prédéfinis fournis par Excel)
       .Name = sNomTS
    End With
    
    Set tsTable = Nothing
    
End Sub
