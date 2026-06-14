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
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

'-------------------------------------------------------------------------------------------------------------------------
' Bibliothèque de procédures / fonctions multi-projets
'-------------------------------------------------------------------------------------------------------------------------
'
' Fonctions relatives aux onglets
'
' DeprotegerFeuille       : ôter la protection d'une feuille.
' ProtegerFeuille         : protéger une feuille.
' EstFeuilleExistante     : vérifie si le nom de l'onglet existe dans le classeur. Exemple : EstFeuilleExistante(activeWorkBook,"Feuil1")
' ValidationExiste        : vérifie si la cellule de la feuille est une liste déroulante. Exemple : ValidationExiste(activeSheet, Range("B1")
' DerniereLigne           : retourne le numéro de la dernière ligne renseignée d'une colonne d'une feuille.
' DerniereColonne         : retourne le numéro de la dernière colonne renseignée d'une ligne d'une feuille.
' NumeroColonne           : convertit les lettres d'une colonne au numéro de colonne correspondant. Exemple : NumeroColonne("A") retourne 1.
' LettreColonne           : convertit un numéro de colonne au format Lettre. Exemple : LettreColonne(1) retourne "A".
' CreerLienHypertexte     : crée un lien hypertexte dans une cellule donnée du classeur, avec un nom affiché.
' AjouterListeDeroulante  : ajoute une liste déroulante dans la feuille.
'
' Fonctions génériques
'
' ExtensionFichier        : retourne l'extension d'un fichier.
' TriBulles               : trie un tableau de chaînes de caractères avec la méthode du tri à bulles.
' TriRapide               : trie un tableau de chaînes de caractères avec la méthode du tri rapide. Cette méthode nécessite d'initialiser des sentinelles avant de trier.
' InitialiserTraitement   : procédure à exécuter au début d'un traitement afin de désactiver le rafraîchissement automatique et les événements. Elle permet d'améliorer les performances en désactivant les rafraîchissements de l'écran en arrière-plan.
' TerminerTraitement      : procédure à exécuter à la fin du traitement afin d'annuler les désactivations réalisées à l'initialisation.
' EstNomExistant          : vérifie si un nom Excel existe dans le classeur.
' ConvertirUrlSharePoint  : convertit les répertoires sous forme d'URL (https://live....) dans un format compatible avec le systèmes de fichiers de Windows.
' FichierExiste           : vérifie si le fichier en paramètre existe physiquement.
' RepertoireExiste        : vérifie si le répertoire en paramètre existe physiquement.
' EnregistrerClasseurSous : enregistre le classeur actif sous le nom sélectionné dans la boîte de dialogue et avec le format prédéfini.

Option Explicit
Option Compare Text

'-------------------------------------------------------------------------------------------------------------------------
' Retourne l'extension d'un fichier
'-------------------------------------------------------------------------------------------------------------------------
' NomFichier : nom physique d'un fichier avec son extension (répertoire facultatif)
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne l'extension du fichier seule (sans le point)
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
' Oter la protection d'une feuille protégée
'-------------------------------------------------------------------------------------------------------------------------
' wsFeuille : Objet feuille d'un classeur Excel
'-------------------------------------------------------------------------------------------------------------------------
' Si la feuille est protégée alors la protection est désactivée
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call DeprotegerFeuille(worksheets("Feuil1")) => supprime la protection de la feuille "Feuil1"
'-------------------------------------------------------------------------------------------------------------------------
Public Sub DeprotegerFeuille(wsFeuille As Worksheet)

    If wsFeuille.ProtectContents = True Then
        wsFeuille.Unprotect
    End If

End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Protéger une feuille d'un classeur
'-------------------------------------------------------------------------------------------------------------------------
' wsFeuille : Objet feuille d'un classeur Excel
'-------------------------------------------------------------------------------------------------------------------------
' Si la feuille n'est pas protégée alors la protection est activée en protégeant l'interface utilisateur mais pas les macros
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' Call ProtegerFeuille(worksheets("Feuil1")) => protège la feuille "Feuil1"
'-------------------------------------------------------------------------------------------------------------------------
Public Sub ProtegerFeuille(wsFeuille As Worksheet)

    If wsFeuille.ProtectContents = False Then
        wsFeuille.Protect UserInterfaceOnly:=True
    End If
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Convertir une lettre de colonne en son numéro équivalent.
' Par exemple, la colonne A correspond au numéro 1, Z à 26, AA à 27, etc
'-------------------------------------------------------------------------------------------------------------------------
' ColonneAlphabet : Lettre(s) de la colonne (entre "A" et "XFD")
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne le numéro de la colonne qui correspond aux lettres communiquées
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
' Convertir un numéro de colonne en lettre(s)
' Par exemple, la colonne 1 correspond au numéro A, 26 à Z, 27 à AA, etc
'-------------------------------------------------------------------------------------------------------------------------
' NumeroColonne : Numéro de la colonne (entre 1 et 16384)
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne les lettres de la colonne qui correspondent au numéro communiqué
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
' Vérifie si un nom de feuille existe déjà dans un classeur
'-------------------------------------------------------------------------------------------------------------------------
' wbClasseur : Objet classeur qui contiendrait la feuille dont on veut vérifier la présence
' NomFeuille : Nom de la feuille (onglet) dont on veut vérifier la présence dans un classeur donné
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne True si la feuille existe dans le classeur sinon False
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' EstFeuilleExistante(ThisWorkBook,"Feuil1") => retourne True si "Feuil1" est présent dans le classeur qui exécute la macro
'-------------------------------------------------------------------------------------------------------------------------
Public Function EstFeuilleExistante(wbClasseur As Workbook, NomFeuille As String) As Boolean

    Dim wsFeuille As Worksheet

    ' Pour chaque feuille présente dans le classeur
    For Each wsFeuille In wbClasseur.Worksheets
        ' Si le nom de la feuille en entrée est identique à celui d'une feuille du classeur (ne pas tenir compte de la casse)
        If UCase$(wsFeuille.Name) = UCase$(NomFeuille) Then
            ' La feuille existe dans le classeur, on retourne le booléen Vrai
            EstFeuilleExistante = True
            Exit Function
        End If
    Next wsFeuille
    EstFeuilleExistante = False
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Initialiser des traitements longs en désactivant le rafraichissement automatique de l'écran et les événements, affichant un sablier
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
' Terminer des traitements longs en activant le rafraichissement automatique de l'écran et les événements, affichant le curseur se souris
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
' Vérifie si une cellule est une liste déroulante
'-------------------------------------------------------------------------------------------------------------------------
' wsFeuille : Objet feuille qui contient la cellule à inspecter
' Cellule   : Objet Cellule dont on veut déterminer si une liste déroulante est présente
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne True si la cellule contient une liste déroulante
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' ValidationExiste(ActiveSheet,Range("B1")) => retourne True si la cellule B1 contient une liste déraoulante
'-------------------------------------------------------------------------------------------------------------------------
Public Function ValidationExiste(wsFeuille As Worksheet, Cellule As Range) As Boolean

    Dim rCible As Range, bFeuilleProtegee As Boolean
 
    ' Sauvegarde l'état de protection de la feuille
    bFeuilleProtegee = wsFeuille.ProtectContents
    ' Déprotéger la feuille afin de pouvoir rechercher les cellules de validation
    If bFeuilleProtegee Then Call DeprotegerFeuille(wsFeuille)
    
    ' Recherche toutes les cellules contenant une liste de validation dans la feuille active et non protégée.
    Set rCible = wsFeuille.Cells.SpecialCells(xlCellTypeAllValidation)
    
    ' Si aucune cellule de validation trouvée dans la feuille
    If rCible Is Nothing Then
        ValidationExiste = False
    Else
        If Intersect(rCible, Cellule) Is Nothing Then
            ValidationExiste = False
        Else
            ValidationExiste = True
        End If
    End If
    
    ' Protéger de nouveau la feuille
    If bFeuilleProtegee Then Call ProtegerFeuille(wsFeuille)

End Function

'-------------------------------------------------------------------------------------------------------------------------
' Vérifie si un nom donné existe dans un classeur
'-------------------------------------------------------------------------------------------------------------------------
' wsClasseur : Objet classeur qui contiendrait le nom cherché
' Nom        : Nom d'une cellule ou plage de cellules (Formules / Gestionnaire de noms)
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne True si le Nom donné existe dans le classeur donné
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
        If nNom.Name = Nom Then
            EstNomExistant = True
            Exit For
        End If
    Next
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Retourne le numéro de la dernière ligne qui contient des données dans une colonne donnée
'-------------------------------------------------------------------------------------------------------------------------
' wsFeuille     : Objet Feuille dans laquelle la recherche sera effectuée
' NumeroColonne : Numéro de la colonne dans laquelle rechercher la dernière donnée présente
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne le numéro de la ligne qui contient la dernière donnée renseignée dans la colonne donnée
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
' Retourne le numéro de la dernière colonne qui contient des données dans une ligne donnée
'-------------------------------------------------------------------------------------------------------------------------
' wsFeuille   : Objet Feuille dans laquelle la recherche sera effectuée
' NumeroLigne : Numéro de la ligne dans laquelle rechercher la dernière donnée présente
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne le numéro de la colonne qui contient la dernière donnée renseignée dans la ligne donnée
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
' Convertir un nom de chemin défini par une URL OneDrive ou SharePoint vers un nom de chemin Windows
' Exemple : https://xxx-my.sharepoint.com/personal/ devient c:\Users\xxxx\OneDrive - xxx
'-------------------------------------------------------------------------------------------------------------------------
' Chemin : chemin d'accès à un répertoire ou fichier
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne le répertoire pour accéder à l'URL en entrée
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' ConvertirUrlSharePoint("https://d.docs.live.net/5938f6g833d79c7d/Documents") => retourne "c:\users\vince\OneDrive\Documents"
'-------------------------------------------------------------------------------------------------------------------------
Public Function ConvertirUrlSharePoint(Chemin As String) As String

    Dim sListeDossiers() As String, iNbDossiers As Integer, lPosDoc As Long, Repertoire As String
    
    ' Si le chemin du fichier commence par http
    If LCase$(Left(Chemin, 4)) = "http" Then
        Select Case True
        ' Espace personnel sur SharePoint (i.e. OneDrive Commercial) ?
        Case Chemin Like "https://*-my.sharepoint.com/personal/*"
            ' Recherche la chaîne "/Documents/documents" afin d'obtenir le début de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, Chemin, "/Documents/Documents/", vbTextCompare) + Len("/Documents")
            ' Le répertoire local est récupéré à partir du 2ème /Documents
            Repertoire = Mid(Chemin, lPosDoc, Len(Chemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDriveCommercial") & Replace(Repertoire, "/", "\")
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
            Repertoire = Mid(Chemin, lPosDoc, Len(Chemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDrive") & Replace(Repertoire, "/", "\")
        End Select
    Else
        ConvertirUrlSharePoint = Chemin
    End If
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Vérifie si un fichier physique existe
'-------------------------------------------------------------------------------------------------------------------------
' NomFichier : Nom du fichier dont l'existence doit être vérifiée (inclure le répertoire avant le nom)
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne True si le fichier existe
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' FichierExiste("c:\Windows\Notepad.exe") => retourne True si ce fichier est présent
'-------------------------------------------------------------------------------------------------------------------------
Public Function FichierExiste(NomFichier) As Boolean
    
    FichierExiste = Dir(NomFichier, vbNormal) <> ""
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Vérifie si un répertoire existe
'-------------------------------------------------------------------------------------------------------------------------
' Repertoire : répertoire dont l'existence doit être vérifiée
'-------------------------------------------------------------------------------------------------------------------------
' La fonction retourne True si le répertoire existe
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' RepertoireExiste("c:\Windows\") => retourne True si ce répertoire est présent
'-------------------------------------------------------------------------------------------------------------------------
Public Function RepertoireExiste(Repertoire As String) As Boolean
    
    RepertoireExiste = Dir(Repertoire, vbDirectory) <> ""
    
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Créer une liste déroulante dans une cellule donnée
'-------------------------------------------------------------------------------------------------------------------------
' Cellule : Objet Cellule (unique) dans lequel la liste déroualnte doit être créée
'-------------------------------------------------------------------------------------------------------------------------
' La procédure crée une liste déroulante constituée des éléments présents dans le nom donné
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' AjouterListeDeroulante(Range("A1"),"Pays",True,True,True)
'-------------------------------------------------------------------------------------------------------------------------
Public Sub AjouterListeDeroulante(Cellule As Range, NomListe As String, IgnorerErreur As Boolean, ListeDansCellule As Boolean, AfficherErreur As Boolean)

    ' Création d'une liste déroulante
    With Cellule.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="" & NomListe
        .IgnoreBlank = IgnorerErreur
        .InCellDropdown = ListeDansCellule
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = AfficherErreur
    End With
End Sub
                                                                                          
'-------------------------------------------------------------------------------------------------------------------------
' Déterminer la liste des lignes sélectionnées
'-------------------------------------------------------------------------------------------------------------------------
' NumeroLigneEntete : Nnuméro de la dernière ligen de l'en-tête de la feuille. Tous les numéros de ligne inférieurs à cette valeur seront ignorés.
'-------------------------------------------------------------------------------------------------------------------------
' La procédure crée un objet ArrayList qui contient l'ensemble des lignes dont au moins une cellule a été sélectionnée
'-------------------------------------------------------------------------------------------------------------------------
' Exemple d'appel :
' ListeLignesSelectionnees(1,True)
'-------------------------------------------------------------------------------------------------------------------------
Public Function ListeLignesSelectionnees(Optional NumeroLigneEntete As Long = 0, Optional TriLignes As Boolean = False) As Object

    Dim lNumeroLigne As Long, IcAdrCell As Long
    Dim aListeAdrCell() As String, sAdresse As String, aPlage() As String, IcPlage As Long, lNumeroLigneDeb As Long, lNumeroLigneFin As Long
    
    Set ListeLignesSelectionnees = CreateObject("System.Collections.ArrayList")
    
    ' Découpe la liste des adresses sélectionnées
    aListeAdrCell = Split(Selection.Address, ",")
    
    ' Pour chaque bloc sélectionné
    For IcAdrCell = LBound(aListeAdrCell) To UBound(aListeAdrCell)
    
        sAdresse = aListeAdrCell(IcAdrCell)
        ' Plage de cellules ?
        If InStr(1, sAdresse, ":") > 0 Then
            ' Découpe la plage de cellules depuis le coin haut gauche vers le coin bas droite
            aPlage = Split(sAdresse, ":")
            lNumeroLigneDeb = Ligne(aPlage(0))
            ' Si c'est une colonne qui a été sélectionnée
            If lNumeroLigneDeb = -1 Then
                Err.Raise -30, , "Ne pas sélectionner une colonne entière"
            End If
            If UBound(aPlage) = 1 Then
                lNumeroLigneFin = Ligne(aPlage(1))
            Else
                lNumeroLigneFin = lNumeroLigneDeb
            End If
        Else
            lNumeroLigneDeb = Ligne(sAdresse)
            lNumeroLigneFin = lNumeroLigneDeb
        End If
        
        For lNumeroLigne = lNumeroLigneDeb To lNumeroLigneFin
            If Not ListeLignesSelectionnees.Contains(lNumeroLigne) Then
                If Range("A" & lNumeroLigne).EntireRow.Hidden = False And lNumeroLigne > NumeroLigneEntete Then
                    ListeLignesSelectionnees.Add lNumeroLigne
                End If
            End If
        Next lNumeroLigne
        
    Next IcAdrCell
    
    If TriLignes = True Then
        ListeLignesSelectionnees.Sort
    End If
    
End Function

' Extraire le numéro de ligne d'une adresse de cellule
Private Function Ligne(sAdresseCell As String) As Long

    Dim aAdresse() As String
    
    aAdresse = Split(sAdresseCell, "$")
    
    If IsNumeric(aAdresse(1)) Then
        Ligne = CLng(aAdresse(2))
    Else
        If UBound(aAdresse) = 2 Then
            Ligne = CLng(aAdresse(2))
        Else
            ' Absence de ligne (colonne sélectionnée)
            Ligne = -1
        End If
    End If
    
End Function
