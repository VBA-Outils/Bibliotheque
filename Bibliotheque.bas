Attribute VB_Name = "Bibliotheque"
'
' https://github.com/VBA-Outils/Bibliotheque
'
' Fonctions g�n�riques VBA
'
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
'
' Copyright (c) 2024, Vincent ROSSET
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

' *----------------------------------------------------------------------------------------------------------*
' * Biblioth�que de proc�dures / fonctions multi-projets                                                     *
' *----------------------------------------------------------------------------------------------------------*
'
' Fonctions relatives aux onglets
'
' DeprotegerFeuille       : �ter la protection d'une feuille.
' ProtegerFeuille         : prot�ger une feuille.
' EstFeuilleExistante     : v�rifie si le nom de l'onglet existe dans le classeur. Exemple : EstFeuilleExistante(activeWorkBook,"Feuil1")
' ValidationExiste        : v�rifie si la cellule de la feuille est une liste d�roulante. Exemple : ValidationExiste(activeSheet, Range("B1")
' DerniereLigne           : retourne le num�ro de la derni�re ligne renseign�e d'une colonne d'une feuille.
' DerniereColonne         : retourne le num�ro de la derni�re colonne renseign�e d'une ligne d'une feuille.
' NumeroColonne           : convertit les lettres d'une colonne au num�ro de colonne correspondant. Exemple : NumeroColonne("A") retourne 1.
' LettreColonne           : convertit un num�ro de colonne au format Lettre. Exemple : LettreColonne(1) retourne "A".
' CreerLienHypertexte     : cr�e un lien hypertexte dans une cellule donn�e du classeur, avec un nom affich�.
' AjouterListeDeroulante  : ajoute une liste d�roulante dans la feuille.
'
' Fonctions g�n�riques
'
' ExtensionFichier        : retourne l'extension d'un fichier.
' TriBulles               : trie un tableau de cha�nes de caract�res avec la m�thode du tri � bulles.
' TriRapide               : trie un tableau de cha�nes de caract�res avec la m�thode du tri rapide. Cette m�thode n�cessite d'initialiser des sentinelles avant de trier.
' InitialiserTraitement   : proc�dure � ex�cuter au d�but d'un traitement afin de d�sactiver le rafra�chissement automatique et les �v�nements. Elle permet d'am�liorer les performances en d�sactivant les rafra�chissements de l'�cran en arri�re-plan.
' TerminerTraitement      : proc�dure � ex�cuter � la fin du traitement afin d'annuler les d�sactivations r�alis�es � l'initialisation.
' EstNomExistant          : v�rifie si un nom Excel existe dans le classeur.
' ConvertirUrlSharePoint  : convertit les r�pertoires sous forme d'URL (https://live....) dans un format compatible avec le syst�mes de fichiers de Windows.
' FichierExiste           : v�rifie si le fichier en param�tre existe physiquement.
' RepertoireExiste        : v�rifie si le r�pertoire en param�tre existe physiquement.
' EnregistrerClasseurSous : enregistre le classeur actif sous le nom s�lectionn� dans la bo�te de dialogue et avec le format pr�d�fini.

Option Explicit
Option Compare Text

' *---------------------------------------------------------------------------------------------------*
' * Retourne l'extension d'un fichier                                                                 *
' *---------------------------------------------------------------------------------------------------*
Public Function ExtensionFichier(sNomFichier As String) As String
    Dim lPosPt As Long
    
    lPosPt = InStrRev(sNomFichier, ".")
    If lPosPt > 0 Then
        ExtensionFichier = LCase$(Right$(sNomFichier, Len(sNomFichier) - lPosPt))
    End If
End Function

' *---------------------------------------------------------------------------------------------------*
' * Oter la protection d'une feuille (si elle est prot�g�e)                                           *
' *---------------------------------------------------------------------------------------------------*
Public Sub DeprotegerFeuille(wFeuille As Worksheet)
    If wFeuille.ProtectContents = True Then wFeuille.Unprotect
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Prot�ger une feuille en autorisant le reformatage des cellules                                    *
' *---------------------------------------------------------------------------------------------------*
Public Sub ProtegerFeuille(wFeuille As Worksheet)
    If wFeuille.ProtectContents = False Then wFeuille.Protect UserInterfaceOnly:=True
End Sub

' *---------------------------------------------------------------------------------------------------*
' * Convertir les letrres d'une colonne au num�ro de colonne correspondant                            *
' *---------------------------------------------------------------------------------------------------*
Public Function NumeroColonne(sColonne As String) As Long

    Dim iIndice As Integer, iNbColonnes As Integer
    
    iNbColonnes = Len(sColonne)
    ' 3 lettres maximun par colonne, et la derni�re colonne pr�sente dans Excel est "XFD"
    If iNbColonnes > 3 Then Exit Function
    If iNbColonnes = 3 And sColonne > "XFD" Then Exit Function
    
    NumeroColonne = 0
    For iIndice = 1 To iNbColonnes
        NumeroColonne = NumeroColonne * 26 + Asc(UCase$(Mid(sColonne, iIndice, 1))) - 64
    Next iIndice
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Convertir un num�ro de colonne au format Lettre                                                   *
' *---------------------------------------------------------------------------------------------------*
Public Function LettreColonne(lNumeroColonne As Long) As String

    Dim l1ereLettre As Long
    Dim l2emeLettre As Long
    Dim l3emeLettre As Long

    If lNumeroColonne > 16384 Or lNumeroColonne < 1 Then Exit Function

    ' Si le num�ro de colonne > 702 alors 3 lettres sont n�cessaires
    ' Entre chaque 1�re lettre (Axx et Bxx) il existe 26*26=676 combinaisons
    ' On calcule d'abord le nombre de colonnes - 26 premi�res colonnes (A � Z) module 676 afin d'obtenir le rang de la 1�re lettre (0 = 2 lettres seulement, 1 = Axx, 2 = Bxx)
    l1ereLettre = (lNumeroColonne - 27) \ 676
    ' Calcul la valeur du num�ro de colonne (des 2�me et 3�me lettre) sans la premi�re lettre
    lNumeroColonne = lNumeroColonne - l1ereLettre * 676
    ' Calcul du r�sultat modulo 26 afin d'obtenir le rang de la 2�me lettre (1 = Ax, 2 = Bx)
    l2emeLettre = (lNumeroColonne - 1) \ 26
    ' Calcul du rang de la 3�me lettre, c'est-�-dire le reste de la division par 26
    l3emeLettre = lNumeroColonne - l2emeLettre * 26
    ' Concat�ne les 3 r�sultats afin d'obtenir les lettres qui correspondent au n� de colonne
    LettreColonne = IIf(l1ereLettre = 0, "", Chr(64 + l1ereLettre)) & IIf(l2emeLettre = 0, "", Chr(64 + l2emeLettre)) + Chr(64 + l3emeLettre)

End Function

' *---------------------------------------------------------------------------------------------------*
' * V�rifier si un nom de feuille existe dans le Classeur actif                                       *
' *---------------------------------------------------------------------------------------------------*
Public Function EstFeuilleExistante(wbClasseur As Workbook, sNomFeuille As String) As Boolean

    Dim wsFeuille As Worksheet

    ' Pour chaque feuille pr�sente dans le classeur
    For Each wsFeuille In wbClasseur.Worksheets
        ' Si le nom de la feuille en entr�e est identique � celui d'une feuille du classeur (ne pas tenir compte de la casse)
        If UCase$(wsFeuille.Name) = UCase$(sNomFeuille) Then
            ' La feuille existe dans le classeur, on retourne le bool�en Vrai
            EstFeuilleExistante = True
            Exit Function
        End If
    Next wsFeuille
    EstFeuilleExistante = False
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Proc�dures d'initialisation et de terminaison d'un traitement                                     *
' *---------------------------------------------------------------------------------------------------*
Public Sub InitialiserTraitement()

    ' Ne plus rafraichir l'�cran
    Application.ScreenUpdating = False
    ' Afficher le curseur d'attente (sablier)
    Application.Cursor = xlWait
    ' Annuler le copier/couper d'Excel qui serait encore actif (cela perturbe certaines actions faites par VBA)
    Application.CutCopyMode = False
    ' Pour toute automatisation, on commence par inhiber les �v�nements, afin de ne pas d�clencher Worksheet_Change
    Application.EnableEvents = False

End Sub

Public Sub TerminerTraitement()

    ' Rafraichier de nouveau l'�cran
    Application.ScreenUpdating = True
    ' Affichier le curseur de souris par d�faut
    Application.Cursor = xlDefault
    ' R�activer les �v�nements
    Application.EnableEvents = True

End Sub

' *--------------------------------------------------------------------------------------------------------------------------*
' * V�rifie si la cellule est une liste d�roulante                                                                           *
' *--------------------------------------------------------------------------------------------------------------------------*
Public Function ValidationExiste(wFeuille As Worksheet, rCellule As Range) As Boolean

    Dim rCible As Range, bFeuilleProtegee As Boolean
 
    ' Sauvegarde l'�tat de protection de la feuille
    bFeuilleProtegee = wFeuille.ProtectContents
    ' D�prot�ger la feuille afin de pouvoir rechercher les cellules de validation
    If bFeuilleProtegee Then Call DeprotegerFeuille(wFeuille)
    
    ' Recherche toutes les cellules contenant une liste de validation dans la feuille active et non prot�g�e.
    Set rCible = wFeuille.Cells.SpecialCells(xlCellTypeAllValidation)
    
    ' Si aucune cellule de validation trouv�e dans la feuille
    If rCible Is Nothing Then
        ValidationExiste = False
    Else
        If Intersect(rCible, rCellule) Is Nothing Then
            ValidationExiste = False
        Else
            ValidationExiste = True
        End If
    End If
    
    ' Prot�ger de nouveau la feuille
    If bFeuilleProtegee Then Call ProtegerFeuille(wFeuille)

End Function

' *--------------------------------------------------------------------------------------------------------------------------*
' * Retourne Vrai si le nom est d�j� d�fini dans le classeur                                                                 *
' *--------------------------------------------------------------------------------------------------------------------------*
Public Function EstNomExistant(wbClasseur As Workbook, sNom As String) As Boolean

    Dim nNom As Name
    
    ' Pour chaque nom pr�sent dans le classeur
    For Each nNom In wbClasseur.Names
        ' Si le nom en entr�e existe dans le classeur
        If nNom.Name = sNom Then
            EstNomExistant = True
            Exit Function
        End If
    Next
    EstNomExistant = False
    
End Function

' *--------------------------------------------------------------------------------------------------------------------------*
' * Recherche de la derni�re ligne renseign�e d'une colonne                                                                  *
' *--------------------------------------------------------------------------------------------------------------------------*
Public Function DerniereLigne(wsFeuille As Worksheet, lNumeroColonne As Long) As Long

    Dim rCellule As Range
    
    ' Dans la colonne n de la feuille
    With wsFeuille.Columns(lNumeroColonne)
        ' Rechercher la ligne pr�c�dente qui contient un texte
        Set rCellule = .Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, LookIn:=xlValues)
        If rCellule Is Nothing Then
            DerniereLigne = 1
        Else
            DerniereLigne = rCellule.Row
        End If
    End With
    
End Function

' *--------------------------------------------------------------------------------------------------------------------------*
' * Recherche de la derni�re colonne renseign�e d'une ligne                                                                  *
' *--------------------------------------------------------------------------------------------------------------------------*
Public Function DerniereColonne(wsFeuille As Worksheet, lNumeroLigne As Long) As Long

    Dim rCellule As Range
    
    ' Dans la ligne n de la feuille
    With wsFeuille.Rows(lNumeroLigne)
        ' Rechercher la colonne pr�c�dente qui contient un texte
        Set rCellule = .Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious, LookIn:=xlValues)
        If rCellule Is Nothing Then
            DerniereColonne = 1
        Else
            DerniereColonne = rCellule.Column
        End If
    End With
    
End Function

' *--------------------------------------------------------------------------------------------------------------------------*
' * Convertir un nom de chemin d�fini par une URL OneDrive ou SharePoint vers un nom de chemin Windows                       *
' * Exemple : https://xxx-my.sharepoint.com/personal/ devient c:\Users\xxxx\OneDrive - xxx                                   *
' *--------------------------------------------------------------------------------------------------------------------------*
Public Function ConvertirUrlSharePoint(sChemin As String) As String

    Dim sListeDossiers() As String, iNbDossiers As Integer, lPosDoc As Long, sRepertoire As String
    
    ' Si le chemin du fichier commence par http
    If LCase$(Left(sChemin, 4)) = "http" Then
        Select Case True
        ' Espace personnel sur SharePoint (i.e. OneDrive Commercial) ?
        Case sChemin Like "https://*-my.sharepoint.com/personal/*":
            ' Recherche la cha�ne "/Documents/documents" afin d'obtenir le d�but de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, sChemin, "/Documents/Documents/", vbTextCompare) + Len("/Documents")
            ' Le r�pertoire local est r�cup�r� � partir du 2�me /Documents
            sRepertoire = Mid(sChemin, lPosDoc, Len(sChemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDriveCommercial") & Replace(sRepertoire, "/", "\")
        ' Espace de travail partag�
        Case sChemin Like "https://weshare*":
            sListeDossiers = Split(sChemin, "/")
            ConvertirUrlSharePoint = "\\" & sListeDossiers(2) & "@SSL\DavWWWRoot"
            For iNbDossiers = 3 To UBound(sListeDossiers)
                ConvertirUrlSharePoint = ConvertirUrlSharePoint & "\" & sListeDossiers(iNbDossiers)
            Next
        Case sChemin Like "https://d.docs.live.net/*":
            ' Recherche la cha�ne "/documents" afin d'obtenir le d�but de l'arborescence dans le dossier des documents
            lPosDoc = InStr(1, sChemin, "/Documents/", vbTextCompare)
            ' Le r�pertoire local est r�cup�r� � partir du 2�me /Documents
            sRepertoire = Mid(sChemin, lPosDoc, Len(sChemin) - lPosDoc + 1)
            ConvertirUrlSharePoint = Environ("OneDrive") & Replace(sRepertoire, "/", "\")
        End Select
    Else
        ConvertirUrlSharePoint = sChemin
    End If
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * V�rifie si un fichier existe                                                                      *
' *---------------------------------------------------------------------------------------------------*
Public Function FichierExiste(sNomFichier) As Boolean
    
    FichierExiste = Dir(sNomFichier, vbNormal) <> ""
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * V�rifie si un r�pertoire existe                                                                   *
' *---------------------------------------------------------------------------------------------------*
Public Function RepertoireExiste(sRepertoire As String) As Boolean
    
    RepertoireExiste = Dir(sRepertoire, vbDirectory) <> ""
    
End Function

' *--------------------------------------------------------------------------------------------------------------------------*
' * Ajout d'une liste d�roulante dans une cellule                                                                            *
' *--------------------------------------------------------------------------------------------------------------------------*
Public Sub AjouterListeDeroulante(rCellule As Range, sNomListe As String, bIgnorerErreur As Boolean, bListeDansCellule As Boolean, bAfficherErreur As Boolean)

    ' Cr�ation d'une liste d�roulante
    With rCellule.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & sNomListe
        .IgnoreBlank = bIgnorerErreur
        .InCellDropdown = bListeDansCellule
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = bAfficherErreur
    End With
End Sub
                                                                                          
' *---------------------------------------------------------------------------------------------------*
' *  Cr�er un lien hypertete dans une cellule                                                         *
' *---------------------------------------------------------------------------------------------------*
Public Function CreerLienHypertexte(NomClasseur As String, NomFeuille As String, AdrCellule As String, Repertoire As String, NomClasseurSource As String, NomFeuilleSource As String, AdrCelluleSource As String)

    Workbooks(NomClasseur).Sheets(NomFeuille).Activate
    Workbooks(NomClasseur).Sheets(NomFeuille).Hyperlinks.Add _
        Anchor:=Workbooks(NomClasseur).Sheets(NomFeuille).Range(AdrCellule), _
        Address:=Repertoire & "\" & NomClasseurSource, _
        SubAddress:=NomFeuilleSource & "!" & AdrCelluleSource, _
        TextetoDisplay:="Link"
    
End Function

' *---------------------------------------------------------------------------------------------------*
' * Enregistrer le classeur/feuille active sous le format et le nom                                   *
' *---------------------------------------------------------------------------------------------------*
' NomClasseur : nom du fichier � enregistrer
' FormatFichier : Format du fichier enregistr�
'   xlCSV                          6   *.csv   CSV
'   xlCSVUTF8                      62  *.csv   UTF8 CSV
'   xlCSVWindows                   23  *.csv   CSV Windows
'   xlExcel12                      50  *.xlsb  Classeur Excel binaire
'   xlHtml                         44  *.html  Format HTML
'   xlOpenXMLStrictWorkbook        61  *.xlsx  (&H3D) Fichier Open XML Strict
'   xlOpenXMLWorkbookMacroEnabled  52  *.xlsm  Classeur Open XML avec macros
'   xlTextWindows                  20  *.txt   Texte Windows
'   xlUnicodeText                  42  *.txt   Texte Unicode Aucune extension de fichier
'   xlXMLSpreadsheet               46  *.xml   Feuille de calcul XML
Public Sub EnregistrerClasseurSous(NomClasseur As String, FormatFichier As Long, NomInitialFichier As String)

    Dim NomFichier As Variant, CheminAcces As String, Filtre As String
    
    ' Lire le r�pertoire du classeur � enregistrer sous
    CheminAcces = Workbooks(NomClasseur).Path
    ' Convertir les URL SP et OneDrive au format de gestion des fichiers de l'OS
    If ConvertirUrlSharePoint(CheminAcces) Then
        MsgBox "Dossier SharePoint non g�r� : """ & CheminAcces & """", vbCritical
        Exit Sub
    End If
    
    ' Initialiser les filtres � appliquer lors de la s�lection du fichier � enregsitrer
    Select Case FormatFichier
        Case xlCSV: Filtre = "CSV, *.csv"
        Case xlCSVUTF8: Filtre = "UTF8 CSV, *.csv"
        Case xlCSVWindows: Filtre = "CSV Windows, *.csv"
        Case xlExcel12: Filtre = "Classeur Excel binaire, *.xlsb"
        Case xlHtml: Filtre = "Format HTML, *.htm; *.html"
        Case xlOpenXMLStrictWorkbook: Filtre = "Fichier Open XML Strict, *.xlsx"
        Case xlOpenXMLWorkbookMacroEnabled: Filtre = "Classeur Open XML avec macros, *.xlsm"
        Case xlTextWindows: Filtre = "Texte Windows, *.txt"
        Case xlUnicodeText: Filtre = "Texte Unicode, *.txt"
        Case xlXMLSpreadsheet: Filtre = "Feuille de calcul XML, *.xml"
        Case Else:
            MsgBox "Format de fichier non pris en charge", vbCritical, "Fin de l'enregistrement sous"
            Exit Sub
    End Select
    ' Changer le r�pertoire
    ChDir CheminAcces
    ' Appel de la fonction pour enregistrer sous
    NomFichier = Application.GetSaveAsFilename(FileFilter:=Filtre, Title:="Enregistrer le fichier sous le nom :", _
                  InitialFileName:=NomInitialFichier)
    ' Si un nom de fichier a �t� s�lectionn� ou saisi
    If NomFichier <> False Then
        ' Interception des erreurs, mais aucune action (e.g. l'utilisateur refuse d'�craser le fichier)
        On Error Resume Next
        ' Enregistrer sans mot de passe
        ActiveWorkbook.SaveAs WriteResPassword:="", Filename:=NomFichier, Password:="", FileFormat:=FormatFichier
    End If
    
End Sub
