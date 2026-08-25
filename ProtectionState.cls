VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProtectionState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' https://github.com/VBA-Outils/Gerer-Tables-Parametrage
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
' Module de classe : ProtectionState
' Descritpion      : Sauvegarder les options de protection d'une feuille, déprotéger une feuille, restaurer les options de
'                    protection lors de l'activation de la protection.
'-------------------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------------------------
' Propriétés :
' +--------------------------------+--------------------+----------------------------------------------------------------+
' | Nom                            | Type               | Description                                                    |
' +--------------------------------+--------------------+----------------------------------------------------------------+
' | ProtectContents                | Boolean            | Protège le contenu des cellules.                               |
' | ProtectDrawingObjects          | Boolean            | Protège les objets de dessin                                   |
' | ProtectScenarios               | Boolean            | Protège les scénarios                                          |
' | AllowFormattingCells           | Boolean            | Autorise le formattage des cellules                            |
' | AllowFormattingColumns         | Boolean            | Autorise le formattage des colonnes                            |
' | AllowFormattingRows            | Boolean            | Autorise le formattage des lignes                              |
' | AllowInsertingColumns          | Boolean            | Autorise l'insertion de colonnes                               |
' | AllowInsertingRows             | Boolean            | Autorise l'insertion de lignes                                 |
' | AllowInsertingHyperlinks       | Boolean            | Autorise l'insertion de liens hypertexte                       |
' | AllowDeletingColumns           | Boolean            | Autorise la suppression de colonnes                            |
' | AllowDeletingRows              | Boolean            | Autorise la suppression de lignes                              |
' | AllowSorting                   | Boolean            | Autorise le tri                                                |
' | AllowFiltering                 | Boolean            | Autorise de filtrer les données                                |
' | AllowUsingPivotTables          | Boolean            |                                                                |
' | UserInterfaceOnly              | Boolean            | Protège la feuille pour l’utilisateur, mais autorise le        |
' |                                |                    | VBA à modifier la feuille.                                     |
' +--------------------------------+--------------------+----------------------------------------------------------------+
'
' Méthodes :
' +--------------------------------+-------------------------------------------------------------------------------------+
' | Nom                            | Description                                                                         |
' +--------------------------------+-------------------------------------------------------------------------------------+
' | UnprotectWorksheet             | Oter la protection d'une feuille                                                    |
' | LoadFromWorksheet              | Sauvegarder les options de protection d'une feuille                                 |
' | ApplyToWorksheet               | Restaurer les options de protection d'une feuille                                   |
' +--------------------------------+-------------------------------------------------------------------------------------+


Option Explicit

' --- Propriétés classiques ---
Public ProtectContents As Boolean
Public ProtectDrawingObjects As Boolean
Public ProtectScenarios As Boolean

' --- Propriétés modernes (Protection object) ---
Public AllowFormattingCells As Boolean
Public AllowFormattingColumns As Boolean
Public AllowFormattingRows As Boolean

Public AllowInsertingColumns As Boolean
Public AllowInsertingRows As Boolean
Public AllowInsertingHyperlinks As Boolean

Public AllowDeletingColumns As Boolean
Public AllowDeletingRows As Boolean

Public AllowSorting As Boolean
Public AllowFiltering As Boolean
Public AllowUsingPivotTables As Boolean

' --- Cas particulier ---
Public UserInterfaceOnly As Boolean

'-------------------------------------------------------------------------------------------------------------------------
' Méthode   : UnprotectWorksheet
' Rôle      : Oter la protection de la feuille
' Paramètre : wsFeuille (Worksheet) -> Feuille protégée
'             Password (string) -> mot de passe de la feuille
'-------------------------------------------------------------------------------------------------------------------------
Public Sub UnprotectWorksheet(wsFeuille As Worksheet, Optional Password As String = "")

    wsFeuille.Unprotect (Password)
    
End Sub

'-------------------------------------------------------------------------------------------------------------------------
' Méthode   : LoadFromWorksheet
' Rôle      : Sauvegarde des paramètres de protection de la feuille
' Paramètre : wsFeuille (Worksheet) -> Feuille protégée
'-------------------------------------------------------------------------------------------------------------------------
Public Sub LoadFromWorksheet(wsFeuille As Worksheet)

    With wsFeuille
        ' Paramètres classiques
        Me.ProtectContents = .ProtectContents
        Me.ProtectDrawingObjects = .ProtectDrawingObjects
        Me.ProtectScenarios = .ProtectScenarios
        ' Paramètres modernes via wsFeuille.Protection
        With .Protection
            Me.AllowFormattingCells = .AllowFormattingCells
            Me.AllowFormattingColumns = .AllowFormattingColumns
            Me.AllowFormattingRows = .AllowFormattingRows
    
            Me.AllowInsertingColumns = .AllowInsertingColumns
            Me.AllowInsertingRows = .AllowInsertingRows
            Me.AllowInsertingHyperlinks = .AllowInsertingHyperlinks
    
            Me.AllowDeletingColumns = .AllowDeletingColumns
            Me.AllowDeletingRows = .AllowDeletingRows
    
            Me.AllowSorting = .AllowSorting
            Me.AllowFiltering = .AllowFiltering
            Me.AllowUsingPivotTables = .AllowUsingPivotTables
        End With
    End With

    ' UserInterfaceOnly n'est jamais lisible
    Me.UserInterfaceOnly = True

End Sub


'-------------------------------------------------------------------------------------------------------------------------
' Méthode   : ApplyToWorksheet
' Rôle      : Restaure la protection de la feuille
' Paramètre : wsFeuille (Worksheet) -> Feuille à protéger
'             Password (string) -> Mot de passe
'-------------------------------------------------------------------------------------------------------------------------
Public Sub ApplyToWorksheet(wsFeuille As Worksheet, Optional Password As String = "")

    ' Reprotéger avec les paramètres sauvegardés
    wsFeuille.Protect _
        Password:=Password, _
        DrawingObjects:=Me.ProtectDrawingObjects, _
        Contents:=Me.ProtectContents, _
        Scenarios:=Me.ProtectScenarios, _
        UserInterfaceOnly:=Me.UserInterfaceOnly, _
        AllowFormattingCells:=Me.AllowFormattingCells, _
        AllowFormattingColumns:=Me.AllowFormattingColumns, _
        AllowFormattingRows:=Me.AllowFormattingRows, _
        AllowInsertingColumns:=Me.AllowInsertingColumns, _
        AllowInsertingRows:=Me.AllowInsertingRows, _
        AllowInsertingHyperlinks:=Me.AllowInsertingHyperlinks, _
        AllowDeletingColumns:=Me.AllowDeletingColumns, _
        AllowDeletingRows:=Me.AllowDeletingRows, _
        AllowSorting:=Me.AllowSorting, _
        AllowFiltering:=Me.AllowFiltering, _
        AllowUsingPivotTables:=Me.AllowUsingPivotTables

End Sub
