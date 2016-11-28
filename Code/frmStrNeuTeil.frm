VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStrNeuTeil 
   Caption         =   "Ereigniss"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   OleObjectBlob   =   "frmStrNeuTeil.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmStrNeuTeil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'General form aesthetics
Private Sub UserForm_Activate()
Dim frmStrNeuTeil As CFormChanger
doMoveTeilAngabe = False

Set frmStrNeuTeil = New CFormChanger

    With frmStrNeuTeil
        .setIconFromWS
        Set .Form = Me
    
    End With

End Sub

Private Sub btnFertig_Click()
    If frmMove.btnEingabe.Enabled = True Then
        newPartEntry
        doMoveTeilAngabe = True
    Else
        MsgBox "Bitte geben Sie die Angaben zum Teil und versuchen wieder", vbInformation, "Teilangaben fehlen"
        doMoveTeilAngabe = False
    End If
    
    frmStrNeuTeil.Hide
    
End Sub

Private Sub btnStoerung_Click()
    Unload frmStrNeuTeil
    If Sheets("OEE").ProtectContents Then
            Sheets("OEE").Unprotect Password:="aczyM4iu"
    End If

    'input error using function[probInput].module[ProblemInput]
     Worksheets("OEE").Range("S" & currRow) = probInput(currColumn)
    doMoveTeilAngabe = True
    
    'frmStrNeuTeil.Hide

End Sub


