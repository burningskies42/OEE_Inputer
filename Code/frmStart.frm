VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStart 
   Caption         =   "Auftrag"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4425
   OleObjectBlob   =   "frmStart.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Form Purpose: Start new entry or continue old one
'Last Change: 13:59 09.08.2016

Public WithEvents btnGrp As MSForms.CommandButton
Attribute btnGrp.VB_VarHelpID = -1

Private Sub btnNeu_Enter()
    Set btnGrp = btnNeu
End Sub

Private Sub btnWeiter_Enter()
    Set btnGrp = btnWeiter
End Sub

Private Sub btnGrp_Click()
    Select Case btnGrp.Name
        Case "btnNeu"
            startEntry
        Case "btnWeiter"
            startEntry (False)
    End Select
    'saveForm
End Sub

'General form aesthetics
Private Sub UserForm_Activate()
Dim frmStart As CFormChanger

Set frmStart = New CFormChanger

    With frmStart
        .setIconFromWS
        Set .Form = Me
    
    End With

End Sub
