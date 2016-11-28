VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlGeplanteStillstand 
   Caption         =   "Geplante Stillstand"
   ClientHeight    =   2955
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4575
   OleObjectBlob   =   "frmFehlGeplanteStillstand.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlGeplanteStillstand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Input of Geplante Stillstand                                                      *
'*  Last Change: 13:22 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Public WithEvents optBtn As MSForms.OptionButton
Attribute optBtn.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------------------
'   General form aesthetics
'----------------------------------------------------------------------------------------------------

Private Sub UserForm_Activate()
Dim frmFehlGeplanteStillstand As CFormChanger

Set frmFehlGeplanteStillstand = New CFormChanger

With frmFehlGeplanteStillstand

    .setIconFromWS
    Set .Form = Me

End With
End Sub

'----------------------------------------------------------------------------------------------------
'   Add all buttons into pattern
'----------------------------------------------------------------------------------------------------

Private Sub optGespraech_Enter()
    Set optBtn = optGespraech
End Sub

Private Sub optPause_Enter()
    Set optBtn = optPause
End Sub

Private Sub optReinigung_Enter()
    Set optBtn = optReinigung
End Sub

Private Sub optSchicht_Enter()
    Set optBtn = optSchicht
End Sub

'----------------------------------------------------------------------------------------------------
'   Write problem descrition into the description column
'----------------------------------------------------------------------------------------------------

Private Sub btnEingabe_Click()
Dim ctrl As Control
Dim problemStr As String

For Each ctrl In frmFehlGeplanteStillstand.Controls
    If TypeOf ctrl Is MSForms.OptionButton Then
        If ctrl.value = True Then
            problemStr = ctrl.Caption

        End If
    End If
Next


innerString = Me.Caption & ", " & problemStr
frmFehlGeplanteStillstand.Hide
End Sub

Private Sub optBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        If optBtn.value = False Then
            optBtn.value = True
        Else
            btnEingabe_Click
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------------------
'Selecting one of the options enables "Eingabe" button
'----------------------------------------------------------------------------------------------------

Private Sub optBtn_Click()
    btnEingabe.Enabled = True
End Sub


