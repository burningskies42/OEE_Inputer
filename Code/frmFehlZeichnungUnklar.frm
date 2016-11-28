VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlZeichnungUnklar 
   Caption         =   "Zeichnung unklar"
   ClientHeight    =   2760
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFehlZeichnungUnklar.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlZeichnungUnklar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Input of unklare Zeichnungsnummer                                                 *
'*  Last Change: 16:30 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------
'   General form aestetics
'----------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
Dim frmFehlZeichnungUnklar As CFormChanger

Set frmFehlZeichnungUnklar = New CFormChanger

With frmFehlZeichnungUnklar

    .setIconFromWS
    Set .Form = Me

End With
End Sub

'----------------------------------------------------------------------------------------------------
'   Key presses
'----------------------------------------------------------------------------------------------------
Private Sub tbZeichNum_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    checkAllTb
End Sub

Private Sub tbZeichnum_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    'Allows input of nums only
    Select Case KeyCode
        Case 8 'Backspace
            If Len(tbZeichnum) > 0 Then
                tbZeichnum.value = Left(tbZeichnum.value, Len(tbZeichnum.value))
            End If
        Case 9 'Tab
        
        Case 13 'Enter
            If btnEingabe.Enabled = True Then
                btnEingabe_Click
            End If
            
        Case 48 To 57, 96 To 105 'Num keys from both numpad and over-letter-array

        Case Else
            KeyCode = 0 'prevents all other inputs
    End Select
    
End Sub

Private Sub checkAllTb()

If Len(tbZeichnum) < 5 Then
    btnEingabe.Enabled = False
    Exit Sub
End If

btnEingabe.Enabled = True
End Sub

Private Sub btnEingabe_Click()

innerString = Me.Caption & ", Zeichnungsummer: " & tbZeichnum

btnEingabe.Enabled = False
tbZeichnum = ""
Me.Hide
End Sub
