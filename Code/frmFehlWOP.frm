VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlWOP 
   Caption         =   "WOP"
   ClientHeight    =   2760
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFehlWOP.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlWOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Label2_Click()

End Sub

'Form Purpose: Report problems of type WOP
'Last Change: 08:23 27.06.2016

'General aestetics
Private Sub UserForm_Activate()
Dim frmFehlWOP As CFormChanger

Set frmFehlWOP = New CFormChanger

frmFehlWOP.setIconFromWS

Set frmFehlWOP.Form = Me


End Sub

'Enabling btnEingabe only when "TBzeichNum" not empty

'Key presses
Private Sub tbZeichNum_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If Len(tbZeichnum) < 5 Then
        btnEingabe.Enabled = False
    Else
        btnEingabe.Enabled = True
    End If

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

Private Sub btnEingabe_Click()
    innerString = Me.Caption & ", Zeich. Num: " & tbZeichnum
    
    btnEingabe.Enabled = False
    tbZeichnum = ""
    frmFehlWOP.Hide
End Sub

