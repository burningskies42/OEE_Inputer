VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlMaterialFehlt 
   Caption         =   "Material fehlt"
   ClientHeight    =   2895
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFehlMaterialFehlt.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlMaterialFehlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Input of Mehrmaschinebedienung                                                    *
'*  Last Change: 16:30 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------


Public WithEvents tbAllg As MSForms.TextBox
Attribute tbAllg.VB_VarHelpID = -1

'General form aestetics
Private Sub UserForm_Activate()
Dim frmFehlMaterialFehlt As CFormChanger

Set frmFehlMaterialFehlt = New CFormChanger

With frmFehlMaterialFehlt

    .setIconFromWS
    Set .Form = Me

End With
End Sub

'Key presses
Private Sub tbAuftragsnum_Enter()
    Set tbAllg = tbAuftragsnum
End Sub

Private Sub tbAllg_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    checkAllTb
End Sub

Private Sub tbAllg_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
'Allows input of nums only
    Select Case KeyCode
        Case 8 'Backspace
            If Len(tbAllg) > 0 Then
                tbAllg.value = Left(tbAllg.value, Len(tbAllg.value))
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
Dim cCont As Control

If Len(tbAllg) < 5 Then
    btnEingabe.Enabled = False
    Exit Sub
End If

btnEingabe.Enabled = True
End Sub

Private Sub btnEingabe_Click()
Dim ctrl As Control
Dim problemStr As String

For Each ctrl In Me.Controls
    If TypeOf ctrl Is MSForms.OptionButton Then
        If ctrl.value = True Then
            problemStr = ctrl.Caption
            ctrl.value = ""
        End If
    End If
Next

innerString = Me.Caption & ", TeilNum: " & tbAuftragsnum
tbAllg.value = ""
Me.Hide
End Sub
