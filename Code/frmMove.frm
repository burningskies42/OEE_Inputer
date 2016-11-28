VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMove 
   ClientHeight    =   4800
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   2400
   OleObjectBlob   =   "frmMove.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'Form Purpose: Display of movement operators + input of produced parts
'Last Change: 15:42 06.06.2016

'Group all textboxes
Public WithEvents SpecTextBox As MSForms.TextBox
Attribute SpecTextBox.VB_VarHelpID = -1

'---------------------------------------------------------------------------------------------------
'   General userform settings
'---------------------------------------------------------------------------------------------------

Private Sub UserForm_Activate()
Dim frmMove As CFormChanger

    Set frmMove = New CFormChanger
    frmMove.setIconFromWS

    Set frmMove.Form = Me
    frmMove.Modal = False
    
    With Me
        .Left = Application.Width - .Width - 20
        .Top = 10
    End With

End Sub

'---------------------------------------------------------------------------------------------------
'   If Teilnummer is too shirt (less than 5 chars) show the warning caption
'---------------------------------------------------------------------------------------------------

Private Sub tbTlNmmr_Change()
    If Len(tbTlNmmr) < 5 Then
        Me.Caption = "Teilangeben fehlen !"
    Else
        Me.Caption = ""
    End If
    
End Sub

Private Sub Allg_QualCB_Click()
    tbTlNmmr.SetFocus
End Sub

'---------------------------------------------------------------------------------------------------
'   Add the textboxes to the grouping
'---------------------------------------------------------------------------------------------------

Private Sub tbTlNmmr_Enter()
    Set SpecTextBox = tbTlNmmr
End Sub
Private Sub tbGutteil_Enter()
    Set SpecTextBox = tbGutteil
End Sub
Private Sub tbAussch_Enter()
    Set SpecTextBox = tbAussch
End Sub
Private Sub tbStckZeit_Enter()
    Set SpecTextBox = tbStckZeit
End Sub

'---------------------------------------------------------------------------------------------------
'   Keyboard input handling
'---------------------------------------------------------------------------------------------------

Private Sub SpecTextBox_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'MsgBox KeyCode
    Select Case KeyCode
    
        'Backspace
        Case 8
            If Len(SpecTextBox) > 0 Then
                'MsgBox Len(SpecTextBox.value)
                SpecTextBox.value = Left(SpecTextBox.value, Len(SpecTextBox.value))
            End If
            
        'Tab
        Case 9
            KeyCode = vbNull
            
            If SpecTextBox.TabIndex < 12 Then
                Controls(SpecTextBox.TabIndex).SetFocus
            Else
                Controls(8).SetFocus
            End If
        
        'Enter
        Case 13
            If btnEingabe.Enabled = True Then
                btnEingabe_Click
            End If
            
        'Cursor Keys
        Case 37 To 40
            moveME (KeyCode)
            KeyCode = vbNull
            
        'Num keys from both numpad and over-letter-array
        Case 48 To 57, 96 To 105
            If Len(SpecTextBox) = 15 Then
                KeyCode = 0
            End If
            
        'Allows for non-integer values in "time for part"
        'also convert decimal points into commas
        Case 110, 188, 190
            If SpecTextBox.Name = tbStckZeit.Name And Len(SpecTextBox) > 0 And _
                Not (SpecTextBox Like "*,") And Not (SpecTextBox Like "*,*") Then
                    KeyCode = 110
            Else
                KeyCode = 0
            End If

        'prevents all other inputs
        Case Else
            KeyCode = 0
            
    End Select

End Sub

'---------------------------------------------------------------------------------------------------
'   Checks conditions to enable the button
'---------------------------------------------------------------------------------------------------

Private Sub SpecTextBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'MsgBox Len(tbTlNmmr) * (Len(tbGutteil) + Len(tbAussch)) * Len(tbStckZeit)
    If (Len(tbGutteil) + Len(tbAussch)) * Len(tbStckZeit) > 0 Then
        If Len(tbTlNmmr) > 4 Then
            frmMove.Height = 260
            btnEingabe.Enabled = True
        Else
            btnEingabe.Enabled = False
            frmMove.Height = 275
        End If
    Else
        frmMove.Height = 260
        btnEingabe.Enabled = False
    End If
    
End Sub

'---------------------------------------------------------------------------------------------------
'   Clears all textboxes
'---------------------------------------------------------------------------------------------------

Public Sub Clr_Btn_Click()
Dim ctrl As Control
    
For Each ctrl In Me.Controls
    If ctrl.Name Like "tb*" Then
        ctrl.Text = ""
    End If
Next

End Sub

'---------------------------------------------------------------------------------------------------
' Movement keys
'---------------------------------------------------------------------------------------------------
Private Sub Btn_Up_Click()
    moveSelection (1), (Allg_QualCB)
End Sub

Private Sub Btn_Down_Click()
    moveSelection (2), (Allg_QualCB)
End Sub

Private Sub Btn_Left_Click()
    moveSelection (3), (Allg_QualCB)
End Sub

Private Sub Btn_Right_Click()
    moveSelection (4), (Allg_QualCB)
End Sub

'---------------------------------------------------------------------------------------------------
'   Movement algorithem - redirects into relevant key presses
'---------------------------------------------------------------------------------------------------
Private Sub moveME(inpCode As Integer)
    Select Case inpCode
         
        Case 38 'UP
            Btn_Up_Click
          
        Case 40 'DOWN
            Btn_Down_Click
            
        Case 37 'LEFT
            Btn_Left_Click
                    
        Case 39 'RIGHT
            Btn_Right_Click
            
    End Select

End Sub

'---------------------------------------------------------------------------------------------------
'   Input data from textboxes - parts are done
'---------------------------------------------------------------------------------------------------

Public Sub btnEingabe_Click()
    newPartEntry
End Sub

'---------------------------------------------------------------------------------------------------
'   Closing of frmMove mid-process: shows promt and saves into .sav file
'---------------------------------------------------------------------------------------------------

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim a As Integer
    
    If CloseMode = vbFormControlMenu Then
        Select Case MsgBox("Wollen Sie die Eingabe innehalten ?", vbYesNo + vbExclamation, "Terminieren")
            Case vbYes

            Case vbNo
                Cancel = True
                Exit Sub
        End Select
    End If
    
    saveForm
    
    toggleFullscreen (False)
    Worksheets("OEE").Protect Password:="aczyM4iu"
    
End Sub
