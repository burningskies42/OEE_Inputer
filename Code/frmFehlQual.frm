VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlQual 
   Caption         =   "Qualitätsproblem"
   ClientHeight    =   4770
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFehlQual.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlQual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Input of Qualitätsprobleme                                                        *
'*  Last Change: 13:22 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Public WithEvents optBtn As MSForms.OptionButton
Attribute optBtn.VB_VarHelpID = -1
Public WithEvents tbAussch As MSForms.TextBox
Attribute tbAussch.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------------------
'   General form aesthetics
'----------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
Dim frmFehlQual As CFormChanger

    Set frmFehlQual = New CFormChanger
    
    With frmFehlQual
        .setIconFromWS
        Set .Form = Me
    
    End With
    
    If Len(frmMove.tbAussch) > 0 Then
        optAusschuss.value = True
        tbAusschN = frmMove.tbTlNmmr
        tbAusschQ = frmMove.tbAussch
        tbAusschT = frmMove.tbStckZeit
        btnEingabe.Enabled = True
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------
'   define option button template
'----------------------------------------------------------------------------------------------------
Private Sub optAusschuss_Enter()
    Set optBtn = optAusschuss
End Sub
Private Sub optNachgearb_Enter()
    Set optBtn = optNachgearb
End Sub
Private Sub optSonst_Enter()
    Set optBtn = optSonst
End Sub

'----------------------------------------------------------------------------------------------------
'   define textbox template
'----------------------------------------------------------------------------------------------------
Private Sub tbAusschQ_Enter()
    Set tbAussch = tbAusschQ
End Sub
Private Sub tbAusschT_Enter()
    Set tbAussch = tbAusschT
End Sub
Private Sub tbAusschN_Enter()
    Set tbAussch = tbAusschN
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


Private Sub optBtn_Click()
    If optBtn = optSonst Then
        tbBeschreib.Enabled = True
        tbBeschreib = ""
    Else
        tbBeschreib.Enabled = False
        tbBeschreib = "Beschreiben Sie bitte das Problem ... "
    End If
    
    tbAusschN.Text = ""
    tbAusschT.Text = ""
    tbAusschQ.Text = ""
        
    If optBtn = optAusschuss Then
        frmAussch.Enabled = True

        lblN.ForeColor = vbBlack
        lblT.ForeColor = vbBlack
        lblQ.ForeColor = vbBlack
    Else
        frmAussch.Enabled = False
        lblT.ForeColor = &H80000011
        lblQ.ForeColor = &H80000011
        lblN.ForeColor = &H80000011
    End If
    
    Select Case optBtn.Name
        Case "optAusschuss"
            btnEingabe.Enabled = False
        Case Else
            btnEingabe.Enabled = True
    End Select
    
End Sub

Private Sub btnEingabe_Click()
'----------------------------------------------------------------------------------------------------
'   Generate report string for comments
'----------------------------------------------------------------------------------------------------
    Dim problemStr As String
    
    problemStr = Me.Caption
    
    Select Case optBtn
        Case optAusschuss
            problemStr = problemStr & ", Ausschuss"
            
        Case optNachgearb
            problemStr = problemStr & ", Nacharbeit möglich"
            
        Case optSonst
            problemStr = problemStr & ", " & tbBeschreib
            
    End Select
    
    innerString = problemStr
    
    tbBeschreib = "Beschreiben Sie bitte das Problem ... "
    btnEingabe.Enabled = False
    problemStr = ""
    Unload Me
    
    If optAusschuss.value Then
        Call frmMove.btnEingabe_Click
    End If
    
End Sub

'---------------------------------------------------------------------------------------------------
'   Keyboard input handling
'---------------------------------------------------------------------------------------------------

Private Sub tbAussch_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'MsgBox KeyCode
    Select Case KeyCode
    
        'Backspace
        Case 8
            If Len(tbAussch) > 0 Then
                'MsgBox Len(SpecTextBox.value)
                tbAussch.value = Left(tbAussch.value, Len(tbAussch.value))
            End If
            
        'Tab
        Case 9
            KeyCode = vbNull
            
            If frmAussch.ActiveControl.TabIndex < 2 Then
                frmAussch.Controls(frmAussch.ActiveControl.TabIndex + 1).SetFocus
            Else
                frmAussch.Controls(0).SetFocus
            End If
        
        'Enter
        Case 13
            If btnEingabe.Enabled = True Then
                btnEingabe_Click
            End If
            Exit Sub
            
        'Num keys from both numpad and over-letter-array
        Case 48 To 57, 96 To 105
            If Len(tbAussch) = 15 Then
                KeyCode = 0
            End If
            
        'Allows for non-integer values in "time for part"
        'also convert decimal points into commas
        Case 110, 188, 190
            If tbAussch.Name = tbAusschT.Name And Len(tbAussch) > 0 And _
                Not (tbAussch Like "*,") And Not (tbAussch Like "*,*") Then
                    KeyCode = 110
            Else
                KeyCode = 0
            End If
 
        'prevents all other inputs
        Case Else
            KeyCode = 0
            
    End Select

End Sub

Private Sub tbAussch_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If Len(tbAusschQ) > 0 And Len(tbAusschT) > 0 Then
        btnEingabe.Enabled = True
    Else
        btnEingabe.Enabled = False
    End If
    
    If Len(tbAussch) > 0 Then
        Select Case Right(tbAussch.Name, 1)
            Case "Q"
                frmMove.tbAussch = CInt(tbAusschQ)
            Case "N"
                frmMove.tbTlNmmr = tbAusschN
            Case "T"
                frmMove.tbStckZeit = CDbl(tbAusschT)
                
        End Select
    End If
    
End Sub
