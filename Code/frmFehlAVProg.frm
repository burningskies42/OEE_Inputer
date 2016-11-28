VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlAVProg 
   Caption         =   "AVProg fehlt"
   ClientHeight    =   3735
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFehlAVProg.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlAVProg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Input of AVProg                                                                   *
'*  Last Change: 13:22 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------


Public WithEvents optBtn As MSForms.OptionButton
Attribute optBtn.VB_VarHelpID = -1

'General form aesthetics
Private Sub UserForm_Activate()
Dim frmFehlAVProg As CFormChanger

Set frmFehlAVProg = New CFormChanger

    With frmFehlAVProg
        .setIconFromWS
        Set .Form = Me
    
    End With

End Sub

'Add all optionbuttons to general option scheme
Private Sub optBear_Enter()
    Set optBtn = optBear
End Sub
Private Sub optHeflt_Enter()
    Set optBtn = optHeflt
End Sub

Private Sub optBtn_Click()
    Select Case optBtn
        Case optBear
            tbBegrund.Enabled = True
            tbBegrund = ""
        
        Case optHeflt
            tbBegrund.Enabled = False
            tbBegrund = "Begründung ..."
    End Select
    
End Sub

Private Sub optBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        If optBtn.value = False Then
            optBtn.value = True
        ElseIf btnEingabe.Enabled = True Then
            btnEingabe_Click
        End If
    End If
End Sub

Private Sub tbZeichNum_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'Enable input button only if textboxes are non empty
    If tbZeichnum = "" Then
        btnEingabe.Enabled = False
    Else
        btnEingabe.Enabled = True
    End If

End Sub


Private Sub btnEingabe_Click()

'Generate report string for comments
    Dim problemStr As String
    
    If optBear = True Then
        problemStr = "AVProg, " & tbBegrund
    Else
        problemStr = "AVProg fehlt, "
    End If
    
    problemStr = problemStr & ", Zeich. Num: " & tbZeichnum.value
    
    tbBegrund = "Begründung ..."
    tbZeichnum = ""
    
    btnEingabe.Enabled = False
    innerString = problemStr
    frmFehlAVProg.Hide
    
End Sub
