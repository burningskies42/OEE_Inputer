VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlRuesten 
   Caption         =   "Rüsten"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   OleObjectBlob   =   "frmFehlRuesten.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlRuesten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Report problems of type WOP                                                       *
'*  Last Change: 07:32 08.08.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Public WithEvents optBtn As MSForms.OptionButton
Attribute optBtn.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------------------
'   General aestetics
'----------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
Dim frmFehlRuesten As CFormChanger

    Set frmFehlRuesten = New CFormChanger
    frmFehlRuesten.setIconFromWS
    
    Set frmFehlRuesten.Form = Me

    tbAuftrNm.Enabled = False
    tbAuftrNm.Text = "Auftragsnummer ..."

End Sub

Private Sub optNeuAuftr_Enter()
    Set optBtn = optNeuAuftr
End Sub

Private Sub optWeitSpann_Enter()
    Set optBtn = optWeitSpann
End Sub

Private Sub optBtn_Change()
    
    If optBtn.Name = "optNeuAuftr" Then
        tbAuftrNm.Enabled = True
        tbAuftrNm.Text = ""
    ElseIf optBtn.Name = "optWeitSpann" Then
        tbAuftrNm.Enabled = False
        tbAuftrNm.Text = "Auftragsnummer ..."
    End If
    
    enbaleEingabeBtn
    
End Sub

Private Sub tbAuftrNm_Change()
    enbaleEingabeBtn
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

Private Sub btnEingabe_Click()
    Select Case optNeuAuftr
        Case True
            innerString = "Neuer Auftrag, " & tbAuftrNm.value
            isNeuerAuftragRuesten = True
        Case False
            innerString = "Weitere Spannung"
    End Select

    frmFehlRuesten.Hide
End Sub

Private Sub enbaleEingabeBtn()
    If (optNeuAuftr.value = True And Len(tbAuftrNm.value) > 0) Or _
        (optWeitSpann.value = True) Then
            btnEingabe.Enabled = True
    Else
        btnEingabe.Enabled = False
    End If
End Sub

