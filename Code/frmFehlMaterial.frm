VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlMaterial 
   Caption         =   "Materialprobleme"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   OleObjectBlob   =   "frmFehlMaterial.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'                                                                                                   *
'   Form Purpose: Input of Materialprobleme                                                         *
'   Last Change: 16:36 29.06.2016                                                                   *
'                                                                                                   *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Public WithEvents optBtn As MSForms.OptionButton
Attribute optBtn.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------------------
'   General form aesthetics
'----------------------------------------------------------------------------------------------------
Private Sub UserForm_Activate()
Dim frmFehlMaterial As CFormChanger

Set frmFehlMaterial = New CFormChanger

With frmFehlMaterial
    '.ShowCloseBtn = False
    '.ShowSysMenu = False
    
    '.IconPath = Application.ActiveWorkbook.Path & "\Uhlmann_Logo.ico"
    .setIconFromWS
    '.ShowIconWS = True
    Set .Form = Me
    
    '.Modal = True
End With
End Sub

Private Sub UserForm_Initialize()

'Set options to drop down list
With Me.comboQual
    .AddItem "Lunker"
    .AddItem "Farbunterschied"
    .AddItem "Materialschlüsse"
    .value = "Lunker"
End With

End Sub

'----------------------------------------------------------------------------------------------------

Private Sub optAuftraege_Enter()
    Set optBtn = optAuftraege
End Sub

Private Sub optFalsch_Enter()
    Set optBtn = optFalsch
End Sub

Private Sub optGeplant_Enter()
    Set optBtn = optGeplant
End Sub

Private Sub optGesaegt_Enter()
    Set optBtn = optGesaegt
End Sub

Private Sub optKrumm_Enter()
    Set optBtn = optKrumm
End Sub

Private Sub optQualitaet_Enter()
    Set optBtn = optQualitaet
End Sub

Private Sub optBtn_Click()

    If optBtn.Name = "optQualitaet" Then
        comboQual.Enabled = True
    Else
        comboQual.Enabled = False
    End If
End Sub

'----------------------------------------------------------------------------------------------------------------------------------
'   Pressing the enter button
'----------------------------------------------------------------------------------------------------------------------------------
Private Sub optBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = 13 Then
        If optBtn.value = False Then
            optBtn.value = True
        ElseIf btnEingabe.Enabled = True Then
            btnEingabe_Click
        End If
    End If
    
End Sub


'----------------------------------------------------------------------------------------------------------------------------------
'   Register data if valid
'----------------------------------------------------------------------------------------------------------------------------------
Private Sub btnEingabe_Click()
Dim ctrl As Control
Dim problemStr As String

    If tbZeichnum.value = "" Then
        MsgBox ("Bitte geben Sie einen Zeichnungsnummer ein"), (vbCritical), ("Zeichnungsnummer fehlt")
        
        Exit Sub
    End If
    
    problemStr = optBtn.Caption
    
    If optBtn.Name = "optQualitaet" Then
        problemStr = problemStr & "-" & comboQual
    End If
    
    problemStr = problemStr & ", Zeich. Num: " & tbZeichnum.value
    
    innerString = problemStr
    
    ' Clear all textboxes
    problemStr = ""
    tbZeichnum = ""
    comboQual = "Lunker"
    comboQual.Enabled = False
    
    frmFehlMaterial.Hide

End Sub

Private Sub tbZeichnum_Change()

    If tbZeichnum.value = "" Then
        btnEingabe.Enabled = False
    Else
        btnEingabe.Enabled = True
    End If
    
End Sub
