VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlMaschinen 
   Caption         =   "Maschinenstörung"
   ClientHeight    =   3090
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFehlMaschinen.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlMaschinen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Input of Maschinenstörung                                                         *
'*  Last Change: 16:02 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Public WithEvents optBtn As MSForms.OptionButton
Attribute optBtn.VB_VarHelpID = -1

'----------------------------------------------------------------------------------------------------
'   General form aesthetics
'----------------------------------------------------------------------------------------------------

Private Sub UserForm_Activate()
Dim frmFehlMaschinen As CFormChanger

Set frmFehlMaschinen = New CFormChanger

With frmFehlMaschinen

    .setIconFromWS
    Set .Form = Me
    
    '.Modal = True
End With
End Sub


'----------------------------------------------------------------------------------------------------
'   Selecting one of the options enables "Eingabe" button
'----------------------------------------------------------------------------------------------------

Private Sub optInstand_Click()
    btnEingabe.Enabled = True
End Sub

Private Sub optSelbst_Click()
    btnEingabe.Enabled = True
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


Private Sub btnEingabe_Click()

    Select Case optBtn
        Case optInstand
            innerString = Me.Caption & ", behoben bei Instandbehaltung"
        
        Case optSelbst
            innerString = Me.Caption & ", selbständig behoben"
            
        Case optWarmlauf
            innerString = Me.Caption & ", Warmlauf"
                    
    End Select
    
    optBtn.value = False
    frmFehlMaschinen.Hide
End Sub


