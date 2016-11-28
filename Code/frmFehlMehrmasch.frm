VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlMehrmasch 
   Caption         =   "Mehrmaschinenbedienung"
   ClientHeight    =   2850
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFehlMehrmasch.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlMehrmasch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Form Purpose: Input of Mehrmaschinebedienung
'Last Change: 08:30 25.10.2016

Public WithEvents optBtn As MSForms.OptionButton
Attribute optBtn.VB_VarHelpID = -1
'General form aestetics
Private Sub UserForm_Activate()
Dim frmFehlMaterial As CFormChanger

Set frmFehlMaterial = New CFormChanger

With frmFehlMaterial

    .setIconFromWS
    Set .Form = Me

End With

End Sub

'   Add option buttons to group
Private Sub optStamm_Enter()
    Set optBtn = optStamm
End Sub
Private Sub optZusatz_Enter()
    Set optBtn = optZusatz
End Sub

Private Sub optBtn_Click()
    If optBtn.Name = "optZusatz" Then
        tbZusatzMasch.Enabled = True
        tbZusatzMasch.Text = ""
        tbZusatzMasch.SetFocus
    Else
        tbZusatzMasch.Text = "Zusatzmaschine nr. ..."
        tbZusatzMasch.Enabled = False
    End If
    
    btnEingabe.Enabled = True
    
End Sub

'   option buttons group events
Private Sub optBtn_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If optBtn.Name = "optZusatz" Then
        tbZusatzMasch.Enabled = True
        tbZusatzMasch.Text = ""
    Else
        tbZusatzMasch.Text = "Zusatzmaschine nr. ..."
        tbZusatzMasch.Enabled = False
    End If
  
    If KeyCode = 13 Then
        If optBtn.value = False Then
            optBtn.value = True
        Else
            btnEingabe_Click
        End If
    End If
End Sub

Private Sub btnEingabe_Click()
Dim ctrl As Control
Dim problemStr As String
'
'    For Each ctrl In Me.Controls
'        If TypeOf ctrl Is MSForms.OptionButton Then
'            If ctrl.value = True Then
'                problemStr = ctrl.Caption
'                ctrl.value = ""
'            End If
'        End If
'    Next

    innerString = Me.Caption

    Select Case optBtn.Name
        Case "optStamm"
            innerString = innerString & " bei Stammmaschine"
        Case "optZusatz"
            innerString = innerString & " bei Zusatzmschine"
            
            If Len(tbZusatzMasch) > 0 And Left(tbZusatzMasch, 6) <> "Zusatz" Then
                  innerString = innerString & ": " & tbZusatzMasch.value
            End If
    End Select
    
    doMoveTeilAngabe = True
    tbZusatzMasch = ""
    optBtn.value = False
    btnEingabe.Enabled = False
    Me.Hide
    
End Sub
