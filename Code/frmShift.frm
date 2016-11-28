VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmShift 
   Caption         =   "Schicht Auswahl"
   ClientHeight    =   2040
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   3945
   OleObjectBlob   =   "frmShift.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Form Purpose: Selection of the current shift
'Last Change: 18:24 08.06.2016

' curr_Shift is a global integer for the shift
' values: 1-früh 2-spät 3-nacht

Public WithEvents radioBtn As MSForms.OptionButton
Attribute radioBtn.VB_VarHelpID = -1
Dim local_curr_Shift As Integer

Private Sub opt_FS_Enter()
    Set radioBtn = opt_FS
End Sub

Private Sub opt_SS_Enter()
    Set radioBtn = opt_SS
End Sub

Private Sub opt_NS_Enter()
    Set radioBtn = opt_NS
End Sub

Private Sub UserForm_Activate()

Dim frmShift As CFormChanger
Set frmShift = New CFormChanger

With frmShift
    .ShowCloseBtn = False
    '.ShowSysMenu = False
    
    '.IconPath = Application.ActiveWorkbook.Path & "\Uhlmann_Logo.ico"
    .setIconFromWS
    Set .Form = Me
    
    'frmShift.Modal = True

End With

'When the form loads, the default value of the shift is selected according to the current time
    
    Select Case Hour(Time)
        Case 5 To 12
            opt_FS.value = True

        Case 13 To 20
            opt_SS.value = True

        Case 21 To 24, 0 To 4
            opt_NS.value = True
    End Select
   
    
End Sub


Private Sub Inp_Btn_Click()

Dim oCtl As MSForms.Control

    For Each oCtl In Me.Controls
        If TypeName(oCtl) = "OptionButton" And oCtl.value = True Then
            curr_Shift = oCtl.Tag
            
            ' Color the appropriate shift cell gray
            Select Case curr_Shift
                Case 1
                    Worksheets("OEE").Range("A7").Interior.Color = vbYellow
                    Worksheets("OEE").Range("B7").Interior.Color = vbWhite
                    Worksheets("OEE").Range("C7").Interior.Color = vbWhite
                Case 2
                    Worksheets("OEE").Range("A7").Interior.Color = vbWhite
                    Worksheets("OEE").Range("B7").Interior.Color = vbYellow
                    Worksheets("OEE").Range("C7").Interior.Color = vbWhite
                    
                Case 3
                    Worksheets("OEE").Range("A7").Interior.Color = vbWhite
                    Worksheets("OEE").Range("B7").Interior.Color = vbWhite
                    Worksheets("OEE").Range("C7").Interior.Color = vbYellow
                           
            End Select
        End If
    Next oCtl
    
    With Worksheets("OEE")
        If RecExists(createKey(.Range("Anlage"), .Range("T2"), curr_Shift)) = True Then
            If MsgBox("Ein Datensatz für die gegebene Anlage, Datum und schicht existiert schon in der Datenbank. Wollen Sie ihn überschreiben ?", vbYesNo, "Datensatz vorhanden") = vbYes Then
                frmShift.Hide
                frmMove.Show
            Else
                frmShift.Hide
                Application.DisplayFullScreen = False
                Worksheets("OEE").Protect Password:="aczyM4iu"
            End If
        Else
            Unload frmShift
            frmMove.Show
        End If
        
        .Range("Schicht") = curr_Shift
        
    End With
    


End Sub

