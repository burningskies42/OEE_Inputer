VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFehlWerkzeug 
   Caption         =   "Werkzeugstörung"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4950
   OleObjectBlob   =   "frmFehlWerkzeug.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFehlWerkzeug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Form Purpose: Input of Werkzeugstörung Beschreibung
'Last Change: 14:33 22.08.2016

Public WithEvents tbWerkzeug As MSForms.TextBox
Attribute tbWerkzeug.VB_VarHelpID = -1
Public WithEvents optBtn As MSForms.OptionButton
Attribute optBtn.VB_VarHelpID = -1

'General aestetics
Private Sub UserForm_Activate()

Dim frmFehlWerkzeug As CFormChanger

Set frmFehlWerkzeug = New CFormChanger
With frmFehlWerkzeug
    .setIconFromWS
    Set .Form = Me
End With

'set default values of text boxes
    tbZeitStnd = Right("00" & Hour(Now()), 2)
    tbZeitMnt = Right("00" & Minute(Now()), 2)
    
    tbEinsatzStnd = Right("00" & Hour(Now()), 2)
    tbEinsatzMnt = Right("00" & Minute(Now()), 2)
    
    tbBstlDay = Right("00" & Day(Now()), 2)
    tbBstlMnt = Right("00" & Month(Now()), 2)
    tbBstlYr = Year(Now()) - 2000
End Sub

Private Sub MultiPage1_Change()
    checkAllTb
    optBtn.value = False
End Sub

'1 Page
Private Sub tbTNumBrch_Enter()
    Set tbWerkzeug = tbTNumBrch
End Sub

Private Sub tbZeichnum_Enter()
    Set tbWerkzeug = tbZeichnum
End Sub

'2 Page
Private Sub tbTNumVersch_Enter()
    Set tbWerkzeug = tbTNumVersch
End Sub
Private Sub tbZeitStnd_Enter()
    Set tbWerkzeug = tbZeitStnd
End Sub
Private Sub tbZeitMnt_Enter()
    Set tbWerkzeug = tbZeitMnt
End Sub

'3 Page
Private Sub tbEinsatzStnd_Enter()
    Set tbWerkzeug = tbEinsatzStnd
End Sub
Private Sub tbEinsatzMnt_Enter()
    Set tbWerkzeug = tbEinsatzMnt
End Sub
Private Sub tbBstlDay_Enter()
    Set tbWerkzeug = tbBstlDay
End Sub
Private Sub tbBstlMnt_Enter()
    Set tbWerkzeug = tbBstlMnt
End Sub
Private Sub tbBstlYr_Enter()
    Set tbWerkzeug = tbBstlYr
End Sub

'4 Page
Private Sub optInstand_Enter()
    Set optBtn = optInstand
End Sub
Private Sub optSelbst_Enter()
    Set optBtn = optSelbst
End Sub
Private Sub optWarmlauf_Enter()
    Set optBtn = optWarmlauf
End Sub

Private Sub optBtn_Click()
    btnEingabe.Enabled = True
End Sub

Private Sub tbWerkzeug_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)


    Select Case tbWerkzeug.Name
        Case "tbTNumVersch", "tbTNumBrch"
            'Allow all input
            
        Case "tbZeichNum"
            Select Case KeyCode
                    Case 8 'Backspace
                        If Len(tbWerkzeug) > 0 Then tbWerkzeug.value = Left(tbWerkzeug.value, Len(tbWerkzeug.value))
                        
                    Case 13 'Enter
                        If btnEingabe.Enabled = True Then btnEingabe_Click
                        
                    Case 48 To 57, 96 To 105 'Num keys from both numpad and over-letter-array
                        'Allows input of 2 nums only in time and date fields
                    
                    Case Else
                        KeyCode = 0
                        
                End Select
            
        Case Else
            'is allowed to have 2 chars at most
            If Len(tbWerkzeug) >= 2 And KeyCode <> 8 And KeyCode <> 13 Then
                KeyCode = 0
            Else
                'Check input according to keycodes
                Select Case KeyCode
                    Case 8 'Backspace
                        If Len(tbWerkzeug) > 0 Then tbWerkzeug.value = Left(tbWerkzeug.value, Len(tbWerkzeug.value))
                        
                    Case 13 'Enter
                        If btnEingabe.Enabled = True Then btnEingabe_Click
                        
                    Case 48 To 57, 96 To 105 'Num keys from both numpad and over-letter-array
                        'Allows input of 2 nums only in time and date fields
                    
                    Case Else
                        KeyCode = 0
                        
                End Select
            End If
    End Select
    
End Sub

Private Sub tbWerkzeug_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    checkAllTb
End Sub

Private Sub checkAllTb()
Dim cCont As Control

    ' Enable Eingabe if either:
    '   1. On pages 1 - 3 all textboxes are full
    '   2. On page 4 one option is selected
    Select Case Me.MultiPage1.value
        Case 0 To 2
            For Each cCont In Me.MultiPage1.Pages(Me.MultiPage1.value).Controls
                If TypeName(cCont) = "TextBox" Then
                    If cCont.value = "" Then
                        btnEingabe.Enabled = False
                        Exit Sub
                    End If
                End If
            Next cCont
            btnEingabe.Enabled = True
            
        Case 3
            For Each cCont In Me.MultiPage1.Pages(Me.MultiPage1.value).Controls
                If TypeName(cCont) = "OptionButton" Then
                    If cCont.value = True Then
                        btnEingabe.Enabled = True
                        Exit For
                    End If
                End If
            Next cCont
            btnEingabe.Enabled = False
            
    End Select
    
End Sub


Private Sub btnEingabe_Click()
Dim ctrl As Control

' Checker is a temp variable for checking that no fields were left blank
Dim checker As String

' checker is given the default value "NULL", which is changed only if all necessary textboxes were filled
checker = "NULL"

checker = errorInput

    ' Malfunction report is not registered until all text boxes are filled
    If checker <> "NULL" Then
            innerString = checker
    Else
        Exit Sub
    End If
    
    'Continue navigation
    init = False ' -------------not in use
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.value = ""
        ElseIf TypeName(ctrl) = "OptionButton" Then
            ctrl.value = False
        End If
    Next ctrl
    
    btnEingabe.Enabled = False
    frmFehlWerkzeug.Hide
End Sub

Public Function errorInput() As String

'errorInput validates the input of the txtboxes and returns the malfunction description
    Dim retString As String
    Dim optB As Object

'Me.MultiPage1.value denotes the 3 tabs of the form
    Select Case Me.MultiPage1.value
    
        Case 0 'Bruch
            If tbTNumBrch = "" Or tbZeichnum = "" Then
                retString = "NULL"
            Else
                retString = Me.MultiPage1.Pages(0).Caption & _
                    ", T-Num: " + tbTNumBrch + " Zeichnum: " + tbZeichnum
            End If
            
        Case 1 'Verschleiß
            If tbTNumVersch = "" Then
                retString = "NULL"
            Else
                retString = Me.MultiPage1.Pages(1).Caption & _
                    ", T-Num: " + tbTNumVersch + " bestellt um " + tbEinsatzStnd + ":" + tbEinsatzMnt
            End If

        Case 2 'Nicht bereit bestellt
            If tbBstlDay = "" Or tbBstlMnt = "" Or tbBstlYr = "" Then
                retString = "NULL"
            Else
                retString = Me.MultiPage1.Pages(2).Caption & _
                    ", Bestell. Datum: " + tbBstlDay + "." + tbBstlMnt + "." + tbBstlYr
            End If
        
        
        Case 3 'Maschinenstörung
            retString = "NULL"
            For Each optB In Me.MultiPage1.Pages(Me.MultiPage1.value).Controls
                If TypeName(optB) = "OptionButton" Then
                    If optB.value = True Then
                        retString = Me.MultiPage1.Pages(3).Caption & ": " & optB.Caption
                    End If
                End If
            Next optB
            
    End Select
    
    errorInput = retString
    
End Function

