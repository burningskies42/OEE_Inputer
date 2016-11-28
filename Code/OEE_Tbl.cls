VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OEE_Tbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Sheet Purpose: Main worksheet                                                                   *
'*  Last Change: 14:43 28.11.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

'****************************************************************************************************
'   Help button
'****************************************************************************************************
Private Sub btnHelp_Click()
    logAction "Help request"
    emailMe
End Sub

'****************************************************************************************************
'   Date combo box
'****************************************************************************************************
Private Sub ComboBox1_Change()
    Application.EnableEvents = True
End Sub

'****************************************************************************************************
'   Start or Restart OEE entry using Function[restartEntry].Module[globalRepository]
'****************************************************************************************************
Private Sub Input_Btn_Click()
    If Application.Sum(Worksheets("OEE").Range("D9:R56")) > 0 Then
        frmStart.Show
    Else
        startEntry
    End If
End Sub

'****************************************************************************************************
'check cells for desired format to trigger the frmCalendar.show routine
'otherwise exit the sub
'****************************************************************************************************
Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)

Dim DateFormats, DF
    
Application.EnableEvents = True
'MsgBox Target.NumberFormat
     
    DateFormats = Array("dd/mm/yy;@", "mmmm d JJJJ")
    For Each DF In DateFormats

        If DF = target.NumberFormat Then
            If frmCalendar.HelpLabel.Caption <> "" Then
                frmCalendar.Height = 191 + frmCalendar.HelpLabel.Height
            Else: frmCalendar.Height = 191
                frmCalendar.Show
            End If
        End If
    Next
End Sub

'****************************************************************************************************
'   Movement on the worksheet
'   -either problem registry
'   -or normal movement
'****************************************************************************************************
Private Sub Worksheet_SelectionChange(ByVal target As Range)
Dim inpt As String

    'define coordiantes of currently selected cell
    setCurrPos
    
    ' Inputbox for Maßnahmen
    If target.Row > 7 And target.Row < 57 And target.Column = 20 And frmMove.Visible = True Then
        frmMove.Hide
        Cells(currRow, currColumn) = InputBox("Geben Sie die Maßnahmen ein", "Maßnahmem", Cells(currRow, currColumn))
        findLastCell.Select
        frmMove.Show
        
    End If
    
    ' Inputbox for Fehlerbeschreibung
    If target.Row > 7 And target.Row < 57 And target.Column = 19 And frmMove.Visible = True Then
        frmMove.Hide
        Cells(currRow, currColumn) = InputBox("Geben Sie die Fehlerbeschreibung ein", "Fehlerbeschreibung", Cells(currRow, currColumn))
        findLastCell.Select
        frmMove.Show
    End If
    
    ' Inputbox for the part-details (amount, duration, ID)
    If (target.Row > 60 And target.Row < 75) And (target.Column < 19) _
                And frmMove.Visible = True Then
        frmMove.Hide
        Cells(currRow, currColumn) = InputBox("Geben Sie die Eingabe ein", "Teileingabe", Cells(currRow, currColumn))
        findLastCell.Select
        frmMove.Show
    End If
    
    If target.Row < 57 And target.Column < 19 Then
        'findLastCell.Select
    End If
    
    'findLastCell.Select
    
End Sub
