Attribute VB_Name = "globalRepository"
Option Explicit
Option Private Module

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Module Purpose: Defintion of global variables                                                   *
'*  Last Change: 08:12 09.08.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

' curr_Shift indicate the shift of the current form
Global curr_Shift As Integer

' init indicates whether the form input is already initiated, or whether an error form was just terminated -
' CURRENTLY INACTIVE
Global init  As Integer

Global currRow As Integer
Global currColumn As Integer
Global isNeuerAuftragRuesten As Boolean
Global doMoveTeilAngabe As Boolean
Global staticDict As cAnlageDict


'Envelope routine for the input
Public Sub startEntry(Optional isNew As Boolean = True)

    Set staticDict = New cAnlageDict

On Error GoTo errHandler:

    ShiftWasEntered = False
    dbPath = ActiveWorkbook.Path & "\OEE_DATABASE\dbOEE_be.mdb"
    
    If frmStart.Visible = True Then frmStart.Hide
    'Sheet is unlocked for editing
    ActiveSheet.Unprotect Password:="aczyM4iu"
    toggleFullscreen (True)
    
    'Setting default value for Zwangsangabe Neuer Auftrag bei Rüsten
    isNeuerAuftragRuesten = False
    
    If isNew = True Then
        'Entire sheet is cleared
        Range("D8:R56") = ""
        
        Range("verf_fs") = ""
        Range("leis_fs") = ""
        Range("qual_fs") = ""
        
        Range("verf_ss") = ""
        Range("leis_ss") = ""
        Range("qual_ss") = ""
        
        Range("verf_ns") = ""
        Range("leis_ns") = ""
        Range("qual_ns") = ""
        
        Range("S8:T56") = ""
        Range("A61:R74") = ""
                   
    'Set table start as selected cell
        Range("D8").value = 2
        Range("D8").Select
    Else
        ' Finds the last cell entered in the previous session.
        ' located in [GlobalRepository]
        findLastCell.Select
    End If
    
'Show the shift selection form form
    If frmMove.Visible = True Then frmMove.Hide
    frmShift.Show
    
Exit Sub
'
errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    'emailimage
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & ", StartEntry"
    saveForm True
End Sub

'Closing the entry routine
Public Sub endEntry()
Dim saveShift As New cDoneShift

On Error GoTo errHandler:

1    Application.DisplayFullScreen = False
2    copyToReport readVals                      'Run [ValueTransfer].[copyToReport] with [OEECalculation].[readVals] as parameter
3    printOEEreport
4    saveForm                                   'Saves to file
5    Worksheets("OEE").Protect Password:="aczyM4iu"
    
    
6   Exit Sub

errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    'emailimage
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & _
    ", endEentry , line " & Erl
    saveForm True
End Sub

'As name suggests
Sub toggleFullscreen(state As Boolean)

    With Application
        .DisplayFullScreen = state
        .DisplayFormulaBar = Not state
        
    End With
        
    With ActiveWindow
        .DisplayWorkbookTabs = Not state
        .DisplayHeadings = Not state
     End With
 
 End Sub

'Prints the OEE report without the control elements
Public Sub printOEEreport()
  
    If MsgBox("Wollen Sie den Formular drucken ?", vbYesNo, "Drucken") = vbYes Then
       With Worksheets("OEE")
       
            'Hide all control elements
           .DrawingObjects.Visible = False
           
           .PageSetup.BlackAndWhite = False
           .PageSetup.LeftFooter = Environ(Application.UserName)
           .Range("Print_Area").PrintOut ActivePrinter:="Adobe PDF"
   
           'MsgBox "Drucken ...", vbInformation, "Druckvorgang"
           .DrawingObjects.Visible = True
       End With
   End If

End Sub
 
'Registers new parts. Code for [frmMove].[btnEingabe]
Public Sub newPartEntry()
 
Dim notepad As Range
Dim i As Range
Dim emptyRow As Integer
Dim outputStr As String
Dim Ausschuss As Integer
1   emptyRow = 0
2  outputStr = ""

On Error GoTo errHandler

3  With frmMove
4        If Sheets("OEE").ProtectContents Then
5            Sheets("OEE").Unprotect Password:="aczyM4iu"
6        End If
7           'Finds next empty cell to register given part
8          'Additionally checks whether this part was alredy registered
9           Set notepad = Worksheets("OEE").Range("A61:A74")
10
11           For Each i In notepad
12              If i = "" Then
13                  Exit For
14              ElseIf i = .tbTlNmmr.Text Then

15                    'A part registry  with the same number was found in the notebook
16                  If MsgBox("Die Auftragsnummer " & .tbTlNmmr.Text & " wurde schon während dieser Schicht registriert. Wollen Sie den Auftrag wieder registrieren ?", _
                          vbYesNo + vbCritical, "Auftrag existiert schon") = vbNo Then
18                        Exit Sub
19                   End If
20               End If
21           Next i
           
           'Read in part details from the textboxes and write them into the notebook
            
22           Range("A" & i.Row) = (.tbTlNmmr)
23          If IsNumeric(.tbStckZeit) Then
24              Range("M" & i.Row) = CDbl(.tbStckZeit)
25           Else
26              MsgBox "Ungültige Werte bei Stückzeit", vbCritical, "Nicht numerische Werte"
27              Exit Sub
28          End If
            
29           If .tbGutteil <> "" Then
30               Range("D" & i.Row) = CInt(.tbGutteil)
31           End If
           
32           If .tbAussch <> "" Then
33               Range("H" & i.Row) = CInt(.tbAussch)
34               Ausschuss = CInt(.tbAussch)
35           Else
36               Ausschuss = 0
37           End If
           
38           outputStr = "Eintrag: " & (.tbTlNmmr) & ", "
           
39           If .tbGutteil <> "" Then
40               outputStr = outputStr & "Gutteile: " & CInt(.tbGutteil) & ", "
41           End If
42           If .tbAussch <> "" Then
43               outputStr = outputStr & "Ausschuss: " & CInt(.tbAussch) & ", "
44           End If
           
45           outputStr = outputStr & "Stückzeit: " & CDbl(.tbStckZeit) & " gemeldet"
           
           'Display confirmation to part registry
46           MsgBox outputStr
               
           'Print out part registry to Maßnahmen
47         setCurrPos
48           'Range("T" & currRow) = "Teilnum. " & (.tbTlNmmr) & " fertig"
           
49           frmMove.tbTlNmmr.SetFocus
    
           'case handling for first row
           
           'First row
50           If currRow = 8 Then
51               moveSelection (2), (.Allg_QualCB)
               
           'Not first row
52           Else
               
               'Not first column
53               If currColumn <> 4 Then
               
                       'Allgemeine Qualitätsprobleme
54                      If frmMove.Allg_QualCB = True Then
55                          If TypeName(currRow) <> "Integer" Or currRow <= 1 Then
56                              setCurrPos
57                          End If
58
59
60                          Worksheets("OEE").Range("R" & Worksheets("OEE").Rows(currRow) + 1) = 1
61                       End If
                       
62                       Worksheets("OEE").Cells(currRow + 1, currColumn).value = 1
63                       Worksheets("OEE").Cells(currRow + 1, currColumn).Select
64                       doMoveTeilAngabe = True
65                       frmMove.Clr_Btn_Click
                   
               'First column
66               Else
                   'Just came from the right (colum offset by 1, is equal 2)
67                   If Worksheets(1).Cells(currRow, currColumn + 1) > 0 Then
68                       moveSelection (2), (.Allg_QualCB)
69                   Else
70                       moveSelection (4), (.Allg_QualCB)
71                   End If
72               End If
               
73           End If
           
74    End With
 
    ' Open frmAusschuss if rejected parts (Ausschusse) were reported
75    If Ausschuss > 0 Then
76            Worksheets("Ausschuss").Range("I6") = frmMove.tbAussch.value
77            Worksheets("Ausschuss").Range("K6") = frmMove.tbAussch.value + frmMove.tbGutteil.value

78        Unload frmMove
79        Worksheets(2).Visible = True
80        Worksheets(1).Visible = False
81        Worksheets(2).Activate
82    End If
    
83  Exit Sub

errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & _
        ", newPartEntry, line " & Erl
        
    saveForm True

End Sub

' Cheacks whether the new part already exists in the new-parts list below
Public Function AuftragExists(auftrNum As Integer) As Boolean
Dim cell As Range
Dim ret As Boolean

ret = False

For Each cell In Worksheets("OEE").Range("neueAufträge")
    If cell.value = auftrNum Then
        ret = True
    End If
Next cell

AuftragExists = ret
End Function

'Find last cell in notebook, where the mover stopped
Public Function findLastCell() As Range
Dim rng As Range
Dim maxRow As Range
Dim maxCell As Range

For Each rng In Worksheets("OEE").Range("D9:Q56").Rows
    If Application.Max(rng) = 0 Then
        Set maxRow = rng.Offset(-1, 0)
        Exit For
    End If
Next rng

If maxRow Is Nothing Then
    Set maxRow = Range("D56:Q56")
End If

With Application
    If .Sum(maxRow) = 1 Or .Sum(maxRow) = 2 Then
        Set maxCell = .Cells(maxRow.Row, .WorksheetFunction.Match(.WorksheetFunction.Max(maxRow), maxRow, 0) + 3)
    Else
        For Each rng In maxRow.Cells
            If (rng.Offset(0, 1) > 0 And rng = 2 And rng.Offset(0, -1) = 0) Or _
                (rng.Offset(0, 1) = 0 And rng = 2 And rng.Offset(0, -1) > 0) Or _
                (rng = 2 And rng.Column = 4) Or _
                (rng = 2 And rng.Column = 17) Then
                    Set maxCell = Range(rng.Address)
            End If
        Next rng
    End If
End With

Set findLastCell = maxCell

End Function

'------------------------------------------------------------------------------------
' Clears the entire form
'------------------------------------------------------------------------------------
Public Sub clearOEEform()
    
    With ActiveWorkbook.Worksheets("OEE")
        Application.DataEntryMode = xlOff
        .Unprotect Password:="aczyM4iu"
        .Range("D8:Y56").ClearContents
        .Range("A61:R74").ClearContents
        .Range("V72:X74").ClearContents
        .Range("A7:C7").Interior.Color = vbWhite
        .Protect Password:="aczyM4iu"
    End With
End Sub


Public Sub setCurrPos()
Dim target As Range
    
    ActiveWorkbook.Worksheets("OEE").Select
    Set target = Selection
    
    currRow = target.Row
    currColumn = target.Column
End Sub

Sub AddCode()
Dim xPro As VBIDE.VBProject
Dim xCom As VBIDE.VBComponent
Dim xMod As VBIDE.CodeModule
Dim xLine As Long
Dim i As Integer

    With ThisWorkbook
        Set xPro = .VBProject
        Set xCom = xPro.VBComponents("Movement")
        Set xMod = xCom.CodeModule

        With xMod

            For i = 1 To xMod.CountOfLines
                .ReplaceLine i, i & "   " & .Lines(i)
            Next i
            
        End With
    End With

End Sub

