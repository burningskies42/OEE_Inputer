Attribute VB_Name = "Movement"
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Module Purpose: Defintion of movement on the timetable                                          *
'*  Last Change: 06:58 11.08.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------


Dim Report As Worksheet 'Set up new worksheet variable.
Dim cell As Range

Public Sub moveSelection(dirc As Integer, allg_qual As Boolean)
'Control of the movement with arrow buttons
'dirc indicates the inputed direction from frmMove

On Error Resume Next:

0      Application.ScreenUpdating = False
1
2   Set Report = Excel.ActiveSheet 'Assign the active sheet to the variable.
3   Set cell = findLastCell

    If Sheets("OEE").ProtectContents Then
        Sheets("OEE").Unprotect Password:="aczyM4iu"
    End If
4
5    Dim bearb As Variant
6
7    setCurrPos
8
9    'Up - used to backtrack/ delete erroneous input
10    If dirc = 1 Then
11
12    'check that the destination is valid
13   If cell.Offset(-1, 0).value > 0 And cell.Row > 8 Then
14
15            'Allgemeine Qualit‰tsprobleme
16           If allg_qual = True Then
17                Report.Range("R" & cell.Row) = ""
18           End If
19
20            cell.value = ""
21            cell.Offset(-1, 0).Select
22
23        End If
24
25       ActiveWindow.ScrollRow = currRow - 7
26
27
28    'DOWN
29    ElseIf dirc = 2 Then
30
31
32
33        If cell.Row < 56 Then
34
35            'Forces new part input if in most upper-left cell
'            With frmMove
'                If cell.Address = "$D$8" And (Len(.tbTlNmmr) * Len(.tbGutteil + .tbAussch) * Len(.tbStckZeit)) = 0 Then
'                    MsgBox "Bitte geben Sie die Teilangabe ein", vbCritical, "Teilangabe"
'                    Exit Sub
'                End If
'            End With

            'Input a malfunction if first movement down in current column
37            If (cell.value = 2) Then
38                If cell.Column > 4 Then
39                     Select Case cell.Column
                            Case 5 ' Fertigung im Einsatz
                                Worksheets("OEE").Range("S" & currRow) = probInput(cell.Column)
                            Case 9  'Personal fehlt
                                Worksheets("OEE").Range("S" & currRow) = probInput(cell.Column)
                            Case 10 'Schlosser T‰tigkeit
                                Worksheets("OEE").Range("S" & currRow) = probInput(cell.Column)
                            Case 16
                                Worksheets("OEE").Range("S" & currRow) = probInput(cell.Column)
                            Case 17
        
                            Case Else
                                frmStrNeuTeil.Show
                        End Select
40
41                Else
42                    doMoveTeilAngabe = True
43                End If
44            End If
45
46            'Move only if either new part entry succeeded or malfunction reported
47            If doMoveTeilAngabe = True Then
48                'Allgemeine Qualit‰tsprobleme
49                If allg_qual = True Then
50                    Report.Range("R" & cell.Row + 1) = 1
51                End If
52
53                cell.Offset(1, 0).value = 1
                  If ActiveSheet.Name = "OEE" Then
54                   cell.Offset(1, 0).Select
                  End If
55                doMoveTeilAngabe = True
56            End If
57          If ActiveSheet.Name = "OEE" And currRow - 7 > 1 Then
58              ActiveWindow.ScrollRow = currRow - 7
59          End If

60        Else
61
62          ' END OF TIMETABLE REACHED
63          '1. Read data from timetable, calculate OEE values and display them in form
64          'Function[readVals].module[OEECalculation]
65            frmMove.Hide
66
67          '2. Throw message when end of time is reached
68            If MsgBox("Der Eintrag ist fertig. Wollen Sie die Eingabe noch bearbeiten ?", vbYesNo, "Ende der Schicht") = vbYes Then
69
70              '3. Copy all relevant values to worksheet[Report]
71              'Fuction[copyToReport].module[ValueTransfer]
72                readVals (False)
73                toggleFullscreen (False)
74                frmBearb.Show vbModeless
75            Else
76                endEntry
77            End If
78
79        End If
80  ' LEFT
81    ElseIf dirc = 3 Then
82
83        If cell.Column > 4 Then
84            If cell.Row = 8 Then
85                cell.value = 0
86                cell.Offset(, -1).value = 2
87                cell.Offset(, -1).Select
88            Else
89
90                If cell.Offset(, -1).value > 0 Then
91                    cell.value = 0
92                Else
93                    cell.Offset(, -1).value = 2
94                End If
95
96                If cell.Column = 11 Then
97                    ifNichtBestellt
98                End If
99                cell.Offset(, -1).Select
100            End If
101       Else
102
103       End If
104
105
106    ' RIGHT
107    ElseIf dirc = 4 Then
108       If cell.Column < 17 Then
109           If cell.Row = 8 Then
110               cell.value = 0
111               cell.Offset(, 1).value = 2
112               cell.Offset(, 1).Select
113            Else
114               'Backtracking on written cells, rewrite
115               If cell.Offset(, 1).value > 0 Then
116                   cell.value = 0
117               Else
118                   cell.Offset(0, 1).value = 2
119               End If
120               If cell.Column = 11 Then
121                   ifNichtBestellt
122               End If
123               cell.Offset(0, 1).Select
124           End If
125       Else
126
127       End If
128
129    ' Impossible movement value
130    Else
131        MsgBox dirc & " is an impossible value"
132
133  End If
134
135    ' Clear all part input on frmMove
136  With frmMove
137       .tbTlNmmr = ""
138       .tbGutteil = ""
140       .tbAussch = ""
141       .tbStckZeit = ""
142   End With
143
144   Application.ScreenUpdating = True
145
146
147 Exit Sub

errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    saveForm True
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & _
        ", moveSelection, line " & Erl

    
End Sub

Private Sub ifNichtBestellt() 'ifVerschlieﬂ()
    Set Report = Excel.ActiveSheet 'Assign the active sheet to the variable.
    Set cell = Selection
    
    Dim looper As Range
    Set looper = cell
    
    Set looper = Report.Range("S" & cell.Row)
    
    Do While InStr(1, looper.value, "Nicht bereit Bestellt") = 0
        Set looper = looper.Offset(-1, 0)
        
        If looper.value <> "" And InStr(1, looper.value, "Nicht bereit Bestellt") = 0 Then
            Exit Sub
        End If
        
    Loop
    
    With Range("S" & (cell.Row - 1))
        If .value <> "" Then
            .value = .value & ", "
        End If
        
        Range("S" & (cell.Row - 1)) = Range("S" & (cell.Row - 1)) & _
            "geliefert um " & Time()
    End With
    
End Sub


