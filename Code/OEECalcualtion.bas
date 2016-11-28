Attribute VB_Name = "OEECalcualtion"
Option Explicit
Option Private Module

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Module Purpose: Input of OEE values from wht worksheet                                          *
'*  Last Change: 07:53 16.09.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------


Public Function readVals(Optional withEnd As Boolean = True) As COEE_val

Dim Report As Worksheet
Dim newOEE As COEE_val
    
    On Error GoTo errHandler:
    
1    Set newOEE = New COEE_val
    
    'Assign the active sheet to the variable.
2    Set Report = Excel.ActiveWorkbook.Worksheets("OEE")
            
    ' Insert all values into a new instance of OEE class
3    With newOEE
    
4        .Anlage = Report.Range("Anlage")
5        .Datum = Report.Range("T2")
        
6        Select Case curr_Shift
            Case 1
                .Schicht = "Früh"
            Case 2
                .Schicht = "Spät"
            Case 3
                .Schicht = "Nacht"

        End Select
        
14        If .Schicht = "" Then
15            Select Case ActiveWorkbook.Worksheets("OEE").Range("Schicht")
                Case 1
                    .Schicht = "Früh"
                Case 2
                    .Schicht = "Spät"
                Case 3
                    .Schicht = "Nacht"
            End Select
23        End If
        
24        .Mehrmaschinenbedienung = Report.Range("E57")
25        .Geplante_Stillstaende = Report.Range("F57")
26        .Ruesten = Report.Range("G57")
27        .Material_fehlt = Report.Range("HH7")
28        .Personal_fehlt = Report.Range("I57")
29        .Schlosser = Report.Range("J57")
30        .Stoerung = Report.Range("K57")
31        .Materialprobleme = Report.Range("L57")
32        .Qualitaetsprobleme = Report.Range("M57")
33        .Zeichnung_unklar = Report.Range("N57")
34        .Avprog_fehlt_ueberarbeiten = Report.Range("O57")
35        .WOP = Report.Range("P57")
36        .Abweichung_Planzeit = Report.Range("Q57")
37        .Allg_Qualitaetsprobleme = Report.Range("R57")
        
38        .StillZeit = Excel.WorksheetFunction.Sum(Report.Range("E57:P57"))
39        .Auftragzeit = Excel.WorksheetFunction.Sum(Report.Range("auftrag_zeit"))
40        .Ausschuss = Excel.WorksheetFunction.Sum(Report.Range("sum_ausschuss"))
41        .Gutteile = Excel.WorksheetFunction.Sum(Report.Range("sum_gutteile"))
        
42        .Laufzeit = .BetriebZeit - .StillZeit
        
43        If (Not IsNumeric(.BetriebZeit)) Or .BetriebZeit <= 0 Then
44            .val_d = 0
45        Else
46            .val_d = .Laufzeit / .BetriebZeit
47        End If
        
48        If (Not IsNumeric(.Laufzeit)) Or .Laufzeit <= 0 Then
49            .val_e = 0
50        Else
51            .val_e = .Auftragzeit / .Laufzeit
52        End If
        
53        .val_f = .Gutteile / Application.Max(.Ausschuss + .Gutteile, 1)
        
54        Select Case .Schicht
            Case "Früh"
                Report.Range("verf_fs") = .val_d
                Report.Range("leis_fs") = .val_e
                Report.Range("qual_fs") = .val_f
            Case "Spät"
                Report.Range("verf_ss") = .val_d
                Report.Range("leis_ss") = .val_e
                Report.Range("qual_ss") = .val_f
            Case "Nacht"
                Report.Range("verf_ns") = .val_d
                Report.Range("leis_ns") = .val_e
                Report.Range("qual_ns") = .val_f
            End Select
                
68    End With
    
69    Set readVals = newOEE
    
70    Exit Function
    
errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    saveForm True
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & _
        ", readVals, line " & Erl

End Function


