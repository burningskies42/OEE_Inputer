Attribute VB_Name = "ValueTransfer"
Option Explicit
Option Private Module

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Module Purpose: Copying data from worksheets("OEE") to worksheets("Report") and sequentially    *
'*  to the db                                                                                       *
'*  Last Change: 06:58 11.08.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Global staticDictTransfer As cAnlageDict
Public dbPath As String

'Copies data from an instance of the COEE_val class and write it to worksheets("Report")
Public Sub copyToReport(newOEE As COEE_val, Optional withTrans As Boolean = True)

Dim shiftDate As Date
Dim Anlage As String
Dim OEE_val As Double
Dim Schicht As Integer

    On Error GoTo errHandler:

1    With newOEE

        'Paste all values into sheet Report
        'Key identifiers + registry data
        If Len(.Anlage) = 0 Or IsDate(DateValue(.Datum)) = False Or (curr_Shift < 1 Or curr_Shift > 3) Then
            Set newOEE = readVals
        Else
2           Sheets("report").Range("A2").value = createKey(.Anlage, DateValue(.Datum), curr_Shift)
        End If
        
3        Sheets("report").Range("B2").value = .Anlage
4        Sheets("report").Range("C2").value = .Datum
5        Sheets("report").Range("D2").value = .Schicht
6        Sheets("report").Range("E2").value = Round(.OEE, 2) * 100
7        Sheets("report").Range("F2").value = Now()
8        Sheets("report").Range("G2").value = Environ$("username")
9
        'All malfunction types
10        Sheets("report").Range("H2").value = .Mehrmaschinenbedienung
11        Sheets("report").Range("I2").value = .Geplante_Stillstaende
12        Sheets("report").Range("J2").value = .Ruesten
13        Sheets("report").Range("K2").value = .Material_fehlt
14        Sheets("report").Range("L2").value = .Personal_fehlt
15        Sheets("report").Range("M2").value = .Schlosser
16        Sheets("report").Range("N2").value = .Stoerung
17        Sheets("report").Range("O2").value = .Materialprobleme
18        Sheets("report").Range("P2").value = .Qualitaetsprobleme
19        Sheets("report").Range("Q2").value = .Zeichnung_unklar
20        Sheets("report").Range("R2").value = .Avprog_fehlt_ueberarbeiten
21        Sheets("report").Range("S2").value = .WOP
22        Sheets("report").Range("T2").value = .Abweichung_Planzeit
23        Sheets("report").Range("U2").value = .Allg_Qualitaetsprobleme
        
        'Success parameters
24        Sheets("report").Range("V2").value = .Gutteile
25        Sheets("report").Range("W2").value = .Ausschuss
26        Sheets("report").Range("X2").value = .Laufzeit
27        Sheets("report").Range("Y2").value = .Auftragzeit
    
28    End With
    
29    ActiveWorkbook.Saved = True
    
    'Initiate upload to database
30    ShiftWasEntered = True
    
    ' export to db if withTrans = true
31    If withTrans Then
32        DoTrans
33    End If
    
34    Exit Sub
    
errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    'emailimage
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & _
        ", copyToReport ,line " & Erl
    saveForm (True)
End Sub

'Uploads data from worksheets("Report") to the database with a sql query
'Database Table: dbOEE; Database location: M:\Austausverzeichnis\Edelmann_l
Public Sub DoTrans()
'export from Report ds to dbOEE in M:\Austausverzeichnis\Edelmann_l

Dim cn As Object
Dim dbWb As String
Dim dbWs As String
Dim scn As String
Dim dsh As String
'Dim dbPath As String
Dim target_db As String
Dim ssql As String
Dim varStr As String
Dim cll As Range

Dim errCnt As Integer
1   errCnt = 0

tryAgain:

On Error GoTo catchError
    
2    Set cn = CreateObject("ADODB.Connection")
3    dbWb = Application.ActiveWorkbook.FullName
4    dbWs = "Report" 'Application.ActiveSheet.Name
    
    'Dependent on ACCESS version
    'scn = "PROVIDER=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";" ' ----- > = Access 2007
5    scn = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";" '----- < Access 2007
6    dsh = "[" & Application.Sheets("report").Name & "$]"
7    target_db = "tblOEE_dev"
    
    'Open connection to db
8    cn.Open scn
    
    'CURRENTLY - if a recordset with an identical ID exits, DELETE it
9    If RecExists(Sheets("report").Range("A2").value) = True Then
10        ssql = "DELETE FROM " & target_db & _
                " WHERE [ID] = " & Sheets("report").Range("A2").value & ";"
11        cn.Execute ssql
12    End If
    
13    varStr = ""
14    For Each cll In Worksheets("Report").Range("A2:Y2")
15        varStr = varStr & "'" & cll & "',"
16    Next cll
    
17    varStr = Left(varStr, Len(varStr) - 1)
    
18    ssql = "INSERT INTO " & target_db & " ([ID], [Anlage], [Datum], [Schicht], [OEE], [Eintrag_Zeit], [Benutzer_Name], " & _
             "[Mehrmaschinenbedienung], [Geplante_Stillstaende], [Ruesten], [Material_fehlt], [Personal_fehlt], [Schlosser], " & _
             "[Stoerung], [Materialprobleme], [Avprog_fehlt_ueberarbeiten], [Qualitaetsprobleme], [Zeichnung_unklar], [WOP], " & _
             "[Abweichung_Planzeit], [Allg_Qualitaetsprobleme], " & _
             "[Gutteile], [Ausschußteile], [Laufzeit], [Auftragszeit])" & _
             " VALUES(" & varStr & ")"
             '"SELECT * FROM [Excel 8.0;HDR=YES;DATABASE=" & dbWb & "]." & dsh
    
    'MsgBox ssql
19    cn.Execute ssql
20    MsgBox "Schicht Daten wurden registriert"
    
    'Terminate connection
21    cn.Close
22    Set cn = Nothing
    
23    Exit Sub
    
catchError:
errCnt = errCnt + 1
    Select Case errCnt
        Case Is < 3
            Application.Wait (Now + TimeValue("0:00:05"))
            GoTo tryAgain
        Case Else
            MsgBox "Keine Verbindung zur Datenbank", vbCritical, "Fehler"
            logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & _
                ", DoTrans, line " & Erl
            Exit Sub
    End Select
            

End Sub


'Checks whether a recordset with an ID = findID exits in the db already, return TRUE if it does
Public Function RecExists(findID As Long) As Boolean

Dim cn As Object
Dim rs As Object
Dim strSql As String
Dim strConnection As String
Dim ans As Boolean
'Dim dbPath As String

Dim errCnt As Integer
errCnt = 0
    
tryAgain:

    On Error GoTo catchError
    
    Set cn = CreateObject("ADODB.Connection")
    strConnection = "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & dbPath
                    
    strSql = "SELECT Count(*) FROM tblOEE_dev WHERE [ID] = " & findID & ";"
    
    'Execute query
    cn.Open strConnection
    Set rs = cn.Execute(strSql)
    
    If rs.Fields(0) > 0 Then
        ans = True
    Else
        ans = False
    End If
    
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    
    RecExists = ans

    Exit Function
    
catchError:
errCnt = errCnt + 1
    Select Case errCnt
        Case Is < 3
            GoTo tryAgain
        Case Else
            MsgBox "Keine Verbindung zur Datenbank", vbCritical, "Fehler"
            logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & _
                ", RecExists"
            Exit Function
    End Select
            
End Function

'Build a unique key for the recordset
Public Function createKey(anlageStr As String, datStr As Date, schichtInt As Integer) As Long
Dim anlageInt As Integer
'Dim schichtInt As Integer
Dim mainKey As Long
    
    If globalRepository.staticDict Is Nothing Then
        Set ValueTransfer.staticDictTransfer = New cAnlageDict
        anlageInt = ValueTransfer.staticDictTransfer.KeyDict(anlageStr)
    Else
        anlageInt = globalRepository.staticDict.KeyDict(anlageStr)
    End If

'MsgBox anlageInt & CLng(DateValue(datStr)) & schichtInt
mainKey = CLng(anlageInt & CLng(DateValue(datStr)) & schichtInt)
createKey = mainKey


End Function


