Attribute VB_Name = "ActionLogging"
Option Explicit
Option Private Module

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Module Purpose: logging usage elements into the database                                        *
'*  Last Change: 07:43 16.09.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Dim errCnt As Integer
Global ShiftWasEntered As Boolean

Public Sub logAction(act As String, Optional comments As String)
'log user open and close of inputer

tryAgain:

Dim cn As Object
Dim dbWb As String
Dim dbWs As String
Dim scn As String
Dim dsh As String
Dim dbPath As String
Dim target_db As String
Dim ssql As String
Dim varStr As String
Dim cll As Range

Dim errCnt As Integer
errCnt = 0

Dim userN As String

Select Case Environ$("username")
    Case "edelmann_l"
        userN = "dev"
    Case Else
        userN = Environ$("username")
End Select

    On Error GoTo catchError
    
    Set cn = CreateObject("ADODB.Connection")
    dbPath = ActiveWorkbook.Path & "\OEE_DATABASE\dbOEE_be.mdb"
    dbWb = Application.ActiveWorkbook.FullName
    dbWs = "Report" 'Application.ActiveSheet.Name
    
    'Dependent on ACCESS version
    'scn = "PROVIDER=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";" ' ----- > = Access 2007
    scn = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath & ";" '----- < Access 2007
    dsh = "[" & Application.Sheets("report").Name & "$]"
    target_db = "tblAction_log"
    
    'Open connection to db
    cn.Open scn
    
    ssql = "INSERT INTO " & target_db & " ( [Timestamp], [Action], [User_ID],[PC_Name], [Comments])" & _
             " VALUES('" & Now & "','" & act & "','" & userN & "','" & Environ$("computername") & "','" & comments & "')"
    
    'MsgBox ssql
    cn.Execute ssql
    
    'Terminate connection
    cn.Close
    Set cn = Nothing
    errCnt = 0
    
    Exit Sub
    
catchError:
errCnt = errCnt + 1
    Select Case errCnt
        Case Is < 3
            Application.Wait (Now + TimeValue("0:00:05"))
            GoTo tryAgain
        Case Else
            MsgBox "Keine Verbindung zur Datenbank", vbCritical, "Fehler"
            Exit Sub
    End Select
            

End Sub

Sub emailimage()

Dim OutApp As Object
Dim OutMail As Object

    'Shift-Print Screen
    Application.SendKeys "(%{1068})"
    
    On Error Resume Next
    
    'Prepare the email
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.Session.Logon
    Set OutMail = OutApp.CreateItem(0)
    
    On Error Resume Next
    
    With OutMail
    .To = getTroubleshooter
    .Subject = "OEE Inputer Fehler bei " & Environ$("username")
    .display
    
    Application.SendKeys "(^v)"
    
    End With
    On Error GoTo 0
    
    OutApp.Session.Logoff
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub

Sub emailMe()
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object
    Dim strBody As String
    

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

'    strbody = "Hi there" & vbNewLine & vbNewLine & _
'              "This is line 1" & vbNewLine & _
'              "This is line 2" & vbNewLine & _
'              "This is line 3" & vbNewLine & _
'              "This is line 4"

    On Error Resume Next
    With OutMail
        .To = getTroubleshooter
        .CC = ""
        .BCC = ""
        .Subject = "OEE Inputer Fehler bei " & Environ$("username")
        .Body = strBody

        .display
        
        
    End With
    
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

