Attribute VB_Name = "pdfSend"
Option Explicit
Option Private Module

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*   CURRENTLY NOT IN USE                                                                           *
'*   Module Purpose: Defintion of global variables                                                  *
'*   Last Change: 08:12 09.08.2016                                                                  *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub AttachActiveSheetPDF(targetEmail As String)
Dim IsCreated As Boolean
Dim i As Long
Dim PdfName As String, pdfPath As String
Dim PdfFile As String, Title As String
Dim OutlApp As Object
 
    ' Not sure for what the Title is
    Title = "Ausschussbericht von " & Application.UserName & ", Auftragsnr. " & Range("C6")
 
    ' Define PDF filename
    pdfPath = ActiveWorkbook.Path & "\Ausschuss_Berichte\"
    'i = InStrRev(PdfFile, ".")
    'If i > 1 Then PdfFile = Left(PdfFile, i - 1)
    
    PdfName = Format(Date, "DDMMYYYY") & "_Ausschuss_Auftr_" & Range("C6") & ".pdf"
    PdfFile = pdfPath & PdfName
    
    
    If Dir(pdfPath, vbDirectory) = "" Then
        Shell ("cmd /c mkdir """ & pdfPath & """")
    End If
    
    ' Export activesheet as PDF
    With ActiveSheet
      .ExportAsFixedFormat Type:=xlTypePDF, FileName:=PdfFile, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End With
 
    ' Use already open Outlook if possible
    On Error Resume Next
    Set OutlApp = GetObject(, "Outlook.Application")
    If Err Then
      Set OutlApp = CreateObject("Outlook.Application")
      IsCreated = True
    End If
    OutlApp.Visible = True
    On Error GoTo 0
 
    ' Prepare e-mail with PDF attachment
    With OutlApp.CreateItem(0)
   
    ' Prepare e-mail
    .Subject = "Ausschussbericht von " & Application.UserName & ", Auftragsnr. " & Range("C6")
    .To = targetEmail ' <-- Put email of the recipient here
    '.CC = "..." ' <-- Put email of 'copy to' recipient here
    .Body = "Dieser E-mail wurde automatisch generiert." & vbLf & vbLf _
          & "Der Auschussbericht ist in PDF Format beigefügt." & vbLf & vbLf _
          & vbLf _
          & Application.UserName & vbLf & vbLf
    .Attachments.Add PdfFile
   
    ' Try to send
    On Error Resume Next
    .Send
    Application.Visible = True
    
    If Err Then
        MsgBox "Beim E-mail Versand ist ein Fehler aufgetreten", vbExclamation, "Fehler"
    Else
        MsgBox "E-mail wurde erfolgreich abgesendet", vbInformation, "Erfolg"
        ' Delete PDF file
        ' Kill PdfFile
    End If
    On Error GoTo 0
   
  End With
 
  ' Delete PDF file
  ' Kill PdfFile
 
  ' Quit Outlook if it was created by this code
  If IsCreated Then OutlApp.Quit
 
  ' Release the memory of object variable
  Set OutlApp = Nothing
 
End Sub

Sub Mail_small_Text_Outlook(strBody As String, strAdd As String, strSbj As String)
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2016
    Dim OutApp As Object
    Dim OutMail As Object
    
    OpenOutlook

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        '.To = getTroubleshooter
        .To = strAdd
        
        '.Subject = "This is the Subject line"
        .Subject = strSbj
        
        .Body = strBody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        .display True
        .Send   'or use .Display
    End With
    On Error GoTo 0
    
    OutApp
    
    Set OutMail = Nothing
    OutApp.Quit
    Set OutApp = Nothing
End Sub

Sub OpenOutlook()
    Dim oOutlook As Object
    Dim OutlookApp As Object
    Dim ret As Long

    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If oOutlook Is Nothing Then
        ret = ShellExecute(Application.hwnd, vbNullString, "Outlook", vbNullString, "C:\", 1)
    End If
    
End Sub



