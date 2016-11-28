VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Purpose: main Workbook                                                                          *
'*  Last Change: 14:43 28.11.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

'****************************************************************************************************
'   Problem reporting by no OEE entry
'****************************************************************************************************
Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim ans As Integer
Dim commnets As String

    If Application.ActiveSheet.Name = Worksheets("Ausschuss").Name Then
        'MsgBox "boom"
        Cancel = True
        Exit Sub
    End If
    'Removes saving promt
    ActiveWorkbook.Saved = True
    
    'Case handling upon closing of form
    Select Case ShiftWasEntered
        Case True
            logAction ("Closed")
        Case False
            If MsgBox("Keine Schicht OEE Werte wurden registriert." & vbNewLine & "Wollen sie einen Fehler melden ?", vbYesNo, "Fehlermeldung") = vbYes Then
                
                logAction ("Closed"), (InputBox("Probleme bei Laufzeit:", "Fehlermeldung"))
            Else
                logAction ("Closed")
            End If
    End Select

End Sub

Private Sub Workbook_Open()

    ActiveWindow.WindowState = xlMaximized
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayGridlines = False

    dbPath = ActiveWorkbook.Path & "\OEE_DATABASE\dbOEE_be.mdb"
    Worksheets("OEE").Range("T2") = Date
    logAction ("Opened")
    ShiftWasEntered = True
End Sub


