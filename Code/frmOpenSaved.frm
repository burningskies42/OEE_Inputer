VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpenSaved 
   Caption         =   "Gespeicherte Schichte"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8700
   OleObjectBlob   =   "frmOpenSaved.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmOpenSaved"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim allKeys() As String
Dim uniqKey As Variant

'--------------------------------------------------------------------------------------------------------------------
'   Upon form open - populate machines listbox
'--------------------------------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
Dim j As Integer, i As Integer
Dim keyExists As Boolean
Dim temp As String

    allKeys = getSavedSheets
    
    For Each uniqKey In Excel.WorksheetFunction.Transpose(Excel.WorksheetFunction.Transpose(Excel.WorksheetFunction.Index(allKeys, 0, 1)))
        keyExists = False
        For j = 0 To lbMachines.ListCount - 1
            If lbMachines.Column(0, j) = uniqKey Then
                keyExists = True
                'Exit For
            End If
        Next j
        If Not keyExists And Len(uniqKey) > 0 Then
            lbMachines.AddItem uniqKey
        End If
    Next uniqKey
    
    With Me.lbMachines
    For j = 0 To .ListCount - 2
             For i = 0 To .ListCount - 2
                 If .List(i) > .List(i + 1) Then
                     temp = .List(i)
                     .List(i) = .List(i + 1)
                     .List(i + 1) = temp
                 End If
             Next i
         Next j
    End With
    
End Sub

'--------------------------------------------------------------------------------------------------------------------
'   First Listbox - Machines - populates listbox dates
'--------------------------------------------------------------------------------------------------------------------
Private Sub lbMachines_Click()
Dim i As Integer
Dim j As Integer
Dim temp As Date
Dim keyExists As Boolean

    lbDates.Clear
    lbShifts.Clear
    
    For i = 1 To UBound(allKeys)
        If lbMachines.value = allKeys(i, 1) Then
            keyExists = False
            
            For j = 0 To lbDates.ListCount - 1
                If allKeys(i, 2) = lbDates.Column(0, j) Then
                    keyExists = True
                End If
            Next j
            If Not keyExists Then
                lbDates.AddItem Format(DateValue(allKeys(i, 2)), "DD.MM.YYYY")
            End If
        End If
    Next i
    
    With Me.lbDates
    For j = 0 To .ListCount - 2
             For i = 0 To .ListCount - 2
                 If DateValue(.List(i)) < DateValue(.List(i + 1)) Then
                     temp = .List(i)
                     .List(i) = .List(i + 1)
                     .List(i + 1) = Format(temp, "DD.MM.YYYY")
                 End If
             Next i
         Next j
    End With
    
End Sub

'--------------------------------------------------------------------------------------------------------------------
'   Second Listbox - Dates - populates listbox shifts
'--------------------------------------------------------------------------------------------------------------------
Private Sub lbDates_Click()
Dim i As Integer

    lbShifts.Clear
    
    For i = 1 To UBound(allKeys)
        If lbMachines.value = allKeys(i, 1) And lbDates.value = allKeys(i, 2) Then
            Select Case allKeys(i, 3)
                Case 1
                    lbShifts.AddItem "Früh"
                Case 2
                    lbShifts.AddItem "Spät"
                Case 3
                    lbShifts.AddItem "Nacht"
            End Select
        End If
    Next i

End Sub

'--------------------------------------------------------------------------------------------------------------------
'   third Listbox - Shifts - opens selected shift
'--------------------------------------------------------------------------------------------------------------------
Private Sub lbShifts_Click()

Dim selectedShift As New cDoneShift
Dim shiftInt As Integer
Dim uniqKey As Long

    
    frmOpenSaved.Hide
    
    Select Case lbShifts.value
        Case "Früh"
            shiftInt = 1
        Case "Spät"
            shiftInt = 2
        Case "Nacht"
            shiftInt = 3
    End Select
    
    ' Disable keyboard
'    Application.OnKey "^d", "KeyboardOn"
'    Application.DataEntryMode = True
    
    
    uniqKey = selectedShift.createKey(lbMachines.value, lbDates.value, shiftInt)
    
    'Load shift
    selectedShift.readFromFile (ActiveWorkbook.Path & "\OEE_DATABASE\saves\" & uniqKey & ".sav")
    
    frmOpenSaved.Hide
    loadSheet (ActiveWorkbook.Path & "\OEE_DATABASE\saves\" & uniqKey & ".sav")
    
    Unload frmOpenSaved
    
'    ' Enable keyboard
'    Application.DataEntryMode = False
    
End Sub

