Attribute VB_Name = "savingSheets"
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Module Purpose: Saving and restoring the entire worksheet                                       *
'*  Last Change: 11:10 14.09.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Global staticDictLoader As cAnlageDict

'--------------------------------------------------------------------------------------------------
' read the entire worksheet into variables
'--------------------------------------------------------------------------------------------------
Public Sub loadSheet(pth As String, Optional withMSO As Boolean = False)

Dim Shift As New cDoneShift
Dim uniqKey As Variant

Dim result As Integer
Dim x As Integer, y As Integer
Dim readRow
Dim arr As Variant
Dim prt As New cDonePart
Dim emptyCell As Range



    clearOEEform
    
'------------------------------------------------------------------------------------------------------------
' Open file dialog - only for .sav file
'------------------------------------------------------------------------------------------------------------
     If withMSO Then
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "Select Test File"
            .Filters.Add "Gespeicherte Bogen", "*.sav"
            .FilterIndex = 1
            .AllowMultiSelect = False
            .InitialFileName = ActiveWorkbook.Path & "\OEE_DATABASE\saves\"
            
            result = .Show
    
            If (result <> 0) Then
                pth = Trim(.SelectedItems.Item(1))
    
            End If
        End With
     End If
'------------------------------------------------------------------------------------------------------------
' read the data from the file into _shift_
'------------------------------------------------------------------------------------------------------------
    If pth = "" Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Shift.readFromFile (pth)
    uniqKey = Shift.extractKey
    arr = Shift.WholeMtrx
    
'------------------------------------------------------------------------------------------------------------
' Write mtrx and commnets to worksheet
'------------------------------------------------------------------------------------------------------------
   
    With Excel.ActiveWorkbook.Worksheets("OEE")
        .Unprotect Password:="aczyM4iu"
        .Range("Anlage") = uniqKey(1)
        .Range("T2") = uniqKey(2)
        .Range("Schicht") = uniqKey(3)
        
        For x = 1 To 49

            readRow = Excel.WorksheetFunction.Transpose( _
                            Excel.WorksheetFunction.Transpose( _
                                Excel.WorksheetFunction.Index(arr, x, 0) _
                            ) _
                        )
            .Range("D8:R56").Rows(x) = readRow
            .Range("D8:R56").Rows(x).Replace What:="0", Replacement:="", MatchCase:=True
            .Range("S8:T56").Cells(x, 1) = Shift.Desc_Mass(1, x)
            .Range("S8:T56").Cells(x, 2) = Shift.Desc_Mass(2, x)

        Next x
    
'------------------------------------------------------------------------------------------------------------
' Write parts to worksheet
'------------------------------------------------------------------------------------------------------------
    
        Set emptyCell = .Range("A61")
        For Each prt In Shift.DoneParts
            While emptyCell.value <> ""
                Set emptyCell = emptyCell.Offset(1, 0)
            Wend
            emptyCell = prt.Nummer
            emptyCell.Offset(0, 1) = prt.Gutteile
            emptyCell.Offset(0, 5) = prt.Ausschusse
            emptyCell.Offset(0, 10) = prt.Stueckzeit
            frmSplash.prgStatus.value = Excel.WorksheetFunction.Min(frmSplash.prgStatus + 5, 100)
        Next prt
        
    End With
    copyToReport readVals, False
    Excel.ActiveWorkbook.Worksheets("OEE").Protect Password:="aczyM4iu"
    'MsgBox "Schicht wurde geladen", vbInformation, "Laden erfolgreich"
    
    ' Close the splash form.
'    loadScreen.TaskDone = True
'    Unload loadScreen
    
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------------------------------------------------------
' Clearform
'------------------------------------------------------------------------------------------------------------
Public Sub clearFormSC()
    Call globalRepository.clearOEEform
End Sub
  
'------------------------------------------------------------------------------------------------------------
' Saves shift data into .sav
'------------------------------------------------------------------------------------------------------------
Public Sub saveForm(Optional withErr As Boolean = False)
1   Dim saveShift As New cDoneShift
2    On Error GoTo errHandler

3   saveShift.readInSheet
4   saveShift.saveDatToFile withErr

5    Exit Sub
    
errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    'emailimage
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & ", SavingSheets"
End Sub

'------------------------------------------------------------------------------------------------------------
'   Returns an array holding shift keys broken to their 3 components
'------------------------------------------------------------------------------------------------------------
Public Function getSavedSheets() As String()
Dim savesList() As String
Dim retKey As Variant
Dim retKeys() As String
Dim savedShift As New cDoneShift
Dim strShift As Variant
Dim shifts As New Collection
Dim staticDict As cAnlageDict

    savesList = GetFileList(ActiveWorkbook.Path & "\OEE_DATABASE\saves\")
    ReDim retKeys(1 To UBound(savesList), 1 To 3)
    
    On Error Resume Next
    For Each strShift In savesList
        savedShift.uniqKey = Left(strShift, Len(strShift) - 4)
        shifts.Add savedShift
        retKey = savedShift.extractKey
        retKeys(shifts.Count, 1) = retKey(1)
        retKeys(shifts.Count, 2) = retKey(2)
        retKeys(shifts.Count, 3) = retKey(3)
        Set savedShift = Nothing
        Set retKey = Nothing
    Next strShift
    
    Erase savesList
    Set shifts = Nothing
    
    getSavedSheets = retKeys()
        
End Function

'------------------------------------------------------------------------------------------------------------
'   Returns an array of filenames that match FileSpec
'   If no matching files are found, it returns False
'-----------------------------------------------------------------------------------------------------------
Private Function GetFileList(FileSpec As String) As Variant

Dim FileArray() As String
Dim FileCount As Integer
Dim FileName As String
    
    On Error GoTo NoFilesFound

    FileCount = 0
    FileName = Dir(FileSpec)
    If FileName = "" Then GoTo NoFilesFound
    
'   Loop until no more matching files are found
    Do While FileName <> ""
        FileCount = FileCount + 1
        ReDim Preserve FileArray(1 To FileCount)
        FileArray(FileCount) = FileName
        FileName = Dir()
    Loop
    GetFileList = FileArray
    Exit Function

'   Error handler
NoFilesFound:
    GetFileList = False
End Function

'------------------------------------------------------------------------------------
'   TEMPORARY - OPEN LOAD FORM
'------------------------------------------------------------------------------------
Public Sub openLoadForm()
Attribute openLoadForm.VB_ProcData.VB_Invoke_Func = "l\n14"
    Set staticDictLoader = New cAnlageDict
    frmOpenSaved.Show
End Sub


