VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDoneShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pUniqKey As Long

Private pMtrx() As Integer
Private pDesc_Mass() As String

Private pDoneParts As New Collection

'---------------------------------------------------------------------------------------------------------
' Class constructor
'---------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
Dim key
    ReDim pMtrx(1 To 49, 1 To 15)
    ReDim pDesc_Mass(1 To 2, 1 To 49)
 
End Sub

'---------------------------------------------------------------------------------------------------------
' UniqKey properties
'---------------------------------------------------------------------------------------------------------
Public Property Get uniqKey() As Long
    uniqKey = pUniqKey
End Property
Public Property Let uniqKey(value As Long)
    pUniqKey = value
End Property

'---------------------------------------------------------------------------------------------------------
' Matrix properties
'---------------------------------------------------------------------------------------------------------
Public Property Get mtrx(indX As Integer, indY As Integer) As Integer
    mtrx = pMtrx(indX, indY)
End Property
Public Property Get WholeMtrx() As Integer()
    WholeMtrx = pMtrx()
End Property

Public Property Let mtrx(indX As Integer, indY As Integer, value As Integer)
    pMtrx(indX, indY) = value
End Property

'---------------------------------------------------------------------------------------------------------
' Desc_Mass properties
'---------------------------------------------------------------------------------------------------------
Public Property Get Desc_Mass(indX As Integer, indY As Integer) As String
    Desc_Mass = pDesc_Mass(indX, indY)
End Property
Public Property Let Desc_Mass(indX As Integer, indY As Integer, value As String)
    pDesc_Mass(indX, indY) = value
End Property

'---------------------------------------------------------------------------------------------------------
' DoneParts properties
'---------------------------------------------------------------------------------------------------------
Public Property Get DoneParts() As Collection
    Set DoneParts = pDoneParts
End Property
Public Property Let DoneParts(value As Collection)
    Set pDoneParts = value
End Property

'---------------------------------------------------------------------------------------------------------
' Read-in Function: reads in all the data on the OEE form
'---------------------------------------------------------------------------------------------------------
Public Function readInSheet()
Dim rng As Range
Dim x As Integer
Dim y As Integer

Dim cell As Range
Dim pPart As New cDonePart
    
    With Excel.ActiveWorkbook.Worksheets("OEE")
        Me.uniqKey = createKey(.Range("Anlage"), .Range("T2"), .Range("Schicht"))

    End With
    
    Set rng = Excel.ActiveWorkbook.Worksheets("OEE").Range("$D$8:$T$56")
    
    'Go through all cells and fill in mtrx
    For x = 1 To 49
        For y = 1 To 15
            Me.mtrx(x, y) = rng.Cells(x, y)
        Next y
        
        ' read in Descriptions
        Me.Desc_Mass(1, x) = rng(x, 16)
        
        'read in measures
        Me.Desc_Mass(2, x) = rng(x, 17)
    Next x
    
    ' read in all done parts
    For Each cell In Excel.ActiveWorkbook.Worksheets("OEE").Range("A61:A74")
        If cell.value <> "" Then
            pPart.Nummer = CStr(cell.value)
            pPart.Gutteile = cell.Offset(0, 1).value
            pPart.Ausschusse = cell.Offset(0, 5).value
            pPart.Stueckzeit = cell.Offset(0, 10).value
            pDoneParts.Add pPart
            Set pPart = Nothing
        End If
    Next cell

End Function

'---------------------------------------------------------------------------------------------------------
' Save to file Function: saves all the sheet data to a .sav file
'---------------------------------------------------------------------------------------------------------
Public Sub saveDatToFile(Optional withError As Boolean = False)
Dim pth As String
Dim errPth As String
Dim x As Integer, y As Integer

Dim fs As Object

    On Error GoTo errHandler
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    errPth = ActiveWorkbook.Path & "\OEE_DATABASE\saves\errors\" & Environ$("username") & _
             "_" & Date & ".csv"
    pth = ActiveWorkbook.Path & "\OEE_DATABASE\saves\" & Me.uniqKey & ".sav"
    
    ' Open save file for writing
    Open pth For Output As #1
        
        ' Write in unique key
        Print #1, "UniqKey: " & Me.uniqKey
        Print #1, ""
        
        ' Write in movement matrix
        Print #1, "moveMatrix:"
        For x = 1 To 49
            For y = 1 To 15
                Select Case y
                    Case 15
                        Write #1, Me.mtrx(x, y)
                    Case Else
                        ' end of row
                        Write #1, Me.mtrx(x, y),
                End Select
            Next y
        Next x
        Print #1, ""
        
        ' Write in descriptions and measures taken
        Print #1, "desc_mass:"
        For x = 1 To 49
            Write #1, Me.Desc_Mass(1, x), Me.Desc_Mass(2, x)
        Next x
        
        Print #1, ""
        
        ' Write in done parts
        Print #1, "DoneParts:"
        
        With Me.DoneParts
            For x = 1 To .Count
                With .Item(x)
                    Write #1, .Nummer, .Gutteile, .Ausschusse, .Stueckzeit
                End With
            Next x
        End With
        
    Close #1
    
    If withError Then
        fs.copyfile pth, errPth
    End If
    
    Exit Sub
    
errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & ", SaveDatToFile"
    
End Sub

Public Sub readFromFile(pth As String)

Dim DataLine As String
Dim x As Integer
Dim currRow As Variant
Dim y As Integer
Dim newPart As New cDonePart

    Open pth For Input As #1
    
    While Not EOF(1)
        Line Input #1, DataLine
        
        If InStr(1, DataLine, "UniqKey:") > 0 Then
            Me.uniqKey = Right(DataLine, Len(DataLine) - Len("UniqKey: "))
        ElseIf InStr(1, DataLine, "moveMatrix:") > 0 Then
            For x = 1 To 49
            
                Line Input #1, DataLine
                currRow = Split(DataLine, ",")
                For y = 1 To 15
                    Me.mtrx(x, y) = currRow(y - 1)
                Next y
            Next x
        ElseIf InStr(1, DataLine, "desc_mass:") > 0 Then
            
            For x = 1 To 49
                Line Input #1, DataLine
                DataLine = Replace(DataLine, Chr(34), "")
                Me.Desc_Mass(1, x) = Left(DataLine, InStr(1, DataLine, ",") - 1)
                Me.Desc_Mass(2, x) = Right(DataLine, InStr(1, StrReverse(DataLine), ",") - 1)
            Next x
        ElseIf InStr(1, DataLine, "DoneParts:") > 0 Then
            x = 1
            While Not EOF(1)
                Line Input #1, DataLine
                DataLine = Replace(DataLine, Chr(34), "")
                newPart.Nummer = Left(DataLine, InStr(1, DataLine, ",") - 1)
                DataLine = Right(DataLine, Len(DataLine) - InStr(1, DataLine, ","))
                newPart.Gutteile = Left(DataLine, InStr(1, DataLine, ",") - 1)
                DataLine = Right(DataLine, Len(DataLine) - InStr(1, DataLine, ","))
                newPart.Ausschusse = Left(DataLine, InStr(1, DataLine, ",") - 1)
                DataLine = Right(DataLine, Len(DataLine) - InStr(1, DataLine, ","))
                newPart.Stueckzeit = CDbl(Replace(DataLine, ".", ","))
                Me.DoneParts.Add newPart
                Set newPart = Nothing
            Wend
        End If
        
    Wend
    
    Close #1
End Sub

'---------------------------------------------------------------------------------------------------------
' Build a unique key for the recordset
'---------------------------------------------------------------------------------------------------------
Public Function createKey(anlageStr As String, datStr As Date, schichtInt As Integer) As Long

Dim anlageInt As Integer
Dim mainKey As Long

'Anlage conversion
    If savingSheets.staticDictLoader Is Nothing Then
        Set savingSheets.staticDictLoader = New cAnlageDict
    End If
    anlageInt = savingSheets.staticDictLoader.KeyDict(anlageStr)

    mainKey = CLng(anlageInt & CLng(DateValue(datStr)) & schichtInt)
    createKey = mainKey


End Function

Public Function extractKey() As String()
Dim strKey As String
Dim str(1 To 3) As String
    
    strKey = CStr(Me.uniqKey)
    
    str(3) = Right(strKey, 1)
    strKey = Left(strKey, Len(strKey) - 1)
    
    str(2) = Format(Right(strKey, 5), "DD.MM.YYYY")
    strKey = CInt(Left(strKey, Len(strKey) - 5))
    
    str(1) = savingSheets.staticDictLoader.KeyDict_ind(CInt(strKey))
    
    
    extractKey = str()

End Function


