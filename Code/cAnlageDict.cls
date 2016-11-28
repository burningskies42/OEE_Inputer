VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAnlageDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim pKeyDict As Object
Dim pKeyDict_ind As Object

'------------------------------------------------------------------------------------------------------------
'   Class constructor
'------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
Dim key

    Set pKeyDict = CreateObject("Scripting.Dictionary")
    Set pKeyDict_ind = CreateObject("Scripting.Dictionary")

    pKeyDict.RemoveAll
    pKeyDict.Add "5735-2", 1
    pKeyDict.Add "5735-3", 2
    pKeyDict.Add "5735-4", 3
    pKeyDict.Add "5736-2", 4
    pKeyDict.Add "5736-3", 5
    pKeyDict.Add "5736-4", 6
    pKeyDict.Add "6300-1", 7
    pKeyDict.Add "6300-2", 8
    pKeyDict.Add "6350-1", 9
    pKeyDict.Add "6350-2", 10
    pKeyDict.Add "6350-3", 11
    pKeyDict.Add "6350-4", 12
    pKeyDict.Add "6390-1", 13
    pKeyDict.Add "6390-2", 14
    pKeyDict.Add "6410-3", 15
    pKeyDict.Add "6420-1", 16
    pKeyDict.Add "6420-2", 17
    pKeyDict.Add "6430-1", 18
    pKeyDict.Add "6430-2", 19
    pKeyDict.Add "6440", 20
    pKeyDict.Add "6450", 21
    pKeyDict.Add "6460-1", 22
    pKeyDict.Add "6460-2", 23
    pKeyDict.Add "6500-1", 24
    pKeyDict.Add "6500-2", 25
    pKeyDict.Add "6500-3", 26
    pKeyDict.Add "6500-4", 27
    pKeyDict.Add "6510-1", 28
    pKeyDict.Add "6510-2", 29
    pKeyDict.Add "6520", 30
    pKeyDict.Add "6650-1", 31
    pKeyDict.Add "6650-2", 32
    pKeyDict.Add "6720-1", 33
    pKeyDict.Add "6720-2", 34
    pKeyDict.Add "6720-3", 35
    pKeyDict.Add "6720-4", 36
    pKeyDict.Add "7991-1", 37
    pKeyDict.Add "7991-2", 38
    
    'Later addition 17/10/2016
    pKeyDict.Add "5720", 39
    pKeyDict.Add "6740", 40
    pKeyDict.Add "7920", 41
    

    For Each key In pKeyDict.keys()
        pKeyDict_ind.Add pKeyDict_ind.Count + 1, key
    Next key
End Sub

'---------------------------------------------------------------------------------------------------------
' UniqKey properties
'---------------------------------------------------------------------------------------------------------
Public Property Get KeyDict(anlageStr As String) As Integer
    KeyDict = pKeyDict(anlageStr)
End Property

Public Property Get KeyDict_ind(anlageKey As Integer) As String
    KeyDict_ind = pKeyDict_ind(anlageKey)
End Property



