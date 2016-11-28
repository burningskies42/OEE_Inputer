VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSettingsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTSemail As String
Private pAusschussEmail As String

'--------------------------------------------------------------------------------------------------------
'   Constructor
'--------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    readinSettings
End Sub

'--------------------------------------------------------------------------------------------------------
'   Troubleshooter Email - property
'--------------------------------------------------------------------------------------------------------
Public Property Get tsEmail() As String
    tsEmail = pTSemail
End Property
Public Property Let tsEmail(val As String)
    pTSemail = val
    writeSettings
End Property

'--------------------------------------------------------------------------------------------------------
'   Ausschuss Email address - property
'--------------------------------------------------------------------------------------------------------
Public Property Get AusschussEmail() As String
    AusschussEmail = pAusschussEmail
End Property
Public Property Let AusschussEmail(val As String)
    pAusschussEmail = val
    writeSettings
End Property

'--------------------------------------------------------------------------------------------------------
'   Extract current settings from the ssettings file
'--------------------------------------------------------------------------------------------------------
Private Sub readinSettings()
Dim settFile As String
Dim textline As String

    settFile = ThisWorkbook.Path & "\Einstellungen\Einstellungen.dat"
    
    Open settFile For Input As #1

    Do Until EOF(1)
        Line Input #1, textline
        
        Select Case whichSetting(textline)
            Case "ts"
                pTSemail = Right(textline, Len(textline) - Len("troubleShooter:= "))
            Case "aus_email"
                pAusschussEmail = Right(textline, Len(textline) - Len("Ausschuss_Email:= "))
        End Select
        
    Loop
    
    Close #1
        
End Sub

'--------------------------------------------------------------------------------------------------------
'   Write current settings to the settings file
'--------------------------------------------------------------------------------------------------------
Private Sub writeSettings()
Dim settFile As String
Dim textline As String

    settFile = ThisWorkbook.Path & "\Einstellungen\Einstellungen.dat"
    
    Open settFile For Output As #1
    
    Print #1, "troubleShooter:= " & Me.tsEmail
    Print #1, "Ausschuss_Email:= " & Me.AusschussEmail
    
    Close #1
    
End Sub

'--------------------------------------------------------------------------------------------------------
'   Identifies which setting was read from the settings file
'--------------------------------------------------------------------------------------------------------
Private Function whichSetting(str As String) As String
Dim ret As Boolean
    
    ' Checks if the curent line is the troubleshooter Setting
    If str Like "troubleShooter:=*" Then
        whichSetting = "ts"
    ElseIf str Like "Ausschuss_Email:=*" Then
        whichSetting = "aus_email"
    End If
    
End Function
