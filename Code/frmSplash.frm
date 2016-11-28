VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   OleObjectBlob   =   "frmSplash.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Set true when the long task is done.
Public TaskDone As Boolean



'Private Sub UserForm_Activate()
'    Dim frmSplash As CFormChanger
'
'    Set frmSplash = New CFormChanger
'
'    With frmSplash
'        .ShowCloseBtn = False
'        .ShowSysMenu = False
'
'        '.IconPath = Application.ActiveWorkbook.Path & "\Uhlmann_Logo.ico"
'        .setIconFromWS
'        '.ShowIconWS = True
'        Set .Form = Me
'
'        '.Modal = True
'    End With
'End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = Not TaskDone
End Sub
