VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBearb 
   Caption         =   "Manuelle Bearbeitung"
   ClientHeight    =   930
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   2415
   OleObjectBlob   =   "frmBearb.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmBearb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Completion on manual changes to done OEE form                                     *
'*  Last Change: 13:22 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Private Sub UserForm_Activate()
    Dim frmBearb As CFormChanger
    
    Set frmBearb = New CFormChanger
    
    With frmBearb
        .ShowCloseBtn = False
        '.ShowSysMenu = False
        
        '.IconPath = Application.ActiveWorkbook.Path & "\Uhlmann_Logo.ico"
        .setIconFromWS
        '.ShowIconWS = True
        Set .Form = Me
        
        '.Modal = True
    End With
    Worksheets("OEE").Unprotect Password:="aczyM4iu"
End Sub


Private Sub bearbFertBtn_Click()
    frmBearb.Hide
    endEntry
End Sub
