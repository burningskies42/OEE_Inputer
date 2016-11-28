VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAusschuss 
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2415
   OleObjectBlob   =   "frmAusschuss.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmAusschuss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Confirming completion of Ausschuss form                                           *
'*  Last Change: 13:22 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Private Sub UserForm_Activate()

    Dim frmAusschuss As CFormChanger
    Set frmAusschuss = New CFormChanger
    
    With frmAusschuss
        .ShowCloseBtn = False
        '.ShowSysMenu = False
        
        '.IconPath = Application.ActiveWorkbook.Path & "\Uhlmann_Logo.ico"
        .setIconFromWS
        '.ShowIconWS = True
        Set .Form = Me
        
        '.Modal = True
    End With

End Sub

Private Sub btnFertig_Click()
Dim doPrint As Variant

    ' Check that there is indeed a "Auftragsnummer"
    If Worksheets("Ausschuss").Range("c6") = "" Then
        ' Dont allow completion if the number is missing
        MsgBox "Bitte geben Sie ein Auftragsnummer ein", vbCritical, "Auftragsnummer Fehlt"
        Exit Sub
    End If

    doPrint = MsgBox("Wollen Sie den Formular drucken ?", vbYesNoCancel, "Ausschuﬂformular")
    
    Select Case doPrint
        Case vbYes
                With Worksheets("Ausschuss")
                    
                    .PageSetup.BlackAndWhite = False
                    .PageSetup.LeftFooter = Environ(Application.UserName)
                    .Range("Ausschuﬂ_Print_Area").PrintOut ActivePrinter:="Adobe PDF"
            
                MsgBox "Printing ..."
            End With
            
'            With Application.Workbooks(1)
'                .Worksheets(1).Visible = True
'                .Worksheets(2).Visible = False
'                .Worksheets(1).Activate
'            End With
            
            'frmAusschuss.Hide
            'frmMove.Show
            
        Case vbNo
'            With Application.Workbooks(1)
'                .Worksheets(1).Visible = True
'                .Worksheets(2).Visible = False
'                .Worksheets(1).Activate
'            End With
            
            'frmAusschuss.Hide
            'frmMove.Show
        Case vbCancel
            Exit Sub
        
    End Select
    
        ' Promt whether the form should be send automatically
    Select Case MsgBox("Wollen Sie den Auschussbericht absenden ?", vbQuestion + vbYesNoCancel, "Ausschuss Melden")
        Case vbYes
            AttachActiveSheetPDF (getAusschussEmail)

                With Application.Workbooks(1)
                    .Worksheets(1).Visible = True
                    .Worksheets(2).Visible = False
                    .Worksheets(1).Activate
                    findLastCell
                End With
        Case vbNo

                With Application.Workbooks(1)
                    .Worksheets(1).Visible = True
                    .Worksheets(2).Visible = False
                    .Worksheets(1).Activate
                    findLastCell
                End With
            
        Case vbCancel
            Exit Sub
    End Select
    
    frmAusschuss.Hide
    frmMove.Show
            
    
End Sub
