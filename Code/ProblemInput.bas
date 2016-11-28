Attribute VB_Name = "ProblemInput"
Option Explicit
Option Private Module

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Module Purpose: Functions to input the different malfunctions - code for Fehl* forms            *
'*  Last Change: 13:11 29.06.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------


Public innerString As String

Public Function probInput(rowCol As Integer) As String

Dim Report As Worksheet 'Set up worksheet variable.
Set Report = Excel.ActiveSheet 'Assign the active sheet to the variable.
        
On Error GoTo errHandler

Dim err_desc As String
innerString = ""
    Select Case rowCol
        Case 5 ' Mehrmaschinenbedienung - input using Function[inptMehrmasch].module[ProblemInput]
            err_desc = inptMehrmasch
        
        Case 6 ' Geplante Stillstand - input using Function[inptGeplanteStillstand].module[ProblemInput]
            err_desc = inptGeplanteStillstand
        
        Case 7 ' Rüsten - input using Function[inptRuesten].module[ProblemInput]
            err_desc = inptRuesten
            
        Case 8 ' Material fehlt - input using Function[inptMaterialFehlt].module[ProblemInput]
            err_desc = inptMaterialFehlt
        
        ' Schlossertätigkeiten - used to be Maschinenstörung
        Case 10
            'err_desc = inptMaschStör - used for Maschinenstörung
            err_desc = "Schlossertätigkeit"
            
        Case 11 'Werkzeugstörung - input using Function[inptWerkzeug].module[ProblemInput]
            err_desc = inptWerkzeug
            
        Case 12 'Materialprobleme - input using Function[inptMaterProb].module[ProblemInput]
            err_desc = inptMaterProb
            
        Case 13 'Qualitätsprobleme - input using Function[inptQual].module[ProblemInput]
            err_desc = inptQual
            
        Case 14 'Zeichnung unklar - input using Function[inptAVProg].module[ProblemInput]
            err_desc = inptZeichUnklar
            
        Case 15 'AVProg fehlt - input using Function[inptAVProg].module[ProblemInput]
            err_desc = inptAVProg
            
        Case 16 'WOP - input using Function[inptWOP].module[ProblemInput]
            err_desc = inptWOP
                    
        Case 17  ' Allg. Abweichung - input using a simple inputbox
            err_desc = Report.Cells(7, rowCol) & ": " & _
            InputBox("Bitte beschreiben Sie die " & Report.Cells(7, rowCol), Report.Cells(7, rowCol))
              
    End Select
    
    If Sheets("OEE").ProtectContents Then
        Sheets("OEE").Unprotect Password:="aczyM4iu"
    End If

    probInput = err_desc
    

Exit Function
errHandler:
    MsgBox "Ein Fehler ist aufgetreten." & vbNewLine & "Ein Fehlerbericht wird jetzt generiert.", vbCritical, "Fehler"
    'emailimage
    logAction "Error", "num: " & Err.Number & ", desc: " & Err.Description & ", src: " & Err.Source & _
        ", probInput"

End Function

' Input frunctions for the different malfunction types

' Input malfunction of type Maschinenstörungen
Public Function inptMaschStör() As String

    'Show the form
    frmFehlMaschinen.Show
    
    'function returns the module-level variable innerString
    inptMaschStör = innerString

End Function


' Input malfunction of type Geplante Stillstand
Public Function inptMehrmasch() As String

    'Show the form
    frmFehlMehrmasch.Show
    
    'function returns the module-level variable innerString
    inptMehrmasch = innerString

End Function

' Input malfunction of type Geplante Stillstand
Public Function inptGeplanteStillstand() As String

    'Show the form
    frmFehlGeplanteStillstand.Show
    
    'function returns the module-level variable innerString
    inptGeplanteStillstand = innerString

End Function

' Input malfunction of type Werkzeugstörung
Public Function inptWerkzeug() As String

    'Show the form
    frmFehlWerkzeug.Show
    
    'function returns the module-level variable innerString
    inptWerkzeug = innerString

End Function

' Input malfunction of type Werkzeugstörung
Public Function inptMaterProb() As String

    'Show the form
    frmFehlMaterial.Show
    
    'function returns the module-level variable innerString
    inptMaterProb = innerString

End Function

' Input malfunction of type AVProg
Public Function inptAVProg() As String

    'Show the form
    frmFehlAVProg.Show
    
    'function returns the module-level variable innerString
    inptAVProg = innerString

End Function

' Input malfunction of type WOP
Public Function inptWOP() As String

    'Show the form
    frmFehlWOP.Show
    
    'function returns the module-level variable innerString
    inptWOP = innerString

End Function

' Input malfunction of type Rüsten
Public Function inptRuesten() As String

    'Show the form
    frmFehlRuesten.Show
    
    'function returns the module-level variable innerString
    inptRuesten = innerString

End Function


' Input malfunction of type Material fehlt
Public Function inptMaterialFehlt() As String

    'Show the form
    frmFehlMaterialFehlt.Show
    
    'function returns the module-level variable innerString
    inptMaterialFehlt = innerString

End Function

' Input malfunction of type Material fehlt
Public Function inptZeichUnklar() As String

    'Show the form
    frmFehlZeichnungUnklar.Show
    
    'function returns the module-level variable innerString
    inptZeichUnklar = innerString

End Function

' Input malfunction of type Qualitätsprobleme
Public Function inptQual() As String

    'Show the form
    frmFehlQual.Show
    
    'function returns the module-level variable innerString
    inptQual = innerString

End Function








