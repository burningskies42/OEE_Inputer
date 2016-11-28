VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    frmAusschuss.Show vbModeless
    setDefaultVals
End Sub

Private Sub Worksheet_Deactivate()
    frmAusschuss.Hide
End Sub

Private Sub setDefaultVals()
Const str As String = "Bitte Wählen"
    Range("C1") = "Allgemeiner Fehler"
    Range("C5") = Format(Worksheets("OEE").Range("T2"), "DD.MM.YYYY")
    Range("C6") = ""
    Range("F5:F6") = ""
    Range("I5") = Worksheets("OEE").Range("Anlage")

    Range("E8") = ""
    Range("E12") = ""
    Range("G27") = ""
    Range("F28:F31") = ""
    Range("C31") = ""
    Range("I31") = ""
    Range("F33:F34") = ""
    Range("C34") = ""
    Range("B39") = ""
    Range("B42") = ""
    
    Range("C8") = str
    Range("C10") = str
    Range("H10") = str
    Range("E11") = str
    Range("F14:F16") = str
    Range("G18:G25") = str
    Range("G19") = ""
    Range("G23") = ""
    Range("D27") = str
    Range("D36") = str
    Range("F45:F46") = str
    
End Sub
