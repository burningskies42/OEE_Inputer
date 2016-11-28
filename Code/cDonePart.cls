VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDonePart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pNummer As String
Private pGutteile As Integer
Private pAusschusse As Integer
Private pStueckzeit As Double

'----------------------------------------------------------------------------------------
' Nummer properties
'----------------------------------------------------------------------------------------
Public Property Get Nummer() As String
    Nummer = pNummer
End Property

Public Property Let Nummer(value As String)
    pNummer = value
End Property

'----------------------------------------------------------------------------------------
' Gutteile properties
'----------------------------------------------------------------------------------------
Public Property Get Gutteile() As Integer
    Gutteile = pGutteile
End Property

Public Property Let Gutteile(value As Integer)
    pGutteile = value
End Property

'----------------------------------------------------------------------------------------
' Ausschusse properties
'----------------------------------------------------------------------------------------
Public Property Get Ausschusse() As Integer
    Ausschusse = pAusschusse
End Property

Public Property Let Ausschusse(value As Integer)
    pAusschusse = value
End Property

'----------------------------------------------------------------------------------------
' Stueckzeit properties
'----------------------------------------------------------------------------------------
Public Property Get Stueckzeit() As Double
    Stueckzeit = pStueckzeit
End Property

Public Property Let Stueckzeit(value As Double)
    pStueckzeit = value
End Property
