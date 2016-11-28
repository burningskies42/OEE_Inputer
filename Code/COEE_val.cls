VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COEE_val"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Identifiers
Private pDatum As Date
Private pAnlage As String
Private pSchicht As String

'Problems
Private pMehrmaschinenbedienung As Integer '---------- 1
Private pGeplante_Stillstaende As Integer '----------- 2
Private pRuesten As Integer '------------------------- 3
Private pMaterial_fehlt As Integer '------------------ 4
Private pPersonal_fehlt As Integer '------------------ 5
Private pSchlosser As Integer '--------------- 6
Private pStoerung As Integer '---------------- 7
Private pMaterialprobleme As Integer '---------------- 8
Private pQualitaetsprobleme As Integer '-------------- 9
Private pZeichnung_unklar As Integer '--------------- 10
Private pAvprog_fehlt_ueberarbeiten As Integer '----- 11
Private pWOP As Integer '---------------------------- 12
Private pAbweichung_Planzeit As Integer '------------ 13
Private pAllg_Qualitaetsprobleme As Integer '-------- 14

'Laufzeit
Private pLaufZeit As Integer

'Betriebzeit
Private pBetriebZeit As Double
'Stillstandzeit
Private pStillZeit As Double

'Auftragzeit
Private pAuftragzeit As Double

'Gutteile
Private pGutteile As Integer

'Ausschuse
Private pAusschuss As Integer

'OEE value
Private pOEE As Double

Private pVal_d As Double
Private pVal_e As Double
Private pVal_f As Double

'----------------------------------------------------------------------------
' Identifiers
'----------------------------------------------------------------------------
'Anlage properties
Public Property Get Anlage() As String
    Anlage = pAnlage
End Property

Public Property Let Anlage(value As String)
    pAnlage = value
End Property

'Datum properties
Public Property Get Datum() As Date
    Datum = pDatum
End Property

Public Property Let Datum(value As Date)
    pDatum = value
End Property

'Schicht properties
Public Property Get Schicht() As String
    Schicht = pSchicht
End Property

Public Property Let Schicht(value As String)
    pSchicht = value
End Property

'----------------------------------------------------------------------------
' OEE properties
'----------------------------------------------------------------------------

'calculate OEE value
Public Property Get OEE() As Double
    pOEE = pVal_d * pVal_e * pVal_f
    OEE = pOEE
End Property

'LaufZeit properties
Public Property Get Laufzeit() As Double
    Laufzeit = pLaufZeit
End Property

Public Property Let Laufzeit(value As Double)
    pLaufZeit = value
End Property

'BetriebZeit properties
Public Property Get BetriebZeit() As Double
     BetriebZeit = 480
End Property

'StillZeit properties
Public Property Get StillZeit() As Double
     StillZeit = pStillZeit
End Property

Public Property Let StillZeit(value As Double)
     pStillZeit = value
End Property

'Auftragzeit properties
Public Property Get Auftragzeit() As Double
     Auftragzeit = pAuftragzeit
End Property

Public Property Let Auftragzeit(value As Double)
     pAuftragzeit = value
End Property

'Gutteile properties
Public Property Get Gutteile() As Integer
     Gutteile = pGutteile
End Property

Public Property Let Gutteile(value As Integer)
     pGutteile = value
End Property

'Ausschusse properties
Public Property Get Ausschuss() As Integer
     Ausschuss = pAusschuss
End Property

Public Property Let Ausschuss(value As Integer)
     pAusschuss = value
End Property

'val_d properties
Public Property Get val_d() As Double
    val_d = pVal_d
End Property

Public Property Let val_d(value As Double)
     pVal_d = value
End Property

'val_e properties
Public Property Get val_e() As Double
    val_e = pVal_e
End Property

Public Property Let val_e(value As Double)
     pVal_e = value
End Property

'val_f properties
Public Property Get val_f() As Double
    val_f = pVal_f
End Property

Public Property Let val_f(value As Double)
     pVal_f = value
End Property

'------------------------------------------------------------------------------------------------------------------------------------------------------------
' Problem duration
'------------------------------------------------------------------------------------------------------------------------------------------------------------


'Mehrmaschinenbedienung properties - 1
Public Property Get Mehrmaschinenbedienung() As Integer
    Mehrmaschinenbedienung = pMehrmaschinenbedienung
End Property

Public Property Let Mehrmaschinenbedienung(value As Integer)
     pMehrmaschinenbedienung = value
End Property

'Geplante_Stillstaende properties - 2
Public Property Get Geplante_Stillstaende() As Integer
    Geplante_Stillstaende = pGeplante_Stillstaende
End Property

Public Property Let Geplante_Stillstaende(value As Integer)
     pGeplante_Stillstaende = value
End Property

'Ruesten properties - 3
Public Property Get Ruesten() As Integer
    Ruesten = pRuesten
End Property

Public Property Let Ruesten(value As Integer)
     pRuesten = value
End Property

'Material_fehlt properties - 4
Public Property Get Material_fehlt() As Integer
    Material_fehlt = pMaterial_fehlt
End Property

Public Property Let Material_fehlt(value As Integer)
     pMaterial_fehlt = value
End Property

'Personal_fehlt properties - 5
Public Property Get Personal_fehlt() As Integer
    Personal_fehlt = pPersonal_fehlt
End Property

Public Property Let Personal_fehlt(value As Integer)
     pPersonal_fehlt = value
End Property

'Schlosser properties - 6
Public Property Get Schlosser() As Integer
    Schlosser = pSchlosser
End Property

Public Property Let Schlosser(value As Integer)
     pSchlosser = value
End Property

'Stoerung  properties - 7
Public Property Get Stoerung() As Integer
    Stoerung = pStoerung
End Property

Public Property Let Stoerung(value As Integer)
     pStoerung = value
End Property

'Materialprobleme properties - 8
Public Property Get Materialprobleme() As Integer
    Materialprobleme = pMaterialprobleme
End Property

Public Property Let Materialprobleme(value As Integer)
     pMaterialprobleme = value
End Property

'Qualitaetsprobleme properties - 9
Public Property Get Qualitaetsprobleme() As Integer
    Qualitaetsprobleme = pQualitaetsprobleme
End Property

Public Property Let Qualitaetsprobleme(value As Integer)
     pQualitaetsprobleme = value
End Property

'Zeichnung_unklar properties - 10
Public Property Get Zeichnung_unklar() As Integer
    Zeichnung_unklar = pZeichnung_unklar
End Property

Public Property Let Zeichnung_unklar(value As Integer)
     pZeichnung_unklar = value
End Property

'Avprog_fehlt_ueberarbeiten properties - 11
Public Property Get Avprog_fehlt_ueberarbeiten() As Integer
    Avprog_fehlt_ueberarbeiten = pAvprog_fehlt_ueberarbeiten
End Property

Public Property Let Avprog_fehlt_ueberarbeiten(value As Integer)
     pAvprog_fehlt_ueberarbeiten = value
End Property

'WOP properties - 12
Public Property Get WOP() As Integer
    WOP = pWOP
End Property

Public Property Let WOP(value As Integer)
     pWOP = value
End Property

'Abweichung_Planzeit properties - 13
Public Property Get Abweichung_Planzeit() As Integer
    Abweichung_Planzeit = pAbweichung_Planzeit
End Property

Public Property Let Abweichung_Planzeit(value As Integer)
     pAbweichung_Planzeit = value
End Property

'Allg_Qualitaetsprobleme properties - 14
Public Property Get Allg_Qualitaetsprobleme() As Integer
    Allg_Qualitaetsprobleme = pAllg_Qualitaetsprobleme
End Property

Public Property Let Allg_Qualitaetsprobleme(value As Integer)
     pAllg_Qualitaetsprobleme = value
End Property

