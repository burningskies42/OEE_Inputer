VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "Calendar Control"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3960
   OleObjectBlob   =   "frmCalendar.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------
'****************************************************************************************************
'*                                                                                                  *
'*  Form Purpose: Calender form, for choosing relevant shift date                                   *
'*  Last Change: 14:50 28.11.2016                                                                   *
'*                                                                                                  *
'****************************************************************************************************
'----------------------------------------------------------------------------------------------------

Dim ThisDay As Date
Dim ThisYear, ThisMth As Date
Dim CreateCal As Boolean
Dim i As Integer

Private Sub UserForm_Activate()
    Dim frmCalendar As CFormChanger
    Set frmCalendar = New CFormChanger
    
    With frmCalendar
        '.ShowCloseBtn = False
        '.ShowSysMenu = False
        
        '.IconPath = Application.ActiveWorkbook.Path & "\Uhlmann_Logo.ico"
        .setIconFromWS
        '.ShowIconWS = True
        Set .Form = Me
        
        '.Modal = True
    End With
End Sub

Private Sub UserForm_Initialize()
    Application.EnableEvents = False
    'starts the form on todays date
    ThisDay = Date
    ThisMth = Format(ThisDay, "mm")
    ThisYear = Format(ThisDay, "yyyy")
    For i = 1 To 12
        CB_Mth.AddItem Format(DateSerial(Year(Date), Month(Date) + i, 0), "mmmm")
    Next
    CB_Mth.ListIndex = Format(Date, "mm") - Format(Date, "mm")
    For i = -20 To 50
        If i = 1 Then CB_Yr.AddItem Format((ThisDay), "yyyy") Else CB_Yr.AddItem _
            Format((DateAdd("yyyy", (i - 1), ThisDay)), "yyyy")
    Next
    CB_Yr.ListIndex = 21
    'Builds the calendar with todays date
    frmCalendar.Width = frmCalendar.Width
    CreateCal = True
    Call Build_Calendar
    Application.EnableEvents = True
End Sub
Private Sub CB_Mth_Change()
    'rebuilds the calendar when the month is changed by the user
    Build_Calendar
End Sub
Private Sub CB_Yr_Change()
    'rebuilds the calendar when the year is changed by the user
    Build_Calendar
End Sub
Private Sub Build_Calendar()
    'the routine that actually builds the calendar each time
    If CreateCal = True Then
    frmCalendar.Caption = " " & CB_Mth.value & " " & CB_Yr.value
    'sets the focus for the todays date button
    CommandButton1.SetFocus
    For i = 1 To 42
        If i < Weekday((CB_Mth.value) & "/1/" & (CB_Yr.value)) Then
            Controls("D" & (i)).Caption = Format(DateAdd("d", (i - Weekday((CB_Mth.value) & "/1/" & (CB_Yr.value))), _
                ((CB_Mth.value) & "/1/" & (CB_Yr.value))), "d")
            Controls("D" & (i)).ControlTipText = Format(DateAdd("d", (i - Weekday((CB_Mth.value) & "/1/" & (CB_Yr.value))), _
                ((CB_Mth.value) & "/1/" & (CB_Yr.value))), "m/d/yy")
        ElseIf i >= Weekday((CB_Mth.value) & "/1/" & (CB_Yr.value)) Then
            Controls("D" & (i)).Caption = Format(DateAdd("d", (i - Weekday((CB_Mth.value) _
                & "/1/" & (CB_Yr.value))), ((CB_Mth.value) & "/1/" & (CB_Yr.value))), "d")
            Controls("D" & (i)).ControlTipText = Format(DateAdd("d", (i - Weekday((CB_Mth.value) & "/1/" & (CB_Yr.value))), _
                ((CB_Mth.value) & "/1/" & (CB_Yr.value))), "dd/mm/yy")
        End If
        If Format(DateAdd("d", (i - Weekday((CB_Mth.value) & "/1/" & (CB_Yr.value))), _
        ((CB_Mth.value) & "/1/" & (CB_Yr.value))), "mmmm") = ((CB_Mth.value)) Then
            If Controls("D" & (i)).BackColor <> &H80000016 Then Controls("D" & (i)).BackColor = &H80000018  '&H80000010
            Controls("D" & (i)).Font.Bold = True
        If Format(DateAdd("d", (i - Weekday((CB_Mth.value) & "/1/" & (CB_Yr.value))), _
            ((CB_Mth.value) & "/1/" & (CB_Yr.value))), "dd/mm/yy") = Format(ThisDay, "dd/mm/yy") Then Controls("D" & (i)).SetFocus
        Else
            If Controls("D" & (i)).BackColor <> &H80000016 Then Controls("D" & (i)).BackColor = &H8000000F
            Controls("D" & (i)).Font.Bold = False
        End If
    Next
    End If
End Sub
Private Sub D1_Click()
    'this sub and the ones following represent the buttons for days on the form
    'retrieves the current value of the individual controltiptext and
    'places it in the active cell
    ActiveCell.value = D1.ControlTipText
    Unload Me
    'after unload you can call a different userform to continue data entry
    'uncomment this line and add a userform named UserForm2
    'Userform2.Show
    
End Sub
Private Sub D2_Click()
    ActiveCell.value = D2.ControlTipText
    Unload Me
    
End Sub
Private Sub D3_Click()
    ActiveCell.value = D3.ControlTipText
    Unload Me
    
End Sub
Private Sub D4_Click()
    ActiveCell.value = D4.ControlTipText
    Unload Me
    
End Sub
Private Sub D5_Click()
    ActiveCell.value = D5.ControlTipText
    Unload Me
    
End Sub
Private Sub D6_Click()
    ActiveCell.value = D6.ControlTipText
    Unload Me
    
End Sub
Private Sub D7_Click()
    ActiveCell.value = D7.ControlTipText
    Unload Me
    
End Sub
Private Sub D8_Click()
    ActiveCell.value = D8.ControlTipText
    Unload Me
    
End Sub
Private Sub D9_Click()
    ActiveCell.value = D9.ControlTipText
    Unload Me
    
End Sub
Private Sub D10_Click()
    ActiveCell.value = D10.ControlTipText
    Unload Me
    
End Sub
Private Sub D11_Click()
    ActiveCell.value = D11.ControlTipText
    Unload Me
    
End Sub
Private Sub D12_Click()
    ActiveCell.value = D12.ControlTipText
    Unload Me
    
End Sub
Private Sub D13_Click()
    ActiveCell.value = D13.ControlTipText
    Unload Me
    
End Sub
Private Sub D14_Click()
    ActiveCell.value = D14.ControlTipText
    Unload Me
    
End Sub
Private Sub D15_Click()
    ActiveCell.value = D15.ControlTipText
    Unload Me
    
End Sub
Private Sub D16_Click()
    ActiveCell.value = D16.ControlTipText
    Unload Me
    
End Sub
Private Sub D17_Click()
    ActiveCell.value = D17.ControlTipText
    Unload Me
    
End Sub
Private Sub D18_Click()
    ActiveCell.value = D18.ControlTipText
    Unload Me
    
End Sub
Private Sub D19_Click()
    ActiveCell.value = D19.ControlTipText
    Unload Me
    
End Sub
Private Sub D20_Click()
    ActiveCell.value = D20.ControlTipText
    Unload Me
    
End Sub
Private Sub D21_Click()
    ActiveCell.value = D21.ControlTipText
    Unload Me
    
End Sub
Private Sub D22_Click()
    ActiveCell.value = D22.ControlTipText
    Unload Me
    
End Sub
Private Sub D23_Click()
    ActiveCell.value = D23.ControlTipText
    Unload Me
    
End Sub
Private Sub D24_Click()
    ActiveCell.value = D24.ControlTipText
    Unload Me
    
End Sub
Private Sub D25_Click()
    ActiveCell.value = D25.ControlTipText
    Unload Me
    
End Sub
Private Sub D26_Click()
    ActiveCell.value = D26.ControlTipText
    Unload Me
    
End Sub
Private Sub D27_Click()
    ActiveCell.value = D27.ControlTipText
    Unload Me
    
End Sub
Private Sub D28_Click()
    ActiveCell.value = D28.ControlTipText
    Unload Me
    
End Sub
Private Sub D29_Click()
    ActiveCell.value = D29.ControlTipText
    Unload Me
    
End Sub
Private Sub D30_Click()
    ActiveCell.value = D30.ControlTipText
    Unload Me
    
End Sub
Private Sub D31_Click()
    ActiveCell.value = D31.ControlTipText
    Unload Me
    
End Sub
Private Sub D32_Click()
    ActiveCell.value = D32.ControlTipText
    Unload Me
    
End Sub
Private Sub D33_Click()
    ActiveCell.value = D33.ControlTipText
    Unload Me
    
End Sub
Private Sub D34_Click()
    ActiveCell.value = D34.ControlTipText
    Unload Me
    
End Sub
Private Sub D35_Click()
    ActiveCell.value = D35.ControlTipText
    Unload Me
    
End Sub
Private Sub D36_Click()
    ActiveCell.value = D36.ControlTipText
    Unload Me
    
End Sub
Private Sub D37_Click()
    ActiveCell.value = D37.ControlTipText
    Unload Me
    
End Sub
Private Sub D38_Click()
    ActiveCell.value = D38.ControlTipText
    Unload Me
    
End Sub
Private Sub D39_Click()
    ActiveCell.value = D39.ControlTipText
    Unload Me
    
End Sub
Private Sub D40_Click()
    ActiveCell.value = D40.ControlTipText
    Unload Me
    
End Sub
Private Sub D41_Click()
    ActiveCell.value = D41.ControlTipText
    Unload Me
    
End Sub
Private Sub D42_Click()
    ActiveCell.value = D42.ControlTipText
    Unload Me
    
End Sub


