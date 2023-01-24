VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form calendar_form 
   BackColor       =   &H80000013&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendar"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2700
   Icon            =   "cal_form.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView calendar 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   86179841
      CurrentDate     =   36378
   End
End
Attribute VB_Name = "calendar_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim return_date As Boolean  ' True indicates simple return
Dim d As Date
Dim cal As calendar_type
Dim double_click As Boolean  ' True when double click has occurred


' 1/28/2017 ES
' It doesn't look like this function is used for anything in the project
' I don't know what it's intended purpose was
Public Function execute_card(m As Integer, d As Integer, y As Integer, n As String) As Boolean
  ' Return the day, month, year if double clicked
  
  return_date = True
  double_click = False
  show vbModal
  If (double_click) Then
    ' Ok or double click hit so save it
    m = cal.Month
    d = cal.day
    y = cal.Year
    n = cal.name
    execute_card = True
  End If
  return_date = False
End Function


Private Sub calendar_DateDblClick(ByVal DateDblClicked As Date)
  Dim d
  Dim da As Integer
  Dim mo As Integer
  Dim yr As Integer
  Dim name As String
  
  ' Go to the selected month/day
  d = DateDblClicked
  yr = Val(Format(d, "yyyy"))
  mo = Val(Format(d, "mm"))
  da = Val(Format(d, "dd"))
  name = Format(d, "dddd")
  
  If (Not return_date) Then
    view.start_day = da
    view.start_month = mo
    view.current_month = mo
    view.start_year = yr
    view.current_year = yr
  
    main_form.going_to_day = True
    main_form.update_entry_tabs
  Else
    cal.day = da
    cal.Month = mo
    cal.Year = yr
    cal.name = name
    double_click = True
    Hide
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyEscape) Then
    Form_Unload 0
  End If
End Sub

Private Sub Form_Load()
  ' Set up the calendar
  Dim lR As Long
  lR = SetTopMostWindow(calendar_form.hWnd, True)
  calendar.Value = date
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Hide
End Sub
