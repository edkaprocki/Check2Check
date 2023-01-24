VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form summary_form 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Summary"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5445
   Icon            =   "summary_form.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   2990
      _Version        =   393216
      Rows            =   7
      Cols            =   5
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "summary_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  'If (KeyCode = vbKeyEscape) Then
    Form_Unload 0
  'End If
End Sub

Private Sub Form_Load()
  ' Set up the form
  Dim lR As Long
  lR = SetTopMostWindow(summary_form.hWnd, True)
  
  With grid
    
    .ColAlignment(0) = flexAlignCenterCenter
    .ColAlignment(1) = flexAlignCenterCenter
    .ColAlignment(2) = flexAlignCenterCenter
    .ColAlignment(3) = flexAlignCenterCenter
    .ColAlignment(4) = flexAlignCenterCenter
    
    .RowHeight(2) = 50
End With
  
  summary_form.width = grid.width + 100
  summary_form.height = grid.height + 400
  
End Sub

Private Sub Form_Resize()
  Dim i
  
  If (summary_form.width < 500) Or (summary_form.height < 375) Then Exit Sub
  grid.Top = 0
  grid.Left = 0
  grid.width = summary_form.width - 100
  grid.height = summary_form.height - 375
  
  For i = 0 To grid.Cols - 1
    grid.ColWidth(i) = (grid.width - 350) / grid.Cols
  Next i
  
End Sub

Public Sub update_summary_display()
  Dim i
  Dim index(8)
  Dim ta_inc As Double, tn_inc As Double
  Dim ta_exp As Double, tn_exp As Double
  
  Dim lR As Long
  lR = SetTopMostWindow(summary_form.hWnd, True)
  
  index(0) = PAID_DONE
  index(1) = PAID_QUESTION
  index(2) = PAID_BLANK
  index(3) = PAID_DASH
  
  ta_inc = 0  ' Total amount
  tn_inc = 0  ' Total number
  ta_exp = 0  ' Total amount
  tn_exp = 0  ' Total number
  
  With grid
    .Redraw = False
    
    .TextMatrix(0, 1) = words(INCOME_N)
    .TextMatrix(0, 2) = words(EXPENSE_N)
    .TextMatrix(0, 3) = words(QTY_INC_N)
    .TextMatrix(0, 4) = words(QTY_EXP_N)
  
    .TextMatrix(1, 0) = words(TOTAL_N)
    .TextMatrix(3, 0) = words(DONE_N)
    .TextMatrix(4, 0) = words(PENDING_N)
    .TextMatrix(5, 0) = words(BLANK_N)
    .TextMatrix(6, 0) = words(SKIP_N)
    
    ' ----- Display the paid amounts -----
    .row = 2
    For i = 0 To MAX_PAID
      ta_inc = ta_inc + summary.income(index(i))
      ta_exp = ta_exp + summary.expense(index(i))
      tn_inc = tn_inc + summary.number_income(index(i))
      tn_exp = tn_exp + summary.number_expense(index(i))
    
      .row = .row + 1
      .Col = 1
      .CellForeColor = amount_color(summary.income(index(i)))
      .Text = currency_s(summary.income(index(i)))
      .CellFontBold = True
      
      .Col = 2
      .CellForeColor = amount_color(summary.expense(index(i)))
      .Text = currency_s(summary.expense(index(i)))
      .CellFontBold = True
      
      .Col = 3
      .Text = Str(summary.number_income(index(i)))
      .CellFontBold = False
      
      .Col = 4
      .Text = Str(summary.number_expense(index(i)))
      .CellFontBold = False
      
    Next i
      
    ' ----- Display total -----
    .row = 1
    .Col = 1
    .CellForeColor = amount_color(ta_inc)
    .Text = currency_s(ta_inc)
    .CellFontBold = True
      
    .Col = 2
    .CellForeColor = amount_color(ta_exp)
    .Text = currency_s(ta_exp)
    .CellFontBold = True
      
    .Col = 3
    .Text = Str(tn_inc)
    .CellFontBold = False
      
    .Col = 4
    .Text = Str(tn_exp)
    .CellFontBold = False
      
    ' ----- Display the grand total -----
    .row = 0
    .Col = 0
    .CellForeColor = amount_color(ta_inc + ta_exp)
    .Text = currency_s(ta_inc + ta_exp)
    .CellFontBold = True
      
    .Redraw = True
    
  End With
   
  ' Update the form caption
  summary_form.Caption = words(SUMMARY_N) + " --- " + main_form.entry_tab.Caption  'TabCaption(i)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Hide
End Sub

Public Sub normal()
  Dim lR As Long
  lR = SetTopMostWindow(summary_form.hWnd, True)
  
  summary_form.WindowState = 0  ' Normal state
End Sub

