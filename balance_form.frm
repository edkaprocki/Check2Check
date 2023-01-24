VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form balance_form 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Balances"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6906
      _Version        =   393216
      Rows            =   15
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
Attribute VB_Name = "balance_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyEscape) Or (KeyCode = vbKeyReturn) Then
    Form_Unload 0
  End If
End Sub

Private Sub Form_Load()
  ' Set up the form
  Dim lR As Long
  lR = SetTopMostWindow(balance_form.hWnd, True)
  
  grid.TextMatrix(0, 1) = words(BEGINNING_N)
  grid.TextMatrix(0, 2) = words(LOW_N)
  grid.TextMatrix(0, 3) = words(ENDING_N)
  grid.TextMatrix(0, 4) = words(CHANGE_N)
  grid.TextMatrix(10, 0) = words(AVERAGE_N)
  
  grid.row = 0
  grid.Col = 1
  grid.CellAlignment = flexAlignCenterCenter
  grid.Col = 2
  grid.CellAlignment = flexAlignCenterCenter
  grid.Col = 3
  grid.CellAlignment = flexAlignCenterCenter
  grid.Col = 4
  grid.CellAlignment = flexAlignCenterCenter
  grid.Col = 2
  grid.CellAlignment = flexAlignCenterCenter
  grid.Col = 3
  grid.CellAlignment = flexAlignCenterCenter
  
  grid.TextMatrix(13, 1) = "============="
  grid.TextMatrix(13, 2) = "============="
  grid.TextMatrix(13, 3) = "============="
  grid.TextMatrix(13, 4) = "============="
  
  grid.TextMatrix(14, 0) = words(AVERAGE_N)
  'grid.TextMatrix(13, 2) = words(LOW_N)
  'grid.TextMatrix(13, 3) = words(ENDING_N)
  'grid.TextMatrix(13, 4) = words(CHANGE_N)

End Sub

Private Sub Form_Resize()
  Dim i
  
  If (balance_form.width < 500) Or (balance_form.height < 375) Then Exit Sub
  grid.Top = 0
  grid.Left = 0
  grid.width = balance_form.width - 100
  grid.height = balance_form.height - 375
  
  For i = 0 To 4
    grid.ColWidth(i) = (grid.width - 350) / 5
  Next i
  
End Sub

Private Sub update_balances()
  Dim i, j, m, y
  Dim average_amounts(14)   ' We add 1 for Low and Delta
  ' Scan through the displayed tabs
    ' Get the starting month and year
    m = view.start_month
    y = view.start_year
    
    For i = 0 To 11
      ' See if we are done with the current year
      m = m + 1
      If (m > 11) Then
        m = 0
        y = y + 1
      End If
    Next i
End Sub


Public Function clear_row_colors(row)
Dim c, r, i

  With grid
    r = .row  ' Save the row and column and restore them when done
    c = .Col
    
    .row = r
    For i = 0 To 4
        .Col = i
        .CellBackColor = vbWhite
    Next i
    
    ' Restore the row and column
    .row = r
    .Col = c
  End With
  
End Function


Public Sub update_balance_display()
  Dim i, j, m, n, mm, mm0, mm1, yy0, yy1, y, yy, b, c
  Dim a As Double
  Dim average_amounts(4)   ' 4 columns of data
  Dim Active_Row As Integer
  Dim Active_Home_Row As Integer
  Dim jj As Integer
  Dim Old_month, old_day, Old_year
  Dim lR As Long
  lR = SetTopMostWindow(balance_form.hWnd, True)
  Dim s As String
  Dim color, white_row, cyan_row, yellow_row
  
  '----- Zero out the averages -----
   For j = 0 To 3
    average_amounts(j) = 0
   Next j

  '----- Zero out colors -----
  white_row = 0
  cyan_row = 0
  
    
    
  '-------------------- Here we go with recursing the grid --------------------------
  With grid
    .Redraw = False
    
    For i = 0 To 11  ' This will cover all months displayed
      .row = i + 1 ' The +1 was to start below the header line
      .Col = 0
      
      clear_row_colors (i)  ' Clear out any previous colors
      
      .Text = main_form.entry_tab.TabCaption(i)
      If (i = main_form.entry_tab.Tab) Then
        ' We are on our current month so hightlight the entire line
        .CellFontBold = True
      Else
       .CellFontBold = False
       Active_Row = i
     End If
    
    '----- Highlight the current selected month -----
    mm0 = Val(Format(.Text, "mm"))
    yy0 = Val(Format(.Text, "yyyy"))
    
    ' ----- Get the current date and year as intergers -----
    mm1 = Val(Format(Now, "mm"))  ' Get today's date
    yy1 = Val(Format(Now, "yyyy"))
    
  
    ' ----- Update a line in the grid -----
    ' Set upthe highlighted line
      If (.CellFontBold = True) Then white_row = .row
      If ((mm0 = mm1) And (yy0 = yy1)) Then cyan_row = .row
      
      ' ----- Set up for beginning balance -----s
      .Col = 1
      .CellForeColor = amount_color(balance_summary(i).beginning)
      .Text = currency_s(balance_summary(i).beginning)
      .CellFontBold = True
      average_amounts(0) = average_amounts(0) + balance_summary(i).beginning
    
      ' ----- Set up for low balance -----
      .Col = 2
      a = balance_summary(i).low
      If (a > 9999999#) Then a = 0
      .CellForeColor = amount_color(a)
      .Text = currency_s(a)
      .CellFontBold = False
      average_amounts(1) = average_amounts(1) + a
       
      ' ----- Set up for ending balance -----
      .Col = 3
      .CellForeColor = amount_color(balance_summary(i).ending)
      .Text = currency_s(balance_summary(i).ending)
      .CellFontBold = True
      average_amounts(2) = average_amounts(2) + balance_summary(i).ending
   
      ' ----- Set up for change -----
      a = balance_summary(i).ending - balance_summary(i).beginning
      .Col = 4
      .CellForeColor = amount_color(a)
      .Text = currency_s(a)
      .CellFontBold = True
      average_amounts(3) = average_amounts(3) + a
   
 
      
      ' ----- See if we are done with the current year -----
      m = m + 1
      If (m > 11) Then
        m = 0
        y = y + 1
      End If
    Next i
    
    
    
   ' ----- Now go highlight the all thebackground colors -----
  For n = 0 To 4: .Col = n: .CellBackColor = vbWhite: Next n
  If ((white_row >= 1) And (white_row <= 13)) Then .row = white_row: For n = 0 To 4: .Col = n: .CellBackColor = vbYellow: Next n
  If ((cyan_row >= 1) And (cyan_row <= 13)) Then .row = cyan_row: For n = 0 To 4: .Col = n: .CellBackColor = vbCyan: Next n
 
    '----- Now set the colors for the average amounts -----
    .row = 14
      For j = 0 To 3  ' 4 columns to show
        a = average_amounts(j) / 12  ' Average  over 12 months
        .Col = j + 1
        .CellForeColor = amount_color(a)
        .Text = currency_s(a)
        .CellFontBold = True
      Next j
  
    .Redraw = True
    
  End With
   
  balance_form.Caption = words(BALANCES_N) + " --- " + main_form.entry_tab.TabCaption(0) + " - " + main_form.entry_tab.TabCaption(11)
  
  ' Now that sthe form is loaded
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Hide
End Sub

Public Sub normal()
  Dim lR As Long
  lR = SetTopMostWindow(balance_form.hWnd, True)
  
  balance_form.WindowState = 0  ' Normal state
End Sub

