VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form reconcile_form 
   Caption         =   "Reconcile Checkbook"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8115
   Icon            =   "reconcile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame sort_by_frame 
      Caption         =   "Sort By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   3195
      Begin VB.OptionButton sort_by_radio 
         Caption         =   "Date"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton sort_by_radio 
         Caption         =   "Check Number"
         Height          =   255
         Index           =   1
         Left            =   1500
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton edit_button 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Picture         =   "reconcile.frx":08CA
      TabIndex        =   20
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton finish_later_button 
      Caption         =   "Finish Later"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   6450
      Picture         =   "reconcile.frx":0C5E
      TabIndex        =   19
      Top             =   180
      Width           =   1590
   End
   Begin VB.Frame bank_statement_frame 
      Caption         =   "Bank Statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   3195
      Begin VB.TextBox difference_box 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1620
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   1320
         Width           =   1395
      End
      Begin VB.TextBox cleared_balance_box 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1620
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   960
         Width           =   1395
      End
      Begin VB.TextBox ending_balance_box 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1620
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox beginning_balance_box 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1620
         TabIndex        =   0
         Text            =   "Text3"
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Difference"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Cleared Balance"
         Height          =   195
         Left            =   300
         TabIndex        =   15
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ending Balance"
         Height          =   195
         Left            =   300
         TabIndex        =   14
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Beginning Balance"
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame cleared_transactions_frame 
      Caption         =   "Cleared Transactions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3480
      TabIndex        =   8
      Top             =   60
      Width           =   2895
      Begin VB.CheckBox show_all_check 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   24
         Top             =   1380
         Width           =   1890
      End
      Begin VB.TextBox deposits_box 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   1380
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox checks_box 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   1380
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox withdrawals_box 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   1380
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1020
         Width           =   1335
      End
      Begin VB.OptionButton cleared_radio 
         Caption         =   "Checks"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   660
         Width           =   1035
      End
      Begin VB.OptionButton cleared_radio 
         Caption         =   "Withdrawals"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   1020
         Width           =   1275
      End
      Begin VB.OptionButton cleared_radio 
         Caption         =   "Deposits"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.CommandButton finish_button 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6480
      Picture         =   "reconcile.frx":105A
      TabIndex        =   6
      Top             =   840
      Width           =   1545
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   6480
      Picture         =   "reconcile.frx":140D
      TabIndex        =   7
      Top             =   1440
      Width           =   1545
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3195
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   12
      FixedCols       =   2
      BackColorSel    =   65535
      ForeColorSel    =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
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
Attribute VB_Name = "reconcile_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const name_width = 5000

Const DATE_COL = 0
Const CHECK_COL = 1
Const CLEARED_COL = 2
Const AMOUNT_COL = 3
Const STATUS_COL = 4
Const EXCLUDE_COL = 5
Const TAGS_COL = 6
Const NAME_COL = 7
Const THIS_COL = 8
Const QDATE_COL = 9
Const INDEX_COL = 10
Const SORT_COL = 11  ' Used for sorting by date since they are stuffed in the table by date

Private Type totals_type
  index As Integer
  sum As Double
  cleared_sum As Double
End Type

Private Type all_type
   records() As r_type
   index() As Integer
End Type

Private all As all_type
Private totals(3) As totals_type

Private current_select As Integer  ' Which one of the 3 is selected
Private beginning_balance As Double
Private ending_balance As Double
Private difference As Double
Private save_flag As Integer  ' 0=finish exit, 1=finish later


Private Sub update_language()
  reconcile_form.Caption = words(RECONCILE_CHECKBOOK_N)
  bank_statement_frame.Caption = words(BANK_STATEMENT_N)
  Label1.Caption = words(BEGINNING_BALANCE_N)
  Label2.Caption = words(ENDING_BALANCE_N)
  Label3.Caption = words(CLEARED_BALANCE_N)
  Label4.Caption = words(DIFFERENCE_N)
  cleared_transactions_frame.Caption = words(CLEARED_TRANSACTIONS_N)
  cleared_radio(2).Caption = words(DEPOSITS_N)
  cleared_radio(0).Caption = words(CHECKS_N)
  cleared_radio(1).Caption = words(WITHDRAWALS_N)
  show_all_check.Caption = words(SHOW_ALL_N)
  edit_button.Caption = words(EDIT_N)
  finish_later_button.Caption = words(FINISH_LATER_N)
  finish_button.Caption = words(FINISH_N)
  CancelButton.Caption = words(CANCEL_N)
  sort_by_frame.Caption = words(SORT_BY_N)
  sort_by_radio(0).Caption = words(DATE_N)
  sort_by_radio(1).Caption = words(CHECK_NUMBER_N)
End Sub


Public Sub clear_table()
  With grid
    .Rows = 2
    .Clear
    .TextMatrix(0, DATE_COL) = words(DATE_N)  '"Date"
    .TextMatrix(0, CHECK_COL) = words(NUMBER_N)  '"Number"
    .TextMatrix(0, CLEARED_COL) = words(CLR_N)  '"CLR"
    .TextMatrix(0, AMOUNT_COL) = words(AMOUNT_N)  '"Amount"
    .TextMatrix(0, STATUS_COL) = words(STATUS_N)  '"Stat"
    .TextMatrix(0, EXCLUDE_COL) = words(EXCLUDE_N)   '"Excl"
    .TextMatrix(0, TAGS_COL) = words(TAGS_N)   '"Tags"
    .TextMatrix(0, NAME_COL) = words(NAME_N)  '"Name"
    .TextMatrix(0, THIS_COL) = "this"
    .TextMatrix(0, QDATE_COL) = "date"
    .TextMatrix(0, INDEX_COL) = "index"
    
    .ColAlignment(DATE_COL) = flexAlignCenterCenter
    .ColAlignment(CHECK_COL) = flexAlignCenterCenter
    .ColAlignment(CLEARED_COL) = flexAlignCenterCenter
    .ColAlignment(STATUS_COL) = flexAlignCenterCenter
    .ColAlignment(EXCLUDE_COL) = flexAlignCenterCenter
    .ColAlignment(TAGS_COL) = flexAlignCenterCenter
    .ColAlignment(NAME_COL) = flexAlignLeftCenter
    
    .row = 0
    .Col = AMOUNT_COL
    .CellAlignment = flexAlignCenterCenter
    .row = 0
    .Col = NAME_COL
    .CellAlignment = flexAlignCenterCenter
    
    .row = 1
    .Col = 2
  End With
End Sub

Private Sub Form_Activate()
  beginning_balance_box.SetFocus
End Sub

Private Sub sort_by_radio_Click(index As Integer)
  If (index = 0) Then sort_by_radio(0).Value = True
  If (index = 1) Then sort_by_radio(1).Value = True
  sort_grid
End Sub

Private Sub Form_Load()
  With grid
    .ColWidth(DATE_COL) = 1000
    .ColWidth(CHECK_COL) = 800
    .ColWidth(CLEARED_COL) = 400
    .ColWidth(AMOUNT_COL) = 1100
    .ColWidth(STATUS_COL) = 600
    .ColWidth(EXCLUDE_COL) = 600
    .ColWidth(TAGS_COL) = 600
    .ColWidth(NAME_COL) = 2000
    .ColWidth(THIS_COL) = 1
    .ColWidth(QDATE_COL) = 1
    .ColWidth(INDEX_COL) = 1
    .ColWidth(SORT_COL) = 1
    Form_Resize
  End With
End Sub

Private Sub Form_Resize()
  ' Resise the form
  If (reconcile_form.width < 5000) Or (reconcile_form.height < 3000) Then Exit Sub
  
  grid.width = reconcile_form.width - 300
  grid.height = reconcile_form.height - 3100
  grid.ColWidth(NAME_COL) = grid.width - name_width
  finish_button.Left = reconcile_form.width - finish_button.width - 200
  finish_later_button.Left = reconcile_form.width - finish_later_button.width - 200
  CancelButton.Left = reconcile_form.width - CancelButton.width - 200
End Sub

Public Sub display()
  Dim s
  Dim n, i, j, r
  
  reconcile_form.MousePointer = vbHourglass

  With grid
    .Redraw = False
    clear_table
    
    If (cleared_radio(0).Value) Then n = 0
    If (cleared_radio(1).Value) Then n = 1
    If (cleared_radio(2).Value) Then n = 2
    current_select = n
    
  
    If (totals(n).index > 0) Then
      r = 1
      For i = 0 To totals(n).index - 1
        
        If (show_all_check.Value = vbChecked) Or (all.records(n, i).paid = 1) Then
          ' Transaction is either done or we are set to show all
          .row = r
          
          .TextMatrix(r, SORT_COL) = r ' Put in an index number so we can sort by date for it
          
          s = get_date(all.records(n, i).month, all.records(n, i).day, all.records(n, i).year)

          .Col = DATE_COL
          .Text = s
    
          .Col = CHECK_COL
          If (all.records(n, i).check >= 0) Then .Text = all.records(n, i).check
          '.CellFontBold = False
    
          .Col = CLEARED_COL
          .CellFontSize = 14
          If (all.records(n, i).cleared = 1) Then .Text = "*"
          If (all.records(n, i).cleared = 2) Then .Text = "X"
          '.CellFontBold = False
    
          .Col = AMOUNT_COL
          .Text = currency_s(all.records(n, i).amount)
          .CellForeColor = amount_color(all.records(n, i).amount)
          '.CellFontBold = False
    
          .Col = STATUS_COL
          j = all.records(n, i).paid
          If (j = 0) Then .Text = ""        ' Blank
          If (j = 1) Then .Text = Chr(149)  ' Dot
          If (j = 2) Then .Text = "?"       ' ?
          If (j = 3) Then .Text = Chr(150)  ' Dash
          .CellFontSize = 14
          .CellFontBold = False
    
          .Col = EXCLUDE_COL
          j = all.records(n, i).exclude
          If (j = False) Then .Text = ""        ' Blank
          If (j = True) Then .Text = Chr(149)  ' Dot
          .CellFontSize = 14
          .CellFontBold = False
    
          .Col = TAGS_COL
          s = ""
          If ((all.records(n, i).tags And 1) <> 0) Then s = "1"
          If ((all.records(n, i).tags And 2) <> 0) Then s = s + "2"
          If ((all.records(n, i).tags And 4) <> 0) Then s = s + "3"
          If ((all.records(n, i).tags And 8) <> 0) Then s = s + "4"
          .Text = s
          .CellFontSize = 8
          .CellFontBold = False
    
          .Col = NAME_COL
          .Text = all.records(n, i).name
          '.CellFontBold = False
    
          .Col = THIS_COL
          .Text = all.records(n, i).this
    
          .Col = QDATE_COL
    
          .Col = INDEX_COL
          .Text = i  ' Save the index so we know which one to work with
    
          .Rows = .Rows + 1
          .row = .row + 1
          r = r + 1
        End If
      Next i
  
    End If
  
  .Redraw = True
  If (.Rows > 2) Then .Rows = .Rows - 1
  End With
  
  update_totals
  reconcile_form.MousePointer = vbDefault

End Sub

Public Sub initialize()
  ReDim all.records(2, data.number_of_records + 10) As r_type
  ReDim all.index(2, data.number_of_records + 10) As Integer
  
  totals(0).index = 0
  totals(0).sum = 0
  totals(0).cleared_sum = 0
  totals(1).index = 0
  totals(1).sum = 0
  totals(1).cleared_sum = 0
  totals(2).index = 0
  totals(2).sum = 0
  totals(2).cleared_sum = 0
  current_select = 0
  clear_table
End Sub

Public Sub add()
  Dim n
  
  n = this.year + this.month * 33 + this.day
  
  If (this.amount > 0) Then
    ' We have a deposit
    all.records(2, totals(2).index) = this
    all.index(2, totals(2).index) = n
    
    totals(2).index = totals(2).index + 1
  Else
    If (this.check > -1) Then
      ' We have a valid check number so put it in the table
      all.records(0, totals(0).index) = this
      all.index(0, totals(0).index) = n
      
      ' Do the totals
      totals(0).index = totals(0).index + 1
        
    Else
      ' We have a withdrawal
      all.records(1, totals(1).index) = this
      all.index(1, totals(1).index) = n
      
      totals(1).index = totals(1).index + 1
    End If
  End If
  
End Sub

Private Sub sort_grid()
  ' Show the sort by frame if we are displaying the checks only
  grid.row = 1
  grid.RowSel = grid.Rows - 1
  sort_by_frame.Visible = cleared_radio(0).Value  ' Show or hide the sort_by panel
  
  If (sort_by_radio(1).Value = True) Then
    grid.Col = CHECK_COL
  Else
    grid.Col = SORT_COL
  End If
  
  grid.Sort = flexSortNumericAscending  ' Do the actual sort now
  'grid.Sort = flexSortNone
  
  grid.row = 1
  grid.Col = 2
  If (reconcile_form.Visible) Then grid.SetFocus
End Sub

Public Sub execute()
  beginning_balance = data.bank_balance_beginning
  ending_balance = data.bank_balance_ending
  
  compute_totals
  display
  cleared_radio(2).Value = True
  grid.row = 1
  grid.Col = 2
  grid.ColSel = grid.Cols - 1
  sort_by_radio(0).Value = True  ' Sort by date on entry
  sort_grid
  
  update_language
  
  show vbModal
End Sub

Private Sub beginning_balance_box_GotFocus()
  beginning_balance_box.SelStart = 0
  beginning_balance_box.SelLength = 20
End Sub

Private Sub beginning_balance_box_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo error_h
  If (KeyCode = vbKeyReturn) Then
    beginning_balance = beginning_balance_box
    ending_balance_box.SetFocus
  End If
  Exit Sub
  
error_h:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
End Sub

Private Sub beginning_balance_box_LostFocus()
  On Error GoTo error_h
  ' Check the value in the box
  beginning_balance = beginning_balance_box
  
  update_totals
  Exit Sub
  
error_h:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
End Sub

Private Sub CancelButton_Click()
  Dim a
  
  a = MsgBox(words(ARE_YOU_SURE_Q_N), _
      vbYesNoCancel + vbQuestion + vbApplicationModal, words(CANCEL_RECONCILIATION_Q_N))
    
  If (a = vbYes) Then
    Hide
    Exit Sub
  End If
End Sub

Private Sub cleared_button_Click()
  Dim index As Integer
  
  ' Mark the selected transactions as cleared
  With grid
    index = Val(.TextMatrix(.row, INDEX_COL))
    'If (records(current_select, index).cleared = 0) Then
    '  totals(current_select).cleared_sum = totals(current_select).cleared_sum + records(current_select, index).amount
    'End If
    
    all.records(current_select, index).cleared = 1
    .TextMatrix(.row, CLEARED_COL) = "*"
  End With
  
  update_totals
End Sub

Private Sub edit_button_Click()
  Dim r As Integer
  Dim row As Integer
  
  ' Edit the record at the current line
  r = -1
  With grid
    row = .row
    If (.TextMatrix(.row, INDEX_COL) <> "") Then
      r = Val(.TextMatrix(.row, INDEX_COL))
    End If
    
    If (r >= 0) Then
      ' We have a valid record so edit it
      this = all.records(current_select, r)
      If (edit_transaction_form.execute) Then
        ' We have a valid edit so save it
        all.records(current_select, r) = this
        display
      End If
      If (row <= .Rows - 1) Then
        .row = row
      Else
         .row = row - 1
      End If
      .Col = 2
    End If
  End With
  
  ' Compute the totals
  compute_totals
  update_totals
  
End Sub

Private Sub ending_balance_box_GotFocus()
  ending_balance_box.SelStart = 0
  ending_balance_box.SelLength = 20
End Sub

Private Sub ending_balance_box_KeyDown(KeyCode As Integer, Shift As Integer)
  On Error GoTo error_h
  If (KeyCode = vbKeyReturn) Then
    ending_balance = ending_balance_box
    grid.SetFocus
  End If
  Exit Sub
  
error_h:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
End Sub

Private Sub ending_balance_box_LostFocus()
  On Error GoTo error_h
  ' Check the value in the box
  ending_balance = ending_balance_box
  
  update_totals
  Exit Sub
  
error_h:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If (reconcile_form.Visible) Then finish_later_button_Click
End Sub

Private Sub grid_DblClick()
  Call grid_KeyDown(vbKeySpace, 0)
End Sub

Private Sub grid_EnterCell()
    If (grid.Col > 1) And (grid.row > 0) Then grid.CellBackColor = vbYellow
End Sub

Private Sub grid_GotFocus()
    grid.ColSel = grid.Cols - 1
End Sub

Private Sub grid_LeaveCell()
  If (grid.Col > 1) And (grid.row > 0) Then grid.CellBackColor = vbWhite
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim t As Integer
  Dim rows_displayed As Integer
  Dim r As Integer
  Dim index As Integer
  
  With grid
  t = .TopRow
  r = .row
  ' Get the current transaction number from the index_col
  index = Val(.TextMatrix(.row, INDEX_COL))
  
  If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeySpace) Then
    ' Toggle the value
    If (all.records(current_select, index).cleared = 0) Then
      all.records(current_select, index).cleared = 1
      .TextMatrix(.row, CLEARED_COL) = "*"
      compute_totals
      'totals(current_select).cleared_sum = totals(current_select).cleared_sum + records(current_select, index).amount
    Else
      all.records(current_select, index).cleared = 0
      .TextMatrix(.row, CLEARED_COL) = ""
      'totals(current_select).cleared_sum = totals(current_select).cleared_sum - records(current_select, index).amount
      compute_totals
    End If
    
    If (KeyCode = vbKeyReturn) And (.row + 1 < .Rows) Then
       .row = .row + 1
       .ColSel = .Cols - 1
    End If
  End If
  
  ' Set the top row
  ' Compute the number or rows displayed
  rows_displayed = (.height - 500) / .RowHeight(0)
  If (.row > t + rows_displayed - 1) Then
    .TopRow = t + 1
  End If
  
  End With
  
  update_totals
End Sub

Private Sub not_cleared_button_Click()
  Dim index As Integer
  
  ' Mark the selected transactions as cleared
  ' Get the current transaction number from the index_col
  With grid
    index = Val(.TextMatrix(.row, INDEX_COL))
    'If (records(current_select, index).cleared > 0) Then
    '  totals(current_select).cleared_sum = totals(current_select).cleared_sum - records(current_select, index).amount
    'End If
    
    all.records(current_select, index).cleared = 0
    .TextMatrix(.row, CLEARED_COL) = ""
  End With
  
  compute_totals
  update_totals
End Sub

Private Sub cleared_radio_Click(index As Integer)
  current_select = index
  clear_table
  display
  Call Form_Resize
  
  sort_grid
  
  grid.row = 1
  grid.Col = 2
  If (reconcile_form.Visible) Then grid.SetFocus
End Sub


Private Sub update_totals()
  Dim sum As Double
  
  ' Update the bank balances
  beginning_balance_box.Text = currency_s(beginning_balance)
  ending_balance_box.Text = currency_s(ending_balance)
  
  ' Update the totals boxes
  checks_box.Text = currency_s(totals(0).cleared_sum)
  withdrawals_box.Text = currency_s(totals(1).cleared_sum)
  deposits_box.Text = currency_s(totals(2).cleared_sum)
  
  sum = totals(0).cleared_sum + totals(1).cleared_sum + totals(2).cleared_sum
  cleared_balance_box.Text = currency_s(sum)
  
  ' Get the total difference
  difference = sum - (ending_balance - beginning_balance)
  difference_box.Text = currency_s(difference)
  
  ' Set the box colors
  beginning_balance_box.ForeColor = amount_color(beginning_balance)
  ending_balance_box.ForeColor = amount_color(ending_balance)
  checks_box.ForeColor = amount_color(totals(0).cleared_sum)
  withdrawals_box.ForeColor = amount_color(totals(1).cleared_sum)
  deposits_box.ForeColor = amount_color(totals(2).cleared_sum)
  cleared_balance_box.ForeColor = amount_color(sum)
  difference_box.ForeColor = amount_color(difference)
End Sub

Private Sub save_reconcilation_data()
  Dim a, i, j
  
  ' If the difference is zero then ask if we want to mark everything as cleared permanently
  a = vbNo  ' Set to not finalize
  If (save_flag = 0) Then  ' Do this only if trying to finish
    If (Abs(difference) < 0.0001) Then
      a = MsgBox(words(DO_YOU_WANT_TO_FINALIZE_Q_N), _
          vbYesNoCancel + vbQuestion + vbApplicationModal, words(SUCCESS_RECONCILIATION_COMPLETE_N))
    Else
      If (save_flag = 0) Then
        a = MsgBox(words(THE_DIFFERENCE_IS_NOT_ZERO_Q_N), _
            vbYesNoCancel + vbQuestion + vbApplicationModal, words(CAUTION_RECONCILIATION_NOT_COMPLETE_N))
        If (a = vbCancel) Or (a = vbNo) Then Exit Sub
      End If
    End If
  End If
  
  ' Save the new cleared contents
  For i = 0 To 2
    If (totals(i).index > 0) Then
      ' We have a transaction so save it
      For j = 0 To totals(i).index - 1
        ' Loop through all the records for this type
        If (all.records(i, j).cleared = 1) And (a = vbYes) Then all.records(i, j).cleared = 2 ' Mark the records as completly cleared
        db(all.records(i, j).this).name = all.records(i, j).name
        db(all.records(i, j).this).amount = all.records(i, j).amount
        db(all.records(i, j).this).cleared = all.records(i, j).cleared
        db(all.records(i, j).this).check = all.records(i, j).check
        db(all.records(i, j).this).paid = all.records(i, j).paid
        db(all.records(i, j).this).exclude = all.records(i, j).exclude
      Next j
    End If
  Next i
  
  data.bank_balance_beginning = beginning_balance
  data.bank_balance_ending = ending_balance
  If (a = vbYes) Then
    ' This completes the reconcilation process so move the ending balance to the beginning balance
    data.bank_balance_beginning = data.bank_balance_ending
    data.bank_balance_ending = 0
  End If
  
  changed_flag = True
  main_form.update_caption
  
  Hide  ' Exit the form
End Sub

Private Sub finish_later_button_Click()
  save_flag = 1
  save_reconcilation_data
End Sub


Private Sub finish_button_Click()
  save_flag = 0
  save_reconcilation_data
End Sub


Private Sub show_all_check_Click()
  display
End Sub

Private Sub compute_totals()
  Dim i As Integer
  Dim j As Integer
  
  For i = 0 To 2
    totals(i).cleared_sum = 0
    For j = 0 To totals(i).index - 1
      If (all.records(i, j).cleared = 1) And _
         (all.records(i, j).paid = 1) And _
         (all.records(i, j).exclude = False) Then totals(i).cleared_sum = totals(i).cleared_sum + all.records(i, j).amount
    Next j
  Next i
  
  
End Sub
