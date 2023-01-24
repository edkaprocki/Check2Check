VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form quick_form 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Quick Accounts"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame status_frame 
      Caption         =   "Status"
      Height          =   735
      Left            =   1140
      TabIndex        =   9
      Top             =   60
      Width           =   1410
      Begin VB.OptionButton status_radio 
         Caption         =   "Pending"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton status_radio 
         Caption         =   "Done"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.CommandButton delete_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "quick.frx":0000
      DownPicture     =   "quick.frx":0102
      Height          =   315
      Left            =   240
      Picture         =   "quick.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Delete Quick Save Account"
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton calendar_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "quick.frx":0306
      DownPicture     =   "quick.frx":0408
      Height          =   315
      Left            =   660
      Picture         =   "quick.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Show Calendar"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton calculator_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "quick.frx":0694
      DownPicture     =   "quick.frx":0796
      Height          =   315
      Left            =   660
      Picture         =   "quick.frx":0898
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Show Calculator"
      Top             =   120
      Width           =   315
   End
   Begin VB.Frame amount_frame 
      Caption         =   "Amount"
      Height          =   735
      Left            =   2655
      TabIndex        =   5
      Top             =   45
      Width           =   1455
      Begin VB.TextBox amount_box 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.CommandButton OKButton 
      Height          =   615
      Left            =   5580
      Picture         =   "quick.frx":0C29
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   1710
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   180
      Width           =   1305
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   420
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   3675
      Left            =   90
      TabIndex        =   1
      Top             =   900
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   6482
      _Version        =   393216
      Rows            =   31
      Cols            =   7
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "quick_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_COL = 0
Const DATE_COL = 1
Const NAME_COL = 2
Const NEEDED_COL = 3
Const LINE1_COL = 4
Const AMOUNT_COL = 5
Const PENDING_COL = 6


Dim editing_a_record As Boolean  ' True when a cell has been double clicked
Dim accounts_changed As Boolean  ' True when any account has changed
Dim accounts As quick_accounts_type  ' Local copy of accounts
Dim last_row, last_col
Dim ok_pressed As Boolean
Public due_date_s As String
Public name_s As String
Public amount As Double
Public status As Integer  ' 0=blank, 1=paid, 2=skip, 3=skip
Dim entry_mode As Integer  ' 0=view/edit, 1=save, 2=deposit
Dim updating_display As Boolean  ' True when updating the display, used to prevent cell entry function
Dim total_amount_in_all_accounts As Double  ' Total accumulated amount
Dim account_selected As Boolean  ' True when user has selected an account
Dim pending_amount(MAX_QUICK_ACCOUNT + 1) As Double  ' Total including all paid and pending up to the currently selected date




Sub setup_form()
  ReDim accounts.account(MAX_QUICK_ACCOUNT + 1)
  Select Case entry_mode
    Case 0  ' Quick view/edit
      grid.SelectionMode = flexSelectionFree
      delete_button.Visible = True
      quick_form.Caption = words(QUICK_ACCOUNTS_VIEW_EDIT_N)  '"Quick Accounts - View / Edit"
      amount_frame.Enabled = True
      amount_frame.Caption = words(TOTAL_SAVED_N)   '"Total Saved"
      amount_box.Enabled = False
      OKButton.Visible = True
      status_frame.Visible = False
      
    Case 1  ' Quick Save
      grid.SelectionMode = flexSelectionByRow
      delete_button.Visible = False
      quick_form.Caption = words(QUICK_SAVE_N)  '"Quick Save"
      amount_frame.Enabled = True
      amount_frame.Caption = words(AMOUNT_N)  '"Amount"
      amount_box.Enabled = True
      OKButton.Visible = False
      status_frame.Visible = True
      
    Case 2  ' Quick Deposit
      grid.SelectionMode = flexSelectionByRow
      delete_button.Visible = False
      quick_form.Caption = words(QUICK_DEPOSIT_N)  '"Quick Deposit"
      amount_frame.Enabled = True
      amount_frame.Caption = words(AMOUNT_N)  '"Amount"
      amount_box.Enabled = True
      OKButton.Visible = False
      status_frame.Visible = True
    
    End Select
    
End Sub


Public Function execute(index As Integer) As Boolean
  ' Entry point when doing quick save/deposit
  Dim n As Integer
  
  ' i = 0 = Quick View/Edit - Modal
  ' i = 1 = Quick Save - Modal
  ' i = 2 = Quick Deposit - Modal
  
  CancelButton.Caption = words(CANCEL_N)
  status_frame.Caption = words(STATUS_N)
  status_radio(0).Caption = words(DONE_N)
  status_radio(1).Caption = words(PENDING_N)
  
  entry_mode = index
  
  accounts_changed = False
  ok_pressed = False
  execute = False
  account_selected = False
  amount_box.Text = "0.00"
  'status_radio(1).Value = True
  
  ' Setup the form for view/save/deposit
  setup_form
  
  update_display
  
  show vbModal
  If (ok_pressed) Then
    changed_flag = True
    main_form.update_caption
  
    If (entry_mode = 0) Then  ' Doing a save
    End If
    
    If (entry_mode = 1) Then  ' Doing a save
      n = Val(grid.TextMatrix(grid.row, 0))
      name_s = "(QS" + Format(n, "00) ") + grid.TextMatrix(grid.row, 2)
      amount = -Abs(amount_box.Text)  ' Make negative for deposit
      If (status_radio(0).Value) Then status = PAID_DONE
      If (status_radio(1).Value) Then status = PAID_QUESTION
    End If
    
    If (entry_mode = 2) Then  ' Doing a deposit
      n = Val(grid.TextMatrix(grid.row, 0))
      name_s = "(QD" + Format(n, "00) ") + grid.TextMatrix(grid.row, 2)
      amount = Abs(Val(amount_box.Text))  ' Make negative for deposit
      If (status_radio(0).Value) Then status = PAID_DONE
      If (status_radio(1).Value) Then status = PAID_QUESTION
    End If
    
    ' Return true only if an account has been selected
    If (account_selected) Then
      execute = True
    Else
      execute = False
    End If
    
  End If
End Function


Private Sub update_display()
  Dim i
  
  updating_display = True
  
  With grid
  .ColAlignment(NUM_COL) = flexAlignCenterCenter
  .ColAlignment(DATE_COL) = flexAlignLeftCenter
  .ColAlignment(NAME_COL) = flexAlignLeftCenter
  .ColAlignment(NEEDED_COL) = flexAlignRightCenter
  .ColAlignment(AMOUNT_COL) = flexAlignRightCenter
  .ColAlignment(PENDING_COL) = flexAlignRightCenter
  
  .ColWidth(NUM_COL) = 300
  .ColWidth(DATE_COL) = 900
  '.ColWidth(NAME_COL) = quick_form.Width - 4550
  .ColWidth(NEEDED_COL) = 1100  '900
  .ColWidth(LINE1_COL) = 50
  .ColWidth(AMOUNT_COL) = 1000  '900
  .ColWidth(PENDING_COL) = 1000  '900
  
  .Clear
  
  ' Set the column headers to bold
  .row = 0
  For i = 0 To .Cols - 1
    .row = 0
    .Col = i
    .CellFontBold = True
  Next i
  
  
  ' Setup the column headers
  .Rows = 1
  .TextMatrix(0, DATE_COL) = words(DUE_DATE_N)   '"Due Date"
  .TextMatrix(0, NAME_COL) = words(NAME_N)  '"Name"
  .TextMatrix(0, NEEDED_COL) = words(AMOUNT_NEEDED_N)   '"Amount Needed"
  .TextMatrix(0, AMOUNT_COL) = words(AMOUNT_SAVED_N)   '"Amount Saved"
  .TextMatrix(0, PENDING_COL) = words(PENDING_AMOUNT_N)  '"Pending Amount"
  .RowHeight(0) = 400
  
  ' Fill up the grid
  .Rows = 1
   For i = 1 To MAX_QUICK_ACCOUNT
     If ((QUICK_ACCOUNTS.account(i).date <> "") Or _
         (QUICK_ACCOUNTS.account(i).name <> "")) Or _
         (QUICK_ACCOUNTS.account(i).needed <> 0) Or _
         (QUICK_ACCOUNTS.account(i).total <> 0) Or _
         (entry_mode = 0) Then
       .Rows = .Rows + 1
       .TextMatrix(.Rows - 1, NUM_COL) = i
       .TextMatrix(.Rows - 1, DATE_COL) = QUICK_ACCOUNTS.account(i).date
       .TextMatrix(.Rows - 1, NAME_COL) = QUICK_ACCOUNTS.account(i).name
      
       If (QUICK_ACCOUNTS.account(i).date <> "") Or _
          (QUICK_ACCOUNTS.account(i).name <> "") Or _
          (QUICK_ACCOUNTS.account(i).needed <> 0) Or _
          (QUICK_ACCOUNTS.account(i).total <> 0) Then
          .TextMatrix(.Rows - 1, NEEDED_COL) = currency_s(QUICK_ACCOUNTS.account(i).needed)
          .TextMatrix(.Rows - 1, AMOUNT_COL) = currency_s(QUICK_ACCOUNTS.account(i).total)
          .TextMatrix(.Rows - 1, PENDING_COL) = currency_s(pending_amount(i))
       Else
          .TextMatrix(.Rows - 1, NEEDED_COL) = ""
          .TextMatrix(.Rows - 1, AMOUNT_COL) = ""
          .TextMatrix(.Rows - 1, PENDING_COL) = ""
       End If
      
       ' Set the color of the total amount saved
       .row = .Rows - 1
       .Col = AMOUNT_COL
       If (QUICK_ACCOUNTS.account(i).total > 0) Then
         ' Change the color of the cell to blue
         grid.CellForeColor = vbBlue
       Else
         grid.CellForeColor = vbBlack
       End If
       
       ' Set the color of the pending amount saved
       .row = .Rows - 1
       .Col = PENDING_COL
       If (pending_amount(i) > 0) Then
         ' Change the color of the cell to blue
         grid.CellForeColor = vbBlue
       Else
         grid.CellForeColor = vbBlack
       End If
       
       .row = .Rows - 1
       .Col = AMOUNT_COL
       .CellFontBold = True

     End If
   Next i
   
   ' Update the total amount saved so far
   If (entry_mode = 0) Then
     amount_box.Text = currency_s(total_amount_in_all_accounts)
   End If
  
  End With  ' grid
  
  ' If no accounts then disable the amount box
  'If (grid.Rows < 2) Then
  '  amount_box.Enabled = False
  'Else
  '  amount_box.Enabled = True
  'End If
  
  updating_display = False
End Sub



Sub validate_grid()
  ' Save all the fields of the grid into the accounts structure
  Dim d As String
  Dim n As String
  Dim needed_s As String
  Dim i
  
  With grid
    For i = 1 To MAX_QUICK_ACCOUNT
      d = .TextMatrix(i, DATE_COL)
      n = .TextMatrix(i, NAME_COL)
      needed_s = .TextMatrix(i, NEEDED_COL)
      
      If ((d <> "") Or (n <> "") Or (needed_s <> "")) Then
        ' We have data here
        accounts.account(i).date = d
        accounts.account(i).name = n
        accounts.account(i).needed = .TextMatrix(i, NEEDED_COL)
      End If
    Next i
  End With
End Sub


Public Sub check_quick_account()
  Dim s As String
  Dim n As Integer
  
  ' Check the record for not excluded
  If (this.exclude = True) Then Exit Sub
  
  ' If this is a normal transaction then see if it is a quick account
  n = -1
  s = UCase(Mid(this.name, 1, 3))
  If ((s = "(QS") Or (s = "(QD")) Then
    n = Val(Mid(this.name, 4, 2))
  End If
    
  If (n > 0) And (n < MAX_QUICK_ACCOUNT) Then
    ' We have a valid quick account number
    this.cleared = 0
    
    If (this.paid = PAID_DONE) Then
      ' We have a transaction marked as done
      QUICK_ACCOUNTS.account(n).total = -(this.amount) + QUICK_ACCOUNTS.account(n).total   ' Make the amount negative since a negative is actually a positive saving
      total_amount_in_all_accounts = -this.amount + total_amount_in_all_accounts
      this.cleared = 2  ' Mark all quick transactions which are done as cleared
    End If
    
    If (this.paid = PAID_QUESTION) Then
       '(this.paid = PAID_DONE) Then
      ' We have a pending amount so now see if it's from today's date or earlier
      If (this.year <= view.current_year) And _
         (this.month <= view.current_month) And _
         (this.day <= view.current_day) Then
            pending_amount(n) = -(this.amount) + pending_amount(n)
      End If
    End If
  End If
If (this.this = -1) Then
n = -1
End If

  save_this  ' Save "this" transaction back to the database
  
End Sub


Public Sub clear_accounts()
  Dim i
  
  total_amount_in_all_accounts = 0
  For i = 0 To MAX_QUICK_ACCOUNT
    QUICK_ACCOUNTS.account(i).total = 0
    pending_amount(i) = 0
  Next i
End Sub




Private Sub amount_box_GotFocus()
  ' Highlight the text
  amount_box.SelStart = 0
  amount_box.SelLength = 100
End Sub

Private Sub amount_box_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Validate the amount if enter pressed
  Dim v As Double
  
  On Error GoTo error_h
  
  If (KeyCode = vbKeyReturn) Then
    v = amount_box.Text
    amount_box.Text = currency_s(v)
  End If
  Exit Sub
  
error_h:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
End Sub

Private Sub calculator_button_Click()
  ' Run the windows calculator
  Dim RetVal
  Dim i
  
  ' See if the calc is already up there and set focus to it if it is
  i = 0
  On Error GoTo error_h
  AppActivate "Calculator"  ' Attempt to switch to the calculator
  
  If (i = 1) Then
    i = 2
    RetVal = Shell("CALC.EXE", 1)    ' Run the calculator
  End If

  Exit Sub

error_h:
  Err.Clear
  If (i = 0) Then
    i = 1
    Resume Next
  Else
    MsgBox words(CALCULATOR_NOT_FOUND_N)
  End If
End Sub

Private Sub calendar_button_Click()
  ' Show the calendar
  calendar_form.show vbModal
End Sub

Private Sub CancelButton_Click()
  Hide
End Sub

Private Sub delete_button_Click()
  Dim r, ok
  Dim s
  
  ' Delete the current account
  ok = False
  r = grid.row
  s = grid.TextMatrix(r, NAME_COL)
  If (MsgBox(words(DELETE_QUICK_SAVE_ACCOUNT_N) + " (" + s + ")?", _
    vbYesNoCancel + vbQuestion + vbApplicationModal, words(DELETE_ACCOUNT_Q_N)) = vbYes) Then
    ' Yes, delete the account
    ok = True
    If ((grid.TextMatrix(r, DATE_COL) <> "") Or (grid.TextMatrix(r, NAME_COL) <> "") Or (grid.TextMatrix(r, NEEDED_COL) <> "")) And _
       (Val(grid.TextMatrix(r, AMOUNT_COL)) <> 0) Then
      ' We want to delete an account but the balance is not zero
      ' so put up an error message
      If (MsgBox(words(ARE_YOU_SURE_Q_QUICK_SAVE_ACCOUNT_IS_NOT_ZERO_N), _
        vbYesNoCancel + vbCritical + vbApplicationModal, words(WARNING_QUICK_SAVE_ACCOUNT_BALANCE_IS_NOT_ZERO_N)) <> vbYes) Then
        ' No, do not delete the account
        ok = False
      End If
    End If
  End If
     
  If (ok) Then
    grid.TextMatrix(r, DATE_COL) = ""
    grid.TextMatrix(r, NAME_COL) = ""
    grid.TextMatrix(r, NEEDED_COL) = ""
    grid.TextMatrix(r, AMOUNT_COL) = ""
    accounts_changed = True
  End If
End Sub

Private Sub Form_Activate()
  Form_Resize
  
  delete_button.Visible = False
  
  Select Case entry_mode
    Case 0  ' Quick view/edit
      If (quick_form.Visible) Then grid.SetFocus
      
    Case 1  ' Quick Save
      If (quick_form.Visible) Then amount_box.SetFocus
      
    Case 2  ' Quick Deposit
      If (quick_form.Visible) Then amount_box.SetFocus
    
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyEscape) Or (KeyCode = vbKeyReturn) Then
    Form_Unload 0
  End If
End Sub

Private Sub Form_Load()
  ' Set up the form
  'Dim lR As Long
  'lR = SetTopMostWindow(quick_form.hwnd, True)
  
  grid.Rows = MAX_QUICK_ACCOUNT
  grid.ColWidth(NUM_COL) = 500
  grid.ColWidth(DATE_COL) = 1000
  grid.ColWidth(NAME_COL) = 1
  grid.ColWidth(NEEDED_COL) = 2000  '900
  grid.ColWidth(AMOUNT_COL) = 2000 ' 900
 
  grid.row = 0
  grid.Col = 1
  grid.ColAlignment(NUM_COL) = flexAlignCenterCenter
  grid.ColAlignment(DATE_COL) = flexAlignLeftCenter
  grid.ColAlignment(NAME_COL) = flexAlignLeftCenter
  grid.ColAlignment(NEEDED_COL) = flexAlignRightCenter
  grid.ColAlignment(AMOUNT_COL) = flexAlignRightCenter
  
End Sub

Private Sub Form_Resize()
  Dim i
  Dim w, w1
  
  If (quick_form.width < 6500) Or (quick_form.height < 2000) Then Exit Sub
  grid.Left = 0
  grid.width = quick_form.width - 100
  grid.height = quick_form.height - grid.Top - 400
  
  grid.ColWidth(NAME_COL) = grid.width - 4700
  OKButton.Left = quick_form.width - OKButton.width - 200
  CancelButton.Left = OKButton.Left - CancelButton.width - 200
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Hide
End Sub

Public Sub normal()
  Dim lR As Long
  lR = SetTopMostWindow(quick_form.hWnd, True)
  
  quick_form.WindowState = 0  ' Normal state
End Sub


Private Sub grid_DblClick()
  If (entry_mode > 0) Then Exit Sub
  
  With grid
    ' If in the name column then allow for editing
    If (.Col = DATE_COL) Or (.Col = NAME_COL) Or (.Col = NEEDED_COL) Then
      editing_a_record = True
      grid_KeyPress (0)
      txtEdit.SelStart = 0
      txtEdit.SelLength = 100
    End If
  End With
End Sub

Private Sub grid_EnterCell()
  
  If (Not updating_display) Then
    account_selected = True
    OKButton.Visible = True
    If (grid.Col > NEEDED_COL) Then grid.Col = NEEDED_COL
    If (grid.row > 0) And _
       (entry_mode = 0) And _
       ((grid.TextMatrix(grid.row, DATE_COL) <> "") Or (grid.TextMatrix(grid.row, NAME_COL) <> "") Or (grid.TextMatrix(grid.row, NEEDED_COL) <> "")) Then
      delete_button.Visible = True  ' Only allow delete when in view/edit mode
    Else
      delete_button.Visible = False
    End If
  End If
End Sub

Private Sub save_record()
  ' Save the data in cell pointed to by last_col, last_row
End Sub

Private Sub grid_GotFocus()
    Dim rec
    Dim n
    Dim ans
    
    ' Data was entered on the grid
    
    On Error GoTo errorh
    If txtEdit.Visible = False Then Exit Sub
    grid = txtEdit
    txtEdit.Visible = False
        
    
    
    ' We just entered data so now see if this is on a new record or exising record
    With grid
      last_row = .row
      last_col = .Col
      '.row = last_row
      '.Col = last_col
      
      'If ((.Col = NAME_COL) And (.TextMatrix(.row, AMOUNT_COL) = "")) Then .TextMatrix(.row, AMOUNT_COL) = "0"
      
      ' We have an existing record so update the data
      If (.Col = NEEDED_COL) Then
        grid = currency_s(txtEdit)
      End If
        
    End With
    
    editing_a_record = False
    
    If (last_col = DATE_COL) Then
      grid.Col = NAME_COL
      grid.row = last_row
      If (grid.TextMatrix(grid.row, NEEDED_COL) = "") Then
          grid.TextMatrix(grid.row, NEEDED_COL) = "0.00"
          'grid.TextMatrix(grid.row, AMOUNT_COL) = "0.00"
      End If
    Else
      If last_col = NAME_COL Then
          grid.Col = NEEDED_COL
          grid.row = last_row
      If (grid.TextMatrix(grid.row, NEEDED_COL) = "") Then
          grid.TextMatrix(grid.row, NEEDED_COL) = "0.00"
          'grid.TextMatrix(grid.row, AMOUNT_COL) = "0.00"
      End If
      Else
          If last_col = NEEDED_COL Then
            'grid.TextMatrix(grid.row, AMOUNT_COL) = "0.00"
            grid.Col = DATE_COL
            If last_row + 1 < grid.Rows Then
                grid.row = last_row + 1
            Else
                grid.row = last_row
            End If
        End If
      End If
    End If
    
    'grid.SetFocus
    
  Exit Sub

errorh:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
End Sub

Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
  ' Key was pressed so save the data
  accounts_changed = True
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
  If (editing_a_record) Then
    MSFlexGridEdit grid, txtEdit, KeyAscii
  Else
      editing_a_record = True
      grid_KeyPress (KeyAscii)
  End If
End Sub

Private Sub grid_LeaveCell()
    If (grid.row > 0) Then grid.CellBackColor = vbWhite
    
    If txtEdit.Visible = False Then Exit Sub
    
    EditKeyCode grid, txtEdit, vbKeyReturn, 0
    If (vbKeyReturn = 27) Then editing_a_record = False  ' See if escape was pressed
    accounts_changed = True
    
    If (grid.Col = NEEDED_COL) Then
      On Error GoTo error_h
      grid = currency_s(txtEdit)
      GoTo continue:
error_h:
      MsgBox words(INVALID_NUMBER_N)
    Else
      grid = txtEdit
    End If
    
continue:
    txtEdit.Visible = False
    editing_a_record = False
End Sub

Private Sub OKButton_Click()
  Dim v As Double
  Dim vmax As Double
  
  On Error GoTo error_h
  
  If (entry_mode = 0) Then
    If (accounts_changed) Then
      validate_grid
      QUICK_ACCOUNTS = accounts
    End If
  Else
    v = amount_box.Text
  End If
  
  If (entry_mode = 1) Or (entry_mode = 2) Then
    ' Validate that the ammount they want to deposit is available
    vmax = grid.TextMatrix(grid.row, AMOUNT_COL)
    If (v <= vmax) Or _
       (entry_mode = 1) Then
      ' We have enough money in accout so let it go through
      ok_pressed = True
    Else
       ' We don't have enough money in account
       MsgBox words(SORRY_NOT_ENOUGH_MONEY_IN_THAT_ACCOUNT_N) + " " + currency_s(vmax)
       Exit Sub
    End If
  End If
  
  'Ask one more time if they want to save this
  If (entry_mode = 0) And (accounts_changed) Then
    If (MsgBox(words(SAVE_QUICK_ACCOUNT_DATA_Q_N), _
      vbOKCancel + vbQuestion + vbApplicationModal, words(SAVE_QUICK_ACCOUNT_DATA_N)) <> vbOK) Then Exit Sub
  End If
  
  ok_pressed = True
  
  Hide
  Exit Sub
  
error_h:
  MsgBox words(INVALID_AMOUNT_N)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode grid, txtEdit, KeyCode, Shift
    If (KeyCode = 27) Then editing_a_record = False  ' See if escape was pressed
    accounts_changed = True
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    ' Delete returns to get rid of beep.
    If KeyAscii = Val(vbCr) Then KeyAscii = 0
End Sub
