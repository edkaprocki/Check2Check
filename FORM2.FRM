VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form main_form 
   Caption         =   "Check 2 Check"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8280
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form2"
   MouseIcon       =   "Form2.frx":0000
   ScaleHeight     =   5970
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "printer"
      Height          =   315
      Left            =   840
      TabIndex        =   25
      Top             =   420
      Width           =   1035
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1860
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox beginning_balance_box 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   120
      Width           =   1275
   End
   Begin VB.TextBox ending_balance_box 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   480
      Width           =   1275
   End
   Begin VB.CommandButton previous_month_button 
      Caption         =   "<"
      Height          =   615
      Left            =   60
      TabIndex        =   17
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton next_month_button 
      Caption         =   ">"
      Height          =   615
      Left            =   7920
      TabIndex        =   16
      Top             =   840
      Width           =   375
   End
   Begin VB.OptionButton transactions_radio 
      Caption         =   "Transactions"
      Height          =   315
      Left            =   2880
      TabIndex        =   15
      Top             =   480
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.OptionButton notes_radio 
      Caption         =   "Notes"
      Height          =   195
      Left            =   4320
      TabIndex        =   14
      Top             =   540
      Width           =   1095
   End
   Begin VB.TextBox notes_box 
      Height          =   1455
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Text            =   "Form2.frx":0442
      Top             =   4200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton calendar_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":044C
      DownPicture     =   "Form2.frx":054E
      Height          =   315
      Left            =   3540
      Picture         =   "Form2.frx":0650
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Show Calendar"
      Top             =   60
      Width           =   315
   End
   Begin MSComCtl2.MonthView calendar 
      Height          =   2370
      Left            =   2040
      TabIndex        =   8
      Top             =   3300
      Visible         =   0   'False
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      StartOfWeek     =   24510465
      CurrentDate     =   36378
   End
   Begin VB.CommandButton delete_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":07DA
      DownPicture     =   "Form2.frx":08DC
      Height          =   315
      Left            =   1740
      Picture         =   "Form2.frx":09DE
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Delete"
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton insert_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":0AE0
      DownPicture     =   "Form2.frx":0BE2
      Height          =   315
      Left            =   1380
      Picture         =   "Form2.frx":0CE4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Insert"
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton paste_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":0DE6
      DownPicture     =   "Form2.frx":0EE8
      Height          =   315
      Left            =   2940
      Picture         =   "Form2.frx":0FEA
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Paste transaction"
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton copy_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":10EC
      DownPicture     =   "Form2.frx":11EE
      Height          =   315
      Left            =   2580
      Picture         =   "Form2.frx":12F0
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Copy transaction"
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton cut_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":13F2
      DownPicture     =   "Form2.frx":14F4
      Height          =   315
      Left            =   2220
      Picture         =   "Form2.frx":15F6
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cut transaction"
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton save_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":16F8
      DownPicture     =   "Form2.frx":17FA
      Height          =   315
      Left            =   780
      Picture         =   "Form2.frx":18FC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton open_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":19FE
      DownPicture     =   "Form2.frx":1B00
      Height          =   315
      Left            =   420
      Picture         =   "Form2.frx":1C02
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open"
      Top             =   60
      Width           =   315
   End
   Begin VB.CommandButton new_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "Form2.frx":1D04
      DownPicture     =   "Form2.frx":1E06
      Height          =   315
      Left            =   60
      Picture         =   "Form2.frx":1F08
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "New"
      Top             =   60
      Width           =   315
   End
   Begin MSComDlg.CommonDialog open_dialog 
      Left            =   7380
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin TabDlg.SSTab entry_tab 
      Height          =   600
      Left            =   480
      TabIndex        =   21
      Top             =   840
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   1058
      _Version        =   393216
      Tabs            =   12
      Tab             =   3
      TabsPerRow      =   12
      TabHeight       =   882
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Jul"
      TabPicture(0)   =   "Form2.frx":200A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Aug"
      TabPicture(1)   =   "Form2.frx":2026
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Sep"
      TabPicture(2)   =   "Form2.frx":2042
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Oct"
      TabPicture(3)   =   "Form2.frx":205E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Nov"
      TabPicture(4)   =   "Form2.frx":207A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Dec"
      TabPicture(5)   =   "Form2.frx":2096
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Jan"
      TabPicture(6)   =   "Form2.frx":20B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Feb"
      TabPicture(7)   =   "Form2.frx":20CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Mar"
      TabPicture(8)   =   "Form2.frx":20EA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Apr"
      TabPicture(9)   =   "Form2.frx":2106
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "May"
      TabPicture(10)  =   "Form2.frx":2122
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Jun"
      TabPicture(11)  =   "Form2.frx":213E
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
   End
   Begin MSFlexGridLib.MSFlexGrid entry_grid 
      Height          =   3315
      Left            =   480
      TabIndex        =   22
      Top             =   1440
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   5847
      _Version        =   393216
      Rows            =   32
      Cols            =   14
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   12583104
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      MergeCells      =   1
   End
   Begin VB.Label beginning_balance_label 
      Alignment       =   1  'Right Justify
      Caption         =   "Beginning Balance"
      Height          =   195
      Left            =   5340
      TabIndex        =   24
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label ending_balance_label 
      Alignment       =   1  'Right Justify
      Caption         =   "Ending Balance"
      Height          =   255
      Left            =   5460
      TabIndex        =   23
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label total_label 
      Caption         =   "Label3"
      Height          =   195
      Left            =   4020
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label last_label 
      Caption         =   "Label3"
      Height          =   195
      Left            =   4560
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label first_label 
      Caption         =   "Label3"
      Height          =   195
      Left            =   4560
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu file_menu 
      Caption         =   "&File"
      Begin VB.Menu new_menu 
         Caption         =   "&New"
      End
      Begin VB.Menu open_menu 
         Caption         =   "&Open"
      End
      Begin VB.Menu close_menu 
         Caption         =   "&Close"
      End
      Begin VB.Menu dummy4 
         Caption         =   "-"
      End
      Begin VB.Menu save_menu 
         Caption         =   "&Save"
      End
      Begin VB.Menu save_as_menu 
         Caption         =   "Save &As"
      End
      Begin VB.Menu dummy1 
         Caption         =   "-"
      End
      Begin VB.Menu exit_menu 
         Caption         =   "E&xit"
      End
      Begin VB.Menu spare1_menu 
         Caption         =   "-"
      End
      Begin VB.Menu recent_file_menu 
         Caption         =   "Recentfile1"
         Index           =   0
      End
      Begin VB.Menu recent_file_menu 
         Caption         =   "Recentfile2"
         Index           =   1
      End
      Begin VB.Menu recent_file_menu 
         Caption         =   "Recentfile3"
         Index           =   2
      End
      Begin VB.Menu recent_file_menu 
         Caption         =   "Recentfile4"
         Index           =   3
      End
   End
   Begin VB.Menu edit_menu 
      Caption         =   "&Edit"
      Begin VB.Menu insert_menu 
         Caption         =   "&Insert"
      End
      Begin VB.Menu delete_menu 
         Caption         =   "&Delete"
      End
      Begin VB.Menu dummy2 
         Caption         =   "-"
      End
      Begin VB.Menu cut_menu 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy_menu 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste_menu 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu dummy3 
         Caption         =   "-"
      End
      Begin VB.Menu cut_month_menu 
         Caption         =   "Cut Month"
      End
      Begin VB.Menu copy_month_menu 
         Caption         =   "Copy Month"
      End
      Begin VB.Menu paste_month_menu 
         Caption         =   "Paste Month"
      End
   End
   Begin VB.Menu view_menu 
      Caption         =   "&View"
      Begin VB.Menu next_month_menu 
         Caption         =   "&Next Month"
      End
      Begin VB.Menu previous_month_menu 
         Caption         =   "&Previous Month"
      End
      Begin VB.Menu vspare1 
         Caption         =   "-"
      End
      Begin VB.Menu transactions_menu 
         Caption         =   "&Transactions"
      End
      Begin VB.Menu notes_menu 
         Caption         =   "&Notes"
      End
      Begin VB.Menu vspare2 
         Caption         =   "-"
      End
      Begin VB.Menu calendar_menu 
         Caption         =   "&Calendar"
      End
      Begin VB.Menu override_columns_menu 
         Caption         =   "&Override Columns"
      End
      Begin VB.Menu sparee1 
         Caption         =   "-"
      End
      Begin VB.Menu goto_month_menu 
         Caption         =   "Go to Month / Year"
      End
   End
   Begin VB.Menu search_menu 
      Caption         =   "&Search"
      Visible         =   0   'False
   End
   Begin VB.Menu help_menu 
      Caption         =   "&Help"
      Begin VB.Menu index_menu 
         Caption         =   "&Index"
         Visible         =   0   'False
      End
      Begin VB.Menu about_menu 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DATE_COL = 0
Const DAY_COL = 1
Const THIS_COL = 2
Const PREV_COL = 3
Const NEXT_COL = 4
Const NAME_COL = 5
Const PAID_COL = 6
Const AMOUNT_COL = 7
Const EXCLUDE_COL = 8
Const DOUBLE_LINE1_COL = 9
Const BALANCE_COL = 10
Const DOUBLE_LINE2_COL = 11
Const OVERRIDE_COL = 12
Const OVERRIDE_AMOUNT_COL = 13

Dim I, j, start_dir, last_row, last_col
Dim changed_flag As Boolean
Dim doing_move As Boolean
Dim move_from_row As Integer
Dim debug_variable As Boolean

Private Sub Command1_Click()
  print_form.Show (vbModal)
End Sub

Private Sub Form_Load()
  Dim commandline As Variant
  Dim no_command_line As Boolean
  
  ' See if the file association already exists
  ' associate_file_type
  
  start_dir = GetSetting("Check 2 Check", "Settings", "Directory", "C:\")
  
  initialize
  
  notes_box.Top = entry_grid.Top
  notes_box.Left = entry_grid.Left
  
  entry_grid.MergeCol(0) = True
  entry_grid.MergeCol(1) = True
  
  entry_grid.ColWidth(DATE_COL) = 500
  entry_grid.ColWidth(DAY_COL) = 500
  entry_grid.ColWidth(THIS_COL) = 1  '300
  entry_grid.ColWidth(PREV_COL) = 1  '300
  entry_grid.ColWidth(NEXT_COL) = 30  '300
  entry_grid.ColWidth(NAME_COL) = 3200
  entry_grid.ColWidth(PAID_COL) = 500
  entry_grid.ColWidth(AMOUNT_COL) = 1000
  entry_grid.ColWidth(EXCLUDE_COL) = 400
  entry_grid.ColWidth(DOUBLE_LINE1_COL) = 30
  entry_grid.ColWidth(BALANCE_COL) = 1000
  entry_grid.ColWidth(DOUBLE_LINE2_COL) = 30
  entry_grid.ColWidth(OVERRIDE_COL) = 400
  entry_grid.ColWidth(OVERRIDE_AMOUNT_COL) = 1000
  
  ' Put up the data for this current month
  For I = 1 To 31
    entry_grid.TextArray(egi(I, 0)) = I
  Next I
  
  'entry_grid.Redraw = True
  
  For I = 0 To MAX_DATA_TABLE
    db(I).this = -1
  Next I
  
  ' Start with no records loaded
  data.first = 0
  data.last = 0
  data.current = 0
  data.number_of_records = 0
  
  For I = 0 To MAX_DATA_TABLE
    db(I).this = -1
  Next I
  
  view.start_month = month(Now)
  view.start_year = year(Now)
  view.current_month = view.start_month
  view.current_year = view.start_year
  
  entry_grid.ColAlignment(PAID_COL) = flexAlignCenterCenter
  entry_grid.ColAlignment(EXCLUDE_COL) = flexAlignCenterCenter
  entry_grid.ColAlignment(OVERRIDE_COL) = flexAlignCenterCenter
  
  copy_of_this.this = -1  ' Start off with nothing in the copy buffer
  entry_tab.Tab = 0
  doing_move = False
  update_entry_tabs
  
  debug_variable = False
  
  ' See if there is a command line file to load
  commandline = Command$()   'GetCommandLine(1)
    
  If (commandline <> "") Then
    open_dialog.FileName = Dir(commandline)
    open_dialog.InitDir = strip_filename(commandline)
    start_dir = strip_filename(commandline)
    ChDir start_dir
    ChDrive get_drive(commandline)
    Call open_the_file(False)
    
    save_to_recent_docs (commandline)
    
    'MsgBox (commandline + "  " + open_dialog.FileName + "  " + open_dialog.InitDir + "  " + start_dir + "  " + get_drive(commandline))
  End If
  
  Get_Recent_Files
  
End Sub

Private Sub about_menu_Click()
  frmAbout.Show 1
End Sub

Private Sub calendar_menu_Click()
  ' Show or hide the calendar
  If (calendar_menu.Checked) Then
    ' Hide it
    calendar_form.Hide
  Else
    ' Show it
    calendar_form.Show
  End If
  calendar_menu.Checked = Not calendar_menu.Checked
End Sub

Private Sub Close_menu_Click()
  ' Close menu clicked
  save_menu_Click
  new_menu_Click
End Sub

Private Sub calendar_button_Click()
  calendar_menu_Click
End Sub

Private Sub copy_button_Click()
  copy_menu_Click
End Sub

Private Sub copy_menu_Click()
  Dim rec As Integer
  
  If (notes_radio.Value = True) Then Exit Sub
  
  ' Copy the selected record
  With entry_grid
    If (.TextMatrix(.Row, THIS_COL) <> "") Then
      ' We have a record to copy
      data.current = Int(.TextMatrix(.Row, THIS_COL))
      get_record
      copy_of_this = this
      
      copy_to_clipboard
    End If
  End With
  
End Sub

Private Sub copy_to_clipboard()
  Dim s
  
  With entry_grid
  
    s = Format(this.month) + "/" + .TextMatrix(.Row, DATE_COL) + "/" + Format(this.year) + "  " + _
       .TextMatrix(.Row, DAY_COL) + "      " + _
       .TextMatrix(.Row, NAME_COL) + "   "
    s = s + "Amount ( " + .TextMatrix(.Row, AMOUNT_COL) + ")   "
    s = s + "Balance ( " + .TextMatrix(.Row, BALANCE_COL) + ")   "
  
    If (this.paid) Then s = s + " Done  "
    If (this.exclude) Then s = s + " Exclude  "
    If (this.override) Then s = s + " Override"
  
  End With
  
  Clipboard.SetText (s)
End Sub

Private Sub copy_month_menu_Click()
  Dim j
  
  If (view.records_in_month = 0) Then
    MsgBox "No transactions to copy"
    Exit Sub
  End If
  
  ' Copy the current month data to the copy_of_month
  j = 0
  For I = 1 To MAX_RECORDS_IN_MONTH
    If (table_image.table(I).day = 0) Then Exit For
    If (table_image.table(I).this >= 0) Then
      data.current = table_image.table(I).this
      get_record  ' Read the selected record and save all the parameters
      copy_of_month.table(j) = this
      j = j + 1
    End If
  Next I
  copy_of_month.table(j).this = -1
    
  copy_of_month.month = view.current_month
  copy_of_month.year = view.current_year
End Sub

Private Sub cut_button_Click()
  cut_menu_Click
End Sub

Private Sub cut_menu_Click()
  If (notes_radio.Value = True) Then Exit Sub
  
  ' Copy and then delete the record
  copy_menu_Click
  delete_menu_Click
  changed_flag = True
  update_caption
End Sub

Private Sub delete_button_Click()
  delete_menu_Click
  changed_flag = True
  update_caption
End Sub

Private Sub delete_menu_Click()
  Dim rec As Integer, t
  
  ' Delete a record
  With entry_grid
    t = .TopRow
    
    .Redraw = False
    If (.TextMatrix(.Row, THIS_COL) <> "") Then
      ' We have a valid record to delete
      rec = .TextMatrix(.Row, THIS_COL)
      delete_record (rec)
      process
    Else
      ' We are deleting a blank line so see if it's the first in the date
      If (.TextMatrix(.Row, DATE_COL) = .TextMatrix(.Row - 1, DATE_COL)) Then
        ' It is not the first so go ahead and delete the entire line
        shuffle_up_entry_grid (.Row)
      End If
    End If
    .Redraw = True
    'If (.Row > 10) Then .TopRow = .Row - 7
    .TopRow = t
  End With
End Sub

Private Sub cut_month_menu_Click()
  Dim j
  
  If (view.records_in_month = 0) Then
    MsgBox "No transactions to cut"
    Exit Sub
  End If
  
  If (MsgBox("Cut all transactions for " + entry_tab.Caption, _
      vbYesNoCancel + vbQuestion + vbApplicationModal, "Delete Month") = vbYes) Then
    ' Yes, delete the entire month
    With table_image
      j = 0
      For I = 1 To MAX_RECORDS_IN_MONTH
        If (.table(I).day = 0) Then Exit For
      
        If (.table(I).this >= 0) Then
          ' We have a record to delete
          
          ' Get the record and save it
          data.current = .table(I).this
          get_record  ' Read the selected record and save all the parameters
          copy_of_month.table(j) = this
          j = j + 1
          
          delete_record (.table(I).this)
        End If
      Next I
      copy_of_month.table(j).this = -1
    
      copy_of_month.month = view.current_month
      copy_of_month.year = view.current_year
  
      ' Delete the notes too
      notes_box.Text = ""
      update_notes
      
      process
      changed_flag = True
      update_caption
    End With
  
  End If
End Sub

Private Sub entry_grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyInsert Then
        insert_menu_Click
    ElseIf KeyCode = vbKeyDelete Then
        delete_menu_Click
    End If
End Sub

Private Sub entry_grid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim r As Integer
  
  ' Calculate the row it is on
  r = Int(y / entry_grid.RowHeight(1)) + entry_grid.TopRow - 1
  If (r < 1) Then Exit Sub
  
  entry_grid.Row = r
  
  ' See if left button pressed for move
  If Button = 1 Then
    
    main_form.MousePointer = 99  '2
    
    move_from_row = entry_grid.Row
    If (entry_grid.TextMatrix(entry_grid.Row, THIS_COL) <> "") Then
      doing_move = True
    End If
  End If
  
  ' See if right button pressed for menu
  If Button = 2 Then
    PopupMenu edit_menu
  End If
End Sub

Private Sub entry_grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim r, day
  
  main_form.MousePointer = vbDefault
  
  If (Button = 1) And (doing_move) Then
    ' Calculate the row it is on
    r = Int(y / entry_grid.RowHeight(1)) + entry_grid.TopRow - 1
    If (move_from_row <> r) And _
       (r >= 1) And (r < entry_grid.Rows) Then
       '(entry_grid.TextMatrix(move_from_row, DATE_COL) <> (entry_grid.TextMatrix(r, DATE_COL))) Then
       
      ' We have a valid move
      entry_grid.Row = move_from_row
      day = entry_grid.TextMatrix(r, DATE_COL)
      cut_menu_Click
      'If (r > move_from_row) Then r = r - 1
      'entry_grid.Row = r
      
      copy_of_this.day = day
      paste_menu_Click
      End If
    
    doing_move = False
  End If
End Sub

Private Sub exit_menu_Click()
  Unload Me
End Sub

Private Function egi(r, C) As Integer
  egi = r * entry_grid.Cols + C
End Function

Private Sub update_entry_tabs()
  Dim y
  
  y = view.start_year
  For I = 0 To 11
    j = view.start_month + I
    
    If (j > 12) Then
    j = j - 12
      y = view.start_year + 1
    End If
    
    entry_tab.TabCaption(I) = MONTH_STRINGS(j) + " " + CStr(y)
  Next I
  process
  entry_grid.Col = NAME_COL
  entry_grid.Row = 1
  
End Sub

Private Sub entry_grid_DblClick()
  Dim r
  
  ' Bring up the transaction form for this record
  'transaction_form.Show vbModal
  
  With entry_grid
    ' If in the name column then allow for editing
    If (.Col = NAME_COL) Or (.Col = AMOUNT_COL) Then
      If (.TextMatrix(.Row, THIS_COL) <> "") Then
      ' We have a valid record so see what it is
      entry_grid_KeyPress (0)
      txtEdit.SelStart = 0
      txtEdit.SelLength = 100
      End If
    End If
    
    ' If in the paid column then toggle this item
    If (.Col = PAID_COL) Then
      If (.TextMatrix(.Row, THIS_COL) <> "") Then
      ' We have a valid record so see what it is
      r = .TextMatrix(.Row, THIS_COL)
      db(r).paid = Not db(r).paid
      changed_flag = True
      update_caption
      process
      End If
    End If
    
    ' If in the include column then toggle this item
    If (.Col = EXCLUDE_COL) Then
      If (.TextMatrix(.Row, THIS_COL) <> "") Then
      ' We have a valid record so see what it is
      r = .TextMatrix(.Row, THIS_COL)
      db(r).exclude = Not db(r).exclude
      changed_flag = True
      update_caption
      process
      End If
    End If
    
    ' If in the override column then toggle this item
    If (.Col = OVERRIDE_COL) Then
      If (.TextMatrix(.Row, THIS_COL) <> "") Then
      ' We have a valid record so see what it is
      r = .TextMatrix(.Row, THIS_COL)
      db(r).override = Not db(r).override
      changed_flag = True
      update_caption
      process
      End If
    End If
  End With
  
End Sub

Sub entry_grid_KeyPress(KeyAscii As Integer)
  MSFlexGridEdit entry_grid, txtEdit, KeyAscii
End Sub

Private Sub update_notes()
  ' See if the current notes file wants to be added or deleted
  For I = 0 To MAX_NOTES
    If (notes(I).month = view.current_month) And _
       (notes(I).year = view.current_year) Then
      ' Yes we found a notes file
      notes(I).s = notes_box.Text
      
      If (notes(I).s = "") Then
        notes(I).month = 0
      End If
      Exit For
    End If
  Next I
  
  If (I = MAX_NOTES + 1) And (notes_box.Text <> "") Then
    ' No notes found so loop through all the notes and find a blank one
    For I = 0 To MAX_NOTES
      If (notes(I).s = "") Then
        ' We found a blank spot
        notes(I).month = view.current_month
        notes(I).year = view.current_year
        notes(I).s = notes_box.Text
        data.number_of_notes = data.number_of_notes + 1
        Exit For
      End If
    Next I
  End If
  
End Sub

Private Sub entry_tab_Click(PreviousTab As Integer)
  update_notes
  process
End Sub

Private Sub Form_Resize()
  If (main_form.WindowState = 1) Then Exit Sub  ' Don't resize if minimized
  If (main_form.Height < 2300) Then Exit Sub
  
  entry_grid.Height = main_form.Height - 2300
  
  notes_box.Height = entry_grid.Height
  notes_box.Width = entry_grid.Width
  
  ' Move the balance boxes
  beginning_balance_box.Left = main_form.Width - 300 - beginning_balance_box.Width
  beginning_balance_label.Left = beginning_balance_box.Left - 100 - beginning_balance_label.Width
  ending_balance_box.Left = main_form.Width - 300 - ending_balance_box.Width
  ending_balance_label.Left = ending_balance_box.Left - 100 - ending_balance_label.Width
  
  ' Resize the next button
  next_month_button.Left = main_form.Width - 300 - next_month_button.Width
  
  ' Resize the entry tab
  entry_tab.Width = next_month_button.Left - 50 - entry_tab.Left

  ' Resize the table
  entry_grid.Width = next_month_button.Left - 50 - entry_grid.Left
  
  ' Resize the name column
  If (override_columns_menu.Checked) Then
    If (entry_grid.Width > 5700) Then entry_grid.ColWidth(NAME_COL) = entry_grid.Width - 5700
  Else
    If (entry_grid.Width > 4300) Then entry_grid.ColWidth(NAME_COL) = entry_grid.Width - 4300
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim answer
  
  SaveSetting "Check 2 Check", "Settings", "Directory", start_dir
  If (changed_flag) Then
    answer = MsgBox("Save Database?", vbYesNo + vbQuestion, "Check 2 Check")
    If answer = vbYes Then
      save_menu_Click
    End If
  End If
  
  Unload calendar_form
End Sub

Private Sub goto_month_menu_Click()
  goto_month_form.Show vbModal

  If (goto_month_form.ok) Then
    view.start_month = goto_month_form.month_combo.ListIndex + 1
    view.start_year = CInt(goto_month_form.year_combo.List(goto_month_form.year_combo.ListIndex))
    update_entry_tabs
  End If
End Sub

Private Sub insert_button_Click()
  insert_menu_Click
End Sub

Private Sub insert_menu_Click()
  With entry_grid
    .Redraw = False
    ' Insert a new row below this current line
    shuffle_down_entry_grid (.Row + 1)
    If (.Row + 1 = .Rows) Then .Rows = .Rows + 1
    .TextMatrix(.Row + 1, DATE_COL) = .TextMatrix(.Row, DATE_COL)
    .TextMatrix(.Row + 1, DAY_COL) = .TextMatrix(.Row, DAY_COL)
    .Row = .Row + 1
    .Redraw = True
    If (.Row > 10) Then .TopRow = .Row - 7
  End With
End Sub

Private Sub new_button_Click()
  new_menu_Click
End Sub

Private Sub new_menu_Click()
  Dim answer
  
  If (changed_flag) Then
    answer = MsgBox("Save Database?", vbYesNoCancel + vbQuestion, "Check 2 Check")
    If answer = vbCancel Then
      Exit Sub
    Else
       If answer = vbYes Then
        save_menu_Click
       End If
    End If
  End If
  
  data.db_name = ""
  changed_flag = False
  update_caption
  
  ' Start with a fresh database
  data.first = 0
  data.last = 0
  data.number_of_records = 0
  
  ' Zero out the data table
  For I = 0 To MAX_DATA_TABLE
    db(I).this = -1
  Next I
  
  ' Zero out the notes file
  For I = 0 To MAX_NOTES
    notes(I).month = 0
    notes(I).s = ""
  Next I
  
  process
End Sub

Private Sub next_month_button_Click()
  next_month_menu_Click
End Sub

Private Sub next_month_menu_Click()
  view.start_month = view.start_month + 1
  If (view.start_month > 12) Then
    view.start_month = 1
    view.start_year = view.start_year + 1
  End If
  update_entry_tabs
End Sub

Private Sub notes_menu_Click()
  ' Display the entry_grid
  If (entry_grid.Visible) Then
    entry_grid.Visible = False
    notes_box.Visible = True
    
    transactions_radio.Value = False
    notes_radio.Value = True
  End If
End Sub

Private Sub notes_radio_Click()
  notes_menu_Click
  update_notes
End Sub

Private Sub open_button_Click()
  open_menu_Click
End Sub

Private Sub open_the_file(showit As Boolean)
  With open_dialog
    If (showit) Then
        .InitDir = start_dir
        .Filter = "Check2Check files | *.c2c"
        .ShowOpen
        start_dir = CurDir()
    End If
    If (.FileName <> "") Then
       ' We have a filename so strip off the extension
      j = InStr(1, .FileName, ".")
      If (j > 0) Then
        data.db_name = Left(.FileName, j - 1)
      End If
      
      ' Zero out the notes file
      For I = 0 To MAX_NOTES
        notes(I).month = 0
        notes(I).s = ""
      Next I
      
      update_caption
      read_database
      process
      changed_flag = False
      update_caption
    End If
  End With

  Save_Recent_Files
  
End Sub
Private Sub open_menu_Click()
  Dim answer
  
  If (changed_flag) Then
    answer = MsgBox("Save Database?", vbYesNoCancel + vbQuestion, "Check 2 Check")
    If answer = vbCancel Then
      Exit Sub
    Else
       If answer = vbYes Then
        save_menu_Click
       End If
    End If
  End If
  
  On Error GoTo error_h
  
  Call open_the_file(True)
  
  Exit Sub
  
error_h:
  ' We hit the cancel button
End Sub

Private Sub override_columns_menu_Click()
  ' Show or not show the override columns
  override_columns_menu.Checked = Not override_columns_menu.Checked
  Form_Resize
End Sub

Private Sub paste_button_Click()
  paste_menu_Click
  
End Sub

Private Sub paste_menu_Click()
  If (notes_radio.Value = True) Then Exit Sub
  
  ' Paste the record now
  If (copy_of_this.this >= 0) Then
    ' We have a valid record
    this = copy_of_this
    this.month = view.current_month
    this.year = view.current_year
    If (Not doing_move) Then  ' Doing a move so use the day that was given by mouse up
      this.day = Int(entry_grid.TextMatrix(entry_grid.Row, DATE_COL))
    End If
    insert_record (-1)
    process
  End If
End Sub

Private Sub paste_month_menu_Click()
  Dim days_adjusted
  
  If (MsgBox("Paste all transactions to " + entry_tab.Caption, _
      vbYesNoCancel + vbQuestion + vbApplicationModal, "Paste Month") = vbYes) Then
    ' Yes, Paste the entire month
  
    ' Paste the month buffer to the new month
    ' Be sure to check for the last day and overruns
    days_adjusted = False
    With copy_of_month
      'If (.month = view.current_month) And (.year = view.current_year) Then
        ' Can't copy to the same month and year
        'MsgBox "Can't paste to the same month and year"
        'Exit Sub
      'End If
    
      For I = 0 To MAX_RECORDS_IN_MONTH
        If (.table(I).this <= -1) Then Exit For
        ' We have a record to transfer
        this = .table(I)
        this.paid = False
        If (this.day > view.number_of_days) Then
          ' We must adjust the days because there are less days in the month
          this.day = view.number_of_days
          days_adjusted = True
        End If
        this.month = view.current_month
        this.year = view.current_year
        insert_record (-1)
      Next I
    End With
  
    process
    changed_flag = True
    update_caption
    If (days_adjusted) Then MsgBox "Adjusted dates to match the current month"
  
  End If
End Sub

Private Sub previous_month_menu_Click()
  view.start_month = view.start_month - 1
  If (view.start_month < 1) Then
    view.start_month = 12
    view.start_year = view.start_year - 1
  End If
  update_entry_tabs
End Sub

Private Sub recent_file_menu_Click(Index As Integer)
  Dim s As String
  
  s = recent_file_menu(Index).Caption
  
  open_dialog.FileName = Dir(s)
  open_dialog.InitDir = strip_filename(s)
  start_dir = strip_filename(s)
  ChDir start_dir
  ChDrive get_drive(s)
  Call open_the_file(False)
  
End Sub

Private Sub save_as_menu_Click()
  Dim answer
  
  On Error GoTo errorh
  
  With open_dialog
    .Filter = "Check2Check files | *.c2c"
    .ShowSave
    start_dir = CurDir()
    If (.FileName <> "") Then
      ' We have a filename so see if it exists
      If Dir(open_dialog.FileName) <> "" Then
        ' File exists so see if we should overwrite it
        If Dir(open_dialog.FileName) <> "" Then
          answer = MsgBox("File Exists, overwrite?", vbYesNoCancel + vbQuestion, "overwrite?")
          If answer = vbYes Then
            Kill (open_dialog.FileName)
          Else
            Exit Sub
          End If
        End If
        
      End If
      
      ' We have a filename so strip off the extension
      j = InStr(1, .FileName, ".")
      If (j > 0) Then
        data.db_name = Left(.FileName, j - 1)
      End If
    Caption = data.db_name
    write_database
    changed_flag = False
    update_caption
    End If
  End With
  
  Save_Recent_Files
  
  Exit Sub
  
errorh:
  
End Sub

Private Sub previous_month_button_Click()
  previous_month_menu_Click
End Sub

Private Sub save_button_Click()
  save_menu_Click
End Sub

Private Sub save_menu_Click()
  If (data.db_name = "") Then
    save_as_menu_Click
  Else
    update_notes
    write_database
    changed_flag = False
    update_caption
    
    Save_Recent_Files
  End If
End Sub

Private Sub transactions_menu_Click()
  ' Display the entry_grid
  If (notes_box.Visible) Then
    notes_box.Visible = False
    entry_grid.Visible = True
    
    notes_radio.Value = False
    transactions_radio.Value = True
  End If
End Sub

Private Sub transactions_radio_Click()
  transactions_menu_Click
  update_notes
End Sub

Sub txtEdit_KeyPress(KeyAscii As Integer)
    ' Delete returns to get rid of beep.
    If KeyAscii = Val(vbCr) Then KeyAscii = 0
End Sub

Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode entry_grid, txtEdit, KeyCode, Shift
    changed_flag = True
    update_caption
End Sub

Sub update_caption()
    If changed_flag = True Then
        If Right(Caption, 1) <> "*" Then
            Caption = main_form.Caption + " *"
        End If
    Else
        Caption = data.db_name + " - Check2Check"
    End If
End Sub

Sub insert_row_into_database()
  Dim r
   
  entry_grid.Redraw = False
   
  On Error GoTo errorh
  
  r = -1
  
  With entry_grid
    ' Insert the current row into the database at the right place
    this.day = .TextMatrix(.Row, DATE_COL)
    this.month = view.current_month
    this.year = view.current_year
    this.name = .TextMatrix(.Row, NAME_COL)
    this.amount = .TextMatrix(.Row, AMOUNT_COL)
    
    ' See if it belongs after any other ones in the same day
    If (.Row > 1) Then
      ' We are not on the first row
      If (this.day = .TextMatrix(.Row - 1, DATE_COL)) Then
        ' Yes there is another record before this one with the same day
        ' Get the record number of the previous one
        r = .TextMatrix(.Row - 1, THIS_COL)  ' r contains the record number to insert after
      End If
    End If
    ' Insert the record now
    If (r = "") Then r = -1
    Call insert_record(r)   ' If r=-1 then insert it at the proper date, if r >= 0 then insert it after that record
  .Redraw = True
  End With
  Exit Sub

errorh:
  MsgBox "Invalid number 2" + Error
  entry_grid.Redraw = True
End Sub


Sub entry_grid_GotFocus()
    Dim rec
    
    On Error GoTo errorh
    If txtEdit.Visible = False Then Exit Sub
    entry_grid = txtEdit
    txtEdit.Visible = False
        
    
    ' We just entered data so now see if this is on a new record or exising record
    With entry_grid
      last_row = .Row
      last_col = .Col
      If ((.Col = NAME_COL) And (.TextMatrix(.Row, AMOUNT_COL) = "")) Then .TextMatrix(.Row, AMOUNT_COL) = "0"
      
      If (.TextMatrix(.Row, THIS_COL) = "") Then
        ' We have a new record
        insert_row_into_database
      Else
        ' We have an existing record so update the data
        rec = .TextMatrix(.Row, THIS_COL)
        If (.Col = NAME_COL) Then db(rec).name = txtEdit
        If (.Col = AMOUNT_COL) Then db(rec).amount = txtEdit
        If (.Col = OVERRIDE_AMOUNT_COL) Then db(rec).override_amount = txtEdit
        End If
        
      rec = .Row
    End With
    
    Call process
    
    entry_grid.Row = rec
    If last_col = NAME_COL Then
        entry_grid.Col = AMOUNT_COL
        entry_grid.Row = last_row
    Else
        insert_menu_Click
        If last_row + 1 < entry_grid.Rows Then
            entry_grid.Row = last_row + 1
        Else
            entry_grid.Row = last_row
        End If
        entry_grid.Col = NAME_COL
    End If
    
    If (entry_grid.Row > 10) Then
      entry_grid.TopRow = entry_grid.Row - 7
    Else
      entry_grid.TopRow = 1
    End If
    
  Exit Sub

errorh:
  MsgBox "Invalid number 1 entered"
End Sub


Sub entry_grid_LeaveCell()
    If txtEdit.Visible = False Then Exit Sub
    entry_grid = txtEdit
    txtEdit.Visible = False
End Sub


Private Sub shuffle_up_entry_grid(n As Integer)
  'entry_grid.Rows = entry_grid.Rows + 1
  entry_grid.RemoveItem (n)
  entry_grid.Row = entry_grid.Row - 1
End Sub


Private Sub shuffle_down_entry_grid(n As Integer)
  With entry_grid
    .AddItem "", n
    .TextMatrix(n, DATE_COL) = .TextMatrix(n - 1, DATE_COL)
    .TextMatrix(n, DAY_COL) = .TextMatrix(n - 1, DAY_COL)
  End With
End Sub


Private Sub put_in_amount_cell(ByVal r As Integer, ByVal C As Integer, ByVal amt As Single)
      entry_grid.Row = r
      entry_grid.Col = C
      entry_grid.CellFontBold = True
      
      entry_grid.TextMatrix(r, C) = currency_s(amt)
      
      If (amt < 0) Then
        ' Change the color of the cell to red
        entry_grid.CellForeColor = vbRed
      Else
        If (amt > 0) Then
          ' Change the color of the cell to blue
          entry_grid.CellForeColor = vbBlue
        Else
          entry_grid.CellForeColor = vbGreen
        End If
      End If
End Sub

Private Sub shuffle_down_table_image(Index As Integer)
  For j = MAX_RECORDS_IN_MONTH To Index Step -1
    table_image.table(j) = table_image.table(j - 1)
  Next j
End Sub

Private Sub stuff_this_in_entry_grid()
  Dim C, r, amt
  
  C = entry_grid.Col
  r = entry_grid.Row
  
  ' Put this in the entry grid
  ' Scan down the entry grid till we find the place where it goes
  For I = entry_grid.Rows - 1 To 1 Step -1
    If (entry_grid.TextMatrix(I, DATE_COL) = this.day) Then
      ' We found a date that matches
      ' See if there is already an entry there
      If (entry_grid.TextMatrix(I, THIS_COL) <> "") Then
        ' We have an entry so shuffle down
        I = I + 1
        shuffle_down_entry_grid (I)
      End If
      
      entry_grid.Row = I
      entry_grid.Col = DATE_COL
      entry_grid.CellFontBold = True
      entry_grid.TextMatrix(I, THIS_COL) = this.this
      entry_grid.TextMatrix(I, PREV_COL) = this.previous
      entry_grid.TextMatrix(I, NEXT_COL) = this.next
      entry_grid.TextMatrix(I, NAME_COL) = this.name
      
      ' Put in the exclude graphic
      entry_grid.Row = I
      entry_grid.Col = EXCLUDE_COL
      If (this.exclude) Then
        Set entry_grid.CellPicture = LoadPicture("c:\check2check\check.bmp")
      Else
        Set entry_grid.CellPicture = LoadPicture("c:\check2check\blank_check.bmp")
      entry_grid.CellPictureAlignment = flexAlignCenterCenter
      End If
      
      ' Put in the override graphic
      entry_grid.Row = I
      entry_grid.Col = OVERRIDE_COL
      If (this.override) Then
        Set entry_grid.CellPicture = LoadPicture("c:\check2check\check.bmp")
      Else
        Set entry_grid.CellPicture = LoadPicture("c:\check2check\blank_check.bmp")
      entry_grid.CellPictureAlignment = flexAlignCenterCenter
      End If
      
      ' Put in the paid graphic
      entry_grid.Row = I
      entry_grid.Col = PAID_COL
      If (this.paid) Then
        Set entry_grid.CellPicture = LoadPicture("c:\check2check\check.bmp")
      Else
        Set entry_grid.CellPicture = LoadPicture("c:\check2check\blank_check.bmp")
      entry_grid.CellPictureAlignment = flexAlignCenterCenter
      End If
      
      Call put_in_amount_cell(I, AMOUNT_COL, this.amount)
      Call put_in_amount_cell(I, BALANCE_COL, this.balance)
      Call put_in_amount_cell(I, OVERRIDE_AMOUNT_COL, this.override_amount)
    
      Exit For
    End If
  Next I
    
  entry_grid.Col = C
  entry_grid.Row = r
  
End Sub


Private Sub calculate()
  Dim balance As Single
  
  balance = 0
  
  ' Start from the beginning and calculate the balances
  If (data.number_of_records > 0) Then
    ' We have at least one record
    data.current = data.first
    get_record  ' Get the first record
    
    ' See if this record should be included as normal
    If (Not db(data.current).exclude) And (Not db(data.current).override) Then
      balance = db(data.current).amount
      db(data.current).balance = balance
    End If
    
    ' See if we have exclude
    If (db(data.current).exclude) Then
      ' We have exclude
      balance = 0
      db(data.current).balance = balance
    End If
    
    ' See if we have override
    If (db(data.current).override) Then
      ' We have override
      balance = db(data.current).override_amount
      db(data.current).balance = balance
    End If
 
    
    
    While get_next_record
      ' Loop though all the remaining records and do the balance
      If (Not db(data.current).exclude) And (Not db(data.current).override) Then
        balance = balance + db(data.current).amount
        db(data.current).balance = balance
      End If
    
      ' See if we have exclude
      If (db(data.current).exclude) Then
        ' We have exclude
        db(data.current).balance = balance
      End If
    
      ' See if we have override
      If (db(data.current).override) Then
         ' We have override
         balance = db(data.current).override_amount
         db(data.current).balance = balance
      End If
    
    Wend
  End If
  
End Sub

Private Sub stuff_this_in_table_image()
  Dim C, r, amt
  
  ' Put this in the entry grid
  ' Scan down the entry grid till we find the place where it goes
  For I = table_image.last To 1 Step -1
    If (table_image.table(I).day = this.day) Or _
       (this.day > table_image.table(table_image.last).day) Then
      ' We found a date that matches
      ' See if there is already an entry there
      If (table_image.table(I).this > -2) Then
        ' We have an entry so shuffle down
        I = I + 1
        shuffle_down_table_image (I)
        table_image.last = table_image.last + 1
      End If
      
      'table_image.last = table_image.last + 1
      table_image.table(I).day = this.day
      table_image.table(I).this = this.this
      table_image.table(I).prev = this.previous
      table_image.table(I).next = this.next
      view.records_in_month = view.records_in_month + 1
      Exit For
    End If
  Next I

End Sub

Private Sub stuff_the_entry_grid()
  entry_grid.Redraw = False

  With entry_grid
  
    .Rows = MAX_RECORDS_IN_MONTH  ' Start with the max number
  
    ' Clear out the entry
    calendar.day = 1
    calendar.year = view.current_year
    calendar.month = view.current_month
    .Clear
    .TextMatrix(0, DATE_COL) = "Date"
    .TextMatrix(0, DAY_COL) = "Day"
    .TextMatrix(0, THIS_COL) = "This"
    .TextMatrix(0, PREV_COL) = "Prev"
    .TextMatrix(0, NEXT_COL) = "Next"
    .TextMatrix(0, NAME_COL) = "Name"
    .TextMatrix(0, PAID_COL) = "Done"
    .TextMatrix(0, AMOUNT_COL) = "Amount"
    .TextMatrix(0, EXCLUDE_COL) = "Excl"
    .TextMatrix(0, BALANCE_COL) = "Bal"
    .TextMatrix(0, OVERRIDE_COL) = "O/R"
    .TextMatrix(0, OVERRIDE_AMOUNT_COL) = "O/R Bal"

    For I = 1 To MAX_RECORDS_IN_MONTH
      If (table_image.table(I).day = 0) Then Exit For
      .Row = I
      .Col = DATE_COL
      .CellFontBold = True
      .TextMatrix(I, DATE_COL) = table_image.table(I).day
    
      calendar.day = table_image.table(I).day
      .TextMatrix(I, DAY_COL) = DAY_STRINGS(calendar.DayOfWeek)
      
      data.current = table_image.table(I).this
      If (data.current >= 0) Then
        ' We have a valid record number
        get_record
        .TextMatrix(I, THIS_COL) = this.this
        .TextMatrix(I, PREV_COL) = this.previous
        .TextMatrix(I, NEXT_COL) = this.next
        .TextMatrix(I, NAME_COL) = this.name
      
      
        Call put_in_amount_cell(I, AMOUNT_COL, this.amount)
        Call put_in_amount_cell(I, BALANCE_COL, this.balance)
        Call put_in_amount_cell(I, OVERRIDE_AMOUNT_COL, this.override_amount)
    
        .Row = I
        
        ' Put in the exclude graphic
        If (this.exclude) Then
          .Col = EXCLUDE_COL
          .CellFontName = "wingdings 2"
          .CellFontSize = 14
          .Text = "P"
        End If
      
        ' Put in the override graphic
        If (this.override) Then
          .Col = OVERRIDE_COL
          .CellFontName = "wingdings 2"
          .CellFontSize = 14
          .Text = "P"
        End If
      
        ' Put in the paid graphic
        If (this.paid) Then
          .Col = PAID_COL
          .CellFontName = "wingdings 2"
          .CellFontSize = 14
          .Text = "P"
        End If
    
    
      End If
            
    Next I
    .Rows = I
    .Redraw = True
  End With
End Sub


Private Sub process()
  Dim end_balance, start_balance
  
  first_label.Caption = data.first
  last_label.Caption = data.last
  total_label.Caption = data.number_of_records
  
  main_form.MousePointer = 11
  
  table_image.last = -1
  'entry_grid.Redraw = False
  
  ' Process and update
  calculate
  
  ' Find the year of the tab we are on
  view.current_month = view.start_month + entry_tab.Tab
  view.current_year = view.start_year
  If (view.current_month > 12) Then
    ' We rolled over into the next year
    view.current_month = view.current_month - 12
    view.current_year = view.current_year + 1
  End If
  
  
  
  On Error GoTo done
  calendar.day = 1
  calendar.year = view.current_year
  calendar.month = view.current_month
  For I = 1 To 100
    table_image.table(I).day = I
    table_image.table(I + 1).day = 0
    table_image.table(I).this = -2
    
    calendar.day = calendar.day + 1
  Next I
  
done:
  table_image.last = I
  view.number_of_days = I
  
  view.records_in_month = 0
  view.last_balance = 0
  beginning_balance_box.Text = currency_s(0)
  ending_balance_box.Text = currency_s(0)
  
  ' Find the first record that matches it
  If (find_first(view.current_month, view.current_year)) Then
    ' We have at least one record
    ' Put record in display
    'stuff_this_in_entry_grid
    stuff_this_in_table_image
    
    beginning_balance_box.Text = currency_s(this.balance - this.amount)
    end_balance = this.balance
    While (find_next())
      'stuff_this_in_entry_grid
      stuff_this_in_table_image
      
      end_balance = this.balance
      ending_balance_box.Text = currency_s(end_balance)
    Wend
  Else
    ' No records found in month
    beginning_balance_box.Text = currency_s(view.last_balance)
  End If
  ending_balance_box.Text = currency_s(view.last_balance)  ' end_balance
  
  'entry_grid.Redraw = True
  stuff_the_entry_grid
  
  ' Now display the notes
  notes_box.Text = ""
  For I = 0 To MAX_NOTES
    If (notes(I).month = view.current_month) And _
       (notes(I).year = view.current_year) Then
      ' We have found the note that belongs to this month
      notes_box.Text = notes(I).s
      Exit For
    End If
    
  Next I
  
  main_form.MousePointer = 1
  
  show_buttons
  
  If (Val(beginning_balance_box.Text) > 0) Then
    beginning_balance_box.ForeColor = vbBlue
  Else
    If (Val(beginning_balance_box.Text) < 0) Then
      beginning_balance_box.ForeColor = vbRed
    Else
      beginning_balance_box.ForeColor = vbGreen
    End If
  End If
  
  If (Val(ending_balance_box.Text) > 0) Then
    ending_balance_box.ForeColor = vbBlue
  Else
    If (Val(ending_balance_box.Text) < 0) Then
      ending_balance_box.ForeColor = vbRed
    Else
      ending_balance_box.ForeColor = vbGreen
    End If
  End If
  
  
  update_summary
End Sub

Function currency_s(ByVal v As Single) As String
  currency_s = Format(v, "###,###,##0.00 ")
End Function

Private Sub update_summary()
  Dim s
  
  ' Update the summary box
  s = "Total transactions: " + Str(data.number_of_records)
  s = s + Chr(10) + " Total notes: " + Str(data.number_of_notes)
  s = s + Chr(10) + " Current file: " + data.db_name
  s = s + Chr(10) + " Max transactions allowed: " + Format(MAX_DATA_TABLE, "##,###,###,###")
  s = s + Chr(10) + " Max transactions allowed in a month: " + Format(MAX_RECORDS_IN_MONTH, "###,###")
  
End Sub

Private Sub show_buttons()
  ' Show or hide the buttons
  If (view.records_in_month = 0) Then
    delete_menu.Enabled = False
    cut_menu.Enabled = False
    copy_menu.Enabled = False
    cut_month_menu.Enabled = False
    copy_month_menu.Enabled = False
    
    delete_button.Visible = False
    cut_button.Visible = False
    copy_button.Visible = False
  Else
    delete_menu.Enabled = True
    cut_menu.Enabled = True
    copy_menu.Enabled = True
    cut_month_menu.Enabled = True
    copy_month_menu.Enabled = True
    
    delete_button.Visible = True
    cut_button.Visible = True
    copy_button.Visible = True
  End If
End Sub

Sub Get_Recent_Files()
  Dim key, I
  Dim s As String
  
  ' Get recent file strings
  For I = 0 To 3
    key = "RecentFile" & I + 1
    s = GetSetting("Check 2 Check", "Settings", key, "Not Used")
    If s <> "Not Used" Then
      ' Update the recent files
      main_form.recent_file_menu(I).Visible = True
      main_form.recent_file_menu(I).Caption = s + ".c2c"
    Else
      main_form.recent_file_menu(I).Visible = False
    End If
  Next I

End Sub

Sub Save_Recent_Files()
  Dim s(4) As String
  
  s(0) = GetSetting("Check 2 Check", "Settings", "RecentFile1", "Not Used")
  s(1) = GetSetting("Check 2 Check", "Settings", "RecentFile2", "Not Used")
  s(2) = GetSetting("Check 2 Check", "Settings", "RecentFile3", "Not Used")
  s(3) = GetSetting("Check 2 Check", "Settings", "RecentFile4", "Not Used")

  ' See if the current one is in the list
  If s(3) = data.db_name Then
    ' Shuffle down
    s(3) = s(2)
    s(2) = s(1)
    s(1) = s(0)
    s(0) = data.db_name
  Else
    If s(2) = data.db_name Then
      s(2) = s(1)
      s(1) = s(0)
      s(0) = data.db_name
    Else
      If s(1) = data.db_name Then
        s(1) = s(0)
        s(0) = data.db_name
      Else
        If s(0) <> data.db_name Then
          s(3) = s(2)
          s(2) = s(1)
          s(1) = s(0)
          s(0) = data.db_name
        End If
      End If
    End If
  End If
  
  SaveSetting "Check 2 Check", "Settings", "RecentFile1", s(0)
  SaveSetting "Check 2 Check", "Settings", "RecentFile2", s(1)
  SaveSetting "Check 2 Check", "Settings", "RecentFile3", s(2)
  SaveSetting "Check 2 Check", "Settings", "RecentFile4", s(3)

  Get_Recent_Files
  'If (s(0) <> "") Then save_to_recent_docs (s(0))
End Sub

