VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form preferences_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   5565
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6705
   Icon            =   "prefer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4980
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   8784
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "prefer.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Paste  Month"
      TabPicture(1)   =   "prefer.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "Cleared Column"
         Height          =   795
         Index           =   2
         Left            =   -72420
         TabIndex        =   27
         Top             =   2640
         Width           =   1935
         Begin VB.ComboBox paste_month_combo 
            Height          =   315
            Index           =   5
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   300
            Width           =   1515
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Check Number Column"
         Height          =   795
         Index           =   1
         Left            =   -74520
         TabIndex        =   25
         Top             =   2640
         Width           =   1935
         Begin VB.ComboBox paste_month_combo 
            Height          =   315
            Index           =   4
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   300
            Width           =   1515
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Status Column"
         Height          =   1875
         Index           =   0
         Left            =   -74520
         TabIndex        =   16
         Top             =   660
         Width           =   4035
         Begin VB.ComboBox paste_month_combo 
            Height          =   315
            Index           =   3
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1320
            Width           =   1755
         End
         Begin VB.ComboBox paste_month_combo 
            Height          =   315
            Index           =   2
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   960
            Width           =   1755
         End
         Begin VB.ComboBox paste_month_combo 
            Height          =   315
            Index           =   1
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   1755
         End
         Begin VB.ComboBox paste_month_combo 
            Height          =   315
            Index           =   0
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Skip -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   600
            TabIndex        =   23
            Top             =   1380
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Pending -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   21
            Top             =   1020
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Done -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   540
            TabIndex        =   19
            Top             =   660
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Blank -->"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   17
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4185
         Left            =   300
         TabIndex        =   3
         Top             =   480
         Width           =   4455
         Begin VB.CheckBox play_sounds_check 
            Caption         =   "Play sounds"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   3060
            Width           =   1935
         End
         Begin VB.CheckBox auto_load_last_file_check 
            Caption         =   "Auto load last file"
            Height          =   375
            Left            =   120
            TabIndex        =   33
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox control_number_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   135
            TabIndex        =   31
            Text            =   "Text1"
            Top             =   3690
            Width           =   1275
         End
         Begin VB.CommandButton change_password_button 
            Caption         =   "Change Password"
            Height          =   435
            Left            =   2640
            TabIndex        =   30
            Top             =   240
            Width           =   1635
         End
         Begin VB.CheckBox save_recovery_file_check 
            Caption         =   "Save recovery file"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   2520
            Width           =   1695
         End
         Begin VB.CheckBox prompt_for_move_check 
            Caption         =   "Prompt for move"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   1815
         End
         Begin VB.CheckBox auto_insert_check 
            Caption         =   "Auto insert new line on transaction entry"
            Height          =   315
            Left            =   120
            TabIndex        =   13
            Top             =   1020
            Width           =   3615
         End
         Begin VB.CheckBox show_splash_screen_check 
            Caption         =   "Show splash screen at startup"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1500
            Width           =   2895
         End
         Begin VB.CheckBox show_name_colors_check 
            Caption         =   "Show names in colors"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   2175
         End
         Begin VB.CheckBox show_override_columns_check 
            Caption         =   "Show override columns at startup"
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   1260
            Width           =   2895
         End
         Begin VB.CheckBox auto_check_done_check 
            Caption         =   "Auto check ""Done"" on amount entry"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   540
            Width           =   3255
         End
         Begin VB.CheckBox prompt_for_delete_check 
            Caption         =   "Prompt for delete"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   300
            Width           =   3135
         End
         Begin VB.CheckBox prompt_for_paste_notes_check 
            Caption         =   "Prompt for paste notes"
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1800
            Width           =   2655
         End
         Begin VB.CheckBox auto_negative_numbers_check 
            Caption         =   "Auto negative numbers in amount column"
            Height          =   225
            Left            =   120
            TabIndex        =   6
            Top             =   2100
            Width           =   3615
         End
         Begin VB.CheckBox auto_check_done_on_check_check 
            Caption         =   "Auto check ""Done"" on check entry"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   2280
            Width           =   3615
         End
         Begin VB.ComboBox notes_font_size_combo 
            Height          =   315
            ItemData        =   "prefer.frx":0902
            Left            =   135
            List            =   "prefer.frx":0904
            TabIndex        =   4
            Text            =   "notes_font_size_combo"
            Top             =   3330
            Width           =   645
         End
         Begin VB.Label Label3 
            Caption         =   "Reference Number"
            Height          =   240
            Left            =   1530
            TabIndex        =   32
            Top             =   3735
            Width           =   1905
         End
         Begin VB.Label Label1 
            Caption         =   "Font size for Notes"
            Height          =   255
            Left            =   810
            TabIndex        =   15
            Top             =   3405
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   5340
      TabIndex        =   1
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Height          =   555
      Left            =   5340
      Picture         =   "prefer.frx":0906
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "preferences_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub update_language()
  prompt_for_move_check.Caption = words(PROMPT_FOR_MOVE_N)
  prompt_for_delete_check.Caption = words(PROMPT_FOR_DELETE_N)
  show_name_colors_check.Caption = words(SHOW_NAMES_IN_COLORS_N)
  auto_check_done_check.Caption = words(AUTO_CHECK_DONE_ON_AMOUNT_N)
  auto_insert_check.Caption = words(AUTO_INSERT_NEW_LINE_N)
  show_override_columns_check.Caption = words(SHOW_OVERRIDE_COLUMNS_N)
  show_splash_screen_check.Caption = words(SHOW_SPLASH_SCREEN_N)
  prompt_for_paste_notes_check.Caption = words(PROMPT_FOR_PASTE_NOTES_N)
  auto_negative_numbers_check.Caption = words(AUTO_NEGATIVE_NUMBERS_N)
  auto_check_done_on_check_check.Caption = words(AUTO_CHECK_DONE_ON_CHECK_N)
  save_recovery_file_check.Caption = words(SAVE_RECOVERY_FILE_N)
  auto_load_last_file_check.Caption = words(AUTO_LOAD_LAST_FILE_N)
  play_sounds_check.Caption = words(PLAY_SOUNDS_N)
  
  change_password_button.Caption = words(CHANGE_PASSWORD_N)
  CancelButton.Caption = words(CANCEL_N)
  
  preferences_form.Caption = words(PREFERENCES_N)
  Label1.Caption = words(FONT_SIZE_NOTES_N)
  Label3.Caption = words(REFERENCE_NUMBER_N)
  SSTab1.TabCaption(0) = words(GENERAL_N)
  SSTab1.TabCaption(1) = words(PASTE_MONTH_N)
  
  Frame2(0).Caption = words(STATUS_COLUMN_N)
  Frame2(1).Caption = words(CHECK_NUMBER_COLUMN_N)
  Frame2(2).Caption = words(CLEARED_COLUMN_N)
  
  Label2(0).Caption = words(BLANK_N)
  Label2(1).Caption = words(DONE_N)
  Label2(2).Caption = words(PENDING_N)
  Label2(3).Caption = words(SKIP_N)
  
End Sub


Private Sub CancelButton_Click()
  Hide
End Sub

Private Sub change_password_button_Click()
  password_form.execute (1)
End Sub

Private Sub Form_Activate()
  Dim i
  
  update_language
  
  auto_insert_check.Value = 0
  If preferences.auto_insert Then auto_insert_check.Value = 1
  
  prompt_for_move_check.Value = 0
  If preferences.prompt_for_move Then prompt_for_move_check.Value = 1
  
  show_splash_screen_check.Value = 0
  If preferences.show_splash_screen Then show_splash_screen_check.Value = 1

  show_name_colors_check.Value = 0
  If preferences.show_name_colors Then show_name_colors_check.Value = 1

  show_override_columns_check.Value = 0
  If preferences.show_override_columns Then show_override_columns_check.Value = 1

  auto_check_done_check.Value = 0
  If preferences.auto_check_done Then auto_check_done_check.Value = 1

  auto_check_done_on_check_check.Value = 0
  If preferences.auto_check_done_on_check Then auto_check_done_on_check_check.Value = 1
  
  prompt_for_delete_check.Value = 0
  If preferences.prompt_for_delete Then prompt_for_delete_check.Value = 1

  prompt_for_paste_notes_check.Value = 0
  If preferences.prompt_for_paste_notes Then prompt_for_paste_notes_check.Value = 1
  
  auto_negative_numbers_check.Value = 0
  If preferences.auto_negative_numbers Then auto_negative_numbers_check.Value = 1

  save_recovery_file_check.Value = 0
  If preferences.save_recovery_file Then save_recovery_file_check.Value = 1

  auto_load_last_file_check.Value = 0
  If preferences.auto_load_last_file Then auto_load_last_file_check.Value = 1

  play_sounds_check.Value = 0
  If preferences.play_sounds Then play_sounds_check.Value = 1

  control_number_box.Text = Val(GetSetting("Check 2 Check", "Settings", "Reference_number", "0"))
  
  notes_font_size_combo.Clear
  notes_font_size_combo.AddItem "8", 0
  notes_font_size_combo.AddItem "10", 1
  notes_font_size_combo.AddItem "12", 2
  notes_font_size_combo.AddItem "14", 3
  notes_font_size_combo.AddItem "16", 4
  notes_font_size_combo.AddItem "18", 5
  notes_font_size_combo.AddItem "20", 6
  notes_font_size_combo.Text = Str(preferences.notes_font_size)
  
  ' Fill up the paste month status combo boxes
  For i = 0 To 3
    paste_month_combo(i).Clear
    paste_month_combo(i).AddItem words(BLANK_N), 0  '"Blank", 0
    paste_month_combo(i).AddItem words(DONE_N), 1  '"Done", 1
    paste_month_combo(i).AddItem words(PENDING_N), 2  '"Pending", 2
    paste_month_combo(i).AddItem words(SKIP_N), 3  '"Skip", 3
    paste_month_combo(i).ListIndex = preferences.paste_month(i)
  Next i
  
  ' Fill up the paste month check number combo boxes
  paste_month_combo(PREF_MONTH_CHECK_NUMBER_INDEX).Clear
  paste_month_combo(PREF_MONTH_CHECK_NUMBER_INDEX).AddItem words(SAME_N), 0  '"Same", 0
  paste_month_combo(PREF_MONTH_CHECK_NUMBER_INDEX).AddItem words(BLANK_N), 1  '"Blank", 1
  paste_month_combo(PREF_MONTH_CHECK_NUMBER_INDEX).ListIndex = preferences.paste_month(PREF_MONTH_CHECK_NUMBER_INDEX)
  
  ' Fill up the paste month cleared combo boxes
  paste_month_combo(PREF_MONTH_CLEARED_INDEX).Clear
  paste_month_combo(PREF_MONTH_CLEARED_INDEX).AddItem words(SAME_N), 0  '"Same", 0
  paste_month_combo(PREF_MONTH_CLEARED_INDEX).AddItem words(BLANK_N), 1  '"Blank", 1
  paste_month_combo(PREF_MONTH_CLEARED_INDEX).ListIndex = preferences.paste_month(PREF_MONTH_CLEARED_INDEX)
  
End Sub

Private Sub OKButton_Click()
  Dim v
  
  ' Save the options
  preferences.auto_insert = auto_insert_check.Value
  preferences.prompt_for_move = prompt_for_move_check.Value
  preferences.show_splash_screen = show_splash_screen_check.Value
  preferences.show_name_colors = show_name_colors_check.Value
  preferences.show_override_columns = show_override_columns_check.Value
  preferences.auto_check_done = auto_check_done_check.Value
  preferences.auto_check_done_on_check = auto_check_done_on_check_check.Value
  preferences.prompt_for_delete = prompt_for_delete_check.Value
  preferences.auto_negative_numbers = auto_negative_numbers_check.Value
  preferences.save_recovery_file = save_recovery_file_check.Value
  preferences.auto_load_last_file = auto_load_last_file_check.Value
  preferences.play_sounds = play_sounds_check.Value
  
  v = Val(notes_font_size_combo.Text)
  If (v > 6) Then preferences.notes_font_size = v
  
  ' Save the paste month options
  preferences.paste_month(0) = paste_month_combo(0).ListIndex
  preferences.paste_month(1) = paste_month_combo(1).ListIndex
  preferences.paste_month(2) = paste_month_combo(2).ListIndex
  preferences.paste_month(3) = paste_month_combo(3).ListIndex
  preferences.paste_month(4) = paste_month_combo(4).ListIndex
  preferences.paste_month(5) = paste_month_combo(5).ListIndex
  
  SaveSetting "Check 2 Check", "Settings", "Reference_number", Val(control_number_box.Text)

  Hide
End Sub

