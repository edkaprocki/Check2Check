VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form main_form 
   Caption         =   "Check 2 Check"
   ClientHeight    =   5970
   ClientLeft      =   2295
   ClientTop       =   1200
   ClientWidth     =   9465
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MouseIcon       =   "main.frx":08CA
   ScaleHeight     =   5970
   ScaleWidth      =   9465
   WindowState     =   2  'Maximized
   Begin VB.CommandButton todays_date_button 
      Height          =   600
      Left            =   5375
      Picture         =   "main.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Go To Today's Date"
      Top             =   360
      Width           =   465
   End
   Begin VB.CommandButton center_month_button 
      Height          =   375
      Left            =   5280
      Picture         =   "main.frx":0D80
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Center Month"
      Top             =   990
      Width           =   600
   End
   Begin VB.CommandButton next_year_button 
      Height          =   375
      Left            =   7560
      Picture         =   "main.frx":10F9
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Next Year"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton previous_year_button 
      Height          =   375
      Left            =   3240
      Picture         =   "main.frx":119D
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Previous Year"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   45
      TabIndex        =   58
      Top             =   360
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton top_button 
      Height          =   375
      Left            =   5870
      Picture         =   "main.frx":1241
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Top of Month"
      Top             =   990
      Width           =   815
   End
   Begin VB.CommandButton bottom_button 
      Height          =   375
      Left            =   4440
      Picture         =   "main.frx":162F
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Bottom of Month"
      Top             =   990
      Width           =   815
   End
   Begin VB.CommandButton next_month_top_button 
      Height          =   375
      Left            =   6840
      Picture         =   "main.frx":1A24
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Next Month Top"
      Top             =   990
      Width           =   975
   End
   Begin VB.CommandButton previous_month_bottom_button 
      Height          =   375
      Left            =   3240
      Picture         =   "main.frx":1DF3
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Previous Month Bottom"
      Top             =   990
      Width           =   975
   End
   Begin VB.Frame characters_left_frame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Characters Left"
      Height          =   690
      Left            =   4320
      TabIndex        =   52
      Top             =   2655
      Width           =   2130
      Begin VB.Label characters_left_label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   285
         Left            =   225
         TabIndex        =   53
         Top             =   315
         Width           =   1635
      End
   End
   Begin VB.TextBox misc_notes_box 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7200
      MaxLength       =   65000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   51
      Text            =   "main.frx":21BF
      Top             =   4230
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame reference_number_frame 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reference Number"
      Height          =   2265
      Left            =   6795
      TabIndex        =   44
      Top             =   1935
      Width           =   2535
      Begin VB.CommandButton reference_number_button 
         Caption         =   "New Ref Number"
         Height          =   375
         Left            =   225
         TabIndex        =   48
         ToolTipText     =   "Insert Control Number"
         Top             =   1665
         Width           =   2085
      End
      Begin VB.CommandButton reference_number_last_button 
         Caption         =   "Last Ref Number"
         Height          =   375
         Left            =   225
         TabIndex        =   47
         ToolTipText     =   "Insert Control Number"
         Top             =   1170
         Width           =   2085
      End
      Begin VB.CommandButton last_name_button 
         Caption         =   "Insert Last Name"
         Height          =   375
         Left            =   225
         TabIndex        =   46
         ToolTipText     =   "Insert Control Number"
         Top             =   675
         Width           =   2085
      End
      Begin VB.Label reference_label 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   285
         Left            =   540
         TabIndex        =   45
         Top             =   360
         Width           =   1500
      End
   End
   Begin VB.Timer auto_save_timer 
      Enabled         =   0   'False
      Left            =   6255
      Top             =   2115
   End
   Begin VB.CommandButton next_month_button 
      Caption         =   "Next Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   5880
      Picture         =   "main.frx":21D0
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Next Month"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton previous_month_button 
      Caption         =   "Previous Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3720
      Picture         =   "main.frx":2693
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Previous Month"
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cardtrak_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":2B52
      DownPicture     =   "main.frx":2C54
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6660
      Picture         =   "main.frx":2D56
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "View Cardtrak"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton quick_view_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":30D4
      DownPicture     =   "main.frx":31D6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6300
      Picture         =   "main.frx":32D8
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "View Quick Accounts"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton view_summary_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":365B
      DownPicture     =   "main.frx":375D
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5940
      Picture         =   "main.frx":385F
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "View Summary"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":3BD9
      DownPicture     =   "main.frx":3CDB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5580
      Picture         =   "main.frx":3DDD
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "View Tags"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton view_balance_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":4157
      DownPicture     =   "main.frx":4259
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5220
      Picture         =   "main.frx":435B
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "View Balances"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton filter_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":46D9
      DownPicture     =   "main.frx":47DB
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4860
      Picture         =   "main.frx":48DD
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Filter Transactions"
      Top             =   0
      Width           =   315
   End
   Begin VB.Frame balance_frame 
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   7020
      TabIndex        =   22
      Top             =   0
      Width           =   2475
      Begin VB.TextBox checkbook_balance_box 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   900
         TabIndex        =   31
         Text            =   "Text2"
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox ending_balance_box 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   900
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox beginning_balance_box 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   270
         Left            =   900
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label checkbook_label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Checkbook"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   780
         Width           =   825
      End
      Begin VB.Label ending_balance_label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ending"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   495
      End
      Begin VB.Label beginning_balance_label 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Beginning"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.CommandButton calculator_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":4C67
      DownPicture     =   "main.frx":4D69
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4500
      Picture         =   "main.frx":4E6B
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Show Calculator"
      Top             =   0
      Width           =   315
   End
   Begin MSComDlg.CommonDialog help_dialog 
      Left            =   120
      Top             =   3300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "c2c.hlp"
   End
   Begin VB.CommandButton undo_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":51FC
      DownPicture     =   "main.frx":52FE
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3660
      Picture         =   "main.frx":5400
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Undo"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton print_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":5932
      DownPicture     =   "main.frx":5A34
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      Picture         =   "main.frx":5B36
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Print"
      Top             =   0
      Width           =   315
   End
   Begin VB.Timer splash_timer 
      Interval        =   15000
      Left            =   600
      Top             =   5280
   End
   Begin MSComDlg.CommonDialog print_dialog 
      Left            =   120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1860
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox notes_box 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Text            =   "main.frx":6068
      Top             =   4185
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton calendar_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":6072
      DownPicture     =   "main.frx":6174
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4140
      Picture         =   "main.frx":6276
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Show Calendar"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton delete_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":6400
      DownPicture     =   "main.frx":6502
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1980
      Picture         =   "main.frx":6604
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Delete"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton insert_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":6706
      DownPicture     =   "main.frx":6808
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      Picture         =   "main.frx":690A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Insert"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton paste_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":6A0C
      DownPicture     =   "main.frx":6B0E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3180
      Picture         =   "main.frx":6C10
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Paste"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton copy_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":6D12
      DownPicture     =   "main.frx":6E14
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2820
      Picture         =   "main.frx":6F16
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Copy"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton cut_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":7018
      DownPicture     =   "main.frx":711A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2460
      Picture         =   "main.frx":721C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cut"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton save_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":731E
      DownPicture     =   "main.frx":7420
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      Picture         =   "main.frx":7522
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton open_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":7624
      DownPicture     =   "main.frx":7726
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   420
      Picture         =   "main.frx":7828
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Open"
      Top             =   0
      Width           =   315
   End
   Begin VB.CommandButton new_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "main.frx":792A
      DownPicture     =   "main.frx":7A2C
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      Picture         =   "main.frx":7B2E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "New"
      Top             =   0
      Width           =   315
   End
   Begin MSComCtl2.MonthView calendar 
      Height          =   2820
      Left            =   1320
      TabIndex        =   15
      Top             =   3480
      Visible         =   0   'False
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      StartOfWeek     =   36765697
      CurrentDate     =   36378
   End
   Begin TabDlg.SSTab entry_tab 
      Height          =   615
      Left            =   855
      TabIndex        =   16
      Top             =   1395
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   1085
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
      TabPicture(0)   =   "main.frx":8060
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Aug"
      TabPicture(1)   =   "main.frx":807C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Sep"
      TabPicture(2)   =   "main.frx":8098
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Oct"
      TabPicture(3)   =   "main.frx":80B4
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Nov"
      TabPicture(4)   =   "main.frx":80D0
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Dec"
      TabPicture(5)   =   "main.frx":80EC
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   "Jan"
      TabPicture(6)   =   "main.frx":8108
      Tab(6).ControlEnabled=   0   'False
      Tab(6).ControlCount=   0
      TabCaption(7)   =   "Feb"
      TabPicture(7)   =   "main.frx":8124
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Mar"
      TabPicture(8)   =   "main.frx":8140
      Tab(8).ControlEnabled=   0   'False
      Tab(8).ControlCount=   0
      TabCaption(9)   =   "Apr"
      TabPicture(9)   =   "main.frx":815C
      Tab(9).ControlEnabled=   0   'False
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "May"
      TabPicture(10)  =   "main.frx":8178
      Tab(10).ControlEnabled=   0   'False
      Tab(10).ControlCount=   0
      TabCaption(11)  =   "Jun"
      TabPicture(11)  =   "main.frx":8194
      Tab(11).ControlEnabled=   0   'False
      Tab(11).ControlCount=   0
   End
   Begin MSComDlg.CommonDialog open_dialog 
      Left            =   0
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid entry_grid 
      Height          =   3525
      Left            =   855
      TabIndex        =   14
      Top             =   1980
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   6218
      _Version        =   393216
      Rows            =   32
      Cols            =   18
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   4194304
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      MergeCells      =   1
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
   Begin VB.Image binder 
      Height          =   43005
      Left            =   45
      Picture         =   "main.frx":81B0
      Top             =   1890
      Width           =   750
   End
   Begin VB.Label misc_notes_label 
      Caption         =   "Misc Notes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1260
      MouseIcon       =   "main.frx":1287C
      MousePointer    =   99  'Custom
      TabIndex        =   50
      ToolTipText     =   "Transactions"
      Top             =   1035
      Width           =   2265
   End
   Begin VB.Label transaction_notes_check_label 
      Alignment       =   2  'Center
      Caption         =   "�"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   2
      Left            =   855
      TabIndex        =   49
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label transaction_notes_check_label 
      Alignment       =   2  'Center
      Caption         =   "�"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   1
      Left            =   855
      TabIndex        =   43
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label transaction_notes_check_label 
      Alignment       =   2  'Center
      Caption         =   "�"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   0
      Left            =   855
      TabIndex        =   42
      Top             =   360
      Width           =   375
   End
   Begin VB.Label monthly_notes_label 
      Caption         =   "Monthly Notes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   1260
      MouseIcon       =   "main.frx":12CBE
      MousePointer    =   99  'Custom
      TabIndex        =   41
      ToolTipText     =   "Notes"
      Top             =   675
      Width           =   2535
   End
   Begin VB.Label transactions_label 
      Caption         =   "Transactions"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   1260
      MouseIcon       =   "main.frx":13100
      MousePointer    =   99  'Custom
      TabIndex        =   40
      ToolTipText     =   "Transactions"
      Top             =   315
      Width           =   2265
   End
   Begin VB.Label todays_date_label 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   45
      MouseIcon       =   "main.frx":13542
      MousePointer    =   99  'Custom
      TabIndex        =   39
      ToolTipText     =   "Go To Today's Date"
      Top             =   1125
      Width           =   795
   End
   Begin VB.Label todays_date_label 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   45
      MouseIcon       =   "main.frx":1384C
      MousePointer    =   99  'Custom
      TabIndex        =   38
      ToolTipText     =   "Go To Today's Date"
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label todays_date_label 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   1
      Left            =   45
      MouseIcon       =   "main.frx":13B56
      MousePointer    =   99  'Custom
      TabIndex        =   37
      ToolTipText     =   "Go To Today's Date"
      Top             =   1620
      Width           =   765
   End
   Begin VB.Label message_label 
      Caption         =   "Message"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   780
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label total_label 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5220
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label last_label 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5940
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label first_label 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6600
      TabIndex        =   9
      Top             =   5700
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
         Shortcut        =   ^S
      End
      Begin VB.Menu save_as_menu 
         Caption         =   "Save &As"
         Shortcut        =   ^A
      End
      Begin VB.Menu dummy1 
         Caption         =   "-"
      End
      Begin VB.Menu view_print_menu 
         Caption         =   "View"
      End
      Begin VB.Menu print_menu 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu spare3 
         Caption         =   "-"
      End
      Begin VB.Menu preferences_menu 
         Caption         =   "Preferences"
      End
      Begin VB.Menu spare1 
         Caption         =   "-"
      End
      Begin VB.Menu database_integrity_menu 
         Caption         =   "Check Database Integrity"
         Visible         =   0   'False
      End
      Begin VB.Menu sparefile1 
         Caption         =   "-"
         Visible         =   0   'False
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
      Begin VB.Menu undo_menu 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu sparee3 
         Caption         =   "-"
      End
      Begin VB.Menu copy_menu 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu cut_menu 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu paste_menu 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu paste_options_menu 
         Caption         =   "Paste Options"
         Begin VB.Menu paste_and_clear_menu 
            Caption         =   "&Paste Pending"
            Index           =   0
         End
         Begin VB.Menu paste_and_clear_menu 
            Caption         =   "Paste and clear"
            Index           =   1
         End
         Begin VB.Menu paste_intact_menu 
            Caption         =   "Paste Intact"
         End
      End
      Begin VB.Menu dummy2 
         Caption         =   "-"
      End
      Begin VB.Menu edit_transaction_menu 
         Caption         =   "&Edit transaction"
         Shortcut        =   ^E
      End
      Begin VB.Menu mark_root_menu 
         Caption         =   "&Mark ""Status"" with..."
         Begin VB.Menu mark_menu 
            Caption         =   "&Blank"
            Index           =   0
         End
         Begin VB.Menu mark_menu 
            Caption         =   "&Done"
            Index           =   1
         End
         Begin VB.Menu mark_menu 
            Caption         =   "&Pending"
            Index           =   2
         End
         Begin VB.Menu mark_menu 
            Caption         =   "&Skip"
            Index           =   3
         End
      End
      Begin VB.Menu tag_menu1 
         Caption         =   "Tags"
         Index           =   0
         Begin VB.Menu set_menu 
            Caption         =   "Set"
            Begin VB.Menu tag_set_menu 
               Caption         =   "&1"
               Index           =   0
            End
            Begin VB.Menu tag_set_menu 
               Caption         =   "&2"
               Index           =   1
            End
            Begin VB.Menu tag_set_menu 
               Caption         =   "&3"
               Index           =   2
            End
            Begin VB.Menu tag_set_menu 
               Caption         =   "&4"
               Index           =   3
            End
         End
         Begin VB.Menu clear_menu 
            Caption         =   "Clear"
            Begin VB.Menu tag_clear_menu 
               Caption         =   "&1"
               Index           =   0
            End
            Begin VB.Menu tag_clear_menu 
               Caption         =   "&2"
               Index           =   1
            End
            Begin VB.Menu tag_clear_menu 
               Caption         =   "&3"
               Index           =   2
            End
            Begin VB.Menu tag_clear_menu 
               Caption         =   "&4"
               Index           =   3
            End
         End
      End
      Begin VB.Menu sparee2 
         Caption         =   "-"
      End
      Begin VB.Menu filter_menu 
         Caption         =   "&Filter Transactions"
         Shortcut        =   ^F
      End
      Begin VB.Menu spar1 
         Caption         =   "-"
      End
      Begin VB.Menu quick_save_menu 
         Caption         =   "Quick Save"
         Index           =   1
      End
      Begin VB.Menu quick_save_menu 
         Caption         =   "Quick Deposit"
         Index           =   2
      End
      Begin VB.Menu view_quick_menu 
         Caption         =   "Quick View/Edit"
      End
      Begin VB.Menu dummy5 
         Caption         =   "-"
      End
      Begin VB.Menu insert_menu 
         Caption         =   "&Insert"
      End
      Begin VB.Menu delete_menu 
         Caption         =   "&Delete"
      End
      Begin VB.Menu spare10 
         Caption         =   "-"
      End
      Begin VB.Menu copy_selected_menu 
         Caption         =   "Copy Selected"
      End
      Begin VB.Menu cut_selected_menu 
         Caption         =   "Cut Selected"
      End
      Begin VB.Menu paste_selected_menu 
         Caption         =   "Paste Selected into current date with ""Status""..."
         Begin VB.Menu paste_selected_option_menu 
            Caption         =   "Blank"
            Index           =   0
         End
         Begin VB.Menu paste_selected_option_menu 
            Caption         =   "Intact"
            Index           =   1
         End
         Begin VB.Menu paste_selected_option_menu 
            Caption         =   "Pending"
            Index           =   2
         End
      End
      Begin VB.Menu dummy3 
         Caption         =   "-"
      End
      Begin VB.Menu copy_month_menu 
         Caption         =   "Copy Month"
      End
      Begin VB.Menu cut_month_menu 
         Caption         =   "Cut Month"
      End
      Begin VB.Menu paste_month_menu 
         Caption         =   "Paste Month"
      End
      Begin VB.Menu paste_month_options_menu 
         Caption         =   "Paste Month Options"
         Begin VB.Menu paste_month_option_root_menu 
            Caption         =   "Paste month with ""Status"" ..."
            Begin VB.Menu paste_month_option_menu 
               Caption         =   "&Blank"
               Index           =   0
            End
            Begin VB.Menu paste_month_option_menu 
               Caption         =   "&Intact"
               Index           =   1
            End
            Begin VB.Menu paste_month_option_menu 
               Caption         =   "&Pending"
               Index           =   2
            End
         End
         Begin VB.Menu paste_month_arrange_menu 
            Caption         =   "Paste Month and Arrange"
         End
      End
      Begin VB.Menu spareee22 
         Caption         =   "-"
      End
      Begin VB.Menu copy_tag_menu 
         Caption         =   "Copy Tags"
         Begin VB.Menu copy_tags_menu 
            Caption         =   "1"
            Index           =   0
         End
         Begin VB.Menu copy_tags_menu 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu copy_tags_menu 
            Caption         =   "3"
            Index           =   2
         End
         Begin VB.Menu copy_tags_menu 
            Caption         =   "4"
            Index           =   3
         End
      End
      Begin VB.Menu cut_tag_menu 
         Caption         =   "Cut Tags"
         Begin VB.Menu cut_tags_menu 
            Caption         =   "1"
            Index           =   0
         End
         Begin VB.Menu cut_tags_menu 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu cut_tags_menu 
            Caption         =   "3"
            Index           =   2
         End
         Begin VB.Menu cut_tags_menu 
            Caption         =   "4"
            Index           =   3
         End
      End
      Begin VB.Menu paste_tags_root_menu 
         Caption         =   "Paste Tags"
         Begin VB.Menu paste_tags_menu 
            Caption         =   "1"
            Index           =   0
         End
         Begin VB.Menu paste_tags_menu 
            Caption         =   "2"
            Index           =   1
         End
         Begin VB.Menu paste_tags_menu 
            Caption         =   "3"
            Index           =   2
         End
         Begin VB.Menu paste_tags_menu 
            Caption         =   "4"
            Index           =   3
         End
      End
      Begin VB.Menu paste_tags_options_menu 
         Caption         =   "Paste Tags Options"
         Begin VB.Menu paste_tags_arrange_menu 
            Caption         =   "Paste Tags and Arrange with ""Status""..."
            Begin VB.Menu paste_tag_arrange_option_menu 
               Caption         =   "Blank"
               Index           =   0
            End
            Begin VB.Menu paste_tag_arrange_option_menu 
               Caption         =   "Intact"
               Index           =   1
            End
            Begin VB.Menu paste_tag_arrange_option_menu 
               Caption         =   "Pending"
               Index           =   2
            End
         End
         Begin VB.Menu paste_tags_current_main_menu 
            Caption         =   "Paste Tags into current date with ""Status""..."
            Begin VB.Menu paste_tags_current_menu 
               Caption         =   "Blank"
               Index           =   0
            End
            Begin VB.Menu paste_tags_current_menu 
               Caption         =   "Intact"
               Index           =   1
            End
            Begin VB.Menu paste_tags_current_menu 
               Caption         =   "Pending"
               Index           =   2
            End
         End
         Begin VB.Menu paste_tags_option_root_menu 
            Caption         =   "Paste Tags with ""Status"" ..."
            Begin VB.Menu paste_tag_option_menu 
               Caption         =   "Blank"
               Index           =   0
            End
            Begin VB.Menu paste_tag_option_menu 
               Caption         =   "Intact"
               Index           =   1
            End
            Begin VB.Menu paste_tag_option_menu 
               Caption         =   "Pending"
               Index           =   2
            End
         End
         Begin VB.Menu paste_tag_menu 
            Caption         =   "Paste Tags with ""Status"" blank"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu spare_cn 
         Caption         =   "-"
      End
      Begin VB.Menu insert_reference_number_menu 
         Caption         =   "Insert Reference Number"
      End
   End
   Begin VB.Menu checkbook_menu 
      Caption         =   "&Checkbook"
      Begin VB.Menu reconcile_menu 
         Caption         =   "&Reconcile"
      End
   End
   Begin VB.Menu cardtrak_main_menu 
      Caption         =   "Cardtrak"
      Begin VB.Menu cardtrak_menu 
         Caption         =   "New Transaction"
         Index           =   0
      End
      Begin VB.Menu cardtrak_menu 
         Caption         =   "Edit Transaction"
         Index           =   1
      End
      Begin VB.Menu cardtrak_menu 
         Caption         =   "Convert to Cardtrack"
         Index           =   2
      End
      Begin VB.Menu cardtrak_menu 
         Caption         =   "Add/Delete/Edit Cards"
         Index           =   3
      End
   End
   Begin VB.Menu view_menu 
      Caption         =   "&View"
      Begin VB.Menu view_quick_accounts_menu 
         Caption         =   "Quick Accounts"
      End
      Begin VB.Menu sparev2 
         Caption         =   "-"
      End
      Begin VB.Menu center_tab_menu 
         Caption         =   "Center Tab"
      End
      Begin VB.Menu vspare3 
         Caption         =   "-"
      End
      Begin VB.Menu next_month_menu 
         Caption         =   "&Next Month"
      End
      Begin VB.Menu next_month_top_menu 
         Caption         =   "Next Month Top"
      End
      Begin VB.Menu next_year_menu 
         Caption         =   "Next Year"
      End
      Begin VB.Menu spare20_menu 
         Caption         =   "-"
      End
      Begin VB.Menu previous_month_menu 
         Caption         =   "&Previous Month"
      End
      Begin VB.Menu previous_month_bottom_menu 
         Caption         =   "Previous Month Bottom"
      End
      Begin VB.Menu previous_year_menu 
         Caption         =   "Previous Year"
      End
      Begin VB.Menu vspare1 
         Caption         =   "-"
      End
      Begin VB.Menu transactions_menu 
         Caption         =   "&Transactions"
         Shortcut        =   ^T
      End
      Begin VB.Menu monthly_notes_menu 
         Caption         =   "&Monthly Notes"
         Shortcut        =   ^M
      End
      Begin VB.Menu misc_notes_menu 
         Caption         =   "Misc &Notes"
         Shortcut        =   ^N
      End
      Begin VB.Menu vspare2 
         Caption         =   "-"
      End
      Begin VB.Menu calendar_menu 
         Caption         =   "&Calendar"
      End
      Begin VB.Menu view_calculator_menu 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu override_columns_menu 
         Caption         =   "&Override Columns"
      End
      Begin VB.Menu sparee1 
         Caption         =   "-"
      End
      Begin VB.Menu goto_month_menu 
         Caption         =   "Go to Month / Year"
         Shortcut        =   ^G
      End
      Begin VB.Menu sparev1 
         Caption         =   "-"
      End
      Begin VB.Menu view_balances_menu 
         Caption         =   "&Balances"
      End
      Begin VB.Menu view_tags_menu 
         Caption         =   "Tags"
      End
      Begin VB.Menu view_summary_menu 
         Caption         =   "Summary"
      End
      Begin VB.Menu cardtrak_summary_menu 
         Caption         =   "Cardtrak Summary"
      End
   End
   Begin VB.Menu search_menu 
      Caption         =   "&Search"
      Visible         =   0   'False
   End
   Begin VB.Menu help_menu 
      Caption         =   "&Help"
      Begin VB.Menu contents_menu 
         Caption         =   "&Contents"
      End
      Begin VB.Menu index_menu 
         Caption         =   "&Index"
      End
      Begin VB.Menu spareh2 
         Caption         =   "-"
      End
      Begin VB.Menu register_menu 
         Caption         =   "&Register Check2Check"
      End
      Begin VB.Menu buy_now_menu 
         Caption         =   "&Buy now"
      End
      Begin VB.Menu update_menu 
         Caption         =   "Check for Updates"
      End
      Begin VB.Menu web_site_menu 
         Caption         =   "Check2Check Web Site"
      End
      Begin VB.Menu register_menu_dash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu language_main_menu 
         Caption         =   "Language"
         Visible         =   0   'False
         Begin VB.Menu language_menu 
            Caption         =   "English"
            Index           =   0
         End
         Begin VB.Menu language_menu 
            Caption         =   "Spanish"
            Index           =   1
         End
      End
      Begin VB.Menu spareh 
         Caption         =   "-"
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

Const debugging = 0  ' 0 don't show, 1 will show this column and the integrity and ct test tab

Const TRANSACTIONS_PAGE = 0
Const NOTES_PAGE = 1
Const CARDTRAK_PAGE = 3

Const NAME_WIDTH_OR = 8400  '7550
Const NAME_WIDTH_NORMAL = 6725  ' 5950 ' 5300

Const MAX_COL = 15

Const DATE_COL = 0
Const DAY_COL = 1
Const THIS_COL = 2
Const PREV_COL = 3
Const NEXT_COL = 4
Const DUE_COL = 5
Const CHECK_COL = 6
Const NAME_COL = 7
Const PAID_COL = 8
Const AMOUNT_COL = 9
Const EXCLUDE_COL = 10
Const TAG_COL = 11
Const CLEARED_COL = 12
Const DOUBLE_LINE1_COL = 13
Const BALANCE_COL = 14
Const DOUBLE_LINE2_COL = 15
Const OVERRIDE_COL = 16
Const OVERRIDE_AMOUNT_COL = 17

Const PM_ARRANGE = 0
Const PM_NORMAL = 1
Const PM_CURRENT = 2

Dim s As String
Dim i, j, k, start_dir, last_row, last_col
Dim doing_move As Boolean  ' True when moving a transaction so we don't ask for prompt
Dim doing_cut As Boolean  ' True when cutting a transaction so we don't ask for prompt
Dim move_from_row As Integer
Dim debug_variable As Boolean
Dim x As Integer, y As Integer  ' Used for x and y for screen and page
Dim transaction_line_count As Integer  ' Count of the number of transaction on a printed or view page
Dim page_type As Integer        ' 0=Transactions, 1= notes
Dim date_s As String  ' Used for printouts
Dim time_s As String  ' Used for printouts
Dim print_destination  ' Either PTR or SRC
Dim form_loading As Boolean  ' True when the form is loading
Dim splash_screen_shown As Boolean
Public going_to_day As Boolean  ' True when desiring to position the cursor on a certain date
Dim startup_directory As String   ' The path that check2check started up in
Dim editing_a_record As Boolean  ' True when a cell has been double clicked
Dim last_column As Integer  ' Last column that was selected by moving the cursor to it, used for cursor moving forward and backward
Dim pasting_tags As Boolean  ' True when we are pasting tags via the menu
Dim last_tag As Integer  ' Tag number of the last tag operated on
Dim paste_month_options As Integer  '0=paste month with arrange,  1=doing a normal paste month, 2=paste into current date
Dim paste_selected_active As Boolean  ' True when pasted selected transactions
Dim items_are_selected As Boolean  ' True when one or more transactions are selected
Dim last_row_selected As Integer  ' Contains the last row that was clicked on
Dim status As Integer  ' Paid status of a quick transaction
Dim number_of_displayed_rows As Integer  ' Number of actual displayed rows
Dim allow_paste_ct As Boolean  ' When true then allow transactions with ct data to be pasted to
Public doing_process As Boolean  ' True when doing process so we know when it's ok to update things
Dim last_transaction_name As String  ' Keep the name when we enter a cell



Public Sub do_debugging(i As Integer)
  If (i > 0) Then
    main_form.database_integrity_menu.Visible = True
    main_form.sparefile1.Visible = True
    main_form.entry_grid.ColWidth(2) = 500   '1  Debugging
  Else
    main_form.database_integrity_menu.Visible = False
    main_form.sparefile1.Visible = False
    main_form.entry_grid.ColWidth(2) = 1  ' Not Debugging
  End If
  
End Sub

Private Sub beginning_balance_box_GotFocus()
  entry_tab.SetFocus
End Sub

Public Function ct_card_number(s As String) As Integer
  Dim n As Integer
  
  ct_card_number = -1
  If (UCase(Mid(s, 2, 2)) = "CT") Then
       n = Val(Mid(s, 4, 2))  ' We have a valid existing credit card
       If (n >= 0) And (n <= MAX_CARDS) Then ct_card_number = n
  End If
End Function

Private Sub bottom_button_Click()
  entry_grid.row = entry_grid.Rows - 1
  entry_grid.TopRow = entry_grid.Rows - 1
End Sub

Private Sub cardtrak_button_Click()
cardtrak_summary_menu_Click
End Sub

Private Sub cardtrak_menu_Click(index As Integer)
  Dim cardn As Integer
  Dim t As Integer
  Dim r, c
  Dim fcn As Integer
  Dim doit As Boolean
  
  If (index = 0) Then fcn = CT_CREATE  ' Make a new transaction
  If (index = 1) Then fcn = CT_EDIT    ' Edit an existing transaction
  If (index = 2) Then fcn = CT_CONVERT ' Convert an existing non-ct to ct transaction
  If (index = 3) Then fcn = CT_ADD     ' Add new cards only
    
  doit = False
  
  With entry_grid
  
    If (fcn = CT_EDIT) Or (fcn = CT_CONVERT) Then
      ' We should be on a transaction
      If (.TextMatrix(.row, THIS_COL) <> "") Then
        ' We are sitting on a real transaction
        t = Val(.TextMatrix(.row, THIS_COL))
        this = db(t)
        cardn = ct_card_number(db(t).name)  ' This is the card number we are editing or -1 for no card number
        If (fcn = CT_CONVERT) And (cardn < 0) Then doit = True
        If (fcn = CT_EDIT) And (cardn >= 0) Then doit = True
      Else
        ' We are not on a real transaction
        t = -1
        this.this = t
        cardn = -1
      End If
    End If
      
    ' See if we are making a whole new ct transaction
    If (fcn = CT_CREATE) Then
      t = -1
      cardn = -1
      this.this = t
      fcn = CT_EDIT
      doit = True
    End If
    
    
    ' See if we are converting the existing record to a ct record
    If ((t >= 0) And (fcn = CT_CONVERT) And (cardn < 0)) Then
      ' Yes we are converting a non-ct record to a ct record
      this.name = "(CT00) " + this.name
      cardn = 0
      fcn = CT_EDIT  ' Let it be edited now
      doit = True
    End If
    
    If (fcn = CT_ADD) Then
      doit = True
    End If
    
'    If ((cardn >= 0) And (cardn <= MAX_CARDS)) Or (t = -1) Then
    If (doit) Then
      ' We have a valid credit card number or we have a blank transaction
      ' We do not have a normal non-ct transaction
      r = entry_grid.row
      c = entry_grid.Col
      ' Setup THIS so that we can get the current information from it
      setup_this_from_current_line
    
      ' ================ Execute the cardtrak form ===============
      If (card_data_form.execute(fcn)) Then
        ' We have hit ok so save the credit card data
        changed_flag = True
        If (fcn = CT_EDIT) Then
          insert_this_record (this.this)  ' Find a place for this record and put it there
          
          update_caption
          process
          entry_grid.row = r  ' Put row selected back to where it was
          entry_grid.Col = c
        End If
      End If
    Else
      ' We are not editing a credit card transaction
    End If
  
  End With
  
  
  
  
  ' Index 1 - Convert the current transaction to a CT transaction


  ' Index 2 - Add/Delete Cards
    ' Go to the card transaction form but don't display the transaction tabs
'    this.name = "(CT00) ADD_CARDS"
'    this.this = -2
'    credit_card_entry_menu_Click
  
  
  

End Sub

Private Sub cardtrak_summary_menu_Click()
  ' Show the tags form
  ct_summary_form.show
  process
End Sub

Private Sub center_month_button_Click()
  center_tab_menu_Click
End Sub

Private Sub center_tab_menu_Click()
    ' Set the center tab to the current date
    Dim m As Integer
    Dim y As Integer
    
    m = view.current_month
    y = view.current_year
    Call set_center_tab(m, y)
End Sub

Private Sub Command1_Click()
  Dim s As String
  s = get_date(4, 1, 0)
  message_label.Caption = s
End Sub

Private Sub ending_balance_box_GotFocus()
  entry_tab.SetFocus
End Sub

Private Sub checkbook_balance_box_GotFocus()
  entry_tab.SetFocus
End Sub

Private Sub buy_now_menu_Click()
  ' Show the reginfo form
  reginfo_form.show vbModal
End Sub

Private Sub calculator_button_Click()
  view_calculator_menu_Click
End Sub

Private Sub checkbook_balance_menu_Click()
  ' Let try to see what the checkbook balance should be
  Dim balance As Double
  Dim i
  
  ' Loop through the entire account and subtract off the cleared items
  ' from the starting balance
  
  balance = 0
  ' Start from the beginning and calculate the balances
  If (data.number_of_records > 0) Then
    ' We have at least one record
    data.current = data.first
    get_record  ' Get the first record
    
    ' See if this record should be included as normal
    If (Not db(data.current).exclude) And (Not db(data.current).override) Then
      If (db(data.current).paid = 1) Then balance = db(data.current).amount
    End If
    
    ' See if we have exclude
    If (db(data.current).exclude) Then
      ' We have exclude
      balance = 0
    Else
      ' See if we have override
      If (db(data.current).override) And (Not db(data.current).exclude) Then
        ' We have override
        If (db(data.current).paid = 1) Then balance = db(data.current).override_amount
      End If
    End If
    
    
    While get_next_record
      ' Loop though all the remaining records and do the balance
      'If (this.this = 79) Then
      'this.this = this.this
      'End If
      If (this.day = 0) Then
      this.day = 0
      End If
    
      
      
      If (Not db(data.current).exclude) And (Not db(data.current).override) Then
        If (db(data.current).paid = 1) Then balance = balance + db(data.current).amount
      End If
    
      ' See if we have override and no exclude
      If (db(data.current).override) And (Not db(data.current).exclude) Then
        ' We have override
        If (db(data.current).paid = 1) Then balance = db(data.current).override_amount
      End If
    
      
    Wend
  End If
  
  'display the balance
  ending_balance_box.Text = currency_s(balance)

End Sub

Private Sub Command2_Click()
  view_tags_menu_Click
End Sub

Private Sub copy_selected_menu_Click()
  Dim j, r, c
  
  ' Loop through all the displayed rows and get the ones that are in bold to copy_of_month
  With entry_grid
  '.MousePointer = flexHourglass
  main_form.MousePointer = vbHourglass '  = 99  '2
  
  .Redraw = False
  'Save the current cell position
  
  r = .row
  c = .Col
  
  j = 0
  For i = 1 To .Rows - 1
    .row = i
    .Col = NAME_COL
    If ((table_image.table(i).this >= 0) And _
        (.CellFontBold)) Then
      data.current = table_image.table(i).this
      get_record  ' Read the selected record and save all the parameters
      
      copy_of_month.table(j) = this
      copy_of_cardtrak_month.table(j) = cards(this.sub_transaction_number)
      j = j + 1
    End If
  Next i
  
  allow_paste_ct = False  ' Don't allow paste of ct data
  
  copy_of_month.table(j).this = -1
  
  copy_of_month.Month = view.current_month
  copy_of_month.Year = view.current_year
  copy_of_month.notes = ""  ' Don't copy any notes

  If (j = 0) Then
    MsgBox words(NO_TRANSACTIONS_COPIED_N)  '"No transactions copied"
  End If
  
  If (r >= .Rows) Then r = .Rows - 1
  .row = r
  .Col = c
  .Redraw = True
  '.MousePointer = flexArrow
  main_form.MousePointer = vbDefault '  = 99  '2
  End With
End Sub


Private Sub cut_selected_menu_Click()
  Dim r, c
  Dim j As Integer
  
  If (view.records_in_month = 0) Then
    MsgBox words(NO_TRANSACTIONS_TO_CUT_N) ' "No transactions to cut"
    Exit Sub
  End If
  
  With entry_grid
  .Redraw = False
  
  r = .row
  c = .Col
  j = 0
  
'  If (MsgBox(words(cut_all_selected_n) + " " + words(transactions_for_n) + " " + entry_tab.Caption, _
'      vbYesNoCancel + vbQuestion + vbApplicationModal, "Delete Month") = vbYes) Then
    ' Yes, delete the entire month
      For i = 1 To .Rows - 1
        .row = i
        .Col = NAME_COL
        If ((table_image.table(i).this >= 0) And _
            (.CellFontBold)) Then
          data.current = table_image.table(i).this
          get_record  ' Read the selected record and save all the parameters
      
          copy_of_month.table(j) = this
          copy_of_cardtrak_month.table(j) = cards(this.sub_transaction_number)
          delete_record (table_image.table(i).this)
          j = j + 1
        End If
      Next i
  
      allow_paste_ct = True  ' Allowing pasting since we are doing a cut
      
      copy_of_month.table(j).this = -1
    
      copy_of_month.Month = view.current_month
      copy_of_month.Year = view.current_year
      copy_of_month.notes = notes_box.Text
      
      ' Save the undo stuff
      undo.what_was_done = WHAT_CUT_MONTH
      undo.copy_of_month = copy_of_month
      undo_cardtrak_month = copy_of_cardtrak_month
      undo_menu.Enabled = True
      undo_button.Visible = undo_menu.Enabled
      undo_menu.Caption = "Undo - Cut Selected"
      Label1.Caption = "?"
      
      
      If (j > 0) Then  ' If we cut any transactions then do process
        process
        changed_flag = True
        update_caption
      End If
      
  If (j = 0) Then
    MsgBox words(NO_TRANSACTIONS_CUT_N)  '"No transactions cut"
  End If
  
  If (r >= .Rows) Then r = .Rows - 1
  .row = r
  .Col = c
  .Redraw = True
  
'  End If
  End With
End Sub

Private Sub copy_tags_menu_Click(index As Integer)
  Dim j
  
  If (view.records_in_month = 0) Then
    MsgBox words(NO_TRANSACTIONS_TO_COPY_N)  ' "No transactions to copy"
    Exit Sub
  End If
  
  ' Copy the current tag data to the copy_of_month
  j = 0
  For i = 1 To MAX_RECORDS_IN_MONTH
    If (table_image.table(i).day = 0) Then Exit For
    If (table_image.table(i).this >= 0) Then
      data.current = table_image.table(i).this
      get_record  ' Read the selected record and save all the parameters
      
      ' See if it matches the tag
      If ((this.tags And tag_mask(index)) <> 0) Then
        ' We have a tag that matches
        copy_of_month.table(j) = this
        copy_of_cardtrak_month.table(j) = cards(this.sub_transaction_number)
        j = j + 1
      End If
    End If
  Next i
  
  copy_of_month.table(j).this = -1
  
  copy_of_month.Month = view.current_month
  copy_of_month.Year = view.current_year
  copy_of_month.notes = ""  ' Don't copy any notes

  If (j = 0) Then
    MsgBox words(NO_TRANSACTIONS_COPIED_N)  '"No transactions copied"
    Exit Sub
  End If
  
End Sub

Private Sub setup_this_from_current_line()
  ' Get the date and put it in 'this'
  this.day = view.current_day
  this.Month = view.current_month
  this.Year = view.current_year
End Sub

Private Sub cut_tags_menu_Click(index As Integer)
  Dim j
  
  If (view.records_in_month = 0) Then
    MsgBox words(NO_TRANSACTIONS_TO_CUT_N)  '"No transactions to cut"
    Exit Sub
  End If
  
  If (MsgBox(words(CUT_ALL_TAGGED_Q_N) + " " + Str(index + 1) + " " + words(TRANSACTIONS_FOR_N) + " " + entry_tab.Caption, _
      vbYesNoCancel + vbQuestion + vbApplicationModal, "Delete Month") = vbYes) Then
    ' Yes, delete the entire month
    With table_image
      j = 0
      For i = 1 To MAX_RECORDS_IN_MONTH
        If (.table(i).day = 0) Then Exit For
      
        If (.table(i).this >= 0) Then
          ' We have a record to delete
          
          ' Get the record and save it
          data.current = .table(i).this
          get_record  ' Read the selected record and save all the parameters
          
          ' See if it matches the tag
          If ((this.tags And tag_mask(index)) <> 0) Then
            ' We have a tag that matches
              copy_of_month.table(j) = this
              copy_of_cardtrak_month.table(j) = cards(this.sub_transaction_number)
              j = j + 1
          
              delete_record (.table(i).this)
          End If
        End If
      Next i
      copy_of_month.table(j).this = -1
    
      copy_of_month.Month = view.current_month
      copy_of_month.Year = view.current_year
      copy_of_month.notes = notes_box.Text
      
      ' Save the undo stuff
      undo.what_was_done = WHAT_CUT_MONTH
      undo.copy_of_month = copy_of_month
      undo_cardtrak_month = copy_of_cardtrak_month
      undo_menu.Enabled = True
      undo_button.Visible = undo_menu.Enabled
      undo_menu.Caption = "Undo - Cut " + MONTH_STRINGS(view.current_month) + " " + Format(view.current_year)
      Label1.Caption = "?"
      
      
      process
      changed_flag = True
      update_caption
    End With
  
  End If
End Sub

Private Sub database_integrity_menu_Click()
  integrity_form.show
End Sub

Private Sub edit_transaction_menu_Click()
  Dim r, c
  
  ' Edit the transaction that the cell is pointing to
  ' Do this by calling the double click method
  With entry_grid
    r = .row
    c = .Col
    If (.TextMatrix(.row, THIS_COL) <> "") Then
      ' We have a valid record to work with
      this = db(Val(.TextMatrix(.row, THIS_COL)))
      
      ' Save the current record number for undo
      undo.r = this
      undo.what_was_done = WHAT_EDIT_TRANSACTION
      
      If (edit_transaction_form.execute) Then
        ' We have a successful edit so save it
        db(this.this) = this
        process
        .row = r
        .Col = c
      End If
      
    End If
  End With
  'entry_grid_DblClick
    
  undo_menu.Enabled = True
  undo_button.Visible = undo_menu.Enabled
  undo_menu.Caption = "Undo Edit Transaction"
  'Label1.Caption = "Undo Edit Transactions"
End Sub

Private Sub entry_grid_Click()
  view.current_day = Val(entry_grid.TextMatrix(entry_grid.row, DATE_COL))
  
  If (entry_grid.Col = DATE_COL) Or _
     (entry_grid.Col = DAY_COL) Or _
     (entry_grid.Col = THIS_COL) Or _
     (entry_grid.Col = PREV_COL) Or _
     (entry_grid.Col = NEXT_COL) Then
       entry_grid.Col = DUE_COL
  End If
  
  If (entry_grid.Col = DOUBLE_LINE1_COL) And (last_column = CLEARED_COL) Then
       entry_grid.Col = CLEARED_COL  'OVERRIDE_COL
  End If
  
  If (entry_grid.Col = DOUBLE_LINE1_COL) And (last_column = OVERRIDE_COL) Then
       entry_grid.Col = CLEARED_COL
  End If
  
  If (entry_grid.Col = BALANCE_COL) Then
       entry_grid.Col = AMOUNT_COL
  End If
  
  If (entry_grid.Col = DOUBLE_LINE2_COL) And (last_column = CLEARED_COL) Then
       entry_grid.Col = OVERRIDE_COL
  End If
  
  If (entry_grid.Col = DOUBLE_LINE2_COL) And (last_column = OVERRIDE_COL) Then
       entry_grid.Col = CLEARED_COL
  End If
  
  last_column = entry_grid.Col
  
  ' Save the last transaction name so it can be pasted into the notes
  If (last_column = NAME_COL) Then
     last_transaction_name = entry_grid.TextMatrix(entry_grid.row, NAME_COL)
  End If
End Sub

Private Sub entry_grid_EnterCell()
  ' Set the cell background color
  With entry_grid
    If (.row > 0) Then entry_grid.CellBackColor = vbYellow
  End With
End Sub

Private Sub entry_grid_RowColChange()
  If (entry_grid.Redraw = True) Then Call entry_grid_Click
End Sub

Private Sub entry_tab_DblClick()
  ' Double clicking a tab will center it
  center_tab_menu_Click
End Sub

Private Sub filter_button_Click()
  filter_menu_Click
End Sub

Private Sub filter_menu_Click()
  filter_results_form.Hide
  If (filter_form.execute) Then process
End Sub

Private Sub Form_Activate()
  If splash_screen_shown = False Then
    If preferences.show_splash_screen Then
      splash_form.show
    Else
      splash_timer.Enabled = False
    End If
  End If
  splash_screen_shown = True
  
  Call WheelHook(main_form)
End Sub

Private Sub Form_Deactivate()
  WheelUnHook
End Sub

Private Sub Form_Load()
  Dim commandline As Variant
  Dim no_command_line As Boolean
  Dim s
  Dim m
  
  paste_month_options = PM_ARRANGE
  splash_screen_shown = False
  paste_selected_active = False
  items_are_selected = False
  ReDim QUICK_ACCOUNTS.account(MAX_QUICK_ACCOUNT + 1)
  
  ' See if the file association already exists
  ' Let Install Maker do the association
  'associate_file_type
  
  main_form.Caption = "Check2Check   v" + Format(major_version) + "." + Format(minor_version) + "   " + version_date_s
  
  ' Increment the startup counter
  m = Val(GetSetting("Check 2 Check", "Settings", "Startups", "1"))
  m = m + 1
  If (m > 30000) Then m = 1  ' Start over
  SaveSetting "Check 2 Check", "Settings", "Startups", Format(m)
  
  ' Get the start directory
  start_dir = GetSetting("Check 2 Check", "Settings", "Directory", "C:\")
  
  strings_initialize
  
  form_loading = True
  
  notes_box.Top = entry_grid.Top
  notes_box.Left = entry_grid.Left
  
  misc_notes_box.Top = entry_grid.Top
  misc_notes_box.Left = entry_grid.Left
  
  entry_grid.MergeCol(0) = True
  entry_grid.MergeCol(1) = True
  
  entry_grid.ColWidth(DATE_COL) = 500
  entry_grid.ColWidth(DAY_COL) = 500
  entry_grid.ColWidth(THIS_COL) = 1 '500   Debugging
  entry_grid.ColWidth(PREV_COL) = 1  '300
  entry_grid.ColWidth(NEXT_COL) = 1  '300
  entry_grid.ColWidth(DUE_COL) = 400
  entry_grid.ColWidth(CHECK_COL) = 600
  entry_grid.ColWidth(NAME_COL) = NAME_WIDTH_NORMAL
  entry_grid.ColWidth(PAID_COL) = 550
  entry_grid.ColWidth(AMOUNT_COL) = 1200
  entry_grid.ColWidth(EXCLUDE_COL) = 400
  entry_grid.ColWidth(TAG_COL) = 600
  entry_grid.ColWidth(CLEARED_COL) = 400
  entry_grid.ColWidth(DOUBLE_LINE1_COL) = 30
  entry_grid.ColWidth(BALANCE_COL) = 1200
  entry_grid.ColWidth(DOUBLE_LINE2_COL) = 30
  entry_grid.ColWidth(OVERRIDE_COL) = 400
  entry_grid.ColWidth(OVERRIDE_AMOUNT_COL) = 1200
  
  do_debugging (debugging)  ' Show or hide the debug stuff, i.e. this
  
  ' Put up the data for this current month
  For i = 1 To 31
    entry_grid.TextArray(egi(i, 0)) = i
  Next i
  
  'entry_grid.Redraw = True
  
  For i = 0 To UBound(db, 1)
    db(i).this = -1
  Next i
  
  ' Start with no records loaded
  data.first = 0
  data.last = 0
  data.current = 0
  data.number_of_records = 0
  
  For i = 0 To UBound(db, 1)
    db(i).this = -1
  Next i
  
  view.start_day = 1
  
  m = Month(Now) - 5
  view.start_year = Year(Now)
  If (m <= 0) Then
    view.start_month = 12 + m
    view.start_year = view.start_year - 1
  Else
    view.start_month = m
  End If
  
  view.current_month = view.start_month
  view.current_year = view.start_year
  
  entry_grid.ColAlignment(PAID_COL) = flexAlignRightCenter
  entry_grid.ColAlignment(EXCLUDE_COL) = flexAlignRightCenter
  entry_grid.ColAlignment(OVERRIDE_COL) = flexAlignRightCenter
  
  paste_month_menu.Enabled = False
  paste_month_option_root_menu.Enabled = False
  
  copy_of_this.this = -1  ' Start off with nothing in the copy buffer
  copy_of_cardtrak.active = False
  
  entry_tab.Tab = 5
  doing_move = False
  update_entry_tabs
  preferences.auto_insert = False
  preferences.prompt_for_move = True
  transactions_menu.Checked = True
  undo_menu.Enabled = False
  undo_button.Visible = False
  editing_a_record = False
  pasting_tags = False
  copy_of_month.table(0).this = -1
  startup_directory = App.Path
  reference_number_frame.Visible = False
  characters_left_frame.Visible = False
  
  misc_notes_box.Text = ""
  
  ' Load in the preferences
  Call load_preferences
  notes_box.FontSize = preferences.notes_font_size
  misc_notes_box.FontSize = preferences.notes_font_size
  debug_variable = False
  
  ' See if there is a command line file to load
  commandline = Command$()
    
  On Error GoTo error_h
  
  If (commandline <> "") And (open_menu.Visible) Then
    open_dialog.Filename = commandline
    open_dialog.InitDir = strip_filename(commandline)
    start_dir = strip_filename(commandline)
    ChDir start_dir
    ChDrive get_drive(commandline)
    Call open_the_file(False)
    
    save_to_recent_docs (commandline)
  End If
  
  ' Set the start language
  language_menu_Click (preferences.language)  ' Display the startup language
  
error_h:
  
  Get_Recent_Files  ' Fill up the recent filenames in the file menu

  ' Open the last file if set for auto load
  If (main_form.recent_file_menu(0).Visible) And _
     (open_menu.Visible) And _
     (preferences.auto_load_last_file) And _
     (Left(s, 8) <> "UNTITLED") Then
    ' Load it now
    On Error GoTo error_h1
    s = main_form.recent_file_menu(0).Caption
    open_dialog.Filename = s
    open_dialog.InitDir = strip_filename(s)
    start_dir = strip_filename(s)
    ChDir start_dir
    ChDrive get_drive(s)
    Call open_the_file(False)
error_h1:
  End If
    
  ' Play the signon sound
  play_sound (1)
  
  'main_form.Caption = "Check 2 Check   v" + Format(major_version) + "." + Format(minor_version) + "   " + version_date_s
End Sub

Private Sub load_preferences()
  preferences.prompt_for_move = ("True" = GetSetting("Check 2 Check", "Settings", "Prompt_for_move", "True"))
  preferences.auto_insert = ("True" = GetSetting("Check 2 Check", "Settings", "Auto_insert", "False"))
  preferences.show_splash_screen = ("True" = GetSetting("Check 2 Check", "Settings", "Show_splash_screen", "True"))
  preferences.show_name_colors = ("True" = GetSetting("Check 2 Check", "Settings", "Show_name_colors", "True"))
  preferences.show_override_columns = ("True" = GetSetting("Check 2 Check", "Settings", "Show_override_columns", "False"))
  preferences.auto_check_done = ("True" = GetSetting("Check 2 Check", "Settings", "Auto_check_done", "True"))
  preferences.auto_check_done_on_check = ("True" = GetSetting("Check 2 Check", "Settings", "Auto_check_done_on_check", "True"))
  preferences.prompt_for_delete = ("True" = GetSetting("Check 2 Check", "Settings", "Prompt_for_delete", "True"))
  preferences.prompt_for_paste_notes = ("True" = GetSetting("Check 2 Check", "Settings", "Prompt_for_paste_notes", "True"))
  preferences.auto_negative_numbers = ("True" = GetSetting("Check 2 Check", "Settings", "Auto_negative_numbers", "True"))
  preferences.save_recovery_file = ("True" = GetSetting("Check 2 Check", "Settings", "Save_recovery_file", "True"))
  preferences.notes_font_size = Val(GetSetting("Check 2 Check", "Settings", "Notes_font_size", "10"))
  preferences.paste_month(0) = Val(GetSetting("Check 2 Check", "Settings", "Paste_month_blank", "0"))  ' Default to blank
  preferences.paste_month(1) = Val(GetSetting("Check 2 Check", "Settings", "Paste_month_done", "2"))  ' Default to pending
  preferences.paste_month(2) = Val(GetSetting("Check 2 Check", "Settings", "Paste_month_pending", "2"))  ' Default to pending
  preferences.paste_month(3) = Val(GetSetting("Check 2 Check", "Settings", "Paste_month_skip", "3"))  ' Default to skip
  preferences.paste_month(4) = Val(GetSetting("Check 2 Check", "Settings", "Paste_month_check", "1"))  ' Default to blank
  preferences.paste_month(5) = Val(GetSetting("Check 2 Check", "Settings", "Paste_month_cleared", "1"))  ' Default to blank
  preferences.language = Val(GetSetting("Check 2 Check", "Settings", "Language", "0"))
  preferences.auto_load_last_file = ("True" = GetSetting("Check 2 Check", "Settings", "Auto_load_last_file", "True"))
  preferences.play_sounds = ("True" = GetSetting("Check 2 Check", "Settings", "Play_sounds", "True"))
  
  override_columns_menu.Checked = preferences.show_override_columns
End Sub

Private Sub save_preferences()
  SaveSetting "Check 2 Check", "Settings", "Prompt_for_move", preferences.prompt_for_move
  SaveSetting "Check 2 Check", "Settings", "Auto_insert", preferences.auto_insert
  SaveSetting "Check 2 Check", "Settings", "Show_splash_screen", preferences.show_splash_screen
  SaveSetting "Check 2 Check", "Settings", "Show_name_colors", preferences.show_name_colors
  SaveSetting "Check 2 Check", "Settings", "Show_override_columns", preferences.show_override_columns
  SaveSetting "Check 2 Check", "Settings", "Auto_check_done", preferences.auto_check_done
  SaveSetting "Check 2 Check", "Settings", "Auto_check_done_on_check", preferences.auto_check_done_on_check
  SaveSetting "Check 2 Check", "Settings", "Prompt_for_delete", preferences.prompt_for_delete
  SaveSetting "Check 2 Check", "Settings", "Prompt_for_paste_notes", preferences.prompt_for_paste_notes
  SaveSetting "Check 2 Check", "Settings", "Auto_negative_numbers", preferences.auto_negative_numbers
  SaveSetting "Check 2 Check", "Settings", "Save_recovery_file", preferences.save_recovery_file
  SaveSetting "Check 2 Check", "Settings", "Notes_font_size", Str(preferences.notes_font_size)
  SaveSetting "Check 2 Check", "Settings", "Paste_month_blank", Str(preferences.paste_month(0))
  SaveSetting "Check 2 Check", "Settings", "Paste_month_done", Str(preferences.paste_month(1))
  SaveSetting "Check 2 Check", "Settings", "Paste_month_pending", Str(preferences.paste_month(2))
  SaveSetting "Check 2 Check", "Settings", "Paste_month_skip", Str(preferences.paste_month(3))
  SaveSetting "Check 2 Check", "Settings", "Paste_month_check", Str(preferences.paste_month(4))
  SaveSetting "Check 2 Check", "Settings", "Paste_month_cleared", Str(preferences.paste_month(5))
  SaveSetting "Check 2 Check", "Settings", "Language", Str(preferences.language)
  SaveSetting "Check 2 Check", "Settings", "Auto_load_last_file", preferences.auto_load_last_file
  SaveSetting "Check 2 Check", "Settings", "Play_sounds", preferences.play_sounds
End Sub

Private Sub about_menu_Click()
  frmAbout.show 1
End Sub

Private Sub calendar_menu_Click()
  ' Show the calendar
  calendar_form.show
End Sub

Private Sub close_menu_Click()
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
  
  ' ----- See if multiple transactions are selected and if yes then go to copy selected
  If (items_are_selected) Then
    copy_selected_menu_Click
    Exit Sub
  End If
  
  
  ' ----- See if we are editing a name or amount
  If (txtEdit.Visible) Then
    ' Yes we are editing a field
    If (txtEdit.SelLength <> 0) Then
      Clipboard.Clear
      Clipboard.SetText txtEdit.SelText
    End If
    Exit Sub
  End If
  
  ' ----- See if we are doing the notes
  If (monthly_notes_menu.Checked = True) Then
    If (notes_box.SelLength <> 0) Then
      Clipboard.Clear
      Clipboard.SetText notes_box.SelText
    End If
    Exit Sub
  End If
  
  ' ----- Copy the selected record
  With entry_grid
    If (.TextMatrix(.row, THIS_COL) <> "") Then
      ' We have a record to copy
      data.current = Int(.TextMatrix(.row, THIS_COL))
      get_record
      If (Not doing_cut) Then this.sub_transaction_number = 0  ' Don't copy the cardtrack information if doing a COPY
      copy_of_this = this
      'copy_of_cardtrak = cards(this.sub_transaction_number)
      
      copy_to_clipboard
      allow_paste_ct = False  ' Since we are copying then don't allow pasting ct data
    End If
  End With
  
End Sub

Private Sub copy_to_clipboard()
  Dim s
  
  With entry_grid
  
    's = Format(this.month) + "/" + .TextMatrix(.row, DATE_COL) + "/" + Format(this.year) + "  " + _
    s = get_date(this.month), val(.TextMatrix(.row, DATE_COL)), Format(this.year)) + "  " + _
       .TextMatrix(.row, DAY_COL) + "      " + _
       .TextMatrix(.row, DUE_COL) + "      " + _
       .TextMatrix(.row, CHECK_COL) + "      " + _
       .TextMatrix(.row, NAME_COL) + "   "
    s = s + "Amount ( " + .TextMatrix(.row, AMOUNT_COL) + ")   "
    s = s + "Balance ( " + .TextMatrix(.row, BALANCE_COL) + ")   "
  
    If (this.paid) Then s = s + " Done  "
    If (this.exclude) Then s = s + " Exclude  "
    If (this.override) Then s = s + " Override"
  
  End With
  
  Clipboard.SetText (s)
End Sub

Private Sub copy_month_menu_Click()
  Dim j As Integer
  
  If (view.records_in_month = 0) Then
    MsgBox words(NO_TRANSACTIONS_TO_COPY_N)  '"No transactions to copy"
    Exit Sub
  End If
  
  ' Copy the current month data to the copy_of_month
  j = 0
  Call save_to_month(0, this)   ' Initialize the month save
  For i = 1 To MAX_RECORDS_IN_MONTH
    If (table_image.table(i).day = 0) Then Exit For
    If (table_image.table(i).this >= 0) Then
      data.current = table_image.table(i).this
      get_record  ' Read the selected record and save all the parameters
      Call save_to_month(1, this)
      j = j + 1
    End If
  Next i
  copy_of_month.table(j).this = -1
    
  copy_of_month.Month = view.current_month
  copy_of_month.Year = view.current_year
  copy_of_month.notes = notes_box.Text
  allow_paste_ct = False
End Sub

Private Sub cut_button_Click()
  cut_menu_Click
End Sub

Private Sub cut_menu_Click()
  ' ----- See if we are editing a name or amount
  If (txtEdit.Visible) Then
    ' Yes we are editing a field
    If (txtEdit.SelLength <> 0) Then
      Clipboard.Clear
      Clipboard.SetText txtEdit.SelText
      txtEdit.SelText = ""
    End If
    Exit Sub
  End If
  
  ' ----- See if we are doing the notes
  If (monthly_notes_menu.Checked = True) Then
    If (notes_box.SelLength <> 0) Then
      Clipboard.Clear
      Clipboard.SetText notes_box.SelText
      notes_box.SelText = ""
    End If
    Exit Sub
  End If
  
  
  ' ----- See if multiple transactions are selected and if yes then go to cut selected
  If (items_are_selected) Then
    cut_selected_menu_Click
    Exit Sub
  End If
  
  ' ----- Copy and then delete the record
  doing_cut = True
  copy_menu_Click
  delete_menu_Click
  doing_cut = False
  changed_flag = True
  update_caption
  undo_menu.Caption = "Undo - Cut transaction"
  allow_paste_ct = True
End Sub

Private Sub delete_button_Click()
  delete_menu_Click
  changed_flag = True
  update_caption
End Sub

Private Sub save_undo_this()
  undo.r = this
  undo.cardtrak = cards(this.sub_transaction_number)
End Sub


Private Sub delete_menu_Click()
  Dim rec As Integer, t
  Dim answer, r
  
  ' See if multiple transactions are selected
  If (items_are_selected) Then
    cut_selected_menu_Click  ' Cut all the selected records
    Exit Sub
  End If
  
  ' Delete a single record
  With entry_grid
    t = .TopRow
    
    .Redraw = False
    If (.TextMatrix(.row, THIS_COL) <> "") Then
      ' We have a valid record so verify the delete with user
      answer = vbYes
      If (preferences.prompt_for_delete) And (Not doing_move) And (Not doing_cut) Then
        answer = MsgBox(words(DELETE_TRANSACTION_Q_N), vbYesNoCancel + vbQuestion, "Check2Check")
      End If
      
      
      If (answer = vbYes) Then
        r = .row
        rec = .TextMatrix(.row, THIS_COL)
        ' Save the current record number for undo
        ' Get the record pointed to by rec
        data.current = rec
        get_record
        Call save_undo_this  ' Save this and ct
        
        undo.what_was_done = WHAT_DELETE_RECORD
        If (doing_cut) Then undo.what_was_done = WHAT_CUT_RECORD
        If (doing_move) Then undo.what_was_done = WHAT_MOVE_RECORD
        undo_menu.Enabled = True
        undo_button.Visible = undo_menu.Enabled
        undo_menu.Caption = "Undo - Delete transaction"
        Label1.Caption = undo.r.this
    
        delete_record (rec)
    
        process
        changed_flag = True
        update_caption
        
        ' Make the same line active
        .Col = NAME_COL
        If (r < .Rows) Then
          .row = r
        Else
          .row = r - 1
        End If
        .Col = NAME_COL
        
      End If
    Else
      ' We are deleting a blank line so see if it's the first in the date
      If (.TextMatrix(.row, DATE_COL) = .TextMatrix(.row - 1, DATE_COL)) Then
        ' It is not the first so go ahead and delete the entire line
        shuffle_up_entry_grid (.row)
      End If
    End If
    .Redraw = True
    'If (.Row > 10) Then .TopRow = .Row - 7
    .TopRow = t
  End With
End Sub

Private Sub cut_month_menu_Click()
  If (view.records_in_month = 0) Then
    MsgBox words(NO_TRANSACTIONS_TO_CUT_N)  '"No transactions to cut"
    Exit Sub
  End If
  
  If (MsgBox(words(CUT_ALL_TRANSACTIONS_FOR_N) + " " + entry_tab.Caption, _
      vbYesNoCancel + vbQuestion + vbApplicationModal, "Delete Month") = vbYes) Then
    ' Yes, delete the entire month
    With table_image
      Call save_to_month(0, this)  ' Initialize month
      For i = 1 To MAX_RECORDS_IN_MONTH
        If (.table(i).day = 0) Then Exit For
      
        If (.table(i).this >= 0) Then
          ' We have a record to delete
          
          ' Get the record and save it
          data.current = .table(i).this
          get_record  ' Read the selected record and save all the parameters
          Call save_to_month(1, this)
          
          delete_record (.table(i).this)
        End If
      Next i
      copy_of_month.table(j).this = -1
    
      copy_of_month.Month = view.current_month
      copy_of_month.Year = view.current_year
      copy_of_month.notes = notes_box.Text
      
      ' Save the undo stuff
      undo.what_was_done = WHAT_CUT_MONTH
      undo.copy_of_month = copy_of_month
      undo_cardtrak_month = copy_of_cardtrak_month
      undo_menu.Enabled = True
      undo_button.Visible = undo_menu.Enabled
      undo_menu.Caption = "Undo - Cut " + MONTH_STRINGS(view.current_month) + " " + Format(view.current_year)
      Label1.Caption = "?"
      
      ' Delete the notes too
      notes_box.Text = ""
      update_notes
      
      process
      allow_paste_ct = True
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
  Dim r As Integer, c As Integer
  
  ' Calculate the row it is on
  r = entry_grid.MouseRow  'Int(y / entry_grid.RowHeight(1)) + entry_grid.TopRow - 1
  If (r < 1) Then Exit Sub
  If (r > entry_grid.Rows - 1) Then Exit Sub
  c = entry_grid.Col
  
  ' Turn off any bold names
  If (False) And (Shift = 0) And (Button = 1) And (items_are_selected) Then
    With entry_grid
      .Redraw = False
      For i = 1 To .Rows - 1
        .row = i
        .Col = NAME_COL
        If (.CellFontBold) Then .CellFontBold = False
      Next i
      .Redraw = True
      items_are_selected = False
      show_buttons
    End With
  End If
  
  entry_grid.row = r
  entry_grid.Col = c
  
  ' See if left button pressed for move
  If Button = 1 Then
    
    main_form.MousePointer = 99  '2
    
    move_from_row = entry_grid.row
    If (entry_grid.TextMatrix(entry_grid.row, THIS_COL) <> "") Then
      ' See if we are tagging or moving
      If (Shift = 0) Then
        doing_move = True
      End If
      
      If (Shift = 1) And (entry_grid.Col = NAME_COL) Then
        ' Shift key held down
        ' Make all the names from the last click bold
        entry_grid.Redraw = False
        For i = last_row_selected To entry_grid.row
          If (table_image.table(i).this >= 0) Then
            ' We have a valid transaction
            entry_grid.row = i
            entry_grid.Col = NAME_COL
            entry_grid.CellFontBold = True
            items_are_selected = True
            show_buttons
          End If
        Next i
        entry_grid.Redraw = True
      End If
      
      If (Shift = 2) And (entry_grid.Col = NAME_COL) Then
        ' Control key held down
        If (entry_grid.CellFontBold) Then  ' Toggle the bold
          entry_grid.CellFontBold = False
        Else
          entry_grid.CellFontBold = True
          items_are_selected = True
          show_buttons
        End If
      End If
      
    End If
  End If
  
  last_row_selected = entry_grid.row
  
  ' See if right button pressed for menu
  If Button = 2 Then
    PopupMenu edit_menu
  End If
End Sub

Private Sub entry_grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim r, day, answer
  
  main_form.MousePointer = vbDefault
  
  If (Button = 1) And (doing_move) Then
    ' Calculate the row it is on
    r = entry_grid.MouseRow  'Int(y / entry_grid.RowHeight(1)) + entry_grid.TopRow - 1
    If (move_from_row <> r) And _
       (r >= 1) And (r < entry_grid.Rows) Then
       
      ' We have a valid move so verify with user
      answer = vbYes
      If (preferences.prompt_for_move) Then
        answer = MsgBox(words(MOVE_TRANSACTION_Q_N), vbYesNoCancel + vbQuestion, "Check2Check")
      End If
      
      If answer = vbYes Then
        ' Moving record
        entry_grid.row = move_from_row
        day = entry_grid.TextMatrix(r, DATE_COL)
        If (items_are_selected) Then
          ' Move selected transactions
          undo.what_was_done = WHAT_MOVE_SELECTED
          cut_selected_menu_Click
          view.current_day = day
          undo.what_was_done = WHAT_MOVE_SELECTED
          allow_paste_ct = True
          paste_selected_option_menu_Click (1)  ' Intact
          undo.what_was_done = WHAT_MOVE_SELECTED
        Else
          ' Move a single transaction
          cut_menu_Click
          copy_of_this.day = day
          allow_paste_ct = True
          paste_intact_menu_Click
        End If
      End If
      
    End If
    
    doing_move = False
  End If
  
  ' Turn off any bold names
  If (Shift = 0) And (Button = 1) And (items_are_selected) Then
    With entry_grid
      .Redraw = False
      For i = 1 To .Rows - 1
        .row = i
        .Col = NAME_COL
        If (.CellFontBold) Then .CellFontBold = False
      Next i
      .Redraw = True
      items_are_selected = False
      show_buttons
    End With
  End If
  

End Sub

Private Sub exit_menu_Click()
  Unload Me
End Sub

Private Function egi(r, c) As Integer
  egi = r * entry_grid.Cols + c
End Function

Public Sub update_entry_tab_names()
  Dim y
  
  y = view.start_year
  For i = 0 To 11
    j = view.start_month + i
    
    If (j > 12) Then
    j = j - 12
      y = view.start_year + 1
    End If
    
    entry_tab.TabCaption(i) = MONTH_STRINGS(j) + " " + CStr(y)
  Next i
End Sub

Public Sub update_entry_tabs()
' ESK 1/27/2017
' I added this line because the variable i was not declared locally and since it was declared
' globally it was being modified and causing a runtime error when double clicking
' on a calendar date.
  Dim i  ' ESK 1/26/2017
  
  update_entry_tab_names
  process
  entry_grid.Col = NAME_COL
  
  ' Go down the list and find the matching day
  entry_grid.row = 1
  If (going_to_day) Then
    entry_grid.row = 1
    For i = 1 To 100  ' Start at row 1 because row 0 has the "Date" string, No month should havr more than 100 rows
      On Error GoTo error_h
 s = entry_grid.TextMatrix(i, DATE_COL)
 
      If (Val(entry_grid.TextMatrix(i, DATE_COL)) = view.start_day) Then
        Exit For
      End If
      GoTo continue
error_h:
      i = 1
      Exit For  ' We had an error
continue:
    Next i
  entry_tab.Tab = 0
  entry_grid.row = i
  entry_grid.SetFocus
  End If
  
  going_to_day = False
  
End Sub

Private Sub entry_grid_DblClick()
  Dim r
  Dim s
  Dim paid_sequence(5)
  Dim row As Integer
  Dim Col As Integer
  
  paid_sequence(0) = 2
  paid_sequence(1) = 3
  paid_sequence(2) = 1
  paid_sequence(3) = 0
  
  ' Bring up the transaction form for this record
  'transaction_form.Show vbModal
  
  With entry_grid
    row = .row
    Col = .Col
    
    ' If this is a cardtrak then do the cardtrak menu
    If (.Col = NAME_COL) And (is_cardtrak_transaction(.TextMatrix(.row, NAME_COL))) Then
      cardtrak_menu_Click (1)
      Exit Sub
    End If
    
    ' If in the check number column
    If (.Col = CHECK_COL) Then
      If (.TextMatrix(.row, THIS_COL) <> "") Then
        ' We have a valid record so see what it is
        If (.TextMatrix(.row, CHECK_COL) <> "") Then
          ' There is a check number so edit it normally
          editing_a_record = True
          entry_grid_KeyPress (0)
          txtEdit.SelStart = 0
          txtEdit.SelLength = 100
        Else
          ' We have a blank check number so insert the current one
          r = .TextMatrix(.row, THIS_COL)
          If (db(r).amount <= 0) Then
            data.last_check_number = data.last_check_number + 1
            db(r).check = data.last_check_number
            If (preferences.auto_check_done_on_check) Then db(r).paid = 1
            changed_flag = True
            update_caption
            process
            ' Put the cursor back here
            .row = row
            .Col = Col
          Else
            s = MsgBox(words(CANT_ASSIGN_A_CHECK_NUMBER_TO_A_DEPOSIT_N), vbOKOnly + vbInformation, words(CHECK_NUMBER_ERROR_N))
          End If
        End If
      End If
    End If
    
    
    
    ' If in the name column then allow for editing
    If (.Col = DUE_COL) Or (.Col = NAME_COL) Or (.Col = AMOUNT_COL) Or (.Col = OVERRIDE_AMOUNT_COL) Then
      If (.TextMatrix(.row, THIS_COL) <> "") Then
        ' We have a valid record so see what it is
        editing_a_record = True
        entry_grid_KeyPress (0)
        txtEdit.SelStart = 0
        txtEdit.SelLength = 100
        End If
    End If
    
    ' If in the paid column then toggle this item
    If (.Col = PAID_COL) Then
      If (.TextMatrix(.row, THIS_COL) <> "") Then
      ' We have a valid record so see what it is
      r = .TextMatrix(.row, THIS_COL)
      db(r).paid = paid_sequence(db(r).paid)  'Not db(r).paid
      If (db(r).paid > 3) Then db(r).paid = 0
      changed_flag = True
      update_caption
      process
      End If
    End If
    
    ' If in the include column then toggle this item
    If (.Col = EXCLUDE_COL) Then
      If (.TextMatrix(.row, THIS_COL) <> "") Then
      ' We have a valid record so see what it is
      r = .TextMatrix(.row, THIS_COL)
      db(r).exclude = Not db(r).exclude
      changed_flag = True
      update_caption
      process
      End If
    End If
    
    ' If in the tags column then toggle this item
    If (.Col = TAG_COL) Then
      If (.TextMatrix(.row, THIS_COL) <> "") Then
        ' We have a valid record so see what it is
        r = .TextMatrix(.row, THIS_COL)
      
        ' Toggle the tag number of the last tag entered
        If ((db(r).tags And tag_mask(last_tag)) = 0) Then
          ' Set the tag
          db(r).tags = db(r).tags Or tag_mask(last_tag)
        Else
          ' Clear the tag
          db(r).tags = db(r).tags And Not (tag_mask(last_tag))
        End If
        changed_flag = True
        update_caption
        process
      End If
    End If
    
    ' If in the cleared column then toggle this item
    If (.Col = CLEARED_COL) Then
      If (.TextMatrix(.row, THIS_COL) <> "") Then
        ' We have a valid record so see what it is
        r = .TextMatrix(.row, THIS_COL)
      
        If (db(r).cleared <= 1) Then
          db(r).cleared = db(r).cleared + 1
          If (db(r).cleared > 1) Then db(r).cleared = 0
          changed_flag = True
          update_caption
          process
        End If
      End If
    End If
    
    ' If in the override column then toggle this item
    If (.Col = OVERRIDE_COL) Then
      If (.TextMatrix(.row, THIS_COL) <> "") Then
      ' We have a valid record so see what it is
      r = .TextMatrix(.row, THIS_COL)
      db(r).override = Not db(r).override
      changed_flag = True
      update_caption
      process
      End If
    End If
  End With
  
  Label1.Caption = Val(r)
End Sub

Sub entry_grid_KeyPress(KeyAscii As Integer)
  If (editing_a_record) Then
    MSFlexGridEdit entry_grid, txtEdit, KeyAscii
  Else
    If (entry_grid.Col = NAME_COL) Then
      If (entry_grid.TextMatrix(entry_grid.row, NAME_COL) <> "") And (KeyAscii <> vbKeyReturn) Then
        insert_menu_Click  ' Insert a new line only if the name is not blank
      End If
      editing_a_record = True
      entry_grid_KeyPress (KeyAscii)
    Else
      If (entry_grid.Col = DUE_COL) Then
        editing_a_record = True
        entry_grid_KeyPress (KeyAscii)
      Else
        If (entry_grid.Col = CHECK_COL) Then
          editing_a_record = True
          entry_grid_KeyPress (KeyAscii)
        Else
          MSFlexGridEdit entry_grid, txtEdit, KeyAscii
        End If
      End If
    End If
  End If
End Sub

Private Sub update_notes()
  ' See if the current notes file wants to be added or deleted
  For i = 0 To MAX_NOTES
    If (notes(i).Month = view.current_month) And _
       (notes(i).Year = view.current_year) Then
      ' Yes we found a notes file
      notes(i).s = notes_box.Text
      
      If (notes(i).s = "") Then
        notes(i).Month = 0
      End If
      Exit For
    End If
  Next i
  
  If (i = MAX_NOTES + 1) And (notes_box.Text <> "") Then
    ' No notes found so loop through all the notes and find a blank one
    For i = 0 To MAX_NOTES
      If (notes(i).s = "") Then
        ' We found a blank spot
        notes(i).Month = view.current_month
        notes(i).Year = view.current_year
        notes(i).s = notes_box.Text
        data.number_of_notes = data.number_of_notes + 1
        Exit For
      End If
    Next i
  End If
  
End Sub

Public Sub entry_tab_Click(PreviousTab As Integer)
  update_notes
  process
End Sub

Private Sub Form_Resize()
  Dim n As Integer
  Dim center_of_buttons As Integer  ' Use to senter the prev/next buttons over the 5th tab
  
  If (main_form.WindowState = 1) Then
    ' We are minimized so minimize all the other windows
    filter_results_form.WindowState = 1
    tags_form.WindowState = 1
    summary_form.WindowState = 1
    balance_form.WindowState = 1
    Exit Sub  ' Don't resize if minimized
  End If
    
  ' ----- Not minimized so make it large
  filter_results_form.normal
  tags_form.normal
  balance_form.normal
  summary_form.normal
  
  If (main_form.height < 2500) Then Exit Sub
  If (main_form.width < 6200) Then Exit Sub
  
  ' ----- Resize the entry grid but do it in increments of the rows
  ' Find how many rows will be displayed
  n = Fix((main_form.height - 2800) / entry_grid.RowHeight(0))
  entry_grid.height = (n * entry_grid.RowHeight(0)) - 100  ' Subtract off a little bit because the last line of the grid was being cut off
  number_of_displayed_rows = n - 1
  
  notes_box.height = entry_grid.height
  notes_box.width = entry_grid.width
  
  ' ----- Resize the table
  entry_grid.width = main_form.width - 200 - entry_grid.Left
  
  ' ----- Resize the entry tab
  entry_tab.width = entry_grid.width - 50
  
  ' ----- Move the balance boxes
  balance_frame.Left = main_form.width - 200 - balance_frame.width
  
  ' ----- Resize the name column
  If (override_columns_menu.Checked) Then
    If (entry_grid.width > NAME_WIDTH_OR) Then entry_grid.ColWidth(NAME_COL) = entry_grid.width - NAME_WIDTH_OR
  Else
    If (entry_grid.width > NAME_WIDTH_NORMAL) Then entry_grid.ColWidth(NAME_COL) = entry_grid.width - NAME_WIDTH_NORMAL
  End If

  ' ----- Resize the notes box
  notes_box.width = entry_grid.width
  notes_box.height = entry_grid.height
  
  ' ----- Resize the misc notes box
  misc_notes_box.width = entry_grid.width
  misc_notes_box.height = entry_grid.height
  
  ' ----- Resize the binder
  binder.height = entry_grid.height
  
  
  ' ----- Resize all the  buttons in the central Next/Previous group -----
  ' ----- Resize the previous and next month buttons
  center_of_buttons = (main_form.width / 2) - 500  ' Subtract off this number of pixels to move the entire button group to the center of the 6th tab
  previous_month_button.Left = center_of_buttons - previous_month_button.width - 300
  next_month_button.Left = center_of_buttons + 200
  todays_date_button.Left = center_of_buttons - (todays_date_button.width / 2) - 40
  
  previous_month_bottom_button.Left = previous_month_button.Left - 300
  bottom_button.Left = previous_month_bottom_button.Left + previous_month_bottom_button.width + 18
  
  top_button.Left = next_month_button.Left + 90
  next_month_top_button.Left = top_button.Left + top_button.width + 18
  
  center_month_button.Left = center_of_buttons - (center_month_button.width / 2) - 40
  
  ' ----- Resize the previous and next year buttons - Added 1/28/2017 ESK
  previous_year_button.Left = previous_month_button.Left - previous_year_button.width - 25
  next_year_button.Left = next_month_button.Left + next_month_button.width + 25
  
  ' ----- Resize the control number frame
  reference_number_frame.Left = main_form.width - reference_number_frame.width - 600
  reference_number_frame.Top = entry_grid.Top + 100
  characters_left_frame.Left = main_form.width - reference_number_frame.width - 600
  characters_left_frame.Top = entry_grid.Top + 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim answer
  
  Call save_preferences
  
  Unload balance_form
  Unload calendar_form
  Unload splash_form
  Unload print_form
  Unload filter_form
  Unload filter_results_form
  Unload tags_form
  Unload summary_form
  Unload integrity_form
  Unload reconcile_form
  Unload edit_transaction_form
  Unload card_data_form
  Unload quick_form
  Unload password_form
  Unload ct_summary_form
  
  
  SaveSetting "Check 2 Check", "Settings", "Directory", start_dir
  If (changed_flag) Then
    answer = MsgBox(words(SAVE_DATABASE_Q_N), vbYesNoCancel + vbQuestion, "Check2Check")
    If answer = vbYes Then
      save_menu_Click
    End If
  End If
  
  play_sound (9)
End Sub

Private Sub goto_month_menu_Click()
  Dim m As Integer
  Dim y As Integer
  
  On Error GoTo error_h  ' This will handle any errors from the goto_month_form

  goto_month_form.show vbModal

  If (goto_month_form.ok) Then
    m = goto_month_form.month_combo.ListIndex + 1  ' Add 1 because months range from 1-12
    y = CInt(goto_month_form.year_combo.Text)
    
    Call set_center_tab(m, y)
    
  End If
  
error_h:
End Sub

Private Sub index_menu_Click()
  ' Show the help file
  help_dialog.HelpFile = startup_directory + "\c2c.hlp"
  help_dialog.HelpKey = ""
  help_dialog.HelpCommand = cdlHelpPartialKey
  help_dialog.ShowHelp
End Sub

Private Sub contents_menu_Click()
  ' Show the help file
  help_dialog.HelpFile = startup_directory + "\c2c.hlp"
  help_dialog.HelpCommand = cdlHelpContents
  help_dialog.ShowHelp
End Sub

Private Sub insert_button_Click()
  insert_menu_Click
End Sub

Private Sub insert_reference_number(what As Integer)
  Dim i As Long
  Dim s1 As String
  Dim s2 As String
  Dim cur As Integer
  
  ' 0 = Insert last number
  ' 1 = Insert last number with name
  ' 2 = Increment and insert
  
  ' Insert the control number
  If (entry_grid.Visible) Then
    ' Entry grid is visible
    ' Insert a new control number on the current transaction name
    With entry_grid
      If (.TextMatrix(.row, THIS_COL) <> "") Then
        ' We have a transaction so add the number to it
        i = Val(GetSetting("Check 2 Check", "Settings", "Reference_number", "0"))
        If (what = 2) Then i = i + 1
        SaveSetting "Check 2 Check", "Settings", "reference_number", Format(i)
        
        cur = Val(.TextMatrix(.row, THIS_COL))
        db(cur).name = db(cur).name + " - (RN-" + Format(i, "###,###,###,###") + ")"
        last_transaction_name = db(cur).name
        changed_flag = True
        update_caption
        process
      End If
    End With
  Else
    If (notes_box.Visible) Then
      ' Entry grid is not visible so we are on the notes
      i = Val(GetSetting("Check 2 Check", "Settings", "Reference_number", "0"))
      If (what = 2) Then i = i + 1
      SaveSetting "Check 2 Check", "Settings", "Reference_number", Format(i)
      ' Get the current position of the cursor in the text box
    
      With notes_box
        cur = .SelStart
        If (what = 1) Then
          ' Insert the last name
          .SelText = Chr(13) + Chr(10) + last_transaction_name + Chr(13) + Chr(10)
        Else
          ' Insert the last number
          .SelText = Chr(13) + Chr(10) + "RN-" + Format(i, "###,###,###,###") + Chr(13) + Chr(10)
        End If
        .SetFocus
        changed_flag = True
        update_caption
      End With
    Else
      If (misc_notes_box.Visible) Then
        ' Entry grid is not visible so we are on the notes
        i = Val(GetSetting("Check 2 Check", "Settings", "Reference_number", "0"))
        If (what = 2) Then i = i + 1
        SaveSetting "Check 2 Check", "Settings", "Reference_number", Format(i)
        ' Get the current position of the cursor in the text box
    
        With misc_notes_box
          cur = .SelStart
          If (what = 1) Then
            ' Insert the last name
            .SelText = Chr(13) + Chr(10) + last_transaction_name + Chr(13) + Chr(10)
          Else
            ' Insert the last number
            .SelText = Chr(13) + Chr(10) + "RN-" + Format(i, "###,###,###,###") + Chr(13) + Chr(10)
          End If
          .SetFocus
          changed_flag = True
          update_caption
        End With
      End If
    End If
  End If
  
  i = Val(GetSetting("Check 2 Check", "Settings", "Reference_number", "0"))
  reference_label.Caption = Format(i, "###,###,###,###")
End Sub

Private Sub update_todays_date()
  todays_date_label(0).Caption = MONTH_STRINGS(Month(Now))  ' Format(Now, "mmm")
  todays_date_label(1).Caption = get_date(Month(Now), day(Now), 0)   'Format(Now, "short date")  '"m/d")
  todays_date_label(2).Caption = DAY_STRINGS(Weekday(Now))  ' Format(Now, "ddd")
End Sub

Private Sub update_language()
  strings_initialize
  
  balance_frame.Caption = words(BALANCE_N)
  beginning_balance_label.Caption = words(BEGINNING_N)
  ending_balance_label.Caption = words(ENDING_N)
  checkbook_label.Caption = words(CHECKBOOK_N)
  
  last_name_button.Caption = words(INSERT_LAST_NAME_N)
  reference_number_last_button.Caption = words(LAST_REFERENCE_NUMBER_N)
  reference_number_button.Caption = words(NEW_REFERENCE_NUMBER_N)
  reference_number_frame.Caption = words(REFERENCE_NUMBER_N)
  characters_left_frame.Caption = words(CHARACTERS_LEFT_N)
  
  transactions_label.Caption = words(TRANSACTIONS_N)
  monthly_notes_label.Caption = words(MONTHLY_NOTES_N)
  misc_notes_label.Caption = words(NOTES_N)
  
  previous_month_button.Caption = words(PREVIOUS_MONTH_N)
  next_month_button.Caption = words(NEXT_MONTH_N)
  
  ' Update the main menu
  file_menu.Caption = words(FILE_N)
  new_menu.Caption = words(NEW_N)
  open_menu.Caption = words(OPEN_N)
  save_menu.Caption = words(SAVE_N)
  save_as_menu.Caption = words(SAVE_AS_N)
  close_menu.Caption = words(CLOSE_N)
  view_print_menu.Caption = words(VIEW_N)
  print_menu.Caption = words(PRINT_N)
  preferences_menu.Caption = words(PREFERENCES_N)
  exit_menu.Caption = words(EXIT_N)
  
  edit_menu.Caption = words(EDIT_N)
  undo_menu.Caption = words(UNDO_N)
  edit_transaction_menu.Caption = words(EDIT_TRANSACTION_N)
  mark_root_menu.Caption = words(MARK_TRANSACTION_WITH_N)
  mark_menu(0).Caption = words(BLANK_N)
  mark_menu(1).Caption = words(DONE_N)
  mark_menu(2).Caption = words(PENDING_N)
  mark_menu(3).Caption = words(SKIP_N)
  tag_menu1(0).Caption = words(TAGS_N)
  set_menu.Caption = words(SET_N)
  tag_set_menu(0).Caption = words(ONE_N)
  tag_set_menu(1).Caption = words(TWO_N)
  tag_set_menu(2).Caption = words(THREE_N)
  tag_set_menu(3).Caption = words(FOUR_N)
  tag_clear_menu(0).Caption = words(CLEAR_N)
  tag_clear_menu(0).Caption = words(ONE_N)
  tag_clear_menu(1).Caption = words(TWO_N)
  tag_clear_menu(2).Caption = words(THREE_N)
  tag_clear_menu(3).Caption = words(FOUR_N)
  filter_menu.Caption = words(FILTER_TRANSACTIONS_N)
  quick_save_menu(1).Caption = words(QUICK_SAVE_N)
  quick_save_menu(2).Caption = words(QUICK_DEPOSIT_N)
  view_quick_menu.Caption = words(QUICK_VIEW_EDIT_N)
  insert_menu.Caption = words(INSERT_N)
  delete_menu.Caption = words(DELETE_N)
  copy_menu.Caption = words(COPY_N)
  cut_menu.Caption = words(CUT_N)
  paste_menu.Caption = words(PASTE_N)
  paste_options_menu.Caption = words(PASTE_OPTIONS_N)
    
  paste_and_clear_menu(0).Caption = words(PASTE_PENDING_N)
  paste_and_clear_menu(1).Caption = words(PASTE_AND_CLEAR_N)
  paste_intact_menu.Caption = words(PASTE_INTACT_N)
  
  copy_selected_menu.Caption = words(COPY_SELECTED_N)
  cut_selected_menu.Caption = words(CUT_SELECTED_N)
  paste_selected_menu.Caption = words(PASTE_SELECTED_N)
  paste_selected_menu.Caption = words(PASTE_SELECTED_INTO_CURRENT_DATE_WITH_STATUS_N)
  paste_selected_option_menu(0).Caption = words(BLANK_N)
  paste_selected_option_menu(1).Caption = words(INTACT_N)
  paste_selected_option_menu(2).Caption = words(PENDING_N)
  copy_month_menu.Caption = words(COPY_MONTH_N)
  cut_month_menu.Caption = words(CUT_MONTH_N)
'
  paste_month_menu.Caption = words(PASTE_MONTH_N)
  
  paste_month_menu.Caption = words(PASTE_MONTH_N)
  'paste_month_menu.Caption = words(PASTE_MONTH_N)
  paste_month_arrange_menu.Caption = words(PASTE_MONTH_AND_ARRANGE_N)
  paste_month_option_root_menu.Caption = words(PASTE_MONTH_WITH_STATUS_N)
  paste_month_option_menu(0).Caption = words(BLANK_N)
  paste_month_option_menu(1).Caption = words(INTACT_N)
  paste_month_option_menu(2).Caption = words(PENDING_N)
  copy_tag_menu.Caption = words(COPY_TAGS_N)
  copy_tags_menu(0).Caption = words(ONE_N)
  copy_tags_menu(1).Caption = words(TWO_N)
  copy_tags_menu(2).Caption = words(THREE_N)
  copy_tags_menu(3).Caption = words(FOUR_N)
  cut_tag_menu.Caption = words(CUT_TAGS_N)
  cut_tags_menu(0).Caption = words(ONE_N)
  cut_tags_menu(1).Caption = words(TWO_N)
  cut_tags_menu(2).Caption = words(THREE_N)
  cut_tags_menu(3).Caption = words(FOUR_N)
  paste_tags_options_menu.Caption = words(PASTE_TAGS_OPTIONS_N)
  
  paste_tags_arrange_menu.Caption = words(PASTE_TAGS_AND_ARRANGE_WITH_STATUS_N)
  paste_tag_arrange_option_menu(0).Caption = words(BLANK_N)
  paste_tag_arrange_option_menu(1).Caption = words(INTACT_N)
  paste_tag_arrange_option_menu(2).Caption = words(PENDING_N)
  paste_tags_current_main_menu.Caption = words(PASTE_TAGS_INTO_CURRENT_DATE_WITH_STATUS_N)
  paste_tags_current_menu(0).Caption = words(BLANK_N)
  paste_tags_current_menu(1).Caption = words(INTACT_N)
  paste_tags_current_menu(2).Caption = words(PENDING_N)
  paste_tags_option_root_menu.Caption = words(PASTE_TAGS_WITH_STATUS_N)
  paste_tag_option_menu(0).Caption = words(BLANK_N)
  paste_tag_option_menu(1).Caption = words(INTACT_N)
  paste_tag_option_menu(2).Caption = words(PENDING_N)
  insert_reference_number_menu.Caption = words(INSERT_REFERENCE_NUMBER_N)
  
  checkbook_menu.Caption = words(CHECKBOOK_N)
  reconcile_menu.Caption = words(RECONCILE_N)
  
  cardtrak_main_menu.Caption = words(CARDTRAK_N)
  cardtrak_menu(0).Caption = words(NEW_TRANSACTION_N)
  cardtrak_menu(1).Caption = words(EDIT_TRANSACTION_N)
  cardtrak_menu(2).Caption = words(CONVERT_TO_CARDTRAK_N)
  cardtrak_menu(3).Caption = words(ADD_DELETE_EDIT_CARDS_N)

  view_menu.Caption = words(VIEW_N)
  view_quick_accounts_menu.Caption = words(QUICK_ACCOUNTS_N)
  next_month_menu.Caption = words(NEXT_MONTH_N)
  previous_month_menu.Caption = words(PREVIOUS_MONTH_N)
  next_month_top_menu.Caption = words(NEXT_MONTH_TOP_N)
  previous_month_bottom_menu.Caption = words(PREVIOUS_MONTH_BOTTOM_N)
  transactions_menu.Caption = words(TRANSACTIONS_N)
  monthly_notes_menu.Caption = words(MONTHLY_NOTES_N)
  misc_notes_menu.Caption = words(NOTES_N)
  calendar_menu.Caption = words(CALENDAR_N)
  view_calculator_menu.Caption = words(CALCULATOR_N)
  override_columns_menu.Caption = words(OVERRIDE_COLUMNS_N)
  goto_month_menu.Caption = words(GO_TO_MONTH_YEAR_N)
  view_balances_menu.Caption = words(BALANCES_N)
  view_tags_menu.Caption = words(TAGS_N)
  view_summary_menu.Caption = words(SUMMARY_N)
  cardtrak_summary_menu.Caption = words(CARDTRAK_SUMMARY_N)
  
  help_menu.Caption = words(HELP_N)
  contents_menu.Caption = words(CONTENTS_N)
  index_menu.Caption = words(INDEX_N)
  register_menu.Caption = words(REGISTER_CHECK2CHECK_N)
  buy_now_menu.Caption = words(BUY_NOW_N)
  update_menu.Caption = words(CHECK_FOR_LATEST_VERSION_VIA_INTERNET_N)
  web_site_menu.Caption = words(CHECK2CHECK_WEB_SITE)
  language_main_menu.Caption = words(LANGUAGE_N)
  language_menu(0).Caption = words(ENGLISH_N)
  language_menu(1).Caption = words(SPANISH_N)
  about_menu.Caption = words(ABOUT_N)
  
  update_entry_tab_names
  update_todays_date
  
  ct_summary_form.update_language
  card_data_form.update_language
End Sub

Private Sub language_menu_Click(index As Integer)
  ' 0=English
  ' 1=Spanish
  preferences.language = index
  update_language
  process
  language_menu(0).Checked = False
  language_menu(1).Checked = False
  language_menu(index).Checked = True
End Sub

Private Sub last_name_button_Click()
  insert_reference_number (1)  ' Insert the last name
End Sub

Private Sub notes_label_Click()
  monthly_notes_menu_Click
End Sub

Private Sub misc_notes_box_Change()
  update_misc_notes_characters_left
End Sub

Private Sub misc_notes_box_Click()
  changed_flag = True
  update_caption
End Sub

Private Sub misc_notes_box_GotFocus()
  misc_notes_box.Text = cards_info(0).name
End Sub

Private Sub misc_notes_box_LostFocus()
  cards_info(0).name = misc_notes_box.Text
End Sub

Private Sub misc_notes_label_Click()
  misc_notes_menu_Click
End Sub

Private Sub monthly_notes_label_Click()
  monthly_notes_menu_Click
End Sub

Private Sub next_month_top_button_Click()
  next_month_top_menu_Click
End Sub

Private Sub next_month_top_menu_Click()
  next_month_menu_Click
  entry_grid.row = 1
  entry_grid.Col = NAME_COL
  entry_grid.TopRow = 1
End Sub

Private Sub next_year_button_Click()
  next_year_menu_Click
End Sub

Private Sub next_year_menu_Click()
  view.start_year = view.start_year + 1  ' Increment the year
  view.start_month = view.start_month - 1  ' Decrement the month knowing that in the next line it will be incremented
  next_month_menu_Click
End Sub

Private Sub paste_menu_Click()
  ' Paste menu
  ' Handle single transactions
  ' Handle multiple selected transactions
  Call paste_and_clear_menu_Click(1)  ' zzzzz I had this commented out and put it back in today 12/13/2017
End Sub

Private Sub previous_month_bottom_button_Click()
  previous_month_bottom_menu_Click
End Sub

Private Sub previous_month_bottom_menu_Click()
  previous_month_menu_Click
  entry_grid.row = entry_grid.Rows - 1
  entry_grid.Col = NAME_COL
  entry_grid.TopRow = entry_grid.Rows - 1
End Sub

Private Sub previous_year_button_Click()
  previous_year_menu_Click
End Sub

Private Sub previous_year_menu_Click()
  view.start_year = view.start_year - 1  ' Decrement the year
  view.start_month = view.start_month + 1  ' Increment the month knowing that in the next line it will be decremented
  previous_month_menu_Click
End Sub

Private Sub reference_number_button_Click()
  insert_reference_number (2)  ' Insert a new number
End Sub

Private Sub insert_reference_number_menu_Click()
  insert_reference_number (2) ' 2 says add a new number
End Sub

Private Sub reference_number_last_button_Click()
  ' Insert the control number but don't increment it
  insert_reference_number (0)  ' Zero says use the current control number without incrementing
End Sub

Private Sub insert_menu_Click()
  With entry_grid
    .Redraw = False
    
    ' See if the current line is non_blank
    If (.TextMatrix(.row, THIS_COL) <> "") Then
      ' Insert a new row below this current line
      shuffle_down_entry_grid (.row + 1)
      If (.row + 1 = .Rows) Then .Rows = .Rows + 1
      .TextMatrix(.row + 1, DATE_COL) = .TextMatrix(.row, DATE_COL)
      .TextMatrix(.row + 1, DAY_COL) = .TextMatrix(.row, DAY_COL)
      .row = .row + 1
    End If
    
    .Redraw = True
    
    'If (.Row > 10) Then .TopRow = .Row - 7
  End With
End Sub

Private Sub mark_menu_Click(index As Integer)
  Dim rec As Integer
  
  
  ' Mark this transaction as necessary
  With entry_grid
    If (.TextMatrix(.row, THIS_COL) = "") Then Exit Sub  ' Do nothing if not on a valid transaction
    
    rec = .TextMatrix(.row, THIS_COL)
    db(rec).paid = index
  End With
  
  changed_flag = True
  update_caption
  
  process
End Sub

Private Sub new_button_Click()
  new_menu_Click
End Sub

Public Sub display_transactions_or_notes(which As Integer)
  ' Make the entry grid visible and the 2 notes invisible
  If (which = 0) Then
    ' Display the transactions
    entry_grid.Visible = True
    notes_box.Visible = False
    misc_notes_box.Visible = False
    
    reference_number_frame.Visible = False
    characters_left_frame.Visible = False
    
    transactions_menu.Checked = True
    monthly_notes_menu.Checked = False
    misc_notes_menu.Checked = False
    
    edit_transaction_menu.Enabled = True
    mark_root_menu.Enabled = True
    
    transactions_label.ForeColor = vbRed
    monthly_notes_label.ForeColor = vbBlue
    misc_notes_label.ForeColor = vbBlue
    transaction_notes_check_label(0).Visible = True
    transaction_notes_check_label(1).Visible = False
    transaction_notes_check_label(2).Visible = False
  End If
    
  If (which = 1) Then
    ' Display the monthly notes
    entry_grid.Visible = False
    notes_box.Visible = True
    misc_notes_box.Visible = False
    
    reference_number_frame.Visible = True
    characters_left_frame.Visible = False
    
    transactions_menu.Checked = False
    monthly_notes_menu.Checked = True
    misc_notes_menu.Checked = False
    
    edit_transaction_menu.Enabled = False
    mark_root_menu.Enabled = False
  
    transactions_label.ForeColor = vbBlue
    monthly_notes_label.ForeColor = vbRed
    misc_notes_label.ForeColor = vbBlue
    transaction_notes_check_label(0).Visible = False
    transaction_notes_check_label(1).Visible = True
    transaction_notes_check_label(2).Visible = False
    
    notes_box.SelStart = 30000  ' Put the cursor at the end of the notes
  End If
  
  If (which = 2) Then
    ' Display the misc notes
    entry_grid.Visible = False
    notes_box.Visible = False
    misc_notes_box.Visible = True
    reference_number_frame.Visible = False
    characters_left_frame.Visible = True
    
    transactions_menu.Checked = False
    monthly_notes_menu.Checked = False
    misc_notes_menu.Checked = True
    
    edit_transaction_menu.Enabled = False
    mark_root_menu.Enabled = False
  
    transactions_label.ForeColor = vbBlue
    monthly_notes_label.ForeColor = vbBlue
    misc_notes_label.ForeColor = vbRed
    
    transaction_notes_check_label(0).Visible = False
    transaction_notes_check_label(1).Visible = False
    transaction_notes_check_label(2).Visible = True
  End If
  
End Sub

Public Sub new_menu_Click()
  Dim answer
  Dim m As Integer
  
  If (changed_flag) Then
    answer = MsgBox(words(SAVE_DATABASE_Q_N), vbYesNoCancel + vbQuestion + vbMsgBoxSetForeground, "Check2Check")
    If answer = vbCancel Then
      Exit Sub
    Else
       If answer = vbYes Then
        save_menu_Click
       End If
    End If
  End If
  
  view.start_day = 1
  m = Month(Now) - 5
  view.start_year = Year(Now)
  If (m <= 0) Then
    view.start_month = 12 + m
    view.start_year = view.start_year - 1
  Else
    view.start_month = m
  End If
  
  view.current_month = view.start_month
  view.current_year = view.start_year
  ' Clear out stuff for a new form
  'copy_of_this.this = -1  ' Start off with nothing in the copy buffer
  entry_tab.Tab = 5
  doing_move = False
  update_entry_tabs
  
  display_transactions_or_notes (0)
  
  undo_menu.Enabled = False
  undo_button.Visible = False
  editing_a_record = False
  
  data.db_name = "Untitled"
  changed_flag = False
  update_caption
  
  ' Start with a fresh database
  data.first = 0
  data.last = 0
  data.number_of_records = 0
  data.number_of_notes = 0
  data.last_check_number = 0
  
  ' Zero out the data table
  For i = 0 To MAX_DATA_TABLE
    db(i).this = -1
  Next i
  
  ' Zero out the notes file
  For i = 0 To MAX_NOTES
    notes(i).Month = 0
    notes(i).s = ""
  Next i

  ' Clear out the misc notes box
  misc_notes_box.Text = ""
  cards_info(0).name = ""
  
  ' Don't zero out the quick accounts
    
  ' Zero out Cardtrak transactions
  card_data_form.zero
  
  process
End Sub

Private Sub next_month_button_Click()
  next_month_menu_Click
End Sub

Public Sub next_month_menu_Click()
  Dim r
  
  r = entry_grid.row
  
  view.start_month = view.start_month + 1
  If (view.start_month > 12) Then
    view.start_month = 1
    view.start_year = view.start_year + 1
  End If
  update_entry_tabs
  
  If (r < entry_grid.Rows) Then
    entry_grid.row = r
  Else
    entry_grid.row = entry_grid.Rows - 1
  End If
  
End Sub

Private Sub notes_box_Click()
  changed_flag = True
  update_caption
End Sub

Private Sub notes_box_KeyPress(KeyAscii As Integer)
  changed_flag = True
End Sub

Private Sub monthly_notes_menu_Click()
  i = Val(GetSetting("Check 2 Check", "Settings", "Reference_number", "0"))
  reference_label.Caption = Format(i, "###,###,###,###")
  
  ' Display the entry_grid
  If (monthly_notes_menu.Checked = False) Then
    display_transactions_or_notes (1)
  End If
End Sub

Private Sub misc_notes_menu_Click()
  i = Val(GetSetting("Check 2 Check", "Settings", "Reference_number", "0"))
  reference_label.Caption = Format(i, "###,###,###,###")
  
  ' Display the entry_grid
  
  If (misc_notes_menu.Checked = False) Then
    display_transactions_or_notes (2)
  End If
End Sub

Private Sub notes_picture_Click()
  monthly_notes_menu_Click
End Sub

Private Sub notes_radio_Click()
  monthly_notes_menu_Click
  update_notes
  
  show_buttons
End Sub

Private Sub open_button_Click()
  open_menu_Click
End Sub

Private Sub open_the_file(showit As Boolean)
  Dim s, found
  
  With open_dialog
    If (showit) Then
        .DefaultExt = "c2c"
        .InitDir = start_dir
        .filter = "Check2Check files | *.c2c"
        .ShowOpen
        start_dir = CurDir()
    End If
    
    s = GetExtension(.Filename)
    found = False
    If (UCase(s) = "C2C") Then
      found = True
    End If
    
    If (.Filename <> "") And (found) Then
       ' We have a filename so strip off the extension
      j = InStr(1, .Filename, ".")
      If (j > 0) Then
        k = InStr(j + 1, .Filename, ".")
        If (k = 0) Then k = j
        data.db_name = Left(.Filename, k - 1)
      End If
      
      ' Zero out the notes file
      For i = 0 To MAX_NOTES
        notes(i).Month = 0
        notes(i).s = ""
      Next i
      
      misc_notes_box.Text = ""
      
      update_caption
      
      On Error GoTo error_h
      s = Dir(.Filename)
      If (s = "") Then
        Err.Description = "File not found"
        GoTo error_h
      End If
      
      If (FileLen(s) <= 0) Then
        Err.Description = "File length is zero"
        GoTo error_h
      End If
     
      read_database
      GoTo continue
      
error_h:
    s = MsgBox(Err.Description, vbOK + vbInformation, words(ERROR_N))
      
continue:
      misc_notes_box.Text = cards_info(0).name
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
    answer = MsgBox(words(SAVE_DATABASE_Q_N), vbYesNoCancel + vbQuestion, "Check2Check")
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

Private Sub paste_and_clear_menu_Click(index As Integer)
  Dim t As r_type
  
  ' See if we are doing the notes
  If (monthly_notes_menu.Checked = True) Then
    paste_intact_menu_Click
    Exit Sub
  End If
  
  
  ' Paste the record now
  If (copy_of_this.this >= 0) Then
    ' We have a valid record
    t = copy_of_this  ' Save a copy of the transaction
    copy_of_this.check = -1
    If (index = 0) Then
      copy_of_this.check = -1
      copy_of_this.cleared = 0
      copy_of_this.paid = PAID_QUESTION
    Else
      copy_of_this.check = -1
      copy_of_this.cleared = 0
      copy_of_this.paid = PAID_BLANK
    End If
    paste_intact_menu_Click
    copy_of_this = t  ' Put it back the way it was
    allow_paste_ct = False
  End If
End Sub

Private Sub paste_button_Click()
  paste_intact_menu_Click
End Sub

Private Sub paste_intact_menu_Click()
  Dim s
  
  changed_flag = True
  update_caption
  
  ' See if we are editing a field
  If (txtEdit.Visible) Then
    ' Yes we are editing a field
    s = Clipboard.GetText
    txtEdit.SelText = Clipboard.GetText
    Exit Sub
  End If
  
  ' See if we are doing the notes
  If (monthly_notes_menu.Checked = True) Then
    s = Clipboard.GetText
    notes_box.SelText = Clipboard.GetText
    Exit Sub
  End If
  
  
  If (copy_of_this.this >= 0) Then
    ' We have a valid record
    this = copy_of_this
    this.Month = view.current_month
    this.Year = view.current_year
    If (Not doing_move) Then  ' Doing a move so use the day that was given by mouse up
      this.day = Int(entry_grid.TextMatrix(entry_grid.row, DATE_COL))
    End If
    If (this.sub_transaction_number > 0) Then cards(this.sub_transaction_number).active = True
    insert_record (-1)
    
    ' Save the current record number for undo
    undo.rec_num = data.current
    If (Not doing_move) Then
      undo.what_was_done = WHAT_PASTE_RECORD
      undo_menu.Caption = "Undo - Paste transaction"
    Else
      undo_menu.Caption = "Undo - Move transaction"
    End If
    
    undo_menu.Enabled = True
    undo_button.Visible = undo_menu.Enabled
    Label1.Caption = undo.rec_num
    
    process
  End If
End Sub

Private Sub paste_month_normal_menu_Click()
  paste_month_options = PM_NORMAL
  paste_month_arrange_menu_Click
End Sub

Private Sub paste_month_arrange()
  Dim days_adjusted
  Dim index As Integer
  
  If (MsgBox(words(PASTE_ALL_TRANSACTIONS_TO_N) + entry_tab.Caption, _
      vbYesNoCancel + vbQuestion + vbApplicationModal, words(PASTE_MONTH_N)) = vbYes) Then
    ' Yes, Paste the entire month
  
    ' Paste the month buffer to the new month
    ' Be sure to check for the last day and overruns
    days_adjusted = False
    With copy_of_month
      For i = 0 To MAX_RECORDS_IN_MONTH
        If (.table(i).this <= -1) Then Exit For
        ' We have a record to transfer
        this = .table(i)
        
        ' See how to set the paid colummn
        this.paid = preferences.paste_month(this.paid)
        
        ' See how to set the check number colummn
        If (preferences.paste_month(PREF_MONTH_CHECK_NUMBER_INDEX) = 1) Then this.check = -1  ' Set to blank
        
        ' See how to set the cleared colummn
        If (preferences.paste_month(PREF_MONTH_CLEARED_INDEX) = 1) Then this.cleared = 0  ' Set to blank
        
        ' Make the current day the due day if we are doing arrange
        If (paste_month_options = PM_ARRANGE) Then
          If (this.due > 0) And (this.due <= 31) Then  ' We are doing arrange
            this.day = this.due
          End If
        End If
        
        ' See if we have to adjust the days
        If (this.day > view.number_of_days) Then
          ' We must adjust the days because there are less days in the month
          this.day = view.number_of_days
          days_adjusted = True
        End If
        this.Month = view.current_month
        this.Year = view.current_year
        
        ' ---- Restore the saved cardtrack ----
        index = insert_cardtrak_record(this, copy_of_cardtrak_month.table(i))  ' Returned index points to the ct record in the ct db
        this.sub_transaction_number = index
        If (Not allow_paste_ct) Then this.sub_transaction_number = 0
        
        insert_record (-1)
        
        ' Save the undo stuff
        If (Not pasting_tags) Then
          undo.what_was_done = WHAT_PASTE_MONTH
        Else
          undo.what_was_done = WHAT_PASTE_TAGS
        End If
        
        undo.copy_of_month.table(i).this = data.current
        undo_cardtrak_month.table(i) = cards(this.sub_transaction_number)
        If (i < MAX_RECORDS_IN_MONTH) Then undo.copy_of_month.table(i + 1).this = -1
        undo_menu.Enabled = True
        undo_button.Visible = undo_menu.Enabled
      Next i
    End With
  
    ' Save the undo notes
    undo.copy_of_month.notes = notes_box.Text
    undo_menu.Caption = "Undo - Paste " + MONTH_STRINGS(view.current_month) + " " + Format(view.current_year)
    undo.copy_of_month.Month = view.current_month
    undo.copy_of_month.Year = view.current_year
    
    If (undo.what_was_done = WHAT_PASTE_MONTH) Then
      ' Only paste notes if it was past month, and not paste tags
      ' Paste the notes now
      If (preferences.prompt_for_paste_notes) Then
        If (MsgBox(words(PASTE_NOTES_Q_N), _
            vbYesNoCancel + vbQuestion + vbApplicationModal, words(PASTE_MONTH_N)) = vbYes) Then
          notes_box.Text = notes_box.Text + copy_of_month.notes
          update_notes
        End If
      End If
    End If
    
    process
    changed_flag = True
    update_caption
    If (days_adjusted) Then MsgBox words(ADJUSTED_DATES_TO_MATCH_THE_CURRENT_MONTH_N)
  
  End If
  paste_month_options = PM_ARRANGE
End Sub

Private Sub paste_month_arrange_menu_Click()
  Call paste_month_arrange
End Sub

Private Sub paste_month_option(index As Integer)
  Dim days_adjusted, ans
  Dim ct_index As Integer
  
  ans = vbYes
  ' Prompt if doing month or tagged paste
  If (Not paste_selected_active) Then
    ans = MsgBox(words(PASTE_ALL_TRANSACTIONS_TO_N) + " " + entry_tab.Caption, _
          vbYesNoCancel + vbQuestion + vbApplicationModal, words(PASTE_MONTH_N))
  End If
  
  If (ans = vbYes) Then
    ' Yes, Paste the entire month
  
    ' Paste the month buffer to the new month
    ' Be sure to check for the last day and overruns
    days_adjusted = False
    With copy_of_month
      'If (.month = view.current_month) And (.year = view.current_year) Then
        ' Can't copy to the same month and year
        'MsgBox words(Cant_paste_to_the_same_month_and_year_n)
        'Exit Sub
      'End If
    
      For i = 0 To MAX_RECORDS_IN_MONTH
        If (.table(i).this <= -1) Then Exit For
        ' We have a record to transfer
        this = .table(i)
        
        ' See how to set the paid colummn
        Select Case index
          Case 0  ' Clear it out
            this.paid = 0
            this.check = -1  ' Blank the check column
          Case 1  ' Intact
            ' Do nothing
          Case 2  ' Question mark
            If (this.paid = 1) Then this.paid = 2
            this.check = -1  ' Blank the check column
        End Select
        
        ' See how to set the cleared colummn
        If (Not doing_move) Then If (preferences.paste_month(PREF_MONTH_CLEARED_INDEX) = 1) Then this.cleared = 0  ' Set to blank
        
        ' Make the current day the due day if we are doing arrange
        If (paste_month_options = PM_ARRANGE) Then
          If (this.due > 0) And (this.due <= 31) Then  ' We are doing arrange
            this.day = this.due
          End If
        End If
        
        ' Make the current day the due day if we are doing arrange
        If (paste_month_options = PM_CURRENT) Then this.day = view.current_day
        
        If (this.day > view.number_of_days) Then
          ' We must adjust the days because there are less days in the month
          this.day = view.number_of_days
          days_adjusted = True
        End If
        
        this.Month = view.current_month
        this.Year = view.current_year
        
        ' ---- Restore the saved cardtrack ----
        If (Not allow_paste_ct) Then
          this.sub_transaction_number = 0
        Else
          If (copy_of_cardtrak_month.table(i).active) Then
            ct_index = insert_cardtrak_record(this, copy_of_cardtrak_month.table(i))  ' Returned index points to the ct record in the ct db
            this.sub_transaction_number = ct_index
          Else
            this.sub_transaction_number = 0
          End If
        End If
        
        ' Save the record now
        insert_record (-1)
        
        ' Save the undo stuff
        If (undo.what_was_done <> WHAT_MOVE_SELECTED) Then
          If (Not pasting_tags) Then
            undo.what_was_done = WHAT_PASTE_MONTH
          Else
            undo.what_was_done = WHAT_PASTE_TAGS
          End If
        End If
        
        If (undo.what_was_done <> WHAT_MOVE_SELECTED) Then
          undo.copy_of_month.table(i).this = data.current
          undo_cardtrak_month.table(i) = cards(this.sub_transaction_number)
          If (i < MAX_RECORDS_IN_MONTH) Then undo.copy_of_month.table(i + 1).this = -1
        Else
          undo.selected_rec_num(i) = data.current
          If (i < MAX_RECORDS_IN_MONTH) Then undo.selected_rec_num(i + 1) = -1
        End If
        
        undo_menu.Enabled = True
        undo_button.Visible = undo_menu.Enabled
      Next i
    End With
  
    ' Save the undo notes
    undo.copy_of_month.notes = notes_box.Text
    undo_menu.Caption = "Undo - Paste and clear " + MONTH_STRINGS(view.current_month) + " " + Format(view.current_year)
    undo.copy_of_month.Month = view.current_month
    undo.copy_of_month.Year = view.current_year
    
    If (undo.what_was_done = WHAT_MOVE_SELECTED) Then
      undo_menu.Caption = "Undo - Move Selected"
    Else
      undo_menu.Caption = "Undo - Paste " + MONTH_STRINGS(view.current_month) + " " + Format(view.current_year)
    End If
    
    If (undo.what_was_done = WHAT_PASTE_MONTH) Then
      ' Only paste notes if it was past month, and not paste tags
      ' Paste the notes now
      If (preferences.prompt_for_paste_notes) Then
        If (MsgBox(words(PASTE_NOTES_Q_N) + " ", _
            vbYesNoCancel + vbQuestion + vbApplicationModal, "Paste Month") = vbYes) Then
          notes_box.Text = notes_box.Text + copy_of_month.notes
          update_notes
        End If
      End If
    End If
    
    process
    changed_flag = True
    update_caption
    allow_paste_ct = False
    If (days_adjusted) Then MsgBox words(ADJUSTED_DATES_TO_MATCH_THE_CURRENT_MONTH_N)
  
  End If
End Sub

Private Sub paste_month_option_menu_Click(index As Integer)
  paste_month_options = PM_NORMAL
  Call paste_month_option(index)
  paste_month_options = PM_ARRANGE
End Sub

Private Sub paste_selected_option_menu_Click(index As Integer)
  paste_selected_active = True
  pasting_tags = True
  paste_month_options = PM_CURRENT
  Call paste_month_option(index)
  pasting_tags = False
  paste_month_options = PM_ARRANGE
  paste_selected_active = False
  allow_paste_ct = False  ' Don't allow any more pastes since we did them all once
End Sub

Private Sub paste_tag_menu_Click()
  pasting_tags = True
  paste_month_menu_click
  pasting_tags = False
End Sub

Private Sub preferences_menu_Click()
  preferences_form.show (vbModal)
  notes_box.FontSize = preferences.notes_font_size
  misc_notes_box.FontSize = preferences.notes_font_size
End Sub

Public Sub previous_month_menu_Click()
  Dim r
  
  r = entry_grid.row
  
  view.start_month = view.start_month - 1
  If (view.start_month < 1) Then
    view.start_month = 12
    view.start_year = view.start_year - 1
  End If
  update_entry_tabs
  
  If (r < entry_grid.Rows) Then
    entry_grid.row = r
  Else
    entry_grid.row = entry_grid.Rows - 1
  End If
End Sub

Private Sub print_button_Click()
  print_menu_Click
End Sub

Private Sub print_menu_Click()
  Dim f, i
  
  printer_error = False
  print_destination = PTR
  
  f = cdlPDNoPageNums Or cdlPDNoSelection Or cdlPDUseDevModeCopies
  'f = cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection Or cdlPDUseDevModeCopies Or cdlPDNoPageNums
  print_dialog.FLAGS = f
  
  On Error GoTo error_h
  print_dialog.Copies = 1
  print_dialog.ShowPrinter
  
  ' We have hit ok so print it out
  main_form.MousePointer = 11
  Call view_printout
  
  print_form.Hide
  main_form.MousePointer = 0
  
  ' Bump up the printer counter
  i = Val(GetSetting("Check 2 Check", "Settings", "Printouts", "0"))
  i = i + 1
  SaveSetting "Check 2 Check", "Settings", "Printouts", Format(i)
  Exit Sub
error_h:
  
  main_form.MousePointer = 0
End Sub
Private Sub clear_this()
  this.amount = 0
  this.check = -1
  this.cleared = 0
  this.day = 1
  this.due = 0
  this.exclude = False
  this.Month = view.current_month
  this.day = view.current_day
  this.override_amount = 0
  this.override = False
  this.paid = 0
  this.tags = 0
  this.sub_transaction_number = 0  ' 3/15/03
  this.Year = view.current_year
End Sub

Private Sub quick_save_menu_Click(index As Integer)
  Dim d As String
  Dim c As Integer
  Dim r As Integer
  
  With entry_grid
  ' Put up the Quick Save form
  If (quick_form.execute(index)) Then
    c = .Col
    
    ' We have a successful save
    clear_this
    this.name = quick_form.name_s
    this.amount = quick_form.amount
    this.paid = quick_form.status
    
    insert_menu_Click
    r = .row
    .TextMatrix(.row, NAME_COL) = this.name
    .TextMatrix(.row, AMOUNT_COL) = this.amount
    status = quick_form.status
    
    insert_row_into_database
    changed_flag = True
    update_caption
  End If
  End With
  process
  entry_grid.row = r
  entry_grid.Col = c
End Sub

Private Sub quick_view_button_Click()
  view_quick_accounts_menu_Click
End Sub

Private Sub recent_file_menu_Click(index As Integer)
  Dim s As String
  Dim answer
  
  If (changed_flag) Then
    answer = MsgBox(words(SAVE_DATABASE_Q_N), vbYesNoCancel + vbQuestion, "Check2Check")
    If answer = vbCancel Then
      Exit Sub
    Else
       If answer = vbYes Then
        save_menu_Click
       End If
    End If
  End If
  
  s = recent_file_menu(index).Caption
  
  On Error GoTo error_h
  
  open_dialog.Filename = s  'Dir(s)
  open_dialog.InitDir = strip_filename(s)
  start_dir = strip_filename(s)
  ChDir start_dir
  ChDrive get_drive(s)
  Call open_the_file(False)
  
  Exit Sub
error_h:
  answer = MsgBox(words(FILE_NOT_FOUND_N), vbOK + vbInformation, "")
End Sub

Private Sub save_as_menu_Click()
  Dim answer
  
  ' Update those things that don't happen when process or lost focus isn't executed
  update_notes
  cards_info(0).name = misc_notes_box.Text
  
  On Error GoTo errorh
  
  With open_dialog
    .filter = "Check2Check files | *.c2c"
    .ShowSave
    start_dir = CurDir()
    If (.Filename <> "") Then
      ' We have a filename so see if it exists
      If Dir(open_dialog.Filename) <> "" Then
        ' File exists so see if we should overwrite it
        If Dir(open_dialog.Filename) <> "" Then
          answer = MsgBox(words(FILE_EXISTS_OVERWRITE_Q_N), vbYesNoCancel + vbQuestion, words(OVERWRITE_Q_N))
          If answer = vbYes Then
            clear_attributes (open_dialog.Filename)
            Kill (open_dialog.Filename)
          Else
            Exit Sub
          End If
        End If
        
      End If
      
      ' We have a filename so strip off the extension
      j = InStr(1, .Filename, ".")
      If (j > 0) Then
        k = InStr(j + 1, .Filename, ".") ' See if there is a 2nd dot
        If (k = 0) Then k = j
        data.db_name = Left(.Filename, k - 1)
      End If
    Caption = data.db_name
    write_database
    If (preferences.save_recovery_file) Then write_backup_database
    changed_flag = False
    update_caption
    End If
  End With
  
  Save_Recent_Files
  
  Exit Sub
  
errorh:
  MsgBox words(ERROR_IN_DELETING_WRITING_FILE_N)
End Sub

Private Sub previous_month_button_Click()
  previous_month_menu_Click
End Sub

Private Sub save_button_Click()
  save_menu_Click
End Sub

Private Sub save_menu_Click()
  If (data.db_name = "") Or (data.db_name = "Untitled") Then
    save_as_menu_Click
  Else
    update_notes
    cards_info(0).name = misc_notes_box.Text
    write_database
    If (preferences.save_recovery_file) Then write_backup_database
    changed_flag = False
    update_caption
    
    Save_Recent_Files
  End If
End Sub

Private Sub splash_timer_Timer()
  splash_timer.Enabled = False
  If splash_form.Visible Then splash_form.Hide
End Sub

Private Sub tag_clear_menu_Click(index As Integer)
  ' 0 - MAX_TAG
  Dim rec As Integer
  Dim s As String
  
  ' Get the transaction number
  With entry_grid
    s = .TextMatrix(.row, THIS_COL)
    If (s = "") Then Exit Sub
    rec = Val(s)
    db(rec).tags = db(rec).tags And Not tag_mask(index)
  End With
  
  changed_flag = True
  update_caption
  
  process

End Sub

Private Sub tag_set_menu_Click(index As Integer)
  ' 0 - MAX_TAG
  Dim rec As Integer
  
  ' Get the transaction number
  With entry_grid
    If (.TextMatrix(.row, THIS_COL) = "") Then Exit Sub  ' Do nothing if not on a valid transaction
    
    rec = .TextMatrix(.row, THIS_COL)
    db(rec).tags = db(rec).tags Or tag_mask(index)
  End With
  
  changed_flag = True
  update_caption
  
  process
End Sub

Private Sub todays_date_button_Click()
  todays_date_label_Click (0)
End Sub

Private Sub todays_date_label_Click(index As Integer)
  ' Make the current view today
  Dim m As Integer
  Dim y As Integer
  
  m = Val(Format(Now, "mm"))  ' Get today's date
  y = Val(Format(Now, "yyyy"))
  
  Call set_center_tab(m, y)
  Exit Sub
End Sub

Private Sub top_button_Click()
  entry_grid.row = 1
  entry_grid.TopRow = 1
End Sub

Private Sub transaction_notes_check_label_Click(index As Integer)
  If (index = 0) Then transactions_menu_Click
  If (index = 1) Then monthly_notes_menu_Click
  If (index = 2) Then misc_notes_menu_Click
End Sub

Private Sub transactions_label_Click()
  transactions_menu_Click
End Sub

Private Sub transactions_menu_Click()
  ' Display the entry_grid
  If (transactions_menu.Checked = False) Then
    display_transactions_or_notes (0)
  End If
End Sub

Private Sub transactions_picture_Click()
  transactions_menu_Click
End Sub

Private Sub transactions_radio_Click()
  transactions_menu_Click
  update_notes
  
  ' Hide the necessary menu items
  show_buttons
End Sub

Sub txtEdit_KeyPress(KeyAscii As Integer)
    ' Delete returns to get rid of beep.
    If KeyAscii = Val(vbCr) Then KeyAscii = 0
End Sub

Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode entry_grid, txtEdit, KeyCode, Shift
    If (KeyCode = 27) Then editing_a_record = False  ' See if escape was pressed
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
    If (.TextMatrix(.row, DUE_COL) = "") Then .TextMatrix(.row, DUE_COL) = "0"
    If (.TextMatrix(.row, CHECK_COL) = "") Then .TextMatrix(.row, CHECK_COL) = "-1"
    If (.TextMatrix(.row, AMOUNT_COL) = "") Then .TextMatrix(.row, AMOUNT_COL) = "0.0"
    
    this.day = .TextMatrix(.row, DATE_COL)
    this.Month = view.current_month
    this.Year = view.current_year
    this.due = .TextMatrix(.row, DUE_COL)
    this.check = .TextMatrix(.row, CHECK_COL)
    this.name = .TextMatrix(.row, NAME_COL)
    this.amount = .TextMatrix(.row, AMOUNT_COL)
    ' See what tags are flagged
    this.tags = 0
    'this.paid = 0
    this.paid = status
    this.override = False
    this.override_amount = 0
    this.sub_transaction_number = 0  ' 3/15/03
    this.cleared = 0
    
    If (preferences.auto_check_done) And (.Col = AMOUNT_COL) Then this.paid = 1
    If (preferences.auto_check_done_on_check) And (.Col = CHECK_COL) Then this.paid = 1
    
    ' See if it belongs after any other ones in the same day
    If (.row > 1) Then
      ' We are not on the first row
      If (this.day = .TextMatrix(.row - 1, DATE_COL)) Then
        ' Yes there is another record before this one with the same day
        ' Get the record number of the previous one
        r = .TextMatrix(.row - 1, THIS_COL)  ' r contains the record number to insert after
      End If
    End If
    ' Insert the record now
    If (r = "") Then r = -1
    Call insert_record(r)   ' If r=-1 then insert it at the proper date, if r >= 0 then insert it after that record
  .Redraw = True
  End With
  Exit Sub

errorh:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
  entry_grid.Redraw = True
End Sub


Sub entry_grid_GotFocus()
    Dim rec
    Dim n
    Dim ans
    
    ' Data was entered on the grid
    ' Data entry
    
    On Error GoTo errorh
    If txtEdit.Visible = False Then Exit Sub
    entry_grid = txtEdit
    txtEdit.Visible = False
        
    
    
    ' We just entered data so now see if this is on a new record or exising record
    With entry_grid
      last_row = .row
      last_col = .Col
      If ((.Col = NAME_COL) And (.TextMatrix(.row, AMOUNT_COL) = "")) Then
        .TextMatrix(.row, AMOUNT_COL) = "0"
        .TextMatrix(.row, PAID_COL) = "?"
      End If
      
      If (.TextMatrix(.row, THIS_COL) = "") Then
        ' We have a new record
        status = PAID_QUESTION
        
        ' ----------- Insert a new row into the database -------------
        If (.Text <> "") Then
          ' Don't allow blank fields to be entered
          insert_row_into_database
        End If
        ' Set the pending flag
        'db(rec).paid = PAID_QUESTION
      Else
        ' We have an existing record so update the data
        rec = .TextMatrix(.row, THIS_COL)
        
        ' -------------------- Due column ---------------------
        If (.Col = DUE_COL) Then
          db(rec).due = 0
          If (txtEdit <> "") Then db(rec).due = txtEdit
        End If
        
        ' -------------------- Check number column ---------------------
        If (.Col = CHECK_COL) Then
          'db(rec).check = -1
          If (db(rec).amount <= 0) Then
            If (txtEdit <> "") Then
              n = db(rec).check
              db(rec).check = txtEdit
              ' If the check number is the same as the previous one then don't set the last check number
              If (n <> db(rec).check) Then data.last_check_number = txtEdit  ' Save the last check number
            Else
              ' No check number entered so see if there is an existing check number, if yes then blank it out, if no then apply this new one
              If (db(rec).check >= 0) Then
                ' We have an existing check number so clear it out
                db(rec).check = -1
              Else
                ' We don't currently have a check number so use the default one
                data.last_check_number = data.last_check_number + 1
                db(rec).check = data.last_check_number
              End If
            End If
            If (preferences.auto_check_done_on_check) Then db(rec).paid = 1
          Else
            s = MsgBox(words(CANT_ASSIGN_A_CHECK_NUMBER_TO_A_DEPOSIT_N), vbOKOnly + vbInformation, words(CHECK_NUMBER_ERROR_N))
          End If
        End If
        
        ' ------------------ Name column ---------------
        If (.Col = NAME_COL) Then db(rec).name = txtEdit
        
        ' ------------------ Amount column ---------------
        If (.Col = AMOUNT_COL) Then
          ' See if txtEdit has a sign on it
          If (Mid(LTrim(txtEdit.Text), 1, 1) <> "-") And _
             (Mid(LTrim(txtEdit.Text), 1, 1) <> "+") Then
            If (db(rec).amount < 0.001) And (preferences.auto_negative_numbers) Then
              txtEdit.Text = "-" + txtEdit.Text
            End If
          End If
            
          ' Save the newly entered amount into the database
          db(rec).amount = txtEdit
          If (preferences.auto_check_done) Then db(rec).paid = 1
        End If
        
        ' --------------------- Override amount column -----------------
        If (.Col = OVERRIDE_AMOUNT_COL) Then db(rec).override_amount = txtEdit
        
        ' ------------------ Paid column -----------------
        If (.Col = PAID_COL) Then
          ' Display the paid column menu
          db(rec).paid = db(rec).paid + 1
          If (db(rec).paid > 3) Then db(rec).paid = 0
        End If
        
        ' ------------------- Exclude column ----------------------
        If (.Col = EXCLUDE_COL) Then db(rec).exclude = Not db(rec).exclude
        
        ' -------------------- Tag column ---------------------
        If (.Col = TAG_COL) Then
          n = Val(txtEdit.Text) - 1
          If (n >= 0) And (n <= MAX_TAG) Then
            last_tag = n  ' Save the tag number for subsequent operations
            If ((db(rec).tags And tag_mask(n)) = 0) Then
              ' Set the tag
              db(rec).tags = db(rec).tags Or tag_mask(n)
            Else
              ' Clear the tag
              db(rec).tags = db(rec).tags And Not (tag_mask(n))
            End If
          Else
            ' Put up a message for invalid tag number
            MsgBox words(INVALID_TAG_NUMBER_N)
          End If
        End If
        
        ' ----------------------- Cleared column ---------------------
        If (.Col = CLEARED_COL) Then
          If (db(rec).cleared = 2) Or (UCase(txtEdit.Text) = "X") Then
            ' We have a perm cleared item already or we want to perm clear it
            ans = MsgBox(words(CHANGING_THIS_MAY_CAUSE_A_PROBLEM_WHEN_YOU_RECONCILE_Q_N), _
                vbYesNoCancel + vbQuestion + vbApplicationModal, words(CAUTION_TRANSACTION_PERMANENTLY_MARKED_N))
            If (ans = vbNo) Or (ans = vbCancel) Then
              txtEdit.Text = "-"
            End If
          End If
          
          ' Mark it the right way
          If (txtEdit.Text = "") Then
            If (db(rec).cleared = 0) Then
              db(rec).cleared = 1
            Else
              If (db(rec).cleared = 1) Then db(rec).cleared = 0
            End If
          End If
          
          If (txtEdit.Text = " ") Then
            db(rec).cleared = 0  ' Mark it as not cleared
          Else
            If (txtEdit.Text = "*") Then
              db(rec).cleared = 1  ' Mark it as cleared
              txtEdit.Text = " "
            Else
              If (UCase(txtEdit.Text) = "X") Then
                db(rec).cleared = 2  ' Mark it as perm cleared
              Else
                ' Invalid key key - Use only space, * and X
                If (txtEdit.Text <> "-") And (txtEdit.Text <> "") Then
                  MsgBox words(INVALID_KEY_N)
                End If
              End If
            End If
          End If
        End If
        
        ' --------------------- Override column -------------------
        If (.Col = OVERRIDE_COL) Then db(rec).override = Not db(rec).override
        End If
        
      rec = .row
    End With
    
    editing_a_record = False
    Call process
    
    entry_grid.row = rec
    
    If (last_col = DUE_COL) Then
      entry_grid.Col = CHECK_COL
      entry_grid.row = last_row
    Else
      If last_col = CHECK_COL Then
          entry_grid.Col = CHECK_COL
          If last_row + 1 < entry_grid.Rows Then
              entry_grid.row = last_row + 1
          Else
              entry_grid.row = last_row
          End If
      Else
        If last_col = NAME_COL Then
            entry_grid.Col = AMOUNT_COL
            entry_grid.row = last_row
        Else
            If (preferences.auto_insert) Then insert_menu_Click
            If last_row + 1 < entry_grid.Rows Then
                entry_grid.row = last_row + 1
            Else
                entry_grid.row = last_row
            End If
            
            ' If it was in the tag column then keep it there
            If (last_col <> TAG_COL) Then
              entry_grid.Col = CHECK_COL
            Else
              entry_grid.Col = TAG_COL
            End If
        End If
      End If
    End If
    
    entry_grid.SetFocus
    ' See if the selected cell is off the screen
    With entry_grid
      If ((.row - .TopRow) >= number_of_displayed_rows) Then
          .TopRow = .TopRow + 1
      End If
    End With
    
    'If (entry_grid.Row > 10) Then
    '  entry_grid.TopRow = entry_grid.Row - 7
    'Else
    '  entry_grid.TopRow = 1
    'End If
    
  Exit Sub

errorh:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
End Sub


Sub entry_grid_LeaveCell()
    If (entry_grid.row > 0) Then entry_grid.CellBackColor = vbWhite
    
    If txtEdit.Visible = False Then Exit Sub
    entry_grid = txtEdit
    txtEdit.Visible = False
    editing_a_record = False
End Sub


Private Sub shuffle_up_entry_grid(n As Integer)
  'entry_grid.Rows = entry_grid.Rows + 1
  entry_grid.RemoveItem (n)
  entry_grid.row = entry_grid.row - 1
End Sub


Private Sub shuffle_down_entry_grid(n As Integer)
  With entry_grid
    .AddItem "", n
    .TextMatrix(n, DATE_COL) = .TextMatrix(n - 1, DATE_COL)
    .TextMatrix(n, DAY_COL) = .TextMatrix(n - 1, DAY_COL)
  End With
End Sub


Private Sub color_name_cell(ByVal r As Integer, ByVal c As Integer, ByVal amt As Double)
      entry_grid.row = r
      entry_grid.Col = NAME_COL
      
      If (amt < 0) Then
        ' Change the color of the cell to red
        entry_grid.CellForeColor = vbRed
      Else
        If (amt > 0) Then
          ' Change the color of the cell to blue
          entry_grid.CellForeColor = vbBlue
        Else
          entry_grid.CellForeColor = vbBlack
        End If
      End If
    entry_grid.Col = c
End Sub

Private Sub put_in_amount_cell(ByVal r As Integer, ByVal c As Integer, ByVal amt As Double)
      Dim color
      
      entry_grid.row = r
      entry_grid.Col = c
      entry_grid.CellFontBold = True
      
      entry_grid.TextMatrix(r, c) = currency_s(amt)
      
      If (amt < 0) Then
        ' Change the color of the cell to red
        color = vbRed
      Else
        If (amt > 0) Then
          ' Change the color of the cell to blue
          color = vbBlue
        Else
          color = vbBlack
        End If
      End If
      
      entry_grid.CellForeColor = color
      
      If (c = AMOUNT_COL) And (preferences.show_name_colors) Then
        entry_grid.Col = NAME_COL
        entry_grid.CellForeColor = color
        entry_grid.Col = c
      End If
End Sub

Private Sub shuffle_down_table_image(index As Integer)
  For j = MAX_RECORDS_IN_MONTH To index Step -1
    table_image.table(j) = table_image.table(j - 1)
  Next j
End Sub

Private Sub calculate()
  Dim balance As Double
  Dim check_balance As Double
  Dim i
'  Dim nnn
  
  ' Clear the quick accounts
  quick_form.clear_accounts
  
  ' Clear the cardtrak summary
  initialize_cardtrak_summary
  
  balance = 0
  check_balance = 0
  For i = 0 To 11
    balance_summary(i).beginning = balance
    balance_summary(i).ending = balance
    balance_summary(i).low = 9999999.99
    balance_summary(i).begin_found = False
    balance_summary(i).end_found = False
  Next i
  view.quick_date_start = view.start_year * 12 + view.start_month
  
  ' ----- Clear the tags -----
  For i = 0 To MAX_TAG
    tags(i).number = 0
    tags(i).total = 0
    tags(i).done = 0
    tags(i).pending = 0
    tags(i).blank = 0
    tags(i).skip = 0
  Next i
    
  ' ----- Clear the summary information -----
  For i = 0 To MAX_PAID
    summary.income(i) = 0
    summary.expense(i) = 0
    summary.number_income(i) = 0
    summary.number_expense(i) = 0
  Next i
  
  ' Start from the beginning and calculate the balances
  If (data.number_of_records > 0) Then
    ' We have at least one record
    data.current = data.first
    get_record  ' Get the first record
    
    ' Check for a quick account
    Call quick_form.check_quick_account
    
    ' See if this record should be included as normal
    If (Not db(data.current).exclude) And (Not db(data.current).override) Then
      balance = db(data.current).amount
      If (db(data.current).paid = 1) Then check_balance = db(data.current).amount
    End If
    
    ' See if we have exclude
    If (db(data.current).exclude) Then
      ' We have exclude
      balance = 0
      check_balance = 0
    Else
      ' See if we have override
      If (db(data.current).override) And (Not db(data.current).exclude) Then
        ' We have override
        balance = db(data.current).override_amount
        If (db(data.current).paid = 1) Then check_balance = db(data.current).override_amount
      End If
    End If
    
      
    db(data.current).balance = balance
    ' Put this into the balance summary
    this.balance = balance
    put_this_in_balance_summary
    put_this_in_cardtrak_summary (cardtrak_filter)
    
'nnn = 0
    While get_next_record
      ' Loop though all the remaining records and do the balance
      
      ' Check for a quick account
'nnn = nnn + 1
'If (this.this = 200) Then
'this.this = 200
'End If

Call quick_form.check_quick_account
      
      If (this.this = -1) And (this.next = 0) And (this.previous = 0) Then
      this.this = this.this
      End If
      'If (this.this = 79) Then
      'this.this = this.this
      'End If
      If (this.day = 0) Then
      this.day = 0
      End If
    
      
      
      If (Not db(data.current).exclude) And (Not db(data.current).override) Then
        balance = balance + db(data.current).amount
        If (db(data.current).paid = 1) Then check_balance = check_balance + db(data.current).amount
      End If
    
      ' See if we have override and no exclude
      If (db(data.current).override) And (Not db(data.current).exclude) Then
        ' We have override
        balance = db(data.current).override_amount
        If (db(data.current).paid = 1) Then check_balance = db(data.current).override_amount
      End If
    
      
      db(data.current).balance = balance
      ' Put this into the balance summary
      this.balance = balance
      put_this_in_balance_summary
      put_this_in_cardtrak_summary (cardtrak_filter)
    
    Wend
  End If
  
  checkbook_balance_box.Text = currency_s(check_balance)
  ' Show the checkbook balance colors
  If (Val(checkbook_balance_box.Text) > 0) Then
    checkbook_balance_box.ForeColor = vbBlue
  Else
    If (Val(checkbook_balance_box.Text) < 0) Then
       checkbook_balance_box.ForeColor = vbRed
    Else
      checkbook_balance_box.ForeColor = vbBlack
    End If
  End If
End Sub

Private Sub stuff_this_in_table_image()
  Dim c, r, amt
  
  ' See if it matches the transaction filter, false means to kill it
  If (filter_check = False) Then
    Exit Sub
  End If
  
  ' ----- Sum up the tag and summary amounts if not excluded -----
  If (this.exclude = False) Then
    If (this.amount > 0) Then
      summary.income(this.paid) = summary.income(this.paid) + this.amount
      summary.number_income(this.paid) = summary.number_income(this.paid) + 1
    Else
      If (this.amount < 0) Then
        summary.expense(this.paid) = summary.expense(this.paid) + this.amount
        summary.number_expense(this.paid) = summary.number_expense(this.paid) + 1
      End If
    End If
    
    If (this.tags <> 0) Then  ' See if we have any tags set in this transaction
      For i = 0 To MAX_TAG  ' Yes we do
        If ((this.tags And tag_mask(i)) <> 0) Then
          tags(i).total = tags(i).total + this.amount
          If (this.paid = PAID_BLANK) Then tags(i).blank = tags(i).blank + this.amount
          If (this.paid = PAID_DONE) Then tags(i).done = tags(i).done + this.amount
          If (this.paid = PAID_QUESTION) Then tags(i).pending = tags(i).pending + this.amount
          If (this.paid = PAID_DASH) Then tags(i).skip = tags(i).skip + this.amount
          tags(i).number = tags(i).number + 1
        End If
      Next i
    End If
  End If
    
  
  ' Put this in the entry grid
  ' Scan down the entry grid till we find the place where it goes
  For i = table_image.last To 1 Step -1
    If (table_image.table(i).day = this.day) Or _
       (this.day > table_image.table(table_image.last).day) Then
      ' We found a date that matches
      ' See if there is already an entry there
      If (table_image.table(i).this > -2) Then
        ' We have an entry so shuffle down
        i = i + 1
        shuffle_down_table_image (i)
        table_image.last = table_image.last + 1
      End If
      
      ' Add this to the total if not excluded
      If (this.exclude = False) Then
        filter.total_amount = filter.total_amount + this.amount
      End If
      
      'table_image.last = table_image.last + 1
      table_image.table(i).day = this.day
      table_image.table(i).this = this.this
      table_image.table(i).prev = this.previous
      table_image.table(i).next = this.next
      view.records_in_month = view.records_in_month + 1
      Exit For
    End If
  Next i

End Sub

Private Sub stuff_the_entry_grid()
  Dim j
  
  entry_grid.Redraw = False

  With entry_grid
  
    .Rows = MAX_RECORDS_IN_MONTH  ' Start with the max number
  
    ' Clear out the entry
    calendar.day = 1
    calendar.Year = view.current_year
    calendar.Month = view.current_month
    .Clear
    For i = 0 To MAX_COL
      .row = 0
      .Col = i
      .CellFontBold = False
      .CellFontSize = 8
      .CellAlignment = flexAlignCenterCenter
    Next i
    
    .TextMatrix(0, DATE_COL) = words(DATE_N)  '"Date"
    .TextMatrix(0, DAY_COL) = words(DAY_N)  '"Day"
    .TextMatrix(0, THIS_COL) = "" ' "This"
    .TextMatrix(0, PREV_COL) = "" ' "Prev"
    .TextMatrix(0, NEXT_COL) = "" ' "Next"
    .TextMatrix(0, DUE_COL) = words(DUE_N)  '"Due"
    .TextMatrix(0, CHECK_COL) = words(CHECK_N)  '"Check"
    .TextMatrix(0, NAME_COL) = words(NAME_N)  '"Name"
    .TextMatrix(0, PAID_COL) = words(STATUS_N)  '"Status"
    .TextMatrix(0, AMOUNT_COL) = words(AMOUNT_N)  '"Amount"
    .TextMatrix(0, EXCLUDE_COL) = words(EXCLUDE_N)  '"Excl"
    .TextMatrix(0, TAG_COL) = words(TAGS_N)  '"Tags"
    .TextMatrix(0, CLEARED_COL) = words(CLR_N)  '"CLR"
    .TextMatrix(0, BALANCE_COL) = words(BALANCE_N)  '"Balance"
    .TextMatrix(0, OVERRIDE_COL) = words(OR_N)  '"O/R"
    .TextMatrix(0, OVERRIDE_AMOUNT_COL) = words(OR_BALANCE_N)  '"O/R Balance"

    For i = 1 To MAX_RECORDS_IN_MONTH
      If (table_image.table(i).day = 0) Then Exit For
      .row = i
      .Col = DATE_COL
      .CellFontBold = True
      .CellAlignment = flexAlignCenterCenter  ' Center justify the name
      .CellFontSize = 10
      
      .Col = DAY_COL
      .CellFontSize = 9
      .CellAlignment = flexAlignCenterCenter  ' Center justify the name
      
      .Col = DUE_COL
      .CellFontSize = 9
      .CellAlignment = flexAlignCenterCenter  ' Right justify the name
      
      .Col = CHECK_COL
      .CellFontSize = 9
      .CellAlignment = flexAlignCenterCenter  ' Right justify the name
      
      .Col = NAME_COL
      .CellFontSize = 9
      .CellAlignment = flexAlignLeftCenter  ' Left justify the name
      .CellFontBold = False
      
      .TextMatrix(i, DATE_COL) = table_image.table(i).day
    
      calendar.day = table_image.table(i).day
      .TextMatrix(i, DAY_COL) = DAY_STRINGS(calendar.DayOfWeek)
      
      data.current = table_image.table(i).this
      If (data.current >= 0) Then
        ' We have a valid record number
        get_record
        .TextMatrix(i, THIS_COL) = this.this
        .TextMatrix(i, PREV_COL) = this.previous
        .TextMatrix(i, NEXT_COL) = this.next
        If (this.check >= 0) Then .TextMatrix(i, CHECK_COL) = Format(this.check)
        .TextMatrix(i, NAME_COL) = this.name
      
        If (this.due <> 0) Then
          .TextMatrix(i, DUE_COL) = this.due
        Else
          .TextMatrix(i, DUE_COL) = ""
        End If
        
        Call put_in_amount_cell(i, AMOUNT_COL, this.amount)
        Call put_in_amount_cell(i, BALANCE_COL, this.balance)
        Call put_in_amount_cell(i, OVERRIDE_AMOUNT_COL, this.override_amount)
    
        .Col = CLEARED_COL
        If (this.cleared = 0) Then .TextMatrix(i, CLEARED_COL) = ""
        If (this.cleared = 1) Then
          .TextMatrix(i, CLEARED_COL) = "*"
          .CellFontSize = 14
        End If
        If (this.cleared = 2) Then
          .TextMatrix(i, CLEARED_COL) = "X"
          .CellFontSize = 10
        End If
        .CellFontBold = True
        .CellAlignment = flexAlignCenterCenter  ' Center justify the name
      
        
        .row = i
        
        ' Put in the exclude graphic
        If (this.exclude) Then
          .Col = EXCLUDE_COL
          .CellAlignment = flexAlignCenterCenter
          .CellFontSize = 16
          .Text = Chr(149)
        End If
      
        ' Put in the override graphic
        If (this.override) Then
          .Col = OVERRIDE_COL
          .CellAlignment = flexAlignCenterCenter
          .CellFontSize = 16
          .Text = Chr(149)
        End If
      
        ' Put in the paid graphic
        .Col = PAID_COL
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 14
        '.Text = Chr(149)
        'If (this.paid = -1) Then this.paid = 1
        If (this.paid = 0) Then .Text = ""        ' Blank
        If (this.paid = 1) Then .Text = Chr(149)  ' Dot
        If (this.paid = 2) Then .Text = "?"       ' ?
        If (this.paid = 3) Then .Text = "--" 'Chr(150)  ' Dash
        
        ' Put in the tags
        .Col = TAG_COL
        .CellAlignment = flexAlignCenterCenter
        .CellFontSize = 8
        .CellFontBold = True
        .Text = ""
        For j = 0 To MAX_TAG
          If ((this.tags And tag_mask(j)) <> 0) Then .Text = .Text + Str(j + 1)
        Next j
      End If
            
    Next i
    .Rows = i
    .Redraw = True
    
  End With
End Sub

Public Sub update_misc_notes_characters_left()
  ' Show the user how many characters are left out of the maximum length of 65000
  characters_left_label.Caption = misc_notes_box.MaxLength - Len(misc_notes_box.Text)
End Sub

Public Sub process()
  Dim end_balance, start_balance, last_balance
  Dim r, c
  
  update_misc_notes_characters_left
  
  doing_process = True
  
  'Display todays date on the main form
  update_todays_date
  
  message_label.Caption = ""
  filter_form.get_filter_parameters  ' Get any current filter parameters
  
  first_label.Caption = data.first
  last_label.Caption = data.last
  total_label.Caption = data.number_of_records
  
  main_form.MousePointer = 11
  
  table_image.last = -1
  'entry_grid.Redraw = False
  
  ' Find the year of the tab we are on
  view.current_month = view.start_month + entry_tab.Tab
  view.current_year = view.start_year
  If (view.current_month > 12) Then
    ' We rolled over into the next year
    view.current_month = view.current_month - 12
    view.current_year = view.current_year + 1
  End If
  
  
  ' Process and update
  calculate
  
  
  On Error GoTo done
  calendar.day = 1
  calendar.Year = view.current_year
  calendar.Month = view.current_month
  For i = 1 To MAX_RECORDS_IN_MONTH
    table_image.table(i).day = i
    table_image.table(i + 1).day = 0
    table_image.table(i).this = -2
    
    calendar.day = calendar.day + 1
  Next i
  
done:
  table_image.last = i
  view.number_of_days = i
  
  view.records_in_month = 0
  view.last_balance = 0
  last_balance = 0
  beginning_balance_box.Text = currency_s(0)
  ending_balance_box.Text = currency_s(0)
  
  ' Find the first record that matches it
  If (find_first(view.current_month, view.current_year)) Then
    ' We have at least one record
    ' Put record in display
    stuff_this_in_table_image
    
    ' This is the first record in the month so see if it was excluded
    If (this.exclude) Then
      beginning_balance_box.Text = currency_s(this.balance)  ' - this.amount)
    Else
      beginning_balance_box.Text = currency_s(this.balance - this.amount)
    End If
    
    end_balance = this.balance
    ending_balance_box.Text = currency_s(view.last_balance)  ' end_balance
    While (find_next())
      stuff_this_in_table_image
      
      end_balance = this.balance
      ending_balance_box.Text = currency_s(end_balance)
    Wend
  Else
    ' No records found in month
    beginning_balance_box.Text = currency_s(view.last_balance)
    ending_balance_box.Text = currency_s(view.last_balance)  ' end_balance
  End If
  
  'entry_grid.Redraw = True
  stuff_the_entry_grid
  
  ' Now display the notes
  notes_box.Text = ""
  For i = 0 To MAX_NOTES
    If (notes(i).Month = view.current_month) And _
       (notes(i).Year = view.current_year) Then
      ' We have found the note that belongs to this month
      notes_box.Text = notes(i).s
      Exit For
    End If
  Next i
  
  main_form.MousePointer = 1
  
  show_buttons
  
  ' Show the beginning balance colors
  If (Val(beginning_balance_box.Text) > 0) Then
    beginning_balance_box.ForeColor = vbBlue
  Else
    If (Val(beginning_balance_box.Text) < 0) Then
      beginning_balance_box.ForeColor = vbRed
    Else
      beginning_balance_box.ForeColor = vbBlack
    End If
  End If
  
  ' Show the ending balance colors
  If (Val(ending_balance_box.Text) > 0) Then
    ending_balance_box.ForeColor = vbBlue
  Else
    If (Val(ending_balance_box.Text) < 0) Then
      ending_balance_box.ForeColor = vbRed
    Else
      ending_balance_box.ForeColor = vbBlack
    End If
  End If
  
  
  ' ----- Update the other status forms -----
  balance_form.update_balance_display  ' Update the balance form
  tags_form.update_tags_display  ' Update the tags form
  summary_form.update_summary_display  ' Update the summary form
  filter_results_form.update_filter_results_display
  If (ct_summary_form.Visible) Then ct_summary_form.update_ct_summary_display
  
  
  ' ----- Update the summary which isn't currently used -----
  update_summary
  
  ' Show if we have any filtered transactions
  If (filter.filtered) Then
    filter_results_form.show
    filter_results_form.total_box.Text = "$" + currency_s(filter.total_amount)
    filter_results_form.filtered_in_box.Text = Format(filter.filtered_in_count)
    filter_results_form.filtered_out_box.Text = Format(filter.filtered_out_count)
    main_form.SetFocus
  Else
    filter_results_form.Hide
  End If
  
  ' Update the tag caption in the entry grid fixed row
  r = entry_grid.row
  c = entry_grid.Col
  entry_grid.TextMatrix(0, TAG_COL) = "Tags-" + Format(last_tag + 1)
  entry_grid.row = r
  entry_grid.Col = c
  
  If (entry_grid.Visible) Then entry_grid.SetFocus
  If (quick_form.Visible) Then quick_form.SetFocus
  If (ct_summary_form.Visible) Then ct_summary_form.make_it_visible
  
  doing_process = False
End Sub

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
  If (items_are_selected) Then
    copy_selected_menu.Enabled = True
    cut_selected_menu.Enabled = True
  Else
    copy_selected_menu.Enabled = False
    cut_selected_menu.Enabled = False
  End If
  
  If (monthly_notes_menu.Checked = True) Then
      ' Doing notes
      cut_menu.Enabled = True
      copy_menu.Enabled = True
      cut_button.Visible = True
      copy_button.Visible = True
      
      insert_menu.Enabled = False
      delete_menu.Enabled = False
      cut_month_menu.Enabled = False
      copy_month_menu.Enabled = False
      paste_month_menu.Enabled = False
      paste_month_option_root_menu.Enabled = False
      
      insert_button.Visible = False
      delete_button.Visible = False
  Else  'Doing transactions
    insert_button.Visible = True
    If (view.records_in_month = 0) Then
      insert_menu.Enabled = False
      delete_menu.Enabled = False
      cut_menu.Enabled = False
      copy_menu.Enabled = False
      cut_month_menu.Enabled = False
      copy_month_menu.Enabled = False
      
      paste_month_menu.Enabled = True
      paste_month_option_root_menu.Enabled = True
      insert_button.Visible = False
      delete_button.Visible = False
      cut_button.Visible = False
      copy_button.Visible = False
    Else
      insert_menu.Enabled = True
      delete_menu.Enabled = True
      cut_menu.Enabled = True
      copy_menu.Enabled = True
      cut_month_menu.Enabled = True
      copy_month_menu.Enabled = True
      paste_month_menu.Enabled = True
      paste_month_option_root_menu.Enabled = True
    
      insert_button.Visible = True
      delete_button.Visible = True
      cut_button.Visible = True
      copy_button.Visible = True
    End If
  End If
End Sub

Sub Get_Recent_Files()
  Dim KEY, i
  Dim s As String
  
  ' Get recent file strings
  For i = 0 To 3
    KEY = "RecentFile" & i + 1
    s = GetSetting("Check 2 Check", "Settings", KEY, "Not Used")
    If s <> "Not Used" Then
      ' Update the recent files
      main_form.recent_file_menu(i).Visible = True
      main_form.recent_file_menu(i).Caption = s + ".c2c"
    Else
      main_form.recent_file_menu(i).Visible = False
    End If
  Next i

End Sub

Sub Save_Recent_Files()
  Dim s(4) As String
  Dim s_in As String
  
  s_in = UCase(data.db_name)
  
  s(0) = UCase(GetSetting("Check 2 Check", "Settings", "RecentFile1", "Not Used"))
  s(1) = UCase(GetSetting("Check 2 Check", "Settings", "RecentFile2", "Not Used"))
  s(2) = UCase(GetSetting("Check 2 Check", "Settings", "RecentFile3", "Not Used"))
  s(3) = UCase(GetSetting("Check 2 Check", "Settings", "RecentFile4", "Not Used"))

  ' See if the current one is in the list
  If s(3) = s_in Then
    ' Shuffle down
    s(3) = s(2)
    s(2) = s(1)
    s(1) = s(0)
    s(0) = s_in
  Else
    If s(2) = s_in Then
      s(2) = s(1)
      s(1) = s(0)
      s(0) = s_in
    Else
      If s(1) = s_in Then
        s(1) = s(0)
        s(0) = s_in
      Else
        If s(0) <> s_in Then
          s(3) = s(2)
          s(2) = s(1)
          s(1) = s(0)
          s(0) = s_in
        End If
      End If
    End If
  End If
  
  SaveSetting "Check 2 Check", "Settings", "RecentFile1", s(0)
  SaveSetting "Check 2 Check", "Settings", "RecentFile2", s(1)
  SaveSetting "Check 2 Check", "Settings", "RecentFile3", s(2)
  SaveSetting "Check 2 Check", "Settings", "RecentFile4", s(3)

  Get_Recent_Files
End Sub

Private Function pad(s As String, i As Integer) As String
  ' Make the returned string i characters long
  pad = s
  While (Len(pad) < i)
    pad = " " + pad
  Wend
End Function

Private Sub print_header_notes()
  'Call print_form.print_box(0, 0, 700, 1000, 2)
  
  ' Print out the data
  print_form.FontName = "Arial"
  Call print_form.print_next(5, 0, 0, data.db_name, 10)
  Call print_form.print_next(5, 25, 0, "Check2Check", 14)
  Call print_form.print_next(225, 13, 0, LONG_MONTH_STRINGS(view.current_month) + " " + CStr(view.current_year), 24)
  Call print_form.print_next(600, 1, 0, date_s, 10)
  Call print_form.print_next(600, 16, 0, time_s, 10)
  Call print_form.print_next(600, 31, 0, "Page " + Format(page_number), 10)
  Call print_form.print_dash(0, 50, 700, 50, 2)
    
  line_count = 0
End Sub

Private Sub print_header_transactions()
  'Call print_form.print_box(0, 0, 700, 1000, 2)
  
  ' Print out the data
  Call print_form.print_next(2, 52, 0, "Date", 10)
  Call print_form.print_next(32, 52, 0, "Day", 10)
  Call print_form.print_next(60, 52, 0, "Due", 10)
  Call print_form.print_next(88, 52, 0, "Check", 10)
  Call print_form.print_next(205, 52, 0, "Name", 10)
  Call print_form.print_next(360, 52, 0, "Status", 10)
  Call print_form.print_next(410, 52, 0, "Amount", 10)
  Call print_form.print_next(465, 52, 0, "Excl", 10)
  Call print_form.print_next(495, 52, 0, "Tags", 10)
  Call print_form.print_next(528, 52, 0, "Clr", 10)
  Call print_form.print_next(555, 52, 0, "Balance", 10)
  Call print_form.print_next(610, 52, 0, "O/R", 10)
  Call print_form.print_next(640, 52, 0, "O/R Bal", 10)
  
  Call print_form.print_dash(0, 70, 700, 70, 2)
End Sub

Private Sub do_new_page()
  page_number = page_number + 1
  
  If (page_number = 1) Then print_form.start_document (print_destination)
  
  If (page_type = 0) Then  ' Doing transactions
    If (page_number <> 1) Then print_form.new_page
    Call print_header_notes
    Call print_header_transactions
    y = 60
    transaction_line_count = 1
  End If
  
  If (page_type = 1) Then  ' Doing notes
    print_form.new_page
    Call print_header_notes
    y = 60
  End If
End Sub

Private Function do_line_count() As Boolean
  line_count = line_count + 1
  transaction_line_count = transaction_line_count + 1
  
  do_line_count = False  ' Set to initially do more pages
  If (line_count > 60) Then
    print_form.exit_value = 1
    
    ' Display the lines if doing transactions
    If (page_type = 0) Then Call display_vertical_lines(y + 62 * 14)
    
    If (print_destination = SCR) Then print_form.show (vbModal)
    If (print_form.exit_value = 0) Then
      do_line_count = True  ' Don't continue with the pages
      print_form.end_document  ' All done with printing or screen
      Exit Function
    End If
      
    do_new_page
    line_count = 0
  End If
End Function

Private Sub display_vertical_lines(y As Integer)
  Dim yy As Integer
  
  ' Draw the vertical lines
  yy = y + j * 14 + 14
  print_form.FontName = "Arial"
  
  Call print_form.print_line(0, y, 700, y, 1)
  Call print_form.print_line(0, 50, 0, y, 1)
  Call print_form.print_line(25, 70, 25, y, 1)  ' Day left
  Call print_form.print_line(57, 70, 57, y, 1)  ' Due left
  Call print_form.print_line(85, 70, 85, y, 1)  ' Check left
  Call print_form.print_line(120, 70, 120, y, 1)  ' Name left
  Call print_form.print_line(370, 70, 370, y, 1)  ' Status left
  Call print_form.print_line(390, 70, 390, y, 1)  ' Amount left
  Call print_form.print_line(470, 70, 470, y, 1)  ' Exclude left
  Call print_form.print_line(490, 70, 490, y, 1)  ' Tags left
  Call print_form.print_line(520, 70, 520, y, 1)  ' Clr left
  Call print_form.print_line(540, 70, 540, y, 3)  ' Balance left
  Call print_form.print_line(610, 70, 610, y, 3)  ' O/R left (balance right)
  Call print_form.print_line(630, 70, 630, y, 1)  ' O/R amount left
  Call print_form.print_line(700, 50, 700, y, 1)
End Sub

Private Sub view_printout()
  Dim last_date_s
  Dim t, m, n, j
  Dim s As String
  
  print_form.last_page = False
  
  page_number = 0
  line_count = 0
  page_type = TRANSACTIONS_PAGE
  date_s = Format(date)
  time_s = Format(Time, "h:mm AMPM")
  
  Call do_new_page
  
  last_date_s = ""
  
  y = 60
  
  ' Print out the month beginning balance
  Call print_form.print_next(130, y + 14, 0, "Beginning Balance", 8)  ' Ending balance
  Call print_form.print_next_right(610, y + 14, beginning_balance_box.Text, 8)  ' Ending balance
  Call print_form.print_line(0, 88, 700, 88, 1)
  y = y + 14
  
  For i = 1 To entry_grid.Rows - 1
    ' Print out each line
    If do_line_count Then Exit Sub
    
    j = transaction_line_count
    
    If (last_date_s <> entry_grid.TextMatrix(i, DATE_COL)) Then
      last_date_s = entry_grid.TextMatrix(i, DATE_COL)
      Call print_form.print_next(5, y + j * 14, 0, entry_grid.TextMatrix(i, DATE_COL), 8) ' Date
      Call print_form.print_next(32, y + j * 14, 0, entry_grid.TextMatrix(i, DAY_COL), 8)  ' Day
    End If
    
    Call print_form.print_next_right(78, y + j * 14, entry_grid.TextMatrix(i, DUE_COL), 8)  ' Due
    Call print_form.print_next_right(117, y + j * 14, entry_grid.TextMatrix(i, CHECK_COL), 8)  ' Check
    Call print_form.print_next(130, y + j * 14, 240, entry_grid.TextMatrix(i, NAME_COL), 8)  ' Name
    Call print_form.print_next_right(383, y + j * 14, entry_grid.TextMatrix(i, PAID_COL), 10)  ' Done
    Call print_form.print_next_right(470, y + j * 14, entry_grid.TextMatrix(i, AMOUNT_COL), 8)  ' Amount
    Call print_form.print_next_right(483, y + j * 14, entry_grid.TextMatrix(i, EXCLUDE_COL), 10)  ' Excl
    Call print_form.print_next(495, y + j * 14, 0, strip_spaces(entry_grid.TextMatrix(i, TAG_COL)), 8)  ' Tags
    Call print_form.print_next_right(534, y + j * 14, entry_grid.TextMatrix(i, CLEARED_COL), 10)  ' Clr
    Call print_form.print_next_right(610, y + j * 14, entry_grid.TextMatrix(i, BALANCE_COL), 8)  ' Balance
    Call print_form.print_next_right(623, y + j * 14, entry_grid.TextMatrix(i, OVERRIDE_COL), 10)  ' O/R
    Call print_form.print_next_right(700, y + j * 14, entry_grid.TextMatrix(i, OVERRIDE_AMOUNT_COL), 8)  ' O/R Balance
    
  Next i
  
  j = j + 1
  
  ' Print out the month ending balance
  Call print_form.print_line(0, y + j * 14, 700, y + j * 14, 1)
  Call print_form.print_next(130, y + j * 14, 0, "Ending Balance", 8)  ' Ending balance
  Call print_form.print_next_right(610, y + j * 14, ending_balance_box.Text, 8)  ' Ending balance
  
  ' Display the lines
  y = y + j * 14 + 14
  print_form.FontName = "Arial"
  Call print_form.print_line(0, y, 700, y, 1)
  Call display_vertical_lines(y)

  '-------------------------------------
  ' Print out the notes
  '-------------------------------------
  y = y + 10
  page_type = NOTES_PAGE
  If do_line_count Then Exit Sub
  Call print_form.print_next(300, y, 0, "Notes", 16)  ' O/R Balance
  
  y = y + 25
  t = notes_box.Text
  m = 1
  n = -1
  For i = 0 To 200
    ' Get the first line of text
    
    m = n + 2
    n = InStr(n + 2, t, Chr(13))
    If (m >= Len(t)) Then Exit For
    s = Mid(t, m, n - m)
    
    If do_line_count Then Exit Sub
    Call print_form.print_next(5, y, 0, s, 8) ' Notes
    y = y + 14
  Next i
  
  print_form.last_page = True
  If (print_destination = SCR) Then print_form.show (vbModal)
  print_form.end_document  ' All done with printing or screen
End Sub


Private Sub undo_button_Click()
  undo_menu_Click
End Sub

Private Sub undo_menu_Click()
  Dim index As Integer
  
  undo.doing_undo = True
  
  ' --------- Undo paste record ---------
  If (undo.what_was_done = WHAT_PASTE_RECORD) Then
    ' Delete the record at undo.rec_num
    delete_record (undo.rec_num)
    process
  End If
  
  '--------- Undo cut / delete record ----------
  If (undo.what_was_done = WHAT_CUT_RECORD) Or _
     (undo.what_was_done = WHAT_DELETE_RECORD) Then
    ' Restore the main and ct records
    this = undo.r
    ' ---- Restore the saved cardtrack ----
    index = insert_cardtrak_record(this, undo.cardtrak)  ' Returned index points to the ct record in the ct db
    this.sub_transaction_number = index
    ' ---- Restore the saved record ---- This must be done after ct is restored
    insert_record (-1)
    process
  End If
  
  ' -------- Undo cut month --------
  If (undo.what_was_done = WHAT_CUT_MONTH) Then
    ' Restore the saved month
    copy_of_month = undo.copy_of_month
    paste_month_menu_click
    process
  End If
  
  ' -------- Undo move record ---------
  If (undo.what_was_done = WHAT_MOVE_RECORD) Then
    ' Put it back where it was
    delete_record (undo.rec_num)  ' Delete the new one
    this = undo.r
    If (this.sub_transaction_number > 0) Then cards(this.sub_transaction_number).active = True
    insert_record (-1)  ' Put the old one back
    process
  End If
  
  ' --------- Undo edit transaction ---------
  If (undo.what_was_done = WHAT_EDIT_TRANSACTION) Then
    ' Put it back the way it was
    db(undo.r.this) = undo.r
    process
  End If
  
  ' --------- Undo paste month / paste tags ---------
  If (undo.what_was_done = WHAT_PASTE_MONTH) Or _
     (undo.what_was_done = WHAT_PASTE_TAGS) Then
    ' Delete the records
    For i = 0 To MAX_RECORDS_IN_MONTH
      If (undo.copy_of_month.table(i).this < 0) Then Exit For
      delete_record (undo.copy_of_month.table(i).this)
    Next i
    view.current_month = undo.copy_of_month.Month
    view.current_year = undo.copy_of_month.Year
    
    ' Only update the notes if it was a paste month
    If (undo.what_was_done = WHAT_PASTE_MONTH) Then
      notes_box.Text = undo.copy_of_month.notes
      update_notes
    End If
    process
  End If
  
  ' ------- Undo move selected --------
  If (undo.what_was_done = WHAT_MOVE_SELECTED) Then
    ' First delete the records that were pasted
    For i = 0 To MAX_RECORDS_IN_MONTH
      If (undo.selected_rec_num(i) < 0) Then Exit For
      delete_record (undo.selected_rec_num(i))
    Next i
    
    ' Next add the records that were originally deleted
    copy_of_month = undo.copy_of_month
    copy_of_cardtrak_month = undo_cardtrak_month
    allow_paste_ct = True
    paste_month_menu_click
    
    process
  End If
  
  undo.what_was_done = WHAT_NONE
  undo_menu.Enabled = False
  undo_button.Visible = undo_menu.Enabled
  undo.doing_undo = False
End Sub

Private Sub update_menu_Click()
  internet_form.show 1
End Sub

Private Sub view_balance_button_Click()
  view_balances_menu_Click
End Sub

Private Sub view_balances_menu_Click()
  process
  balance_form.show
End Sub

Private Sub view_calculator_menu_Click()
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

Private Sub view_print_menu_Click()
  printer_error = False
  print_destination = SCR
  Call view_printout
End Sub


Public Sub EvaluationExpired()
  ' Evaluation period has expired so don't allow for opening files
  open_menu.Enabled = False
  open_button.Enabled = False
  recent_file_menu(0).Enabled = False
  recent_file_menu(1).Enabled = False
  recent_file_menu(2).Enabled = False
  recent_file_menu(3).Enabled = False
  
  ' Show the register menu
  register_menu.Visible = True
  register_menu_dash.Visible = True
  
  ' Clear out any open database
  new_menu_Click
End Sub

Private Sub register_menu_Click()
  ' Display the registration form
  If (Not register_form.ok_to_run_form) Then EvaluationExpired
    
  If (register_form.is_registered = True) Then
    ' We have a registered version
    open_menu.Enabled = True
    open_button.Enabled = True
    recent_file_menu(0).Enabled = True
    recent_file_menu(1).Enabled = True
    recent_file_menu(2).Enabled = True
    recent_file_menu(3).Enabled = True
  
    ' Hide the register menu
    'register_menu.Visible = False
    'register_menu_dash.Visible = False
  End If
End Sub


Private Sub view_quick_accounts_menu_Click()
  ' Show the quick accounts form
  process
  quick_form.execute (0)
End Sub

Private Sub view_quick_menu_Click()
  view_quick_accounts_menu_Click
End Sub

Private Sub view_summary_button_Click()
  view_summary_menu_Click
End Sub

Private Sub view_summary_menu_Click()
  ' Show the summary form
  process
  summary_form.show
End Sub

Private Sub view_tags_menu_Click()
  ' Show the tags form
  process
  tags_form.show
End Sub

Private Sub reconcile_menu_Click()
  ' Show all the unreconciled transactions
  Dim i
  
  reconcile_form.initialize
  
  ' Start from the beginning
  If (data.number_of_records > 0) Then
    ' We have at least one record
    data.current = data.first
    get_record  ' Get the first record
    
    If (db(data.current).cleared < 2) Then reconcile_form.add
    
    ' Loop though all the remaining records
    While get_next_record
        If (db(data.current).cleared < 2) Then reconcile_form.add
    Wend
  End If

  ' Display the reconcile form
  reconcile_form.execute
  process
End Sub

Private Sub paste_tag_arrange_option_menu_Click(index As Integer)
  paste_month_options = PM_ARRANGE
  pasting_tags = True
  Call paste_month_option(index)
  pasting_tags = False
End Sub


Private Sub paste_tag_option_menu_Click(index As Integer)
  pasting_tags = True
  paste_month_options = PM_NORMAL
  Call paste_month_option(index)
  pasting_tags = False
  paste_month_options = PM_ARRANGE
End Sub


Private Sub paste_tags_current_menu_Click(index As Integer)
  pasting_tags = True
  paste_month_options = PM_CURRENT
  Call paste_month_option(index)
  pasting_tags = False
  paste_month_options = PM_ARRANGE
End Sub


Private Sub web_site_menu_Click()
  ' Start up the browser and go to check2check.com
  Call Navigate(Me, "http://www.mycheck2check.com")
End Sub


Private Sub paste_month_menu_click()
  Dim days_adjusted, answer
  Dim index As Integer
  
  If (undo.doing_undo) Then
    ' Setup for doing the undo
    view.current_month = copy_of_month.Month
    view.current_year = copy_of_month.Year
  End If
  
  answer = vbYes
  If (Not undo.doing_undo) Then
    answer = (MsgBox(words(PASTE_ALL_TRANSACTIONS_TO_N) + " " + entry_tab.Caption, _
        vbYesNoCancel + vbQuestion + vbApplicationModal, "Paste Month"))
  End If
  
  If (answer = vbYes) Then
    ' Yes, Paste the entire month
    
    ' Paste the month buffer to the new month
    ' Be sure to check for the last day and overruns
    days_adjusted = False
    With copy_of_month
      For i = 0 To MAX_RECORDS_IN_MONTH
        If (.table(i).this <= -1) Then Exit For
        ' We have a record to transfer
        this = .table(i)
        
        If (Not undo.doing_undo) Then
        End If
        
        If (this.day > view.number_of_days) Then
          ' We must adjust the days because there are less days in the month
          this.day = view.number_of_days
          days_adjusted = True
        End If
        
        If (undo.doing_undo) Then
          this.Month = copy_of_month.Month
          this.Year = copy_of_month.Year
        Else
          this.Month = view.current_month
          this.Year = view.current_year
          this.paid = 0  ' Clear out the done column if not doing undo
          this.check = -1  ' Clear out the check column if not doing undo
        End If
        
        ' ---- Restore the saved cardtrack ----
        If (Not allow_paste_ct) Then
          this.sub_transaction_number = 0
        Else
          If (copy_of_cardtrak_month.table(i).active) Then
            index = insert_cardtrak_record(this, copy_of_cardtrak_month.table(i))  ' Returned index points to the ct record in the ct db
            this.sub_transaction_number = index
          Else
            this.sub_transaction_number = 0
          End If
        End If
        
        ' ---- Restore the saved record ---- This must be done after ct is restored
        insert_record (-1)
        
        ' Save the undo stuff
        If (Not undo.doing_undo) Then
          If (Not pasting_tags) Then
            undo.what_was_done = WHAT_PASTE_MONTH
          Else
            undo.what_was_done = WHAT_PASTE_TAGS
          End If
          
          undo.copy_of_month.table(i).this = data.current
          undo_cardtrak_month.table(i) = cards(this.sub_transaction_number)
          If (i < MAX_RECORDS_IN_MONTH) Then undo.copy_of_month.table(i + 1).this = -1
          undo_menu.Enabled = True
          undo_button.Visible = undo_menu.Enabled
        End If
      Next i
    End With
  
    ' Save the undo notes
    undo_menu.Caption = "Undo - Paste " + MONTH_STRINGS(view.current_month) + " " + Format(view.current_year)
    undo.copy_of_month.notes = notes_box.Text
    undo.copy_of_month.Month = view.current_month
    undo.copy_of_month.Year = view.current_year
    
    If (undo.what_was_done = WHAT_PASTE_MONTH) Then
      ' Only paste notes if it was past month, and not paste tags
      ' Paste the notes now
      If (Not undo.doing_undo) Then
        answer = vbNo
        If (preferences.prompt_for_paste_notes) Then
          answer = MsgBox(words(PASTE_NOTES_Q_N) + " ", vbYesNoCancel + vbQuestion + vbApplicationModal, words(PASTE_MONTH_N))
        End If
    
        If (answer = vbYes) Then
          notes_box.Text = notes_box.Text + copy_of_month.notes
          update_notes
        End If
      Else
        notes_box.Text = copy_of_month.notes
        update_notes
      End If
    End If
    
    process
    changed_flag = True
    update_caption
    allow_paste_ct = False
    If (days_adjusted) Then MsgBox words(ADJUSTED_DATES_TO_MATCH_THE_CURRENT_MONTH_N)
  
  End If
End Sub




