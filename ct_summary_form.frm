VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ct_summary_form 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Cardtrak"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   8505
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1980
      Top             =   4635
   End
   Begin MSComDlg.CommonDialog print_dialog 
      Left            =   1035
      Top             =   4725
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton print_button 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6795
      TabIndex        =   44
      Top             =   135
      Width           =   1140
   End
   Begin VB.CommandButton month_button 
      Height          =   330
      Index           =   1
      Left            =   5715
      Picture         =   "ct_summary_form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   135
      Width           =   1005
   End
   Begin VB.CommandButton month_button 
      Height          =   330
      Index           =   0
      Left            =   4590
      Picture         =   "ct_summary_form.frx":00AA
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   135
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Balance"
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
      Left            =   0
      TabIndex        =   4
      Top             =   540
      Width           =   9420
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   11
         Left            =   5760
         MouseIcon       =   "ct_summary_form.frx":0153
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   43
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   10
         Left            =   5355
         MouseIcon       =   "ct_summary_form.frx":045D
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   42
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   9
         Left            =   4950
         MouseIcon       =   "ct_summary_form.frx":0767
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   41
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   8
         Left            =   4545
         MouseIcon       =   "ct_summary_form.frx":0A71
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   40
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   7
         Left            =   4140
         MouseIcon       =   "ct_summary_form.frx":0D7B
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   39
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   6
         Left            =   3735
         MouseIcon       =   "ct_summary_form.frx":1085
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   38
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   5
         Left            =   3330
         MouseIcon       =   "ct_summary_form.frx":138F
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   37
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   4
         Left            =   2925
         MouseIcon       =   "ct_summary_form.frx":1699
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   36
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   3
         Left            =   2520
         MouseIcon       =   "ct_summary_form.frx":19A3
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   35
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   2
         Left            =   2115
         MouseIcon       =   "ct_summary_form.frx":1CAD
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   34
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   1
         Left            =   1710
         MouseIcon       =   "ct_summary_form.frx":1FB7
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   33
         Top             =   315
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         Height          =   1005
         Index           =   0
         Left            =   1305
         MouseIcon       =   "ct_summary_form.frx":22C1
         MousePointer    =   99  'Custom
         ScaleHeight     =   945
         ScaleWidth      =   270
         TabIndex        =   32
         Top             =   315
         Width           =   330
      End
      Begin VB.Frame Frame2 
         Caption         =   "Average"
         Height          =   510
         Left            =   6255
         TabIndex        =   28
         Top             =   1215
         Width           =   1500
         Begin VB.Label graph_average_label 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   29
            Top             =   225
            Width           =   1320
         End
      End
      Begin VB.ComboBox high_low_combo 
         Height          =   315
         Left            =   7830
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   765
         Width           =   1410
      End
      Begin VB.Frame change_panel 
         Caption         =   "Change"
         Height          =   510
         Left            =   7830
         TabIndex        =   25
         Top             =   1215
         Width           =   1500
         Begin VB.Label graph_change_label 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   26
            Top             =   225
            Width           =   1320
         End
      End
      Begin VB.ComboBox graph_combo 
         Height          =   315
         Left            =   7830
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   225
         Width           =   1410
      End
      Begin VB.Frame low_panel 
         Caption         =   "Low"
         Height          =   510
         Left            =   6255
         TabIndex        =   22
         Top             =   675
         Width           =   1500
         Begin VB.Label graph_low_label 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   23
            Top             =   225
            Width           =   1320
         End
      End
      Begin VB.Frame high_panel 
         Caption         =   "High"
         Height          =   510
         Left            =   6255
         TabIndex        =   20
         Top             =   135
         Width           =   1500
         Begin VB.Label graph_high_label 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   21
            Top             =   225
            Width           =   1320
         End
      End
      Begin VB.Line Line2 
         X1              =   90
         X2              =   1170
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label bar_click_label 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   45
         Top             =   1530
         Width           =   1050
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   1305
         MouseIcon       =   "ct_summary_form.frx":25CB
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   1710
         MouseIcon       =   "ct_summary_form.frx":28D5
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2115
         MouseIcon       =   "ct_summary_form.frx":2BDF
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2520
         MouseIcon       =   "ct_summary_form.frx":2EE9
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2925
         MouseIcon       =   "ct_summary_form.frx":31F3
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   3330
         MouseIcon       =   "ct_summary_form.frx":34FD
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   3735
         MouseIcon       =   "ct_summary_form.frx":3807
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   4140
         MouseIcon       =   "ct_summary_form.frx":3B11
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   4545
         MouseIcon       =   "ct_summary_form.frx":3E1B
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   4950
         MouseIcon       =   "ct_summary_form.frx":4125
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   5355
         MouseIcon       =   "ct_summary_form.frx":442F
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_month_label 
         Alignment       =   2  'Center
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   5760
         MouseIcon       =   "ct_summary_form.frx":4739
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label graph_max_label 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label graph_zero_label 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   6
         Top             =   1170
         Width           =   870
      End
      Begin VB.Label graph_mid_label 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   5
         Top             =   720
         Width           =   1050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   1260
         X2              =   6120
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Shape rectangle 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   1140
         Left            =   1260
         Top             =   270
         Width           =   4875
      End
   End
   Begin VB.CommandButton close_button 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9600
      TabIndex        =   3
      Top             =   135
      Width           =   1500
   End
   Begin VB.CommandButton process_button 
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5130
      TabIndex        =   2
      Top             =   4770
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox card_name_combo 
      Height          =   315
      ItemData        =   "ct_summary_form.frx":4A43
      Left            =   90
      List            =   "ct_summary_form.frx":4A45
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   4470
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   1755
      Left            =   45
      TabIndex        =   0
      Top             =   2430
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   3096
      _Version        =   393216
      Rows            =   100
      Cols            =   11
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      MergeCells      =   2
      AllowUserResizing=   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ct_summary_form.frx":4A47
   End
End
Attribute VB_Name = "ct_summary_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const N_COL = 0
Const NAME_COL = 1
Const BALANCE_COL = 2
Const PAID_COL = 3
Const PURCHASES_COL = 4
Const INTEREST_COL = 5
Const LATE_COL = 6
Const MIN_COL = 7
Const NET_COL = 8  ' Refers to (Paid - Purchases - Interest - Late)
Const SPACE_COL = 9  ' This column is not used
Const EST_BALANCE_COL = 10

Const TRANSACTIONS_PAGE = 0
Const NOTES_PAGE = 1
Const CARDTRAK_PAGE = 3

Dim last_selected_index As Integer  ' Index number of the name combo box that was selected last
Dim updating As Boolean  ' True when updating the display so don't respond to click events
Dim detailed_month_name As String  ' Name of active month
Dim x As Integer, y As Integer, y1 As Integer, y2 As Integer  ' Used for x and y for screen and page
Dim transaction_line_count As Integer  ' Count of the number of transaction on a printed or view page
Dim page_type As Integer        ' 0=Transactions, 1= notes
Dim date_s As String  ' Used for printouts
Dim time_s As String  ' Used for printouts
Dim print_destination  ' Either PTR or SRC
Dim pic As Picture  ' Used for capturing the frame to print
Dim form_has_focus  ' True when this form gets focus
Dim Bar_Value(12) As Double



Public Sub update_language()
  graph_combo.Clear
  graph_combo.AddItem (words(BALANCE_N))  '"Balance")
  graph_combo.AddItem (words(PAID_N))  '"Paid")
  graph_combo.AddItem (words(PURCHASES_N))  '"Purchases")
  graph_combo.AddItem (words(INTEREST_N))  '"Interest")
  graph_combo.AddItem (words(LATE_N))  '"Late")
  graph_combo.AddItem (words(MINIMUM_N))  '"Minimum")
  graph_combo.AddItem ("Net")
  graph_combo.AddItem (words(EST_BAL_N))  '"Est Bal")
  graph_combo.ListIndex = 0
  
  high_low_combo.Clear
  high_low_combo.AddItem (words(HIGH_LOW_N))  '"High/low")
  high_low_combo.AddItem (words(BEGIN_END_N))  '"Begin/End")
  high_low_combo.ListIndex = 0

  print_button.Caption = words(PRINT_N)
  close_button.Caption = words(CLOSE_N)
End Sub


Private Sub card_name_combo_Click()
  ' We just changed to a new card so process it
  If (Not updating) Then
    process_button_Click
  End If
End Sub

Private Sub close_button_Click()
  Hide
End Sub

Private Sub Form_Activate()
  form_has_focus = True
End Sub

Private Sub Form_Deactivate()
  form_has_focus = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyEscape) Or (KeyCode = vbKeyReturn) Then
    Form_Unload 0
  End If
End Sub

Private Sub Form_Load()
  updating = True
  
  ' Set up the form
  'With grid
  '  update_titles
  'End With
  
  last_selected_index = 0
  
  update_language
  
  high_low_combo.ListIndex = 0
  
  updating = False
End Sub

Private Sub Form_Resize()
  Dim i
  Static width
  Static height
  Static buttons
  
  With ct_summary_form
    If (width = 0) Then width = .width
    If (height = 0) Then height = .height
    If (buttons = 0) Then buttons = .width - month_button(0).Left
  
    If (ct_summary_form.width < width) Then .width = width
    If (ct_summary_form.height < 2800) Then .height = 2800
  
    close_button.Left = ct_summary_form.width - close_button.width - 300
    print_button.Left = close_button.Left - print_button.width - 50
    month_button(1).Left = print_button.Left - month_button(1).width - 50
    month_button(0).Left = month_button(1).Left - month_button(0).width - 50
    card_name_combo.width = month_button(0).Left - 150
    Frame1.Left = (.width - Frame1.width) / 2
  End With
  
  ' Change the width of the name col
  With grid
    .width = ct_summary_form.width - 250
    If (ct_summary_form.height > 3000) Then .height = ct_summary_form.height - .Top - 375
    For i = 0 To .Cols - 1
      .ColWidth(i) = 1000
    Next i
    i = .width - (.ColWidth(0) * 9.4) '8.5)
    If (i > 0) Then .ColWidth(1) = i
    .ColWidth(SPACE_COL) = 50
  End With
  
End Sub

Private Sub update_titles()
  'update_language
  With grid
    .TextMatrix(0, 0) = ""
    .TextMatrix(0, NAME_COL) = words(NAME_N)  '"Name"
    .TextMatrix(0, BALANCE_COL) = words(BALANCE_N)  '"Balance"
    .TextMatrix(0, PAID_COL) = words(PAID_N)  '"Paid"
    .TextMatrix(0, PURCHASES_COL) = words(PURCHASES_N)  '"Purchases"
    .TextMatrix(0, INTEREST_COL) = words(INTEREST_N)  '"Interest"
    .TextMatrix(0, LATE_COL) = words(LATE_N)  '"Late"
    .TextMatrix(0, MIN_COL) = words(MINIMUM_N)  '"Minimum"
    .TextMatrix(0, NET_COL) = "Net"  'words(NET_N)  'esk "Net"
    .TextMatrix(0, EST_BALANCE_COL) = words(EST_BAL_N)  '"Est Bal"
    
    .TextMatrix(1, 0) = words(TOTAL_N)  '"Total"
    
    .row = 0
    .Col = 0
    '.ColWidth(0) = 1000
    .CellAlignment = flexAlignCenterCenter
    .Col = 1
    '.ColWidth(1) = 1400
    .CellAlignment = flexAlignCenterCenter
    .Col = 2
    .RowHeight(2) = 50
    .CellAlignment = flexAlignCenterCenter
    .Col = 3
    '.ColWidth(1) = 1400
    .CellAlignment = flexAlignCenterCenter
    .Col = 4
    .CellAlignment = flexAlignCenterCenter
    .Col = 5
    .CellAlignment = flexAlignCenterCenter
    .Col = 6
    .CellAlignment = flexAlignCenterCenter
    .Col = 7
    .CellAlignment = flexAlignCenterCenter
    .Col = 8
    .CellAlignment = flexAlignCenterCenter
    .Col = 9
    .CellAlignment = flexAlignCenterCenter
    .Col = 10
    .CellAlignment = flexAlignCenterCenter
    
    .ColAlignment(0) = flexAlignCenterCenter
  End With
End Sub

Private Sub fill_combo()
  ' Fill up al the names of cards into the combo box
  Dim i As Integer
  
  last_selected_index = card_name_combo.ListIndex
  
  card_name_combo.Clear
  card_name_combo.AddItem (words(ALL_CARDS_N))  '"All Cards")
  For i = 1 To MAX_CARDS
    If (cards_info(i).active) Then
      card_name_combo.AddItem ("(CT" + Format(i, "00") + ") " + cards_info(i).name)
    End If
  Next i
  If (last_selected_index < 0) Then last_selected_index = 0
  If (card_name_combo.ListCount > last_selected_index) Then
    card_name_combo.ListIndex = last_selected_index
  Else
    card_name_combo.ListIndex = 0
    card_name_combo.Text = words(ALL_CARDS_N)  '"All Cards"
  End If
End Sub

Private Sub update_bar(i As Integer, percent As Double, item As Integer)
  Static Bottom As Integer
  Static height As Integer
  Dim h As Integer
  Dim incoming_percent As Double
  
  incoming_percent = percent
  
  If (percent < 0) Then percent = -percent
  
  With Picture1(i)
    .Visible = False
    If (Bottom = 0) Then
      ' First time through so get the bottom and height values
      Bottom = .Top + .height
      height = rectangle.height
      'rectangle.height = height
    End If
  
    h = height * percent * 0.95
    .height = h
    .Top = (Bottom - h) + 50
    
    ' Set the fill color, blue is higher than the previous one, green is less
    If (i = 0) Then .BackColor = vbBlue
    If (i > 0) Then
      If (.height >= Picture1(i - 1).height) Then
        .BackColor = vbBlue
      Else
        .BackColor = vbGreen
      End If
    End If
    If (item = 6) Then ' Do this only for the Net amount
      If (incoming_percent >= 0) Then
        .BackColor = vbBlue
      Else
        .BackColor = vbRed
      End If
    End If
    
    .Visible = True
    
  End With
End Sub

Private Sub update_graph()
  Dim max As Double
  Dim i As Integer
  Dim low As Double
  Dim begin As Double
  Dim ending As Double
  Dim v(12) As Double
  Dim sum As Double
  Dim sum_count As Integer
  
  low = 1000000
  max = -1000000  'esk0.1
  begin = -555555
  ending = -555555
  sum = 0
  sum_count = 0
  
  ' Find the maximum, minimum, begin, and ending values
  For i = 0 To 11
    If (graph_combo.ListIndex = 0) Then
      v(i) = cardtrak_monthly_summary(i).balance
    ElseIf (graph_combo.ListIndex = 1) Then v(i) = cardtrak_monthly_summary(i).paid
    ElseIf (graph_combo.ListIndex = 2) Then v(i) = cardtrak_monthly_summary(i).purchases
    ElseIf (graph_combo.ListIndex = 3) Then v(i) = cardtrak_monthly_summary(i).interest
    ElseIf (graph_combo.ListIndex = 4) Then v(i) = cardtrak_monthly_summary(i).late
    ElseIf (graph_combo.ListIndex = 5) Then v(i) = cardtrak_monthly_summary(i).minimum
    ElseIf (graph_combo.ListIndex = 6) Then v(i) = -cardtrak_monthly_summary(i).paid - cardtrak_monthly_summary(i).purchases - cardtrak_monthly_summary(i).interest - cardtrak_monthly_summary(i).late  ' esk
    ElseIf (graph_combo.ListIndex = 7) Then v(i) = cardtrak_monthly_summary(i).balance + cardtrak_monthly_summary(i).paid
    End If
  
    'esk If (v(i) < 0) Then v(i) = -v(i)
    
    If (Abs(v(i)) > max) Then max = Abs(v(i))
    If (Abs((v(i)) > 0.01) And (Abs(v(i)) < low)) Then low = Abs(v(i))
    
    ' Get the beginning and ending values, throw out zeros
    If (v(i) <> 0) Then
      sum = sum + v(i)
      sum_count = sum_count + 1
      
      If (begin = -555555) Then begin = v(i)
      ending = v(i)
    End If
  Next i
  
  ' ----------- Plot it ---------
  For i = 0 To 11
    If max = 0 Then max = 1000  ' Arbitrarily use a big number because if the card is unused then max will be zero
    Call update_bar(i, v(i) / max, graph_combo.ListIndex)
    Bar_Value(i) = v(i)
  Next i
  
  ' --------- Label the graph ---------
  Frame1.Caption = graph_combo.Text
  graph_max_label.Caption = currency_s(max)
  graph_mid_label.Caption = currency_s(max / 2)
  graph_zero_label.Caption = currency_s(0#)
  
  ' -------- Label the x axis -------
  For i = 0 To 11
    grid.row = 3 + i
    grid.Col = 0
    graph_month_label(i).Caption = Left(grid.Text, 3)
    graph_month_label(i).FontBold = grid.CellFontBold
    If (grid.CellFontBold) Then
      graph_month_label(i).ForeColor = vbBlue
      bar_click_label.Caption = currency_s(v(i))
    Else
      graph_month_label(i).ForeColor = vbBlack
    End If
  Next i
  
  ' --------- Average Value --------
  If (sum_count > 0) Then
    graph_average_label.Caption = currency_s(sum / sum_count)
  Else
    graph_average_label.Caption = currency_s(0#)
  End If
  
  ' ------- High/Low Values -------
  If (max < 0.11) Then max = 0
  If (low > 999999) Then low = 0
  If (high_low_combo.ListIndex = 0) Then
    ' Show the High low change limits
    high_panel.Caption = words(HIGH_N)  '"High"
    low_panel.Caption = words(LOW_N)  '"Low"
    change_panel.Caption = words(CHANGE_N)  '"Change"
    
    graph_high_label.Caption = currency_s(max)
    graph_low_label.Caption = currency_s(low)
  
    graph_change_label.Caption = currency_s(max - low)
  End If
  
  
  ' ---------- Begin/Ending Values ----------
  If (begin = -555555) Then begin = 0
  If (ending = -555555) Then ending = 0
  
  If (high_low_combo.ListIndex = 1) Then
    ' Show the Start End change limits
    high_panel.Caption = words(BEGINNING_N)  '"Beginning"
    low_panel.Caption = words(ENDING_N)  '"Ending"
    change_panel.Caption = words(CHANGE_N)  '"Change"
    
    graph_high_label.Caption = currency_s(begin)
    graph_low_label.Caption = currency_s(ending)
  
    graph_change_label.Caption = currency_s(ending - begin)
  End If
End Sub

Public Sub update_ct_summary_display()
  Dim i As Integer
  Dim m As Integer
  Dim sum As Double
  Dim s As String
  Dim sum_paid As Double
  Dim sum_purchases As Double
  Dim sum_interest As Double
  Dim sum_late As Double
  Dim sum_net As Double
  Dim sum_minimum As Double
  Dim net As Double
  
  'update_language
  
  updating = True
  
  fill_combo  ' Fill up all the card names in the combo box
  
  detailed_month_name = ""
  sum_paid = 0
  sum_purchases = 0
  sum_interest = 0
  sum_late = 0
  sum_minimum = 0
  
  With grid
    .Redraw = False
    .Clear
    .Rows = 20  ' Start with 20
    update_titles
    
    ' --------- Display the monthly summaries --------
    For i = 0 To 11
      .row = i + 3
      .Col = 0
      
      .Text = main_form.entry_tab.TabCaption(i)
      If (i = main_form.entry_tab.Tab) Then
        .CellFontBold = True
        .CellForeColor = vbBlue
        detailed_month_name = .Text
        .Col = 1
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = vbYellow
      Else
        .CellFontBold = False
        .CellForeColor = vbBlack
        .Col = 1
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .CellBackColor = vbWhite
      End If
      
      .Col = NAME_COL
      .CellAlignment = flexAlignLeftCenter
      If (is_cardtrak_transaction(card_name_combo.Text)) Then
        .Text = " " + strip_off_ct_number(card_name_combo.Text)
      Else
        .Text = " " + card_name_combo.Text
      End If
      
      .Col = BALANCE_COL
      .CellForeColor = amount_color(cardtrak_monthly_summary(i).balance)
      .Text = currency_s(cardtrak_monthly_summary(i).balance)
      .CellFontBold = True
    
      .Col = PAID_COL
      .CellForeColor = amount_color(-cardtrak_monthly_summary(i).paid)
      .Text = currency_s(-cardtrak_monthly_summary(i).paid)
      .CellFontBold = True
      sum_paid = sum_paid - cardtrak_monthly_summary(i).paid
      
      .Col = PURCHASES_COL
      .CellForeColor = amount_color(cardtrak_monthly_summary(i).purchases)
      .Text = currency_s(cardtrak_monthly_summary(i).purchases)
      .CellFontBold = True
      sum_purchases = sum_purchases + cardtrak_monthly_summary(i).purchases
      
      .Col = INTEREST_COL
      .CellForeColor = amount_color(cardtrak_monthly_summary(i).interest)
      .Text = currency_s(cardtrak_monthly_summary(i).interest)
      .CellFontBold = True
      sum_interest = sum_interest + cardtrak_monthly_summary(i).interest
      
      .Col = LATE_COL
      .CellForeColor = amount_color(cardtrak_monthly_summary(i).late)
      .Text = currency_s(cardtrak_monthly_summary(i).late)
      .CellFontBold = True
      sum_late = sum_late + cardtrak_monthly_summary(i).late

      .Col = MIN_COL
      .CellForeColor = amount_color(cardtrak_monthly_summary(i).minimum)
      .Text = currency_s(cardtrak_monthly_summary(i).minimum)
      .CellFontBold = True
      sum_minimum = sum_minimum + cardtrak_monthly_summary(i).minimum
      
      net = -cardtrak_monthly_summary(i).paid - cardtrak_monthly_summary(i).purchases - cardtrak_monthly_summary(i).interest - cardtrak_monthly_summary(i).late
      .Col = NET_COL
      .CellForeColor = amount_color(net)
      .Text = currency_s(net)
      .CellFontBold = True
      sum_net = sum_net + net
            
      .Col = EST_BALANCE_COL
      .CellForeColor = amount_color(cardtrak_monthly_summary(i).balance + cardtrak_monthly_summary(i).paid)
      .Text = currency_s(cardtrak_monthly_summary(i).balance + cardtrak_monthly_summary(i).paid)
      .CellFontBold = True
      
    Next i
  
    ' --------- Show the totals --------=
    .TextMatrix(1, PAID_COL) = currency_s(sum_paid)
    .TextMatrix(1, PURCHASES_COL) = currency_s(sum_purchases)
    .TextMatrix(1, INTEREST_COL) = currency_s(sum_interest)
    .TextMatrix(1, LATE_COL) = currency_s(sum_late)
    .TextMatrix(1, MIN_COL) = currency_s(sum_minimum)
    .TextMatrix(1, NET_COL) = currency_s(sum_net)
  
    ' --------- Display the individual cards now ----------
    's = "----- " + detailed_month_name + " Detailed Summary -----"
    s = detailed_month_name + " " + words(DETAILED_SUMMARY_N)
    .row = .row + 2  ' Skip a row
    .Rows = .row + 1
    .ColAlignment(1) = flexAlignLeftCenter
    .TextMatrix(.row, 1) = s  ' Allow these cells to be merged into one
    .TextMatrix(.row, 2) = s
    .TextMatrix(.row, 3) = s
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeRow(.row) = True
    .Col = 1
    .CellFontBold = True
    
    ' ---------- Display the total for the month -----------
    .Rows = .Rows + 1
    .row = .Rows - 1
    .TextMatrix(.row, 0) = words(TOTAL_N)
        
    .Col = BALANCE_COL
    .Text = currency_s(cardtrak_summary_single_month.balance)
      
    .Col = PAID_COL
    .Text = currency_s(-cardtrak_summary_single_month.paid)
      
    .Col = PURCHASES_COL
    .Text = currency_s(cardtrak_summary_single_month.purchases)
      
    .Col = INTEREST_COL
    .Text = currency_s(cardtrak_summary_single_month.interest)
      
    .Col = LATE_COL
    .Text = currency_s(cardtrak_summary_single_month.late)
     
    .Col = MIN_COL
    .Text = currency_s(cardtrak_summary_single_month.minimum)
        
    net = -cardtrak_summary_single_month.paid - cardtrak_summary_single_month.purchases - cardtrak_summary_single_month.interest - cardtrak_summary_single_month.late
    .Col = NET_COL
    .Text = currency_s(net)
    
    .Col = EST_BALANCE_COL
    .Text = currency_s(cardtrak_summary_single_month.balance + cardtrak_summary_single_month.paid)
        
    .Rows = .Rows + 1
    .row = .row + 1
    .RowHeight(.row) = 50
    
    ' --------- Display the individual cards now ----------
    For i = 1 To MAX_CARDS
      If (cardtrak_summary(i).active) Then
        ' We have data for this card so display it
        ' Put the name in the combo box
        
        .Rows = .Rows + 1
        .row = .Rows - 1
        .TextMatrix(.row, 0) = "CT" + Format(i, "00")
        .Col = NAME_COL
        .Text = " " + strip_off_ct_number(cardtrak_summary(i).name)
      
        .Col = BALANCE_COL
        .CellForeColor = amount_color(cardtrak_summary(i).balance)
        .Text = currency_s(cardtrak_summary(i).balance)
        .CellFontBold = True
      
        .Col = PAID_COL
        .CellForeColor = amount_color(-cardtrak_summary(i).paid)
        .Text = currency_s(-cardtrak_summary(i).paid)
        .CellFontBold = True
      
        .Col = PURCHASES_COL
        .CellForeColor = amount_color(cardtrak_summary(i).purchases)
        .Text = currency_s(cardtrak_summary(i).purchases)
        .CellFontBold = True
      
        .Col = INTEREST_COL
        .CellForeColor = amount_color(cardtrak_summary(i).interest)
        .Text = currency_s(cardtrak_summary(i).interest)
        .CellFontBold = True
      
        .Col = LATE_COL
        .CellForeColor = amount_color(cardtrak_summary(i).late)
        .Text = currency_s(cardtrak_summary(i).late)
        .CellFontBold = True
        
        .Col = MIN_COL
        .CellForeColor = amount_color(cardtrak_summary(i).minimum)
        .Text = currency_s(cardtrak_summary(i).minimum)
        .CellFontBold = True
        
        .Col = NET_COL
        net = -cardtrak_summary(i).paid - cardtrak_summary(i).purchases - cardtrak_summary(i).interest - cardtrak_summary(i).late
        .CellForeColor = amount_color(net)
        .Text = currency_s(net)
        .CellFontBold = True
      
        .Col = EST_BALANCE_COL
        .CellForeColor = amount_color(cardtrak_summary(i).balance + cardtrak_summary(i).paid)
        .Text = currency_s(cardtrak_summary(i).balance + cardtrak_summary(i).paid)
        .CellFontBold = True
        
      End If
    Next i
    
    
    .Rows = .row + 1
    .Redraw = True
    
  End With
  
  update_graph
  
  'ct_summary_form.Caption = "Cardtrak Summary --- " + main_form.entry_tab.Caption  'TabCaption(i)
  ct_summary_form.Caption = words(CARDTRAK_SUMMARY_N) + " --- " + main_form.entry_tab.TabCaption(0) + " - " + main_form.entry_tab.TabCaption(11)
  
  updating = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Hide
End Sub


Private Sub graph_combo_Click()
  If (Not updating) Then process_button_Click
End Sub

Private Sub graph_month_label_Click(index As Integer)
  main_form.entry_tab.Tab = index
End Sub

Private Sub grid_DblClick()
  With grid
    If ((.row >= 3) And (.row < 15)) Then
      main_form.entry_tab.Tab = .row - 3
    End If
  End With
  
End Sub

Private Sub high_low_combo_Click()
  If (Not updating) Then process_button_Click
End Sub

Private Sub month_button_Click(index As Integer)
  If (index = 0) Then
    main_form.previous_month_menu_Click
  Else
    main_form.next_month_menu_Click
  End If
End Sub

Private Sub Picture1_Click(index As Integer)
  main_form.entry_tab.Tab = index
End Sub

Private Sub Picture1_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  ' esk
  ' We moved the mouse over this item so display tool tip
  Dim delta As Double
  If index = 0 Then
    delta = Bar_Value(index)
  Else
    delta = Bar_Value(index) - Bar_Value(index - 1)
  End If
  Picture1(index).ToolTipText = " " + currency_s(Bar_Value(index)) + " / " + currency_s(delta)
End Sub

Private Sub process_button_Click()
  ' see if we have a name to filter on
  last_selected_index = card_name_combo.ListIndex
  cardtrak_filter = Val(Mid(card_name_combo.Text, 4, 4))
  main_form.process
End Sub

Private Sub graph_change_label_Click()
  ' View the printout
  ' Print out the cardtrak stuff
  Dim f, i
  
  printer_error = False
  print_destination = SCR
  
  On Error GoTo error_h
  ct_summary_form.Refresh
  
  ' We have hit ok so print it out
  main_form.MousePointer = 11
  Call view_printout
  
  print_form.Hide
  main_form.MousePointer = 0
  
  Exit Sub
error_h:
  
  main_form.MousePointer = 0
  MsgBox (Err.Description)
  
End Sub

Private Sub print_button_Click()
  ' Print out the cardtrak stuff
  Dim f, i
  
  printer_error = False
  print_destination = PTR ' SCR
  
  f = cdlPDNoPageNums Or cdlPDNoSelection Or cdlPDUseDevModeCopies
  'f = cdlPDHidePrintToFile Or cdlPDNoPageNums Or cdlPDNoSelection Or cdlPDUseDevModeCopies Or cdlPDNoPageNums
  print_dialog.FLAGS = f
  
  On Error GoTo error_h
  print_dialog.Copies = 1
  print_dialog.ShowPrinter
  
  ct_summary_form.Refresh
  
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
  'MsgBox (Err.Description)
  
End Sub


Private Sub do_new_page()
  page_number = page_number + 1
  
  If (page_number = 1) Then print_form.start_document (print_destination)
  
  If (page_number <> 1) Then print_form.new_page
  Call print_header_cardtrak
  y = 60
  transaction_line_count = 1

End Sub

Private Sub display_vertical_lines(y1 As Integer, y2 As Integer)
  Dim yy As Integer
  
  ' Draw the vertical lines
  'yy = y + j * 14 + 14
  print_form.FontName = "Arial"
  
  Call print_form.print_line(0, y1, 0, y2, 1)
  Call print_form.print_line(60, y1, 60, y2, 1)
  
  Call print_form.print_line(180, y1, 180, y2, 1)
  Call print_form.print_line(245, y1, 245, y2, 1)
  Call print_form.print_line(310, y1, 310, y2, 1)
  Call print_form.print_line(375, y1, 375, y2, 1)
  Call print_form.print_line(440, y1, 440, y2, 1)
  Call print_form.print_line(505, y1, 505, y2, 1)
  Call print_form.print_line(570, y1, 570, y2, 1)
  Call print_form.print_line(635, y1, 635, y2, 1)
  Call print_form.print_line(700, y1, 700, y2, 1)
  
  Call print_form.print_line(0, y2, 700, y2, 1)

End Sub

Private Sub print_header_cardtrak()
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

Private Sub view_printout()
  Dim last_date_s
  Dim t, m, n, j
  Dim s As String
  Dim i As Integer
  Dim w1, w2 As Integer
  
  print_form.last_page = False
  
  page_number = 0
  line_count = 0
  page_type = CARDTRAK_PAGE
  date_s = Format(date)
  time_s = Format(Time, "h:mm AMPM")
  
  Call do_new_page
  
  last_date_s = ""
  
  y = 60
  Call print_form.print_next(175, y, 0, ct_summary_form.Caption, 14)
  
  y = y + 28
  
  ' ----- Now copy the graph frame out -----
  print_form.print_picture CaptureWindow(Frame1.hWnd, False, 0, 0, _
        fTwipsToPixels(Frame1.width, DIRECTION_HORIZONTAL), _
        fTwipsToPixels(Frame1.height, DIRECTION_VERTICAL)), 30, y
  
  ' ----- Copy out the table of numbers now -----
  y = 240  ' Top of text
  y1 = y  ' Top of vertical lines
  
  With grid
    For i = 0 To .Rows - 1
    Call print_form.print_next_right(55, y, .TextMatrix(i, 0), 9)  ' Get the date
    Call print_form.print_next(65, y, 170, Left(.TextMatrix(i, 1), 19), 9)  ' Get the name and strip it to this many characters
    
    If (i = 1) Then Call print_form.print_line(0, y, 700, y, 2)
    If (i = 15) Then Call print_form.print_line(0, y, 700, y, 1)
    If (i = 18) Then Call print_form.print_line(0, y, 700, y, 2)
    
    If (i <> 16) Then  ' Skip the line that says Summary...
      Call print_form.print_next_right(245, y, .TextMatrix(i, 2), 9)  ' Balance
      Call print_form.print_next_right(310, y, .TextMatrix(i, 3), 9)  ' Paid
      Call print_form.print_next_right(375, y, .TextMatrix(i, 4), 9)  ' Purchases
      Call print_form.print_next_right(440, y, .TextMatrix(i, 5), 9)  ' Interest
      Call print_form.print_next_right(505, y, .TextMatrix(i, 6), 9)  ' Late
      Call print_form.print_next_right(570, y, .TextMatrix(i, 7), 9)  ' Minimum
      Call print_form.print_next_right(635, y, .TextMatrix(i, 8), 9)  ' Net
      'Call print_form.print_next_right(700, y, .TextMatrix(i, 9), 9)  ' This column is not used
      Call print_form.print_next_right(700, y, .TextMatrix(i, 10), 9)  ' Est Bal
    End If
    y = y + 14
    Next i
  End With
  
  ' ----- Display the lines -----
  print_form.FontName = "Arial"
  Call print_form.print_line(0, y1, 700, y1, 1)
  Call display_vertical_lines(y1, y)
  
  
  print_form.last_page = True
  If (print_destination = SCR) Then print_form.show (vbModal)
  print_form.end_document  ' All done with printing or screen
End Sub

Public Sub make_it_visible()
  ct_summary_form.SetFocus
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Timer1.Enabled = False
  On Error GoTo error_h
  'If (Not ct_summary_form.) Then
    ' Set this form to the active one if it's not currently active
    If (Not form_has_focus) Then ct_summary_form.SetFocus
  'End If
error_h:
End Sub
