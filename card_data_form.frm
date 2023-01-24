VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{32A4927E-FB95-11D1-BF5B-00A024982E5B}#121.0#0"; "axGrid.ocx"
Begin VB.Form card_data_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credit Card Transcations"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7305
   Icon            =   "card_data_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton copy_name_button 
      Caption         =   "Copy"
      Height          =   330
      Left            =   45
      TabIndex        =   52
      Top             =   45
      Width           =   555
   End
   Begin VB.TextBox trans_name_box 
      Height          =   315
      Left            =   630
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   360
      Width           =   4275
   End
   Begin VB.CommandButton OK_Button 
      Height          =   495
      Left            =   6180
      Picture         =   "card_data_form.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   1035
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   60
      Width           =   1035
   End
   Begin TabDlg.SSTab tab_strip 
      Height          =   5595
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9869
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Credit Card"
      TabPicture(0)   =   "card_data_form.frx":06BD
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "card_name_label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "card_name_label"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "new_card_button"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "delete_card_button"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "select_card_button"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "info_box"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "card_name_grid"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtedit"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "select_radio(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "select_radio(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Payment"
      TabPicture(1)   =   "card_data_form.frx":06D9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "notes_label"
      Tab(1).Control(1)=   "finance_frame"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(3)=   "notes_box"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Transactions"
      TabPicture(2)   =   "card_data_form.frx":06F5
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "interest_grid"
      Tab(2).Control(1)=   "transaction_grid"
      Tab(2).Control(2)=   "transactions_label"
      Tab(2).Control(3)=   "interest_label"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Database"
      TabPicture(3)   =   "card_data_form.frx":0711
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "test_fill_button"
      Tab(3).Control(1)=   "test_grid"
      Tab(3).Control(2)=   "test_total_label"
      Tab(3).Control(3)=   "Label1"
      Tab(3).Control(4)=   "test_current_index_label"
      Tab(3).ControlCount=   5
      Begin VB.OptionButton select_radio 
         Caption         =   "Edit"
         Height          =   240
         Index           =   1
         Left            =   5625
         TabIndex        =   54
         Top             =   765
         Width           =   1140
      End
      Begin VB.OptionButton select_radio 
         Caption         =   "Select"
         Height          =   285
         Index           =   0
         Left            =   5625
         TabIndex        =   53
         Top             =   495
         Width           =   1230
      End
      Begin VB.CommandButton test_fill_button 
         Caption         =   "Fill Test Grid"
         Height          =   465
         Left            =   -74775
         TabIndex        =   50
         Top             =   855
         Width           =   1860
      End
      Begin MSFlexGridLib.MSFlexGrid test_grid 
         Height          =   3120
         Left            =   -74730
         TabIndex        =   49
         Top             =   2250
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   5503
         _Version        =   393216
         Rows            =   2001
         Cols            =   7
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
      End
      Begin VB.TextBox txtedit 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3960
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   960
         Width           =   495
      End
      Begin MSFlexGridLib.MSFlexGrid card_name_grid 
         Height          =   2295
         Left            =   240
         TabIndex        =   43
         Top             =   1200
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   3
         ScrollTrack     =   -1  'True
      End
      Begin VB.TextBox notes_box 
         Height          =   2175
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   3240
         Width           =   6615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Payment Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   -71160
         TabIndex        =   24
         Top             =   480
         Width           =   3075
         Begin VB.ComboBox status_combo 
            Height          =   315
            ItemData        =   "card_data_form.frx":072D
            Left            =   1800
            List            =   "card_data_form.frx":073D
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   1770
            TabIndex        =   29
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   1800
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   1800
            TabIndex        =   27
            Text            =   "1"
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   1800
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   1800
            TabIndex        =   25
            Text            =   "1"
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label status_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Status"
            Height          =   255
            Left            =   60
            TabIndex        =   36
            Top             =   1860
            Width           =   1695
         End
         Begin VB.Label check_number_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Check Number"
            Height          =   255
            Left            =   60
            TabIndex        =   35
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label amount_paid_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount Paid"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1260
            Width           =   1635
         End
         Begin VB.Label date_paid_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Date Paid"
            Height          =   255
            Left            =   60
            TabIndex        =   33
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label amount_due_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount Due"
            Height          =   255
            Left            =   60
            TabIndex        =   32
            Top             =   660
            Width           =   1695
         End
         Begin VB.Label date_due_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Date Due"
            Height          =   255
            Left            =   60
            TabIndex        =   31
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame finance_frame 
         Caption         =   "Finance Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   3435
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2100
            TabIndex        =   17
            Text            =   "0.00"
            Top             =   1500
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   2100
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   2100
            TabIndex        =   15
            Text            =   "0.00"
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2100
            TabIndex        =   14
            Text            =   "0.00"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2100
            TabIndex        =   13
            Text            =   "0.00"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox finance_box 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   2100
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label late_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Late Charges"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label finance_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Finance Charges"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   1260
            Width           =   1935
         End
         Begin VB.Label payments_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Payments && Credits"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label purchases_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Purchases && Advances"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   60
            TabIndex        =   20
            Top             =   660
            Width           =   1995
         End
         Begin VB.Label previous_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Previous Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label current_label 
            Alignment       =   1  'Right Justify
            Caption         =   "Current Balance"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1860
            Width           =   1935
         End
      End
      Begin VB.TextBox info_box 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   3840
         Width           =   6555
      End
      Begin VB.CommandButton select_card_button 
         Caption         =   "Select This Card"
         Height          =   375
         Left            =   3780
         TabIndex        =   7
         Top             =   495
         Width           =   1635
      End
      Begin VB.CommandButton delete_card_button 
         Caption         =   "Delete This Card"
         Height          =   375
         Left            =   2025
         TabIndex        =   6
         Top             =   480
         Width           =   1635
      End
      Begin VB.CommandButton new_card_button 
         Caption         =   "Add New Card"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin axGridControl.axgrid interest_grid 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   37
         Top             =   3375
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         Cols            =   6
         Row             =   10
         Redraw          =   -1  'True
         ShowGrid        =   -1  'True
         GridSolid       =   -1  'True
         GridLineColor   =   12632256
         AllowSelection  =   0   'False
         BackColorFixed  =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin axGridControl.axgrid transaction_grid 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   39
         Top             =   780
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3836
         Rows            =   200
         Cols            =   4
         Row             =   10
         Redraw          =   -1  'True
         ShowGrid        =   -1  'True
         GridSolid       =   -1  'True
         GridLineColor   =   12632256
         AllowSelection  =   0   'False
         BackColorFixed  =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label test_total_label 
         Caption         =   "Total"
         Height          =   240
         Left            =   -74730
         TabIndex        =   51
         Top             =   1890
         Width           =   2760
      End
      Begin VB.Label Label1 
         Caption         =   "Current CT Number:"
         Height          =   240
         Left            =   -74775
         TabIndex        =   48
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label test_current_index_label 
         Caption         =   "Label1"
         Height          =   285
         Left            =   -73290
         TabIndex        =   47
         Top             =   585
         Width           =   870
      End
      Begin VB.Label card_name_label 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   255
         Left            =   6075
         TabIndex        =   46
         Top             =   3555
         Width           =   735
      End
      Begin VB.Label notes_label 
         Caption         =   "Notes"
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
         Left            =   -74760
         TabIndex        =   42
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label transactions_label 
         Caption         =   "Transactions"
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
         Left            =   -74760
         TabIndex        =   40
         Top             =   480
         Width           =   3435
      End
      Begin VB.Label interest_label 
         Caption         =   "Interest"
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
         Left            =   -74760
         TabIndex        =   38
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "Name / Address / Phone numbers  / Credit limits / Comments"
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
         Left            =   300
         TabIndex        =   10
         Top             =   3540
         Width           =   6495
      End
      Begin VB.Label card_name_label1 
         Caption         =   "Card Name"
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
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Label incoming_name_label 
      Caption         =   "Label1"
      Height          =   255
      Left            =   675
      TabIndex        =   8
      Top             =   90
      Width           =   4215
   End
   Begin VB.Label card_number_label 
      Caption         =   "Label1"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   405
      Width           =   495
   End
   Begin VB.Menu card_menu 
      Caption         =   "Card"
      Begin VB.Menu new_card_menu 
         Caption         =   "Add New Card"
      End
      Begin VB.Menu delete_card_menu 
         Caption         =   "Delete Card"
      End
      Begin VB.Menu spare1 
         Caption         =   "-"
      End
      Begin VB.Menu delete_all_cards_menu 
         Caption         =   "Delete All Cards"
      End
      Begin VB.Menu spare3 
         Caption         =   "-"
      End
      Begin VB.Menu exit_menu 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit_menu 
      Caption         =   "Edit"
      Begin VB.Menu clear_transaction_menu 
         Caption         =   "Clear Transactions"
      End
      Begin VB.Menu clear_interest_menu 
         Caption         =   "Clear Interest"
      End
      Begin VB.Menu spare2 
         Caption         =   "-"
      End
      Begin VB.Menu clear_all_menu 
         Caption         =   "Clear All"
      End
   End
End
Attribute VB_Name = "card_data_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Define the transaction grid columns
Private Const T_N_COL = 0
Private Const T_NAME_COL = 1
Private Const T_AMOUNT_COL = 2
Private Const T_DATE_COL = 3
Private Const T_POSTED_COL = 4

Private Const I_N_COL = 0
Private Const I_NAME_COL = 1
Private Const I_BALANCE_COL = 2
Private Const I_RATE_COL = 3
Private Const I_PAYMENTS_COL = 4
Private Const I_CHARGES_COL = 5
Private Const I_INTEREST_COL = 6

Private Const A_N_COL = 0
Private Const A_NAME_COL = 1
Private Const A_ACCOUNT_COL = 2

Private Const TEST_N_COL = 0
Private Const TEST_ACTIVE_COL = 1
Private Const TEST_NAME_COL = 2
Private Const TEST_SUB_TRANS_COL = 3
Private Const TEST_DATE_COL = 4

Private Const TRANS_MAX = 200
Private Const INT_MAX = 10

Const AXALIGNLEFT = 0
Const AXALIGHCENTER = 1
Const AXALIGNRIGHT = 2

Private current_day As Integer
Private current_month As Integer
Private current_year As Integer
Private current_name_s As String
Private current_month_index As Integer

Private adding_cards_flag As Boolean  ' True when only adding cards and not doing transactions
Private changed As Boolean  ' True if anything has changed
Private row_index(MAX_CARDS)  ' Stores the row that each card is listed on
Private card_name_clicked As Boolean
Private info_last_card_number As Integer  ' Card number when info box is entered
Private last_col As Integer  ' Shows which column we were editing last
Private last_row As Integer  ' Shows which row we were editing last
Private last_card As Integer  ' Shows the last card we were working on

Private cardn As Integer  ' Used to save the card info
Private info_card_number As Integer  ' Card number currently being viewed
Private editing_a_record As Boolean  ' True when editing a msflexgrid

Private ok_hit  ' True for normal exit, false for cancel
Private editing_a_new_transaction  ' true when doing a new card db transaction

' Main working variables
Private local_cards_info(MAX_CARDS) As card_info_type  ' Working copy of the card info
Private local_card As card_transaction_type  ' Working copy for the card months transactions
Private local_this As r_type  ' This is the working variable
Private check_number_enter_counter As Integer  ' Counts how many enters were hit on an empty check number box to insert the auto check number


Public Sub update_language()
  card_data_form.Caption = words(CREDIT_CARD_TRANSACTIONS_N)
  card_menu.Caption = words(CARD_N)
  new_card_menu.Caption = words(ADD_NEW_CARD_N)
  delete_card_menu.Caption = words(DELETE_CARD_N)
  delete_all_cards_menu.Caption = words(DELETE_ALL_CARDS_N)
  exit_menu.Caption = words(EXIT_N)
  edit_menu.Caption = words(EDIT_N)
  clear_transaction_menu.Caption = words(CLEAR_TRANSACTIONS_N)
  clear_interest_menu.Caption = words(CLEAR_INTEREST_N)
  clear_all_menu.Caption = words(CLEAR_ALL_N)
  copy_name_button.Caption = words(COPY_N)
  CancelButton.Caption = words(CANCEL_N)
  ' ok_button.Caption = words(OK_N)
  tab_strip.TabCaption(0) = words(CREDIT_CARD_N)
  tab_strip.TabCaption(1) = words(PAYMENT_N)
  tab_strip.TabCaption(2) = words(TRANSACTIONS_N)
  new_card_button.Caption = words(ADD_NEW_CARD_N)
  delete_card_button.Caption = words(DELETE_THIS_CARD_N)
  select_card_button.Caption = words(SELECT_THIS_CARD_N)
  card_name_label1.Caption = words(CARD_NAME_N)
  select_radio(0).Caption = words(SELECT_N)
  select_radio(1).Caption = words(EDIT_N)
  Label18.Caption = words(NAME_ADDRESS_PHONE_NUMBERS_N)
  finance_frame.Caption = words(FINANCE_INFORMATION_N)
  previous_label.Caption = words(PREVIOUS_BALANCE_N)
  purchases_label.Caption = words(PURCHASES_ADVANCES_N)
  payments_label.Caption = words(PAYMENTS_CREDITS_N)
  finance_label.Caption = words(FINANCE_CHARGES_N)
  late_label.Caption = words(LATE_CHARGES_N)
  current_label.Caption = words(CURRENT_BALANCE_N)
  date_due_label.Caption = words(DATE_DUE_N)
  amount_due_label.Caption = words(AMOUNT_DUE_N)
  date_paid_label.Caption = words(DATE_PAID_N)
  amount_paid_label.Caption = words(AMOUNT_PAID_N)
  check_number_label.Caption = words(CHECK_NUMBER_N)
  status_label.Caption = words(STATUS_N)
  status_combo.Clear
  status_combo.AddItem (words(BLANK_N))
  status_combo.AddItem (words(DONE_N))
  status_combo.AddItem (words(PENDING_N))
  status_combo.AddItem (words(SKIP_N))
  notes_label.Caption = words(NOTES_N)
  transactions_label.Caption = words(TRANSACTIONS_N)
  interest_label.Caption = words(INTEREST_N)
  
  
  
End Sub


Private Sub CancelButton_Click()
  Dim ans As Integer
  
  If (changed) Then
    ans = MsgBox(words(SAVE_CHANGES_Q_N), vbYesNoCancel + vbQuestion, words(SAVE_Q_N))
    If (ans = vbYes) Then
      ok_button_Click
      Exit Sub
    End If
    If (ans = vbCancel) Then Exit Sub
  End If
  Hide
End Sub

Private Sub update_card_info()
  info_box.Text = local_cards_info(info_card_number).notes
  card_name_label.Caption = "CT" + Format(info_card_number, "00")
End Sub


Private Sub card_name_grid_Click()
  With card_name_grid
    changed = True
    If (card_name_grid.row > 0) Then card_name_grid.CellBackColor = vbYellow
    last_card = Val(.TextMatrix(.row, 0))
    last_col = .Col
    last_row = .row
  End With

End Sub

Private Sub card_name_grid_EnterCell()
  If (card_name_grid.row > 0) Then card_name_grid.CellBackColor = vbYellow
  With card_name_grid
    last_card = Val(.TextMatrix(.row, 0))
    last_col = .Col
    last_row = .row
    info_card_number = Val(.TextMatrix(.row, 0))
    update_card_info
  End With
End Sub

Private Sub card_name_grid_LeaveCell()
  If (card_name_grid.row > 0) Then card_name_grid.CellBackColor = vbWhite
End Sub

Private Sub card_name_grid_DblClick()
  If (select_radio(0).Value = False) Then
    ' Edit the cell contents
    With card_name_grid
      last_card = Val(.TextMatrix(.row, 0))
      last_col = .Col
      last_row = .row
       editing_a_record = True
       card_name_grid_keypress (0)
       txtEdit.SelStart = 0
       txtEdit.SelLength = 100
    End With
  Else
    ' We are selecting this card
    select_card_button_Click
  End If
  
End Sub

Private Sub card_name_grid_keypress(KeyAscii As Integer)
  If (editing_a_record) Then
    MSFlexGridEdit card_name_grid, txtEdit, KeyAscii
  Else
    editing_a_record = True
    card_name_grid_keypress (KeyAscii)
  End If
End Sub

Private Sub card_name_grid_GotFocus()
  Dim n As Integer
  card_name_clicked = True
End Sub



Private Sub clear_all_menu_Click()
  clear_interest_menu_Click
  clear_transaction_menu_Click
End Sub

Private Sub clear_interest_menu_Click()
  Dim i As Integer
  
  ' Clear out all the interest fields
  For i = 1 To MAX_CARD_INTEREST
    local_card.interest(i).name = ""
    local_card.interest(i).balance = ""
    local_card.interest(i).percent = ""
    local_card.interest(i).payments = ""
    local_card.interest(i).charges = ""
    local_card.interest(i).interest = ""
  Next i
  changed = True
  update_interest_grid
End Sub

Private Sub clear_transaction_menu_Click()
  ' Clear out all the transactions
  local_card.transactions = ""
  update_transaction_grid
  changed = True
End Sub

Private Sub copy_name_button_Click()
  ' Copy the incoming name label into the trans name box
  trans_name_box.Text = incoming_name_label.Caption
End Sub

Private Sub delete_all_cards_menu_Click()
  ' Delete all the cards
  ' Don't delete any of the ct transactions
  
  Dim i As Integer
  
  If (MsgBox(words(DELETE_ALL_CREDIT_CARDS_AND_INFORMATION_Q_N), vbYesNoCancel + vbQuestion, words(DELETE_Q_N)) = vbYes) Then
    ' Yes, delete the card
    If (MsgBox(words(ARE_YOU_ABSOLUTELY_SURE_Q_N), vbYesNoCancel + vbQuestion, words(DELETE_Q_N)) = vbYes) Then
      For i = 0 To MAX_CARDS
        local_cards_info(i).active = False
      Next i
      MsgBox words(ALL_CREDIT_CARDS_HAVE_BEEN_DELETED_N)
      update_display
      changed = True
    End If
  End If
  
End Sub

Private Sub delete_card_button_Click()
  ' Delete the current card
  Dim n As Integer
  
  changed = True
  ' We have a row selected so now prompt to make sure
  n = Val(card_name_grid.TextMatrix(card_name_grid.row, 0))
  
  If (MsgBox(words(DELETE_N) + " '" + local_cards_info(n).name + "'?", vbYesNoCancel + vbQuestion, words(DELETE_Q_N)) = vbYes) Then
    ' Yes, delete the card
    'delete_card (n)
    local_cards_info(n).active = False
     
    update_display
    changed = True
  End If
End Sub




Private Function validate_field(index As Integer) As Boolean
  On Error GoTo error_h
  validate_field = False
  
  If (index = 0) Then local_card.previous_balance = finance_box(index)
  If (index = 1) Then local_card.total_purchases = finance_box(index)
  If (index = 2) Then local_card.total_payments = finance_box(index)
  If (index = 3) Then local_card.total_interest = finance_box(index)
  If (index = 4) Then local_card.total_late = finance_box(index)
  If (index = 5) Then local_card.new_balance = finance_box(index)
  If (index = 6) Then local_card.date_due = finance_box(index)
  If (index = 7) Then local_card.amount_due = finance_box(index)
  'If (index = 8) Then local_card.date_paid = finance_box(index)
  If (index = 8) Then local_card.amount_paid = finance_box(index)
  If (index = 9) Then
    If (finance_box(index) <> "") Then
      local_card.check_number = finance_box(index)  ' We have a valid check number
      ' Mark the transaction as pending
      local_card.status = PAID_DONE
    Else
      local_card.check_number = -1  ' Blank check number so make it -1
    End If
  End If
  
  validate_field = True
  Exit Function
  
error_h:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
End Function

Private Sub delete_card_menu_Click()
  delete_card_button_Click
End Sub

Private Sub exit_menu_Click()
  ok_button_Click
End Sub

Private Sub finance_box_DblClick(index As Integer)
  If (index = 9) And (finance_box(9).Text = "") Then
    data.last_check_number = data.last_check_number + 1
    finance_box(9).Text = data.last_check_number
    If (preferences.auto_check_done_on_check) Then status_combo.ListIndex = 1
  End If
End Sub

Private Sub finance_box_GotFocus(index As Integer)
  check_number_enter_counter = 0
  finance_box(index).SelStart = 0
  finance_box(index).SelLength = 100
  changed = True
End Sub

Private Sub finance_box_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
  If (index = 9) Then
    ' We are doing the check number box
    If (finance_box(9).Text = "") And (KeyCode = vbKeyReturn) Then
      check_number_enter_counter = check_number_enter_counter + 1
      If (check_number_enter_counter > 1) Then
          check_number_enter_counter = 0
          data.last_check_number = data.last_check_number + 1
          finance_box(9).Text = data.last_check_number
          If (preferences.auto_check_done_on_check) Then status_combo.ListIndex = 1
      End If
    End If
  
  End If
  
  If (KeyCode = vbKeyReturn) Then
    If (validate_field(index)) Then
        If (index < 9) Then finance_box(index + 1).SetFocus
        update_balance_boxes
    End If
  End If
  
  If (KeyCode = vbKeyEscape) Then
    ' Put the field back the way it was
    update_balance_boxes
  End If
End Sub

Private Sub finance_box_LostFocus(index As Integer)
  validate_field (index)
End Sub



Private Sub Form_Activate()
  Dim i
  
  card_name_clicked = False
  editing_a_record = False
  txtEdit.Visible = False  ' Hide the text box used for editing msgflexgrid
  
  select_radio(0).Value = True
  
  ' ----- Set up the transaction grid -----
  With transaction_grid
    For i = 1 To TRANS_MAX
      .TextMatrix(i, 0) = Format(i)
    Next i
    
    .ColWidth(T_NAME_COL) = 185
    .ColAlign(2) = AXALIGNRIGHT
    .ColAlign(3) = AXALIGNRIGHT
    .ColAlign(4) = AXALIGNRIGHT
    
    .TextMatrix(0, T_NAME_COL) = words(NAME_N)
    .TextMatrix(0, T_AMOUNT_COL) = words(AMOUNT_N)
    .TextMatrix(0, T_DATE_COL) = words(DATE_N)
    .TextMatrix(0, T_POSTED_COL) = words(POSTED_N)
   
    .Refresh
  End With
  
  ' ----- Set up the interest grid -----
  With interest_grid
    For i = 1 To INT_MAX
      .TextMatrix(i, I_N_COL) = Format(i)
    Next i
    
    .ColWidth(I_NAME_COL) = 100
    .ColWidth(I_BALANCE_COL) = 55
    .ColWidth(I_RATE_COL) = 55
    .ColWidth(I_PAYMENTS_COL) = 55
    .ColWidth(I_CHARGES_COL) = 55
    .ColWidth(I_INTEREST_COL) = 55
    
    .ColAlign(I_BALANCE_COL) = AXALIGNRIGHT
    .ColAlign(I_RATE_COL) = AXALIGNRIGHT
    .ColAlign(I_PAYMENTS_COL) = AXALIGNRIGHT
    .ColAlign(I_CHARGES_COL) = AXALIGNRIGHT
    .ColAlign(I_INTEREST_COL) = AXALIGNRIGHT
    
    .TextMatrix(0, I_NAME_COL) = words(NAME_N)
    .TextMatrix(0, I_BALANCE_COL) = words(BALANCE_N)
    .TextMatrix(0, I_RATE_COL) = words(RATE_PERCENT_N)
    .TextMatrix(0, I_PAYMENTS_COL) = words(PAYMENT_N)
    .TextMatrix(0, I_CHARGES_COL) = words(CHARGES_N)
    .TextMatrix(0, I_INTEREST_COL) = words(INTEREST_N)
    
    .Refresh
  End With
  
  ' ----- Set up the card name grid -----
  With card_name_grid
    .ColWidth(0) = 500
    .ColWidth(1) = 3600
    .ColWidth(2) = 2000
    .ColAlignment(1) = flexAlignLeftCenter
    .ColAlignment(2) = flexAlignLeftCenter
  End With
    
  ' Show the test tab only if the help about integrity has been clicked
  tab_strip.TabVisible(3) = main_form.database_integrity_menu.Visible
  
End Sub

Private Sub update_balance_boxes()
  finance_box(0).Text = currency_s(local_card.previous_balance)
  finance_box(1).Text = currency_s(local_card.total_purchases)
  finance_box(2).Text = currency_s(local_card.total_payments)
  finance_box(3).Text = currency_s(local_card.total_interest)
  finance_box(4).Text = currency_s(local_card.total_late)
  finance_box(5).Text = currency_s(local_card.new_balance)
  finance_box(6).Text = local_card.date_due
  finance_box(7).Text = currency_s(local_card.amount_due)
  finance_box(8).Text = currency_s(local_card.amount_paid)
  finance_box(9).Text = ""
  If (local_card.check_number >= 0) Then finance_box(9).Text = local_card.check_number
  status_combo.ListIndex = local_card.status
  finance_box(11).Text = local_card.date_paid
End Sub

Private Sub update_interest_grid()
  Dim i As Integer
  
  With interest_grid
    For i = 1 To MAX_CARD_INTEREST
      .TextMatrix(i, I_NAME_COL) = local_card.interest(i).name
      .TextMatrix(i, I_BALANCE_COL) = s_or_currency(local_card.interest(i).balance)
      .TextMatrix(i, I_RATE_COL) = s_or_currency(local_card.interest(i).percent)
      .TextMatrix(i, I_PAYMENTS_COL) = s_or_currency(local_card.interest(i).payments)
      .TextMatrix(i, I_CHARGES_COL) = s_or_currency(local_card.interest(i).charges)
      .TextMatrix(i, I_INTEREST_COL) = s_or_currency(local_card.interest(i).interest)
    Next i
    .Refresh
  End With
End Sub

Private Function valid_number(s As String) As Boolean
  ' Return true if the incoming string is a valid currency number
  On Error GoTo error_h
  s = currency_s(s)
  valid_number = True
  Exit Function
error_h:
  valid_number = False
End Function
        
Private Function s_or_currency(s As String) As String
  ' Return the currency string if s is a valid number or the string itself if it's not
  s_or_currency = s
  If (valid_number(s)) Then s_or_currency = currency_s(s)
End Function


Private Sub update_transaction_grid()
  Dim i As Integer
  Dim s As String
  
  Call load_axgrid_from_string(local_card.transactions, transaction_grid)
  With transaction_grid
    For i = 1 To .Rows
      ' Display the transaction amount
      s = .TextMatrix(i, 2)
      If (s <> "") Then
        .TextMatrix(i, 2) = s_or_currency(s)
      End If
    Next i
  .Refresh
  End With
End Sub

Private Sub update_test_tab()
  test_current_index_label.Caption = Format(local_card.card_this)
End Sub

Private Sub update_display()
  Dim i
  Dim s As String
  Dim n
  
  ' Put the caption on the card name grid
  card_name_grid.TextMatrix(0, 1) = words(NAME_N)
  card_name_grid.TextMatrix(0, 2) = words(ACCOUNT_N)
  
  ' Display the form caption
  card_data_form.Caption = words(CREDIT_CARD_INFORMATION_N) + " - " + get_date(current_month, current_day, current_year)
  
  
  ' Update the card information
  update_card_info
  
  ' Put up the transaction grid info
  update_transaction_grid
  
  ' Put up the interest grid info
  update_interest_grid
  
  ' Put up the balances and charges
  update_balance_boxes
    
  ' Put up the notes
  notes_box.Text = local_card.notes
    
  fill_card_select_grid
    
  ' Update all the fields on the test tab
  update_test_tab
  
End Sub

Private Sub fill_card_select_grid()
  Dim n As Integer
  Dim i As Integer
  ' Fill up the card name combo
  With card_name_grid
  .Clear
  .TextMatrix(0, 1) = words(NAME_N)
  .TextMatrix(0, 2) = words(ACCOUNT_N)
  
  n = 0
  .Rows = 1
  For i = 1 To MAX_CARDS
    row_index(i) = 0
    If (local_cards_info(i).active) Then
      ' Add this to the combo
      n = n + 1
      .Rows = .Rows + 1
      '.row = .Rows - 1
      .TextMatrix(.Rows - 1, 0) = i
      .TextMatrix(.Rows - 1, 1) = (local_cards_info(i).name)
      .TextMatrix(.Rows - 1, 2) = (local_cards_info(i).account)
      row_index(i) = .Rows - 1 ' Save the row number that this card is on
    End If
  Next i
  
  ' See if we have any cards defined. If not then create one
  If ((n = 0) And (info_card_number = 0)) Then
    ' No we have no cards defined
    info_card_number = 1  ' Default to card 1
    local_card.name = "New Credit Card"
    local_card.card_number = 1
  End If
  
  ' Update the card info notes
  info_card_number = Val(.TextMatrix(.row, 0))
  update_card_info
  
  ' ----- Hide or unhide some buttons -----
  If (n = 0) Then
    delete_card_button.Visible = False
    select_card_button.Visible = False
    delete_card_menu.Enabled = False
    delete_all_cards_menu.Enabled = False
    select_radio(0).Visible = False
    select_radio(1).Visible = False
  Else
    delete_card_button.Visible = True
    select_card_button.Visible = True
    delete_card_menu.Enabled = True
    delete_all_cards_menu.Enabled = True
    select_radio(0).Visible = True
    select_radio(1).Visible = True
  End If
  
  If (n = MAX_CARDS) Then
    new_card_button.Visible = False
    new_card_menu.Enabled = False
  Else
    new_card_button.Visible = True
    new_card_menu.Enabled = True
  End If
  
  ' Put up the transaction name
  card_number_label.Caption = "CT" + Format(local_card.card_number, "00")
  trans_name_box.Text = local_card.name
    
  card_name_label.Caption = "CT" + Format(info_card_number, "00")
  End With
End Sub


Private Function create_new_card_transaction() As Integer
  Dim i As Integer
  
  ' Loop through the card transactions and find and empty slot
  ' Return the slot number
  ' Fill in the fields with default data
  local_card.card_this = 0
  
  For i = 1 To MAX_CARD_TRANSACTIONS
    ' Loop through all the card transactions and find the matching one
    If (Not cards(i).active) Then Exit For
  Next i
    
  If (i <= MAX_CARD_TRANSACTIONS) Then
    ' We found a slot for the new record so initialize the cards record
    local_card.card_this = i  ' Save the record number in the card db
    local_card.active = True
    
    'card.card_number = n
    local_card.name = ""
    local_card.check_number = -1
    local_card.cleared = 0
    local_card.amount_due = 0#
    local_card.date_due = 1
    local_card.exclude = False
    local_card.total_late = 0#
    local_card.amount_paid = 0#
    local_card.previous_balance = 0#
    local_card.status = 2
    local_card.tags = 0
    local_card.previous_balance = 0
    local_card.new_balance = 0
    local_card.total_purchases = 0
    local_card.total_payments = 0
    local_card.total_interest = 0
    local_card.total_late = 0#
    local_card.transactions = ""
    
    local_card.date_due = local_this.day
    local_card.date_paid = local_this.day
        
    For i = 0 To MAX_CARD_INTEREST
      local_card.interest(i).balance = ""
      local_card.interest(i).percent = ""
      local_card.interest(i).payments = ""
      local_card.interest(i).charges = ""
      local_card.interest(i).interest = ""
    Next i
        
    local_card.notes = ""
    
    editing_a_new_transaction = True
  'MsgBox words(New_transaction_entered_n)  ' ????
  End If
  
  create_new_card_transaction = local_card.card_this
End Function

Private Sub copy_this_data()
  ' Copy the incoming this to local_card
  local_card.active = True
  
  local_card.name = local_this.name
  local_card.amount_paid = local_this.amount
  local_card.check_number = local_this.check
  local_card.cleared = local_this.cleared
  local_card.exclude = local_this.exclude
  local_card.status = local_this.paid
  local_card.tags = local_this.tags
  local_card.date_due = local_this.due
  local_card.date_paid = local_this.day
  info_card_number = cardn
  local_card.card_number = cardn
End Sub

Private Function validate_record() As Boolean
  Dim i
  
  ' Check the incoming main record and see if it matches with the card db
  ' Return the master record number if it matches
  ' Return -1 if there isn't a record in the card db
  validate_record = False
  editing_a_new_transaction = False
  
  If (local_this.sub_transaction_number = 0) Then
    ' We are working on a new blank incoming transaction
    local_card.card_this = create_new_card_transaction
  Else
    ' We do point to a record in the card db
    ' Check the subtransaction number to see if it points to the card db
    local_card = cards(local_this.sub_transaction_number)  ' Get the record from the card transaction db
    
    If (local_card.card_this <> local_this.sub_transaction_number) Then
      ' The card record doesn't match the one it should so create a new one
      local_card.card_this = create_new_card_transaction
    Else
      copy_this_data
    End If
  End If
  
  ' See if we had an incoming transaction
  If (local_this.this >= 0) Then
    ' Yes we did so get the name and number
    copy_this_data
    editing_a_new_transaction = False
  Else
    local_this.name = ""
  End If
  
  
  If (local_card.card_this > 0) Then validate_record = True
End Function

Private Sub initialize_local_this()
  ' Set up this as if it's a new unused record
  local_this.this = -1
  local_this.sub_transaction_number = 0
End Sub

Public Function execute(fcn As Integer) As Boolean
  Dim i, k
  'Dim r As Integer
  
  ' We enter here with THIS set to the record we are editing
  ' If fcn = 0 CT_EDIT    then we are editing an existing record
  ' If fcn = 1 CT_BLANK   then we are making a blank record
  ' If fcn = 2 CT_ADD     then we are adding new cards only
  ' If fcn = 3 CT_CONVERT then we are converting a normal transaction to a ct transaction
  ' Return true if the edit was ok or false if canceled
  ' THIS is the actual data we are working on
  
  adding_cards_flag = False
  tab_strip.TabVisible(1) = True
  tab_strip.TabVisible(2) = True
  If (fcn = CT_ADD) Then
    this.name = ""
    adding_cards_flag = True
    tab_strip.TabVisible(1) = False
    tab_strip.TabVisible(2) = False
  End If
  
  ' Get a copy of in cards_info database
  For i = 1 To MAX_CARDS
    local_cards_info(i) = cards_info(i)
  Next i
  
  cardn = 0
  initialize_local_this
  If (this.this >= 0) Then
    ' We have an existing record to work on
    local_this = this  'db(r)  ' Make a working copy of incoming transaction
    If (UCase(Mid(local_this.name, 2, 2)) = "CT") Then
      cardn = Val(Mid(local_this.name, 4, 2))
    
      ' Strip off the card number
      local_this.name = strip_off_ct_number(local_this.name)
    End If
  End If
  
  ' Set the dates to the current view date
  current_day = view.current_day
  current_month = view.current_month
  current_year = view.current_year
  local_this.day = current_day
  local_this.Month = current_month
  local_this.Year = current_year
  last_card = 1
  
  ' ----- Here we have n=the card number or n=0 for a new one -----
  If (validate_record()) Then
    ' We have a valid record so edit it
    If (editing_a_new_transaction) Or (fcn = CT_ADD) Then
      tab_strip.Tab = 0
    Else
      tab_strip.Tab = 1
    End If
  Else
    ' We have an error and cannot create a new record
  End If
  
  copy_name_button.Visible = False
  incoming_name_label.Visible = False
  incoming_name_label.Caption = local_this.name
  
  update_display
  
  ok_hit = False
  changed = False
  
  ' =============== Show the form =============
  card_data_form.show vbModal
  
  If (ok_hit) Then
    If (adding_cards_flag = False) Then
      ' We are editing a transaction so save the data
      ' ----- Save the card transaction -----
      If (local_card.amount_paid > 0) Then local_card.amount_paid = -local_card.amount_paid
      local_card.transactions = save_axgrid_into_string(transaction_grid)
      local_card.notes = notes_box.Text
      local_card.name = "(CT" + Format(local_card.card_number, "00") + ") " + trans_name_box.Text
      local_card.active = True
      cards(local_card.card_this) = local_card  ' Save the working variable into the permanent one
    
      ' ----- Copy the fields to THIS -----
      local_this.amount = local_card.amount_paid
      local_this.name = local_card.name
      local_this.sub_transaction_number = local_card.card_this
      this = local_this
    End If
    
    ' ----- Save the cards_info database -----
    For i = 0 To MAX_CARDS
      cards_info(i) = local_cards_info(i)
    Next i
  Else
    ' Cancel out
  End If
   
  execute = ok_hit
End Function


Private Function validate_form() As Boolean
  validate_form = True
  If (adding_cards_flag) Then
    Exit Function
  End If
  
  ' Validate any fields that needs to be filled in
  If (trans_name_box.Text = "") Then
    MsgBox words(MUST_HAVE_A_TRANSACTION_NAME_N)
    validate_form = False
  End If
  
  If (local_card.card_number = 0) Then
    MsgBox words(MUST_SELECT_A_CREDIT_CARD_N)
    validate_form = False
  End If
  
End Function





Private Sub info_box_GotFocus()
  info_last_card_number = info_card_number
  changed = True
End Sub

Private Sub info_box_Validate(Cancel As Boolean)
  local_cards_info(info_last_card_number).notes = info_box.Text
  Cancel = False
End Sub

Private Sub interest_grid_AfterEdit(row As Integer, Col As Integer, NewValue As String)
  With interest_grid
    If (.Col < I_INTEREST_COL) Then
      .Col = .Col + 1
    Else
      If (.row < .Rows) Then
        .row = .row + 1
        .Col = I_NAME_COL
      End If
    End If
  End With
    interest_grid_LostFocus
End Sub

Private Sub interest_grid_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn) Then
    interest_grid_LostFocus
  End If
End Sub

Private Sub interest_grid_LostFocus()
  Dim k
  
  With interest_grid
    For k = 1 To MAX_CARD_INTEREST
      local_card.interest(k).name = .TextMatrix(k, I_NAME_COL)
      local_card.interest(k).balance = .TextMatrix(k, I_BALANCE_COL)
      local_card.interest(k).percent = .TextMatrix(k, I_RATE_COL)
      local_card.interest(k).payments = .TextMatrix(k, I_PAYMENTS_COL)
      local_card.interest(k).charges = .TextMatrix(k, I_CHARGES_COL)
      local_card.interest(k).interest = .TextMatrix(k, I_INTEREST_COL)
    Next k
  End With
  update_interest_grid
  changed = True
End Sub

Private Sub new_card_menu_Click()
  new_card_button_Click
End Sub

Private Sub notes_box_LostFocus()
  local_card.notes = notes_box.Text
  changed = True
End Sub

Private Sub save_button_Click()
  local_card.transactions = save_axgrid_into_string(transaction_grid)
  
  update_display
End Sub


Private Sub select_card_button_Click()
  Dim n As Integer
  
  changed = True
  ' Get the current row and make this the current card
  ' We have a row selected so now prompt to make sure
  n = Val(card_name_grid.TextMatrix(card_name_grid.row, 0))
  
  If (MsgBox(words(SELECT_CARD_N) + " '" + local_cards_info(n).name + "'?", vbYesNoCancel + vbQuestion, words(SELECT_THIS_CARD_Q_N)) = vbYes) Then
    ' Yes, use this one
    local_card.name = local_cards_info(n).name
    local_card.card_number = n
    incoming_name_label.Visible = True
    copy_name_button.Visible = True
    update_display
  End If
  
End Sub

Private Sub status_combo_Change()
  changed = True
End Sub

Private Sub test_fill_button_Click()
  Dim i As Integer
  Dim ct_num As Integer
  Dim j As Integer, k As Integer
  Dim error_flag As Boolean
  Dim total As Integer
  Dim error_count As Integer
  Dim first_error As Integer
  
  ' Fill in the test grid with ct transactions
  error_flag = False
  error_count = 0
  total = 0
  
  With test_grid
    .Redraw = False
    .ColAlignment(TEST_NAME_COL) = flexAlignLeftCenter
    .TextMatrix(0, TEST_N_COL) = ""
    .TextMatrix(0, TEST_ACTIVE_COL) = "Active"
    .TextMatrix(0, TEST_NAME_COL) = "Name"
    .TextMatrix(0, TEST_SUB_TRANS_COL) = "Main"
    .TextMatrix(0, TEST_DATE_COL) = "Date"
    
    For i = 1 To MAX_CARD_TRANSACTIONS
      .TextMatrix(i, TEST_N_COL) = i
      .TextMatrix(i, TEST_ACTIVE_COL) = cards(i).active
      .TextMatrix(i, TEST_NAME_COL) = cards(i).name
      .TextMatrix(i, TEST_SUB_TRANS_COL) = ""
    Next i
    
    ' Scan through the entire main db and find matches for the ct transactions
    For i = 0 To MAX_DATA_TABLE
      If (db(i).this >= 0) And (db(i).sub_transaction_number > 0) Then
        ' We found a matching transaction
        total = total + 1
        ct_num = db(i).sub_transaction_number
        If (ct_num > 0) Then
          If (.TextMatrix(ct_num, TEST_SUB_TRANS_COL) <> "") Then
            ' We have an error, already have a filled in main number
            If (error_flag = False) Then first_error = ct_num
            error_flag = True
            error_count = error_count + 1
          End If
          
          .TextMatrix(ct_num, TEST_SUB_TRANS_COL) = .TextMatrix(ct_num, TEST_SUB_TRANS_COL) + Format(i)
          .TextMatrix(ct_num, TEST_DATE_COL) = get_date(db(i).Month, db(i).day, db(i).Year)
          
          ' Check for an error
          If (cards(ct_num).active = False) Then
            ' Yes we have a reference ct but the active flag is false
            If (error_flag = False) Then first_error = ct_num
            error_flag = True
            error_count = error_count + 1
          End If
        End If
      End If
    Next i
    
    .Redraw = True
    
  If (error_flag) Then MsgBox (Format(error_count) + " " + words(ERRORS_IN_CT_DATABASE_N) + " " + Format(first_error))
  test_total_label = "Total CT Records: " + Format(total)
  End With
End Sub


Private Sub trans_name_box_Click()
  incoming_name_label.Visible = True
  changed = True
End Sub

Private Sub transaction_grid_AfterEdit(row As Integer, Col As Integer, NewValue As String)
  ' Take us to the next column
  With transaction_grid
    If (.Col < T_POSTED_COL) Then
      .Col = .Col + 1
    Else
      If (.row < .Rows) Then
        .row = .row + 1
        .Col = T_NAME_COL
      End If
    End If
  End With
  transaction_grid_LostFocus
End Sub

Private Sub transaction_grid_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn) Then
    transaction_grid_LostFocus
  End If
End Sub

Private Sub transaction_grid_LostFocus()
  local_card.transactions = save_axgrid_into_string(transaction_grid)
  update_transaction_grid
  changed = True
End Sub

Private Sub ok_button_Click()
  ok_hit = True
  
  ' Save the transaction name
  If (local_this.this < 0) Then
    ' We have a new transaction so clear out things
    local_this.cleared = 0
    local_this.tags = 0
    local_this.exclude = False
  End If
  
  local_this.name = card_number_label + trans_name_box.Text
  local_this.amount = local_card.amount_paid
  local_this.check = local_card.check_number
  'rec.cleared = card.cleared
  'rec.exclude = card.exclude
  local_this.paid = status_combo.ListIndex
  'rec.tags = card.tags
  local_this.due = local_card.date_due
  'local_this.day = local_card.date_paid
  If (local_this.amount > 0) Then local_this.amount = -local_this.amount
  
  
  If (validate_form = False) Then ok_hit = False
  
  If (ok_hit) Then
    this = local_this
    Hide
  End If
End Sub


Private Sub new_card_button_Click()
  Dim i
  
  ' Add a new card
  If (MsgBox(words(ADD_NEW_CARD_Q_N), vbYesNoCancel + vbQuestion, words(ADD_Q_N)) = vbYes) Then
    changed = True
    ' Find the first free slot
      For i = 1 To MAX_CARDS
        If (Not local_cards_info(i).active) Then
          ' Create the new card now
          local_cards_info(i).active = True
          local_cards_info(i).name = "New card " + Format(i)
          local_cards_info(i).account = ""
          local_cards_info(i).created_day = day(Now)  ' Save the date this card was created
          local_cards_info(i).created_month = Month(Now)
          local_cards_info(i).created_year = Year(Now)
          
          update_display
          
          ' Select the row and name cell of this new card
          card_name_grid.row = row_index(i)
          card_name_grid.Col = 1
          last_row = card_name_grid.row
          last_col = card_name_grid.Col
          
          update_card_info
          card_name_grid.SetFocus  ' Put the focus here so we can edit the name now
          
          Exit For
        End If
      Next i
   
    If (i > MAX_CARDS) Then
      ' No free card slots
      MsgBox words(SORRY_NO_SPACE_FOR_CARDS_N)
    End If
  End If
End Sub


Private Sub tab_strip_Click(PreviousTab As Integer)
  ' We left a tab so save the data
  ' Prompt to save the data
'  If (PreviousTab = 1) Then
'    ' We have left the card info tab
'    local_cards_info(local_card.card_number).name = name_box.Text
'    local_cards_info(local_card.card_number).account = account_box.Text
'    local_cards_info(local_card.card_number).notes = info_box.Text
'    fill_card_select_table  ' Do this because the name may have changed
'  End If
   
End Sub


Private Sub txtedit_GotFocus()
  last_col = card_name_grid.Col
End Sub

Sub txtEdit_KeyPress(KeyAscii As Integer)
    ' Delete returns to get rid of beep.
    If KeyAscii = Val(vbCr) Then KeyAscii = 0
End Sub

Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    EditKeyCode card_name_grid, txtEdit, KeyCode, Shift
    If (KeyCode = 27) Then editing_a_record = False  ' See if escape was pressed
    changed_flag = True
End Sub

Private Sub txtedit_Validate(Cancel As Boolean)
  ' Get here if we switch out of the text box
  Dim n As Integer
  If txtEdit.Visible = False Then Exit Sub
  txtEdit.Visible = False
  With card_name_grid
    .TextMatrix(last_row, last_col) = txtEdit
    If (last_col = 1) Then
      local_cards_info(last_card).name = txtEdit
      .Col = 2
    Else
      If (last_col = 2) Then
        local_cards_info(last_card).account = txtEdit
        .Col = 1
      End If
    End If
  End With
  Cancel = False
End Sub

Private Sub txtEdit_LostFocus()
  ' Get here when we hit enter
  ' We just switched out of this text box so save the text
  If txtEdit.Visible = False Then Exit Sub
  txtEdit.Visible = False
  With card_name_grid
    .TextMatrix(last_row, last_col) = txtEdit
    If (last_col = 1) Then
      local_cards_info(last_card).name = txtEdit
      .Col = 2
    Else
      If (last_col = 2) Then
         local_cards_info(last_card).account = txtEdit
        .Col = 1
      End If
    End If
  End With
End Sub


Public Sub zero()
  Dim i As Integer
  
  ' Zero out all the ct transactions but not the cards
  For i = 0 To MAX_CARD_TRANSACTIONS
    cards(i).active = False
    cards(i).name = ""
  Next i
End Sub
