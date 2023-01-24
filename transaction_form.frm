VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form transaction_form 
   Caption         =   "Transaction Entry Form"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   Icon            =   "transaction_form.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3915
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   5040
      TabIndex        =   24
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton next_month_button 
      Caption         =   ">>"
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
      Left            =   5220
      TabIndex        =   23
      Top             =   3660
      Width           =   555
   End
   Begin VB.CommandButton previous_month_button 
      Caption         =   "<<"
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
      Left            =   4560
      TabIndex        =   22
      Top             =   3660
      Width           =   555
   End
   Begin MSACAL.Calendar calendar 
      Height          =   2115
      Left            =   3720
      TabIndex        =   21
      Top             =   1560
      Width           =   2835
      _Version        =   524288
      _ExtentX        =   5001
      _ExtentY        =   3731
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   1999
      Month           =   2
      Day             =   2
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   0   'False
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   20
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   540
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type"
      Height          =   915
      Left            =   3000
      TabIndex        =   13
      Top             =   540
      Width           =   1515
      Begin VB.OptionButton Option3 
         Caption         =   "Deposit"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1155
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Withdrawal"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frequency"
      Height          =   1935
      Left            =   180
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
      Begin VB.OptionButton Option1 
         Caption         =   "Annual"
         Height          =   195
         Index           =   8
         Left            =   1980
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Semi Annual"
         Height          =   195
         Index           =   7
         Left            =   1980
         TabIndex        =   11
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Quarterly"
         Height          =   195
         Index           =   6
         Left            =   1980
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Monthly"
         Height          =   195
         Index           =   5
         Left            =   1980
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "4 Weeks"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1500
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Twice Montly"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2 Weeks"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   900
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Weekly"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "None"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.TextBox date_box 
      Height          =   345
      Left            =   1140
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   960
      Width           =   1470
   End
   Begin VB.TextBox amount_box 
      Height          =   345
      Left            =   1140
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   540
      Width           =   1470
   End
   Begin VB.TextBox name_box 
      Height          =   345
      Left            =   1140
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3390
   End
   Begin VB.Label Label3 
      Caption         =   "Next Date"
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
      Left            =   120
      TabIndex        =   18
      Top             =   1020
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
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
      Left            =   540
      TabIndex        =   17
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Amount"
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
      Left            =   420
      TabIndex        =   16
      Top             =   600
      Width           =   675
   End
End
Attribute VB_Name = "transaction_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub calendar_DblClick()
  date_box.Text = Str(calendar.month) + " /" + Str(calendar.day) + " /" + Str(calendar.year)
End Sub

Private Sub Command2_Click()
'gstrDBName = "transactions.mdb"

End Sub

Private Sub Command3_Click()
  ' Add the data to the database
  
End Sub

Private Sub next_month_button_Click()
If calendar.month < 12 Then
  calendar.month = calendar.month + 1
Else
  calendar.month = 1
  calendar.year = calendar.year + 1
End If
End Sub

Private Sub previous_month_button_Click()
If calendar.month > 1 Then
  calendar.month = calendar.month - 1
Else
  calendar.month = 12
  calendar.year = calendar.year - 1
End If
End Sub
