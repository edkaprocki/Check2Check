VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form cardtrack_form 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Card Track"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6525
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
      Width           =   1275
      Begin VB.OptionButton status_radio 
         Caption         =   "Pending"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   480
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton status_radio 
         Caption         =   "Done"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.CommandButton delete_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "cardtrack.frx":0000
      DownPicture     =   "cardtrack.frx":0102
      Height          =   315
      Left            =   240
      Picture         =   "cardtrack.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Delete Quick Save Account"
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton calendar_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "cardtrack.frx":0306
      DownPicture     =   "cardtrack.frx":0408
      Height          =   315
      Left            =   660
      Picture         =   "cardtrack.frx":050A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Show Calendar"
      Top             =   480
      Width           =   315
   End
   Begin VB.CommandButton calculator_button 
      BackColor       =   &H00C0C0C0&
      DisabledPicture =   "cardtrack.frx":0694
      DownPicture     =   "cardtrack.frx":0796
      Height          =   315
      Left            =   660
      Picture         =   "cardtrack.frx":0898
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Show Calculator"
      Top             =   120
      Width           =   315
   End
   Begin VB.Frame amount_frame 
      Caption         =   "Amount"
      Height          =   735
      Left            =   2580
      TabIndex        =   5
      Top             =   60
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
      Left            =   5340
      Picture         =   "cardtrack.frx":0C29
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   1035
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   180
      Width           =   1035
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
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   6482
      _Version        =   393216
      Rows            =   31
      Cols            =   7
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
   End
End
Attribute VB_Name = "cardtrack_form"
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


