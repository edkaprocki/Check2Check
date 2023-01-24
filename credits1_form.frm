VERSION 5.00
Begin VB.Form credits1_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Credits"
   ClientHeight    =   3960
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7995
   ClipControls    =   0   'False
   Icon            =   "credits1_form.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2733.263
   ScaleMode       =   0  'User
   ScaleWidth      =   7507.717
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      Picture         =   "credits1_form.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   3000
      TabIndex        =   1
      Top             =   180
      Width           =   3000
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   585
      Left            =   3420
      Picture         =   "credits1_form.frx":0B4D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3300
      Width           =   1260
   End
   Begin VB.Label Label7 
      Caption         =   "on this time with great love and joy."
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
      Left            =   180
      TabIndex        =   8
      Top             =   2640
      Width           =   3075
   End
   Begin VB.Label Label6 
      Caption         =   "years to come we'll look back"
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
      Left            =   180
      TabIndex        =   7
      Top             =   2340
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   2835
      Left            =   5580
      Picture         =   "credits1_form.frx":0F00
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "on this project and I know in the"
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
      Left            =   180
      TabIndex        =   6
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "great pleasure working together"
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
      Left            =   180
      TabIndex        =   5
      Top             =   1740
      Width           =   2715
   End
   Begin VB.Label Label3 
      Caption         =   "son, Matt Kaprocki. It's been a"
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
      Left            =   180
      TabIndex        =   4
      Top             =   1440
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   "Ed Kaprocki with help from my"
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
      Left            =   180
      TabIndex        =   3
      Top             =   1140
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Designed and programmed by"
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
      Left            =   180
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   2835
      Left            =   3300
      Picture         =   "credits1_form.frx":BA9E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   7437.288
      Y1              =   2194.893
      Y2              =   2194.893
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   7380.945
      Y1              =   2153.48
      Y2              =   2153.48
   End
End
Attribute VB_Name = "credits1_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  main_form.do_debugging (1)
End Sub
