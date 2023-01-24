VERSION 5.00
Begin VB.Form reginfo_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check2Check Registration"
   ClientHeight    =   4485
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7260
   ClipControls    =   0   'False
   Icon            =   "reginfo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "reginfo.frx":08CA
   ScaleHeight     =   3095.626
   ScaleMode       =   0  'User
   ScaleWidth      =   6817.515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1920
      Picture         =   "reginfo.frx":0D0C
      ScaleHeight     =   495
      ScaleWidth      =   3075
      TabIndex        =   8
      Top             =   180
      Width           =   3075
   End
   Begin VB.PictureBox buy_now_pic 
      AutoSize        =   -1  'True
      Height          =   555
      Left            =   5760
      MouseIcon       =   "reginfo.frx":0F8F
      MousePointer    =   99  'Custom
      Picture         =   "reginfo.frx":1299
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   1200
      Width           =   1275
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   60
      Picture         =   "reginfo.frx":1E50
      ScaleHeight     =   1264.2
      ScaleMode       =   0  'User
      ScaleWidth      =   1053.5
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   525
      Left            =   3000
      Picture         =   "reginfo.frx":26CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"reginfo.frx":2A7D
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   300
      TabIndex        =   7
      Top             =   1920
      Width           =   6675
      WordWrap        =   -1  'True
   End
   Begin VB.Label email_label 
      Alignment       =   2  'Center
      Caption         =   "support@mycheck2check.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1560
      MouseIcon       =   "reginfo.frx":2BD7
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label www_label 
      Alignment       =   2  'Center
      Caption         =   "www.mycheck2check.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1560
      MouseIcon       =   "reginfo.frx":2EE1
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   900
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "$19.95"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   5490
      TabIndex        =   4
      Top             =   540
      Width           =   1635
   End
   Begin VB.Label version_label 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1980
      TabIndex        =   2
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   788.803
      X2              =   6013.687
      Y1              =   2609.022
      Y2              =   2609.022
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   774.718
      X2              =   5985.516
      Y1              =   2567.61
      Y2              =   2567.61
   End
End
Attribute VB_Name = "reginfo_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  version_label.Caption = "Version " + Format(major_version) + "." + Format(minor_version) + " - " + version_date_s
End Sub

Private Sub buy_now_pic_Click()
  Call Navigate(Me, "http://www.mycheck2check.com/buynow.htm")
End Sub

Private Sub www_label_Click()
  Call Navigate(Me, "http://www.mycheck2check.com")
End Sub

Private Sub email_label_Click()
  Call Navigate(Me, "mailto:support@mycheck2check.com")
End Sub


