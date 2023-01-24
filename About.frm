VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Check 2 Check"
   ClientHeight    =   4365
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5790
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3012.801
   ScaleMode       =   0  'User
   ScaleWidth      =   5437.109
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1305
      Picture         =   "About.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   3000
      TabIndex        =   3
      Top             =   90
      Width           =   3000
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1800
      Left            =   135
      Picture         =   "About.frx":0B4D
      ScaleHeight     =   1264.2
      ScaleMode       =   0  'User
      ScaleWidth      =   1053.5
      TabIndex        =   1
      Top             =   765
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   585
      Left            =   2295
      Picture         =   "About.frx":13C7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label ct_label 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label total_notes_label 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   13
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label next_check_label 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label total_transactions_label 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label last_transaction_date_label 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label first_transaction_date_label 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label www_label 
      Alignment       =   2  'Center
      Caption         =   "www.mycheck2check.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1710
      MouseIcon       =   "About.frx":177A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label email_label 
      Alignment       =   2  'Center
      Caption         =   "support@mycheck2check.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1755
      MouseIcon       =   "About.frx":1A84
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label credits2_label 
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   3555
      Width           =   315
   End
   Begin VB.Label credits1_label 
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   3555
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Copyright © 2000-2019 QuickSoft, Inc."
      Height          =   255
      Left            =   1305
      TabIndex        =   4
      Top             =   3135
      Width           =   3315
   End
   Begin VB.Label version_label 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   375
      TabIndex        =   2
      Top             =   2745
      Width           =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.569
      Y1              =   2412.312
      Y2              =   2412.312
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   2370.898
      Y2              =   2370.898
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub credits1_label_Click()
  credits1_form.show vbModal
End Sub

Private Sub credits2_label_Click()
  credits2_form.show vbModal
End Sub

Private Sub email_label_Click()
  Call Navigate(Me, "mailto:support@check2check.com")
End Sub

Private Sub Form_Activate()
  Dim i As Integer
  Dim n As Integer
  
  first_transaction_date_label.Caption = words(FIRST_TRANSACTION_N) + " - " + get_date(db(data.first).Month, db(data.first).day, db(data.first).Year)
  last_transaction_date_label.Caption = words(LAST_TRANSACTION_N) + " - " + get_date(db(data.last).Month, db(data.last).day, db(data.last).Year)
  total_transactions_label.Caption = words(TOTAL_TRANSACTIONS_N) + " - " + Format(data.number_of_records)
  total_notes_label.Caption = words(TOTAL_NOTES_N) + " - " + Format(data.number_of_notes)
  next_check_label.Caption = words(NEXT_CHECK_NUMBER_N) + " - " + Format(data.last_check_number + 1)

  ' Add up all the cardtrak records
  n = 0
  For i = 1 To MAX_CARD_TRANSACTIONS
    ' Loop through all the card transactions and if they don't point to an active main then delete it
    If (cards(i).active) Then n = n + 1
  Next i
  ct_label.Caption = words(TOTAL_CARDTRAKS_N) + " - " + Format(n)
  
End Sub

Private Sub Form_Load()
  version_label.Caption = words(VERSION_N) + " " + Format(major_version) + "." + Format(minor_version) + " - " + version_date_s
End Sub

Private Sub www_label_Click()
  Call Navigate(Me, "http://www.mycheck2check.com")
End Sub
