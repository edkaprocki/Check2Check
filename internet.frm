VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form internet_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check2Check"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "internet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_ok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btn_no 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      TabIndex        =   2
      Top             =   780
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btn_yes 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   900
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   735
   End
   Begin InetCtlsObjects.Inet internet1 
      Left            =   2430
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   -45
      TabIndex        =   4
      Top             =   630
      Visible         =   0   'False
      Width           =   5820
      WordWrap        =   -1  'True
   End
   Begin VB.Label internet_label 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "internet_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As String  ' String to send back
  

Private Sub btn_no_Click()
  Unload Me
End Sub

Private Sub btn_ok_Click()
  Unload Me
End Sub

Private Sub btn_yes_Click()
  internet_label.Caption = "Loading..."
  Call Navigate(Me, "http://www.mycheck2check.com")
  Unload Me
End Sub

Private Sub Form_Activate()
  Dim version, web_version, goto_url As String
  On Error GoTo errhandler
  version = major_version & "." & minor_version
  
  setup_stuff
  
  btn_yes.Visible = False
  btn_no.Visible = False
  btn_ok.Visible = False
  
  internet_label.Caption = "Connecting..."
  internet1.AccessType = icUseDefault
  web_version = internet1.OpenURL("http://www.mycheck2check.com/cgi-bin/users.cgi?data=" + s, 0)
  web_version = internet1.OpenURL("http://www.mycheck2check.com/version.txt", 0)
  If web_version = "" Then GoTo errhandler
  If Str(web_version) > Str(version) Then
    internet_label.Caption = "Newer version is available: " + web_version + " Go to website now?"
    btn_yes.Visible = True
    btn_no.Visible = True
  Else
    internet_label.Caption = "No newer versions are available!"
    btn_ok.Visible = True
  End If
  Exit Sub
  
errhandler:
  If Error = "" Then
    internet_label.Caption = "Could not retrieve file!  Check your internet connection and try again."
  Else: internet_label.Caption = Error
  End If
  btn_ok.Visible = True
End Sub

Private Sub setup_stuff()
  s = date
  s = s + "," + register_form.get_registration_date_s  ' Install date
  s = s + "," + GetSetting("Microsoft", "QSIEKR", "QSIEKR", "Exclude")  ' Valid registration
  s = s + "," + GetSetting("Check 2 Check", "Settings", "Regcode", "00000000000000000000")  ' Registration code
  s = s + "," + GetSetting("Check 2 Check", "Settings", "Startups", "0")  ' Number of startups
  s = s + "," + Format(data.number_of_records)
  s = s + "," + Format(data.number_of_notes)
  s = s + "," + GetSetting("Check 2 Check", "Settings", "Printouts", "0")
  s = s + "," + GetSetting("Check 2 Check", "Settings", "Name", "noname")
  
  Label1.Caption = s
End Sub

