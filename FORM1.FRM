VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   3060
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Width           =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Database Maintenance"
      Height          =   375
      Left            =   5940
      TabIndex        =   129
      Top             =   60
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transactions"
      Height          =   375
      Left            =   840
      TabIndex        =   128
      Top             =   60
      Width           =   1875
   End
   Begin VB.CheckBox override_check 
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   127
      Top             =   1560
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   41
      Left            =   7320
      TabIndex        =   126
      Text            =   "15"
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   41
      Left            =   6840
      TabIndex        =   125
      Text            =   "$1000.25"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   41
      Left            =   6840
      TabIndex        =   124
      Text            =   "$1000.25"
      Top             =   5220
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   40
      Left            =   6300
      TabIndex        =   123
      Text            =   "15"
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   40
      Left            =   5820
      TabIndex        =   122
      Text            =   "$1000.25"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   40
      Left            =   5820
      TabIndex        =   121
      Text            =   "$1000.25"
      Top             =   5220
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   39
      Left            =   5220
      TabIndex        =   120
      Text            =   "15"
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   39
      Left            =   4740
      TabIndex        =   119
      Text            =   "$1000.25"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   39
      Left            =   4740
      TabIndex        =   118
      Text            =   "$1000.25"
      Top             =   5220
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   38
      Left            =   4200
      TabIndex        =   117
      Text            =   "15"
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   38
      Left            =   3720
      TabIndex        =   116
      Text            =   "$1000.25"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   38
      Left            =   3720
      TabIndex        =   115
      Text            =   "$1000.25"
      Top             =   5220
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   37
      Left            =   3120
      TabIndex        =   114
      Text            =   "15"
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   37
      Left            =   2640
      TabIndex        =   113
      Text            =   "$1000.25"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   37
      Left            =   2640
      TabIndex        =   112
      Text            =   "$1000.25"
      Top             =   5220
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   36
      Left            =   2100
      TabIndex        =   111
      Text            =   "15"
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   36
      Left            =   1620
      TabIndex        =   110
      Text            =   "$1000.25"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   36
      Left            =   1620
      TabIndex        =   109
      Text            =   "$1000.25"
      Top             =   5220
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   35
      Left            =   1020
      TabIndex        =   108
      Text            =   "15"
      Top             =   4860
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   35
      Left            =   540
      TabIndex        =   107
      Text            =   "$1000.25"
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   35
      Left            =   540
      TabIndex        =   106
      Text            =   "$1000.25"
      Top             =   5220
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   34
      Left            =   7320
      TabIndex        =   105
      Text            =   "15"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   34
      Left            =   6840
      TabIndex        =   104
      Text            =   "$1000.25"
      Top             =   4380
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   34
      Left            =   6840
      TabIndex        =   103
      Text            =   "$1000.25"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   33
      Left            =   6300
      TabIndex        =   102
      Text            =   "15"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   33
      Left            =   5820
      TabIndex        =   101
      Text            =   "$1000.25"
      Top             =   4380
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   33
      Left            =   5820
      TabIndex        =   100
      Text            =   "$1000.25"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   32
      Left            =   5220
      TabIndex        =   99
      Text            =   "15"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   32
      Left            =   4740
      TabIndex        =   98
      Text            =   "$1000.25"
      Top             =   4380
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   32
      Left            =   4740
      TabIndex        =   97
      Text            =   "$1000.25"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   31
      Left            =   4200
      TabIndex        =   96
      Text            =   "15"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   31
      Left            =   3720
      TabIndex        =   95
      Text            =   "$1000.25"
      Top             =   4380
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   31
      Left            =   3720
      TabIndex        =   94
      Text            =   "$1000.25"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   30
      Left            =   3120
      TabIndex        =   93
      Text            =   "15"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   30
      Left            =   2640
      TabIndex        =   92
      Text            =   "$1000.25"
      Top             =   4380
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   30
      Left            =   2640
      TabIndex        =   91
      Text            =   "$1000.25"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   29
      Left            =   2100
      TabIndex        =   90
      Text            =   "15"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   29
      Left            =   1620
      TabIndex        =   89
      Text            =   "$1000.25"
      Top             =   4380
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   29
      Left            =   1620
      TabIndex        =   88
      Text            =   "$1000.25"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   28
      Left            =   1020
      TabIndex        =   87
      Text            =   "15"
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   28
      Left            =   540
      TabIndex        =   86
      Text            =   "$1000.25"
      Top             =   4380
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   28
      Left            =   540
      TabIndex        =   85
      Text            =   "$1000.25"
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   27
      Left            =   7320
      TabIndex        =   84
      Text            =   "15"
      Top             =   3540
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   27
      Left            =   6840
      TabIndex        =   83
      Text            =   "$1000.25"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   27
      Left            =   6840
      TabIndex        =   82
      Text            =   "$1000.25"
      Top             =   3900
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   26
      Left            =   6300
      TabIndex        =   81
      Text            =   "15"
      Top             =   3540
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   26
      Left            =   5820
      TabIndex        =   80
      Text            =   "$1000.25"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   26
      Left            =   5820
      TabIndex        =   79
      Text            =   "$1000.25"
      Top             =   3900
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   25
      Left            =   5220
      TabIndex        =   78
      Text            =   "15"
      Top             =   3540
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   25
      Left            =   4740
      TabIndex        =   77
      Text            =   "$1000.25"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   25
      Left            =   4740
      TabIndex        =   76
      Text            =   "$1000.25"
      Top             =   3900
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   24
      Left            =   4200
      TabIndex        =   75
      Text            =   "15"
      Top             =   3540
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   24
      Left            =   3720
      TabIndex        =   74
      Text            =   "$1000.25"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   24
      Left            =   3720
      TabIndex        =   73
      Text            =   "$1000.25"
      Top             =   3900
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   23
      Left            =   3120
      TabIndex        =   72
      Text            =   "15"
      Top             =   3540
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   23
      Left            =   2640
      TabIndex        =   71
      Text            =   "$1000.25"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   23
      Left            =   2640
      TabIndex        =   70
      Text            =   "$1000.25"
      Top             =   3900
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   22
      Left            =   2100
      TabIndex        =   69
      Text            =   "15"
      Top             =   3540
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   22
      Left            =   1620
      TabIndex        =   68
      Text            =   "$1000.25"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   22
      Left            =   1620
      TabIndex        =   67
      Text            =   "$1000.25"
      Top             =   3900
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   21
      Left            =   1020
      TabIndex        =   66
      Text            =   "15"
      Top             =   3540
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   21
      Left            =   540
      TabIndex        =   65
      Text            =   "$1000.25"
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   21
      Left            =   540
      TabIndex        =   64
      Text            =   "$1000.25"
      Top             =   3900
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   20
      Left            =   7320
      TabIndex        =   63
      Text            =   "15"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   20
      Left            =   6840
      TabIndex        =   62
      Text            =   "$1000.25"
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   20
      Left            =   6840
      TabIndex        =   61
      Text            =   "$1000.25"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   19
      Left            =   6300
      TabIndex        =   60
      Text            =   "15"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   19
      Left            =   5820
      TabIndex        =   59
      Text            =   "$1000.25"
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   19
      Left            =   5820
      TabIndex        =   58
      Text            =   "$1000.25"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   18
      Left            =   5220
      TabIndex        =   57
      Text            =   "15"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   18
      Left            =   4740
      TabIndex        =   56
      Text            =   "$1000.25"
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   18
      Left            =   4740
      TabIndex        =   55
      Text            =   "$1000.25"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   17
      Left            =   4200
      TabIndex        =   54
      Text            =   "15"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   17
      Left            =   3720
      TabIndex        =   53
      Text            =   "$1000.25"
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   17
      Left            =   3720
      TabIndex        =   52
      Text            =   "$1000.25"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   16
      Left            =   3120
      TabIndex        =   51
      Text            =   "15"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   16
      Left            =   2640
      TabIndex        =   50
      Text            =   "$1000.25"
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   16
      Left            =   2640
      TabIndex        =   49
      Text            =   "$1000.25"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   15
      Left            =   2100
      TabIndex        =   48
      Text            =   "15"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   15
      Left            =   1620
      TabIndex        =   47
      Text            =   "$1000.25"
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   15
      Left            =   1620
      TabIndex        =   46
      Text            =   "$1000.25"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   14
      Left            =   1020
      TabIndex        =   45
      Text            =   "15"
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   14
      Left            =   540
      TabIndex        =   44
      Text            =   "$1000.25"
      Top             =   3060
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   14
      Left            =   540
      TabIndex        =   43
      Text            =   "$1000.25"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   13
      Left            =   7320
      TabIndex        =   42
      Text            =   "15"
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   13
      Left            =   6840
      TabIndex        =   41
      Text            =   "$1000.25"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   13
      Left            =   6840
      TabIndex        =   40
      Text            =   "$1000.25"
      Top             =   2580
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   12
      Left            =   6300
      TabIndex        =   39
      Text            =   "15"
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   12
      Left            =   5820
      TabIndex        =   38
      Text            =   "$1000.25"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   12
      Left            =   5820
      TabIndex        =   37
      Text            =   "$1000.25"
      Top             =   2580
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   11
      Left            =   5220
      TabIndex        =   36
      Text            =   "15"
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   11
      Left            =   4740
      TabIndex        =   35
      Text            =   "$1000.25"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   11
      Left            =   4740
      TabIndex        =   34
      Text            =   "$1000.25"
      Top             =   2580
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   10
      Left            =   4200
      TabIndex        =   33
      Text            =   "15"
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   10
      Left            =   3720
      TabIndex        =   32
      Text            =   "$1000.25"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   10
      Left            =   3720
      TabIndex        =   31
      Text            =   "$1000.25"
      Top             =   2580
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   9
      Left            =   3120
      TabIndex        =   30
      Text            =   "15"
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   9
      Left            =   2640
      TabIndex        =   29
      Text            =   "$1000.25"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   9
      Left            =   2640
      TabIndex        =   28
      Text            =   "$1000.25"
      Top             =   2580
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   8
      Left            =   2100
      TabIndex        =   27
      Text            =   "15"
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   8
      Left            =   1620
      TabIndex        =   26
      Text            =   "$1000.25"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   8
      Left            =   1620
      TabIndex        =   25
      Text            =   "$1000.25"
      Top             =   2580
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   7
      Left            =   1020
      TabIndex        =   24
      Text            =   "15"
      Top             =   2220
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   7
      Left            =   540
      TabIndex        =   23
      Text            =   "$1000.25"
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   7
      Left            =   540
      TabIndex        =   22
      Text            =   "$1000.25"
      Top             =   2580
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   6
      Left            =   7320
      TabIndex        =   21
      Text            =   "15"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   6
      Left            =   6840
      TabIndex        =   20
      Text            =   "$1000.25"
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   6
      Left            =   6840
      TabIndex        =   19
      Text            =   "$1000.25"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   5
      Left            =   6300
      TabIndex        =   18
      Text            =   "15"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   5
      Left            =   5820
      TabIndex        =   17
      Text            =   "$1000.25"
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   5
      Left            =   5820
      TabIndex        =   16
      Text            =   "$1000.25"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   4
      Left            =   5220
      TabIndex        =   15
      Text            =   "15"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   4
      Left            =   4740
      TabIndex        =   14
      Text            =   "$1000.25"
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   4
      Left            =   4740
      TabIndex        =   13
      Text            =   "$1000.25"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   3
      Left            =   4200
      TabIndex        =   12
      Text            =   "15"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   3
      Left            =   3720
      TabIndex        =   11
      Text            =   "$1000.25"
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   3
      Left            =   3720
      TabIndex        =   10
      Text            =   "$1000.25"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   2
      Left            =   3120
      TabIndex        =   9
      Text            =   "15"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   2
      Left            =   2640
      TabIndex        =   8
      Text            =   "$1000.25"
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   2
      Left            =   2640
      TabIndex        =   7
      Text            =   "$1000.25"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   1
      Left            =   2100
      TabIndex        =   6
      Text            =   "15"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Text            =   "$1000.25"
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   1
      Left            =   1620
      TabIndex        =   4
      Text            =   "$1000.25"
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox line1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   0
      Left            =   1020
      TabIndex        =   3
      Text            =   "15"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox line2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   0
      Left            =   540
      TabIndex        =   2
      Text            =   "$1000.25"
      Top             =   1740
      Width           =   855
   End
   Begin VB.TextBox line3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   165
      Index           =   0
      Left            =   540
      TabIndex        =   1
      Text            =   "$1000.25"
      Top             =   1920
      Width           =   855
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   7695
      _Version        =   524288
      _ExtentX        =   13573
      _ExtentY        =   8916
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   1998
      Month           =   11
      Day             =   28
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   16711680
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   16711680
      GridLinesColor  =   -2147483633
      ShowDateSelectors=   -1  'True
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  transaction_form.Show
End Sub

Private Sub Command2_Click()
  frmtransaction_table.Show
End Sub
