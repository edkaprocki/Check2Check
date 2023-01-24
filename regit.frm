VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Check2Check Registration"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2640
      TabIndex        =   6
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   435
      Left            =   660
      TabIndex        =   5
      Top             =   1440
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   960
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   3435
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Check 2 Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   540
      TabIndex        =   2
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

