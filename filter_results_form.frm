VERSION 5.00
Begin VB.Form filter_results_form 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtered Results"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3585
   Icon            =   "filter_results_form.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox filtered_in_box 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox filtered_out_box 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   900
      Width           =   1695
   End
   Begin VB.TextBox total_box 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1740
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   60
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Displayed Qty"
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
      TabIndex        =   5
      Top             =   540
      Width           =   1395
   End
   Begin VB.Label label21 
      Alignment       =   1  'Right Justify
      Caption         =   "Hidden Qty"
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
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Displayed Total"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "filter_results_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyEscape) Or (KeyCode = vbKeyReturn) Then
    filter_results_form.Hide
  End If
End Sub

Private Sub Form_Load()
  ' Set up the form
  Dim lR As Long
  lR = SetTopMostWindow(filter_results_form.hwnd, True)

End Sub

Public Sub update_filter_results_display()
  Dim lR As Long
  lR = SetTopMostWindow(filter_results_form.hwnd, True)
End Sub

Public Sub normal()
  Dim lR As Long
  lR = SetTopMostWindow(filter_results_form.hwnd, True)
  
  filter_results_form.WindowState = 0  ' Normal state
End Sub

