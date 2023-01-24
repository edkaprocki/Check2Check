VERSION 5.00
Begin VB.Form splash_form 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   Icon            =   "splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   4575
      Left            =   0
      Picture         =   "splash.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7635
   End
End
Attribute VB_Name = "splash_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
  If (splash_form.Visible) Then splash_form.Hide
  'Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  Hide
End Sub

Private Sub Form_Load()
  ' Set up the splash form
  Dim lR As Long
  lR = SetTopMostWindow(splash_form.hwnd, True)
End Sub

Private Sub Image1_Click()
  Hide
End Sub
