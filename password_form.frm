VERSION 5.00
Begin VB.Form password_form 
   Caption         =   "Enter Password"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "password_form.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ok_button 
      Height          =   555
      Left            =   3240
      Picture         =   "password_form.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cancel_button 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   3240
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame bottom_frame 
      Caption         =   "Enter new password twice"
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   1620
      Width           =   2775
      Begin VB.TextBox password2_box 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "################"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   300
         PasswordChar    =   "*"
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox password1_box 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "################"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   300
         PasswordChar    =   "*"
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame top_frame 
      Caption         =   "Password"
      Height          =   1395
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2835
      Begin VB.TextBox password_box 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "################"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "password_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim entry_mode As Integer
Public password As String
Dim ok As Boolean


Private Sub update_language()
  top_frame.Caption = words(PASSWORD_N)
  bottom_frame.Caption = words(ENTER_NEW_PASSWORD_N)
  password_form.Caption = words(ENTER_PASSWORD_N)
  cancel_button.Caption = words(CANCEL_N)
End Sub


Public Function execute(n As Integer) As Boolean
  ok = False
  entry_mode = n
  If (n = 0) Then
    top_frame.Visible = True
    bottom_frame.Visible = False
  End If
  
  If (n = 1) Then
    top_frame.Visible = False
    bottom_frame.Visible = True
  End If
  
  update_language
  
  show vbModal
  execute = ok
End Function


Private Sub cancel_button_Click()
  password = ""
  password_box.Text = ""
  password1_box.Text = ""
  password2_box.Text = ""
  Hide
End Sub

Private Sub Form_Activate()
  password = ""
  password_box.Text = ""
  password1_box.Text = ""
  password2_box.Text = ""
End Sub

Private Sub ok_button_Click()
  If (entry_mode = 0) Then
    password = flip_password(UCase(password_box.Text))
    ok = True
    Hide
  End If
  
  If (entry_mode = 1) Then
    If (UCase(password1_box.Text)) = (UCase(password2_box.Text)) Then
      data.password = flip_password(UCase(password1_box.Text))
      password = data.password
      changed_flag = True
      main_form.update_caption
      ok = True
      Hide
    Else
      MsgBox words(PASSWORDS_DONT_MATCH_N)
    End If
  End If
  
End Sub

Private Sub password_box_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn) Then
    ok_button_Click
  End If
End Sub

Private Sub password1_box_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn) Then
    password2_box.SetFocus
  End If
End Sub

Private Sub password2_box_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn) Then
    ok_button_Click
  End If
End Sub

Private Function flip_password(s As String) As String
  ' Return the encrypted / decrypted string
  Dim s1 'As String
  Dim i
  Dim c1 As Byte
  
  s1 = ""
  For i = 1 To Len(s)
    c1 = Asc(Mid(s, i, 1))
    c1 = c1 Xor &HFF
    s1 = s1 + Chr(c1)
  Next i
  flip_password = s1
End Function
