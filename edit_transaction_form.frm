VERSION 5.00
Begin VB.Form edit_transaction_form 
   Caption         =   "Edit Transaction"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   Icon            =   "edit_transaction_form.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame tags_frame 
      Caption         =   "Tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   6840
      TabIndex        =   23
      Top             =   1395
      Width           =   1005
      Begin VB.CheckBox tag_check 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   27
         Top             =   1200
         Width           =   555
      End
      Begin VB.CheckBox tag_check 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   26
         Top             =   900
         Width           =   555
      End
      Begin VB.CheckBox tag_check 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   25
         Top             =   600
         Width           =   555
      End
      Begin VB.CheckBox tag_check 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   24
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame exclude_frame 
      Caption         =   "Exclude"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   5715
      TabIndex        =   20
      Top             =   1395
      Width           =   1050
      Begin VB.OptionButton exclude_radio 
         Caption         =   "Yes"
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
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   540
         Width           =   735
      End
      Begin VB.OptionButton exclude_radio 
         Caption         =   "No"
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
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame cleared_frame 
      Caption         =   "Cleared"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   2475
      TabIndex        =   13
      Top             =   1395
      Width           =   1455
      Begin VB.OptionButton cleared_radio 
         Caption         =   "No"
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
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton cleared_radio 
         Caption         =   "Yes"
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
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   540
         Width           =   795
      End
      Begin VB.OptionButton cleared_radio 
         Caption         =   "Finished"
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
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   840
         Width           =   1170
      End
   End
   Begin VB.Frame status_frame 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   4005
      TabIndex        =   8
      Top             =   1395
      Width           =   1635
      Begin VB.OptionButton status_radio 
         Caption         =   "Skip"
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
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1140
         Width           =   1140
      End
      Begin VB.OptionButton status_radio 
         Caption         =   "Pending"
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
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton status_radio 
         Caption         =   "Done"
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
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   540
         Width           =   1245
      End
      Begin VB.OptionButton status_radio 
         Caption         =   "Blank"
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
         Index           =   0
         Left            =   225
         TabIndex        =   9
         Top             =   225
         Width           =   1305
      End
   End
   Begin VB.Frame amount_frame 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   6
      Top             =   1395
      Width           =   2355
      Begin VB.TextBox amount_box 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Frame name_frame 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   6420
      Begin VB.TextBox name_box 
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   225
         Width           =   6045
      End
   End
   Begin VB.Frame check_number_frame 
      Caption         =   "Check Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   2
      Top             =   2385
      Width           =   2310
      Begin VB.TextBox number_box 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1860
      End
   End
   Begin VB.CommandButton OKButton 
      Height          =   555
      Left            =   6615
      Picture         =   "edit_transaction_form.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   6615
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label check_label 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   19
      Top             =   225
      Width           =   2505
   End
   Begin VB.Label amount_label 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1890
      TabIndex        =   18
      Top             =   225
      Width           =   2055
   End
   Begin VB.Label date_label 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   17
      Top             =   225
      Width           =   1770
   End
End
Attribute VB_Name = "edit_transaction_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private first_pass As Boolean  ' First time in this routine
Private ok_pressed As Boolean  ' We have hit ok to exit
Private check_number_enter_counter  ' Counts the number of enters hit for the check number box


Private Sub update_language()
  name_frame.Caption = words(NAME_N)
  amount_frame.Caption = words(AMOUNT_N)
  check_number_frame.Caption = words(CHECK_NUMBER_N)
  cleared_frame.Caption = words(CLEARED_N)
  status_frame.Caption = words(STATUS_N)
  exclude_frame.Caption = words(EXCLUDE_N)
  tags_frame.Caption = words(TAGS_N)
  
  status_radio(0).Caption = words(BLANK_N)
  status_radio(1).Caption = words(DONE_N)
  status_radio(2).Caption = words(PENDING_N)
  status_radio(3).Caption = words(SKIP_N)
  
  cleared_radio(0).Caption = words(NO_N)
  cleared_radio(1).Caption = words(YES_N)
  cleared_radio(2).Caption = words(FINISHED_N)
  
  exclude_radio(0).Caption = words(NO_N)
  exclude_radio(1).Caption = words(YES_N)
  
  CancelButton.Caption = words(CANCEL_N)
End Sub


Private Sub amount_box_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn) Then
    'this.amount = amount_box
    'update_display
    amount_box_LostFocus
  End If
End Sub

Private Sub amount_box_LostFocus()
  ' Update the amount
  On Error GoTo error_h
  If (amount_box.Text <> "") Then
    this.amount = amount_box
  End If
  update_display
  Exit Sub
error_h:
  MsgBox words(INVALID_NUMBER_ENTERED_N)  '"Invalid number entered"
  amount_box.SetFocus
End Sub

Private Sub CancelButton_Click()
  Hide
End Sub

Private Sub cleared_radio_Click(index As Integer)
  If (first_pass) Then Exit Sub
  
  If (index = 2) Or (this.cleared = 2) Then
    ' We are attempting to change a finished transaction
    If (MsgBox(words(ARE_YOU_SURE_Q_N), vbYesNo + vbQuestion, words(CHANGE_FINISHED_STATUS_Q_N)) = vbYes) Then
      ' Put it back the way it was
      this.cleared = index
      update_display
      Exit Sub
    End If
  Else
    this.cleared = index
  End If
  update_display
End Sub

Private Sub exclude_radio_Click(index As Integer)
  If (index = 0) Then
    this.exclude = False
  Else
    this.exclude = True
  End If
End Sub

Private Sub name_box_LostFocus()
  this.name = name_box.Text
End Sub

Private Sub number_box_DblClick()
  If (number_box.Text = "") Then
    data.last_check_number = data.last_check_number + 1
    this.check = data.last_check_number
    If (preferences.auto_check_done_on_check) Then
      this.paid = 1
      update_display
    End If
  End If
End Sub

Private Sub number_box_GotFocus()
  check_number_enter_counter = 0
End Sub

Private Sub number_box_KeyDown(KeyCode As Integer, Shift As Integer)
  If (number_box.Text = "") And (KeyCode = vbKeyReturn) Then
    check_number_enter_counter = check_number_enter_counter + 1
    If (check_number_enter_counter > 1) Then
        check_number_enter_counter = 0
        data.last_check_number = data.last_check_number + 1
        this.check = data.last_check_number
        If (preferences.auto_check_done_on_check) Then
          this.paid = 1
          update_display
        End If
    End If
  End If
  
  If (KeyCode = vbKeyReturn) Then
    number_box_LostFocus
    update_display
  End If
End Sub

Private Sub number_box_LostFocus()
  On Error GoTo error_h
  
  If (number_box.Text <> "") Then
    this.check = number_box.Text
  Else
    this.check = -1
  End If
  update_display
  Exit Sub
error_h:
  MsgBox words(INVALID_NUMBER_ENTERED_N)
  number_box.SetFocus
End Sub

Private Sub OKButton_Click()
  ok_pressed = True
  Hide
End Sub

Public Function execute() As Boolean
  ok_pressed = False
  first_pass = True
  execute = False
  date_label.Caption = words(DATE_N) + ": " + get_date(this.month, this.day, this.year)
  amount_label.Caption = words(AMOUNT_N) + ": " + currency_s(this.amount)
  amount_label.ForeColor = amount_color(this.amount)
  If (this.check >= 0) Then
    check_label.Caption = words(CHECK_NUMBER_N) + ": " + Str(this.check)
  Else
    check_label.Caption = "Check: "
  End If
  
  update_language
  
  update_display
  first_pass = False
  
  edit_transaction_form.show vbModal
  If (ok_pressed) Then execute = True
End Function

Private Sub update_display()
  name_box.Text = this.name
  amount_box.Text = currency_s(this.amount)
  status_radio(this.paid).Value = True
  cleared_radio(this.cleared).Value = True
  If (this.check >= 0) Then
    number_box.Text = this.check
  Else
    number_box.Text = ""
  End If
  
  If (this.exclude) Then
    exclude_radio(1).Value = True
  Else
    exclude_radio(0).Value = True
  End If
  
  
  tag_check(0).Value = vbUnchecked
  tag_check(1).Value = vbUnchecked
  tag_check(2).Value = vbUnchecked
  tag_check(3).Value = vbUnchecked
  If ((this.tags And 1) <> 0) Then tag_check(0).Value = vbChecked
  If ((this.tags And 2) <> 0) Then tag_check(1).Value = vbChecked
  If ((this.tags And 4) <> 0) Then tag_check(2).Value = vbChecked
  If ((this.tags And 8) <> 0) Then tag_check(3).Value = vbChecked
  
  amount_box.ForeColor = amount_color(this.amount)
End Sub

Private Sub status_radio_LostFocus(index As Integer)
  this.paid = index
  update_display
End Sub

Private Sub tag_check_LostFocus(index As Integer)
  this.tags = this.tags And &HF0
  If (tag_check(0).Value = vbChecked) Then this.tags = this.tags Or 1
  If (tag_check(1).Value = vbChecked) Then this.tags = this.tags Or 2
  If (tag_check(2).Value = vbChecked) Then this.tags = this.tags Or 4
  If (tag_check(3).Value = vbChecked) Then this.tags = this.tags Or 8
  update_display
End Sub
