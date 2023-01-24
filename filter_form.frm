VERSION 5.00
Begin VB.Form filter_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filter"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
   Icon            =   "filter_form.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton clear_filter_button 
      Caption         =   "Clear Filter"
      Height          =   555
      Left            =   180
      TabIndex        =   22
      Top             =   2880
      Width           =   1410
   End
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
      Height          =   975
      Left            =   3375
      TabIndex        =   17
      Top             =   1710
      Width           =   1575
      Begin VB.CheckBox tag0_check 
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
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   300
         Width           =   555
      End
      Begin VB.CheckBox tag1_check 
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
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox tag2_check 
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
         Height          =   255
         Left            =   900
         TabIndex        =   19
         Top             =   300
         Width           =   495
      End
      Begin VB.CheckBox tag3_check 
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
         Height          =   255
         Left            =   900
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.OptionButton deposit_radio 
      Caption         =   "Deposit"
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
      Left            =   3630
      TabIndex        =   7
      Top             =   900
      Width           =   1155
   End
   Begin VB.OptionButton withdrawal_radio 
      Caption         =   "Withdrawal"
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
      Left            =   3630
      TabIndex        =   6
      Top             =   600
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.CommandButton ok_button 
      Height          =   555
      Left            =   3555
      Picture         =   "filter_form.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1395
   End
   Begin VB.CommandButton cancel_button 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   1890
      TabIndex        =   5
      Top             =   2880
      Width           =   1410
   End
   Begin VB.TextBox amount_to_box 
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
      Left            =   2475
      TabIndex        =   2
      Top             =   660
      Width           =   975
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
      Height          =   975
      Left            =   180
      TabIndex        =   15
      Top             =   1680
      Width           =   3045
      Begin VB.CheckBox dash_check 
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
         Height          =   255
         Left            =   1620
         TabIndex        =   11
         Top             =   540
         Width           =   1305
      End
      Begin VB.CheckBox question_check 
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
         Height          =   255
         Left            =   1620
         TabIndex        =   10
         Top             =   270
         Width           =   1365
      End
      Begin VB.CheckBox done_check 
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
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   540
         Width           =   1350
      End
      Begin VB.CheckBox blank_check 
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
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   1350
      End
   End
   Begin VB.TextBox check_box 
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
      Left            =   1890
      TabIndex        =   3
      Top             =   1170
      Width           =   975
   End
   Begin VB.TextBox amount_from_box 
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
      Left            =   945
      TabIndex        =   1
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox name_box 
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
      Left            =   945
      TabIndex        =   0
      Top             =   180
      Width           =   3990
   End
   Begin VB.Label to_label 
      Alignment       =   2  'Center
      Caption         =   "to"
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
      Left            =   2070
      TabIndex        =   16
      Top             =   720
      Width           =   195
   End
   Begin VB.Label check_label 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   45
      TabIndex        =   14
      Top             =   1215
      Width           =   1815
   End
   Begin VB.Label amount_label 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   45
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label name_label 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   45
      TabIndex        =   12
      Top             =   240
      Width           =   825
   End
End
Attribute VB_Name = "filter_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim filt As filter_type
Dim valid_fields As Boolean  ' True when all fields are valid
Dim cancel_it As Boolean  ' Set to false when ok was hit



Public Sub get_filter()
  Dim t As Double
  
  On Error GoTo errorh:
  
  valid_fields = True
  filt.name = UCase(name_box.Text)
  filt.amount_to = Val(amount_to_box.Text)
  filt.amount_from = Val(amount_from_box.Text)
  
  If (check_box.Text = "") Then
    filt.check = -1
  Else
    filt.check = Val(check_box.Text)
  End If
  
  ' Read in the amount
  If (amount_from_box.Text <> "") Then
    filt.amount_from = Val(amount_from_box.Text)
    If (amount_to_box.Text <> "") Then
      filt.amount_to = Val(amount_to_box.Text)
    Else
      filt.amount_to = filt.amount_from
    End If
    If (withdrawal_radio.Value) Then
      ' Switch them
      t = filt.amount_to
      filt.amount_to = filt.amount_from * -1#
      filt.amount_from = t * -1#
    End If
  Else
    ' We don't have an amount from checked
    filt.amount_from = -55  ' 55 for both indicates no checking
    filt.amount_to = -55
  End If
  
  filt.status_blank = blank_check.Value
  filt.status_done = done_check.Value
  filt.status_question = question_check.Value
  filt.status_dash = dash_check.Value
  filt.filtered = False
  filt.filtered_out_count = 0
  filt.filtered_in_count = 0
  filt.total_amount = 0
  
  filt.tags(0) = tag0_check.Value
  filt.tags(1) = tag1_check.Value
  filt.tags(2) = tag2_check.Value
  filt.tags(3) = tag3_check.Value
  
  
  filter.filtered = False  ' Set to show we have no filtered transactions
  filter.filtered_out_count = 0
  filter.filtered_in_count = 0
  filter.total_amount = 0
  Exit Sub
  
errorh:
  valid_fields = False
End Sub

Public Function execute() As Boolean
  filter_form.Caption = words(FILTER_N)
  name_label.Caption = words(NAME_N)
  amount_label.Caption = words(AMOUNT_N)
  to_label.Caption = words(TO_N)
  withdrawal_radio.Caption = words(WITHDRAWAL_N)
  deposit_radio.Caption = words(DEPOSIT_N)
  status_frame.Caption = words(STATUS_N)
  tags_frame.Caption = words(TAGS_N)
  blank_check.Caption = words(BLANK_N)
  done_check.Caption = words(DONE_N)
  question_check.Caption = words(PENDING_N)
  dash_check.Caption = words(SKIP_N)
  cancel_button.Caption = words(CANCEL_N)
  clear_filter_button.Caption = words(CLEAR_FILTER_N)
  check_label.Caption = words(CHECK_NUMBER_N)
  'ok_button.Caption = words(OK_N)
  
  filt = filter
  filt.filtered = False  ' Set to show we have no filtered transactions
  filt.filtered_out_count = 0
  filt.filtered_in_count = 0
  filt.total_amount = 0
  name_box.SelStart = 0
  name_box.SelLength = 100
  filt.check = -1
  cancel_it = True
  
  show vbModal
  execute = True
  If cancel_it Then execute = False
End Function

Private Sub amount_from_box_GotFocus()
  amount_from_box.SelStart = 0
  amount_from_box.SelLength = 100
End Sub

Private Sub amount_to_box_GotFocus()
  amount_to_box.SelStart = 0
  amount_to_box.SelLength = 100
End Sub

Private Sub cancel_button_Click()
  Hide
End Sub

Private Sub check_box_GotFocus()
  check_box.SelStart = 0
  check_box.SelLength = 100
End Sub

Private Sub clear_filter_button_Click()
  ' Clear out all the boxes
  name_box.Text = ""
  amount_from_box.Text = ""
  amount_to_box.Text = ""
  check_box.Text = ""
  blank_check.Value = Unchecked
  done_check.Value = Unchecked
  question_check.Value = Unchecked
  dash_check.Value = Unchecked
  tag0_check.Value = Unchecked
  tag1_check.Value = Unchecked
  tag2_check.Value = Unchecked
  tag3_check.Value = Unchecked
End Sub

Private Sub Form_Activate()
  name_box.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyReturn) Then
    ok_button_Click
  Else
    If (KeyCode = vbKeyEscape) Then
      cancel_button_Click
    End If
  End If
End Sub

Private Sub name_box_GotFocus()
  name_box.SelStart = 0
  name_box.SelLength = 100
End Sub

Private Sub ok_button_Click()
  get_filter
  If (valid_fields) Then
    filter = filt
    cancel_it = False
    Hide
  Else
    MsgBox words(INVALID_ENTRY_N)
  End If
End Sub

Public Sub get_filter_parameters()
  get_filter
  filter = filt
  check_filter_active
End Sub

Public Sub check_filter_active()
  filter.active = False
  If (filter.name <> "") Then filter.active = True
  If ((filter.amount_from <> -55) And (filter.amount_to) <> -55) Then filter.active = True
  If (filter.check > -1) Then filter.active = True
  
  filter.status_ignore = True
  If (filter.status_blank Or _
      filter.status_done Or _
      filter.status_question Or _
      filter.status_dash) Then
        
        filter.status_ignore = False
        filter.active = True
  End If
  
  filter.tags_ignore = True
  If (filter.tags(0) Or _
      filter.tags(1) Or _
      filter.tags(2) Or _
      filter.tags(3)) Then
        
        filter.tags_ignore = False
        filter.active = True
  End If
  
End Sub


