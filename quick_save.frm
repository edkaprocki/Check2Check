VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form quick_save_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quick Save"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame amount_frame 
      Caption         =   "Amount"
      Height          =   735
      Left            =   180
      TabIndex        =   5
      Top             =   120
      Width           =   1395
      Begin VB.TextBox amount_box 
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   300
         Width           =   1035
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2595
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4577
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   4800
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Height          =   555
      Left            =   6240
      Picture         =   "quick_save.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton pending_radio 
      Caption         =   "Pending"
      Height          =   195
      Left            =   2340
      TabIndex        =   1
      Top             =   540
      Width           =   1035
   End
   Begin VB.OptionButton done_radio 
      Caption         =   "Done"
      Height          =   195
      Left            =   2340
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "quick_save_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DATE_COL = 1
Const NAME_COL = 2
Const NEEDED_COL = 3
Const AMOUNT_COL = 4

Public due_date_s As String
Public name_s As String
Public needed As Double
Public amount As Double  ' Amount to save or deposit
Dim ok_pressed As Boolean


Private Sub CancelButton_Click()
  Hide
End Sub

Private Sub OKButton_Click()
  Dim v As Double
  
  On Error GoTo error_h
  v = amount_box.Text
  ok_pressed = True
  Hide
  Exit Sub
  
error_h:
  MsgBox "Invalid amount"
End Sub

Public Function execute(index As Integer) As Boolean
  Dim n As Integer
  
  ' i = 1 = Quick Save
  ' i = 2 = Quick Deposit
  
  ok_pressed = False
  execute = False
  update_display
  
  quick_save_form.show vbModal
  If (ok_pressed) Then
    If (index = 1) Then  ' Doing a save
      name_s = "QS" + Format(grid.row, "00:") + grid.TextMatrix(grid.row, 2)
      amount = -Abs(Val(amount_box.Text))  ' Make negative for deposit
    End If
    If (index = 2) Then  ' Doing a deposit
      n = Val(grid.TextMatrix(grid.row, 0))
      name_s = "QD" + Format(n, "00:") + grid.TextMatrix(grid.row, 2)
      amount = Abs(Val(amount_box.Text))  ' Make negative for deposit
    End If
    
    execute = True
  End If
End Function

Private Sub update_display()
  Dim i
  
  quick_save_form.Caption = words(QUICK_SAVE_N)
  CancelButton.Caption = words(CANCEL_N)
  
  With grid
  .ColAlignment(0) = flexAlignCenterCenter
  .ColAlignment(1) = flexAlignLeftCenter
  .ColAlignment(2) = flexAlignLeftCenter
  .ColAlignment(3) = flexAlignRightCenter
  .ColAlignment(4) = flexAlignRightCenter
  
  .ColWidth(0) = 300
  .ColWidth(1) = 900
  .ColWidth(2) = quick_save_form.width - 3650
  .ColWidth(3) = 900
  .ColWidth(4) = 900
  
  .TextMatrix(0, 1) = words(DATE_N)  '"Date"
  .TextMatrix(0, 2) = words(NAME_N)  '"Name"
  .TextMatrix(0, 3) = words(AMOUNT_NEEDED_N)   '"Needed"
  .TextMatrix(0, 4) = words(AMOUNT_SAVED_N)  '"Saved"
  
  ' Fill up the list box
  .Rows = 1
   For i = 1 To MAX_QUICK_ACCOUNT
     If (QUICK_ACCOUNTS.account(i).date <> "") And (QUICK_ACCOUNTS.account(i).name <> "") Then
      .Rows = .Rows + 1
      .TextMatrix(.Rows - 1, 0) = i
      .TextMatrix(.Rows - 1, 1) = QUICK_ACCOUNTS.account(i).date
      .TextMatrix(.Rows - 1, 2) = QUICK_ACCOUNTS.account(i).name
      .TextMatrix(.Rows - 1, 3) = currency_s(QUICK_ACCOUNTS.account(i).needed)
      .TextMatrix(.Rows - 1, 4) = currency_s(QUICK_ACCOUNTS.account(i).total)
      End If
   Next i
  
  End With  ' grid
End Sub
