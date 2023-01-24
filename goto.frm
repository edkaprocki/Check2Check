VERSION 5.00
Begin VB.Form goto_month_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Go To Month / Year"
   ClientHeight    =   1245
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3420
   Icon            =   "goto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton TodayButton 
      Caption         =   "Today"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   850
   End
   Begin VB.ComboBox month_combo 
      Height          =   315
      ItemData        =   "goto.frx":08CA
      Left            =   240
      List            =   "goto.frx":08F2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox year_combo 
      Height          =   315
      ItemData        =   "goto.frx":0958
      Left            =   1800
      List            =   "goto.frx":095A
      TabIndex        =   2
      Text            =   "year_combo"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   850
   End
   Begin VB.CommandButton OKButton 
      Height          =   495
      Left            =   2400
      Picture         =   "goto.frx":095C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   845
   End
End
Attribute VB_Name = "goto_month_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public ok As Boolean
Dim last_good_year_s As String



Private Sub CancelButton_Click()
  ok = False
  year_combo.Text = last_good_year_s
  Hide
End Sub

Private Sub Form_GotFocus()
  ok = False
  year_combo.Text = last_good_year_s
End Sub

Private Sub Form_Load()
Dim y As Integer
Dim list_start_year As Integer

' Fill up the year combo box
' This is only done once at form load
' Do it this way instead of manually entering the years in the list properties
' Also make the years descending because when it's clicked on we usually want
' to go to earlier years and the list shows this better
' The list of years can simply be changed here without any other concerns
  For y = 2051 To 1990 Step -1
    year_combo.AddItem (Str(y))
  Next y
  
  ' Get the year that is first in the list
  ' Note that it's populated from high to low from above
  list_start_year = CInt(year_combo.List(0))

  ' Now set the year combo box index by taking the first year and subtracting off the current year
  year_combo.ListIndex = list_start_year - Val(Format(Now, "yyyy"))

  ' Now set the month combo box index to the month that we are currently at
  month_combo.ListIndex = Val(Format(Now, "mm")) - 1   ' Subtract 1 since the months range from 1-12
  
  ' Save the text of the current year_combo in case we need to restore it due to
  ' the user typing in an error
  last_good_year_s = year_combo.Text
  
End Sub

Private Sub OKButton_Click()
' Validate the year number since it could have been typed in error or out of range
Dim y As Integer

  On Error GoTo error_h
  y = CInt(year_combo.Text)
  If ((y >= 1800) And (y <= 2200)) Then  ' Arbitrarily pick 2 years that we must be between
    last_good_year_s = year_combo.Text
    ok = True
    Hide
    Exit Sub  ' Everything is good and the year is within limits so exit
  End If

error_h:
  ' We have an error so put up an error message and don't close this form
  MsgBox words(INVALID_NUMBER_ENTERED_N)
  year_combo.SetFocus
End Sub

Private Sub TodayButton_Click()
Dim list_start_year As Integer

  ' Set the month_combo box
  month_combo.ListIndex = Val(Format(Now, "mm") - 1)  ' ListIndex starts at 0 but month starts at 1 so subtract 1
  
  ' Get the year that is first in the list
  ' Note that it's populated from high to low in formload above
  list_start_year = CInt(year_combo.List(0))

  ' Now set the year combo box index by taking the first year and subtracting off the current year
  year_combo.ListIndex = list_start_year - Val(Format(Now, "yyyy"))
  
  ' Act as if we hit the OK button with the current date
  OKButton_Click
End Sub
