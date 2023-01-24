VERSION 5.00
Begin VB.Form register_form 
   Caption         =   "Check2Check Registration"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   Icon            =   "register.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   840
      Picture         =   "register.frx":08CA
      ScaleHeight     =   435
      ScaleWidth      =   3075
      TabIndex        =   9
      Top             =   60
      Width           =   3075
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   120
      Picture         =   "register.frx":0B4D
      ScaleHeight     =   555
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   60
      Width           =   615
   End
   Begin VB.CommandButton info_button 
      Caption         =   "Purchase Information"
      Height          =   435
      Left            =   1680
      TabIndex        =   6
      Top             =   1620
      Width           =   1335
   End
   Begin VB.CommandButton cancel_button 
      Caption         =   "Run Check2Check"
      Height          =   435
      Left            =   3120
      TabIndex        =   5
      Top             =   1620
      Width           =   1335
   End
   Begin VB.CommandButton register_button 
      Caption         =   "Register"
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   1620
      Width           =   1335
   End
   Begin VB.TextBox code_box 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   3435
   End
   Begin VB.TextBox name_box 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   3435
   End
   Begin VB.Label days_left_label 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   540
      Width           =   2595
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
      Left            =   45
      TabIndex        =   3
      Top             =   1200
      Width           =   885
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
      Left            =   45
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "register_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const REGISTRATION_DAYS = 30
Const DAYS_TO_SUBTRACT = 15386
Const KEY = "13762190523907458905"

Dim mydate
Dim cur_date_long As Double
Dim install_date As Double
Dim s As String


Private Sub update_language()
  Label1.Caption = words(NAME_N)
  Label3.Caption = words(CODE_N)
  register_button.Caption = words(REGISTER_N)
  info_button.Caption = words(PURCHASE_INFORMATION_N)
  cancel_button.Caption = words(RUN_CHECK2CHECK_N)
End Sub


Public Function get_registration_date_s() As String
  install_date = Val(GetSetting("Microsoft", "QSIEKR", "QSIEKD", "0"))
  If (install_date > 36600) Then
    get_registration_date_s = CDate(install_date - 15386)
  Else
    get_registration_date_s = "0/0/0"
  End If
End Function

Private Sub cancel_button_Click()
  Unload Me
End Sub


Private Sub info_button_Click()
  ' Show the reginfo form
  reginfo_form.show vbModal
End Sub


Private Sub register_button_Click()
  Dim valid As Boolean
  Dim mydate
  
  valid = False
  
  ' See if stuff has been entered
  If (name_box.Text <> "") And (code_box.Text <> "") Then
    
    ' Perform the algorithm here
    valid = check_code(name_box.Text, KEY, code_box.Text)
     
    If (valid) Then
      ' We have a valid registration code entered
      ' Save the entry in the registry that this program is registered
      SaveSetting "Microsoft", "QSIEKR", "QSIEKR", "QSIREG"
      SaveSetting "Check 2 Check", "Settings", "Regcode", code_box.Text
      SaveSetting "Check 2 Check", "Settings", "Name", name_box.Text
      MsgBox words(SUCCESS_THANK_YOU_FOR_REGISTERING_N)
      Unload Me
    Else
      ' Not a valid name/code combination
      MsgBox words(SORRY_NOT_A_VALID_NAME_CODE_N)
    End If
  Else
    MsgBox words(NOT_ENOUGH_INFORMATION_N)
  End If
  
End Sub


Public Function is_registered() As Boolean
  Dim reg As Boolean
  ' See if this is a valid registered program
  reg = ("QSIREG" = GetSetting("Microsoft", "QSIEKR", "QSIEKR", "EXCLUDE"))
  If (reg = False) Then
    ' Pop up a message box begging for money
    MsgBox "=======================" & vbCrLf & "Thank you for using Check2Check." & vbCrLf & "  This is a fully functional version." & vbCrLf & "      Please consider purchasing." & vbCrLf & "                     Thank you!" & vbCrLf & "======================="
    
 End If

  is_registered = True  ' Always allow this program to run
End Function


Public Function ok_to_run() As Boolean
  ' Return true if ok to run, false if outside of trial period
  
  If (is_registered) Then
    ' We have a fully registered program so proceed normally
    ok_to_run = True
    Exit Function
  End If
  
  ' This sub will log this program as being run the first time
  ' See if it has been run at all
  s = GetSetting("Microsoft", "QSIEKR", "QSIEKD", "NOT FOUND")
  If (s <> "NOT FOUND") Then
    ' We have been run before so see how many days left
    If (get_days_left >= 0) Then
      ' We are still within the run time
      ok_to_run = True
    Else
      ok_to_run = False
    End If
  Else
    ' We have not been run before so set up the parameters and save in registry
    cur_date_long = date
    SaveSetting "Microsoft", "QSIEKR", "QSIEKR", "Exclude"  ' Not registered
    SaveSetting "Microsoft", "QSIEKR", "QSIEKD", Format(cur_date_long + DAYS_TO_SUBTRACT) ' Date first ran
    SaveSetting "Check 2 Check", "Settings", "Regcode", code_box.Text
    SaveSetting "Check 2 Check", "Settings", "Name", name_box.Text
    ok_to_run = True
  End If
  
  get_days_left
  
  update_language
  
  show vbModal  ' Show this form
  
End Function


Public Function ok_to_run_form() As Boolean
  ' See if we are registered
  If (is_registered) Then
    get_days_left
    update_language
    
    show vbModal
    ok_to_run_form = True
    Exit Function
  End If
  
  ok_to_run_form = ok_to_run
  Exit Function
End Function


Public Function get_days_left() As Integer
  mydate = date
  cur_date_long = mydate  ' Get the date
  
  install_date = Val(GetSetting("Microsoft", "QSIEKR", "QSIEKD", "0"))
  If (install_date > 36600) Then
    If (install_date <> 0) Then
      ' We have been run before so see how many days left
      install_date = install_date - DAYS_TO_SUBTRACT
      get_days_left = REGISTRATION_DAYS - (cur_date_long - install_date)
    Else
      get_days_left = REGISTRATION_DAYS
    End If
  Else
    get_days_left = -1  ' Set to show expired
  End If
  
  ' Show the days left
  If (is_registered) Then
    days_left_label.Caption = "Licensed copy"
  Else
    If (get_days_left >= 0) Then
      days_left_label.Caption = "Days Left: " + Format(get_days_left)
    Else
      days_left_label.Caption = "Evaluation period expired"
    End If
  End If
End Function



' This program is a code generator for Quicksoft, Inc.
'
' Developed by Ed Kaprocki
' Copyright 2000 QuickSoft, Inc.
'
' There are 2 main public entry points
'
'  Function check_code (name_string, key_string, check_string) as boolean
'     This function returns true or false depending on whether the check_string matches
'     the calculated string.
'
'  Function calculate_code(name_string, key_string) As String
'     This function returns a string of the calculated registration code
'-----------------------------------------------------
'  Inputs
'     name_string
'       This is the name of the user that is registering the software. This
'       name may be any length. All leading, trailing and embedded spaces are
'       ignored and do not affect the check code. Upper and lower case result in
'       the same check code.

'     key_string is a string of numbers which is used to calculate the check code.
'       The length of KEY determines how many check digits in the check code.
'       Valid digits are 0 - 9.
'
'     check_string is a string of alpha characters that is being checked to see if
'       it matches the calculated check code. This is not case sensitive.
'-----------------------------------------------------
'  Outputs
'    Function Check_Code returns true or false. See description above.
'
'    Function Calculate_Code returns a string of the the actual registration code
'
'-----------------------------------------------------

Public Function calculate_code(name_in As String, key_in As String) As String
  Dim i, j As Integer
  Dim v As Long
  Dim c1 As Long
  Dim n1 As Long
  Dim sum As Long
  Dim n As String  ' Main name string variable
  Dim k As String  ' Main key string variable
  Dim r As String  ' Accumulated final code string
  Dim ns As String
  
  calculate_code = ""
  
  ' Convert name to uppercase
  n = UCase(name_in)
  k = UCase(key_in)
  
  ' Strip leading and trailing spaces
  n = RTrim(LTrim(n))
  
  ' Remove any spaces in the name
  ns = ""
  For i = 1 To Len(n)
    If (Mid(n, i, 1) <> " ") Then ns = ns + Mid(n, i, 1)
  Next i
  n = ns
  
  ' Make sure the name length is > 0 and key length is > 0
  If (Len(n) = 0) Or (Len(k) = 0) Then
    Exit Function
  End If
  
  
  ' ----- Now do the algorithm -----
  sum = 0
  For j = 1 To Len(n)
    sum = sum + Asc(Mid(n, j, 1))
  
    r = ""
    For i = 1 To Len(k)
      n1 = Asc(Mid(k, i, 1))
    
      ' Convert v to an alpha character A-Z
      v = (((n1 + sum + (i * j)) Mod 26)) + Asc("A")
    
      r = r + Chr(v)
    Next i
  Next j
 
  calculate_code = UCase(r)
End Function


Public Function check_code(name_in As String, key_in As String, check_in As String) As Boolean
  check_code = False
  If (Len(check_in) = Len(key_in)) And (Len(key_in) > 0) Then
    check_code = (UCase(check_in) = calculate_code(UCase(name_in), UCase(key_in)))
  End If
End Function


