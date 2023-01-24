VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form integrity_form 
   Caption         =   "Database Integrity"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   Icon            =   "integrity_form.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton view_all_button 
      Caption         =   "View All Records"
      Height          =   495
      Left            =   3180
      TabIndex        =   6
      Top             =   120
      Width           =   1155
   End
   Begin VB.CommandButton ok_button 
      Height          =   495
      Left            =   6120
      Picture         =   "integrity_form.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton test_button 
      Caption         =   "Check Database"
      Height          =   495
      Left            =   1860
      TabIndex        =   4
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton delete_bad_records_button 
      Caption         =   "Delete Bad Records"
      Height          =   495
      Left            =   4380
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2235
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   3942
      _Version        =   393216
      Cols            =   7
      ScrollTrack     =   -1  'True
   End
   Begin VB.Label bad_label 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Good_label 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "integrity_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnt


Public Sub add()
  cnt = cnt + 1
  With grid
    .Col = 0
    .Text = Str(cnt)
    .Col = 1
    .Text = Str(this.this)
    .Col = 2
    .Text = Str(this.year)
    .Col = 3
    .Text = Str(this.month)
    .Col = 4
    .Text = Str(this.day)
    .Col = 5
    .Text = this.name
    .Col = 6
    .CellAlignment = flexAlignRightCenter

    .Text = currency_s(this.amount)
    
    .Rows = .Rows + 1
    .Row = .Row + 1
    
  End With
End Sub

Public Sub init()
  delete_bad_records_button.Visible = False
  grid.Clear
  With grid
    .Rows = 2
    .Row = 0
    .Col = 1
    .Text = "Record"
    .Col = 2
    .Text = "Year"
    .Col = 3
    .Text = "Month"
    .Col = 4
    .Text = "Day"
    .Col = 5
    .Text = "Name"
    .Col = 6
    .Text = "Amount"
  
    grid.Rows = 2
    grid.Row = 1
  End With
End Sub

Private Sub delete_bad_records_button_Click()
  integrity_form.MousePointer = vbHourglass
  grid.Row = 1
  grid.Col = 1
  While (grid.Text <> "")
    delete_record (Val(grid.Text))
    grid.Row = grid.Row + 1
    changed_flag = True
    main_form.update_caption
  Wend
  integrity_form.MousePointer = vbDefault
  grid.Clear
  delete_bad_records_button.Visible = False
End Sub

Private Sub Form_Resize()
  Dim i
  
  If (integrity_form.Width < 500) Or ((integrity_form.Height - 375 - grid.Top) < 0) Then Exit Sub
  'grid.Top = 0
  grid.Left = 0
  grid.Width = integrity_form.Width - 100
  grid.Height = integrity_form.Height - 375 - grid.Top
  
  For i = 0 To grid.Cols - 1
    grid.ColWidth(i) = (grid.Width - 350) / grid.Cols
  Next i
  
End Sub

Private Sub ok_button_Click()
  Hide
End Sub

Private Sub test_button_Click()
  Dim good, bad
  Dim done
  Dim first_pass
  
  integrity_form.MousePointer = vbHourglass
  init
  
  grid.Redraw = False  ' Don't draw the grid yet
  
  ' Loop through the entire database and test it
  good = 0
  bad = 0
  cnt = 0
  done = False
  first_pass = True
  If (data.number_of_records > 0) Then
    ' We have at least one record
    While done = False
    
      If (first_pass) Then
        data.current = data.first
        get_record  ' Get the first record
      Else
        If (get_next_record = False) Then done = True
      End If
      
      If (done = False) Then
        ' See if this record has valid parameters
        If (this.day = 0) Then
          ' We have a problem
          'delete_record
          bad = bad + 1
          integrity_form.add
        Else
          ' We have a valid first record
          good = good + 1
        End If
      End If
      first_pass = False
    Wend
  End If
    
  integrity_form.Good_label.Caption = "Good:" + Str(good)
  integrity_form.bad_label.Caption = "Bad: " + Str(bad)

  If (bad <> 0) Then delete_bad_records_button.Visible = True
  
  grid.Redraw = True
  integrity_form.MousePointer = vbDefault
  
End Sub

Private Sub view_all_button_Click()
  Dim done
  Dim first_pass
  
  integrity_form.MousePointer = vbHourglass
  init
  
  grid.Redraw = False  ' Don't draw the grid yet
  
' Loop through the entire database and test it
  cnt = 0
  done = False
  first_pass = True
  If (data.number_of_records > 0) Then
    ' We have at least one record
    While done = False
    
      If (first_pass) Then
        data.current = data.first
        get_record  ' Get the first record
      Else
        If (get_next_record = False) Then done = True
      End If
      
      If (done = False) Then
        ' See if this record has valid parameters
        add
      End If
      first_pass = False
    Wend
  End If
    
  Good_label.Caption = "Records:" + Str(cnt)
  integrity_form.bad_label.Caption = ""
  
  grid.Redraw = True
  integrity_form.MousePointer = vbDefault
  
End Sub
