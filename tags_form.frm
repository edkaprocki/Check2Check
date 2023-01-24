VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form tags_form 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Tags"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   8
      Cols            =   5
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "tags_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If (KeyCode = vbKeyEscape) Or (KeyCode = vbKeyReturn) Then
    Form_Unload 0
  End If
End Sub


Private Sub Form_Load()
  ' Set up the form
  Dim lR As Long
  lR = SetTopMostWindow(tags_form.hWnd, True)
  
  With grid
    
    .row = 0
    .Col = 0
    '.ColWidth(0) = 1000
    .CellAlignment = flexAlignCenterCenter
    .Col = 1
    '.ColWidth(1) = 1400
    .CellAlignment = flexAlignCenterCenter
    .Col = 2
    .RowHeight(2) = 50
    .CellAlignment = flexAlignCenterCenter
    .Col = 3
    '.ColWidth(1) = 1400
    .CellAlignment = flexAlignCenterCenter
    .Col = 4
    .CellAlignment = flexAlignCenterCenter
    
    .ColAlignment(0) = flexAlignCenterCenter
  End With
  
  'tags_form.Width = grid.Width + 100
  'tags_form.Height = grid.Height + 400
  
  ' Set up the tag mask that is used by many functions
  tag_mask(0) = 1
  tag_mask(1) = 2
  tag_mask(2) = 4
  tag_mask(3) = 8
  tag_mask(4) = 16
  tag_mask(5) = 32
  tag_mask(6) = 64
  tag_mask(7) = 128
  
End Sub

Private Sub Form_Resize()
  Dim i
  
  If (tags_form.width < 500) Or (tags_form.height < 875) Then Exit Sub
  grid.Top = 0
  grid.Left = 0
  grid.width = tags_form.width - 100
  grid.height = tags_form.height - 375 ' - 500
  
  For i = 0 To grid.Cols - 1
    grid.ColWidth(i) = (grid.width - 350) / grid.Cols
  Next i
End Sub

Public Sub update_tags_display()
  Dim i
  Dim sum As Double
  
  Dim lR As Long
  lR = SetTopMostWindow(tags_form.hWnd, True)
  
  With grid
    .Redraw = False
    
    .TextMatrix(0, 0) = ""
    .TextMatrix(1, 0) = words(TOTAL_N)
    .TextMatrix(2, 0) = ""
    .TextMatrix(3, 0) = words(DONE_N)
    .TextMatrix(4, 0) = words(PENDING_N)
    .TextMatrix(5, 0) = words(BLANK_N)
    .TextMatrix(6, 0) = words(SKIP_N)
    .TextMatrix(7, 0) = words(QTY_N)
    
    .TextMatrix(0, 1) = words(TAGS_N) + " 1"
    .TextMatrix(0, 2) = words(TAGS_N) + " 2"
    .TextMatrix(0, 3) = words(TAGS_N) + " 3"
    .TextMatrix(0, 4) = words(TAGS_N) + " 4"
    sum = 0
    
    For i = 0 To MAX_TAG
      .Col = i + 1
      
      sum = sum + tags(i).total

      .row = 1
      .CellForeColor = amount_color(tags(i).total)
      .Text = currency_s(tags(i).total)
      .CellFontBold = True
      
      .row = 3
      .CellForeColor = amount_color(tags(i).done)
      .Text = currency_s(tags(i).done)
      .CellFontBold = True
      
      .row = 4
      .CellForeColor = amount_color(tags(i).pending)
      .Text = currency_s(tags(i).pending)
      .CellFontBold = True
      
      .row = 5
      .CellForeColor = amount_color(tags(i).blank)
      .Text = currency_s(tags(i).blank)
      .CellFontBold = True
      
      .row = 6
      .CellForeColor = amount_color(tags(i).skip)
      .Text = currency_s(tags(i).skip)
      .CellFontBold = True
      
      .row = 7
      .Text = tags(i).number
      .CellFontBold = False
      .CellAlignment = flexAlignRightCenter
    
    Next i
  
    .row = 0
    .Col = 0
    .CellForeColor = amount_color(sum)
    .Text = currency_s(sum)
    .CellFontBold = True
    
    .Redraw = True
    
  End With
  
  tags_form.Caption = words(TAGS_N) + " --- " + main_form.entry_tab.Caption  'TabCaption(i)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Hide
End Sub

Public Sub normal()
  Dim lR As Long
  lR = SetTopMostWindow(tags_form.hWnd, True)
  
  tags_form.WindowState = 0  ' Normal state
End Sub
