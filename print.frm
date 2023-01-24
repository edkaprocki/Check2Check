VERSION 5.00
Begin VB.Form print_form 
   Caption         =   "Print Preview"
   ClientHeight    =   3690
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   7155
   Icon            =   "print.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   180
      Max             =   100
      TabIndex        =   4
      Top             =   2880
      Width           =   6495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2655
      LargeChange     =   10
      Left            =   6720
      Max             =   100
      TabIndex        =   3
      Top             =   180
      Width           =   255
   End
   Begin VB.Frame frame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   6555
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   15500
         Left            =   60
         ScaleHeight     =   1033
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   713
         TabIndex        =   5
         Top             =   60
         Width           =   10700
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Next Page"
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "print_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim scale_x As Single
Dim scale_y As Single
Dim next_x As Integer
Dim next_y As Integer
Public exit_value As Integer  ' 0=canceled, 1=next page
Public last_page As Boolean  'true=last page, false=no

Dim where As Integer
Dim margin_x As Integer
Dim margin_y As Integer
Option Explicit

Private Sub CancelButton_Click()
  Hide
End Sub

Private Sub Form_activate()
  Dim i
  
  exit_value = 0  ' Set for cancel
  
  If (where = PTR) Then
    Printer.FontName = "Arial"
    Printer.DrawWidth = 1
  End If
  
  pic.FontName = "Arial"
  pic.DrawWidth = 1
  
  OKButton.Visible = Not last_page
  
End Sub

Private Sub Form_Resize()
  If (print_form.WindowState = 1) Then Exit Sub  ' Don't resize if minimized
  
  CancelButton.Top = print_form.height - CancelButton.height - 500
  OKButton.Top = print_form.height - OKButton.height - 500
  
  VScroll1.Left = print_form.width - 500
  
  HScroll1.width = VScroll1.Left - HScroll1.Left
  
  HScroll1.Top = OKButton.Top - 400
  VScroll1.height = HScroll1.Top - VScroll1.Top
  
  frame.width = HScroll1.width
  frame.height = VScroll1.height
    
End Sub

Private Sub HScroll1_Scroll()
  pic.Left = -(pic.width - frame.width) * HScroll1.Value / 100#
End Sub

Private Sub OKButton_Click()
  exit_value = 1
  Hide
End Sub

Function newx(ByVal x As Integer) As Integer
  newx = (x + margin_x) * scale_x
End Function

Function newy(y As Integer) As Integer
  newy = (y + margin_y) * scale_y
End Function

Sub get_scale()
  With Printer
    margin_x = 50
    margin_y = 50
    scale_x = .width / 700# * 0.85
    scale_y = .height / 1000# * 0.85
  End With
End Sub

Sub print_next(x As Integer, y As Integer, w As Integer, s As String, fs As Integer)
  ' Print out this string but don't exceed the width
  Dim i, j
  
  If printer_error Then Exit Sub
  
  next_x = x
  next_y = y
  
  On Error GoTo error_h
    
  If (where = SCR) Then
    pic.CurrentX = next_x
    pic.CurrentY = next_y
    pic.FontSize = fs
    pic.FontName = print_form.FontName
    ' Shorten the string if it exceeds the width
    While (pic.TextWidth(s) > (w)) And (w > 0)
      s = Left(s, Len(s) - 1)
    Wend
    
    pic.Print s
  Else
    Printer.CurrentX = newx(next_x)
    Printer.CurrentY = newy(next_y)
    Printer.FontSize = fs
    Printer.FontName = print_form.FontName
    Printer.FontBold = True
    ' Shorten the string if it exceeds the width
    While (pic.TextWidth(s) > (w)) And (w > 0)
      s = Left(s, Len(s) - 1)
    Wend
    
    Printer.Print s;
  End If
  Exit Sub
  
error_h:
  printer_error = True
End Sub

Sub print_next_right(x As Integer, y As Integer, s As String, fs As Integer)
  Dim i, j, w
  
  If printer_error Then Exit Sub
  
  If (where = SCR) Then
    pic.FontSize = fs
    pic.Font = FontName
    w = pic.TextWidth(s)
    Call print_next(x - w, y, 0, s, fs)
  Else
    Printer.FontSize = fs
    Printer.FontName = print_form.FontName
    Printer.FontBold = True
    w = Printer.TextWidth(s)
    Printer.CurrentX = newx(x) - w
    Printer.CurrentY = newy(y)
    Printer.Print s;
  End If
End Sub

Sub print_line(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, w As Integer)

    If printer_error Then Exit Sub
  
    If (where = SCR) Then
      pic.DrawWidth = w
      pic.DrawStyle = vbSolid
      pic.Line (x1, y1)-(x2, y2)
    Else
      Printer.DrawWidth = w + 1
      Printer.DrawStyle = vbSolid
      Printer.Line (newx(x1), newy(y1))-(newx(x2), newy(y2))
    End If
End Sub

Sub print_dash(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, w As Integer)
    If printer_error Then Exit Sub
  
    If (where = SCR) Then
      pic.DrawWidth = w
      pic.DrawStyle = vbDot
      pic.Line (x1, y1)-(x2, y2)
    Else
      Printer.DrawWidth = w + 1
      Printer.DrawStyle = vbSolid
      Printer.Line (newx(x1), newy(y1))-(newx(x2), newy(y2))
    End If
End Sub

Sub print_box(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer, w As Integer)
  If printer_error Then Exit Sub
  
  If (where = SCR) Then
    pic.DrawWidth = w
    pic.DrawStyle = vbSolid
    pic.Line (x1, y1)-(x2, y2), , B
  Else
    Printer.DrawWidth = w
    Printer.DrawStyle = vbSolid
    Printer.Line (newx(x1), newy(y1))-(newx(x2), newy(y2)), , B
  End If
End Sub

Private Sub VScroll1_Scroll()
  If (where = SCR) Then
    pic.Top = -(pic.height - frame.height) * VScroll1.Value / 100#
  End If
End Sub


Sub new_page()
  If printer_error Then Exit Sub
  
  If (where = SCR) Then
    pic.Cls
  Else
    pic.Cls
    Printer.NewPage
  End If
End Sub


Sub start_document(i As Integer)
  If printer_error Then Exit Sub
  
  ' SCR=screen, PTR=printer
  where = i
  
  If (where = SCR) Then
    pic.Cls
  Else
    get_scale
  End If
End Sub

Sub end_document()
  If (where = SCR) Then
    'pic.Cls
  Else
    Printer.EndDoc
  End If
End Sub

Sub print_picture(p As Picture, x As Integer, y As Integer)
  Dim w As Integer
  
  If (where = SCR) Then
    pic.PaintPicture p, x, y
  Else
    get_scale
    Printer.PaintPicture p, newx(x), newy(y)
  End If
End Sub
