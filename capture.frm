VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   240
      Left            =   1530
      TabIndex        =   6
      Top             =   945
      Width           =   1140
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   945
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   285
      Left            =   2745
      TabIndex        =   4
      Top             =   585
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   285
      Left            =   1530
      TabIndex        =   3
      Top             =   585
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   585
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   1140
      Left            =   495
      ScaleHeight     =   1080
      ScaleWidth      =   4815
      TabIndex        =   1
      Top             =   2385
      Width           =   4875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   315
      TabIndex        =   0
      Top             =   90
      Width           =   2445
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
      '--------------------------------------------------------------------
      ' Capture the entire screen
      Private Sub Command1_Click()
         Set Picture1.Picture = CaptureScreen()
      End Sub

      ' Capture the entire form including title and border
      Private Sub Command2_Click()
         Set Picture1.Picture = CaptureForm(Me)
      End Sub

      ' Capture the client area of the form
      Private Sub Command3_Click()
         Set Picture1.Picture = CaptureClient(Me)
      End Sub

      ' Capture the active window after two seconds
      Private Sub Command4_Click()
         MsgBox "Two seconds after you close this dialog " & _
            "the active window will be captured."

         ' Wait for two seconds
         Dim EndTime As Date
         EndTime = DateAdd("s", 2, Now)
         Do Until Now > EndTime
            DoEvents
         Loop

         Set Picture1.Picture = CaptureActiveWindow()

         ' Set focus back to form
         Me.SetFocus
      End Sub

      ' Print the current contents of the picture box
      Private Sub Command5_Click()
         PrintPictureToFitPage Printer, Picture1.Picture
         Printer.EndDoc
      End Sub

      ' Clear out the picture box
      Private Sub Command6_Click()
         Set Picture1.Picture = Nothing
      End Sub

      ' Initialize the form and controls
      Private Sub Form_Load()
         Me.Caption = "Capture and Print Example"
         Command1.Caption = "&Screen"
         Command2.Caption = "&Form"
         Command3.Caption = "&Client"
         Command4.Caption = "&Active"
         Command5.Caption = "&Print"
         Command6.Caption = "C&lear"
         Picture1.AutoSize = True
      End Sub
      '--------------------------------------------------------------------





