Attribute VB_Name = "Module1"
Option Explicit

' This module is used to create hyperlinks and
' to open applications
' Use it as follows:
'
' Call Navigate(Me, "http://www.vbthunder.com")
'
' If you want to, say, start up a .DOC file with Microsoft Word, you can change the code to read:
' Call Navigate(Me, "C:\MyPath\testdoc.doc")


Public Const SW_SHOW = 1

Public Declare Function ShellExecute Lib _
"shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long


'Now everything for the command is set up.
'Add a public subroutine for use throughout
'your project like this:
Public Sub Navigate(frm As Form, ByVal WebPageURL As String)
     Dim hBrowse As Long
     hBrowse = ShellExecute(frm.hWnd, "open", WebPageURL, "", "", SW_SHOW)
End Sub
'--end block--'
 


