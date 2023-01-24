Attribute VB_Name = "recent"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2000 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const SHARD_PIDL = &H1&
Public Const SHARD_PATH = &H2&

Public Declare Function SHAddToRecentDocs Lib "shell32.dll" _
  (ByVal dwFlags As Long, ByVal dwData As String) As Long
'--end block--'
   
Public Sub save_to_recent_docs(name_s As String)
   Dim r As Long
   Dim shFlag As Long
   Dim shData As String
   Dim saveit As Boolean
   
   shFlag = SHARD_PATH

   If (name_s <> "") Then saveit = True
   
   Select Case saveit
      Case True:
        'We're about to add an item, so
        'set the 'file name' from Text1
         shData = (name_s)

      Case False:
        'User selected delete.
        'Pass a null string as the data
         shData = vbNullString
      
   End Select
  
  'update or clear the list
   r = SHAddToRecentDocs(shFlag, shData)

      
End Sub
'--end block--'
   



