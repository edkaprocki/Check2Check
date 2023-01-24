Attribute VB_Name = "Module9"
Option Explicit

'----------------------------------------------------------------
'Author: Dr. John A. Nyhart
'work  : john_nyhart@medicalogic.com
'home  : jnyhart@spessart.com
'web   : www.spessart.com/users/jnyhart/john1.htm
'Posted:7/18/97
'
'How do I play a WAV file with VB?
'----------------------------------------------------------------
Private Const SND_APPLICATION = &H80 ' look for application specific association
Private Const SND_ALIAS = &H10000 ' name is a WIN.INI [sounds] entry
Private Const SND_ALIAS_ID = &H110000 ' name is a WIN.INI [sounds] entry identifier
Private Const SND_ASYNC = &H1 ' play asynchronously
Private Const SND_FILENAME = &H20000 ' name is a file name
Private Const SND_LOOP = &H8 ' loop the sound until next sndPlaySound
Private Const SND_MEMORY = &H4 ' lpszSoundName points to a memory file
Private Const SND_NODEFAULT = &H2 ' silence not default, if sound not found
Private Const SND_NOSTOP = &H10 ' don't stop any currently playing sound
Private Const SND_NOWAIT = &H2000 ' don't wait if the driver is busy
Private Const SND_PURGE = &H40 ' purge non-static events for task
Private Const SND_RESOURCE = &H40004 ' name is a resource name or atom
Private Const SND_SYNC = &H0 ' play synchronously (default)
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long


Public Sub play_sound(i As Integer)
If (i = 0) Or (preferences.play_sounds = False) Then Exit Sub

If (i = 1) Then PlaySound App.Path + "\welcome.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
If (i = 2) Then PlaySound App.Path + "\sdwms.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
If (i = 9) Then PlaySound App.Path + "\sound999.wav", ByVal 0&, SND_FILENAME Or SND_ASYNC
End Sub



