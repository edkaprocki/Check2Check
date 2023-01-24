Attribute VB_Name = "Module7"
' Wheel Mouse support
' These procedures came from the web.

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
ByVal lpPrevWndFunc As Long, _
ByVal hWnd As Long, _
ByVal Msg As Long, _
ByVal Wparam As Long, _
ByVal Lparam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Public Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_MBUTTON = &H10
Public Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Dim LocalHwnd As Long
Dim LocalPrevWndProc As Long
Dim MyForm As Form

'Now copy the following functions into the same code module.

Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long

    Dim MouseKeys As Long
    Dim Rotation As Long
    Dim Xpos As Long
    Dim Ypos As Long
    
    If Lmsg = WM_MOUSEWHEEL Then
        MouseKeys = Wparam And 65535
        Rotation = Wparam / 65536
        Xpos = Lparam And 65535
        Ypos = Lparam / 65536
        'MyForm.MouseWheel MouseKeys, Rotation, Xpos, Ypos
        MouseWheel MouseKeys, Rotation, Xpos, Ypos
    End If
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, Wparam, Lparam)
End Function

Public Sub WheelHook(PassedForm As Form)
' ESK 1/26/2017
' For some unknown reason back in 2005 this Exit Sub was inserted below which would prevent
' mouse wheel scrolling of the FlexGrid on the main form. By commenting out this line the
' mouse wheel scrolling is operational again. I am using the touch pad on my laptop to verify it works. Maybe
' there was some issue when using a regular wheel mouse. I'm going to leave this wheel scrolling
' active. The stable version, 3.4011, used this function with no known problems.
'Exit Sub
    
    On Error Resume Next

    Set MyForm = PassedForm
    LocalHwnd = PassedForm.hWnd
    LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub


Public Sub WheelUnHook()
    Dim WorkFlag As Long

    On Error Resume Next
    WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
    Set MyForm = Nothing
End Sub


' FlexGrid mouse wheel support
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim NewValue As Long
    Dim Lstep As Single

    On Error Resume Next

    With MyForm.entry_grid
        Lstep = .height / .RowHeight(0)
        Lstep = Int(Lstep)
        If Lstep < 10 Then
            Lstep = 10
        End If
        
        Lstep = 1
        If Rotation > 0 Then
            NewValue = .TopRow - Lstep
            If NewValue < 1 Then
                NewValue = 1
            End If
        Else
            NewValue = .TopRow + Lstep
            If NewValue > .Rows - 1 Then
                NewValue = .Rows - 1
            End If
        End If
        
        .TopRow = NewValue
    End With
End Sub

