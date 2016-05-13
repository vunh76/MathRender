Attribute VB_Name = "modWheelHook"
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const MK_CONTROL = &H8
Private Const MK_LBUTTON = &H1
Private Const MK_RBUTTON = &H2
Private Const MK_MBUTTON = &H10
Private Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Private LocalHwnd As Long
Private LocalPrevWndProc As Long
Public Scroll As VScrollBar

Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long
    Dim MouseKeys As Long
    Dim Rotation As Long
    Dim Xpos As Long
    Dim Ypos As Long
    Dim delta As Long
    If Lmsg = WM_MOUSEWHEEL Then
        MouseKeys = Wparam And 65535
        Rotation = Wparam / 65536
        Xpos = Lparam And 65535
        Ypos = Lparam / 65536
        Dim NewValue As Long
        On Error Resume Next
        With Scroll
            If MouseKeys = MK_CONTROL Then
              delta = .LargeChange
            Else
              delta = .SmallChange
            End If
            If Rotation > 0 Then
                NewValue = .Value - delta
                If NewValue < .Min Then
                    NewValue = .Min
                End If
            Else
                NewValue = .Value + delta
                If NewValue > .Max Then
                    NewValue = .Max
                End If
            End If
            .Value = NewValue
        End With
    End If
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, Wparam, Lparam)
End Function

Public Sub WheelHook(PassedForm As PictureBox)
    On Error Resume Next
    LocalHwnd = PassedForm.hWnd
    LocalPrevWndProc = SetWindowLong(LocalHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub WheelUnHook()
    Dim WorkFlag As Long
    On Error Resume Next
    WorkFlag = SetWindowLong(LocalHwnd, GWL_WNDPROC, LocalPrevWndProc)
End Sub

