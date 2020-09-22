Attribute VB_Name = "Hooking"
Option Explicit

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Const GWL_WNDPROC = (-4)
Public Const WM_COPYDATA = &H4A
Public lpPrevWndProc As Long, gHW As Long, OtherInstanceHwnd As Long, Hooked As Boolean

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub Hook()
    
    gHW = Form1.Hwnd
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
    Hooked = True
    
End Sub

Public Sub Unhook()
          
    Dim temp As Long
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)

End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
          
    If uMsg = WM_COPYDATA Then
        Call MySub(lParam)
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)

End Function

Sub MySub(lParam As Long)
          
    Dim cds As COPYDATASTRUCT
    Dim buf(1 To 255) As Byte, a As String

    Call CopyMemory(cds, ByVal lParam, Len(cds))

    Select Case cds.dwData
        Case 1
            Debug.Print "got a 1"
        Case 2
            Debug.Print "got a 2"
        Case 3
            Call CopyMemory(buf(1), ByVal cds.lpData, cds.cbData)
            a = StrConv(buf, vbUnicode)
            a = Left(a, InStr(1, a, Chr(0)) - 1)
            Call Form1.mnuRestore_Click
            MsgBox "Message received from second instance:" & vbLf & vbLf & a & _
                vbLf & vbLf & "This file could also be processed here!", vbInformation
    End Select

End Sub


