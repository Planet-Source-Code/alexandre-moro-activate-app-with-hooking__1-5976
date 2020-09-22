VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Minimize to systray, register .tst file type and starts hooking..."
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Not Hooking"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   915
   End
   Begin VB.Menu mnu_1 
      Caption         =   "mnu_1"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore window"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit!"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Declare Function ShowWindowAsync& Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long)
'Const SW_HIDE = 0    'Hides the window. Activation passes to another window.
'Const SW_MINIMIZE = 6     'Minimizes the window. Activation passes to another window.
Private Const SW_RESTORE = 9    'Displays a window at its original size and location and activates it.
'Const SW_SHOW = 5   'Displays a window at its current size and location, and activates it.
'Const SW_SHOWMAXIMIZED = 3      'Maximizes a window and activates it.
'Const SW_SHOWMINIMIZED = 2      'Minimizes a window and activates it.
'Const SW_SHOWMINNOACTIVE = 7    'Minimizes a window without changing the active window.
'Const SW_SHOWNA = 8     'Displays a window at its current size and location. Does not change the active window.
'Const SW_SHOWNOACTIVATE = 4     'Displays a window at its most recent size and location. Does not change the active window.
'Const SW_SHOWNORMAL = 1     'Same as SW_RESTORE.

Private Type NOTIFYICONDATA
    cbSize As Long
    Hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Msg As Long, nid As NOTIFYICONDATA, j As Long, OpenError As Boolean

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONDBLCLK = &H206

Private Sub Command1_Click()
    
    'Only activates hooking in the compiled version.
    'IDE crashes if you close the prog without unhooking first!
    'Also, do not hook twice, else the program crashes!
    If Not Hooked Then
        AssociateFileType "tst", False
        Label1.Caption = "Hooking..."
        Label1.Refresh
        Hook
        Command1.Caption = "Minimize to systray"
    End If
    
    'System tray routines
    nid = SetNotifyIconData(Me.Hwnd, vbNull, NIF_MESSAGE Or NIF_ICON Or NIF_TIP, WM_MOUSEMOVE, Me.Icon, "Hook Test" & vbNullChar)
    j = Shell_NotifyIcon(NIM_ADD, nid)
    
    Me.WindowState = vbMinimized
    
End Sub

Private Sub Form_Activate()

    If StrComp((Right(Command, 3)), "tst", vbTextCompare) = 0 Then MsgBox "File " & Command & " could be processed here!", vbInformation
        
End Sub

Private Sub Form_Load()

    If App.PrevInstance Then
        'This will only work in the COMPILED version,
        'IDE creates a class named "ThunderFormDC"
        OtherInstanceHwnd = fActivateWindowClass("ThunderRT6FormDC", "Hook Test")
        If OtherInstanceHwnd = 0 Then
            MsgBox "Problem activating the other instance!", vbCritical
        Else
            If Command <> "" And StrComp((Right(Command, 3)), "tst", vbTextCompare) = 0 Then
                Dim cds As COPYDATASTRUCT, ThWnd As Long, buf(1 To 255) As Byte, a As String
                ' Get the hWnd of the target application
                ThWnd = OtherInstanceHwnd
                a = Command
                'Copy the string into a byte array, converting it to ASCII
                CopyMemory buf(1), ByVal a, Len(a)
                cds.dwData = 3
                cds.cbData = Len(a) + 1
                cds.lpData = VarPtr(buf(1))
                SendMessage OtherInstanceHwnd, WM_COPYDATA, Me.Hwnd, cds
            End If
        End If
        OpenError = True
        Unload Me
        End
    End If

    'Only now defines the caption, else FindWindow will find just this window
    Me.Caption = "Hook Test"
    
End Sub


Private Function SetNotifyIconData(Hwnd As Long, ID As Long, Flags As Long, CallbackMessage As Long, Icon As Long, tip As String) As NOTIFYICONDATA
          
    Dim nidTemp As NOTIFYICONDATA
    nidTemp.cbSize = Len(nidTemp)
    nidTemp.Hwnd = Hwnd
    nidTemp.uId = ID
    nidTemp.uFlags = Flags
    nidTemp.uCallBackMessage = CallbackMessage
    nidTemp.hIcon = Icon
    nidTemp.szTip = tip & Chr$(0)
    SetNotifyIconData = nidTemp
          
End Function

Private Function fActivateWindowClass(psClassname As String, App As String) As Long
    
    Dim Hwnd As Long
    Hwnd = FindWindow(psClassname, App)
    
    If Hwnd > 0 Then
        ShowWindowAsync Hwnd, SW_RESTORE
        SetForegroundWindow Hwnd
    End If
    
    fActivateWindowClass = Hwnd
    
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.ScaleMode = vbPixels Then
        Msg = X
    Else
        Msg = X / Screen.TwipsPerPixelX
    End If

    Select Case Msg
        Case WM_LBUTTONUP '514 restore form window
        ShowWindowAsync Me.Hwnd, SW_RESTORE
        SetForegroundWindow Me.Hwnd
        Case WM_LBUTTONDBLCLK '515 restore form window
        ShowWindowAsync Me.Hwnd, SW_RESTORE
        SetForegroundWindow Me.Hwnd
        Case WM_RBUTTONUP '517 display popup menu
        Me.PopupMenu Me.mnu_1
    End Select
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If OpenError Then Exit Sub
        
    If MsgBox("This will end the program." & vbLf & vbLf & "Are you sure?", vbYesNo + vbQuestion) = vbYes Then
        If Hooked Then Unhook
        Shell_NotifyIcon NIM_DELETE, nid
        Unload Me
        End
    Else
        Cancel = True
    End If
        
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then Me.Hide
    
End Sub

Private Sub mnuQuit_Click()

    Unload Me
    
End Sub

Sub mnuRestore_Click()

    ShowWindowAsync Me.Hwnd, SW_RESTORE
    SetForegroundWindow Me.Hwnd
    
End Sub


