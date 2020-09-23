Attribute VB_Name = "Module1"
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Global wHook As Long
Private Const VK_SHIFT = &H10

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Const WM_COPYDATA = &H4A
Private Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Function KeyboardProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim FindParent As Long
Dim tFindParent As Long
Dim ThePropHwnd As Long
Dim CpyData As COPYDATASTRUCT
Dim Doneit As Boolean
    
    If idHook < 0 Then
        KeyboardProc = CallNextHookEx(wHook, idHook, wParam, ByVal lParam)
    Else

If Not (lParam And &H80000000) Then
    If lParam And &H20000000 Then
        If wParam >= 65 And wParam <= 90 Then
            'Debug.Print Chr(wParam), lParam And -&H4000, wParam, GetForegroundWindow
            tFindParent = GetForegroundWindow
            Do
                FindParent = tFindParent
                tFindParent = GetParent(FindParent)
            Loop Until tFindParent = 0
            'Debug.Print FindParent, Form1.hwnd
            ThePropHwnd = GetProp(FindParent, "SPCKEY&" & Chr(wParam))
            If ThePropHwnd > 0 Then
                Doneit = True
                Call SendMessage(ThePropHwnd, &H555, ByVal wParam, lParam)
            End If
        End If
    End If
End If

        If Not Doneit Then
            KeyboardProc = CallNextHookEx(wHook, idHook, wParam, ByVal lParam)
        End If
    End If
End Function
