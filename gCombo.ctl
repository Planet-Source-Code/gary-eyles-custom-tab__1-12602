VERSION 5.00
Begin VB.UserControl gCombo 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "gCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const WS_EX_STATICEDGE = &H20000
Const WS_EX_TRANSPARENT = &H20&
Const WS_CHILD = &H40000000
Const CW_USEDEFAULT = &H80000000
Const SW_NORMAL = 1
Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    Y As Long
    X As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long


Private Type INITCOMMONCONTROLSEXSt
    dwSize As Long
    dwICC As Long
End Type

Private Const ICC_TAB_CLASSES = &H8

Private Declare Function InitCommonControlsEx Lib "comctl32" (INIT As INITCOMMONCONTROLSEXSt) As Long


Private Const TCM_FIRST = &H1300
Private Const TCM_INSERTITEMA = (TCM_FIRST + 7)
Private Const TCM_INSERTITEMW = (TCM_FIRST + 62)
Private Const TCM_DELETEITEM = (TCM_FIRST + 8)

Private Type TCITEMHEADER
    mask As Long
    lpReserved1 As Long
    lpReserved2 As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
End Type



Private Type TCITEMA
    mask As Long
'#if (_WIN32_IE >= 0x0300)
    dwState As Long
    dwStateMask As Long
'#Else
'    UINT lpReserved1;
'    UINT lpReserved2;
'#End If
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type

Private Type TCITEMW
    mask As Long
'#if (_WIN32_IE >= 0x0300)
    dwState As Long
    dwStateMask As Long
'#Else
'    UINT lpReserved1;
'    UINT lpReserved2;
'#End If
    pszText As String
    cchTextMax As String
    iImage As String
    lParam As Long
End Type

'#If UNICODE Then
'    Private Const TCITEM = TCITEMW
'    Private Const LPTCITEM = LPTCITEMW
'#Else
'    Private Const TCITEM = TCITEMA
'    Private Const LPTCITEM = LPTCITEMA
'#End If


Private Const WM_DRAWITEM = &H2B


Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const TCM_ADJUSTRECT = (TCM_FIRST + 40)
Private Const TCM_SETMINTABWIDTH = (TCM_FIRST + 49)
Private Const TCM_SETITEMA = (TCM_FIRST + 6)
'#define TCM_SETITEMW            (TCM_FIRST + 61)

Private Const TCM_SETITEMW = (TCM_FIRST + 61)
Private Const LVIF_TEXT = &H1
Private Const TCM_SETEXTENDEDSTYLE = (TCM_FIRST + 52)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long

Private Declare Function gSendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Const WM_SETFONT = &H30

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" _
    (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, _
    ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, _
    ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, _
    ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, _
    ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, _
    ByVal lpszFace As String) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Dim mWnd As Long
Dim tmpFont As Long

Private Const TCM_GETROWCOUNT = (TCM_FIRST + 44)
Private Const TCM_GETITEMRECT = (TCM_FIRST + 10)

Private Const WM_NOTIFY = &H4E
Private Const WM_COMMAND = &H111

Private Type NMHDR
    hwndFrom As Long
    idFrom As Long
    code As Long
End Type

Private Const TCM_GETITEM = (TCM_FIRST + 60)
Private Const TCM_GETCURSEL = (TCM_FIRST + 11)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Implements ISubclass
Private m_emr As EMsgResponse
'Event Declarations:
Event TabChange(TabIndex As Long, TabText As String)

Public Function pTop() As Long
Dim TmpRect As RECT
'Dim pTop    As Long

Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
pTop = (TmpRect.Bottom - TmpRect.Top) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
End Function

Public Sub DeleteTab(Index As Long)
SendMessage mWnd, TCM_DELETEITEM, Index, 0
End Sub

Public Sub InsertTab(TheString As String)
Dim ss As String
Dim TmpStuff As TCITEMW

ss = TheString
TmpStuff.mask = LVIF_TEXT
TmpStuff.pszText = StrConv(ss, vbUnicode)
TmpStuff.cchTextMax = Len(ss)

Call SendMessage(mWnd, TCM_INSERTITEMW, 0&, TmpStuff)
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse

End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Debug.Print "HERE " & wParam & " " & lParam

If iMsg = WM_NOTIFY Then
    'Dim TmpS As NMHDR
    'CopyMemory TmpS, lParam, LenB(TmpS)
    'Debug.Print TmpS.code
    'Debug.Print TmpS.idFrom
    'Debug.Print TmpS.hwndFrom & " " & mWnd
    
If lParam = 8387264 Then
    Dim cIndex As Long
    Dim TmpS As TCITEMW
    cIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)
    TmpS.mask = LVIF_TEXT
    
    Dim TmpString As String
    TmpString = Space(255)
    TmpS.cchTextMax = 255
    'TmpS.pszText = StrPtr(TmpString)
    TmpS.pszText = TmpString
    Call SendMessage(mWnd, TCM_GETITEM, cIndex, TmpS)
    TmpString = (StrConv(TmpS.pszText, vbFromUnicode))
    TmpString = Mid(TmpString, 1, InStr(1, TmpString, Chr$(134), vbTextCompare) - 2)
    'Debug.Print TmpString
    RaiseEvent TabChange(cIndex, TmpString)
End If

ElseIf iMsg = WM_COMMAND Then
    Debug.Print "COMMAND"
Else
    Debug.Print "DRAW"
    Exit Function
End If

CallOldWindowProc hwnd, iMsg, wParam, lParam
End Function

Private Sub UserControl_Initialize()
Dim CS As CREATESTRUCT
Dim dStyle As Long
dStyle = WS_CHILD 'Or &H200
'dStyle = dStyle Or &H0 Or &H1000 Or &H40 'Or &H4 Or &H100
'dStyle = dStyle Or &H2000 'Ownerdraw
    
mWnd = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TRANSPARENT, "ComboBoxEx32", "", dStyle, 0, 0, 300, 200, UserControl.hwnd, 0, App.hInstance, CS)
ShowWindow mWnd, SW_NORMAL
    
'tmpFont = CreateFont(Font.Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font.Name)
'SendMessage mWnd, WM_SETFONT, tmpFont, 0

'AttachMessage Me, UserControl.hwnd, WM_NOTIFY
'AttachMessage Me, UserControl.hwnd, WM_COMMAND
'AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
End Sub

Private Sub UserControl_Resize()
MoveWindow mWnd, 0, 0, ScaleWidth, ScaleHeight, 1
End Sub

Private Sub UserControl_Terminate()
DetachMessage Me, UserControl.hwnd, WM_COMMAND
DetachMessage Me, UserControl.hwnd, WM_NOTIFY
DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
    
    DeleteObject tmpFont
    DestroyWindow mWnd
End Sub

Public Property Get hwnd() As Long
    hwnd = mWnd
End Property

