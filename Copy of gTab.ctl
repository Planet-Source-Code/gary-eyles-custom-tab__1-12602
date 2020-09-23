VERSION 5.00
Begin VB.UserControl gTab 
   Appearance      =   0  'Flat
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "gTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum TabOrig
    tTop = 0
    tBottom = 1
    tLeft = 2
    tRight = 3
End Enum

Dim TabOr As TabOrig
Dim RotateText As Boolean
Dim tButtonStyle As Boolean
Dim tButtonHighlight As Boolean

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_BTNTEXT = 18

Const ICC_LISTVIEW_CLASSES = &H1       ' listview, header
Const ICC_TREEVIEW_CLASSES = &H2       ' treeview, tooltips
Const ICC_BAR_CLASSES = &H4            ' toolbar, statusbar, trackbar, tooltips
Const ICC_TAB_CLASSES = &H8            ' tab, tooltips
Const ICC_UPDOWN_CLASS = &H10          ' updown
Const ICC_PROGRESS_CLASS = &H20        ' progress
Const ICC_HOTKEY_CLASS = &H40          ' hotkey
Const ICC_ANIMATE_CLASS = &H80         ' animate
Const ICC_WIN95_CLASSES = &HFF
Const ICC_DATE_CLASSES = &H100         ' month picker, date picker, time picker, updown
Const ICC_USEREX_CLASSES = &H200       ' comboex
Const ICC_COOL_CLASSES = &H400         ' rebar (coolbar) control
Const ICC_INTERNET_CLASSES = &H800
Const ICC_PAGESCROLLER_CLASS = &H1000      ' page scroller
Const ICC_NATIVEFNTCTL_CLASS = &H2000      ' native font control
Private Type InitCommonControlsExType
    dwSize As Long 'size of this structure
    dwICC As Long 'flags indicating which classes to be initialized
End Type
'Private Const WS_VISIBLE = &H10000000
'Private Const WS_CHILD = &H40000000
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function InitCommonControlsEx Lib "comctl32" (init As InitCommonControlsExType) As Boolean
'Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const WS_VISIBLE = &H10000000

Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&

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
    y As Long
    x As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long


Private Type INITCOMMONCONTROLSEXSt
    dwSize As Long
    dwICC As Long
End Type

'Private Const ICC_TAB_CLASSES = &H8

'Private Declare Function InitCommonControlsEx Lib "comctl32" (init As INITCOMMONCONTROLSEXSt) As Long


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
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long

Private Declare Function gSendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

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

Private Const NM_FIRST = -0&
Private Const NM_RCLICK = (NM_FIRST - 5)

Private Const TCM_GETITEM = (TCM_FIRST + 60)
Private Const TCM_GETCURSEL = (TCM_FIRST + 11)



Private Const TCS_SCROLLOPPOSITE = &H1
Private Const TCS_BOTTOM = &H2
Private Const TCS_RIGHT = &H2
Private Const TCS_MULTISELECT = &H4
'#if (_WIN32_IE >= 0x0400)
Private Const TCS_FLATBUTTONS = &H8
'#End If
Private Const TCS_FORCEICONLEFT = &H10
Private Const TCS_FORCELABELLEFT = &H20
'#if (_WIN32_IE >= 0x0300)
Private Const TCS_HOTTRACK = &H40
Private Const TCS_VERTICAL = &H80
'#End If
Private Const TCS_TABS = &H0
Private Const TCS_BUTTONS = &H100
Private Const TCS_SINGLELINE = &H0
Private Const TCS_MULTILINE = &H200
Private Const TCS_RIGHTJUSTIFY = &H0
Private Const TCS_FIXEDWIDTH = &H400
Private Const TCS_RAGGEDRIGHT = &H800
Private Const TCS_FOCUSONBUTTONDOWN = &H1000
Private Const TCS_OWNERDRAWFIXED = &H2000
Private Const TCS_TOOLTIPS = &H4000
Private Const TCS_FOCUSNEVER = &H8000

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Const WM_LBUTTONDOWN = &H201

Private Const WM_PARENTNOTIFY = &H210

Private Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hDC As Long
        rcItem As RECT
        itemData As Long
End Type

Private Declare Function GetModuleFileName Lib "kernel32" _
    Alias "GetModuleFileNameA" _
    ( _
    ByVal hModule As Long, _
    ByVal lpFileName As String, _
    ByVal nSize As Long _
    ) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CENTER = &H1
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_CALCRECT = &H400

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

Private Const TCN_FIRST = -551
Private Const TCN_SELCHANGE = (TCN_FIRST - 1)
Private Const TCN_SELCHANGING = (TCN_FIRST - 2)
Private Const NM_CLICK = (NM_FIRST - 2)
Private Const TCM_SETCURSEL = (TCM_FIRST + 12)
Private Const TCM_DELETEALLITEMS = (TCM_FIRST + 9)

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Dim OldTab As Long

Implements ISubclass
Private m_emr As EMsgResponse

Private Const GWL_STYLE = (-16)
Private Const TCM_GETITEMCOUNT = (TCM_FIRST + 4)

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Event TabChange(TabIndex As Long, TabText As String)
Event Resize()

Public Function CountTabs() As Long
CountTabs = SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0)
End Function

Public Function DeleteAllTabs() As Long
DeleteAllTabs = SendMessage(mWnd, TCM_DELETEALLITEMS, 0, 0)
UserControl_Resize
End Function

Private Function InVBDesignEnvironment() As Boolean
    
    Dim strFileName As String
    Dim lngCount As Long
    
    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = Left(strFileName, lngCount)
    
    InVBDesignEnvironment = False


    If UCase(Right(strFileName, 7)) = "VB5.EXE" Then
        InVBDesignEnvironment = True
    ElseIf UCase(Right(strFileName, 7)) = "VB6.EXE" Then
        InVBDesignEnvironment = True
    End If
End Function

Public Function ButtonHighlight(Optional YesNo As Boolean) As Boolean
If IsMissing(YesNo) Then
    ButtonHighlight = tButtonHighlight
Else
    tButtonHighlight = YesNo
    Call RefreshTabs
End If
End Function

Public Function RefreshTabs()
Dim TmpRect As RECT

Call GetWindowRect(mWnd, TmpRect)
Call RedrawWindow(mWnd, TmpRect, &H100, 1)
End Function

Public Function tRotateText(Optional YesNo As Boolean) As Boolean
If IsMissing(YesNo) Then
    tRotateText = RotateText
Else
    RotateText = YesNo
End If

UserControl_Resize
End Function

Public Function pTop() As Long
Dim TmpRect As RECT

If TabOr = tTop Then
    Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
    pTop = (TmpRect.Bottom - TmpRect.Top) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
Else
    pTop = 0
End If
End Function

Public Function pLeft() As Long
Dim TmpRect As RECT

If TabOr = tLeft Then
    Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
    'pLeft = (TmpRect.Bottom - TmpRect.Top) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
    pLeft = (TmpRect.Right - TmpRect.Left) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
Else
    pLeft = 0
End If
End Function

Public Function pRight() As Long
Dim TmpRect As RECT

If TabOr = tRight Then
    Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
    'pLeft = (TmpRect.Bottom - TmpRect.Top) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
    pRight = (TmpRect.Right - TmpRect.Left) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
    pRight = UserControl.ScaleWidth - pRight
Else
    pRight = UserControl.ScaleWidth
End If
End Function

Public Function pBottom() As Long
Dim TmpRect As RECT

If TabOr = tBottom Then
    Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
    pBottom = (TmpRect.Bottom - TmpRect.Top) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
    pBottom = UserControl.ScaleHeight - pBottom
Else
    pBottom = UserControl.ScaleHeight
End If
End Function

Public Sub DeleteTab(Index As Long)
Dim cIndex As Long
cIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)

SendMessage mWnd, TCM_DELETEITEM, Index, 0

UserControl_Resize

If cIndex = Index Then
    RaiseEvent TabChange(-1, "")
End If
End Sub

Public Sub InsertTab(TheString As String, Optional Index As Long = 0)
Dim ss As String
Dim TmpStuff As TCITEMW

Dim cIndex As Long
cIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)

ss = TheString
TmpStuff.mask = LVIF_TEXT
TmpStuff.pszText = StrConv(ss, vbUnicode)
TmpStuff.cchTextMax = Len(ss)

Call SendMessage(mWnd, TCM_INSERTITEMW, Index, TmpStuff)

UserControl_Resize

If cIndex = -1 Then
    RaiseEvent TabChange(Index, TheString)
End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)

End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse

End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Debug.Print "HERE " & iMsg & " " & wParam & " " & lParam

If iMsg = WM_NOTIFY Then
    'Debug.Print "Notify"
    
    Dim TmpSs As NMHDR
    CopyMemory TmpSs, ByVal lParam, Len(TmpSs)
    'Debug.Print "Code " & TmpSs.code & " " & TCN_SELCHANGE
    'Debug.Print "idFrom " & TmpSs.idFrom
    'Debug.Print "hwndFrom " & TmpSs.hwndFrom & " " & mWnd
    'MsgBox TmpSs.code & vbTab & TmpSs.hwndFrom & vbTab & TmpSs.idFrom
        
If TmpSs.code = NM_CLICK Then
'If TmpSs.code = TCN_SELCHANGE Then
        Dim gcIndex As Long
        gcIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)
    Debug.Print gcIndex

'        If cIndex <> OldTab And OldTab <> -1 Then
'            OldTab = -1
            Dim TmpS As TCITEMW
            TmpS.mask = LVIF_TEXT
            Dim TmpString As String
            TmpString = Space(255)
            TmpS.cchTextMax = 255
            TmpS.pszText = TmpString
            Call SendMessage(mWnd, TCM_GETITEM, gcIndex, TmpS)
            TmpString = (StrConv(TmpS.pszText, vbFromUnicode))
            TmpString = Mid(TmpString, 1, InStr(1, TmpString, Chr$(134), vbTextCompare) - 2)
            RaiseEvent TabChange(gcIndex, TmpString)
End If

        
'        Dim cIndex As Long
'        cIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)
'        If cIndex <> OldTab And OldTab <> -1 Then
'            OldTab = -1
'            Dim TmpS As TCITEMW
'            TmpS.mask = LVIF_TEXT
'            Dim TmpString As String
'            TmpString = Space(255)
'            TmpS.cchTextMax = 255
'            TmpS.pszText = TmpString
'            Call SendMessage(mWnd, TCM_GETITEM, cIndex, TmpS)
'            TmpString = (StrConv(TmpS.pszText, vbFromUnicode))
'            TmpString = Mid(TmpString, 1, InStr(1, TmpString, Chr$(134), vbTextCompare) - 2)
'            RaiseEvent TabChange(cIndex, TmpString)
'        End If

ElseIf iMsg = WM_DRAWITEM Then
    Dim lpds As DRAWITEMSTRUCT
    'Dim TmpText As String
    
    Call CopyMemory(lpds, ByVal lParam, Len(lpds))
    
'Debug.Print lpds.CtlType
If lpds.CtlID <> 101 Then
    
            Dim cIndex2 As Long
            cIndex2 = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)
            
            Dim TmpS2 As TCITEMW
            TmpS2.mask = LVIF_TEXT
            Dim TmpString2 As String
            TmpString2 = Space(255)
            TmpS2.cchTextMax = 255
            TmpS2.pszText = TmpString2
            Call SendMessage(mWnd, TCM_GETITEM, lpds.itemID, TmpS2)
            TmpString2 = (StrConv(TmpS2.pszText, vbFromUnicode))
            TmpString2 = Mid(TmpString2, 1, InStr(1, TmpString2, Chr$(134), vbTextCompare) - 2)
    
    
    
       
    Dim TmpRect As RECT
    LSet TmpRect = lpds.rcItem
    

    If cIndex2 <> lpds.itemID Then
        SetTextColor lpds.hDC, GetSysColor(COLOR_BTNTEXT)
        FillRectEx lpds.hDC, lpds.rcItem, GetSysColor(COLOR_BTNFACE)
        
If Not tButtonStyle Then
        If TabOr = tTop Then
            TmpRect.Top = TmpRect.Top + 4
        ElseIf TabOr = tBottom Then
            TmpRect.Top = TmpRect.Top - 4
        End If
End If

    Else
        SetTextColor lpds.hDC, GetSysColor(COLOR_BTNTEXT)
        FillRectEx lpds.hDC, lpds.rcItem, IIf(tButtonHighlight = True, GetSysColor(COLOR_BTNHIGHLIGHT), GetSysColor(COLOR_BTNFACE))
        'DrawGradient lpds.hDC, lpds.rcItem, QBColor(4), QBColor(7), True
    End If
    
    SetBkMode lpds.hDC, 1
    'SetTextColor lpds.hDC, 0
    
            Dim oldFont As Long
            Dim fnt As New CLogFont
            Set fnt.LogFont = UserControl.Font
            
            If TabOr = tLeft Or TabOr = tRight Then
                If RotateText Then
                    fnt.Rotation = 270
                Else
                    fnt.Rotation = 90
                End If
            End If
            
            oldFont = SelectObject(lpds.hDC, fnt.Handle)
    
    If TabOr = tLeft Or TabOr = tRight Then
        'Call DrawText(lpds.hDC, TmpString2, Len(TmpString2), TmpRect, DT_SINGLELINE)
        Call DrawText(lpds.hDC, TmpString2, Len(TmpString2), TmpRect, DT_SINGLELINE Or DT_CALCRECT)
        
        Dim TmpX, TmpY As Long
        TmpX = ((lpds.rcItem.Right - lpds.rcItem.Left) / 2) - (TmpRect.Bottom - TmpRect.Top) / 2
        TmpY = ((lpds.rcItem.Bottom - lpds.rcItem.Top) / 2) - (TmpRect.Right - TmpRect.Left) / 2
        'Debug.Print "TmpX " & TmpX & vbTab & TmpY & vbTab & TmpString2
        
        If cIndex2 <> lpds.itemID Then
If Not tButtonStyle Then
            If TabOr = tLeft Then
                TmpX = TmpX + 2
            ElseIf TabOr = tRight Then
                TmpX = TmpX - 2
            End If
End If
        End If
        
        If fnt.Rotation = 270 Then
            TmpX = TmpX + (TmpRect.Bottom - TmpRect.Top)
            TmpY = TmpY + (TmpRect.Right - TmpRect.Left)
        End If
        
        Call TextOut(lpds.hDC, lpds.rcItem.Left + TmpX, lpds.rcItem.Bottom - TmpY, TmpString2, Len(TmpString2))
    Else
        Call DrawText(lpds.hDC, TmpString2, Len(TmpString2), TmpRect, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER)
    End If
    
    
        fnt.CleanUp
    
         'If Not (m_Font Is Nothing) Then
            Call SelectObject(lpds.hDC, oldFont)
         'End If
         
    Exit Function
End If
    
ElseIf iMsg = WM_PARENTNOTIFY Then
    'If LoWord(wParam) = WM_LBUTTONDOWN Then
    '    OldTab = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)
    'End If
End If

CallOldWindowProc hWnd, iMsg, wParam, lParam
End Function

Private Sub UserControl_Initialize()
    Const IE3_INSTALLED = True
    If IE3_INSTALLED = True Then
        Dim initcc As InitCommonControlsExType
        initcc.dwSize = Len(initcc)
        initcc.dwICC = ICC_TAB_CLASSES
        InitCommonControlsEx initcc
    Else
        InitCommonControls
    End If

Dim CS As CREATESTRUCT
Dim dStyle As Long
dStyle = WS_CHILD Or WS_VISIBLE
dStyle = dStyle Or TCS_FOCUSONBUTTONDOWN Or TCS_MULTILINE
dStyle = dStyle Or TCS_OWNERDRAWFIXED 'Or TCS_HOTTRACK
       
'mWnd = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TRANSPARENT, "SysTabControl32", "", dStyle, 0, 0, 300, 200, UserControl.hwnd, 0, App.hInstance, CS)
mWnd = CreateWindowEx(0, "SysTabControl32", "", dStyle, 0, 0, 300, 200, UserControl.hWnd, 0, App.hInstance, CS)
ShowWindow mWnd, SW_NORMAL
   
tmpFont = CreateFont(Font.Size, 0, 900, 900, 0, 0, 0, 0, 0, 6, 0, 0, 0, Font.Name)
SendMessage mWnd, WM_SETFONT, tmpFont, 0

'MsgBox InVBDesignEnvironment
'MsgBox UserControl.Ambient.UserMode

    'AttachMessage Me, UserControl.hWnd, WM_NOTIFY
    'AttachMessage Me, UserControl.hWnd, WM_DRAWITEM
    'AttachMessage Me, UserControl.hWnd, WM_PARENTNOTIFY
End Sub

Private Sub UserControl_Resize()
MoveWindow mWnd, 0, 0, ScaleWidth, ScaleHeight, 1

RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
If UserControl.Ambient.UserMode Then
    AttachMessage Me, UserControl.hWnd, WM_NOTIFY
    AttachMessage Me, UserControl.hWnd, WM_DRAWITEM
    AttachMessage Me, UserControl.hWnd, WM_PARENTNOTIFY
End If
End Sub

Private Sub UserControl_Terminate()
'DetachMessage Me, UserControl.hwnd, WM_COMMAND
DetachMessage Me, UserControl.hWnd, WM_NOTIFY
DetachMessage Me, UserControl.hWnd, WM_DRAWITEM
DetachMessage Me, UserControl.hWnd, WM_PARENTNOTIFY
    
DeleteObject tmpFont
DestroyWindow mWnd
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = mWnd
End Property

Private Function HiWord(LongIn As Long) As Integer
     HiWord% = (LongIn& And &HFFFF0000) \ &H10000
End Function

Private Function LoWord(LongIn As Long) As Integer
  Dim l As Long
  
  l& = LongIn& And &HFFFF&
  
  If l& > &H7FFF Then
       LoWord% = l& - &H10000
  Else
       LoWord% = l&
  End If
End Function

Public Function TabIndex(Optional Index As Long) As Long
If IsMissing(Index) Then
    TabIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)
Else
    TabIndex = SendMessage(mWnd, TCM_SETCURSEL, Index, 0)
End If
End Function

Private Sub FillRectEx(hDC As Long, rc As RECT, Color As Long)
'Also based on Paul DiLascia's
'a good idea to simplify the calls to FillRect
  Dim OldBrush As Long
  Dim NewBrush As Long
  
  NewBrush& = CreateSolidBrush(Color&)
  Call FillRect(hDC&, rc, NewBrush&)
  Call DeleteObject(NewBrush&)
End Sub

Public Function GetStyle() As TabOrig
    GetStyle = TabOr
End Function

Public Function GetStyleButton() As Boolean
GetStyleButton = tButtonStyle
End Function

Public Sub ChangeStyle(NewStyle As TabOrig, Optional ButtonStyle As Boolean = False)
Dim dStyle As Long
dStyle = WS_CHILD Or WS_VISIBLE
dStyle = dStyle Or TCS_FOCUSONBUTTONDOWN Or TCS_MULTILINE
dStyle = dStyle Or TCS_OWNERDRAWFIXED 'Or TCS_HOTTRACK

If ButtonStyle = True Then
    dStyle = dStyle Or TCS_BUTTONS
    tButtonStyle = True
ElseIf ButtonStyle = False Then 'Or IsMissing(ButtonStyle) Then
    tButtonStyle = False
End If

If NewStyle = tBottom Then
    dStyle = dStyle Or TCS_BOTTOM
ElseIf NewStyle = tLeft Then
    dStyle = dStyle Or TCS_VERTICAL
ElseIf NewStyle = tRight Then
    dStyle = dStyle Or TCS_VERTICAL Or TCS_RIGHT
ElseIf NewStyle = tTop Then
    'dStyle = dStyle Or TCS_BOTTOM
End If

Call SetWindowLong(mWnd, GWL_STYLE, dStyle)

TabOr = NewStyle

UserControl_Resize
'MoveWindow mWnd, 0, 0, ScaleWidth, ScaleHeight, 1
'RefreshTabs
End Sub

Private Sub DrawGradient( _
      ByVal hDC As Long, _
      ByRef rct As RECT, _
      ByVal lEndColour As Long, _
      ByVal lStartColour As Long, _
      ByVal bVertical As Boolean _
   )
Dim lStep As Long
Dim lPos As Long, lSize As Long
Dim bRGB(1 To 3) As Integer
Dim bRGBStart(1 To 3) As Integer
Dim dR(1 To 3) As Double
Dim dPos As Double, d As Double
Dim hBr As Long
Dim tR As RECT
   
   LSet tR = rct
   If bVertical Then
      lSize = (tR.Bottom - tR.Top)
   Else
      lSize = (tR.Right - tR.Left)
   End If
   lStep = lSize \ 255
   If (lStep < 3) Then
       lStep = 3
   End If
       
   bRGB(1) = lStartColour And &HFF&
   bRGB(2) = (lStartColour And &HFF00&) \ &H100&
   bRGB(3) = (lStartColour And &HFF0000) \ &H10000
   bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
   dR(1) = (lEndColour And &HFF&) - bRGB(1)
   dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
   dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)
        
   For lPos = lSize To 0 Step -lStep
      ' Draw bar:
      If bVertical Then
         tR.Top = tR.Bottom - lStep
      Else
         tR.Left = tR.Right - lStep
      End If
      If tR.Top < rct.Top Then
         tR.Top = rct.Top
      End If
      If tR.Left < rct.Left Then
         tR.Left = rct.Left
      End If
      
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hDC, tR, hBr
      DeleteObject hBr
            
      dPos = ((lSize - lPos) / lSize)
      If bVertical Then
         tR.Bottom = tR.Top
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      Else
         tR.Right = tR.Left
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      End If
      
   Next lPos

End Sub

