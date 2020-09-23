VERSION 5.00
Begin VB.UserControl gTab 
   Appearance      =   0  'Flat
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "gTab.ctx":0000
End
Attribute VB_Name = "gTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreatePopupMenu Lib "user32" () As Long
'Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Const MF_STRING = &H0&
Private Const MF_BYPOSITION = &H400&
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Private Const TPM_RIGHTALIGN = &H8&
Private Const TPM_LEFTALIGN = &H0&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_CHECKED = &H8&
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Const WM_COPYDATA = &H4A
Private Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As Long
End Type

Private Const TTF_IDISHWND = &H1
Private Const TTF_CENTERTIP = &H2
Private Const TTF_RTLREADING = &H4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_DI_SETITEM = &H8000
Private Const TTS_ALWAYSTIP = 1

Private Const LPSTR_TEXTCALLBACK As Long = -1
Private Const CW_USEDEFAULT = &H80000000

Private Const WM_ACTIVATE = &H6
Private Const WM_CHAR = &H102

Private Const WS_DISABLED = &H8000000
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_VISIBLE = &H10000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_STATICEDGE = &H20000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_CHILD = &H40000000
Private Const WM_PRINT = &H317
Private Const SW_NORMAL = 1

Private Type RECT
        left As Long
        tOp As Long
        Right As Long
        Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uId As Long
    RECT As RECT
    hInst As Long
    lpszText As Long
    lParam As Long
End Type

Private Const WM_USER = &H400
Private Const TTM_SETTOOLINFOW = (WM_USER + 54)
Private Const TTM_ACTIVATE = (WM_USER + 1)
Private Const TTM_ADDTOOLW = (WM_USER + 50)

Enum TabOrig
    ttop = 0
    tbottom = 1
    tleft = 2
    tRight = 3
End Enum

Dim TabOr As TabOrig
Dim RotateText As Boolean
Dim tButtonStyle As Boolean
Dim tButtonHighlight As Boolean

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

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

Private Declare Sub InitCommonControls Lib "COMCTL32" ()
Private Declare Function InitCommonControlsEx Lib "COMCTL32" (init As InitCommonControlsExType) As Boolean

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
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Type INITCOMMONCONTROLSEXSt
    dwSize As Long
    dwICC As Long
End Type

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
    dwState As Long
    dwStateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type

Private Type TCITEMW
    mask As Long
    dwState As Long
    dwStateMask As Long
    pszText As String
    cchTextMax As String
    iImage As String
    lParam As Long
End Type

Private Const WM_DRAWITEM = &H2B

Private Const TCM_ADJUSTRECT = (TCM_FIRST + 40)
Private Const TCM_SETMINTABWIDTH = (TCM_FIRST + 49)
Private Const TCM_SETITEMA = (TCM_FIRST + 6)

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

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

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
Private Const TCS_FLATBUTTONS = &H8
Private Const TCS_FORCEICONLEFT = &H10
Private Const TCS_FORCELABELLEFT = &H20
Private Const TCS_HOTTRACK = &H40
Private Const TCS_VERTICAL = &H80
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

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
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
        hdc As Long
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

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CENTER = &H1
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_CALCRECT = &H400

Private Const TCN_FIRST = -551
Private Const TCN_SELCHANGE = (TCN_FIRST - 1)
Private Const TCN_SELCHANGING = (TCN_FIRST - 2)
Private Const NM_CLICK = (NM_FIRST - 2)
Private Const TCM_SETCURSEL = (TCM_FIRST + 12)
Private Const TCM_DELETEALLITEMS = (TCM_FIRST + 9)

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Dim OldTab As Long
Dim TipHwnd As Long

Implements ISubclass
Private m_emr As EMsgResponse

Private Const GWL_STYLE = (-16)
Private Const TCM_GETITEMCOUNT = (TCM_FIRST + 4)
Private Const TCM_SETTOOLTIPS = (TCM_FIRST + 46)
Private Const TCM_SETIMAGELIST = (TCM_FIRST + 3)
Private Const TCM_GETIMAGELIST = (TCM_FIRST + 2)

Private Type NMLVGETINFOTIP_NOSTRING
   hdr As NMHDR
   pszText As Long
   cchTextMax As Long
   iItem As Long
End Type

Private Const WS_TABSTOP = &H10000
Private Const TCM_SETITEMSIZE = (TCM_FIRST + 41)
Private Const TCM_SETPADDING = (TCM_FIRST + 43)

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public gsInfoTipBuffer As String

Private Type NMCUSTOMDRAWINFO
    NMHDR As NMHDR
    dwDrawStage As Long
    hdc As Long
    rc As RECT
    dwItemSpec As Long
    uItemState As Long
    lItemlParam As Long
End Type

Private Type NMTTCUSTOMDRAW
    nmcd As NMCUSTOMDRAWINFO
    uDrawFlags As Long
End Type

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function DrawStateString Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lpString As String, _
   ByVal cbStringLen As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal cx As Long, _
   ByVal cy As Long, _
   ByVal fuFlags As Long) As Long

Private Type TCHITTESTINFO
    pt As POINTAPI
    flags As Long
End Type

Private Const TCIF_IMAGE = &H2
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_HITTEST = (WM_USER + 55)
Private Const TCM_HITTEST = (TCM_FIRST + 13)

Dim TabTips() As String
Dim TabImage() As Long
Dim tImgX As Long
Dim tImgY As Long
Dim OldTabSize As Long
Dim tEnabled As Boolean
Dim Inserting As Boolean
Dim IconPlace As Boolean
Dim OpTrue  As Boolean
Dim HasBeenDraw As Boolean
Dim HotTracking As Boolean
Dim pHwnd As Long

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type

Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long
Private Declare Function GetCharWidth Lib "gdi32" Alias "GetCharWidthA" (ByVal hdc As Long, ByVal wFirstChar As Long, ByVal wLastChar As Long, lpBuffer As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Event gTabChange(tTabIndex As Long, tTabString As String)
Event Resize()
'Default Property Values:
Const m_def_ScrollOpposite = 0
'Property Variables:
Dim m_ScrollOpposite As Boolean

Private Type Size
        cx As Long
        cy As Long
End Type
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Const WM_SYSKEYDOWN = &H104

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Const WH_KEYBOARD = 2
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Sub DoMenu()
Dim hMenu As Long
Dim cc As Long
Dim tRect As RECT
Dim xy As POINTAPI

If Not tEnabled Then Exit Sub

GetCursorPos xy

hMenu = CreatePopupMenu

Dim aString As String

'Dim zCnt As Long
Dim zItmInfo As TCITEMW
Dim zaStr As Long
Dim zaStrInfo As String
Dim zTmpString As String
       
Dim CurS As Long

CurS = SendMessage(mWnd, TCM_GETCURSEL, ByVal 0, ByVal 0)
       
For cc = 0 To SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0) - 1
    'aString = "Gaz " & cc '& vbNull
    
    zItmInfo.mask = LVIF_TEXT
    zTmpString = Space(255)
    zItmInfo.cchTextMax = 255
    zItmInfo.pszText = zTmpString
    Call SendMessage(mWnd, TCM_GETITEM, cc, zItmInfo)
    zTmpString = (StrConv(zItmInfo.pszText, vbFromUnicode))
    zTmpString = Mid(zTmpString, 1, InStr(1, zTmpString, Chr$(134), vbTextCompare) - 2)
    
    If cc = CurS Then
        Call InsertMenu(hMenu, &HFFFFFFFF, MF_CHECKED Or MF_BYPOSITION Or MF_STRING, cc + 1, zTmpString)
    Else
        Call InsertMenu(hMenu, &HFFFFFFFF, MF_BYPOSITION Or MF_STRING, cc + 1, zTmpString)
    End If
    'Debug.Print SetMenuItemBitmaps(hMenu, cc, MF_BYCOMMAND, Form1.Picture1.Picture, Form1.Picture1.Picture)
Next

Dim Ret As Long
Ret = TrackPopupMenu(hMenu, TPM_LEFTALIGN Or &H100, xy.x, xy.y, ByVal 0, mWnd, tRect)

    zItmInfo.mask = LVIF_TEXT
    zTmpString = Space(255)
    zItmInfo.cchTextMax = 255
    zItmInfo.pszText = zTmpString
    Call SendMessage(mWnd, TCM_GETITEM, Ret - 1, zItmInfo)
    zTmpString = (StrConv(zItmInfo.pszText, vbFromUnicode))
    zTmpString = Mid(zTmpString, 1, InStr(1, zTmpString, Chr$(134), vbTextCompare) - 2)

    SetFocus mWnd
    Call SendMessage(mWnd, TCM_SETCURSEL, ByVal Ret - 1, 0)
    RaiseEvent gTabChange(Ret - 1, zTmpString)

Call DestroyMenu(hMenu)
End Sub

Public Function IconPlacement(Optional AtLeft As Boolean) As Boolean
    If IsMissing(AtLeft) Then
        IconPlacement = IconPlace
    Else
        IconPlace = AtLeft
    End If

    RefreshTabs
End Function

Public Function GetImagelist() As Long
    GetImagelist = SendMessage(mWnd, TCM_GETIMAGELIST, 0, 0)
End Function

Public Function Enabled(Optional YesNo) As Boolean
If IsMissing(YesNo) Then
    Enabled = tEnabled
    Exit Function
Else
    If CInt(YesNo) < -1 Or CInt(YesNo) > 0 Then
        Exit Function
    End If
    tEnabled = YesNo
End If

Dim dStyle As Long
dStyle = WS_CHILD Or WS_VISIBLE
dStyle = dStyle Or TCS_FOCUSONBUTTONDOWN Or TCS_MULTILINE
dStyle = dStyle Or TCS_OWNERDRAWFIXED 'Or TCS_HOTTRACK
dStyle = dStyle Or TCS_TOOLTIPS

If HotTracking Then
    dStyle = dStyle Or TCS_HOTTRACK
End If

If tButtonStyle = True Then
    dStyle = dStyle Or TCS_BUTTONS
End If

If TabOr = tbottom Then
    dStyle = dStyle Or TCS_BOTTOM
ElseIf TabOr = tleft Then
    dStyle = dStyle Or TCS_VERTICAL
ElseIf TabOr = tRight Then
    dStyle = dStyle Or TCS_VERTICAL Or TCS_RIGHT
End If

If SendMessage(mWnd, TCM_GETIMAGELIST, 0, 0) <> 0 Then
    If TabOr = tleft Or TabOr = tRight Then
        SendMessage mWnd, TCM_SETITEMSIZE, 0, ByVal MAKELONG(0, tImgY + 4)
        SendMessage mWnd, TCM_SETPADDING, 0, ByVal MAKELONG(tImgX, 0)
    Else
        SendMessage mWnd, TCM_SETITEMSIZE, 0, ByVal MAKELONG(0, tImgY + 4)
        SendMessage mWnd, TCM_SETPADDING, 0, ByVal MAKELONG(tImgX, 0)
    End If
End If

If tEnabled = False Then
    dStyle = dStyle Or WS_DISABLED
End If

Call SetWindowLong(mWnd, GWL_STYLE, dStyle)

Call RefreshTabs
End Function

Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
  MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
  MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

Public Function SetImageList(Optional TheHml As vbalImageList = Nothing) As Long

If TheHml Is Nothing Then
    SendMessage mWnd, TCM_SETPADDING, ByVal 0, ByVal MAKELONG(5, 0)
    SetImageList = SendMessage(mWnd, TCM_SETIMAGELIST, 0, ByVal 0)
    SendMessage mWnd, TCM_SETITEMSIZE, 0, ByVal MAKELONG(0, 18)
    tImgX = 0
    tImgY = 0
Else
    SetImageList = SendMessage(mWnd, TCM_SETIMAGELIST, 0, ByVal TheHml.hIml)
    OldTabSize = SendMessage(mWnd, TCM_SETITEMSIZE, 0, ByVal MAKELONG(0, TheHml.IconSizeY + 4))
    SendMessage mWnd, TCM_SETPADDING, 0, ByVal MAKELONG(TheHml.IconSizeX, 0)
    tImgX = TheHml.IconSizeX
    tImgY = TheHml.IconSizeY
End If

RefreshTabs
UserControl_Resize
End Function

Public Function CountTabs() As Long
    CountTabs = SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0)
End Function

Public Function DeleteAllTabs() As Long
DeleteAllTabs = SendMessage(mWnd, TCM_DELETEALLITEMS, 0, 0)

If DeleteAllTabs > 0 Then
    ReDim TabTips(0)
    ReDim TabImage(0)
End If

UserControl_Resize
End Function

Private Function InVBDesignEnvironment() As Boolean
    Dim strFileName As String
    Dim lngCount As Long
    
    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
    strFileName = left(strFileName, lngCount)
    
    InVBDesignEnvironment = False


    If UCase(Right(strFileName, 7)) = "VB5.EXE" Then
        InVBDesignEnvironment = True
    ElseIf UCase(Right(strFileName, 7)) = "VB6.EXE" Then
        InVBDesignEnvironment = True
    End If
End Function

Public Function ButtonHighlight(Optional YesNo) As Boolean
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
Call RedrawWindow(mWnd, TmpRect, ByVal &H100, ByVal 1)

'Resize message
SendMessage mWnd, &H5, 0, 0
End Function

Public Sub SetTooltipBkColor(tColor As Long)
SendMessage TipHwnd, TTM_SETTIPBKCOLOR, tColor, 0
End Sub

Public Sub SetTooltipTextColor(tColor As Long)
SendMessage TipHwnd, TTM_SETTIPTEXTCOLOR, tColor, 0
End Sub

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

If TabOr = ttop Or TabOr = tbottom Then
    Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
    
    If OpTrue And tButtonStyle = False Then
        If TabOr = ttop Then
            pTop = (TmpRect.Bottom - TmpRect.tOp)
        ElseIf TabOr = tbottom Then
            pTop = (TmpRect.Bottom - TmpRect.tOp) * (SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0) - 1)
        End If
    Else
        If TabOr = ttop Then
            pTop = (TmpRect.Bottom - TmpRect.tOp) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
        Else
            pTop = 0
        End If
    End If
Else
    pTop = 0
End If
End Function

Public Function pLeft() As Long
Dim TmpRect As RECT

If TabOr = tleft Or TabOr = tRight Then
    Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
    
    If OpTrue And tButtonStyle = False Then
        If TabOr = tleft Then
            pLeft = (TmpRect.Right - TmpRect.left) '* SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
        ElseIf TabOr = tRight Then
            pLeft = (TmpRect.Right - TmpRect.left) * (SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0) - 1)
        End If
    Else
        If TabOr = tleft Then
            pLeft = (TmpRect.Right - TmpRect.left) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
        Else
            pLeft = 0
        End If
    End If
Else
    pLeft = 0
End If
End Function

Public Function pRight() As Long
Dim TmpRect As RECT

If TabOr = tRight Or TabOr = tleft Then
    Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
    
    If OpTrue And tButtonStyle = False Then
        If TabOr = tRight Then
            pRight = (TmpRect.Right - TmpRect.left) '* SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
            pRight = UserControl.ScaleWidth - pRight
        ElseIf TabOr = tleft Then
            pRight = (TmpRect.Right - TmpRect.left) * (SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0) - 1)
            pRight = UserControl.ScaleWidth - pRight
        End If
    Else
        If TabOr = tRight Then
            pRight = (TmpRect.Right - TmpRect.left) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
            pRight = UserControl.ScaleWidth - pRight
        Else
            pRight = UserControl.ScaleWidth
        End If
    End If
Else
    pRight = UserControl.ScaleWidth
End If
End Function

Public Function pBottom() As Long
Dim TmpRect As RECT

If TabOr = tbottom Or TabOr = ttop Then
    Call SendMessage(mWnd, TCM_GETITEMRECT, 0, TmpRect)
    
    If OpTrue And tButtonStyle = False Then
        If TabOr = ttop Then
            pBottom = (TmpRect.Bottom - TmpRect.tOp) * (SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0) - 1)
            pBottom = UserControl.ScaleHeight - pBottom
        ElseIf TabOr = tbottom Then
            pBottom = (TmpRect.Bottom - TmpRect.tOp)
            pBottom = UserControl.ScaleHeight - pBottom
        End If
    Else
        If TabOr = tbottom Then
            pBottom = (TmpRect.Bottom - TmpRect.tOp) * SendMessage(mWnd, TCM_GETROWCOUNT, 0, 0)
            pBottom = UserControl.ScaleHeight - pBottom
        Else
            pBottom = UserControl.ScaleHeight
        End If
    End If
Else
    pBottom = UserControl.ScaleHeight
End If
End Function

Public Sub DeleteTab(Index As Long)
Dim cIndex As Long
cIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)

If SendMessage(mWnd, TCM_DELETEITEM, Index, 0) > 0 Then
    Dim Cnt As Long
    Index = Index + 1
    For Cnt = Index To UBound(TabTips) - 1
        TabTips(Cnt) = TabTips(Cnt + 1)
        TabImage(Cnt) = TabImage(Cnt + 1)
    Next
    ReDim Preserve TabTips(UBound(TabTips) - 1)
    ReDim Preserve TabImage(UBound(TabImage) - 1)
End If
    
UserControl_Resize

'If cIndex = Index Then
    RaiseEvent gTabChange(-1, "")
'End If
End Sub

Public Sub InsertTab(TheString As String, Optional TheToolTip As String = "", Optional ImageIndex As Long = -1, Optional Index As Long = 0)
On Error GoTo CatchErr

Inserting = True

Dim HasA As Long
Dim HasAStr As String

Dim ss As String
Dim TmpStuff As TCITEMW

Dim cIndex As Long
cIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)

ss = TheString
TmpStuff.mask = LVIF_TEXT 'Or TCIF_IMAGE
TmpStuff.pszText = StrConv(ss, vbUnicode)
TmpStuff.cchTextMax = Len(ss)
'TmpStuff.iImage = ImageIndex

If SendMessage(mWnd, TCM_INSERTITEMW, Index, TmpStuff) > -1 Then
    
    Dim Cnt As Long
    If Index = 0 Then
        ReDim Preserve TabTips(SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0))
        ReDim Preserve TabImage(SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0))
        For Cnt = UBound(TabTips) To 2 Step -1
            TabTips(Cnt) = TabTips(Cnt - 1)
            TabImage(Cnt) = TabImage(Cnt - 1)
        Next
        TabTips(1) = TheToolTip
        TabImage(1) = ImageIndex
        
        HasA = InStr(1, TheString, "&", vbTextCompare)
        If HasA > 0 Then
            HasAStr = Mid(TheString, HasA, 2)
            'SetProp pHwnd, "SPCKEY" & UCase(HasAStr), ByVal 1
            SetProp UserControl.Parent.hwnd, "SPCKEY" & UCase(HasAStr), ByVal UserControl.hwnd
            'Debug.Print "SPCKEY" & HasAStr, pHwnd
        End If
    Else
        Index = Index + 1
        ReDim Preserve TabTips(SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0))
        ReDim Preserve TabImage(SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0))
        For Cnt = UBound(TabTips) To Index + 1 Step -1
            TabTips(Cnt) = TabTips(Cnt - 1)
            TabImage(Cnt) = TabImage(Cnt - 1)
        Next
        TabTips(Index) = TheToolTip
        TabImage(Index) = ImageIndex
        
        HasA = InStr(1, TheString, "&", vbTextCompare)
        If HasA > 0 Then
            HasAStr = Mid(TheString, HasA, 2)
            SetProp UserControl.Parent.hwnd, "SPCKEY" & UCase(HasAStr), ByVal UserControl.hwnd
        End If

    End If

End If

Inserting = False

UserControl_Resize
RefreshTabs

If cIndex = -1 Then
    RaiseEvent gTabChange(Index, TheString)
End If

Exit Sub
CatchErr:

Debug.Print Err.Description
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = m_emr
    
    'Debug.Print "EMR " & m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If Inserting Then Exit Function

If iMsg = &H555 Then
    'Debug.Print "Character"

Dim zCnt As Long
Dim zItmInfo As TCITEMW
Dim zaStr As Long
Dim zaStrInfo As String
Dim zTmpString As String
    
For zCnt = 0 To SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0) - 1
    zItmInfo.mask = LVIF_TEXT
    zTmpString = Space(255)
    zItmInfo.cchTextMax = 255
    zItmInfo.pszText = zTmpString
    Call SendMessage(mWnd, TCM_GETITEM, zCnt, zItmInfo)
    zTmpString = (StrConv(zItmInfo.pszText, vbFromUnicode))
    zTmpString = Mid(zTmpString, 1, InStr(1, zTmpString, Chr$(134), vbTextCompare) - 2)
    
    zaStr = InStr(1, zTmpString, "&", vbTextCompare)
    
    If UCase(Mid(zTmpString, zaStr + 1, 1)) = Chr(wParam) Then
        SetFocus mWnd
        Call SendMessage(mWnd, TCM_SETCURSEL, ByVal zCnt, 0)
        RaiseEvent gTabChange(zCnt, zTmpString)
        Dim SendClick As NMHDR
        SendClick.code = NM_CLICK
        SendClick.hwndFrom = mWnd
        SendClick.idFrom = zCnt
        'Call SendMessage(UserControl.hwnd, WM_NOTIFY, ByVal 0, SendClick)
        Exit Function
        
    End If
Next
    
Exit Function
End If


If iMsg = WM_PRINT Then
    Debug.Print "WM_PRINT"
End If

If iMsg = WM_NOTIFY Then
    Dim TmpSs As NMHDR
    CopyMemory TmpSs, ByVal lParam, Len(TmpSs)
    
    'If TmpSs.code = (NM_FIRST - 18) Then
    '    Debug.Print "Char"
    'End If

    'Message saying the tooltip is about to be shown
    'not needed though
    
    'If TmpSs.code = -520 - 1 Then
    '    Debug.Print "Tip Show"
    'End If
        
    'Tried drawing my own tooltips, can't get it to
    'work though, feel free to help me
    
    'If TmpSs.code = (NM_FIRST - 12) Then
    '    Dim ttDraw As NMTTCUSTOMDRAW
    '    CopyMemory ttDraw, ByVal lParam, Len(ttDraw)
    '    Debug.Print "DRAW " & ttDraw.uDrawFlags
    '    Debug.Print ttDraw.nmcd.dwDrawStage
    '    If ttDraw.nmcd.dwDrawStage = 1 Then
    '        'ttDraw.nmcd.dwDrawStage = &H4
    '        'CopyMemory ByVal lParam, ttDraw, Len(ttDraw)
    '        ISubclass_WindowProc = &H4
    '        m_emr = &H4
    '        'CallOldWindowProc = &H4
    '        'WindowProc = &H4
    '        GoTo TheEnd
    '    End If
    '    Exit Function
    'End If
        
    If TmpSs.code = -520 Then
        Dim tLVIT As NMLVGETINFOTIP_NOSTRING
        Dim sTip As String
         
        CopyMemory tLVIT, ByVal lParam, Len(tLVIT)

        On Error Resume Next
        sTip = TabTips(tLVIT.hdr.idFrom + 1)
                 
        If sTip <> "" Then
            tLVIT.cchTextMax = Len(sTip)
            gsInfoTipBuffer = StrConv(sTip, vbFromUnicode)
            tLVIT.pszText = StrPtr(gsInfoTipBuffer)
            CopyMemory ByVal lParam, tLVIT, Len(tLVIT)
        End If
    End If
        
    If TmpSs.code = NM_CLICK Then
        Dim gcIndex As Long
        Dim TmpS As TCITEMW
        Dim TmpString As String
        
        gcIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)
            
        TmpS.mask = LVIF_TEXT
        TmpString = Space(255)
        TmpS.cchTextMax = 255
        TmpS.pszText = TmpString
        Call SendMessage(mWnd, TCM_GETITEM, gcIndex, TmpS)
        TmpString = (StrConv(TmpS.pszText, vbFromUnicode))
        TmpString = Mid(TmpString, 1, InStr(1, TmpString, Chr$(134), vbTextCompare) - 2)
        RaiseEvent gTabChange(gcIndex, TmpString)
    End If

ElseIf iMsg = WM_DRAWITEM Then
    Dim lpds As DRAWITEMSTRUCT
    Dim ItsHot As Boolean
       
    Call CopyMemory(lpds, ByVal lParam, Len(lpds))
       
    If tEnabled Then
        SetTextColor lpds.hdc, GetSysColor(18)
        Dim gHitT As TCHITTESTINFO
        Dim gRect As RECT
        GetCursorPos gHitT.pt
        GetWindowRect mWnd, gRect
        gHitT.pt.x = gHitT.pt.x - gRect.left
        gHitT.pt.y = gHitT.pt.y - gRect.tOp
        gHitT.flags = &H2 Or &H4
        If SendMessage(mWnd, TCM_HITTEST, ByVal 0, gHitT) = lpds.itemID And HotTracking Then
            'SetTextColor lpds.hdc, QBColor(12)
            SetTextColor lpds.hdc, GetSysColor(26)
            ItsHot = True
        End If
    Else
        SetTextColor lpds.hdc, GetSysColor(14)
    End If

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
            FillRectEx lpds.hdc, lpds.rcItem, GetSysColor(COLOR_BTNFACE)
        
            If Not tButtonStyle Then
                If TabOr = ttop Then
                    TmpRect.tOp = TmpRect.tOp + 3
                ElseIf TabOr = tbottom Then
                    TmpRect.tOp = TmpRect.tOp - 3
            End If
        End If

    Else
        FillRectEx lpds.hdc, lpds.rcItem, IIf(tButtonHighlight = True And tEnabled = True, GetSysColor(COLOR_BTNHIGHLIGHT), GetSysColor(COLOR_BTNFACE))
        'DrawGradient lpds.hdc, lpds.rcItem, QBColor(4), QBColor(7), True
    End If
    
    SetBkMode lpds.hdc, 1
       
    Dim oldFont As Long
    Dim fnt As New CLogFont
    Set fnt.LOGFONT = UserControl.Font
            
    If TabOr = tleft Or TabOr = tRight Then
        If RotateText Then
            fnt.Rotation = 270
        Else
            fnt.Rotation = 90
        End If
    End If
            
    oldFont = SelectObject(lpds.hdc, fnt.Handle)
    
    If TabOr = tleft Or TabOr = tRight Then
        Call DrawText(lpds.hdc, TmpString2, Len(TmpString2), TmpRect, DT_SINGLELINE Or DT_CALCRECT Or DT_NOPREFIX)
        
        Dim TmpX, TmpY As Long
        TmpX = ((lpds.rcItem.Right - lpds.rcItem.left) / 2) - (TmpRect.Bottom - TmpRect.tOp) / 2
        TmpY = ((lpds.rcItem.Bottom - lpds.rcItem.tOp) / 2) - (TmpRect.Right - TmpRect.left) / 2
               
        If cIndex2 <> lpds.itemID Then
            If Not tButtonStyle Then
                If TabOr = tleft Then
                    TmpX = TmpX + 2
                ElseIf TabOr = tRight Then
                    TmpX = TmpX - 2
                End If
            End If
        End If
        
        If fnt.Rotation = 270 Then
            TmpX = TmpX + (TmpRect.Bottom - TmpRect.tOp)
            TmpY = TmpY + (TmpRect.Right - TmpRect.left)
        End If
        
        If SendMessage(mWnd, TCM_GETIMAGELIST, 0, 0) <> 0 Then
            TmpY = TmpY + tImgX / 2
        End If
        
        If IconPlace Then
            TmpY = TmpY - tImgY
        End If
        
Dim Findc As Long
Findc = InStr(1, TmpString2, "&", vbTextCompare)
If Findc > 0 Then
    Dim sChar As String * 1
    Dim stChars As String
    sChar = Mid(TmpString2, Findc + 1, 1)
    stChars = Mid(TmpString2, 1, Findc - 1)
    
'Debug.Print TmpString2, stChars, sChar
    
    TmpString2 = Mid(TmpString2, 1, Findc - 1) & Mid(TmpString2, Findc + 1, Len(TmpString2) - Findc)

    Dim dummy As POINTAPI
    Dim lX As Long
    Dim lY As Long
    Dim cWidth As Long
    Dim ctWidth As Long
    lX = lpds.rcItem.left + TmpX + (TmpRect.Bottom - TmpRect.tOp)
    
    If RotateText Then
        lX = lX - (TmpRect.Bottom - TmpRect.tOp) * 2
    Else
        lX = lX - 1
    End If
    
    Call GetCharWidth(lpds.hdc, Asc(sChar), Asc(sChar), cWidth)
    
    'Dim cc As Long
    'Dim TmpWidth As Long
    Dim TheWidth As Long
    'For cc = 1 To Len(stChars)
    '    Call GetCharWidth(lpds.hdc, Asc(Mid(stChars, cc, 1)), Asc(Mid(stChars, cc, 1)), TmpWidth)
    '    TheWidth = TheWidth + TmpWidth
    'Next
    
    Dim sSize As Size
    Call GetTextExtentPoint32(lpds.hdc, stChars, Len(stChars), sSize)
    TheWidth = sSize.cx
    
    Dim NewPen As Long
    Dim OldPen As Long
    
    If tEnabled Then
        NewPen = CreatePen(0, 1, IIf(ItsHot, GetSysColor(26), GetSysColor(18)))
    Else
        NewPen = CreatePen(0, 1, GetSysColor(14))
    End If
    OldPen = SelectObject(lpds.hdc, NewPen)
    
    If RotateText Then
        lY = lpds.rcItem.Bottom - TmpY
        MoveToEx lpds.hdc, lX, lY + cWidth + TheWidth - 2, dummy
        LineTo lpds.hdc, lX, lY + TheWidth - 1
    Else
        lY = lpds.rcItem.Bottom - TmpY
        MoveToEx lpds.hdc, lX, lY - TheWidth, dummy
        LineTo lpds.hdc, lX, lY - cWidth - TheWidth
    End If
    
    Call SelectObject(lpds.hdc, OldPen)
    DeleteObject NewPen
    
    If Not tEnabled Then
        NewPen = CreatePen(0, 1, GetSysColor(17))
        OldPen = SelectObject(lpds.hdc, NewPen)
        If RotateText Then
            MoveToEx lpds.hdc, lX - 1, lY + cWidth + TheWidth - 2 + 1, dummy
            LineTo lpds.hdc, lX - 1, lY + TheWidth - 1 + 1
        Else
            MoveToEx lpds.hdc, lX - 1, lY - TheWidth + 1, dummy
            LineTo lpds.hdc, lX - 1, lY - cWidth - TheWidth + 1
        End If
        Call SelectObject(lpds.hdc, OldPen)
        DeleteObject NewPen
    End If
    
End If
        
        Call TextOut(lpds.hdc, lpds.rcItem.left + TmpX, lpds.rcItem.Bottom - TmpY, TmpString2, Len(TmpString2))
        
        If Not tEnabled Then
            SetTextColor lpds.hdc, GetSysColor(17)
            Call TextOut(lpds.hdc, lpds.rcItem.left + TmpX - 1, lpds.rcItem.Bottom - TmpY + 1, TmpString2, Len(TmpString2))
        End If
    Else
    
        If SendMessage(mWnd, TCM_GETIMAGELIST, 0, 0) <> 0 Then
            If IconPlace Then
                TmpRect.Right = TmpRect.Right - tImgX
            Else
                TmpRect.left = TmpRect.left + tImgX
            End If
        End If
    
        Dim aTmpRct As RECT
        LSet aTmpRct = TmpRect
        
        Call DrawText(lpds.hdc, TmpString2, Len(TmpString2), TmpRect, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER)
        
        If Not tEnabled Then
            aTmpRct.Bottom = aTmpRct.Bottom - 1
            aTmpRct.left = aTmpRct.left - 1
            aTmpRct.Right = aTmpRct.Right - 1
            aTmpRct.tOp = aTmpRct.tOp - 1
            SetTextColor lpds.hdc, GetSysColor(17)
            Call DrawText(lpds.hdc, TmpString2, Len(TmpString2), aTmpRct, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER)
        End If
        
        Call DrawText(lpds.hdc, TmpString2, Len(TmpString2), TmpRect, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER Or DT_CALCRECT)
        
    End If
    
    
        fnt.CleanUp
    
        Call SelectObject(lpds.hdc, oldFont)

    If SendMessage(mWnd, TCM_GETIMAGELIST, 0, 0) <> 0 Then
        Dim IconY As Long
        Dim IconX As Long

        If TabOr = tleft Or TabOr = tRight Then
            IconY = lpds.rcItem.Bottom + 3
            IconY = IconY - TmpY + 3
            If RotateText Then
                IconY = IconY + (TmpRect.Right - TmpRect.left)
            End If
    
            If IconPlace Then
                IconY = IconY - (TmpRect.Right - TmpRect.left)
                IconY = IconY - tImgY / 2
                IconY = IconY - tImgY - 3
            End If

        Else
            IconY = (lpds.rcItem.Bottom - lpds.rcItem.tOp) / 2
            IconY = IconY + lpds.rcItem.tOp
            IconY = IconY - (tImgY / 2)
        End If

        If TabOr = tleft Or TabOr = tRight Then
            IconX = lpds.rcItem.left + (lpds.rcItem.Right - lpds.rcItem.left) / 2
            IconX = IconX - tImgX / 2
            If TabOr = tleft Then
                IconX = IconX + 1
            Else
                IconX = IconX - 1
            End If
        Else
            IconX = lpds.rcItem.left + (lpds.rcItem.Right - lpds.rcItem.left) / 2
            IconX = IconX - (TmpRect.Right - TmpRect.left) / 2
            IconX = IconX - tImgX
            IconX = IconX + 3
        
            If IconPlace Then
                IconX = IconX + (TmpRect.Right - TmpRect.left)
                IconX = IconX + tImgX
                IconX = IconX - 6
            End If
        End If

        If cIndex2 <> lpds.itemID And OpTrue = False Then
            If Not tButtonStyle Then
                If TabOr = ttop Then
                    IconY = IconY + 3
                ElseIf TabOr = tbottom Then
                    IconY = IconY - 3
                End If
            End If
        End If

        If tEnabled Then
            ImageList_Draw SendMessage(mWnd, TCM_GETIMAGELIST, 0, 0), _
            TabImage(lpds.itemID + 1), lpds.hdc, _
            IconX, IconY, &H1 Or &H40
        Else
            Dim hIcon As Long
            hIcon = ImageList_GetIcon(SendMessage(mWnd, TCM_GETIMAGELIST, 0, 0), TabImage(lpds.itemID + 1), 0)
            Call DrawState(lpds.hdc, 0, 0, hIcon, 0, IconX, IconY, 16, 16, DST_ICON Or DSS_DISABLED)
            DeleteObject hIcon
        End If
    End If
                  
    Exit Function
    End If
End If

TheEnd:

CallOldWindowProc hwnd, iMsg, wParam, lParam
End Function

Private Sub UserControl_EnterFocus()
    On Error Resume Next
    'SetForegroundWindow UserControl.Parent.hwnd
    SetFocus mWnd
End Sub

Private Sub UserControl_Resize()
If Not UserControl.Ambient.UserMode Then
    If Not HasBeenDraw Then
        CreateAll
        HasBeenDraw = True
    End If
End If

MoveWindow mWnd, 0, 0, ScaleWidth, ScaleHeight, 1

RaiseEvent Resize
End Sub

Private Sub UserControl_Show()
If UserControl.Ambient.UserMode Then
    AttachMessage Me, UserControl.hwnd, WM_NOTIFY
    AttachMessage Me, UserControl.hwnd, WM_DRAWITEM
    AttachMessage Me, UserControl.hwnd, WM_PRINT
    AttachMessage Me, UserControl.hwnd, &H555
    
    pHwnd = UserControl.Parent.hwnd
    Debug.Print "Parent HWND = " & pHwnd
       
    If wHook = 0 Then
        wHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, App.ThreadID)
    End If
    
    'AttachMessage Me, UserControl.hwnd, WM_PARENTNOTIFY
    'AttachMessage Me, TipHwnd, &H1
    'SetParent TipHwnd, UserControl.Parent.hwnd
End If
End Sub

Private Sub UserControl_Terminate()
Dim Cnt As Long
Dim ItmInfo As TCITEMW
Dim aStr As Long
Dim aStrInfo As String
Dim TmpString As String

For Cnt = 0 To SendMessage(mWnd, TCM_GETITEMCOUNT, 0, 0) - 1
    ItmInfo.mask = LVIF_TEXT
    TmpString = Space(255)
    ItmInfo.cchTextMax = 255
    ItmInfo.pszText = TmpString
    Call SendMessage(mWnd, TCM_GETITEM, Cnt, ItmInfo)
    TmpString = (StrConv(ItmInfo.pszText, vbFromUnicode))
    TmpString = Mid(TmpString, 1, InStr(1, TmpString, Chr$(134), vbTextCompare) - 2)
    
    aStr = InStr(1, TmpString, "&", vbTextCompare)
    If aStr > 0 Then
        aStrInfo = Mid(TmpString, aStr, 2)
        If GetProp(pHwnd, "SPCKEY" & aStrInfo) > 0 Then
            RemoveProp pHwnd, "SPCKEY" & aStrInfo
        End If
    End If
Next

'Call SendMessage(mWnd, TCM_GETITEM, gcIndex, TmpS)
'TmpString = (StrConv(TmpS.pszText, vbFromUnicode))
'TmpString = Mid(TmpString, 1, InStr(1, TmpString, Chr$(134), vbTextCompare) - 2)
'RaiseEvent TabChange(gcIndex, TmpString)
        
DetachMessage Me, UserControl.hwnd, WM_NOTIFY
DetachMessage Me, UserControl.hwnd, WM_DRAWITEM
DetachMessage Me, UserControl.hwnd, WM_PRINT
DetachMessage Me, UserControl.hwnd, &H555
UnhookWindowsHookEx wHook
'DetachMessage Me, UserControl.hwnd, WM_COMMAND
'DetachMessage Me, UserControl.hwnd, WM_PARENTNOTIFY
'DetachMessage Me, TipHwnd, &H1
    
DeleteObject tmpFont
DestroyWindow TipHwnd
DestroyWindow mWnd
End Sub

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = mWnd
End Property

Private Function HiWord(LongIn As Long) As Integer
     HiWord% = (LongIn& And &HFFFF0000) \ &H10000
End Function

Private Function LoWord(dwValue As Long) As Integer
  CopyMemory LoWord, dwValue, 2
End Function

Public Function tTabIndex(Optional Index) As Long
If IsMissing(Index) Then
    tTabIndex = SendMessage(mWnd, TCM_GETCURSEL, 0, 0)
Else
    If IsNumeric(Index) Then
        tTabIndex = SendMessage(mWnd, TCM_SETCURSEL, Index, 0)
    End If
End If
End Function

Private Sub FillRectEx(hdc As Long, rc As RECT, Color As Long)
'Also based on Paul DiLascia's
'a good idea to simplify the calls to FillRect
  Dim OldBrush As Long
  Dim NewBrush As Long
  
  NewBrush& = CreateSolidBrush(Color&)
  Call FillRect(hdc&, rc, NewBrush&)
  Call DeleteObject(NewBrush&)
End Sub

Public Function GetStyle() As TabOrig
    GetStyle = TabOr
End Function

Public Function GetStyleButton() As Boolean
    GetStyleButton = tButtonStyle
End Function

Public Function HotTrack(Optional YesNo) As Boolean
Dim dStyle As Long
Dim NewStyle As TabOrig

If IsMissing(YesNo) Then
    HotTrack = HotTracking
    Exit Function
Else
    If CInt(YesNo) < -1 Or CInt(YesNo) > 0 Then
        Exit Function
    End If
End If

NewStyle = TabOr

dStyle = WS_CHILD Or WS_VISIBLE
dStyle = dStyle Or TCS_FOCUSONBUTTONDOWN Or TCS_MULTILINE
dStyle = dStyle Or TCS_OWNERDRAWFIXED 'Or TCS_HOTTRACK
dStyle = dStyle Or TCS_TOOLTIPS

If YesNo Then
    dStyle = dStyle Or TCS_HOTTRACK
    HotTracking = True
Else
    HotTracking = False
End If

If tButtonStyle = True Then
    dStyle = dStyle Or TCS_BUTTONS
End If

If NewStyle = tbottom Then
    dStyle = dStyle Or TCS_BOTTOM
ElseIf NewStyle = tleft Then
    dStyle = dStyle Or TCS_VERTICAL
ElseIf NewStyle = tRight Then
    dStyle = dStyle Or TCS_VERTICAL Or TCS_RIGHT
ElseIf NewStyle = ttop Then
    'dStyle = dStyle Or TCS_BOTTOM
End If

If tEnabled = False Then
    dStyle = dStyle Or WS_DISABLED
End If

Call SetWindowLong(mWnd, GWL_STYLE, dStyle)
End Function

Public Sub ChangeStyle(NewStyle As TabOrig, Optional ButtonStyle As Boolean = False)
Dim dStyle As Long
dStyle = WS_CHILD Or WS_VISIBLE
dStyle = dStyle Or TCS_FOCUSONBUTTONDOWN Or TCS_MULTILINE
dStyle = dStyle Or TCS_OWNERDRAWFIXED
dStyle = dStyle Or TCS_TOOLTIPS
dStyle = dStyle 'Or TCS_SCROLLOPPOSITE
'Debug.Print "STYLE " & GetWindowLong(mWnd, GWL_STYLE)

If HotTracking Then
    dStyle = dStyle Or TCS_HOTTRACK
End If

If ButtonStyle = True Then
    dStyle = dStyle Or TCS_BUTTONS
    tButtonStyle = True
ElseIf ButtonStyle = False Then 'Or IsMissing(ButtonStyle) Then
    tButtonStyle = False
End If

If NewStyle = tbottom Then
    dStyle = dStyle Or TCS_BOTTOM
ElseIf NewStyle = tleft Then
    dStyle = dStyle Or TCS_VERTICAL
ElseIf NewStyle = tRight Then
    dStyle = dStyle Or TCS_VERTICAL Or TCS_RIGHT
ElseIf NewStyle = ttop Then
    'dStyle = dStyle Or TCS_BOTTOM
End If

If SendMessage(mWnd, TCM_GETIMAGELIST, 0, 0) <> 0 Then
    If NewStyle = tleft Or NewStyle = tRight Then
        SendMessage mWnd, TCM_SETITEMSIZE, 0, ByVal MAKELONG(0, tImgY + 4)
        SendMessage mWnd, TCM_SETPADDING, 0, ByVal MAKELONG(tImgX, 0)
    Else
        SendMessage mWnd, TCM_SETITEMSIZE, 0, ByVal MAKELONG(0, tImgY + 4)
        SendMessage mWnd, TCM_SETPADDING, 0, ByVal MAKELONG(tImgX, 0)
    End If
End If

If tEnabled = False Then
    dStyle = dStyle Or WS_DISABLED
End If

Call SetWindowLong(mWnd, GWL_STYLE, dStyle)

TabOr = NewStyle

UserControl_Resize
'MoveWindow mWnd, 0, 0, ScaleWidth, ScaleHeight, 1
'RefreshTabs
End Sub

Private Sub DrawGradient( _
      ByVal hdc As Long, _
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
      lSize = (tR.Bottom - tR.tOp)
   Else
      lSize = (tR.Right - tR.left)
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
         tR.tOp = tR.Bottom - lStep
      Else
         tR.left = tR.Right - lStep
      End If
      If tR.tOp < rct.tOp Then
         tR.tOp = rct.tOp
      End If
      If tR.left < rct.left Then
         tR.left = rct.left
      End If
      
      hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
      FillRect hdc, tR, hBr
      DeleteObject hBr
            
      dPos = ((lSize - lPos) / lSize)
      If bVertical Then
         tR.Bottom = tR.tOp
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      Else
         tR.Right = tR.left
         bRGB(1) = bRGBStart(1) + dR(1) * dPos
         bRGB(2) = bRGBStart(2) + dR(2) * dPos
         bRGB(3) = bRGBStart(3) + dR(3) * dPos
      End If
      
   Next lPos

End Sub

Private Sub CreateAll()
Dim initcc As InitCommonControlsExType
initcc.dwSize = Len(initcc)
initcc.dwICC = ICC_TAB_CLASSES
InitCommonControlsEx initcc

Dim CS As CREATESTRUCT
Dim dStyle As Long
dStyle = WS_CHILD Or WS_VISIBLE
dStyle = dStyle Or TCS_MULTILINE
dStyle = dStyle Or TCS_OWNERDRAWFIXED
dStyle = dStyle Or TCS_TOOLTIPS
dStyle = dStyle Or TCS_FOCUSONBUTTONDOWN

If ScrollOpposite Then
Debug.Print "HERE"
    dStyle = dStyle Or TCS_SCROLLOPPOSITE
    OpTrue = True
End If
       
mWnd = CreateWindowEx(0, "SysTabControl32", "", dStyle, 0, 0, 300, 200, UserControl.hwnd, 0, App.hInstance, CS)

Dim pHwnd As Long
pHwnd = UserControl.Parent.hwnd
        
TipHwnd = CreateWindowEx(0&, "tooltips_class32", "", TTS_ALWAYSTIP, _
    CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, _
    pHwnd, 0&, App.hInstance, 0&)

Dim TmpTipInfo As TOOLINFO
TmpTipInfo.cbSize = Len(TmpTipInfo)
TmpTipInfo.uFlags = TTF_IDISHWND + TTF_SUBCLASS
TmpTipInfo.lpszText = LPSTR_TEXTCALLBACK
TmpTipInfo.hwnd = mWnd
TmpTipInfo.uId = mWnd
TmpTipInfo.hInst = App.hInstance

Call SendMessage(TipHwnd, TTM_ADDTOOLW, 0, TmpTipInfo)

Call SendMessage(mWnd, TCM_SETTOOLTIPS, TipHwnd, 0)
Call SendMessage(TipHwnd, TTM_ACTIVATE, 1, mWnd)

ShowWindow mWnd, SW_NORMAL
   
tmpFont = CreateFont(Font.Size, 0, 900, 900, 0, 0, 0, 0, 0, 6, 0, 0, 0, Font.Name)
SendMessage mWnd, WM_SETFONT, tmpFont, 0

tEnabled = True
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,1,0,0
Public Property Get ScrollOpposite() As Boolean
    ScrollOpposite = m_ScrollOpposite
End Property

Public Property Let ScrollOpposite(ByVal New_ScrollOpposite As Boolean)
    If Ambient.UserMode Then Err.Raise 382
    m_ScrollOpposite = New_ScrollOpposite
    PropertyChanged "ScrollOpposite"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ScrollOpposite = m_def_ScrollOpposite

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_ScrollOpposite = PropBag.ReadProperty("ScrollOpposite", m_def_ScrollOpposite)

If UserControl.Ambient.UserMode Then
    If Not HasBeenDraw Then
        CreateAll
        HasBeenDraw = True
    End If
End If
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ScrollOpposite", m_ScrollOpposite, m_def_ScrollOpposite)
End Sub

