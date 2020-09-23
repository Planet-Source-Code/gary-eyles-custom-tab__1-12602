VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ILDMonoHDC As Long
Private m_ILDMonoHBMP As Long
Private m_ILDMonoHBMPOld As Long

Private m_ILDColorHDC As Long
Private m_ILDColorHBMP As Long
Private m_ILDColorHBMPOld As Long

Public Enum eilIconSize
  Size16 = 16
  Size32 = 32
End Enum

Public Enum eilIconState
  Normal = 0
  Disabled = 1
End Enum

Private m_hIml As Long

Private Const ILC_MASK = &H1
Private Const ILC_COLOR = &H0
Private Const ILC_COLORDDB = &H0
Private Const ILC_COLOR4 = &H4
Private Const ILC_COLOR8 = &H8
Private Const ILC_COLOR16 = &H10
Private Const ILC_COLOR24 = &H18
Private Const ILC_COLOR32 = &H20
 
Private Const CLR_NONE = -1
Private Const CLR_DEFAULT = -16777216
Private Const CLR_HILIGHT = -16777216

Public Enum ImageTypes
  IMAGE_BITMAP = 0
  IMAGE_ICON = 1
  IMAGE_CURSOR = 2
  'IMAGE_ENHMETAFILE = 3
End Enum
 
Private Declare Function ImageList_SetBkColor Lib "COMCTL32" (ByVal hImageList As Long, ByVal clrBk As Long) As Long
Private Declare Function ImageList_GetBkColor Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Private Declare Function ImageList_DrawEx Lib "COMCTL32" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Private Declare Function ImageList_Convert Lib "COMCTL32" Alias "ImageList_Draw" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Private Declare Function ImageList_Create Lib "COMCTL32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_Replace Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hbmImage As Long, ByVal hBmMask As Long) As Long
Private Declare Function ImageList_Add Lib "COMCTL32" (ByVal hImageList As Long, ByVal hbmImage As Long, hBmMask As Long) As Long
Private Declare Function ImageList_Remove Lib "COMCTL32" (ByVal hImageList As Long, ImgIndex As Long) As Long
Private Type IMAGEINFO
    hBitmapImage As Long
    hBitmapMask As Long
    cPlanes As Long
    cBitsPerPixel As Long
    rcImage As RECT
End Type
Private Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        pImageInfo As IMAGEINFO _
    ) As Long
Private Declare Function ImageList_AddIcon Lib "COMCTL32" (ByVal hIml As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32" (ByVal hImageList As Long, ByVal ImgIndex As Long, hBmMask As Long) As Long
Private Declare Function ImageList_SetImageCount Lib "COMCTL32" (ByVal hImageList As Long, uNewCount As Long)
Private Declare Function ImageList_GetImageCount Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_Destroy Lib "COMCTL32" (ByVal hImageList As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Private Declare Function ImageList_SetIconSize Lib "COMCTL32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long

Private Declare Function ImageList_LoadImage Lib "COMCTL32" Alias "ImageList_LoadImageA" (ByVal hInst As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long)

Public Function Create(ByVal hdc As Long, ByVal ImgSize As eilIconSize) As Boolean
     
     ' Do we already have an image list?
    Destroy

    'Create the Imagelist
    m_hIml = ImageList_Create(ImgSize, ImgSize, ILC_MASK, 4, 4)
    If (m_hIml <> 0) And (m_hIml <> -1) Then
        pMakeWorkDCS hdc, ImgSize
    Else
        m_hIml = 0
    End If
    
End Function
Public Sub Destroy()
    If (hIml <> 0) Then
        ImageList_Destroy hIml
        pClearUpWorkDCS
        m_hIml = 0
    End If
End Sub
Public Sub DrawImage( _
        ByVal iImgIndex As Long, _
        ByVal hdc As Long, _
        ByVal xPixels As Integer, _
        ByVal yPixels As Integer, _
        Optional ByVal bSelected = False, _
        Optional ByVal bDisabled = False _
    )
Dim hIcon As Long
Dim lFlags As Long
Dim lhIml As Long

   lhIml = hIml
    
    lFlags = ILD_TRANSPARENT
    If (bSelected) Or (bDisabled) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    
    If (bDisabled) Then
        lFlags = lFlags Or ILD_SELECTED
        ImageList_DrawEx _
              lhIml, _
              iImgIndex, _
              hdc, _
              xPixels, yPixels, 0, 0, _
              CLR_NONE, GetSysColor(COLOR_WINDOW), _
              lFlags
    Else
        ImageList_Draw _
            lhIml, _
            iImgIndex, _
            hdc, _
            xPixels, _
            yPixels, _
            lFlags
    End If
End Sub
Public Sub DrawImage2( _
        ByVal iImgIndex As Long, _
        ByVal hdc As Long, _
        ByVal xPixels As Integer, _
        ByVal yPixels As Integer, _
        Optional ByVal bSelected = False, _
        Optional ByVal bCut = False, _
        Optional ByVal bDisabled = False, _
        Optional ByVal hExternalIml As Long = 0 _
    )
Dim hIcon As Long
Dim lFlags As Long
Dim lhIml As Long

    If (hExternalIml <> 0) Then
        lhIml = hExternalIml
    Else
        lhIml = hIml
    End If
    
    lFlags = ILD_TRANSPARENT
    If (bSelected) Or (bCut) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    
    If (bCut) Then
        lFlags = lFlags Or ILD_SELECTED
        ImageList_DrawEx _
              lhIml, _
              iImgIndex, _
              hdc, _
              xPixels, yPixels, 0, 0, _
              CLR_NONE, GetSysColor(COLOR_WINDOW), _
              lFlags
    ElseIf (bDisabled) Then
        ' todo
            ' use drawstate...
    Else
        ImageList_Draw _
            lhIml, _
            iImgIndex, _
            hdc, _
            xPixels, _
            yPixels, _
            lFlags
    End If
End Sub

Public Property Get IconSize() As Integer
Dim ImgHeight As Long, ImgWidth As Long
    ImageList_GetIconSize hIml, ImgHeight, ImgWidth
    IconSize = ImgHeight
End Property
Public Property Get ImageCount() As Integer
    ImageCount = ImageList_GetImageCount(hIml)
End Property
Public Sub RemoveImage(ByVal Index As Integer)
    ImageList_Remove hIml, ByVal Index
End Sub
Public Sub Clear()
    ImageList_Remove hIml, -1
End Sub
Public Function AddFromFile( _
        ByVal sFileName As String, _
        ByVal iType As ImageTypes, _
        Optional ByVal bMapSysColors As Boolean = False, _
        Optional ByVal lBackColor As OLE_COLOR = -1 _
    ) As Long
Dim hImage As Long
Dim un2 As Long
    
    un2 = LR_LOADFROMFILE
    ' Load the image from file:
    If bMapSysColors Then
        un2 = un2 Or LR_LOADMAP3DCOLORS
    End If
    hImage = LoadImage(App.hInstance, sFileName, iType, 0, 0, un2)
    If (hImage <> 0) Then
        If (iType = IMAGE_BITMAP) Then
            ' And add it to the image list:
            AddFromFile = ImageList_AddMasked(hIml, hImage, lBackColor)
        ElseIf (iType = IMAGE_ICON) Then
            AddFromFile = ImageList_AddIcon(hIml, hImage)
        End If
    Else
        AddFromFile = -1
    End If
    
End Function
Public Function AddFromPictureBox( _
        ByVal hdc As Long, _
        pic As Object, _
        Optional ByVal LeftPixels As Long = 0, _
        Optional ByVal TopPixels As Long = 0, _
        Optional ByVal lBackColor As OLE_COLOR = -1 _
    ) As Long
Dim lHDC As Long
Dim lhBmp As Long, lhBmpOld As Long
Dim tBM As BITMAP
Dim lAColor As Long
Dim lW As Long, lH As Long
Dim hBrush As Long
Dim tR As RECT
Dim lR As Long
Dim lIconSize As Long
Dim lBPixel As Long
    
    lIconSize = IconSize
    ' Create a DC to hold the bitmap to transfer into the image list:
    lHDC = CreateCompatibleDC(hdc)
    If (lHDC <> 0) Then
        ' Create a bitmap compatible with the current device
        ' to copy the picture into:
        'GetObjectAPI pic.Picture.Handle, LenB(tBm), tBm
        'tBm.bmBits = 0
        'tBm.bmWidth = lIconSize
        'tBm.bmHeight = lIconSize
        'lHbmp = CreateBitmapIndirect(tBm)
        lhBmp = CreateCompatibleBitmap(hdc, lIconSize, lIconSize)
        If (lhBmp <> 0) Then
            ' Get the backcolor to use:
            If (lBackColor = -1) Then
                ' None specified, use the colour at 0,0:
                lBackColor = GetPixel(pic.hdc, 0, 0)
            Else
                ' Try to get the specified backcolor:
                If OleTranslateColor(lBackColor, 0, lAColor) Then
                    ' Failed- use default of silver
                    lBackColor = &HC0C0C0
                Else
                    ' Set to GDI version of OLE Color
                    lBackColor = lAColor
                End If
            End If
            ' Select the bitmap into the DC
            lhBmpOld = SelectObject(lHDC, lhBmp)
            ' Clear the background:
            hBrush = CreateSolidBrush(lBackColor)
            tR.Right = lIconSize: tR.Bottom = lIconSize
            FillRect lHDC, tR, hBrush
            DeleteObject hBrush
            
            ' Get the source picture's dimension:
            GetObjectAPI pic.Picture.Handle, LenB(tBM), tBM
            lW = 16
            lH = 16
            If (lW + LeftPixels > tBM.bmWidth) Then
                lW = tBM.bmWidth - LeftPixels
            End If
            If (lH + TopPixels > tBM.bmHeight) Then
                lH = tBM.bmHeight - TopPixels
            End If
            If (lW > 0) And (lH > 0) Then
                ' Blt from the picture into the bitmap:
                lR = BitBlt(lHDC, 0, 0, lW, lH, hdc, LeftPixels, TopPixels, SRCCOPY)
                Debug.Assert (lR <> 0)
            End If
            
            ' We now have the image in the bitmap, so select it out of the DC:
            SelectObject lHDC, lhBmpOld
            ' And add it to the image list:
            lR = ImageList_AddMasked(hIml, lhBmp, lBackColor)
            Debug.Assert (lR <> -1)
            AddFromPictureBox = lR
                
            DeleteObject lhBmp
        End If
        ' Clear up the DC:
        DeleteObject lHDC
    End If
    
End Function
Public Property Get hIml() As Long
    hIml = m_hIml
End Property
Private Sub pMakeWorkDCS( _
        ByVal lHDCBasis As Long, _
        ByVal lIconSize As Long _
    )
    m_ILDMonoHDC = CreateCompatibleDC(0)
    If (m_ILDMonoHDC <> 0) Then
        m_ILDMonoHBMP = CreateCompatibleBitmap(m_ILDMonoHDC, lIconSize, lIconSize * 3)
        If (m_ILDMonoHBMP <> 0) Then
            m_ILDMonoHBMPOld = SelectObject(m_ILDMonoHDC, m_ILDMonoHBMP)
        End If
    End If
    
    m_ILDColorHDC = CreateCompatibleDC(lHDCBasis)
    If (m_ILDColorHDC <> 0) Then
        m_ILDColorHBMP = CreateCompatibleBitmap(lHDCBasis, lIconSize, lIconSize * 2)
        If (m_ILDColorHBMP <> 0) Then
            m_ILDColorHBMPOld = SelectObject(m_ILDColorHDC, m_ILDColorHBMP)
        End If
    End If
End Sub
Private Sub pClearUpWorkDCS()
    If (m_ILDMonoHDC <> 0) Then
        If (m_ILDMonoHBMP <> 0) Then
            SelectObject m_ILDMonoHDC, m_ILDMonoHBMPOld
            DeleteObject m_ILDMonoHBMP
        End If
        DeleteObject m_ILDMonoHDC
    End If
    If (m_ILDColorHDC <> 0) Then
        If (m_ILDColorHBMP <> 0) Then
            SelectObject m_ILDColorHDC, m_ILDColorHBMPOld
            DeleteObject m_ILDColorHBMP
        End If
        DeleteObject m_ILDColorHDC
    End If
End Sub
Private Sub pImageListDrawIconDisabled( _
        ByVal lHDC As Long, _
        ByVal hIml As Long, _
        ByVal iiconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lSize As Long _
    )
Dim tR As RECT
Dim hBrush As Long
Dim lStyle As Long

    ' Firstly, create the mask & image:
    ' Draw the image into the top square of the mono DC:
    BitBlt m_ILDMonoHDC, 0, 0, lSize, lSize * 3, m_ILDMonoHDC, 0, 0, WHITENESS
    
    lStyle = ILD_IMAGE
    ImageList_Draw hIml, iiconIndex, m_ILDMonoHDC, 0, 0, lStyle
    ' Draw the Mask into the second square:
    lStyle = ILD_MASK
    ImageList_Draw hIml, iiconIndex, m_ILDMonoHDC, 0, lSize, lStyle
    ' Or the mask & mono image together:
    BitBlt m_ILDMonoHDC, 0, 0, lSize, lSize, m_ILDMonoHDC, 0, lSize, SRCPAINT
    ' Invert the thing:
   'BitBlt m_ILDMonoHDC, 0, lSize * 2, lSize, lSize, m_ILDMonoHDC, 0, 0, WHITENESS
    BitBlt m_ILDMonoHDC, 0, lSize * 2, lSize, lSize, m_ILDMonoHDC, 0, lSize, SRCINVERT

    ' Now create white & button shadow copies of it:
    BitBlt m_ILDColorHDC, 0, 0, lSize, lSize, m_ILDMonoHDC, 0, lSize * 2, SRCCOPY
    hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNSHADOW))
    tR.left = 0
    tR.Right = lSize
    tR.tOp = lSize
    tR.Bottom = lSize * 2
    FillRect m_ILDColorHDC, tR, hBrush
    DeleteObject hBrush
    BitBlt m_ILDColorHDC, 0, lSize, lSize, lSize, m_ILDMonoHDC, 0, lSize * 2, SRCAND
    BitBlt m_ILDColorHDC, 0, lSize, lSize, lSize, m_ILDMonoHDC, 0, lSize, SRCPAINT
        
    ' Finally, we blit the disabled verson to the DC:
    ' Draw white version, offset by 1 pixel in x & y:
    BitBlt lHDC, lX + 1, lY + 1, lSize - 1, lSize - 1, m_ILDColorHDC, 0, 0, SRCPAINT
    ' Draw mask for dark version:
    BitBlt lHDC, lX, lY, lSize, lSize, m_ILDColorHDC, 0, 0, SRCPAINT
    ' Finally draw the button shadow version:
    BitBlt lHDC, lX, lY, lSize, lSize, m_ILDColorHDC, 0, lSize, SRCAND

End Sub

Private Sub Class_Terminate()
    Destroy
End Sub
