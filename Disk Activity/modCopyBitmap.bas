Attribute VB_Name = "modCopyBitmap"
Option Explicit
' Voir http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=64964&lngWId=1
    
Public Declare Function BitBlt Lib "GDI32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "GDI32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BitmapInfo, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "GDI32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BitmapInfo, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetDesktopWindow Lib "User32.dll" () As Long
Private Declare Function GetDC Lib "User32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "User32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Type BitmapInfoHeader ' 40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BitmapInfo
    bmiHeader As BitmapInfoHeader
    bmiColors(255) As Long
End Type

Private Type Bitmap ' 24 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Enum enCopyBMPMode
    cbmSame = &H0
    cbmDDB = &H1
    cbmDIB = &H2
End Enum

Public Enum enCopyBMPStretchMode
    cbsmBlackOnWhite = &H1
    cbsmWhiteOnBlack = &H2
    cbsmColourOnColour = &H3
    cbsmHalftone = &H4
End Enum

Private Const DIB_RGB_COLORS As Long = &H0
Private Const BI_BITFIELDS As Long = &H3
'

Public Function CopyBitmap(ByVal inBMP As Long, _
                           Optional ByVal inCopyMode As enCopyBMPMode = cbmSame, _
                           Optional ByVal inNewWidth As Long = 0, _
                           Optional ByVal inNewHeight As Long = 0, _
                           Optional ByVal inStretchMode As enCopyBMPStretchMode = cbsmHalftone) As Long
    Dim BMInf As Bitmap
    Dim hSrcDC As Long, hSrcOldBMP As Long
    Dim hDstDC As Long, hDstBMP As Long, hDstOldBMP As Long
    Dim DIBInf As BitmapInfo
    Dim NumCol As Long
    Dim DeskWnd As Long, DeskDC As Long
    Dim UseWidth As Long, UseHeight As Long
    Dim BlitRet As Long
    
    ' Get some information about the source Bitmap
    If (GetObject(inBMP, Len(BMInf), BMInf) = 0) Then Exit Function
    
    ' Create some surfaces and select source Bitmap
    hSrcDC = CreateCompatibleDC(0)
    hDstDC = CreateCompatibleDC(0)
    hSrcOldBMP = SelectObject(hSrcDC, inBMP)
    
    ' Get size of result Bitmap
    If (inNewWidth) Then UseWidth = Abs(inNewWidth) Else UseWidth = BMInf.bmWidth
    If (inNewHeight) Then UseHeight = Abs(inNewHeight) Else UseHeight = BMInf.bmHeight
    
    If (hSrcOldBMP) Then ' If we're matching the format, select which the source was
        If (inCopyMode = cbmSame) Then inCopyMode = IIf(BMInf.bmBits, cbmDIB, cbmDDB)
        
        If (inCopyMode = cbmDIB) Then ' Set DIB header size
            DIBInf.bmiHeader.biSize = Len(DIBInf.bmiHeader)
            
            ' Get DIB header information from source
            If (GetDIBits(hSrcDC, inBMP, 0, 0, ByVal 0&, DIBInf, DIB_RGB_COLORS)) Then
                If (((DIBInf.bmiHeader.biBitCount > 0) And _
                    (DIBInf.bmiHeader.biBitCount <= 8)) Or _
                    (DIBInf.bmiHeader.biCompression = BI_BITFIELDS)) Then
                    NumCol = DIBInf.bmiHeader.biClrUsed ' Get palette information
                    Call GetDIBits(hSrcDC, inBMP, 0, 0, ByVal 0&, DIBInf, DIB_RGB_COLORS)
                    DIBInf.bmiHeader.biClrUsed = NumCol
                End If
                
                With DIBInf.bmiHeader
                    .biWidth = UseWidth
                    .biHeight = UseHeight ' Fill DIB header with any new information about the size
                    .biSizeImage = ((((UseWidth * .biBitCount) + &H1F) And Not &H1F) \ &H8) * UseHeight
                End With
                
                ' Create new DIB with same header as source
                hDstBMP = CreateDIBSection(hDstDC, DIBInf, DIB_RGB_COLORS, 0, 0, 0)
            End If
        Else
            DeskWnd = GetDesktopWindow()
            DeskDC = GetDC(DeskWnd) ' Create new DDB compatible with the screen
            hDstBMP = CreateCompatibleBitmap(DeskDC, UseWidth, UseHeight)
            Call ReleaseDC(DeskWnd, DeskDC)
        End If
        
        ' Select new Bitmap into destination DC
        hDstOldBMP = SelectObject(hDstDC, hDstBMP)
        
        If (hDstOldBMP) Then ' Copy source to destination
            If ((UseWidth = BMInf.bmWidth) And (UseHeight = BMInf.bmHeight)) Then ' Blit
                BlitRet = BitBlt(hDstDC, 0, 0, BMInf.bmWidth, BMInf.bmHeight, hSrcDC, 0, 0, vbSrcCopy)
            Else ' Stretch-blit
                Call SetStretchBltMode(hDstDC, inStretchMode)
                BlitRet = StretchBlt(hDstDC, 0, 0, UseWidth, UseHeight, _
                    hSrcDC, 0, 0, BMInf.bmWidth, BMInf.bmHeight, vbSrcCopy)
            End If
            
            If (BlitRet) Then ' Return copied Bitmap to caller
                CopyBitmap = hDstBMP
            Else ' Copy failed, destroy destination Bitmap
                Call DeleteObject(hDstBMP)
            End If
            
            ' De-select destination Bitmap
            Call SelectObject(hDstDC, hDstOldBMP)
        Else ' Something went wrong
            Call DeleteObject(hDstBMP)
        End If
        
        ' De-select source
        Call SelectObject(hSrcDC, hSrcOldBMP)
    End If
    
    ' Destroy surfaces
    Call DeleteDC(hDstDC)
    Call DeleteDC(hSrcDC)
End Function
