Attribute VB_Name = "modGraphics"
Option Explicit

Public Const SRCCOPY = &HCC0020
Public Const S_OK As Long = &H0

Public Declare Function DeleteObject Lib "GDI32.dll" (ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "GDI32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Public Declare Function OleCreatePictureIndirect Lib "OLEPro32.dll" (ByRef PicDesc As Any, ByRef RefIID As Guid, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
Public Declare Function CreateIconIndirect Lib "User32.dll" (ByRef pIconInfo As IconInfo) As Long
Public Declare Function DrawIcon Lib "User32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function CreateCompatibleDC Lib "GDI32.dll" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "GDI32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "GDI32.dll" (ByVal hDC As Long) As Long
Public Declare Function SetTextColor Lib "GDI32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkColor Lib "GDI32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "GDI32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Public Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type PictDescIcon
    cbSizeOfStruct As Long
    picType As Long
    hIcon As Long
End Type

Public Type IconInfo
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Public Type Bitmap ' 24 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
'

'  Chroma-key transparent blit 1.1
' Written by Mike D Sutton of EDais
'     Microsoft Visual Basic MVP
'
' E-Mail: EDais@mvps.org
' WWW: Http://www.mvps.org/EDais/
'
' Written: 14/03/2002
' Last edited: 19/04/2002
'
'About:
' TransparentBlt() without the memory leaks ;)
'
'Version history:
' Version 1.1 (17/08/2003):
'   Re-wrote ChromaBlt() to use the new masking function (Below)
'   Added Back-buffered draw option to ChromaBlt()
'
'   GetColMask() - Generates a 1-bpp mask bitmap, based on an
'                  existing image and mask colour.
'
' Version 1.02 (19/04/2002):
'   Fixed another small bug when drawing on a DC with a non-black
'   foreground colour.
'
' Version 1.01 (15/03/2002):
'   Fixed small problem in ChromaBlt() where it wasn't rendering
'   the mask correctly over other images.
'
' Version 1.0 (14/03/2002):
'   ChromaBlt() - Blit's an image onto a DC, excluding a defined
'                 'transparent' colour.
'
'You use this code at your own risk, I don't accept any
' responsibility for anything nasty it may do to your machine!
'
'Please don't rip my work off...  I'm distributing this library
' free of charge because I think it can help other developers,
' this doesn't give you the right to take credit for it.  By
' all means use it, yes, but please don't claim it's your own
' work or charge for it.  If you do create anything interesting
' with it then feel free to send me it, if I receive any nice
' source code I'll post it on the site (With your permission)
' and of course you'll get full credit for it.
'
'Visit my site for any updates to this an more strange graphics
' related VB code, comments and suggestions always welcome!

Public Function GetColMask(ByVal inDC As Long, _
                           ByVal inX As Long, _
                           ByVal inY As Long, _
                           ByVal inWidth As Long, _
                           ByVal inHeight As Long, _
                           ByVal inMaskCol As Long) As Long
    
    Dim MaskDC As Long, MaskBMP As Long, OldMask As Long
    Dim OldBack As Long
    
    ' Make sure the input sizes are valid
    If ((inWidth < 1) Or (inHeight < 1)) Then Exit Function
    
    ' Create a new DC
    MaskDC = CreateCompatibleDC(inDC)
    
    If (MaskDC) Then ' Create a new 1-bpp Bitmap (DDB)
        MaskBMP = CreateBitmap(inWidth, inHeight, 1, 1, ByVal 0&)
        
        If (MaskBMP) Then ' Select Bitmap into DC
            OldMask = SelectObject(MaskDC, MaskBMP)
            
            If (OldMask) Then ' Set mask colour
                OldBack = SetBkColor(inDC, inMaskCol)
                
                ' Generate mask image
                If (BitBlt(MaskDC, 0, 0, inWidth, inHeight, inDC, _
                    inX, inY, vbSrcCopy) <> 0) Then GetColMask = MaskBMP
                
                ' Clean up
                Call SetBkColor(inDC, OldBack)
                Call SelectObject(MaskDC, OldMask)
            End If
            
            ' Something went wrong, destroy mask Bitmap
            If (GetColMask = 0) Then Call DeleteObject(MaskBMP)
        End If
        
        ' Destroy temporary DC
        Call DeleteDC(MaskDC)
    End If
End Function

Public Function ChromaBlt(ByVal outDC As Long, _
                          ByVal inX As Long, _
                          ByVal inY As Long, _
                          ByVal outWidth As Long, _
                          ByVal outHeight As Long, _
                          ByVal inSrcDC As Long, _
                          ByVal inSrcX As Long, _
                          ByVal inSrcY As Long, _
                          ByVal inSrcWidth As Long, _
                          ByVal inSrcHeight As Long, _
                          ByVal inChromaKey As Long, _
                          Optional ByVal inDoubleBuffer As Boolean = False) As Boolean
    
    Dim MaskDC As Long, MaskBMP As Long, OldMaskBMP As Long
    Dim SpriteDC As Long, SpriteBMP As Long, OldSpriteBMP As Long
    Dim OldFGCol As Long, OldBkCol As Long
    
    If (inDoubleBuffer) Then
        ' Create a back-buffer to perform drawing to
        SpriteDC = CreateCompatibleDC(0)
        
        If (SpriteDC) Then
            SpriteBMP = CreateCompatibleBitmap(outDC, outWidth, outHeight)
            
            If (SpriteBMP) Then
                OldSpriteBMP = SelectObject(SpriteDC, SpriteBMP)
                
                If (OldSpriteBMP) Then ' Prepare back-buffer and copy background to it
                    Call BitBlt(SpriteDC, 0, 0, outWidth, outHeight, outDC, inX, inY, vbSrcCopy)
                    Call SetStretchBltMode(SpriteDC, GetStretchBltMode(outDC)) ' Sync. stretch blit modes
                    
                    ' Re-call the routine and draw to the back-buffer instead
                    ChromaBlt = ChromaBlt(SpriteDC, 0, 0, outWidth, outHeight, inSrcDC, _
                        inSrcX, inSrcY, inSrcWidth, inSrcHeight, inChromaKey, False)
                    
                    ' Draw result to destination DC, clean up and quit
                    If (ChromaBlt) Then Call BitBlt(outDC, inX, inY, outWidth, outHeight, SpriteDC, 0, 0, vbSrcCopy)
                    Call SelectObject(SpriteDC, OldSpriteBMP)
                End If
                
                Call DeleteObject(SpriteBMP)
            End If
            
            Call DeleteDC(SpriteDC)
        End If
        
        Exit Function
    End If
    
    ' Generate mask bitmap based on existing DC and mask colour
    MaskBMP = GetColMask(inSrcDC, inSrcX, inSrcY, inSrcWidth, inSrcHeight, inChromaKey)
    If (MaskBMP = 0) Then Exit Function ' Mask genration failed
    
    ' Create temp DC's
    MaskDC = CreateCompatibleDC(outDC)
    SpriteDC = CreateCompatibleDC(outDC)
    
    If ((MaskDC <> 0) And (SpriteDC <> 0)) Then ' Create temp Bitmap's
        SpriteBMP = CreateCompatibleBitmap(outDC, inSrcWidth, inSrcHeight)
        
        If (SpriteBMP <> 0) Then ' Select Bitmap's into DC's
            OldMaskBMP = SelectObject(MaskDC, MaskBMP)
            OldSpriteBMP = SelectObject(SpriteDC, SpriteBMP)
            
            If ((OldMaskBMP <> 0) And (OldSpriteBMP <> 0)) Then ' All set up
                Call SetBkColor(SpriteDC, inChromaKey)
                
                ' Make copy of existing image
                Call BitBlt(SpriteDC, 0, 0, inSrcWidth, inSrcHeight, inSrcDC, inSrcX, inSrcY, vbSrcCopy)
                
                ' Draw inverted mask over sprite image
                Call BitBlt(SpriteDC, 0, 0, inSrcWidth, inSrcHeight, MaskDC, 0, 0, vbSrcInvert)
                
                ' Temporarily set foreground and background colours of
                ' render DC to black and white respectively (For mask)
                OldFGCol = SetTextColor(outDC, vbBlack)
                OldBkCol = SetBkColor(outDC, vbWhite)
                
                ' Composite mask onto existing image and overlay sprite
                If ((inSrcWidth = outWidth) And (inSrcHeight = outHeight)) Then ' Straight blit
                    Call BitBlt(outDC, inX, inY, inSrcWidth, inSrcHeight, MaskDC, 0, 0, vbSrcAnd)
                    Call BitBlt(outDC, inX, inY, inSrcWidth, inSrcHeight, SpriteDC, 0, 0, vbSrcInvert)
                Else ' StretchBlt() up to desired size
                    Call StretchBlt(outDC, inX, inY, outWidth, outHeight, MaskDC, 0, 0, inSrcWidth, inSrcHeight, vbSrcAnd)
                    Call StretchBlt(outDC, inX, inY, outWidth, outHeight, SpriteDC, 0, 0, inSrcWidth, inSrcHeight, vbSrcInvert)
                End If
                
                ' Re-set colours on render DC
                Call SetTextColor(outDC, OldFGCol)
                Call SetBkColor(outDC, OldBkCol)
                
                ' De-select objects
                Call SelectObject(MaskDC, OldMaskBMP)
                Call SelectObject(SpriteDC, OldSpriteBMP)
                
                ' If it got this far then all went well
                ChromaBlt = True
            End If
        End If
        
        ' Clean up temp sprite Bitmap
        Call DeleteObject(SpriteBMP)
    End If
    
    ' Clean up mask Bitmap
    Call DeleteObject(MaskBMP)
    
    ' Clean up temp DC's
    Call DeleteDC(MaskDC)
    Call DeleteDC(SpriteDC)
End Function

' Take a HBITMAP and return an HICON
Public Function BitmapToIcon(ByVal inBMP As Long, _
                             Optional ByVal inTransCol As Long = vbBlack) As Long
    Dim IconInf As IconInfo
    Dim BMInf As Bitmap
    Dim hSrcDC As Long, hSrcBMP As Long, hSrcOldBMP As Long
    Dim hMaskDC As Long, hMaskBMP As Long, hMaskOldBMP As Long

    ' Requires:
    '   modChromaBlt  - For mask extraction
    '   modCopyBitmap - For DDB copy

    ' Get some information about this Bitmap and create a mask the same size
    If (GetObject(inBMP, Len(BMInf), BMInf) = 0) Then Exit Function

    ' Create a copy of the original Bitmap as a DDB that we can play with
    hSrcBMP = CopyBitmap(inBMP, cbmDDB)

    ' Create DC's and select source copy
    hSrcDC = CreateCompatibleDC(0)
    hMaskDC = CreateCompatibleDC(0)
    hSrcOldBMP = SelectObject(hSrcDC, hSrcBMP)

    If (hSrcOldBMP) Then ' Extract a colour mask from source copy
        hMaskBMP = GetColMask(hSrcDC, 0, 0, BMInf.bmWidth, BMInf.bmHeight, inTransCol)
        hMaskOldBMP = SelectObject(hMaskDC, hMaskBMP)

        If (hMaskOldBMP) Then ' Overlay inverted mask over source
            Call SetTextColor(hSrcDC, vbWhite)
            Call SetBkColor(hSrcDC, vbBlack)
            Call BitBlt(hSrcDC, 0, 0, BMInf.bmWidth, BMInf.bmHeight, hMaskDC, 0, 0, vbSrcAnd)
            Call SelectObject(hMaskDC, hMaskOldBMP) ' De-select mask
        End If

        ' De-select source copy
        Call SelectObject(hSrcDC, hSrcOldBMP)
    End If

    ' Destroy DC's
    Call DeleteDC(hMaskDC)
    Call DeleteDC(hSrcDC)

    With IconInf ' Set some information about the icon
        .fIcon = True
        .hbmMask = hMaskBMP
        .hbmColor = hSrcBMP
    End With

    ' Create the icon and destroy the temp mask
    BitmapToIcon = CreateIconIndirect(IconInf)

    ' Destroy interim Bitmaps
    Call DeleteObject(hMaskBMP)
    Call DeleteObject(hSrcBMP)
End Function

