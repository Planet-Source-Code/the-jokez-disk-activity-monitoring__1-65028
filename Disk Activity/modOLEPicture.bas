Attribute VB_Name = "modOLEPicture"
Option Explicit
' Voir http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=64964&lngWId=1

Private Declare Function OleCreatePictureIndirect Lib "OLEPro32.dll" (ByRef PicDesc As Any, ByRef RefIID As Guid, ByVal fPictureOwnsHandle As Long, ByRef IPic As IPicture) As Long
Private Declare Function GetObjectType Lib "GDI32.dll" (ByVal hGDIObj As Long) As Long
Private Declare Function GetIconInfo Lib "User32.dll" (ByVal hIcon As Long, ByRef pIconInfo As IconInfo) As Long
Private Declare Function GetMetaFileBitsEx Lib "GDI32.dll" (ByVal hMF As Long, ByVal nSize As Long, ByRef lpvData As Any) As Long
Private Declare Function SetWinMetaFileBits Lib "GDI32.dll" (ByVal cbBuffer As Long, ByRef lpbBuffer As Byte, ByVal hDCRef As Long, lpMFP As MetaFilePict) As Long
Private Declare Function GetEnhMetaFileHeader Lib "GDI32.dll" (ByVal hEMF As Long, ByVal cbBuffer As Long, ByRef lpEMH As EnhMetaHeader) As Long
Private Declare Function DeleteEnhMetaFile Lib "GDI32.dll" (ByVal hEMF As Long) As Long

Private Declare Function CreateEnhMetaFile Lib "GDI32.dll" Alias "CreateEnhMetaFileA" (ByVal hDCRef As Long, ByVal lpFileName As String, ByRef lpRect As Any, ByVal lpDescription As String) As Long
Private Declare Function CloseEnhMetaFile Lib "GDI32.dll" (ByVal hDC As Long) As Long
Private Declare Function PlayMetaFile Lib "GDI32.dll" (ByVal hDC As Long, ByVal hMF As Long) As Long

Private Type PictDescGeneirc
    pdgSize As Long
    pdcPicType As Long
    pdcHandle As Long
    pdcExtraA As Long ' xExt for metafile, hPal for Bitmap
    pdcExtraB As Long ' yExt for metafile
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type IconInfo
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Type MetaFilePict
    mm As Long
    xExt As Long
    yExt As Long
    hMF As Long
End Type

Private Type RectL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SizeL
    cx As Long
    cy As Long
End Type

Private Type EnhMetaHeader
    iType As Long
    nSize As Long
    rclBounds As RectL
    rclFrame As RectL
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As SizeL
    szlMillimeters As SizeL
End Type

Private Const OBJ_BITMAP As Long = &H7
Private Const OBJ_METAFILE As Long = &H9
Private Const OBJ_ENHMETAFILE As Long = &HD

Private Const PICTYPE_BITMAP As Long = &H1
Private Const PICTYPE_METAFILE As Long = &H2
Private Const PICTYPE_ICON As Long = &H3
Private Const PICTYPE_ENHMETAFILE As Long = &H4

Private Const S_OK As Long = &H0

Public Function GDIToPicture(ByVal inGDIObj As Long, _
    Optional ByVal inOwnObj As Boolean = True, _
    Optional ByVal inPal As Long = &H0) As IPicture
    Dim IconInf As IconInfo
    Dim PicDesc As PictDescGeneirc
    Dim PictureGUID As Guid
    Dim RetPic As IPicture
    Dim TempEMF As Long
    Dim MetaHead As EnhMetaHeader
    
    Select Case GetObjectType(inGDIObj)
        Case OBJ_BITMAP
            PicDesc.pdgSize = 16
            PicDesc.pdcPicType = PICTYPE_BITMAP
            PicDesc.pdcExtraA = inPal
        Case OBJ_METAFILE ' UNTESTED!
            PicDesc.pdgSize = 20
            PicDesc.pdcPicType = PICTYPE_METAFILE
            
            ' WMF objects don't store bounds information so perform
            ' temporary conversion to EMF and read header structure
            TempEMF = WMFToEMF(inGDIObj)
            If (TempEMF) Then
                Call GetEnhMetaFileHeader(TempEMF, Len(MetaHead), MetaHead)
                PicDesc.pdcExtraA = MetaHead.rclBounds.Right
                PicDesc.pdcExtraB = MetaHead.rclBounds.Bottom
                Call DeleteEnhMetaFile(TempEMF)
            End If
        Case OBJ_ENHMETAFILE
            PicDesc.pdgSize = 12
            PicDesc.pdcPicType = PICTYPE_ENHMETAFILE
        Case Else ' Test for icon/cursor
            If (GetIconInfo(inGDIObj, IconInf)) Then
                PicDesc.pdgSize = 12
                PicDesc.pdcPicType = PICTYPE_ICON
                
                ' Clean up Bitmap copies
                Call DeleteObject(IconInf.hbmColor)
                Call DeleteObject(IconInf.hbmMask)
            End If
    End Select
    
    ' Couldn't match this object against known types
    If (PicDesc.pdgSize = 0) Then Exit Function
    
    ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With PictureGUID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(3) = &HAA
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    ' Set object handle
    PicDesc.pdcHandle = inGDIObj
    
    If (OleCreatePictureIndirect(PicDesc, PictureGUID, _
        inOwnObj, RetPic) = S_OK) Then Set GDIToPicture = RetPic
    Set RetPic = Nothing
End Function

Private Function WMFToEMF(ByVal inWMF As Long) As Long
    Dim EMetaDC As Long
    
    ' Create a new Enhanced metafile device context
    EMetaDC = CreateEnhMetaFile(0, vbNullString, ByVal 0&, vbNullString)
    Call PlayMetaFile(EMetaDC, inWMF)
    WMFToEMF = CloseEnhMetaFile(EMetaDC)
    
    If (WMFToEMF = 0) Then ' If first method fails, try copy method
        Dim WMFSize As Long, WMFData() As Byte
        Dim MetaInf As MetaFilePict
        
        ' Query WMF data size
        WMFSize = GetMetaFileBitsEx(inWMF, 0, ByVal 0&)
        If (WMFSize) Then ' Allocate data buffer and extract WMF data
            ReDim WMFData(WMFSize - 1) As Byte
            Call GetMetaFileBitsEx(inWMF, WMFSize, WMFData(0))
            
            MetaInf.hMF = inWMF ' Convert WMF data to EMF
            WMFToEMF = SetWinMetaFileBits(WMFSize, WMFData(0), 0, MetaInf)
        End If
    End If
End Function
