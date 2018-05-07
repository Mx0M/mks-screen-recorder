Public Class ScreenCapture
    '#include-once

    ' #INDEX# =======================================================================================================================
    ' Title .........: GDIPlus_Constants
    ' AutoIt Version : 3.2
    ' Language ......: English
    ' Description ...: Constants for GDI+
    ' Author(s) .....: Valik, Gary Frost
    ' ===============================================================================================================================

    ' #CONSTANTS# ===================================================================================================================
    ' Pen Dash Cap Types
    Public Const GDIP_DASHCAPFLAT As Integer = 0 ' A square cap that squares off both ends of each dash
    Public Const GDIP_DASHCAPROUND As Integer = 2 ' A circular cap that rounds off both ends of each dash
    Public Const GDIP_DASHCAPTRIANGLE As Integer = 3 ' A triangular cap that points both ends of each dash

    ' Pen Dash Style Types
    Public Const GDIP_DASHSTYLESOLID As Integer = 0 ' A solid line
    Public Const GDIP_DASHSTYLEDASH As Integer = 1 ' A dashed line
    Public Const GDIP_DASHSTYLEDOT As Integer = 2 ' A dotted line
    Public Const GDIP_DASHSTYLEDASHDOT As Integer = 3 ' An alternating dash-dot line
    Public Const GDIP_DASHSTYLEDASHDOTDOT As Integer = 4 ' An alternating dash-dot-dot line
    Public Const GDIP_DASHSTYLECUSTOM As Integer = 5 ' A a user-defined, custom dashed line

    ' Enocder Parameter GUIDs
    Public Const GDIP_EPGCHROMINANCETABLE As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
    Public Const GDIP_EPGCOLORDEPTH As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
    Public Const GDIP_EPGCOMPRESSION As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
    Public Const GDIP_EPGLUMINANCETABLE As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
    Public Const GDIP_EPGQUALITY As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
    Public Const GDIP_EPGRENDERMETHOD As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
    Public Const GDIP_EPGSAVEFLAG As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
    Public Const GDIP_EPGSCANMETHOD As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
    Public Const GDIP_EPGTRANSFORMATION As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
    Public Const GDIP_EPGVERSION As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"

    ' Encoder Parameter Types
    Public Const GDIP_EPTBYTE As Integer = 1 ' 8 bit unsigned integer
    Public Const GDIP_EPTASCII As Integer = 2 ' Null terminated character string
    Public Const GDIP_EPTSHORT As Integer = 3 ' 16 bit unsigned integer
    Public Const GDIP_EPTLONG As Integer = 4 ' 32 bit unsigned integer
    Public Const GDIP_EPTRATIONAL As Integer = 5 ' Two longs (numerator, denomintor)
    Public Const GDIP_EPTLONGRANGE As Integer = 6 ' Two longs (low, high)
    Public Const GDIP_EPTUNDEFINED As Integer = 7 ' Array of bytes of any type
    Public Const GDIP_EPTRATIONALRANGE As Integer = 8 ' Two ratationals (low, high)

    ' GDI Error Codes
    Public Const GDIP_ERROK As Integer = 0 ' Method call was successful
    Public Const GDIP_ERRGENERICERROR As Integer = 1 ' Generic method call error
    Public Const GDIP_ERRINVALIDPARAMETER As Integer = 2 ' One of the arguments passed to the method was not valid
    Public Const GDIP_ERROUTOFMEMORY As Integer = 3 ' The operating system is out of memory
    Public Const GDIP_ERROBJECTBUSY As Integer = 4 ' One of the arguments in the call is already in use
    Public Const GDIP_ERRINSUFFICIENTBUFFER As Integer = 5 ' A buffer is not large enough
    Public Const GDIP_ERRNOTIMPLEMENTED As Integer = 6 ' The method is not implemented
    Public Const GDIP_ERRWIN32ERROR As Integer = 7 ' The method generated a Microsoft Win32 error
    Public Const GDIP_ERRWRONGSTATE As Integer = 8 ' The object is in an invalid state to satisfy the API call
    Public Const GDIP_ERRABORTED As Integer = 9 ' The method was aborted
    Public Const GDIP_ERRFILENOTFOUND As Integer = 10 ' The specified image file or metafile cannot be found
    Public Const GDIP_ERRVALUEOVERFLOW As Integer = 11 ' The method produced a numeric overflow
    Public Const GDIP_ERRACCESSDENIED As Integer = 12 ' A write operation is not allowed on the specified file
    Public Const GDIP_ERRUNKNOWNIMAGEFORMAT As Integer = 13 ' The specified image file format is not known
    Public Const GDIP_ERRFONTFAMILYNOTFOUND As Integer = 14 ' The specified font family cannot be found
    Public Const GDIP_ERRFONTSTYLENOTFOUND As Integer = 15 ' The specified style is not available for the specified font
    Public Const GDIP_ERRNOTTRUETYPEFONT As Integer = 16 ' The font retrieved is not a TrueType font
    Public Const GDIP_ERRUNSUPPORTEDGDIVERSION As Integer = 17 ' The installed GDI+ version is incompatible
    Public Const GDIP_ERRGDIPLUSNOTINITIALIZED As Integer = 18 ' The GDI+ API is not in an initialized state
    Public Const GDIP_ERRPROPERTYNOTFOUND As Integer = 19 ' The specified property does not exist in the image
    Public Const GDIP_ERRPROPERTYNOTSUPPORTED As Integer = 20 ' The specified property is not supported

    ' Encoder Value Types
    Public Const GDIP_EVTCOMPRESSIONLZW As Integer = 2 ' TIFF: LZW compression
    Public Const GDIP_EVTCOMPRESSIONCCITT3 As Integer = 3 ' TIFF: CCITT3 compression
    Public Const GDIP_EVTCOMPRESSIONCCITT4 As Integer = 4 ' TIFF: CCITT4 compression
    Public Const GDIP_EVTCOMPRESSIONRLE As Integer = 5 ' TIFF: RLE compression
    Public Const GDIP_EVTCOMPRESSIONNONE As Integer = 6 ' TIFF: No compression
    Public Const GDIP_EVTTRANSFORMROTATE90 As Integer = 13 ' JPEG: Lossless 90 degree clockwise rotation
    Public Const GDIP_EVTTRANSFORMROTATE180 As Integer = 14 ' JPEG: Lossless 180 degree clockwise rotation
    Public Const GDIP_EVTTRANSFORMROTATE270 As Integer = 15 ' JPEG: Lossless 270 degree clockwise rotation
    Public Const GDIP_EVTTRANSFORMFLIPHORIZONTAL As Integer = 16 ' JPEG: Lossless horizontal flip
    Public Const GDIP_EVTTRANSFORMFLIPVERTICAL As Integer = 17 ' JPEG: Lossless vertical flip
    Public Const GDIP_EVTMULTIFRAME As Integer = 18 ' Multiple frame encoding
    Public Const GDIP_EVTLASTFRAME As Integer = 19 ' Last frame of a multiple frame image
    Public Const GDIP_EVTFLUSH As Integer = 20 ' The encoder object is to be closed
    Public Const GDIP_EVTFRAMEDIMENSIONPAGE As Integer = 23 ' TIFF: Page frame dimension

    ' Image Codec Flags constants
    Public Const GDIP_ICFENCODER As Integer = &H1 ' The codec supports encoding (saving)
    Public Const GDIP_ICFDECODER As Integer = &H2 ' The codec supports decoding (reading)
    Public Const GDIP_ICFSUPPORTBITMAP As Integer = &H4 ' The codec supports raster images (bitmaps)
    Public Const GDIP_ICFSUPPORTVECTOR As Integer = &H8 ' The codec supports vector images (metafiles)
    Public Const GDIP_ICFSEEKABLEENCODE As Integer = &H10 ' The encoder requires a seekable output stream
    Public Const GDIP_ICFBLOCKINGDECODE As Integer = &H20 ' The decoder has blocking behavior during the decoding process
    Public Const GDIP_ICFBUILTIN As Integer = &H10000 ' The codec is built in to GDI+
    Public Const GDIP_ICFSYSTEM As Integer = &H20000 ' Not used in GDI+ version 1.0
    Public Const GDIP_ICFUSER As Integer = &H40000 ' Not used in GDI+ version 1.0

    ' Image Lock Mode constants
    Public Const GDIP_ILMREAD As Integer = &H1 ' A portion of the image is locked for reading
    Public Const GDIP_ILMWRITE As Integer = &H2 ' A portion of the image is locked for writing
    Public Const GDIP_ILMUSERINPUTBUF As Integer = &H4 ' The buffer is allocated by the user

    ' LineCap constants
    Public Const GDIP_LINECAPFLAT As Integer = &H0 ' Specifies a flat cap
    Public Const GDIP_LINECAPSQUARE As Integer = &H1 ' Specifies a square cap
    Public Const GDIP_LINECAPROUND As Integer = &H2 ' Specifies a circular cap
    Public Const GDIP_LINECAPTRIANGLE As Integer = &H3 ' Specifies a triangular cap
    Public Const GDIP_LINECAPNOANCHOR As Integer = &H10 ' Specifies that the line ends are not anchored
    Public Const GDIP_LINECAPSQUAREANCHOR As Integer = &H11 ' Specifies that the line ends are anchored with a square
    Public Const GDIP_LINECAPROUNDANCHOR As Integer = &H12 ' Specifies that the line ends are anchored with a circle
    Public Const GDIP_LINECAPDIAMONDANCHOR As Integer = &H13 ' Specifies that the line ends are anchored with a diamond
    Public Const GDIP_LINECAPARROWANCHOR As Integer = &H14 ' Specifies that the line ends are anchored with arrowheads
    Public Const GDIP_LINECAPCUSTOM As Integer = &HFF ' Specifies that the line ends are made from a CustomLineCap

    ' Pixel Format constants
    Public Const GDIP_PXF01INDEXED As Integer = &H30101 ' 1 bpp, indexed
    Public Const GDIP_PXF04INDEXED As Integer = &H30402 ' 4 bpp, indexed
    Public Const GDIP_PXF08INDEXED As Integer = &H30803 ' 8 bpp, indexed
    Public Const GDIP_PXF16GRAYSCALE As Integer = &H101004 ' 16 bpp, grayscale
    Public Const GDIP_PXF16RGB555 As Integer = &H21005 ' 16 bpp' 5 bits for each RGB
    Public Const GDIP_PXF16RGB565 As Integer = &H21006 ' 16 bpp' 5 bits red, 6 bits green, and 5 bits blue
    Public Const GDIP_PXF16ARGB1555 As Integer = &H61007 ' 16 bpp' 1 bit for alpha and 5 bits for each RGB component
    Public Const GDIP_PXF24RGB As Integer = &H21808 ' 24 bpp' 8 bits for each RGB
    Public Const GDIP_PXF32RGB As Integer = &H22009 ' 32 bpp' 8 bits for each RGB. No alpha.
    Public Const GDIP_PXF32ARGB As Integer = &H26200A ' 32 bpp' 8 bits for each RGB and alpha
    Public Const GDIP_PXF32PARGB As Integer = &HD200B ' 32 bpp' 8 bits for each RGB and alpha, pre-mulitiplied
    Public Const GDIP_PXF48RGB As Integer = &H10300C ' 48 bpp' 16 bits for each RGB
    Public Const GDIP_PXF64ARGB As Integer = &H34400D ' 64 bpp' 16 bits for each RGB and alpha
    Public Const GDIP_PXF64PARGB As Integer = &H1C400E ' 64 bpp' 16 bits for each RGB and alpha, pre-multiplied

    ' ImageFormat constants (Publicly Unique Identifier (GUID))
    Public Const GDIP_IMAGEFORMAT_UNDEFINED As String = "{B96B3CA9-0728-11D3-9D7B-0000F81EF32E}" ' Windows GDI+ is unable to determine the format.
    Public Const GDIP_IMAGEFORMAT_MEMORYBMP As String = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}" ' Image was constructed from a memory bitmap.
    Public Const GDIP_IMAGEFORMAT_BMP As String = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}" ' Microsoft Windows Bitmap (BMP) format.
    Public Const GDIP_IMAGEFORMAT_EMF As String = "{B96B3CAC-0728-11D3-9D7B-0000F81EF32E}" ' Enhanced Metafile (EMF) format.
    Public Const GDIP_IMAGEFORMAT_WMF As String = "{B96B3CAD-0728-11D3-9D7B-0000F81EF32E}" ' Windows Metafile Format (WMF) format.
    Public Const GDIP_IMAGEFORMAT_JPEG As String = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}" ' Joint Photographic Experts Group (JPEG) format.
    Public Const GDIP_IMAGEFORMAT_PNG As String = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}" ' Portable Network Graphics (PNG) format.
    Public Const GDIP_IMAGEFORMAT_GIF As String = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}" ' Graphics Interchange Format (GIF) format.
    Public Const GDIP_IMAGEFORMAT_TIFF As String = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}" ' Tagged Image File Format (TIFF) format.
    Public Const GDIP_IMAGEFORMAT_EXIF As String = "{B96B3CB2-0728-11D3-9D7B-0000F81EF32E}" ' Exchangeable Image File (EXIF) format.
    Public Const GDIP_IMAGEFORMAT_ICON As String = "{B96B3CB5-0728-11D3-9D7B-0000F81EF32E}" ' Microsoft Windows Icon Image (ICO)format.

    ' ImageType constants
    Public Const GDIP_IMAGETYPE_UNKNOWN As Integer = 0
    Public Const GDIP_IMAGETYPE_BITMAP As Integer = 1
    Public Const GDIP_IMAGETYPE_METAFILE As Integer = 2

    ' ImageFlags flags constants
    Public Const GDIP_IMAGEFLAGS_NONE As Integer = &H0 ' no format information.
    Public Const GDIP_IMAGEFLAGS_SCALABLE As Integer = &H1 ' image can be scaled.
    Public Const GDIP_IMAGEFLAGS_HASALPHA As Integer = &H2 ' pixel data contains alpha values.
    Public Const GDIP_IMAGEFLAGS_HASTRANSLUCENT As Integer = &H4 ' pixel data has alpha values other than 0 (transparent) and 255 (opaque).
    Public Const GDIP_IMAGEFLAGS_PARTIALLYSCALABLE As Integer = &H8 ' pixel data is partially scalable with some limitations.
    Public Const GDIP_IMAGEFLAGS_COLORSPACE_RGB As Integer = &H10 ' image is stored using an RGB color space.
    Public Const GDIP_IMAGEFLAGS_COLORSPACE_CMYK As Integer = &H20 ' image is stored using a CMYK color space.
    Public Const GDIP_IMAGEFLAGS_COLORSPACE_GRAY As Integer = &H40 ' image is a grayscale image.
    Public Const GDIP_IMAGEFLAGS_COLORSPACE_YCBCR As Integer = &H80 ' image is stored using a YCBCR color space.
    Public Const GDIP_IMAGEFLAGS_COLORSPACE_YCCK As Integer = &H100 ' image is stored using a YCCK color space.
    Public Const GDIP_IMAGEFLAGS_HASREALDPI As Integer = &H1000 ' dots per inch information is stored in the image.
    Public Const GDIP_IMAGEFLAGS_HASREALPIXELSIZE As Integer = &H2000 ' pixel size is stored in the image.
    Public Const GDIP_IMAGEFLAGS_READONLY As Integer = &H10000 ' pixel data is read-only.
    Public Const GDIP_IMAGEFLAGS_CACHING As Integer = &H20000 ' pixel data can be cached for faster access.

    ' ===============================================================================================================================

    ' ===============================================================================================================================
    ' #VARIABLES# ===================================================================================================================
    Public giBMPFormat As Integer = GDIP_PXF24RGB
    Public giJPGQuality As Integer = 100
    Public giTIFColorDepth As Integer = 24
    Public giTIFCompression As Integer = GDIP_EVTCOMPRESSIONLZW
    ' ===============================================================================================================================

    ' #CONSTANTS# ===================================================================================================================
    Public Const __SCREENCAPTURECONSTANT_SM_CXSCREEN As Integer = 0
    Public Const __SCREENCAPTURECONSTANT_SM_CYSCREEN As Integer = 1
    Public Const __SCREENCAPTURECONSTANT_SRCCOPY As Integer = &HCC0020
    ' ===============================================================================================================================

    '
    Public Declare Function GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As IntPtr
    Public Declare Function GetDC Lib "user32" Alias "GetDC" (ByVal hwnd As IntPtr) As IntPtr
    Public Declare Function CreateCompatibleDC Lib "gdi32" Alias "CreateCompatibleDC" (ByVal hdc As IntPtr) As IntPtr
    Public Declare Function CreateCompatibleBitmap Lib "gdi32" Alias "CreateCompatibleBitmap" (ByVal hdc As IntPtr, ByVal nWidth As Integer, ByVal nHeight As Integer) As IntPtr
    Public Declare Function SelectObject Lib "gdi32" Alias "SelectObject" (ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    Public Declare Function BitBlt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As IntPtr, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Integer) As Integer
    Public Declare Function GetCursorInfo Lib "User32" Alias "GetCursorInfo" (ByRef pci As PCURSORINFO) As Integer
    Public Declare Function CopyIcon Lib "user32" Alias "CopyIcon" (ByVal hIcon As IntPtr) As IntPtr
    Public Declare Function DrawIcon Lib "user32" Alias "DrawIcon" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal hIcon As Integer) As Integer
    Public Declare Function DestroyIcon Lib "user32" Alias "DestroyIcon" (ByVal hIcon As IntPtr) As Integer
    Public Declare Function ReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hwnd As IntPtr, ByVal hdc As IntPtr) As Integer
    Public Declare Function DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As IntPtr) As Integer
    Public Declare Function GetIconInfo Lib "user32" Alias "GetIconInfo" (ByVal hIcon As IntPtr, ByVal piconinfo As ICONINFO) As Integer
    Public Declare Function DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As IntPtr) As Integer
    Structure PCURSORINFO
        Public cbSize As Integer
        Public flags As Integer
        Public hCursor As Integer
        Public ptScreenPos As Point
    End Structure
    Structure ICONINFO
        Public fIcon As Boolean
        Public xHotspot As Integer
        Public yHotspot As Integer
        Public hbmMask As Integer
        Public hbmColor As Integer
    End Structure

    Public Shared Function Capture(ByVal Bounds As Rectangle, Optional ByVal Cursor As Boolean = True) As Image
        Return Capture(Bounds.X, Bounds.Y, Bounds.X + Bounds.Width, Bounds.Y + Bounds.Height, Cursor)
    End Function
    Public Shared Function Capture(ByVal Location As Point, ByVal Size As Size, Optional ByVal Cursor As Boolean = True) As Image
        Return CaptureScreen(Location.X, Location.Y, Location.X + Size.Width, Location.Y + Size.Height, Cursor)
    End Function
    Public Shared Function Capture(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Optional ByVal Cursor As Boolean = True) As Image
        Return CaptureScreen(X, Y, X + Width, Y + Height, Cursor)
    End Function
    Private Shared Function CaptureScreen(ByVal iLeft As Integer, ByVal iTop As Integer, ByVal iRight As Integer, ByVal iBottom As Integer, ByVal fCursor As Boolean) As Image
        If iRight < iLeft Then Return Nothing
        If iBottom < iTop Then Return Nothing

        Dim iW = (iRight - iLeft) + 1
        Dim iH = (iBottom - iTop) + 1
        Dim hWnd = GetDesktopWindow()
        Dim hDDC = GetDC(hWnd)
        Dim hCDC = CreateCompatibleDC(hDDC)
        Dim hBMP = CreateCompatibleBitmap(hDDC, iW, iH)
        SelectObject(hCDC, hBMP)
        BitBlt(hCDC, 0, 0, iW, iH, hDDC, iLeft, iTop, __SCREENCAPTURECONSTANT_SRCCOPY)

        If fCursor Then
            Dim aCursor As New PCURSORINFO
            GetCursorInfo(aCursor)
            If aCursor.flags Then
                Dim hIcon = CopyIcon(aCursor.hCursor)
                Dim aIcon As New ICONINFO
                GetIconInfo(hIcon, aIcon)
                DeleteObject(aIcon.hbmMask) ' delete bitmap mask return by _WinAPI_GetIconInfo()
                DrawIcon(hCDC, aCursor.ptScreenPos.X - aIcon.xHotspot - iLeft, aCursor.ptScreenPos.Y - aIcon.yHotspot - iTop, hIcon)
                DestroyIcon(hIcon)
            End If
        End If

        ReleaseDC(hWnd, hDDC)
        DeleteDC(hCDC)
        If hBMP = 0 Then Return New Bitmap(iW, iH)
        CaptureScreen = Bitmap.FromHbitmap(hBMP)
        DeleteObject(hBMP)
    End Function

End Class
