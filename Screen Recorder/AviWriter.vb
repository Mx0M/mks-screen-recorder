Imports System.Runtime.InteropServices
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection

Public Class AviWriter
    Public Const StreamtypeVIDEO As Integer = 1935960438
    Public Const OF_SHARE_DENY_WRITE As Integer = 32
    Public Const BMP_MAGIC_COOKIE As Integer = 19778

    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Public Structure RECTstruc
        Public left As UInt32
        Public top As UInt32
        Public right As UInt32
        Public bottom As UInt32
    End Structure
    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Public Structure BITMAPINFOHEADERstruc
        Public biSize As UInt32
        Public biWidth As Int32
        Public biHeight As Int32
        Public biPlanes As Int16
        Public biBitCount As Int16
        Public biCompression As UInt32
        Public biSizeImage As UInt32
        Public biXPelsPerMeter As Int32
        Public biYPelsPerMeter As Int32
        Public biClrUsed As UInt32
        Public biClrImportant As UInt32
    End Structure
    <StructLayout(LayoutKind.Sequential, Pack:=1)> _
    Public Structure AVISTREAMINFOstruc
        Public fccType As UInt32
        Public fccHandler As UInt32
        Public dwFlags As UInt32
        Public dwCaps As UInt32
        Public wPriority As UInt16
        Public wLanguage As UInt16
        Public dwScale As UInt32
        Public dwRate As UInt32
        Public dwStart As UInt32
        Public dwLength As UInt32
        Public dwInitialFrames As UInt32
        Public dwSuggestedBufferSize As UInt32
        Public dwQuality As UInt32
        Public dwSampleSize As UInt32
        Public rcFrame As RECTstruc
        Public dwEditCount As UInt32
        Public dwFormatChangeCount As UInt32
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=64)> _
        Public szName As UInt16()
    End Structure
    'Initialize the AVI library
    <DllImport("avifil32.dll")> _
    Public Shared Sub AVIFileInit()
    End Sub
    'Open an AVI file
    <DllImport("avifil32.dll", PreserveSig:=True)> _
    Public Shared Function AVIFileOpen(ByRef ppfile As Integer, ByVal szFile As [String], ByVal uMode As Integer, ByVal pclsidHandler As Integer) As Integer
    End Function
    'Create a new stream in an open AVI file
    <DllImport("avifil32.dll")> _
    Public Shared Function AVIFileCreateStream(ByVal pfile As Integer, ByRef ppavi As IntPtr, ByRef ptr_streaminfo As AVISTREAMINFOstruc) As Integer
    End Function
    'Set the format for a new stream
    <DllImport("avifil32.dll")> _
    Public Shared Function AVIStreamSetFormat(ByVal aviStream As IntPtr, ByVal lPos As Int32, ByRef lpFormat As BITMAPINFOHEADERstruc, ByVal cbFormat As Int32) As Integer
    End Function
    'Write a sample to a stream
    <DllImport("avifil32.dll")> _
    Public Shared Function AVIStreamWrite(ByVal aviStream As IntPtr, ByVal lStart As Int32, ByVal lSamples As Int32, ByVal lpBuffer As IntPtr, ByVal cbBuffer As Int32, ByVal dwFlags As Int32, _
         ByVal dummy1 As Int32, ByVal dummy2 As Int32) As Integer
    End Function
    'Release an open AVI stream
    <DllImport("avifil32.dll")> _
    Public Shared Function AVIStreamRelease(ByVal aviStream As IntPtr) As Integer
    End Function
    'Release an open AVI file
    <DllImport("avifil32.dll")> _
    Public Shared Function AVIFileRelease(ByVal pfile As Integer) As Integer
    End Function
    'Close the AVI library
    <DllImport("avifil32.dll")> _
    Public Shared Sub AVIFileExit()
    End Sub
    Public Declare Function AVIStreamGetFrame Lib "avifil32.dll" (ByVal pgf As IntPtr, ByVal lPos As Integer) As Integer
    <DllImport("avifil32.dll")> _
    Public Shared Sub AVIStreamGetFrameClose(ByVal pget As IntPtr)
    End Sub
    <DllImport("avifil32.dll")> _
    Public Shared Function AVIStreamGetFrameOpen(ByVal pavi As IntPtr, ByVal lpbiWanted As IntPtr) As IntPtr
    End Function
    <DllImport("avifil32.dll")> _
    Public Shared Function AVIStreamLength(ByVal pavi As IntPtr) As Integer
    End Function
    Public Declare Function GetLastError Lib "kernel32" Alias "GetLastError" () As Integer
    Private aviFile As Integer = 0
    Private aviStream As IntPtr = IntPtr.Zero
    Private frameRate As UInt32 = 0
    Private countFrames As Integer = 0
    Private width As Integer = 0
    Private height As Integer = 0
    Private stride As UInt32 = 0
    Private fccType As UInt32 = StreamtypeVIDEO
    Private fccHandler As UInt32 = 1668707181
    Private strideInt As Integer
    Private strideU As UInteger
    Private heightU As UInteger
    Private widthU As UInteger
    Private pget As IntPtr
    Public Property PixelFormat As PixelFormat = Imaging.PixelFormat.Format24bppRgb
    Public Sub OpenAVI(ByVal fileName As String, ByVal frameRate As UInt32, ByVal SampleBitmap As Bitmap, Optional ByVal Mode As Integer = 4097)
        Me.frameRate = frameRate
        Me.countFrames = 0
        AVIFileInit()
        Dim OpeningError As Integer = AVIFileOpen(aviFile, fileName, Mode, 0)
        If OpeningError <> 0 Then
            Throw New Exception("Error in AVIFileOpen: " & OpeningError)
        End If

        Dim bmpData As BitmapData = SampleBitmap.LockBits(New Rectangle(0, 0, SampleBitmap.Width, SampleBitmap.Height), ImageLockMode.[ReadOnly], PixelFormat)
        Dim bmpDatStride As UInteger = bmpData.Stride
        Me.stride = bmpDatStride
        Me.width = SampleBitmap.Width
        Me.height = SampleBitmap.Height
        CreateStream()

        SampleBitmap.UnlockBits(bmpData)
        pget = AVIStreamGetFrameOpen(aviStream, Nothing)
    End Sub
    Public Sub AddFrame(ByVal bmp As Bitmap)
        Dim bmpData As BitmapData = bmp.LockBits(New Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.[ReadOnly], PixelFormat)
        strideInt = stride

        Dim writeResult As Integer = AVIStreamWrite(aviStream, countFrames, 1, bmpData.Scan0, strideInt * height, 0, 0, 0)
        If Not writeResult = 0 Then
            Throw New Exception("Error in AVIStreamWrite: " & writeResult)
        End If

        bmp.UnlockBits(bmpData)
        'System.Math.Max(System.Threading.Interlocked.Increment(countFrames), countFrames - 1)
        countFrames += 1
    End Sub

    Private Structure BITMAPFILEHEADER
        Dim bfType As Integer
        Dim bfSize As Integer
        Dim bfReserved1 As Integer
        Dim bfReserved2 As Integer
        Dim bfOffBits As Integer
    End Structure

    Private Structure BITMAPINFOHEADER
        Dim biSize As Integer
        Dim biWidth As Integer
        Dim biHeight As Integer
        Dim biPlanes As Integer
        Dim biBitCount As Integer
        Dim biCompression As Integer
        Dim biSizeImage As Integer
        Dim biXPelsPerMeter As Integer
        Dim biYPelsPerMeter As Integer
        Dim biClrUsed As Integer
        Dim biClrImportant As Integer
    End Structure

    Public Const BI_RGB = 0&

    Public Function GetFrame(ByVal Index As Integer) As Bitmap
        Dim pDib As IntPtr = AVIStreamGetFrame(pget, Index)
        Return BitmapFromDIB(pDib)
    End Function
    Public Declare Function GdipCreateBitmapFromGdiDib Lib "GdiPlus.dll" Alias "GdipCreateBitmapFromGdiDib" (ByVal pBIH As IntPtr, ByVal pPix As IntPtr, ByRef pBitmap As IntPtr) As Integer
    Private Function BitmapFromDIB(ByVal pDIB As IntPtr) As Bitmap
        Dim pPix As IntPtr = New IntPtr(pDIB.ToInt32() + Marshal.SizeOf(GetType(BITMAPINFOHEADER)))
        Dim mi As MethodInfo = GetType(Bitmap).GetMethod("FromGDIplus", BindingFlags.Static Or BindingFlags.NonPublic)
        If IsNothing(mi) Then Return Nothing
        Dim pBmp As IntPtr = IntPtr.Zero
        Dim status As Integer = GdipCreateBitmapFromGdiDib(pDIB, pPix, pBmp)
        If status = 0 Or Not pBmp = IntPtr.Zero Then
            Return mi.Invoke(Nothing, {pBmp})
        Else
            Return Nothing
        End If
    End Function
    Public Function FramesCount() As Integer
        Return AVIStreamLength(aviStream)
    End Function
    Private Sub CreateStream()
        Dim strhdr As New AVISTREAMINFOstruc()
        strhdr.fccType = fccType
        strhdr.fccHandler = fccHandler
        strhdr.dwScale = 1
        strhdr.dwRate = frameRate
        strideU = stride
        heightU = height
        strhdr.dwSuggestedBufferSize = stride * strideU
        strhdr.dwQuality = 10000

        heightU = height
        widthU = width
        strhdr.rcFrame.bottom = heightU
        strhdr.rcFrame.right = widthU
        strhdr.szName = New UInt16(64) {}
        Dim createResult As Integer = AVIFileCreateStream(aviFile, aviStream, strhdr)
        If createResult <> 0 Then
            Throw New Exception("Error in AVIFileCreateStream: " + createResult.ToString())
        End If
        Dim bi As New BITMAPINFOHEADERstruc()
        Dim bisize As UInteger = Marshal.SizeOf(bi)
        bi.biSize = bisize
        bi.biWidth = width
        bi.biHeight = height
        bi.biPlanes = 1
        bi.biBitCount = 24

        strideU = stride
        heightU = height
        bi.biSizeImage = strideU * heightU
        Dim formatResult As Integer = AVIStreamSetFormat(aviStream, 0, bi, Marshal.SizeOf(bi))
        If formatResult <> 0 Then
            Throw New Exception("Error in AVIStreamSetFormat: " + formatResult.ToString())
        End If
    End Sub
    Public Sub Close()
        If Not aviStream = IntPtr.Zero Then
            AVIStreamRelease(aviStream)
            aviStream = IntPtr.Zero
        End If
        If Not pget = IntPtr.Zero Then
            AVIStreamGetFrameClose(pget)
            pget = IntPtr.Zero
        End If
        If Not aviFile = 0 Then
            AVIFileRelease(aviFile)
            aviFile = 0
        End If
        countFrames = 0
        AVIFileExit()
    End Sub
    Public Sub Flush()
        If Not aviStream = IntPtr.Zero Then
            AVIStreamRelease(aviStream)
        End If
        If Not aviFile = 0 Then
            AVIFileRelease(aviFile)
        End If
    End Sub
End Class