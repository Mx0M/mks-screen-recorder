Imports System
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging

Public Class Form1
    Dim myframerate As UInteger = 24
    Public Declare Function GetLastError Lib "kernel32" Alias "GetLastError" () As Integer
    Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As StringBuilder, ByVal cchBuffer As Integer) As Integer
    Public Declare Function AVIFileOpen Lib "Avifil32.dll" Alias "AVIFileOpenA" (ByRef ppfile As IntPtr, ByVal szFile As String, ByVal mode As UInteger, ByVal pclsidHandler As IntPtr) As Integer
    Public Declare Function AVIFileRelease Lib "Avifil32.dll" Alias "AVIFileRelease" (ByVal pfile As IntPtr) As ULong
    Public Declare Function AVIFileCreateStream Lib "Avifil32.dll" Alias "AVIFileCreateStream" (ByVal pfile As IntPtr, ByRef ppavi As IntPtr, ByRef psi As AviStreamInfo) As Integer
    Public Declare Function AVIStreamWrite Lib "Avifil32.dll" Alias "AVIStreamWrite" (ByVal pavi As IntPtr, ByVal lStart As Long, ByVal lSamples As Long, ByVal lpBuffer As IntPtr, ByVal cbBuffer As Long, ByVal dwFlags As Integer, ByRef plSampWritten As Long, ByRef plBytesWritten As LoaderOptimization) As Integer
    Public Declare Function AVIFileInit Lib "avifil32.dll" Alias "AVIFileInit" () As Integer
    Public Declare Function AVIFileExit Lib "avifil32.dll" Alias "AVIFileExit" () As Integer

    Public Declare Function GetCursorInfo Lib "User32" Alias "GetCursorInfo" (ByRef pci As PCURSORINFO) As Boolean
    Public Declare Function GetCursorInfo Lib "User32" Alias "GetCursorInfo" (ByVal pci As IntPtr) As Boolean
    Public Declare Function CopyIcon Lib "user32" Alias "CopyIcon" (ByVal hIcon As IntPtr) As IntPtr
    Public Declare Function DrawIcon Lib "user32" Alias "DrawIcon" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal hIcon As Integer) As Integer
    Public Declare Function DestroyIcon Lib "user32" Alias "DestroyIcon" (ByVal hIcon As IntPtr) As Integer
    Public Declare Function GetIconInfo Lib "user32" Alias "GetIconInfo" (ByVal hIcon As IntPtr, ByRef piconinfo As ICONINFO) As Integer
    Public Declare Function DeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As IntPtr) As Integer
    Public Declare Function ReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hwnd As IntPtr, ByVal hdc As IntPtr) As Integer
    Public Declare Function DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As IntPtr) As Integer
    Public Declare Function GetAsyncKeyState Lib "user32" Alias "GetAsyncKeyState" (ByVal vKey As Integer) As Boolean

    Structure PCURSORINFO
        Dim cbSize As Integer
        Dim flags As Integer
        Dim hCursor As IntPtr
        Dim ptScreenPos As Point
    End Structure
    Structure ICONINFO
        Dim fIcon As Boolean
        Dim xHotspot As Integer
        Dim yHotspot As Integer
        Dim hbmMask As Integer
        Dim hbmColor As Integer
    End Structure


#Region "OF_Constants"
    Public Const OF_CREATE As Integer = &H1000
    Public Const OF_PARSE As Integer = &H100
    Public Const OF_READ As Integer = &H0
    Public Const OF_READWRITE As Integer = &H2
    Public Const OF_WRITE As Integer = &H1
    Public Const OF_VERIFY As Integer = &H400
    Public Const OF_SHARE_EXCLUSIVE As Integer = &H10
    Public Const OF_SHARE_DENY_WRITE As Integer = &H20
    Public Const OF_SHARE_DENY_READ As Integer = &H30
    Public Const OF_SHARE_DENY_NONE As Integer = &H40
    Public Const OF_SHARE_COMPAT As Integer = &H0
    Public Const OF_REOPEN As Integer = &H8000
    Public Const OF_PROMPT As Integer = &H2000
    Public Const OF_EXIST As Integer = &H4000
    Public Const OF_DELETE As Integer = &H200
    Public Const OF_CANCEL As Integer = &H800
    Public Const VK_LBUTTON As Integer = &H1
    Public Const VK_RBUTTON As Integer = &H2

#End Region
    Public Structure AviStreamInfo
        Dim fccType As Integer
        Dim fccHandler As Integer
        Dim dwFlags As Integer
        Dim dwCaps As Integer
        Dim wPriority As Integer
        Dim wLanguage As Integer
        Dim dwScale As Integer
        Dim dwRate As Integer
        Dim dwStart As Integer
        Dim dwLength As Integer
        Dim dwInitialFrames As Integer
        Dim dwSuggestedBufferSize As Integer
        Dim dwQuality As Integer
        Dim dwSampleSize As Integer
        Dim rcFrame As Rectangle
        Dim dwEditCount As Integer
        Dim dwFormatChangeCount As Integer
        Dim szName As Char()
    End Structure

    Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding(1252)
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        penw.SetLineCap(Drawing2D.LineCap.Round, Drawing2D.LineCap.Round, Drawing2D.DashCap.Round)
        ComboBox1.SelectedIndex = 0

        'MessageBox.Show(myframerate)
    End Sub

    Public Function OpenAvi(ByVal FileName As String) As IntPtr
        If System.IO.File.Exists(FileName) Then
            System.IO.File.Delete(FileName)
        End If
        System.IO.File.WriteAllText(FileName, "", enc)

        AVIFileInit()
        Dim h As Integer = AVIFileOpen(OpenAvi, FileName, OF_CREATE + 1, Nothing)
        If GetLastError Then
            Throw New ArgumentException("Error when creating avi file, Error: " & GetLastError)
        End If
        If Not h = 0 Then
            Throw New ArgumentException("Error when creating avi file, ErrorAvi: " & OpenAvi.ToInt32)
        End If
        If OpenAvi = 0 Then
            Throw New ArgumentException("Error when creating avi file, handle: 0")
        End If
    End Function
    Public Function CreateStreamAvi(ByVal hAvi As IntPtr) As IntPtr
        Dim LPASI As New AviStreamInfo
        Dim bmp As New Bitmap(200, 200)
        Dim bmpDat As BitmapData = bmp.LockBits(New Rectangle(0, 0, 200, 200), ImageLockMode.ReadOnly, bmp.PixelFormat)
        Dim h As IntPtr = 0
        With LPASI
            .fccType = Convert.ToUInt32(1935960438)
            .fccHandler = Convert.ToUInt32(808810089)
            .rcFrame = New Rectangle(0, 0, 200, 200)
            .dwRate = Convert.ToUInt32(25)
            .dwSuggestedBufferSize = bmpDat.Stride * bmp.Height
            .dwQuality = 100
            h = AVIFileCreateStream(hAvi, CreateStreamAvi, LPASI)
        End With
        bmp.Dispose()
        If GetLastError Or Not h = 0 Then
            Throw New ArgumentException("Error when creating avi stream")
        End If
        If GetLastError Or CreateStreamAvi = 0 Then
            Throw New ArgumentException("Error when creating avi stream")
        End If
    End Function
    Public Function AddAviFrame(ByVal hAviStream As IntPtr, ByVal bmp As Bitmap, ByVal Index As Integer) As Boolean
        Dim bmpData As BitmapData = bmp.LockBits(New Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.[ReadOnly], PixelFormat.Format24bppRgb)
        If Index = 0 Then
            Dim bmpDatStride As UInteger = bmpData.Stride
        End If
        Dim writeResult As Integer = AVIStreamWrite(hAviStream, Index, 1, bmpData.Scan0, bmpData.Stride * Height, 0, 0, 0)

        If GetLastError Then
            Throw (New ArgumentException("Error when add frame"))
        End If

        bmp.UnlockBits(bmpData)

        Return True
    End Function


    Public ReadOnly Property ApplicationDir As String
        Get
            Return System.IO.Path.GetDirectoryName(Application.ExecutablePath)
        End Get
    End Property

    Dim aviBuff As String = Path.GetFullPath("record.avi")
    Dim objAvi As AviWriter
    Public sz As New Size(200, 200)
    Dim bmps As Bitmap()
    Dim bmps2 As Bitmap()
    Dim bmpx As Integer = 0
    Dim bmpx2 As Integer = 0
    Dim totalbmp As Integer = 0
    Dim bg As Boolean = False
    Dim fCursor As Boolean = True
    Dim sw As New Stopwatch
    Dim sw2 As New Stopwatch
    Dim curpos As Point
    Dim mdown As Integer
    Dim mup As Integer
    Dim mdowncolor As Color = Color.Lime
    Dim mupcolor As Color = Color.Aqua
    Dim oldbrsh As New SolidBrush(Color.FromArgb(50, Color.Red))

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            '  If CheckBox3.Checked Then
            '  If GetAsyncKeyState(VK_LBUTTON) Then
            'mdown = 150
            ' ElseIf GetAsyncKeyState(VK_RBUTTON) Then
            ' mup = 150
            ' End If
            ' End If

            Dim p As Point = curpos
            Dim np As Point = MousePosition


            'np.X -= 5
            'np.Y -= 5
            Select Case CameraMode
                Case 0
                    If Abs(p.X - np.X) > sz.Width Then
                        p.X += (np.X - p.X) / 8
                    Else
                        p.X += (np.X - p.X) / 16
                    End If

                    If Abs(p.Y - np.Y) > sz.Height Then
                        p.Y += (np.Y - p.Y) / 8
                    Else
                        p.Y += (np.Y - p.Y) / 16
                    End If
                Case 1
                    p = np
            End Select

            curpos = p
            If sw2.Elapsed.TotalMilliseconds >= 200 Then
                sw2.Reset()
                If FrameCap.BrushX = 0 Then
                    FrameCap.Brush = Brushes.Red
                    FrameCap.BrushX = 1
                    FrameCap.Draw()
                Else
                    FrameCap.Brush = Brushes.Lime
                    FrameCap.BrushX = 0
                    FrameCap.Draw()
                End If
                sw2.Start()
            End If
            FrameCap.Location = New Point(curpos.X - sz.Width / 2, curpos.Y - sz.Height / 2)
            Dim bmp As New Bitmap(sz.Width, sz.Height)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.CopyFromScreen(p.X - sz.Width / 2, p.Y - sz.Height / 2, 0, 0, bmp.Size)
                UpdateCur(g)

                If fCursor Then
                    Dim mpos As New Point(CInt(sz.Width / 2) + (np.X - p.X), CInt(sz.Height / 2) + (np.Y - p.Y))
                    If Not mdown = 0 Then
                        g.FillEllipse(New SolidBrush(Color.FromArgb(mdown, mdowncolor)), New Rectangle(New Point(mpos.X - 25, mpos.Y - 25), New Size(50, 50)))
                        mdown -= 10
                        If mdown < 0 Then mdown = 0
                    End If
                    If Not mup = 0 Then
                        g.FillEllipse(New SolidBrush(Color.FromArgb(mup, mupcolor)), New Rectangle(New Point(mpos.X - 25, mpos.Y - 25), New Size(50, 50)))
                        mup -= 10
                        If mup < 0 Then mup = 0
                    End If
                    Dim acur As New PCURSORINFO
                    acur.cbSize = Marshal.SizeOf(acur)
                    GetCursorInfo(acur)
                    Dim ico As Icon = Icon.FromHandle(acur.hCursor)
                    g.DrawIcon(ico, mpos.X, mpos.Y)
                    ico.Dispose()
                    'g.FillRectangle(oldbrsh, New Rectangle(New Point(0, 0), sz))
                End If
            End Using
            'InvertBitmap(bmp)

            If bg Then
                ReDim Preserve bmps2(bmpx2)
                bmps(bmpx2) = bmp
                bmpx2 += 1
            Else
                ReDim Preserve bmps(bmpx)
                bmps(bmpx) = bmp
                bmpx += 1
            End If

            If sz.Width > Panel1.Width Or sz.Height > Panel1.Height Then
                If sz.Width > sz.Height Then
                    Dim genw As Integer = sz.Height / 200 * Panel1.Height
                    Panel1.CreateGraphics.DrawImage(bmp, New Rectangle(0, Panel1.Height / 2 - genw / 2, Panel1.Width, genw))
                Else
                    Dim genw As Integer = sz.Width / 200 * Panel1.Width
                    Panel1.CreateGraphics.DrawImage(bmp, New Rectangle(Panel1.Width / 2 - genw / 2, 0, genw, Panel1.Height))
                End If
            Else
                Panel1.CreateGraphics.DrawImage(bmp, CInt(Panel1.Width / 2 - sz.Width / 2), CInt(Panel1.Height / 2 - sz.Height / 2))
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            bg = True
            Dim tm As Boolean = Timer1.Enabled
            If Not bmpx = 0 Then
                Timer1.Enabled = False
                For Each i In bmps
                    i.RotateFlip(RotateFlipType.RotateNoneFlipY)
                    objAvi.AddFrame(i)
                    i.Dispose()
                Next
                bmps = Nothing
                totalbmp += bmpx
                bmpx = 0
            End If


            bg = False

            If Not bmpx2 = 0 Then
                For Each i In bmps2
                    i.RotateFlip(RotateFlipType.RotateNoneFlipY)
                    objAvi.AddFrame(i)
                    i.Dispose()
                Next
                bmps2 = Nothing
                totalbmp += bmpx2
                bmpx2 = 0
            End If

            If Not tm Then
                Timer2.Enabled = False
            Else
                Timer1.Enabled = True
            End If
            Label3.Text = sw.Elapsed.ToString
            Label4.Text = totalbmp
            Label5.Text = SizeName(Microsoft.VisualBasic.FileLen(aviBuff), 2)
        Catch ex As Exception

        End Try
    End Sub

    Public Function SizeName(ByVal Size As Long, Optional ByVal Decimals As Integer = -1) As String
        If Decimals < 0 Then
            Decimals = 15
        End If
        Select Case Size
            Case 0 To 1023
                Return Math.Round(Size, Decimals) & " Bytes"
            Case 1024 To 1024 * 1024 - 1
                Return Math.Round(Size / 1024, Decimals) & " KB"
            Case 1024 * 1024 To 1024 * 1024 * 1024 - 1
                Return Math.Round(Size / 1024 / 1024, Decimals) & " MB"
            Case 1024 * 1024 * 1024 To CLng(1024 * 1024 * 1024) * 1024 - 1
                Return Math.Round(Size / 1024 / 1024 / 1024, Decimals) & " GB"
            Case CLng(1024 * 1024 * 1024) * 1024 To CLng(1024 * 1024 * 1024) * 1024 * 1024 - 1
                Return Math.Round(Size / 1024 / 1024 / 1024 / 1024, Decimals) & " TB"
            Case CLng(1024 * 1024 * 1024) * 1024 * 1024 To CLng(1024 * 1024 * 1024) * 1024 * 1024 * 1024 - 1
                Return Math.Round(Size / 1024 / 1024 / 1024 / 1024 / 1024, Decimals) & " PB"
            Case Else
                Return 0
        End Select
    End Function
    Private Function InRange(ByVal Location As Point, ByVal Bounds As Rectangle) As Boolean
        If Location.X >= Bounds.X And Location.X <= Bounds.X + Bounds.Width And Location.Y >= Bounds.Y And Location.Y <= Bounds.Y + Bounds.Height Then Return True
        Return False
    End Function
    Private Function InRange(ByVal Rect As Rectangle, ByVal Bounds As Rectangle) As Boolean
        If (Rect.X >= Bounds.X And Rect.X <= Bounds.X + Bounds.Width And Rect.Y >= Bounds.Y And Rect.Y <= Bounds.Y + Bounds.Height) Or (Rect.X + Rect.Width >= Bounds.X And Rect.X <= Bounds.X + Rect.Width + Bounds.Width And Rect.Y + Rect.Height >= Bounds.Y And Rect.Y + Rect.Height <= Bounds.Y + Bounds.Height) Then Return True
        Return False
    End Function


    Public Function InvertNewBitmap(ByVal b As Bitmap) As Bitmap
        InvertNewBitmap = New Bitmap(b)
        Dim Color As Color = Nothing
        For i As Integer = 0 To InvertNewBitmap.Width - 1
            For i2 As Integer = 0 To InvertNewBitmap.Height - 1
                Color = InvertNewBitmap.GetPixel(i, i2)
                InvertNewBitmap.SetPixel(i, i2, Color.FromArgb(Color.A, 255 - Color.R, 255 - Color.G, 255 - Color.B))
            Next
        Next
    End Function
    Public Sub InvertBitmap(ByRef b As Bitmap)
        Dim Color As Color = Nothing
        For i As Integer = 0 To b.Width - 1
            For i2 As Integer = 0 To b.Height - 1
                Color = b.GetPixel(i, i2)
                b.SetPixel(i, i2, Color.FromArgb(Color.A, 255 - Color.R, 255 - Color.G, 255 - Color.B))
            Next
        Next
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            Dim framecount As Integer = objAvi.FramesCount
            Dim bmp As Bitmap = Nothing
            Dim samplebmp As New Bitmap(sz.Width, sz.Height)
            For i As Integer = 0 To framecount
                bmp = objAvi.GetFrame(i)
                'InvertBitmap(bmp)
                Panel1.CreateGraphics.DrawImage(bmp, 0, 0)
            Next
        Catch ex As Exception

        End Try
    End Sub
    Public Function Abs(ByVal n As Integer) As Integer
        If n < 0 Then
            Return -n
        Else
            Return n
        End If
    End Function

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If FrameCap.Visible Then
            FrameCap.Hide()
        Else
            curpos = MousePosition
            FrameCap.CanResize = True
            FrameCap.Size = sz
            FrameCap.Location = New Point(curpos.X - sz.Width / 2, curpos.Y - sz.Height / 2)
            FrameCap.Show()
        End If
    End Sub

    Dim CameraMode As Integer = 0


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try


            If FrameCap.Visible Then
                FrameCap.Hide()
            Else
                curpos = MousePosition
                FrameCap.CanResize = True
                FrameCap.Size = sz
                FrameCap.Location = New Point(curpos.X - sz.Width / 2, curpos.Y - sz.Height / 2)
                FrameCap.Show()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox1.Text) And IsNumeric(TextBox2.Text) Then
            sz.Width = TextBox1.Text
            sz.Height = TextBox2.Text
            FrameCap.Size = sz
        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If IsNumeric(TextBox1.Text) And IsNumeric(TextBox2.Text) Then
            sz.Width = TextBox1.Text
            sz.Height = TextBox2.Text
            FrameCap.Size = sz
        End If
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonsave.Click
        Try


            If File.Exists(aviBuff) Then
                Dim fl As New SaveFileDialog()
                If fl.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                    System.IO.File.Copy(aviBuff, fl.FileName, True)
                End If
            Else
                MsgBox("Cannot find '" & aviBuff & "'", MsgBoxStyle.Exclamation)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        Try


            If Not IsNothing(objAvi) Then objAvi.Close()
            If File.Exists(aviBuff) Then
                Try
                    File.Delete(aviBuff)
                Catch ex As Exception
                    MsgBox("Cannot delete '" & aviBuff & "'. Please delete it manually", MsgBoxStyle.Exclamation)
                End Try
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        fCursor = CheckBox1.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked Then
            Me.Opacity = 1
        Else
            Me.Opacity = 0.99
        End If
    End Sub

    Dim curves As Curve()
    Dim bgd As Boolean
    Structure Curve
        Dim Cur As Point()
    End Structure
    Public Sub AddPoint(ByRef ps As Point(), ByVal p As Point)
        Try


            If IsNothing(ps) Then
                ps = {p, p, p}
            Else
                Dim u As Integer = ps.Length
                ReDim Preserve ps(u)
                ps(u) = p
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub AddCur(ByRef ps As Curve(), ByVal p As Curve)
        Try


            If IsNothing(ps) Then
                ps = {p}
            Else
                Dim u As Integer = ps.Length
                ReDim Preserve ps(u)
                ps(u) = p
            End If
        Catch ex As Exception

        End Try
    End Sub
    Dim penw As New Pen(Color.FromArgb(150, Color.Red), 8)

    Public Sub UpdateCur(Optional ByVal g As Graphics = Nothing)
        Try


            If IsNothing(g) Then g = Panel1.CreateGraphics
            If Not IsNothing(curves) Then
                For Each i In curves
                    If Not IsNothing(i.Cur) Then
                        g.DrawCurve(penw, i.Cur)
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
    End Sub
    Dim img As New Bitmap(200, 200)
    Dim gimg As Graphics = Graphics.FromImage(img)

    Private Sub Panel1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseDown
        Try


            If e.Button = Windows.Forms.MouseButtons.Left Then
                Dim cur As New Curve
                AddPoint(cur.Cur, e.Location)
                AddCur(curves, cur)
                If Not Timer1.Enabled And Not Timer3.Enabled Then
                    gimg.Clear(Color.Black)
                    UpdateCur(gimg)
                    Me.Panel1.BackgroundImage = img
                    Me.Panel1.CreateGraphics.DrawImage(img, 0, 0)
                Else
                    UpdateCur()
                End If
                bgd = True
            ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
                curves = Nothing
                If Not Timer1.Enabled And Not Timer3.Enabled Then
                    gimg.Clear(Color.Black)
                    Me.Panel1.BackgroundImage = img
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Panel1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseMove
        Try


            If GetAsyncKeyState(VK_LBUTTON) And Not bgd Then
                If Not IsNothing(curves) Then
                    Dim ix As Integer = curves.Length - 1
                    Dim cur As Curve = curves(ix)
                    AddPoint(cur.Cur, e.Location)
                    curves(ix) = cur
                    If Not Timer1.Enabled And Not Timer3.Enabled Then
                        gimg.Clear(Color.Black)
                        UpdateCur(gimg)
                        Me.Panel1.CreateGraphics.DrawImage(img, 0, 0)
                    End If
                End If
            Else
                bgd = False
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Panel1_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Panel1.MouseUp
        Try


            If Not Timer1.Enabled And Not Timer3.Enabled Then
                gimg.Clear(Color.Black)
                UpdateCur(gimg)
                Me.Panel1.BackgroundImage = img
                Me.Panel1.CreateGraphics.DrawImage(img, 0, 0)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Try


            Dim t As New Date(CLng(CInt((TrackBar1.Value) / 52) * 10000000))
            ToolTip1.Show(t.TimeOfDay.ToString, TrackBar1)
            Dim frameX As Integer = TrackBar1.Value
            If frameX >= frameY Then
                Timer3.Enabled = False
                Button2.Text = "Play"
                frameX = 0
                TrackBar1.Value = 0
            Else
                frameX += 1
                Dim bmp As Bitmap = objAvi.GetFrame(frameX)
                If IsNothing(bmp) Then
                    Timer3.Enabled = False
                    Button2.Text = "Play"
                    frameX = 0
                    TrackBar1.Value = 0
                    MsgBox("Error!, Cannot read frame")
                Else
                    Me.Panel1.CreateGraphics.DrawImage(bmp, 0, 0)
                    TrackBar1.Value = frameX
                    t = New Date(CLng(CInt((frameY) / 52) * 10000000))
                    Dim t2 As New Date(CLng(CInt((frameX) / 52) * 10000000))
                    Label9.Text = t.TimeOfDay.ToString & "/" & t2.TimeOfDay.ToString
                End If
            End If
            Me.frameX = frameX
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try


            If Not IsNothing(objAvi) Then
                objAvi.Close()
                Button5.Enabled = False
                TrackBar1.Enabled = False
                Label9.Enabled = False
                Button6.Enabled = False

                Button2.Enabled = False
                Button1.Enabled = True
                buttonsave.Enabled = True
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button5_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If Button5.Text = "Play" Then
            Timer3.Enabled = True
            Button5.Text = "Pause"
        Else
            Timer3.Enabled = False
            Button5.Text = "Play"
        End If
    End Sub

    Dim frameX As Integer = 0
    Dim frameY As Integer = 0
    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Try


            If frameX >= frameY Then
                frameX = 0
                TrackBar1.Value = 0
            Else
                frameX += 1
                Dim bmp As Bitmap = objAvi.GetFrame(frameX)
                If IsNothing(bmp) Then
                    Timer3.Enabled = False
                    Button5.Text = "Play"
                    frameX = 0
                    TrackBar1.Value = 0
                    MsgBox("Error!, Cannot read frame")
                Else
                    Dim g As Graphics = Graphics.FromImage(bmp)
                    UpdateCur(g)
                    g.Dispose()
                    Me.Panel1.CreateGraphics.DrawImage(bmp, 0, 0)
                    TrackBar1.Value = frameX
                    myframerate = ComboBox1.SelectedItem
                    Dim t As New Date(CLng(CInt((frameY) / myframerate) * 10000000)) 'frame rateset 25
                    Dim t2 As New Date(CLng(CInt((frameX) / myframerate) * 10000000)) 'frame rate set 25
                    Label9.Text = t.TimeOfDay.ToString & "/" & t2.TimeOfDay.ToString
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try


            If Button1.Text = "Start" Then
                If File.Exists(aviBuff) Then
                    File.Delete(aviBuff)
                End If
                If CameraMode = 0 And CameraMode = 1 Then
                    curpos = MousePosition
                ElseIf CameraMode = 2 Then
                    curpos = New Point(FrameCap.Location.X + sz.Width / 2, FrameCap.Location.Y + sz.Height / 2)
                End If
                If Not CheckBox2.Checked Then Me.Opacity = 0.99
                If Not IsNothing(objAvi) Then objAvi.Close()
                sz.Width = TextBox1.Text
                sz.Height = TextBox2.Text
                FrameCap.Size = sz
                FrameCap.Location = New Point(curpos.X - sz.Width / 2, curpos.Y - sz.Height / 2)
                FrameCap.CanResize = False
                FrameCap.Show()
                FrameCap.Refresh()
                objAvi = New AviWriter
                ' If myframerate = "" Or myframerate = 0 Then
                myframerate = ComboBox1.SelectedItem
                'End If
                objAvi.OpenAVI(aviBuff, myframerate, New Bitmap(sz.Width, sz.Height)) 'framerate 53
                Button1.Text = "Stop"
                buttonsave.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button2.Enabled = False
                TrackBar1.Enabled = False
                Label9.Enabled = False
                Button5.Enabled = False
                Button6.Enabled = False

                Timer1.Enabled = True
                Timer2.Enabled = True
                sw.Reset()
                sw.Start()
                sw2.Reset()
                sw2.Start()

            Else
                Timer1.Enabled = False
                sw.Stop()
                sw2.Stop()
                Me.Opacity = 1
                Button1.Text = "Start"
                Button1.Enabled = False
                Button3.Enabled = True
                Button4.Enabled = True
                Button2.Enabled = True
                TrackBar1.Enabled = True
                Label9.Enabled = True
                Button5.Enabled = True
                Button6.Enabled = True
                FrameCap.CanResize = True
                FrameCap.Refresh()

                Dim frmsc As Integer = objAvi.FramesCount()
                Dim t As New Date(CLng(CInt((frmsc) / 52) * 10000000))
                Dim t2 As New Date(0)
                TrackBar1.Maximum = frmsc
                TrackBar1.Value = 0
                frameX = 0
                frameY = frmsc - 1
                Label9.Text = t.TimeOfDay.ToString & "/" & t2.TimeOfDay.ToString
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Timer3.Enabled = False
        Button5.Text = "Play"
        frameX = 0
    End Sub

    Private Sub TextBox1_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Timer1_Tick_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try



            If CheckBox2.Checked Then
                If GetAsyncKeyState(VK_LBUTTON) Then
                    mdown = 150
                ElseIf GetAsyncKeyState(VK_RBUTTON) Then
                    mup = 150
                End If
            End If

            Dim p As Point = curpos
            Dim np As Point = MousePosition


            'np.X -= 5
            'np.Y -= 5
            Select Case CameraMode
                Case 0
                    If Abs(p.X - np.X) > sz.Width Then
                        p.X += (np.X - p.X) / 8
                    Else
                        p.X += (np.X - p.X) / 16
                    End If

                    If Abs(p.Y - np.Y) > sz.Height Then
                        p.Y += (np.Y - p.Y) / 8
                    Else
                        p.Y += (np.Y - p.Y) / 16
                    End If
                Case 1
                    p = np
            End Select

            curpos = p
            If sw2.Elapsed.TotalMilliseconds >= 200 Then
                sw2.Reset()
                If FrameCap.BrushX = 0 Then
                    FrameCap.Brush = Brushes.Red
                    FrameCap.BrushX = 1
                    FrameCap.Draw()
                Else
                    FrameCap.Brush = Brushes.Lime
                    FrameCap.BrushX = 0
                    FrameCap.Draw()
                End If
                sw2.Start()
            End If
            FrameCap.Location = New Point(curpos.X - sz.Width / 2, curpos.Y - sz.Height / 2)
            Dim bmp As New Bitmap(sz.Width, sz.Height)
            Using g As Graphics = Graphics.FromImage(bmp)
                g.CopyFromScreen(p.X - sz.Width / 2, p.Y - sz.Height / 2, 0, 0, bmp.Size)
                UpdateCur(g)

                If fCursor Then
                    Dim mpos As New Point(CInt(sz.Width / 2) + (np.X - p.X), CInt(sz.Height / 2) + (np.Y - p.Y))
                    If Not mdown = 0 Then
                        g.FillEllipse(New SolidBrush(Color.FromArgb(mdown, mdowncolor)), New Rectangle(New Point(mpos.X - 25, mpos.Y - 25), New Size(50, 50)))
                        mdown -= 10
                        If mdown < 0 Then mdown = 0
                    End If
                    If Not mup = 0 Then
                        g.FillEllipse(New SolidBrush(Color.FromArgb(mup, mupcolor)), New Rectangle(New Point(mpos.X - 25, mpos.Y - 25), New Size(50, 50)))
                        mup -= 10
                        If mup < 0 Then mup = 0
                    End If
                    Dim acur As New PCURSORINFO
                    acur.cbSize = Marshal.SizeOf(acur)
                    GetCursorInfo(acur)
                    Dim ico As Icon = Icon.FromHandle(acur.hCursor)
                    g.DrawIcon(ico, mpos.X, mpos.Y)
                    ico.Dispose()
                    'g.FillRectangle(oldbrsh, New Rectangle(New Point(0, 0), sz))
                End If
            End Using
            'InvertBitmap(bmp)

            If bg Then
                ReDim Preserve bmps2(bmpx2)
                bmps(bmpx2) = bmp
                bmpx2 += 1
            Else
                ReDim Preserve bmps(bmpx)
                bmps(bmpx) = bmp
                bmpx += 1
            End If

            If sz.Width > Panel1.Width Or sz.Height > Panel1.Height Then
                If sz.Width > sz.Height Then
                    Dim genw As Integer = sz.Height / 200 * Panel1.Height
                    Panel1.CreateGraphics.DrawImage(bmp, New Rectangle(0, Panel1.Height / 2 - genw / 2, Panel1.Width, genw))
                Else
                    Dim genw As Integer = sz.Width / 200 * Panel1.Width
                    Panel1.CreateGraphics.DrawImage(bmp, New Rectangle(Panel1.Width / 2 - genw / 2, 0, genw, Panel1.Height))
                End If
            Else
                Panel1.CreateGraphics.DrawImage(bmp, CInt(Panel1.Width / 2 - sz.Width / 2), CInt(Panel1.Height / 2 - sz.Height / 2))
            End If
        Catch castex As System.InvalidCastException
        Catch sysx As System.ArgumentException
        Catch ex As Exception
        End Try

    End Sub

    Private Sub Timer2_Tick_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try


            bg = True
            Dim tm As Boolean = Timer1.Enabled
            If Not bmpx = 0 Then
                Timer1.Enabled = False
                For Each i In bmps
                    i.RotateFlip(RotateFlipType.RotateNoneFlipY)
                    objAvi.AddFrame(i)
                    i.Dispose()
                Next
                bmps = Nothing
                totalbmp += bmpx
                bmpx = 0
            End If


            bg = False

            If Not bmpx2 = 0 Then
                For Each i In bmps2
                    i.RotateFlip(RotateFlipType.RotateNoneFlipY)
                    objAvi.AddFrame(i)
                    i.Dispose()
                Next
                bmps2 = Nothing
                totalbmp += bmpx2
                bmpx2 = 0
            End If

            If Not tm Then
                Timer2.Enabled = False
            Else
                Timer1.Enabled = True
            End If
            Label3.Text = sw.Elapsed.ToString
            Label4.Text = totalbmp
            Label5.Text = SizeName(Microsoft.VisualBasic.FileLen(aviBuff), 2)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton3_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        CameraMode = 1
    End Sub

    Private Sub RadioButton2_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        CameraMode = 2
    End Sub

    Private Sub RadioButton1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        CameraMode = 0
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Form1_MinimumSizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MinimumSizeChanged


    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Hide()
        End If
        If CheckBox3.Checked = True Then
            Me.Hide()
            NotifyIcon1.Visible = False
            FrameCap.Close()
            Timer4.Enabled = True
        End If
    End Sub

    Private Sub NotifyIcon1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseClick

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        If Me.WindowState = FormWindowState.Minimized Then
            Me.Show()
            Me.WindowState = FormWindowState.Normal
            FrameCap.Show()
            Me.Show()
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged

    End Sub
    Sub hotkey()
        Try


            Dim Hotkeyss As Boolean
            Hotkeyss = My.Computer.Keyboard.CtrlKeyDown AndAlso My.Computer.Keyboard.AltKeyDown AndAlso CBool(GetAsyncKeyState(119))
            If Hotkeyss = True Then
                Me.Show()
                Me.WindowState = FormWindowState.Normal
                Me.Show()
                Timer4.Enabled = False
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        hotkey()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class
