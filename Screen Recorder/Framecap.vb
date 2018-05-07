Public Class FrameCap
    Public Property WidthFr As Integer = 20
    Public Property Brush As Brush = Brushes.Lime
    Public Property BrushX As Integer = 0
    Public Property CanResize As Boolean
    Public Sub Draw()
        Dim g As Graphics = Me.CreateGraphics
        g.FillRectangle(Brush, CInt(Me.Width / 2 - WidthFr / 2), 0, WidthFr, 2)
        g.FillRectangle(Brush, 0, CInt(Me.Height / 2 - WidthFr / 2), 2, WidthFr)
        g.FillRectangle(Brush, CInt(Me.Width / 2 - WidthFr / 2), Me.Height - 2, WidthFr, 2)
        g.FillRectangle(Brush, Me.Width - 2, CInt(Me.Height / 2 - WidthFr / 2), 2, WidthFr)
        g.FillRectangle(Brush, 0, 0, WidthFr, 2)
        g.FillRectangle(Brush, 0, 0, 2, WidthFr)
        g.FillRectangle(Brush, Me.Width - WidthFr, 0, WidthFr, 2)
        g.FillRectangle(Brush, Me.Width - 2, 0, 2, WidthFr)
        g.FillRectangle(Brush, 0, Me.Height - 2, WidthFr, 2)
        g.FillRectangle(Brush, 0, Me.Height - WidthFr, 2, WidthFr)
        g.FillRectangle(Brush, Me.Width - WidthFr, Me.Height - 2, WidthFr, 2)
        g.FillRectangle(Brush, Me.Width - 2, Me.Height - WidthFr, 2, WidthFr)
        If CanResize Then g.DrawIcon(Icon.FromHandle(Cursors.SizeAll.CopyHandle), New Rectangle(Me.Width / 2 - 20, Me.Height / 2 - 20, 40, 40))
    End Sub
    Private Sub FrameCap_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        Draw()
    End Sub
    Private Sub FrameCap_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        If CanResize Then
            If InRange(e.Location, New Rectangle(CInt(Me.Width / 2 - WidthFr / 2), 0, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNS
            ElseIf InRange(e.Location, New Rectangle(0, CInt(Me.Height / 2 - WidthFr / 2), 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeWE
            ElseIf InRange(e.Location, New Rectangle(CInt(Me.Width / 2 - WidthFr / 2), Me.Height - 2, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNS
            ElseIf InRange(e.Location, New Rectangle(Me.Width - 2, CInt(Me.Height / 2 - WidthFr / 2), 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeWE
            ElseIf InRange(e.Location, New Rectangle(0, 0, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNWSE
            ElseIf InRange(e.Location, New Rectangle(0, 0, 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeNWSE
            ElseIf InRange(e.Location, New Rectangle(Me.Width - WidthFr, 0, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNESW
            ElseIf InRange(e.Location, New Rectangle(Me.Width - 2, 0, 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeNESW
            ElseIf InRange(e.Location, New Rectangle(0, Me.Height - 2, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNESW
            ElseIf InRange(e.Location, New Rectangle(0, Me.Height - WidthFr, 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeNESW
            ElseIf InRange(e.Location, New Rectangle(Me.Width - WidthFr, Me.Height - 2, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNWSE
            ElseIf InRange(e.Location, New Rectangle(Me.Width - 2, Me.Height - WidthFr, 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeNWSE
            ElseIf InRange(e.Location, New Rectangle(Me.Width / 2 - 20, Me.Height / 2 - 20, 40, 40)) Then
                Me.Cursor = Cursors.SizeAll
            Else
                Me.Cursor = Cursors.Default
            End If
        Else
            If Not Me.Cursor = Cursors.Default Then
                Me.Cursor = Cursors.Default
            End If
        End If
    End Sub
    Private Function InRange(ByVal Location As Point, ByVal Bounds As Rectangle) As Boolean
        If Location.X >= Bounds.X And Location.X <= Bounds.X + Bounds.Width And Location.Y >= Bounds.Y And Location.Y <= Bounds.Y + Bounds.Height Then Return True
        Return False
    End Function
    Private Function InRange(ByVal Rect As Rectangle, ByVal Bounds As Rectangle) As Boolean
        If (Rect.X >= Bounds.X And Rect.X <= Bounds.X + Bounds.Width And Rect.Y >= Bounds.Y And Rect.Y <= Bounds.Y + Bounds.Height) Or (Rect.X + Rect.Width >= Bounds.X And Rect.X <= Bounds.X + Rect.Width + Bounds.Width And Rect.Y + Rect.Height >= Bounds.Y And Rect.Y + Rect.Height <= Bounds.Y + Bounds.Height) Then Return True
        Return False
    End Function

    Public Declare Function ReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Boolean
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As Integer) As Int32
    Public Declare Function GetLastError Lib "kernel32" Alias "GetLastError" () As Integer

    Public Const WM_NCHITTEST As Integer = &H84
    Public Const WM_NCLBUTTONDOWN As Integer = &HA1
    Public Const HTCAPTION = 2
    Public Const HTGROWBOX = 4
    Private Const HTBORDER As Integer = 18
    Private Const HTBOTTOM As Integer = 15
    Private Const HTBOTTOMLEFT As Integer = 16
    Private Const HTBOTTOMRIGHT As Integer = 17
    Private Const HTLEFT As Integer = 10
    Private Const HTRIGHT As Integer = 11
    Private Const HTTOP As Integer = 12
    Private Const HTTOPLEFT As Integer = 13
    Private Const HTTOPRIGHT As Integer = 14

    Private Sub FrameCap_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        If CanResize Then
            Dim smConstants() As Integer = {HTLEFT, HTTOPLEFT, HTTOP, HTTOPRIGHT, HTRIGHT, HTBOTTOMRIGHT, HTBOTTOM, HTBOTTOMLEFT}
            Dim cursorIndex As Integer

            If InRange(e.Location, New Rectangle(CInt(Me.Width / 2 - WidthFr / 2), 0, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNS
                cursorIndex = 2
            ElseIf InRange(e.Location, New Rectangle(0, CInt(Me.Height / 2 - WidthFr / 2), 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeWE
                cursorIndex = 0
            ElseIf InRange(e.Location, New Rectangle(CInt(Me.Width / 2 - WidthFr / 2), Me.Height - 2, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNS
                cursorIndex = 6
            ElseIf InRange(e.Location, New Rectangle(Me.Width - 2, CInt(Me.Height / 2 - WidthFr / 2), 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeWE
                cursorIndex = 4
            ElseIf InRange(e.Location, New Rectangle(0, 0, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNWSE
                cursorIndex = 1
            ElseIf InRange(e.Location, New Rectangle(0, 0, 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeNWSE
                cursorIndex = 1
            ElseIf InRange(e.Location, New Rectangle(Me.Width - WidthFr, 0, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNESW
                cursorIndex = 3
            ElseIf InRange(e.Location, New Rectangle(Me.Width - 2, 0, 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeNESW
                cursorIndex = 3
            ElseIf InRange(e.Location, New Rectangle(0, Me.Height - 2, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNESW
                cursorIndex = 7
            ElseIf InRange(e.Location, New Rectangle(0, Me.Height - WidthFr, 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeNESW
                cursorIndex = 7
            ElseIf InRange(e.Location, New Rectangle(Me.Width - WidthFr, Me.Height - 2, WidthFr, 2)) Then
                Me.Cursor = Cursors.SizeNWSE
                cursorIndex = 5
            ElseIf InRange(e.Location, New Rectangle(Me.Width - 2, Me.Height - WidthFr, 2, WidthFr)) Then
                Me.Cursor = Cursors.SizeNWSE
                cursorIndex = 5
            ElseIf InRange(e.Location, New Rectangle(Me.Width / 2 - 20, Me.Height / 2 - 20, 40, 40)) Then
                Me.Cursor = Cursors.SizeAll
                Form1.RadioButton2.Checked = True
            Else
                Me.Cursor = Cursors.Default
            End If
            If Cursor.Current = Cursors.SizeAll Then
                ReleaseCapture()
                SendMessage(Me.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0)
            Else
                ReleaseCapture()
                SendMessage(Me.Handle, WM_NCLBUTTONDOWN, smConstants(cursorIndex), 0)
            End If
        End If
    End Sub

    Private Sub FrameCap_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If CanResize Then
            Form1.sz = Me.Size
            Form1.TextBox1.Text = Me.Width
            Form1.TextBox2.Text = Me.Height
        End If
    End Sub

    Private Sub FrameCap_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class