Imports System
Imports System.Reflection
Imports System.Threading

Public Class BufferedForm : Inherits Form
    Public Shared PaintCounter As Integer = 0, BGColor As New SolidBrush(Color.FromKnownColor(KnownColor.Control)), BGRects As List(Of Rectangle)
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            cp.Style = cp.Style Or &H2000000 And Not 33554432
            Return cp
        End Get
    End Property 'CreateParams

    Public Sub New()
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
        UpdateStyles()
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        PaintCounter += 1
    End Sub

    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        MyBase.OnResize(e)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
        'If Not exp Is Nothing Then e.Graphics.FillRectangle(BGColor, 0, 0, Width, exp.AddressBar.Top) Else e.Graphics.FillRectangle(BGColor, Me.ClientRectangle)
        e.Graphics.FillRectangle(BGColor, Me.ClientRectangle)
    End Sub

    Private Sub BufferedForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        '
        'BufferedForm
        '
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Name = "BufferedForm"
        Me.ResumeLayout(False)

    End Sub
End Class

Public Class EmptyForm : Inherits Form
    Public Event Resized(ByVal sender As Object, ByVal e As EventArgs)

    Public Structure POINTAPI ' This holds the logical cursor information
        Public x As Integer
        Public y As Integer
    End Structure
    Public Declare Function GetCursorPos Lib "user32" (ByRef point As POINTAPI) As Boolean

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            cp.Style = cp.Style Or &H2000000 And Not 33554432
            Return cp
        End Get
    End Property 'CreateParams

    Public ControlBar As FormBar, CF As String
    Public BorderThickness As Integer = 5, BarColor As New SolidBrush(Color.LightBlue), BorderColor As Pen
    Public Locked As Boolean = False
    Private DragX As Integer = -1, DragY As Integer = -1, MouseDownX As Integer = -1, MouseDownY As Integer = -1
    Private ResizeX As Integer = -1, ResizeY As Integer = -1, ResizeType As Byte = 0
    Private OL As Integer, OT As Integer, OW As Integer, OH As Integer
    Private CornerRects As New List(Of Rectangle)
    Private _backBuffer As Bitmap
    Public PMouse As POINTAPI
    Public WWidth, WHeight As Integer
    Private bMinimizing As Boolean = False, bMaximizing As Boolean = False

    Public Sub New(ByRef ContentFolder As String, ByRef x As Integer, ByRef y As Integer, ByRef w As Integer, ByRef h As Integer)
        Try
            FormBorderStyle = FormBorderStyle.None
            Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
            UpdateStyles()

            OL = x
            OT = y
            OW = w
            OH = h
            WWidth = w - (BorderThickness * 2)
            WHeight = h - (BorderThickness * 2)
            BackColor = Color.White
            BorderColor = New Pen(BarColor.Color)
            MinimumSize = New Size(200, 200)
            TransparencyKey = Color.FromArgb(255, 255, 0, 128)
            CF = ContentFolder
            ControlBar = New FormBar(Me, "Explorer woo", CF)

            AddHandler Me.Load, AddressOf EF_Load
            AddHandler Me.MouseUp, AddressOf EF_MouseUp
            AddHandler Me.MouseDown, AddressOf EF_MouseDown
            AddHandler Me.MouseMove, AddressOf EF_MouseMove
            AddHandler Me.LostFocus, AddressOf EF_LostFocus
            AddHandler Me.GotFocus, AddressOf EF_GotFocus
            AddHandler Me.MouseLeave, AddressOf EF_MouseLeave
        Catch ex As Exception
            MsgBox("Empty Form New: " & ex.Message)
        End Try
    End Sub

    Public Sub EF_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
        MouseDownX = -1
        MouseDownY = -1
        DragX = -1
        DragY = -1
        ResizeX = -1
        ResizeY = -1
    End Sub

    Public Sub EF_GotFocus(ByVal sender As Object, ByVal e As EventArgs)
        MouseDownX = -1
        MouseDownY = -1
        DragX = -1
        DragY = -1
        ResizeX = -1
        ResizeY = -1
    End Sub

    Public Sub EF_Load(ByVal sender As Object, ByVal e As EventArgs)
        SetFormBounds(OL, OT, OW, OH)
    End Sub

    Private Sub SetFormBounds(ByRef l As Integer, ByRef t As Integer, ByRef w As Integer, ByRef h As Integer)
        If w < Me.MinimumSize.Width Then
            w = Me.MinimumSize.Width
            l = Left
        End If
        If h < Me.MinimumSize.Height Then
            h = Me.MinimumSize.Height
            t = Top
        End If
        If l = Left AndAlso t = Top AndAlso w = Width AndAlso h = Height Then Exit Sub
        ControlBar.SetBar(w)
        SetBounds(l, t, w, h)
        CornerRects = New List(Of Rectangle) From {{New Rectangle(0, 0, BorderThickness, 15)}, {New Rectangle(0, 0, 15, BorderThickness)}, {New Rectangle(Width - 15, 0, 15, BorderThickness)}, {New Rectangle(Width - BorderThickness, 0, BorderThickness, 15)}, {New Rectangle(0, Height - 15, BorderThickness, 15)}, {New Rectangle(BorderThickness, Height - BorderThickness, 10, BorderThickness)}, {New Rectangle(Width - BorderThickness, Height - 15, BorderThickness, 15)}, {New Rectangle(Width - 15, Height - BorderThickness, 10, BorderThickness)}}
    End Sub

    Protected Overrides Sub OnSizeChanged(ByVal e As EventArgs)
        ControlBar.SetBar(Width) 'Sets the width of controlbar to new width, and updates position of the 3 top-right form buttons
        If Not (_backBuffer Is Nothing) Then
            _backBuffer.Dispose()
            _backBuffer = Nothing
        End If
        If bMinimizing = True Then
            bMinimizing = False
            MyBase.OnSizeChanged(e)
        ElseIf Width = OW AndAlso Height = OH Then
            MyBase.OnSizeChanged(e)
        Else
            RaiseEvent Resized(Me, e) 'Resizes controls in custom handler, in this example, it is unused - with controls, they don't flicker when resized though
            MyBase.OnSizeChanged(e)
        End If
    End Sub

    Private Sub EF_MouseLeave(ByVal sender As Object, ByVal e As EventArgs)
        MouseDownX = -1
        MouseDownY = -1
        DragX = -1
        DragY = -1
        ResizeX = -1
        ResizeY = -1
    End Sub

    Private Sub EF_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        GetCursorPos(PMouse)
        If Locked = True Then Exit Sub
        If DragX <> -1 Then
            SetBounds(OL + (PMouse.x - DragX), OT + (PMouse.y - DragY), Width, Height)
        ElseIf ResizeX <> -1 Then
            If ResizeType = 1 Then 'NW
                SetFormBounds(OL + (PMouse.x - ResizeX), OT + (PMouse.y - ResizeY), OW - (PMouse.x - ResizeX), OH + (ResizeY - PMouse.y))
            ElseIf ResizeType = 2 Then 'NE
                SetFormBounds(OL, OT + (PMouse.y - ResizeY), OW + (PMouse.x - ResizeX), OH + (ResizeY - PMouse.y))
            ElseIf ResizeType = 3 Then 'SE
                SetFormBounds(OL + (PMouse.x - ResizeX), OT, OW - (PMouse.x - ResizeX), OH + (PMouse.y - ResizeY))
            ElseIf ResizeType = 4 Then 'SW
                SetFormBounds(OL, OT, OW + (PMouse.x - ResizeX), OH + (PMouse.y - ResizeY))
            ElseIf ResizeType = 5 Then 'North
                SetFormBounds(OL, OT + (PMouse.y - ResizeY), OW, OH + (ResizeY - PMouse.y))
            ElseIf ResizeType = 6 Then 'South
                SetFormBounds(OL, OT, OW, OH + (PMouse.y - ResizeY))
            ElseIf ResizeType = 7 Then 'West
                SetFormBounds(OL + (PMouse.x - ResizeX), OT, OW + (ResizeX - PMouse.x), OH)
            ElseIf ResizeType = 8 Then 'East
                SetFormBounds(OL, OT, OW + (PMouse.x - ResizeX), OH)
            End If
        Else
            For i = 0 To 1
                If CornerRects(i).Contains(e.X, e.Y) Then
                    Cursor.Current = Cursors.SizeNWSE
                    Exit Sub
                End If
            Next
            For i = 2 To 3
                If CornerRects(i).Contains(e.X, e.Y) Then
                    Cursor.Current = Cursors.SizeNESW
                    Exit Sub
                End If
            Next
            For i = 4 To 5
                If CornerRects(i).Contains(e.X, e.Y) Then
                    Cursor.Current = Cursors.SizeNESW
                    Exit Sub
                End If
            Next
            For i = 6 To 7
                If CornerRects(i).Contains(e.X, e.Y) Then
                    Cursor.Current = Cursors.SizeNWSE
                    Exit Sub
                End If
            Next
            If e.Y <= BorderThickness Or e.Y >= (Height - BorderThickness) Then
                Cursor.Current = Cursors.SizeNS
            ElseIf e.X <= BorderThickness Or e.X >= (Width - BorderThickness) Then
                Cursor.Current = Cursors.SizeWE
            End If
        End If
    End Sub

    Private Sub EF_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)

        MouseDownX = e.X
        MouseDownY = e.Y
        If Locked = False Then
            Dim lol As POINTAPI
            GetCursorPos(lol)
            OL = Left
            OT = Top
            OW = Width
            OH = Height
            For i = 0 To 1
                If CornerRects(i).Contains(e.X, e.Y) Then
                    Cursor.Current = Cursors.SizeNWSE
                    ResizeX = lol.x
                    ResizeY = lol.y
                    ResizeType = 1
                    Exit Sub
                End If
            Next
            For i = 2 To 3
                If CornerRects(i).Contains(e.X, e.Y) Then
                    Cursor.Current = Cursors.SizeNESW
                    ResizeX = lol.x
                    ResizeY = lol.y
                    ResizeType = 2
                    Exit Sub
                End If
            Next
            For i = 4 To 5
                If CornerRects(i).Contains(e.X, e.Y) Then
                    Cursor.Current = Cursors.SizeNESW
                    ResizeX = lol.x
                    ResizeY = lol.y
                    ResizeType = 3
                    Exit Sub
                End If
            Next
            For i = 6 To 7
                If CornerRects(i).Contains(e.X, e.Y) Then
                    Cursor.Current = Cursors.SizeNWSE
                    ResizeX = lol.x
                    ResizeY = lol.y
                    ResizeType = 4
                    Exit Sub
                End If
            Next
            If e.Y <= BorderThickness Then
                Cursor.Current = Cursors.SizeNS
                ResizeX = lol.x
                ResizeY = lol.y
                ResizeType = 5
                Exit Sub
            ElseIf e.Y >= (Height - BorderThickness) Then
                Cursor.Current = Cursors.SizeNS
                ResizeX = lol.x
                ResizeY = lol.y
                ResizeType = 6
                Exit Sub
            End If
            If e.X <= BorderThickness Then
                Cursor.Current = Cursors.SizeWE
                ResizeX = lol.x
                ResizeY = lol.y
                ResizeType = 7
                Exit Sub
            ElseIf e.X >= (Height - BorderThickness) Then
                Cursor.Current = Cursors.SizeWE
                ResizeX = lol.x
                ResizeY = lol.y
                ResizeType = 8
                Exit Sub
            End If
        End If
        If e.Y < ControlBar.Height Then
            If ControlBar.ExitButton.Rect.Contains(e.X, e.Y) = False AndAlso ControlBar.MaximizeButton.Rect.Contains(e.X, e.Y) = False AndAlso ControlBar.MinimizeButton.Rect.Contains(e.X, e.Y) = False Then
                If Locked = True Then Exit Sub
                DragX = (e.X + Left)
                DragY = (e.Y + Top)
            End If
        End If

    End Sub

    Private Sub EF_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        DragX = -1
        DragY = -1
        ResizeX = -1
        ResizeY = -1
        If MouseDownX < 0 Or MouseDownY < 0 Or e.X < 0 Or e.Y < 0 Then Exit Sub
        Dim w As Integer = Width, h As Integer = Height, l As Integer = Left, t As Integer = Top
        If ControlBar.ExitButton.Rect.Contains(e.X, e.Y) AndAlso ControlBar.ExitButton.Rect.Contains(MouseDownX, MouseDownY) Then
            Close()
        ElseIf ControlBar.MaximizeButton.Rect.Contains(e.X, e.Y) AndAlso ControlBar.MaximizeButton.Rect.Contains(MouseDownX, MouseDownY) Then
            If Width = OW AndAlso Height = OH Then
                w = Screen.PrimaryScreen.WorkingArea.Width
                h = Screen.PrimaryScreen.WorkingArea.Height
                l = 0
                t = 0
                Locked = True
            Else
                Locked = False
                w = OW
                h = OH
                l = OL
                t = OT
            End If
            SetFormBounds(l, t, w, h)
            ControlBar.SetBar(Width)
            Me.Refresh()
        ElseIf ControlBar.MinimizeButton.Rect.Contains(e.X, e.Y) AndAlso ControlBar.MinimizeButton.Rect.Contains(MouseDownX, MouseDownY) Then
            bMinimizing = True
            WindowState = FormWindowState.Minimized
        End If
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        If _backBuffer Is Nothing Then
            _backBuffer = New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
        End If
        Dim g As Graphics = Graphics.FromImage(_backBuffer)
        g.Clear(SystemColors.Control)
        'Draw Control Box
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
        g.FillRectangle(BarColor, 0, 0, Width, ControlBar.Height)
        If ControlBar.Title <> "" Then g.DrawString(ControlBar.Title, ControlBar.Font, ControlBar.FontBrush, ControlBar.TextLeft, ControlBar.TextTop)
        g.DrawImage(FormBar.bmpCorners(0), 0, 0) 'Makes transparent corner, very small bitmap created at run-time
        g.DrawImage(FormBar.bmpCorners(1), Width - FormBar.bmpCorners(0).Width, 0)
        'Draw Control Box buttons top right
        If ControlBar.ExitButton.Enabled = True Then g.DrawImage(ControlBar.ExitButton.Img, ControlBar.ExitButton.Rect.X, ControlBar.ExitButton.Rect.Y)
        If ControlBar.MaximizeButton.Enabled = True Then g.DrawImage(ControlBar.MaximizeButton.Img, ControlBar.MaximizeButton.Rect.X, ControlBar.MaximizeButton.Rect.Y)
        If ControlBar.MinimizeButton.Enabled = True Then g.DrawImage(ControlBar.MinimizeButton.Img, ControlBar.MinimizeButton.Rect.X, ControlBar.MinimizeButton.Rect.Y)
        If Not ControlBar.Ico Is Nothing Then g.DrawImage(ControlBar.Ico, 5, 5) 'Draw Control Box icon (program icon) if it is set

        'Draw the form border
        For i = 0 To BorderThickness - 1
            g.DrawLine(BorderColor, i, ControlBar.Height, i, Height - 1)
            g.DrawLine(BorderColor, Width - 1 - i, ControlBar.Height, Width - 1 - i, Height - 1)
            g.DrawLine(BorderColor, BorderThickness, Height - 1 - i, Width - BorderThickness, Height - 1 - i)
        Next
        g.Dispose()
        e.Graphics.DrawImageUnscaled(_backBuffer, 0, 0)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
    End Sub

    Public Sub FB_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
        ResizeX = -1
        ResizeY = -1
        DragX = -1
        DragY = -1
    End Sub

    Public Shared Function MeasureAString(ByRef Width0Height1Both2 As Byte, ByRef str As String, ByRef tFont As Font) As Object
        Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
            Dim SFormat As New System.Drawing.StringFormat
            Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
            Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
            SFormat.SetMeasurableCharacterRanges(range)
            Dim regions() As Region = gx.MeasureCharacterRanges(str, tFont, rect, SFormat)
            rect = regions(0).GetBounds(gx)

            If Width0Height1Both2 = 0 Then Return rect.Right + 1 'gx.MeasureString(str, Font, 50000000, lolz).Width
            If Width0Height1Both2 = 1 Then Return rect.Bottom + 1 'gx.MeasureString(str, Font, 50000000, lolz).Height
            If Width0Height1Both2 = 2 Then Return New SizeF(rect.Right + 1, rect.Bottom + 1) 'gx.MeasureString(str, Font, 50000000, lolz)
        End Using
        Return -1
    End Function

    Public Class FormBar
        Public MaximizeButton As Btn, MinimizeButton As Btn, ExitButton As Btn
        Public Font As New Font("Comci Sans MS", 10), ButtonBorderColor As New Pen(Color.Black), ButtonInnerColor As New Pen(Color.White)
        Public TextTop As Integer = 0, TextLeft As Integer = 30, FontBrush As New SolidBrush(Color.Black)
        Public ButtonSize As Integer = 18
        Public Title As String = "", Ico As Bitmap
        Public TC As Color
        Public Width As Integer, Height As Integer, Left As Integer, Top As Integer
        Public Parent As EmptyForm
        Public Shared bmpCorners(1) As Bitmap

        Public Sub New(ByRef prnt As EmptyForm, ByRef sTitle As String, ByRef CFolder As String, Optional ByRef tIco As Bitmap = Nothing, Optional ByRef bMaximize As Boolean = True, Optional ByRef bMinimize As Boolean = True, Optional ByRef bExit As Boolean = True, Optional ByVal tBarHeight As Integer = 25, Optional ByRef x As Integer = 0, Optional ByRef y As Integer = 0)
            If CFolder.Last <> "\" Then CFolder &= "\"
            Parent = prnt
            Title = sTitle
            Left = x
            Top = y
            Width = Parent.Width - Left
            Height = tBarHeight
            Font = New Font("Comic Sans MS", 10)
            TextTop = (Height - MeasureAString(1, Title, Font)) / 2
            TC = prnt.TransparencyKey
            Try
                If tIco Is Nothing Then Ico = Image.FromFile(CFolder & "FormIcon.png") Else Ico = tIco
            Catch ex As Exception
                tIco = Nothing
            End Try
            Dim btnTop As Integer = (Height - ButtonSize) / 2 + 1
            ExitButton = New Btn(Width - ButtonSize - 5, btnTop, Image.FromFile(CFolder & "ExitButton.png"), True)
            MaximizeButton = New Btn(ExitButton.Rect.Left - ButtonSize - 5, btnTop, Image.FromFile(CFolder & "MaximizeButton.png"), True)
            MinimizeButton = New Btn(MaximizeButton.Rect.Left - ButtonSize - 5, btnTop, Image.FromFile(CFolder & "MinimizeButton.png"), True)

            'Set corner transparency
            bmpCorners(0) = New Bitmap(4, 4)
            bmpCorners(0).SetPixel(0, 0, TC)
            bmpCorners(0).SetPixel(1, 0, TC)
            bmpCorners(0).SetPixel(2, 0, TC)
            bmpCorners(0).SetPixel(3, 0, TC)

            bmpCorners(0).SetPixel(0, 1, TC)
            bmpCorners(0).SetPixel(0, 2, TC)
            bmpCorners(0).SetPixel(0, 3, TC)

            bmpCorners(0).SetPixel(1, 1, TC)
            bmpCorners(1) = bmpCorners(0).Clone
            bmpCorners(1).RotateFlip(RotateFlipType.RotateNoneFlipX)
        End Sub

        Public Sub SetBar(ByRef w As Integer)
            Width = w
            ExitButton.Rect.X = Width - ButtonSize - 5
            MaximizeButton.Rect.X = ExitButton.Rect.X - ButtonSize - 5
            MinimizeButton.Rect.X = MaximizeButton.Rect.X - ButtonSize - 5
        End Sub

        Public Class Btn
            Public Img As Bitmap, Rect As Rectangle
            Public Enabled As Boolean = True, Down As Boolean = False
            Public Sub New(ByRef l As Integer, ByRef t As Integer, ByRef tImg As Bitmap, ByRef tEnabled As Boolean)
                Img = tImg
                Rect = New Rectangle(l, t, Img.Width, Img.Height)
            End Sub
        End Class
    End Class
End Class

Public Class DropDownMenu
    Inherits System.Windows.Forms.Control

    Public Event MenuItemClicked(ByVal sender As DropDownMenu, ByVal e As MenuItemClickedArgs)

    Public Items As New List(Of MItem)
    Public Rect As Rectangle
    Public WBuffer As Integer = 5, TBuffer As Integer = 5
    Public LongestWidth As Integer
    Public BoxWidth As Integer = 200, BoxHeight As Integer = 30
    Public TextHeight As Integer = 0, TextTop As Integer = 0
    Public ItemBorders As Boolean = True, BorderColor As New Pen(Color.DarkGray), BorderThickness As Integer = 1
    Public FontColor As SolidBrush
    Public HighlightBuffer As Integer
    Public LColumnWidth As Integer = 30
    Public InactiveFontColor As SolidBrush
    Public Shared BGColor As New SolidBrush(Color.White), HighlightColor As New SolidBrush(Color.FromArgb(150, Color.Gray)), HighlightBorderColor As New Pen(Color.Gray)
    Public Shared ActiveDDM As DropDownMenu
    Public Shared ContentFolder As String
    Public Shared OMGLol As Integer = 0
    Private HighlightedItem As Integer = -1, CheckBufferW As Integer, CheckBufferH As Integer

    Public Sub New(ByRef prnt As Form, ByRef tItem As String, Optional ByRef Checkable As Boolean = False, Optional ByRef sTag As Object = Nothing)
        Parent = prnt
        Width = 0
        Height = 0
        SetFont(New Font("Comic Sans MS", 10, FontStyle.Regular), Color.Black)
        TextHeight = MeasureAString(1, "yI", Font)
        TextTop = (BoxHeight - TextHeight) / 2
        MItem.CheckedImg = Image.FromFile(ContentFolder & "Checkmark.png")
        MItem.CheckedImg.MakeTransparent(Color.FromArgb(255, 255, 0, 128))
        CheckBufferW = (LColumnWidth - MItem.CheckedImg.Width) / 2
        CheckBufferH = (BoxHeight - MItem.CheckedImg.Height) / 2
        AddItem(tItem, Checkable, sTag)
        Visible = False
        AddHandler Me.Paint, AddressOf DDM_Paint
        AddHandler Me.MouseUp, AddressOf DDM_MouseUp
        AddHandler Me.MouseMove, AddressOf DDM_MouseMove
        AddHandler Me.MouseLeave, AddressOf DDM_MouseLeave
        AddHandler Me.LostFocus, AddressOf DDM_LostFocus
    End Sub

    Public Sub SetFont(ByRef tFont As Font, Optional ByRef tFontColor As Color = Nothing)
        Font = tFont
        If tFontColor = Nothing Then FontColor = New SolidBrush(Color.Black) Else FontColor = New SolidBrush(tFontColor)
        InactiveFontColor = New SolidBrush(Color.FromArgb(100, FontColor.Color))
    End Sub

    Private Sub DDM_MouseLeave(ByVal sender As Object, ByVal e As EventArgs)
        HighlightedItem = -1
        Me.Refresh()
    End Sub

    Private Sub DDM_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
        Dim b As Integer = HighlightedItem
        If Me.ClientRectangle.Contains(e.X, e.Y) = True Then
            Dim index As Integer = Math.Floor(e.Y / BoxHeight)
            HighlightedItem = index
        Else
            HighlightedItem = -1
        End If
        If b <> HighlightedItem Then Me.Refresh()
    End Sub

    Private Sub DDM_MouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        If Me.ClientRectangle.Contains(e.X, e.Y) = True Then
            Dim ev As MenuItemClickedArgs
            Dim index As Integer = Math.Floor(e.Y / BoxHeight)
            If Items(index).Active = False Then Exit Sub
            If Items(index).Checkable = True Then
                Items(index).Checked = Not Items(index).Checked
                ev = New MenuItemClickedArgs(Items(index), index, True)
            Else
                ev = New MenuItemClickedArgs(Items(index), index, False)
            End If
            RaiseEvent MenuItemClicked(Me, ev)
        End If
        Dismiss()
    End Sub

    Public Function GetItemByName(ByRef sName As String) As MItem
        Dim sName2 As String = sName.ToUpper
        For i = 0 To Items.Count - 1
            If Items(i).Text.ToUpper = sName2 Then Return Items(i)
        Next
        Return Nothing
    End Function

    Private Sub DDM_LostFocus(ByVal sender As Object, ByVal e As EventArgs)
        Dismiss()
    End Sub

    Public Sub Dismiss()
        ActiveDDM = Nothing
        Me.Visible = False
    End Sub

    Public Sub ShowF(ByRef l As Integer, ByRef t As Integer)
        If Visible = True Then
            Dim omg As String = ""
        End If
        Dim HModifier As Integer = 0, WModifier As Integer = 0, h As Integer = Items.Count * BoxHeight
        If TypeOf Parent Is Form Then
            Dim lol As Form = Parent
            If lol.MaximizeBox Or lol.MinimizeBox Then
                HModifier = 38
            End If
            If lol.FormBorderStyle = FormBorderStyle.Sizable Then
                'WModifier = 16
            End If
        End If
        If l + BoxWidth > Parent.Width - WModifier Then l = Parent.Width - BoxWidth - WModifier
        If t + h > Parent.Height - HModifier Then t = Parent.Height - HModifier - h

        SetBounds(l, t, BoxWidth, h)
        ActiveDDM = Me
        Visible = True
        BringToFront()
        Me.Focus()
    End Sub

    Public Sub AddItem(ByRef sText As String, Optional ByRef Checkable As Boolean = False, Optional ByRef tTag As Object = Nothing)
        Dim newItem As MItem = Nothing
        If Items.Count = 0 Then
            newItem = New MItem(Nothing, sText, Checkable, tTag)
        Else
            newItem = New MItem(Items(Items.Count - 1), sText, Checkable, tTag)
        End If
        Items.Add(newItem)
        Height = Items.Count * BoxHeight
    End Sub

    Private Sub DDM_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        OMGLol += 1
        For i = 0 To Items.Count - 1
            If ItemBorders = True Then
                e.Graphics.DrawRectangle(BorderColor, LColumnWidth, i * BoxHeight, BoxWidth - LColumnWidth - 1, BoxHeight - 1)
                e.Graphics.DrawRectangle(BorderColor, 0, 0, LColumnWidth, Height - 1)
            End If
            If Items(i).Active = True Then e.Graphics.DrawString(Items(i).Text, Font, FontColor, LColumnWidth + WBuffer, (i * BoxHeight) + TextTop) Else e.Graphics.DrawString(Items(i).Text, Font, InactiveFontColor, LColumnWidth + WBuffer, (i * BoxHeight) + TextTop)
            If HighlightedItem = i Then
                Dim hr As New Rectangle(BorderThickness + HighlightBuffer, (i * BoxHeight) + TextTop, Width - (BorderThickness * 2) - HighlightBuffer - 1, TextHeight)
                e.Graphics.DrawRectangle(HighlightBorderColor, hr)
                e.Graphics.FillRectangle(HighlightColor, hr.Left + 2, hr.Top + 2, hr.Width - 3, hr.Height - 3)
            End If
            If Items(i).Checked = True Then e.Graphics.DrawImage(MItem.CheckedImg, CheckBufferW, (i * BoxHeight) + CheckBufferH)
        Next
        e.Graphics.DrawLine(BorderColor, LColumnWidth, 0, LColumnWidth, Height - 1)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
        e.Graphics.FillRectangle(BGColor, Me.ClientRectangle)
    End Sub

    Public Overloads Sub AppendItems(ByVal sNames As List(Of String))
        For i = 0 To sNames.Count - 1
            Dim newItem As MItem
            If Items.Count = 0 Then newItem = New MItem(Nothing, sNames(i)) Else newItem = New MItem(Items(Items.Count - 1), sNames(i))
            Items.Add(newItem)
        Next
    End Sub

    Public Overloads Sub AppendItems(ByVal arrItems As List(Of MItem), Optional ByRef InsertIndex As Integer = -1)
        If InsertIndex = -1 Then
            If Items.Count > 0 Then
                arrItems(0).PItem = Items(Items.Count - 1)
                Items(Items.Count - 1).NItem = arrItems(0)
            End If
            For i = 0 To arrItems.Count - 1
                If i > 0 AndAlso arrItems(i).PItem Is Nothing Then
                    arrItems(i).PItem = arrItems(i - 1)
                    arrItems(i - 1).NItem = arrItems(i)
                End If
                Items.Add(arrItems(i))
            Next
        End If
        Height = Items.Count * BoxHeight
    End Sub

    Public Class MenuItemClickedArgs
        Inherits System.EventArgs
        Public Item As MItem, index As Integer, CheckChanged As Boolean
        Public Sub New(ByRef tItem As MItem, ByRef tIndex As Integer, ByRef tCheckChanged As Boolean)
            Item = tItem
            index = tIndex
            CheckChanged = tCheckChanged
        End Sub
    End Class

    Public Class MItem
        Public Shared CheckedImg As Bitmap
        Public Checked As Boolean = False, Checkable As Boolean = False, Active As Boolean = True
        Public ico As Bitmap
        Public Text As String, Tag As Object
        Public PItem As MItem, NItem As MItem, SubItem As MItem

        Public Sub New(ByRef PreviousItem As MItem, ByRef sText As String, Optional ByRef tCheckable As Boolean = False, Optional ByVal tTag As Object = Nothing)
            If Not PItem Is Nothing Then PItem.NItem = Me
            PItem = PreviousItem
            Text = sText
            Checkable = tCheckable
            Tag = tTag
        End Sub
    End Class
End Class

Public Class ScrollBar
    Inherits System.Windows.Forms.Control
    Implements IDisposable

    Private Shadows disposed As Boolean = False
    Protected Overrides Sub Dispose( _
       ByVal disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                Parent = Nothing
                img = Nothing
                ScrollArrows(0) = Nothing
                ScrollArrows(1) = Nothing
                If Not tmr Is Nothing Then
                    tmr.Abort()
                    If Not tmr Is Nothing Then tmr = Nothing
                End If
                For i = 0 To Delegates.Count - 1
                    RemoveHandler ScrollBarMoved, Delegates(0)
                Next
            End If
        End If
        Me.disposed = True
    End Sub

    Public Custom Event ScrollBarMoved As EventHandler '(ByRef sbSender As ScrollBar, ByVal e As BarMoveArgs)
        AddHandler(ByVal value As EventHandler)
            Me.Events.AddHandler("ScrollBarMovedEvent", value)
            Delegates.Add(value)
        End AddHandler
        RemoveHandler(ByVal value As eventhandler)
            Me.Events.RemoveHandler("ScrollBarMovedEvent", value)
            Delegates.Remove(value)
        End RemoveHandler

        RaiseEvent(ByVal sender As Object, ByVal e As BarMoveArgs)
            CType(Me.Events("ScrollBarMovedEvent"), EventHandler).Invoke(sender, e)
        End RaiseEvent
    End Event

    Private Delegates As New ArrayList
    Private ScrollDownX As Integer = 0, ScrollDownY As Integer = 0, Covered As Integer = 0
    Private ScrollRect As Rectangle
    Public BackBarColor As New SolidBrush(Color.FromArgb(255, 214, 231, 236)), BarOutlineColor As Color = Color.DarkBlue, BarColor As Color = Color.LightBlue
    Public min As Integer, max As Integer
    Public BType As Integer '0 = horizontal, 1 = vertical
    Public FieldMax As Single
    Public BarSize As Integer
    Public ScrollArrows(1) As ScrollArrow
    Public ScrollIncrement As Integer = 0
    Public ScrollPosition As Integer = 0
    Public TimerDelayDefault As Integer = 400
    Private tmr As Thread
    Private TimerDelay As Integer, TimerInterval As Integer = 50

    Private img As Bitmap
    Private TimerEventType As Integer = 0
    Private NotShown As Integer = 0
    Private BarAreaLeft As Integer = 0, MinBarSize As Integer = 35
    Private MouseX As Integer, MouseY As Integer
    Private blnFirstMove As Boolean = False

    Public Sub New(ByRef prnt As Control, ByRef tX As Integer, ByVal tY As Integer, ByRef tW As Integer, ByRef tH As Integer, ByRef iType As Integer, ByRef FieldSize As Single, Optional ByRef tBarSize As Integer = 16)
        Parent = prnt
        Left = tX
        Top = tY
        Width = tW
        Height = tH
        BType = iType
        BarSize = tBarSize
        FieldMax = FieldSize
        DoubleBuffered = True
        If ScrollArrow.Arrows Is Nothing Then SetArrows()
        Dim BorderPen As New Pen(BarOutlineColor)
        If iType = 0 Then 'Horizontal
            min = Left + BarSize
            max = Width - BarSize
            ScrollArrows(0) = New ScrollArrow(0, 0, 0)
            ScrollArrows(1) = New ScrollArrow(Width - BarSize, 0, 1)
            Dim Shown As Single = CSng(tW) / FieldMax
            NotShown = FieldMax - tW
            ScrollIncrement = (NotShown / 10)
            img = New Bitmap(Math.Max(CInt(Shown * (max - min)), MinBarSize), BarSize)
            BarAreaLeft = (max - min) - img.Width

            Using gx As Graphics = Graphics.FromImage(img)
                gx.FillRectangle(New SolidBrush(BarColor), 2, 1, img.Width - 4, img.Height - 2)
                gx.DrawLine(New Pen(BarColor), 1, 3, 1, img.Height - 3)
                gx.DrawLine(New Pen(BarColor), img.Width - 2, 3, img.Width - 2, img.Height - 3)

                gx.DrawLine(New Pen(BarOutlineColor), 3, 0, img.Width - 4, 0)
                gx.DrawLine(New Pen(BarOutlineColor), 3, img.Height - 1, img.Width - 4, img.Height - 1)
                gx.DrawLine(New Pen(BarOutlineColor), 0, 3, 0, img.Height - 4)
                gx.DrawLine(New Pen(BarOutlineColor), img.Width - 1, 3, img.Width - 1, img.Height - 4)
                'FillArea(img, img.Width / 2, img.Height / 2, BarColor, Color.FromArgb(0, 0, 0, 0))
                Dim mid As Integer = img.Width / 2
                gx.DrawLine(New Pen(BarOutlineColor), mid, 3, mid, img.Height - 4)
                For i = 1 To 1
                    gx.DrawLine(New Pen(BarOutlineColor), mid + (i * 2), 3, mid + (i * 2), img.Height - 4)
                    gx.DrawLine(New Pen(Color.FromArgb(50, BarOutlineColor)), mid + (i * 2) - 1, 3, mid + (i * 2) - 1, img.Height - 4)
                    gx.DrawLine(New Pen(BarOutlineColor), mid - (i * 2), 3, mid - (i * 2), img.Height - 4)
                    gx.DrawLine(New Pen(Color.FromArgb(50, BarOutlineColor)), mid - (i * 2) + 1, 3, mid - (i * 2) + 1, img.Height - 4)
                Next
            End Using
            ScrollRect = New Rectangle(min, 0, img.Width, img.Height)
        Else 'Vertical

        End If
        img.SetPixel(1, 2, BarOutlineColor)
        img.SetPixel(2, 1, BarOutlineColor)
        img.SetPixel(img.Width - 2, 2, BarOutlineColor)
        img.SetPixel(img.Width - 3, 1, BarOutlineColor)
        img.SetPixel(1, img.Height - 3, BarOutlineColor)
        img.SetPixel(2, img.Height - 2, BarOutlineColor)
        img.SetPixel(img.Width - 2, img.Height - 3, BarOutlineColor)
        img.SetPixel(img.Width - 3, img.Height - 2, BarOutlineColor)


        AddHandler Me.MouseDown, AddressOf Scroll_Click
        AddHandler Me.Paint, AddressOf SB_Paint
        AddHandler Me.MouseMove, AddressOf SB_MouseMove
    End Sub

    Private Sub SB_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)
        MouseX = e.X
        MouseY = e.Y
        If TimerEventType = 2 Then
            If MouseButtons <> Windows.Forms.MouseButtons.Left Then
                TimerEventType = -1
                Exit Sub
            End If
            Dim OP As Integer = ScrollPosition
            ScrollPosition = Math.Max(Covered + ((ScrollDownX - MouseX) / BarAreaLeft) * (FieldMax - Width), -1 * NotShown)
            ScrollPosition = Math.Min(ScrollPosition, 0)
            If OP <> ScrollPosition Then MoveBar(OP, min + ((ScrollPosition / NotShown) * -1 * BarAreaLeft))
        End If
    End Sub

    Private Sub SB_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
        For i = 0 To 1
            e.Graphics.DrawImage(ScrollArrow.Arrows(i + (BType * 2)), ScrollArrows(i).x, ScrollArrows(i).y)
        Next
        e.Graphics.DrawImage(img, ScrollRect.X, ScrollRect.Y)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As System.Windows.Forms.PaintEventArgs)
        If Not ScrollRect = Nothing Then
            e.Graphics.FillRectangle(BackBarColor, 0, 0, ScrollRect.X + 3, BarSize)
            e.Graphics.FillRectangle(BackBarColor, ScrollRect.X + img.Width - 3, 0, Width - ScrollRect.X - img.Width + 3, BarSize)
        Else
            e.Graphics.FillRectangle(BackBarColor, Me.ClientRectangle)
        End If
    End Sub

    Private Sub tmr_tick(ByVal delay As Integer)
        If delay > 0 Then System.Threading.Thread.Sleep(delay)
        If MouseButtons <> Windows.Forms.MouseButtons.Left Then
            TimerEventType = -1
            System.Threading.Thread.CurrentThread.Abort()
        End If
        Dim OP As Integer = ScrollPosition
        Select Case TimerEventType
            Case 0
                ScrollPosition = Math.Min(ScrollPosition + ScrollIncrement, 0)
            Case 1
                ScrollPosition = Math.Max(ScrollPosition - ScrollIncrement, -1 * NotShown)
        End Select
        If OP <> ScrollPosition Then
            ScrollRect.X = min + ((ScrollPosition / NotShown) * -1 * BarAreaLeft)
            Me.Invoke(New Action(Of String)(AddressOf UpdateBar), CStr(OP))
        End If
        tmr_tick(50)
    End Sub

    Private Sub UpdateBar(ByVal OP As Integer)
        RaiseEvent ScrollBarMoved(Me, New BarMoveArgs(ScrollPosition - OP))
        Me.Refresh()
    End Sub

    Public Sub UpdateScroll(ByRef position As Integer)
        ScrollPosition = position
        MoveBar(0, min + ((ScrollPosition / NotShown) * -1 * BarAreaLeft))
    End Sub

    Public Sub MoveBar(ByRef OP As Integer, ByRef newPosition As Integer)
        ScrollRect.X = newPosition
        RaiseEvent ScrollBarMoved(Me, New BarMoveArgs(ScrollPosition - OP))
        Me.Refresh()
    End Sub

    Public Class BarMoveArgs
        Inherits System.EventArgs
        Public difference As Integer
        Public Sub New(ByRef dif As Integer)
            difference = dif
        End Sub
    End Class

    Public Class ArrowClicked
        Public Arrow As ScrollArrow
        Public Sub New(ByRef tArrow As ScrollArrow)
            Arrow = tArrow
        End Sub
    End Class

    Public Sub ScrollInDirection(ByRef direction As Integer)
        Dim OP As Integer = ScrollPosition
        If direction = 0 Then
            ScrollPosition = Math.Min(ScrollPosition + ScrollIncrement, 0)
        Else
            ScrollPosition = Math.Max(ScrollPosition - ScrollIncrement, -1 * NotShown)
        End If
        MoveBar(OP, min + ((ScrollPosition / NotShown) * -1 * BarAreaLeft))
    End Sub

    Private Sub Scroll_Click(ByVal scrSender As ScrollBar, ByVal e As MouseEventArgs)
        Dim OP As Integer = ScrollPosition
        blnFirstMove = True
        TimerEventType = -1
        For i = 0 To 1
            If (e.X >= ScrollArrows(i).x AndAlso e.X < ScrollArrows(i).x + BarSize) AndAlso (e.Y >= ScrollArrows(i).y AndAlso e.Y < ScrollArrows(i).y + BarSize) Then
                If i = 0 Then
                    ScrollPosition = Math.Min(ScrollPosition + ScrollIncrement, 0)
                    TimerEventType = 0
                Else
                    TimerEventType = 1
                    ScrollPosition = Math.Max(ScrollPosition - ScrollIncrement, -1 * NotShown)
                End If
                MoveBar(OP, min + ((ScrollPosition / NotShown) * -1 * BarAreaLeft))
                Dim newThread As New Thread(AddressOf tmr_tick)
                newThread.IsBackground = True
                newThread.Start(TimerDelayDefault)
                Exit Sub
            End If
        Next
        If ScrollRect.Contains(e.X, e.Y) Then
            ScrollDownX = e.X
            ScrollDownY = e.Y
            Covered = ScrollPosition
            TimerEventType = 2
        End If
    End Sub

    Private Sub SetArrows()
        ScrollArrow.Arrows = New List(Of Bitmap)
        Dim bmpArrow As New Bitmap(BarSize, BarSize)
        Using gx As Graphics = Graphics.FromImage(bmpArrow)
            'gx.FillRectangle(New SolidBrush(Color.FromArgb(255, 255, 0, 128)), 0, 0, bmpArrow.Width, bmpArrow.Height)
            Dim mid As Integer = BarSize / 2
            gx.DrawLine(ScrollArrow.ArrowColor, 0, mid, BarSize - 1, 0)
            gx.DrawLine(ScrollArrow.ArrowColor, 0, mid, BarSize - 1, BarSize - 1)
            gx.DrawLine(ScrollArrow.ArrowColor, BarSize - 1, 0, BarSize - 1, BarSize - 1)
        End Using
        FillArea(bmpArrow, bmpArrow.Width - 2, bmpArrow.Height / 2, ScrollArrow.ArrowColor.Color, Color.FromArgb(0, 0, 0, 0))
        ScrollArrow.Arrows.Add(bmpArrow.Clone)
        bmpArrow.RotateFlip(RotateFlipType.RotateNoneFlipX)
        ScrollArrow.Arrows.Add(bmpArrow.Clone)
    End Sub

    Public Class ScrollArrow
        Public x As Integer, y As Integer, Tag As Integer
        Public Shared ArrowColor As New Pen(Color.DarkBlue)
        Public Shared Arrows As List(Of Bitmap)

        Public Sub New(ByVal tX As Integer, ByVal tY As Integer, ByRef tTag As Integer)
            x = tX
            y = tY
            Tag = tTag
        End Sub
    End Class
End Class

Public Class TabBrowser
    Inherits System.Windows.Forms.Control

    Event AddBoxClicked(ByRef sender As TabBrowser, ByVal e As EventArgs)
    Event PageChanged(ByRef sender As TabBrowser, ByVal e As EventArgs)

    Public BorderColor As Pen
    Public FontColor As SolidBrush
    Public BorderThickness As Integer
    Public Pages As New List(Of Page)
    Public SelectedPage As Page
    Public SelectedIndex As Integer = -1
    Public CenterTabText As Boolean = False, EnableAddBox As Boolean = True
    Public LastPage As Page
    Public HHeight As Integer
    Public Shared HighlightColor As SolidBrush

    Private TabHeight As Integer = 20
    Private TabWidth As Integer = 100, TabWidthMax = 100, TabWidthMin = 20, TabLBuffer As Integer = 3
    Private DoubleBorder As Integer
    Private ExitButtonXW As Integer = 7, ExitButtonXH As Integer = 7, ExitButtonW As Integer = 11, ExitButtonH As Integer = 11, ExitButtonTBuffer As Integer, ExitButtonLBuffer As Integer, ExitButtonColor As New Pen(Color.Black)
    Private HighlightedExitButton As Integer = -1, HighlightedHeader As Integer = -1
    Private AddBox As NewPageTab
    Private ScrollArrows(1) As ScrollArrow
    Private ControlBrush As New SolidBrush(Color.FromKnownColor(KnownColor.Control))
    Private ItemsShownMax As Integer = 0, TrueTabHeaderWidth As Integer = 0
    Private MinBuffer As Single, ScrollPosition As Integer = 0
    Private TabOrientation As Byte = 0 '0 = left oriented, 1 = right oriented

    Public Sub New(ByRef tbName As String, ByRef prnt As Control, ByRef l As Integer, ByRef t As Integer, ByRef w As Integer, ByRef h As Integer, Optional ByRef sPage As String = "Blank")
        Name = tbName
        If HighlightColor Is Nothing Then HighlightColor = New SolidBrush(Color.FromArgb(60, Color.LightBlue))
        Parent = prnt
        HHeight = h - TabHeight
        SetBounds(l, t, w, TabHeight)
        Font = New Font("Microsoft Sans Serif", 8.25)
        FontColor = New SolidBrush(Parent.ForeColor)
        SetBorderStyle(1, Color.DarkBlue)
        ScrollArrows(0) = New ScrollArrow(0, 0, TabHeight, 0)
        SetAddBoxSettings(True)

        ExitButtonTBuffer = (TabHeight - ExitButtonH) / 2 - 1
        TrueTabHeaderWidth = Width - (ScrollArrows(0).rect.Width * 2) - TabHeight
        Dim omg As Single = TrueTabHeaderWidth / TabWidthMin
        ItemsShownMax = Math.Ceiling(omg)
        MinBuffer = Math.Ceiling((ItemsShownMax - omg) * TabWidthMin)
        If MinBuffer = 0 Then MinBuffer = TabWidthMin

        AddPage(sPage, Nothing)
    End Sub

    Public Sub SetBorderStyle(ByRef thickness As Integer, Optional ByRef clr As Color = Nothing)
        BorderThickness = thickness
        DoubleBorder = BorderThickness * 2
        If clr <> Nothing Then BorderColor = New Pen(clr)
    End Sub

    Public Sub SetAddBoxSettings(ByRef enabled As Boolean)
        EnableAddBox = enabled
        Dim vis As Boolean = False
        If Not ScrollArrows(1) Is Nothing Then
            vis = ScrollArrows(1).visible
        End If
        If enabled = True Then
            ScrollArrows(1) = New ScrollArrow(Width - (TabHeight * 2), 0, TabHeight, 0)
        Else
            ScrollArrows(1) = New ScrollArrow(Width - (TabHeight * 2), 0, TabHeight, 0)
        End If
        ScrollArrows(1).visible = vis
    End Sub

    Public Function IsSelected(ByRef tPage As Page) As Boolean
        If SelectedPage Is tPage Then Return True
        Return False
    End Function

    Public Sub SetPageName(ByRef tPage As Page, ByRef sName As String)
        tPage.Header = New Page.TabHeader(tPage, sName)
        Me.Refresh()
    End Sub

    Public Sub AddPage(ByRef sName As String, Optional ByRef tFont As Font = Nothing, Optional ByRef tFontColor As Color = Nothing, Optional ByRef tBGColor As Color = Nothing, Optional ByRef bRefresh As Boolean = True, Optional ByRef bChangeSelection As Boolean = True)
        If Pages.Count = 1 AndAlso Pages(0).Header.Text = "Blank" AndAlso Pages(0).Controls.Count = 0 Then
            Pages(0).Dispose()
            If Not Pages(0) Is Nothing Then Pages(0) = Nothing
            Pages.RemoveAt(0)
        End If
        Pages.Add(New Page(Me, sName, tFont, tFontColor, tBGColor))
        LastPage = Pages(Pages.Count - 1)
        If SelectedIndex = -1 Or SelectedPage Is Nothing Then
            SelectedIndex = 0
            SelectedPage = Pages(0)
        Else
            If bChangeSelection = True Then
                SelectedIndex = Pages.Count - 1
                SelectedPage = LastPage
                SelectedPage.BringToFront()
            End If
        End If
        UpdateView()
        If ScrollArrows(0).visible = True AndAlso (SelectedIndex * TabWidth) + TabWidth + ScrollPosition > ScrollArrows(1).rect.Left Then
            ScrollPosition = -1 * ((Pages.Count * TabWidth) - TrueTabHeaderWidth - ScrollArrows(0).rect.Width)
        End If
        If bRefresh = True Then Refresh()
    End Sub

    Public Sub SelectPage(ByRef index As Integer, Optional ByRef bRefresh As Boolean = True)
        If SelectedIndex = index Then Exit Sub
        SelectedIndex = index
        SelectedPage = Pages(index)
        SelectedPage.BringToFront()
        UpdateView()
        RaiseEvent PageChanged(Me, New EventArgs)
        If bRefresh = True Then Me.Refresh()
    End Sub

    Private Sub TB_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles Me.MouseLeave
        Dim OHH As Integer = HighlightedHeader, OHEB As Integer = HighlightedExitButton
        HighlightedExitButton = -1
        HighlightedHeader = -1
        If HighlightedExitButton <> OHEB Or HighlightedHeader <> OHH Then Me.Refresh()
    End Sub

    Private Sub TB_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseMove
        If e.Y > TabHeight Then Exit Sub
        Dim index As Integer = Math.Floor((e.X - ScrollPosition) / TabWidth)
        Dim OHH As Integer = HighlightedHeader, OHEB As Integer = HighlightedExitButton
        HighlightedHeader = index
        If e.X >= (index * TabWidth) + ExitButtonLBuffer + ScrollPosition AndAlso e.X <= (index * TabWidth) + TabWidth - 1 + ScrollPosition And e.Y >= ExitButtonTBuffer And e.Y <= ExitButtonTBuffer + ExitButtonH Then
            If HighlightedExitButton <> index Then HighlightedExitButton = index
        Else
            If HighlightedExitButton <> -1 Then HighlightedExitButton = -1
        End If
        If HighlightedExitButton <> OHEB Or HighlightedHeader <> OHH Then Me.Refresh()
    End Sub

    Private Sub TB_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Me.MouseDown
        If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
        If e.Y > TabHeight Then Exit Sub
        If AddBox.rect.Contains(e.X, e.Y) Then
            RaiseEvent AddBoxClicked(Me, New EventArgs)
            Exit Sub
        End If
        If ScrollArrows(0).visible = True Then
            For i = 0 To 1
                If ScrollArrows(i).rect.Contains(e.X, e.Y) Then
                    If i = 0 Then 'left arrow clicked
                        If TabOrientation = 1 Then ScrollPosition += MinBuffer Else ScrollPosition += TabWidthMin
                        ScrollPosition = Math.Min(ScrollPosition, ScrollArrows(0).rect.Width)
                        TabOrientation = 0
                    Else 'right arrow clicked
                        If TabOrientation = 0 Then ScrollPosition -= MinBuffer Else ScrollPosition -= TabWidthMin
                        ScrollPosition = Math.Max(ScrollPosition, -1 * ((Pages.Count * TabWidth) - TrueTabHeaderWidth - ScrollArrows(0).rect.Width))
                        TabOrientation = 1
                    End If
                    Me.Refresh()
                    Exit Sub
                End If
            Next
        End If
        Dim index As Integer = Math.Floor((e.X - ScrollPosition) / TabWidth)
        If index >= Pages.Count Or index = SelectedIndex Then Exit Sub
        If Pages.Count > 1 Then
            If e.X >= (index * TabWidth) + ExitButtonLBuffer + ScrollPosition AndAlso e.X <= (index * TabWidth) + TabWidth - 1 + ScrollPosition And e.Y >= ExitButtonTBuffer And e.Y <= ExitButtonTBuffer + ExitButtonH Then
                RemovePage(index)
                Exit Sub
            End If
        End If
        SelectPage(index, True)
    End Sub

    Public Sub RemovePage(ByRef index As Integer)
        Dim pg As Page = Pages(index)
        For Each cntrl As Control In pg.Controls
            RecursiveDispose(cntrl)
        Next

        Pages(index).Dispose()
        If Not Pages(index) Is Nothing Then Pages(index) = Nothing
        Pages.RemoveAt(index)
        LastPage = Pages(Pages.Count - 1)

        Dim newIndex As Integer = SelectedIndex
        SelectedIndex = -1

        If newIndex = index Then
            If index = 0 Then
                SelectPage(0, True)
            Else
                SelectPage(index - 1, True)
            End If
        Else
            If index < newIndex Then
                SelectPage(newIndex - 1, True)
            Else
                UpdateView()
                Me.Refresh()
            End If
        End If
    End Sub

    Private Sub RecursiveDispose(ByRef tControl As Control)
        For Each cntrl As Control In tControl.Controls
            RecursiveDispose(cntrl)
        Next
        tControl.Parent.Controls.Remove(tControl)
        tControl.Dispose()
        If Not tControl Is Nothing Then tControl = Nothing
    End Sub

    Private Sub UpdateView()
        Dim SM As Integer = TabWidthMax * Pages.Count + TabHeight - 1
        If SM > Width Then
            TabWidth = ((Width - TabHeight) / (TabWidthMax * Pages.Count)) * TabWidthMax
            AddBox = New NewPageTab(Width - TabHeight - 1, 0, TabHeight)
            If TabWidth < TabWidthMin Then
                TabWidth = TabWidthMin
                ScrollArrows(0).visible = True
                ScrollArrows(1).visible = True
                If ScrollPosition = 0 Then ScrollPosition = ScrollArrows(0).rect.Width
            Else
                If ScrollPosition <> 0 Then
                    ScrollArrows(0).visible = False
                    ScrollArrows(1).visible = False
                    ScrollPosition = 0
                End If
            End If
        Else
            TabWidth = TabWidthMax
            AddBox = New NewPageTab(Pages.Count * TabWidth, 0, TabHeight)
        End If
        ExitButtonLBuffer = TabWidth - ExitButtonW - 3
        If AddBox.rect.Right >= Width - 2 Then AddBox.rect.X = Width - AddBox.rect.Width - 1
    End Sub

    Private Sub TB_Paint(ByVal sender As Object, ByVal e As PaintEventArgs) Handles Me.Paint
        e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
        Dim CX As Integer = 1 + ScrollPosition
        e.Graphics.DrawLine(BorderColor, 1, TabHeight - 1, Width - 1, TabHeight - 1) 'Bottom tab line
        e.Graphics.FillRectangle(ControlBrush, 1, 1, Width - 2, TabHeight - 2)
        For i = 0 To Pages.Count - 1
            Dim rect As Rectangle = New Rectangle(CX, 1, TabWidth, TabHeight - 2)
            If CX + TabWidth > 0 Then
                e.Graphics.FillRectangle(Pages(i).BGColor, rect)
                e.Graphics.DrawString(Pages(i).Header.Text, Pages(i).Font, Pages(i).Header.clr, rect.Left + Pages(i).Header.TextRect.Left, Pages(i).Header.TextRect.Top)
                If HighlightedHeader = i Then e.Graphics.FillRectangle(HighlightColor, CX, 1, TabWidth, TabHeight - 2)
                Dim ER As Rectangle = New Rectangle(CX + ExitButtonLBuffer, ExitButtonTBuffer, ExitButtonW, ExitButtonH)
                If HighlightedExitButton = i Then e.Graphics.DrawRectangle(BorderColor, ER)
                e.Graphics.DrawLine(BorderColor, ER.Left + 2, ER.Top + 2, ER.Right - 2, ER.Bottom - 2)
                e.Graphics.DrawLine(BorderColor, ER.Right - 2, ER.Top + 2, ER.Left + 2, ER.Bottom - 2)
            End If
            CX += TabWidth
            If (ScrollArrows(0).visible = True AndAlso CX > ScrollArrows(1).rect.Left) Or (CX > Width - AddBox.rect.Width) Then  Else e.Graphics.DrawLine(BorderColor, rect.Right - 1, 1, rect.Right - 1, TabHeight - 1)
        Next
        e.Graphics.DrawLine(New Pen(SelectedPage.BGColor.Color), SelectedIndex * TabWidth + 1, TabHeight - 1, (SelectedIndex * TabWidth) + TabWidth - 1, TabHeight - 1)

        e.Graphics.DrawLine(BorderColor, 1, 0, Pages.Count * TabWidth, 0)
        For i = 0 To BorderThickness - 1
            e.Graphics.DrawLine(BorderColor, i, 0, i, Height - 1)
            e.Graphics.DrawLine(BorderColor, Width - 1 - i, TabHeight, Width - 1 - i, Height - 1)
        Next
        If AddBox.rect.Right < Width - 1 Then e.Graphics.FillRectangle(ControlBrush, AddBox.rect.Right + 1, 0, Width - AddBox.rect.Right - 1, TabHeight - 1)
        If EnableAddBox = True Then
            e.Graphics.FillRectangle(ControlBrush, AddBox.rect)
            e.Graphics.DrawRectangle(BorderColor, AddBox.rect)
            e.Graphics.DrawLine(BorderColor, AddBox.L1X1, AddBox.L1Y, AddBox.L1X2, AddBox.L1Y)
            e.Graphics.DrawLine(BorderColor, AddBox.L2X, AddBox.L2Y1, AddBox.L2X, AddBox.L2Y2)
        End If

        If ScrollArrows(0).visible = True Then
            For i = 0 To 1
                e.Graphics.FillRectangle(ControlBrush, ScrollArrows(i).rect)
                e.Graphics.DrawRectangle(BorderColor, ScrollArrows(i).rect)
            Next
            e.Graphics.DrawLine(BorderColor, ScrollArrows(0).X1, ScrollArrows(0).Y2, ScrollArrows(0).X2, ScrollArrows(0).Y1)
            e.Graphics.DrawLine(BorderColor, ScrollArrows(0).X1, ScrollArrows(0).Y2, ScrollArrows(0).X2, ScrollArrows(0).Y3)
            e.Graphics.DrawLine(BorderColor, ScrollArrows(1).X2, ScrollArrows(1).Y2, ScrollArrows(1).X1, ScrollArrows(1).Y1)
            e.Graphics.DrawLine(BorderColor, ScrollArrows(1).X2, ScrollArrows(1).Y2, ScrollArrows(1).X1, ScrollArrows(1).Y3)
        End If
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
    End Sub

    Public Sub SetBoundsTo(ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
        HHeight = h - TabHeight
        SetBounds(x, y, w, TabHeight)
        For i = 0 To Pages.Count - 1
            Pages(i).SetBoundsTo(Left, Bottom, Width, HHeight)
        Next
        SelectedPage.Refresh()
    End Sub

    Public Shared Function MeasureAString(ByRef Width0Height1Both2 As Byte, ByRef str As String, ByRef tFont As Font) As Object
        Using gx As Graphics = Graphics.FromImage(New Bitmap(1, 1))
            Dim SFormat As New System.Drawing.StringFormat
            Dim rect As New System.Drawing.RectangleF(0, 0, 6000, 6000)
            Dim range() As CharacterRange = New CharacterRange() {New CharacterRange(0, str.Length)}
            SFormat.SetMeasurableCharacterRanges(range)
            Dim regions() As Region = gx.MeasureCharacterRanges(str, tFont, rect, SFormat)
            rect = regions(0).GetBounds(gx)

            If Width0Height1Both2 = 0 Then Return rect.Right + 1 'gx.MeasureString(str, Font, 50000000, lolz).Width
            If Width0Height1Both2 = 1 Then Return rect.Bottom + 1 'gx.MeasureString(str, Font, 50000000, lolz).Height
            If Width0Height1Both2 = 2 Then Return New SizeF(rect.Right + 1, rect.Bottom + 1) 'gx.MeasureString(str, Font, 50000000, lolz)
        End Using
        Return -1
    End Function

    Public Class ScrollArrow
        Public rect As Rectangle, visible As Boolean
        Public X1 As Integer, X2 As Integer, Y1 As Integer, Y2 As Integer, Y3 As Integer
        Public Sub New(ByRef tX As Integer, ByRef tY As Integer, ByRef sz As Integer, ByRef iType As Integer)
            rect = New Rectangle(tX, tY, sz - 1, sz - 1)
            Dim midB As Integer = sz / 2
            Dim lineSZ As Integer = (sz / 4) - 2
            X1 = midB - lineSZ + tX
            X2 = midB + lineSZ + tX
            Y2 = midB + tY
            Y1 = midB - lineSZ + tY
            Y3 = midB + lineSZ + tY
            visible = False
        End Sub
    End Class

    Public Class NewPageTab
        Public rect As Rectangle
        Public L1X1, L1X2, L1Y, L2Y1, L2X, L2Y2 As Integer
        Public Sub New(ByRef x As Integer, ByVal y As Integer, ByVal sz As Integer)
            rect = New Rectangle(x, y, sz - 1, sz - 1)
            Dim midB As Integer = sz / 2
            Dim lineSZ As Integer = ((sz - midB) / 2) - 1
            L1X1 = midB - lineSZ + x
            L1X2 = midB + lineSZ + x
            L1Y = midB + y
            L2X = midB + x
            L2Y1 = midB - lineSZ + y
            L2Y2 = midB + lineSZ + y
        End Sub
    End Class

    Public Class Page
        Inherits System.Windows.Forms.Control
        Implements IDisposable

        Private Shadows disposed As Boolean = False
        Protected Overrides Sub Dispose( _
           ByVal disposing As Boolean)
            If Not Me.disposed Then
                If disposing Then
                    If TB.SelectedPage Is Me Then TB.SelectedPage = Nothing
                    If TB.LastPage Is Me Then TB.LastPage = Nothing
                    Header.Parent = Nothing
                    Header = Nothing
                    TB = Nothing
                    Parent = Nothing
                End If

                ' Free your own state (unmanaged objects).
                ' Set large fields to null.
            End If
            Me.disposed = True
        End Sub

        Public Header As TabHeader
        Public TB As TabBrowser
        Public BorderThickness As Integer = 1, BorderColor As Pen, WWidth As Integer, WHeight As Integer
        Public BGColor As SolidBrush
        Public Shared Count As Integer = 0
        Private DoubleBorder As Integer
        Private bCovered As Boolean = False

        Public Sub SetBoundsTo(ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer)
            WWidth = w - DoubleBorder
            WHeight = h - BorderThickness
            SetBounds(x, y, w, h)
        End Sub

        Public Sub New(ByRef prnt As TabBrowser, ByRef sName As String, Optional ByRef tFont As Font = Nothing, Optional ByRef tFontColor As Color = Nothing, Optional ByRef tBGColor As Color = Nothing)
            Count += 1
            Name = prnt.Name & "Page" & Count
            Parent = prnt.Parent
            TB = prnt
            If tFont Is Nothing Then Font = Parent.Font Else Font = tFont
            If tFontColor = Nothing Then ForeColor = TB.FontColor.Color Else ForeColor = tFontColor
            If tBGColor = Nothing Then BGColor = New SolidBrush(Color.FromKnownColor(KnownColor.Control)) Else BGColor = New SolidBrush(tBGColor)
            BorderColor = TB.BorderColor

            DoubleBorder = BorderThickness * 2
            SetBoundsTo(TB.Left, TB.Bottom, TB.Width, TB.HHeight)
            Header = New TabHeader(Me, sName)
        End Sub

        Private Sub P_ControlAdded(ByVal sender As Control, ByVal e As EventArgs) Handles Me.ControlAdded
            If sender.ClientRectangle = Me.ClientRectangle Then bCovered = True
        End Sub

        Private Sub P_Paint(ByVal sender As Object, ByVal e As PaintEventArgs) Handles Me.Paint
            For i = 0 To BorderThickness - 1
                e.Graphics.DrawLine(BorderColor, i, 0, i, Height - 1)
                e.Graphics.DrawLine(BorderColor, Width - 1 - i, 0, Width - 1 - i, Height - 1)
                e.Graphics.DrawLine(BorderColor, BorderThickness, Height - 1 - i, Width - BorderThickness, Height - 1 - i)
            Next
        End Sub

        Public Class TabHeader
            Public Shared Height As Integer = 20
            Public Parent As Page
            Public Text As String, TextRect As Rectangle
            Public clr As SolidBrush

            Public Sub New(ByRef prnt As Page, ByRef sName As String)
                Parent = prnt
                Text = sName
                clr = New SolidBrush(Parent.ForeColor)
                Dim SS As SizeF = TabBrowser.MeasureAString(2, Text, Parent.Font)
                Dim HBuffer As Integer = 1 + ((Parent.TB.TabHeight - SS.Height) / 2)
                Dim w As Integer = Parent.TB.TabWidth - Parent.TB.TabLBuffer - Parent.TB.ExitButtonXW - 3
                If Parent.TB.CenterTabText = True Then
                Else
                    TextRect = New Rectangle(Parent.TB.TabLBuffer, HBuffer, Math.Min(SS.Width, w), SS.Height)
                End If
                If TextRect.Width = w Then
                    Dim pct As Single = w / SS.Width
                    Dim newLine As String = Text.Substring(0, Math.Min(pct * Text.Length + 1, Text.Length))
                    If newLine.Length < 4 Then
                        newLine = Text.Chars(0) & "."
                        Text = newLine
                    Else
                        For i = 1 To 3
                            If newLine.Length < i Then Exit For
                            Text = newLine.Remove(newLine.Length - i)
                            Text &= Text.PadRight(i, ".")
                        Next
                    End If
                End If
                TextRect.Width += 5
            End Sub
        End Class
    End Class

End Class

Public Class TextField
    Inherits System.Windows.Forms.Control

    Public FontColor As New SolidBrush(Color.Red), BGColor As New SolidBrush(Color.LightBlue), BorderPen As New Pen(Color.DarkBlue)
    Public BorderThickness As Integer, WWidth As Integer, WHeight As Integer
    Public FontSpaceBuffer As Integer
    Public Sub New(ByRef prnt As Control, ByRef rect As Rectangle, ByRef tFont As Font)
        Parent = prnt
        'Dim h As Integer = Math.Max(rect.Height, System.Windows.Forms.SystemInformation.ca
        SetBounds(rect.X, rect.Y, rect.Width, rect.Height)
        SetBorderThickness(1)
    End Sub

    Public Sub SetBorderThickness(ByRef w As Integer)
        BorderThickness = w
        WWidth = Width - (BorderThickness * 2)
        WHeight = Height - (BorderThickness * 2)
    End Sub

    Public Sub TF_Paint(ByVal sender As Object, ByVal e As PaintEventArgs) Handles Me.Paint
        For i = 0 To BorderThickness - 1
            e.Graphics.DrawRectangle(BorderPen, i, i, Width - (i * 2) - 1, Height - (i * 2) - 1)
        Next
        e.Graphics.FillRectangle(BGColor, BorderThickness, BorderThickness, WWidth, WHeight)
    End Sub

    Protected Overrides Sub OnPaintBackground(ByVal pevent As PaintEventArgs)
    End Sub

    Public Class Letter
        Public c As Char, sz As SizeF
        Public Sub New(ByRef character As Char, ByRef size As SizeF)
            c = character
            sz = size
        End Sub
    End Class
End Class
