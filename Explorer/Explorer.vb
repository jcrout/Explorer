Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Reflection
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Threading
Imports System.IO

Public Class frmMain : Inherits BufferedForm
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H2000000
            cp.Style = cp.Style Or &H2000000 And Not 33554432
            Return cp
        End Get
    End Property

    Dim exp As Explorer
    Public Shared OMGLol As Integer = 0
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint, True)
        UpdateStyles()
        FormBorderStyle = FormBorderStyle.Sizable
        Me.MinimumSize = New Size(316, 338)
        Dim w As Integer = 650, h As Integer = 650
        exp = New Explorer(Me, (Screen.PrimaryScreen.WorkingArea.Width / 2) - (w / 2), (Screen.PrimaryScreen.WorkingArea.Height / 2) - (h / 2), w, h)
    End Sub
End Class

Public Class Explorer
#Region "Explorer"

    Public Event ObjectCreated(ByVal sender As Object, ByVal e As EventArgs)
    Public Event TabChanged(ByVal sender As Object, ByVal e As EventArgs)

    Public Structure POINTAPI 'Cursor API
        Public x As Integer
        Public y As Integer
    End Structure
    Public Declare Function GetCursorPos Lib "user32" (ByRef point As POINTAPI) As Boolean

    Private Const SHGFI_ICON As Integer = &H100
    Private Const SHGFI_SMALLICON As Integer = &H1
    Private Const SHGFI_LARGEICON As Integer = &H0

    Public Function GetShellIcon(ByVal path As String) As Bitmap
        Dim info As SHFILEINFO = New SHFILEINFO()
        Dim retval As IntPtr = SHGetFileInfo(path, 0, info, Marshal.SizeOf(info), SHGFI_LARGEICON Or SHGFI_ICON)
        If retval = IntPtr.Zero Then Throw New ApplicationException("Could not retrieve icon")
        Dim cargt() As Type = {GetType(IntPtr)}
        Dim ci As ConstructorInfo = GetType(Icon).GetConstructor(BindingFlags.NonPublic Or BindingFlags.Instance, Nothing, cargt, Nothing)
        Dim cargs() As Object = {info.IconHandle}
        Dim icon As Icon = CType(ci.Invoke(cargs), Icon)
        Return icon.ToBitmap
    End Function

    '' P/Invoke declaration
    Private Structure SHFILEINFO
        Public IconHandle As IntPtr
        Public IconIndex As Integer
        Public Attributes As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public DisplayString As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> _
        Public TypeName As String
    End Structure

    Private Declare Auto Function SHGetFileInfo Lib "Shell32.dll" (ByVal path As String, _
      ByVal attributes As Integer, ByRef info As SHFILEINFO, ByVal infoSize As Integer, ByVal flags As Integer) As IntPtr

    Public Preferences As ExpPreferences
    Public Field As FieldDisplay
    Public AddressBar As TopBar
    Public DisplayBar As BottomBar
    Public Initializing As Boolean = True
    Public TBMain As TabBrowser
    Public Page As String = ""
    Public StartPages As New List(Of String) From {{"C:\"}}
    Public HomePage As String = "C:\"
    Public Parent As Control, FParent As Form
    Public Width As Integer, Height As Integer
    Public ContentFolder As String = ""
    Public ShowSystemFolders As Boolean = False
    Public TB As TextBox, TBHandledFirst As Boolean = False
    Public DCTime As Integer = 0

    Public MyFont As Font, FontColor As Color
    Public MinW As Integer = 100, MinH As Integer = 100
    Public TimerTest As Timer
    Public TimerThreadTest As Thread, TimerInterval As Integer
    Public WWidth As Integer, WHeight As Integer
    Private BorderThickness As Integer = 16
    Private TBESC As Boolean = False

    Public WinRARExtract As ProcessStartInfo
    Public WinRARCompress As ProcessStartInfo
    Public DDMHandledFirst As Boolean = False
    Public CBoard As New ClipBoardX
    Public SF As New SafeFile()
    Public MouseHoverItem As Object
    Private temp As Object
    Private StartingTBMainHeight As Integer = 155
    Private bMinimizing As Boolean = False

#Region "New Explorer"
    Public Sub NewExplorer()
        FParent = Parent.FindForm
        FindContentFolder(Application.StartupPath)
        SetFileGroups()
        LoadPreferences()
        SetWinRARProcessInfo()
        FieldDisplay.SetDDMs()
        DropDownMenu.ContentFolder = ContentFolder
        UpdateSize()
        DCTime = System.Windows.Forms.SystemInformation.DoubleClickTime
        TBMain = New TabBrowser("TBMain", Parent, 0, 100, WWidth, WHeight - StartingTBMainHeight)
        TB = New TextBox
        TB.Parent = Parent
        TB.Font = MyFont
        TB.ForeColor = FontColor
        TB.BorderStyle = BorderStyle.None
        TB.Visible = False
        AddHandler TB.KeyDown, AddressOf TB_KeyDown
        AddHandler TB.LostFocus, AddressOf TB_LostFocus

        AddressBar = New TopBar(Me, Parent, 0, TBMain.Top - 35, WWidth, 35)
        DisplayBar = New BottomBar(Me)

        For i = 0 To StartPages.Count - 1
            NewTab(StartPages(i), False)
        Next

        AddHandler TBMain.AddBoxClicked, AddressOf AddBoxClicked
        AddHandler TBMain.PageChanged, AddressOf PageChanged
        AddHandler Parent.Resize, AddressOf ParentResized

        AddressBar.LoadProfile(Field)

        AddOmniHandlers(DisplayBar)
        AddOmniHandlers(AddressBar)
        AddOmniHandlers(Parent)
        AddOmniHandlers(TBMain)

        SetTestItems()
    End Sub

    Public Sub New(ByRef prnt As Control)
        Parent = prnt
        NewExplorer()
    End Sub

    Public Sub New(ByRef prnt As Control, ByRef l As Integer, ByRef t As Integer, ByRef w As Integer, ByRef h As Integer)
        Parent = prnt
        Parent.SetBounds(l, t, w, h)
        NewExplorer()
    End Sub

    Private Sub LoadPreferences()
        Preferences = New ExpPreferences
        Preferences.SetHighlightColors(Color.MediumSeaGreen)
        MyFont = Preferences.MyFont
        FontColor = Preferences.FontColor
        TabBrowser.HighlightColor = New SolidBrush(Preferences.HighlightHoverColor)
        FieldDisplay.BGColor = New SolidBrush(Preferences.FieldBackcolor)
        FieldDisplay.HighlightColor = New SolidBrush(Preferences.HighlightColor)
        FieldDisplay.HoverHighlightColor = New SolidBrush(Preferences.HighlightHoverColor)
        FieldDisplay.HighlightBorderColor = New Pen(Preferences.HighlightBorderColor)
        TopBar.BGColor = New SolidBrush(Preferences.TopBarBackColor)
        TopBar.AddressBarColor = New SolidBrush(Preferences.AddressBarColor)
        TopBar.HighlightColor = New SolidBrush(Preferences.HighlightHoverColor)
        TopBar.HighlightBorderColor = New Pen(Preferences.HighlightBorderColor)
        BottomBar.BorderColor = New Pen(Preferences.BorderColor)
        BottomBar.BGColor = New SolidBrush(Preferences.BGColor)
        'BottomBar.MyFont = Preferences.MyFont
        DropDownMenu.BGColor = New SolidBrush(Preferences.BGColor)
        DropDownMenu.HighlightColor = New SolidBrush(Preferences.HighlightHoverColor)
        DropDownMenu.HighlightBorderColor = New Pen(Preferences.HighlightBorderColor)
    End Sub

#End Region

#Region "WinRAR"
    Private Sub SetWinRARProcessInfo()
        WinRARExtract = New ProcessStartInfo("C:\Program Files (x86)\WinRAR\WinRAR.exe")
        WinRARCompress = New ProcessStartInfo("C:\Program Files (x86)\WinRAR\WinRAR.exe")
    End Sub

    Public Sub ExtractFiles(ByVal sCompressedFile As String, Optional ByVal sDestination As String = "")
        Dim newThread As New Thread(AddressOf ExtractFilesThread)
        newThread.IsBackground = True
        newThread.Start(sCompressedFile)
    End Sub

    Private Sub ExtractFilesThread(ByVal sArgs As String)
        Dim sCompressedFile As String = ""
        If sArgs.Contains("|") Then
        Else
            sCompressedFile = sArgs
        End If
        Dim sFName As String = sCompressedFile.Substring(0, sCompressedFile.LastIndexOf("\") + 1) ' & sCompressedFile.Substring(sCompressedFile.LastIndexOf("\") + 1, sCompressedFile.LastIndexOf(".") - sCompressedFile.LastIndexOf("\") - 1)
        Dim sName As String = sCompressedFile.Substring(sCompressedFile.LastIndexOf("\") + 1, sCompressedFile.LastIndexOf(".") - sCompressedFile.LastIndexOf("\") - 1)
        If Not System.IO.Directory.Exists(sFName) Then
            My.Computer.FileSystem.CreateDirectory(sFName)
        End If
        WinRARExtract.Arguments = "x -ad " & Chr(34) & sCompressedFile & Chr(34) & " " & Chr(34) & sFName & Chr(34)
        System.Diagnostics.Process.Start(WinRARExtract).WaitForExit()

        Dim di As New System.IO.DirectoryInfo(sFName & sName) 'Detect redundant folders such as C:\Downloads\Downloads\
        Dim arrFo() As System.IO.DirectoryInfo = di.GetDirectories
        If arrFo.Count = 1 AndAlso arrFo(0).Name = di.Name Then
            For Each fo As System.IO.DirectoryInfo In arrFo(0).GetDirectories
                My.Computer.FileSystem.MoveDirectory(fo.FullName, sFName & sName & "\" & fo.Name)
            Next
            For Each fi As System.IO.FileInfo In arrFo(0).GetFiles
                My.Computer.FileSystem.MoveDirectory(fi.FullName, sFName & sName & "\" & fi.Name)
            Next
        End If
        System.Threading.Thread.CurrentThread.Abort()
    End Sub

    Public Sub CompressFile(ByRef sFile As String, ByVal sDestination As String)
        ' For i = 0 To sFiles.Count - 1
        'WinRARCompress.Arguments = "a -rr10 -s c:\backup.rar " & sFiles(i)
        'System.Diagnostics.Process.Start(WinRARCompress)
        'Next
    End Sub

    Public Sub CompressFiles(ByRef sFiles As List(Of String))
        For i = 0 To sFiles.Count - 1
            WinRARCompress.Arguments = "a -rr10 -s c:\backup.rar " & sFiles(i)
            System.Diagnostics.Process.Start(WinRARCompress)
        Next
    End Sub
#End Region

#Region "Font"
    Public Sub SetFont() 'ByRef tFont As Font, Optional ByRef tFontColor As Color = Nothing, Optional ByRef bRefresh As Boolean = False)
        Dim newForm As New Form()
        AddHandler newForm.Load, AddressOf frmFont_Load
        temp = New List(Of Object) From {{MyFont}, {FontColor}}
        newForm.ShowDialog()

        If Not temp Is Nothing AndAlso TypeOf temp Is List(Of Object) Then
            Dim lstObj As List(Of Object) = temp
            Dim tFont As Font = lstObj(0)
            Dim tFontColor As Color = lstObj(1)
            If Not tFontColor = Nothing Then FontColor = tFontColor Else FontColor = Color.Black
            MyFont = tFont
            For i = 0 To TBMain.Pages.Count - 1
                Dim FD As FieldDisplay = TBMain.Pages(i).Controls(0)
                FD.SetFont(MyFont, FontColor, True)
                FD.LoadPage(FD.Page, 3)
            Next
            AddressBar.SetFont(MyFont, FontColor)
            AddressBar.SetURL(Field, AddressBar.URL)
        End If
    End Sub

    Private Sub frmFont_Load(ByVal sender As Form, ByVal e As EventArgs)
        sender.SetBounds(Parent.Left, Parent.Top, 200, 200)
        Dim lblFont As New Label
        lblFont.Parent = sender
        lblFont.AutoSize = True
        lblFont.Location = New Point(5, 5)
        lblFont.Text = "Font Family: "
        Dim cbFontFamily As New ComboBox()
        cbFontFamily.Name = "cbFontFamily"
        cbFontFamily.Parent = sender
        cbFontFamily.SetBounds(lblFont.Right + 5, 5, 100, 20)
        Dim index As Integer = 0, count As Integer = 0
        For Each ff As FontFamily In System.Drawing.FontFamily.Families
            cbFontFamily.Items.Add(ff.Name)
            If MyFont.FontFamily.Name = ff.Name Then index = count
            count += 1
        Next
        cbFontFamily.SelectedIndex = index

        Dim lblFontSize As New Label()
        lblFontSize.Parent = sender
        lblFontSize.AutoSize = True
        lblFontSize.Text = "Size: "
        lblFontSize.Left = 5
        lblFontSize.Top = cbFontFamily.Bottom + 5
        Dim txtFontSize As New TextBox()
        txtFontSize.Name = "txtFontSize"
        txtFontSize.Parent = sender
        txtFontSize.SetBounds(cbFontFamily.Left, cbFontFamily.Bottom + 5, cbFontFamily.Width, cbFontFamily.Height)
        txtFontSize.Text = MyFont.Size

        Dim lblColor As New Label
        lblColor.Parent = sender
        lblColor.AutoSize = True
        lblColor.Text = "Color: "
        lblColor.Left = 5
        lblColor.Top = txtFontSize.Bottom + 5

        Dim btnColor As New Button
        btnColor.Parent = sender
        btnColor.Name = "btnColor"
        btnColor.Text = "Color"
        btnColor.BackColor = Color.White
        btnColor.ForeColor = FontColor
        btnColor.SetBounds(cbFontFamily.Left, txtFontSize.Bottom + 5, cbFontFamily.Width, cbFontFamily.Height)

        Dim btnCancel As New Button()
        btnCancel.Parent = sender
        btnCancel.Text = "Cancel"
        btnCancel.SetBounds(5, btnColor.Bottom + 25, 85, 22)

        Dim btnUpdate As New Button()
        btnUpdate.Parent = sender
        btnUpdate.Text = "Update"
        btnUpdate.SetBounds(btnColor.Right - 85, btnColor.Bottom + 25, 85, 22)

        sender.AcceptButton = btnUpdate
        sender.CancelButton = btnCancel

        AddHandler btnColor.Click, AddressOf frmFont_btnColor_Click
        AddHandler btnUpdate.Click, AddressOf frmFont_UpdateCancel_Click
        AddHandler btnCancel.Click, AddressOf frmFont_UpdateCancel_Click
    End Sub

    Private Sub frmFont_UpdateCancel_Click(ByVal sender As Button, ByVal e As EventArgs)
        Dim frmFont As Form = sender.FindForm
        If sender.Text = "Cancel" Then
            frmFont.Close()
        Else
            Dim cb As ComboBox = frmFont.Controls("cbFontFamily")
            Dim tb As TextBox = frmFont.Controls("txtFontSize")
            Dim btn As Button = frmFont.Controls("btnColor")
            Dim lstObj As List(Of Object) = temp
            Try
                lstObj.Item(0) = New Font(cb.SelectedItem.ToString, tb.Text)
                lstObj.Item(1) = btn.ForeColor
                frmFont.Close()
            Catch ex As Exception
                MsgBox("Error setting new font: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub frmFont_btnColor_Click(ByVal sender As Button, ByVal e As EventArgs)
        Dim c As Color = frmFont_SetFontColor()
        Dim btnColor As Button = sender.Parent.Controls("btnColor")
        btnColor.ForeColor = c
    End Sub

    Private Function frmFont_SetFontColor() As Color
        Dim CD As New ColorDialog()
        CD.ShowDialog()
        Return CD.Color
    End Function
#End Region

#Region "Testing"
    Public LBLLOL As New Label, LBLAB As New Label, LBLDDM As New Label, LBLFRM As New Label

    Private Sub SetTestItems()
        LBLLOL.Parent = Parent
        LBLLOL.Left = 1
        LBLLOL.Top = 1
        LBLLOL.Width = 75
        LBLLOL.Height = 30
        LBLLOL.BorderStyle = BorderStyle.FixedSingle
        LBLAB = New Label
        LBLAB.Parent = Parent
        LBLAB.Left = 80
        LBLAB.Top = 1
        LBLAB.Width = 75
        LBLAB.Height = 30
        LBLAB.BorderStyle = BorderStyle.FixedSingle
        LBLDDM = New Label
        LBLDDM.Parent = Parent
        LBLDDM.Left = 160
        LBLDDM.Top = 1
        LBLDDM.Width = 75
        LBLDDM.Height = 30
        LBLDDM.BorderStyle = BorderStyle.FixedSingle
        LBLFRM = New Label
        LBLFRM.Parent = Parent
        LBLFRM.Left = 240
        LBLFRM.Top = 1
        LBLFRM.Width = 75
        LBLFRM.Height = 30
        LBLFRM.BorderStyle = BorderStyle.FixedSingle
        TimerInterval = 1000
        TimerThreadTest = New Thread(AddressOf ThreadTimerTick)
        TimerThreadTest.IsBackground = True
        TimerThreadTest.Start()
    End Sub

    Private Sub ThreadTimerTick()
        Do
            System.Threading.Thread.Sleep(TimerInterval)
            TTick()
        Loop
    End Sub

    Private Sub TTick()
        If Not Field Is Nothing Then
            If Not LBLLOL Is Nothing Then
                If LBLLOL.InvokeRequired Then
                    Try
                        LBLLOL.Invoke(New Action(AddressOf TTick))
                    Catch ex As Exception
                    End Try
                Else
                    LBLLOL.Text = "FD: " & FieldDisplay.OMGlol
                    LBLAB.Text = "TB: " & TopBar.OMGlol
                    LBLDDM.Text = "DDM: " & DropDownMenu.OMGLol
                    ' LBLFRM.Text = "Form: " & frmMainOld.OMGLol
                End If
            End If
        End If
    End Sub
#End Region

    Private Sub UpdateSize()
        If TypeOf Parent Is Form Then
            Dim lol As Form = Parent
            If lol.FormBorderStyle = FormBorderStyle.Sizable Then
                WWidth = Parent.Width - 16
            End If
            If lol.MaximizeBox Or lol.MinimizeBox Or lol.Text <> "" Then
                WHeight = Parent.Height - 38
            End If
        Else
            WWidth = Parent.Width
            WHeight = Parent.Height
        End If
        Height = WHeight
        Width = WWidth
    End Sub

    Public Sub ResizeFields()
        TBMain.SetBoundsTo(0, TBMain.Top, WWidth, WHeight - StartingTBMainHeight)
        For i = 0 To TBMain.Pages.Count - 1
            Dim FD As FieldDisplay = TBMain.Pages(i).Controls(0)
            FD.SetBounds(TBMain.Pages(i).BorderThickness, 0, TBMain.Pages(i).WWidth, TBMain.Pages(i).WHeight)
        Next
        If AddressBar.Width <> TBMain.Width Then AddressBar.Width = TBMain.Width
    End Sub

    Private Sub ParentResized(ByVal sender As Object, ByVal e As EventArgs)
        Dim LP As TabBrowser.Page = TBMain.LastPage
        Dim larg As Rectangle = LP.Bounds
        If Parent.FindForm.WindowState = FormWindowState.Minimized Then
            bMinimizing = True
            Exit Sub
        Else
            If bMinimizing = True Then
                bMinimizing = False
                Exit Sub
            End If
        End If
        UpdateSize()
        ResizeFields()
        DisplayBar.UpdatePosition()
    End Sub

    Public Sub RemoveHandlers(ByRef sender As Control)
        RemoveHandler sender.MouseUp, AddressOf OmniMouseUp
        RemoveHandler sender.MouseDown, AddressOf OmniMouseDown
        RemoveHandler sender.MouseEnter, AddressOf OmniMouseEnter
        RemoveHandler sender.MouseWheel, AddressOf OmniMouseWheel
        RemoveHandler sender.LostFocus, AddressOf OmniLostFocus
        If TypeOf sender Is FieldDisplay Then
            Dim omg As FieldDisplay = sender
            If Field Is omg Then Field = Nothing
            'RemoveHandler omg.SelectedItemChanged, DisplayBar.FieldItemChanged
            'RemoveHandler omg.DirectoryChanged, AddressOf DisplayBar.DirectoryChanged
        End If
        If TB.Parent Is sender Then TB.Parent = Nothing
        If MouseHoverItem Is sender Then MouseHoverItem = Nothing
    End Sub

    Private Sub AddOmniHandlers(ByRef sender As Control)
        AddHandler sender.MouseUp, AddressOf OmniMouseUp
        AddHandler sender.MouseDown, AddressOf OmniMouseDown
        AddHandler sender.MouseEnter, AddressOf OmniMouseEnter
        AddHandler sender.MouseWheel, AddressOf OmniMouseWheel
        AddHandler sender.LostFocus, AddressOf OmniLostFocus
    End Sub

    Private Sub OmniMouseWheel(ByVal sender As Object, ByVal e As MouseEventArgs)
        If MouseHoverItem Is Nothing Then Exit Sub
        If TypeOf MouseHoverItem Is FieldDisplay Then
            Dim F As FieldDisplay = MouseHoverItem
            If F.Scrollbars(0) Is Nothing Then Exit Sub
            If e.Delta < 0 Then
                F.Scrollbars(0).ScrollInDirection(0)
            Else
                F.Scrollbars(0).ScrollInDirection(1)
            End If
        End If
    End Sub

    Private Sub OmniMouseEnter(ByVal sender As Object, ByVal e As EventArgs)
        MouseHoverItem = sender
    End Sub

    Private Sub OmniMouseUp(ByVal sender As Object, ByVal e As MouseEventArgs)
        Try
            If e.Button = Windows.Forms.MouseButtons.XButton1 AndAlso Not AddressBar.History.PNode Is Nothing Then
                Field.LoadPage(AddressBar.History.PNode.URL, -1)
            End If
            If e.Button = Windows.Forms.MouseButtons.XButton2 Then If Not AddressBar.History.NNode Is Nothing Then Field.LoadPage(AddressBar.History.NNode.URL, 1)
        Catch ex As Exception
            MsgBox("Sub OmniMouseUp: " & ex.Message)
        End Try
    End Sub

    Private Sub OmniMouseDown(ByVal sender As Object, ByVal e As MouseEventArgs)
        If TBHandledFirst = True Or DDMHandledFirst = True Then
            TBHandledFirst = False
            DDMHandledFirst = False
            Exit Sub
        End If
        Try
            If Not DropDownMenu.ActiveDDM Is Nothing Then DropDownMenu.ActiveDDM.Dismiss()

            If TB.Visible = True Then TB.Visible = False
        Catch ex As Exception
            MsgBox("Sub OmniMouseDown: " & ex.Message)
        End Try
        TBHandledFirst = False
        DDMHandledFirst = False
    End Sub

    Private Sub OmniLostFocus(ByVal sender As Object, ByVal e As EventArgs)
        If Not Field Is Nothing Then
            Field.bSelecting = False
            Field.bDrag = False
            If Not Field.MouseRect = Nothing AndAlso Field.MouseRect.Width > 3 Or Field.MouseRect.Height > 3 Then
                Field.MouseRect = Nothing
                Field.Refresh()
            End If
        End If
    End Sub

    Private Sub PageChanged(ByRef tbSender As TabBrowser, ByVal e As EventArgs)
        If Not DropDownMenu.ActiveDDM Is Nothing AndAlso DropDownMenu.ActiveDDM.Visible = True Then DropDownMenu.ActiveDDM.Dismiss()
        Field = TBMain.SelectedPage.Controls(0)
        AddressBar.LoadProfile(Field)
        Field.RefreshPage()
        AddressBar.Refresh()
        DisplayBar.UpdateItems()
        RaiseEvent TabChanged(tbSender, e)
    End Sub

    Private Sub FindContentFolder(ByRef sPath As String)
        Dim di As New System.IO.DirectoryInfo(sPath)
        Dim arrFo() As System.IO.DirectoryInfo = di.GetDirectories
        For Each fo As System.IO.DirectoryInfo In arrFo
            If ContentFolder <> "" Then Exit Sub
            If fo.Name.ToUpper = "CONTENT" Then
                ContentFolder = fo.FullName & "\"
                LoadIcons()
                Exit Sub
            End If
        Next
        If ContentFolder = "" Then FindContentFolder(sPath.Substring(0, sPath.LastIndexOf("\")))
    End Sub

    Private Sub LoadIcons()
        If Not FieldDisplay.Item.ItemIcons Is Nothing Then Exit Sub
        Dim di As New System.IO.DirectoryInfo(ContentFolder)
        Dim arrFi() As System.IO.FileInfo = di.GetFiles
        FieldDisplay.Item.ItemIcons = New List(Of StringAndObject)
        For Each fi As System.IO.FileInfo In arrFi
            If fi.Name.StartsWith("Ico") Then
                Dim bmp As Bitmap = Image.FromFile(fi.FullName)
                bmp.MakeTransparent(Color.FromArgb(255, 255, 0, 128))
                Dim lolz As String = fi.Name.Substring(3, fi.Name.LastIndexOf(".") - 3)
                If lolz = "Error" Then FieldDisplay.Item.ErrorIcon = FieldDisplay.Item.ItemIcons.Count
                If lolz <> "Folder" Then lolz = lolz.Insert(0, ".").ToLower
                FieldDisplay.Item.ItemIcons.Add(New StringAndObject(lolz, bmp))
            End If
        Next
    End Sub

    Public Sub AddBoxClicked(ByRef tbSender As TabBrowser, ByVal e As EventArgs)
        NewTab(HomePage)
    End Sub

    Public Sub DDMCreated(ByVal ddmSender As DropDownMenu, ByVal e As EventArgs)
        RaiseEvent ObjectCreated(ddmSender, e)
    End Sub

    Public Sub NewTab(ByRef sPage As String, Optional ByRef bChangeTab As Boolean = True)
        Dim PN As String = ""
        Page = sPage
        If Page.Last <> "\" Then
            PN = Page.Substring(Page.LastIndexOf("\") + 1)
            Page &= "\"
        Else
            Dim yorg As String = Page.Remove(Page.Length - 1)
            PN = yorg.Substring(yorg.LastIndexOf("\") + 1)
        End If
        TBMain.AddPage(PN, MyFont, FontColor, Preferences.FieldBackcolor, True, bChangeTab)
        If Not DropDownMenu.ActiveDDM Is Nothing AndAlso DropDownMenu.ActiveDDM.Visible = True Then DropDownMenu.ActiveDDM.Dismiss()
        Dim newTab As New FieldDisplay(Me, TBMain.LastPage, Page)
        AddHandler newTab.SelectedItemChanged, AddressOf DisplayBar.FieldItemChanged
        AddHandler newTab.DirectoryChanged, AddressOf DisplayBar.DirectoryChanged
        If bChangeTab = True Or Field Is Nothing Then
            Field = newTab
            DisplayBar.UpdateItems()
        End If
        AddOmniHandlers(TBMain.LastPage)
        AddOmniHandlers(newTab)
        AddHandler newTab.DDMCreated, AddressOf DDMCreated
        RaiseEvent ObjectCreated(newTab, New EventArgs)
    End Sub

    Private Sub SetTB(ByRef prnt As Control, ByRef l As Integer, ByRef t As Integer, ByRef w As Integer, ByRef txt As String, Optional ByRef tFont As Font = Nothing, Optional ByRef tFontColor As Color = Nothing)
        TB.Parent = prnt
        TB.Left = l
        TB.Top = t
        TB.Width = w
        If Not tFont Is Nothing Then TB.Font = tFont
        If tFontColor.A <> 0 Then TB.ForeColor = tFontColor
        TB.Text = txt
        TB.BringToFront()
        TB.Visible = True
        TB.Focus()
        TB.SelectionStart = TB.Text.Length
    End Sub

    Private Sub TB_KeyDown(ByVal sender As TextBox, ByVal e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            If TypeOf TB.Parent Is TopBar Then
                If TB.Text.Length = 0 Then Exit Sub
                Dim URL As String = TB.Text
                Try
                    If System.IO.Directory.Exists(URL) Then
                        Field.LoadPage(URL)
                    ElseIf System.IO.File.Exists(URL) Then
                        Try
                            System.Diagnostics.Process.Start(URL)
                        Catch ex As Exception

                        End Try
                    Else
                        MsgBox("Invalid address.", MsgBoxStyle.OkOnly, "omg")
                    End If
                Catch ex As Exception
                    MsgBox("Invalid address.", MsgBoxStyle.OkOnly, "omg")
                End Try
            ElseIf TypeOf TB.Parent Is FieldDisplay Then

            End If
            TB.Visible = False
        ElseIf e.KeyCode = Keys.Escape Then
            TBESC = True
            TB.Visible = False
        End If
    End Sub

    Private Sub TB_LostFocus(ByVal sender As TextBox, ByVal e As EventArgs)
        If TBESC = True Then
            TBESC = False
            Exit Sub
        End If
        If TypeOf sender.Parent Is FieldDisplay Then
            Dim FD As FieldDisplay = sender.Parent
            If FD.TextEventType = 0 Then
                If TB.Text <> FD.TextItem.Text Then
                    Try
                        If FD.TextItem.Extension = "Folder" Then
                            My.Computer.FileSystem.RenameDirectory(FD.TextItem.FullPath, TB.Text)
                            FD.LoadPage(FD.Page)
                        Else
                            'My.Computer.FileSystem.RenameFile
                        End If

                    Catch ex As Exception

                    End Try
                End If
            End If
        End If
        TB.Visible = False
    End Sub

    Protected Shared Sub GetFormCoords(ByRef obj As Control, ByRef pt As Point)
        Dim obj2 As Control = obj
        Do
            If obj2.Parent Is Nothing Then Exit Sub
            pt.X += obj2.Left
            pt.Y += obj2.Top
            If TypeOf obj2.Parent Is EmptyForm Then
                Exit Sub
            Else
                obj2 = obj2.Parent
            End If
        Loop
    End Sub

    Public Shared Sub HighlightArea(ByRef gx As Graphics, ByRef rect As Rectangle, ByRef hColor As SolidBrush, ByRef bColor As Pen, Optional ByRef mode As Integer = 0)
        gx.DrawRectangle(bColor, rect)
        If mode = 0 Then gx.FillRectangle(hColor, rect.Left + 1, rect.Top + 1, rect.Width - 1, rect.Height - 1) Else gx.FillRectangle(hColor, rect.Left + 2, rect.Top + 2, rect.Width - 3, rect.Height - 3)
    End Sub

    Public Class ExpPreferences
        Public OpenInNewTabsOnly As Boolean = True
        Public MyFont As Font = New Font("Comic Sans MS", 10)
        Public FontColor As Color = Color.Black
        Public FieldBackcolor As Color = Color.White
        Public TopBarBackColor As Color = Color.White
        Public AddressBarColor As Color = Color.White
        Public MenuBackColor As Color = Color.MediumPurple
        Public HighlightColor As Color
        Public HighlightHoverColor As Color
        Public HighlightBorderColor As Color
        Public BorderColor As Color = Color.Blue
        Public BGColor As Color = Color.White
        Public HighlighAlpha, HoverHighlightAlpha As Integer

        Public Sub SetHighlightColors(ByRef clr As Color, Optional ByRef tHighlightAlpha As Integer = 150, Optional ByRef tHoverHighlightAlpha As Integer = 100)
            HighlighAlpha = tHighlightAlpha
            HoverHighlightAlpha = tHoverHighlightAlpha
            HighlightColor = Color.FromArgb(tHighlightAlpha, clr)
            HighlightHoverColor = Color.FromArgb(tHoverHighlightAlpha, clr.R, clr.G, clr.B)
            HighlightBorderColor = clr
        End Sub

        Public Sub SavePreferences()

        End Sub
    End Class

    Public Function CapitalizeEveryWord(ByRef sLine As String)
        sLine = sLine.Trim
        sLine = sLine.Insert(0, sLine.Substring(0, 1).ToUpper)
        sLine = sLine.Remove(1, 1)
        Dim pIndex As Integer = 0
        Do
            Dim index As Integer = sLine.IndexOf(" ", pIndex)
            If index = -1 Then Return sLine
            If sLine.Chars(index + 1) = " " Then
                pIndex += 1
                Continue Do
            End If
            sLine = sLine.Insert(index + 1, sLine.Substring(index + 1, 1).ToUpper)
            sLine = sLine.Remove(index + 2, 1)
            pIndex = index + 1
        Loop
        Return sLine
    End Function

#End Region

#Region "File Groups"
    Public Shared MasterFileGroups As List(Of FileGroup), GenericFiles As FileGroup

    Public Sub SetFileGroups()
        FileGroup.PropertyList = New Hashtable(287)
        Dim strFileName As String = ContentFolder & "IcoError.png"
        Dim objShell As Object = CreateObject("Shell.Application")
        Dim objFolder As Object = objShell.Namespace(Path.GetDirectoryName(strFileName))
        Dim lastindex As Integer = 0
        Dim arrlines As New List(Of String), sOMG As String = ""
        For i = 0 To 286
            Dim omg As String = CapitalizeEveryWord(objFolder.GetDetailsOf(strFileName, i).ToString)
            FileGroup.PropertyList.Add(i, omg)
            arrlines.Add(i & ", " & omg & vbNewLine)
        Next
        ' arrlines.Reverse()
        ' For i = 0 To arrlines.Count - 1
        '     sOMG &= arrlines(i)
        ' Next
        '  Clipboard.SetText(sOMG)

        Dim props As List(Of Integer) = Enumerable.Range(6, 280).ToList
        GenericFiles = New FileGroup("Generic Files", props, New List(Of FileType) From {{New FileType("Generic", "Generic")}})

        MasterFileGroups = New List(Of FileGroup)
        MasterFileGroups.Add(New FileGroup("Document", New List(Of Integer), New List(Of FileType) From {{New FileType(".txt", "ASCII Text")}, {New FileType(".rtf", "Rich Text Format")}}))
        MasterFileGroups.Add(New FileGroup("Programs", New List(Of Integer) From {{25}, {271}}, New List(Of FileType) From {{New FileType(".exe", "PC Application")}}))
        MasterFileGroups.Add(New FileGroup("Videos", New List(Of Integer) From {{27}, {28}, {283}, {284}, {285}, {286}}, New List(Of FileType) From {{New FileType(".wmv", "Windows Media Video")}, {New FileType(".mp4", "ISO compliant MPEG-4 streams with AAC audio")}, {New FileType(".avi", "Audio-Video Interleave")}, {New FileType(".mpg", "Motion Pictures Experts Group", New List(Of String) From {{".mpeg"}})}, {New FileType(".mkv", "Matroska")}, {New FileType(".flv", "Flash Video")}, {New FileType(".mov", "Apple Quicktime")}, {New FileType(".rm", "Real Media")}, {New FileType(".ogm", "OGG Media")}}))
        MasterFileGroups.Add(New FileGroup("Images", New List(Of Integer) From {{31}, {160}, {161}, {162}, {163}, {164}}, New List(Of FileType) From {{New FileType(".jpg", "Joint Photographic Experts Group", New List(Of String) From {{".jpeg"}})}, {New FileType(".png", "Public Network Graphic")}, {New FileType(".bmp", "Windows BitMap")}, {New FileType(".gif", "Graphics Interchange Format")}}))
        MasterFileGroups.Add(New FileGroup("Music", New List(Of Integer) From {{13}, {14}, {15}, {16}, {19}, {20}, {21}, {26}, {27}, {28}}, New List(Of FileType) From {{New FileType(".mp3", "MPEG-1 or MPEG-2 Audio Layer II")}, {New FileType(".wav", "Microsoft Wave")}, {New FileType(".m4a", "Apple Lossless")}}))

        MasterFileGroups(2).ShowOrder = New List(Of Integer) From {{0}, {2}, {1}, {27}, {283}, {285}, {28}, {286}, {284}, {4}, {5}}
    End Sub

    Public Function GetFileType(ByRef sItem As String) As FileType
        Dim sExtension As String = Path.GetExtension(sItem)
        For Each FG As FileGroup In MasterFileGroups
            Dim FT As FileType = FG.FileTypes.Item(sExtension)
            If FT Is Nothing Then Continue For
            Return FT
        Next
        Return GenericFiles.FileTypes("Generic")
    End Function

    Public Overloads Function GetFileInfo(ByRef sItem As String, ByRef sFileType As FileType) As List(Of TwoStrings)
        Return sFileType.GetAttributes(sItem, sFileType.Parent.Properties, 1)
    End Function

    Public Overloads Function GetFileInfo(ByRef sItem As String) As List(Of TwoStrings)
        Dim sExtension As String = Path.GetExtension(sItem)
        For Each FG As FileGroup In MasterFileGroups
            Dim FT As FileType = FG.FileTypes.Item(sExtension)
            If FT Is Nothing Then Continue For
            Return FT.GetAttributes(sItem, FG.Properties, 1) 'All numbers: Enumerable.Range(0, 286).ToList
        Next
        Dim GT As FileType = GenericFiles.FileTypes.Item("Generic")
        Return GT.GetAttributes(sItem, GenericFiles.Properties, 1)
        Return Nothing
    End Function

    Public Class FileGroup
        Public Name As String
        Public Properties As New List(Of Integer)
        Public ShowOrder As New List(Of Integer)
        Public FileTypes As New Hashtable()
        Public Shared PropertyList As Hashtable

        Public Sub New(ByRef sGroupName As String, ByRef sAttributes As List(Of Integer), Optional ByRef arrFileTypes As List(Of FileType) = Nothing, Optional ByRef bDefaultShowOrder As Boolean = True)
            Name = sGroupName
            Properties.AddRange(Enumerable.Range(0, 6))
            If bDefaultShowOrder = True Then ShowOrder = New List(Of Integer) From {{0}, {2}, {1}, {3}, {4}, {5}}
            If bDefaultShowOrder = True Then
                For i = 0 To sAttributes.Count - 1
                    Properties.Add(sAttributes(i))
                    ShowOrder.Add(sAttributes(i))
                Next
            Else
                For i = 0 To sAttributes.Count - 1
                    Properties.Add(sAttributes(i))
                Next
            End If
            If Not arrFileTypes Is Nothing Then
                For i = 0 To arrFileTypes.Count - 1
                    FileTypes.Add(arrFileTypes(i).Extension, New FileType(arrFileTypes(i).Extension, arrFileTypes(i).Name, arrFileTypes(i).Extensions, Me))
                Next
            End If
        End Sub

        Public Sub AddFileType(ByRef sExtension As String, ByRef sName As String, Optional ByRef sAlternateExtensions As List(Of String) = Nothing)
            FileTypes.Add(sExtension, New FileType(sName, sExtension, sAlternateExtensions, Me))
        End Sub

        Public Sub AddFileTypes(ByRef sFileTypes As List(Of FileType))
            For i = 0 To sFileTypes.Count - 1
                FileTypes.Add(sFileTypes(i).Extension, sFileTypes(i))
            Next
        End Sub

    End Class

    Public Class FileType
        Public Parent As FileGroup
        Public Name As String
        Public Extension As String, Extensions As New List(Of String)

        Public Sub New(ByRef sExtension As String, ByRef sName As String, Optional ByRef sAlternateExtensions As List(Of String) = Nothing, Optional ByRef prnt As FileGroup = Nothing)
            Parent = prnt
            Name = sName
            Extension = sExtension
            Extensions.Add(Extension)
            If Not sAlternateExtensions Is Nothing Then
                For i = 0 To sAlternateExtensions.Count - 1
                    Extensions.Add(sAlternateExtensions(i))
                Next
            End If
        End Sub

        Public Function OrganizeToShowOrder(ByRef sList As Hashtable) As List(Of TwoStrings)
            Dim OList As New List(Of TwoStrings)
            For i = 0 To Parent.ShowOrder.Count - 1
                Dim index() As String = sList(Parent.ShowOrder(i))
                OList.Add(New TwoStrings(index(0), index(1)))
            Next
            Return OList
        End Function

        Public Function GetAttributes(ByVal ItemFullPath As String, ByRef AIndexes As List(Of Integer), ByRef SearchType As Byte) As List(Of TwoStrings)
            Dim sList As New Hashtable(Parent.Properties.Count - 1)
            Dim sTwoStrings(1) As String
            Try
                Dim fi As New FileInfo(ItemFullPath)
                Dim shl As Shell32.Shell = New Shell32.Shell
                Dim dir As Shell32.Folder = shl.[NameSpace](fi.DirectoryName)
                Dim itm As Shell32.FolderItem = dir.Items().Item(fi.Name)
                Dim itm2 As Shell32.ShellFolderItem = DirectCast(itm, Shell32.ShellFolderItem)
                If SearchType = 0 Then
                    For i = 0 To AIndexes.Count - 1
                        ReDim sTwoStrings(1)
                        sTwoStrings(0) = FileGroup.PropertyList(AIndexes(i))
                        sTwoStrings(1) = dir.GetDetailsOf(itm2, AIndexes(i)).ToString
                        sList.Add(AIndexes(i), sTwoStrings)
                    Next
                Else
                    For i = 0 To AIndexes.Count - 1
                        ReDim sTwoStrings(1)
                        sTwoStrings(1) = dir.GetDetailsOf(itm2, AIndexes(i)).ToString
                        sTwoStrings(0) = FileGroup.PropertyList(AIndexes(i))
                        If sTwoStrings(1) = "" Then Continue For
                        sList.Add(AIndexes(i), sTwoStrings)
                    Next
                End If
                Return OrganizeToShowOrder(sList)
            Catch ex As Exception
                MsgBox("Error: Could not read fileinfo of file " & Path.GetFileName(ItemFullPath))
                Return Nothing
            End Try
        End Function
    End Class
#End Region








    Public Class TopBar
        Inherits System.Windows.Forms.Control

        Public Shared AddressBarColor As SolidBrush, BGColor As SolidBrush
        Public AddressSegments As List(Of SegmentBox), AddressBarBorderThickness As Integer = 1
        Public AddressSegmentArrows As List(Of SegmentArrow)
        Public BorderThickness As Integer = 1, BorderColor As New Pen(Color.Black)
        Public FontColor As SolidBrush
        Public AddressBarBorderColor As New Pen(Color.Black)
        Public History As PageNode, HistoryList As DropDownMenu
        Public HistoryArrow As DownArrow
        Public URLBox As Rectangle, AddressBarInnerRect As Rectangle
        Public Arrows(1) As Arrow, ArrowSize As Integer = 26, SegmentArrowW As Integer = 5, SegmentArrowH As Integer = 5
        Public URL As String
        Public Exp As Explorer
        Public TopNode As PageNode, NodeCount As Integer
        Public Shared HighlightColor As SolidBrush, HighlightBorderColor As Pen, HighlightMode As Integer = 1
        Public Shared AddressBarRightBuffer As Integer = 5, AddressBarTop As Integer = 7
        Private HighlightedItem As Object, Initializing As Boolean = True
        Public Shared FieldProfiles As New Hashtable, FProfile As FieldProfile
        Public Shared OMGlol As Integer = 0 'Paint event counter

        Public Class FieldProfile
            Public Field As FieldDisplay
            Public TB As TopBar
            Public URL As String
            Public TopNode As PageNode, NodeCount As Integer
            Public History As PageNode
            Public AddressSegments As List(Of SegmentBox)
            Public AddressSegmentArrows As List(Of SegmentArrow)

            Public Sub New(ByRef tParent As TopBar, ByRef tField As FieldDisplay)
                TB = tParent
                Field = tField
            End Sub

            Public Sub SetProfileItems(ByVal tURL As String, ByRef tHistory As PageNode, ByVal tNodeCount As Integer, ByVal tAS As List(Of TopBar.SegmentBox), ByVal tAA As List(Of TopBar.SegmentArrow))
                URL = tURL
                History = tHistory.Clone
                Dim obj As PageNode = History
                Do
                    If obj.NNode Is Nothing Then Exit Do
                    obj = obj.NNode
                Loop
                TopNode = obj
                NodeCount = tNodeCount
                AddressSegments = tAS
                AddressSegmentArrows = tAA
            End Sub

            Public Sub Remove()
                FieldProfiles.Remove(Field.ID)
                Me.Field = Nothing
                Me.TB = Nothing
                Me.URL = Nothing
                Me.History.Dispose()
                Me.History = Nothing
                Me.TopNode = Nothing
                Me.AddressSegmentArrows = Nothing
                Me.AddressSegments = Nothing
            End Sub
        End Class

        Public Sub LoadProfile(ByRef tField As FieldDisplay)
            FProfile = FieldProfiles(tField.ID)
            If FProfile Is Nothing Then Exit Sub
            URL = FProfile.URL
            History = FProfile.History
            TopNode = FProfile.TopNode
            NodeCount = FProfile.NodeCount
            AddressSegments = FProfile.AddressSegments
            AddressSegmentArrows = FProfile.AddressSegmentArrows
            UpdateHistoryList()
        End Sub

        Public Sub CreateProfile(ByRef tField As FieldDisplay)
            If Not FieldProfiles.Contains(tField.ID) Then
                FProfile = New FieldProfile(Me, tField)
                FProfile.SetProfileItems(URL, History, NodeCount, AddressSegments, AddressSegmentArrows)
                FieldProfiles.Add(tField.ID, FProfile)
            Else
                FProfile = FieldProfiles(tField.ID)
                FProfile.SetProfileItems(URL, History, NodeCount, AddressSegments, AddressSegmentArrows)
            End If
        End Sub

        Public Sub New(ByRef tExp As Explorer, ByRef prnt As Control, ByRef x As Integer, ByRef y As Integer, ByRef w As Integer, ByRef h As Integer)
            Exp = tExp
            Parent = prnt
            SetBounds(x, y, w, h)
            Arrow.Width = 26
            Arrow.Height = 26
            SetArrowImages()
            SetFont(Exp.MyFont, Exp.FontColor)

            Arrows(0) = New Arrow(8, ((Height / 2) - (Arrow.Height / 2)))
            Arrows(1) = New Arrow(Arrows(0).rect.Right + 5, Arrows(0).rect.Y)
            HistoryArrow = New DownArrow(Arrows(1).rect.Right, BorderThickness, DownArrow.Img.Width + 10, (Height - (BorderThickness * 2)))  '(Arrows(1).rect.Right + 5, Arrows(1).rect.Y + ((Arrows(1).rect.Height / 2) - (SegmentArrowH / 2)), Nothing, 1)
            SetAddressBarSize(New Rectangle(HistoryArrow.rect.Right, AddressBarTop, Width - HistoryArrow.rect.Right - 5 - 50, Height - (AddressBarTop * 2)))
            AddHandler Me.Paint, AddressOf TB_Paint
            AddHandler Me.MouseUp, AddressOf TB_MouseUp
            AddHandler Me.MouseDown, AddressOf TB_MouseDown
            AddHandler Me.MouseMove, AddressOf TB_MouseMove
            AddHandler Me.MouseLeave, AddressOf TB_MouseLeave
            AddHandler Me.Resize, AddressOf TB_Resize
        End Sub

        Public Sub SetAddressBarSize(ByRef rect As Rectangle)
            URLBox = rect
            AddressBarInnerRect = New Rectangle(URLBox.X + 1, URLBox.Y + 1, URLBox.Width - 1, URLBox.Height - 1)
            SegmentBox.height = AddressBarInnerRect.Height - 1
        End Sub

        Public Sub SetFont(ByRef tFont As Font, Optional ByRef clr As Color = Nothing)
            Font = tFont
            SegmentBox.MyFont = Font
            If clr = Nothing Then FontColor = New SolidBrush(Color.Black) Else FontColor = New SolidBrush(clr)
        End Sub

        Private Sub TB_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
            OMGlol += 1
            e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
            For i = 0 To BorderThickness - 1
                e.Graphics.DrawRectangle(BorderColor, i, i, Width - 1 - (i * 2), Height - 1 - (i * 2))
            Next

            For i = 0 To 1
                If Arrows(i).Enabled Then e.Graphics.DrawImage(Arrow.Images(i), Arrows(i).rect.X, Arrows(i).rect.Y) Else e.Graphics.DrawImage(Arrow.DisabledImages(i), Arrows(i).rect.X, Arrows(i).rect.Y)
            Next
            If NodeCount < 2 Then e.Graphics.DrawImage(DownArrow.InactiveImg, HistoryArrow.DX, HistoryArrow.DY) Else e.Graphics.DrawImage(DownArrow.Img, HistoryArrow.DX, HistoryArrow.DY)
            e.Graphics.DrawRectangle(AddressBarBorderColor, URLBox)
            e.Graphics.FillRectangle(AddressBarColor, AddressBarInnerRect)
            If URL <> "" Then
                For i = 0 To AddressSegments.Count - 1
                    If AddressSegments(i).Visible = False Then Continue For
                    e.Graphics.DrawString(AddressSegments(i).Text, SegmentBox.MyFont, FontColor, AddressSegments(i).rect.X, AddressSegments(i).rect.Y)
                    e.Graphics.DrawImage(SegmentArrow.Img, AddressSegmentArrows(i).rect.X + SegmentArrow.LBuffer, AddressSegmentArrows(i).rect.Y + SegmentArrow.TBuffer)
                Next
            End If
            If Not HighlightedItem Is Nothing Then
                If TypeOf HighlightedItem Is DownArrow Then Explorer.HighlightArea(e.Graphics, HighlightedItem.hrect, HighlightColor, HighlightBorderColor, HighlightMode) Else Explorer.HighlightArea(e.Graphics, HighlightedItem.rect, HighlightColor, HighlightBorderColor, HighlightMode)
            End If
        End Sub

        Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
            e.Graphics.FillRectangle(BGColor, BorderThickness, BorderThickness, Width - (BorderThickness * 2), Height - (BorderThickness * 2))
        End Sub

        Public Sub TB_Resize(ByVal sender As Object, ByVal e As EventArgs)
            SetAddressBarSize(New Rectangle(HistoryArrow.rect.Right, AddressBarTop, Width - HistoryArrow.rect.Right - 5 - 50, Height - (AddressBarTop * 2)))
            SetURL(FProfile.Field, URL, 2)
        End Sub

        Private Sub TB_MouseMove(ByVal tbSender As TopBar, ByVal e As MouseEventArgs)
            If HistoryArrow.rect.Contains(e.X, e.Y) Then
                If TopNode.PNode Is Nothing Then Exit Sub
                If HighlightedItem Is HistoryArrow Then Exit Sub
                HighlightedItem = HistoryArrow
                Me.Refresh()
                Exit Sub
            End If

            For i = 0 To AddressSegments.Count - 1
                If AddressSegmentArrows(i).rect.Contains(e.X, e.Y) Then
                    If AddressSegmentArrows(i) Is HighlightedItem Then Exit Sub
                    HighlightedItem = AddressSegmentArrows(i)
                    Me.Refresh()
                    Exit Sub
                ElseIf AddressSegments(i).rect.Contains(e.X, e.Y) Then
                    If AddressSegments(i) Is HighlightedItem Then Exit Sub
                    HighlightedItem = AddressSegments(i)
                    Me.Refresh()
                    Exit Sub
                End If
            Next
            If Not HighlightedItem Is Nothing Then
                HighlightedItem = Nothing
                Me.Refresh()
            End If
        End Sub

        Private Sub TB_MouseUp(ByVal tbSender As TopBar, ByVal e As MouseEventArgs)
            If e.Button <> Windows.Forms.MouseButtons.Left Then Exit Sub
            Dim MPT As POINTAPI
            GetCursorPos(MPT)
            Dim MPoint As New Point(e.X, e.Y)
            For i = 0 To 1
                If Arrows(i).rect.Contains(MPoint) Then
                    If Arrows(i).Enabled = True Then
                        If i = 0 Then
                            Exp.Field.LoadPage(History.PNode.URL, -1)
                        Else
                            Exp.Field.LoadPage(History.NNode.URL, 1)
                        End If
                    End If
                    Exit Sub
                End If
            Next
            For i = AddressSegments.Count - 1 To 0 Step -1
                If AddressSegments(i).rect.Contains(MPoint) Then
                    If Exp.Field.Page <> AddressSegments(i).FullPath Then Exp.Field.LoadPage(AddressSegments(i).FullPath)
                    Exit Sub
                ElseIf AddressSegmentArrows(i).rect.Contains(MPoint) Then
                    'Dim di As New System.IO.DirectoryInfo
                    MsgBox("Arrow " & i)
                    Exit Sub
                End If
            Next
            If URLBox.Contains(MPoint) And e.X > AddressSegmentArrows(AddressSegmentArrows.Count - 1).rect.Right Then
                Exp.SetTB(Me, URLBox.Left + 6, AddressSegments(0).rect.Y + 1, URLBox.Width - 6, URL, SegmentBox.MyFont, FontColor.Color)
            End If
        End Sub

        Private Sub TB_MouseDown(ByVal tbSender As TopBar, ByVal e As MouseEventArgs)
            Dim MPoint As New Point(e.X, e.Y)
            If HistoryArrow.rect.Contains(MPoint) Then
                If NodeCount < 2 Then Exit Sub
                If HistoryList.Visible = True Then
                    HistoryList.Dismiss()
                    Exit Sub
                End If
                Exp.DDMHandledFirst = True
                Dim FormCoords As New Point(e.X, e.Y)
                GetFormCoords(Me, FormCoords)
                HistoryList.ShowF(FormCoords.X, FormCoords.Y)
                Exit Sub
            End If
        End Sub

        Private Sub TB_MouseLeave(ByVal tbSender As TopBar, ByVal e As EventArgs)
            If Not HighlightedItem Is Nothing Then
                HighlightedItem = Nothing
                Refresh()
            End If
        End Sub

        Public Sub SetURL(ByRef tField As FieldDisplay, ByRef sPage As String, Optional ByRef direction As Integer = 0)
            Try
                If sPage.Last <> "\" Then sPage &= "\"
                URL = sPage
                Dim sLine As String = URL
                Dim PIndex As Integer = 0
                AddressSegments = New List(Of SegmentBox)
                AddressSegmentArrows = New List(Of SegmentArrow)
                Dim modifier As Integer = 0
                Do
                    Dim index As Integer = URL.IndexOf("\", PIndex)
                    If index = -1 Then Exit Do
                    Dim sName As String = URL.Substring(0, index)
                    If AddressSegments.Count = 0 Then
                        modifier = MeasureAString(1, sName, SegmentBox.MyFont)
                        modifier = URLBox.Y + ((URLBox.Height / 2) - (modifier / 2))
                        AddressSegments.Add(New SegmentBox(URLBox.X + 5, modifier, sName))
                    Else
                        AddressSegments.Add(New SegmentBox(AddressSegmentArrows(AddressSegmentArrows.Count - 1).rect.Right, modifier, sName))
                    End If
                    AddressSegmentArrows.Add(New SegmentArrow(AddressSegments(AddressSegments.Count - 1).rect.Right, AddressBarTop + AddressBarBorderThickness, AddressSegments(AddressSegments.Count - 1)))
                    PIndex = index + 1
                    If index = -1 Then Exit Do
                Loop
                If AddressSegmentArrows(AddressSegmentArrows.Count - 1).rect.Right > (URLBox.Right - AddressBarRightBuffer) Then
                    Dim Shift As Integer = 0
                    Dim dif As Integer = AddressSegmentArrows(AddressSegmentArrows.Count - 1).rect.Right - URLBox.Right - AddressBarRightBuffer
                    For i = 0 To AddressSegments.Count - 1
                        If AddressSegments(i).rect.Left - dif < URLBox.Left Then
                            AddressSegments(i).Visible = False
                        Else
                            dif = AddressSegments(i).rect.X - URLBox.X - 5
                            AddressSegments(i).rect.X -= dif
                            AddressSegmentArrows(i).rect.X -= dif
                            If (AddressSegmentArrows(AddressSegmentArrows.Count - 1).rect.Right - dif) > (URLBox.Right - AddressBarRightBuffer) Then
                                dif += (AddressSegmentArrows(i).rect.Right + 2 - URLBox.X)
                                AddressSegments(i).Visible = False
                            End If

                            For i2 = i + 1 To AddressSegments.Count - 1
                                AddressSegments(i2).rect.X -= dif
                                AddressSegmentArrows(i2).rect.X -= dif
                            Next
                            Exit For
                        End If
                    Next
                End If
                If Not History Is Nothing AndAlso Not FieldProfiles.Contains(tField.ID) Then
                    History = Nothing
                    NodeCount = 0
                    Arrows(0).Enabled = False
                    Arrows(1).Enabled = False
                End If
                If direction <> 2 Then
                    If Not History Is Nothing Then
                        If direction = -1 Then History = History.PNode
                        If direction = 1 Then History = History.NNode
                        If direction = 0 Then
                            If Not History.NNode Is Nothing Then History.NNode.Dispose()
                            History = New PageNode(Me, History, URL)
                        End If
                    Else
                        History = New PageNode(Me, Nothing, URL)
                    End If
                End If
                UpdateHistoryList()
                CreateProfile(tField)
                If Initializing = True Then
                    Initializing = False
                Else
                    Me.Refresh()
                End If
            Catch ex As Exception
                MsgBox("TopBar Sub SetURL: " & ex.Message)
            End Try
        End Sub

        Public Sub UpdateHistoryList()
            If NodeCount < 2 Then Exit Sub
            Dim obj As PageNode = TopNode
            HistoryList = New DropDownMenu(Exp.FParent, obj.Text, False, obj)
            If History Is obj Then HistoryList.Items(0).Checked = True
            Dim SimilarIndex As Integer = -1, bFound As Boolean = False
            Do
                If obj.PNode Is Nothing Then Exit Do
                If Not System.IO.Directory.Exists(obj.PNode.URL) Then
                    MsgBox("Directory no longer exists.", MsgBoxStyle.OkOnly, "Error")
                    Dim lol As PageNode = obj.PNode
                    If Not lol.PNode Is Nothing Then
                        obj.PNode = lol.PNode
                        lol.PNode.NNode = obj
                    Else
                        obj.PNode = Nothing
                        lol = Nothing
                    End If
                End If
                obj = obj.PNode
                HistoryList.AddItem(obj.Text, False, obj)
                If History Is obj Then
                    HistoryList.Items(HistoryList.Items.Count - 1).Checked = True
                    bFound = True
                ElseIf History.URL = obj.URL Then
                    SimilarIndex = HistoryList.Items.Count - 1
                End If
            Loop
            If bFound = False AndAlso SimilarIndex <> -1 Then
                HistoryList.Items(SimilarIndex).Checked = True
            End If
            If NodeCount > 1 Then
                If Not History.PNode Is Nothing Then Arrows(0).Enabled = True Else Arrows(0).Enabled = False
                If Not History.NNode Is Nothing Then Arrows(1).Enabled = True Else Arrows(1).Enabled = False
            End If
            HistoryList.Visible = False
            AddHandler HistoryList.MenuItemClicked, AddressOf DDM_Click
        End Sub

        Private Sub DDM_Click(ByVal ddmSender As DropDownMenu, ByVal e As DropDownMenu.MenuItemClickedArgs)
            If e.Item.Checked = True Then Exit Sub
            History = e.Item.Tag
            HighlightedItem = Nothing
            If History.URL = URL Then
                UpdateHistoryList()
                Me.Refresh()
            Else
                Exp.Field.LoadPage(History.URL, 2)
            End If
        End Sub

        Public Sub SetArrowImages()
            If Not SegmentArrow.Img Is Nothing Then Exit Sub
            Dim bmp As New Bitmap(ArrowSize, ArrowSize)
            Dim ArrowPen As New Pen(Arrow.ArrowColor)
            Dim mid As Integer = ArrowSize / 2
            Using gx As Graphics = Graphics.FromImage(bmp)
                gx.FillEllipse(Arrow.BackColor, 0, 0, ArrowSize, ArrowSize)
                gx.DrawLine(ArrowPen, 5, mid, mid, mid - 7)
                gx.DrawLine(ArrowPen, 5, mid, mid, mid + 7)
                gx.DrawLine(ArrowPen, mid, mid - 6, mid, mid - 2)
                gx.DrawLine(ArrowPen, mid, mid + 6, mid, mid + 2)
                gx.DrawLine(ArrowPen, mid, mid - 2, mid + 7, mid - 2)
                gx.DrawLine(ArrowPen, mid, mid + 2, mid + 7, mid + 2)
                gx.DrawLine(ArrowPen, mid + 7, mid - 2, mid + 7, mid + 2)
            End Using
            FillArea(bmp, mid, mid, ArrowPen.Color, Arrow.BackColor.Color)
            Arrow.Images(0) = bmp.Clone
            bmp.RotateFlip(RotateFlipType.RotateNoneFlipX)
            Arrow.Images(1) = bmp.Clone

            For i = 0 To 1
                Dim dbmp As Bitmap = Arrow.Images(i).Clone
                For r = 0 To dbmp.Width - 1
                    For c = 0 To dbmp.Height - 1
                        If dbmp.GetPixel(r, c) = Arrow.BackColor.Color Then dbmp.SetPixel(r, c, Color.FromArgb(45, Arrow.BackColor.Color))
                    Next
                Next
                Arrow.DisabledImages(i) = dbmp
            Next

            ArrowPen = New Pen(SegmentArrow.ArrowColor)
            bmp = New Bitmap(SegmentArrowW, SegmentArrowH)
            mid = SegmentArrowH / 2
            Using gx As Graphics = Graphics.FromImage(bmp)
                gx.DrawLine(ArrowPen, 0, 0, SegmentArrowW - 1, mid)
                gx.DrawLine(ArrowPen, 0, SegmentArrowH - 1, SegmentArrowW - 1, mid)
                gx.DrawLine(ArrowPen, 0, 0, 0, SegmentArrowH)
            End Using
            FillArea(bmp, 1, mid, ArrowPen.Color, Color.FromArgb(0, 0, 0, 0))
            SegmentArrow.Img = bmp.Clone
            SegmentArrow.Width = bmp.Width + SegmentArrow.LBuffer + SegmentArrow.RBuffer
            SegmentArrow.Height = Height - (AddressBarBorderThickness * 2) - (AddressBarTop * 2)
            SegmentArrow.TBuffer = (SegmentArrow.Height / 2) - (SegmentArrowH / 2)

            bmp = New Bitmap(9, 6)
            Using gx As Graphics = Graphics.FromImage(bmp)
                mid = Math.Floor(bmp.Width / 2)
                gx.DrawLine(ArrowPen, 0, 0, bmp.Width - 1, 0)
                gx.DrawLine(ArrowPen, 0, 0, mid, bmp.Height - 1)
                gx.DrawLine(ArrowPen, bmp.Width - 1, 0, mid, bmp.Height - 1)
            End Using
            FillArea(bmp, bmp.Width / 2, 2, ArrowPen.Color, Color.FromArgb(0, 0, 0, 0))
            DownArrow.Img = bmp.Clone
            For r = 0 To bmp.Width - 1
                For c = 0 To bmp.Height - 1
                    If bmp.GetPixel(r, c) <> Color.FromArgb(0, 0, 0, 0) Then bmp.SetPixel(r, c, Color.FromArgb(50, ArrowPen.Color))
                Next
            Next
            DownArrow.InactiveImg = bmp
        End Sub

        Public Class PageNode
            Implements ICloneable

            Public PNode As PageNode, Text As String, URL As String, NNode As PageNode
            Private TB As TopBar

            Public Sub New(ByRef TopBar As TopBar, ByRef PreviousNode As PageNode, ByVal PageName As String, Optional ByRef NextNode As PageNode = Nothing)
                TB = TopBar
                URL = PageName
                PNode = PreviousNode
                NNode = NextNode
                If Not PNode Is Nothing Then
                    PNode.NNode = Me
                End If
                If URL.Last <> "\" Then
                    Text = URL.Substring(URL.LastIndexOf("\" + 1))
                Else
                    Dim line As String = URL.Substring(0, URL.Length - 1)
                    Text = line.Substring(line.LastIndexOf("\") + 1)
                End If
                TB.TopNode = Me
                TB.NodeCount += 1
            End Sub

            Public Sub Dispose()
                TB.NodeCount -= 1
                TB = Nothing
                PNode = Nothing
                URL = ""
                Text = ""
                If Not NNode Is Nothing Then
                    NNode.Dispose()
                    NNode = Nothing
                End If
                Me.Finalize()
            End Sub

            Public Function Clone() As Object Implements System.ICloneable.Clone
                Return Me.MemberwiseClone()
            End Function
        End Class

        Public Class SegmentBox
            Public rect As Rectangle
            Public Text As String
            Public FullPath As String
            Public Visible As Boolean = True
            Public Shared MyFont As Font, height As Integer

            Public Sub New(ByRef x As Integer, ByRef y As Integer, ByRef sPath As String)
                FullPath = sPath
                If FullPath.Last <> "\" Then
                    Text = FullPath.Substring(FullPath.LastIndexOf("\") + 1)
                    FullPath &= "\"
                Else
                    Dim orgo As String = FullPath.Substring(0, FullPath.Length - 1)
                    FullPath = orgo.Substring(orgo.LastIndexOf("\") + 1)
                End If
                Dim sz As SizeF = MeasureAString(2, Text, MyFont)
                rect = New Rectangle(x, y, sz.Width, height)
            End Sub
        End Class

        Public Class DownArrow
            Public Shared Img As Bitmap, InactiveImg As Bitmap
            Public rect As Rectangle, hrect As Rectangle, DX As Integer, DY As Integer
            Public Sub New(ByRef x As Integer, ByVal y As Integer, ByRef w As Integer, ByRef h As Integer)
                rect = New Rectangle(x, y, w, h)
                DX = x + (w - Img.Width) / 2
                DY = (h - Img.Height) / 2
                hrect = New Rectangle(x, DY - 6, w - 1, Img.Height + 10)
            End Sub
        End Class

        Public Class SegmentArrow
            Public Shared LBuffer As Integer = 7, RBuffer As Integer = 5, TBuffer As Integer
            Public Shared Width As Integer, Height As Integer, Img As Bitmap, ArrowColor As Color = Color.Black
            Public SBox As SegmentBox
            Public rect As Rectangle, Visible As Boolean = True, DDM As DropDownMenu
            Public Sub New(ByRef tX As Integer, ByRef tY As Integer, ByVal tSegmentBox As SegmentBox, Optional ByRef tType As Byte = 0)
                SBox = tSegmentBox
                If tType = 0 Then rect = New Rectangle(tX, tY, Width, Height) Else rect = New Rectangle(tX, tY, Width, Height)
            End Sub
        End Class

        Public Class Arrow
            Public Shared Width As Integer, Height As Integer, Images(1) As Bitmap, DisabledImages(1) As Bitmap, BackColor As New SolidBrush(Color.FromArgb(255, 0, 0, 255)), ArrowColor As Color = Color.White
            Public rect As Rectangle, Enabled As Boolean = False
            Public Sub New(ByRef tX As Integer, ByRef tY As Integer)
                rect = New Rectangle(tX, tY, Width, Height)
            End Sub
        End Class
    End Class


    Public Class FieldDisplay 'Display the folder items
        Inherits System.Windows.Forms.Control

        Implements IDisposable
        Private Shadows disposed As Boolean = False
        Protected Overrides Sub Dispose( _
           ByVal disposing As Boolean)
            If Not Me.disposed Then
                If disposing Then
                    NextID = ID
                    If ActiveField Is Me Then ActiveField = Nothing
                    Dim FProf As TopBar.FieldProfile = TopBar.FieldProfiles(ID)
                    Dim yango As Hashtable = TopBar.FieldProfiles
                    FProf.Remove()
                    FProf = Nothing
                    'If Not TopBar.FProfile Is Nothing AndAlso TopBar.FProfile.Field.ID = ID Then TopBar.FProfile = Nothing
                    Exp.RemoveHandlers(Me)
                    Exp = Nothing
                    Files = Nothing
                    Folders = Nothing
                    Parent = Nothing
                    PParent = Nothing
                    For i = 0 To Items.Count - 1
                        Items = Nothing
                    Next
                    SelectedItems = Nothing
                    OldSelectedItems = Nothing
                    OldShiftItems = Nothing
                    HoverItem = Nothing
                    Columns = Nothing
                    di = Nothing
                    If Not Scrollbars(0) Is Nothing Then Scrollbars(0).Dispose()
                    If Not Scrollbars(1) Is Nothing Then Scrollbars(1).Dispose()

                    If Not DCTimer Is Nothing Then
                        DCTimer.Abort()
                        If Not DCTimer Is Nothing Then DCTimer = Nothing
                    End If
                    If Not RenameTimer Is Nothing Then
                        RenameTimer.Abort()
                        If Not RenameTimer Is Nothing Then RenameTimer = Nothing
                    End If
                    If Not Delegates(0) Is Nothing Then
                        For i = Delegates(0).Count - 1 To 0 Step -1
                            RemoveHandler SelectedItemChanged, Delegates(0).Item(i)
                        Next
                        Delegates(0) = Nothing
                    End If
                    If Not Delegates(1) Is Nothing Then
                        For i = Delegates(1).Count - 1 To 0 Step -1
                            RemoveHandler PageLoaded, Delegates(1).Item(i)
                        Next
                        Delegates(1) = Nothing
                    End If
                    If Not Delegates(2) Is Nothing Then
                        For i = Delegates(2).Count - 1 To 0 Step -1
                            RemoveHandler DirectoryChanged, Delegates(2).Item(i)
                        Next
                        Delegates(2) = Nothing
                    End If
                End If
                ' Free your own state (unmanaged objects).
                ' Set large fields to null.
            End If
            Me.disposed = True
        End Sub

        Private Delegates(3) As ArrayList

        Public Custom Event SelectedItemChanged As EventHandler
            AddHandler(ByVal value As EventHandler)
                Me.Events.AddHandler("SelectedItemChangedEvent", value)
                If Delegates(0) Is Nothing Then Delegates(0) = New ArrayList
                Delegates(0).Add(value)
            End AddHandler
            RemoveHandler(ByVal value As EventHandler)
                Me.Events.RemoveHandler("SelectedItemChangedEvent", value)
                Delegates(0).Remove(value)
            End RemoveHandler
            RaiseEvent(ByVal sender As Object, ByVal e As EventArgs)
                If Not Delegates(0) Is Nothing Then CType(Me.Events("SelectedItemChangedEvent"), EventHandler).Invoke(sender, e)
            End RaiseEvent
        End Event

        Public Custom Event PageLoaded As EventHandler
            AddHandler(ByVal value As EventHandler)
                Me.Events.AddHandler("PageLoadedEvent", value)
                If Delegates(1) Is Nothing Then Delegates(1) = New ArrayList
                Delegates(1).Add(value)
            End AddHandler
            RemoveHandler(ByVal value As EventHandler)
                Me.Events.RemoveHandler("PageLoadedEvent", value)
                Delegates(1).Remove(value)
            End RemoveHandler
            RaiseEvent(ByVal sender As Object, ByVal e As System.EventArgs)
                If Not Delegates(1) Is Nothing Then CType(Me.Events("PageLoadedEvent"), EventHandler).Invoke(sender, e)
            End RaiseEvent
        End Event

        Public Custom Event DirectoryChanged As EventHandler
            AddHandler(ByVal value As EventHandler)
                Me.Events.AddHandler("DirectoryChangedEvent", value)
                If Delegates(2) Is Nothing Then Delegates(2) = New ArrayList
                Delegates(2).Add(value)
            End AddHandler
            RemoveHandler(ByVal value As EventHandler)
                Me.Events.RemoveHandler("DirectoryChangedEvent", value)
                Delegates(2).Remove(value)
            End RemoveHandler
            RaiseEvent(ByVal sender As Object, ByVal e As System.EventArgs)
                If Not Delegates(2) Is Nothing Then CType(Me.Events("DirectoryChangedEvent"), EventHandler).Invoke(sender, e)
            End RaiseEvent
        End Event

        Public Custom Event DDMCreated As EventHandler
            AddHandler(ByVal value As EventHandler)
                Me.Events.AddHandler("DDMCreatedEvent", value)
                If Delegates(3) Is Nothing Then Delegates(3) = New ArrayList
                Delegates(3).Add(value)
            End AddHandler
            RemoveHandler(ByVal value As EventHandler)
                Me.Events.RemoveHandler("DDMCreatedEvent", value)
                Delegates(3).Remove(value)
            End RemoveHandler
            RaiseEvent(ByVal sender As DropDownMenu, ByVal e As System.EventArgs)
                If Not Delegates(3) Is Nothing Then CType(Me.Events("DDMCreatedEvent"), EventHandler).Invoke(sender, e)
            End RaiseEvent
        End Event

        Public Exp As Explorer
        Public PParent As TabBrowser.Page
        Public WLeft As Integer, WTop As Integer, WWidth As Integer, WHeight As Integer
        Public BorderThickness As Integer = 0, BorderColor As New Pen(Color.Black)
        Public Page As String = ""
        Public PageName As String = ""
        Public di As System.IO.DirectoryInfo
        Public Folders() As System.IO.DirectoryInfo
        Public Files() As System.IO.FileInfo
        Public Items As List(Of Item), SelectedItems As New List(Of Item)
        Public Columns As New List(Of Point), StartColumn As Integer, ColumnHeight As Integer, ColumnWidth As Integer, MaxColumnWidth As Integer = 510, ColumnHSpacing As Integer = 15, ColumnVSpacing As Integer = 3
        Public bScrollHorizontal As Boolean = False
        Public TBuffer As Integer = 10, LBuffer As Integer = 10
        Public Scrollbars(1) As ScrollBar
        Public TT As ToolTip, HoverItem As Item, bDrag As Boolean = False, bSelecting As Boolean = False
        Public FontColor As SolidBrush
        Public TextHeight As Integer, TextTop As Integer
        Public TextItem As Item, TextEventType As Byte = 0
        Public MouseRect As Rectangle
        Public ID As Integer = 0
        Public Shared IcoW As Integer = 16, IcoH As Integer = 16
        Public Shared HighlightColor As SolidBrush, HoverHighlightColor As SolidBrush, HighlightBorderColor As Pen
        Public Shared DCInterval As Integer = 0, BGColor As SolidBrush
        Public Shared HighlightMode As Integer = 1
        Public Shared ActiveField As FieldDisplay
        Public Shared Count As Integer = 0, NextID As Integer = -1
        Public Shared DDMFileEvents As New List(Of String), DDMWinRarEvents As New List(Of String), DDMPaste As New List(Of String)

        Public Shared OMGlol As Integer = 0

        Private MouseDownX As Integer = -1, MouseDownY As Integer = -1, bDoubleClicked As Boolean = False
        Private MPT As Point
        Private DCTimer As Thread, RenameTimer As Thread
        Private DCX As Integer = -1, DCY As Integer = -1
        Private imgHeight As Integer = 0, OldSelectedItems As List(Of Item)
        Private ShiftStart As Integer = -1, OldShiftItems As New List(Of Item)
        Private Initializing As Boolean = True
        Private Shared CopyType As Byte = 0, CopyStart As FieldDisplay

        Public Sub New(ByRef E As Explorer, ByRef prnt As Control, Optional ByRef sPage As String = "")
            ID = Count
            Count += 1
            Exp = E
            Parent = prnt
            Name = "Field"
            If TypeOf prnt Is TabBrowser.Page Then
                PParent = prnt
                SetBounds(PParent.BorderThickness, 0, PParent.WWidth, PParent.WHeight)
            Else
                SetBounds(0, 0, prnt.Width, prnt.Height)
            End If

            Scrollbars(0) = Nothing
            Scrollbars(1) = Nothing
            DCInterval = Exp.DCTime

            WLeft = BorderThickness
            WTop = BorderThickness
            WWidth = Width - (BorderThickness * 2)
            WHeight = Height - (BorderThickness * 2)
            SetFont(Exp.MyFont, Exp.FontColor, False)

            TT = New ToolTip()
            If sPage <> "" Then LoadPage(sPage)
            ActiveField = Me
            AddHandler Me.Resize, AddressOf F_Resizing
        End Sub

        Public Sub F_Paint(ByVal sender As Object, ByVal e As PaintEventArgs) Handles Me.Paint
            If Initializing = False AndAlso Not Exp.Field Is Me Then Exit Sub
            Using bmp As New Bitmap(Me.ClientSize.Width, Me.ClientSize.Height)
                Using gfx As Graphics = Graphics.FromImage(bmp)
                    For i = 0 To BorderThickness - 1
                        gfx.DrawRectangle(BorderColor, i, i, Width - (i * 2) - 1, Height - (i * 2) - 1)
                    Next
                    gfx.FillRectangle(BGColor, BorderThickness, BorderThickness, Width - (BorderThickness * 2), Height - (BorderThickness * 2))
                    OMGlol += 1

                    'Handle click-drag highlighting
                    If MouseDownX <> -1 Then
                        Explorer.HighlightArea(gfx, MouseRect, HighlightColor, HighlightBorderColor, HighlightMode)
                    End If

                    'Draw folder contents
                    gfx.TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
                    MPT = GetItemRanges(0, Width)
                    For i = MPT.X To MPT.Y
                        gfx.DrawImage(Items(i).GetIcon, Items(i).rect.X, Items(i).rect.Y + imgHeight)
                        gfx.DrawString(Items(i).Text, Font, FontColor, Items(i).TextRect)
                    Next

                    For i = 0 To SelectedItems.Count - 1
                        Explorer.HighlightArea(gfx, SelectedItems(i).rect, HighlightColor, HighlightBorderColor, HighlightMode)
                    Next
                    If Not HoverItem Is Nothing Then If HoverItem.Highlighted = False Then Explorer.HighlightArea(gfx, HoverItem.rect, HoverHighlightColor, HighlightBorderColor, HighlightMode)
                End Using
                e.Graphics.DrawImage(bmp, 0, 0)
            End Using

        End Sub

        Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)

        End Sub

        Private Sub F_Resizing(ByVal sender As Object, ByVal e As EventArgs)
            WWidth = Width - (BorderThickness * 2)
            RefreshPage()
        End Sub

        Private Sub F_MouseMove(ByVal fdSender As Object, ByVal e As MouseEventArgs) Handles Me.MouseMove
            If StartColumn = -1 Or Columns.Count = 0 Then Exit Sub
            If e.X < 0 Or e.Y < 0 Then Exit Sub
            If bSelecting = True Then
                If Not HoverItem Is Nothing Then HoverItem = Nothing
                Dim LX As Integer = Math.Min(MouseDownX, e.X), LY As Integer = Math.Min(MouseDownY, e.Y)
                MouseRect = New Rectangle(LX, LY, Math.Max(MouseDownX, e.X) - LX, Math.Max(MouseDownY, e.Y) - LY)
            End If
            If MouseRect <> Nothing Then MPT = GetItemRanges(MouseRect.Left, MouseRect.Right) Else MPT = GetItemRanges(e.X, e.X)

            Dim count As Integer = SelectedItems.Count
            If bSelecting = True Then
                DeselectItems(False)
                For i = MPT.X To MPT.Y
                    If Items(i).rect.IntersectsWith(MouseRect) Then 'Select this item, it is in the mouse select-box
                        SelectItem(Items(i))
                    End If
                Next
                If ModifierKeys = Keys.Control Then
                    For i = 0 To OldSelectedItems.Count - 1
                        If OldSelectedItems(i).rect.IntersectsWith(MouseRect) Then
                            DeselectItem(OldSelectedItems(i))
                        Else
                            If Not SelectedItems.Contains(OldSelectedItems(i)) Then
                                OldSelectedItems(i).Highlighted = True
                                SelectItem(OldSelectedItems(i))
                            End If
                        End If
                    Next
                End If
                If SelectedItems.Count <> count Then
                    RaiseEvent SelectedItemChanged(Me, New EventArgs)
                End If
                Refresh()
                Exit Sub
            ElseIf bDrag = True Then
                Exit Sub
            Else
                For i = MPT.X To MPT.Y
                    If Items(i).rect.Contains(e.X, e.Y) Then
                        If Items(i).Highlighted = True Then
                            If Not HoverItem Is Nothing Then
                                HoverItem = Nothing
                                Refresh()
                            End If
                            Exit Sub
                        End If
                        If HoverItem Is Items(i) Then Exit Sub
                        If TextItem Is Items(i) Then
                            If Not HoverItem Is Nothing Then
                                HoverItem = Nothing
                                Refresh()
                            End If
                            Exit Sub
                        End If
                        HoverItem = Items(i)
                        Refresh()
                        Exit Sub
                    End If
                Next
            End If
            If Not HoverItem Is Nothing Then
                HoverItem = Nothing
                Refresh()
            End If
        End Sub

        Private Sub F_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles Me.MouseLeave
            If Not HoverItem Is Nothing Then
                HoverItem = Nothing
                Refresh()
            End If
        End Sub

        Private Sub RenameTimerThread(ByVal tItem As Item)
            System.Threading.Thread.Sleep(900)
            Me.Invoke(New Action(Of Item)(AddressOf RenameItemEvent), tItem)
            System.Threading.Thread.CurrentThread.Abort()
        End Sub

        Private Sub F_MouseDown(ByVal fdSender As FieldDisplay, ByVal e As MouseEventArgs) Handles Me.MouseDown
            Me.Focus()
            If Not RenameTimer Is Nothing AndAlso RenameTimer.ThreadState.ToString.Contains("Background") Then
                Dim omglolomg As String = RenameTimer.ThreadState.ToString
                RenameTimer.Abort()
            End If
            If e.Button <> Windows.Forms.MouseButtons.Left AndAlso e.Button <> Windows.Forms.MouseButtons.Right Then Exit Sub
            If e.Button = Windows.Forms.MouseButtons.Left Then
                If DCTimer Is Nothing Then
                    DCTimer = New Thread(AddressOf DCTimer_Tick)
                    DCTimer.IsBackground = True
                    DCTimer.Start()
                Else
                    If SelectedItems.Count = 1 AndAlso SelectedItems(0).rect.Contains(e.X, e.Y) Then bDoubleClicked = True
                    TimerReset()
                End If
            End If
            If Not HoverItem Is Nothing Then HoverItem = Nothing 'Deselect hover item, check if should be reselected later
            MouseDownX = e.X
            MouseDownY = e.Y

            MPT = GetItemRanges(e.X, e.X)
            Dim count As Integer = SelectedItems.Count
            Dim sMod As String = ModifierKeys.ToString
            For i = MPT.X To MPT.Y
                If Items(i).rect.Contains(e.X, e.Y) Then
                    If e.Button = Windows.Forms.MouseButtons.Left Then
                        bDrag = True
                        If ShiftStart = -1 Then ShiftStart = i
                    End If
                    If sMod = "None" Then
                        If Not SelectedItems.Contains(Items(i)) Then 'New item selected, deselect other items
                            HoverItem = Nothing
                            DeselectItems(False)
                            SelectItem(Items(i))
                            ShiftStart = i
                            RaiseEvent SelectedItemChanged(Me, New EventArgs)
                            Refresh()
                            Exit Sub
                        Else 'Same item selected again
                            If e.Button = Windows.Forms.MouseButtons.Left Then
                                If bDoubleClicked = True Then Exit Sub
                                If SelectedItems.Count > 1 Then
                                    DeselectItems(False)
                                    SelectItem(Items(i))
                                    ShiftStart = i
                                    RaiseEvent SelectedItemChanged(Me, New EventArgs)
                                    Me.Refresh()
                                Else
                                    RenameTimer = New Thread(AddressOf RenameTimerThread)
                                    RenameTimer.IsBackground = True
                                    RenameTimer.Start(Items(i))
                                End If
                                Exit Sub
                            End If
                        End If
                    Else
                        HoverItem = Nothing
                        If sMod.Contains("Control") And Not sMod.Contains("Shift") Then
                            ShiftStart = i
                        End If
                        If sMod.Contains("Shift") Then
                            If sMod.Contains("Control") Then
                                HighlightShift(i, False)
                            Else
                                HighlightShift(i, True)
                            End If
                        Else
                            If Items(i).Highlighted = True Then
                                DeselectItem(Items(i))
                            Else
                                SelectItem(Items(i))
                            End If
                        End If
                        If count <> SelectedItems.Count Then
                            If SelectedItems.Count <> count Then RaiseEvent SelectedItemChanged(Me, New EventArgs)
                            Refresh()
                        End If
                    End If
                    Exit Sub
                End If
            Next
            OldSelectedItems = New List(Of Item)
            If Not sMod.Contains("Control") AndAlso Not sMod.Contains("Shift") Then
                DeselectItems(count <> SelectedItems.Count)
            Else
                For i = 0 To SelectedItems.Count - 1
                    OldSelectedItems.Add(SelectedItems(i))
                Next
            End If
            bSelecting = True
            MouseRect = New Rectangle(MouseDownX, MouseDownY, 1, 1)
            If count <> SelectedItems.Count Then
                If SelectedItems.Count <> count Then RaiseEvent SelectedItemChanged(Me, New EventArgs)
                Refresh()
            End If
        End Sub

        Private Sub F_MouseUp(ByVal fdSender As FieldDisplay, ByVal e As MouseEventArgs) Handles Me.MouseUp
            bDrag = False
            bSelecting = False
            If MouseDownX <> -1 Then
                If bDoubleClicked = True AndAlso SelectedItems.Count = 1 Then
                    If SelectedItems(0).Extension = "Folder" Then
                        LoadPage(SelectedItems(0).FullPath)
                    Else
                        If SelectedItems(0).Extension.Contains(".") Then System.Diagnostics.Process.Start(SelectedItems(0).FullPath, SelectedItems(0).Args)
                    End If
                End If
                If e.Button = Windows.Forms.MouseButtons.Right Then
                    If MouseRect.Width < 5 AndAlso MouseRect.Height < 5 Then
                        If SelectedItems.Count > 0 Then
                            Dim newDDM As DropDownMenu = Nothing
                            If SelectedItems.Count > 1 Then newDDM = New DropDownMenu(Exp.FParent, "Open in New Tabs", False) Else newDDM = New DropDownMenu(Exp.FParent, "Open", False)
                            If Not Exp.CBoard.GetData Is Nothing Then newDDM.AppendItems(DDMPaste)
                            newDDM.Tag = SelectedItems
                            newDDM.AppendItems(DDMFileEvents)
                            If SelectedItems.Count <> 1 Then newDDM.GetItemByName("Rename").Active = False
                            ShowDDM(newDDM, New Point(e.X, e.Y))
                        Else
                            Dim newDDM As New DropDownMenu(Exp.FParent, "Paste", False)
                            If Exp.CBoard.GetData Is Nothing Then
                                newDDM.Items(0).Active = False
                            Else
                                If Not TypeOf Exp.CBoard.GetData Is List(Of Item) Then newDDM.Items(0).Active = False
                            End If
                            newDDM.AddItem("Set Font")
                            newDDM.Tag = SelectedItems
                            ShowDDM(newDDM, New Point(e.X, e.Y))
                        End If
                    End If
                End If

                bDoubleClicked = False
                MouseDownX = -1
                MouseDownY = -1
                If SelectedItems.Count = 0 Then
                    ShiftStart = -1
                End If
                If Not MouseRect = Nothing AndAlso MouseRect.Width > 1 AndAlso MouseRect.Height > 1 Then
                    MouseRect = Nothing
                    Refresh()
                Else
                    MouseRect = Nothing
                End If
            End If
        End Sub

        Public Sub SetFont(ByRef tFont As Font, Optional ByRef clr As Color = Nothing, Optional ByRef ChangeIcons As Boolean = False)
            Font = tFont
            If clr = Nothing Then FontColor = New SolidBrush(Color.Black) Else FontColor = New SolidBrush(clr)
            Dim TSize As SizeF = MeasureAString(2, "jy0", Font)
            TextHeight = TSize.Height + ColumnVSpacing
            TextTop = 0
            If ChangeIcons = True Then
                IcoH = TSize.Height - 2
                IcoW = TSize.Height - 2
                If Item.ItemIcons.Count > 2 Then Item.ItemIcons.RemoveRange(2, Item.ItemIcons.Count - 2)
                If Item.ItemIcons(0).Object1.Width = IcoW Then Exit Sub
                For i = 0 To 1
                    Dim bmp As Bitmap = Item.ItemIcons(i).Object1
                    Dim bmp2 As New Bitmap(IcoW, IcoH)
                    Using gx As Graphics = Graphics.FromImage(bmp2)
                        Dim omg As New System.Drawing.Imaging.ImageAttributes()
                        gx.InterpolationMode = InterpolationMode.HighQualityBicubic
                        gx.DrawImage(bmp, 0, 0, IcoW, IcoH)
                    End Using
                    Item.ItemIcons(i).Object1 = bmp2
                Next
            End If
        End Sub

        Public Shared Sub SetDDMs()
            DDMFileEvents.AddRange(New List(Of String) From {{"Cut"}, {"Copy"}, {"Rename"}, {"Delete"}})
            DDMPaste.AddRange(New List(Of String) From {{"Paste"}})
        End Sub

        Private Sub DCTimer_Tick()
            System.Threading.Thread.Sleep(DCInterval)
            TimerReset()
            DCTimer = Nothing
        End Sub

        Private Sub TimerReset()
            DCX = -1
            DCY = -1
        End Sub

        Private Sub OpenPages(ByRef tItems As List(Of Item), ByRef iType As Integer)
            Dim NewPath As String = ""
            For i = 0 To tItems.Count - 1
                If tItems(i).Extension = "Folder" Then
                    If NewPath <> "" Then
                        Exp.NewTab(tItems(i).FullPath)
                    Else
                        NewPath = tItems(i).FullPath
                        If iType = 1 Then Exp.NewTab(NewPath)
                    End If
                Else
                    System.Diagnostics.Process.Start(tItems(i).FullPath, tItems(i).Args)
                End If
            Next
            tItems = Nothing
            If iType = 0 AndAlso NewPath <> "" Then LoadPage(NewPath)
        End Sub

        Private Sub DDM_Clicked(ByVal ddmSender As DropDownMenu, ByVal e As DropDownMenu.MenuItemClickedArgs)
            Dim tItems As List(Of Item) = ddmSender.Tag
            If e.Item.Text = "Open" Then
                If tItems(0).Extension = "Folder" Then
                    LoadPage(tItems(0).FullPath)
                Else
                    System.Diagnostics.Process.Start(tItems(0).FullPath, tItems(0).Args)
                End If
            ElseIf e.Item.Text.Contains("Open") AndAlso e.Item.Text.Contains("New Tab") Then
                If Exp.Preferences.OpenInNewTabsOnly = True Then OpenPages(tItems, 1) Else OpenPages(tItems, 0)
            ElseIf e.Item.Text = "Cut" Then
                If Not tItems Is Nothing Then
                    CopyType = 1
                    CopyStart = Me
                    Exp.CBoard.SetData(tItems)
                End If
            ElseIf e.Item.Text = "Copy" Then
                If Not tItems Is Nothing Then
                    CopyType = 0
                    CopyStart = Me
                    Exp.CBoard.SetData(tItems)
                End If
            ElseIf e.Item.Text = "Paste" Then
                Dim obj As Object = Exp.CBoard.GetData
                If TypeOf obj Is List(Of Item) Then
                    Dim objItems As List(Of Item) = obj
                    Dim sDestination As String = ""
                    Dim iType As Integer = -1
                    If tItems.Count = 0 Then
                        Dim newThread As New Thread(AddressOf PasteFilesToDestThread)
                        newThread.Start(New List(Of Object) From {{objItems}, {Page}, {CopyType}})
                    Else
                        Dim newThread As New Thread(AddressOf PasteFilesToThread)
                        newThread.Start(New List(Of Object) From {{objItems}, {tItems}, {CopyType}})
                    End If
                End If
            ElseIf e.Item.Text = "Rename" Then
                Dim itm As Item = tItems(0)
                RenameItemEvent(itm)
            ElseIf e.Item.Text = "Delete" Then
                Dim sMessage As String = ""
                If tItems.Count = 1 Then sMessage = "Are you sure you want to delete this item?" & vbNewLine & vbNewLine & tItems(0).Text Else sMessage = "Are you sure you want to delete these " & tItems.Count & " files?"
                If MessageBox.Show(sMessage, "Do it?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Exit Sub
                Dim newThread As New Thread(AddressOf DeleteFiles)
                newThread.Start(tItems)
            ElseIf e.Item.Text = "Set Font" Then
                Exp.SetFont()
            End If
        End Sub

        Private Sub DeleteFiles(ByVal arrList As List(Of Item))
            Dim sf As New SafeFile
            Dim count As Integer = arrList.Count
            For i = 0 To arrList.Count - 1
                If arrList(i).Extension = "Folder" Then
                    If sf.DeleteItem(arrList(i).FullPath, 1, False) = True Then count -= 1
                Else
                    If sf.DeleteItem(arrList(i).FullPath, 0, False) = True Then count -= 1
                End If
            Next
            If count <> arrList.Count Then Me.Invoke(New Action(Of String)(AddressOf Me.LoadPage), Page)
            System.Threading.Thread.CurrentThread.Abort()
        End Sub

        Private Sub PasteFilesToDestThread(ByVal arrStuff As List(Of Object)) 'ByVal arrItems As List(Of Item), ByVal sDestination As String, ByVal tCopyType As Byte)
            Dim arrItems As List(Of Item) = arrStuff(0)
            Dim sDestination As String = arrStuff(1)
            Dim tCopyType As Integer = arrStuff(2)
            Dim SF As New SafeFile
            For i = 0 To arrItems.Count - 1
                Dim ItemType As Integer = 0
                If arrItems(i).Extension = "Folder" Then ItemType = 1
                If tCopyType = 1 Then
                    SF.MoveItem(arrItems(i).FullPath, sDestination, ItemType)
                Else
                    SF.CopyItemTo(arrItems(i).FullPath, sDestination, ItemType)
                End If
            Next
            Me.Invoke(New Action(Of String)(AddressOf Me.LoadPage), Page)
            Exp.CBoard.Clear()
            System.Threading.Thread.CurrentThread.Abort()
        End Sub

        Private Sub PasteFilesToThread(ByVal arrStuff As List(Of Object))
            Dim arrItems As List(Of Item) = arrStuff(0)
            Dim arrDestinations As List(Of Item) = arrStuff(1)
            Dim tCopyType As Integer = arrStuff(2)
            Dim SF As New SafeFile

            For i = 0 To arrDestinations.Count - 1
                For i2 = 0 To arrItems.Count - 1
                    Dim ItemType As Integer = 0
                    If arrItems(i).Extension = "Folder" Then ItemType = 1
                    If tCopyType = 1 Then
                        SF.MoveItem(arrItems(i2).FullPath, arrDestinations(i).FullPath, ItemType)
                    Else
                        SF.CopyItemTo(arrItems(i2).FullPath, arrDestinations(i).FullPath, ItemType)
                    End If
                Next
            Next
            Me.Invoke(New Action(Of String)(AddressOf Me.LoadPage), Page)
            Exp.CBoard.Clear()
            System.Threading.Thread.CurrentThread.Abort()
        End Sub

        Private Sub ShowDDM(ByRef ddm As DropDownMenu, ByRef pt As Point)
            If Not DropDownMenu.ActiveDDM Is Nothing Then DropDownMenu.ActiveDDM.Dismiss()
            GetFormCoords(Me, pt)
            ' Exp.DDMHandledFirst = True
            AddHandler ddm.MenuItemClicked, AddressOf DDM_Clicked
            RaiseEvent DDMCreated(ddm, New EventArgs)
            ddm.ShowF(pt.X, pt.Y)
        End Sub

        Private Sub RenameItemEvent(ByVal tItem As Item)
            SelectedItems.Remove(tItem)
            tItem.Highlighted = False
            Refresh()
            Exp.SetTB(Me, tItem.TextRect.X + 2, tItem.TextRect.Y, tItem.TextSize.Width, tItem.Text, Font, FontColor.Color)
            Exp.TB.SelectionStart = 0
            Exp.TB.SelectionLength = Exp.TB.Text.Length
            Exp.TBHandledFirst = True
            TextItem = tItem
        End Sub

        Private Sub HighlightShift(ByVal i As Integer, Optional ByRef reset As Boolean = True)
            If ShiftStart = i Then Exit Sub
            If reset = True Then
                For i2 = 0 To OldShiftItems.Count - 1
                    OldShiftItems(i2).Highlighted = False
                    SelectedItems.Remove(OldShiftItems(i2))
                Next
                OldShiftItems = New List(Of Item)
            End If
            If ShiftStart < i Then
                For i2 = ShiftStart + 1 To i
                    If Not SelectedItems.Contains(Items(i2)) Then
                        SelectItem(i2)
                        OldShiftItems.Add(Items(i2))
                    End If
                Next
            Else
                For i2 = i To ShiftStart - 1
                    If Not SelectedItems.Contains(Items(i2)) Then
                        SelectItem(i2)
                        OldShiftItems.Add(Items(i2))
                    End If
                Next
            End If
        End Sub

        Private Function GetItemRanges(ByRef x As Integer, ByRef y As Integer) As Point
            Try
                Dim SC As Integer = -1, EC As Integer = -1
                For i = StartColumn To Columns.Count - 1
                    If SC = -1 AndAlso Columns(i).X > x Then
                        If i = 0 Then SC = Columns(i).Y Else SC = Columns(i - 1).Y
                    End If
                    If Columns(i).X >= y Then
                        EC = Columns(i).Y
                        Exit For
                    End If
                Next
                If SC = -1 Then SC = Columns(Columns.Count - 1).Y
                If EC = -1 Then EC = Items.Count
                EC -= 1
                Return New Point(SC, EC)
            Catch ex As Exception
                MsgBox("GetItemRanges: " & ex.Message)
                Return Nothing
            End Try
        End Function

        Public Sub UpdateView()
            Try
                bScrollHorizontal = False
                If Not Scrollbars(0) Is Nothing Then Scrollbars(0).Dispose()
                Dim RunType As Integer = 0
                Columns = New List(Of Point)
                Columns.Add(New Point(0, 0))
                StartColumn = -1
                Do
                    Dim TotalHeight As Integer = BorderThickness + TBuffer
                    Dim LongestWidth As Integer = BorderThickness + LBuffer
                    Dim CX As Integer = LBuffer, CY As Integer = BorderThickness + TBuffer
                    Dim modifier As Integer = 0
                    If bScrollHorizontal = True Then
                        ColumnHeight = Height - BorderThickness - 16
                        If Not Scrollbars(0) Is Nothing Then modifier = Scrollbars(0).ScrollPosition Else modifier = 0
                        CX += modifier
                    Else
                        ColumnHeight = Height - BorderThickness
                    End If
                    For i = 0 To Items.Count - 1
                        If (CY + Items(i).TextSize.Height) >= ColumnHeight Then
                            CX += LongestWidth + ColumnHSpacing
                            If bScrollHorizontal = False AndAlso CX > (Width - BorderThickness) Then
                                bScrollHorizontal = True
                                StartColumn = 0
                                Continue Do
                            End If
                            If RunType = 1 Then
                                If StartColumn = -1 Then If CX >= 0 Then If Columns.Count = 0 Then StartColumn = 0 Else StartColumn = Columns.Count - 1
                                Columns.Add(New Point(CX, i))
                            End If
                            CY = BorderThickness + TBuffer
                            LongestWidth = 0
                        End If
                        Items(i).rect = New Rectangle(CX + IcoW, CY, Math.Min(IcoW + Items(i).TextSize.Width, MaxColumnWidth), TextHeight)
                        Items(i).TextRect = New Rectangle(Items(i).rect.X + IcoW, Items(i).rect.Y + TextTop, MaxColumnWidth - IcoW, Items(i).TextSize.Height)
                        LongestWidth = Math.Min(Math.Max(LongestWidth, Items(i).rect.Width), MaxColumnWidth)
                        CY += Items(i).TextSize.Height + ColumnVSpacing
                    Next
                    If RunType = 0 Then
                        RunType = 1
                        Continue Do
                    End If
                    CX += LongestWidth + ColumnHSpacing
                    If bScrollHorizontal = False AndAlso CX > (Width - BorderThickness) Then bScrollHorizontal = True
                    If bScrollHorizontal = True Then
                        Scrollbars(0) = New ScrollBar(Me, BorderThickness, Height - BorderThickness - 16, WWidth, 16, 0, CX + 10)
                        AddHandler Scrollbars(0).ScrollBarMoved, AddressOf ScrollBarMoved
                    End If
                    Exit Do
                Loop
                If StartColumn = -1 Then
                    StartColumn = 0
                End If
            Catch ex As Exception
                MsgBox("Field Sub Update View: " & ex.Message)
            End Try

        End Sub

        Private Sub ScrollBarMoved(ByVal sbSender As ScrollBar, ByVal e As ScrollBar.BarMoveArgs)
            UpdateScrollPosition(e.difference)
            Refresh()
        End Sub

        Private Sub UpdateScrollPosition(ByRef modifier As Integer)
            StartColumn = -1
            For i = 0 To Items.Count - 1
                Items(i).rect.X += modifier
                Items(i).TextRect.X += modifier
            Next
            For i = 0 To Columns.Count - 1
                Columns(i) = New Point(Columns(i).X + modifier, Columns(i).Y)
                If StartColumn = -1 Then If Columns(i).X >= 0 Then StartColumn = i
            Next
            If StartColumn = -1 Then StartColumn = 0
        End Sub

        Private Sub SelectItem(ByRef tItem As Item)
            If Not SelectedItems.Contains(tItem) Then
                tItem.Highlighted = True
                SelectedItems.Add(tItem)
            End If
        End Sub

        Private Sub SelectItem(ByRef i As Integer)
            If Not SelectedItems.Contains(Items(i)) Then
                Items(i).Highlighted = True
                SelectedItems.Add(Items(i))
            End If
        End Sub

        Private Overloads Sub DeselectItem(ByRef tItem As Item)
            If tItem.Highlighted = True Then
                tItem.Highlighted = False
                SelectedItems.Remove(tItem)
            End If
        End Sub

        Private Overloads Sub DeselectItem(ByRef i As Integer)
            If Items(i).Highlighted = True Then
                Items(i).Highlighted = False
                SelectedItems.Remove(Items(i))
            End If
        End Sub

        Public Sub UpdateSizes()
            For i = 0 To Items.Count - 1
                Items(i).TextSize = MeasureAString(2, Items(i).Text, Font)
            Next
            UpdateView()
            Refresh()
        End Sub

        Public Sub PopulateItems()
            Folders = di.GetDirectories
            Files = di.GetFiles
            Items = New List(Of Item)
            bScrollHorizontal = False
            For Each fo As System.IO.DirectoryInfo In Folders
                Try
                    If Exp.ShowSystemFolders = False AndAlso fo.Attributes.ToString.Contains("System") Then Continue For
                Catch ex As Exception
                End Try
                Dim newSZ As SizeF = MeasureAString(2, fo.Name, Font)
                Dim newItem As New Item(fo, newSZ)
                Items.Add(newItem)
            Next
            For Each fi As System.IO.FileInfo In Files
                Try
                    Dim name As String = If(fi.Extension = "", fi.Name.Remove(fi.Name.Length - fi.Extension.Length), fi.Name)
                    Dim newSZ As SizeF = MeasureAString(2, name, Font)
                    Dim newItem As New Item(fi, newSZ)
                    Items.Add(newItem)
                Catch ex As Exception

                End Try
            Next
            If Items.Count > 0 Then imgHeight = ((Items(0).TextSize.Height) / 2) - (IcoH / 2)
        End Sub

        Public Sub DeselectItems(Optional ByRef raiseanevent As Boolean = True)
            For i = 0 To SelectedItems.Count - 1
                SelectedItems(i).Highlighted = False
            Next
            SelectedItems = New List(Of Item)
            If raiseanevent = True Then RaiseEvent SelectedItemChanged(Me, New EventArgs)
        End Sub

        Public Sub RefreshPage()
            Dim TempArray As New List(Of String), PScroll As Integer = 0
            If Not Scrollbars(0) Is Nothing Then PScroll = Scrollbars(0).ScrollPosition
            For i = 0 To SelectedItems.Count - 1
                TempArray.Add(SelectedItems(i).Text)
            Next
            DeselectItems(False)
            PopulateItems()
            UpdateView()
            If PScroll <> 0 And Not Scrollbars(0) Is Nothing Then
                Scrollbars(0).UpdateScroll(PScroll)
            End If
            Dim index As Integer = 0
            For i = 0 To TempArray.Count - 1
                For i2 = index To Items.Count - 1
                    If Items(i2).Text = TempArray(i) Then
                        index = i
                        Items(i2).Highlighted = True
                        SelectedItems.Add(Items(i2))
                    End If
                Next
            Next
        End Sub

        Public Sub LoadPage(ByVal sPage As String, Optional ByVal direction As Integer = 0, Optional ByVal bRefresh As Boolean = True)
            If Not ActiveField Is Me Then ActiveField = Me
            If Not System.IO.Directory.Exists(sPage) Then
                MsgBox("Directory no longer exists.")
                Exp.AddressBar.UpdateHistoryList()
                Exit Sub
            End If
            If Not HoverItem Is Nothing Then HoverItem = Nothing
            If Not sPage.Contains("\") Then sPage &= "\"
            If sPage.Last <> "\" Then sPage &= "\"
            Page = sPage
            di = New System.IO.DirectoryInfo(Page)
            For i = 0 To 1
                If Not Scrollbars(i) Is Nothing Then
                    Scrollbars(i).Dispose()
                    Scrollbars(i) = Nothing
                End If
            Next
            DeselectItems(False)
            If bRefresh = True Then PopulateItems()
            UpdateView()
            If Initializing = True Then
                Me.Refresh()
                Initializing = False
            Else
                If bRefresh = True Then
                    Refresh()
                    If direction < 3 Then
                        If Not PParent Is Nothing Then
                            If PParent.Header.Text <> di.Name Then Exp.TBMain.SetPageName(Parent, di.Name)
                        End If
                    End If
                End If
            End If
            RaiseEvent DirectoryChanged(Me, New EventArgs)
            Exp.AddressBar.SetURL(Me, sPage, direction)
        End Sub

        Public Class Item
            Public IconIndex As Integer = -1
            Public Text As String, FullPath As String, Args As String
            Public Extension As String
            Public TextSize As SizeF
            Public rect As Rectangle, TextRect As Rectangle
            Public Highlighted As Boolean = False
            Public Shared ItemIcons As List(Of StringAndObject), ErrorIcon As Integer = -1

            Public Sub New(ByVal fo As System.IO.DirectoryInfo, ByVal sSize As SizeF, Optional ByVal arguments As String = "")
                Text = fo.Name
                FullPath = fo.FullName
                Extension = "Folder"
                NewItem(sSize, arguments)
            End Sub

            Public Sub New(ByVal fi As System.IO.FileInfo, ByVal sSize As SizeF, Optional ByVal arguments As String = "")
                Text = fi.Name.Remove(fi.Name.Length - fi.Extension.Length)
                FullPath = fi.FullName
                Extension = fi.Extension
                NewItem(sSize, arguments)
            End Sub

            Public Sub NewItem(ByVal sSize As SizeF, ByVal arguments As String)
                TextSize = sSize
                Args = arguments
                SetIconIndex()
            End Sub

            Private Sub SetIconIndex()
                For i = 0 To ItemIcons.Count - 1
                    If ItemIcons(i).String1 = Extension Then
                        IconIndex = i
                        Exit Sub
                    End If
                Next
                Try
                    'Dim ico As Icon
                    Dim bmp As Bitmap = System.Drawing.Icon.ExtractAssociatedIcon(FullPath).ToBitmap
                    Dim bmp2 As New Bitmap(IcoW, IcoH)
                    Using gx As Graphics = Graphics.FromImage(bmp2)
                        Dim omg As New System.Drawing.Imaging.ImageAttributes()
                        gx.InterpolationMode = InterpolationMode.HighQualityBicubic
                        gx.DrawImage(bmp, 0, 0, IcoW, IcoH)
                    End Using
                    ItemIcons.Add(New StringAndObject(Extension, bmp2))
                    IconIndex = ItemIcons.Count - 1
                Catch ex As Exception
                    IconIndex = ErrorIcon
                End Try
            End Sub

            Public Function GetIcon() As Bitmap
                Return ItemIcons(IconIndex).Object1
            End Function

            Public Function IconsContains(ByVal sExtension As String) As Boolean
                For i = 0 To ItemIcons.Count - 1
                    If ItemIcons(i).String1 = sExtension Then Return True
                Next
                Return False
            End Function
        End Class

    End Class


    Public Class BottomBar 'Below the field, displays item information
        Inherits System.Windows.Forms.Control

        Public Exp As Explorer
        Public Items As New List(Of TwoStrings)
        Public NoPaint As Boolean = False

        Public Shared ItemTemplates As New List(Of Item)
        Public Shared BorderThickness As Integer, DoubleBorder As Integer, BorderColor As Pen, BGColor As SolidBrush
        Public Shared FontColor As New SolidBrush(Color.Black), NameFontColor As New SolidBrush(Color.Red), DescriptorFont As New Font("Microsoft Sans Serif", 8.25), ItemFont As New Font("Microsoft Sans Serif", 10)
        Public Shared OmgLOL As Integer = 0
        Private Shared TextHeight As Integer, WHeight As Integer = 0
        Private IconImg As Bitmap, IconLBuffer As Integer = 15, IconTBuffer As Integer = 5, IconW As Integer, IconH As Integer
        Private LItemBuffer As Integer, ItemTBuffer As Integer = 3
        Private di As System.IO.DirectoryInfo
        Private FieldItemCount As Integer = 0
        Private EmptyList As List(Of String)
        Private ItemVSpacing As Integer = 2, ItemHSpacing As Integer = 10
        Private DrawFormat As New System.Drawing.StringFormat

        Public Sub New(ByRef tExplorer As Explorer)
            Exp = tExplorer
            Parent = Exp.Parent
            SetBorder(1)
            UpdatePosition() 'SetBounds(0, Exp.TBMain.Bottom, Exp.WWidth, Exp.WHeight - Exp.TBMain.Bottom)
            DrawFormat.Alignment = StringAlignment.Far
            If Not Exp.Field Is Nothing Then
                di = New System.IO.DirectoryInfo(Exp.Field.Page)
                UpdateItems()
            End If
            AddHandler Me.MouseMove, AddressOf B_MouseMove
            AddHandler Me.Paint, AddressOf Me.B_Paint
        End Sub

        Public Shared Sub SetFont(ByRef tFont As Font, ByRef tFontColor As Color)
            DescriptorFont = tFont
            FontColor = New SolidBrush(tFontColor)
        End Sub

        Public Sub UpdatePosition()
            NoPaint = False
            SetBounds(0, Exp.TBMain.Pages(0).Bottom, Exp.WWidth, Exp.WHeight - Exp.TBMain.Pages(0).Bottom)
            IconH = Height - (IconTBuffer * 2)
            IconW = IconH
            WHeight = Height - (BorderThickness * 2)
        End Sub

        Public Sub DirectoryChanged(ByVal sender As Object, ByVal e As EventArgs)
            di = New System.IO.DirectoryInfo(Exp.Field.Page)
            UpdateItems()
        End Sub

        Private Sub FieldChanged(ByVal sender As Object, ByVal e As EventArgs)
            di = New System.IO.DirectoryInfo(Exp.Field.Page)
            UpdateItems()
        End Sub

        Public Overloads Sub AddItem(ByRef sName As String, Optional ByRef sValue As String = "", Optional ByRef sVisible As Boolean = True)
            ' Items.Add(New Item(sName, New List(Of String) From {{sValue}}, sVisible))
        End Sub

        Public Overloads Sub AddItem(ByRef sName As String, ByRef sValues As List(Of String), Optional ByRef sVisible As Boolean = True)
            '     Items.Add(New Item(sName, sValues, sVisible))
        End Sub

        Private Function ThumbnailCallback() As Boolean
            Return False
        End Function

        Public Sub UpdateItems()
            Items = New List(Of TwoStrings)
            If Exp.Field Is Nothing Then Exit Sub
            If Exp.Field.SelectedItems Is Nothing Then Exit Sub
            FieldItemCount = Exp.Field.SelectedItems.Count
            If FieldItemCount = 0 Then
                SetIconImg(Explorer.FieldDisplay.Item.ItemIcons(1).Object1)
                EmptyList = GetDirectoryReport()
            ElseIf FieldItemCount = 1 Then
                Dim tItem As FieldDisplay.Item = Exp.Field.SelectedItems(0)
                If tItem.Extension <> "Folder" Then
                    Dim FT As FileType = Exp.GetFileType(tItem.FullPath)
                    Items = Exp.GetFileInfo(tItem.FullPath, FT)
                    ' Dim sLol As String = "", iCount As Integer = 0
                    'For i = 0 To lolo.Count - 1
                    'sLol &= lolo(i) & vbNewLine
                    'Next
                    'Clipboard.SetText(sLol)
                    '    MsgBox(sLol)
                    Items(0).String1 = ""
                    Items(1).String1 = ""
                    If FT.Parent.Name = "Images" Then
                        Dim img As Image = Image.FromFile(tItem.FullPath)
                        Dim w As Integer = (IconH / img.Height) * img.Width
                        IconImg = img.GetThumbnailImage(w, IconH, New System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback), IntPtr.Zero)
                    Else
                        SetIconImg(tItem.GetIcon)
                    End If
                Else
                    SetIconImg(tItem.GetIcon)
                End If
            Else

            End If
            Me.Refresh()
        End Sub

        Private Sub SetIconImg(ByRef bmp As Bitmap)
            IconImg = New Bitmap(IconW, IconH)
            Using gx As Graphics = Graphics.FromImage(IconImg)
                gx.DrawImage(bmp, 0, 0, IconW, IconH)
            End Using
            LItemBuffer = IconLBuffer + IconW + 15
        End Sub

        Public Sub FieldItemChanged(ByVal sender As Object, ByVal e As EventArgs)
            UpdateItems()
        End Sub

        Public Sub SetBorder(ByRef thickness As Integer)
            BorderThickness = thickness
            DoubleBorder = thickness * 2
        End Sub

        Private Sub B_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs)

        End Sub

        Private Function GetDirectoryReport() As List(Of String)
            Dim arrList As New List(Of String) From {{(Exp.Field.Folders.Count + Exp.Field.Files.Count) & " items: " & Exp.Field.Folders.Count & " Folders, " & Exp.Field.Files.Count & " Files"}}
            Dim arrItems As New List(Of String)
            For Each fi As System.IO.FileInfo In Exp.Field.Files

            Next
            Return arrList
        End Function

        Public MaxColumnWidth As Integer = 500

        Public Sub B_Paint(ByVal sender As Object, ByVal e As PaintEventArgs)
            '  e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
            If NoPaint = True Then Exit Sub
            OmgLOL += 1
            For i = 0 To BorderThickness - 1
                e.Graphics.DrawRectangle(BorderColor, i, i, Width - (i * 2) - 1, Height - (i * 2) - 1)
            Next
            If Not Exp.Field Is Nothing Then
                If IconImg Is Nothing Then UpdateItems()
                e.Graphics.DrawImage(IconImg, IconLBuffer, IconTBuffer)
                Dim CX As Integer = LItemBuffer, CY As Integer = ItemTBuffer, SSDesc As New SizeF, SSItem As New SizeF
                If FieldItemCount = 0 Then
                    SSDesc = MeasureAString(2, EmptyList(0), DescriptorFont)
                    For i = 0 To EmptyList.Count - 1
                        SSDesc = MeasureAString(2, EmptyList(0), DescriptorFont)
                        e.Graphics.DrawString(EmptyList(i), ItemFont, FontColor, CX, CY)
                        CY += SSDesc.Height + ItemVSpacing 'Item Buffer
                        If CY + SSDesc.Height > WHeight Then
                            CY = ItemTBuffer
                            CX += SSDesc.Width + ItemHSpacing
                        End If
                    Next
                ElseIf FieldItemCount = 1 Then
                    Dim DescRect As New RectangleF, ItemRect As New RectangleF, longestwidth As Integer = 0, w As Integer = 0
                    If Items.Count > 0 Then 'File selected, not folder

                        For i = 0 To 1
                            If i = 0 Then SSItem = MeasureAString(2, Items(i).String2, ItemFont) Else SSItem = MeasureAString(2, Items(i).String2, DescriptorFont)
                            If CY + SSItem.Height > WHeight Then
                                CY = ItemTBuffer
                                CX += longestwidth + ItemHSpacing
                                longestwidth = 0
                            End If
                            w = Math.Min(SSItem.Width + 5, MaxColumnWidth)
                            longestwidth = Math.Max(longestwidth, w)
                            If i = 0 Then e.Graphics.DrawString(Items(i).String2, ItemFont, FontColor, CX, CY) Else e.Graphics.DrawString(Items(i).String2, DescriptorFont, FontColor, CX, CY)
                            CY += SSItem.Height + ItemVSpacing
                        Next
                        Dim Rects As New List(Of Rectangle), Points As New List(Of Point)
                        For i = 2 To Items.Count - 1
                            SSDesc = MeasureAString(2, Items(i).String1, DescriptorFont)
                            If CY + SSItem.Height > WHeight Then
                                CY = ItemTBuffer
                                CX += longestwidth + ItemHSpacing
                                longestwidth = 0
                            End If
                        Next


                    Else 'EmptyList in use, folder selected

                    End If
                End If
            End If
            e.Graphics.DrawString(OmgLOL, Font, FontColor, 15, 15)
        End Sub

        Protected Overrides Sub OnPaintBackground(ByVal e As PaintEventArgs)
            e.Graphics.FillRectangle(BGColor, BorderThickness, BorderThickness, Width - DoubleBorder, Height - DoubleBorder)
        End Sub

        Public Class Item
            Public Name As String, VisibleName As Boolean
            Public Values As New List(Of String)

            Public Sub New(ByRef sName As String, ByRef sValues As List(Of String), Optional ByRef sVisibleName As Boolean = True)
                Name = sName
                Values = sValues
                VisibleName = sVisibleName
            End Sub
        End Class

    End Class

    Public Class SideBar 'Tree View on side

    End Class

End Class



'Supporting components

Public Class ClipBoardX
    Public obj As Object
    Public Sub New()
        obj = Nothing
    End Sub
    Public Sub New(ByRef tObject As Object)
        obj = tObject
    End Sub
    Public Sub SetData(ByRef tObject As Object)
        obj = tObject
    End Sub
    Public Function GetData() As Object
        Return obj
    End Function
    Public Sub Clear()
        obj = Nothing
    End Sub
End Class

Public Class SafeFile

    Public ErrorTitle As String
    Public ErrorStringStart As String

    Public Sub New(Optional ByRef sTitle As String = "stop messing up!", Optional ByRef sErrorStringStartPhrase As String = "Error: ")
        ErrorTitle = sTitle
        ErrorStringStart = sErrorStringStartPhrase
    End Sub

    Public Function CheckFileExists(ByVal sFile As String, Optional ByVal blnShowError As Boolean = False) As Boolean
        Try
            If System.IO.File.Exists(sFile) Then Return True
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not check file exists for file " & sFile & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
        If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not find file " & Chr(34) & sFile & Chr(34), ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return False
    End Function

    Public Function DeleteItem(ByVal sItem As String, ByVal File0Folder1 As Integer, Optional ByVal blnShowDialogs As Boolean = True, Optional ByVal blnShowError As Boolean = True) As Boolean
        Try
            If File0Folder1 = 0 Then
                If System.IO.File.Exists(sItem) = True Then
                    If blnShowDialogs = True Then My.Computer.FileSystem.DeleteFile(sItem, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin) Else My.Computer.FileSystem.DeleteFile(sItem, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    Return True
                Else
                    If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not find file " & sItem, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
            Else
                If System.IO.Directory.Exists(sItem) = True Then
                    If blnShowDialogs = True Then My.Computer.FileSystem.DeleteDirectory(sItem, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin) Else My.Computer.FileSystem.DeleteDirectory(sItem, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    Return True
                Else
                    If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not find directory " & sItem, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
            End If
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not delte file " & Chr(34) & sItem & Chr(34) & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Public Function MoveItems(ByVal lstItems As List(Of String), ByVal sNewLocation As String, ByVal File0Folder1 As Integer)
        Try
            For i = 0 To lstItems.Count - 1
                MoveItem(lstItems(i), sNewLocation, File0Folder1)
            Next
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function CopyItemTo(ByVal sItem As String, ByVal sNewLocation As String, ByVal File0Folder1 As Integer)
        Try
            If sItem.Last = "\" Then sItem = sItem.Remove(sItem.Length - 1)
            Dim sName As String = sItem.Substring(sItem.LastIndexOf("\") + 1)
            If sNewLocation.Last <> "\" Then sNewLocation &= "\"
            sNewLocation &= sName
            sNewLocation = GetNewName(sNewLocation, File0Folder1)
            If File0Folder1 = 0 Then
                My.Computer.FileSystem.CopyFile(sItem, sNewLocation, FileIO.UIOption.AllDialogs)
            Else
                My.Computer.FileSystem.CopyDirectory(sItem, sNewLocation, FileIO.UIOption.AllDialogs)
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function MoveItem(ByVal sItem As String, ByVal sNewLocation As String, ByVal File0Folder1 As Integer, Optional ByRef blnShowError As Boolean = True) As Boolean
        Try
            If sItem.Last = "\" Then sItem = sItem.Remove(sItem.Length - 1)
            Dim sName As String = sItem.Substring(sItem.LastIndexOf("\") + 1)

            If File0Folder1 = 0 Then
                If sNewLocation.Last <> "\" Then sNewLocation &= "\"
                sNewLocation &= sName
                My.Computer.FileSystem.MoveFile(sItem, sNewLocation, FileIO.UIOption.AllDialogs)
            Else
                My.Computer.FileSystem.MoveDirectory(sItem, sNewLocation, FileIO.UIOption.AllDialogs)
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function GetNewName(ByVal sName As String, ByVal File0Folder1 As Integer) As String
        If System.IO.File.Exists(sName) Then
            Dim BaseName As String = sName.Substring(0, sName.LastIndexOf("."))
            Dim ExtName As String = ""
            If File0Folder1 = 0 Then ExtName = sName.Substring(sName.LastIndexOf("."), sName.Length - sName.LastIndexOf("."))
            Dim i As Short = 2
            Do
                If i = 10000 Then Exit Do
                If System.IO.File.Exists(BaseName & " (" & i & ")" & ExtName) = False Then
                    sName = BaseName & " (" & i & ")" & ExtName
                    Return sName
                Else
                    i += 1
                End If
            Loop
        Else
            Return sName
        End If
        Return sName
    End Function

    Public Function RenameFile(ByVal sOldFileName As String, ByVal sNewFileName As String, Optional ByRef blnHandleSameFileName As Boolean = False, Optional ByVal blnShowError As Boolean = True) As Boolean
        Try
            If sNewFileName.Contains("\") Then sNewFileName = sNewFileName.Substring(sNewFileName.LastIndexOf("\") + 1, sNewFileName.Length - sNewFileName.LastIndexOf("\"))
            If System.IO.File.Exists(sNewFileName) Then
                If blnHandleSameFileName = True Then
                    Dim BaseName As String = sNewFileName.Substring(0, sNewFileName.LastIndexOf("."))
                    Dim ExtName As String = sNewFileName.Substring(sNewFileName.LastIndexOf("."), sNewFileName.Length - sNewFileName.LastIndexOf("."))
                    Dim i As Short = 2
                    Do
                        If i = 10000 Then Exit Do
                        If System.IO.File.Exists(BaseName & " (" & i & ")" & ExtName) = False Then
                            sNewFileName = BaseName & " (" & i & ")" & ExtName
                            Exit Do
                        Else
                            i += 1
                        End If
                    Loop
                Else
                    If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not rename file " & Chr(34) & sOldFileName & Chr(34) & " to " & Chr(34) & sNewFileName & Chr(34) & " because a file with that name already exists.", ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
            End If
            My.Computer.FileSystem.RenameFile(sOldFileName, sNewFileName)
            Return True
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not rename file " & Chr(34) & sOldFileName & Chr(34) & " to " & Chr(34) & sNewFileName & Chr(34) & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Public Function SW_CreateText(ByVal sFile As String, ByRef sw As System.IO.StreamWriter, Optional ByVal blnDeletePrevFile As Boolean = True, Optional ByVal blnShowError As Boolean = True) As System.IO.StreamWriter
        Try
            If blnDeletePrevFile = True Then If DeleteItem(sFile, blnShowError) = False Then Return Nothing
            sw = System.IO.File.CreateText(sFile)
            Return sw
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not create text file from streamwriter " & Chr(34) & sFile & Chr(34) & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function SR_OpenText(ByVal sFile As String, ByRef sr As System.IO.StreamReader, Optional ByVal blnShowError As Boolean = True) As System.IO.StreamReader
        Try
            sr = System.IO.File.OpenText(sFile)
            Return sr
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not open text file from streamreader " & Chr(34) & sFile & Chr(34) & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Public Function GetArrayLinesFromStream(ByVal sFile As String, Optional ByRef blnRemoveEmptyLines As Boolean = False, Optional ByVal blnTrim As Boolean = False, Optional ByVal blnSort As Boolean = False, Optional ByVal blnShowError As Boolean = True) As List(Of String)
        Try
            Dim arrTemp As New List(Of String)
            Using sr As System.IO.StreamReader = System.IO.File.OpenText(sFile)
                Do While sr.Peek <> -1
                    Dim sLine As String = sr.ReadLine
                    If blnRemoveEmptyLines = True Then If sLine = "" Then Continue Do
                    If blnTrim = True Then arrTemp.Add(sLine.Trim) Else arrTemp.Add(sLine)
                Loop
            End Using
            If blnSort = True Then arrTemp.Sort()
            Return arrTemp
        Catch ex As Exception
            If blnShowError = True Then MessageBox.Show(ErrorStringStart & "Could not read file " & Chr(34) & sFile & Chr(34) & " in streamreader." & vbNewLine & vbNewLine & ex.Message, ErrorTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

End Class