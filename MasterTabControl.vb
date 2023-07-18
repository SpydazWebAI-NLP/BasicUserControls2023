Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms
Imports BasicUserControls.Controls

Namespace TabControls

    Public Class MasterTabControl
        Private Shared Speaker As New Speech.Synthesis.SpeechSynthesizer
        Private Shared Speech_on As Boolean = True

        Public Class IEGetFavorites

            Public Shared Function GetFavoritesfromIE() As List(Of String)
                Dim list As New List(Of String)
                Try
                    For Each favorites As String In
                My.Computer.FileSystem.GetFiles(System.Environment.GetFolderPath(Environment.SpecialFolder.Favorites),
                                                FileIO.SearchOption.SearchAllSubDirectories, "*.url")
                        Using sr As System.IO.StreamReader = New System.IO.StreamReader(favorites)
                            While Not sr.EndOfStream
                                Dim line As String = sr.ReadLine()
                                If line.Contains("BASEURL=") Then
                                    list.Add(line.Replace("BASEURL=", ""))
                                    Exit While
                                End If
                            End While
                        End Using
                    Next
                Catch
                    Return Nothing

                End Try
                Return list
            End Function

        End Class

        Public Class WebbrowserFunction
            Inherits WebBrowser

            Private Sub webBrowserDocComplete() Handles Me.DocumentCompleted
                Dim tabpage As TabPage = Me.Tag
                'set length of tab max
                If Me.DocumentTitle.Length > 25 Then
                    tabpage.Text = Me.DocumentTitle.Substring(0, 25)
                Else
                    tabpage.Text = Me.DocumentTitle
                End If

            End Sub

            Public Sub New()
                Me.ScriptErrorsSuppressed = True
            End Sub

        End Class

#Region "TABS-Main Functions - Universal"

#Region "CheckTab"

        Public Shared Function IsWebTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsWebTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "Browser" Then Return True
            End If

        End Function

        Public Shared Function IsTestTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsTestTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "TEST" Then Return True
            End If

        End Function

        Public Shared Function IsTextTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsTextTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "TEXT" Then Return True
                If TabDocumentBrowser.SelectedTab.Text = "TextBox" Then Return True
            End If
        End Function

        Public Shared Function IsPluginTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsPluginTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "PLUGIN" Then Return True
            End If

        End Function

        Public Shared Function IsAvatarTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsAvatarTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "AVATAR" Then Return True
            End If

        End Function

        Public Shared Function IsUpdateTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsUpdateTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "UPDATE" Then Return True
            End If

        End Function

        Public Shared Function IsDeviceTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsDeviceTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "DEVICE" Then Return True
            End If

        End Function

        Public Shared Function IsScriptTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsScriptTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "Script" Or TabDocumentBrowser.SelectedTab.Text = "SCRIPT" Then Return True
            End If
        End Function

        Public Shared Function IsSnippitTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsSnippitTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "Snippit" Or TabDocumentBrowser.SelectedTab.Text = "SNIPPIT" Then Return True
            End If

        End Function

        ''' <summary>
        ''' Used to check if Tab is in collection (to disable duplicate tabs)
        ''' </summary>
        ''' <param name="Tabs"></param>
        ''' <param name="Text">title of tab</param>
        ''' <returns></returns>
        Public Shared Function CheckTab(ByRef Tabs As TabControl.TabPageCollection, ByRef Text As String) As Boolean
            For Each item In Tabs
                If item.text = Text Then
                    Return True
                End If
            Next
            Return False
        End Function

        Public Shared Function CheckTabExists(ByRef Sender As TabControl, Page As TabPage) As Boolean
            CheckTabExists = False
            For Each item As TabPage In Sender.TabPages
                If item.Name = Page.Name Then
                    Return True

                End If
            Next

        End Function

#End Region

#Region "Tab General"

        Public Shared Function GetSyntaxTextBox(ByRef Sender As TabPage) As SpydazWebTextBox
            Try
                Return CType(Sender.Controls.Item("Body"), SpydazWebTextBox)
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Shared Function GetRichTextBox(ByRef Sender As TabPage) As RichTextBox
            Try
                Return CType(Sender.Controls.Item("Body"), RichTextBox)
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Shared Function GetInteliTextBox(ByRef Sender As TabPage) As SpydazWebTextBox
            Try
                Return CType(Sender.Controls.Item("Body"), SpydazWebTextBox)
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Shared Sub RemoveAllTabs(ByRef Sender As TabControl)

            For Each Page As TabPage In Sender.TabPages

                If Page.Text = "HOME" = False And
                Page.Text = "COMPILER" = False Then
                    Sender.TabPages.Remove(Page)
                End If
            Next

        End Sub

        Public Shared Sub RemoveAllTabsButThis(ByRef Sender As TabControl, ByRef TabCount As Integer)
            For Each Page As TabPage In Sender.TabPages
                If Not Page.Name = Sender.SelectedTab.Name Then
                    If Page.Text = "HOME" = False And
                Page.Text = "COMPILER" = False Then
                        Sender.TabPages.Remove(Page)
                    End If
                End If
            Next
        End Sub

        Public Shared Sub CloseTab_Click(sender As TabControl)
            If sender.SelectedTab.Text <> "HOME" Or sender.SelectedTab.Text <> "COMPILER" Then
                RemoveTab(sender, sender.TabPages.Count)
                Speaker.Speak("Closing tab")
            End If
        End Sub

        Public Shared Sub RemoveTab(ByRef Sender As TabControl, ByRef Tabcount As Integer)
            If Sender.SelectedTab.Text <> "HOME" Or
            Sender.SelectedTab.Text <> "COMPILER" Or
           Sender.SelectedTab.Text <> "AI CHAT" Or
           Sender.SelectedTab.Text <> "WELCOME" Then
                If Not Sender.TabPages.Count = 1 Then
                    Sender.TabPages.Remove(Sender.SelectedTab)
                    Tabcount = Sender.TabPages.Count
                Else
                End If
            End If
        End Sub

        Public Shared Function GetCurrentTabText(ByRef Sender As TabControl) As String
            If IsTextTab(Sender) = True Or
            IsAvatarTab(Sender) = True Or
            IsDeviceTab(Sender) = True Or
            IsUpdateTab(Sender) = True Or
            IsSnippitTab(Sender) = True Or
            IsTestTab(Sender) = True Or
            IsPluginTab(Sender) = True Or
            IsScriptTab(Sender) = True Then
                Return CType(Sender.SelectedTab.Controls.Item("Body"), RichTextBox).Text
            Else
                Try
                    Return CType(Sender.SelectedTab.Controls.Item("Body"), RichTextBox).Text
                Catch ex As Exception

                End Try
            End If
            Return Nothing
        End Function

        Public Shared Function GetCurrentTabSyntaxText(ByRef Sender As TabControl) As String
            If IsTextTab(Sender) = True Or
            IsAvatarTab(Sender) = True Or
            IsDeviceTab(Sender) = True Or
            IsUpdateTab(Sender) = True Or
            IsSnippitTab(Sender) = True Or
            IsTestTab(Sender) = True Or
            IsPluginTab(Sender) = True Or
            IsScriptTab(Sender) = True Then
                Return CType(Sender.SelectedTab.Controls.Item("Body"), SpydazWebTextBox).Text
            Else
                Try
                    Return CType(Sender.SelectedTab.Controls.Item("Body"), SpydazWebTextBox).Text
                Catch ex As Exception

                End Try
            End If
            Return Nothing
        End Function

        Public Shared Function GetCurrentTabRichTextBox(ByRef Sender As TabControl) As RichTextBox
            If IsTextTab(Sender) = True Or
            IsAvatarTab(Sender) = True Or
            IsSnippitTab(Sender) = True Or
            IsDeviceTab(Sender) = True Or
            IsUpdateTab(Sender) = True Or
            IsTestTab(Sender) = True Or
            IsPluginTab(Sender) = True Or
             IsScriptTab(Sender) = True Then

                Return CType(Sender.SelectedTab.Controls.Item("Body"), RichTextBox)
            Else
                Try
                    Return CType(Sender.SelectedTab.Controls.Item("Body"), RichTextBox)
                Catch ex As Exception

                End Try
            End If
            Return Nothing
        End Function

        Public Shared Function GetCurrentTabSyntaxTextBox(ByRef Sender As TabControl) As SpydazWebTextBox
            If IsTextTab(Sender) = True Or
            IsAvatarTab(Sender) = True Or
            IsSnippitTab(Sender) = True Or
            IsDeviceTab(Sender) = True Or
            IsUpdateTab(Sender) = True Or
            IsTestTab(Sender) = True Or
            IsPluginTab(Sender) = True Or
             IsScriptTab(Sender) = True Then

                Return CType(Sender.SelectedTab.Controls.Item("Body"), SpydazWebTextBox)
            Else
                Try
                    Return CType(Sender.SelectedTab.Controls.Item("Body"), SpydazWebTextBox)
                Catch ex As Exception

                End Try
            End If
            Return Nothing
        End Function

        Public Shared Function GetCurrentTabInteliTextBox(ByRef Sender As TabControl) As SpydazWebTextBox
            If IsTextTab(Sender) = True Or
            IsAvatarTab(Sender) = True Or
            IsSnippitTab(Sender) = True Or
            IsDeviceTab(Sender) = True Or
            IsUpdateTab(Sender) = True Or
            IsTestTab(Sender) = True Or
            IsPluginTab(Sender) = True Or
             IsScriptTab(Sender) = True Then

                Return CType(Sender.SelectedTab.Controls.Item("Body"), SpydazWebTextBox)
            Else
                Try
                    Return CType(Sender.SelectedTab.Controls.Item("Body"), SpydazWebTextBox)
                Catch ex As Exception

                End Try
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Used to insert Scripts etc to the TextBox
        ''' </summary>
        ''' <param name="Text"></param>
        Public Shared Sub InsertText(ByRef TabDocumentBrowser As TabControl, ByRef Text As String)
            If IsTextTab(TabDocumentBrowser) = True Or
            IsAvatarTab(TabDocumentBrowser) = True Or
            IsDeviceTab(TabDocumentBrowser) = True Or
            IsUpdateTab(TabDocumentBrowser) = True Or
            IsTestTab(TabDocumentBrowser) = True Or
               IsSnippitTab(TabDocumentBrowser) = True Or
                IsScriptTab(TabDocumentBrowser) = True Or
            IsPluginTab(TabDocumentBrowser) = True Then

                Dim insertText = vbNewLine & Text & vbNewLine
                Dim RichTextBoxCodeWindow As RichTextBox = GetCurrentTabRichTextBox(TabDocumentBrowser)
                Dim insertPos As Integer = RichTextBoxCodeWindow.SelectionStart
                RichTextBoxCodeWindow.Text = RichTextBoxCodeWindow.Text.Insert(insertPos, insertText)
                RichTextBoxCodeWindow.SelectionStart = insertPos + insertText.Length
                'Previously for Highlighting Code
                RichTextBoxCodeWindow.HighlightInternalSyntax
            Else
                If GetCurrentTabInteliTextBox(TabDocumentBrowser) IsNot Nothing Then
                    Dim insertText = vbNewLine & Text & vbNewLine
                    Dim RichTextBoxCodeWindow As SpydazWebTextBox = GetCurrentTabInteliTextBox(TabDocumentBrowser)
                    Dim insertPos As Integer = RichTextBoxCodeWindow.SelectionStart
                    RichTextBoxCodeWindow.Text = RichTextBoxCodeWindow.Text.Insert(insertPos, insertText)
                    RichTextBoxCodeWindow.SelectionStart = insertPos + insertText.Length
                Else
                    If GetCurrentTabSyntaxTextBox(TabDocumentBrowser) IsNot Nothing Then
                        Dim insertText = vbNewLine & Text & vbNewLine
                        Dim RichTextBoxCodeWindow As SpydazWebTextBox = GetCurrentTabSyntaxTextBox(TabDocumentBrowser)
                        Dim insertPos As Integer = RichTextBoxCodeWindow.SelectionStart
                        RichTextBoxCodeWindow.Text = RichTextBoxCodeWindow.Text.Insert(insertPos, insertText)
                        RichTextBoxCodeWindow.SelectionStart = insertPos + insertText.Length
                    Else

                    End If
                End If
            End If

        End Sub

        ''' <summary>
        ''' Used to insert Scripts etc to the TextBox
        ''' </summary>
        ''' <param name="Text"></param>
        Public Shared Sub InsertSyntaxText(ByRef TabDocumentBrowser As TabControl, ByRef Text As String)
            If IsTextTab(TabDocumentBrowser) = True Or
            IsAvatarTab(TabDocumentBrowser) = True Or
            IsDeviceTab(TabDocumentBrowser) = True Or
            IsUpdateTab(TabDocumentBrowser) = True Or
            IsTestTab(TabDocumentBrowser) = True Or
               IsSnippitTab(TabDocumentBrowser) = True Or
                IsScriptTab(TabDocumentBrowser) = True Or
            IsPluginTab(TabDocumentBrowser) = True Then

                Dim insertText = vbNewLine & Text & vbNewLine
                Dim RichTextBoxCodeWindow As SpydazWebTextBox = GetCurrentTabSyntaxTextBox(TabDocumentBrowser)
                Dim insertPos As Integer = RichTextBoxCodeWindow.SelectionStart
                RichTextBoxCodeWindow.Text = RichTextBoxCodeWindow.Text.Insert(insertPos, insertText)
                RichTextBoxCodeWindow.SelectionStart = insertPos + insertText.Length
                'Previously for Highlighting Code
            Else
                If GetCurrentTabSyntaxTextBox(TabDocumentBrowser) IsNot Nothing Then
                    Dim insertText = vbNewLine & Text & vbNewLine
                    Dim RichTextBoxCodeWindow As SpydazWebTextBox = GetCurrentTabSyntaxTextBox(TabDocumentBrowser)
                    Dim insertPos As Integer = RichTextBoxCodeWindow.SelectionStart
                    RichTextBoxCodeWindow.Text = RichTextBoxCodeWindow.Text.Insert(insertPos, insertText)
                    RichTextBoxCodeWindow.SelectionStart = insertPos + insertText.Length
                End If
            End If

        End Sub

        Public Shared Sub SaveCurrentTab(ByRef Sender As TabControl, ByRef Filename As String)
            If IsTextTab(Sender) = True Or
            IsAvatarTab(Sender) = True Or
            IsDeviceTab(Sender) = True Or
            IsSnippitTab(Sender) = True Or
            IsUpdateTab(Sender) = True Or
            IsTestTab(Sender) = True Or
            IsPluginTab(Sender) = True Or
             IsScriptTab(Sender) = True Then
                GetCurrentTabRichTextBox(Sender).SaveFile(Filename, RichTextBoxStreamType.RichText)
            Else
            End If
        End Sub

        Public Shared Sub SaveCurrentSyntaxTextTab(ByRef Sender As TabControl, ByRef Filename As String)
            If IsTextTab(Sender) = True Or
            IsAvatarTab(Sender) = True Or
            IsDeviceTab(Sender) = True Or
            IsSnippitTab(Sender) = True Or
            IsUpdateTab(Sender) = True Or
            IsTestTab(Sender) = True Or
            IsPluginTab(Sender) = True Or
             IsScriptTab(Sender) = True Then
                GetCurrentTabSyntaxTextBox(Sender).SaveFile(Filename, RichTextBoxStreamType.RichText)
            Else
            End If
        End Sub

        Public Shared Sub OpenNewSyntaxDocumentTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, Filename As String)
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ForeColor = Color.MidnightBlue

            Dim NewPage As New TabPage()
            Tabcount += 1

            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.LoadFile(Filename, RichTextBoxStreamType.PlainText)
            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
        End Sub

        Public Shared Sub OpenNewDocumentTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, Filename As String)
            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ForeColor = Color.MidnightBlue

            Dim NewPage As New TabPage()
            Tabcount += 1

            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.LoadFile(Filename, RichTextBoxStreamType.PlainText)
            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
        End Sub

        Public Shared Sub OpenSyntaxTextCurrentTab(ByRef Sender As TabControl, ByRef Filename As String)
            GetCurrentTabSyntaxTextBox(Sender).LoadFile(Filename, RichTextBoxStreamType.PlainText)
            Speaker.Speak("Loading file" & Filename)
        End Sub

        Public Shared Sub OpenTextCurrentTab(ByRef Sender As TabControl, ByRef Filename As String)
            GetCurrentTabRichTextBox(Sender).LoadFile(Filename, RichTextBoxStreamType.PlainText)
            Speaker.Speak("Loading file" & Filename)
        End Sub

        Public Shared Sub AddEmptyTextTab(ByRef Sender As TabControl)
            Speaker.Speak("Opening New Text Tab")
            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Body.ForeColor = Color.MidnightBlue
            Dim DocumentText As String = "Document " & Sender.TabPages.Count + 1
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            ' Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddEmptySyntaxTextTab(ByRef Sender As TabControl)
            Speaker.Speak("Opening New Tab")
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Body.ForeColor = Color.MidnightBlue
            Dim DocumentText As String = "Document " & Sender.TabPages.Count + 1
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            ' Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddEmptySyntaxTextTab(ByRef Sender As TabControl, ByRef Syntax As List(Of String))
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Body.ForeColor = Color.MidnightBlue
            Dim DocumentText As String = "Document " & Sender.TabPages.Count + 1
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            ' Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddEmptyInteliTextTab(ByRef Sender As TabControl, ByRef Syntax As List(Of String))
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Body.ForeColor = Color.MidnightBlue
            Dim DocumentText As String = "Document " & Sender.TabPages.Count + 1
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            ' Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Function GetCurrentTabBrowserControl(ByRef Sender As TabControl) As WebBrowser
            If IsWebTab(Sender) = True Then
                Return CType(Sender.SelectedTab.Controls.Item("Body"), WebBrowser)
            Else
            End If
            Return Nothing
        End Function

        Public Shared Sub AddWebTab(ByRef Sender As TabControl, ByRef Tabcount As Integer)
            Dim Body As New WebBrowser
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Tabcount += 1
            Body.ScriptErrorsSuppressed = True
            Body.GoHome()
            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.Tag = TAB()
            '  Body.ContextMenuStrip = rcMenu
            NewPage.Text = "Browser"

            Sender.TabPages.Add(NewPage)
            Sender.SelectedTab = NewPage

        End Sub

        Public Shared Sub LoadCross(ByRef sender As TabControl)
            'Set the Mode of Drawing as Owner Drawn
            sender.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
            'Add the Handler to draw the Image on Tab Pages
            AddHandler sender.DrawItem, AddressOf TabWebBrowsers_DrawItem
        End Sub

        Public Shared Sub TabWebBrowsers_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs)
            Dim _imgHitArea As Point = New Point(13, 2)
            Dim _imageLocation As Point = New Point(15, 5)
            Try
                'Close Image to draw
                Dim img As Image = New Bitmap(My.Resources.close)
                Dim r As Rectangle = e.Bounds

                r = sender.GetTabRect(e.Index)
                r.Offset(2, 2)
                Dim TitleBrush As Brush = New SolidBrush(Color.Black)
                Dim f As Font = sender.Font
                Dim title As String = sender.TabPages(e.Index).Text
                e.Graphics.DrawString(title, f, TitleBrush, New PointF(r.X, r.Y))
                e.Graphics.DrawImage(img, New Point(r.X + (sender.GetTabRect(e.Index).Width - _imageLocation.X), _imageLocation.Y))
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        Public Shared Sub TabWebBrowsers_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
            Dim tc As TabControl = CType(sender, TabControl)
            Dim p As Point = e.Location
            Dim _tabWidth As Integer
            Dim _imgHitArea As Point = New Point(13, 2)
            Dim _imageLocation As Point = New Point(15, 5)
            _tabWidth = sender.GetTabRect(tc.SelectedIndex).Width - (_imgHitArea.X)
            Dim r As Rectangle = sender.GetTabRect(tc.SelectedIndex)
            r.Offset(_tabWidth, _imgHitArea.Y)
            r.Width = 16
            r.Height = 16
            If r.Contains(p) Then
                Dim TabP As TabPage = DirectCast(tc.TabPages.Item(tc.SelectedIndex), TabPage)
                tc.TabPages.Remove(TabP)
            End If
        End Sub

        Public Shared Sub AddWebTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Url As String)
            Dim Body As New WebBrowser
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ScriptErrorsSuppressed = True
            Dim NewPage As New TabPage()
            Tabcount += 1
            Body.Navigate(Url)
            ' Body.ContextMenuStrip = rcMenu
            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.Tag = TAB()
            NewPage.Text = "Browser"
            'Sender.SelectedTab.Tag = Url
            Sender.TabPages.Add(NewPage)
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddWebSourceTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Source As String)

            Dim Body As New System.Windows.Forms.WebBrowser
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ScriptErrorsSuppressed = True
            Dim NewPage As New TabPage()
            Tabcount += 1

            Body.DocumentText = Source
            Body.Refresh()
            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.Tag = TAB()
            NewPage.Text = "Browser"

            Sender.TabPages.Add(NewPage)
            Sender.SelectedTab = NewPage
            ' Body.ContextMenuStrip = rcMenu
        End Sub

        Public Shared Sub AddPopulatedTextTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Text As String, ByRef Title As String)
            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Speaker.Speak(Title)
            Body.Text = Text
            Dim NewPage As New TabPage()

            Body.ForeColor = Color.SaddleBrown
            Body.HighlightInternalSyntax()
            Dim DocumentText As String = Title & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = Title
            'Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddPopulatedInteliTextTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Text As String, ByRef Title As String, Syntax As List(Of String))
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill

            Body.Text = Text
            Dim NewPage As New TabPage()

            Body.ForeColor = Color.SaddleBrown

            Dim DocumentText As String = Title & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Speaker.Speak(Title)
            Sender.TabPages.Add(NewPage)
            NewPage.Text = Title
            'Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddPopulatedSyntaxTextTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Text As String, ByRef Title As String)
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Speaker.Speak(Title)
            Body.Text = Text
            Dim NewPage As New TabPage()

            Body.ForeColor = Color.SaddleBrown
            Body.HighlightInternalSyntax()
            Dim DocumentText As String = Title & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = Title
            'Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddPopulatedSyntaxTextTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Text As String, ByRef Title As String, ByRef Syntax As List(Of String))
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill

            Body.Text = Text
            Dim NewPage As New TabPage()
            Speaker.Speak(Title)

            Body.ForeColor = Color.SaddleBrown
            Body.HighlightInternalSyntax()
            Dim DocumentText As String = Title & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = Title
            'Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

#End Region

#Region "Files"

        Public Shared Function LoadFile(ByRef Sender As TabControl, ByRef Filter As String) As String
            Dim Scriptfile As String = ""
            Dim txt As String = ""
            Speaker.Speak("Select file")
            Select Case Filter
                Case "Json"
                    Filter = "Json files (*.Json)|*.Json"
                Case = "Html"
                    Filter = "Html Source files (*.htm)|*.htm"
                Case "Aib"
                    Filter = "BrainScript files (*.Aib)|*.Aib"
                Case "Text"
                    Filter = "Text files (*.txt)|*.txt"
                Case "VB"
                    Filter = "VB files (*.vb)|*.vb"
                Case "C#"
                    Filter = "C# files (*.cs)|*.cs"
                Case "All"
                    Filter = "All files (*.*)|*.*"
                Case "CSV"
                    Filter = "comma seperated data files (*.csv)|*.csv"
                Case "XML"
                    Filter = "X Markup Language (*.xml)|*.xml"
                Case "JavaScript"
                    Filter = "Javascript Language (*.js)|*.js"
            End Select
            Dim Ofile As New OpenFileDialog
            With Ofile
                .Filter = Filter
                If (.ShowDialog() = DialogResult.OK) Then
                    Scriptfile = .FileName
                End If
            End With
            Try
                txt = My.Computer.FileSystem.ReadAllText(Scriptfile)

                AddSyntaxTextTab(Sender, Sender.TabPages.Count, txt, Scriptfile)
            Catch ex As Exception
                '   Speaker.Speak("Select file")
                '  MsgBox(ex.ToString,, "Error")
            End Try
            Return txt
        End Function

        Public Shared Sub SaveScript(ByRef Script As String)
            Try
                Dim ScriptFile As String = ""
                Dim S As New SaveFileDialog
                With S
                    .Filter = "All Files (*.*)|*.*"
                    If (.ShowDialog() = DialogResult.OK) Then
                        ScriptFile = .FileName
                    End If
                End With
                My.Computer.FileSystem.WriteAllText(ScriptFile, Script, True)
            Catch ex As Exception
                MsgBox(ex.ToString,, "error")
            End Try
        End Sub

        Public Shared Sub SaveFile(ByRef sender As TabControl)
            If IsTextTab(sender) = True Or
    IsAvatarTab(sender) = True Or
    IsDeviceTab(sender) = True Or
    IsSnippitTab(sender) = True Or
    IsUpdateTab(sender) = True Or
    IsTestTab(sender) = True Or
    IsPluginTab(sender) = True Or
     IsScriptTab(sender) = True Then
                SaveScript(GetCurrentTabRichTextBox(sender).Text)
                Speaker.Speak("Saving")
            Else
            End If
        End Sub

#End Region

#Region "Web-Tab Functions"

        Public Shared Sub Navigate(ByRef sender As TabControl, ByRef Text As String)
            AddWebTab(sender, sender.TabCount, Text)
        End Sub

        Public Shared Sub GetHtmlSource(ByRef sender As TabControl)
            If IsWebTab(sender) = True Then
                '   If Speech_on = True Then Speaker.SpeakAsync("Getting Source code")
                Dim SwBrowser = GetCurrentTabBrowserControl(sender)
                Dim _source = SwBrowser.DocumentText
                'AddPopulatedTextTab(sender, sender.TabPages.Count, _source, "TEXT")
                AddTextTab(sender, sender.TabPages.Count, _source, "SCRIPT")
            Else
            End If
        End Sub

        Public Shared Sub OpenHtmlSource(ByRef sender As TabControl)
            If IsWebTab(sender) = True Then
                Dim SwBrowser = GetCurrentTabBrowserControl(sender)
                Dim _source = SwBrowser.DocumentText
                AddPopulatedTextTab(sender, sender.TabCount, _source, "SOURCE")
            Else
            End If
        End Sub

        Public Shared Sub SaveWebpage(ByRef Sender As TabControl)
            If IsWebTab(Sender) = True Then

                Dim SwBrowser = GetCurrentTabBrowserControl(Sender)
                Dim alltext As String = SwBrowser.DocumentText
                Try
                    Dim ScriptFile As String = ""
                    Dim S As New SaveFileDialog
                    With S
                        .Filter = "Enter Filename : (*.Html)|*.Html"
                        If (.ShowDialog() = DialogResult.OK) Then
                            ScriptFile = .FileName
                        End If
                    End With
                    My.Computer.FileSystem.WriteAllText(ScriptFile, alltext, True)
                Catch ex As Exception
                    MsgBox(ex.ToString,, "error")
                End Try
            Else

            End If
        End Sub

        Public Shared Sub SaveSource(ByRef Sender As TabControl)
            Dim alltext As String = ""
            Try
                Try
                    Dim SwBrowser = GetCurrentTabBrowserControl(Sender)
                    alltext = SwBrowser.DocumentText
                Catch ex As Exception
                    Dim SwBrowser = GetCurrentTabRichTextBox(Sender)
                    alltext = SwBrowser.Text
                End Try

                Dim ScriptFile As String = ""
                Dim S As New SaveFileDialog
                With S
                    .Filter = "Enter Filename : (*.txt)|*.txt"
                    If (.ShowDialog() = DialogResult.OK) Then
                        ScriptFile = .FileName
                    End If
                End With
                My.Computer.FileSystem.WriteAllText(ScriptFile, alltext, True)
            Catch ex As Exception

            End Try

        End Sub

        Public Shared Function GetImageFromUrl(ByRef Url As String) As Image
            Using webClient = New WebClient()
                Dim image_stream As New MemoryStream(webClient.DownloadData(Url))
                Return Image.FromStream(image_stream)
            End Using
        End Function

        Public Shared Function GetPicture(ByRef Url As String) As Image
            Using webClient = New WebClient()
                Try
                    Dim image_stream As New MemoryStream(webClient.DownloadData(Url))
                    Return Image.FromStream(image_stream)
                Catch ex As Exception
                    Return Nothing
                End Try

            End Using
        End Function

        Public Shared Function GrabLinks(Sender As TabControl) As List(Of String)
            Dim WebLinks As New List(Of String)
            Dim SwBrowser = GetCurrentTabBrowserControl(Sender)
            If SwBrowser IsNot Nothing Then
                Dim links = SwBrowser.Document.Links
                ' Dim images As HtmlElementCollection = SwBrowser.Document.Images
                For Each link As HtmlElement In links
                    WebLinks.Add(link.GetAttribute("href").ToString)
                    '  frmGrabber.Visible = False
                Next
                'Dim links As HtmlElementCollection = SW_Browser.Document.Links
            End If
            Return WebLinks
        End Function

        Public Shared Function GrabLinks(Sender As WebBrowser) As List(Of String)
            Dim WebLinks As New List(Of String)
            Dim SwBrowser = Sender
            If SwBrowser IsNot Nothing Then
                Dim links = SwBrowser.Document.Links
                ' Dim images As HtmlElementCollection = SwBrowser.Document.Images
                For Each link As HtmlElement In links
                    WebLinks.Add(link.GetAttribute("href").ToString)
                    '  frmGrabber.Visible = False
                Next
                'Dim links As HtmlElementCollection = SW_Browser.Document.Links
            End If
            Return WebLinks
        End Function

        Public Shared Function GrabUrlImages(ByRef sender As WebBrowser) As List(Of String)
            Dim LinkStr As New List(Of String)
            Dim SwBrowser = sender
            If SwBrowser IsNot Nothing Then
                ' Dim images As HtmlElementCollection = SwBrowser.Document.
                Dim links As HtmlElementCollection = SwBrowser.Document.Links
                For Each link As HtmlElement In links
                    LinkStr.Add(link.GetAttribute("img"))
                Next
            End If

            Return LinkStr
        End Function

        Public Shared Function GrabUrlImages(ByRef sender As TabControl) As List(Of String)
            Dim LinkStr As New List(Of String)
            Dim SwBrowser = GetCurrentTabBrowserControl(sender)
            If SwBrowser IsNot Nothing Then
                ' Dim images As HtmlElementCollection = SwBrowser.Document.
                Dim links As HtmlElementCollection = SwBrowser.Document.Links
                For Each link As HtmlElement In links
                    LinkStr.Add(link.GetAttribute("img"))
                Next
            End If

            Return LinkStr
        End Function

        Public Shared Function GetText(sender As TabControl) As String
            If IsWebTab(sender) = True Then
                Dim SW_Browser = GetCurrentTabBrowserControl(sender)
                Return SW_Browser.Document.Body.InnerText
            End If
            Return Nothing
        End Function

        Public Shared Function GetText(sender As WebBrowser) As String
            Dim SW_Browser = sender
            Return SW_Browser.Document.Body.InnerText
        End Function

        Public Shared Function GrabSource(ByRef sender As TabControl) As String
            Dim _source As String = ""

            If IsWebTab(sender) = True Then
                If Speech_on = True Then Speaker.SpeakAsync("Getting Source code")
                Dim SwBrowser = GetCurrentTabBrowserControl(sender)
                Return SwBrowser.DocumentText
            Else
            End If
            Return _source
        End Function

        Public Shared Function DownloadWebPage(ByRef Url As String) As String
            Dim request As WebRequest = WebRequest.Create(Url)
            Dim html As String = ""
            Using response As WebResponse = request.GetResponse()
                Using reader As New StreamReader(response.GetResponseStream())
                    html = reader.ReadToEnd()
                End Using
            End Using
            Return html
        End Function

        Public Shared Sub DownloadWebPage(ByRef Url As String, Reciever As TabControl)
            Dim request As WebRequest = WebRequest.Create(Url)
            Dim html As String = ""
            Using response As WebResponse = request.GetResponse()
                Using reader As New StreamReader(response.GetResponseStream())
                    html = reader.ReadToEnd()
                End Using
            End Using
            AddPopulatedTextTab(Reciever, Reciever.TabPages.Count, html, "SCRIPT")

        End Sub

        Public Shared Function DownloadSource(ByRef Url As String) As String
            Dim LinkStr As String = ""
            Dim SwBrowser As New WebBrowser
            SwBrowser.ScriptErrorsSuppressed = True
            SwBrowser.Navigate(Url)
            If SwBrowser.DocumentText IsNot Nothing Then
                LinkStr = SwBrowser.DocumentText
            End If
            Return LinkStr
        End Function

        Public Shared Function OpenPicfolder(ByRef Path As String) As List(Of String)
            Dim JpgEntries As String() = Directory.GetFiles(Path, "*.jpg")
            Dim PngEntries As String() = Directory.GetFiles(Path, "*.png")
            Dim BmpEntries As String() = Directory.GetFiles(Path, "*.bmp")
            Dim str As New List(Of String)
            For Each item In JpgEntries
                str.Add(item)
            Next
            For Each item In BmpEntries
                str.Add(item)
            Next
            For Each item In PngEntries
                str.Add(item)
            Next
            Return str
        End Function

#End Region

#Region "Browser Events"

        ' Hides script errors without hiding other dialog boxes.
        Private Shared Sub SuppressScriptErrorsOnly(ByVal browser As WebBrowser)

            ' Ensure that ScriptErrorsSuppressed is set to false.
            browser.ScriptErrorsSuppressed = True

            ' Handle DocumentCompleted to gain access to the Document object.
            AddHandler browser.DocumentCompleted,
        AddressOf browser_DocumentCompleted

        End Sub

        Private Shared Sub browser_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)

            AddHandler CType(sender, WebBrowser).Document.Window.Error,
        AddressOf Window_Error

        End Sub

        Private Shared Sub Window_Error(ByVal sender As Object, ByVal e As HtmlElementErrorEventArgs)

            ' Ignore the error and suppress the error dialog box.
            e.Handled = True

        End Sub

#End Region

#End Region

#Region "Cut/CopyPaste"

        ''' <summary>
        ''' Used to insert Scripts etc to the TextBox
        ''' </summary>
        ''' <param name="Text"></param>
        Public Sub InsertTextTab(ByRef Sender As TabControl, ByRef Text As String)
            AddPopulatedSyntaxTextTab(Sender, Sender.TabPages.Count, Text, "Snippit")
        End Sub

        Public Shared Sub Undo(ByRef Sender As TabControl)
            If GetCurrentTabRichTextBox(Sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(Sender).CanUndo = True Then
                        GetCurrentTabRichTextBox(Sender).Undo()
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(Sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(Sender).CanUndo = True Then
                            GetCurrentTabSyntaxTextBox(Sender).Undo()
                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                Else
                    If GetCurrentTabInteliTextBox(Sender) IsNot Nothing Then
                        Try
                            If GetCurrentTabInteliTextBox(Sender).CanUndo = True Then
                                GetCurrentTabInteliTextBox(Sender).Undo()
                            End If
                        Catch ex As Exception
                            '   MsgBox(ex.Message)
                        End Try
                    End If
                End If
            End If

        End Sub

        Public Shared Sub Redo(ByRef Sender As TabControl)
            If GetCurrentTabRichTextBox(Sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(Sender).CanUndo = True Then
                        GetCurrentTabRichTextBox(Sender).Redo()
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(Sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(Sender).CanUndo = True Then
                            GetCurrentTabSyntaxTextBox(Sender).Redo()
                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                Else

                End If
            End If

        End Sub

        Public Shared Sub Cut(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        GetCurrentTabRichTextBox(sender).Cut()
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            GetCurrentTabSyntaxTextBox(sender).Cut()
                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                Else
                    If GetCurrentTabInteliTextBox(sender) IsNot Nothing Then
                        Try
                            If GetCurrentTabInteliTextBox(sender).SelectedText IsNot Nothing Then
                                GetCurrentTabInteliTextBox(sender).Cut()
                            End If
                        Catch ex As Exception
                            '   MsgBox(ex.Message)
                        End Try
                    End If
                End If
            End If

        End Sub

        Public Shared Sub Copy(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        GetCurrentTabRichTextBox(sender).Copy()
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            GetCurrentTabSyntaxTextBox(sender).Copy()
                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                Else
                    If GetCurrentTabInteliTextBox(sender) IsNot Nothing Then
                        Try
                            If GetCurrentTabInteliTextBox(sender).SelectedText IsNot Nothing Then
                                GetCurrentTabInteliTextBox(sender).Copy()
                            End If
                        Catch ex As Exception
                            '   MsgBox(ex.Message)
                        End Try
                    End If
                End If

            End If

            'If PLUGIN_RESPONSE_LIST.SelectedItem IsNot Nothing Then
            '    PLUGIN_RESPONSE_LIST.SelectedItem.Copy()
            'End If
        End Sub

        Public Shared Sub Paste(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    GetCurrentTabRichTextBox(sender).Paste()
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        GetCurrentTabSyntaxTextBox(sender).Paste()
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                Else
                    If GetCurrentTabInteliTextBox(sender) IsNot Nothing Then
                        Try
                            If GetCurrentTabInteliTextBox(sender).SelectedText IsNot Nothing Then
                                GetCurrentTabInteliTextBox(sender).Paste()
                            End If
                        Catch ex As Exception
                            '   MsgBox(ex.Message)
                        End Try
                    End If
                End If
            End If

        End Sub

        Public Shared Sub SelectAll(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    GetCurrentTabRichTextBox(sender).SelectAll()
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        GetCurrentTabSyntaxTextBox(sender).SelectAll()
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                Else
                    If GetCurrentTabInteliTextBox(sender) IsNot Nothing Then
                        Try
                            If GetCurrentTabInteliTextBox(sender) IsNot Nothing Then
                                GetCurrentTabInteliTextBox(sender).SelectAll()
                            End If
                        Catch ex As Exception
                            '   MsgBox(ex.Message)
                        End Try
                    End If
                End If
            End If

        End Sub

#End Region

        Public Shared Sub InsertPicture(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                'Use the code 'InsertPicture()' to show the dialog and insert a picture
                Try
                    Dim GetPicture As New OpenFileDialog
                    GetPicture.Filter = "PNGs (*.png), Bitmaps (*.bmp), GIFs (*.gif), JPEGs (*.jpg)|*.bmp;*.gif;*.jpg;*.png|PNGs (*.png)|*.png|Bitmaps (*.bmp)|*.bmp|GIFs (*.gif)|*.gif|JPEGs (*.jpg)|*.jpg"
                    GetPicture.FilterIndex = 1
                    GetPicture.InitialDirectory = "C:\"
                    If GetPicture.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Dim SelectedPicture As String = GetPicture.FileName
                        Dim Picture As Bitmap = New Bitmap(SelectedPicture)
                        Dim cboard As Object = Clipboard.GetData(System.Windows.Forms.DataFormats.Text)
                        Clipboard.SetImage(Picture)
                        Dim PictureFormat As DataFormats.Format = DataFormats.GetFormat(DataFormats.Bitmap)
                        If GetCurrentTabRichTextBox(sender).CanPaste(PictureFormat) Then
                            GetCurrentTabRichTextBox(sender).Paste(PictureFormat)
                        End If
                        Clipboard.Clear()
                        Clipboard.SetText(cboard)
                    End If
                Catch ex As Exception
                End Try
            Else
            End If
        End Sub

#Region "BOLD/UPPERCASE/LOWERCASE/ITALICS/readaloud"

#Region "Text"

        Public Shared Sub Bold(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        Dim CurrentFont = GetCurrentTabRichTextBox(sender).Font.FontFamily
                        Dim CurrentSize = GetCurrentTabRichTextBox(sender).Font.Size
                        'Sets the selected text to bold
                        If GetCurrentTabRichTextBox(sender).SelectionFont.Style = FontStyle.Bold Then
                            GetCurrentTabRichTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Regular)
                        Else
                            GetCurrentTabRichTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Bold)
                        End If

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            Dim CurrentFont = GetCurrentTabSyntaxTextBox(sender).Font.FontFamily
                            Dim CurrentSize = GetCurrentTabSyntaxTextBox(sender).Font.Size
                            'Sets the selected text to bold
                            If GetCurrentTabSyntaxTextBox(sender).SelectionFont.Style = FontStyle.Bold Then
                                GetCurrentTabSyntaxTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Regular)
                            Else
                                GetCurrentTabSyntaxTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Bold)
                            End If

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub Italic(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        Dim CurrentFont = GetCurrentTabRichTextBox(sender).Font.FontFamily
                        Dim CurrentSize = GetCurrentTabRichTextBox(sender).Font.Size
                        'Sets the selected text to bold
                        If GetCurrentTabRichTextBox(sender).SelectionFont.Style = FontStyle.Italic Then
                            GetCurrentTabRichTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Regular)
                        Else
                            GetCurrentTabRichTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Italic)
                        End If

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            Dim CurrentFont = GetCurrentTabSyntaxTextBox(sender).Font.FontFamily
                            Dim CurrentSize = GetCurrentTabSyntaxTextBox(sender).Font.Size
                            'Sets the selected text to bold
                            If GetCurrentTabSyntaxTextBox(sender).SelectionFont.Style = FontStyle.Italic Then
                                GetCurrentTabSyntaxTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Regular)
                            Else
                                GetCurrentTabSyntaxTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Italic)
                            End If

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub Strikeout(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        Dim CurrentFont = GetCurrentTabRichTextBox(sender).Font.FontFamily
                        Dim CurrentSize = GetCurrentTabRichTextBox(sender).Font.Size
                        'Sets the selected text to bold
                        If GetCurrentTabRichTextBox(sender).SelectionFont.Style = FontStyle.Strikeout Then
                            GetCurrentTabRichTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Regular)
                        Else
                            GetCurrentTabRichTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Strikeout)
                        End If

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            Dim CurrentFont = GetCurrentTabSyntaxTextBox(sender).Font.FontFamily
                            Dim CurrentSize = GetCurrentTabSyntaxTextBox(sender).Font.Size
                            'Sets the selected text to bold
                            If GetCurrentTabSyntaxTextBox(sender).SelectionFont.Style = FontStyle.Strikeout Then
                                GetCurrentTabSyntaxTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Regular)
                            Else
                                GetCurrentTabSyntaxTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Strikeout)
                            End If

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If

            End If
        End Sub

        Public Shared Sub Underline(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        Dim CurrentFont = GetCurrentTabRichTextBox(sender).Font.FontFamily
                        Dim CurrentSize = GetCurrentTabRichTextBox(sender).Font.Size
                        'Sets the selected text to bold
                        If GetCurrentTabRichTextBox(sender).SelectionFont.Style = FontStyle.Underline Then
                            GetCurrentTabRichTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Regular)
                        Else
                            GetCurrentTabRichTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Underline)
                        End If

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            Dim CurrentFont = GetCurrentTabSyntaxTextBox(sender).Font.FontFamily
                            Dim CurrentSize = GetCurrentTabSyntaxTextBox(sender).Font.Size
                            'Sets the selected text to bold
                            If GetCurrentTabSyntaxTextBox(sender).SelectionFont.Style = FontStyle.Underline Then
                                GetCurrentTabSyntaxTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Regular)
                            Else
                                GetCurrentTabSyntaxTextBox(sender).SelectionFont = New Font(CurrentFont, CurrentSize, FontStyle.Underline)
                            End If

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub ReadAloud(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        If GetCurrentTabRichTextBox(sender).SelectedText = "" Then
                            'Speaks the entire document (if no text is selected)
                            Dim SAPI2
                            SAPI2 = CreateObject("SAPI.spvoice")
                            SAPI2.Speak(GetCurrentTabRichTextBox(sender).Text)
                            SAPI2.Speak("Status : " & "Finished speaking document")
                        Else
                            'Speaks the selected text
                            Dim SAPI
                            SAPI = CreateObject("SAPI.spvoice")
                            SAPI.Speak(GetCurrentTabRichTextBox(sender).SelectedText)
                            SAPI.Speak("Status : " & "Finished speaking document")
                        End If
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                            If GetCurrentTabRichTextBox(sender).SelectedText = "" Then
                                'Speaks the entire document (if no text is selected)
                                Dim SAPI2
                                SAPI2 = CreateObject("SAPI.spvoice")
                                SAPI2.Speak(GetCurrentTabRichTextBox(sender).Text)
                                SAPI2.Speak("Status : " & "Finished speaking document")
                            Else
                                'Speaks the selected text
                                Dim SAPI
                                SAPI = CreateObject("SAPI.spvoice")
                                SAPI.Speak(GetCurrentTabRichTextBox(sender).SelectedText)
                                SAPI.Speak("Status : " & "Finished speaking document")
                            End If
                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub ReadAloud(ByRef Text As String)
            'Speaks the selected text
            Dim SAPI
            SAPI = CreateObject("SAPI.spvoice")
            SAPI.Speak(Text)
            SAPI.Speak("Status : " & "Finished speaking document")
        End Sub

        Public Shared Sub ReadAloud(ByRef sender As TextBox)
            If sender IsNot Nothing Then
                Try
                    If sender.SelectedText IsNot Nothing Then
                        If sender.SelectedText = "" Then
                            'Speaks the entire document (if no text is selected)
                            Dim SAPI2
                            SAPI2 = CreateObject("SAPI.spvoice")
                            SAPI2.Speak(sender.Text)
                            SAPI2.Speak("Status : " & "Finished speaking document")
                        Else
                            'Speaks the selected text
                            Dim SAPI
                            SAPI = CreateObject("SAPI.spvoice")
                            SAPI.Speak(sender.SelectedText)
                            SAPI.Speak("Status : " & "Finished speaking document")
                        End If
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else

            End If
        End Sub

        Public Shared Sub ReadAloud(ByRef sender As RichTextBox)
            If sender IsNot Nothing Then
                Try
                    If sender.SelectedText IsNot Nothing Then
                        If sender.SelectedText = "" Then
                            'Speaks the entire document (if no text is selected)
                            Dim SAPI2
                            SAPI2 = CreateObject("SAPI.spvoice")
                            SAPI2.Speak(sender.Text)
                            SAPI2.Speak("Status : " & "Finished speaking document")
                        Else
                            'Speaks the selected text
                            Dim SAPI
                            SAPI = CreateObject("SAPI.spvoice")
                            SAPI.Speak(sender.SelectedText)
                            SAPI.Speak("Status : " & "Finished speaking document")
                        End If
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else

            End If
        End Sub

        Public Async Sub ReadAloudAsync(ByVal sender As RichTextBox)
            Await Task.Run(Sub() ReadAloud(sender))
        End Sub

        Public Async Sub ReadAloudAsync(ByVal sender As TextBox)
            Await Task.Run(Sub() ReadAloud(sender))
        End Sub

        Public Async Sub ReadAloudAsync(ByVal sender As String)
            Await Task.Run(Sub() ReadAloud(sender))
        End Sub

        Public Async Sub ReadAloudAsync(ByVal sender As TabControl)
            Await Task.Run(Sub() ReadAloud(sender))
        End Sub

        Public Shared Sub ToLower(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then

                        GetCurrentTabRichTextBox(sender).SelectedText.ToLower()
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            End If
        End Sub

        Public Shared Sub ToUpper(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then

                        GetCurrentTabRichTextBox(sender).SelectedText.ToUpper()
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then

                            GetCurrentTabSyntaxTextBox(sender).SelectedText.ToUpper()
                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

#End Region

        Public Shared Sub Subscript(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        'Subscript
                        GetCurrentTabRichTextBox(sender).SelectionCharOffset = -10
                        GetCurrentTabRichTextBox(sender).SelectedText = GetCurrentTabRichTextBox(sender).SelectedText
                        GetCurrentTabRichTextBox(sender).SelectionCharOffset = 0
                        GetCurrentTabRichTextBox(sender).SelectionCharOffset = 10
                        GetCurrentTabRichTextBox(sender).SelectionCharOffset = 0

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            'Subscript
                            GetCurrentTabSyntaxTextBox(sender).SelectionCharOffset = -10
                            GetCurrentTabSyntaxTextBox(sender).SelectedText = GetCurrentTabRichTextBox(sender).SelectedText
                            GetCurrentTabSyntaxTextBox(sender).SelectionCharOffset = 0
                            GetCurrentTabSyntaxTextBox(sender).SelectionCharOffset = 10
                            GetCurrentTabSyntaxTextBox(sender).SelectionCharOffset = 0

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub Superscript(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then
                        'Subscript
                        'Superscript
                        GetCurrentTabRichTextBox(sender).SelectionCharOffset = 10
                        GetCurrentTabRichTextBox(sender).SelectedText = GetCurrentTabSyntaxTextBox(sender).SelectedText
                        GetCurrentTabRichTextBox(sender).SelectionCharOffset = 0
                        GetCurrentTabRichTextBox(sender).SelectionCharOffset = -10
                        GetCurrentTabRichTextBox(sender).SelectionCharOffset = 0

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            'Subscript
                            'Superscript
                            GetCurrentTabSyntaxTextBox(sender).SelectionCharOffset = 10
                            GetCurrentTabSyntaxTextBox(sender).SelectedText = GetCurrentTabSyntaxTextBox(sender).SelectedText
                            GetCurrentTabSyntaxTextBox(sender).SelectionCharOffset = 0
                            GetCurrentTabSyntaxTextBox(sender).SelectionCharOffset = -10
                            GetCurrentTabSyntaxTextBox(sender).SelectionCharOffset = 0

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

#Region "Alignment"

        Public Shared Sub LeftAlignment(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then

                        GetCurrentTabRichTextBox(sender).SelectionAlignment = HorizontalAlignment.Left

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then

                            GetCurrentTabSyntaxTextBox(sender).SelectionAlignment = HorizontalAlignment.Left

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub RightAlignment(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then

                        GetCurrentTabRichTextBox(sender).SelectionAlignment = HorizontalAlignment.Right

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then

                            GetCurrentTabSyntaxTextBox(sender).SelectionAlignment = HorizontalAlignment.Right

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub CenterAlignment(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then

                        GetCurrentTabRichTextBox(sender).SelectionAlignment = HorizontalAlignment.Left

                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then

                            GetCurrentTabSyntaxTextBox(sender).SelectionAlignment = HorizontalAlignment.Left

                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

#End Region

#Region "find"

        Public Shared Sub Find(ByRef sender As TabControl, ByRef Str As String)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)

                Try
                    GetCurrentTabRichTextBox(sender).Find(Str)
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                    '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                    '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                    '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)

                    Try
                        GetCurrentTabSyntaxTextBox(sender).Find(Str)
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub FindMatchCase(ByRef sender As TabControl, ByRef Str As String)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)

                Try

                    GetCurrentTabRichTextBox(sender).Find(Str, RichTextBoxFinds.MatchCase)
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                    '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                    '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                    '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)

                    Try

                        GetCurrentTabSyntaxTextBox(sender).Find(Str, RichTextBoxFinds.MatchCase)
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub FindNoHighlight(ByRef sender As TabControl, ByRef Str As String)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)
                Try
                    GetCurrentTabRichTextBox(sender).Find(Str, RichTextBoxFinds.NoHighlight)
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                    '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                    '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                    '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)
                    Try
                        GetCurrentTabSyntaxTextBox(sender).Find(Str, RichTextBoxFinds.NoHighlight)
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub FindReverse(ByRef sender As TabControl, ByRef Str As String)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)
                Try
                    GetCurrentTabRichTextBox(sender).Find(Str, RichTextBoxFinds.Reverse)
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                    '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                    '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                    '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)
                    Try
                        GetCurrentTabSyntaxTextBox(sender).Find(Str, RichTextBoxFinds.Reverse)
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        Public Shared Sub FindWholeWord(ByRef sender As TabControl, ByRef Str As String)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)
                Try
                    GetCurrentTabRichTextBox(sender).Find(Str, RichTextBoxFinds.Reverse)
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    'OPTIONAL : Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.MatchCase)
                    '  Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.NoHighlight)
                    '    Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.Reverse)
                    '   Or  Form1.RichTextBoxPrintCtrl1.Find(ComboBox1.Text, RichTextBoxFinds.WholeWord)
                    Try
                        GetCurrentTabSyntaxTextBox(sender).Find(Str, RichTextBoxFinds.Reverse)
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

#End Region

        Public Sub SetTextColor(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Dim ColorDialog1 As New ColorDialog
                'Set the color of the selected text
                If ColorDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

                    GetCurrentTabRichTextBox(sender).SelectionColor = ColorDialog1.Color

                End If
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Dim ColorDialog1 As New ColorDialog
                    'Set the color of the selected text
                    If ColorDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then

                        GetCurrentTabSyntaxTextBox(sender).SelectionColor = ColorDialog1.Color

                    End If
                End If

            End If

        End Sub

#End Region

        Public Shared Sub ChangeFontSize(ByRef Sender As TabControl, ByRef FntSize As Integer)
            If GetCurrentTabRichTextBox(Sender) IsNot Nothing Then
                If Sender.TabPages.Count > 0 Then
                    If GetCurrentTabRichTextBox(Sender) IsNot Nothing Then
                        GetCurrentTabRichTextBox(Sender).SelectionFont = New Font(FntSize.ToString, FntSize, GetCurrentTabRichTextBox(Sender).SelectionFont.Style)
                    End If
                End If
            Else
                If Sender.TabPages.Count > 0 Then
                    If GetCurrentTabSyntaxTextBox(Sender) IsNot Nothing Then
                        GetCurrentTabSyntaxTextBox(Sender).SelectionFont = New Font(FntSize.ToString, FntSize, GetCurrentTabSyntaxTextBox(Sender).SelectionFont.Style)
                    End If
                End If
            End If
        End Sub

        Public Shared Sub ChangeFont(ByRef Sender As TabControl, ByRef SenderFontFamily As String)
            Dim ComboFonts As System.Drawing.Font
            If GetCurrentTabRichTextBox(Sender) IsNot Nothing Then
                Try
                    ComboFonts = GetCurrentTabRichTextBox(Sender).SelectionFont
                    GetCurrentTabRichTextBox(Sender).SelectionFont = New System.Drawing.Font(SenderFontFamily, GetCurrentTabRichTextBox(Sender).SelectionFont.Size, GetCurrentTabRichTextBox(Sender).SelectionFont.Style)
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(Sender) IsNot Nothing Then
                    Try
                        ComboFonts = GetCurrentTabSyntaxTextBox(Sender).SelectionFont
                        GetCurrentTabSyntaxTextBox(Sender).SelectionFont = New System.Drawing.Font(SenderFontFamily, GetCurrentTabSyntaxTextBox(Sender).SelectionFont.Size, GetCurrentTabSyntaxTextBox(Sender).SelectionFont.Style)
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
            End If
        End Sub

        ''' <summary>
        ''' Compiles Code from Text TabControl
        ''' </summary>
        ''' <param name="Sender">Tabcontrol Is used to return error Result </param>
        ''' <param name="CodeBlock">Text To be Compiled</param>
        ''' <param name="iClassName">ClassName Required For Executing Code in Memeory</param>
        ''' <param name="iMethodName">Required for executing Code in Memenory</param>
        ''' <param name="Assemblies">Reffereced Assemblies</param>
        ''' <param name="EmbededFiles">Reffereced Embedded Files</param>
        ''' <param name="CompileType">EXE/MEM/DLL</param>
        Public Shared Sub VB_Tab_CodeCompiler(ByRef Sender As TabControl, CodeBlock As String,
                                          Optional ByRef CompileType As String = "DLL",
                                          Optional ByRef iClassName As String = "MainClass",
                                          Optional iMethodName As String = "Execute",
                                          Optional Assemblies As List(Of String) = Nothing,
                                          Optional ByRef EmbededFiles As List(Of String) = Nothing)

            Const Lang As String = "VisualBasic"
            Speaker.Speak("Compiler Started")
            'The CreateProvider method returns a CodeDomProvider instance for the specificed language name.
            'This instance is later used when we have finished specifying the parmameter values.
            Dim Compiler As CodeDom.Compiler.CodeDomProvider = CodeDom.Compiler.CodeDomProvider.CreateProvider(Lang)
            'Optionally, this could be called Parameters.
            'Think of parameters as settings or specific values passed to a method (later passsed to CompileAssemblyFromSource method).
            Dim Settings As New CodeDom.Compiler.CompilerParameters
            'The CompileAssemblyFromSource method returns a CompilerResult that will be stored in this variable.
            Dim Result As CodeDom.Compiler.CompilerResults = Nothing
            Dim ExecuteableName As String = ""
            If EmbededFiles Is Nothing Then EmbededFiles = New List(Of String)

            Try
                Speaker.Speak("Adding Embedded resources")
                'handle Embedded Resources
                If EmbededFiles IsNot Nothing And EmbededFiles.Count > 0 Then
                    For Each str As String In EmbededFiles
                        Settings.EmbeddedResources.Add(str)
                    Next
                End If
            Catch ex As Exception
                Speaker.Speak("There is a problem with a resource")
            End Try

            Try
                Speaker.Speak("Adding Assemblys")
                If Assemblies IsNot Nothing And Assemblies.Count > 0 Then
                    For Each str As String In Assemblies
                        Settings.ReferencedAssemblies.Add(str)
                    Next
                End If
            Catch ex As Exception
                Speaker.Speak("There is a problem with an refference assembly")
            End Try
            'Must Always be added
            Settings.ReferencedAssemblies.Add("System.Windows.Forms.dll")
            Settings.ReferencedAssemblies.Add("Microsoft.VisualBasic.dll")
            Settings.ReferencedAssemblies.Add("System.dll")

#Region "COMPILER"

            Select Case CompileType
                Case "EXE"
                    Speaker.Speak("Compile EXE")
                    Dim S As New SaveFileDialog
                    With S

                        .Filter = "Executable (*.exe)|*.exe"
                        If (.ShowDialog() = DialogResult.OK) Then
                            ExecuteableName = .FileName
                        End If
                    End With
                    'Library type options : /target:library, /target:win, /target:winexe
                    'Generates an executable instead of a class library.
                    'Compiles in memory.
                    Settings.GenerateInMemory = True
                    Settings.GenerateExecutable = True
                    Settings.CompilerOptions = " /target:winexe"
                    'Set the assembly file name / path
                    Settings.OutputAssembly = ExecuteableName
                    'Read the documentation for the result again variable.
                    'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
                    Result = Compiler.CompileAssemblyFromSource(Settings, CodeBlock)
                Case "MEM"
                    Speaker.Speak("Executing")
                    'Compiles in memory.
                    Settings.GenerateInMemory = True
                    'Read the documentation for the result again variable.
                    'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
                    Result = Compiler.CompileAssemblyFromSource(Settings, CodeBlock)

                    'If Assembly error = true then
                    'Return Errors
                    'Else
                    If Result.Errors.HasErrors = True Then

                        Speaker.Speak("There are some issues to be handled")

                        For Each ITEM In Result.Errors
                            AddPopulatedTextTab(Sender, Sender.TabPages.Count, ITEM.ERRORTEXT, "ERROR")
                        Next
                    Else
                        'Creates assembly
                        Dim objAssembly As System.Reflection.Assembly = Result.CompiledAssembly

                        Dim objTheClass As Object = objAssembly.CreateInstance(iClassName)
                        'CheckInstance
                        If objTheClass Is Nothing Then
                            '     MsgBox("Can't load class...MainClass")
                            Speaker.Speak("Error i Can't load class, " & iClassName)
                            AddPopulatedTextTab(Sender, Sender.TabPages.Count, "Can't load class..." & iClassName, "ERROR")
                            Exit Sub
                        End If
                        'Trys to execute
                        Try
                            objTheClass.GetType.InvokeMember(iMethodName,
                            System.Reflection.BindingFlags.InvokeMethod, Nothing, objTheClass, Nothing)
                        Catch ex As Exception
                            '  MsgBox("Error:" & ex.Message)
                            Speaker.Speak("Error i Can't load Main Function, " & iMethodName)
                            AddPopulatedTextTab(Sender, Sender.TabPages.Count, "Error:" & ex.Message, "ERROR")
                            Speaker.Speak("Woops")
                        End Try
                    End If

                Case "DLL"
                    Speaker.Speak("Create Library")
                    Dim S As New SaveFileDialog
                    With S
                        .Filter = "Executable (*.Dll)|*.Dll"
                        If (.ShowDialog() = DialogResult.OK) Then
                            ExecuteableName = .FileName
                        End If
                    End With
                    'Library type options : /target:library, /target:win, /target:winexe
                    'Generates an executable instead of a class library.
                    'Compiles in memory.
                    Settings.GenerateInMemory = False
                    Settings.GenerateExecutable = False
                    Settings.CompilerOptions = " /target:library"
                    'Set the assembly file name / path
                    Settings.OutputAssembly = ExecuteableName
                    'Read the documentation for the result again variable.
                    'Calls the CompileAssemblyFromSource that will compile the specified source code using the parameters specified in the settings variable.
                    Result = Compiler.CompileAssemblyFromSource(Settings, CodeBlock)
            End Select

#End Region

#Region "Errors"

            'Determines if we have any errors when compiling if so loops through all of the CompileErrors in the Reults variable and displays their ErrorText property.
            If (Result.Errors.Count <> 0) Then
                Speaker.Speak("Error Exception")
                AddPopulatedTextTab(Sender, Sender.TabPages.Count, vbNewLine & "Exception Occured!", "ERROR")

                '  MessageBox.Show("Exception Occured!", "Whoops!", MessageBoxButtons.OK, MessageBoxIcon.Information)
                For Each E As CodeDom.Compiler.CompilerError In Result.Errors

                    AddPopulatedTextTab(Sender, Sender.TabPages.Count, vbNewLine & E.ErrorText, "ERROR")

                Next
            ElseIf (Result.Errors.Count = 0) And CompileType = "EXE" Or CompileType = "DLL" Then
                Speaker.Speak("file saved " & ExecuteableName)
                AddPopulatedTextTab(Sender, Sender.TabPages.Count, vbNewLine & "file saved at " & ExecuteableName, "ERROR")

            End If

#End Region

        End Sub

        Public Shared Async Sub VB_Tab_CodeCompilerAsync(ByVal Sender As TabControl, CodeBlock As String,
                                          Optional ByVal CompileType As String = "DLL",
                                          Optional ByVal iClassName As String = "MainClass",
                                          Optional iMethodName As String = "Execute",
                                          Optional Assemblies As List(Of String) = Nothing,
                                         Optional ByVal EmbededFiles As List(Of String) = Nothing)
            Try
                Speaker.Speak("Compiling")
                Await Task.Run(Sub() VB_Tab_CodeCompiler(Sender, CodeBlock, CompileType, iClassName, iMethodName, Assemblies, EmbededFiles))
            Catch ex As Exception
            End Try
        End Sub

        Enum SearchEngine
            SearchTextAol = 1
            SearchTextGoogle = 2
            SearchTextBing = 3
            Searchwikipedia = 4
            SearchYoutube = 5
            SearchGoogleMaps = 6
            WebSiteSearch = 7
            FilipinoSearch = 8
            StockSearch = 9
            BookSearch = 10
            YellowPagesSearch = 11
            BbcSearch = 12
            PersonSearch = 13
            ProductSearch = 14
            VideoSearch = 15
            RadioSearch = 16
            NewsSearch = 17
            PicSearch = 18
            PhoneSearch = 19
            BusinessSearch = 20
        End Enum


        Public Shared Sub ADD_WebCrawlerUrl(ByRef sender As CheckedListBox, ByRef Url As String)
            Dim Found As Boolean = False
            For Each item In sender.Items
                If item = Url And Found <> True Then
                    Found = True
                Else
                End If
            Next
            If Found = False Then
                sender.Items.Add(Url)
            End If
        End Sub



        Public Shared Function GET_WebCrawlerUrls(ByRef sender As CheckedListBox) As List(Of String)
            Dim Str As New List(Of String)
            For Each item In sender.CheckedItems
                Str.Add(item)
            Next
            Return Str
        End Function

        Public Shared Function DataGridtoCSV(ByRef Sender As DataGridView) As String
            'Build the CSV file data as a Comma separated string.
            Dim csv As String = String.Empty

            'Add the Header row for CSV file.
            For Each column As DataGridViewColumn In Sender.Columns
                csv += column.HeaderText & ","c
            Next

            'Add new line.
            csv += vbCr & vbLf

            'Adding the Rows
            For Each row As DataGridViewRow In Sender.Rows
                For Each cell As DataGridViewCell In row.Cells
                    'Add the Data rows.
                    csv += cell.Value.ToString().Replace(",", ";") & ","c
                Next

                'Add new line.
                csv += vbCr & vbLf
            Next

            Return csv

        End Function

#Region "Data Tabs"

        ''' <summary>
        ''' Current number of tabpages
        ''' </summary>
        Public TabPages As Integer = 0

        ''' <summary>
        ''' Adds a current new tab(document)
        ''' </summary>
        ''' <param name="Sender">tab control</param>
        ''' <param name="Tabcount">tab count</param>
        Public Shared Sub AddDataTab(ByRef Sender As TabControl, ByRef Tabcount As Integer)
            Dim Body As New DataGridView
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()

            Body.ForeColor = Color.MidnightBlue

            Dim DocumentText As String = "Data" & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "Data"
            Sender.SelectedTab = NewPage
        End Sub

        ''' <summary>
        ''' Adds a current new tab(document)
        ''' </summary>
        ''' <param name="Sender">tab control</param>
        ''' <param name="Tabcount">tab count</param>
        Public Shared Sub AddDataTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Data As Object)
            Dim Body As New DataGridView
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Tabcount += 1
            Body.ForeColor = Color.MidnightBlue
            Body.DataSource = Data
            Dim DocumentText As String = "Data" & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "Data"
            Sender.SelectedTab = NewPage
        End Sub

        ''' <summary>
        ''' Returns text held in current tab
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <returns>text</returns>
        Public Shared Function GetCurrentTabDataGridView(ByRef Sender As TabControl) As DataGridView
            If IsDataTab(Sender) = True Then
                Return CType(Sender.SelectedTab.Controls.Item("Body"), DataGridView)
            Else
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Checks if tab is web or text or data
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function IsDataTab(ByRef Sender As TabControl) As Boolean
            IsDataTab = False
            If Sender.SelectedTab.Text = "Data" Then Return True
        End Function

#End Region

        Public Shared Sub LoadPictureBox(ByRef Sender As PictureBox, ByRef UrlList As List(Of Image))
            Speaker.Speak("Selecting Picture")
            Sender.Image = UrlList(0)
        End Sub

        Public Shared Sub LoadListbox(ByRef Sender As ListBox, ByRef UrlList As List(Of String))
            For Each item In UrlList
                Sender.Items.Add(item)
            Next
        End Sub

        Public Shared Sub LoadCheckedListBox(ByRef Sender As CheckedListBox, ByRef UrlList As List(Of String))
            For Each item In UrlList
                Sender.Items.Add(item)
            Next

        End Sub

        Public Shared Sub LoadTextBox(ByRef Sender As TextBox, ByRef _source As String)
            Sender.Text = _source
        End Sub

    End Class

    Public Class WebSite

        Public Enum CrawlerOption
            Html
            SourceCode
            Text
            Images
            HtmlLinks
            ImageLinks
            Browser
        End Enum

        Public Url As String
        Public imageLinks As List(Of String)
        Public images As List(Of Image)
        Public WebPage As String
        Public links As List(Of String)
        Public TextStr As String
        Public WebSource As String

        Public Sub New()
            Url = ""
            imageLinks = New List(Of String)
            images = New List(Of Image)
            links = New List(Of String)
            WebPage = ""
            TextStr = ""
        End Sub

        Public Sub New(ByRef Url As String)
            Me.Url = Url
            ' Crawl(Url)
        End Sub

        Public Shared Function CrawlSite(ByVal url) As WebSite
            Dim NewSite As New WebSite
            NewSite.Url = url
            Try
                NewSite.imageLinks = DownloadImageLinks(url)
            Catch ex As Exception

            End Try
            Try
                NewSite.images = DownloadImagesToList(url)
            Catch ex As Exception

            End Try
            Try
                NewSite.links = DownloadLinks(url)
            Catch ex As Exception

            End Try
            Try
                NewSite.TextStr = DownloadText(url)
            Catch ex As Exception

            End Try
            Try
                NewSite.WebSource = DownloadWebSource(url)
            Catch ex As Exception
            End Try
            Try
                NewSite.WebPage = DownloadHtml(url)
            Catch ex As Exception

            End Try
            Return NewSite
        End Function

        Public Function Crawl(ByVal url As String) As Boolean
            Me.Url = url
            Try
                Me.imageLinks = DownloadImageLinks(url)
            Catch ex As Exception

            End Try
            Try
                Me.images = DownloadImagesToList(url)
            Catch ex As Exception

            End Try

            Try
                Me.links = DownloadLinks(url)
            Catch ex As Exception

            End Try
            Try
                Me.TextStr = DownloadText(url)
            Catch ex As Exception

            End Try
            Return True
        End Function

        Public Loaded As Boolean
        Public Shared Property Speaker As Object

#Region "Download_Images BackGround"

        Public Shared Sub DownloadImagesBackGround(ByVal url As String)
            Dim SwBrowser = New WebBrowser
            SwBrowser.ScriptErrorsSuppressed = True
            SwBrowser.Navigate(url)
            AddHandler SwBrowser.DocumentCompleted, AddressOf browserImageLoad_DocumentCompleted
        End Sub

        Private Shared Sub browserImageLoad_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)
            Dim browser As WebBrowser = CType(sender, WebBrowser)
            Dim collection As HtmlElementCollection
            Dim imgListString As List(Of HtmlElement) = New List(Of HtmlElement)()
            If browser IsNot Nothing Then
                If browser.Document IsNot Nothing Then
                    collection = browser.Document.GetElementsByTagName("img")
                    If collection IsNot Nothing Then
                        For Each element As HtmlElement In collection
                            Dim wClient As WebClient = New WebClient()
                            Dim urlDownload As String = element.GetAttribute("src")
                            wClient.DownloadFile(urlDownload, urlDownload.Substring(urlDownload.LastIndexOf("/"c)))
                        Next
                    End If
                    Dim LinkStr As New List(Of String)
                    Dim images As HtmlElementCollection = browser.Document.Images
                    For Each link As HtmlElement In images
                        LinkStr.Add(link.GetAttribute("img"))
                    Next
                    Dim links As HtmlElementCollection = browser.Document.Links
                    For Each link As HtmlElement In links
                        LinkStr.Add(link.GetAttribute("img"))
                    Next
                    Dim StrPath = Application.StartupPath & "\Images"
                    Dim Counter As Integer = 0
                    Parallel.ForEach(LinkStr, Sub(item)
                                                  Counter += 1
                                                  Try
                                                      DownloadImage(item).Save(StrPath & "\" & CleanUrl(browser.Url.ToString) & "\" & Counter & ".png")
                                                  Catch ex As Exception
                                                  End Try
                                              End Sub)
                End If
            End If
        End Sub

#End Region

#Region "Download Html Links Background"

        Public Shared Sub DownloadHtmlLinksBackGround(ByVal url As String)
            Dim SwBrowser = New WebBrowser
            SwBrowser.ScriptErrorsSuppressed = True
            SwBrowser.Navigate(url)
            AddHandler SwBrowser.DocumentCompleted, AddressOf browserLinksLoad_DocumentCompleted
        End Sub

        Private Shared Sub browserLinksLoad_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)
            Dim browser As WebBrowser = CType(sender, WebBrowser)
            Dim collection As HtmlElementCollection
            If browser IsNot Nothing Then
                If browser.Document IsNot Nothing Then
                    collection = browser.Document.GetElementsByTagName("img")
                    If collection IsNot Nothing Then
                        For Each element As HtmlElement In collection
                            Dim wClient As WebClient = New WebClient()
                            Dim urlDownload As String = element.GetAttribute("src")
                            wClient.DownloadFile(urlDownload, urlDownload.Substring(urlDownload.LastIndexOf("/"c)))
                        Next
                    End If
                    Dim LinkStr As New List(Of String)
                    Dim HRefFromImgs As HtmlElementCollection = browser.Document.Images
                    For Each link As HtmlElement In HRefFromImgs
                        LinkStr.Add(link.GetAttribute("href"))
                    Next
                    Dim HRefFromlnks As HtmlElementCollection = browser.Document.Links
                    For Each link As HtmlElement In HRefFromlnks
                        LinkStr.Add(link.GetAttribute("href"))
                    Next
                    Dim imgFromImgs As HtmlElementCollection = browser.Document.Images
                    For Each link As HtmlElement In imgFromImgs
                        LinkStr.Add(link.GetAttribute("img"))
                    Next
                    Dim imgFromLinks As HtmlElementCollection = browser.Document.Links
                    For Each link As HtmlElement In imgFromLinks
                        LinkStr.Add(link.GetAttribute("img"))
                    Next
                    Dim Str As String = ""
                    Parallel.ForEach(LinkStr, Sub(item)
                                                  Str &= item & vbNewLine
                                              End Sub)
                    My.Computer.FileSystem.WriteAllText(CleanUrl(browser.Url.ToString) & ".txt", Str, True)
                End If
            End If
        End Sub

#End Region

        ''' <summary>
        ''' Cleans up URL
        ''' </summary>
        ''' <param name="Url">www.google.com to www_google_com</param>
        ''' <returns></returns>
        Public Shared Function CleanUrl(ByVal Url As String) As String
            Dim Str As String = ""
            Str = Url.Replace("Http://", "")
            Str = Str.Replace(".", "_")
            Str = Str.Replace("/", "_")
            Str = Str.Replace(":", "_")
            Return Str
        End Function

#Region "Main Functions"

#Region "Text"

        ''' <summary>
        ''' Download text from url webpage
        ''' </summary>
        ''' <param name="url">http://www.spydazweb.co.uk</param>
        ''' <returns></returns>
        Public Shared Function DownloadText(ByVal url As String) As String
            Dim LinkStr As String = ""
            Try
                Dim Str As String = DownloadWebSource(url)
                Dim SwBrowser = New WebBrowser
                SwBrowser.ScriptErrorsSuppressed = True
                SwBrowser.DocumentText = Str
                Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
                SwBrowser.Document.Write(Str)
                If SwBrowser IsNot Nothing Then
                    LinkStr = SwBrowser.Document.Body.InnerText
                End If
            Catch ex As Exception

            End Try

            Return LinkStr
        End Function

        ''' <summary>
        ''' Download text from url webpage
        ''' </summary>
        ''' <param name="url">http://www.spydazweb.co.uk</param>
        Public Shared Sub DownloadTextToFile(ByVal url As String)
            Dim LinkStr As String = ""
            Try
                Dim Str As String = DownloadWebSource(url)
                Dim SwBrowser = New WebBrowser
                SwBrowser.ScriptErrorsSuppressed = True
                SwBrowser.DocumentText = Str
                Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
                SwBrowser.Document.Write(Str)
                If SwBrowser IsNot Nothing Then
                    LinkStr = SwBrowser.Document.Body.InnerText
                    Dim fl_Name = CleanUrl(url)
                    fl_Name = Application.StartupPath & "\Downloads\" & fl_Name & "\Text\"
                    If (Not System.IO.Directory.Exists(fl_Name)) Then
                        System.IO.Directory.CreateDirectory(fl_Name)
                    End If
                    Try
                        File.WriteAllText(fl_Name & "_Text_" & ".txt", LinkStr)
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString, "Error")
                    End Try
                End If
            Catch ex As Exception

            End Try
        End Sub

#End Region

#Region "SourceCode"

        ''' <summary>
        ''' Downloads Html fomr Url (Slightliy Formatted)
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns>html screen (formatted)</returns>
        Public Shared Function DownloadHtml(ByVal Url As String) As String
            Dim LinkStr As String = ""
            Try
                Dim Str As String = DownloadWebSource(Url)
                Dim Browser = New WebBrowser
                Browser.ScriptErrorsSuppressed = True
                Browser.DocumentText = Str
                Dim htmlDoc As HtmlDocument = Browser.Document.OpenNew(True)
                Browser.Document.Write(Str)
                If Browser IsNot Nothing Then
                    LinkStr = Browser.Document.Body.InnerHtml
                End If
            Catch ex As Exception

            End Try
            Return LinkStr
        End Function

        Public Shared Async Function DownloadHtmlAsync(ByVal Url As String) As Task(Of String)
            Dim html As String = ""
            Try
                html = Await Task.Run(Function() DownloadHtml(Url))
            Catch ex As Exception

            End Try
            Return html
        End Function

        ''' <summary>
        ''' Download Web Page source (raw)
        ''' </summary>
        ''' <param name="Url"></param>
        ''' <returns></returns>
        Public Shared Function DownloadWebSource(ByVal Url As String) As String
            Dim html As String = ""
            Try
                Dim request As WebRequest = WebRequest.Create(Url)

                Using response As WebResponse = request.GetResponse()
                    Using reader As New StreamReader(response.GetResponseStream())
                        Try
                            html = reader.ReadToEnd()
                        Catch
                        End Try

                        'Dim fl_Name = CleanUrl(Url)
                        'fl_Name = Application.StartupPath & "\Downloads\" & fl_Name & "\Source\"
                        'If (Not System.IO.Directory.Exists(fl_Name)) Then
                        '    System.IO.Directory.CreateDirectory(fl_Name)
                        'End If
                        'Try
                        '    File.WriteAllText(fl_Name & "_Source_" & ".html", html)
                        'Catch ex As Exception
                        '    MessageBox.Show(ex.ToString, "Error")
                        'End Try

                    End Using
                End Using
            Catch ex As Exception

            End Try

            Return html
        End Function

        Public Shared Async Function DownloadWebSourceAsync(ByVal Url As String) As Task(Of String)
            Dim html As String = ""
            Try
                html = Await Task.Run(Function() DownloadWebSource(Url))
            Catch ex As Exception

            End Try
            Return html
        End Function

        ''' <summary>
        ''' Download Web Page source (raw)
        ''' </summary>
        ''' <param name="Url"></param>
        Public Shared Sub DownloadWebSourceToFile(ByVal Url As String)
            Try
                Dim request As WebRequest = WebRequest.Create(Url)
                Dim html As String = ""
                Using response As WebResponse = request.GetResponse()
                    Using reader As New StreamReader(response.GetResponseStream())
                        html = reader.ReadToEnd()
                        Dim fl_Name = CleanUrl(Url)
                        fl_Name = Application.StartupPath & "\Downloads\" & fl_Name & "\html\"
                        If (Not System.IO.Directory.Exists(fl_Name)) Then
                            System.IO.Directory.CreateDirectory(fl_Name)
                        End If
                        Try
                            File.WriteAllText(fl_Name & "_WebSource_" & ".html", html)
                        Catch ex As Exception
                            MessageBox.Show(ex.ToString, "Error")
                        End Try
                    End Using
                End Using
            Catch ex As Exception

            End Try

        End Sub

        ''' <summary>
        ''' Downloads Html fomr Url (Slightliy Formatted)
        ''' </summary>
        ''' <param name="url"></param>
        Public Shared Sub DownloadHtmlToFile(ByVal Url As String)
            Dim LinkStr As String = ""
            Try
                Dim Str As String = DownloadWebSource(Url)
                Dim SwBrowser = New WebBrowser
                SwBrowser.ScriptErrorsSuppressed = True
                SwBrowser.DocumentText = Str
                Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
                SwBrowser.Document.Write(Str)
                If SwBrowser IsNot Nothing Then
                    LinkStr = SwBrowser.Document.Body.InnerHtml
                    Dim fl_Name = CleanUrl(Url)
                    fl_Name = Application.StartupPath & "\Downloads\" & fl_Name & "\html\"
                    If (Not System.IO.Directory.Exists(fl_Name)) Then
                        System.IO.Directory.CreateDirectory(fl_Name)
                    End If
                    Try
                        File.WriteAllText(fl_Name & "_html_" & ".html", LinkStr)
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString, "Error")
                    End Try
                End If
            Catch ex As Exception

            End Try
        End Sub

#End Region

#Region "Links"

        ''' <summary>
        ''' Download Html Links from Url
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns>List of Links</returns>
        Public Shared Function DownloadLinks(ByVal url As String) As List(Of String)
            Dim LinkStr As New List(Of String)
            Dim Str As String = DownloadWebSource(url)

            Dim SwBrowser = New WebBrowser
            SwBrowser.ScriptErrorsSuppressed = True
            SwBrowser.DocumentText = Str
            Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
            SwBrowser.Document.Write(Str)
            If SwBrowser IsNot Nothing Then
                Dim HrefFromImgs As HtmlElementCollection = SwBrowser.Document.Images
                For Each link As HtmlElement In HrefFromImgs
                    LinkStr.Add(link.GetAttribute("href"))
                Next
                Dim HrefFromLinks As HtmlElementCollection = SwBrowser.Document.Links
                For Each link As HtmlElement In HrefFromLinks
                    LinkStr.Add(link.GetAttribute("href"))
                Next
            End If
            Return LinkStr
        End Function

        Public Shared Async Function DownloadLinksAsync(ByVal url As String) As Task(Of List(Of String))
            Dim LinkStr As New List(Of String)
            Dim Str As String = Await DownloadWebSourceAsync(url)

            Dim SwBrowser = New WebBrowser
            SwBrowser.ScriptErrorsSuppressed = True
            SwBrowser.DocumentText = Str
            Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
            SwBrowser.Document.Write(Str)
            If SwBrowser IsNot Nothing Then
                Dim HrefFromImgs As HtmlElementCollection = SwBrowser.Document.Images
                For Each link As HtmlElement In HrefFromImgs
                    LinkStr.Add(link.GetAttribute("href"))
                Next
                Dim HrefFromLinks As HtmlElementCollection = SwBrowser.Document.Links
                For Each link As HtmlElement In HrefFromLinks
                    LinkStr.Add(link.GetAttribute("href"))
                Next
            End If
            Return LinkStr
        End Function

        ''' <summary>
        ''' Attmpts to get list of image links
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Shared Function DownloadImageLinks(ByVal url As String) As List(Of String)
            Dim LinkStr As New List(Of String)
            Try
                Dim Str As String = DownloadWebSource(url)
                Dim SwBrowser = New WebBrowser
                SwBrowser.ScriptErrorsSuppressed = True
                SwBrowser.DocumentText = Str
                Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
                SwBrowser.Document.Write(Str)
                If SwBrowser IsNot Nothing Then
                    Dim HrefFromImgs As HtmlElementCollection = SwBrowser.Document.Images
                    For Each link As HtmlElement In HrefFromImgs
                        LinkStr.Add(link.GetAttribute("href"))
                    Next
                    'Dim imgFromLinks As HtmlElementCollection = SwBrowser.Document.Links
                    'For Each link As HtmlElement In imgFromLinks
                    '    LinkStr.Add(link.GetAttribute("img"))
                    'Next
                End If
            Catch ex As Exception

            End Try
            Return LinkStr
        End Function

        Public Shared Async Function DownloadImageLinksAsync(ByVal url As String) As Task(Of List(Of String))
            Dim LinkStr As New List(Of String)
            Dim Str As String = Await DownloadWebSourceAsync(url)
            Dim SwBrowser = New WebBrowser
            SwBrowser.ScriptErrorsSuppressed = True
            SwBrowser.DocumentText = Str
            Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
            SwBrowser.Document.Write(Str)
            If SwBrowser IsNot Nothing Then
                Dim HrefFromImgs As HtmlElementCollection = SwBrowser.Document.Images
                For Each link As HtmlElement In HrefFromImgs
                    LinkStr.Add(link.GetAttribute("href"))
                Next
                'Dim imgFromLinks As HtmlElementCollection = SwBrowser.Document.Links
                'For Each link As HtmlElement In imgFromLinks
                '    LinkStr.Add(link.GetAttribute("img"))
                'Next
            End If
            Return LinkStr
        End Function

        ''' <summary>
        ''' Download Html Links from Url
        ''' </summary>
        ''' <param name="url"></param>
        Public Shared Sub DownloadLinksToFile(ByVal url As String)
            Dim LinkStr As New List(Of String)
            Try
                Dim Str As String = DownloadWebSource(url)
                Dim SwBrowser = New WebBrowser
                SwBrowser.ScriptErrorsSuppressed = True
                SwBrowser.DocumentText = Str
                Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
                SwBrowser.Document.Write(Str)
                If SwBrowser IsNot Nothing Then
                    Dim HrefFromImgs As HtmlElementCollection = SwBrowser.Document.Images
                    For Each link As HtmlElement In HrefFromImgs
                        LinkStr.Add(link.GetAttribute("href"))
                    Next
                    Dim HrefFromLinks As HtmlElementCollection = SwBrowser.Document.Links
                    For Each link As HtmlElement In HrefFromLinks
                        LinkStr.Add(link.GetAttribute("href"))
                    Next
                    Dim StrLnks As String = ""
                    For Each item In LinkStr
                        StrLnks &= item & vbNewLine

                    Next
                    Dim fl_Name = CleanUrl(url)
                    fl_Name = Application.StartupPath & "\Downloads\" & fl_Name & "\Links\"
                    If (Not System.IO.Directory.Exists(fl_Name)) Then
                        System.IO.Directory.CreateDirectory(fl_Name)
                    End If
                    Try
                        File.WriteAllText(fl_Name & "_Links_" & ".txt", StrLnks)
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString, "Error")
                    End Try
                End If
            Catch ex As Exception

            End Try
        End Sub

        ''' <summary>
        ''' Attempts to get list of image links
        ''' </summary>
        ''' <param name="url"></param>
        ''' <returns></returns>
        Public Shared Function DownloadImageLinksToFile(ByVal url As String) As List(Of String)
            Dim LinkStr As New List(Of String)
            Try
                Dim Str As String = DownloadWebSource(url)
                Dim SwBrowser = New WebBrowser
                SwBrowser.ScriptErrorsSuppressed = True
                SwBrowser.DocumentText = Str
                Dim htmlDoc As HtmlDocument = SwBrowser.Document.OpenNew(True)
                SwBrowser.Document.Write(Str)
                If SwBrowser IsNot Nothing Then
                    Dim HrefFromImgs As HtmlElementCollection = SwBrowser.Document.Images
                    For Each link As HtmlElement In HrefFromImgs
                        LinkStr.Add(link.GetAttribute("href"))
                    Next
                    Dim StrLnks As String = ""
                    For Each item In LinkStr
                        StrLnks &= item & vbNewLine

                    Next
                    Dim fl_Name = CleanUrl(url)
                    fl_Name = Application.StartupPath & "\Downloads\" & fl_Name & "\Links\"
                    If (Not System.IO.Directory.Exists(fl_Name)) Then
                        System.IO.Directory.CreateDirectory(fl_Name)
                    End If
                    Try
                        File.WriteAllText(fl_Name & "_Links_" & ".txt", StrLnks)
                    Catch ex As Exception
                        MessageBox.Show(ex.ToString, "Error")
                    End Try
                End If
            Catch ex As Exception

            End Try
            Return LinkStr
        End Function

#End Region

#Region "Files"

        ''' <summary>
        ''' Download urlLink
        ''' </summary>
        ''' <param name="UrlLink">Fully qualified link to item</param>
        ''' <param name="Filename">filename for item</param>
        ''' <returns></returns>
        Public Shared Function DownloadFile(ByRef UrlLink As String, ByRef Filename As String) As Boolean
            Try
                ' Create a new WebClient instance.
                Dim myWebClient As New WebClient()
                Try
                    myWebClient.DownloadFile(UrlLink, Filename)
                    Return True
                Catch ex As Exception
                    Return False
                End Try
            Catch ex As Exception

            End Try
            Return False
        End Function

#End Region

#Region "Images"

        Public Shared Function DownloadImagesToList(ByRef Url As String) As List(Of Image)
            DownloadImagesToList = New List(Of Image)
            Dim lst = DownloadImageLinks(Url)
            Parallel.ForEach(lst, Sub(item)
                                      DownloadImagesToList.Add(DownloadImage(item))
                                  End Sub)
        End Function

        Public Shared Sub DownloadImages(ByRef Url As String)
            Dim iurl As String = Url
            Dim Foldername As String = CleanUrl(iurl)

            Foldername = Application.StartupPath & "\Downloads\" & Foldername & "\images\"
            If (Not System.IO.Directory.Exists(Foldername)) Then
                System.IO.Directory.CreateDirectory(Foldername)
            End If

            Try
                Dim links = DownloadImageLinks(Url)
                Dim Counter As Integer = 0
                For Each item In links
                    Try
                        Dim fi As New FileInfo(Foldername & Counter & ".png")
                        If fi.Exists Then
                            fi.Delete()
                        End If
                        WebSite.DownloadFile(item, fi.FullName)
                        Counter += 1
                    Catch ex As Exception

                    End Try
                Next
            Catch ex As Exception

            End Try

        End Sub

        ''' <summary>
        ''' Download Image from url http://www.spydazweb.co.uk/image.gif
        ''' </summary>
        ''' <param name="Url"></param>
        ''' <returns>image</returns>
        Public Shared Function DownloadImage(ByVal Url As String) As Image
            Try
                Return MasterTabControl.GetPicture(Url)
            Catch ex As Exception
                Try
                    Using webClient = New WebClient()
                        Try
                            Dim image_stream As New MemoryStream(webClient.DownloadData(Url))
                            Return Image.FromStream(image_stream)
                        Catch ex1 As Exception
                            Return Nothing
                        End Try
                    End Using
                Catch ex2 As Exception

                End Try

            End Try

            Return Nothing

        End Function

        Public Shared Sub DownloadImagesToFiles(ByRef Url As String)
            Try
                DownloadImages(Url)
            Catch ex As Exception

            End Try

        End Sub

#End Region

#End Region

#Region "Querys"

        Public Shared Function GetQuery(ByVal Url As String) As WebSite
            Dim x As New WebSite(Url)
            Try
                x = WebSite.CrawlSite(Url)
            Catch ex As Exception
            End Try
            Return x
        End Function

        Public Shared Function GetQuerys(ByVal Urls As List(Of String)) As List(Of WebSite)
            Dim Querys As New List(Of WebSite)
            Parallel.ForEach(Urls, Sub(item)
                                       Querys.Add(GetQuery(item))
                                   End Sub)
            Return Querys
        End Function

        Public Shared Function GetQuerysParallel(ByVal Urls As List(Of String)) As List(Of WebSite)
            Dim Querys As New List(Of WebSite)
            Parallel.ForEach(Urls, Sub(item)
                                       Try
                                           Dim x = WebSite.GetQuery(item)
                                           Querys.Add(x)
                                       Catch ex As Exception

                                       End Try
                                   End Sub)
            Return Querys
        End Function

#End Region

    End Class
    ''' <summary>
    ''' Colors Words in RichText Box
    ''' </summary>
    Public Class SyntaxHighlighter

#Region "Public Fields"

        Public Shared SyntaxTerms() As String = ({"SPYDAZ", "ABS", "ACCESS", "ADDITEM", "ADDNEW", "ALIAS", "AND", "ANY", "APP", "APPACTIVATE", "APPEND", "APPENDCHUNK", "ARRANGE", "AS", "ASC", "ATN", "BASE", "BEEP", "BEGINTRANS", "BINARY", "BYVAL", "CALL", "CASE", "CCUR", "CDBL", "CHDIR", "CHDRIVE", "CHR", "CHR$", "CINT", "CIRCLE", "CLEAR", "CLIPBOARD", "CLNG", "CLOSE", "CLS", "COMMAND", "
COMMAND$", "COMMITTRANS", "COMPARE", "CONST", "CONTROL", "CONTROLS", "COS", "CREATEDYNASET", "CSNG", "CSTR", "CURDIR$", "CURRENCY", "CVAR", "CVDATE", "DATA", "DATE", "DATE$", "DATESERIAL", "DATEVALUE", "DAY", "
DEBUG", "DECLARE", "DEFCUR", "CEFDBL", "DEFINT", "DEFLNG", "DEFSNG", "DEFSTR", "DEFVAR", "DELETE", "DIM", "DIR", "DIR$", "DO", "DOEVENTS", "DOUBLE", "DRAG", "DYNASET", "EDIT", "ELSE", "ELSEIF", "END", "ENDDOC", "ENDIF", "
ENVIRON$", "EOF", "EQV", "ERASE", "ERL", "ERR", "ERROR", "ERROR$", "EXECUTESQL", "EXIT", "EXP", "EXPLICIT", "FALSE", "FIELDSIZE", "FILEATTR", "FILECOPY", "FILEDATETIME", "FILELEN", "FIX", "FOR", "FORM", "FORMAT", "
FORMAT$", "FORMS", "FREEFILE", "FUNCTION", "GET", "GETATTR", "GETCHUNK", "GETDATA", "DETFORMAT", "GETTEXT", "GLOBAL", "GOSUB", "GOTO", "HEX", "HEX$", "HIDE", "HOUR", "IF", "IMP", "INPUT", "INPUT$", "INPUTBOX", "INPUTBOX$", "
INSTR", "INT", "INTEGER", "IS", "ISDATE", "ISEMPTY", "ISNULL", "ISNUMERIC", "KILL", "LBOUND", "LCASE", "LCASE$", "LEFT", "LEFT$", "LEN", "LET", "LIB", "LIKE", "LINE", "LINKEXECUTE", "LINKPOKE", "LINKREQUEST", "
LINKSEND", "LOAD", "LOADPICTURE", "LOC", "LOCAL", "LOCK", "LOF", "LOG", "LONG", "LOOP", "LSET", "LTRIM",
                        "LTRIM$", "ME", "MID", "MID$", "MINUTE", "MKDIR", "MOD", "MONTH", "MOVE", "MOVEFIRST", "MOVELAST", "MOVENEXT", "MOVEPREVIOUS",
                        "MOVERELATIVE", "MSGBOX", "NAME", "NEW", "NEWPAGE", "NEXT", "NEXTBLOCK", "NOT", "NOTHING",
                        "NOW", "NULL", "OCT", "OCT$", "ON", "OPEN", "OPENDATABASE", "OPTION", "OR", "OUTPUT", "POINT", "PRESERVE", "PRINT",
                        "PRINTER", "PRINTFORM", "PRIVATE", "PSET", "PUT", "PUBLIC", "QBCOLOR", "RANDOM", "RANDOMIZE", "READ", "REDIM", "REFRESH",
                        "REGISTERDATABASE", "REM", "REMOVEITEM", "RESET", "RESTORE", "RESUME", "RETURN", "RGB", "RIGHT", "RIGHT$", "RMDIR", "RND",
                        "ROLLBACK", "RSET", "RTRIM", "RTRIM$", "SAVEPICTURE", "SCALE", "SECOND", "SEEK", "SELECT", "SENDKEYS", "SET", "SETATTR",
                        "SETDATA", "SETFOCUS", "SETTEXT", "SGN", "SHARED",
                        "SHELL", "SHOW", "SIN", "SINGLE", "SPACE", "SPACE$", "SPC", "SQR",
                        "STATIC", "STEP", "STOP", "STR", "STR$", "STRCOMP", "STRING", "STRING$", "SUB",
                        "SYSTEM", "TAB", "TAN", "TEXT", "TEXTHEIGHT", "TEXTWIDTH", "THEN", "TIME", "TIME$",
                        "TIMER", "TIMESERIAL", "TIMEVALUE", "TO", "TRIM",
                        "TRIM$", "TRUE", "TYPE", "TYPEOF", "UBOUND", "UCASE", "UCASE$", "UNLOAD",
                        "UNLOCK", "UNTIL", "UPDATE", "USING", "VAL", "VARIANT", "VARTYPE", "WEEKDAY", "WEND", "WHILE", "WIDTH",
                        "WRITE", "XOR", "YEAR", "ZORDER", "ADDHANDLER", "ADDRESSOF", "ALIAS", "AND", "ANDALSO", "AS", "BYREF",
                        "BOOLEAN", "BYTE", "BYVAL", "CALL", "CASE", "CATCH", "CBOOL", "CBYTE", "CCHAR", "CDATE",
                        "CDEC", "CDBL", "CHAR", "CINT", "CLASS", "CLNG", "COBJ", "CONST", "CONTINUE", "CSBYTE",
                        "CSHORT", "CSNG", "CSTR", "CTYPE", "CUINT", "CULNG", "CUSHORT", "DATE", "DECIMAL", "DECLARE",
                        "DEFAULT", "DELEGATE", "DIM", "DIRECTCAST", "DOUBLE", "DO", "EACH", "ELSE", "ELSEIF", "END",
                        "ENDIF", "ENUM", "ERASE", "ERROR", "EVENT", "EXIT", "FALSE", "FINALLY", "FOR", "FRIEND",
                        "FUNCTION", "GET", "GETTYPE", "GLOBAL", "GOSUB", "GOTO", "HANDLES", "IF", "IMPLEMENTS",
                        "IMPORTS", "IN", "INHERITS", "INTEGER", "INTERFACE", "IS", "ISNOT", "LET", "LIB", "LIKE",
                        "LONG", "LOOP", "ME", "MOD", "MODULE", "MUSTINHERIT", "MUSTOVERRIDE", "MYBASE", "MYCLASS",
                        "NAMESPACE", "NARROWING", "NEW", "NEXT", "NOT", "NOTHING", "NOTINHERITABLE", "NOTOVERRIDABLE",
                        "OBJECT", "ON", "OF", "OPERATOR", "OPTION", "OPTIONAL", "OR", "ORELSE", "OVERLOADS",
                        "OVERRIDABLE", "OVERRIDES", "PARAMARRAY", "PARTIAL", "PRIVATE", "PROPERTY", "PROTECTED",
                        "PUBLIC", "RAISEEVENT", "READONLY", "REDIM", "REM", "REMOVEHANDLER", "RESUME", "RETURN",
                        "SBYTE", "SELECT", "SET", "SHADOWS", "SHARED", "SHORT", "SINGLE", "STATIC", "STEP", "STOP",
                        "STRING", "STRUCTURE", "SUB", "SYNCLOCK", "THEN", "THROW", "TO", "TRUE", "TRY", "TRYCAST",
                        "TYPEOF", "WEND", "VARIANT", "UINTEGER", "ULONG", "USHORT", "USING", "WHEN", "WHILE", "WIDENING",
                        "WITH", "WITHEVENTS", "WRITEONLY",
                        "XOR", "#CONST", "#ELSE", "#ELSEIF", "#END", "#IF"})

#End Region

#Region "Private Fields"

        Private Shared indexOfSearchText As Integer = 0

        Private Shared start As Integer = 0

        Private mGrammar As New List(Of String)

#End Region

#Region "Public Methods"

        Public Shared Sub ColorSearchTerm(ByRef SearchStr As String, Rtb As RichTextBox)
            ColorSearchTerm(SearchStr, Rtb, Color.CadetBlue)
        End Sub

        Public Shared Sub ColorSearchTerm(ByRef SearchStr As String, Rtb As RichTextBox, ByRef MyColor As Color)
            Dim startindex As Integer = 0
            start = 0
            indexOfSearchText = 0

            If SearchStr <> "" Then

                SearchStr = SearchStr & " "
                If SearchStr.Length > 0 Then
                    startindex = FindText(Rtb, ProperCase(SearchStr), start, Rtb.Text.Length)
                End If
                If SearchStr.Length > 0 And startindex = 0 Then
                    startindex = FindText(Rtb, LCase(SearchStr), start, Rtb.Text.Length)
                End If
                If SearchStr.Length > 0 And startindex = 0 Then
                    startindex = FindText(Rtb, UCase(SearchStr), start, Rtb.Text.Length)
                End If
                If SearchStr.Length > 0 And startindex = 0 Then
                    startindex = FindText(Rtb, SearchStr, start, Rtb.Text.Length)
                End If
                ' If string was found in the RichTextBox, highlight it
                If startindex >= 0 Then
                    ' Set the highlight color as red
                    Rtb.SelectionColor = MyColor

                    ' Find the end index. End Index = number of characters in textbox
                    Dim endindex As Integer = SearchStr.Length
                    ' Highlight the search string

                    Rtb.Select(startindex, endindex)
                    Rtb.SelectedText.ToUpper()
                    ' mark the start position after the position of last search string
                    start = startindex + endindex

                End If
            Else
            End If
            Rtb.Select(Rtb.TextLength, Rtb.TextLength)
        End Sub

        Public Shared Sub HighlightInternalSyntax(ByRef sender As RichTextBox)

            For Each Word As String In SyntaxTerms
                HighlightWord(sender, Word)
                HighlightWord(sender, LCase(Word))
                HighlightWord(sender, ProperCase(Word))
            Next

        End Sub

        Public Shared Sub HighlightWord(ByRef sender As RichTextBox, ByRef word As String)

            Dim index As Integer = 0
            While index < sender.Text.LastIndexOf(word)
                'find
                sender.Find(word, index, sender.TextLength, RichTextBoxFinds.WholeWord)
                'select and color
                sender.SelectionColor = Color.OrangeRed
                index = sender.Text.IndexOf(word, index) + 1
            End While
        End Sub

        ''' <summary>
        ''' Searches For Internal Syntax
        ''' </summary>
        ''' <param name="Rtb"></param>
        ''' <remarks></remarks>
        Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox)
            'Searches Basic Syntax
            For Each wrd As String In SyntaxTerms
                start = 0
                indexOfSearchText = 0
                ColorSearchTerm(wrd, Rtb)
            Next
            For Each wrd As String In SyntaxTerms
                start = 0
                indexOfSearchText = 0
                ColorSearchTerm(wrd, Rtb)
            Next
        End Sub

        ''' <summary>
        ''' Searches For Internal Syntax
        ''' </summary>
        ''' <param name="Rtb"></param>
        ''' <remarks></remarks>
        Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef Terms As List(Of String))
            'Searches Basic Syntax
            For Each wrd As String In SyntaxTerms
                start = 0
                indexOfSearchText = 0
                ColorSearchTerm(wrd, Rtb)
            Next
            For Each wrd As String In Terms
                start = 0
                indexOfSearchText = 0
                ColorSearchTerm(UCase(wrd), Rtb)
            Next
        End Sub

        ''' <summary>
        ''' Searches for Specific Word to colorize (Blue)
        ''' </summary>
        ''' <param name="Rtb"></param>
        ''' <param name="SearchStr"></param>
        ''' <remarks></remarks>
        Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef SearchStr As String)
            start = 0
            indexOfSearchText = 0
            ColorSearchTerm(SearchStr, Rtb)
        End Sub

        ''' <summary>
        ''' Searches for Specfic word to colorize specified color
        ''' </summary>
        ''' <param name="Rtb"></param>
        ''' <param name="SearchStr"></param>
        ''' <param name="MyColor"></param>
        ''' <remarks></remarks>
        Public Shared Sub SearchSyntax(ByRef Rtb As RichTextBox, ByRef SearchStr As String, MyColor As Color)

            start = 0
            indexOfSearchText = 0
            ColorSearchTerm(SearchStr, Rtb, MyColor)

        End Sub

#End Region

#Region "Private Methods"

        Private Shared Function FindText(ByRef Rtb As RichTextBox, SearchStr As String,
                                                                            ByVal searchStart As Integer, ByVal searchEnd As Integer) As Integer

            ' Unselect the previously searched string
            If searchStart > 0 AndAlso searchEnd > 0 AndAlso indexOfSearchText >= 0 Then
                Rtb.Undo()
            End If

            ' Set the return value to -1 by default.
            Dim retVal As Integer = -1

            ' A valid starting index should be specified. if indexOfSearchText = -1, the end of search
            If searchStart >= 0 AndAlso indexOfSearchText >= 0 Then

                ' A valid ending index
                If searchEnd > searchStart OrElse searchEnd = -1 Then

                    ' Find the position of search string in RichTextBox
                    indexOfSearchText = Rtb.Find(SearchStr, searchStart, searchEnd, RichTextBoxFinds.WholeWord)

                    ' Determine whether the text was found in richTextBox1.
                    If indexOfSearchText <> -1 Then
                        ' Return the index to the specified search text.
                        retVal = indexOfSearchText
                    End If

                End If
            End If
            Return retVal

        End Function

        ''' <summary>
        ''' Returns Propercase Sentence
        ''' </summary>
        ''' <param name="TheString">String to be formatted</param>
        ''' <returns></returns>
        Private Shared Function ProperCase(ByRef TheString As String) As String
            ProperCase = UCase(Left(TheString, 1))

            For i = 2 To Len(TheString)

                ProperCase = If(Mid(TheString, i - 1, 1) = " ", ProperCase & UCase(Mid(TheString, i, 1)), ProperCase & LCase(Mid(TheString, i, 1)))
            Next i
        End Function

#End Region

    End Class
    ''' <summary>
    ''' Tabpage Extensions
    ''' </summary>
    Public Module TabExt
        <Extension>
        Public Function AddForm_Click(ByRef NewForm As Form, ByRef Title As String, ByRef iTabcontrol As TabControl)
            ' Create a new instance of the form to be opened


            Dim newTabPage As New TabPage(Title)

            ' Set the Dock property of the form to fill the TabPage
            NewForm.TopLevel = False
            NewForm.FormBorderStyle = FormBorderStyle.None
            NewForm.Dock = DockStyle.Fill
            ' Create a new TabPage

            ' Add the new form to the new TabPage
            newTabPage.Controls.Add(NewForm)
            iTabcontrol.TabPages.Add(newTabPage)
            NewForm.Show()
            iTabcontrol.Visible = True
            Return iTabcontrol
        End Function

#Region "textTabs"

        ''' <summary>
        ''' Returns text held in current tab
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <returns>text</returns>
        <System.Runtime.CompilerServices.Extension()>
        Public Function GetCurrentTabRichTextBox(ByRef Sender As TabControl) As RichTextBox
            If IsTextTab(Sender) = True Then
                Return CType(Sender.SelectedTab.Controls.Item("Body"), RichTextBox)
            Else
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Checks if tab is web or text
        ''' </summary>
        ''' <returns></returns>
        <System.Runtime.CompilerServices.Extension()>
        Public Function IsTextTab(ByRef Sender As TabControl) As Boolean
            If CType(Sender.SelectedTab.Controls.Item("Body"), RichTextBox) IsNot Nothing Then
                IsTextTab = True
            Else
                IsTextTab = False
            End If
        End Function

        <System.Runtime.CompilerServices.Extension()>
        Public Sub BOld(ByRef Sender As TabControl)
            Dim BoldFont As New Font(GetCurrentTabRichTextBox(Sender).SelectionFont.Name, GetCurrentTabRichTextBox(Sender).SelectionFont.Size, FontStyle.Bold)
            Dim RegularFont As New Font(GetCurrentTabRichTextBox(Sender).SelectionFont.Name, GetCurrentTabRichTextBox(Sender).SelectionFont.Size, FontStyle.Regular)

            GetCurrentTabRichTextBox(Sender).SelectionFont = If(GetCurrentTabRichTextBox(Sender).SelectionFont.Bold, RegularFont, BoldFont)
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub Itallic(ByRef Sender As TabControl)
            Dim ItalicFont As New Font(GetCurrentTabRichTextBox(Sender).SelectionFont.Name, GetCurrentTabRichTextBox(Sender).SelectionFont.Size, FontStyle.Italic)
            Dim RegularFont As New Font(GetCurrentTabRichTextBox(Sender).SelectionFont.Name, GetCurrentTabRichTextBox(Sender).SelectionFont.Size, FontStyle.Regular)

            GetCurrentTabRichTextBox(Sender).SelectionFont = If(GetCurrentTabRichTextBox(Sender).SelectionFont.Italic, RegularFont, ItalicFont)
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub UnderLine(ByRef Sender As TabControl)
            Dim UnderlineFont As New Font(GetCurrentTabRichTextBox(Sender).SelectionFont.Name, GetCurrentTabRichTextBox(Sender).SelectionFont.Size, FontStyle.Underline)
            Dim RegularFont As New Font(GetCurrentTabRichTextBox(Sender).SelectionFont.Name, GetCurrentTabRichTextBox(Sender).SelectionFont.Size, FontStyle.Regular)

            GetCurrentTabRichTextBox(Sender).SelectionFont = If(GetCurrentTabRichTextBox(Sender).SelectionFont.Underline, RegularFont, UnderlineFont)
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub UpperCase(ByRef Sender As TabControl)
            GetCurrentTabRichTextBox(Sender).SelectedText = GetCurrentTabRichTextBox(Sender).SelectedText.ToUpper()
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub LowerCase(ByRef Sender As TabControl)
            GetCurrentTabRichTextBox(Sender).SelectedText = GetCurrentTabRichTextBox(Sender).SelectedText.ToLower()
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub Undo(ByRef Sender As TabControl)
            GetCurrentTabRichTextBox(Sender).Undo()
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub Redo(ByRef Sender As TabControl)
            GetCurrentTabRichTextBox(Sender).Redo()
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub Cut(ByRef Sender As TabControl)
            GetCurrentTabRichTextBox(Sender).Cut()
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub Copy(ByRef Sender As TabControl)
            GetCurrentTabRichTextBox(Sender).Copy()
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub Paste(ByRef Sender As TabControl)
            GetCurrentTabRichTextBox(Sender).Paste()
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub SelectAll(ByRef Sender As TabControl)
            GetCurrentTabRichTextBox(Sender).SelectAll()
            SyntaxHighlighter.HighlightInternalSyntax(GetCurrentTabRichTextBox(Sender))
        End Sub

        ''' <summary>
        ''' Closes all tabs
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <param name="Tabcount">number of tabs</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub RemoveAllTabs(ByRef Sender As TabControl, ByRef Tabcount As Integer)

            For Each Page As TabPage In Sender.TabPages
                Sender.TabPages.Remove(Page)
            Next
            Tabcount = 0
            AddTextTab(Sender, Tabcount)
        End Sub

        ''' <summary>
        ''' removes all tabs except current
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <param name="TabCount">number of tabs</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub RemoveAllTabsButThis(ByRef Sender As TabControl, ByRef TabCount As Integer)
            For Each Page As TabPage In Sender.TabPages
                If Not Page.Name = Sender.SelectedTab.Name Then
                    Sender.TabPages.Remove(Page)
                    TabCount = 0 + 1
                Else
                End If
            Next
        End Sub

        ''' <summary>
        ''' Adds new tab with data
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <param name="Tabcount">number of pages</param>
        ''' <param name="Text">data</param>
        ''' <param name="Title">title</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub AddTextTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Text As String, ByRef Title As String)
            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.HighlightInternalSyntax()
            Body.Text = Text
            Dim NewPage As New TabPage()
            Tabcount += 1
            Body.ForeColor = Color.MidnightBlue

            Dim DocumentText As String = Title & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            Sender.SelectedTab = NewPage
        End Sub

        <System.Runtime.CompilerServices.Extension()>
        Public Sub AddSyntaxTextTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Text As String, ByRef Title As String)
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.HighlightInternalSyntax()
            Body.Text = Text
            Dim NewPage As New TabPage()
            Tabcount += 1
            Body.ForeColor = Color.MidnightBlue

            Dim DocumentText As String = Title & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            Sender.SelectedTab = NewPage
        End Sub

        ''' <summary>
        ''' Adds a current new tab(document) (Empty)
        ''' </summary>
        ''' <param name="Sender">tab control</param>
        ''' <param name="Tabcount">tab count</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub AddTextTab(ByRef Sender As TabControl, ByRef Tabcount As Integer)
            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Tabcount += 1
            Body.ForeColor = Color.MidnightBlue

            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            Sender.SelectedTab = NewPage
        End Sub

        ''' <summary>
        ''' saves data in current tab to file
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <param name="Filename">filename</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub SaveCurrentTab(ByRef Sender As TabControl, ByRef Filename As String)
            If IsTextTab(Sender) = True Then
                GetCurrentTabRichTextBox(Sender).SaveFile(Filename, RichTextBoxStreamType.RichText)
            Else
            End If
        End Sub

        ''' <summary>
        ''' loads file data into current tab
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <param name="Filename">data file</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub OpenFileInCurrent(ByRef Sender As TabControl, ByRef Filename As String)

            GetCurrentTabRichTextBox(Sender).LoadFile(Filename, RichTextBoxStreamType.PlainText)

        End Sub

        ''' <summary>
        ''' opens file in new tab
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <param name="Tabcount">number of tabs</param>
        ''' <param name="Filename">file location</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub OpenFileInNewTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Filename As String)

            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ForeColor = Color.MidnightBlue

            Dim NewPage As New TabPage()
            Tabcount += 1

            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.LoadFile(Filename, RichTextBoxStreamType.PlainText)
            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
        End Sub

#End Region

#Region "Data Tab Functions"

        ''' <summary>
        ''' Returns text held in current tab
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <returns>text</returns>
        <System.Runtime.CompilerServices.Extension()>
        Public Function GetCurrentTabDataGridView(ByRef Sender As TabControl) As DataGridView
            If IsDataTab(Sender) = True Then
                Return CType(Sender.SelectedTab.Controls.Item("Body"), DataGridView)
            Else
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' Returns text held in current tab
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <returns>text</returns>
        <System.Runtime.CompilerServices.Extension()>
        Public Function GetCurrentTabDataTable(ByRef Sender As TabControl) As DataTable
            Dim DGV As New DataGridView
            If IsDataTab(Sender) = True Then
                DGV = CType(Sender.SelectedTab.Controls.Item("Body"), DataGridView)
            Else
            End If
            If DGV IsNot Nothing Then
                Return Convert_DataGridToTable(DGV)
            End If
            Return Nothing
        End Function

        <System.Runtime.CompilerServices.Extension()>
        Public Function Convert_DataGridToTable(ByRef DGV As DataGridView) As DataTable
            'Creating DataTable.
            Dim dt As New DataTable()
            '     Dim DGV As DataGridView = GetCurrentTabDataGridView(TabDocumentBrowser)
            'Adding the Columns.
            For Each column As DataGridViewColumn In DGV.Columns
                dt.Columns.Add(column.HeaderText, column.ValueType)
            Next

            'Adding the Rows.
            For Each row As DataGridViewRow In DGV.Rows
                dt.Rows.Add()
                For Each cell As DataGridViewCell In row.Cells
                    dt.Rows(dt.Rows.Count - 1)(cell.ColumnIndex) = cell.Value.ToString()
                Next
            Next
            Return dt
        End Function

        ''' <summary>
        ''' Checks if tab is web or text or data
        ''' </summary>
        ''' <returns></returns>
        <System.Runtime.CompilerServices.Extension()>
        Public Function IsDataTab(ByRef Sender As TabControl) As Boolean
            IsDataTab = False
            If Sender.SelectedTab.Text = "Data" Then Return True
        End Function

        ''' <summary>
        ''' Adds a current new tab(document)
        ''' </summary>
        ''' <param name="Sender">tab control</param>
        ''' <param name="Tabcount">tab count</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub AddDataTab(ByRef Sender As TabControl, ByRef Tabcount As Integer)
            Dim Body As New DataGridView
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Tabcount += 1
            Body.ForeColor = Color.MidnightBlue

            Dim DocumentText As String = "Data" & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "Data"
            Sender.SelectedTab = NewPage
        End Sub

        ''' <summary>
        ''' Adds a current new tab(document)
        ''' </summary>
        ''' <param name="Sender">tab control</param>
        ''' <param name="Tabcount">tab count</param>
        <System.Runtime.CompilerServices.Extension()>
        Public Sub AddDataTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Data As Object)
            Dim Body As New DataGridView
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Tabcount += 1
            Body.ForeColor = Color.MidnightBlue
            Body.DataSource = Data
            Dim DocumentText As String = "Data" & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "Data"
            Sender.SelectedTab = NewPage
        End Sub

#End Region

#Region "WebTabs"

        ''' <summary>
        ''' Checks if tab is web or text
        ''' </summary>
        ''' <returns></returns>
        <Runtime.CompilerServices.Extension()>
        Public Function IsWebTab(ByRef Sender As TabControl) As Boolean
            IsWebTab = False
            If Sender.SelectedTab.Text = "Browser" Then Return True
        End Function

        ''' <summary>
        ''' adds Webtab
        ''' </summary>
        ''' <param name="Sender">tabcontrol</param>
        ''' <param name="Tabcount">number of pages</param>
        ''' <param name="Source">html source text</param>
        <Runtime.CompilerServices.Extension()>
        Public Sub AddWebSourceTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Source As String)

            Dim Body As New System.Windows.Forms.WebBrowser
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ScriptErrorsSuppressed = True
            Dim NewPage As New TabPage()
            Tabcount += 1

            Body.DocumentText = Source
            Body.Refresh()
            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.Tag = TAB()
            NewPage.Text = "Browser"
            Sender.TabPages.Add(NewPage)
            Sender.SelectedTab = NewPage

        End Sub

#End Region

#Region "Trimming"

        <Runtime.CompilerServices.Extension()>
        Public Function TrimLow(ByVal text As String, Optional ByVal Max As Integer = 0) As String
            IsNull(text)
            If text.Length() < Max Or Max = 0 Then
                text = text.Trim().ToLower()
            End If
            If Max = -1 Then
                text = text.Trim().ToLower()
            End If
            Return text
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function TrimUp(ByVal text As String, Optional ByVal Max As Integer = 0) As String
            IsNull(text)
            If text.Length() < Max Or Max = 0 Then
                text = text.Trim().ToUpper()
            End If
            If Max = -1 Then
                text = text.Trim().ToUpper()
            End If
            Return text
        End Function

#End Region

        <Runtime.CompilerServices.Extension()>
        Public Sub IsNull(ByRef Var As String, Optional ByVal value As String = Nothing)
            If IsNothing(Var) = True Then
                If IsNothing(value) = False Then
                    Var = value
                Else
                    Var = ""
                End If
            End If
        End Sub

        <Runtime.CompilerServices.Extension()>
        Public Sub IsNull(ByRef Var As String(), Optional ByVal value As String() = Nothing)
            If IsNothing(Var) = True Then
                If IsNothing(value) = False Then
                    ReDim Var(value.Length)
                    Var = value
                Else
                    Var = {}
                End If
            End If
        End Sub

        <Runtime.CompilerServices.Extension()>
        Public Sub IsNull(ByRef Var As List(Of String), Optional ByVal value As List(Of String) = Nothing)
            If IsNothing(Var) = True Then
                If IsNothing(value) = False Then
                    Var = value
                Else
                    Var = New List(Of String)
                End If
            End If
        End Sub

        <Runtime.CompilerServices.Extension()>
        Public Function Join(ByVal Input As String(), ByVal deliminator As String) As String
            IsNull(deliminator)
            IsNull(Input)

            Dim Results = ""
            Dim leng = Input.Length() - 1

            For i = 0 To leng
                If i = leng Then
                    Results = Results + Input(i)
                Else
                    Results = Results + Input(i) + deliminator
                End If
            Next

            Return Results
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function Join(ByVal Items As List(Of String), ByVal deliminator As String) As String
            IsNull(Items)
            IsNull(deliminator)

            Dim Results = ""
            Dim leng = Items.Count() - 1

            For i = 0 To leng
                If i = leng Then
                    Results = Results + Items(i)
                Else
                    Results = Results + Items(i) + deliminator
                End If
            Next

            Return Results
        End Function

#Region "Numbers"

        <Runtime.CompilerServices.Extension()>
        Public Function GetGreater(ByVal ints As Integer()) As Integer
            Dim results As Integer = 0

            For Each i In ints
                If i > results Then results = i
            Next

            Return results
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function GetLesser(ByVal ints As Integer()) As Integer
            Dim results As Integer = GetGreater(ints)

            For Each i In ints
                If i < results Then results = i
            Next

            Return results
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function GetGreater(ByVal ints As Double()) As Double
            Dim results As Double = 0

            For Each i In ints
                If i > results Then results = i
            Next

            Return results
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function GetLesser(ByVal ints As Double()) As Double
            Dim results As Double = GetGreater(ints)

            For Each i In ints
                If i < results Then results = i
            Next

            Return results
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function GetGreater(ByVal ints As Decimal()) As Decimal
            Dim results As Decimal = 0

            For Each i In ints
                If i > results Then results = i
            Next

            Return results
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function GetLesser(ByVal ints As Decimal()) As Decimal
            Dim results As Decimal = GetGreater(ints)

            For Each i In ints
                If i < results Then results = i
            Next

            Return results
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function IsGreater(ByVal base As Integer, ByVal int As Integer) As Boolean
            If base < int Then
                Return True
            End If

            Return False
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function IsLesser(ByVal base As Integer, ByVal int As Integer) As Boolean
            If base > int Then
                Return True
            End If

            Return False
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function IsGreater(ByVal base As Double, ByVal int As Double) As Boolean
            If base < int Then
                Return True
            End If

            Return False
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function IsLesser(ByVal base As Double, ByVal int As Double) As Boolean
            If base > int Then
                Return True
            End If

            Return False
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function IsGreater(ByVal base As Decimal, ByVal int As Decimal) As Boolean
            If base < int Then
                Return True
            End If

            Return False
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function IsLesser(ByVal base As Decimal, ByVal int As Decimal) As Boolean
            If base > int Then
                Return True
            End If

            Return False
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function NumbersToWords(ByVal strNumber As String) As String

            Dim NumL = strNumber.Length()

            If NumL < 70 Then
                Dim f1 = ""
                Dim Line1 = ""
                Dim preLine = ""
                Dim r = 0

                If NumL <= 3 Then
                    Dim Results = ToWordsHundreds(strNumber)
                    Return (Results)
                Else

                    Dim Remd = NumL Mod (3)
                    If Remd = 0 Then Remd = 3

                    Dim Nub1 = strNumber.Length - Remd
                    Dim strNumb = Mid(strNumber, Remd + 1, strNumber.Length - Remd + 1)

                    For i = Nub1 To 0 Step -3
                        f1 = Mid(strNumb, (strNumb.Length() - i) + 1, 3)
                        Line1 += ToWordsHundreds(f1) & " " & ToWordsBigNumbers((i - 3) / 3) & " "
                    Next

                    f1 = Mid(strNumber, 1, Remd)
                    preLine = ToWordsHundreds(f1) & " " & ToWordsBigNumbers(Nub1 / 3) & " "
                    Line1 = preLine & Line1

                End If

                Return (Line1)
            Else

                Return ""
            End If
        End Function

        'Private
        <Runtime.CompilerServices.Extension()>
        Public Function ToWordsHundreds(ByVal strNumber As String) As String
            strNumber = Trim(strNumber)

            Dim strFirst = ""
            Dim strSecond = ""
            Dim strThird = ""
            Dim strTeens = ""
            Dim strSingles = ""

            If strNumber.Length = 1 Then

                'singles
                strSingles = ToWordsOnes(strNumber)
                Return (strSingles)
            End If

            If strNumber.Length = 2 Then

                'Teens
                strTeens = ToWordsTeens(strNumber)

                If strTeens > "" Then
                    Return (strTeens)
                End If

                Dim first1 = strNumber(0)
                Dim second1 = strNumber(1)

                '20 through 99
                strFirst = ToWordsTens(first1)
                strSecond = ToWordsOnes(second1)

                Return (strFirst & "-" & strSecond)
            End If

            If strNumber.Length = 3 Then

                'hundreds
                Dim first1 = strNumber(0)
                Dim second1 = strNumber(1)
                Dim third1 = strNumber(2)

                strFirst = ToWordsOnes(first1) & " hundred and "

                If second1 = "1" Then
                    strSecond = ToWordsTeens(second1 & third1)
                    Return (strFirst & strSecond)
                ElseIf second1 = "0" Then
                    strThird = ToWordsOnes(third1)
                    Return (strFirst & strThird)
                Else
                    strSecond = ToWordsTens(second1)
                    strThird = ToWordsOnes(third1)
                    Return (strFirst & strSecond & "-" & strThird)
                End If

            End If

            Return ("")
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function ToWordsOnes(ByVal strNumber As String) As String

            Dim strSingles = ""

            'single digit
            Select Case strNumber
                Case "0"
                    Return "zero"
                Case "1"
                    Return "one"
                Case "2"
                    Return "two"
                Case "3"
                    Return "three"
                Case "4"
                    Return "four"
                Case "5"
                    Return "five"
                Case "6"
                    Return "six"
                Case "7"
                    Return "seven"
                Case "8"
                    Return "eight"
                Case "9"
                    Return "nine"
            End Select

            Return ""
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function ToWordsTens(ByVal strNumber As String) As String

            Dim strTens = ""

            Select Case strNumber
                Case "2"
                    Return "twenty"
                Case "3"
                    Return "thirty"
                Case "4"
                    Return "fourty"
                Case "5"
                    Return "fifty"
                Case "6"
                    Return "sixty"
                Case "7"
                    Return "seventy"
                Case "8"
                    Return "eighty"
                Case "9"
                    Return "ninety"
            End Select

            Return ""
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function ToWordsTeens(ByVal strNumber As String) As String

            Dim strTeens = ""

            '10s through 19
            Select Case strNumber
                Case "10"
                    Return "ten"
                Case "11"
                    Return "eleven"
                Case "12"
                    Return "twelve"
                Case "13"
                    Return "thirteen"
                Case "14"
                    Return "fourteen"
                Case "15"
                    Return "fifteen"
                Case "16"
                    Return "sixteen"
                Case "17"
                    Return "seventeen"
                Case "18"
                    Return "eighteen"
                Case "19"
                    Return "nineteen"
            End Select

            Return ""
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function ToWordsBigNumbers(ByVal strNumber As String) As String
            Dim strBig = ""

            Select Case strNumber
                Case "1"
                    Return "thousand"
                Case "2"
                    Return "million"
                Case "3"
                    Return "billion"
                Case "4"
                    Return "trillion"
                Case "5"
                    Return "quadrillion"
                Case "6"
                    Return "quintillion"
                Case "7"
                    Return "sextillion"
                Case "8"
                    Return "septillion"
                Case "9"
                    Return "octillion"
                Case "10"
                    Return "nonillion"
                Case "11"
                    Return "decillion"
                Case "12"
                    Return "undecillion"
                Case "13"
                    Return "duodecillion"
                Case "14"
                    Return "tredecillion"
                Case "15"
                    Return "quattuordecillion"
                Case "16"
                    Return "quindecillion"
                Case "17"
                    Return "sexdecillion"
                Case "18"
                    Return "septendecillion"
                Case "19"
                    Return "octodecillion"
                Case "20"
                    Return "novemdecillion"
                Case "21"
                    Return "vigintillion"
                Case "22"
                    Return "centillion"
            End Select

            Return ""
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function Remove(ByVal text As String, ByVal toRemove As String) As String
            Return Replace(text, toRemove, "")
        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function ToNumbsOnes(ByVal strNumber As String) As Long
            'single digit
            Select Case strNumber
                Case "zero"
                    Return 0
                Case "one"
                    Return 1
                Case "two"
                    Return 2
                Case "three"
                    Return 3
                Case "four"
                    Return 4
                Case "five"
                    Return 5
                Case "six"
                    Return 6
                Case "seven"
                    Return 7
                Case "eight"
                    Return 8
                Case "nine"
                    Return 9
                Case Else
                    Return -1
            End Select

        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function ToNumbsTens(ByVal strNumber As String) As Long
            Select Case strNumber
                Case "twenty"
                    Return 20
                Case "thirty"
                    Return 30
                Case "fourty"
                    Return 40
                Case "fifty"
                    Return 50
                Case "sixty"
                    Return 60
                Case "seventy"
                    Return 70
                Case "eighty"
                    Return 80
                Case "ninety"
                    Return 90
                Case Else
                    Return -1
            End Select

        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function ToNumbsTeens(ByVal strNumber As String) As Long
            '10s through 19
            Select Case strNumber
                Case "ten"
                    Return 10
                Case "eleven"
                    Return 11
                Case "twelve"
                    Return 12
                Case "thirteen"
                    Return 13
                Case "fourteen"
                    Return 14
                Case "fifteen"
                    Return 15
                Case "sixteen"
                    Return 16
                Case "seventeen"
                    Return 17
                Case "eighteen"
                    Return 18
                Case "nineteen"
                    Return 19
                Case Else
                    Return -1
            End Select

        End Function

        <Runtime.CompilerServices.Extension()>
        Public Function ToNumbsBigNumbers(ByVal intNumber As String) As Decimal

            Select Case intNumber
                Case "thousand"
                    Return 1000D
                Case "million"
                    Return 1000000D
                Case "billion"
                    Return 1000000000D
                Case "trillion"
                    Return 1000000000000D
                Case "quadrillion"
                    Return 1000000000000000D
                Case "quintillion"
                    Return 1000000000000000000D
                Case "sextillion"
                    Return 1000000000000000000000D
                Case "septillion"
                    Return 1000000000000000000000000D
                Case "octillion"
                    Return 1000000000000000000000000000D
                Case Else
                    Return -1
            End Select

        End Function

#End Region

        Public Function AlphabetLC() As String()
            Return {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}
        End Function

        Public Function AlphabetUC() As String()
            Return {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
        End Function

        ''' <summary>
        '''This is a way of opening a dialog form in the center of the main form
        '''without using Me.StartPosition = FormStartPosition.CenterParent
        ''' EG:
        ''' 1 = Dim center = Local.FindCenter(New Point(0, 0), MouseInfo.ScreenSize(), New Size(Width, Height))
        ''' EG
        ''' 2=      Location() = New Point(x, y)
        ''' 2=      Location() = Local.FindCenter(New Point(sender.Left, sender.Top), New Size(sender.Width, sender.Height), Size())
        ''' </summary>
        ''' <param name="formLocation"></param>
        ''' <param name="formSize"></param>
        ''' <param name="dialogSize"></param>
        ''' <returns></returns>
        Public Function FindCenter(ByVal formLocation As Point, ByVal formSize As Size, ByVal dialogSize As Size) As Point
            'This is a way of opening a dialog form in the center of the main form
            'without using Me.StartPosition = FormStartPosition.CenterParent
            Dim centerPoint As New Point

            Dim half_form_X = formSize.Width / 2
            Dim half_form_Y = formSize.Height / 2

            Dim half_Diag_X = dialogSize.Width / 2
            Dim half_Diag_Y = dialogSize.Height / 2

            centerPoint.X = formLocation.X + half_form_X - half_Diag_X
            centerPoint.Y = formLocation.Y + half_form_Y - half_Diag_Y

            Return centerPoint
        End Function

        ''' <summary>
        '''This is a way of opening a dialog form in the center of the main form
        '''without using Me.StartPosition = FormStartPosition.CenterParent
        ''' </summary>
        ''' <param name="formLocation"></param>
        ''' <param name="formSize"></param>
        ''' <returns></returns>
        Public Function FindCenter(ByVal formLocation As Point, ByVal formSize As Size) As Point
            'This is a way of opening a dialog form in the center of the main form
            'without using Me.StartPosition = FormStartPosition.CenterParent
            Dim centerPoint As New Point

            Dim half_form_X = formSize.Width / 2
            Dim half_form_Y = formSize.Height / 2

            centerPoint.X = formLocation.X + half_form_X
            centerPoint.Y = formLocation.Y + half_form_Y

            Return centerPoint
        End Function

        ''' <summary>
        '''This is a way of opening a dialog form in the center of the main form
        '''without using Me.StartPosition = FormStartPosition.CenterParent
        ''' </summary>
        ''' <param name="form"></param>
        ''' <param name="Dialog"></param>
        ''' <returns></returns>
        <Runtime.CompilerServices.Extension()>
        Public Function FindCenter(ByVal form As Control, ByVal Dialog As Control) As Point

            Dim centerPoint As New Point

            Dim half_form_X = form.Width / 2
            Dim half_form_Y = form.Height / 2

            Dim half_Diag_X = Dialog.Width / 2
            Dim half_Diag_Y = Dialog.Height / 2

            centerPoint.X = form.Margin.Left + (half_form_X - half_Diag_X)
            centerPoint.Y = form.Margin.Top + (half_form_Y - half_Diag_Y)

            Return centerPoint
        End Function

    End Module

End Namespace