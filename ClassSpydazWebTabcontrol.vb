Imports System.Drawing
Imports System.IO
Imports System.Net
Imports System.Windows.Forms
Imports BasicUserControls.Controls

Namespace Controls

    Public Class SpydazWebTabcontrol
        Inherits TabControl

#Region "Speech Control"

        Private Shared Speaker As New Speech.Synthesis.SpeechSynthesizer
        Private Shared Speech_on As Boolean = True

#End Region

#Region "CheckTab types"

        Public Shared Sub _DownloadFile(ByVal url As String,
                                     ByVal localFilePath As String)

            Try
                If Not String.IsNullOrWhiteSpace(url) AndAlso
                    Not String.IsNullOrWhiteSpace(localFilePath) Then

                    Dim fi As New FileInfo(localFilePath)

                    If fi.Exists Then
                        fi.Delete()
                    End If

                    Dim request As WebRequest = WebRequest.Create(url)

                    Using response As HttpWebResponse = DirectCast(request.GetResponse, HttpWebResponse)
                        Using rStream As Stream = response.GetResponseStream
                            Using ms As New MemoryStream
                                rStream.CopyTo(ms)
                                Dim data As Byte() = ms.ToArray()
                                File.WriteAllBytes(fi.FullName, data)
                            End Using
                        End Using
                    End Using
                End If
            Catch ex As Exception
                MessageBox.Show(String.Format("An error occurred:{0}{0}{1}",
                                              vbCrLf, ex.Message),
                                              "Exception Thrown",
                                              MessageBoxButtons.OK,
                                              MessageBoxIcon.Warning)

            End Try

        End Sub

        ''' <summary>
        ''' Adds a current new tab(document)
        ''' </summary>
        ''' <param name="Sender">tab control</param>
        Public Shared Sub AddDataTab(ByRef Sender As TabControl)
            Dim Body As New DataGridView
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()

            Body.ForeColor = Color.Black

            Dim DocumentText As String = "Data"
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
        Public Shared Sub AddDataTab(ByRef Sender As TabControl, ByRef Data As Object)
            Dim Body As New DataGridView
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Body.ForeColor = Color.Black
            Body.DataSource = Data
            Dim DocumentText As String = "Data"
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Sender.TabPages.Add(NewPage)
            NewPage.Text = "Data"
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddEmptySyntaxTextTab(ByRef Sender As TabControl)
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 14, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Body.ForeColor = Color.Black
            Dim DocumentText As String = "Document " & Sender.TabPages.Count + 1
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            ' Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddEmptyTextTab(ByRef Sender As TabControl)
            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Body.ForeColor = Color.Black
            Dim DocumentText As String = "Document " & Sender.TabPages.Count + 1
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            ' Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
        End Sub

        Public Shared Sub AddPopulatedSyntaxTextTab(ByRef Sender As TabControl, ByRef Text As String, ByRef Title As String)
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 14, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill

            Body.Text = Text
            Dim NewPage As New TabPage()

            Body.ForeColor = Color.Black

            Dim DocumentText As String = Title & Sender.TabPages.Count
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = Title

            Sender.SelectedTab = NewPage
            If Speech_on = True Then Speaker.SpeakAsync("Laoding Script Page")
        End Sub

        Public Shared Sub AddPopulatedTextTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, ByRef Text As String, ByRef Title As String)
            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill

            Body.Text = Text
            Dim NewPage As New TabPage()

            Body.ForeColor = Color.Black

            Dim DocumentText As String = Title & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)

            Sender.TabPages.Add(NewPage)
            NewPage.Text = Title
            'Body.ContextMenuStrip = rcMenu
            Sender.SelectedTab = NewPage
            If Speech_on = True Then Speaker.SpeakAsync("Loading Text Page")
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

        Public Shared Sub AddWebTab(ByRef Sender As TabControl)
            Dim Body As New WebBrowser
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Dim NewPage As New TabPage()
            Dim Tabcount = Sender.TabPages.Count
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
            If Speech_on = True Then Speaker.SpeakAsync("Opening Browser Page")
        End Sub

        Public Shared Sub AddWebTab(ByRef Sender As TabControl, ByRef Url As String)
            Dim Body As New WebBrowser
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ScriptErrorsSuppressed = True
            Dim NewPage As New TabPage()
            Dim Tabcount = Sender.TabPages.Count
            Body.Navigate(Url)
            ' Body.ContextMenuStrip = rcMenu
            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.Tag = TAB()
            NewPage.Text = "Browser"
            Sender.TabPages.Add(NewPage)
            Sender.SelectedTab = NewPage
            If Speech_on = True Then Speaker.SpeakAsync("Opening Browser Page")
        End Sub

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
                        GetCurrentTabRichTextBox(sender).Cut()
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
                            GetCurrentTabSyntaxTextBox(sender).Cut()
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
                If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
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

        Public Shared Sub CloseTab_Click(sender As TabControl)

            RemoveTab(sender)

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
                End If
            End If
        End Sub

        Public Shared Sub DownloadFile(ByRef URl As String, ByRef filename As String)
            Using webClient = New WebClient()
                Dim bytes = webClient.DownloadData(URl) 'again variable here
                File.WriteAllBytes(filename, bytes) 'save to desktop or specialfolder.

            End Using
        End Sub

        ''' <summary>
        ''' Name of Download Folder
        ''' </summary>
        ''' <param name="FolderImages"></param>
        Public Shared Sub DownloadImages(ByRef sender As TabControl, ByRef FolderImages As String)
            Dim SW_Browser = GetCurrentTabBrowserControl(sender)
            Dim wcObj As New System.Net.WebClient() 'Create New Web Client Object
            Dim hecImages As HtmlElementCollection = SW_Browser.Document.GetElementsByTagName("img") 'Browse Through HTML Tags
            Dim strWebTitle As String 'Web Page Title
            Dim strPath1 As String 'Folder Path
            Dim strPath2 As String 'Sub Folder Path
            Dim strPath3 As String 'Sub Folder Path
            strWebTitle = SW_Browser.DocumentTitle  'Obtain Web Page Title
            strPath1 = FolderImages  'Create Folder Named Web Page Title
            strPath2 = strPath1 & "\Images\" & strWebTitle 'Create Images Sub Folder
            strPath3 = strPath2 & "\" & "files\"

            Dim diTitle As DirectoryInfo = Directory.CreateDirectory(strPath1)
            Dim diImages As DirectoryInfo = Directory.CreateDirectory(strPath2)
            Dim difiles As DirectoryInfo = Directory.CreateDirectory(strPath3)
            For i As Integer = 0 To hecImages.Count - 1 'Loop Through All IMG Elements Found

                'Download Image(s) To Specified Path

                wcObj.DownloadFile(hecImages(i).GetAttribute("src"), strPath3 & "\" & i.ToString() & ".png")
            Next

        End Sub

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

        Public Shared Function GetCurrentTabBrowserControl(ByRef Sender As TabControl) As WebBrowser
            If IsWebTab(Sender) = True Then
                Return CType(Sender.SelectedTab.Controls.Item("Body"), WebBrowser)
            Else
            End If
            Return Nothing
        End Function

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

        Public Shared Sub GetHtmlSource(ByRef sender As TabControl)
            If IsWebTab(sender) = True Then
                '   If Speech_on = True Then Speaker.SpeakAsync("Getting Source code")
                Dim SwBrowser = GetCurrentTabBrowserControl(sender)
                Dim _source = SwBrowser.DocumentText
                AddPopulatedSyntaxTextTab(sender, _source, "SCRIPT")
            Else
            End If
        End Sub

        Public Shared Function GetPicture(ByRef Url As String) As Image
            Using webClient = New WebClient()
                Dim image_stream As New MemoryStream(webClient.DownloadData(Url))
                Return Image.FromStream(image_stream)
            End Using
        End Function

        Public Shared Function GetRichTextBox(ByRef Sender As TabPage) As RichTextBox
            Try
                Return CType(Sender.Controls.Item("Body"), RichTextBox)
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Shared Function GetSyntaxTextBox(ByRef Sender As TabPage) As SpydazWebTextBox
            Try
                Return CType(Sender.Controls.Item("Body"), SpydazWebTextBox)
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Shared Function GetText(sender As TabControl) As String
            Dim SW_Browser = GetCurrentTabBrowserControl(sender)
            Return SW_Browser.Document.Body.InnerText
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

            End If
            Return WebLinks
        End Function

        Public Shared Function GrabUrlImages(ByRef sender As TabControl) As List(Of String)
            Dim LinkStr As New List(Of String)
            Dim SwBrowser = GetCurrentTabBrowserControl(sender)
            If SwBrowser IsNot Nothing Then
                Dim images As HtmlElementCollection = SwBrowser.Document.Images
                For Each link As HtmlElement In images
                    LinkStr.Add(link.GetAttribute("img"))
                Next
            End If
            'Dim links As HtmlElementCollection = SW_Browser.Document.Links

            Return LinkStr
        End Function

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
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
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
                            If GetCurrentTabSyntaxTextBox(sender).CanPaste(PictureFormat) Then
                                GetCurrentTabSyntaxTextBox(sender).Paste(PictureFormat)
                            End If
                            Clipboard.Clear()
                            Clipboard.SetText(cboard)
                        End If
                    Catch ex As Exception
                    End Try
                Else
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
            Else
            End If

        End Sub

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
            Else
            End If

        End Sub

        Public Shared Function IsAvatarTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsAvatarTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "AVATAR" Then Return True
            End If

        End Function

        ''' <summary>
        ''' Checks if tab is web or text or data
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function IsDataTab(ByRef Sender As TabControl) As Boolean
            IsDataTab = False
            If Sender.SelectedTab.Text = "Data" Then Return True
        End Function

        Public Shared Function IsDeviceTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsDeviceTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "DEVICE" Then Return True
            End If

        End Function

        Public Shared Function IsPluginTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsPluginTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "PLUGIN" Then Return True
            End If

        End Function

        Public Shared Function IsScriptTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsScriptTab = False

            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If GetCurrentTabSyntaxTextBox(TabDocumentBrowser) IsNot Nothing Or GetCurrentTabRichTextBox(TabDocumentBrowser) IsNot Nothing Then Return True

            End If
        End Function

        Public Shared Function IsSnippitTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsSnippitTab = False

            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If GetCurrentTabSyntaxTextBox(TabDocumentBrowser) IsNot Nothing Or GetCurrentTabRichTextBox(TabDocumentBrowser) IsNot Nothing Then Return True

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
                If GetCurrentTabSyntaxTextBox(TabDocumentBrowser) IsNot Nothing Or GetCurrentTabRichTextBox(TabDocumentBrowser) IsNot Nothing Then Return True
            End If
        End Function

        Public Shared Function IsUpdateTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsUpdateTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "UPDATE" Then Return True
            End If

        End Function

#End Region

#Region "All tabs"

        Public Shared Function IsWebTab(ByRef TabDocumentBrowser As TabControl) As Boolean
            IsWebTab = False
            If TabDocumentBrowser.SelectedTab IsNot Nothing Then
                If TabDocumentBrowser.SelectedTab.Text = "Browser" Then Return True
            End If

        End Function

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

        Public Shared Sub LoadCheckedListBox(ByRef Sender As CheckedListBox, ByRef UrlList As List(Of String))
            For Each item In UrlList
                Sender.Items.Add(item)
            Next

        End Sub

        Public Shared Function LoadFile(ByRef Sender As TabControl, ByRef Filter As String) As String
            Dim Scriptfile As String = ""
            Dim txt As String = ""
            Select Case Filter
                Case "Json"
                    Filter = "Json files (*.Json)|*.Json"
                    If Speech_on = True Then Speaker.SpeakAsync("Open Json file?")
                Case = "Html"
                    Filter = "Html Source files (*.htm)|*.htm"
                    If Speech_on = True Then Speaker.SpeakAsync("Open Html file?")
                Case "Aib"
                    Filter = "BrainScript files (*.Aib)|*.Aib"
                    If Speech_on = True Then Speaker.SpeakAsync("Open BrainScript?")
                Case "Text"
                    Filter = "Text files (*.txt)|*.txt"
                    If Speech_on = True Then Speaker.SpeakAsync("Open Text file?")
                Case "VB"
                    Filter = "VB files (*.vb)|*.vb"
                    If Speech_on = True Then Speaker.SpeakAsync("Open Visual Basic Script file?")
                Case "C#"
                    Filter = "C# files (*.cs)|*.cs"
                    If Speech_on = True Then Speaker.SpeakAsync("Open C Sharp Files?")
                Case "All"
                    Filter = "All files (*.*)|*.*"
                    If Speech_on = True Then Speaker.SpeakAsync("Open Any File?")
                Case "CSV"
                    Filter = "comma seperated data files (*.csv)|*.csv"
                    If Speech_on = True Then Speaker.SpeakAsync("Open CSV file?")
                Case "XML"
                    Filter = "X Markup Language (*.xml)|*.xml"
                    If Speech_on = True Then Speaker.SpeakAsync("Open X M L file?")
                Case "JavaScript"
                    Filter = "Javascript Language (*.js)|*.js"
                    If Speech_on = True Then Speaker.SpeakAsync("Open Java file?")
            End Select
            Dim Ofile As New OpenFileDialog
            With Ofile
                .Filter = Filter
                If (.ShowDialog() = DialogResult.OK) Then
                    Scriptfile = .FileName
                End If
            End With
            If Speech_on = True Then Speaker.SpeakAsync("OK!, Opening File!")
            Try
                txt = My.Computer.FileSystem.ReadAllText(Scriptfile)

                AddPopulatedSyntaxTextTab(Sender, txt, Scriptfile)
            Catch ex As Exception
                If Speech_on = True Then Speaker.SpeakAsync("Open file Error!")
                '  MsgBox(ex.ToString,, "Error")
            End Try
            Return txt
        End Function

        Public Shared Sub LoadListbox(ByRef Sender As ListBox, ByRef UrlList As List(Of String))
            For Each item In UrlList
                Sender.Items.Add(item)
            Next
        End Sub

        Public Shared Sub LoadPictureBox(ByRef Sender As PictureBox, ByRef UrlList As List(Of Image))
            Speaker.Speak("Selecting Picture")
            Sender.Image = UrlList(0)
        End Sub

        Public Shared Sub LoadTextBox(ByRef Sender As TextBox, ByRef _source As String)
            Sender.Text = _source
        End Sub

        Public Shared Sub Navigate(ByRef sender As TabControl, ByRef Text As String)
            AddWebTab(sender, Text)
        End Sub

        Public Shared Sub OpenHtmlSource(ByRef sender As TabControl, ByVal _source As String)
            If IsWebTab(sender) = True Then
                Dim SwBrowser = GetCurrentTabBrowserControl(sender)
                SwBrowser.DocumentText = _source
            Else
            End If

        End Sub

        Public Shared Sub OpenNewDocumentTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, Filename As String)
            Dim Body As New RichTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ForeColor = Color.Black

            Dim NewPage As New TabPage()
            Tabcount += 1

            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.LoadFile(Filename, RichTextBoxStreamType.PlainText)
            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            If Speech_on = True Then Speaker.SpeakAsync("Opening Text Page")
        End Sub

        Public Shared Sub OpenNewSyntaxDocumentTab(ByRef Sender As TabControl, ByRef Tabcount As Integer, Filename As String)
            Dim Body As New SpydazWebTextBox
            Body.Font = New Font(Sender.Font.Name, 12, FontStyle.Regular)
            Body.Name = NameOf(Body)
            Body.Dock = DockStyle.Fill
            Body.ForeColor = Color.Black

            Dim NewPage As New TabPage()
            Tabcount += 1

            Dim DocumentText As String = "Document " & Tabcount
            NewPage.Name = DocumentText
            NewPage.Text = DocumentText
            NewPage.Controls.Add(Body)
            Body.LoadFile(Filename, RichTextBoxStreamType.PlainText)
            Sender.TabPages.Add(NewPage)
            NewPage.Text = "TextBox"
            If Speech_on = True Then Speaker.SpeakAsync("Opening Script Page")
        End Sub

        Public Shared Function OpenPicfolder(ByRef Path As String) As List(Of String)
            Dim gifEntries As String() = Directory.GetFiles(Path, "*.gif")
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
            For Each item In gifEntries
                str.Add(item)
            Next
            For Each item In PngEntries
                str.Add(item)
            Next
            Return str
        End Function

        Public Shared Sub OpenSyntaxTextCurrentTab(ByRef Sender As TabControl, ByRef Filename As String)
            GetCurrentTabSyntaxTextBox(Sender).LoadFile(Filename, RichTextBoxStreamType.PlainText)
        End Sub

        Public Shared Sub OpenTextCurrentTab(ByRef Sender As TabControl, ByRef Filename As String)
            GetCurrentTabRichTextBox(Sender).LoadFile(Filename, RichTextBoxStreamType.PlainText)
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
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then
                            If GetCurrentTabSyntaxTextBox(sender).SelectedText = "" Then
                                'Speaks the entire document (if no text is selected)
                                Dim SAPI2
                                SAPI2 = CreateObject("SAPI.spvoice")
                                SAPI2.Speak(GetCurrentTabSyntaxTextBox(sender).Text)
                                SAPI2.Speak("Status : " & "Finished speaking document")
                            Else
                                'Speaks the selected text
                                Dim SAPI
                                SAPI = CreateObject("SAPI.spvoice")
                                SAPI.Speak(GetCurrentTabSyntaxTextBox(sender).SelectedText)
                                SAPI.Speak("Status : " & "Finished speaking document")
                            End If
                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
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
                End If
            End If

        End Sub

        Public Shared Sub RemoveAllTabs(ByRef Sender As TabControl)

            For Each Page As TabPage In Sender.TabPages

                Sender.TabPages.Remove(Page)

            Next

        End Sub

        Public Shared Sub RemoveAllTabsButThis(ByRef Sender As TabControl)
            For Each Page As TabPage In Sender.TabPages

                If Page IsNot Sender.SelectedTab Then
                    Sender.TabPages.Remove(Page)

                End If

            Next
        End Sub

        Public Shared Sub RemoveTab(ByRef Sender As TabControl)
            If Sender.SelectedTab.Text <> "HOME" Or Sender.SelectedTab.Text <> "COMPILER" Then
                If Sender.TabPages.Count > 0 Then
                    Sender.TabPages.Remove(Sender.SelectedTab)
                Else
                End If
            End If
        End Sub

#Region "Tab Close"

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
            Else
            End If
        End Sub

        Public Shared Sub SaveScript(ByRef Script As String)
            Try
                Dim ScriptFile As String = ""
                Dim S As New SaveFileDialog
                With S
                    .Filter = "All Files (*.*)|*.*"

                    If Speech_on = True Then Speaker.SpeakAsync("Choose file Name")
                    If (.ShowDialog() = DialogResult.OK) Then
                        ScriptFile = .FileName
                    End If
                End With
                My.Computer.FileSystem.WriteAllText(ScriptFile, Script, True)
                If Speech_on = True Then Speaker.SpeakAsync("Saving file ")
            Catch ex As Exception

                If Speech_on = True Then Speaker.SpeakAsync("Save file Error!")
            End Try
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
                    .Filter = "Enter Filename : (*.*)|*.*"
                    If (.ShowDialog() = DialogResult.OK) Then
                        ScriptFile = .FileName
                    End If
                End With
                My.Computer.FileSystem.WriteAllText(ScriptFile, alltext, True)
            Catch ex As Exception

            End Try

        End Sub

        Public Shared Sub SaveWebpage(ByRef Sender As TabControl)
            If IsWebTab(Sender) = True Then

                Dim SwBrowser = GetCurrentTabBrowserControl(Sender)
                Dim alltext As String = SwBrowser.DocumentText
                Try
                    Dim ScriptFile As String = ""
                    Dim S As New SaveFileDialog
                    With S
                        .Filter = "Enter Filename : (*.*)|*.*"
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
                        GetCurrentTabRichTextBox(sender).SelectedText = GetCurrentTabRichTextBox(sender).SelectedText
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
                            GetCurrentTabSyntaxTextBox(sender).SelectedText = GetCurrentTabRichTextBox(sender).SelectedText
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

        Public Shared Sub ToLower(ByRef sender As TabControl)
            If GetCurrentTabRichTextBox(sender) IsNot Nothing Then
                Try
                    If GetCurrentTabRichTextBox(sender).SelectedText IsNot Nothing Then

                        GetCurrentTabRichTextBox(sender).SelectedText.ToLower()
                    End If
                Catch ex As Exception
                    '   MsgBox(ex.Message)
                End Try
            Else
                If GetCurrentTabSyntaxTextBox(sender) IsNot Nothing Then
                    Try
                        If GetCurrentTabSyntaxTextBox(sender).SelectedText IsNot Nothing Then

                            GetCurrentTabSyntaxTextBox(sender).SelectedText.ToLower()
                        End If
                    Catch ex As Exception
                        '   MsgBox(ex.Message)
                    End Try
                End If
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
                End If
            End If

        End Sub

        Public Sub InsertSyntaxTextTab(ByRef Sender As TabControl, ByRef Text As String)
            AddPopulatedSyntaxTextTab(Sender, Text, "SCRIPT")
        End Sub

        ''' <summary>
        ''' Used to insert Scripts etc to the TextBox
        ''' </summary>
        ''' <param name="Text"></param>
        Public Sub InsertTextTab(ByRef Sender As TabControl, ByRef Text As String)
            AddPopulatedTextTab(Sender, Sender.TabPages.Count, Text, "TEXT")
        End Sub

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

        Private Shared Sub browser_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs)

            AddHandler CType(sender, WebBrowser).Document.Window.Error,
            AddressOf Window_Error

        End Sub

        ' Hides script errors without hiding other dialog boxes.
        Private Shared Sub SuppressScriptErrorsOnly(ByVal browser As WebBrowser)

            ' Ensure that ScriptErrorsSuppressed is set to false.
            browser.ScriptErrorsSuppressed = True

            ' Handle DocumentCompleted to gain access to the Document object.
            AddHandler browser.DocumentCompleted,
            AddressOf browser_DocumentCompleted

        End Sub

        Private Shared Sub Window_Error(ByVal sender As Object, ByVal e As HtmlElementErrorEventArgs)

            ' Ignore the error and suppress the error dialog box.
            e.Handled = True

        End Sub

        Private Sub SpydazWebTabcontrol_DrawItem(sender As Object, e As DrawItemEventArgs) Handles Me.DrawItem
            'Set the Mode of Drawing as Owner Drawn
            DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
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

        Private Sub SpydazWebTabcontrol_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
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

        Private Sub SpydazWebTabcontrol_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Me.SelectedIndexChanged
            If Me.TabPages.Count = 0 Then
                Me.Visible = False
            End If
        End Sub

#End Region

#End Region

    End Class

End Namespace

Public Class WebSite

    Public imageLinks As List(Of String)

    Public images As List(Of Image)

    Public links As List(Of String)

    Public Loaded As Boolean

    Public TextStr As String

    Public Url As String

    Public WebPage As String

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

    Public Enum CrawlerOption
        Html
        SourceCode
        Text
        Images
        HtmlLinks
        ImageLinks
        Browser
    End Enum

    Public Shared Property Speaker As Object

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
            If Browser.Document.Body IsNot Nothing Then
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

    Public Shared Sub DownloadHtmlLinksBackGround(ByVal url As String)
        Dim SwBrowser = New WebBrowser
        SwBrowser.ScriptErrorsSuppressed = True
        SwBrowser.Navigate(url)
        AddHandler SwBrowser.DocumentCompleted, AddressOf browserLinksLoad_DocumentCompleted
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

    ''' <summary>
    ''' Download Image from url http://www.spydazweb.co.uk/image.gif
    ''' </summary>
    ''' <param name="Url"></param>
    ''' <returns>image</returns>
    Public Shared Function DownloadImage(ByVal Url As String) As Image
        Try
            Return SpydazWebTabcontrol.GetPicture(Url)
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

    Public Shared Sub DownloadImagesBackGround(ByVal url As String)
        Dim SwBrowser = New WebBrowser
        SwBrowser.ScriptErrorsSuppressed = True
        SwBrowser.Navigate(url)
        AddHandler SwBrowser.DocumentCompleted, AddressOf browserImageLoad_DocumentCompleted
    End Sub

    Public Shared Sub DownloadImagesToFiles(ByRef Url As String)
        Try
            DownloadImages(Url)
        Catch ex As Exception

        End Try

    End Sub

    Public Shared Function DownloadImagesToList(ByRef Url As String) As List(Of Image)
        DownloadImagesToList = New List(Of Image)
        Dim lst = DownloadImageLinks(Url)
        Parallel.ForEach(lst, Sub(item)
                                  DownloadImagesToList.Add(DownloadImage(item))
                              End Sub)
    End Function

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
            If SwBrowser.Document.Body IsNot Nothing Then
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

#Region "Download_Images BackGround"

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

End Class