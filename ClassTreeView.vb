
Imports System.Windows.Forms

Namespace Controls

    Namespace AdvancedTreeViewNodes

        Public Enum AdvancedTreeViewNodeType
            ComboBox
            TextBox
            Button
        End Enum

        ''' <summary>
        ''' Still need work to map the button press handling
        ''' </summary>
        Public Class ButtonNode
            Inherits Windows.Forms.TreeNode
            Private m_ComboBox As Button = New Button()

            Public Sub New()
            End Sub

            Public Sub New(text As String)
                MyBase.New()
                Button.Text = text
            End Sub

            Public Sub New(text As String, children() As Windows.Forms.TreeNode)
                MyBase.New(text, children)
                Button.Text = text
            End Sub

            Public Sub New(text As String, imageIndex As Integer, selectedImageIndex As Integer)
                MyBase.New(text, imageIndex, selectedImageIndex)
                Button.Text = text
            End Sub

            Public Sub New(text As String, imageIndex As Integer, selectedImageIndex As Integer, children() As Windows.Forms.TreeNode)
                MyBase.New(text, imageIndex, selectedImageIndex, children)
                Button.Text = text
            End Sub

            Protected Sub New(serializationInfo As Runtime.Serialization.SerializationInfo, context As Runtime.Serialization.StreamingContext)
                MyBase.New(serializationInfo, context)
            End Sub

            Public Property Button As Button
                Get

                    Return Me.m_ComboBox
                End Get
                Set(ByVal value As Button)
                    Me.m_ComboBox = value

                End Set
            End Property

            Public ReadOnly Property Type As AdvancedTreeViewNodeType
                Get
                    Return AdvancedTreeViewNodeType.Button
                End Get
            End Property

        End Class

        Public Class ComboBoxNode
            Inherits Windows.Forms.TreeNode
            Private m_ComboBox As ComboBox = New ComboBox()

            Public Sub New()
            End Sub

            Public Sub New(text As String)
                MyBase.New(text)
            End Sub

            Public Sub New(text As String, children() As Windows.Forms.TreeNode)
                MyBase.New(text, children)
            End Sub

            Public Sub New(text As String, imageIndex As Integer, selectedImageIndex As Integer)
                MyBase.New(text, imageIndex, selectedImageIndex)
            End Sub

            Public Sub New(text As String, imageIndex As Integer, selectedImageIndex As Integer, children() As Windows.Forms.TreeNode)
                MyBase.New(text, imageIndex, selectedImageIndex, children)
            End Sub

            Protected Sub New(serializationInfo As Runtime.Serialization.SerializationInfo, context As Runtime.Serialization.StreamingContext)
                MyBase.New(serializationInfo, context)
            End Sub

            Public Property ComboBox As ComboBox
                Get

                    Me.m_ComboBox.DropDownStyle = ComboBoxStyle.DropDown
                    Return Me.m_ComboBox
                End Get
                Set(ByVal value As ComboBox)
                    Me.m_ComboBox = value

                End Set
            End Property

            Public ReadOnly Property Type As AdvancedTreeViewNodeType
                Get
                    Return AdvancedTreeViewNodeType.ComboBox
                End Get
            End Property

        End Class

        Public Class TextBoxNode
            Inherits Windows.Forms.TreeNode
            Private m_ComboBox As TextBox = New TextBox()

            Public Sub New()
            End Sub

            Public Sub New(text As String)
                MyBase.New()
                TextBox.Text = text
            End Sub

            Public Sub New(text As String, children() As Windows.Forms.TreeNode)
                MyBase.New(text, children)

            End Sub

            Public Sub New(text As String, imageIndex As Integer, selectedImageIndex As Integer)
                MyBase.New(text, imageIndex, selectedImageIndex)

            End Sub

            Public Sub New(text As String, imageIndex As Integer, selectedImageIndex As Integer, children() As Windows.Forms.TreeNode)
                MyBase.New(text, imageIndex, selectedImageIndex, children)

            End Sub

            Protected Sub New(serializationInfo As Runtime.Serialization.SerializationInfo, context As Runtime.Serialization.StreamingContext)
                MyBase.New(serializationInfo, context)
            End Sub

            Public Property TextBox As TextBox
                Get

                    Return Me.m_ComboBox
                End Get
                Set(ByVal value As TextBox)
                    Me.m_ComboBox = value

                End Set
            End Property

            Public ReadOnly Property Type As AdvancedTreeViewNodeType
                Get
                    Return AdvancedTreeViewNodeType.TextBox
                End Get
            End Property

        End Class

    End Namespace

End Namespace