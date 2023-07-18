
Imports System.Windows.Forms
Imports BasicUserControls.Controls.AdvancedTreeViewNodes

Namespace Controls
    Public Class SpydazWebAdvancedTreeView
        Inherits TreeView
        Private m_CurrentButtonNode As ButtonNode = Nothing

        Private m_CurrentComboBoxNode As ComboBoxNode = Nothing

        Private m_CurrentTextBoxNode As TextBoxNode = Nothing

        Public Sub New()
            MyBase.New()
        End Sub

        Protected Overrides Sub OnMouseWheel(ByVal e As MouseEventArgs)
            HideComboBox()
            MyBase.OnMouseWheel(e)
        End Sub

        Protected Overrides Sub OnNodeMouseClick(ByVal e As TreeNodeMouseClickEventArgs)
            If TypeOf e.Node Is ComboBoxNode Then
                Me.m_CurrentComboBoxNode = CType(e.Node, ComboBoxNode)
                Me.Controls.Add(Me.m_CurrentComboBoxNode.ComboBox)
                Me.m_CurrentComboBoxNode.ComboBox.SetBounds(Me.m_CurrentComboBoxNode.Bounds.X - 1, Me.m_CurrentComboBoxNode.Bounds.Y - 2, Me.m_CurrentComboBoxNode.Bounds.Width + 25, Me.m_CurrentComboBoxNode.Bounds.Height)
                AddHandler Me.m_CurrentComboBoxNode.ComboBox.SelectedValueChanged, New EventHandler(AddressOf ComboBox_SelectedValueChanged)
                AddHandler Me.m_CurrentComboBoxNode.ComboBox.DropDownClosed, New EventHandler(AddressOf ComboBox_DropDownClosed)
                Me.m_CurrentComboBoxNode.ComboBox.Show()
                Me.m_CurrentComboBoxNode.ComboBox.DroppedDown = True
            End If
            If TypeOf e.Node Is TextBoxNode Then
                Me.m_CurrentTextBoxNode = CType(e.Node, TextBoxNode)
                Me.Controls.Add(Me.m_CurrentTextBoxNode.TextBox)
                Me.m_CurrentTextBoxNode.TextBox.SetBounds(Me.m_CurrentTextBoxNode.Bounds.X - 1, Me.m_CurrentTextBoxNode.Bounds.Y - 2, Me.m_CurrentTextBoxNode.Bounds.Width + m_CurrentTextBoxNode.TextBox.Text.Length + 100, Me.m_CurrentTextBoxNode.Bounds.Height)
                Me.m_CurrentTextBoxNode.TextBox.Show()
                ' AddHandler Me.m_CurrentTextBoxNode.TextBox.TextChanged, New EventHandler(AddressOf TextBox_TextChanged)
                ' Public Sub TextBox_TextChanged(sender As Object, e As EventArgs)
            End If
            If TypeOf e.Node Is ButtonNode Then
                Me.m_CurrentButtonNode = CType(e.Node, ButtonNode)
                Me.Controls.Add(Me.m_CurrentButtonNode.Button)
                Me.m_CurrentButtonNode.Button.SetBounds(Me.m_CurrentButtonNode.Bounds.X - 1, Me.m_CurrentButtonNode.Bounds.Y - 2, Me.m_CurrentButtonNode.Bounds.Width + 25, Me.m_CurrentButtonNode.Bounds.Height)
                Me.m_CurrentButtonNode.Button.Show()
                ' AddHandler Me.m_CurrentTextBoxNode.TextBox.TextChanged, New EventHandler(AddressOf TextBox_TextChanged)
                ' Public Sub TextBox_TextChanged(sender As Object, e As EventArgs)
            End If
            MyBase.OnNodeMouseClick(e)
        End Sub

        Private Sub ComboBox_DropDownClosed(ByVal sender As Object, ByVal e As EventArgs)
            HideComboBox()
        End Sub

        Private Sub ComboBox_SelectedValueChanged(ByVal sender As Object, ByVal e As EventArgs)
            HideComboBox()
        End Sub

        Private Sub HideComboBox()
            If Me.m_CurrentComboBoxNode IsNot Nothing Then

                Me.m_CurrentComboBoxNode.Text = Me.m_CurrentComboBoxNode.ComboBox.Text
                Me.m_CurrentComboBoxNode.ComboBox.Hide()
                Me.m_CurrentComboBoxNode.ComboBox.DroppedDown = False
                Me.Controls.Remove(Me.m_CurrentComboBoxNode.ComboBox)
                Me.m_CurrentComboBoxNode = Nothing
            End If
        End Sub

    End Class

End Namespace