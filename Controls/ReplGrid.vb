Imports System.Windows.Forms

Public Class ReplGrid
    Public Property My_Cells As New List(Of InteractiveReplNode)
    Public ReadOnly Property CellCount As Integer
        Get
            Return My_Cells.Count
        End Get
    End Property
    Public Sub ADD_NEW_CELL()
        Dim iInteractiveReplNode As New InteractiveReplNode
        My_Cells.Add(iInteractiveReplNode)
        ADD_CELL(iInteractiveReplNode)
    End Sub
    Public Sub ADD_CELL(ByRef iInteractiveReplNode As InteractiveReplNode)
        My_Cells.Add(iInteractiveReplNode)
        iInteractiveReplNode.Dock = DockStyle.Top
        GroupBoxCells.Controls.Add(iInteractiveReplNode)
    End Sub

    Public Sub ExecuteCells()

        Dim iCode As String = ""
        For Each item In My_Cells
            iCode += item.Code + vbNewLine
        Next
        Dim Proj As String = "Public namespace " &
            iCode &
            " End namespace"

    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        ADD_NEW_CELL()
    End Sub
End Class
