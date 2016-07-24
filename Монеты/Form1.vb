Public Class MainForm
    Private OperationBool As Boolean = False

    Private Sub МонетыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles МонетыToolStripMenuItem.Click
        G_coins.Show()
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ПодразделенияToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПодразделенияToolStripMenuItem.Click
        Form4.Show()
    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        'Only for debugging
        'For tech only 
        'Need to be deleted in release ver
        'Form7.Show()
        'Form7.Activate()
    End Sub

    Private Sub КонтрагентыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles КонтрагентыToolStripMenuItem.Click
        Form5.Show()
    End Sub

    Private Sub ВидыВалютToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВидыВалютToolStripMenuItem.Click
        Form6.Show()
    End Sub

    Private Sub ОПрограммеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        AboutBox1.Show()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If OperationBool Then
            ToolStripButton3.Visible = False
            ToolStripButton4.Visible = False
        Else
            ToolStripButton3.Visible = True
            ToolStripButton4.Visible = True
        End If
        OperationBool = Not OperationBool
    End Sub
End Class