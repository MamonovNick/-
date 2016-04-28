Public Class MainForm
    Private Sub МонетыToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles МонетыToolStripMenuItem.Click
        G_coins.Show()
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        Me.Close()
    End Sub
End Class