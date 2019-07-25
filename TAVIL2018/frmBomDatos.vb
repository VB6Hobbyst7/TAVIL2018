Imports System.Windows.Forms

Public Class frmBomDatos
    Private Sub frmBomDatos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "LISTAD DE PIEZAS - v" & cfg._appversion
    End Sub

    Private Sub frmBomDatos_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        frmBo = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
