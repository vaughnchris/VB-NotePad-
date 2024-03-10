Public Class AboutForm
    Private Sub AboutForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblVersion.Text = Application.ProductVersion
        lblAuthor.Text = My.Resources.ProgramAuthor
        txtDescription.Text = My.Resources.ProgramDescription
    End Sub
End Class