﻿Public Class MainForm
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles NewToolStripMenuItem.Click
        ''first check to see if there is an open file to save
        If DocumentData.DocumentChanged Then
            If TrySave(DocumentData) Then
                ''open a new file
                DocumentData.FileName = ""
                DocumentData.FileContents = ""
                DocumentData.DocumentChanged = False
            End If
        End If
    End Sub
    Private Function TrySave(documentData As FileData) As Boolean
        Dim result As DialogResult = MessageBox.Show("Do you want to save changes to " &
                                                     documentData.FileName & "?", "Save Changes?",
                                                     MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Yes Then
            SaveFile()
        ElseIf result = DialogResult.Cancel Then
            Return False
        End If
        Return True
    End Function
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        If DocumentData.DocumentChanged Then
            If (TrySave(DocumentData)) Then
                If OpenFile() Then
                    Me.Text = My.Resources.MainFormTitle & " - " &
                        DocumentData.FileName
                    Me.txtDoc.Text = DocumentData.FileContents
                End If
            End If
        End If
        If OpenFile() Then
            Me.Text = My.Resources.MainFormTitle & " - " &
                DocumentData.FileName
            Me.txtDoc.Text = DocumentData.FileContents
        End If
    End Sub
    Private Sub txtDoc_TextChanged(sender As Object, e As EventArgs) Handles txtDoc.TextChanged
        DocumentData.FileContents = txtDoc.Text
        DocumentData.DocumentChanged = True
        Me.Text = My.Resources.MainFormTitle & " - " & DocumentData.FileName & "*"
    End Sub
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'set the form title from the resources file and set the document data to a new file
        Me.Text = My.Resources.MainFormTitle & " - New File"
        DocumentData = New FileData("", "")
        txtDoc_FontChanged(txtDoc, Nothing)
    End Sub
    Private Sub txtDoc_MouseWheel(sender As Object, e As MouseEventArgs) Handles txtDoc.MouseWheel
        If e.Delta > 0 Then
            txtDoc.Font = New Font(txtDoc.Font.FontFamily, txtDoc.Font.Size + 1)

        Else
            txtDoc.Font = New Font(txtDoc.Font.FontFamily, txtDoc.Font.Size - 1)
        End If
    End Sub
    Private Sub FontFaceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontFaceToolStripMenuItem.Click
        Dim dlgFont As New FontDialog
        dlgFont.Font = txtDoc.Font
        If dlgFont.ShowDialog = DialogResult.OK Then
            txtDoc.Font = dlgFont.Font
        End If
    End Sub
    Private Sub txtDoc_FontChanged(sender As Object, e As EventArgs) Handles txtDoc.FontChanged
        tssFontFace.Text = txtDoc.Font.FontFamily.Name
        tssFontSize.Text = txtDoc.Font.Size.ToString
    End Sub
    Private Sub AutoWrapTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoWrapTextToolStripMenuItem.Click
        AboutToolStripMenuItem.Checked = Not AutoWrapTextToolStripMenuItem.Checked
        txtDoc.WordWrap = AutoWrapTextToolStripMenuItem.Checked
        tssWordWrap.Text = If(AutoWrapTextToolStripMenuItem.Checked, "On", "Off")
    End Sub
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If SaveFile() Then
            Me.Text = My.Resources.MainFormTitle & " - " & DocumentData.FileName
        End If
    End Sub
    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        If SaveFileAs() Then
            Me.Text = My.Resources.MainFormTitle & " - " & DocumentData.FileName
        End If
    End Sub
    Private Sub File_DropDownOpening(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.DropDownOpening
        If DocumentData.DocumentChanged Then
            SaveToolStripMenuItem.Enabled = True
        Else
            SaveToolStripMenuItem.Enabled = False
        End If
        If DocumentData.FileName = "" Then
            SaveAsToolStripMenuItem.Enabled = False
        Else
            SaveAsToolStripMenuItem.Enabled = True
        End If
    End Sub
End Class