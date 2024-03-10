Imports System.Drawing.Printing
Imports System.IO

Public Class MainForm

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Dim newTab As TabPageTemplate = New TabPageTemplate()
        DocumentTabs.TabPages.Add(newTab)
        DocumentTabs.SelectedTab = newTab
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim frmAbout As New AboutForm
        frmAbout.ShowDialog()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub DocumentTabs_ControlAdded(sender As Object, e As ControlEventArgs) Handles DocumentTabs.ControlAdded
        Dim tab As TabPageTemplate = CType(e.Control, TabPageTemplate)
        AutoWrapTextToolStripMenuItem_Click(sender, Nothing)
        tab.DocumentTextBox.Focus()
    End Sub

    Private Sub AutoWrapTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoWrapTextToolStripMenuItem.Click
        Dim tempTab As TabPageTemplate = TryCast(DocumentTabs.SelectedTab, TabPageTemplate)
        tempTab.DocumentTextBox.WordWrap = AutoWrapTextToolStripMenuItem.Checked
        tempTab.DocumentTextBox.ScrollBars = If(tempTab.DocumentTextBox.WordWrap, ScrollBars.None, ScrollBars.Vertical)
        tempTab.StatusBar.tssWordWrap.Text = If(tempTab.DocumentTextBox.WordWrap, "On", "Off")
    End Sub
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Dim newTab As TabPageTemplate = CType(Me.DocumentTabs.SelectedTab, TabPageTemplate)
        If SaveFile(newTab) Then
            Me.Text = My.Resources.MainFormTitle & " - " & newTab.FileName
        End If
    End Sub
    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim newTab As TabPageTemplate = New TabPageTemplate(CType(Me.DocumentTabs.SelectedTab, TabPageTemplate))
        If SaveFileAs(newTab) Then
            Me.Text = My.Resources.MainFormTitle & " - " & newTab.FileName
            Me.DocumentTabs.TabPages.Add(newTab)
            Me.DocumentTabs.SelectedTab = newTab
        Else
            newTab.Dispose()
        End If
    End Sub
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim newTab As TabPageTemplate = New TabPageTemplate
        If OpenFile(newTab) Then
            Me.DocumentTabs.TabPages.Add(newTab)
            Me.DocumentTabs.SelectedTab = newTab
        Else
            newTab.Dispose()
        End If
    End Sub
    Private Sub File_DropDownOpening(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.DropDownOpening
        Dim tab As TabPageTemplate = CType(Me.DocumentTabs.SelectedTab, TabPageTemplate)
        If tab IsNot Nothing Then
            If tab.DocumentChanged Then
                SaveToolStripMenuItem.Enabled = True
            Else
                SaveToolStripMenuItem.Enabled = False
            End If
            If tab.FileName = "" Then
                SaveAsToolStripMenuItem.Enabled = False
            Else
                SaveAsToolStripMenuItem.Enabled = True
            End If
            If tab.FileContents = "" Then
                PrintPreviewToolStripMenuItem.Enabled = False
            Else
                PrintPreviewToolStripMenuItem.Enabled = True
            End If
        End If
        If My.Settings.RecentFiles Is Nothing OrElse My.Settings.RecentFiles.Count = 0 Then
            OpenRecentToolStripMenuItem.Enabled = False
        Else
            OpenRecentToolStripMenuItem.Enabled = True
            OpenRecentToolStripMenuItem.DropDownItems.Clear()
            For Each file As String In My.Settings.RecentFiles
                Dim item As New ToolStripMenuItem
                item.Text = file.Substring(file.LastIndexOf("\") + 1)
                item.Tag = file
                AddHandler item.Click, AddressOf OpenRecentFile
                OpenRecentToolStripMenuItem.DropDownItems.Add(item)
            Next
        End If
    End Sub
    Private Sub OpenRecentFile(sender As Object, e As EventArgs)
        Dim item As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        If File.Exists(item.Tag.ToString) Then
            Dim newTab As TabPageTemplate = New TabPageTemplate
            newTab.FileName = item.Tag.ToString
            newTab.FileContents = File.ReadAllText(item.Tag.ToString)
            Me.DocumentTabs.TabPages.Add(newTab)
            Me.DocumentTabs.SelectedTab = newTab
        Else
            MessageBox.Show($"The file {item.Tag.ToString} does not exist, or has been moved.",
                            "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error)
            My.Settings.RecentFiles.Remove(item.Tag.ToString)
            My.Settings.Save()
        End If
    End Sub
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        ''get the current tab
        Dim tab As TabPageTemplate = CType(Me.DocumentTabs.SelectedTab, TabPageTemplate)
        ''if the document has been changed, prompt to save
        If tab.DocumentChanged Then
            Dim result As DialogResult = MessageBox.Show("Do you want to save changes to " & tab.FileName & "?",
                                                         "Save Changes?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                SaveFile(tab)
            ElseIf result = DialogResult.Cancel Then
                Return
            End If
        End If
    End Sub

    Private Sub DocumentTabs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DocumentTabs.SelectedIndexChanged
        Me.Text = $"{My.Resources.MainFormTitle} - {CType(Me.DocumentTabs.SelectedTab, TabPageTemplate).FileName}"
    End Sub
End Class
