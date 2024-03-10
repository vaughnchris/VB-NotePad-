Imports System.Drawing.Printing
Imports System.IO

Public Class MainForm
    Private _DocumentData As FileData = New FileData("", "")
    Private Property DocumentData As FileData
        Get
            Return _DocumentData
        End Get
        Set
            _DocumentData = Value
            RefreshState()
        End Set
    End Property
#Region "Printer Operations"
    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim strContents As String = txtDoc.Text
        Dim charactersOnPage As Integer = 0
        Dim linesPerPage As Integer = 0

        ' Sets the value of charactersOnPage to the number of characters  
        ' of stringToPrint that will fit within the bounds of the page.
        e.Graphics.MeasureString(strContents, Me.Font, e.MarginBounds.Size,
            StringFormat.GenericTypographic, charactersOnPage, linesPerPage)

        ' Draws the string within the bounds of the page
        e.Graphics.DrawString(strContents, Me.Font, Brushes.Black,
            e.MarginBounds, StringFormat.GenericTypographic)

        ' Remove the portion of the string that has been printed.
        strContents = strContents.Substring(charactersOnPage)

        ' Check to see if more pages are to be printed.
        e.HasMorePages = strContents.Length > 0
    End Sub
#End Region
#Region "Form Events"
    ''' <summary>
    ''' Sets the form title and the document data to a 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'set the form title 
        Me.Text = My.Resources.MainFormTitle & " - Unsaved File"
        txtDoc.Font = New Font(My.Resources.InitialFontFace,
                CInt(My.Resources.InitialFontSize))
        txtDoc_FontChanged(txtDoc, Nothing)
        txtDoc.Focus()
        txtDoc.Select(0, 0)
    End Sub
    ''' <summary>
    ''' Refresh the document data as the form is updated
    ''' keeping them synchronized
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub txtDoc_TextChanged(sender As Object, e As EventArgs)
        Dim data As FileData = Me.DocumentData
        data.FileContents = txtDoc.Text
        data.DocumentChanged = True
        Me.DocumentData = data
        If DocumentData.FileContents.Length > 1 AndAlso
            (DocumentData.FileContents(DocumentData.FileContents.Length - 2) = "."c OrElse
             DocumentData.FileContents(DocumentData.FileContents.Length - 2) = "?"c OrElse
             DocumentData.FileContents(DocumentData.FileContents.Length - 2) = "!"c) AndAlso
             DocumentData.FileContents(DocumentData.FileContents.Length - 1) = " "c Then
            tssSentanceCount.Text = DocumentData.SentanceCount.ToString
        End If
        If DocumentData.FileContents.Length > 0 AndAlso
            DocumentData.FileContents(DocumentData.FileContents.Length - 1) = " "c Then
            tssWordCount.Text = DocumentData.WordCount.ToString
        End If
        Dim strDocumentTitle As String = If(DocumentData.FileName = "", "Unsaved File", DocumentData.FileName)
        Me.Text = My.Resources.MainFormTitle & " - " & strDocumentTitle & "*"
    End Sub

    Private Sub txtDoc_MouseWheel(sender As Object, e As MouseEventArgs)
        If e.Delta > 0 AndAlso
            txtDoc.Font.Size + 1 < System.Single.MaxValue Then
            txtDoc.Font = New Font(txtDoc.Font.FontFamily, txtDoc.Font.Size + 1)

        Else
            If txtDoc.Font.Size - 1 > 1 Then
                txtDoc.Font = New Font(txtDoc.Font.FontFamily, txtDoc.Font.Size - 1)
            End If
        End If
    End Sub
    Private Sub txtDoc_FontChanged(sender As Object, e As EventArgs)
        tssFontFace.Text = txtDoc.Font.FontFamily.Name
        tssFontSize.Text = txtDoc.Font.Size.ToString & "pt"
    End Sub

#End Region
#Region "Menu Events"
#Region "File Menu"
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles NewToolStripMenuItem.Click
        ''first check to see if there is an open file to save
        If DocumentData.DocumentChanged Then
            If TrySave(DocumentData) Then
                ''clear the document data and the text box
                ClearData()
            End If
        End If
    End Sub
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        If DocumentData.DocumentChanged Then
            If TrySave(DocumentData) Then
                If OpenFile(DocumentData) Then
                    UpdateRecentFiles(DocumentData.FileName)
                    UpdateFormTitleAndContent(DocumentData)
                    UpdateStatistics(DocumentData)
                End If
            End If
        Else
            OpenFile(DocumentData)
            UpdateFormTitleAndContent(DocumentData)
        End If
    End Sub

    Private Sub UpdateRecentFiles(fileName As String)
        If My.Settings.RecentFiles Is Nothing Then
            My.Settings.RecentFiles = New Specialized.StringCollection
        End If
        If My.Settings.RecentFiles.Count >= 10 Then
            My.Settings.RecentFiles.RemoveAt(0)
        End If
        My.Settings.RecentFiles.Add(fileName)
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

    Private Sub UpdateFormTitleAndContent(data As FileData)
        Dim strDocumentTitle As String = If(data.FileName = "", "Unsaved File", data.FileName)
        Me.Text = $"{My.Resources.MainFormTitle} - {strDocumentTitle}"
        Me.txtDoc.Text = data.FileContents
    End Sub

    Private Sub UpdateStatistics(data As FileData)
        Me.tssWordCount.Text = data.WordCount.ToString()
        Me.tssSentanceCount.Text = data.SentanceCount.ToString()
    End Sub
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If SaveFile(DocumentData) Then
            Me.Text = My.Resources.MainFormTitle & " - " & DocumentData.FileName
        End If
    End Sub
    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        If SaveFileAs(DocumentData) Then
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
        If DocumentData.FileContents = "" Then
            PrintPreviewToolStripMenuItem.Enabled = False
        Else
            PrintPreviewToolStripMenuItem.Enabled = True
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
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If DocumentData.DocumentChanged Then
            If TrySave(DocumentData) Then
                Me.Close()
            End If
        Else
            Me.Close()
        End If
    End Sub
    Private Sub PrintPreviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        Dim strContents As String = txtDoc.Text
        PrintDocument1.DocumentName = DocumentData.FileName
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub
    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog = DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub
#End Region
#Region "Edit Menu"
    Private Sub DuplicateWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DuplicateWindowToolStripMenuItem.Click
        Dim newForm As New MainForm
        Dim newDocument As New FileData("", txtDoc.Text)
        newForm.DocumentData = newDocument
        newForm.Show()
    End Sub
#End Region
#Region "Settings Menu"
    Private Sub FontFaceToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles FontFaceToolStripMenuItem.Click, tssFontFace.Click
        Dim dlgFont As New FontDialog
        dlgFont.Font = txtDoc.Font
        If dlgFont.ShowDialog = DialogResult.OK Then
            txtDoc.Font = dlgFont.Font
        End If
    End Sub
    Private Sub AutoWrapTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoWrapTextToolStripMenuItem.Click
        AboutToolStripMenuItem.Checked = Not AutoWrapTextToolStripMenuItem.Checked
        txtDoc.WordWrap = AutoWrapTextToolStripMenuItem.Checked
        tssWordWrap.Text = If(AutoWrapTextToolStripMenuItem.Checked, "On", "Off")
    End Sub
#End Region
#Region "Help Menu"

#End Region
#End Region
#Region "User Methods"
    Private Sub RefreshState()
        txtDoc.Text = DocumentData.FileContents
        Dim strDocumentTitle As String = If(DocumentData.FileName = "", "Unsaved File", DocumentData.FileName)
        Me.Text = $"{My.Resources.MainFormTitle} - {strDocumentTitle}"
    End Sub
    Private Sub ClearData()
        Dim data As FileData = Me.DocumentData
        With data
            .DocumentChanged = False
            .FileName = ""
            .FileContents = ""
        End With
        Me.DocumentData = data
        Me.Text = My.Resources.MainFormTitle & " - New File"
        Me.txtDoc.Text = ""
    End Sub
    Private Function TrySave(documentData As FileData) As Boolean
        Dim result As DialogResult = MessageBox.Show("Do you want to save changes to " &
                                                     documentData.FileName & "?", "Save Changes?",
                                                     MessageBoxButtons.YesNoCancel)
        If result = DialogResult.Yes Then
            SaveFile(documentData)
        ElseIf result = DialogResult.Cancel Then
            Return False
        End If
        Return True
    End Function
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim frmAbout As New AboutForm
        frmAbout.ShowDialog()

    End Sub
    Private Sub OpenRecentFile(sender As Object, e As EventArgs)
        Dim item As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        If File.Exists(item.Tag.ToString) Then
            Dim data As New FileData("", "")
            data.FileName = item.Tag.ToString
            data.FileContents = File.ReadAllText(data.FileName)
            data.DocumentChanged = False
            Me.DocumentData = data
            UpdateFormTitleAndContent(DocumentData)
            UpdateStatistics(DocumentData)
        End If

    End Sub

    Private Sub tsbFindNext_Click(sender As Object, e As EventArgs) Handles tsbFindNext.Click
        Static intStart As Integer = 0
        Dim intFound As Integer
        If tstFind.TextBox.Text = "" OrElse
            intStart >= tstFind.TextBox.Text.Length Then
            Exit Sub
        Else

            intFound = txtDoc.Text.IndexOf(tstFind.TextBox.Text, intStart)
            If intFound = -1 Then
                MessageBox.Show("No more occurances found")
                intStart = 0
            Else
                txtDoc.Select(intFound, tstFind.TextBox.Text.Length)
                intStart = intFound + tstFind.TextBox.Text.Length
            End If
        End If

    End Sub

#End Region
End Class
