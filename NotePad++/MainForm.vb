Imports System.Drawing.Printing

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
        Me.Text = My.Resources.MainFormTitle & " - New File"
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
    Private Sub txtDoc_TextChanged(sender As Object, e As EventArgs) Handles txtDoc.TextChanged
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
        Me.Text = My.Resources.MainFormTitle & " - " & DocumentData.FileName & "*"
    End Sub

    Private Sub txtDoc_MouseWheel(sender As Object, e As MouseEventArgs) Handles txtDoc.MouseWheel
        If e.Delta > 0 AndAlso
            txtDoc.Font.Size + 1 < System.Single.MaxValue Then
            txtDoc.Font = New Font(txtDoc.Font.FontFamily, txtDoc.Font.Size + 1)

        Else
            If txtDoc.Font.Size - 1 > 1 Then
                txtDoc.Font = New Font(txtDoc.Font.FontFamily, txtDoc.Font.Size - 1)
            End If
        End If
    End Sub
    Private Sub txtDoc_FontChanged(sender As Object, e As EventArgs) Handles txtDoc.FontChanged
        tssFontFace.Text = txtDoc.Font.FontFamily.Name
        tssFontSize.Text = txtDoc.Font.Size.ToString & "pt"
    End Sub

#End Region
#Region "File Operations"
    Private Function OpenFile(data As FileData) As Boolean
        Dim dlgOpen As New OpenFileDialog
        With dlgOpen
            .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .FilterIndex = 1
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Title = "Open File"
        End With
        If dlgOpen.ShowDialog = DialogResult.OK Then
            data.FileName = dlgOpen.FileName
            data.FileContents = My.Computer.FileSystem.ReadAllText(data.FileName)
            data.DocumentChanged = False
            Me.DocumentData = data
            Return True
        End If
        Return False
    End Function
    Private Function SaveFile(data As FileData) As Boolean
        If data.FileName = "" Then
            Return SaveFileAs(data)
        Else
            My.Computer.FileSystem.WriteAllText(data.FileName, data.FileContents, False)
            data.DocumentChanged = False
            Me.DocumentData = data
            Return True
        End If
    End Function
    Private Function SaveFileAs(data As FileData) As Boolean
        Dim dlgSave As New SaveFileDialog
        With dlgSave
            .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            .FilterIndex = 1
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Title = "Save File As"
        End With
        If dlgSave.ShowDialog = DialogResult.OK Then
            data.FileName = dlgSave.FileName
            My.Computer.FileSystem.WriteAllText(data.FileName, data.FileContents, False)
            data.DocumentChanged = False
            Me.DocumentData = data
            Return True
        End If
        Return False
    End Function
#End Region
#Region "Menu Events"
#Region "File Menu"
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles NewToolStripMenuItem.Click
        ''first check to see if there is an open file to save
        If DocumentData.DocumentChanged Then
            If TrySave(DocumentData) Then
                ''celar the document data and the text box
                ClearData()
            End If
        End If
    End Sub
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        If DocumentData.DocumentChanged Then
            If (TrySave(DocumentData)) Then
                If OpenFile(DocumentData) Then
                    Me.Text = My.Resources.MainFormTitle & " - " &
                        DocumentData.FileName
                    Me.txtDoc.Text = DocumentData.FileContents
                End If
            End If
        End If
        If OpenFile(DocumentData) Then
            Me.Text = My.Resources.MainFormTitle & " - " &
                DocumentData.FileName
            Me.txtDoc.Text = DocumentData.FileContents
        End If
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
    'Private Sub UpdateStatistics()
    '    tssWordCount.Text = DocumentData.WordCount.ToString
    '    tssSentanceCount.Text = DocumentData.SentanceCount.ToString
    'End Sub
    Private Sub RefreshState()
        txtDoc.Text = DocumentData.FileContents
        Me.Text = My.Resources.MainFormTitle & " - " & DocumentData.FileName
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

    Private Sub tssFontFace_Click(sender As Object, e As EventArgs) Handles tssFontFace.Click

    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim frmAbout As New AboutForm
        frmAbout.ShowDialog()

    End Sub




#End Region
End Class
