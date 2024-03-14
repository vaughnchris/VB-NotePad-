Imports System.ComponentModel
Imports System.Drawing.Printing
Imports System.IO

Public Class MainForm
    Private Property CurrentTab As TabPageTemplate
#Region "Form Control Events"
    ''' <summary>
    ''' Load the form and set the form's location, size, and window state to the last saved settings.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With My.Settings
            Me.Location = .StartLocation
            Me.WindowState = .StartWindowState
            Me.Size = .StartSize
        End With
        Dim SplashForm As New Splash
        SplashForm.ShowDialog()

    End Sub
    ''' <summary>
    ''' Store the form's location, size, and window state to the settings when the form is closing.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MainForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        With My.Settings
            .StartLocation = Me.Location
            .StartWindowState = Me.WindowState
            .StartSize = Me.Size
            .Save()
        End With
    End Sub
    ''' <summary>
    ''' Sets the new tabs autowrap property to synch with the menu item 
    ''' and sets the current tab to the new tab and set focus to the new tab's text box.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DocumentTabs_ControlAdded(sender As Object, e As ControlEventArgs) Handles DocumentTabs.ControlAdded
        ''set the current tab to the added tab
        Dim tab As TabPageTemplate = CType(e.Control, TabPageTemplate)
        Me.CurrentTab = tab
        AutoWrapTextToolStripMenuItem_Click(tab, Nothing)
        Me.Text = $"{My.Resources.MainFormTitle} - {CurrentTab.FileName}"
        ''set the current tab to the added tab
        DocumentTabs.SelectedTab = tab
        tab.DocumentTextBox.Focus()

    End Sub
    ''' <summary>
    ''' Updates the title bar of the form to reflect the current tab's file name 
    ''' when the selected tab is changed.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DocumentTabs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DocumentTabs.SelectedIndexChanged
        ''set the current tab to the selected tab
        Me.CurrentTab = CType(DocumentTabs.SelectedTab, TabPageTemplate)
        If Me.CurrentTab IsNot Nothing Then
            Me.Text = $"{My.Resources.MainFormTitle} - {CurrentTab.FileName}"
        End If

    End Sub
#End Region
#Region "Form Menu Events"
#Region "File Menu Events"
    ''' <summary>
    ''' Sets file logic for the file menu items to be enabled or disabled based on the current tab's state.
    ''' And the recent files list to be populated if there are recent files.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub File_DropDownOpening(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.DropDownOpening
        ''if there are no tabs open, disable the close menu item
        If Me.CurrentTab IsNot Nothing Then
            ''if the document has changed since the last save, enable the save menu item
            If Me.CurrentTab.DocumentChanged Then
                SaveToolStripMenuItem.Enabled = True
            Else
                SaveToolStripMenuItem.Enabled = False
            End If
            ''if the file has a file name, enable the save as menu item
            If Me.CurrentTab.FileName = UnsavedFileText Then
                SaveAsToolStripMenuItem.Enabled = False
            Else
                SaveAsToolStripMenuItem.Enabled = True
            End If
            ''if the file has no contents, disable the print menu items
            If Me.CurrentTab.FileContents = "" Then
                PrintPreviewToolStripMenuItem.Enabled = False
                PrintToolStripMenuItem.Enabled = False
            Else
                PrintPreviewToolStripMenuItem.Enabled = True
                PrintToolStripMenuItem.Enabled = True
            End If
            ''enable the close menu item
            CloseToolStripMenuItem.Enabled = True
        Else
            ''if there are no tabs open
            SaveToolStripMenuItem.Enabled = False
            SaveAsToolStripMenuItem.Enabled = False
            PrintPreviewToolStripMenuItem.Enabled = False
            PrintToolStripMenuItem.Enabled = False
            CloseToolStripMenuItem.Enabled = False
            NewToolStripMenuItem.Enabled = True
            OpenToolStripMenuItem.Enabled = True
        End If
        If My.Settings.RecentFiles IsNot Nothing AndAlso
                My.Settings.RecentFiles.Count > 0 Then
            OpenRecentToolStripMenuItem.DropDownItems.Clear()
            For Each file As String In My.Settings.RecentFiles
                Dim item As New ToolStripMenuItem
                item.Text = Path.GetFileName(file)
                item.Tag = file
                If CurrentTab IsNot Nothing AndAlso
                    item.Tag.ToString = CurrentTab.FileName Then
                    item.Enabled = False
                End If
                AddHandler item.Click, AddressOf OpenRecentFile
                OpenRecentToolStripMenuItem.DropDownItems.Add(item)
            Next
            OpenRecentToolStripMenuItem.Enabled = True
        Else
            OpenRecentToolStripMenuItem.Enabled = False
        End If
    End Sub
    ''' <summary>
    ''' Removes the recent files list when the file menu is closed.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FileToolStripMenuItem_DropDownClosed(sender As Object, e As EventArgs) Handles FileToolStripMenuItem.DropDownClosed
        OpenRecentToolStripMenuItem.DropDownItems.Clear()
    End Sub
    ''' <summary>
    ''' Creates a new tab and adds it to the tab control. with no file contents.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        ''create a new tab and add it to the tab control
        Dim newTab As TabPageTemplate = New TabPageTemplate()
        DocumentTabs.TabPages.Add(newTab)
    End Sub
    ''' <summary>
    ''' Creates a new tab and adds it to the tab control. with the file contents
    '''  selected by the user.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim newTab As TabPageTemplate = New TabPageTemplate
        ''if the file opens successfully, add its tab page to the tab control
        If newTab.OpenFile Then
            Me.DocumentTabs.TabPages.Add(newTab)
        Else
            ''if the file does not open successfully, dispose of the tab
            newTab.Dispose()
        End If
    End Sub
    ''' <summary>
    ''' Shared event to open a recent file from the recent files list of MenuItems.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OpenRecentFile(sender As Object, e As EventArgs)
        ''open the file from the recent files list starting at the tag of the menu item
        ''the tag is the full path to the file
        Dim selectedPath As String = CType(sender, ToolStripMenuItem).Tag.ToString
        ''if the file exists, open it
        If File.Exists(selectedPath) Then
            ''create a new tab from the file
            Dim newTab As TabPageTemplate = New TabPageTemplate
            newTab.FileName = selectedPath
            ''read the file into the tab's text box
            newTab.FileContents = File.ReadAllText(selectedPath)
            ''add the tab to the tab control
            Me.DocumentTabs.TabPages.Add(newTab)
            Me.CurrentTab = newTab
        Else
            ''if the file does not exist, remove it from the recent files list
            My.Settings.RecentFiles.Remove(selectedPath)
            ''save the settings
            My.Settings.Save()
        End If
    End Sub
    ''' <summary>
    ''' Saves the data from the current tab to the file it was opened from or 
    ''' a new file if it has not been saved before.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If Me.CurrentTab.SaveFile() Then
            UpdateRecentPaths()
            Me.Text = $"{My.Resources.MainFormTitle} - {Me.CurrentTab.FileName}"
        End If
    End Sub
    ''' <summary>
    ''' Updates the recent files list with the current tab's file path.
    ''' </summary>
    Private Sub UpdateRecentPaths()
        If My.Settings.RecentFiles Is Nothing Then
            My.Settings.RecentFiles = New Specialized.StringCollection
        End If
        With My.Settings.RecentFiles
            If Not .Contains(CurrentTab.FileName) Then
                .Insert(0, CurrentTab.FileName)
                If .Count > 10 Then
                    .RemoveAt(.Count - 1)
                End If
            End If
        End With
        My.Settings.Save()
    End Sub
    ''' <summary>
    ''' Save a copy of the current tab's file contents to a new file
    ''' and create a new tab with the new file data.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        ''create a new tab from the current tab using the copy constructor
        Dim tabCopy As TabPageTemplate = New _
            TabPageTemplate(Me.CurrentTab)
        ''if the file is saved successfully, add the tab to the tab control
        If tabCopy.SaveFileAs Then
            Me.UpdateRecentPaths()
            Me.DocumentTabs.TabPages.Add(tabCopy)
        Else
            ''if the file is not saved successfully, dispose of the tab
            tabCopy.Dispose()
        End If
    End Sub
    ''' <summary>
    ''' Saves the data from the current tab to the file if it has changed since the last save.
    ''' Then removes the tab from the tab control and disposes of the tab.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        ''if the current tab is not nothing
        If CurrentTab IsNot Nothing Then
            ''if the document has been changed, prompt to save
            If CurrentTab.DocumentChanged Then
                Dim result As DialogResult = MessageBox.Show("Do you want to save changes to " & CurrentTab.FileName & "?",
                                                             "Save Changes?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                ''if the user wants to save the file, save it
                If result = DialogResult.Yes Then
                    ''save the file
                    CurrentTab.SaveFile()
                    ''remove the tab from the tab control
                    Me.DocumentTabs.TabPages.Remove(CurrentTab)
                    ''dispose of the tab
                    CurrentTab.Dispose()
                    ''if the user does not want to save the file, remove the tab from the tab control
                ElseIf result = DialogResult.Cancel Then
                    ''do nothing
                    Return
                End If
                ''if the document has not been changed since last save,
                '', remove the tab from the tab control
            Else
                Me.DocumentTabs.TabPages.Remove(CurrentTab)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Checks if any of the tabs have changed since the last save.
    ''' If so, prompts the user to save the file.
    ''' Then closes the application.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        For Each tab As TabPageTemplate In DocumentTabs.TabPages
            If tab.DocumentChanged Then
                Dim result As DialogResult = MessageBox.Show("Do you want to save changes to " & tab.FileName & "?",
                                                             "Save Changes?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                If result = DialogResult.Yes Then
                    tab.SaveFile()
                ElseIf result = DialogResult.Cancel Then
                    Return
                End If
            End If
        Next
        Me.Close()
    End Sub

#Region "Printing Methods"
    ''' <summary>
    ''' Manages the printing of the current tab's file contents.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ''print the contents of the current tab
        Dim strContents As String = Me.CurrentTab.FileContents
        ''set the value of charactersOnPage to the number of characters
        Dim charactersOnPage As Integer = 0
        ''set the value of linesPerPage to the number of lines
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
    ''' <summary>
    ''' Manages the printing of the current tab's file contents by
    ''' displaying the print dialog and printing the file contents.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog = DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub
    ''' <summary>
    ''' Manages the printing of the current tab's file contents by
    ''' displaying the print preview dialog and printing the file contents.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub PrintPreviewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub
#End Region
#End Region
#Region "Edit Menu Events"
    ''' <summary>
    ''' Sets the edit menu items to be enabled or disabled based on the current tab's state.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub EditToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.DropDownOpening
        If CurrentTab IsNot Nothing Then
            UndoToolStripMenuItem.Enabled = CurrentTab.DocumentTextBox.CanUndo

            If CurrentTab.DocumentTextBox.SelectionLength > 0 Then
                CutToolStripMenuItem.Enabled = True
                CopyToolStripMenuItem.Enabled = True
            Else
                CutToolStripMenuItem.Enabled = False
                CopyToolStripMenuItem.Enabled = False
            End If
            If Clipboard.ContainsText Then
                PasteToolStripMenuItem.Enabled = True
            Else
                PasteToolStripMenuItem.Enabled = False
            End If
            If CurrentTab.FileContents.Length > 0 Then
                Me.DuplicateWindowToolStripMenuItem.Enabled = True
            Else
                Me.DuplicateWindowToolStripMenuItem.Enabled = False
            End If
        Else
            UndoToolStripMenuItem.Enabled = False
            RedoToolStripMenuItem.Enabled = False
            CutToolStripMenuItem.Enabled = False
            CopyToolStripMenuItem.Enabled = False
            PasteToolStripMenuItem.Enabled = False
            SelectAllToolStripMenuItem.Enabled = False
            DuplicateWindowToolStripMenuItem.Enabled = False
        End If

    End Sub
    ''' <summary>
    ''' Undoes the last action in the current tab's text box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        If CurrentTab.DocumentTextBox.CanUndo Then
            CurrentTab.DocumentTextBox.Undo()
        End If
    End Sub
    ''' <summary>
    ''' Redoes the last action in the current tab's text box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click

    End Sub
    ''' <summary>
    ''' Copy the selected text to the clipboard from 
    ''' the current tab's text box.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        ''copy the selected text to the clipboard
        CurrentTab.DocumentTextBox.Copy()
    End Sub
    ''' <summary>
    ''' Cut the selected text to the clipboard from 
    ''' the current tab's text box.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        'cut the selected text to the clipboard
        CurrentTab.DocumentTextBox.Cut()
    End Sub
    ''' <summary>
    ''' Paste the content from the clipboard into the current tab's text box 
    ''' at the current cursor position.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        ''insert the content from the clipboard into the text box
        CurrentTab.DocumentTextBox.Paste()
    End Sub
    ''' <summary>
    ''' Select all the text in the current tab's text box.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        ''copy everything from the textbox to the clipboard
        CurrentTab.DocumentTextBox.SelectAll()
    End Sub
    ''' <summary>
    ''' Duplicate the current tab and contents, and add it to the tab control.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DuplicateWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DuplicateWindowToolStripMenuItem.Click
        ''create a new tab from the current tab using the copy constructor and add it to the tab control
        Dim tabCopy As New TabPageTemplate(Me.CurrentTab)
        Me.DocumentTabs.TabPages.Add(tabCopy)
    End Sub

#End Region
#Region "Settings Menu Events"
    ''' <summary>
    ''' Sets the settings menu items to be enabled or disabled based on the current tab's state.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SettingsToolStripMenuItem_DropDownOpening(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.DropDownOpening
        If Me.CurrentTab Is Nothing Then
            Me.FontFaceToolStripMenuItem.Enabled = False
            Me.AutoWrapTextToolStripMenuItem.Enabled = False
        Else
            Me.FontFaceToolStripMenuItem.Enabled = True
            Me.AutoWrapTextToolStripMenuItem.Enabled = True
        End If
    End Sub
    ''' <summary>
    ''' Changes the font of the current tab's text box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FontFaceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontFaceToolStripMenuItem.Click
        If CurrentTab IsNot Nothing Then
            Dim dlgFont As New FontDialog
            dlgFont.Font = CurrentTab.DocumentTextBox.Font
            If dlgFont.ShowDialog() = DialogResult.OK Then
                CurrentTab.DocumentTextBox.Font = dlgFont.Font
            End If
        End If
    End Sub
    ''' <summary>
    ''' Toggles the word wrap property of the current tab's text box.
    ''' </summary>
    ''' <param name="sender">The Selected MenuItem</param>
    ''' <param name="e">Event Args</param>
    Private Sub AutoWrapTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoWrapTextToolStripMenuItem.Click
        If CurrentTab IsNot Nothing Then
            ''toggle the word wrap property of the current tab's text box
            CurrentTab.DocumentTextBox.WordWrap = AutoWrapTextToolStripMenuItem.Checked
            ''toggle the scroll bars
            Me.SuspendLayout()
            CurrentTab.DocumentTextBox.ScrollBars = If(CurrentTab.DocumentTextBox.WordWrap, ScrollBars.None, ScrollBars.Vertical)
            ''update the status bar
            CurrentTab.StatusBar.tssWordWrap.Text = If(CurrentTab.DocumentTextBox.WordWrap, "On", "Off")
            Me.ResumeLayout()

        End If
    End Sub
#End Region
#Region "Help Menu Events"
    ''' <summary>
    ''' Displays the about form.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        Dim frmAbout As New AboutForm
        frmAbout.ShowDialog()

    End Sub
#End Region
#End Region
End Class
