Imports System.IO
Imports Microsoft.VisualBasic.FileIO

''' <summary>
''' Inherits and extends the TabPage control to include a TextBox and a StatusBar and 
''' to manage the data for a single text file.
''' </summary>
''' <remarks>Author: Christopher Vaughn</remarks>
''' <updated>3/11/2024</updated>
Public Class TabPageTemplate : Inherits TabPage
#Region "Declarations"
    ''' <summary>
    ''' Constant for the name of an unsaved file.
    ''' </summary>
    ''' <summary>
    ''' A structure to hold the data for the file associated with this tab.
    ''' </summary>
    Private _Data As New FileData(UnsavedFileText, "")
    ''' <summary>
    ''' The text box to display and interreact with the file contents.
    ''' </summary>
    Public WithEvents DocumentTextBox As TextBox
    ''' <summary>
    ''' The status bar to display information about the file.
    ''' </summary>
    Friend WithEvents StatusBar As MyStatus
    ''' <summary>
    ''' The full path to the file associated with this tab 
    ''' or the constant UNSAVED_FILE if the file has not been saved yet.
    ''' </summary>
    ''' <returns>Full path to the tabs file location</returns>
    Friend Property FileName As String
        Get
            Return _Data.FileName
        End Get
        Set
            _Data.FileName = Value
            If FileName = UnsavedFileText Then
                Me.Text = FileName
            Else
                Me.Text = FileName.Substring(Value.LastIndexOf("\") + 1)
            End If
        End Set
    End Property
    ''' <summary>
    ''' The ascii text of the file associated with this tab. The 
    ''' value is synchronized with the text box contents in the tab.
    ''' </summary>
    ''' <returns>Ascii text containing the file contents</returns>
    Friend Property FileContents As String
        Get
            Return _Data.FileContents
        End Get
        Set
            Static firstLoad As Boolean = True
            _Data.FileContents = Value
            If firstLoad Then
                firstLoad = False
                Me.DocumentTextBox.Text = Value
                Me.DocumentChanged = False
            End If

        End Set
    End Property
    ''' <summary>
    ''' Tracks if the file text contents have changed since the last save.
    ''' </summary>
    ''' <returns>True if the file contents have changed</returns>
    Friend Property DocumentChanged As Boolean
        Get
            Return _Data.DocumentChanged
        End Get
        Set(value As Boolean)
            _Data.DocumentChanged = value
        End Set
    End Property
    ''' <summary>
    ''' Calculates the number of words in the file contents.
    ''' </summary>
    ''' <returns>Read-only count of how many words are currently in the contents string</returns>
    Friend ReadOnly Property WordCount As Integer
        Get
            Return _Data.WordCount
        End Get
    End Property
    ''' <summary>
    ''' Calculates the number of sentances in the file contents.
    ''' </summary>
    ''' <returns>Read-only count of how many sentances are currently in the contents string</returns>
    Friend ReadOnly Property SentanceCount As Integer
        Get
            Return _Data.SentanceCount
        End Get
    End Property
#End Region
#Region "Constructors/Destructor"
    ''' <summary>
    ''' Default constructor to initialize the tab page with a text box and a status bar.
    ''' </summary>
    Friend Sub New()
        ''Initialize the tab page and status bar
        Me.DocumentTextBox = New System.Windows.Forms.TextBox()
        Me.StatusBar = New MyStatus()
        ''suspend layout to avoid flicker while adding controls
        Me.SuspendLayout()
        ''initialize the tabs file path data
        Me.FileName = UnsavedFileText
        ''initialize the text box
        Me.DocumentTextBox.Name = "DocumentTextBox"
        Me.DocumentTextBox.Text = String.Empty
        Me.DocumentTextBox.Dock = DockStyle.Fill
        Me.DocumentTextBox.Multiline = True
        Me.DocumentTextBox.ScrollBars = ScrollBars.Vertical
        Me.DocumentTextBox.TabIndex = 0
        'add the text box to the tab page's controls collection
        Me.Controls.Add(Me.DocumentTextBox)
        ''initialize the status bar
        Me.StatusBar.Dock = DockStyle.Bottom
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.tssFontFace.Text = DocumentTextBox.Font.Name
        Me.StatusBar.tssFontSize.Text = DocumentTextBox.Font.Size.ToString("##") & "pt"
        Me.StatusBar.tssWordCount.Text = Me.WordCount.ToString
        Me.StatusBar.tssSentanceCount.Text = Me.SentanceCount.ToString
        ''add the status bar to the tab page's controls collection
        Me.Controls.Add(Me.StatusBar)
        ''resume layout to display the controls
        Me.ResumeLayout(False)
        Me.PerformLayout()
        ''add event handlers for the text box and status bar
        AddHandler Me.StatusBar.tssFontFace.Click, AddressOf StatusBar_tssFontFace_Click
        AddHandler Me.DocumentTextBox.TextChanged, AddressOf DocumentTextBox_TextChanged
        AddHandler Me.DocumentTextBox.MouseWheel, AddressOf DocumentTextBox_MouseWheel
        AddHandler Me.DocumentTextBox.FontChanged, AddressOf DocumentTextBox_FontChanged
    End Sub
    ''' <summary>
    ''' Constructor to initialize the tab page with a text box and a status bar and 
    ''' to set the file contents to the given string with no file name.
    ''' </summary>
    ''' <param name="content"></param>
    Friend Sub New(content As String)
        Me.New()
        Me.FileContents = content
    End Sub
    ''' <summary>
    ''' Constructor to initialize the tab page with a text box and a status bar and 
    ''' to set the file contents to the given string and the file name to the given string.
    ''' </summary>
    ''' <param name="fileName">The path and filename for the file associated with this tab</param>
    ''' <param name="content">The ascii text that is the content of the file being edited in this tab</param>
    Friend Sub New(fileName As String, content As String)
        Me.New()
        Me.FileName = fileName
        Me.FileContents = content
    End Sub
    ''' <summary>
    ''' 'Set the file data for this tab to the given file data during construction.
    ''' </summary>
    ''' <param name="data">A file data object</param>
    Friend Sub New(data As FileData)
        Me.New()
        Me._Data = data
    End Sub
    ''' <summary>
    ''' Copy constructor to initialize the tab page with a text box and a status bar and 
    ''' to set the file contents and file name to an unsaved constant.
    ''' </summary>
    ''' <param name="other"></param>
    Friend Sub New(other As TabPageTemplate)
        Me.New()
        Me.FileName = UnsavedFileText
        Me.FileContents = other.FileContents
        Me.StatusBar.tssSentanceCount.Text = Me.SentanceCount.ToString
        Me.StatusBar.tssWordCount.Text = Me.WordCount.ToString
    End Sub
    ''' <summary>
    ''' Dispose of the text box and status bar controls when the tab page is disposed.
    ''' </summary>
    Friend Sub Dispose()
        Me.DocumentTextBox.Dispose()
        Me.StatusBar.Dispose()
        MyBase.Dispose()
    End Sub
#End Region
#Region "Control Events"
    ''' <summary>
    ''' Updates the status bar with the current font face and size when the font is changed.
    ''' </summary>
    ''' <param name="sender">The tab control's documents textbox</param>
    ''' <param name="e">Event Args</param>
    Private Sub DocumentTextBox_FontChanged(sender As Object, e As EventArgs)
        StatusBar.tssFontFace.Text = DocumentTextBox.Font.Name
        StatusBar.tssFontSize.Text = DocumentTextBox.Font.Size.ToString("##") & "pt"
    End Sub
    ''' <summary>
    ''' Increases or decreases the font size of the text box when the mouse wheel is scrolled 
    ''' while the text box has focus.
    ''' </summary>
    ''' <param name="sender">The tab control's documents textbox</param>
    ''' <param name="e">Event Args</param>
    Private Sub DocumentTextBox_MouseWheel(sender As Object, e As MouseEventArgs)
        If e.Delta > 0 AndAlso
            DocumentTextBox.Font.Size + 1 < System.Single.MaxValue Then
            DocumentTextBox.Font = New Font(DocumentTextBox.Font.FontFamily, DocumentTextBox.Font.Size + 1)

        Else
            If DocumentTextBox.Font.Size - 1 > 1 Then
                DocumentTextBox.Font = New Font(DocumentTextBox.Font.FontFamily, DocumentTextBox.Font.Size - 1)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Handles the font face menu item click event to change the font face of the text box.
    ''' </summary>
    ''' <param name="sender">The status bar font face label</param>
    ''' <param name="e">Event Args</param>
    Private Sub StatusBar_tssFontFace_Click(sender As Object, e As EventArgs)
        Dim dlgFont As New FontDialog
        dlgFont.Font = DocumentTextBox.Font
        If dlgFont.ShowDialog = DialogResult.OK Then
            DocumentTextBox.Font = dlgFont.Font
        End If
    End Sub
    ''' <summary>
    ''' Keeps the document data synchronized with the text box contents and updates the 
    ''' status bar word and sentance counts when the text box contents change. 
    ''' Looks for paterns that indicate the end of a sentance and updates the sentance count and
    ''' looks for paterns that indicate the end of a word and updates the word count.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DocumentTextBox_TextChanged(sender As Object, e As EventArgs)
        Me.FileContents = Me.DocumentTextBox.Text
        Me._Data.DocumentChanged = True
        ''if the last character of the text property is not a % then make it one.
        If Me.Text.Length > 0 AndAlso
            Me.Text(Me.Text.Length - 1) <> "*"c Then
            Me.Text &= " *"
        End If
        If Me.FileContents.Length > 1 AndAlso
            (Me.FileContents(Me.FileContents.Length - 2) = "."c OrElse
             Me.FileContents(Me.FileContents.Length - 2) = "?"c OrElse
             Me.FileContents(Me.FileContents.Length - 2) = "!"c) AndAlso
             Me.FileContents(Me.FileContents.Length - 1) = " "c Then
            Me.StatusBar.tssSentanceCount.Text = Me.SentanceCount.ToString
        End If
        If Me.FileContents.Length > 0 AndAlso
            Me.FileContents(Me.FileContents.Length - 1) = " "c Then
            Me.StatusBar.tssWordCount.Text = Me.WordCount.ToString
        End If
    End Sub
#End Region
#Region "File Operations"
    ''' <summary>
    ''' Opens a file dialog to select a file to open and loads the file contents into the text box.
    ''' </summary>
    ''' <returns>True if the file was opened successfully</returns>
    Public Function OpenFile() As Boolean
        Dim ofd As New OpenFileDialog
        Try
            ofd.Title = "Open Text File"
            ofd.InitialDirectory = SpecialDirectories.MyDocuments
            ofd.Filter = "Text Files (*.txt)|*.txt | All Files (*.*)|*.*"
            ofd.FilterIndex = 1
            If ofd.ShowDialog = DialogResult.OK Then
                Me.FileName = ofd.FileName
                Me.FileContents = File.ReadAllText(ofd.FileName)
                Me.DocumentTextBox.Text = Me.FileContents
                Me.DocumentChanged = False
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Saves the file contents to the file associated with this tab. If the file has not been saved before, 
    ''' prompts the user to save as a new file.
    ''' </summary>
    ''' <returns>True if the file was successfully saved.</returns>
    Public Function SaveFile() As Boolean
        Dim sfd As New SaveFileDialog
        Try
            ''if the file has not been saved before, prompt to save as
            If Me.FileName = UnsavedFileText Then
                sfd.Title = "Save As New Document"
                sfd.InitialDirectory = SpecialDirectories.MyDocuments
                sfd.Filter = "Text Files (*.txt)|*.txt | All Files (*.*)|*.*"
                sfd.FilterIndex = 1
                ''if the user cancels, return false
                If sfd.ShowDialog = DialogResult.OK Then
                    File.WriteAllText(sfd.FileName, Me.FileContents)
                    Me.FileName = sfd.FileName
                    Me.DocumentChanged = False

                    Return True
                Else
                    Return False
                End If
            End If
            ''if the file has been saved before, save it
            File.WriteAllText(Me.FileName, Me.FileContents)
            Me.DocumentChanged = False
            Me.FileName = Me.FileName
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Function SaveFileAs() As Boolean
        Dim sfd As New SaveFileDialog
        Try
            sfd.Title = "Save Text File As"
            sfd.InitialDirectory = SpecialDirectories.MyDocuments
            sfd.Filter = "Text Files (*.txt)|*.txt | All Files (*.*)|*.*"
            sfd.FilterIndex = 1
            If sfd.ShowDialog = DialogResult.OK Then
                File.WriteAllText(sfd.FileName, Me.FileContents)
                Me.FileName = sfd.FileName
                Me.DocumentChanged = False
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
#End Region
End Class
