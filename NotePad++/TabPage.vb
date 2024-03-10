Public Class TabPageTemplate : Inherits TabPage
#Region "Declarations"
    Public Const UNSAVED_FILE As String = "Unsaved File"
    Private _Data As New FileData
    Friend Property FileName As String
        Get
            Return _Data.FileName
        End Get
        Set
            _Data.FileName = Value
            Me.Text = If(Value = UNSAVED_FILE, UNSAVED_FILE, Value.Substring(Value.LastIndexOf("/") + 1))
        End Set
    End Property
    Friend Property FileContents As String
        Get
            Return _Data.FileContents
        End Get
        Set
            _Data.FileContents = Value
        End Set
    End Property
    Friend Property DocumentChanged As Boolean
        Get
            Return _Data.DocumentChanged
        End Get
        Set(value As Boolean)
            _Data.DocumentChanged = value
        End Set
    End Property
    Friend ReadOnly Property WordCount As Integer
        Get
            Return _Data.WordCount
        End Get
    End Property
    Friend ReadOnly Property SentanceCount As Integer
        Get
            Return _Data.SentanceCount
        End Get
    End Property
    Public WithEvents DocumentTextBox As TextBox
    Friend WithEvents StatusBar As MyStatus
#End Region
#Region "Constructors"
    Friend Sub New()
        Me.DocumentTextBox = New System.Windows.Forms.TextBox()
        Me.StatusBar = New MyStatus()
        Me.SuspendLayout()
        Me.FileName = UNSAVED_FILE
        Me.DocumentTextBox.Name = "DocumentTextBox"
        Me.DocumentTextBox.Dock = DockStyle.Fill
        Me.DocumentTextBox.Multiline = True
        Me.DocumentTextBox.ScrollBars = ScrollBars.Vertical
        Me.DocumentTextBox.TabIndex = 0
        Me.Controls.Add(Me.DocumentTextBox)
        Me.StatusBar.Dock = DockStyle.Bottom
        Me.StatusBar.Name = "StatusBar"
        Me.Controls.Add(Me.StatusBar)
        Me.StatusBar.tssFontFace.Text = DocumentTextBox.Font.Name
        Me.StatusBar.tssFontSize.Text = DocumentTextBox.Font.Size.ToString("##") & "pt"
        Me.StatusBar.tssWordCount.Text = "0"
        Me.StatusBar.tssSentanceCount.Text = "0"
        Me.ResumeLayout(False)
        Me.PerformLayout()
        AddHandler Me.StatusBar.tssFontFace.Click, AddressOf StatusBar_tssFontFace_Click
        AddHandler Me.DocumentTextBox.TextChanged, AddressOf DocumentTextBox_TextChanged
        ''AddHandler Me.DocumentTextBox.KeyDown, AddressOf DocumentTextBox_KeyDown
        AddHandler Me.DocumentTextBox.MouseWheel, AddressOf DocumentTextBox_MouseWheel
        AddHandler Me.DocumentTextBox.FontChanged, AddressOf DocumentTextBox_FontChanged
    End Sub

    Friend Sub New(content As String)
        Me.New()
        Me.FileContents = content
    End Sub
    Friend Sub New(data As FileData)
        Me.New()
        Me._Data = data
    End Sub
    Friend Sub New(other As TabPageTemplate)
        Me.New()
        Me.FileName = other.FileName
        Me.FileContents = other.FileContents
    End Sub
#End Region
#Region "TextBox Events"
    Private Sub DocumentTextBox_FontChanged(sender As Object, e As EventArgs)
        StatusBar.tssFontFace.Text = DocumentTextBox.Font.Name
        StatusBar.tssFontSize.Text = DocumentTextBox.Font.Size.ToString("##") & "pt"
    End Sub

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
    Private Sub StatusBar_tssFontFace_Click(sender As Object, e As EventArgs)
        Dim dlgFont As New FontDialog
        dlgFont.Font = DocumentTextBox.Font
        If dlgFont.ShowDialog = DialogResult.OK Then
            DocumentTextBox.Font = dlgFont.Font
        End If
    End Sub

    Private Sub DocumentTextBox_TextChanged(sender As Object, e As EventArgs)
        Me.FileContents = Me.DocumentTextBox.Text
        Me._Data.DocumentChanged = True
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
        'Dim strDocumentTitle As String = If(Me.FileName = "", UNSAVED_FILE, Me.FileName.Substring(Me.FileName.LastIndexOf("/") + 1))
        'Me.Text = $"{strDocumentTitle}*"
    End Sub

#End Region
End Class
