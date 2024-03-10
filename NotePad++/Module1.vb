Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Module ModMain
    Structure FileData
        Private _FileContents As String
        Public Property FileName As String
        Public Property FileContents As String
            Get
                Return _FileContents
            End Get
            Set
                _FileContents = Value
                DocumentChanged = True
            End Set
        End Property
        Public Property DocumentChanged As Boolean
        Public ReadOnly Property IsNew As Boolean
            Get
                Return FileName = ""
            End Get
        End Property
        Public ReadOnly Property WordCount As Integer
            Get
                Return FileContents.Split(" "c).Length - 1
            End Get
        End Property
        Public ReadOnly Property SentanceCount As Integer
            Get
                SentanceCount = 0
                For i As Integer = 0 To FileContents.Length - 1
                    If FileContents(i) = "."c Or FileContents(i) = "!"c Or FileContents(i) = "?"c Then
                        If FileContents(i + 1) = " "c Then
                            SentanceCount += 1
                        End If
                    End If
                Next
            End Get
        End Property
        Sub New(ByVal FileName As String, ByVal FileContents As String)
            Me.FileName = FileName
            Me.FileContents = FileContents
            Me.DocumentChanged = False
        End Sub
    End Structure
    ''open a text file and return its contents
    Function OpenFile(DocumentData As FileData) As Boolean
        Dim ofd As New OpenFileDialog
        Try
            ofd.Title = "Open Text File"
            ofd.InitialDirectory = SpecialDirectories.MyDocuments
            ofd.Filter = "Text Files|*.txt"
            ofd.Filter = "All Files|*.*"

            If ofd.ShowDialog = DialogResult.OK Then
                DocumentData.FileName = ofd.FileName
                DocumentData.FileContents = File.ReadAllText(ofd.FileName)
                DocumentData.DocumentChanged = False
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    ''save a string to a text file
    Function SaveFile(DocumentData As FileData) As Boolean
        Dim sfd As New SaveFileDialog
        Try
            ''if the file has not been saved before, prompt to save as
            If DocumentData.FileName = "" Then
                sfd.Title = "Save As New Document"
                sfd.InitialDirectory = SpecialDirectories.MyDocuments
                sfd.Filter = "Text Files|*.txt"
                ''if the user cancels, return false
                If sfd.ShowDialog = DialogResult.OK Then
                    File.WriteAllText(sfd.FileName, DocumentData.FileContents)
                    DocumentData.FileName = sfd.FileName
                    DocumentData.DocumentChanged = False
                    Return True
                Else
                    Return False
                End If
            End If
            ''if the file has been saved before, save it
            File.WriteAllText(DocumentData.FileName, DocumentData.FileContents)
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    ''prompt to save as
    Function SaveFileAs(DocumentData As FileData) As Boolean
        Dim sfd As New SaveFileDialog
        Try
            sfd.Title = "Save Text File As"
            sfd.InitialDirectory = SpecialDirectories.MyDocuments
            sfd.Filter = "Text Files|*.txt"
            If sfd.ShowDialog = DialogResult.OK Then
                File.WriteAllText(sfd.FileName, DocumentData.FileContents)
                DocumentData.FileName = sfd.FileName
                DocumentData.DocumentChanged = False
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
End Module
