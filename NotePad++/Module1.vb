''' <summary>
''' This module contains code shared by all forms in the application.
''' </summary>
''' <remarks>Author: Christopher Vaughn</remarks>
Module ModMain
#Region "Type Definitions"
    ''' <summary>
    ''' A structure to hold the data for a single file.
    ''' </summary>
    Public Class FileData
        Private _FileContents As String
        Private _FileName As String
        Public Property FileName As String
            Get
                Return _FileName
            End Get
            Set
                _FileName = Value
            End Set
        End Property
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
    End Class
#End Region
#Region "Global Variables"
    Public UnsavedFileText As String = My.Resources.UnsavedFileText
#End Region
End Module
