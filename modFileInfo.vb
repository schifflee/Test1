Imports System.IO

Module modFileInfo

    Public Enum CompareByOptions
        FileName
        LastWriteTime
        Length
    End Enum

    Public Class CompareFileInfoEntries
        Implements IComparer

        Private compareBy As CompareByOptions = CompareByOptions.FileName


        Public Sub New(ByVal cBy As CompareByOptions)
            compareBy = cBy
        End Sub

        Public Overridable Overloads Function Compare(ByVal file1 As Object, ByVal file2 As Object) As Integer Implements IComparer.Compare
            'Convert file1 and file2 to FileInfo entries
            Dim f1 As FileInfo = CType(file1, FileInfo)
            Dim f2 As FileInfo = CType(file2, FileInfo)

            'Compare the file names
            Select Case compareBy
                Case CompareByOptions.FileName
                    Return String.Compare(f1.Name, f2.Name)
                Case CompareByOptions.LastWriteTime
                    Return DateTime.Compare(f1.LastWriteTime, f2.LastWriteTime)
                Case CompareByOptions.Length
                    Return f1.Length - f2.Length
            End Select
        End Function
    End Class

    'New stuff
    'http://msdn.microsoft.com/en-us/library/wz42302f.aspx
    Public Function aGetFiles(ByVal directoryInfo As System.IO.DirectoryInfo, ByVal searchPatterns() As String) As System.IO.FileInfo()
        Return bGetFiles(directoryInfo, searchPatterns, System.IO.SearchOption.TopDirectoryOnly)
    End Function

    Public Function bGetFiles(ByVal directoryInfo As System.IO.DirectoryInfo, ByVal searchPatterns() As String, ByVal searchOptions As System.IO.SearchOption) As System.IO.FileInfo()
        Dim oFileListing As New List(Of System.IO.FileInfo)
        For Each sSearchPattern As String In searchPatterns
            oFileListing.AddRange(directoryInfo.GetFiles(sSearchPattern, searchOptions))
        Next
        Return oFileListing.ToArray
    End Function
    'end New stuff

End Module
