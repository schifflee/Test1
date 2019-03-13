Public Class frmChromReporterChoice

    Public boolCancel As Boolean = True

    Private Sub frmChromReporterChoice_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        Dim str1 As String

        str1 = "Enter a Sciex" & ChrW(8482) & " Analyst .rdb file:"
        Me.lblRDB.Text = str1





    End Sub

    Private Sub cmdBrowse_Click(sender As Object, e As EventArgs) Handles cmdBrowseRDB.Click

        Dim strFilter As String
        Dim strFileName As String
        Dim str1 As String
        Dim str2 As String

        strFilter = "Sciex Analyst.rdb file (*.rdb)|*.rdb"
        strFileName = "*.rdb"

        Dim strPath As String = ""

        str2 = ReturnDirectoryBrowse(True, strPath, strFilter, strFileName, True)

        If Len(str2) = 0 Then
            GoTo end1
        End If

        Me.txtRDB.Text = str2

end1:

    End Sub

    Private Sub cmdBrowseDirectory_Click(sender As Object, e As EventArgs) Handles cmdBrowseDirectory.Click

        Dim strFilter As String
        Dim strFileName As String
        Dim str1 As String
        Dim str2 As String

        strFilter = "All files (*.*)|*.*"
        strFileName = "*.*"

        Dim strPath As String = ""

        str2 = ReturnDirectoryBrowse(False, strPath, strFilter, strFileName, True)

        If Len(str2) = 0 Then
            GoTo end1
        End If

        Me.txtDestinationPath.Text = str2

end1:

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        If ValidateForm() Then

            boolCancel = False
            Me.Visible = False

        End If



    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True

        Me.Visible = False

    End Sub

    Function ValidateForm() As Boolean

        ValidateForm = False

        Dim var1
        Dim strM As String = ""
        var1 = Me.txtRDB.Text
        If Len(var1) = 0 Then
            strM = "The Sciex" & ChrW(8482) & "Analyst .rdb file entry cannot be blank."
            GoTo end1
        End If
        If System.IO.File.Exists(var1) Then
        Else
            strM = "The Sciex" & ChrW(8482) & "Analyst .rdb file:" & ChrW(10) & ChrW(10) & var1 & ChrW(10) & ChrW(10) & "does not exist."
            GoTo end1
        End If

        var1 = Me.txtWordFileName.Text
        If Len(var1) = 0 Then
            strM = "The Word file name entry cannot be blank."
            GoTo end1
        End If
        If StrComp(var1, ".docx", CompareMethod.Text) = 0 Then
            strM = var1 & " is not a valid Word file name entry."
            GoTo end1
        End If
        'no special characters
        If HasSpecialCharacters(var1.ToString) Then
            GoTo end1
        End If


        var1 = Me.txtDestinationPath.Text
        If Len(var1) = 0 Then
            strM = "The Word destination path entry cannot be blank."
            GoTo end1
        End If
        If System.IO.Directory.Exists(var1) Then
        Else
            strM = "The Word destination path:" & ChrW(10) & ChrW(10) & var1 & ChrW(10) & ChrW(10) & "does not exist."
            GoTo end1
        End If


        ValidateForm = True

end1:

        If Len(strM) = 0 Then
        Else
            MsgBox(strM, vbInformation, "Invalid entry...")
        End If

    End Function

    Function HasSpecialCharacters(ByVal strVal As String) As Boolean

        HasSpecialCharacters = False
        Dim intL As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim varC
        Dim boolGo1 As Boolean = False
        Dim boolGo2 As Boolean = False
        Dim boolGo3 As Boolean = False
        Dim boolGo4 As Boolean = False

        intL = Len(strVal)
        For Count1 = 1 To intL
            str1 = Mid(strVal, Count1, 1)
            varC = AscW(str1)
            boolGo1 = False
            boolGo2 = False
            boolGo3 = False
            boolGo4 = False
            If (varC > 64 And varC < 91) Or (varC > 60 And varC < 123) Then 'letters OK
                boolGo1 = True
            End If

            If (varC > 47 And varC < 58) Then 'numbers OK
                boolGo2 = True
            End If

            If varC = 92 Or varC = 45 Then '_,-
                boolGo3 = True
            End If

            If boolGo1 Or boolGo2 Or boolGo3 Then
                HasSpecialCharacters = False
            Else
                HasSpecialCharacters = True
                Exit For
            End If

        Next

        If HasSpecialCharacters Then
            Dim strM As String
            strM = "Special characters are not allowed." & ChrW(10) & ChrW(10) & "' " & str1 & " ' is considered a special character."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        End If


    End Function

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles txtWordFileName.TextChanged

    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles lblEnterWordName.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.txtRDB.Text = "G:\Analyst Data\Projects\00176\Demo\Results\20131107-set2$3.rdb"
        'Me.txtRDB.Text = "G:\Analyst Data\Projects\00006\PKData_20140430\14PK0397_X2LCMS12\Results\14PK0397.rdb"
        Me.txtWordFileName.Text = "Test02"
        Me.txtDestinationPath.Text = "C:\Users\Gubbs.GUBBSINC\Desktop\Word Chrom Practice\"


    End Sub
End Class