
Imports System.Text.RegularExpressions

Public Class frmReportTemplateAdd

    Public boolShowVersions As Boolean = False
    Public boolCancel As Boolean = False


    Private Sub frmReportTemplateAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call DoubleBufferControl(Me, "dgv")

        Call ControlDefaults(Me)

        Call SizeForm()

        Call DoEnables()

        Call frmReportTemplateAdd_ToolTipSet()

    End Sub

    Private Sub chkChoose_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Sub panListShow()




    End Sub

    Sub SizeForm()

        Dim a, b, c, d, e, f

        If boolShowVersions Then
            a = Me.dgvVersions.Top + Me.dgvVersions.Height
        Else
            a = Me.lblVersions.Top - 1
        End If

        Me.panList.Height = a

        Dim borderwidth As Single = (Me.Width - Me.ClientSize.Width) / 2
        Dim titleBarHeight As Single = Me.Height - Me.ClientSize.Height - borderwidth

        b = Me.panList.Top + Me.panList.Height + 10

        Me.gbChoice.Height = b

        c = Me.gbChoice.Top + Me.gbChoice.Height + 10 + borderwidth + titleBarHeight

        Me.Height = c


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim strM As String
        Dim intR As String

        strM = "Click Yes for True, No for False"
        intR = MsgBox(strM, vbYesNoCancel, "Enter...")

        If intR = 6 Then 'yes
            boolShowVersions = True
        ElseIf intR = 7 Then 'no
            boolShowVersions = False
        Else
            GoTo end1
        End If

        Call SizeForm()

        Call ReportStatementsChange()

end1:

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        If ValidateOK() Then

            boolCancel = False

            Me.Visible = False

        End If

    End Sub

    Function ValidateOK() As Boolean

        ValidateOK = False

        Dim strPath As String
        Dim strM As String
        Dim strF As String
        Dim str1 As String

        If Me.rbBlank.Checked Then
            'do nothing
        ElseIf Me.rbDocument.Checked Then
            'ensure path exists
            strPath = Me.txtFilePath.Text

            If My.Computer.FileSystem.FileExists(strPath) Then
            Else
                strM = "The file:" & ChrW(10) & ChrW(10) & strPath & ChrW(10) & ChrW(10) & "does not exist."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If


        ElseIf Me.rbTemplate.Checked Then

        Else
            GoTo end1
        End If

        'now inspect strName
        Dim strName As String = Me.txtName.Text

        If Len(strName) = 0 Then
            strM = "Report Template Name cannot be blank."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        'check for special characters
        If HasSpecialCharacters(strName) > 0 Then
            strM = "Report Template Name cannot contain special characters." & ChrW(10) & ChrW(10) & "Acceptable characters are: [space], a - z, A - Z, 0 - 9, [dash], [underscore]"

            '[space]
            'a -z
            'A -Z
            '0-9
            '-
            '_

            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        'Cannot be a replicate of an existing name
        strF = "CHARTITLE = '" & strName & "'"
        Dim rows() As DataRow = tblWordStatements.Select(strF)
        If rows.Length > 0 Then
            str1 = NZ(rows(0).Item("CHARWORDSTATEMENT"), "Active")
            strM = "The entered Report Template Name '" & strName & "' already exists as an " & str1 & " report template."
            strM = strM & ChrW(10) & "Please enter a different name."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If


        ValidateOK = True

end1:

    End Function

    Function HasSpecialCharacters(strC As String) As Short

        'This function is used in ReportTableConfig - AutoAssign columns cell validating

        HasSpecialCharacters = 0

        '        These are valid characters:

        '[space]
        'a -z
        'A -Z
        '0-9
        '-
        '_


        'http://stackoverflow.com/questions/3701018/remove-special-characters-from-a-string
        'Dim cleanString As String = Regex.Replace(yourString, "[^A-Za-z0-9\-/]", "")
        Dim cs As String ' = Regex.Replace(strC, "[^A-Za-z0-9\-_. ]", "")
        cs = Regex.Replace(strC, "[^A-Za-z0-9\-_ ]", "")
        'Dim cs As String = Regex.Replace(strC, "[^A-Za-z0-9\-_. ]", "")

        Dim int1 As Short
        int1 = Math.Abs(Len(cs) - Len(strC))
        HasSpecialCharacters = int1

    End Function

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True

        Me.Visible = False

    End Sub

    Private Sub dgvReportStatements_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReportStatements.CellContentClick

    End Sub

    Private Sub dgvReportStatements_SelectionChanged(sender As Object, e As EventArgs) Handles dgvReportStatements.SelectionChanged

        Call ReportStatementsChange()

    End Sub

    Sub ReportStatementsChange()


        Dim id As Int64
        Dim dgv1 As DataGridView = Me.dgvReportStatements
        Dim dgv2 As DataGridView = Me.dgvVersions
        Dim intRow As Int32

        If dgv1.CurrentRow Is Nothing Then
            If dgv1.RowCount = 0 Then
                GoTo end1
            Else
                dgv1.Rows(0).Selected = True
            End If
        End If

        intRow = dgv1.CurrentRow.Index

        id = dgv1("ID_TBLWORDSTATEMENTS", intRow).Value

        Dim strF As String
        Dim strS As String

        strF = "ID_TBLWORDSTATEMENTS = " & id
        strS = "INTWORDVERSION DESC"

        Dim dv As DataView

        Try

            dv = dgv2.DataSource
            dv.RowFilter = strF
            dv.Sort = strS

            dgv1.AutoResizeRows()
            dgv2.AutoResizeRows()
            dgv2.AutoResizeColumns()

        Catch ex As Exception

        End Try


end1:

    End Sub

    Sub BrowseForFile()

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim strF As String
        Dim strPath As String

        Dim boolIsFile As Boolean = True

        'get default path
        strPath = "C:\"

        Dim strFilter As String
        Dim strFileName As String
        'Me.ov1.OpenFileDialog("Microsoft Word Files(*.doc*)|*.doc*||")
        strFilter = "Word files (*.doc*)|*.doc*"
        strFileName = "*.doc*"

        If boolIsFile = -1 Then
            str2 = ReturnDirectoryBrowse(True, strPath, strFilter, strFileName, True)
        Else
            str2 = ReturnDirectoryBrowse(False, strPath, strFilter, strFileName, True)
        End If

        If Len(str2) = 0 Then
        Else
            Me.txtFilePath.Text = str2
        End If

    End Sub

    Private Sub cmdBrowse_Click(sender As Object, e As EventArgs) Handles cmdBrowse.Click

        Call BrowseForFile()

    End Sub

    Sub DoEnables()

        If Me.rbBlank.Checked Then

            Me.txtFilePath.Enabled = False
            'Me.panList.Enabled = False
            Me.cmdBrowse.Enabled = False

        ElseIf Me.rbDocument.Checked Then

            Me.txtFilePath.Enabled = True
            'Me.panList.Enabled = False
            Me.cmdBrowse.Enabled = True

        ElseIf Me.rbTemplate.Checked Then

            Me.txtFilePath.Enabled = False
            'Me.panList.Enabled = True
            Me.cmdBrowse.Enabled = False

        End If

    End Sub

    Private Sub rbBlank_CheckedChanged(sender As Object, e As EventArgs) Handles rbBlank.CheckedChanged

        Call DoEnables()

    End Sub

    Private Sub rbDocument_CheckedChanged(sender As Object, e As EventArgs) Handles rbDocument.CheckedChanged

        Call DoEnables()

    End Sub

    Private Sub rbTemplate_CheckedChanged(sender As Object, e As EventArgs) Handles rbTemplate.CheckedChanged

        Call DoEnables()

    End Sub

    Private Sub frmReportTemplateAdd_ToolTipSet()

        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As New ToolTip()

        ' Set up the delays for the ToolTip.
        'toolTip1.AutoPopDelay = 5000
        'toolTip1.InitialDelay = 250
        'toolTip1.ReshowDelay = 50

        toolTip1.AutomaticDelay = intToolTipDelay
        'toolTip1.UseFading = False
        'tooltip1.
        'toolTip1.BackColor = Color.Goldenrod
        'toolTip1.IsBalloon = True
        ' Force the ToolTip text to be displayed whether or not the form is active.
        toolTip1.ShowAlways = True

        Dim strM As String

        strM = "Report Template Name cannot contain special characters." & ChrW(10) & "Acceptable characters are: [space], a - z, A - Z, 0 - 9, [dash], [underscore]"

        Try

            toolTip1.SetToolTip(Me.txtName, strM)
            toolTip1.SetToolTip(Me.lblEnter, strM)

        Catch ex As Exception

        End Try

    End Sub

End Class