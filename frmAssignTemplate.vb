Option Compare Text

Public Class frmAssignTemplate
    Public boolCancel As Boolean = True
    Public boolAllowApply As Boolean = False
    Public gLabel As String = ""

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Call TemplateOK()

    End Sub

    Sub TemplateOK()

        Dim int1 As Short
        Dim dgv As DataGridView = Me.dgvTemplates

        'ensure an item is selected
        'int1 = Me.lbxTemplates.SelectedIndex

        Dim dgr As DataGridViewRow

        Try
            dgr = dgv.SelectedRows(0)
            boolCancel = False
            Me.Visible = False
        Catch ex As Exception
            MsgBox("Select an item.", MsgBoxStyle.Information, "Select an item")
        End Try

    End Sub

    Sub SetOG()

        If Me.rbApply.Checked Then
            Me.cmdOK.Visible = True
        Else
            Me.cmdOK.Visible = False
        End If

    End Sub

    Sub SetLabel()

        Me.lbl1.Text = gLabel

    End Sub

    Private Sub frmAssignTemplate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim tbl As System.Data.DataTable
        Dim dtblS As System.Data.DataTable = tblStudies
        Dim rowsS() As DataRow
        Dim strFS As String

        Dim strF As String
        Dim strS As String
        Dim int1 As Short
        Dim Count1 As Short
        Dim var1, var2, var3
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        Call ControlDefaults(Me)

        Dim dgv As DataGridView = Me.dgvTemplates

        Dim dtbl1 As New System.Data.DataTable

        For Count1 = 1 To 2

            Select Case Count1
                Case 1
                    str1 = "charStudyTemplate"
                    str2 = "Study Template"
                Case 2
                    str1 = "charWatsonStudy"
                    str2 = "Watson Study"
            End Select

            Dim col1 As New DataColumn
            col1.ColumnName = str1
            col1.Caption = str2
            col1.DataType = System.Type.GetType("System.String")
            dtbl1.Columns.Add(col1)

        Next

        tbl = tblTemplates
        strF = "boolActive = -1"
        strS = "charTemplateName ASC"
        Dim rowsT() As DataRow = tbl.Select(strF, strS)

        For Count1 = 0 To rowsT.Length - 1
            var1 = rowsT(Count1).Item("charTemplateName")
            var2 = rowsT(Count1).Item("ID_TBLSTUDIES")
            strFS = "ID_TBLSTUDIES = " & var2
            rowsS = dtblS.Select(strFS)
            var3 = rowsS(0).Item("CHARWATSONSTUDYNAME")
            'Me.lbxTemplates.Items.Add(var1)
            ' Me.lbxTemplates.Items.AddRange(New Object() {var1.ToString, var3.ToString})
            Dim nr As DataRow = dtbl1.NewRow
            nr.BeginEdit()
            nr("charStudyTemplate") = var1
            nr("charWatsonStudy") = var3
            dtbl1.Rows.Add(nr)

        Next

        dgv.DataSource = dtbl1
        For Count1 = 1 To 2
            Select Case Count1
                Case 1
                    str1 = "charStudyTemplate"
                    str2 = "Study Template"
                Case 2
                    str1 = "charWatsonStudy"
                    str2 = "Watson Study"
            End Select

            dgv.Columns(Count1 - 1).HeaderText = str2
        Next

        dgv.RowHeadersWidth = dgv.RowHeadersWidth * 0.5
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font(dgv.Font, FontStyle.Bold)

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'do Home Tab
        'Word Document Template

        Call SetOG()

    End Sub

    Private Sub lbxTemplates_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)

        If Me.rbApply.Checked Then
            Call TemplateOK()
        End If

    End Sub

    Private Sub rbView_CheckedChanged(sender As Object, e As EventArgs) Handles rbView.CheckedChanged

        Call SetOG()

    End Sub

    Private Sub rbApply_CheckedChanged(sender As Object, e As EventArgs) Handles rbApply.CheckedChanged

        Call SetOG()

    End Sub
End Class