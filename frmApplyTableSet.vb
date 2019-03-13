Public Class frmApplyTableSet

    Public boolCancel As Boolean = True

    Private Sub frmApplyTableSet_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Call FillGrid()

    End Sub

    Private Sub FillGrid()

        Dim boolT As Boolean = True

        If Me.rbTemplate.Checked Then
            boolT = True
        Else
            boolT = False
        End If

        Dim dgv1 As DataGridView = Me.dgvSource
        Dim dgvRT As DataGridView = frmH.dgvReportStatementWord
        Dim dgvS As DataGrid = frmH.dgStudies
        Dim dv As DataView
        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim str1 As String
        Dim str2 As String

        Dim int1 As Int16

        Dim dtbl As DataTable

        Dim var1

        If boolT Then
            dgv1.DataSource = dgvRT.DataSource
        Else
            dgv1.DataSource = dgvS.DataSource
        End If

        dgv1.ColumnHeadersDefaultCellStyle.Font = New Font(dgv1.Font, FontStyle.Bold)
        dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        For Count1 = 0 To dgv1.ColumnCount - 1
            dgv1.Columns(Count1).Visible = False
        Next

        If boolT Then

            str1 = "CHARTITLE"
            str2 = "Report Template"

            dgv1.Columns(str1).Visible = True
            Try
                dgv1.Columns(str1).HeaderText = str2
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try
            Try
                dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try
            '
            '

        End If



    End Sub

    Private Sub rbTemplate_CheckedChanged(sender As Object, e As EventArgs) Handles rbTemplate.CheckedChanged

        Call FillGrid()

    End Sub

    Private Sub rbStudy_CheckedChanged(sender As Object, e As EventArgs) Handles rbStudy.CheckedChanged

        Call FillGrid()

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click

        boolCancel = False

        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click

        boolCancel = True

        Me.Visible = False

    End Sub
End Class