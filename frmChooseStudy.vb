Option Compare Text

Public Class frmChooseStudy
    Public boolCancel As Boolean = True
    Dim boolCont As Boolean = True
    Private Sub frmChooseStudy_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        boolCancel = True

        Call ControlDefaults(Me)

        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim rows() As DataRow
        Dim intRows As Short
        Dim Count1 As Short
        Dim str1 As String

        tbl = tblConfigReportType
        strF = "id_tblConfigReportType < 1000"
        rows = tbl.Select(strF)
        intRows = rows.Length

        Me.lvStudyType.View = System.Windows.Forms.View.List

        For Count1 = 0 To intRows - 1
            str1 = NZ(rows(Count1).Item("charReportType"), "Sample Analysis")
            Me.lvStudyType.Items.Add(str1)
        Next

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        boolCancel = False
        Me.Visible = False

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        boolCancel = True
        Me.Visible = False
    End Sub

    Private Sub lvStudyType_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles lvStudyType.ItemCheck
        'deselect everything else
        Dim introws As Short
        Dim Count1 As Short

        If boolCont = False Then
            Exit Sub
        End If
        boolCont = False
        introws = Me.lvStudyType.Items.Count
        For Count1 = 0 To introws - 1

            If Me.lvStudyType.Items(Count1).Selected Then
            Else
                Me.lvStudyType.Items(Count1).Checked = False
            End If

            If Count1 = introws - 1 Then
                boolCont = True
            End If
        Next

    End Sub

End Class