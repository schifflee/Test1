Option Compare Text

Public Class frmConfigCompliance

    Private Sub frmConfigCompliance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ConfigMOFandRFC()

        Call DoEnableCompliance()

    End Sub

    Sub DoEnableCompliance()

        If Me.rbAuditTrailOn.Checked Then
            Me.gbESig.Enabled = True
        Else
            Me.gbESig.Enabled = False
        End If

        If Me.rbESigOn.Checked Then
            Me.panESigOptions.Enabled = True
        Else
            Me.panESigOptions.Enabled = False
        End If

    End Sub

    Sub ConfigMOFandRFC()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView

        dgv1 = Me.dgvMOS
        dgv2 = Me.dgvRFC

        Dim dtbl1 As System.Data.DataTable = tblMeaningOfSig
        Dim dtbl2 As System.Data.DataTable = tblReasonForChange

        Dim dv1 As System.Data.DataView = New DataView(dtbl1)
        Dim dv2 As System.Data.DataView = New DataView(dtbl2)

        dv1.AllowDelete = False
        dv1.AllowEdit = False
        dv1.AllowNew = False

        dv2.AllowDelete = False
        dv2.AllowEdit = False
        dv2.AllowNew = False

        dgv1.DataSource = dv1
        dgv2.DataSource = dv2

        dgv1.Columns("ID_TBLMEANINGOFSIG").Visible = False
        dgv1.Columns("CHARMEANINGOFSIG").Visible = True
        dgv1.Columns("CHARMEANINGOFSIG").HeaderText = "Meaning of Signature"
        dgv1.Columns("INTORDER").Visible = True
        dgv1.Columns("INTORDER").HeaderText = "Order"
        dgv1.Columns("BOOLINCLUDE").Visible = False

        dgv2.Columns("ID_TBLREASONFORCHANGE").Visible = False
        dgv2.Columns("CHARREASONFORCHANGE").Visible = True
        dgv2.Columns("CHARREASONFORCHANGE").HeaderText = "Reason For Change"
        dgv2.Columns("INTORDER").Visible = True
        dgv2.Columns("INTORDER").HeaderText = "Order"
        dgv2.Columns("BOOLINCLUDE").Visible = False

    End Sub

    Private Sub rbAuditTrailOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAuditTrailOff.CheckedChanged

        Call DoEnableCompliance()

    End Sub

    Private Sub rbESigOff_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbESigOff.CheckedChanged

        Call DoEnableCompliance()

    End Sub
End Class