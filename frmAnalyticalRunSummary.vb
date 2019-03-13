Option Compare Text



Public Class frmAnalyticalRunSummary

    Public boolFormLoad As Boolean = False

    Private Sub frmAnalyticalRunSummary_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call DoubleBufferControl(Me, "dgv")

        Call FormLoad()

    End Sub

    Sub SetControls()

        Dim a, b, c, d
        Dim h, w


        h = Screen.PrimaryScreen.WorkingArea.Height
        w = Screen.PrimaryScreen.WorkingArea.Width

        a = w - 50
        Me.Left = 25
        Me.Width = a

        Me.Top = 25
        b = h - 50
        Me.Height = b


    End Sub

    Sub FormatControls()

        Dim dgv As DataGridView = Me.dgvAnalRunSummary

        dgv.Columns("Samples").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        dgv.AutoResizeColumns()

    End Sub

    Sub FormLoad()

        boolFormLoad = True

        Call SetControls()

        'Dim dgvS As DataGridView
        'Dim dgvD As DataGridView

        'dgvS = frmH.dgvAnalyticalRunSummary
        'dgvD = Me.dgvAnalRunSummary

        'dgvD.DataSource = dgvS.DataSource


        Me.chkAll.Checked = frmH.chkAll.Checked
        Me.chkAccepted.Checked = frmH.chkAccepted.Checked
        Me.chkRejected.Checked = frmH.chkRejected.Checked
        Me.chkRegrPerformed.Checked = frmH.chkRegrPerformed.Checked
        Me.chkNoRegrPerformed.Checked = frmH.chkNoRegrPerformed.Checked
        Me.chkPSAE.Checked = frmH.chkPSAE.Checked

        boolFormLoad = False

        Call FillAnalRunSum()

        Call FormatControls()

        Me.dgvAnalRunSummary.Columns.Item("Samples").HeaderText = "Run Description"

    End Sub

    Sub FillAnalRunSum()

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgvD As DataGridView = Me.dgvAnalRunSummary
        Dim dgvS As DataGridView = frmH.dgvAnalyticalRunSummary

        Dim strF As String

        Dim int1 As Int16
        Dim Count1 As Int16
        Dim var1, var2

        Dim strF1 As String = ""
        Dim strF2 As String = ""
        Dim strF3 As String = ""
        Dim strF4 As String = ""
        Dim strF5 As String = ""
        Dim strF6 As String = ""
        Dim strFTot As String = ""

        If Me.chkAll.Checked Then
            strF1 = "RUNTYPEID > 0 AND boolInThisRunsAssayID = 'Yes' AND RUNANALYTEREGRESSIONSTATUS > -2" 'do -2 here because blank row has -1
            '20180716 LEE:
            'new logic
            strF1 = "RUNTYPEID > 0 AND RUNANALYTEREGRESSIONSTATUS > -2" 'do -2 here because blank row has -1
            strFTot = strF1
        Else
            If Me.chkAccepted.Checked Then
                strF2 = "RUNANALYTEREGRESSIONSTATUS = 3"
                If Len(strFTot) = 0 Then
                    strFTot = strF2
                Else
                    strFTot = strFTot & " OR " & strF2
                End If
            End If
            If Me.chkRejected.Checked Then
                strF3 = "RUNANALYTEREGRESSIONSTATUS = 4"
                If Len(strFTot) = 0 Then
                    strFTot = strF3
                Else
                    strFTot = strFTot & " OR " & strF3
                End If
            End If
            If Me.chkRegrPerformed.Checked Then
                strF4 = "RUNANALYTEREGRESSIONSTATUS = 2"
                If Len(strFTot) = 0 Then
                    strFTot = strF4
                Else
                    strFTot = strFTot & " OR " & strF4
                End If
            End If
            If Me.chkNoRegrPerformed.Checked Then
                strF5 = "RUNANALYTEREGRESSIONSTATUS = 1"
                If Len(strFTot) = 0 Then
                    strFTot = strF5
                Else
                    strFTot = strFTot & " OR " & strF5
                End If
            End If
            If Me.chkPSAE.Checked Then
                strF6 = "RUNTYPEID = 3"
                If Len(strFTot) = 0 Then
                    strFTot = strF6
                Else
                    'strFTot = "(" & strFTot & ") AND " & strF6
                    strFTot = "(" & strFTot & ") OR " & strF6
                End If
            Else
                strF6 = "RUNTYPEID <> 3"
                If Len(strFTot) = 0 Then
                    strFTot = "RUNANALYTEREGRESSIONSTATUS = -10" 'strF6' strF6
                Else
                    strFTot = "(" & strFTot & ") AND " & strF6
                End If
            End If


            '20180716 LEE:
            If Len(strFTot) = 0 Then
                strFTot = "RUNANALYTEREGRESSIONSTATUS = -10"
            Else
                'strFTot = "(" & strFTot & ") AND boolInThisRunsAssayID = 'Yes' OR RUNANALYTEREGRESSIONSTATUS = -1" 'RUNANALYTEREGRESSIONSTATUS = -1 to get blank rows
                '20180717 LEE
                strFTot = "(" & strFTot & ") OR RUNANALYTEREGRESSIONSTATUS = -1" 'RUNANALYTEREGRESSIONSTATUS = -1 to get blank rows
            End If

        End If


        ''''''''''''''''console.writeline(strF)
        Dim dtbl As DataTable = tblAnalRunSum.Copy

        Dim dv2 As System.Data.DataView = New DataView(dtbl)
        dv2.RowFilter = strFTot
        dgvD.DataSource = dv2

        For Count1 = 0 To dgvS.Columns.Count - 1
            dgvD.Columns(Count1).Visible = dgvS.Columns(Count1).Visible
        Next


    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim dgv As DataGridView = Me.dgvAnalRunSummary
        Dim Count1 As Short

        For Count1 = 0 To dgv.Columns.Count - 1
            'console.writeline(dgv.Columns(Count1).Name)
        Next


    End Sub

    Private Sub chkAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkAll.CheckedChanged

        Call FillAnalRunSum()

    End Sub

    Private Sub chkAccepted_CheckedChanged(sender As Object, e As EventArgs) Handles chkAccepted.CheckedChanged

        Call FillAnalRunSum()

    End Sub

    Private Sub chkRejected_CheckedChanged(sender As Object, e As EventArgs) Handles chkRejected.CheckedChanged

        Call FillAnalRunSum()

    End Sub

    Private Sub chkRegrPerformed_CheckedChanged(sender As Object, e As EventArgs) Handles chkRegrPerformed.CheckedChanged

        Call FillAnalRunSum()

    End Sub

    Private Sub chkNoRegrPerformed_CheckedChanged(sender As Object, e As EventArgs) Handles chkNoRegrPerformed.CheckedChanged

        Call FillAnalRunSum()

    End Sub

    Private Sub chkPSAE_CheckedChanged(sender As Object, e As EventArgs) Handles chkPSAE.CheckedChanged

        Call FillAnalRunSum()

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click

        Me.Dispose()

    End Sub
End Class