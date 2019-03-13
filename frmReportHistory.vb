Option Compare Text

Public Class frmReportHistory

    Private Sub frmReportHistory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call DoubleBufferControl(Me, "dgv")

        Call FormLoad()

    End Sub

    Sub FormLoad()

        Dim dgv As DataGridView
        Dim dtbl As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim intRow As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String

        If id_tblReports = 0 Then
            intRow = frmH.dgvReports.CurrentRow.Index
            id_tblReports = frmH.dgvReports("ID_TBLREPORTS", intRow).Value
        End If

        dgv = Me.dgvReportHistory
        dtbl = tblReportHistory
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTS = " & id_tblReports
        strS = "DTREPORTGENERATED DESC"

        Dim dv As System.Data.DataView = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)

        dgv.DataSource = dv

        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns(Count1).Visible = False
        Next

        str1 = "ID_TBLREPORTS"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).DisplayIndex = 0
        dgv.Columns(str1).HeaderText = "Report ID"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        str1 = "DTREPORTGENERATED"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).DisplayIndex = 1
        dgv.Columns(str1).HeaderText = "Date/Time"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Dim strFormat As String = LDateFormat & " HH:mm:ss tt"
        dgv.Columns(str1).DefaultCellStyle.Format = strFormat

        str1 = "CHARREPORTTITLE"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).DisplayIndex = 2
        dgv.Columns(str1).HeaderText = "Report Title"
        'dgv.Columns("CHARREPORTTITLE").MinimumWidth = dgv.Width * 0.4

        str1 = "CHARREPORTGENERATEDSTATUS"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).DisplayIndex = 3
        dgv.Columns(str1).HeaderText = "Report Type"

        str1 = "CHARUSERID"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).DisplayIndex = 4
        dgv.Columns(str1).HeaderText = "User ID"

        str1 = "CHARUSERNAME"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).DisplayIndex = 5
        dgv.Columns(str1).HeaderText = "User Name"

        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True


        'ID_TBLREPORTHISTORY
        'ID_TBLSTUDIES
        'ID_TBLREPORTS
        'CHARREPORTGENERATEDSTATUS
        'UPSIZE_TS
        'DTREPORTGENERATED
        'CHARREPORTTITLE
        'CHARUSERID
        'CHARUSERNAME

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'dgv.AutoResizeColumns()
        dgv.Columns("CHARREPORTTITLE").Width = dgv.Width * 0.3
        dgv.Columns("DTREPORTGENERATED").Width = 55
        dgv.RowHeadersWidth = 25
        dgv.Columns("ID_TBLREPORTS").Width = 50
        dgv.Columns("CHARUSERID").Width = 60
        dgv.Columns("CHARUSERNAME").Width = 75
        dgv.AutoResizeRows()


        'check for lblWarning

        'Call CheckWatsonRecordsFromReportHistory(True)
        str1 = frmH.lblWarning.Text
        'replace some CR
        str2 = Replace(str1, ChrW(10) & ChrW(10), ChrW(10), 1, -1, CompareMethod.Text)
        str1 = str2
        str2 = Replace(str1, "View Report History", "View Unerlying Data...", 1, -1, CompareMethod.Text)
        Me.lblWarning.Text = str2

        Me.lblWarning.ForeColor = frmH.lblWarning.ForeColor
        Me.lblWarning.BackColor = frmH.lblWarning.BackColor
        Me.cmdVerify.Visible = True
        'If InStr(1, str1, "Warning", CompareMethod.Text) > 0 Then
        '    Me.cmdVerify.Visible = True
        'Else
        '    Me.cmdVerify.Visible = False
        'End If


        If gboolER Then
            Me.cmdShowReports.Visible = True
            'Me.cmdShowReports.Visible = True
            Me.cmdShowReports.Left = Me.lblTitle.Left + Me.lblTitle.Width + 25
            Me.lblWarning.Left = Me.cmdShowReports.Left + Me.cmdShowReports.Width + 10

        Else
            Me.cmdShowReports.Visible = False
            'Me.cmdShowReports.Visible = False
            Me.lblWarning.Left = Me.lblTitle.Left + Me.lblTitle.Width + 25
        End If
        Me.cmdVerify.Left = Me.lblWarning.Left + Me.lblWarning.Width + 10
        'If Me.cmdVerify.Visible Then
        '    Me.cmdExit.Left = Me.cmdVerify.Left + Me.cmdVerify.Width + 10
        'Else
        '    Me.cmdExit.Left = Me.lblWarning.Left + Me.lblWarning.Width + 10
        'End If
        Me.cmdExit.Left = Me.cmdVerify.Left + Me.cmdVerify.Width + 10

        Me.lblWarning.Visible = True



    End Sub


    Private Sub cmdVerify_Click(sender As System.Object, e As System.EventArgs) Handles cmdVerify.Click

        Dim dt As Date
        Dim intRow As Int32
        Dim intCol As Short
        Dim dgv As DataGridView = Me.dgvReportHistory
        Dim strM As String
        Dim var1

        intRow = 0

        Dim rows() As DataRow
        Dim dtR As Date
        Dim strF As String
        Dim strS As String
        Dim str1 As String
        Dim str2 As String

        Dim strCol As String = ReturnWatsonCheckColumn()

        'NEED dtR!!!!!


        'check to proceed
        str2 = Me.lblWarning.Text
        If InStr(1, str2, "Warning", CompareMethod.Text) > 0 Then
        Else
            'If gboolER Then
            '    str1 = "The last Final Report saved is current." & ChrW(10) & ChrW(10)
            '    'str1 = str1 & "No Watson samples have been modified since the last Final Report was saved:" & ChrW(10) & Format(dtR, "dd-MMM-yyyy hh:mm tt")
            '    str1 = str1 & str2
            'Else
            '    str1 = "The last Final Report generated is current." & ChrW(10) & ChrW(10)
            '    str1 = str1 & "No Watson samples have been modified since the last Final Report was generated:" & ChrW(10) & Format(dtR, "dd-MMM-yyyy hh:mm tt")
            '    str1 = str1 & str2
            'End If
            'strM = str1
            strM = str2
            MsgBox(strM, vbInformation, "Status...")
            GoTo end1
        End If

        Dim frm As New frmValidateReport
        frm.dtCutOff = gWatsonCutOffDt

        Call frm.FormLoad()

        frm.ShowDialog()

end1:


    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick

        Dim vc
        vc = Me.lblWarning.BackColor

        If vc = Color.FromArgb(255, 224, 192) Then
            Me.lblWarning.BackColor = Color.FromArgb(192, 255, 192)
        Else
            Me.lblWarning.BackColor = Color.FromArgb(255, 224, 192)
        End If

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Visible = False
    End Sub

    Private Sub cmdShowReports_Click(sender As Object, e As EventArgs) Handles cmdShowReports.Click

        Dim strM As String
        If BOOLVIEWFINALREPORT Then
        Else
            strM = "User is not allowed to view reports."
            MsgBox(strM, vbInformation, "Invalid Action...")
            GoTo end1
        End If

        Dim strP As String = ""

        Dim frm As New frmDocumentCompare

        frm.boolTemplate = False
        frm.gDoc = "Final Report"
        'frm.txtLoadedDocDescription.Text = frm.gDoc
        frm.gReport = strP
        frm.strPrevForm = "ReportHistory"
        frm.frm = Me
        frm.boolFromReportHistory = True
        frm.lblInstructions01.Visible = False
        'frm.gbSaveType.Visible = True
        frm.gbLoad.Visible = True
        frm.Text = "Document Control"

        Call frm.EnableControls(True)
        Cursor.Current = Cursors.Default
        frm.Show(Me)


        GoTo end1

        Dim var1
        Try
            frm.Dispose()
        Catch ex As Exception
            var1 = var1
        End Try

        Call ClearTemp()

        Cursor.Current = Cursors.Default
        Try
            Me.Activate()
        Catch ex As Exception

        End Try

        Try
            Me.Visible = True
        Catch ex As Exception

        End Try

end1:


    End Sub
End Class