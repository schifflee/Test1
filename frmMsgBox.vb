Option Compare Text

Public Class frmMsgBox
    Public boolCancel As Boolean = True
    Public boolKill As Boolean = False
    Public boolTable As Boolean = False
    Public dv As DataView
    Public boolFormLoad As Boolean = True
    Public intRowDGV As Short = 0
    Public tblDGV As DataTable
    Public rowsDGV() As DataRow

    Private Sub lbxAnalytes_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles lbxAnalytes.ItemCheck

        If boolFormLoad Then
            Exit Sub
        End If

        Call lbxAnalytesChange(Me.lbxAnalytes.SelectedIndex, e.NewValue)

end1:

    End Sub


    Sub lbxAnalytesChange(intI As Short, intCheck As Short)

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgv As DataGridView = frmH.dgvReportTableConfiguration

        Dim Count1 As Short
        Dim tbl As DataTable = dv.Table ' dv.ToTable
        Dim int1 As Short = tbl.Columns.Count
        Dim str1 As String
        Dim boolI As Boolean
        Dim intIdv As Short
        Dim intStart As Short

        For Count1 = int1 - 1 To 0 Step -1
            str1 = tbl.Columns(Count1).ColumnName
            If StrComp(str1, "CHARTABLENAME", CompareMethod.Text) = 0 Then
                intStart = Count1 + 1
                Exit For
            End If
        Next

        Dim idR As Long = dv(intRowDGV).Item("ID_TBLREPORTTABLE")
        Dim strF As String = "ID_TBLREPORTTABLE = " & idR

        'boolI = lbxAnalytes.GetItemCheckState(intI)

        If intCheck = 1 Then
            boolI = True
        Else
            boolI = False
        End If

        intIdv = intStart + intI

        rowsDGV(0).BeginEdit()
        rowsDGV(0).Item(intIdv) = boolI
        rowsDGV(0).EndEdit()

end1:

    End Sub

    Private Sub cmdSelect_Click(sender As Object, e As EventArgs) Handles cmdSelect.Click

        Call FillListBox(False, True, False)

    End Sub

    Private Sub cmdDeselect_Click(sender As Object, e As EventArgs) Handles cmdDeselect.Click

        Call FillListBox(False, False, True)

    End Sub


    Sub FillListBox(boolSelectDGV, boolSelectAll, boolSelectNone)

        If boolTable Then
            Me.panAnalytes.Visible = True
        Else
            Me.panAnalytes.Visible = False
            Dim a, b, c, d

            a = Me.gb1.Left
            b = Me.gb1.Width

            Dim BorderWidth As Integer = (Me.Width - Me.ClientSize.Width) / 2
            Dim TitlebarHeight As Integer = Me.Height - Me.ClientSize.Height - 2 * BorderWidth

            Me.Width = a + b + a + (BorderWidth * 2)

            Me.StartPosition = FormStartPosition.CenterScreen

            GoTo end1

        End If

        Dim lbx As CheckedListBox = Me.lbxAnalytes

        Dim dgv As DataGridView = frmH.dgvReportTableConfiguration
        Dim idR As Long
        Dim intRow As Short

        Dim boolFL As Boolean = boolFormLoad

        Try
            intRowDGV = dgv.CurrentRow.Index
        Catch ex As Exception
            Me.panAnalytes.Visible = False
            GoTo end1
        End Try

        'Start
        'ID_TBLREPORTTABLE
        'ID_TBLSTUDIES
        'ID_TBLCONFIGREPORTTABLES
        'INTORDER
        'CHARPAGEORIENTATION
        'BOOLINCLUDE
        'BOOLREQUIRESSAMPLEASSIGNMENT
        'UPSIZE_TS
        'CHARSTABILITYPERIOD
        'CHARHEADINGTEXT
        'CHARSTYLE
        'INTEGNUM
        'CHARFCID
        'BOOLPLACEHOLDER
        'CHARTABLENAME
        'CRC015
        'Sirolimus_C1
        'Sirolimus_C2
        'Sirolimus_C3
        'End

        'last non-analyte column: CHARTABLENAME

        Dim Count1 As Short
        Dim tbl As DataTable = dv.ToTable
        Dim int1 As Short = tbl.Columns.Count
        Dim str1 As String
        Dim boolI As Boolean
        Dim intI As Short = -1
        Dim intStart As Short

        idR = dgv("ID_TBLREPORTTABLE", intRowDGV).Value

        'If boolFormLoad Then
        '    tblDGV = dv.Table
        '    rowsDGV = tblDGV.Select("ID_TBLREPORTTABLE = " & idR)
        'End If

        If boolSelectDGV Then
            lbx.Items.Clear()
        End If


        For Count1 = int1 - 1 To 0 Step -1
            str1 = tbl.Columns(Count1).ColumnName
            If StrComp(str1, "CHARTABLENAME", CompareMethod.Text) = 0 Then
                intStart = Count1 + 1
                Exit For
            End If
        Next

        boolFormLoad = True

        For Count1 = intStart To int1 - 1

            str1 = tbl.Columns(Count1).ColumnName
            If boolSelectDGV Then
                lbx.Items.Add(str1)
            End If

            boolI = dv(intRowDGV).Item(str1)
            intI = intI + 1
            If boolSelectDGV Then
                boolFormLoad = True
                If boolI Then
                    lbx.SetItemCheckState(intI, CheckState.Checked)
                Else
                    lbx.SetItemCheckState(intI, CheckState.Unchecked)
                End If
            ElseIf boolSelectAll Then
                boolFormLoad = True
                lbx.SetItemCheckState(intI, CheckState.Checked)
                boolFormLoad = False
                Call lbxAnalytesChange(intI, 1)
            ElseIf boolSelectNone Then
                boolFormLoad = True
                lbx.SetItemCheckState(intI, CheckState.Unchecked)
                boolFormLoad = False
                Call lbxAnalytesChange(intI, 0)
            End If

        Next

        boolFormLoad = boolFL

end1:

    End Sub

    Private Sub frmMsgBox_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        If boolTable Then

            If Me.boolCancel Then
                rowsDGV(0).RejectChanges()
            End If

        End If

end1:

    End Sub

    Private Sub frmMsgBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        boolFormLoad = True

        Call ControlDefaults(Me)

        'first get template
        Dim boolT As Boolean = ApplyReportTemplate(Me, "Report Prelim")

        'Legend
        boolSampleName01 = False
        boolExcludeTableNumbers = False
        boolExcludeTableTitles = False
        boolExcludeEntireTableTitle = False
        boolIncludeWaterMark = False
        boolExcludeCoverPage = False
        boolDisableWarnings = False
        boolExcludeHeaderFooter = False
        'boolForceWatermark = False

        Me.cbxReportTemplate.MaxDropDownItems = 8

        boolCancel = True

        'position form
        Dim t1, t2, h1, l1, l2

        t1 = frmH.panPrepareReportOutside.Top
        l1 = frmH.panPrepareReportOutside.Left
        h1 = frmH.panPrepareReportOutside.Height

        t2 = t1 + h1 + 50
        l2 = l1

        Me.Top = t2
        Me.Left = l2

        Dim str1 As String
        str1 = "Create as PDF (must have Word" & ChrW(8482) & " 2007 or greater)"
        Me.chkDoPDF.Text = str1

        boolVerbose = False

        Call FillcbxReportTemplate()

        boolKill = False
        If FillcbxReportTemplate() Then
            Call CheckPermissions()

            If frmH.chkReadOnlyTables.Checked Then
                gboolReadOnlyTables = True
                Me.chkReadOnlyTables.Checked = True
            Else
                gboolReadOnlyTables = False
                Me.chkReadOnlyTables.Checked = False
            End If
        Else
            boolKill = True
        End If

        Call AdvSettings()
        Call frmMsgBox_ToolTipSet()

        Dim dgv As DataGridView = frmH.dgvReportTableConfiguration
        dv = dgv.DataSource

        Dim boolAE As Boolean = dv.AllowEdit 'debug

        'dv.AllowEdit = True

        Call FillListBox(True, False, False)

        Me.AcceptButton = Me.cmdOK
        Me.cmdOK.Focus()

        boolFormLoad = False

    End Sub

    Function FillcbxReportTemplate() As Boolean

        FillcbxReportTemplate = True

        Dim dv As DataView = frmH.dgvReportStatementWord.DataSource
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strM As String
        Dim strD As String

        Dim dgvRS As DataGridView = frmH.dgvReportStatements
        Dim dgvRSW As DataGridView = frmH.dgvReportStatementWord

        Me.cbxReportTemplate.Items.Clear()

        'The default in dgvRS may not be active, but allow anyway
        Dim boolHit As Boolean
        'first check to see default exists in active
        boolHit = False
        strD = ""
        Try
            strD = dgvRS("CHARSTATEMENT", 0).Value ' this is dgvRSW("CHARTITLE")
        Catch ex As Exception
            strD = ""
        End Try

        'check to see if it already exists
        boolHit = False
        For Count1 = 0 To dgvRSW.RowCount - 1
            str2 = dgvRSW("CHARTITLE", Count1).Value
            If StrComp(strD, str2, CompareMethod.Text) = 0 Then
                boolHit = True
                Exit For
            End If
        Next

        If boolHit Then
        Else
            'add this to cbx
            Me.cbxReportTemplate.Items.Add(strD)
        End If

        For Count1 = 0 To dv.Count - 1
            str1 = dv(Count1).Item("CHARTITLE")
            Me.cbxReportTemplate.Items.Add(str1)
        Next

        ''determine if default is available
        'boolHit = False
        'str2 = NZ(frmH.dgvReportStatements("CHARSTATEMENT", 0).Value, "NA")
        'For Count1 = 0 To Me.cbxReportTemplate.Items.Count - 1
        '    str1 = Me.cbxReportTemplate.Items(Count1).ToString
        '    If StrComp(str2, str1, CompareMethod.Text) = 0 Then
        '        boolHit = True
        '        Exit For
        '    End If
        'Next

        'If boolHit Then
        'Else
        '    strM = "The default Report Template for this study" & ChrW(10) & ChrW(10) & str2 & ChrW(10) & ChrW(10)
        '    strM = strM & "Does not exist in the list of available Report Templates." & ChrW(10) & ChrW(10)
        '    strM = strM & "It is probable that the Report Template was renamed or deleted." & ChrW(10) & ChrW(10)
        '    strM = strM & "Please select the Report Template Configuration page and assign a new default Report Template for this study."
        '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        '    FillcbxReportTemplate = False
        '    GoTo end1
        'End If

        'select assigned report
        Me.cbxReportTemplate.Text = strD

end1:


    End Function

    Sub CheckPermissions()

        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short
        Dim strL1 As String
        Dim strL2 As String
        Dim strUser As String
        Dim strM As String

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        Dim rows() As DataRow
        rows = tblPermissions.Select(strF)

        Dim rowsU() As DataRow
        strF = "id_tblUserAccounts = " & id_tblUserAccounts
        rowsU = tblUserAccounts.Select(strF)

        'BOOLALLOWREPORTGENERATION
        'BOOLALLOWPDFREPORT

        'BOOLALLOWPDFREPORT = NZ(rows(0).Item("BOOLALLOWPDFREPORT"), 0)
        'BOOLALLOWREPORTGENERATION = NZ(rows(0).Item("BOOLALLOWREPORTGENERATION"), 0)
        'BOOLALLOWREPORTGENERATION = NZ(rows(0).Item("BOOLALLOWREPORTGENERATION"), 0)



        If BOOLALLOWPDFREPORT = 0 And BOOLALLOWREPORTGENERATION = 0 Then 'disallow everything
            strM = "User '" & strUser & "' does not have permission to generate any kind of report"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            boolCancel = True
            Me.Visible = False
        Else
            If BOOLALLOWREPORTGENERATION = 0 And BOOLALLOWPDFREPORT <> 0 Then 'disallow Word, force pdf
                Me.chkDoPDF.Checked = True
                Me.chkDoPDF.Enabled = False
            ElseIf BOOLALLOWREPORTGENERATION <> 0 And BOOLALLOWPDFREPORT = 0 Then 'don't allow pdf
                Me.chkDoPDF.Checked = False
                Me.chkDoPDF.Enabled = False
            ElseIf BOOLALLOWREPORTGENERATION <> 0 And BOOLALLOWPDFREPORT <> 0 Then 'allow everything
            End If
        End If

        If BOOLALLOWPDFREPORT = 0 Then
            Me.lblPDF.Text = "User IS NOT allowed to generate PDF"
        Else
            Me.lblPDF.Text = "User IS allowed to generate PDF"
        End If


        'If BOOLALLOWFINALREPORTWORD = 0 Then
        '    Me.lblWord.Text = "User IS NOT allowed to generate Microsoft" & ChrW(174) & " Word document"
        'Else
        '    Me.lblWord.Text = "User IS allowed to generate Microsoft" & ChrW(174) & " Word document"
        'End If

        If BOOLALLOWFINALREPORTWORD = 0 Then
            Me.lblWord.Text = "User IS NOT allowed to open document in Microsoft" & ChrW(174) & " Word"
        Else
            Me.lblWord.Text = "User IS allowed to open document in Microsoft" & ChrW(174) & " Word"
        End If

        If BOOLFORCEWATERMARK = 0 Then
            Me.lblWatermark.Text = "User IS NOT forced to use watermarks"
        Else
            Me.lblWatermark.Text = "User IS forced to use watermarks"
        End If

        'User is forced to create document as PDF
        If BOOLFORCEFINALREPORTPDF Then
            Me.lblForcePDF.Text = "User IS forced to create document as PDF"
        Else
            Me.lblForcePDF.Text = "User IS NOT forced to create document as PDF"
        End If

        If BOOLFORCEWATERMARK Then
            Me.chkWatermark.Checked = True
            Me.chkWatermark.Enabled = False
        End If

        If Me.chkWatermark.Checked Then
            boolIncludeWaterMark = True
        Else
            boolIncludeWaterMark = False
        End If

        If BOOLFORCEFINALREPORTPDF Then
            Me.chkDoPDF.Checked = True
            Me.chkDoPDF.Enabled = False
        End If

        If Me.chkDoPDF.Checked Then
            gDoPDF = True
        Else
            gDoPDF = False
        End If

    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        Dim str1 As String
        Dim intR As Short

        If BOOLFORCEFINALREPORTPDF Then
            str1 = "User permissions are set to force document as PDF."
            str1 = str1 & ChrW(10) & ChrW(10) & "After report generation, you will be directed to a PDF window."
            intR = MsgBox(str1, MsgBoxStyle.OkCancel, "Opening in PDF...")
            If intR = 1 Then
            Else
                GoTo end1
            End If
        End If

        If Me.chkDoPDF.Checked Then
            gDoPDF = True
        Else
            gDoPDF = False
        End If

        boolCancel = False
        boolDoFormulas = Me.chkFormulas.Checked
        If Me.chkVerbose.Checked Then
            boolVerbose = True
        Else
            boolVerbose = False
        End If

        If Me.chkDisableWarnings.Checked Then
            boolDisableWarnings = True
        Else
            boolDisableWarnings = False
        End If

        'deactivate for now
        'If Me.chkShortSampleName.Checked Then
        '    boolSampleName01 = True
        'Else
        '    boolSampleName01 = False
        'End If

        boolSampleName01 = False

        Call SetReportFormats()

        Me.boolCancel = False

        Me.Visible = False

end1:

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.boolCancel = True
        boolDoFormulas = Me.chkFormulas.Checked
        Me.Visible = False

    End Sub

    Private Sub chkFormulas_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFormulas.CheckedChanged
        If Me.chkFormulas.Checked Then
            boolDoFormulas = True
        Else
            boolDoFormulas = False
        End If
    End Sub

    Private Sub chkHyperlink_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHyperlink.CheckedChanged
        If Me.chkHyperlink.Checked Then
            boolDoHyperlinks = True
        Else
            boolDoHyperlinks = False
        End If

    End Sub


    Sub SetReportFormats()

        'legend
        'Public boolExcludeTableNumbers As Boolean = False
        'Public boolExcludeTableTitles As Boolean = False
        'Public boolExcludeEntireTableTitle As Boolean = False

        'boolExcludeTableNumbers = False
        'boolExcludeTableTitles = False
        'boolExcludeEntireTableTitle = False
        'boolIncludeWaterMark = False
        'boolExcludeCoverPage = False
        'boolDisableWarnings = False
        'boolExcludeHeaderFooter = false

        If Me.chkDisableWarnings.Checked Then
            boolDisableWarnings = True
        Else
            boolDisableWarnings = False
        End If

        If Me.chkExcludeCoverPages.Checked Then
            boolExcludeCoverPage = True
        Else
            boolExcludeCoverPage = False
        End If

        If Me.chkExcludeTableNumbers.Checked Then
            boolExcludeTableNumbers = True
        Else
            boolExcludeTableNumbers = False
        End If

        If Me.chkExcludeEntireTableTitle.Checked Then
            boolExcludeEntireTableTitle = True
        Else
            boolExcludeEntireTableTitle = False
        End If

        If Me.chkExcludeTableTitles.Checked Then
            boolExcludeTableTitles = True
        Else
            boolExcludeTableTitles = False
        End If

        If Me.chkExcludeHeaderFooter.Checked Then
            boolExcludeHeaderFooter = True
        Else
            boolExcludeHeaderFooter = False
        End If

        If Me.chkWatermark.Checked Then
            boolIncludeWaterMark = True
        Else
            boolIncludeWaterMark = False
        End If

        If Me.chkReadOnlyTables.Checked Then
            gboolReadOnlyTables = True
            frmH.chkReadOnlyTables.Checked = True
        Else
            gboolReadOnlyTables = False
            frmH.chkReadOnlyTables.Checked = False
        End If

    End Sub


    Private Sub cmdSaveSelections_Click(sender As System.Object, e As System.EventArgs) Handles cmdSaveSelections.Click

        'MsgBox("Under construction", MsgBoxStyle.Information, "Under construction...")

        Dim strModule As String

        strModule = "Report Prelim"

        Try
            Call RecordReportPrelim(Me, strModule)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub chkAdvSettings_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAdvSettings.CheckedChanged

        Call AdvSettings()

    End Sub

    Sub AdvSettings()

        '20190222 LEE:
        'Tweak size of this window
        Dim boolER As Boolean = False 'Entire Report
        Dim str1 As String = Me.lblText.Text
        If InStr(1, str1, "Entire", CompareMethod.Text) > 0 Or InStr(1, str1, "Example", CompareMethod.Text) > 0 Then
            boolER = True
        End If


        Dim a, b, c, d, e, f

        '20190219 LEE:
        e = Me.panAnalytes.Top + Me.panAnalytes.Height + 50

        If Me.chkAdvSettings.Checked Then

            Me.pan1.Top = Me.gb1.Top + Me.gb1.Height + 20
            Me.gb1.Visible = True

            a = Me.pan1.Top + Me.pan1.Height + 50

        Else

            Me.gb1.Visible = False

            If boolTable Then
                b = Me.panAnalytes.Top + Me.panAnalytes.Height + 50
                Me.pan1.Top = b
            Else
                Me.pan1.Top = Me.gb1.Top

            End If

            a = Me.pan1.Top + Me.pan1.Height + 50

        End If

        If Me.gb1.Visible Then
            Me.pan1.Top = Me.gb1.Top + Me.gb1.Height + 10
        Else
            Me.pan1.Top = Me.gb1.Top
        End If

        a = Me.pan1.Top + Me.pan1.Height + 50

        If boolER Then
            f = a
        Else
            If e > a Then
                f = e
            Else
                f = a
            End If
        End If
      

        Me.Height = f

    End Sub
    Sub frmMsgBox_ToolTipSet()
        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As New ToolTip()

        ' Set up the delays for the ToolTip.
        toolTip1.AutomaticDelay = intToolTipDelay
        toolTip1.ShowAlways = True

        'General Buttons
        toolTip1.SetToolTip(Me.cmdSaveSelections, "All studies configured in StudyDoc")
        toolTip1.SetToolTip(Me.cmdSaveSelections, "Make these the default settings when generating a report.")
        toolTip1.SetToolTip(Me.chkDisableWarnings, "Select to stop Warning/messages from appearing as pop-ups during report generation")
        toolTip1.SetToolTip(Me.chkWatermark, "Watermark appears on all pages as: ""Draft <date> <time>"" ")
        toolTip1.SetToolTip(Me.chkReadOnlyTables, "Show tables as inserted graphics (i.e. numbers cannot be changed) ")
        toolTip1.SetToolTip(Me.chkExcludeTableNumbers, "e.g. ‘Summary of Regression Constants’ instead of " _
                            & vbCrLf & "‘Table 4: Summary of Regression Constants’")
        toolTip1.SetToolTip(Me.chkExcludeTableTitles, "e.g. ‘Table 4’ instead of ‘Table 4:  Summary of Regression Constants’")
        toolTip1.SetToolTip(Me.chkExcludeEntireTableTitle, "Remove table titles (Table # + caption) completely")
        toolTip1.SetToolTip(Me.chkExcludeCoverPages, "Applies only when generating a single table/section.")
        toolTip1.SetToolTip(Me.chkExcludeHeaderFooter, "When generating a single table/section, " _
                            & vbCrLf & "do not include cover page header/footer.")
        toolTip1.SetToolTip(Me.cbxReportTemplate, "Use this to select an alternative Word template to use." _
                            & vbCrLf & "This must be compatible with the data provided.")

    End Sub


    Private Sub lbxAnalytes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxAnalytes.SelectedIndexChanged

    End Sub
End Class