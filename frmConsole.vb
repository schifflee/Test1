Option Compare Text

Imports System.Linq.Expressions
Imports System.Data
Imports System.Linq
Imports System.Data.SqlClient


Public Class frmConsole

    Public tblDashboardReport As New System.Data.DataTable


    Sub CreatetblDashboard()

        Dim str1 As String
        Dim tbl As System.Data.DataTable

        tbl = tblDashboardReport

        str1 = "ID_TBLDASHBOARDREPORT"
        If tbl.Columns.Contains(str1) Then
        Else
            Dim col As New DataColumn
            col.ColumnName = str1
            col.DataType = System.Type.GetType("System.Int64")
            col.AllowDBNull = True 'False
            tbl.Columns.Add(col)

            str1 = "ID_TBLSTUDIES" 'TBLSTUDIES.ID_TBLSTUDIES
            Dim col3 As New DataColumn
            col3.ColumnName = str1
            col3.DataType = System.Type.GetType("System.Int64")
            col3.AllowDBNull = True 'False
            tbl.Columns.Add(col3)

            str1 = "CHARTITLE" 'TBLSTUDIES.CHARWATSONSTUDYNAME
            Dim col1 As New DataColumn
            col1.ColumnName = str1
            col1.DataType = System.Type.GetType("System.String")
            col1.AllowDBNull = True 'False
            tbl.Columns.Add(col1)

            str1 = "DTSTUDYSTARTDATE" 'TBLDATA.DTSTUDYSTARTDATE
            Dim col2 As New DataColumn
            col2.ColumnName = str1
            col2.DataType = System.Type.GetType("System.DateTime")
            col2.AllowDBNull = True 'False
            tbl.Columns.Add(col2)

            str1 = "CHARREPORTNUMBER" 'TBLREPORTS.CHARREPORTNUMBER
            Dim col5 As New DataColumn
            col5.ColumnName = str1
            col5.DataType = System.Type.GetType("System.String")
            col5.AllowDBNull = True 'False
            tbl.Columns.Add(col5)

            str1 = "DTREPORTDRAFTISSUEDATE" 'TBLREPORTS.DTREPORTDRAFTISSUEDATE
            Dim col6 As New DataColumn
            col6.ColumnName = str1
            col6.DataType = System.Type.GetType("System.DateTime")
            col6.AllowDBNull = True 'False
            tbl.Columns.Add(col6)

            str1 = "DTREPORTFINALISSUEDATE" 'TBLREPORTS.DTREPORTFINALISSUEDATE
            Dim col7 As New DataColumn
            col7.ColumnName = str1
            col7.DataType = System.Type.GetType("System.DateTime")
            col7.AllowDBNull = True 'False
            tbl.Columns.Add(col7)

            str1 = "DTSTUDYENDDATE" 'TBLDATA.DTSTUDYENDDATE
            Dim col4 As New DataColumn
            col4.ColumnName = str1
            col4.DataType = System.Type.GetType("System.DateTime")
            col4.AllowDBNull = True 'False
            tbl.Columns.Add(col4)

        End If

        Dim dv As System.Data.DataView = New DataView(tbl)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        Dim dgv As DataGridView

        dgv = Me.dgvDashboard

        dgv.DataSource = dv

        Try
            dgv.Columns("ID_TBLDASHBOARDREPORT").Visible = False

            dgv.Columns("ID_TBLSTUDIES").Visible = False

            dgv.Columns("CHARTITLE").Visible = True
            dgv.Columns("CHARTITLE").HeaderText = "Study ID"
            'dgv.Columns("CHARTITLE").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill

            dgv.Columns("DTSTUDYSTARTDATE").Visible = True
            dgv.Columns("DTSTUDYSTARTDATE").HeaderText = "Study Start Date"
            dgv.Columns("DTSTUDYSTARTDATE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.Columns("CHARREPORTNUMBER").Visible = True
            dgv.Columns("CHARREPORTNUMBER").HeaderText = "Report Number"

            dgv.Columns("DTREPORTDRAFTISSUEDATE").Visible = True
            dgv.Columns("DTREPORTDRAFTISSUEDATE").HeaderText = "Report Draft Date"
            dgv.Columns("DTREPORTDRAFTISSUEDATE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.Columns("DTREPORTFINALISSUEDATE").Visible = True
            dgv.Columns("DTREPORTFINALISSUEDATE").HeaderText = "Report Final Date"
            dgv.Columns("DTREPORTFINALISSUEDATE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.Columns("DTSTUDYENDDATE").Visible = True
            dgv.Columns("DTSTUDYENDDATE").HeaderText = "Study End Date"
            dgv.Columns("DTSTUDYENDDATE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

            dgv.AutoResizeColumns()

        Catch ex As Exception

        End Try



    End Sub


    Sub FillDashboardTable()

        Dim str1 As String
        Dim tblDB As System.Data.DataTable
        Dim rowsDB() As DataRow
        Dim intRowsDB As Int64
        Dim id As Int64

        Dim tbl1 As System.Data.DataTable 'TBLSTUDIES
        Dim rows1() As DataRow
        Dim intRows1 As Int64
        Dim tbl2 As System.Data.DataTable 'TBLDATA
        Dim rows2() As DataRow
        Dim intRows2 As Int64
        Dim tbl3 As System.Data.DataTable 'TBLREPORTS
        Dim rows3() As DataRow
        Dim intRows3 As Int64

        Dim Count1 As Int64
        Dim Count2 As Int64
        Dim Count3 As Int64

        Dim int1 As Int64
        Dim int2 As Int64
        Dim int3 As Int64

        Dim strF As String
        Dim strS As String

        Dim numTotGWSC As Int16
        Dim numTotOSC As Int16

        Dim var1, var2, var3

        Call CreatetblDashboard()

        tblDB = tblDashboardReport
        Try
            tbl1 = tblStudies
            tbl2 = tblData
            tbl3 = tblReports

        Catch ex As Exception

            GoTo end1

        End Try

        tblDB.Clear()

        Try

            strF = "DTSTUDYSTARTDATE IS NOT NULL AND DTSTUDYENDDATE IS NULL"
            strF = "DTSTUDYENDDATE IS NULL"
            strF = "ID_TBLSTUDIES > -1"
            strS = "DTSTUDYSTARTDATE ASC"
            rows2 = tbl2.Select(strF, strS)
            intRows2 = rows2.Length

            If intRows2 = 0 Then
            Else

                id = 0

                Try
                    For Count1 = 0 To intRows2 - 1

                        var1 = rows2(Count1).Item("ID_TBLSTUDIES")
                        strF = "ID_TBLSTUDIES = " & var1

                        rows3 = tbl3.Select(strF)
                        int3 = rows3.Length

                        rows1 = tbl1.Select(strF)
                        If rows1.Length = 0 Then
                            var2 = "aa"
                        Else
                            var2 = rows1(0).Item("CHARWATSONSTUDYNAME")
                        End If


                        Dim nrow As DataRow = tblDB.NewRow
                        id = id + 1
                        nrow.BeginEdit()
                        nrow.Item("ID_TBLDASHBOARDREPORT") = id
                        nrow.Item("CHARTITLE") = var2

                        var2 = rows2(Count1).Item("DTSTUDYSTARTDATE")
                        nrow.Item("DTSTUDYSTARTDATE") = var2

                        var2 = rows2(Count1).Item("DTSTUDYENDDATE")
                        nrow.Item("DTSTUDYENDDATE") = var2

                        If int3 = 0 Then
                            var2 = DBNull.Value ' rows3(Count1).Item("DTREPORTDRAFTISSUEDATE")
                            nrow.Item("DTREPORTDRAFTISSUEDATE") = var2

                            var2 = DBNull.Value ' rows3(Count1).Item("DTREPORTFINALISSUEDATE")
                            nrow.Item("DTREPORTFINALISSUEDATE") = var2

                            var2 = DBNull.Value ' rows3(Count1).Item("CHARREPORTNUMBER")
                            nrow.Item("CHARREPORTNUMBER") = var2

                        Else
                            var2 = rows3(0).Item("DTREPORTDRAFTISSUEDATE")
                            nrow.Item("DTREPORTDRAFTISSUEDATE") = var2

                            var2 = rows3(0).Item("DTREPORTFINALISSUEDATE")
                            nrow.Item("DTREPORTFINALISSUEDATE") = var2

                            var2 = rows3(0).Item("CHARREPORTNUMBER")
                            nrow.Item("CHARREPORTNUMBER") = var2
                        End If

                        nrow.EndEdit()
                        tblDB.Rows.Add(nrow)

                    Next
                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1
                End Try

                'now set pb's
                Me.pbTotalGuWuStudies.Maximum = intRows2
                Me.pbTotalGuWuStudies.Value = intRows2
                Me.lblTotalGuWuStudiesCount.Text = intRows2
                numTotGWSC = intRows2


                Me.pbFinalReport.Maximum = intRows2
                Me.pbFinalReport.Value = intRows2

                strF = "DTREPORTFINALISSUEDATE IS NOT NULL"
                Erase rowsDB
                rowsDB = tblDB.Select(strF)
                intRowsDB = rowsDB.Length

                If Me.rbAbsolute.Checked Then
                    If intRowsDB = 0 Then
                        If numTotGWSC = 0 Then
                            Me.lblFinalCount.Text = "0 of 0"
                        Else
                            Me.lblFinalCount.Text = "0 of " & Format(numTotGWSC, "#,###")
                        End If
                    Else
                        If numTotGWSC = 0 Then
                            Me.lblFinalCount.Text = Format(intRowsDB, "#,###") & " of 0"
                        Else
                            Me.lblFinalCount.Text = Format(intRowsDB, "#,###") & " of " & Format(numTotGWSC, "#,###")
                        End If
                    End If

                Else
                    If intRows2 = 0 Then
                        var2 = RoundToDecimalRAFZ(0, 0) & "% (" & Format(intRowsDB, "#,###") & " of " & Format(numTotGWSC, "#,###") & ")"
                    Else
                        var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (" & Format(intRowsDB, "#,###") & " of " & Format(numTotGWSC, "#,###") & ")"
                    End If

                    Me.lblFinalCount.Text = var2
                End If
                Me.pbFinalReport.Value = intRowsDB



                strF = "DTREPORTFINALISSUEDATE IS NULL"
                Erase rowsDB
                rowsDB = tblDB.Select(strF)
                intRowsDB = rowsDB.Length
                numTotOSC = intRowsDB
                Me.lblTotalOpenStudiesCount.Text = intRowsDB
                'If Me.rbAbsolute.Checked Then
                '    Me.lblTotalOpenStudiesCount.Text = intRowsDB
                'Else
                '    var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "%"
                '    Me.lblTotalOpenStudiesCount.Text = var2
                'End If
                Me.pbTotalOpenStudies.Maximum = intRowsDB
                Me.pbTotalOpenStudies.Value = intRowsDB
                Me.lblTotalOpenStudiesCount.Text = Format(numTotOSC, "#,###") 'intRowsDB

                intRows2 = intRowsDB

                Me.pbInProgressReport.Maximum = intRows2
                Me.pbDraftReport.Maximum = intRows2
                'Me.pbFinalReport.Maximum = intRows2
                'lblTotalOpenStudiesCount

                strF = "DTREPORTDRAFTISSUEDATE IS NULL AND DTREPORTFINALISSUEDATE IS NULL"
                Erase rowsDB
                rowsDB = tblDB.Select(strF)
                intRowsDB = rowsDB.Length
                If Me.rbAbsolute.Checked Then
                    If intRowsDB = 0 Then
                        If numTotOSC = 0 Then
                            Me.lblInProgressCount.Text = "0 of 0" ' intRowsDB
                        Else
                            Me.lblInProgressCount.Text = "0 of " & Format(numTotOSC, "#,###") ' intRowsDB
                        End If
                    Else
                        If numTotOSC = 0 Then
                            Me.lblInProgressCount.Text = Format(intRowsDB, "#,###") & " of 0" ' intRowsDB
                        Else
                            Me.lblInProgressCount.Text = Format(intRowsDB, "#,###") & " of " & Format(numTotOSC, "#,###") ' intRowsDB
                        End If

                    End If
                Else
                    If intRows2 = 0 Then
                        If numTotOSC = 0 Then
                            If intRowsDB = 0 Then
                                var2 = RoundToDecimalRAFZ(0, 0) & "% (0 of 0)"
                            Else
                                var2 = RoundToDecimalRAFZ(0, 0) & "% (" & Format(intRowsDB, "#,###") & " of 0)"
                            End If

                        Else
                            If intRowsDB = 0 Then
                                var2 = RoundToDecimalRAFZ(0, 0) & "% (0 of " & Format(numTotOSC, "#,###") & ")"
                            Else
                                var2 = RoundToDecimalRAFZ(0, 0) & "% (" & Format(intRowsDB, "#,###") & " of " & Format(numTotOSC, "#,###") & ")"
                            End If

                        End If
                    Else
                        If numTotOSC = 0 Then
                            If intRowsDB = 0 Then
                                var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (0 of 0)"
                            Else
                                var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (" & Format(intRowsDB, "#,###") & " of 0)"
                            End If

                        Else
                            If intRowsDB = 0 Then
                                var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (0 of " & Format(numTotOSC, "#,###") & ")"
                            Else
                                var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (" & Format(intRowsDB, "#,###") & " of " & Format(numTotOSC, "#,###") & ")"
                            End If

                        End If

                    End If

                    Me.lblInProgressCount.Text = var2
                End If
                Me.pbInProgressReport.Value = intRowsDB

                strF = "DTREPORTDRAFTISSUEDATE IS NOT NULL"
                Erase rowsDB
                rowsDB = tblDB.Select(strF)
                intRowsDB = rowsDB.Length
                If Me.rbAbsolute.Checked Then
                    If intRowsDB = 0 Then
                        If numTotOSC = 0 Then
                            Me.lblInDraftCount.Text = "0 of 0" 'intRowsDB
                        Else
                            Me.lblInDraftCount.Text = "0 of " & Format(numTotOSC, "#,###") 'intRowsDB
                        End If
                    Else
                        If numTotOSC = 0 Then
                            Me.lblInDraftCount.Text = Format(intRowsDB, "#,###") & " of 0" 'intRowsDB
                        Else
                            Me.lblInDraftCount.Text = Format(intRowsDB, "#,###") & " of " & Format(numTotOSC, "#,###") 'intRowsDB
                        End If
                    End If

                Else
                    If intRowsDB = 0 Then
                        'for some reason, Format(intRowsDB, "#,###") for zero value returns blank space
                        If intRows2 = 0 Then
                            If numTotOSC = 0 Then
                                var2 = RoundToDecimalRAFZ(0, 0) & "% (0 of 0)"
                            Else
                                var2 = RoundToDecimalRAFZ(0, 0) & "% (0 of " & Format(numTotOSC, "#,###") & ")"
                            End If

                        Else
                            If numTotOSC = 0 Then
                                var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (0 of 0)"
                            Else
                                var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (0 of " & Format(numTotOSC, "#,###") & ")"
                            End If

                        End If

                    Else
                        If intRows2 = 0 Then
                            If numTotOSC = 0 Then
                                var2 = RoundToDecimalRAFZ(0, 0) & "% (" & Format(intRowsDB, "#,###") & " of 0)"
                            Else
                                var2 = RoundToDecimalRAFZ(0, 0) & "% (" & Format(intRowsDB, "#,###") & " of " & Format(numTotOSC, "#,###") & ")"
                            End If

                        Else
                            If numTotOSC = 0 Then
                                var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (" & Format(intRowsDB, "#,###") & " of 0)"
                            Else
                                var2 = RoundToDecimalRAFZ(intRowsDB / intRows2 * 100, 0) & "% (" & Format(intRowsDB, "#,###") & " of " & Format(numTotOSC, "#,###") & ")"
                            End If

                        End If

                    End If

                    Me.lblInDraftCount.Text = var2
                End If
                Me.pbDraftReport.Value = intRowsDB


            End If

            Call FillDashboardDGV()
        Catch ex As Exception
            var1 = ex.Message
        End Try


end1:

    End Sub

    Sub frmConsole_ToolTipSet()

        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As New ToolTip()
        Dim Count1 As Short
        Dim intE As Short
        Dim ctl As Control
        Dim strM As String

        toolTip1.AutomaticDelay = intToolTipDelay
        'toolTip1.UseFading = False
        'tooltip1.
        'toolTip1.BackColor = Color.Goldenrod
        'toolTip1.IsBalloon = True
        ' Force the ToolTip text to be displayed whether or not the form is active.
        toolTip1.ShowAlways = True

        Try

            toolTip1.SetToolTip(Me.pbTotalGuWuStudies, "All studies configured in StudyDoc")
            toolTip1.SetToolTip(Me.pbFinalReport, "Studies with a Final Report Date")
            toolTip1.SetToolTip(Me.pbTotalOpenStudies, "Studies without a Final Report Date")
            toolTip1.SetToolTip(Me.pbInProgressReport, "Studies without a Draft Report Date" & vbCrLf & "and without a Final Report Date")
            toolTip1.SetToolTip(Me.pbDraftReport, "Studies with a Draft Report Date, " & vbCrLf & "but without a Final Report Date")
            toolTip1.SetToolTip(Me.rbTotalGuWuStudies, "All studies configured in StudyDoc")
            toolTip1.SetToolTip(Me.rbFinalReports, "Studies with a Final Report Date")
            toolTip1.SetToolTip(Me.rbTotalOpenStudies, "Studies without a Final Report Date")
            toolTip1.SetToolTip(Me.rbInProgressStudies, "Studies without a Draft Report Date" & vbCrLf & "and without a Final Report Date")
            toolTip1.SetToolTip(Me.rbDraftReports, "Studies with a Draft Report Date, " & vbCrLf & "but without a Final Report Date")
            toolTip1.SetToolTip(Me.lblpbTotalGuWuStudies, "All studies configured in StudyDoc")
            toolTip1.SetToolTip(Me.lblpbFinalReport, "Studies with a Final Report Date")
            toolTip1.SetToolTip(Me.lblpbTotalOpenStudies, "Studies without a Final Report Date")
            toolTip1.SetToolTip(Me.lblpbInProgressReport, "Studies without a Draft Report Date" & vbCrLf & "and without a Final Report Date")
            toolTip1.SetToolTip(Me.lblpbDraftReport, "Studies with a Draft Report Date, " & vbCrLf & "but without a Final Report Date")

            'Buttons
            toolTip1.SetToolTip(Me.cmdLogin, "Log in or log out (without exiting)")
            toolTip1.SetToolTip(Me.cmdChangePassword, "Change current user’s password")
            toolTip1.SetToolTip(Me.lblReportWriter, "Enter Report Writer main screen")
            toolTip1.SetToolTip(Me.cmdReportWriter, "Enter Report Writer main screen")
            toolTip1.SetToolTip(Me.lblConfig, "Set Users, Permissions, & E-signature settings")
            toolTip1.SetToolTip(Me.cmdConfig, "Set Users, Permissions, & E-signature settings")
            toolTip1.SetToolTip(Me.lblAuditTrail, "Shortcut to Audit Trail")
            toolTip1.SetToolTip(Me.cmdAuditTrail, "Shortcut to Audit Trail")
            toolTip1.SetToolTip(Me.lblExit, "Exit StudyDoc application")
            toolTip1.SetToolTip(Me.cmdExit, "Exit StudyDoc application")
            toolTip1.SetToolTip(Me.rbASCReportDB, "Sort in ascending order (e.g. A-Z) ")
            toolTip1.SetToolTip(Me.rbDESCReportDB, "Sort in descending order (e.g. Z-A) ")

            'Dashboard Grid
            Me.dgvDashboard.Columns("CHARTITLE").ToolTipText = "Watson Study ID"
            Me.dgvDashboard.Columns("DTSTUDYSTARTDATE").ToolTipText = "Study Start Date: Date entered in the " _
                                & vbCrLf & """Add/Edit Top Level Data"" page"
            Me.dgvDashboard.Columns("CHARREPORTNUMBER").ToolTipText = "Report Number entered in the study's " _
                                & vbCrLf & """Choose Study & Report"" tab"
            Me.dgvDashboard.Columns("DTREPORTDRAFTISSUEDATE").ToolTipText = "Report Draft Date:  Draft date entered " _
                                & vbCrLf & "for the study in the ""Choose Study & Report"" tab"
            Me.dgvDashboard.Columns("DTREPORTFINALISSUEDATE").ToolTipText = "Report Final Date:  Issue date entered " _
                                & vbCrLf & "for the study in the ""Choose Study & Report"" tab"
            Me.dgvDashboard.Columns("DTSTUDYENDDATE").ToolTipText = "Study End Date: Date entered " _
                                & vbCrLf & "in the ""Add/Edit Top Level Data"" page"
        Catch ex As Exception

        End Try

    End Sub

    Sub FillDashboardDGV()

        Dim dgv As DataGridView
        Dim strF As String
        Dim strS As String
        Dim strOrder As String
        Dim dv As System.Data.DataView

        If Me.rbASCReportDB.Checked Then
            strOrder = "ASC"
        Else
            strOrder = "DESC"
        End If

        Dim bool1 As Boolean = False
        Dim bool2 As Boolean = False
        Dim bool3 As Boolean = False
        Dim bool4 As Boolean = False
        Dim bool5 As Boolean = False

        dgv = Me.dgvDashboard

        If Me.rbTotalGuWuStudies.Checked Then
            strF = "ID_TBLDASHBOARDREPORT > 0"
            bool1 = True
        ElseIf Me.rbTotalOpenStudies.Checked Then
            strF = "DTREPORTFINALISSUEDATE IS NULL"
            bool2 = True
        ElseIf Me.rbInProgressStudies.Checked Then
            strF = "DTREPORTDRAFTISSUEDATE IS NULL AND DTREPORTFINALISSUEDATE IS NULL"
            bool3 = True
        ElseIf Me.rbDraftReports.Checked Then
            strF = "DTREPORTDRAFTISSUEDATE IS NOT NULL"
            bool4 = True
        ElseIf Me.rbFinalReports.Checked Then
            strF = "DTREPORTFINALISSUEDATE IS NOT NULL"
            bool5 = True
        End If

        If bool1 Then
            Me.lblpbTotalGuWuStudies.Font = New Font(Me.lblpbTotalGuWuStudies.Font, FontStyle.Bold)
        Else
            Me.lblpbTotalGuWuStudies.Font = New Font(Me.lblpbTotalGuWuStudies.Font, FontStyle.Regular)
        End If

        If bool2 Then
            Me.lblpbTotalOpenStudies.Font = New Font(Me.lblpbTotalOpenStudies.Font, FontStyle.Bold)
        Else
            Me.lblpbTotalOpenStudies.Font = New Font(Me.lblpbTotalOpenStudies.Font, FontStyle.Regular)
        End If

        If bool3 Then
            Me.lblpbInProgressReport.Font = New Font(Me.lblpbInProgressReport.Font, FontStyle.Bold)
        Else
            Me.lblpbInProgressReport.Font = New Font(Me.lblpbInProgressReport.Font, FontStyle.Regular)
        End If

        If bool4 Then
            Me.lblpbDraftReport.Font = New Font(Me.lblpbDraftReport.Font, FontStyle.Bold)
        Else
            Me.lblpbDraftReport.Font = New Font(Me.lblpbDraftReport.Font, FontStyle.Regular)
        End If

        If bool5 Then
            Me.lblpbFinalReport.Font = New Font(Me.lblpbFinalReport.Font, FontStyle.Bold)
        Else
            Me.lblpbFinalReport.Font = New Font(Me.lblpbFinalReport.Font, FontStyle.Regular)
        End If

        Select Case Me.cbxSortDashboard.Text
            Case "Study ID"
                strS = "CHARTITLE " & strOrder
            Case "Study Start Date"
                strS = "DTSTUDYSTARTDATE " & strOrder & ", CHARTITLE ASC"
            Case "Study End Date"
                strS = "DTSTUDYENDDATE " & strOrder & ", CHARTITLE ASC"
            Case "Report Number"
                strS = "CHARREPORTNUMBER " & strOrder & ", CHARTITLE ASC"
            Case "Report Draft Date"
                strS = "DTREPORTDRAFTISSUEDATE " & strOrder & ", CHARTITLE ASC"
            Case "Report Final Date"
                strS = "DTREPORTFINALISSUEDATE " & strOrder & ", CHARTITLE ASC"
        End Select

        dv = dgv.DataSource

        If dv Is Nothing Then
        Else
            dv.RowFilter = strF
            dv.Sort = strS
        End If

        dgv.RowHeadersWidth = 25

        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()



    End Sub

    Private Sub frmConsole_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Try
            Call FillDashboardTable()

        Catch ex As Exception

            Dim var1
            var1 = ex.Message
            var1 = var1
        End Try

    End Sub

    Sub FillcbxDashboard()

        Me.cbxSortDashboard.Items.Add("Study ID")
        Me.cbxSortDashboard.Items.Add("Study Start Date")
        Me.cbxSortDashboard.Items.Add("Study End Date")
        Me.cbxSortDashboard.Items.Add("Report Number")
        Me.cbxSortDashboard.Items.Add("Report Draft Date")
        Me.cbxSortDashboard.Items.Add("Report Final Date")

        Me.cbxSortDashboard.SelectedIndex = 0

    End Sub

    Private Sub frmConsole_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'no need to perform if user is already logged off
        If Len(gUserID) = 0 Then
        Else
            Dim arr1
            Dim var1
            Dim con As New ADODB.Connection
            Dim constr As String

            If boolGuWuAccess Then
                constr = constrIni
            ElseIf boolGuWuSQLServer Then
                constr = constrIni ' "Provider=SQLOLEDB;" & constrIni
            ElseIf boolGuWuOracle Then
                constr = constrIniGuWuODBC
            End If

            Try
                con.Open(constr)
            Catch ex As Exception
                var1 = ex.Message
            End Try

            Call SaveLoginAttempt(arr1, con, 1, False, True)

            Try
                con.Close()
                con = Nothing
            Catch ex As Exception

            End Try
        End If


    End Sub

    Sub PlaceControls()

        Dim a, b, c, d

        a = Me.cmdReportWriter.Top + Me.cmdReportWriter.Height

        b = Me.cmdConfig.Top

        c = b - a

        Me.cmdExit.Top = Me.cmdAuditTrail.Top + Me.cmdAuditTrail.Height + c

    End Sub

    Private Sub frmConsole_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        'Call TempAssSamples() 'temp for doing auto assign samples

        Call DoubleBufferControl(Me, "dgv")
        Call DoubleBufferControl(Me, "pb")

        Call ControlDefaults(Me)

        Me.lbl2.Text = GetStudyDocHeader(True)

        Dim str1 As String
        Dim strF As String
        Dim rowP() As DataRow
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim str5 As String
        Dim tblP As System.Data.DataTable
        Dim tblU As System.Data.DataTable
        Dim rowU() As DataRow

        Call SetFormPos(Me)

        str1 = "Study" ' "Gubbs" & ChrW(174) & " Inc"
        Me.lbl1.Text = str1
        str1 = "Doc" ' & ChrW(8482) ' "GuWu"
        Me.lbl1a.Text = str1
        str1 = ChrW(8482) 'ChrW(174)
        Me.lbl1b.Text = str1


        gWorkstation = My.Computer.Name

        frmC = Me

        Call FillcbxDashboard()

        Call CreateAuditTrailTemp()

        Dim frm As New frmHome_01

        Call frm.FormLoad()

        Text = GetCaption("Console")
        MeCaption = Text
        Me.Text = MeCaption

        frm.Visible = False

        If boolInitLogIn Then
            Me.cmdLogin.Text = "&Log Off"
        Else
            Me.cmdLogin.Text = "&Log On"
        End If

        '20160504 LEE: Don't open StudyDesign anymore

        'Dim frm1 As New frmSDHome
        'Call frm1.FormLoad()
        'frm1.Visible = False

        'add some blank rows to a few tables
        Call AddBlankRows()

        'do dashboard stuff
        Try
            Call CreatetblDashboard()
            Call FillDashboardTable()
        Catch ex As Exception

        End Try

        Call CorrectOrder01()

        'don't need this here. gets called in frmHome.load
        'Call DatabaseCorrections() 'updates database stuff

        Call PlaceControls()

        'set audittrail and esig
        Call SetAuditESig()

        Call frmConsole_ToolTipSet()

        Me.cmdExit.Focus()

        Call DisplayLogin()

        'SendKeys.Send("%")
        ''20171113 LEE:
        ''for some reason, panDashboard isn't anchoring bottom properly
        'Dim a, b, c, d

        'a = Me.panDashboard.Top
        'b = Me.Height
        'c = b - a - 100
        'Me.panDashboard.Height = c


    End Sub

    Sub DisplayLogin()

        Me.cmdChangePassword.Visible = Not (gboolLDAP)

    End Sub


    Sub AddBlankRows()

        Dim dtbl As System.Data.DataTable

        dtbl = tblConfigReportType
        Dim row As DataRow = dtbl.NewRow
        row.BeginEdit()
        row("ID_TBLCONFIGREPORTTYPE") = 0
        row("CHARREPORTTYPE") = ""
        row("BOOLINCLUDE") = -1
        row.EndEdit()
        dtbl.Rows.Add(row)

        dtbl = tblGuWuStudyStat
        Dim row1 As DataRow = dtbl.NewRow
        row1.BeginEdit()
        row1("ID_TBLGUWUSTUDYSTAT") = 0
        row1("CHARSTATUS") = ""
        row1.EndEdit()
        dtbl.Rows.Add(row1)

        dtbl = tblGuWuStudyDesignType
        Dim row2 As DataRow = dtbl.NewRow
        row2.BeginEdit()
        row2("ID_TBLGUWUSTUDYDESIGNTYPE") = 0
        row2("CHARSTUDYDESIGNTYPE") = ""
        row2.EndEdit()
        dtbl.Rows.Add(row2)


        'ID_TBLGUWUSTUDYSTAT

    End Sub

    Function AllowGuest() As Boolean

        '20160711 LEE: Deprecated

        AllowGuest = False
        Dim str1 As String
        If InStr(1, Me.cmdLogin.Text, "On", CompareMethod.Text) > 0 Then
            str1 = "Guest not allowed to proceed."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            AllowGuest = False
        Else
            AllowGuest = True
        End If

    End Function

    Private Sub cmdStudyDesigner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStudyDesigner.Click

        Dim str1 As String

        'If AllowGuest() Then
        'Else
        '    Exit Sub
        'End If


        str1 = "Under Construction..."
        MsgBox(str1, MsgBoxStyle.Information, "Under Construction...")
        Exit Sub

        frmSD.Visible = True

        If id_tblPersonnel = 0 Or id_tblUserAccounts = 0 Then
            'frmH.cbxExampleReport.Enabled = False
        Else
            'frmH.cbxExampleReport.Enabled = True
        End If

        frmSD.WindowState = FormWindowState.Maximized
        Me.Visible = False

    End Sub

    Private Sub cmdReportWriter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportWriter.Click

        'If AllowGuest() Then
        'Else
        '    Exit Sub
        'End If

        Dim strM As String

        If BOOLCONSOLERW Then
        Else
            strM = "User does not have permission to enter the Report Writer window."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Dim str1 As String
        Dim var1

        Call PositionProgress()

        frmH.lblProgress.Text = "Opening Report Writer module..."
        frmH.lblProgress.Refresh()
        frmH.lblProgress.Visible = True
        frmH.lblProgress.Refresh()

        'check
        var1 = frmH.dgvwStudy.CurrentCell 'debug
        var1 = var1

        'If id_tblPersonnel = 0 Or id_tblUserAccounts = 0 Then
        '    frmH.cbxExampleReport.Enabled = False
        '    Call LockAll(True)
        'Else
        '    frmH.cbxExampleReport.Enabled = True
        'End If

        Call LockAll(True, True)
        'make cmdEdit enablde
        Call LockHomeTab(True)

        str1 = GetCaption("ReportWriter")
        frmH.Text = str1 ' GetCaption("ReportWriter")

        var1 = frmH.dgvwStudy.CurrentCell 'debug
        var1 = var1
        Try
            frmH.WindowState = FormWindowState.Maximized
            Call frmH.FormLoad2()
            'pesky
            Call SetPanAction()
            frmH.dgvwStudy.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            frmH.dgvwStudy.ClearSelection()
            boolFormLoad = True
            frmH.Visible = True
            frmH.dgvwStudy.ClearSelection()
            boolFormLoad = False
        Catch ex As Exception
            Dim frm As New frmHome_01
            frm.Show()
            frmH.dgvwStudy.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        End Try

        frmH.WindowState = FormWindowState.Maximized

        'pesky
        Call SetPanAction()

        Pause(0.2)

        Me.Visible = False

        frmH.lblProgress.Visible = False

        frmH.lblProgress.Refresh()



    End Sub

    Private Sub cmdConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConfig.Click

        'If AllowGuest() Then
        'Else
        '    Exit Sub
        'End If



        Dim strM As String


        'If id_tblPersonnel = 0 Or id_tblUserAccounts = 0 Then
        '    strM = "Guest is not allowed access to Administration"
        '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        '    Exit Sub
        'End If

        'check to ensure user is allowed
        If BOOLADMINISTRATIONADMIN Then
        Else
            strM = "User does not have permission to enter the StudyDoc Administration window."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Dim frm As New frmAdministration

        frm.frmName = Me.Name

        frm.lblGlobalParameters.Text = "StudyDoc Global Parameters"

        frm.cmdExit.Text = "E&xit"

        'Call frm.FormLoad()

        frm.ShowDialog()
        'Try
        '    frm.ShowDialog()
        'Catch ex As Exception
        '    strM = ex.Message
        '    MsgBox(strM)
        'End Try

        frm.Dispose()

    End Sub

    Private Sub cmdLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLogin.Click

        Call Login()

        Call DisplayLogin()

    End Sub

    Sub Login()

        Dim strT As String
        strT = cmdLogin.Text
        Dim str2 As String
        Dim boolDo As Boolean
        Dim var1
        Dim strM As String

        boolDo = True

        If StrComp(strT, "&Log On", CompareMethod.Text) = 0 Then 'log in

            Dim frm As New frmLogon
            If boolFormLoad Then
                frm.StartPosition = FormStartPosition.CenterScreen
            End If
            frm.ShowDialog()
            frm.Visible = False
            Me.Refresh()

            'Me.cmdLogin.Select()
            'SendKeys.Send("%(A)")
            Cursor.Current = Cursors.Default


            If frm.boolCancel Then
                boolDo = False
            Else
                cmdLogin.Text = "&Log Off"
                Refresh()
                'SendKeys.Send("%")

                Cursor.Current = Cursors.WaitCursor

                'record constants
                id_tblPersonnel = frm.idP
                id_tblUserAccounts = idU
                id_tblPermissions = frm.idPerm

                Dim tblP As System.Data.DataTable
                Dim tblU As System.Data.DataTable
                Dim rowP() As DataRow
                Dim rowU() As DataRow
                Dim str1 As String
                Dim str3 As String
                Dim str4 As String
                Dim str5 As String
                Dim strF As String
                Dim tblPerm As System.Data.DataTable
                Dim rowPerm() As DataRow

                tblP = tblPersonnel
                tblU = tblUserAccounts
                tblPerm = tblPermissions

                'find user account
                strF = "id_tblUserAccounts = " & id_tblUserAccounts
                rowU = tblU.Select(strF)
                str1 = rowU(0).Item("charUserID")

                Dim idP As Int64
                Dim strF1 As String

                strF1 = "ID_TBLPERMISSIONS = " & id_tblPermissions

                rowPerm = tblPerm.Select(strF1)

                'find user
                strF = "id_tblPersonnel = " & id_tblPersonnel
                rowP = tblP.Select(strF)
                str2 = rowP(0).Item("charFirstName")
                str3 = NZ(rowP(0).Item("charMiddleName"), "")
                str4 = rowP(0).Item("charLastName")
                If Len(str3) = 0 Then
                    str5 = str2 & " " & str4
                Else
                    str5 = str2 & " " & str3 & " " & str4
                End If

                str2 = GetStudyDocHeader(False)
                If gboolLDAP Then
                    str3 = rowU(0).Item("CHARNETWORKACCOUNT")
                    gUserLabel = " - Network User: " & str5 & " logged in as " & str3
                Else
                    gUserLabel = " - StudyDoc User: " & str5 & " logged in as " & str1
                End If

                str2 = str2 & " v" & GetVersion() & gUserLabel
                Text = str2

                gUserName = str5
                gUserID = str1

                'Text = GetCaption("ReportWriter")
                Text = GetCaption("Console")

                MeCaption = Text

                '20160711 LEE: set permissions
                Call SetPermissions(True)

                Dim boolB As Boolean
                If BOOLALLOWPDFREPORT = False And BOOLALLOWREPORTGENERATION = False Then
                    boolB = False
                Else
                    boolB = True
                End If

                Call LockReportGeneration(Not (boolB))

            End If

            'select something else
            'dgvwStudy.Select()

            'Me.cbxStudy.Select()
            Me.Refresh()
            'Me.cbxStudy.Focus()
            'Dim x, y
            'x = Me.cbxStudy.Location.X
            'y = Me.cbxStudy.Location.Y + Me.cbxStudy.Top + Me.cbxStudy.Height
            'Cursor.Position = new system.drawing.point(x + 50, y)

            'SendKeys.Send("%")

            frm.Dispose()

        Else 'log out

            boolDo = False

            Dim intInput As String
            strM = "Are you sure you wish to log off?"
            strM = strM & ChrW(10) & ChrW(10) & "StudyDoc will close."
            intInput = MsgBox(strM, MsgBoxStyle.YesNo, "Log off check...")
            If intInput = 6 Then 'continue

                gUserID = ""
                gUserName = ""

                Cursor.Current = Cursors.WaitCursor

                Call DoThis("Logoff")

                Dim constr As String
                If boolGuWuAccess Then
                    constr = constrIni
                ElseIf boolGuWuSQLServer Then
                    constr = constrIni ' "Provider=SQLOLEDB;" & constrIni
                ElseIf boolGuWuOracle Then
                    constr = constrIniGuWuODBC
                End If

                Dim arr1
                Dim con As New ADODB.Connection
                Try
                    con.Open(constr)
                Catch ex As Exception
                    var1 = ex.Message
                End Try

                Call SaveLoginAttempt(arr1, con, 1, True, False)

                Try
                    con.Close()
                    con = Nothing
                Catch ex As Exception

                End Try

                id_tblPersonnel = 0
                id_tblUserAccounts = -1
                idU = -1


                Call SetPermissions(False)

                str2 = GetStudyDocHeader(False)

                'gUserLabel = " - User: Guest with Read Only permissions"
                gUserLabel = " - User: No User Logged In"
                str2 = str2 & " v" & GetVersion() & gUserLabel
                Text = str2

                cmdLogin.Text = "&Log On"
                'select cmdLogin
                cmdLogin.Select()

                End

            Else
            End If


        End If

        'pesky
        'call FillDataTabData(ByVal boolFromReset As Boolean)
        If boolDo Then
            Call FillDataTabData(True)
            Call AssessSampleAssignment()
            Call ReportStatementsFillCharSection() 'pesky
        End If

        'Me.cbxStudy.Focus()
        'SendKeys.Send("%")

        Cursor.Current = Cursors.Default

        MeCaption = Text

    End Sub

    Private Sub cmdChangePassword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdChangePassword.Click

        Dim strM As String

        If id_tblUserAccounts < 1 Then
            strM = "'No User' has no password."
            MsgBox(strM, MsgBoxStyle.Information, "'No User' has no password...")
            Exit Sub
        End If

        If gboolLDAP Then
            strM = "This user is logged in using Windows Authentication."
            strM = strM & ChrW(10) & ChrW(10)
            strM = strM & "Password cannot be changed here."
            MsgBox(strM, MsgBoxStyle.Information, "Password cannot be changed...")
            Exit Sub
        End If

        Dim frm As New frmPasswordChange
        'frm.boolFromAdmin = False
        'frm.boolFromChgPswd = True

        frm.chkFromAdmin.Checked = False
        frm.chkFromChgPswd.Checked = True


        frm.txtID.Text = id_tblUserAccounts

        'first ensure user is allowed to change password
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim bool As Short
        Dim boolGo As Boolean
        Dim var1
        Dim dtNow As Date
        Dim dtRes As Date
        Dim dt As Date
        Dim str1 As String
        Dim ts As TimeSpan

        tbl = tblUserAccounts
        strF = "id_tblUserAccounts = " & id_tblUserAccounts
        rows = tbl.Select(strF)
        bool = rows(0).Item("boolUserCannotChangePassword")
        boolGo = True
        If bool = -1 Then 'user is not allowed to change password
            MsgBox("User '" & GetUserID() & "' doesn't have permission to change password.", MsgBoxStyle.Information, "No permissions...")
            boolGo = False
        Else

        End If

        If boolGo Then 'check for Password change restriction policy
            Dim tbl1 As System.Data.DataTable
            Dim rows1() As DataRow
            tbl1 = tblUserAccounts
            strF = "id_tblUserAccounts = " & id_tblUserAccounts
            rows1 = tbl1.Select(strF)

            'if user is forced to change password at next login, then policy is ignored
            bool = rows1(0).Item("boolChangePasswordAtNextLogon")
            If bool = -1 Then 'ignore policy
            Else
                dt = rows1(0).Item("dtTimeStamp")
                tbl = tblConfiguration
                Erase rows
                strF = "charConfigCategory = 'Password Settings' AND charConfigTitle = 'Password change restriction (minutes)'"
                rows = tbl.Select(strF)
                var1 = rows(0).Item("charConfigValue") 'restriction minutes
                dtRes = DateAdd(DateInterval.Minute, CDbl(var1), dt)
                dtNow = Now
                ts = dtNow - dtRes

                If ts.Minutes < 0 Then
                    str1 = "Sorry. The StudyDoc Password Restriction Policy will not allow you to change your password for another "
                    str1 = str1 & CStr(-ts.Minutes) & " minutes."
                    MsgBox(str1, MsgBoxStyle.Information, "Password restriction policy violation...")
                    boolGo = False
                End If
            End If
        End If

        If boolGo Then
            frm.ShowDialog()
            frm.Dispose()
            Refresh()
            'SendKeys.Send("%")

        Else
            frm.Dispose()
        End If
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        Call DoExit()

    End Sub

    Sub DoExit()

        'must do this to fire SaveLogin code
        Me.Close()

        Dim var1
        'var1 = MsgBox("Do you wish to exit?", MsgBoxStyle.YesNo, "Exit StudyDoc Study Designer...")
        'Me.Refresh()

        var1 = 6

        If var1 = 6 Then 'Yes then

            End

        ElseIf var1 = 7 Then

        End If

    End Sub

    Private Sub mnuAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuAbout.Click

        Dim frm As New frmAbout

        frm.ShowDialog()

        frm.Dispose()

    End Sub

    Private Sub cmdResults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResults.Click

        'If AllowGuest() Then
        'Else
        '    Exit Sub
        'End If

        Dim strM As String

        strM = "Under Construction."
        MsgBox(strM, MsgBoxStyle.Information, "Under Construction...")

    End Sub


    Public Function GetDirectoryViaBrowseDlg(ByVal bstrDlgTitle As String, ByVal bstrInitialDir As String) As String

    End Function

    Private Sub rbPercent_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbPercent.CheckedChanged

        Call FillDashboardTable()

    End Sub

    Private Sub cbxSortDashboard_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSortDashboard.SelectedIndexChanged

        Call FillDashboardDGV()

    End Sub

    Private Sub rbASCReportDB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbASCReportDB.CheckedChanged

        Call FillDashboardDGV()

    End Sub

    Private Sub rbTotalOpenStudies_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTotalOpenStudies.CheckedChanged

        Call FillDashboardDGV()

    End Sub

    Private Sub rbInProgressStudies_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbInProgressStudies.CheckedChanged

        Call FillDashboardDGV()

    End Sub

    Private Sub rbDraftReports_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDraftReports.CheckedChanged

        Call FillDashboardDGV()

    End Sub

    Private Sub rbFinalReports_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFinalReports.CheckedChanged

        Call FillDashboardDGV()

    End Sub

    Private Sub cmdAuditTrail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAuditTrail.Click

        Dim strM As String
        strM = "Under construction..."



        'If AllowGuest() Then
        'Else
        '    Exit Sub
        'End If

        If BOOLCONSOLEAUDITTRAIL Then
        Else
            strM = "User does not have permission to enter the StudyDoc Audit Trail window."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Dim frm As New frmAuditTrail

        frm.strForm = Me.Name

        frm.ShowDialog()

        Try
            frm.Dispose()
        Catch ex As Exception

        End Try


    End Sub

    Private Sub cmdAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAbout.Click

        Dim frm As New frmAbout

        frm.ShowDialog()

        Try
            frm.Dispose()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub rbTotalGuWuStudies_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbTotalGuWuStudies.CheckedChanged

        Call FillDashboardDGV()

    End Sub

    Private Sub lblReportWriter_Click(sender As System.Object, e As System.EventArgs) Handles lblReportWriter.Click

        Me.cmdReportWriter_Click(sender, e)

    End Sub

    Private Sub lblReportWriter_MouseEnter(sender As Object, e As System.EventArgs) Handles lblReportWriter.MouseEnter

        Dim f = New Font(Me.lblReportWriter.Font.Name, Me.lblReportWriter.Font.Size, FontStyle.Bold Or FontStyle.Underline)
        Me.lblReportWriter.Font = f
        Me.lblReportWriter.Cursor = Cursors.Hand

    End Sub

    Private Sub lblReportWriter_MouseLeave(sender As Object, e As System.EventArgs) Handles lblReportWriter.MouseLeave

        Dim f = New Font(Me.lblReportWriter.Font.Name, Me.lblReportWriter.Font.Size, FontStyle.Bold)
        Me.lblReportWriter.Font = f
        Me.lblReportWriter.Cursor = Cursors.Default

    End Sub

    Private Sub lblConfig_Click_1(sender As System.Object, e As System.EventArgs) Handles lblConfig.Click

        Me.cmdConfig_Click(sender, e)

    End Sub

    Private Sub lblConfig_MouseEnter(sender As Object, e As System.EventArgs) Handles lblConfig.MouseEnter

        Dim f = New Font(Me.lblConfig.Font.Name, Me.lblConfig.Font.Size, FontStyle.Bold Or FontStyle.Underline)
        Me.lblConfig.Font = f
        Me.lblConfig.Cursor = Cursors.Hand

    End Sub

    Private Sub lblConfig_MouseLeave(sender As Object, e As System.EventArgs) Handles lblConfig.MouseLeave

        Dim f = New Font(Me.lblConfig.Font.Name, Me.lblConfig.Font.Size, FontStyle.Bold)
        Me.lblConfig.Font = f
        Me.lblConfig.Cursor = Cursors.Default

    End Sub

    Private Sub lblAuditTrail_Click(sender As System.Object, e As System.EventArgs) Handles lblAuditTrail.Click

        Me.cmdAuditTrail_Click(sender, e)

    End Sub

    Private Sub lblAuditTrail_MouseEnter(sender As Object, e As System.EventArgs) Handles lblAuditTrail.MouseEnter

        Dim f = New Font(Me.lblAuditTrail.Font.Name, Me.lblAuditTrail.Font.Size, FontStyle.Bold Or FontStyle.Underline)
        Me.lblAuditTrail.Font = f
        Me.lblAuditTrail.Cursor = Cursors.Hand

    End Sub

    Private Sub lblAuditTrail_MouseLeave(sender As Object, e As System.EventArgs) Handles lblAuditTrail.MouseLeave

        Dim f = New Font(Me.lblAuditTrail.Font.Name, Me.lblAuditTrail.Font.Size, FontStyle.Bold)
        Me.lblAuditTrail.Font = f
        Me.lblAuditTrail.Cursor = Cursors.Default

    End Sub


    Private Sub lblExit_Click(sender As System.Object, e As System.EventArgs) Handles lblExit.Click

        Me.cmdExit_Click(sender, e)

    End Sub

    Private Sub lblExit_MouseEnter(sender As Object, e As System.EventArgs) Handles lblExit.MouseEnter

        Dim f = New Font(Me.lblExit.Font.Name, Me.lblExit.Font.Size, FontStyle.Bold Or FontStyle.Underline)
        Me.lblExit.Font = f
        Me.lblExit.Cursor = Cursors.Hand

    End Sub

    Private Sub lblExit_MouseLeave(sender As Object, e As System.EventArgs) Handles lblExit.MouseLeave

        Try
            Dim f = New Font(Me.lblExit.Font.Name, Me.lblExit.Font.Size, FontStyle.Bold)
            Me.lblExit.Font = f
            Me.lblExit.Cursor = Cursors.Default
        Catch ex As Exception

        End Try


    End Sub

    Private Sub cmdReportWriter_MouseEnter(sender As Object, e As System.EventArgs) Handles cmdReportWriter.MouseEnter

        Me.cmdReportWriter.Cursor = Cursors.Hand

    End Sub

    Private Sub cmdReportWriter_MouseLeave(sender As Object, e As System.EventArgs) Handles cmdReportWriter.MouseLeave

        Me.cmdReportWriter.Cursor = Cursors.Default

    End Sub

    Private Sub cmdConfig_MouseEnter(sender As Object, e As System.EventArgs) Handles cmdConfig.MouseEnter

        Me.cmdConfig.Cursor = Cursors.Hand

    End Sub

    Private Sub cmdConfig_MouseLeave(sender As Object, e As System.EventArgs) Handles cmdConfig.MouseLeave

        Me.cmdConfig.Cursor = Cursors.Default

    End Sub

    Private Sub cmdAuditTrail_MouseEnter(sender As Object, e As System.EventArgs) Handles cmdAuditTrail.MouseEnter

        Me.cmdAuditTrail.Cursor = Cursors.Hand

    End Sub

    Private Sub cmdAuditTrail_MouseLeave(sender As Object, e As System.EventArgs) Handles cmdAuditTrail.MouseLeave

        Me.cmdAuditTrail.Cursor = Cursors.Default

    End Sub

    Private Sub cmdExit_MouseEnter(sender As Object, e As System.EventArgs) Handles cmdExit.MouseEnter

        Me.cmdExit.Cursor = Cursors.Hand

    End Sub

    Private Sub cmdExit_MouseLeave(sender As Object, e As System.EventArgs) Handles cmdExit.MouseLeave

        Me.cmdExit.Cursor = Cursors.Default

    End Sub

    Private Sub lblpbTotalGuWuStudies_click(sender As System.Object, e As System.EventArgs) Handles lblpbTotalGuWuStudies.Click
        Me.rbTotalGuWuStudies.Select()
    End Sub
    Private Sub lblpbFinalReport_click(sender As System.Object, e As System.EventArgs) Handles lblpbFinalReport.Click
        Me.rbFinalReports.Select()
    End Sub
    Private Sub lblpbTotalOpenStudies_click(sender As System.Object, e As System.EventArgs) Handles lblpbTotalOpenStudies.Click
        Me.rbTotalOpenStudies.Select()
    End Sub
    Private Sub lblpbInProgressReport_click(sender As System.Object, e As System.EventArgs) Handles lblpbInProgressReport.Click
        Me.rbInProgressStudies.Select()
    End Sub
    Private Sub lblpbDraftReportt_click(sender As System.Object, e As System.EventArgs) Handles lblpbDraftReport.Click
        Me.rbDraftReports.Select()
    End Sub


    Private Sub lblpbTotalGuWuStudies_MouseEnter(sender As Object, e As System.EventArgs) Handles lblpbTotalGuWuStudies.MouseEnter
        With Me.lblpbTotalGuWuStudies
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold Or FontStyle.Underline)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Underline)
            End If
            .Font = f
            .Cursor = Cursors.Hand
        End With
    End Sub

    Private Sub lblpbFinalReport_MouseEnter(sender As Object, e As System.EventArgs) Handles lblpbFinalReport.MouseEnter
        With Me.lblpbFinalReport
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold Or FontStyle.Underline)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Underline)
            End If
            .Font = f
            .Cursor = Cursors.Hand
        End With
    End Sub

    Private Sub lblpbTotalOpenStudies_MouseEnter(sender As Object, e As System.EventArgs) Handles lblpbTotalOpenStudies.MouseEnter
        With Me.lblpbTotalOpenStudies
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold Or FontStyle.Underline)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Underline)
            End If
            .Font = f
            .Cursor = Cursors.Hand
        End With
    End Sub
    Private Sub lblpbInProgressReport_MouseEnter(sender As Object, e As System.EventArgs) Handles lblpbInProgressReport.MouseEnter
        With Me.lblpbInProgressReport
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold Or FontStyle.Underline)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Underline)
            End If
            .Font = f
            .Cursor = Cursors.Hand
        End With
    End Sub
    Private Sub lblpbDraftReport_MouseEnter(sender As Object, e As System.EventArgs) Handles lblpbDraftReport.MouseEnter
        With Me.lblpbDraftReport
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold Or FontStyle.Underline)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Underline)
            End If
            .Font = f
            .Cursor = Cursors.Hand
        End With
    End Sub

    Private Sub lblpbTotalGuWuStudiest_MouseLeave(sender As Object, e As System.EventArgs) Handles lblpbTotalGuWuStudies.MouseLeave
        With Me.lblpbTotalGuWuStudies
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Regular)
            End If
            .Font = f
            .Cursor = Cursors.Default
        End With
    End Sub
    Private Sub lblpbFinalReport_MouseLeave(sender As Object, e As System.EventArgs) Handles lblpbFinalReport.MouseLeave
        With Me.lblpbFinalReport
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Regular)
            End If
            .Font = f
            .Cursor = Cursors.Default
        End With
    End Sub
    Private Sub lblpbTotalOpenStudies_MouseLeave(sender As Object, e As System.EventArgs) Handles lblpbTotalOpenStudies.MouseLeave
        With Me.lblpbTotalOpenStudies
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Regular)
            End If
            .Font = f
            .Cursor = Cursors.Default
        End With
    End Sub
    Private Sub lblpbInProgressReport_MouseLeave(sender As Object, e As System.EventArgs) Handles lblpbInProgressReport.MouseLeave
        With Me.lblpbInProgressReport
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Regular)
            End If
            .Font = f
            .Cursor = Cursors.Default
        End With
    End Sub
    Private Sub lblpbDraftReport_MouseLeave(sender As Object, e As System.EventArgs) Handles lblpbDraftReport.MouseLeave
        With Me.lblpbDraftReport
            Dim f
            If .Font.Bold = True Then
                f = New Font(.Font.Name, .Font.Size, FontStyle.Bold)
            Else
                f = New Font(.Font.Name, .Font.Size, FontStyle.Regular)
            End If
            .Font = f
            .Cursor = Cursors.Default
        End With
    End Sub

    Private Sub pbTotalGuWuStudies_click(sender As System.Object, e As System.EventArgs) Handles pbTotalGuWuStudies.Click
        Me.rbTotalGuWuStudies.Select()
    End Sub
    Private Sub pbFinalReport_click(sender As System.Object, e As System.EventArgs) Handles pbFinalReport.Click
        Me.rbFinalReports.Select()
    End Sub
    Private Sub pbTotalOpenStudies_click(sender As System.Object, e As System.EventArgs) Handles pbTotalOpenStudies.Click
        Me.rbTotalOpenStudies.Select()
    End Sub
    Private Sub pbInProgressReport_click(sender As System.Object, e As System.EventArgs) Handles pbInProgressReport.Click
        Me.rbInProgressStudies.Select()
    End Sub
    Private Sub pbDraftReportt_click(sender As System.Object, e As System.EventArgs) Handles pbDraftReport.Click
        Me.rbDraftReports.Select()
    End Sub

    Private Sub frmConsole_Resize(sender As Object, e As EventArgs) Handles Me.Resize

        Me.dgvDashboard.AutoResizeColumns()
        Me.dgvDashboard.AutoResizeRows()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim intR As Short = InputBox("Enter radius", "Enter radius", 25)
        Call buttonBorderRadius(Me.cmdConfig, intR)

    End Sub
End Class