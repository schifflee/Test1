Option Compare Text

Imports System
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Word
Imports System.Data.DataTable

Public Class frmOutliers

    Public tblAnalytes As New System.Data.DataTable
    Public boolFormLoad As Boolean = True
    Public tblResults As New System.Data.DataTable
    Public tblResultsTemp As New System.Data.DataTable
    Public boolFirstRow As Boolean = False
    Public tblCritZ As New System.Data.DataTable
    Public boolSummary As Boolean = False
    Public tblAllSummary As New System.Data.DataTable
    Public boolAllSummary As Boolean = False
    Public intConc As Short
    Public intArea As Short
    Public intISArea As Short
    Public boolHold As Boolean = False
    Public id_tblResults As Int64
    Dim boolChangeOutlier As Boolean = False

    Sub SetProgress()

        Dim dgv As DataGridView = Me.dgvResults

        Me.lblProgress.Location = dgv.Location
        Me.lblProgress.Size = dgv.Size

        If boolFormLoad Then
            Me.panFormLoadProgress.Dock = DockStyle.Fill
            Me.panFormLoadProgress.BringToFront()
            Me.panFormLoadProgress.Refresh()
        End If


    End Sub

    Sub PlaceForm()

        ' Retrieve the working rectangle from the Screen class
        ' using the PrimaryScreen and the WorkingArea properties. 
        Dim workingRectangle As System.Drawing.Rectangle = Screen.PrimaryScreen.WorkingArea

        ' Set the size of the form slightly less than size of 
        ' working rectangle.
        'Me.Size = New System.Drawing.Size(workingRectangle.Width - 10, workingRectangle.Height - 10)
        Dim a, b, c, d

        a = workingRectangle.Width
        b = Me.Width
        c = (a - b) / 2

        Me.Size = New System.Drawing.Size(Me.Width, workingRectangle.Height - 10)

        ' Set the location so the entire form is visible.
        Me.Location = New System.Drawing.Point(c, 5)


    End Sub

    Sub FillSort()

        Me.cbxSortSummaryAll.Items.Add("Analyte")
        Me.cbxSortSummaryAll.Items.Add("Table Name")
        Me.cbxSortSummaryAll.SelectedIndex = 0

    End Sub

    Private Sub frmOutliers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        Call PlaceForm()

        Dim var1

        boolFormLoad = True

        Call FillSort()

        Call DoubleBufferControl(Me, "dgv")

        Me.rbGrubbs.Checked = True 'do this because the rb change event is getting triggered later for some reason

        Me.tblAnalytes = tblAnalytesHome.Copy

        'complete legend
        Dim str1 As String

        str1 = "(1) Grubbs, Frank (February 1969), Procedures for Detecting Outlying Observations in Samples, Technometrics, Vol. 11, No. 1, pp. 1-21."
        str1 = str1 & ChrW(10) & "(2) W. J. Dixon, Processing Data for Outliers, Biometrics, Vol 9, No. 1, 74-89, 1953"

        Me.lblLegend.Text = str1

        'Me.lblProgress.Visible = True
        If boolFormLoad Then
            str1 = "Building tables..."
            Me.lblFormLoadProgress.Text = str1
            Me.panFormLoadProgress.Visible = True
            Me.panFormLoadProgress.Refresh()
            Me.lblFormLoadProgress.Refresh()
        End If

        Call FillCL()

        Call Initialize_tblCritZ()

        Call LoadGTable()

        Call SetGTable()

        Call VisRBs()

        Call Initialize_tblResults()

        Call InitializedgvAnalytes(True)

        Call FilldgvTables() 'will call filldgvanalytes

        Try
            Call Fill_tblResultsFull() 'this fills tblResults with all data
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        Call GatherAll()

        'now suspend form
        'Me.SuspendLayout()

        Dim int1 As Int32
        int1 = tblResults.Rows.Count

        Call ConfigResultsdgv()

        'Call RowSelect(-1, -1, "AA", True, 0)

        Call SummaryAllHeaders()

        str1 = "Summarizing Results..."
        Me.lblProgress.Text = str1
        Me.lblProgress.Refresh()
        If boolFormLoad Then
            Me.lblFormLoadProgress.Text = str1
            Me.lblFormLoadProgress.Refresh()
        End If

        boolFormLoad = False

        Call SelectTableRow()

        Call ToolTipSet()

        Call EnableStuff()

        'pesky
        Me.dgvSummary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Call FilterSummary()
        Call SetColWidths()
        Call dgvAutoSizeCols(Me.dgvSummary)

        Me.lblProgress.Visible = False
        Me.panFormLoadProgress.Visible = False


        'Me.ResumeLayout()

        str1 = str1 'debug

    End Sub

    Sub FillCL()

        Me.cbxCL.Items.Add(90)
        Me.cbxCL.Items.Add(95)
        Me.cbxCL.Items.Add(99)

        Me.cbxCL.SelectedIndex = 0

    End Sub

    Sub dgvAutoSizeCols(dgv As DataGridView)

        dgv.AutoResizeColumns()

    End Sub

    Sub SetStatsColumns(ByRef dgv As DataGridView)

        Dim stats1, stats2 As String
        Dim stats1Tooltip, stats2Tooltip As String
        Dim strSD As String

        Try
            stats1 = "Z"
            stats2 = "Crit" & ChrW(10) & "Z"
            If Me.rbGrubbs.Checked Then
                stats1 = "Z"
                stats2 = "Crit" & ChrW(10) & "Z"
                stats1Tooltip = "Calculated Z ratio: absolute value of [(mean-value)/SD]"
                stats2Tooltip = "Criteria for Z (based on n samples," _
                    & vbCrLf & "P<0.05 for values higher than this)"
            ElseIf Me.rbStdDev.Checked Then
                strSD = Me.txtStdDev.Text
                stats1 = "-" & strSD & ChrW(10) & "SD"
                stats1Tooltip = "Lower Limit (" & strSD & " standard deviations below the mean)"
                stats2 = "+" & strSD & ChrW(10) & "SD"
                stats2Tooltip = "Upper Limit (" & strSD & " standard deviations above the mean)"
            ElseIf Me.rbDixon.Checked Then
                stats1 = "R"
                stats1Tooltip = "Calculated R (Note: In Dixon test, only points farthest" _
                     & vbCrLf & "from mean are considered.  Others are shown as 'NA')"
                stats2 = "Crit" & ChrW(10) & "R"
                stats2Tooltip = "Criteria for R (based on n samples," _
                    & vbCrLf & "P<0.05 for values higher than this)"
            End If

            dgv.Columns("STATS1").Visible = True
            dgv.Columns("STATS1").HeaderText = stats1 '"Z"
            dgv.Columns("STATS1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv.Columns("STATS1").SortMode = DataGridViewColumnSortMode.NotSortable
            dgv.Columns("STATS1").ToolTipText = stats1Tooltip

            dgv.Columns("STATS2").Visible = True
            dgv.Columns("STATS2").HeaderText = stats2 '"CritZ"
            dgv.Columns("STATS2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv.Columns("STATS2").SortMode = DataGridViewColumnSortMode.NotSortable
            dgv.Columns("STATS2").ToolTipText = stats2Tooltip
        Catch ex As Exception

        End Try

    End Sub

    Sub RowSelect(ByVal idTS As Int64, ByVal idT As Int64, ByVal charAnalyte As String, ByVal boolAll As Boolean, ByVal intRow As Short)

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView

        dgv1 = Me.dgvTables
        dgv2 = Me.dgvResults

        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS1 As String
        Dim strS2 As String
        Dim idRT As Integer
        Dim intL1 As Int16
        Dim intL2 As Int16
        Dim int1 As Short
        Dim int2 As Short
        Dim dv As system.data.dataview
        Dim Count1 As Short
        Dim var1, var2, var3
        Dim intARow As Short
        Dim dgvA As DataGridView
        Dim strAnal As String
        Dim strAnalC As String
        Dim intConc As Short
        Dim intArea As Short
        Dim intISArea As Short
        Dim varIntStd
        Dim str2 As String
        Dim dgvT As DataGridView
        'Dim intRow As Short
        Dim boolAssigned As Boolean
        Dim AnalyteIndex As Short
        Dim MasterAssayID As Int64

        Dim stats1 As String
        Dim stats2 As String
        Dim strSD As String

        Dim intGroup As Short

        Dim idTConfig As Short

        boolAssigned = False

        Dim intAnalyteID As Int64
        Dim strMatrix As String

        Cursor.Current = Cursors.WaitCursor
        'Try
        varIntStd = "No"

        Dim boolRAS As Boolean 'requires assigned samples

        dgvA = Me.dgvAnalytes
        dgvT = Me.dgvTables
        strAnal = ""

        'remove dgv2 datasource
        dgv2.SuspendLayout()

        'don't need to do this
        strF = "INTGROUP < 0"
        dv = dgvResults.DataSource

        int1 = Me.tblResults.Rows.Count 'debug

        int1 = dv.Count 'debug

        dv.RowFilter = strf

        idRT = idT

        If dgvA.CurrentRow Is Nothing Then
            intARow = 0
        Else
            intARow = dgvA.CurrentRow.Index
        End If

        If boolAll Then
            'idRT = idT
            strAnal = charAnalyte

        Else
            'find dgvA current row


            ''find dgvT current row
            'If dgvT.CurrentRow Is Nothing Then
            '    intRow = 0
            'Else
            '    intRow = dgvT.CurrentRow.Index
            'End If

            int1 = dgv1("BOOLREQUIRESSAMPLEASSIGNMENT", intRow).Value

            If intRow = -1 Or dgvA.RowCount = 0 Then
                strF1 = "ID_TBLSTUDIES = -1"
            Else
                'idRT = dgv1("ID_TBLREPORTTABLE", intRow).Value
                strAnal = dgvA("ANALYTEDESCRIPTION", intARow).Value
                varIntStd = dgvA("IsIntStd", intARow).Value
            End If
        End If

        intAnalyteID = NZ(dgvA("ANALYTEID", intARow).Value, -1)
        strMatrix = NZ(dgvA("MATRIX", intARow).Value, "NA")
        intGroup = dgvA("INTGROUP", intARow).Value

        'determine if data should be obtained from assignedsamples
        int1 = dgv1("BOOLREQUIRESSAMPLEASSIGNMENT", intRow).Value
        idTConfig = dgv1("ID_TBLCONFIGREPORTTABLES", intRow).Value
        If int1 = 0 Then
            boolAssigned = False
        Else
            boolAssigned = True
        End If

        'determine if table requires assigned samples
        Dim rowAS() As DataRow
        strF1 = "ID_TBLCONFIGREPORTTABLES = " & idTConfig
        rowAS = tblConfigReportTables.Select(strF1)
        Dim intAS As Short
        intAS = rowAS(0).Item("BOOLREQUIRESSAMPLEASSIGNMENT")
        Dim boolReqAssignment As Boolean
        If intAS = 0 Then
            boolReqAssignment = False
        Else
            boolReqAssignment = True
        End If

        'remove datasource from dgv2

        'dgv2.DataSource = Nothing
        'dgv2.DataBindings.Clear()
        'dgv2.Refresh()

        Dim strFG As String

        strFG = "INTGROUP = " & intGroup

        strF1 = "ID_TBLREPORTTABLE = " & idRT & " AND INTGROUP = " & intGroup
        str2 = "CHARHELPER1 ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
        Try
            dv.RowFilter = strF1
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        int1 = Me.tblResults.Rows.Count 'debug
        int1 = int1
        int1 = dv.Count 'debug
        int1 = int1

        Try
            dv.AllowNew = False
            dv.AllowEdit = False
            dv.AllowDelete = False

            'dgv2.DataSource = dv

            Call ConfigResultsdgv()

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        intConc = dgv1("BOOLSHOWCONC", intRow).Value
        intArea = dgv1("BOOLSHOWAREA", intRow).Value
        intISArea = dgv1("BOOLINCLUDEIS", intRow).Value
        Dim idC As Int64
        idC = dgv1("ID_TBLCONFIGREPORTTABLES", intRow).Value


        'make all columns invisible
        For Count1 = 0 To dgv2.ColumnCount - 1
            dgv2.Columns(Count1).Visible = False
            dgv2.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgv2.Columns(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        Next

        dgv2.Columns("INTNUMBER").Visible = True
        dgv2.Columns("INTNUMBER").HeaderText = "#"

        dgv2.Columns("RUNID").Visible = True
        dgv2.Columns("RUNID").HeaderText = "Run" & ChrW(10) & "ID"

        dgv2.Columns("INTN").Visible = True
        dgv2.Columns("INTN").HeaderText = "n"

        dgv2.Columns("RUNSAMPLEORDERNUMBER").Visible = True
        dgv2.Columns("RUNSAMPLEORDERNUMBER").HeaderText = "Seq #"

        dgv2.Columns("CHARHELPER1").Visible = True
        dgv2.Columns("CHARHELPER1").HeaderText = "Level"

        dgv2.Columns("CHARHELPER2").Visible = True
        dgv2.Columns("CHARHELPER2").HeaderText = "Ident2"

        dgv2.Columns("NOMCONC").Visible = True
        dgv2.Columns("NOMCONC").HeaderText = "Nom" & ChrW(10) & "Conc"

        SetStatsColumns(dgv2)

        dgv2.Columns("CHAROUTLIER").Visible = True
        dgv2.Columns("CHAROUTLIER").HeaderText = "Outlier" & ChrW(10) & "(X)" ' "Outlier (X)"

        'dgv2.Columns("CONCENTRATION").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

        'dgv2.Columns("ANALYTEAREA").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

        'dgv2.Columns("INTERNALSTANDARDAREA").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight


        Dim boolRCConc As Boolean
        Dim boolRCPA As Boolean
        Dim boolRCPARatio As Boolean
        Dim boolIncludeISTbl As Boolean

        Call ReturnColumnTypes(boolRCConc, boolRCPA, boolRCPARatio, boolIncludeISTbl, idT) 'this will return bools

        If intConc = -1 Then
            dgv2.Columns.Item("CONCENTRATION").Visible = True
            dgv2.Columns("CONCENTRATION").HeaderText = "Conc."
            dgv2.Columns.Item("ANALYTEAREA").Visible = False
            dgv2.Columns.Item("INTERNALSTANDARDAREA").Visible = False
        ElseIf intArea = -1 Or intISArea = -1 Then
            If StrComp(varIntStd, "No", CompareMethod.Text) = 0 Then
                If boolRCPARatio Then
                    dgv2.Columns.Item("CONCENTRATION").Visible = True
                    dgv2.Columns("CONCENTRATION").HeaderText = "Peak" & ChrW(10) & "Area" & ChrW(10) & "Ratio"
                    dgv2.Columns.Item("ANALYTEAREA").Visible = False
                    dgv2.Columns.Item("INTERNALSTANDARDAREA").Visible = False
                Else
                    dgv2.Columns.Item("CONCENTRATION").Visible = False
                    dgv2.Columns("CONCENTRATION").HeaderText = "Conc."
                    dgv2.Columns.Item("ANALYTEAREA").Visible = True
                    dgv2.Columns("ANALYTEAREA").HeaderText = "Peak" & ChrW(10) & "Area"
                    dgv2.Columns.Item("INTERNALSTANDARDAREA").Visible = False
                    dgv2.Columns("INTERNALSTANDARDAREA").HeaderText = "IS" & ChrW(10) & "Peak" & ChrW(10) & "Area"
                End If
            Else
                dgv2.Columns.Item("CONCENTRATION").Visible = False
                dgv2.Columns("CONCENTRATION").HeaderText = "Conc."
                dgv2.Columns.Item("ANALYTEAREA").Visible = False
                dgv2.Columns("ANALYTEAREA").HeaderText = "Peak" & ChrW(10) & "Area"
                dgv2.Columns.Item("INTERNALSTANDARDAREA").Visible = True
                dgv2.Columns("INTERNALSTANDARDAREA").HeaderText = "IS" & ChrW(10) & "Peak" & ChrW(10) & "Area"
            End If
        End If


        'dgv2.AutoResizeColumns()

        'this will sync dgvResults and dgvSummary headingtext, among other things
        Call SyncSummaryColumns()

        'Make minwidths
        Dim dgvS As DataGridView = Me.dgvSummary
        dgvS.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvS.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        Call SetColWidths()

        'now see if there are blank identification columns
        Erase rows1
        Dim str1 As String

        str1 = "NOMCONC"
        strF = str1 & " IS NULL"
        rows1 = tblResults.Select(strF)
        int1 = rows1.Length
        'If rows1.Length = 0 Then
        '    dgv2.Columns(str1).Visible = True
        'Else
        '    dgv2.Columns(str1).Visible = False
        'End If

        str1 = "CHARHELPER1"
        strF = str1 & " IS NULL"
        'Erase rows1
        'rows1 = tblResults.Select(strF)
        'int1 = rows1.Length
        'If rows1.Length = 0 Then
        '    dgv2.Columns(str1).Visible = True
        'Else
        '    dgv2.Columns(str1).Visible = False

        'End If

        str1 = "CHARHELPER2"
        strF = str1 & " IS NULL"
        Erase rows1
        rows1 = tblResults.Select(strF)
        int1 = rows1.Length
        If rows1.Length = 0 Then
            dgv2.Columns(str1).Visible = True
        Else
            dgv2.Columns(str1).Visible = False

        End If

        Call SetStatsColumns(Me.dgvResults)

        dgv2.AutoResizeRows()
        'dgv2.AutoResizeColumns()

        dgv2.ResumeLayout()

        Cursor.Current = Cursors.Default

        If boolFirstRow = False Then
            boolFirstRow = True
        End If

        Call ChangeSummaryRow()

        Call SyncSummaryColumns()

        dgv2.AutoResizeColumns() 'dgvResults

        'pesky
        dgvS.AutoResizeColumns()



        'Catch ex As Exception
        '    Dim strM As String
        '    strM = "Hmmm. There seems to be a problem preparing the Outlier information for this table."
        '    strM = strM & "Error: " & ex.Message
        '    MsgBox(strM, MsgBoxStyle.Information, "Problem...")
        'End Try


    End Sub

    Sub ReturnColumnTypes(ByRef boolRCConc As Boolean, ByRef boolRCPA As Boolean, ByRef boolRCPARatio As Boolean, ByRef boolIncludeISTbl As Boolean, idT As Int64)

        Dim dtbl As System.Data.DataTable = tblTableProperties
        Dim strF As String
        Dim var1, var2, var3, var4

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idT

        Dim rows() As DataRow = dtbl.Select(strF)
        If rows.Length = 0 Then
            boolRCConc = True
            boolRCPA = False
            boolRCPARatio = False
            boolIncludeISTbl = False
        Else
            var1 = NZ(rows(0).Item("BOOLRCCONC"), -1)
            var2 = NZ(rows(0).Item("BOOLRCPA"), 0)
            var3 = NZ(rows(0).Item("BOOLRCPARATIO"), 0)
            var4 = NZ(rows(0).Item("BOOLINCLUDEISTBL"), 0)

            If var1 = 0 Then
                boolRCConc = False
            Else
                boolRCConc = True
            End If

            If var2 = 0 Then
                boolRCPA = False
            Else
                boolRCPA = True
            End If

            If var3 = 0 Then
                boolRCPARatio = False
            Else
                boolRCPARatio = True
            End If

            If var4 = 0 Then
                boolIncludeISTbl = False
            Else
                boolIncludeISTbl = True
            End If

        End If



    End Sub

    Sub SetColWidths()

        Dim dgv2 As DataGridView = Me.dgvResults
        Dim dgvS As DataGridView = Me.dgvSummary
        Dim var1

        Try

            dgv2.Columns("CHAROUTLIER").MinimumWidth = 50
            dgv2.Columns("RUNID").MinimumWidth = 40
            dgv2.Columns("NOMCONC").MinimumWidth = 40
            dgv2.Columns("STATS1").MinimumWidth = 40
            dgv2.Columns("STATS2").MinimumWidth = 40
            dgv2.Columns("CONCENTRATION").MinimumWidth = 50
            dgv2.Columns("ANALYTEAREA").MinimumWidth = 50
            dgv2.Columns("INTERNALSTANDARDAREA").MinimumWidth = 50
            dgv2.Columns("INTNUMBER").MinimumWidth = 25
            dgv2.Columns("CHARHELPER1").MinimumWidth = 60

            dgvS.Columns("CHAROUTLIER").MinimumWidth = 50
            dgvS.Columns("RUNID").MinimumWidth = 40
            dgvS.Columns("NOMCONC").MinimumWidth = 40
            dgvS.Columns("STATS1").MinimumWidth = 40
            dgvS.Columns("STATS2").MinimumWidth = 40
            dgvS.Columns("CONCENTRATION").MinimumWidth = 50
            dgvS.Columns("ANALYTEAREA").MinimumWidth = 50
            dgvS.Columns("INTERNALSTANDARDAREA").MinimumWidth = 50
            dgvS.Columns("INTNUMBER").MinimumWidth = 25
            dgvS.Columns("CHARHELPER1").MinimumWidth = 60

            'dgv2.Columns("INTNUMBER").Width = 20
            'dgv2.Columns("RUNID").Width = 25
            'dgv2.Columns("RUNSAMPLEORDERNUMBER").Width = 30

            'dgv2.Columns("NOMCONC").Width = 40
            'dgv2.Columns("CONCENTRATION").Width = 40
            'dgv2.Columns("ANALYTEAREA").Width = 75
            'dgv2.Columns("INTERNALSTANDARDAREA").Width = 75


            'dgv2.Columns("CHARHELPER1").Width = 40
            'dgv2.Columns("CHARHELPER2").Width = 40

            'dgv2.Columns("STATS1").Width = 30
            'dgv2.Columns("STATS2").Width = 30
            ''dgv2.Columns("CHAROUTLIER").Width = 50
            'dgv2.Columns("INTN").Width = 20
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


    End Sub


    Sub Fill_tblResults_QCStds(ByVal boolReport As Boolean, ByVal charAnalyte As String, ByVal AnalyteIndex As Short, ByVal MasterAssayID As Int64, ByRef intRow As Int64, intGroup As Short, intAnalyteID As Int64, strMatrix As String, idCT As Int64, idT As Int64, ByRef arrResults As Object, ByVal intColsR As Short, arrColsR As Object, boolAS As Boolean, strTable As String, intConc As Short, intArea As Short, intISArea As Short)

        'intRow = tbleResults Row #
        ''ID_TBLCONFIGREPORTTABLES	ID_TBLREPORTTABLE	CHARHEADINGTEXT	ANALYTEDESCRIPTION
        'idCT, idT
        'boolAS = True if use tblAssignedSamples

        Dim dtbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim intL1 As Short
        Dim dtbl2 As System.Data.DataTable
        Dim rows2() As DataRow
        Dim intL2 As Short
        Dim dtbl3 As System.Data.DataTable
        Dim rows3() As DataRow
        Dim intL3 As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim strF As String
        Dim strF1 As String
        Dim strS1 As String
        Dim strF2 As String
        Dim strS2 As String
        Dim strF3 As String
        Dim strS3 As String
        Dim strF4 As String
        Dim strF5 As String

        Dim dv As System.Data.DataView
        Dim intRows As Int16
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1, var2, var3
        Dim rows() As DataRow
        Dim strComp1 As String
        Dim strComp2 As String
        Dim strA As String
        Dim strB As String
        Dim strC As String
        Dim strAll As String
        Dim varA, varB, varC, varD
        Dim nomConc As Decimal
        Dim intLUA As Short
        Dim intLUC As Short

        Dim AssayID As Int16
        Dim intRunNum As Short
        Dim varASSID
        Dim numAF As Decimal
        Dim idOld As Int64
        Dim intNumber As Int16
        Dim intNumberO As Int16

        Dim boolRCConc As Boolean
        Dim boolRCPA As Boolean
        Dim boolRCPARatio As Boolean
        Dim boolIncludeISTbl As Boolean

        Call ReturnColumnTypes(boolRCConc, boolRCPA, boolRCPARatio, boolIncludeISTbl, idT) 'this will return bools

        Dim arrCols()
        ReDim arrCols(Me.tblResults.Columns.Count)

        For Count1 = 1 To Me.tblResults.Columns.Count
            str1 = Me.tblResults.Columns(Count1 - 1).ColumnName
            arrCols(Count1) = str1
        Next

        Dim dtblR As System.Data.DataTable

        'Dim intCols As Short = Me.tblResults.Columns.Count

        If boolHold Then
            Exit Sub
        End If

        If boolReport Then
            dtblR = Me.tblResultsTemp
        Else
            dtblR = Me.tblResults
        End If

        Dim boolIsIS As Boolean = False

        '20160304 LEE: No, do not clear
        'dtblR.Clear()

        If intAnalyteID = -1 Then
            boolIsIS = True
            var1 = var1
        End If

        dtbl1 = tblBCQCStdsAssayID ' tblQCAI
        dtbl2 = tblQCReps
        If boolAS Then
            dtbl3 = tblAssignedSamples
        Else
            dtbl3 = tblBCQCConcs
        End If
        'dtbl3 = tblBCQCConcs

        ''get filtered tblQCAI
        'strF2 = "AnalyteIndex = " & AnalyteIndex & " AND MASTERASSAYID = " & MasterAssayID & " AND ASSAYID = " & varASSID
        ''strF2 = "ANALYTEDESCRIPTION = '" & charAnalyte & "'"
        'strS2 = "CONCENTRATION ASC"
        'rows1 = dtbl1.Select(strF2, strS2)
        'intL1 = rows1.Length

        Dim rowsNomConc() As DataRow
        Dim strFNomConc As String
        Dim intRunID As Int16
        Dim intLevel As Short
        Dim numNomConc As Decimal

        'retrieve only accepted values from dtbl3
        intRunNum = 0

        'need further delienation depending on idCT
        'CHARTABLENAME	                                                                        CHARHELPER	ID_TBLCONFIGREPORTTABLES	NUMHELPERNUMBER	
        'CHARTABLENAME	                                                                        CHARHELPER	                ID_CT	NUMHELPERNUMBER	

        'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision	                    QC LLOQ	                    11	1	
        'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision	                    QC Low		                11	1	
        'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision	                    QC Mid		                11	1	
        'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision	                    QC High		                11	1	
        'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision	                    Outlier		                11	2	?
        'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision	                    [Clear]		                11	2	?
        'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision	                    QC Diln		                11	1	

        'Summary of Combined Recovery	                                                        QC		                    13	1	
        'Summary of Combined Recovery	                                                        RS - Recovery Standard	    13	1	

        'Summary of True Recovery	                                                            QC		                    14	1	
        'Summary of True Recovery	                                                            PES - Post Extraction Spike	14	1	

        'Summary of Suppression/Enhancement	                                                    RS - Recovery Standard	    15	1	
        'Summary of Suppression/Enhancement	                                                    PES - Post Extraction Spike	15	1	

        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 1		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 2		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 3		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 4		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 5		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 6		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 7		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 8		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 9		                17	1	
        'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments	Lot 10		                17	1	

        '[Period Temp] Stock Solution Stability Assessment	                                    Old Stock Solution	        22	1	
        '[Period Temp] Stock Solution Stability Assessment	                                    New Stock Solution	        22	1	

        '[Period Temp] Spiking Solution Stability Assessment	                                Old Spiking Solution	    23	1	
        '[Period Temp] Spiking Solution Stability Assessment	                                New Spiking Solution	    23	1	

        '[Period Temp] Long-Term QC Std Storage Stability	                                    T(0) Analysis	            29	2	
        '[Period Temp] Long-Term QC Std Storage Stability	                                    Long Term Analysis	        29	2	

        'Ad Hoc QC Stability Comparison Table	                                                [Term 2]                    32	2	

        'Selectivity in Individual Lots Table v1	                                            Blank With IS	            34	2	
        'Selectivity in Individual Lots Table v1	                                            Blank WithOut IS	        34	2	
        'Selectivity in Individual Lots Table v1	                                            LLOQ	                    34	2	
        'Selectivity in Individual Lots Table v1	                                            Lot 1	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 2	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 3	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 4	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 5	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 6	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 7	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 8	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 9	                    34	1	
        'Selectivity in Individual Lots Table v1	                                            Lot 10	                    34	1	

        'Carryover in Individual Lots Table v1	                                                LLOQ	                    35	1	
        'Carryover in Individual Lots Table v1	                                                ULOQ	                    35	1	
        'Carryover in Individual Lots Table v1	                                                Blank	                    35	1	


        'has Term 1's
        Dim rowsT1() As DataRow
        Dim tblT1 As System.Data.DataTable
        Dim intT1 As Short = 0
        Dim intT1a As Short = 1
        Dim boolT1 As Boolean = False
        Dim CountT1 As Short

        Dim rowsT2() As DataRow
        Dim tblT2 As System.Data.DataTable
        Dim intT2 As Short = 0
        Dim intT2a As Short = 1
        Dim boolT2 As Boolean = False
        Dim CountT2 As Short

        Erase rowsT1
        Dim strFT1 As String
        Dim strFT2 As String
        Dim dvT1 As DataView
        Dim dvT2 As DataView
        Dim strS As String
        strFT1 = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND NUMHELPERNUMBER = 1"
        Select Case idCT
            Case 13, 14, 15, 17, 22, 23, 29, 32, 34, 35
                strS = "ID_TBLASSIGNEDSAMPLESHELPER ASC"
                dvT1 = New DataView(tblAssignedSamplesHelper, strFT1, strS1, DataViewRowState.CurrentRows)
                tblT1 = dvT1.ToTable("a", True, "CHARHELPER")
                intT1 = tblT1.Rows.Count
                If boolIsIS Then
                    var1 = var1 'debug
                End If
        End Select
        If intT1 > 0 Then
            boolT1 = True
            intT1a = intT1
        End If

        'has Term 2's
        strFT2 = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND NUMHELPERNUMBER = 2"
        Select Case idCT
            Case 29, 32, 34
                strS = "ID_TBLASSIGNEDSAMPLESHELPER ASC"
                dvT2 = New DataView(tblAssignedSamplesHelper, strFT2, strS1, DataViewRowState.CurrentRows)
                tblT2 = dvT2.ToTable("a", True, "CHARHELPER")
                intT2 = tblT2.Rows.Count
        End Select
        If intT2 > 0 Then
            boolT2 = True
            intT2a = intT2
        End If

        'NOTE: ANOVA table must be done by NomConc and RunID
        'may have more than one Mid, so need to evaluate nomconc
        Dim boolRunID As Boolean = False
        Dim CountRunID As Short
        Dim dvRunID As DataView
        Dim tblRunID As System.Data.DataTable
        Dim strFRunID As String
        Dim intRID As Short = 0
        Dim intRIDa As Short = 1
        Select Case idCT
            Case 11
                boolRunID = True
                'get unique runids in tblassignedsamples
                'strFRunID = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND INTGROUP = " & intGroup
                strFRunID = "ID_TBLREPORTTABLE = " & idT & " AND INTGROUP = " & intGroup
                dvRunID = New DataView(tblAssignedSamples, strFRunID, "RUNID ASC, NOMCONC ASC", DataViewRowState.CurrentRows)
                tblRunID = dvRunID.ToTable("a", True, "RUNID", "NOMCONC")
                intRID = tblRunID.Rows.Count
                If intRID = 0 Then
                    boolRunID = False
                    intRIDa = 0
                End If
        End Select
        If intRID > 0 Then
            boolRunID = True
            intRIDa = intRID
        End If

        Select Case idCT
            Case 13, 14, 15, 17, 22, 23
                strS3 = "CHARHELPER1 ASC, NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            Case Else
                strS3 = "NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
        End Select

        Dim strHelper2 As String = ""

        For CountT2 = 1 To intT2a

            strF = ""
            strF1 = ""
            strHelper2 = ""
            If boolT2 Then
                str1 = tblT2.Rows(CountT2 - 1).Item("CHARHELPER")
                strF = "CHARHELPER2 = '" & str1 & "'"
                strF1 = strF
            End If

            For CountT1 = 1 To intT1a

                strF = ""
                strF2 = strF1
                If boolT1 Then
                    str1 = tblT1.Rows(CountT1 - 1).Item("CHARHELPER")
                    If Len(strF1) = 0 Then
                        strF = "CHARHELPER1 = '" & str1 & "'"
                    Else
                        strF = strF1 & " AND CHARHELPER1 = '" & str1 & "'"
                    End If
                    strF2 = strF
                End If

                For CountRunID = 1 To intRIDa

                    strF = ""
                    strF3 = strF2
                    If boolRunID Then
                        str1 = tblRunID.Rows(CountRunID - 1).Item("RUNID").ToString
                        str2 = tblRunID.Rows(CountRunID - 1).Item("NOMCONC").ToString
                        If Len(strF2) = 0 Then
                            strF = "RUNID = " & str1 & " AND NOMCONC = " & str2
                        Else
                            strF = strF2 & " AND RUNID = " & str1 & " AND NOMCONC = " & str2
                        End If
                        strF3 = strF
                    End If


                    '20160304 LEE:
                    'new filter for groups
                    strF = ""
                    strF4 = strF3
                    If boolAS Then
                        If Len(strF3) = 0 Then
                            strF = "INTGROUP = " & intGroup & " AND ID_TBLREPORTTABLE = " & idT
                        Else
                            strF = strF3 & " AND INTGROUP = " & intGroup & " AND ID_TBLREPORTTABLE = " & idT
                        End If
                        If boolIsIS Then
                            strF = strF & " AND CHARANALYTE = '" & CleanText(charAnalyte) & "'"
                        End If
                        strF4 = strF
                    Else
                        str2 = GetASSAYIDFilter(intGroup, False, False)
                        If Len(strF) = 0 Then
                            strF = "ANALYTEID = " & intAnalyteID & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND " & str2
                        Else
                            strF = strF3 & " AND ANALYTEID = " & intAnalyteID & " AND RUNTYPEID <> 3 AND RUNANALYTEREGRESSIONSTATUS = 3 AND " & str2
                        End If
                        strF4 = strF
                    End If

                    Try
                        rows3 = dtbl3.Select(strF4, strS3)
                    Catch ex As Exception
                        var1 = ex.Message
                        var1 = var1
                    End Try
                    intL3 = rows3.Length

                    str1 = ""
                    str2 = ""
                    strA = ""
                    strB = ""
                    strC = ""
                    strComp1 = ""
                    strComp2 = ""

                    Dim intCt As Short
                    Dim intLCt As Short
                    Dim var10, var11, var12
                    Dim numAnalyteArea As Double = 0
                    Dim numIntStdArea As Double = 0

                    Dim boolDo As Boolean
                    Dim intColNomConc As Short
                    Dim intColConc As Short
                    Dim intNoLevel As Int16 = 100
                    Dim intColH1 As Short

                    intCt = 0
                    intLCt = 0

                    If idCT = 13 Then
                        var1 = var1 'debug
                    End If

                    intNumber = 0
                    intRows = intL3

                    For Count1 = 0 To intL3 - 1

                        intNumber = intNumber + 1
                        intCt = intCt + 1

                        intRow = intRow + 1
                        If intRow > UBound(arrResults, 2) Then
                            ReDim Preserve arrResults(intColsR - 1, UBound(arrResults, 2) + 1000)
                        End If

                        id_tblResults = id_tblResults + 10

                        intLevel = 0
                        intRunID = 0

                        numAnalyteArea = 0
                        numIntStdArea = 0

                        'Note: if boolT2=true, then CHARHELPER2 must be evaluated before CHARHELPER1
                        If boolT2 Then
                            Try
                                var2 = rows3(Count1).Item("CHARHELPER2")
                            Catch ex As Exception
                                var2 = ""
                            End Try
                            strHelper2 = var2
                            var2 = Nothing
                        End If

                        Try
                            For Count2 = 1 To intColsR ' 20

                                str1 = UCase(arrColsR(1, Count2))
                                str2 = UCase(arrColsR(2, Count2))
                                int1 = arrColsR(3, Count2)

                                Select Case idCT
                                    Case 13, 14, 15, 22, 23
                                        Select Case str1
                                            Case "CHARHELPER1"
                                                str2 = str1
                                        End Select

                                End Select

                                If StrComp(str1, "CHARHELPER1", CompareMethod.Text) = 0 Then
                                    var1 = var1 'debug
                                End If

                                var2 = Nothing
                                boolDo = True

                                Select Case str1
                                    Case "INTNUMBER"
                                        var2 = intNumber
                                    Case "ID_TBLRESULTS"
                                        var2 = id_tblResults
                                    Case "NOMCONC" '3
                                        str1 = "NOMCONC"
                                        'var1 = System.Type.GetType("System.Decimal")
                                        str2 = "NOMCONC"
                                        intColNomConc = int1
                                        Try
                                            var2 = rows3(Count1).Item(str2)
                                        Catch ex As Exception
                                            var2 = ""
                                        End Try
                                        strA = CStr(NZ(var2, "")) '****
                                        var2 = Nothing
                                    Case "ID_TBLCONFIGREPORTTABLES" '17
                                        var2 = idCT
                                    Case "ANALYTEDESCRIPTION" '20
                                        var2 = charAnalyte
                                    Case "ID_TBLRESULTS"
                                        var2 = id_tblResults
                                    Case "INTGROUP"
                                        var2 = intGroup
                                    Case "ID_TBLREPORTTABLE" '18
                                        var2 = idT
                                    Case "ID_TBLASSIGNEDSAMPLES"
                                        var2 = intCt
                                        'Case "CHARHELPER1" 'do CHARHELPER1 later
                                        '    strB = CStr(NZ(var2, ""))
                                    Case "CHARHELPER2"
                                        Try
                                            var2 = rows3(Count1).Item(str2)
                                        Catch ex As Exception
                                            var2 = ""
                                        End Try
                                        strC = CStr(NZ(var2, "")) '****
                                        strHelper2 = strC
                                        If Len(strHelper2) = 0 Then
                                        Else
                                            var1 = var1 'debug
                                        End If
                                        'var2 = Nothing
                                        'Case "CHARHELPER1"
                                        '    Try
                                        '        var2 = rows3(Count1).Item(str2)
                                        '    Catch ex As Exception
                                        '        var2 = ""
                                        '    End Try
                                        '    Select Case idCT
                                        '        Case 13, 14, 15, 22, 23
                                        '            intLevel = intNoLevel
                                        '        Case Else
                                        '            intLevel = NZ(var2, 0)
                                        '    End Select
                                        '    strB = CStr(NZ(var2, ""))

                                    Case "CHARHEADINGTEXT"
                                        var2 = strTable
                                    Case Else
                                        If dtbl3.Columns.Contains(str2) Then
                                            boolDo = True
                                            Try
                                                var2 = rows3(Count1).Item(str2)
                                            Catch ex As Exception
                                                var1 = ex.Message
                                                var1 = var1
                                                var2 = Nothing
                                            End Try

                                            Select Case str1
                                                Case "ANALYTEAREA"
                                                    numAnalyteArea = SigFigArea(NZ(var2, 0), LSigFigArea, True, False)
                                                    var2 = numAnalyteArea
                                                Case "INTERNALSTANDARDAREA"
                                                    numIntStdArea = SigFigArea(NZ(var2, 0), LSigFigArea, True, False)
                                                    var2 = numIntStdArea
                                                Case "RUNID"
                                                    intRunID = NZ(var2, 0)
                                                Case "CONCENTRATION"
                                                    intColConc = int1
                                                    numAF = NZ(rows3(Count1).Item("ALIQUOTFACTOR"), 1)
                                                    var3 = var2 / numAF
                                                    var2 = SigFigOrDecString(NZ(var3, 0), LSigFig, False)
                                                    var2 = var2
                                                Case "CHARHELPER1"
                                                    Select Case idCT
                                                        Case 13, 14, 15, 22, 23
                                                            intLevel = intNoLevel
                                                        Case Else
                                                            intLevel = NZ(var2, 0)
                                                    End Select
                                                    strB = CStr(NZ(var2, ""))
                                                    'replace var2
                                                    If dtbl3.Columns.Contains(str1) Then
                                                        var3 = NZ(rows3(Count1).Item(str1), "")
                                                        If Len(var3) = 0 Then
                                                            var3 = var3
                                                        Else
                                                            var2 = var3
                                                        End If
                                                        If Len(strHelper2) = 0 Then
                                                        Else
                                                            'add strhelper2 to charhelper1
                                                            var2 = strHelper2 & " " & var2
                                                        End If
                                                    End If
                                            End Select

                                        End If

                                End Select

                                Select Case str1
                                    Case "CHARHELPER1"
                                        intColH1 = int1
                                        ''console.writeline("idCT: " & idCT & " CHARHELPER1: " & var2)
                                End Select

                                arrResults(Count2 - 1, intRow) = var2

                            Next Count2

                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try


                        Try

                            If intConc = -1 Then
                                'do nothing
                            ElseIf intArea = -1 Or intISArea = -1 Then
                                var10 = numAnalyteArea ' NewRow.Item("ANALYTEAREA")
                                var11 = numIntStdArea ' NewRow.Item("INTERNALSTANDARDAREA")
                                'this is area ratio
                                If boolRCPARatio Then
                                    If var11 = 0 Then
                                        var12 = 0
                                        var2 = 0
                                    Else
                                        var12 = var10 / var11
                                        var2 = SigFigAreaRatio(var12, LSigFigAreaRatio, False, False)
                                    End If
                                    arrResults(intColConc, intRow) = var2
                                End If

                            End If

                            'If boolRCPARatio Then
                            '    var10 = numAnalyteArea ' NewRow.Item("ANALYTEAREA")
                            '    var11 = numIntStdArea ' NewRow.Item("INTERNALSTANDARDAREA")
                            '    'this is area ratio

                            '    If var11 = 0 Then
                            '        var12 = 0
                            '        var2 = 0
                            '    Else
                            '        var12 = var10 / var11
                            '        var2 = SigFigAreaRatio(var12, LSigFigAreaRatio, True, False)
                            '    End If
                            '    arrResults(intColConc, intRow) = var2
                            'End If

                            'Select Case idCT
                            '    Case 13, 14, 15, 22, 23
                            '        var10 = numAnalyteArea ' NewRow.Item("ANALYTEAREA")
                            '        var11 = numIntStdArea ' NewRow.Item("INTERNALSTANDARDAREA")
                            '        'this is area ratio
                            '        If var11 = 0 Then
                            '            var12 = Format(0, "0.0000000")
                            '        Else
                            '            var12 = Format(RoundToDecimalRAFZ(var10 / var11, 7), "0.0000000")
                            '        End If

                            '        var2 = var12
                            '        'newRow.Item("CONCENTRATION") = var12
                            '        arrResults(intColConc, intRow) = var2
                            'End Select

                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try

                        'fix CHARHELPER1
                        If idCT = 12 Then
                            var1 = var1 'debug
                        End If
                        strFNomConc = "ANALYTEID = " & intAnalyteID & " AND LEVELNUMBER = " & intLevel & " AND RUNID = " & intRunID
                        rowsNomConc = dtbl1.Select(strFNomConc)
                        If rowsNomConc.Length = 0 Then
                            var1 = var1
                        Else
                            var1 = NZ(rowsNomConc(0).Item("ID"), "")
                            If Len(var1) = 0 Then
                            Else
                                'newRow.Item("NOMCONC") = var1
                                If Len(strHelper2) = 0 Then
                                    arrResults(intColH1, intRow) = var1
                                End If

                            End If
                        End If

                        'dtbl1 = tblBCQCStdsAssayID 
                        If boolAS Then
                            var2 = rows3(Count1).Item("NOMCONC")
                            arrResults(intColNomConc, intRow) = var2
                        Else
                            'strFNomConc = "ANALYTEID = " & intAnalyteID & " AND LEVELNUMBER = " & intLevel & " AND RUNID = " & intRunID
                            'rowsNomConc = dtbl1.Select(strFNomConc)
                            If rowsNomConc.Length = 0 Then
                                var1 = var1
                            Else
                                var1 = rowsNomConc(0).Item("CONCENTRATION")
                                If IsDBNull(var1) Then
                                Else
                                    'newRow.Item("NOMCONC") = var1
                                    arrResults(intColNomConc, intRow) = var1
                                End If
                            End If
                        End If

                        strAll = strA & strB & strC
                        If Count1 = 0 Then
                            strComp1 = strAll
                            strComp2 = strAll
                        Else
                            strComp2 = strAll
                        End If

                        If StrComp(strComp1, strComp2, CompareMethod.Text) = 0 Then 'ignore

                        Else 'correct data and insert a blank row

                            idOld = id_tblResults
                            intNumber = 1
                            id_tblResults = id_tblResults + 10
                            intRow = intRow + 1
                            intNoLevel = intNoLevel + 1

                            If intRow > UBound(arrResults, 2) Then
                                ReDim Preserve arrResults(intColsR - 1, UBound(arrResults, 2) + 1000)
                            End If

                            'move row of data ahead
                            Try
                                For Count3 = 0 To intColsR - 1
                                    var1 = arrResults(Count3, intRow - 1) 'debug
                                    arrResults(Count3, intRow) = arrResults(Count3, intRow - 1)
                                    str1 = UCase(arrCols(Count3 + 1))
                                    Select Case str1
                                        Case "ID_TBLRESULTS"
                                            arrResults(Count3, intRow) = id_tblResults
                                        Case "INTNUMBER"
                                            arrResults(Count3, intRow) = intNumber
                                    End Select
                                    Select Case idCT
                                        Case 13, 14, 15, 22, 23
                                            'need to replace intnolevel
                                            Select Case str1
                                                Case "CHARHELPER1"
                                                    'arrResults(Count3, intRow) = intNoLevel
                                            End Select
                                    End Select
                                Next
                            Catch ex As Exception
                                var1 = ex.Message
                                var1 = var1
                            End Try


                            'clear old data
                            For Count3 = 0 To intColsR - 1
                                arrResults(Count3, intRow - 1) = Nothing
                            Next

                            intCt = intCt + 1

                            For Count2 = 1 To intColsR
                                boolDo = True
                                str1 = UCase(arrCols(Count2))
                                Select Case str1
                                    Case "ID_TBLRESULTS"
                                        arrResults(Count2 - 1, intRow - 1) = idOld
                                    Case "ID_TBLASSIGNEDSAMPLES"
                                        arrResults(Count2 - 1, intRow - 1) = intCt
                                    Case "INTGROUP"
                                        arrResults(Count2 - 1, intRow - 1) = intGroup
                                    Case "ID_TBLREPORTTABLE"
                                        arrResults(Count2 - 1, intRow - 1) = idT
                                    Case Else
                                        boolDo = False
                                End Select

                            Next

                            strComp1 = strComp2

                        End If

                    Next Count1

                    var1 = var1 'debug

                    If intRows = 0 Then 'ignore
                    Else
                        'just add a row, no swapping
                        'add another row
                        'Dim newRowb As DataRow = dtblR.'newRow
                        'newRowb.BeginEdit()

                        idOld = id_tblResults
                        id_tblResults = id_tblResults + 10

                        intRow = intRow + 1
                        intNoLevel = intNoLevel + 1

                        If intRow > UBound(arrResults, 2) Then
                            ReDim Preserve arrResults(intColsR - 1, UBound(arrResults, 2) + 1000)
                        End If

                        intCt = intCt + 1

                        For Count2 = 1 To intColsR
                            str1 = UCase(arrCols(Count2))
                            Select Case str1
                                Case "ID_TBLRESULTS"
                                    arrResults(Count2 - 1, intRow) = id_tblResults
                                Case "ID_TBLASSIGNEDSAMPLES"
                                    arrResults(Count2 - 1, intRow) = intCt
                                Case "INTGROUP"
                                    arrResults(Count2 - 1, intRow) = intGroup
                                Case "ID_TBLREPORTTABLE"
                                    arrResults(Count2 - 1, intRow) = idT
                            End Select
                        Next

                        intNumber = 0

                    End If

                    var1 = var1 'debug

                Next CountRunID

                var1 = var1 'debug

            Next CountT1

        Next CountT2



    End Sub

    Sub Fill_tblResultsFull()

        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable

        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS1 As String
        Dim strS2 As String
        Dim idRT As Integer
        Dim intL1 As Int16
        Dim intL2 As Int16
        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2, var3
        Dim intARow As Short

        Dim str1 As String
        Dim str2 As String

        Dim strAnal As String
        Dim intConc As Short
        Dim intArea As Short
        Dim intISArea As Short
        Dim varIntStd

        'Dim intRow As Short
        Dim boolAssigned As Boolean
        Dim AnalyteIndex As Short
        Dim MasterAssayID As Int64

        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim Count3 As Int16

        Dim stats1 As String
        Dim stats2 As String
        Dim strSD As String
        Dim strP1 As String
        Dim strP2 As String
        Dim strP3 As String
        Dim intGroup As Short

        boolAssigned = False

        Cursor.Current = Cursors.WaitCursor
        'Try
        varIntStd = "No"

        Dim boolRAS As Boolean = False 'requires assigned samples

        Dim dgvA As DataGridView
        Dim dgvT As DataGridView
        Dim dgvR As DataGridView
        dgvA = Me.dgvAnalytes
        dgvT = Me.dgvTables
        dgvR = Me.dgvResults

        dtbl1 = tblAssignedSamples

        'use array to fill tblResults

        Dim intColsR As Short = Me.tblResults.Columns.Count
        Dim intRow As Int64
        Dim arrResults(intColsR - 1, 1000) As Object
        Dim arrColsR(3, intColsR)

        Call SetProgress()

        For Count2 = 1 To intColsR

            str1 = tblResults.Columns(Count2 - 1).ColumnName
            str2 = str1
            int1 = 0

            Select Case str1

                Case "NOMCONC" '3

                Case "CHARHELPER1" ' 4
                    str2 = "ASSAYLEVEL" '"CHARHELPER1"

                Case "CONCENTRATION" '6

            End Select
            arrColsR(1, Count2) = str1
            arrColsR(2, Count2) = str2
            arrColsR(3, Count2) = Count2 - 1
        Next


        strAnal = ""

        Dim strAnalC As String
        Dim idCRT As Int64
        Dim boolReqAssignment As Boolean

        Dim intAnalyteID As Int64
        Dim strMatrix As String
        Dim strTable As String

        Dim boolAS As Boolean = False

        Call SetProgress()

        If boolFormLoad Then
            Me.panFormLoadProgress.Visible = True
        Else
            Me.lblProgress.Visible = True
        End If

        Me.Visible = True
        Me.Refresh()

        intRow = -1

        strP1 = "Building tables..."
        Me.lblProgress.Text = strP1
        Me.lblProgress.Refresh()
        If boolFormLoad Then
            Me.lblFormLoadProgress.Text = strP1
            Me.lblFormLoadProgress.Refresh()

        End If

        'have to go through tblanalhome to get intstd

        For Count1 = 1 To tblAnalytesHome.Rows.Count

            Try
                strAnal = tblAnalytesHome.Rows(Count1 - 1).Item("ORIGINALANALYTEDESCRIPTION")
                strAnalC = tblAnalytesHome.Rows(Count1 - 1).Item("ANALYTEDESCRIPTION")
                intAnalyteID = NZ(tblAnalytesHome.Rows(Count1 - 1).Item("ANALYTEID"), -1)
                intGroup = tblAnalytesHome.Rows(Count1 - 1).Item("INTGROUP")
                strMatrix = NZ(tblAnalytesHome.Rows(Count1 - 1).Item("MATRIX"), "NA")
                MasterAssayID = tblAnalytesHome.Rows(Count1 - 1).Item("MasterAssayID")
                AnalyteIndex = NZ(tblAnalytesHome.Rows(Count1 - 1).Item("AnalyteIndex"), -1)
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

            strP1 = "Building tables..." & ChrW(10) & ChrW(10)
            strP1 = strP1 & "Evalutating " & Count1 & " of " & dgvA.RowCount & " Analytes:" & ChrW(10) & strAnalC & "..."


            'now go through each table
            For Count2 = 1 To dgvT.RowCount

                idCRT = dgvT("ID_TBLCONFIGREPORTTABLES", Count2 - 1).Value
                idRT = dgvT("ID_TBLREPORTTABLE", Count2 - 1).Value
                int1 = dgvT("BOOLREQUIRESSAMPLEASSIGNMENT", Count2 - 1).Value
                strTable = dgvT("CHARHEADINGTEXT", Count2 - 1).Value

                intConc = dgvT("BOOLSHOWCONC", Count2 - 1).Value
                intArea = dgvT("BOOLSHOWAREA", Count2 - 1).Value
                intISArea = dgvT("BOOLINCLUDEIS", Count2 - 1).Value

                strP2 = "Evalutating " & Count2 & " of " & dgvT.RowCount & " Tables:" & ChrW(10) & strTable
                strP3 = strP1 & ChrW(10) & ChrW(10) & strP2

                If int1 = 0 Then
                    boolRAS = False
                Else
                    boolRAS = True
                End If

                'determine if table requires assigned samples
                Dim rowAS1() As DataRow
                strF1 = "ID_TBLCONFIGREPORTTABLES = " & idCRT
                rowAS1 = tblConfigReportTables.Select(strF1)
                int1 = rowAS1(0).Item("BOOLREQUIRESSAMPLEASSIGNMENT")

                If int1 = 0 Then
                    boolReqAssignment = False
                Else
                    boolReqAssignment = True
                End If

                Try
                    boolAS = False
                    If boolRAS Or boolReqAssignment Then
                        boolAS = True
                    End If

                    Call Fill_tblResults_QCStds(False, strAnalC, AnalyteIndex, MasterAssayID, intRow, intGroup, intAnalyteID, strMatrix, idCRT, idRT, arrResults, intColsR, arrColsR, boolAS, strTable, intConc, intArea, intISArea)

                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1
                End Try

            Next

        Next

        'now filter this to 0 and assign to dgvResults


        Try
            ReDim Preserve arrResults(intColsR - 1, intRow)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        'assign array to tblresults
        Me.tblResults.Clear()
        Me.tblResults.AcceptChanges()
        Me.tblResults.BeginLoadData()

        ' Add the new row to the rows collection.
        Dim arr1()
        ReDim arr1(intColsR - 1)
        Try
            For Count1 = 0 To intRow
                For Count2 = 0 To intColsR - 1
                    arr1(Count2) = arrResults(Count2, Count1)
                Next
                Me.tblResults.LoadDataRow(arr1, False)
                var1 = var1
            Next
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        'do crits
        str1 = "Applying Stats Method..."
        Me.lblProgress.Text = str1
        Me.lblProgress.Refresh()
        If boolFormLoad Then
            str1 = "Applying Stats Method..."
            Me.lblFormLoadProgress.Text = str1
            Me.lblFormLoadProgress.Refresh()
        End If

        Try
            Call FillOutlier()
            'If Me.rbGrubbs.Checked Then
            '    Call FillCritZ(Me.tblResults, intRow, False, "AA")
            'ElseIf Me.rbStdDev.Checked Then
            '    Call FillStdDev(Me.tblResults, intRow, False, "AA")
            'ElseIf Me.rbDixon.Checked Then
            '    Call FillDixon(Me.tblResults, intRow, False, "AA")
            'End If
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        int1 = tblResults.Rows.Count

        strF = "INTGROUP > 0"
        Dim dv As DataView = New DataView(tblResults, strF, "", DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        int1 = dv.Count 'debug
        int1 = int1

        strF = "INTGROUP < 0"
        dv.RowFilter = strF
        dgvR.DataSource = dv

        int1 = dv.Count 'debug
        int1 = int1

        Try
            Call FillSummary(tblResults)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


        '' ''debug
        'make text file
        Dim strPath As String
        strPath = ""
        strPath = "C:\LabIntegrity\Test.csv"
        Dim fs As FileStream = File.Create(strPath)

        Dim strP As String

        ''debug
        'var1 = ""
        'For Count1 = 0 To Me.tblResults.Columns.Count - 1
        '    var2 = Me.tblResults.Columns(Count1).ColumnName
        '    var1 = var1 & "," & var2
        'Next
        'strP = var1 & ChrW(10)


        'For Count2 = 0 To Me.tblResults.Rows.Count - 1
        '    var1 = ""
        '    For Count1 = 0 To Me.tblResults.Columns.Count - 1
        '        var2 = Me.tblResults.Rows(Count2).Item(Count1)
        '        var1 = var1 & "," & var2
        '    Next
        '    strP = strP & var1 & ChrW(10)
        'Next
        'Dim info As Byte() = New UTF8Encoding(True).GetBytes(strP)
        'fs.Write(info, 0, info.Length)
        'fs.Close()
        'fs.Dispose()

        var1 = var1
        ' ''Me.lblProgress.Visible = False

    End Sub


    Sub DoStdDev(tblNC As System.Data.DataTable, strF As String, ByRef dtblR As System.Data.DataTable, strCount As String, boolRCPARatio As Boolean, varIntStd As Object, intConc As Short, intArea As Short, intISArea As Short)

        Dim dtbl As System.Data.DataTable
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim Count4 As Int32
        Dim Count5 As Int32

        Dim dbl1 As Double
        Dim dblCt As Double
        Dim n As Short
        Dim intRows As Short
        Dim var1, var2, var3
        Dim dgv1 As DataGridView
        'Dim intRow As Short
        'Dim strCount As String
        Dim dgvA As DataGridView
        Dim intARow As Short
        'Dim varIntStd

        Dim dblAve As Double
        Dim dblSDP As Double
        Dim dblSDM As Double
        Dim dblStDev As Double

        Dim strAve As String
        Dim strSDP As String
        Dim strSDM As String
        Dim strStDev As String

        Dim arrZ()
        Dim int1 As Short
        'Dim rowsCritZ() As DataRow
        Dim strF1 As String
        Dim strF2 As String
        Dim decCritZ As Decimal
        Dim nS As Short
        Dim num1 As Decimal
        Dim str1 As String
        Dim str2 As String
        Dim numSD As Decimal
        Dim numSDCalc As Decimal
        Dim numNomConc As Decimal

        numSD = CDec(Me.txtStdDev.Text)

        For Count5 = 0 To tblNC.Rows.Count - 1

            numNomConc = tblNC.Rows(Count5).Item("NOMCONC")
            strF1 = strF & " AND NOMCONC = " & numNomConc
            Dim Rows() As DataRow = dtblR.Select(strF1)

            intRows = Rows.Length

            n = 0
            dblCt = 0

            For Count1 = 0 To intRows - 1
                n = n + 1
                dbl1 = NZ(Rows(Count1).Item(strCount), 0)
                dblCt = dblCt + dbl1
            Next Count1

            If n = 0 Then 'exit
                Exit For
            End If
            'fill in Z data
            ReDim arrZ(n)
            'fill arrz
            int1 = 0
            For Count2 = 0 To intRows - 1
                int1 = int1 + 1
                dbl1 = Rows(Count2).Item(strCount)
                arrZ(int1) = dbl1
            Next
            dblAve = dblCt / n
            If n < 2 Then
                dblStDev = 0
            Else
                dblStDev = StdDev(n, arrZ)
            End If


            Dim bool As Boolean
            bool = boolFormLoad
            boolFormLoad = True
            'dblSDP = RoundToDecimalRAFZ(dblAve + (numSD * dblStDev), 1)
            'dblSDM = RoundToDecimalRAFZ(dblAve - (numSD * dblStDev), 1)

            'dblSDP = SigFigOrDec(dblAve + (numSD * dblStDev), LSigFig, False)
            'dblSDM = SigFigOrDec(dblAve - (numSD * dblStDev), LSigFig, False)

            dblSDP = dblAve + (numSD * dblStDev)
            dblSDM = dblAve - (numSD * dblStDev)

            If intConc = -1 Then
                dblSDP = CDec(SigFigOrDec(dblSDP, LSigFig, False))
                dblSDM = CDec(SigFigOrDec(dblSDM, LSigFig, False))
            ElseIf intArea = -1 Or intISArea = -1 Then
                If StrComp(varIntStd, "No", CompareMethod.Text) = 0 Then
                    If boolRCPARatio Then
                        dblSDP = CDec(SigFigAreaRatio(dblSDP, LSigFigAreaRatio, True, False))
                        dblSDM = CDec(SigFigAreaRatio(dblSDM, LSigFigAreaRatio, True, False))
                    Else
                        dblSDP = CDec(SigFigArea(dblSDP, LSigFigArea, True, False))
                        dblSDM = CDec(SigFigArea(dblSDM, LSigFigArea, True, False))
                    End If
                Else
                    dblSDP = CDec(SigFigArea(dblSDP, LSigFigArea, True, False))
                    dblSDM = CDec(SigFigArea(dblSDM, LSigFigArea, True, False))
                End If
            Else
                dblSDP = CDec(SigFigOrDec(dblSDP, LSigFig, False))
                dblSDM = CDec(SigFigOrDec(dblSDM, LSigFig, False))
            End If

            For Count2 = 0 To intRows - 1

                dbl1 = Rows(Count2).Item(strCount)
                Rows(Count2).BeginEdit()
                'num1 = CDec(Format(dblZ, "0.00"))
                Rows(Count2).Item("INTN") = NZ(n, 0)
                If n < 2 Then
                    Rows(Count2).Item("STATS1") = "NA"
                    Rows(Count2).Item("STATS2") = "NA"
                Else

                    If intConc = -1 Then
                        strSDM = SigFigOrDecString(CDec(dblSDM), LSigFig, False)
                        strSDP = SigFigOrDecString(CDec(dblSDP), LSigFig, False)
                        Rows(Count2).Item("STATS1") = strSDM 'SigFigOrDecString(dblSDM, LSigFig, False)
                        Rows(Count2).Item("STATS2") = strSDP 'SigFigOrDecString(dblSDP, LSigFig, False)
                    ElseIf intArea = -1 Or intISArea = -1 Then
                        If StrComp(varIntStd, "No", CompareMethod.Text) = 0 Then
                            If boolRCPARatio Then
                                strSDM = SigFigAreaRatio(dblSDM, LSigFigAreaRatio, False, False) ' Format(RoundToDecimalRAFZ(dblSDM, 0), "0")
                                strSDP = SigFigAreaRatio(dblSDP, LSigFigAreaRatio, False, False) 'Format(RoundToDecimalRAFZ(dblSDP, 0), "0")
                            Else
                                strSDM = SigFigArea(dblSDM, LSigFigArea, False, False) ' Format(RoundToDecimalRAFZ(dblSDM, 0), "0")
                                strSDP = SigFigArea(dblSDP, LSigFigArea, False, False) 'Format(RoundToDecimalRAFZ(dblSDP, 0), "0")
                                'strSDM = Format(RoundToDecimalRAFZ(dblSDM, 0), "0")
                                'strSDP = Format(RoundToDecimalRAFZ(dblSDP, 0), "0")
                            End If
                        Else
                            strSDM = SigFigArea(dblSDM, LSigFigArea, False, False) ' Format(RoundToDecimalRAFZ(dblSDM, 0), "0")
                            strSDP = SigFigArea(dblSDP, LSigFigArea, False, False) 'Format(RoundToDecimalRAFZ(dblSDP, 0), "0")
                            'strSDM = Format(RoundToDecimalRAFZ(dblSDM, 0), "0")
                            'strSDP = Format(RoundToDecimalRAFZ(dblSDP, 0), "0")
                        End If
                        Rows(Count2).Item("STATS1") = strSDM 'Format(dblSDM, "0") ' SigFigOrDecString(dblSDP, LSigFig, False)
                        Rows(Count2).Item("STATS2") = strSDP 'Format(dblSDP, "0") 'SigFigOrDecString(dblSDM, LSigFig, False)
                    Else
                        strSDM = SigFigOrDecString(CDec(dblSDM), LSigFig, False)
                        strSDP = SigFigOrDecString(CDec(dblSDP), LSigFig, False)
                        Rows(Count2).Item("STATS1") = strSDM 'SigFigOrDecString(dblSDM, LSigFig, False)
                        Rows(Count2).Item("STATS2") = strSDP 'SigFigOrDecString(dblSDP, LSigFig, False)
                    End If

                    If CDec(dbl1) < CDec(dblSDM) Or CDec(dbl1) > CDec(dblSDP) Then
                        Rows(Count2).Item("CHAROUTLIER") = "X"
                    End If
                End If

                Rows(Count2).EndEdit()
            Next
            boolFormLoad = bool
            n = 0
            dblCt = 0

        Next Count5

    End Sub


    Sub DoDixon(tblNC As System.Data.DataTable, strF As String, ByRef dtblR As System.Data.DataTable, strCount As String, boolRCPARatio As Boolean)


        'Adapted from W. J. Dixon, "Processing Data for Outliers", Biometrics, Vol 9, No. 1, 74-89, 1953

        '20190226 LEE:
        'Need to allow three confidence intervals: 90, 95, 99


        Dim Count1 As Short
        Dim dbl1 As Double
        Dim dblCt As Double
        Dim n As Short
        Dim intRows As Short
        Dim var1, var2, var3
        Dim dgv1 As DataGridView
        'Dim intRow As Short
        Dim intConc As Short
        Dim intArea As Short
        Dim intISArea As Short
        'Dim strCount As String
        Dim dgvA As DataGridView
        Dim intARow As Short
        'Dim varIntStd
        Dim Count2 As Short
        Dim dblAve As Double
        Dim dblZ As Double
        Dim dblStDev As Double
        Dim arrZ()
        Dim int1 As Short
        Dim rowsCritZ() As DataRow
        'Dim strF As String
        Dim decCritR As Decimal
        Dim nS As Short
        Dim num1 As Decimal
        Dim str1 As String
        Dim str2 As String

        Dim Min1 As Double
        Dim Min2 As Double
        Dim Min3 As Double
        Dim Max1 As Double
        Dim Max2 As Double
        Dim Max3 As Double
        Dim HiDev As Double
        Dim LoDev As Double
        Dim strSuspect As String = "LOW"
        Dim valSuspect As Double
        Dim valNearest1 As Double
        Dim valFurthest1 As Double
        Dim valNearest2 As Double
        Dim valFurthest2 As Double

        Dim Count5 As Int32
        Dim numNomConc As Decimal
        Dim strF1 As String
        Dim strS As String

        Try
            For Count5 = 0 To tblNC.Rows.Count - 1

                numNomConc = tblNC.Rows(Count5).Item("NOMCONC")
                strF1 = strF & " AND NOMCONC = " & numNomConc
                Dim Rows() As DataRow = dtblR.Select(strF)
                intRows = Rows.Length

                n = 0
                dblCt = 0

                For Count1 = 0 To intRows - 1
                    n = n + 1
                    dbl1 = NZ(Rows(Count1).Item(strCount), 0)
                    dblCt = dblCt + dbl1
                Next

                'fill in Z data
                ReDim arrZ(n)
                'fill arrz
                int1 = 0
                For Count2 = 0 To intRows - 1
                    int1 = int1 + 1
                    dbl1 = NZ(Rows(Count2).Item(strCount), 0)
                    arrZ(int1) = dbl1
                Next

                'find Min1
                Min1 = 999999999999
                For Count2 = 1 To int1
                    dbl1 = arrZ(Count2)
                    If dbl1 < Min1 Then
                        Min1 = dbl1
                    End If
                Next

                'find Min2
                Min2 = 999999999999
                For Count2 = 1 To int1
                    dbl1 = arrZ(Count2)
                    If dbl1 < Min2 And dbl1 <> Min1 Then
                        Min2 = dbl1
                    End If
                Next

                'find Min3
                Min3 = 999999999999
                For Count2 = 1 To int1
                    dbl1 = arrZ(Count2)
                    If dbl1 < Min3 And dbl1 <> Min1 And dbl1 <> Min2 Then
                        Min3 = dbl1
                    End If
                Next

                'find Max1
                Max1 = -999999999999
                For Count2 = 1 To int1
                    dbl1 = arrZ(Count2)
                    If dbl1 > Max1 Then
                        Max1 = dbl1
                    End If
                Next

                'find Max2
                Max2 = -999999999999
                For Count2 = 1 To int1
                    dbl1 = arrZ(Count2)
                    If dbl1 > Max2 And dbl1 <> Max1 Then
                        Max2 = dbl1
                    End If
                Next

                'find Max3
                Max3 = 9 - 99999999999
                For Count2 = 1 To int1
                    dbl1 = arrZ(Count2)
                    If dbl1 > Max3 And dbl1 <> Max1 And dbl1 <> Max2 Then
                        Max3 = dbl1
                    End If
                Next

                dblAve = dblCt / n

                If n < 2 Then
                    dblStDev = 0
                Else
                    dblStDev = StdDev(n, arrZ)
                End If

                HiDev = Max1 - dblAve
                LoDev = dblAve - Min1

                If HiDev > LoDev Then
                    strSuspect = "HIGH"
                ElseIf LoDev > HiDev Then
                    strSuspect = "LOW"
                Else
                    strSuspect = "SAME"
                End If

                Select Case strSuspect
                    Case "HIGH"
                        valSuspect = Max1
                        valNearest1 = Max2
                        valFurthest1 = Min1
                        valNearest2 = Max3
                        valFurthest2 = Min2
                    Case "LOW"
                        valSuspect = Min1
                        valNearest1 = Min2
                        valFurthest1 = Max1
                        valNearest2 = Min3
                        valFurthest2 = Max2
                    Case Else
                        valSuspect = -1
                End Select

                'calculate R
                'Q = gap/range = ([n+1]-[n])/([range])
                Dim sampleR As Single
                If n >= 3 And n <= 7 Then
                    sampleR = RoundToDecimalRAFZ((valSuspect - valNearest1) / (valSuspect - valFurthest1), 3)
                ElseIf n >= 8 And n <= 10 Then
                    sampleR = RoundToDecimalRAFZ((valSuspect - valNearest1) / (valSuspect - valFurthest2), 3)
                ElseIf n >= 11 And n <= 13 Then
                    sampleR = RoundToDecimalRAFZ((valSuspect - valNearest2) / (valSuspect - valFurthest2), 3)
                Else
                    sampleR = 0
                End If


                'Critical values for R	
                'Set size	90% Confidence Interval
                '3	0.886
                '4	0.679
                '5	0.557
                '6	0.482
                '7	0.434
                '8	0.479
                '9	0.441
                '10	0.409
                '11	0.517
                '12	0.490
                '13	0.467

                '20190226 LEE
                'https://sebastianraschka.com/Articles/2014_dixon_test.html
                'Critical Values
                'N Q90% Q95% Q99%
                '3 0.941 0.97 0.994 
                '4 0.765 0.829 0.926 
                '5 0.642 0.71 0.821 
                '6 0.56 0.625 0.74 
                '7 0.507 0.568 0.68 
                '8 0.468 0.526 0.634 
                '9 0.437 0.493 0.598 
                '10 0.412 0.466 0.568 
                '11 0.392 0.444 0.542 
                '12 0.376 0.426 0.522 
                '13 0.361 0.41 0.503 

                decCritR = 0
                Dim intCL As Short = Me.cbxCL.SelectedIndex
                Select Case intCL
                    Case 0 '90% CL
                        Select Case n
                            Case 3
                                decCritR = 0.941
                            Case 4
                                decCritR = 0.765
                            Case 5
                                decCritR = 0.642
                            Case 6
                                decCritR = 0.56
                            Case 7
                                decCritR = 0.507
                            Case 8
                                decCritR = 0.468
                            Case 9
                                decCritR = 0.437
                            Case 10
                                decCritR = 0.412
                            Case 11
                                decCritR = 0.392
                            Case 12
                                decCritR = 0.376
                            Case 13
                                decCritR = 0.361

                        End Select
                    Case 1 '95% CL
                        Select Case n
                            Case 3
                                decCritR = 0.97
                            Case 4
                                decCritR = 0.829
                            Case 5
                                decCritR = 0.71
                            Case 6
                                decCritR = 0.625
                            Case 7
                                decCritR = 0.568
                            Case 8
                                decCritR = 0.526
                            Case 9
                                decCritR = 0.493
                            Case 10
                                decCritR = 0.466
                            Case 11
                                decCritR = 0.444
                            Case 12
                                decCritR = 0.426
                            Case 13
                                decCritR = 0.41

                        End Select
                    Case 2 '99% CL
                        Select Case n
                            Case 3
                                decCritR = 0.994
                            Case 4
                                decCritR = 0.926
                            Case 5
                                decCritR = 0.821
                            Case 6
                                decCritR = 0.74
                            Case 7
                                decCritR = 0.68
                            Case 8
                                decCritR = 0.634
                            Case 9
                                decCritR = 0.598
                            Case 10
                                decCritR = 0.568
                            Case 11
                                decCritR = 0.542
                            Case 12
                                decCritR = 0.522
                            Case 13
                                decCritR = 0.503

                        End Select
                End Select

                'Select Case n
                '    Case 3
                '        decCritR = 0.886
                '    Case 4
                '        decCritR = 0.679
                '    Case 5
                '        decCritR = 0.557
                '    Case 6
                '        decCritR = 0.482
                '    Case 7
                '        decCritR = 0.434
                '    Case 8
                '        decCritR = 0.479
                '    Case 9
                '        decCritR = 0.441
                '    Case 10
                '        decCritR = 0.409
                '    Case 11
                '        decCritR = 0.517
                '    Case 12
                '        decCritR = 0.49
                '    Case 13
                '        decCritR = 0.467

                'End Select

                Dim bool As Boolean
                Dim charX As String
                bool = boolFormLoad
                boolFormLoad = True
                For Count2 = 0 To intRows - 1

                    dbl1 = CDec(NZ(Rows(Count2).Item(strCount), 0))

                    Select Case strSuspect
                        Case "HIGH"

                        Case "LOW"

                        Case Else

                    End Select

                    var1 = Format(RoundToDecimalRAFZ(sampleR, 3), "0.000") ' Format(SigFigOrDecString(sampleR, LSigFig, False), GetRegrDecStr(LSigFig)) ' 
                    var2 = Format(RoundToDecimalRAFZ(decCritR, 3), "0.000")

                    If dblStDev = 0 Then
                        var1 = 1 'debugging
                    Else
                        If CDec(decCritR) = 0 Or (n < 3 Or n > 13) Then
                            num1 = 0
                            var1 = "NA"
                            'decCritR = "NA"
                            var2 = "NA"
                            charX = ""
                        Else

                            If CDec(dbl1) = CDec(valSuspect) Then
                                'var1 = Format(RoundToDecimalRAFZ(sampleR, 3), "0.000") ' Format(SigFigOrDecString(sampleR, LSigFig, False), GetRegrDecStr(LSigFig)) ' 
                                'var2 = Format(RoundToDecimalRAFZ(decCritR, 3), "0.000")
                                If CDec(sampleR) > CDec(decCritR) Then
                                    charX = "X"
                                Else
                                    charX = ""
                                    var1 = "NA"
                                    'decCritR = "NA"
                                    var2 = "NA"
                                End If
                            Else
                                num1 = 0
                                var1 = "NA"
                                'decCritR = "NA"
                                var2 = "NA"
                                charX = ""
                            End If

                            'dblZ = (Math.Abs(dbl1 - dblAve)) / dblStDev
                            'Rows(Count2).BeginEdit()
                            'num1 = CDec(Format(dblZ, "0.00"))
                            'var1 = num1
                            'var2 = decCritR
                        End If

                        Rows(Count2).BeginEdit()
                        Rows(Count2).Item("STATS1") = var1
                        Rows(Count2).Item("STATS2") = var2
                        Rows(Count2).Item("INTN") = NZ(n, 0)
                        Rows(Count2).Item("CHAROUTLIER") = charX
                        'If num1 > decCritR Then
                        '    Rows(Count2).Item("CHAROUTLIER") = "X"
                        'End If
                        Rows(Count2).EndEdit()
                    End If
                Next Count2

                boolFormLoad = bool
                n = 0
                dblCt = 0

            Next Count5
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try






    End Sub

    Sub DoCritZ(tblNC As System.Data.DataTable, strF As String, ByRef dtblR As System.Data.DataTable, strCount As String, boolRCPARatio As Boolean)

        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count5 As Int32
        Dim numNomConc As Decimal
        Dim strF1 As String
        Dim strF2 As String
        Dim intRows As Int32
        Dim dblCt As Double
        Dim n As Int32
        Dim dbl1 As Double
        Dim arrZ()
        Dim int1 As Int32
        Dim int2 As Int32
        Dim dblAve As Double
        Dim dblStDev As Double
        Dim decCritZ As Double
        Dim nS As Int16
        Dim rowsCritZ() As DataRow
        Dim var1, var2, var3, var4
        Dim num1 As Double
        Dim num2 As Double
        Dim dblZ As Double

        For Count5 = 0 To tblNC.Rows.Count - 1

            numNomConc = tblNC.Rows(Count5).Item("NOMCONC")
            strF1 = strF & " AND NOMCONC = " & numNomConc
            Dim Rows() As DataRow = dtblR.Select(strF1)

            intRows = Rows.Length

            n = 0
            dblCt = 0


            For Count1 = 0 To intRows - 1

                Rows(Count1).BeginEdit()
                n = n + 1
                dbl1 = NZ(Rows(Count1).Item(strCount), 0)
                dblCt = dblCt + dbl1
                Rows(Count1).EndEdit()
            Next

            ReDim arrZ(n)
            'fill arrz
            int1 = 0
            For Count2 = 0 To intRows - 1
                Rows(Count2).BeginEdit()
                int1 = int1 + 1
                dbl1 = NZ(Rows(Count2).Item(strCount), 0)
                arrZ(int1) = dbl1
                Rows(Count2).EndEdit()
            Next
            dblAve = dblCt / n
            If n < 2 Then
                dblStDev = 0
            Else
                dblStDev = StdDev(n, arrZ)
            End If

            If n < 3 Then 'ignore
                decCritZ = 0
            Else
                'get critz
                If n <= 40 Then
                    nS = n
                ElseIf n <= 50 Then
                    nS = 50
                ElseIf n <= 60 Then
                    nS = 60
                ElseIf n <= 70 Then
                    nS = 70
                ElseIf n <= 80 Then
                    nS = 80
                ElseIf n <= 90 Then
                    nS = 90
                ElseIf n <= 100 Then
                    nS = 100
                ElseIf n <= 110 Then
                    nS = 110
                ElseIf n <= 120 Then
                    nS = 120
                ElseIf n <= 130 Then
                    nS = 130
                ElseIf n <= 140 Then
                    nS = 140
                Else
                    nS = 140
                End If
                strF2 = "N = " & nS
                rowsCritZ = Me.tblCritZ.Select(strF2)
                decCritZ = rowsCritZ(0).Item("CriticalZ")
            End If

            Dim bool As Boolean
            bool = boolFormLoad
            boolFormLoad = True
            For Count2 = 0 To intRows - 1

                dbl1 = NZ(Rows(Count2).Item(strCount), 0)

                If dblStDev = 0 Then
                    var1 = 1 'debugging
                Else
                    If decCritZ = 0 Or n < 2 Then
                        num1 = 0
                        var1 = "NA"
                        'decCritZ = "NA"
                        var2 = "NA"
                    Else
                        dblZ = CDec((Math.Abs(dbl1 - dblAve)) / dblStDev)
                        num1 = CDec(RoundToDecimalRAFZ(dblZ, 2))
                        var1 = Format(RoundToDecimalRAFZ(dblZ, 3), "0.00") ' num1
                        var2 = CDec(decCritZ)
                    End If

                    Rows(Count2).BeginEdit()

                    Rows(Count2).Item("STATS1") = var1
                    Rows(Count2).Item("STATS2") = var2
                    Rows(Count2).Item("INTN") = NZ(n, 0)
                    If CDec(num1) > CDec(decCritZ) Then
                        Rows(Count2).Item("CHAROUTLIER") = "X"
                    End If

                    Rows(Count2).EndEdit()
                End If
            Next
            boolFormLoad = bool
            n = 0
            dblCt = 0

        Next Count5

    End Sub

    Sub FillOutlier()

        Dim dtbl As System.Data.DataTable
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim Count3 As Int32
        Dim Count4 As Int32
        Dim Count5 As Int32

        Dim dbl1 As Double
        Dim dblCt As Double
        Dim n As Short
        Dim intRows As Short
        Dim var1, var2, var3
        Dim dgv1 As DataGridView
        'Dim intRow As Short
        Dim intConc As Short
        Dim intArea As Short
        Dim intISArea As Short
        Dim strCount As String
        Dim dgvA As DataGridView
        Dim dgvT As DataGridView
        Dim intARow As Short
        'Dim varIntStd
        Dim dblAve As Double
        Dim dblZ As Double
        Dim dblStDev As Double
        Dim arrZ()
        Dim int1 As Short
        Dim rowsCritZ() As DataRow

        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim strF4 As String
        Dim strF5 As String

        Dim decCritZ As Decimal
        Dim nS As Short
        Dim num1 As Decimal
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String

        Dim strS1 As String
        Dim strS2 As String
        Dim strS3 As String
        Dim strS4 As String

        Dim boolAS As Boolean

        dgvA = Me.dgvAnalytes
        dgvT = Me.dgvTables

        Dim intAnalyteID As Int64
        Dim idT As Int64
        Dim idCT As Int64
        Dim intGroup As Short
        Dim strMatrix As String
        Dim numNomConc As Decimal
        Dim varIntStd

        Dim strP1 As String
        Dim strP2 As String
        Dim strP3 As String
        Dim strTable As String
        Dim strAnalC As String
        Dim boolRAS As Boolean
        Dim boolReqAssignment As Boolean

        Dim dtblR As System.Data.DataTable = Me.tblResults

        'first clear all outliers in dtblr
        Try
            For Count1 = 0 To dtblR.Rows.Count - 1
                dtblR.Rows(Count1).BeginEdit()
                dtblR.Rows(Count1).Item("CHAROUTLIER") = Nothing
                dtblR.Rows(Count1).Item("BOOLOUTLIER") = False
                dtblR.Rows(Count1).EndEdit()
            Next
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        dtblR.AcceptChanges()

        Dim boolIsIS As Boolean

        Dim tblA As System.Data.DataTable = tblAnalytesHome

        For Count3 = 0 To tblA.Rows.Count - 1

            intAnalyteID = NZ(tblA.Rows(Count3).Item("ANALYTEID"), -1)
            If intAnalyteID = -1 Then
                boolIsIS = True
            Else
                boolIsIS = False
            End If
            intGroup = tblA.Rows(Count3).Item("INTGROUP")
            varIntStd = tblA.Rows(Count3).Item("IsIntStd")
            strAnalC = tblA.Rows(Count3).Item("ANALYTEDESCRIPTION")

            'strP1 = "Applying Stats Method..." & ChrW(10) & ChrW(10)
            'strP1 = strP1 & "Evalutating " & Count3 + 1 & " of " & dgvA.RowCount & " Analytes:" & ChrW(10) & strAnalC & "..."

            For Count4 = 0 To dgvT.Rows.Count - 1

                idT = dgvT("ID_TBLREPORTTABLE", Count4).Value
                idCT = dgvT("ID_TBLCONFIGREPORTTABLES", Count4).Value
                strTable = dgvT("CHARHEADINGTEXT", Count4).Value

                intConc = dgvT("BOOLSHOWCONC", Count4).Value
                intArea = dgvT("BOOLSHOWAREA", Count4).Value
                intISArea = dgvT("BOOLINCLUDEIS", Count4).Value


                Dim boolRCConc As Boolean
                Dim boolRCPA As Boolean
                Dim boolRCPARatio As Boolean
                Dim boolIncludeISTbl As Boolean

                Call ReturnColumnTypes(boolRCConc, boolRCPA, boolRCPARatio, boolIncludeISTbl, idT) 'this will return bools

                strCount = "CONCENTRATION"

                If intConc = -1 Then
                    strCount = "CONCENTRATION"
                ElseIf intArea = -1 Or intISArea = -1 Then
                    If StrComp(varIntStd, "No", CompareMethod.Text) = 0 Then
                        If boolRCPARatio Then
                            strCount = "Concentration"
                        Else
                            strCount = "ANALYTEAREA"
                        End If
                    Else
                        strCount = "INTERNALSTANDARDAREA"
                    End If
                End If

                'has Term 1's
                Dim rowsT1() As DataRow
                Dim tblT1 As System.Data.DataTable
                Dim intT1 As Short = 0
                Dim intT1a As Short = 1
                Dim boolT1 As Boolean = False
                Dim CountT1 As Short

                Dim rowsT2() As DataRow
                Dim tblT2 As System.Data.DataTable
                Dim intT2 As Short = 0
                Dim intT2a As Short = 1
                Dim boolT2 As Boolean = False
                Dim CountT2 As Short

                Erase rowsT1
                Dim strFT1 As String
                Dim strFT2 As String
                Dim dvT1 As DataView
                Dim dvT2 As DataView
                Dim strS As String
                strFT1 = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND NUMHELPERNUMBER = 1"
                Select Case idCT
                    Case 13, 14, 15, 17, 22, 23, 29, 32, 34, 35
                        strS = "ID_TBLASSIGNEDSAMPLESHELPER ASC"
                        dvT1 = New DataView(tblAssignedSamplesHelper, strFT1, strS1, DataViewRowState.CurrentRows)
                        tblT1 = dvT1.ToTable("a", True, "CHARHELPER")
                        intT1 = tblT1.Rows.Count
                        If boolIsIS Then
                            var1 = var1
                        End If
                End Select
                If intT1 > 0 Then
                    boolT1 = True
                    intT1a = intT1
                End If

                'has Term 2's
                strFT2 = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND NUMHELPERNUMBER = 2"
                Select Case idCT
                    Case 29, 32, 34
                        strS = "ID_TBLASSIGNEDSAMPLESHELPER ASC"
                        dvT2 = New DataView(tblAssignedSamplesHelper, strFT2, strS1, DataViewRowState.CurrentRows)
                        tblT2 = dvT2.ToTable("a", True, "CHARHELPER")
                        intT2 = tblT2.Rows.Count
                End Select
                If intT2 > 0 Then
                    boolT2 = True
                    intT2a = intT2
                End If

                'NOTE: ANOVA table must be done by NomConc and RunID
                'may have more than one Mid, so need to evaluate nomconc
                Dim boolRunID As Boolean = False
                Dim CountRunID As Short
                Dim dvRunID As DataView
                Dim tblRunID As System.Data.DataTable
                Dim strFRunID As String
                Dim intRID As Short = 0
                Dim intRIDa As Short = 1
                Dim strHelper2 As String

                Select Case idCT
                    Case 11
                        boolRunID = True
                        'get unique runids in tblassignedsamples
                        'strFRunID = "ID_TBLCONFIGREPORTTABLES = " & idCT & " AND INTGROUP = " & intGroup
                        strFRunID = "ID_TBLREPORTTABLE = " & idT & " AND INTGROUP = " & intGroup
                        dvRunID = New DataView(tblAssignedSamples, strFRunID, "RUNID ASC, NOMCONC ASC", DataViewRowState.CurrentRows)
                        tblRunID = dvRunID.ToTable("a", True, "RUNID", "NOMCONC")
                        intRID = tblRunID.Rows.Count
                        If intRID = 0 Then
                            boolRunID = False
                            intRIDa = 0
                        End If
                End Select
                If intRID > 0 Then
                    boolRunID = True
                    intRIDa = intRID
                End If

                Select Case idCT
                    Case 13, 14, 15, 17, 22, 23
                        strS3 = "CHARHELPER1 ASC, NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                    Case Else
                        strS3 = "NOMCONC ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
                End Select

                For CountT2 = 1 To intT2a

                    strF = ""
                    strF1 = ""
                    strHelper2 = ""
                    If boolT2 Then
                        str1 = tblT2.Rows(CountT2 - 1).Item("CHARHELPER")
                        strF = "CHARHELPER2 = '" & str1 & "'"
                        strF1 = strF
                        strHelper2 = str1
                    End If

                    'Note: At this point, if bootT2=true
                    'CHARHELPER1 has been changed to CHARHELPER2 & CHRW(10) & CHARHELPER1
                    For CountT1 = 1 To intT1a

                        strF = ""
                        strF2 = strF1
                        If boolT1 Then
                            str1 = tblT1.Rows(CountT1 - 1).Item("CHARHELPER")
                            If boolT2 Then
                                str1 = strHelper2 & " " & str1
                            End If
                            If Len(strF1) = 0 Then
                                strF = "CHARHELPER1 = '" & str1 & "'"
                            Else
                                strF = strF1 & " AND CHARHELPER1 = '" & str1 & "'"
                            End If
                            strF2 = strF
                        End If

                        For CountRunID = 1 To intRIDa

                            strF = ""
                            strF3 = strF2
                            If boolRunID Then
                                str1 = tblRunID.Rows(CountRunID - 1).Item("RUNID").ToString
                                str2 = tblRunID.Rows(CountRunID - 1).Item("NOMCONC").ToString
                                If Len(strF2) = 0 Then
                                    strF = "RUNID = " & str1 & " AND NOMCONC = " & str2
                                Else
                                    strF = strF2 & " AND RUNID = " & str1 & " AND NOMCONC = " & str2
                                End If
                                strF3 = strF
                            End If


                            '20160304 LEE:
                            'new filter for groups
                            strF = ""
                            strF4 = strF3
                            If Len(strF3) = 0 Then
                                strF = "INTGROUP = " & intGroup & " AND ID_TBLREPORTTABLE = " & idT
                            Else
                                strF = strF3 & " AND INTGROUP = " & intGroup & " AND ID_TBLREPORTTABLE = " & idT
                            End If
                            strF4 = strF
                            strF5 = strF4 & " AND NOMCONC IS NOT NULL"

                            'strF = "INTGROUP = " & intGroup & " AND ANALYTEID = " & intAnalyteID & " AND ID_TBLREPORTTABLE = " & idT & " AND ID_TBLCONFIGREPORTTABLES = " & idCT
                            'strF = "INTGROUP = " & intGroup & " AND ID_TBLREPORTTABLE = " & idT & " AND ID_TBLCONFIGREPORTTABLES = " & idCT & " AND NOMCONC IS NOT NULL"

                            'now need to find unique nomConcs
                            Dim dvNC As DataView
                            Try
                                dvNC = New DataView(dtblR, strF5, "", DataViewRowState.CurrentRows)
                            Catch ex As Exception
                                var1 = ex.Message
                                var1 = var1
                            End Try

                            int1 = dvNC.Count 'debug

                            If int1 = 0 Then
                            Else
                                Dim tblNC As System.Data.DataTable = dvNC.ToTable("a", True, "NOMCONC")

                                If Me.rbGrubbs.Checked Then
                                    Call DoCritZ(tblNC, strF5, dtblR, strCount, boolRCPARatio)
                                ElseIf Me.rbStdDev.Checked Then
                                    'Call FillStdDev(Me.tblResults, intRow, False, "AA")
                                    Call DoStdDev(tblNC, strF5, dtblR, strCount, boolRCPARatio, varIntStd, intConc, intArea, intISArea)
                                ElseIf Me.rbDixon.Checked Then
                                    'Call FillDixon(Me.tblResults, intRow, False, "AA")
                                    Call DoDixon(tblNC, strF5, dtblR, strCount, boolRCPARatio)
                                End If
                            End If

                        Next CountRunID

                    Next CountT1

                Next CountT2

            Next Count4

        Next Count3

    End Sub

    Sub FillSummary(ByVal dtblR As System.Data.DataTable)

        'dtblr = Me.tblResults

        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim dtbl As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim var1, var2

        dgv = Me.dgvSummary
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'dtbl = Me.tblResults
        dtbl = dtblR 'either tblResults or tblResultsTemp

        var1 = dtblR.Rows.Count 'debug

        Dim intGroup As Short
        Dim dgvA As DataGridView = Me.dgvAnalytes
        Dim intRow As Short
        intGroup = dgvA("INTGROUP", dgvA.CurrentRow.Index).Value

        strF = "CHAROUTLIER = 'X' AND INTGROUP = " & intGroup
        'strS = "NOMCONC ASC, CHARHELPER1 ASC, CHARHELPER2 ASC,  RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
        strS = "NOMCONC ASC, CHARHELPER1 ASC, CHARHELPER2 ASC,  RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
        strS = "RUNID ASC, CHARHELPER1 ASC, CHARHELPER2 ASC,  RUNID ASC, RUNSAMPLEORDERNUMBER ASC"


        Dim dvR As DataView = Me.dgvResults.DataSource
        strS = dvR.Sort

        Dim dv As System.Data.DataView = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        dgv.SuspendLayout()
        dgv.DataSource = dv

        'dgv1 = Me.dgvResults

        'boolSummary = False
        'If boolSummary Then
        'Else
        '    For Count1 = 0 To dgv1.ColumnCount - 1
        '        dgv.Columns(Count1).HeaderText = dgv1.Columns(Count1).HeaderText
        '        dgv.Columns(Count1).Visible = dgv1.Columns(Count1).Visible
        '        dgv.Columns(Count1).SortMode = dgv1.Columns(Count1).SortMode
        '        dgv.Columns(Count1).DefaultCellStyle.Alignment = dgv1.Columns(Count1).DefaultCellStyle.Alignment
        '    Next

        'End If

        dgv.RowHeadersWidth = 25

        Call SetColWidths()

        dgv.AutoResizeRows()
        dgv.AutoResizeColumns()
        dgv.ResumeLayout()

        If boolSummary = False Then
            boolSummary = True
        End If

        'select row
        Try
            Call SelectTableRow()
        Catch ex As Exception

        End Try


    End Sub

    Sub SyncSummaryColumns()

        'this configures dgvSummary

        Dim dgv1 As DataGridView
        Dim dgv As DataGridView
        Dim Count1 As Short

        dgv = Me.dgvSummary
        dgv1 = Me.dgvResults

        For Count1 = 0 To dgv1.ColumnCount - 1

            Try
                dgv.Columns(Count1).HeaderText = dgv1.Columns(Count1).HeaderText
                dgv.Columns(Count1).Visible = dgv1.Columns(Count1).Visible
                dgv.Columns(Count1).SortMode = dgv1.Columns(Count1).SortMode
                dgv.Columns(Count1).DefaultCellStyle.Alignment = dgv1.Columns(Count1).DefaultCellStyle.Alignment
                dgv.Columns(Count1).Width = dgv1.Columns(Count1).Width
                dgv.Columns(Count1).ToolTipText = dgv1.Columns(Count1).ToolTipText
            Catch ex As Exception

            End Try
        Next

        'do not show intN
        dgv.Columns("INTN").Visible = False
        dgv.AutoResizeColumns()

        dgv1.AutoResizeRows()

    End Sub

    Sub Initialize_tblResults()

        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim var1

        If tblResults.Columns.Count > 0 Then
            Exit Sub
        End If

        tbl = Me.tblResults
        tbl1 = Me.tblAllSummary
        tbl2 = Me.tblResultsTemp

        str1 = ""
        str2 = ""
        var1 = ""

        id_tblResults = 0

        For Count1 = -5 To 18

            Select Case Count1
                'Case -5
                '    str1 = "intConc"
                '    var1 = System.Type.GetType("System.Int16")
                '    str2 = "intConc"
                'Case -4
                '    str1 = "intArea"
                '    var1 = System.Type.GetType("System.Int16")
                '    str2 = "intArea"
                Case -5
                    str1 = "INTNUMBER"
                    var1 = System.Type.GetType("System.Int16")
                    str2 = "#"
                Case -4
                    str1 = "ID_TBLRESULTS"
                    var1 = System.Type.GetType("System.Int64")
                    str2 = "ID_TBLRESULTS"
                Case -3
                    str1 = "ID_TBLCONFIGREPORTTABLES"
                    var1 = System.Type.GetType("System.Int64")
                    str2 = "ID_TBLCONFIGREPORTTABLES"
                Case -2
                    str1 = "ID_TBLREPORTTABLE"
                    var1 = System.Type.GetType("System.Int64")
                    str2 = "ID_TBLREPORTTABLE"
                Case -1
                    str1 = "CHARHEADINGTEXT"
                    var1 = System.Type.GetType("System.String")
                    str2 = "Table"
                Case 0
                    str1 = "ANALYTEDESCRIPTION"
                    var1 = System.Type.GetType("System.String")
                    str2 = "Analyte"
                Case 1
                    str1 = "RUNID"
                    var1 = System.Type.GetType("System.Int16")
                    str2 = "Run ID"
                Case 2
                    'str1 = "RUNSAMPLESEQUENCENUMBER"
                    'var1 = System.Type.GetType("System.Int16")
                    'str2 = "Seq#"

                    str1 = "RUNSAMPLEORDERNUMBER"
                    var1 = System.Type.GetType("System.Int16")
                    str2 = "Seq#"
                Case 3
                    str1 = "NOMCONC"
                    var1 = System.Type.GetType("System.Decimal")
                    str2 = "ID1"
                Case 4
                    str1 = "CHARHELPER1"
                    var1 = System.Type.GetType("System.String")
                    str2 = "ID2"
                Case 5
                    str1 = "CHARHELPER2"
                    var1 = System.Type.GetType("System.String")
                    str2 = "ID3"
                Case 6
                    str1 = "CONCENTRATION"
                    var1 = System.Type.GetType("System.Decimal")
                    str2 = "Conc"
                Case 7
                    str1 = "ANALYTEAREA"
                    var1 = System.Type.GetType("System.Double")
                    str2 = "ANALYTEAREA"
                Case 8
                    str1 = "INTERNALSTANDARDAREA"
                    var1 = System.Type.GetType("System.Double")
                    str2 = "INTERNALSTANDARDAREA"
                Case 9
                    str1 = "STATS1"
                    'var1 = System.Type.GetType("System.Decimal")
                    var1 = System.Type.GetType("System.String")
                    str2 = "STATS1"
                Case 10
                    str1 = "STATS2"
                    'var1 = System.Type.GetType("System.Decimal")
                    var1 = System.Type.GetType("System.String")
                    str2 = "STATS2"
                Case 11
                    str1 = "BOOLOUTLIER"
                    var1 = System.Type.GetType("System.Byte")
                    str2 = "Is Outlier"
                Case 12
                    str1 = "CHAROUTLIER"
                    var1 = System.Type.GetType("System.String")
                    str2 = "Is Outlier"
                Case 13
                    str1 = "ID_TBLASSIGNEDSAMPLES"
                    var1 = System.Type.GetType("System.Int64")
                    str2 = "ID_TBLASSIGNEDSAMPLES"
                Case 14
                    str1 = "BOOLSHOWCONC"
                    var1 = System.Type.GetType("System.Byte")
                    str2 = "BOOLSHOWCONC"
                Case 15
                    str1 = "BOOLSHOWAREA"
                    var1 = System.Type.GetType("System.Byte")
                    str2 = "BOOLSHOWAREA"
                Case 16
                    str1 = "BOOLINCLUDEIS"
                    var1 = System.Type.GetType("System.Byte")
                    str2 = "BOOLINCLUDEIS"
                Case 17
                    str1 = "INTGROUP"
                    var1 = System.Type.GetType("System.Int16")
                    str2 = "INTGROUP"

                Case 18
                    str1 = "INTN"
                    var1 = System.Type.GetType("System.Int16")
                    str2 = "n"

            End Select

            Dim col As New DataColumn
            col.ColumnName = str1
            col.Caption = str2
            col.DataType = var1
            col.AllowDBNull = True
            tbl.Columns.Add(col)

            Dim col4 As New DataColumn
            col4.ColumnName = str1
            col4.Caption = str2
            col4.DataType = var1
            col4.AllowDBNull = True
            tbl1.Columns.Add(col4)

            Dim col5 As New DataColumn
            col5.ColumnName = str1
            col5.Caption = str2
            col5.DataType = var1
            col5.AllowDBNull = True
            tbl2.Columns.Add(col5)

        Next

    End Sub

    Sub InitializedgvAnalytes(boolDV As Boolean)

        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim strS As String
        Dim strF As String
        Dim boolAA As Boolean

        'dgv.AllowUserToResizeRows = True
        'dgv.AllowUserToResizeColumns = True

        dgv = Me.dgvAnalytes
        dgv.AllowUserToResizeRows = True
        dgv.AllowUserToResizeColumns = True
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgv.RowHeadersWidth = 25

        strS = "IsIntStd ASC, AnalyteDescription ASC"

        'If boolViewOnly Then
        '    strF = "IsIntStd = 'No'"
        'Else
        '    strF = "IsIntStd = 'No' or IsIntStd = 'Yes'"
        'End If

        strF = "IsIntStd = 'No' or IsIntStd = 'Yes'"

        Dim dv1 As System.Data.DataView = New DataView(Me.tblAnalytes)
        dv1.RowFilter = strF
        dv1.Sort = strS

        dv1.AllowDelete = False
        dv1.AllowEdit = False

        boolAA = boolCont
        boolCont = False
        If boolDV Then
            dgv.DataSource = dv1
        End If

        'dgv.DataSource = dv1
        boolCont = boolAA

        int1 = dgv.RowCount

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        dgv.Columns.Item("AnalyteDescription").Visible = True
        dgv.Columns.Item("AnalyteDescription").HeaderText = "Analyte"
        'dgv.Columns.Item("AnalyteID").Visible = True
        'dgv.Columns.Item("AnalyteID").HeaderText = "ID"
        'dgv.Columns.Item("AnalyteIndex").Visible = True
        'dgv.Columns.Item("AnalyteIndex").HeaderText = "Index"
        'dgv.Columns.Item("MasterAssayID").Visible = True
        'dgv.Columns.Item("MasterAssayID").HeaderText = "MAssayID"
        dgv.Columns.Item("INTGROUP").Visible = True
        dgv.Columns.Item("INTGROUP").HeaderText = "Group"
        Dim wd1, wd2
        wd1 = dgv.RowHeadersWidth
        wd2 = dgv.Width - wd1
        'dgv.Columns.item("AnalyteDescription").MinimumWidth = wd2 * 0.95
        dgv.AutoResizeRows()
        dgv.AutoResizeColumns()

    End Sub

    Sub FilldgvTables()

        Dim dv As System.Data.DataView
        'Dim dv1 as system.data.dataview
        Dim dgv As DataGridView
        Dim strF As String
        Dim strS As String
        Dim intRows As Short
        Dim intRows_frmh As Short
        Dim intRow As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim rows() As DataRow
        Dim var1, var2, var3

        '*****
        'Dim dtbl1 as System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF1 As String
        Dim strF2 As String
        Dim strF3 As String
        Dim strS1 As String
        Dim strS2 As String
        Dim int2 As Int64
        Dim boolSamples As Short
        Dim boolQC As Short
        Dim boolCS As Short
        Dim intCt As Short

        Dim tblOut As New System.Data.DataTable

        Dim dvA As System.Data.DataView = frmH.dgvReportTableConfiguration.DataSource
        Dim dtbl1 As System.Data.DataTable = dvA.ToTable


        'dtbl1 = tblReportTable
        dtbl2 = tblConfigReportTables

        strF1 = "ID_TBLSTUDIES = " & id_tblStudies
        strS1 = "INTORDER ASC"

        rows1 = dtbl1.Select(strF1, strS1)
        intRows = rows1.Length

        intRows = dtbl1.Rows.Count

        strF3 = strF1 & " AND ("
        intCt = 0
        For Count1 = 0 To intRows - 1
            int1 = dtbl1.Rows(Count1).Item("ID_TBLCONFIGREPORTTABLES")

            Select Case int1
                Case 1, 2, 3, 5, 6, 7, 33, 34, 35, 36, 37, 38 'skip
                    GoTo nextCount1

            End Select



            strF2 = "ID_TBLCONFIGREPORTTABLES = " & int1
            Erase rows2
            rows2 = dtbl2.Select(strF2)
            boolSamples = rows2(0).Item("BOOLSAMPLES")
            boolQC = rows2(0).Item("BOOLQCSTATS")
            boolCS = rows2(0).Item("BOOLCSSTATS")
            boolCS = 0 'Calibration standards do not have outliers
            If boolQC = -1 Or boolCS = -1 Then 'include

                intCt = intCt + 1
                If intCt = 1 Then
                    strF3 = strF3 & "ID_TBLCONFIGREPORTTABLES = " & int1
                Else
                    strF3 = strF3 & " OR ID_TBLCONFIGREPORTTABLES = " & int1
                End If
            End If

nextCount1:

        Next
        strF3 = strF3 & ")"
        If intCt = 0 Then
            str1 = "There are no tables in this study that have Outlier type data."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        'dgv = Me.dgvTables
        'dgv.RowHeadersWidth = 25
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        'dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        'dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        'make a dv
        Dim dv1 As System.Data.DataView = New DataView(dtbl1, strF3, strS1, DataViewRowState.CurrentRows)

        '*****

        'dv = frmH.dgvReportTableConfiguration.DataSource
        Dim tbl As System.Data.DataTable = dv1.ToTable("a")

        'add some columns to tbl
        Dim col1 As New DataColumn
        col1.ColumnName = "BOOLSHOWAREA"
        'col1.DataType = System.Type.GetType("System.Double")
        tbl.Columns.Add(col1)
        Dim col2 As New DataColumn
        col2.ColumnName = "BOOLSHOWCONC"
        'col2.DataType = System.Type.GetType("System.Double")
        tbl.Columns.Add(col2)
        Dim col3 As New DataColumn
        col3.ColumnName = "BOOLINCLUDEIS"
        'col3.DataType = System.Type.GetType("System.Double")
        tbl.Columns.Add(col3)
        'Dim col4 As New DataColumn
        'col4.ColumnName = "ID_TBLREPORTTABLES"
        'col4.DataType = System.Type.GetType("System.INT64")
        'tbl.Columns.Add(col4)

        'enter data in new columns
        int1 = tbl.Rows.Count
        For Count1 = 0 To int1 - 1
            str1 = tbl.Rows.Item(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            strF = "ID_TBLCONFIGREPORTTABLES = '" & str1 & "'"

            'str1 = tbl.Rows.Item(Count1).Item("ID_TBLREPORTTABLES")
            'strF = "ID_TBLREPORTTABLES = '" & str1 & "'"

            rows = tblConfigReportTables.Select(strF)
            For Count2 = 0 To rows.Length - 1
                tbl.Rows.Item(Count1).BeginEdit()
                tbl.Rows.Item(Count1).Item("BOOLSHOWAREA") = rows(0).Item("BOOLSHOWAREA")
                tbl.Rows.Item(Count1).Item("BOOLSHOWCONC") = rows(0).Item("BOOLSHOWCONC")
                tbl.Rows.Item(Count1).Item("BOOLINCLUDEIS") = rows(0).Item("BOOLINCLUDEIS")
                tbl.Rows.Item(Count1).EndEdit()
            Next
        Next

        ''enter more data in new columns
        'For Count1 = 0 To int1 - 1
        '    str1 = tbl.Rows.Item(Count1).Item("ID_TBLREPORTTABLES")
        '    strF = "ID_TBLREPORTTABLES = '" & str1 & "'"
        '    rows = tbl.Select(strF)
        '    tbl.Rows.Item(Count1).BeginEdit()
        '    tbl.Rows.Item(Count1).Item("ID_TBLREPORTTABLES") = rows(0).Item("ID_TBLREPORTTABLES")
        '    tbl.Rows.Item(Count1).EndEdit()
        'Next

        'strF = "boolRequiresSampleAssignment = -1" ' & True
        strF = "boolRequiresSampleAssignment = " & True 'Leave as true. Underlying table has boolean
        'strS = "ID_TBLCONFIGREPORTTABLES ASC"
        'strS = "ORDER ASC"
        'Dim dv1 as system.data.dataview = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)

        'strS = "INTORDER ASC"
        'Dim dv1 as system.data.dataview = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)
        Dim dv2 As System.Data.DataView = New DataView(tbl)

        dv2.AllowDelete = False
        dv2.AllowNew = False
        dgv = Me.dgvTables

        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

        Dim boolAA As Boolean
        boolAA = boolCont
        boolCont = False
        dgv.DataSource = dv2
        boolCont = boolAA

        intRows = dv1.Count
        int1 = dgv.Columns.Count

        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next
        dgv.Columns.Item("CHARHEADINGTEXT").Visible = True
        dgv.Columns.Item("CHARHEADINGTEXT").HeaderText = "Table"
        dgv.Columns.Item("ID_TBLREPORTTABLE").Visible = False
        dgv.RowHeadersWidth = 25

        Dim wd1, wd2
        wd1 = dgv.RowHeadersWidth
        wd2 = dgv.Width - wd1
        dgv.Columns.Item("CHARHEADINGTEXT").MinimumWidth = wd2 * 0.95
        dgv.AutoResizeRows()


        ''don't select a row
        'If boolFormLoad Then
        '    'find selected row
        '    intRows_frmh = frmH.dgvReportTableConfiguration.Rows.Count
        '    If intRows_frmh = 0 Then
        '        intRow = -1
        '    ElseIf frmH.dgvReportTableConfiguration.CurrentRow Is Nothing Then
        '        intRow = 0
        '    Else
        '        intRow = frmH.dgvReportTableConfiguration.CurrentRow.Index
        '    End If
        '    'now record table id
        '    var1 = frmH.dgvReportTableConfiguration.Rows.Item(intRow).Cells("BOOLREQUIRESSAMPLEASSIGNMENT").Value
        '    var2 = frmH.dgvReportTableConfiguration.Rows.Item(intRow).Cells("ID_TBLCONFIGREPORTTABLES").Value
        '    If var1 = -1 Then
        '        'find var2 in dgv
        '        For Count1 = 0 To dgv.Rows.Count - 1
        '            var3 = dgv.Item("ID_TBLCONFIGREPORTTABLES", Count1).Value
        '            If var3 = var2 Then
        '                intRow = Count1
        '                Exit For
        '            End If
        '        Next
        '    Else
        '        intRow = 0
        '    End If
        'Else
        '    'select first row
        '    intRow = 0
        'End If
        'dgv.Select()
        'If intRows = 0 Or intRow = -1 Then
        'Else
        '    dgv.Rows.Item(intRow).Cells("CHARHEADINGTEXT").Selected = True
        'End If

        Call FilldgvAnalytes()

        'Call SetAnalysisResultsTable(wStudyID)

        'Call FillAnalyticalRuns()

        If boolFormLoad Then
        Else
            'Call FillAssignedSamples()
        End If

    End Sub

    Sub FillTblAnalytes(ByVal idT As Int64, ByVal boolIS As Boolean, ByVal intRow As Short)
        Dim tbl As System.Data.DataTable
        Dim Count1 As Short
        Dim bool As Boolean
        Dim dv As System.Data.DataView
        'Dim intRow As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim var1
        Dim strF As String
        'Dim boolIS As Boolean
        Dim int1 As Short
        Dim intS As Short
        Dim intE As Short
        Dim str1 As String
        Dim strS As String
        Dim dgv As DataGridView
        Dim boolISIn As Boolean
        Dim boolAA As Boolean
        Dim intSetRow As Short
        Dim boolH As Boolean

        dgv = Me.dgvAnalytes
        tbl = Me.tblAnalytes
        dv = Me.dgvTables.DataSource

        ''record initial row of dgvAnalytes
        'If dgv.RowCount = 0 Then
        '    intSetRow = 0
        'ElseIf dgv.CurrentRow Is Nothing Then
        '    intSetRow = 0
        'Else
        '    intSetRow = dgv.CurrentRow.Index
        'End If

        'intRow = 0
        'If Me.dgvTables.CurrentRow Is Nothing Then
        '    intRow = 0
        'Else
        '    intRow = Me.dgvTables.CurrentRow.Index
        'End If
        'int1 = Me.dgvTables("BOOLINCLUDEIS", intRow).Value

        'If int1 = -1 Then
        '    boolIS = True
        'Else
        '    boolIS = False
        'End If

        If intRow = -1 Then 'there are no tables selected to assign samples
        Else
            strF = ""
            Count3 = 0
            For Count1 = 0 To tbl.Rows.Count - 1
                var1 = tbl.Rows.Item(Count1).Item("AnalyteDescription")
                str1 = NZ(tbl.Rows.Item(Count1).Item("IsIntStd"), "")

                If boolIS Then
                    If StrComp(str1, "No", CompareMethod.Text) = 0 Then
                        bool = dv(intRow).Item(var1)
                        If bool Then
                            Count3 = Count3 + 1
                            If Count3 = 1 Then
                                strF = strF & "ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                            Else
                                strF = strF & " OR ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                            End If
                        End If
                    Else
                        strF = strF & " OR ANALYTEDESCRIPTION = '" & var1 & "'"
                    End If
                Else
                    'bool = dv(intRow).Item(arrAnalytes(1, Count1))
                    If StrComp(str1, "No", CompareMethod.Text) = 0 Then
                        bool = dv(intRow).Item(var1)
                        If bool Then
                            Count3 = Count3 + 1
                            If Count3 = 1 Then
                                strF = strF & "ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                            Else
                                strF = strF & " OR ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                            End If
                        End If
                    End If
                End If
            Next

            'inspect current dv
            Dim dv2 As System.Data.DataView = New DataView(tbl, strF, "ANALYTEDESCRIPTION", DataViewRowState.CurrentRows)
            dv2.AllowDelete = False
            dv2.AllowNew = False

            boolH = boolHold
            boolHold = True
            dgv.DataSource = dv2
            dgv.ReadOnly = True
            boolHold = boolH

        End If

        'select first row
        If dgv.RowCount = 0 Then
        Else

            'dgv.CurrentCell = dgv.Rows(0).Cells(0)
            boolAA = boolCont
            boolCont = False
            boolH = boolHold
            boolHold = True

            dgv.Rows(0).Selected = True
            'set initial row

            dgv.CurrentCell = dgv.Rows.Item(0).Cells(0)

            boolHold = boolH

            boolCont = boolAA
        End If

    End Sub

    Sub FilldgvAnalytes()

        Dim tbl As System.Data.DataTable
        Dim Count1 As Short
        Dim bool As Boolean
        Dim dv As System.Data.DataView
        Dim intRow As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim var1
        Dim strF As String
        Dim boolIS As Boolean
        Dim int1 As Short
        Dim intS As Short
        Dim intE As Short
        Dim str1 As String
        Dim strS As String
        Dim dgv As DataGridView
        Dim boolISIn As Boolean
        Dim boolAA As Boolean
        Dim intSetRow As Short

        dgv = Me.dgvAnalytes
        tbl = Me.tblAnalytes
        dv = Me.dgvTables.DataSource

        'record initial row of dgvAnalytes
        If dgv.RowCount = 0 Then
            intSetRow = 0
        ElseIf dgv.CurrentRow Is Nothing Then
            intSetRow = 0
        Else
            intSetRow = dgv.CurrentRow.Index
        End If

        intRow = 0

        If Me.dgvTables.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvTables.CurrentRow.Index
        End If
        int1 = Me.dgvTables("BOOLINCLUDEIS", intRow).Value

        If int1 = -1 Then
            boolIS = True
        Else
            boolIS = False
        End If

        If intRow = -1 Then 'there are no tables selected to assign samples
        Else
            strF = ""
            Count3 = 0
            For Count1 = 0 To tbl.Rows.Count - 1
                var1 = tbl.Rows.Item(Count1).Item("AnalyteDescription")
                str1 = NZ(tbl.Rows.Item(Count1).Item("IsIntStd"), "")

                If boolIS Then
                    If StrComp(str1, "No", CompareMethod.Text) = 0 Then
                        bool = dv(intRow).Item(var1)
                        If bool Then
                            Count3 = Count3 + 1
                            If Count3 = 1 Then
                                strF = strF & "ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                            Else
                                strF = strF & " OR ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                            End If
                        End If
                    Else
                        'strF = strF & " OR ANALYTEDESCRIPTION = '" & var1 & "'"
                        If Len(strF) = 0 Then
                            strF = "ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                        Else
                            strF = strF & " OR ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                        End If
                    End If
                Else
                    'bool = dv(intRow).Item(arrAnalytes(1, Count1))
                    If StrComp(str1, "No", CompareMethod.Text) = 0 Then
                        bool = dv(intRow).Item(var1)
                        If bool Then
                            Count3 = Count3 + 1
                            If Count3 = 1 Then
                                strF = strF & "ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                            Else
                                strF = strF & " OR ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                            End If
                        End If
                    End If
                End If
            Next

            'inspect current dv
            Dim dv2 As System.Data.DataView = New DataView(tbl, strF, "ANALYTEDESCRIPTION", DataViewRowState.CurrentRows)
            dv2.AllowDelete = False
            dv2.AllowNew = False

            Me.dgvAnalytes.DataSource = dv2
            Me.dgvAnalytes.ReadOnly = True


            'str1 = dv2.RowFilter
            'If StrComp(str1, strF, CompareMethod.Text) = 0 Then 'ignore
            'Else
            '    strS = "IsIntStd ASC, AnalyteDescription ASC"
            '    Dim dv1 as system.data.dataview = New DataView(Me.tblAnalytes)
            '    dv1.RowFilter = strF
            '    dv1.Sort = strS
            '    dv1.AllowDelete = False
            '    dv1.AllowNew = False
            '    boolAA = boolCont
            '    boolCont = False
            '    dgv.DataSource = dv1
            '    dgv.AutoResizeColumns()
            '    boolCont = boolAA
            'End If
        End If

        'select first row
        If dgv.RowCount = 0 Then
        Else

            'dgv.CurrentCell = dgv.Rows(0).Cells(0)
            dgv.Rows(0).Selected = True
            'set initial row
            boolAA = boolCont
            boolCont = False
            'If boolFromdgvTable Then
            '    'str1 = dgv.Rows.item(intSetRow).Cells("AnalyteDescription").Value
            '    If dgv.RowCount - 1 < intSetRow Then
            '        int1 = dgv.RowCount
            '        Dim str2 As String
            '        Dim boolGo As Boolean
            '        boolGo = False
            '        str1 = strAnalFromTable
            '        For Count1 = 0 To int1 - 1
            '            str2 = dgv.Rows.Item(Count1).Cells("AnalyteDescription").Value
            '            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
            '                boolGo = True
            '                Exit For
            '            End If
            '            If boolGo Then
            '                dgv.CurrentCell = dgv.Rows.Item(Count1).Cells(0)
            '            Else
            '                dgv.CurrentCell = dgv.Rows.Item(0).Cells(0)
            '            End If
            '        Next
            '    Else
            '        str1 = dgv.Rows.Item(intSetRow).Cells("AnalyteDescription").Value
            '        If StrComp(str1, strAnalFromTable, CompareMethod.Text) = 0 Then
            '            dgv.CurrentCell = dgv.Rows.Item(intSetRow).Cells("AnalyteDescription")
            '        Else 'look for stranalfromtable
            '            int1 = dgv.RowCount
            '            Dim str2 As String
            '            Dim boolGo As Boolean
            '            boolGo = False
            '            str1 = strAnalFromTable
            '            For Count1 = 0 To int1 - 1
            '                str2 = dgv.Rows.Item(Count1).Cells("AnalyteDescription").Value
            '                If StrComp(str1, str2, CompareMethod.Text) = 0 Then
            '                    boolGo = True
            '                    Exit For
            '                End If
            '                If boolGo Then
            '                    dgv.CurrentCell = dgv.Rows.Item(Count1).Cells(0)
            '                Else
            '                    dgv.CurrentCell = dgv.Rows.Item(0).Cells(0)
            '                End If
            '            Next
            '        End If
            '    End If
            'Else
            '    dgv.CurrentCell = dgv.Rows.Item(0).Cells(0)
            'End If

            dgv.CurrentCell = dgv.Rows.Item(0).Cells(0)

            boolCont = boolAA
        End If


    End Sub

    Sub ConfigResultsdgv()

        Dim dgv As DataGridView

        dgv = Me.dgvResults
        dgv.RowHeadersWidth = 25
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells


    End Sub


    Private Sub dgvTables_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTables.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        'If boolHold Then
        '    Exit Sub
        'End If

        Dim bool As Boolean

        Cursor.Current = Cursors.WaitCursor

        bool = boolFormLoad
        boolFormLoad = True
        Call FilldgvAnalytes()
        Cursor.Current = Cursors.WaitCursor
        boolFormLoad = bool

        Dim intRow As Short

        intRow = Me.dgvTables.CurrentRow.Index

        Dim idT As Int64
        idT = Me.dgvTables("ID_TBLREPORTTABLE", intRow).Value

        Cursor.Current = Cursors.WaitCursor
        Call RowSelect(id_tblStudies, idT, "A", False, intRow)
        Cursor.Current = Cursors.WaitCursor

        Call FilterSummary()
        Cursor.Current = Cursors.Default

    End Sub


    Private Sub dgvAnalytes_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAnalytes.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Call SelectAnalyteRow()

        'If boolHold Then
        '    Exit Sub
        'End If


    End Sub

    Sub SelectAnalyteRow()

        Dim intRow As Short
        Dim intRowT As Short

        intRow = Me.dgvAnalytes.CurrentRow.Index
        intRowT = Me.dgvTables.CurrentRow.Index

        Dim idT As Int64
        idT = Me.dgvTables("ID_TBLREPORTTABLE", intRowT).Value

        'Call RowSelect(id_tblStudies, Me.dgvTables("ID_TBLREPORTTABLE", intRowT).Value, Me.dgvAnalytes("ANALYTEDESCRIPTION", intRow).Value, False)
        Call RowSelect(id_tblStudies, idT, "A", False, intRowT)

        Call FilterSummary()

        Me.dgvSummary.SuspendLayout()
        Me.dgvSummary.AutoResizeColumns()
        Me.dgvSummary.ResumeLayout()

    End Sub


    Sub FilterSummary()

        Dim dgvA As DataGridView = Me.dgvAnalytes
        Dim dgvS As DataGridView = Me.dgvSummary
        Dim dv As DataView = dgvS.DataSource
        Dim dgvT As DataGridView = Me.dgvTables
        Dim dgvR As DataGridView = Me.dgvResults


        Dim strF As String
        Dim intGroup As Short
        Dim idT As Int64
        Dim var1
        Dim intCol As Short
        Dim Count1 As Short

        Try
            intGroup = dgvA("INTGROUP", dgvA.CurrentRow.Index).Value
            idT = dgvT("ID_TBLREPORTTABLE", dgvT.CurrentRow.Index).Value
            strF = "INTGROUP = " & intGroup & " AND CHAROUTLIER = 'X' AND ID_TBLREPORTTABLE = " & idT
            dv.RowFilter = strF

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        If dgvS.RowCount = 0 Then
            'select first row in dgvR
            If dgvR.RowCount = 0 Then
            Else

                intCol = 0
                For Count1 = 0 To dgvR.Columns.Count - 1
                    If dgvR.Columns(Count1).Visible Then
                        intCol = Count1
                        Exit For
                    End If
                Next

                dgvR.CurrentCell = dgvR.Rows(0).Cells(intCol)
                dgvR.CurrentRow.Selected = True

            End If
        End If

    End Sub

    Sub Initialize_tblCritZ()

        'If tblCritZ.Rows.Count > 0 Then
        '    Exit Sub
        'End If

        Dim tbl As System.Data.DataTable
        tbl = tblCritZ

        'add columns
        If tbl.Columns.Contains("N") Then
        Else
            Dim col1 As New DataColumn
            col1.ColumnName = "N"
            col1.Caption = "N"
            col1.DataType = System.Type.GetType("System.Int16")
            tbl.Columns.Add(col1)
            Dim col2 As New DataColumn
            col2.ColumnName = "CriticalZ"
            col2.Caption = "Crit" & ChrW(10) & "Z"
            col2.DataType = System.Type.GetType("System.Decimal")
            tbl.Columns.Add(col2)

            'add additional columns for Dixon
            Dim col3 As New DataColumn
            col3.ColumnName = "95P"
            col3.Caption = "95P"
            col3.DataType = System.Type.GetType("System.Decimal")
            tbl.Columns.Add(col3)

            Dim col4 As New DataColumn
            col4.ColumnName = "99P"
            col4.Caption = "99P"
            col4.DataType = System.Type.GetType("System.Decimal")
            tbl.Columns.Add(col4)

        End If

        tbl.Clear()

        Dim Count1 As Short
        Dim var1, var2, var3, var4

        var1 = ""
        var2 = ""
        If Me.rbGrubbs.Checked Then

            For Count1 = 1 To 48
                Dim newRow As DataRow = tbl.NewRow
                newRow.BeginEdit()
                Select Case Count1
                    Case 1
                        var1 = 3
                        var2 = 1.15
                    Case 2
                        var1 = 4
                        var2 = 1.48
                    Case 3
                        var1 = 5
                        var2 = 1.71
                    Case 4
                        var1 = 6
                        var2 = 1.89
                    Case 5
                        var1 = 7
                        var2 = 2.02
                    Case 6
                        var1 = 8
                        var2 = 2.13
                    Case 7
                        var1 = 9
                        var2 = 2.21
                    Case 8
                        var1 = 10
                        var2 = 2.29
                    Case 9
                        var1 = 11
                        var2 = 2.34
                    Case 10
                        var1 = 12
                        var2 = 2.41
                    Case 11
                        var1 = 13
                        var2 = 2.46
                    Case 12
                        var1 = 14
                        var2 = 2.51
                    Case 13
                        var1 = 15
                        var2 = 2.55
                    Case 14
                        var1 = 16
                        var2 = 2.59
                    Case 15
                        var1 = 17
                        var2 = 2.62
                    Case 16
                        var1 = 18
                        var2 = 2.65
                    Case 17
                        var1 = 19
                        var2 = 2.68
                    Case 18
                        var1 = 20
                        var2 = 2.71
                    Case 19
                        var1 = 21
                        var2 = 2.73
                    Case 20
                        var1 = 22
                        var2 = 2.76
                    Case 21
                        var1 = 23
                        var2 = 2.78
                    Case 22
                        var1 = 24
                        var2 = 2.8
                    Case 23
                        var1 = 25
                        var2 = 2.82
                    Case 24
                        var1 = 26
                        var2 = 2.84
                    Case 25
                        var1 = 27
                        var2 = 2.86
                    Case 26
                        var1 = 28
                        var2 = 2.88
                    Case 27
                        var1 = 29
                        var2 = 2.89
                    Case 28
                        var1 = 30
                        var2 = 2.91
                    Case 29
                        var1 = 31
                        var2 = 2.92
                    Case 30
                        var1 = 32
                        var2 = 2.94
                    Case 31
                        var1 = 33
                        var2 = 2.95
                    Case 32
                        var1 = 34
                        var2 = 2.97
                    Case 33
                        var1 = 35
                        var2 = 2.98
                    Case 34
                        var1 = 36
                        var2 = 2.99
                    Case 35
                        var1 = 37
                        var2 = 3
                    Case 36
                        var1 = 38
                        var2 = 3.01
                    Case 37
                        var1 = 39
                        var2 = 3.03
                    Case 38
                        var1 = 40
                        var2 = 3.04
                    Case 39
                        var1 = 50
                        var2 = 3.13
                    Case 40
                        var1 = 60
                        var2 = 3.2
                    Case 41
                        var1 = 70
                        var2 = 3.26
                    Case 42
                        var1 = 80
                        var2 = 3.31
                    Case 43
                        var1 = 90
                        var2 = 3.35
                    Case 44
                        var1 = 100
                        var2 = 3.38
                    Case 45
                        var1 = 110
                        var2 = 3.42
                    Case 46
                        var1 = 120
                        var2 = 3.44
                    Case 47
                        var1 = 130
                        var2 = 3.47
                    Case 48
                        var1 = 140
                        var2 = 3.49

                End Select

                newRow.Item("N") = var1
                newRow.Item("CriticalZ") = var2

                tbl.Rows.Add(newRow)
            Next

        ElseIf Me.rbDixon.Checked Then

            '20190226 LEE:
            'Added 95 and 99 CL
            For Count1 = 1 To 11
                Dim newRow As DataRow = tbl.NewRow
                newRow.BeginEdit()
                Select Case Count1
                    Case 1
                        var1 = 3
                        var2 = 0.941
                        var3 = 0.97
                        var4 = 0.994
                    Case 2
                        var1 = 4
                        var2 = 0.765
                        var3 = 0.829
                        var4 = 0.926
                    Case 3
                        var1 = 5
                        var2 = 0.642
                        var3 = 0.71
                        var4 = 0.821
                    Case 4
                        var1 = 6
                        var2 = 0.56
                        var3 = 0.625
                        var4 = 0.74
                    Case 5
                        var1 = 7
                        var2 = 0.507
                        var3 = 0.568
                        var4 = 0.68
                    Case 6
                        var1 = 8
                        var2 = 0.468
                        var3 = 0.526
                        var4 = 0.634
                    Case 7
                        var1 = 9
                        var2 = 0.437
                        var3 = 0.493
                        var4 = 0.598
                    Case 8
                        var1 = 10
                        var2 = 0.412
                        var3 = 0.466
                        var4 = 0.568
                    Case 9
                        var1 = 11
                        var2 = 0.392
                        var3 = 0.444
                        var4 = 0.542
                    Case 10
                        var1 = 12
                        var2 = 0.376
                        var3 = 0.426
                        var4 = 0.522
                    Case 11
                        var1 = 13
                        var2 = 0.361
                        var3 = 0.41
                        var4 = 0.503

                End Select

                newRow.Item("N") = var1
                newRow.Item("CriticalZ") = var2
                newRow.Item(2) = var3
                newRow.Item(3) = var4

                tbl.Rows.Add(newRow)
            Next

        End If

    End Sub

    Private Sub dgvSummary_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSummary.CellClick

        If boolFormLoad Then
            Exit Sub
        End If

        Call ChangeSummaryRow()

    End Sub


    Private Sub dgvSummary_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSummary.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Call ChangeSummaryRow()

    End Sub

    Sub ChangeSummaryRow()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim id1 As Int64
        Dim id2 As Int64
        Dim Count1 As Short
        Dim intRow As Short
        Dim intCol As Short

        dgv1 = Me.dgvResults

        intCol = 0
        For Count1 = 0 To dgv1.Columns.Count - 1
            If dgv1.Columns(Count1).Visible Then
                intCol = Count1
                Exit For
            End If
        Next

        dgv2 = Me.dgvSummary

        If dgv2.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv2.CurrentRow Is Nothing Then
            Exit Sub
        End If

        intRow = dgv2.CurrentRow.Index

        id2 = NZ(dgv2("ID_TBLRESULTS", intRow).Value, -1)
        For Count1 = 0 To dgv1.Rows.Count - 1
            id1 = NZ(dgv1("ID_TBLRESULTS", Count1).Value, -1)
            If id1 = id2 Then
                'If Count1 = dgv1.CurrentRow.Index Then
                '    'need to find analyte
                '    Dim idT As Int64
                '    Dim charAnalyte As String
                '    Dim int1 As Short
                '    'call rowselect(id_tblstudies,idT
                'Else
                '    dgv1.CurrentCell = dgv1.Rows(Count1).Cells(intCol)
                '    dgv1.CurrentRow.Selected = True
                'End If
                dgv1.CurrentCell = dgv1.Rows(Count1).Cells(intCol)
                dgv1.CurrentRow.Selected = True
                Exit For
            End If
        Next

        'pesky
        dgv2.AutoResizeColumns()

    End Sub

    Sub GatherAll()

        'this sub should simply populate dgvAllSummary
        Dim strF As String
        Dim dgv As DataGridView = Me.dgvAllSummary
        Dim str1 As String
        Dim strS As String

        strS = ReturnSummarAllSort()

        strF = "CHAROUTLIER = 'X'"
        Dim dv As DataView

        Dim var1
        Try
            dv = New DataView(Me.tblResults, strF, strS, DataViewRowState.CurrentRows)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try
        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        Dim ctRows As Int16
        ctRows = dv.Count

        Me.dgvAllSummary.DataSource = dv

        If ctRows = 1 Then
            Me.lblStatus.Text = ctRows & " Outlier found"
        Else
            Me.lblStatus.Text = ctRows & " Outliers found"
        End If

    End Sub


    Sub SummaryAllHeaders()

        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim dgv3 As DataGridView
        Dim dgv4 As DataGridView

        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable

        Dim str1 As String
        Dim str2 As String
        Dim ctRows As Short

        Dim idT As Int64
        Dim strAnal As String

        Dim var1

        Dim intOrig As Short


        dgv3 = Me.dgvSummary
        dgv4 = Me.dgvAllSummary

        Me.lblStatus.Visible = True

        'If boolAllSummary Then
        'Else

        Call Config_dgvSummaryAll()

        For Count1 = 0 To dgv4.Columns.Count - 1
            dgv4.Columns(Count1).Visible = False
            dgv4.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'For Count1 = 0 To dgv3.Columns.Count - 1
        '    str1 = dgv3.Columns(Count1).Name
        '    'dgv4.Columns(str1).Visible = dgv3.Columns(str1).Visible
        '    dgv4.Columns(str1).SortMode = dgv3.Columns(str1).SortMode
        '    dgv4.Columns(str1).DefaultCellStyle.Alignment = dgv3.Columns(str1).DefaultCellStyle.Alignment
        '    Try
        '        dgv4.Columns(str1).HeaderText = dgv3.Columns(str1).HeaderText

        '    Catch ex As Exception

        '    End Try

        'Next

        Try

            dgv4.Columns("CHARHEADINGTEXT").Visible = True
            dgv4.Columns("CHARHEADINGTEXT").HeaderText = "Table"
            dgv4.Columns("CHARHEADINGTEXT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

            dgv4.Columns("ANALYTEDESCRIPTION").Visible = True
            dgv4.Columns("ANALYTEDESCRIPTION").HeaderText = "Analyte"
            dgv4.Columns("ANALYTEDESCRIPTION").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

            dgv4.Columns("RUNID").Visible = True
            dgv4.Columns("RUNID").HeaderText = "Run ID"
            dgv4.Columns("RUNID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            'dgv4.Columns("RUNSAMPLESEQUENCENUMBER").Visible = True
            'dgv4.Columns("RUNSAMPLESEQUENCENUMBER").HeaderText = "Seq #"
            'dgv4.Columns("RUNSAMPLESEQUENCENUMBER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv4.Columns("RUNSAMPLESEQUENCENUMBER").SortMode = DataGridViewColumnSortMode.NotSortable

            dgv4.Columns("RUNSAMPLEORDERNUMBER").Visible = True
            dgv4.Columns("RUNSAMPLEORDERNUMBER").HeaderText = "Seq #"
            dgv4.Columns("RUNSAMPLEORDERNUMBER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv4.Columns("CHARHELPER1").Visible = True
            dgv4.Columns("CHARHELPER1").HeaderText = "Level" '"Ident1"
            dgv4.Columns("CHARHELPER1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv4.Columns("CHARHELPER2").Visible = False ' True
            dgv4.Columns("CHARHELPER2").HeaderText = "Ident2"
            dgv4.Columns("CHARHELPER2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv4.Columns("NOMCONC").Visible = True
            dgv4.Columns("NOMCONC").HeaderText = "Nom Conc"
            dgv4.Columns("NOMCONC").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            Call SetStatsColumns(dgv4)

            dgv4.Columns("CHAROUTLIER").Visible = True
            dgv4.Columns("CHAROUTLIER").HeaderText = "Outlier (X)"
            dgv4.Columns("CHAROUTLIER").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

            dgv4.Columns("CONCENTRATION").Visible = True
            str1 = "Conc. or" & ChrW(10) & "Area Ratio"
            str1 = Replace(str1, " ", ChrW(160), 1, -1)
            dgv4.Columns("CONCENTRATION").HeaderText = str1 '"Conc. or" & ChrW(10) & "Area Ratio"
            dgv4.Columns("CONCENTRATION").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            dgv4.Columns("ANALYTEAREA").Visible = True
            dgv4.Columns("ANALYTEAREA").HeaderText = "Peak Area"
            dgv4.Columns("ANALYTEAREA").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            dgv4.Columns("INTERNALSTANDARDAREA").Visible = True
            dgv4.Columns("INTERNALSTANDARDAREA").HeaderText = "IS Peak Area"
            dgv4.Columns("INTERNALSTANDARDAREA").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight




            'If intConc = -1 Then
            '    dgv4.Columns.Item("CONCENTRATION").Visible = True
            '    dgv4.Columns("CONCENTRATION").HeaderText = "Conc."
            '    dgv4.Columns.Item("ANALYTEAREA").Visible = False
            '    dgv4.Columns("ANALYTEAREA").HeaderText = "Peak Area"
            '    dgv4.Columns.Item("INTERNALSTANDARDAREA").Visible = False
            '    dgv4.Columns("INTERNALSTANDARDAREA").HeaderText = "IS Peak Area"
            'ElseIf intArea = -1 Or intISArea = -1 Then
            '    If StrComp(varIntStd, "No", CompareMethod.Text) = 0 Then
            '        dgv4.Columns.Item("CONCENTRATION").Visible = False
            '        dgv4.Columns("CONCENTRATION").HeaderText = "Conc."
            '        dgv4.Columns.Item("ANALYTEAREA").Visible = True
            '        dgv4.Columns("ANALYTEAREA").HeaderText = "Peak Area"
            '        dgv4.Columns.Item("INTERNALSTANDARDAREA").Visible = False
            '        dgv4.Columns("INTERNALSTANDARDAREA").HeaderText = "IS Peak Area"
            '    Else
            '        dgv4.Columns.Item("CONCENTRATION").Visible = False
            '        dgv4.Columns("CONCENTRATION").HeaderText = "Conc."
            '        dgv4.Columns.Item("ANALYTEAREA").Visible = False
            '        dgv4.Columns("ANALYTEAREA").HeaderText = "Peak Area"
            '        dgv4.Columns.Item("INTERNALSTANDARDAREA").Visible = True
            '        dgv4.Columns("INTERNALSTANDARDAREA").HeaderText = "IS Peak Area"
            '    End If
            'End If


            'If intConc = -1 Then
            '    dgv4.Columns.Item("CONCENTRATION").Visible = True
            '    dgv4.Columns("CONCENTRATION").HeaderText = "Conc."
            '    dgv4.Columns.Item("ANALYTEAREA").Visible = False
            '    dgv4.Columns("ANALYTEAREA").HeaderText = "Peak Area"
            '    dgv4.Columns.Item("INTERNALSTANDARDAREA").Visible = False
            '    dgv4.Columns("INTERNALSTANDARDAREA").HeaderText = "IS Peak Area"
            'ElseIf intArea = -1 Or intISArea = -1 Then
            '    If StrComp(varIntStd, "No", CompareMethod.Text) = 0 Then
            '        dgv4.Columns.Item("CONCENTRATION").Visible = False
            '        dgv4.Columns("CONCENTRATION").HeaderText = "Conc."
            '        dgv4.Columns.Item("ANALYTEAREA").Visible = True
            '        dgv4.Columns("ANALYTEAREA").HeaderText = "Peak Area"
            '        dgv4.Columns.Item("INTERNALSTANDARDAREA").Visible = False
            '        dgv4.Columns("INTERNALSTANDARDAREA").HeaderText = "IS Peak Area"
            '    Else
            '        dgv4.Columns.Item("CONCENTRATION").Visible = False
            '        dgv4.Columns("CONCENTRATION").HeaderText = "Conc."
            '        dgv4.Columns.Item("ANALYTEAREA").Visible = False
            '        dgv4.Columns("ANALYTEAREA").HeaderText = "Peak Area"
            '        dgv4.Columns.Item("INTERNALSTANDARDAREA").Visible = True
            '        dgv4.Columns("INTERNALSTANDARDAREA").HeaderText = "IS Peak Area"
            '    End If

            'End If


            'End If

            dgv4.AutoResizeColumns()

            dgv4.Columns("CHARHEADINGTEXT").Width = Me.dgvTables.Columns("CHARHEADINGTEXT").Width ' * 0.7
            dgv4.Columns("RUNID").Width = 30
            dgv4.Columns("NOMCONC").Width = 50
            dgv4.Columns("RUNSAMPLEORDERNUMBER").Width = 30

            dgv4.AutoResizeRows()

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try



    End Sub

    Sub Config_dgvSummaryAll()
        If boolAllSummary = False Then
            Dim dgv As DataGridView

            dgv = Me.dgvAllSummary
            dgv.RowHeadersWidth = 25
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells

        End If
    End Sub

    Private Sub cmdReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReport.Click

        Dim strPath As Object
        Dim str1 As String
        Dim boolDetailed As Boolean
        Dim boolAll As Boolean

        Dim frmO As New frmAskOutlierReport
        frmO.ShowDialog()

        boolAll = True
        If frmO.boolCancel Then
            frmO.Dispose()
            Me.Refresh()
            Exit Sub
        Else
            If frmO.rbDetailed.Checked Then
                boolDetailed = True
                If frmO.rbAll.Checked Then
                    boolAll = True
                Else
                    boolAll = False
                End If
            Else
                boolDetailed = False
            End If
            frmO.Dispose()
            Me.Refresh()
        End If

        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim boolSTB As Boolean = True
        Try
            boolSTB = wd.Application.ShowWindowsInTaskbar
        Catch ex As Exception

        End Try

        Cursor.Current = Cursors.WaitCursor

        strPath = GetNewTempFile(True)
        strPath = Replace(strPath, ".xml", ".docx", 1, -1, CompareMethod.Text)
        wd.Documents.Add()


        If Me.rbGrubbs.Checked Or Me.rbStdDev.Checked Or Me.rbDixon.Checked Then
            Call DoGrubbsReport(wd, boolDetailed, boolAll)

        End If

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        'wd.ActiveDocument.SaveAs(FileName:=strPath, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument)
        wd.ActiveDocument.SaveAs(FileName:=strPath, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
        'wdFormatXMLDocument

        Try
            wd.Application.ShowWindowsInTaskbar = boolSTB
        Catch ex As Exception

        End Try

        'wd.Application.Visible = True
        'wd.Application.Activate()


        wd.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)

        Cursor.Current = Cursors.Default

        Call OpenAFR(strPath, "", False, boolSTB, True, False)

        'Dim frm As New frmWordStatement

        'frm.boolOutlier = True
        'frm.boolSTB = boolSTB
        'frm.boolReport = True
        'frm.strReport = strPath
        ''frm.panEditReports.Visible = False
        'frm.pan2.Visible = True
        'frm.cmdFieldCode.Visible = False
        'frm.cmdAddStatement.Visible = False
        'frm.cmdSave.Visible = False

        'frm.panList.Visible = False
        'frm.panSave.Visible = False

        'frm.Text = "StudyDoc Outlier Report"
        'frm.lblSection.Text = ""
        'frm.boolEdit = False

        'frm.Show()

        'Cursor.Current = Cursors.Default


    End Sub

    Sub DoGrubbsReport(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal boolDetailed As Boolean, ByVal boolAllTables As Boolean)

        Dim strStudy As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim var1, var2, var3
        Dim intSumCols As Short
        Dim intSumRows As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim dgv3 As DataGridView
        Dim dgv4 As DataGridView
        Dim boolAll As Boolean
        Dim intRows As Short
        Dim strStatusOrig As String
        Dim boolZ As Boolean

        Dim dtbl1 As System.Data.DataTable
        Dim dtbl2 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF2 As String
        Dim strF3 As String

        Dim strMethod As String

        Dim stats1, stats2 As String
        Dim strSD As String

        Dim intCt As Short

        stats1 = "Z"
        stats2 = "Crit Z"
        If Me.rbGrubbs.Checked Then
            strMethod = "Grubbs Test"
            stats1 = "Z"
            stats2 = "Crit Z"
        ElseIf Me.rbStdDev.Checked Then
            strSD = Me.txtStdDev.Text
            stats1 = "-" & strSD & " SD"
            stats2 = "+" & strSD & " SD"
            strMethod = "Standard Deviation +/-" & strSD & " SD"
        ElseIf Me.rbDixon.Checked Then
            stats1 = "R"
            stats2 = "Crit R"
            var1 = Me.cbxCL.Items(Me.cbxCL.SelectedIndex).ToString
            strMethod = "Dixon Q Test (" & var1 & "% confidence level)"
        End If

        Dim lm, rm, pw, rt

        If boolDetailed Then
        Else
            With wd
                'landscape summary all page
                With .Selection.PageSetup
                    .Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape
                End With
            End With

        End If

        lm = wd.ActiveDocument.Sections(1).PageSetup.LeftMargin
        rm = wd.ActiveDocument.Sections(1).PageSetup.RightMargin
        pw = wd.ActiveDocument.PageSetup.PageWidth
        rt = pw - lm - rm

        strStudy = frmH.dgvwStudy("StudyName", frmH.dgvwStudy.CurrentRow.Index).Value
        dgv1 = Me.dgvAllSummary

        strStatusOrig = Me.lblStatus.Text


        Dim strP1 As String
        Dim strP2 As String
        Dim strP3 As String
        Dim strP4 As String
        Dim strPO As String = Me.lblProgress.Text
        strP1 = "Creating Stats Report:" & ChrW(10) & ChrW(10)
        Me.lblProgress.Text = strP1
        Me.lblProgress.Visible = True
        Me.lblProgress.Refresh()

        With wd



            'create Header with page numbers
            If wd.ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
                wd.ActiveWindow.Panes(2).Close()
            End If
            If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow. _
                ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            End If
            wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader

            .Selection.TypeText(Text:="Outlier Report for Study " & strStudy)
            .Selection.TypeParagraph()
            .Selection.TypeText(Text:="Outlier Method:  " & strMethod)
            .Selection.TypeParagraph()
            .Selection.TypeText(Text:=Format(Now, "MMMM dd, yyyy   hh:mm:ss tt"))

            .Selection.ParagraphFormat.TabStops.ClearAll()

            '.Selection.ParagraphFormat.TabStops(InchesToPoints(6)).Position = 432 ' InchesToPoints(7)
            .Selection.ParagraphFormat.TabStops.Add(rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
            .Selection.TypeText(Text:=vbTab & "Page ")
            .Selection.Fields.Add(Range:=.Selection.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
            .Selection.TypeText(Text:=" of ")
            .Selection.Fields.Add(Range:=.Selection.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)

            .Selection.TypeParagraph()

            .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
            End With

            .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

            .Selection.TypeParagraph()
            If boolDetailed Then
                .Selection.TypeParagraph()
                .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

                'landscape summary all page
                With .Selection.PageSetup
                    .Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape
                End With
            End If


            'create All Table Summary Section
            'count visible rows
            intSumCols = 0

            For Count1 = 0 To dgv1.Columns.Count - 1
                If dgv1.Columns(Count1).Visible Then
                    intSumCols = intSumCols + 1
                End If
            Next
            intSumRows = dgv1.Rows.Count
            If intSumRows = 0 Then
                boolAll = False
                intRows = 1
            Else
                boolAll = True
                intRows = intSumRows
            End If
            Try
                .Selection.Style = wd.ActiveDocument.Styles("Heading 1")
            Catch ex As Exception

            End Try
            .Selection.TypeText(Text:="Summary Stats Method Results for All Study Tables")
            .Selection.TypeParagraph()
            .Selection.Style = wd.ActiveDocument.Styles("Normal")

            'add a table
            wrdSelection = .Selection
            .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intRows + 1, NumColumns:=intSumCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
            .Selection.Tables.Item(1).Select()
            .Selection.Font.Size = 10
            .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
            .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints

            'first align entire table
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

            'now align first column
            .Selection.Tables.Item(1).Columns(1).Select()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom


            'now select first row
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.SelectRow()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.Font.Bold = True
            .Selection.Rows.HeadingFormat = True

            '.Selection.Tables.Item(1).Columns.Item(1).Width = 86
            .Selection.Tables.Item(1).Select()

            ''remove border, but leave top and bottom and sides
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone

            'enter column headers
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            int1 = 0
            For Count1 = 0 To dgv1.ColumnCount - 1
                If dgv1.Columns(Count1).Visible Then
                    int1 = int1 + 1
                    .Selection.Tables.Item(1).Cell(1, int1).Select()
                    .Selection.TypeText(Text:=dgv1.Columns(Count1).HeaderText)
                End If
            Next

            'enter data into all summary table
            If boolAll Then
                int2 = 1
                For Count2 = 0 To dgv1.RowCount - 1
                    int1 = 0
                    int2 = int2 + 1
                    For Count1 = 0 To dgv1.ColumnCount - 1
                        If dgv1.Columns(Count1).Visible Then
                            int1 = int1 + 1
                            .Selection.Tables.Item(1).Cell(int2, int1).Select()
                            .Selection.TypeText(Text:=NZ(dgv1(Count1, Count2).Value, ""))
                        End If
                    Next
                Next
            Else
                .Selection.Tables.Item(1).Cell(2, 1).Select()
                .Selection.TypeText(Text:="No Outliers in this study")
            End If

            .Selection.Tables.Item(1).Select()
            'autofit table
            Call AutoFitTable(wd, False)

            If boolDetailed Then 'continue

            Else
                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                .ActiveWindow.ActivePane.View.Zoom.Percentage = 75
                GoTo end1
            End If


            'create detailed sections
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            .Selection.TypeParagraph()
            .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
            With .Selection.PageSetup
                .Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait
            End With

            dgv1 = Me.dgvTables
            dgv2 = Me.dgvAnalytes
            dgv3 = Me.dgvSummary
            dgv4 = Me.dgvResults

            Dim idT As Int64
            Dim strAnal As String
            Dim strTable As String
            Dim varColor
            Dim strL As String
            Dim strF As String
            Dim strS As String
            Dim rowsT() As DataRow
            Dim varIntStd
            Dim str2 As String

            Dim idC As Int64

            varColor = .Selection.Font.Color

            Try
                .Selection.Style = wd.ActiveDocument.Styles("Heading 1")
            Catch ex As Exception

            End Try
            .Selection.TypeText(Text:="Detailed Stats Method Results for Individual Study Tables")
            .Selection.TypeParagraph()

            Dim intRow As Short
            Dim boolIS As Boolean
            Dim rowsA() As DataRow
            Dim strFA As String
            Dim strSA As String
            Dim intS As Short
            Dim intE As Short
            Dim intGroup As Short

            intRow = dgv1.CurrentRow.Index

            If boolAllTables Then
                intS = 0
                intE = dgv1.RowCount - 1
            Else
                intS = intRow
                intE = intRow
            End If

            For Count1 = intS To intE

                strP2 = "Table " & Count1 + 1 & " of " & dgv1.RowCount & " tables"
                Me.lblStatus.Text = strP2
                Me.lblStatus.Refresh()
                Me.lblProgress.Text = strP1 & strP2
                Me.lblProgress.Refresh()


                idT = dgv1("ID_TBLREPORTTABLE", Count1).Value
                strTable = dgv1("CHARHEADINGTEXT", Count1).Value
                Try
                    .Selection.Style = wd.ActiveDocument.Styles("Heading 2")
                Catch ex As Exception

                End Try
                .Selection.TypeText(Text:=strTable)
                .Selection.TypeParagraph()

                int1 = Me.dgvTables("BOOLINCLUDEIS", Count1).Value

                If int1 = -1 Then
                    boolIS = True
                    strFA = "IsIntStd = 'Yes' OR IsIntStd = 'No'"
                Else
                    boolIS = False
                    strFA = "IsIntStd = 'No'"
                End If
                strSA = "IsIntStd ASC, ANALYTEDESCRIPTION ASC"
                'Call FillTblAnalytes(idT, boolIS, Count1)
                'Erase rowsA
                Try
                    rowsA = tblAnalytesHome.Select(strFA, strSA)
                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1
                End Try

                'dgv2

                Dim strField As String

                'wd.Visible = True

                For Count2 = 0 To rowsA.Length - 1

                    strP3 = "Analyte " & Count2 + 1 & " of " & rowsA.Length & " analytes"
                    'Me.lblStatus.Text = Count1 + 1 & " of " & dgv1.RowCount & " tables...Analyte " & Count2 + 1 & " of " & rowsA.Length
                    Me.lblStatus.Text = strP2 & "..." & strP3
                    Me.lblStatus.Refresh()
                    strP4 = strP1 & strP2 & ChrW(10) & ChrW(10) & strP3
                    Me.lblProgress.Text = strP4
                    Me.lblProgress.Refresh()

                    Me.lblProgress.Text = strP1 & strP2
                    Me.lblProgress.Refresh()

                    strAnal = rowsA(Count2).Item("ANALYTEDESCRIPTION")
                    varIntStd = rowsA(Count2).Item("IsIntStd")
                    intGroup = rowsA(Count2).Item("INTGROUP")

                    dtbl1 = Me.tblAllSummary
                    'dtbl1.Clear()
                    dtbl2 = Me.tblResultsTemp

                    strF = "CHAROUTLIER = 'X' AND INTGROUP = " & intGroup ' & " AND ID_TBLCONFIGREPORTTABLES IS NOT NULL"
                    'strS = "NOMCONC ASC, CHARHELPER1 ASC, CHARHELPER2 ASC,  RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
                    strS = "NOMCONC ASC, CHARHELPER1 ASC, CHARHELPER2 ASC,  RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

                    strS = "ID_TBLRESULTS ASC"

                    Cursor.Current = Cursors.WaitCursor
                    Try
                        .Selection.Style = wd.ActiveDocument.Styles("Heading 3")
                    Catch ex As Exception

                    End Try
                    .Selection.TypeText(Text:="Analyte: " & strAnal)
                    .Selection.TypeParagraph()
                    .Selection.Style = wd.ActiveDocument.Styles("Normal")

                    ''intRows = Me.tblResults.Rows.Count
                    'intRows = dtbl2.Rows.Count

                    'If intRows = 0 Then
                    '    boolZ = True
                    '    intRows = 1
                    'Else
                    '    boolZ = False
                    'End If

                    strF = "INTGROUP = " & intGroup & " AND ID_TBLREPORTTABLE = " & idT
                    Erase rowsT
                    rowsT = Me.tblResults.Select(strF, strS)
                    intRows = rowsT.Length

                    If intRows = 0 Then
                        boolZ = True
                        intRows = 1
                    Else
                        boolZ = False
                    End If

                    'intSumCols = 0
                    'For Count3 = 0 To dgv4.ColumnCount - 1
                    '    If dgv4.Columns(Count3).Visible Then
                    '        intSumCols = intSumCols + 1
                    '    End If
                    'Next

                    intSumCols = 9

                    'wd.Visible = True

                    'add a table
                    wrdSelection = .Selection
                    .ActiveDocument.Tables.Add(Range:=wrdSelection.Range, NumRows:=intRows + 1, NumColumns:=intSumCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent)
                    .Selection.Tables.Item(1).Select()
                    .Selection.Font.Size = 10
                    .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                    .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints

                    'first align entire table
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                    'If boolZ Then
                    '    'now align first column
                    '    .Selection.Tables.Item(1).Columns(1).Select()
                    '    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    '    .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                    'End If


                    'now select first row
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.SelectRow()
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                    .Selection.Font.Bold = True
                    .Selection.Rows.HeadingFormat = True

                    '.Selection.Tables.Item(1).Columns.Item(1).Width = 86
                    .Selection.Tables.Item(1).Select()

                    ''remove border, but leave top and bottom and sides
                    '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone
                    '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.wdlinestyle.wdLineStyleNone

                    'enter column headers
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    '.Selection.TypeText(Text:="Analyte")
                    'determine str2
                    intConc = dgv1("BOOLSHOWCONC", Count1).Value
                    intArea = dgv1("BOOLSHOWAREA", Count1).Value
                    intISArea = dgv1("BOOLINCLUDEIS", Count1).Value
                    idC = dgv1("ID_TBLCONFIGREPORTTABLES", Count1).Value

                    'If idC = 22 Or idC = 23 Then
                    '    str2 = "Area Ratio"
                    '    strField = "CONCENTRATION"
                    'Else
                    '    If intConc = -1 Then
                    '        str2 = "Conc"
                    '        strField = "CONCENTRATION"
                    '    ElseIf intArea = -1 Or intISArea = -1 Then
                    '        If StrComp(varIntStd, "No", CompareMethod.Text) = 0 Then
                    '            str2 = "Peak Area"
                    '            strField = "ANALYTEAREA"
                    '        Else
                    '            str2 = "IS Peak Area"
                    '            strField = "INTERNALSTANDARDAREA"
                    '        End If
                    '    End If
                    'End If

                    '*****

                    Dim boolRCConc As Boolean
                    Dim boolRCPA As Boolean
                    Dim boolRCPARatio As Boolean
                    Dim boolIncludeISTbl As Boolean

                    Call ReturnColumnTypes(boolRCConc, boolRCPA, boolRCPARatio, boolIncludeISTbl, idT) 'this will return bools

                    If intConc = -1 Then
                        str2 = "Conc"
                        strField = "CONCENTRATION"
                    ElseIf intArea = -1 Or intISArea = -1 Then
                        If boolRCPARatio Then
                            str2 = "Peak Area Ratio"
                            strField = "CONCENTRATION"
                        Else
                            If StrComp(varIntStd, "No", CompareMethod.Text) = 0 Then
                                str2 = "Peak Area"
                                strField = "ANALYTEAREA"
                            Else
                                str2 = "IS Peak Area"
                                strField = "INTERNALSTANDARDAREA"
                            End If
                        End If
                    End If

                    '*****

                    Dim arrInfo(intRows + 1, intSumCols)

                    For Count3 = 1 To intSumCols
                        Select Case Count3
                            Case 1
                                str1 = "Run ID"
                            Case 2
                                str1 = "Seq #"
                            Case 3
                                str1 = "Nom Conc"
                            Case 4
                                str1 = "Level" ' "Ident1"
                                'Case 5
                                '    str1 = "Ident2"
                            Case 5
                                str1 = str2
                            Case 6
                                str1 = stats1 '"Z"
                            Case 7
                                str1 = stats2 '"Crit Z"
                            Case 8
                                'str1 = "Outlier" & ChrW(10) & "(X)"
                                str1 = "Outlier (X)"
                        End Select
                        .Selection.Tables.Item(1).Cell(1, Count3).Select()
                        .Selection.TypeText(Text:=str1)
                        arrInfo(1, Count3) = str1
                    Next
                    'int1 = 0
                    'For Count3 = 0 To dgv3.ColumnCount - 1
                    '    If dgv3.Columns(Count3).Visible Then
                    '        int1 = int1 + 1
                    '        .Selection.Tables.Item(1).Cell(1, int1).Select()
                    '        .Selection.TypeText(Text:=dgv3.Columns(Count3).HeaderText)
                    '    End If
                    'Next

                    'now enter data
                    strL = Me.lblStatus.Text
                    If boolZ Then
                        'select 2nd row
                        .Selection.Tables.Item(1).Cell(2, 1).Select()
                        .Selection.SelectRow()
                        .Selection.Cells.Merge()
                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                        .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                        .Selection.TypeText(Text:="No data availble")

                    Else
                        intCt = -1
                        'For Count3 = 0 To dtbl2.Rows.Count - 1
                        For Count3 = 0 To rowsT.Length - 1
                            If Count3 > intCt Then
                                'Me.lblStatus.Text = strL & "...Row " & Count3 & " of " & dtbl2.Rows.Count - 1
                                If Count3 = 0 Then
                                    intCt = 10
                                Else
                                    intCt = intCt + 10
                                End If
                                Me.lblStatus.Refresh()
                            ElseIf Count3 = rowsT.Length - 1 Then
                                'Me.lblStatus.Text = strL & "...Row " & Count3 & " of " & dtbl2.Rows.Count - 1
                                Me.lblStatus.Refresh()
                            End If

                            '''.Selection.Tables.Item(1).Cell(Count3 + 2, 1).Select()
                            '.Selection.TypeText(Text:=strAnal)
                            int1 = 0
                            For Count4 = 0 To intSumCols - 1 'dgv4.Columns.Count - 1
                                Select Case Count4 + 1
                                    Case 1
                                        str1 = "RUNID"
                                    Case 2
                                        'str1 = "RUNSAMPLESEQUENCENUMBER"
                                        str1 = "RUNSAMPLEORDERNUMBER"
                                    Case 3
                                        str1 = "NOMCONC"
                                    Case 4
                                        str1 = "CHARHELPER1"
                                        'Case 5
                                        '    str1 = "CHARHELPER2"
                                    Case 5
                                        str1 = strField
                                    Case 6
                                        str1 = "STATS1"
                                    Case 7
                                        str1 = "STATS2"
                                    Case 8
                                        str1 = "CHAROUTLIER"
                                End Select
                                int1 = int1 + 1
                                var1 = NZ(rowsT(Count3).Item(str1), "")

                                '''.Selection.Tables.Item(1).Cell(Count3 + 2, int1).Select()
                                If StrComp(str1, "CHAROUTLIER", CompareMethod.Text) = 0 Then
                                    If StrComp(var1, "X", CompareMethod.Text) = 0 Then
                                        .Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorRed
                                        .Selection.Font.Bold = True
                                    End If
                                Else

                                End If
                                '''.Selection.TypeText(Text:=var1)
                                Select Case Count4 + 1
                                    Case 5
                                        If Len(var1) = 0 Then 'means this is a blank line
                                        Else
                                            If intConc = -1 Then
                                                var1 = SigFigOrDecString(var1, LSigFig, False)
                                            ElseIf intArea = -1 Or intISArea = -1 Then
                                                If boolRCPARatio Then
                                                    var1 = SigFigAreaRatio(var1, LSigFigAreaRatio, False, False)
                                                Else
                                                    var1 = SigFigArea(var1, LSigFigArea, False, False)
                                                End If
                                            End If
                                        End If

                                End Select
                                arrInfo(Count3 + 2, Count4 + 1) = var1

                                'If dgv4.Columns(Count4).Visible Then
                                '    int1 = int1 + 1
                                '    var1 = NZ(dgv4(Count4, Count3).Value, "")
                                '    str1 = dgv4.Columns(Count4).Name

                                '    .Selection.Tables.Item(1).Cell(Count3 + 2, int1).Select()
                                '    If StrComp(str1, "CHAROUTLIER", CompareMethod.Text) = 0 Then
                                '        If StrComp(var1, "X", CompareMethod.Text) = 0 Then
                                '            .Selection.Font.Color = Microsoft.Office.Interop.Word.wdcolor.wdColorRed
                                '            .Selection.Font.Bold = True
                                '        End If
                                '    Else

                                '    End If
                                '    .Selection.TypeText(Text:=var1)
                                'End If
                            Next
                        Next
                    End If

                    '*****copy/paste data into table
                    'send strpaste to clipboard
                    Dim strPaste As String
                    Dim CountA As Int16
                    Dim CountB As Int16

                    strPaste = ""
                    For CountA = 1 To intRows + 1
                        For CountB = 1 To intSumCols
                            If CountB = 1 And CountA = 1 Then
                                strPaste = arrInfo(CountA, CountB) & ChrW(9)
                            ElseIf CountB = intSumCols Then
                                strPaste = strPaste & arrInfo(CountA, CountB) & ChrW(10)
                            ElseIf CountB = intSumCols And CountA = intRows + 1 Then
                                strPaste = strPaste & arrInfo(CountA, CountB)
                            Else
                                strPaste = strPaste & arrInfo(CountA, CountB) & ChrW(9)
                            End If
                        Next
                    Next

                    '''''''console.writeline("Start")
                    '''''''console.writeline(strPaste)
                    '''''''console.writeline("End")

                    'wdd.visible = True

                    Try
                        Clipboard.Clear()
                    Catch ex As Exception

                    End Try
                    Try
                        Clipboard.SetText(strPaste, TextDataFormat.Text)
                    Catch ex As Exception

                    End Try
                    'select appropriate rows
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
                    .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
                    'paste from clipboard
                    Try
                        .Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault)
                    Catch ex As Exception

                    End Try


                    'the paste action removes the range object and any table formatting, must reset it
                    Call GlobalTableParaFormat(wd)
                    .Selection.Tables.Item(1).Select()
                    .Selection.Font.Size = 10
                    .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False
                    .Selection.Tables.Item(1).Columns.PreferredWidth = Microsoft.Office.Interop.Word.WdPreferredWidthType.wdPreferredWidthPoints

                    'first align entire table
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                    'now align first column
                    '.Selection.Tables.Item(1).Columns(1).Select()

                    'wdd.visible = True
                    '.Selection.Tables.Item(1).Columns(1).Select()
                    ''Why?
                    'Try
                    '    .Selection.Tables.Item(1).Columns(1).Select()
                    '    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    '    .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                    'Catch ex As Exception

                    'End Try
                    '.Selection.Tables.Item(1).Columns(1).Select()
                    '.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    '.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom

                    'now select first row
                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.SelectRow()
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
                    .Selection.Font.Bold = True
                    .Selection.Rows.HeadingFormat = True

                    '*****

                    'go back and add table information to top of table

                    .Selection.Tables.Item(1).Cell(1, 1).Select()
                    .Selection.InsertRowsAbove(1)
                    .Selection.Cells.Merge()
                    .Selection.SelectCell()
                    .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom


                    With .Selection.ParagraphFormat
                        .LeftIndent = 72 / 4 * 3 'InchesToPoints(0.99)
                        .SpaceBeforeAuto = False
                        .SpaceAfterAuto = False
                    End With
                    With .Selection.ParagraphFormat
                        .SpaceBeforeAuto = False
                        .SpaceAfterAuto = False
                        .FirstLineIndent = -72 / 4 * 3 'InchesToPoints(-0.99)
                    End With
                    .Selection.TypeText(Text:="Table:" & vbTab & strTable)
                    .Selection.TypeParagraph()
                    .Selection.TypeText(Text:="Analyte:" & vbTab & strAnal)

                    .Selection.Tables.Item(1).Select()
                    'autofit table
                    Call AutoFitTable(wd, False)


                    MoveOneCellDown(wd)
                    .Selection.TypeParagraph()
                    .Selection.TypeParagraph()

                Next

                If Count1 <> dgv1.RowCount - 1 Then
                    .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
                End If

            Next

            'move home
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            'insert TOC
            Call TOC(wd, rt)

end1:

            Me.lblStatus.Text = strStatusOrig
            Me.lblProgress.Visible = False
            Me.lblProgress.Text = strPO


        End With
    End Sub

    Sub TOC(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal rt As Object)
        Dim wrdS As Microsoft.Office.Interop.Word.Selection

        wrdS = wd.Selection
        With wd

            .ActiveDocument.Tables.Add(Range:=wrdS.Range, NumRows:=2, NumColumns:=1, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)


            'remove border
            .Selection.Tables.Item(1).Select()
            'remove borders
            Call removeAllBorders(wd, False)

            'select first row
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.SelectRow()
            .Selection.Rows.HeadingFormat = True

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Font.Size = .Selection.Font.Size + 2
            .Selection.Font.Bold = True
            .Selection.TypeText(Text:="Table of Contents")
            .Selection.TypeParagraph()
            .Selection.Font.Size = .Selection.Font.Size - 2
            .Selection.Font.Bold = False
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
            .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
            .Selection.TypeText(Text:="Section")
            .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
            '.Selection.ParagraphFormat.TabStops.Add(Position:=72 * 6.5, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight)
            .Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight)
            .Selection.TypeText(Text:=vbTab)
            .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle
            .Selection.TypeText(Text:="Page Number")
            .Selection.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone
            .Selection.TypeParagraph()
            .Selection.Tables.Item(1).Cell(2, 1).Select()
            'With .ActiveDocument
            '    .TablesOfContents.Add(Range:=.Selection.Range, RightAlignPageNumbers:= _
            '        True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
            '        LowerHeadingLevel:=3, IncludePageNumbers:=True, AddedStyles:="", _
            '        UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
            '        True)
            '    .TablesOfContents(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
            '    .TablesOfContents.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent

            'End With

            wrdS = wd.Selection
            With .ActiveDocument
                .TablesOfContents.Add(Range:=wrdS.Range, RightAlignPageNumbers:= _
                    True, UseHeadingStyles:=True, UpperHeadingLevel:=1, _
                    LowerHeadingLevel:=3, IncludePageNumbers:=True, AddedStyles:="", _
                    UseHyperlinks:=True, HidePageNumbersInWeb:=True, UseOutlineLevels:= _
                    True)
                .TablesOfContents(1).TabLeader = Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots
                .TablesOfContents.Format = Microsoft.Office.Interop.Word.WdIndexType.wdIndexIndent

            End With
        End With

    End Sub

    Private Sub dgvAllSummary_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAllSummary.SelectionChanged
        If boolFormLoad Then
            Exit Sub
        End If
        'boolHold = True
        Call SelectTableRow()
        'boolHold = False

    End Sub

    Sub SelectTableRow()

        Dim dgvAS As DataGridView = Me.dgvAllSummary
        Dim dgvA As DataGridView = Me.dgvAnalytes
        Dim dgvT As DataGridView = Me.dgvTables


        Dim intRowAS As Short = 0
        Dim intRowA As Short = 0
        Dim intRowT As Short = 0
        Dim intRowS As Short = 0

        Dim Count1 As Short
        Dim Count2 As Short

        Dim intCol As Short

        Dim var1
        Dim id As Int64
        Dim id1 As Int64
        Dim id2 As Int64

        Dim idT As Int64
        Dim intGroup As Short
        Dim idR As Int64

        Dim boolT As Boolean = boolFormLoad
        boolFormLoad = True

        If dgvAS.Rows.Count = 0 Then
            GoTo end1
        End If

        If dgvAS.CurrentRow Is Nothing Then
            intRowAS = 0
        Else
            intRowAS = dgvAS.CurrentRow.Index
        End If

        Try
            idT = dgvAS("ID_TBLREPORTTABLE", intRowAS).Value
            intGroup = dgvAS("INTGROUP", intRowAS).Value
            idR = dgvAS("ID_TBLRESULTS", intRowAS).Value

            'first select dgvt
            intRowT = -1
            For Count1 = 0 To dgvT.Rows.Count - 1
                id = dgvT("ID_TBLREPORTTABLE", Count1).Value
                If id = idT Then
                    intRowT = Count1
                    Exit For
                End If
            Next
            If intRowT = -1 Then
                GoTo end1
            End If

            'now select dgva
            intRowA = -1
            For Count1 = 0 To dgvA.Rows.Count - 1
                id = dgvA("INTGROUP", Count1).Value
                If id = intGroup Then
                    intRowA = Count1
                    Exit For
                End If
            Next
            If intRowA = -1 Then
                GoTo end1
            End If



            'select rows
            'dgvT
            For Count1 = 0 To dgvT.Columns.Count - 1
                If dgvT.Columns(Count1).Visible Then
                    intCol = Count1
                    Exit For
                End If
            Next
            If intRowT = -1 Then
            Else
                dgvT.CurrentCell = dgvT.Rows(intRowT).Cells(intCol)
                dgvT.Rows(intRowT).Selected = True
            End If

            'dgva
            For Count1 = 0 To dgvA.Columns.Count - 1
                If dgvA.Columns(Count1).Visible Then
                    intCol = Count1
                    Exit For
                End If
            Next
            If intRowA = -1 Then
            Else
                dgvA.CurrentCell = dgvA.Rows(intRowA).Cells(intCol)
                dgvA.Rows(intRowA).Selected = True
            End If

            boolFormLoad = boolT

            'call selection change
            Call SelectAnalyteRow()

            'now select correct dgvSummary row
            Dim dgvS As DataGridView = Me.dgvSummary

            Try
                If dgvS.Rows.Count < 1 Then
                Else
                    'now select dgvS
                    intRowS = -1
                    For Count1 = 0 To dgvS.Rows.Count - 1
                        id = dgvS("ID_TBLRESULTS", Count1).Value
                        If id = idR Then
                            intRowS = Count1
                            Exit For
                        End If
                    Next
                    If intRowS = -1 Then
                        GoTo end1
                    End If

                    For Count1 = 0 To dgvS.Columns.Count - 1
                        If dgvS.Columns(Count1).Visible Then
                            intCol = Count1
                            Exit For
                        End If
                    Next
                    dgvS.CurrentCell = dgvS.Rows(intRowS).Cells(intCol)
                    dgvS.Rows(intRowS).Selected = True
                End If
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


end1:

        boolFormLoad = boolT

    End Sub


    Private Sub lblGrubbsTable_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblGrubbsTable.LinkClicked

        'Me.lblGTable.Text = "Grubbs Test Critical Z Values"

        'Call LoadGTable()

        'Call SetGTable()

        Me.panGTable.Visible = True

    End Sub

    Sub SetGTable()

        Me.panGTable.Top = Me.dgvTables.Top
        Me.panGTable.Left = Me.dgvTables.Left
        Me.panGTable.Height = Me.dgvTables.Height
        Me.panGTable.Width = Me.dgvTables.Width

    End Sub

    Sub LoadGTable()

        Dim strL As String
        Dim dgv As DataGridView

        dgv = Me.dgvGTable

        If Me.rbGrubbs.Checked Then
            strL = "Grubbs, Frank (February 1969), Procedures for Detecting Outlying Observations in Samples, Technometrics, Vol. 11, No. 1, pp. 1-21."
            Me.lblSource.Text = strL

            Dim dv As System.Data.DataView = New DataView(Me.tblCritZ)

            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False

            dgv.DataSource = dv

        ElseIf Me.rbDixon.Checked Then

            strL = "W. J. Dixon, Processing Data for Outliers, Biometrics, Vol 9, No. 1, 74-89, 1953"
            strL = strL & "  https://sebastianraschka.com/Articles/2014_dixon_test.html"
            Me.lblSource.Text = strL

            Dim dv As System.Data.DataView = New DataView(Me.tblCritZ)

            dv.AllowDelete = False
            dv.AllowEdit = False
            dv.AllowNew = False

            dgv.DataSource = dv

        End If

        Try
            'format columns
            If Me.rbGrubbs.Checked Then
                dgv.Columns(0).HeaderText = "n Samples"
                dgv.Columns(1).HeaderText = "Crit Z Value"

                'format column 2
                dgv.Columns(1).DefaultCellStyle.Format = "0.00"

                dgv.Columns(2).Visible = False
                dgv.Columns(3).Visible = False

            ElseIf Me.rbDixon.Checked Then
                dgv.Columns(0).HeaderText = "n Samples"
                'dgv.Columns(1).HeaderText = "Crit R Value"
                '20190226 LEE:
                dgv.Columns(1).HeaderText = "90%"
                dgv.Columns(2).HeaderText = "95%"
                dgv.Columns(3).HeaderText = "99%"
                'format column 2, 3, 4
                dgv.Columns(1).DefaultCellStyle.Format = "0.000"
                dgv.Columns(2).DefaultCellStyle.Format = "0.000"
                dgv.Columns(3).DefaultCellStyle.Format = "0.000"

                dgv.Columns(2).Visible = True
                dgv.Columns(3).Visible = True

                'autosize columns
                dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Me.panGTable.Visible = False
    End Sub


    Sub EnableStuff()

        If Me.rbGrubbs.Checked Then

            Me.lblGrubbsTable.Enabled = True

            Me.lblStdDev.Enabled = False
            Me.txtStdDev.Enabled = False

            Me.lblDixonTable.Enabled = False

        ElseIf Me.rbStdDev.Checked Then
            Me.lblStdDev.Enabled = True
            Me.txtStdDev.Enabled = True

            Me.lblDixonTable.Enabled = False

            Me.lblGrubbsTable.Enabled = False

        ElseIf Me.rbDixon.Checked Then

            Me.lblDixonTable.Enabled = True
            Me.lblGrubbsTable.Enabled = False

            Me.lblStdDev.Enabled = False
            Me.txtStdDev.Enabled = False

        End If

    End Sub

    Sub VisRBs()

        If boolFormLoad Then
            Exit Sub
        End If

        Call EnableStuff()

        Dim boolF As Boolean
        boolF = boolFirstRow
        boolFirstRow = False


        'Call SummaryAllHeaders()

        '20160306 LEE:
        'don't do this
        'Call GatherAll()
        'do this instead
        Call FillOutlier()

        Call SummaryAllHeaders()

        'don't need this
        'Call SelectTableRow()

        boolFirstRow = boolF

        '*****

        'If boolFormLoad Then
        '    Exit Sub
        'End If

        'Call EnableStuff()

        'Dim boolF As Boolean
        'boolF = boolFirstRow
        'boolFirstRow = False


        'Call SummaryAllHeaders()

        'Call GatherAll()

        'Call SelectTableRow()

        'boolFirstRow = boolF


    End Sub


    Sub ChangeOutlier()

        If boolChangeOutlier Then
            boolChangeOutlier = False
            Exit Sub
        End If

        boolChangeOutlier = True

        Dim str1 As String
        Dim str2 As String
        Dim strM As String
        Dim boolCont As Boolean = False
        Dim int1 As Short
        Dim strM1 As String
        Dim strM2 As String
        Dim boolM As Boolean = False

        str1 = "Changing Stats Method..."
        Me.lblProgress.Text = str1
        Me.lblProgress.Visible = True
        Me.lblProgress.Refresh()

        Dim dgvR As DataGridView = Me.dgvResults

        dgvR.SuspendLayout()

        Dim var1, var2

        If Me.rbGrubbs.Checked Then

            Call Initialize_tblCritZ()

            Me.lblGTable.Text = "Grubbs Test Critical Z Values"

            Call LoadGTable()

            Call SetGTable()

            Call VisRBs()

        ElseIf Me.rbDixon.Checked Then


            Me.lblGTable.Text = "Dixon Test Critical R Values for Confidence Levels" ' & ChrW(10) & "(90% Confidence Interval)"

            Call Initialize_tblCritZ()

            Call LoadGTable()

            Call SetGTable()

            Call VisRBs()

        ElseIf Me.rbStdDev.Checked Then


            str1 = Me.txtStdDev.Text

            'strM = "Do you wish to use the current Standard Deviation value of " & str1 & "?"
            'strM = strM & ChrW(10) & ChrW(10) & "If not, enter a new value and click OK"

            'strM1 = "The SD value must a number > 0"

            'strM2 = "You have cycled this loop too many times." & Chr(10)
            'strM2 = strM2 & "The original SD value of " & str1 & " will be used."

            int1 = 0
            Do Until boolCont
                int1 = int1 + 1
                'str2 = InputBox(strM, "Accept SD value?", str1)

                str2 = str1
                If Len(str2) = 0 Then
                    boolM = True
                    GoTo end1
                End If

                If IsNumeric(str2) Then
                Else
                    boolM = True
                    GoTo end1
                End If

                If CDec(str2) > 0 Then
                Else
                    boolM = True
                    GoTo end1
                End If

                boolM = False
                boolCont = True

end1:
                If boolM Then
                    MsgBox(strM1, MsgBoxStyle.Exclamation, "Invalid entry...")
                    Me.Refresh()
                End If

                If int1 > 4 Then
                    Me.Refresh()
                    MsgBox(strM2, MsgBoxStyle.Information, "End of loop...")
                    str2 = str1
                    boolCont = True
                End If
            Loop

            Me.Refresh()

            Me.txtStdDev.Text = str2 'this will call validating event that triggers 

            Call VisRBs()

        End If

        Try
            Call FillSummary(tblResults)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        str1 = "Summarizing Results..."
        Me.lblProgress.Text = str1
        Me.lblProgress.Refresh()
        If boolFormLoad Then
            Me.lblFormLoadProgress.Text = str1
            Me.lblFormLoadProgress.Refresh()
        End If

        Try
            Call GatherAll()
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        Me.dgvResults.SuspendLayout()
        Call SetStatsColumns(Me.dgvResults)

        Me.dgvResults.SuspendLayout()
        Me.dgvResults.AutoResizeRows()
        'Me.dgvResults.AutoResizeColumns()

        Me.dgvResults.ResumeLayout()

        If boolFormLoad Then
        Else
            Me.lblProgress.Visible = False
            Me.lblProgress.Refresh()
        End If

    End Sub

    Private Sub txtStdDev_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtStdDev.Validating

        Dim var1
        Dim boolE As Boolean = False
        Dim strM As String

        If boolFormLoad Then
            Exit Sub
        End If

        var1 = Me.txtStdDev.Text
        If Len(var1) = 0 Then
            boolE = True
            GoTo end1
        End If

        If IsNumeric(var1) Then
        Else
            boolE = True
            GoTo end1
        End If

        If var1 <= 0 Then
            boolE = True
            GoTo end1
        End If

        boolChangeOutlier = False
        Call ChangeOutlier()

end1:
        If boolE Then
            strM = "Entry must be number > 0"
            MsgBox(strM, MsgBoxStyle.Exclamation, "Invalid entry...")
            e.Cancel = True
        Else

        End If


    End Sub

    Private Sub txtStdDev_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStdDev.Validated

        'Call VisRBs()

    End Sub

    Private Sub rbStdDev_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbStdDev.CheckedChanged

        Call ChangeOutlier()

    End Sub
    Private Sub rbGrubbs_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbGrubbs.CheckedChanged

        Call ChangeOutlier()

    End Sub

    Private Sub rbDixon_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbDixon.CheckedChanged

        Call ChangeOutlier()

    End Sub

    Private Sub cbxCL_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxCL.SelectedIndexChanged

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.rbDixon.Checked Then
            boolChangeOutlier = False
            Call ChangeOutlier()
            boolChangeOutlier = False
        End If

    End Sub

    Private Sub lblDixonTable_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblDixonTable.LinkClicked

        'Me.lblGTable.Text = "Dixon Test Critical R Values" & ChrW(10) & "(90% Confidence Interval)"

        'Call LoadGTable()

        'Call SetGTable()

        Me.panGTable.Visible = True

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub ToolTipSet()

        ' Create the ToolTip and associate with the Form container.
        Dim toolTip1 As New ToolTip()

        toolTip1.AutomaticDelay = intToolTipDelay
        toolTip1.ShowAlways = True

        Try
            toolTip1.SetToolTip(Me.lblDixonTable, "View table of R criteria vs. number of samples (for P=0.05)")
            toolTip1.SetToolTip(Me.lblGrubbsTable, "View table of Z criteria vs. number of samples (for P=0.05)")

            Me.dgvResults.Columns.Item("CHAROUTLIER").ToolTipText = "Outliers (according to this test) are marked with an 'X'"
            Me.dgvResults.Columns.Item("CONCENTRATION").ToolTipText = "Calculated Concentration"
            Me.dgvAllSummary.Columns.Item("CHAROUTLIER").ToolTipText = "Outliers (according to this test) are marked with an 'X'"
            Me.dgvAllSummary.Columns.Item("CONCENTRATION").ToolTipText = "Calculated Concentration"
        Catch ex As Exception

        End Try

    End Sub

    Private Sub gbStatsMethod_Enter(sender As Object, e As EventArgs) Handles gbStatsMethod.Enter

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click

        Me.Dispose()
        frmH.Activate()

    End Sub

    Private Sub dgvAnalytes_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAnalytes.CellContentClick

    End Sub

    Private Sub dgvSummary_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvSummary.CellContentClick

    End Sub

    Private Sub txtStdDev_TextChanged(sender As Object, e As EventArgs) Handles txtStdDev.TextChanged

    End Sub

    Private Sub cbxSortSummaryAll_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSortSummaryAll.SelectedIndexChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Dim boolT As Boolean = boolFormLoad
        boolFormLoad = True
        Call SortSummaryAll()
        boolFormLoad = boolT

    End Sub

    Sub SortSummaryAll()

        Dim str1 As String
        Dim strS As String

        Dim dv As DataView = Me.dgvAllSummary.DataSource
        strS = ReturnSummarAllSort()

        dv.Sort = strS


    End Sub

    Function ReturnSummarAllSort() As String

        Dim str1 As String

        Dim selectedItem As Object
        selectedItem = Me.cbxSortSummaryAll.SelectedItem

        str1 = selectedItem.ToString ' Me.cbxSortSummaryAll.Text
        If InStr(1, str1, "Analyte", CompareMethod.Text) > 0 Then
            ReturnSummarAllSort = "ANALYTEDESCRIPTION ASC, CHARHEADINGTEXT ASC"
        Else
            ReturnSummarAllSort = "CHARHEADINGTEXT ASC, ANALYTEDESCRIPTION ASC"
        End If

    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        MsgBox(Me.cbxCL.SelectedIndex)


    End Sub


End Class