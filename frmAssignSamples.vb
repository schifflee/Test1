Option Compare Text

Public Class frmAssignSamples

    Public boolCancelButton As Boolean = False
    Public boolCancel As Boolean = False
    Public boolX As Boolean = False
    Public tblAnalytes As New System.Data.DataTable
    Public boolFormLoad As Boolean = True
    Public boolCont As Boolean = True
    Public lastAnalyteDesc As String
    Public tblAnalysisResults As New System.Data.DataTable 'use in Assign Samples
    Public boolFromdgvTable As Boolean = False
    Public strAnalFromTable As String = ""
    'Public cbxxHelper1 As New DataGridViewComboBoxCell
    Public tblHelper1 As New System.Data.DataTable
    Public tblNomConc As New System.Data.DataTable
    Public boolFromFilter As Boolean = False '
    Public boolFromClearFilters As Boolean = False
    Public boolViewOnly As Boolean = False
    Public tblStudiesA As System.Data.DataTable
    Public numOrigWatsonID As Int64
    Public boolHold As Boolean = False
    Public boolAnalOK(6, 500)
    '0=TRUE/FALSE,1=analytedescription, 2=analyteindex, 3=masterassayid, 4=ANALYTEID, 5=wStudyID, 6=StudyName
    Public ctAnalOK As Short = 0
    Public AnalIS(3) '1=analytedescription, 2=analyteindex, 3=masterassayid
    Public boolDoAccCrit As Boolean = False
    Public strSortISR As String
    Public strSortNonISR As String

    Public boolFromChangeStudy As Boolean = False

    Public booldgvReportTableCancel As Boolean = False

    Dim intNumMatrices As Short
    Dim gintGroup As Short = 0
    Dim gboolFiltersCleared As Boolean = False

    Dim boolFormLoad1 As Boolean = False

    Dim boolOriginal As Boolean = True

    Dim boolAutoAssign As Boolean = False 'True if AutoAssign action is in progress

    Public boolDontChange As Boolean = False

    Public nSDStudyID As Int32
    Public nWStudyID As Int32
    Public nStudyName As String


    'Note: Me.txtStudyID.Text is either id_tblStudies or the study that has been switched using cbxStudy

    Function ColumnOrder(strColumnName As String) As Short




    End Function

    ' Nick Addition
    Sub SetToEditMode()

        Me.cmdEdit.Enabled = False
        Me.cmdEdit.BackColor = System.Drawing.Color.Gray
        Me.cmdOK.Enabled = True
        Me.cmdOK.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdReset.Enabled = True
        Me.cmdReset.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.Enabled = False
        Me.cmdExit.BackColor = System.Drawing.Color.Gray

        Me.cmdAuto.Enabled = True
        Me.cmdAuto.BackColor = System.Drawing.Color.Gainsboro

    End Sub

    Sub SetToNonEditMode()

        Me.cmdEdit.Enabled = True
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOK.Enabled = False
        Me.cmdOK.BackColor = System.Drawing.Color.Gray
        Me.cmdReset.Enabled = False
        Me.cmdReset.BackColor = System.Drawing.Color.Gray
        Me.cmdExit.Enabled = True
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro

        Me.cmdAuto.Enabled = False
        Me.cmdAuto.BackColor = System.Drawing.Color.Gray

    End Sub


    Private Sub frmAssignSamples_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If boolX Then
            e.Cancel = True
        Else
            Call CloseWindow()
        End If

        'Try
        '    frmAnalyticalRunSummary.Close()
        'Catch ex As Exception

        'End Try



    End Sub

    Sub CountSamples()

        Dim var1, var2

        var1 = Me.dgvAnalyticalRuns.RowCount
        var2 = Format(var1, "#,##0")
        Me.txtSSNum.Text = var2
        Me.txtSSNum.Refresh()


    End Sub

    Public Sub PlaceControls(ByVal boolViewOnly As Boolean)

        Dim a, intPanAssSamplesTop, b, c, d, h, w
        Dim dgvAnalyticalRuns As DataGridView
        Dim dgvAssignedSamples As DataGridView
        Dim boolH As Boolean
        Dim IntDgvARwidth, IntDgvASwidth As Short
        Dim IntVertSBWidth As Short
        Dim intCols, Count1 As Short
        Dim var1

        IntVertSBWidth = 19 ' Set the vertical scrollbar width
        dgvAnalyticalRuns = Me.dgvAnalyticalRuns
        dgvAssignedSamples = Me.dgvAssignedSamples
        dgvAssignedSamples.ScrollBars = ScrollBars.Both

        '20160226 LEE: Align some stuff
        a = Me.lblAnalRuns.Left + Me.lblAnalRuns.Width + 5
        Me.grpBoxFilters.Left = a

        a = Me.grpBoxFilters.Left + Me.grpBoxFilters.Width + 5
        Me.panCbxStudy.Left = a

        'First, Make relevant controls visible if just viewing runs
        If boolviewonly Then  'Just View Runs

            Try

                boolH = boolHold
                'boolHold = True
                Me.chkSampleType.Checked = True
                Me.chkAssayLevel.Checked = True
                Me.chkDilFactor.Checked = True
                Me.chkAnalysisDate.Checked = True
                Me.chkFlag.Checked = True
                Me.chkAnalRT.Checked = True
                Me.chkISRT.Checked = True

                'Me.gbShowColumns.Visible = False
                'Me.panAccCrit.Visible = False
                'Me.dgvTables.Visible = False
                'Me.lblTables.Visible = False
                'Me.panAssSamples.Visible = False
                'Me.panLabels.Visible = False
                'Me.cmdOK.Visible = False
                'Me.cmdEdit.Visible = False
                'Me.cmdReset.Visible = False
                'Me.panCbxStudy.Visible = False

                Me.panAccCrit.Visible = False
                Me.panLabels.Visible = False
                Me.panCbxStudy.Visible = False
                Me.panExtra.Visible = False

                Me.gbShowColumns.Visible = False

                Me.cmdEdit.Visible = False
                Me.cmdOK.Visible = False
                Me.cmdReset.Visible = False
                Me.cmdAuto.Visible = False

                Me.dgvTables.Visible = False
                Me.lblTables.Visible = False
                Me.lblColored.Visible = False

                'boolHold = boolH

                'set dgvAnalyticalRuns h and w
                Dim fw, fh

                fw = Me.Width
                fh = Me.Height

                a = Me.panAssSamples.Top + Me.panAssSamples.Height
                b = Me.dgvAnalyticalRuns.Top
                c = Me.dgvAnalyticalRuns.Left
                Me.dgvAnalyticalRuns.Height = fh - b - 50
                Me.dgvAnalyticalRuns.Width = fw - c - 50


                ''Move Analytes control
                'lbl3.Left = dgvAnalyticalRuns.Left
                'lbl3.Top = 24
                'dgvAnalytes.Left = dgvAnalyticalRuns.Left
                'dgvAnalytes.Top = lbl3.Top + lbl3.Height
                'panEdit.Left = dgvAnalytes.Left + dgvAnalytes.Width + 10 - cmdExit.Left
                'panEdit.Top = lbl3.Top

            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

        Else
            ''Move Analytes control back to original positions
            'dgvAnalytes.SetBounds(681, 55, 189, 133)
            'lbl3.Left = dgvAnalytes.Left
            'panEdit.SetBounds(531, 6, 286, 25)

            Dim incr As Short = 5
            Me.dgvAnalytes.Left = Me.dgvTables.Left + Me.dgvTables.Width + incr
            lbl3.Left = dgvAnalytes.Left
            Me.panEdit.Left = (Me.dgvAnalytes.Left + Me.dgvAnalytes.Width) - Me.panEdit.Width


        End If

        'Calculate space we have to work with
        'h,w are height of the form
        h = Me.Height
        w = Me.Width


        If (Not (boolviewonly)) Then   'ASSIGN ANALYTICAL RUNS

            'HEIGHT Calculations (Assign Analytical Runs...)

            'Set a as distance from top of Analytical Runs Table to bottom of the frame.  Then set its height at 45% of that space.
            a = h - dgvAnalyticalRuns.Top
            dgvAnalyticalRuns.Height = a * 0.45

            Try

                'See Assigned Samples pane below Analytical runs
                Me.panAssSamples.Top = dgvAnalyticalRuns.Top + dgvAnalyticalRuns.Height + 10

                'Set Assigned Samples Panel height so that it reaches to within 50 pixels of bottom of frame
                c = h - Me.panAssSamples.Top
                Me.panAssSamples.Height = c - 50

                'Set Labels panel height such that it reaches to within 50 pixels of bottom of frame
                Me.panLabels.Top = h - Me.panLabels.Height - 50

                'WIDTH Calculations (Assign Analytical Runs...)

                'Calculate width of dgvAnalyticalRuns table
                intCols = dgvAnalyticalRuns.Columns.Count
                IntDgvARwidth = 0
                For Count1 = 0 To intCols - 1
                    If dgvAnalyticalRuns.Columns(Count1).Visible Then
                        IntDgvARwidth = IntDgvARwidth + dgvAnalyticalRuns.Columns(Count1).Width
                    End If
                Next
                IntDgvARwidth = IntDgvARwidth + IntVertSBWidth + dgvAnalyticalRuns.RowHeadersWidth

                'Find width of Assigned Samples grid
                intCols = dgvAssignedSamples.Columns.Count
                IntDgvASwidth = 0

                For Count1 = 0 To intCols - 1
                    If dgvAssignedSamples.Columns(Count1).Visible Then
                        IntDgvASwidth = IntDgvASwidth + dgvAssignedSamples.Columns(Count1).Width
                    End If
                Next
                IntDgvASwidth = IntDgvASwidth + IntVertSBWidth + dgvAssignedSamples.RowHeadersWidth

                'If Assigned Samples is not as wide as Analytical Runs, use Analytical Runs width
                If (IntDgvARwidth > IntDgvASwidth) Then
                    IntDgvASwidth = IntDgvARwidth
                End If

                Dim intPanLabelSpacing As Short
                intPanLabelSpacing = 10  'Space between the Assigned Samples table and the labels
                'If panLabels don't fit, put them at right of frame, and narrow the 2 tables
                If Me.panAssSamples.Left + IntDgvASwidth + Me.panLabels.Width + intPanLabelSpacing > w Then
                    Me.panLabels.Left = w - Me.panLabels.Width - Me.panAssSamples.Left - 25
                    dgvAnalyticalRuns.Width = Me.panLabels.Left - intPanLabelSpacing - dgvAnalyticalRuns.Left
                    Me.panAssSamples.Width = dgvAnalyticalRuns.Width
                Else
                    'If it's going to hit the "Sort by:" box, keep it away.
                    If d < Me.cbxSortAssigneSamples.Left + Me.cbxSortAssigneSamples.Width Then
                        d = Me.cbxSortAssigneSamples.Left + Me.cbxSortAssigneSamples.Width + intPanLabelSpacing
                    End If

                    'Otherwise, set the width, and set the panLabels to the right of the tables
                    dgvAnalyticalRuns.Width = IntDgvASwidth
                    Me.panAssSamples.Width = IntDgvASwidth
                    Me.panLabels.Left = Me.panAssSamples.Left + Me.panAssSamples.Width + 10
                End If

                'Align show columns with buttons above (panExtra)
                Me.gbShowColumns.Top = Me.panExtra.Top + Me.panExtra.Height
                Me.gbShowColumns.Left = Me.panExtra.Left

                'Align AcceptanceCriteria with the labels
                Me.panAccCrit.Top = Me.panLabels.Top - Me.panAccCrit.Height - 1
                Me.panAccCrit.Left = Me.panLabels.Left

                'If it impinges on 'Change Study' box, align it with Show columns instead
                Dim intLeftLimit As Short = Me.panCbxStudy.Left + Me.panCbxStudy.Width
                If (Me.panAccCrit.Top < panCbxStudy.Top + panCbxStudy.Height) Then
                    If (Me.panAccCrit.Left < intLeftLimit) Then
                        'Move beside the labels panel
                        Me.panAccCrit.Left = intLeftLimit + 5
                        Me.panAccCrit.Top = Me.panCbxStudy.Top 'Align with the study button (looks better)
                    End If
                End If

                'Set width and height of the assign samples table within the assigned samples panel
                dgvAssignedSamples.Width = Me.panAssSamples.Width
                dgvAssignedSamples.Height = Me.panAssSamples.Height - dgvAssignedSamples.Top - 1
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try


        Else 'VIEW RUNS

            'screw all this

            'Try
            '    boolHold = boolH
            '    Call ShowColumnsGroupBox()

            '    'WIDTH Calculation (View Runs)

            '    'Calculate width of dgvAnalyticalRuns table  (after running ShowColumnsGroupBox)
            '    intCols = dgvAnalyticalRuns.Columns.Count
            '    IntDgvARwidth = 0
            '    For Count1 = 0 To intCols - 1
            '        If dgvAnalyticalRuns.Columns(Count1).Visible Then
            '            IntDgvARwidth = IntDgvARwidth + dgvAnalyticalRuns.Columns(Count1).Width
            '        End If
            '    Next
            '    IntDgvARwidth = IntDgvARwidth + IntVertSBWidth + dgvAnalyticalRuns.RowHeadersWidth


            '    'If Analytical Runs Table is wider than screen, fit exactly on screen
            '    If (IntDgvARwidth + dgvAnalyticalRuns.Left < w) Then
            '        dgvAnalyticalRuns.Width = IntDgvARwidth
            '    Else
            '        dgvAnalyticalRuns.Width = w - dgvAnalyticalRuns.Left - 20
            '    End If

            '    'HEIGHT Calculations
            '    dgvAnalyticalRuns.Height = h - dgvAnalyticalRuns.Top - 60
            'Catch ex As Exception
            '    var1 = ex.Message
            '    var1 = var1
            'End Try

        End If

        Me.lblProgress.Top = Me.dgvAnalyticalRuns.Top
        Me.lblProgress.Height = Me.dgvAnalyticalRuns.Height
        Me.lblProgress.Left = Me.dgvAnalyticalRuns.Left
        Me.lblProgress.Width = Me.dgvAnalyticalRuns.Width

        Me.lblWait.Top = Me.lblProgress.Top
        Me.lblWait.Height = Me.lblProgress.Height
        Me.lblWait.Left = Me.lblProgress.Left
        Me.lblWait.Width = Me.lblProgress.Width

    End Sub

    Private Sub frmAssignSamples_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '20190304 LEE:
        'Begin allowing Regression Constants table (2) to have sample assignment

        Call ControlDefaults(Me)

        Call DoubleBufferControl(Me, "dgv")

        'double buffer all datagridviews in this form
        Dim strC As String
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            strC = Mid(ctrl.Name, 1, 3)
            If StrComp(strC, "dgv", CompareMethod.Text) = 0 Then
                Call modExtensionMethods.DoubleBufferedControl(ctrl, True)
            End If
        Next

        Dim int1 As Int64
        Dim strF As String
        int1 = tblAnalysisResultsHome.Rows.Count 'for debugging
        If int1 = 0 Then
            strF = "No Watson run data exists for this study."
            MsgBox(strF, MsgBoxStyle.Information, "Action terminated...")
            'Me.Visible = False
            'Exit Sub
            Call CloseWindow()
        End If
        Dim dt As Date
        Dim int2 As Int64
        Dim bool As Boolean
        Dim boolA As Short
        Dim frm As New frmErrorMsg
        Dim strM As String
        Dim ctP As Int64
        Dim dgv As DataGridView
        Dim var1
        Dim intRows As Int64
        Dim intRow As Int64

        Dim str1, str2 As String
        Dim w, h, w1, h1

        'Note: For using a network service account, look in help for networkcredential

        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        Me.Left = 0
        Me.Top = 0

        Me.Width = w '* 0.9
        Me.Height = h ' * 0.9
        'Me.Left = (w - Me.Width) / 2
        'Me.Top = (h - Me.Height) / 2

        'Me.WindowState = FormWindowState.Maximized

        Cursor.Current = Cursors.WaitCursor

        frm.cmdOK.Visible = False
        If boolViewOnly Then
            frm.lblErr.Text = "Opening View Analytical Runs window..."
            frm.Text = "Opening View Analytical Runs..."
        Else
            frm.lblErr.Text = "Opening Assigned Samples window..."
            frm.Text = "Opening Assigned Samples..."
        End If
        ctP = 1
        frm.pb1.Value = ctP
        frm.pb1.Maximum = 9
        frm.pb1.Visible = True
        frm.Show()
        frm.Refresh()

        Me.cbxFilterRunID.Items.Add("[None]")
        Me.cbxFilterSampleType.Items.Add("[None]")
        Me.cbxFilterSampleType.Items.Add("QC")
        Me.cbxFilterSampleType.Items.Add("STANDARD")

        'intialize some dgv parameters
        dgv = Me.dgvAnalyticalRuns
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        '20181108 LEE
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        dgv = Me.dgvAssignedSamples
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        '20181108 LEE
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        'dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.None)
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Me.lblWait.Left = 314
        Me.lblWait.Top = 197

        'do this first

        Call FillAccStatus()

        Dim boolX As Boolean = False
        'If boolViewOnly Then
        If boolX Then

            Try
                'populate cbxStudy
                strM = "Retrieving Studies"
                frm.lblErr.Text = "Opening View Analytical Runs..." ' & strM

                Me.tblAnalysisResults = tblAnalysisResultsHome.Copy

                '20160822 LEE:
                'set to case sensitive
                Me.tblAnalysisResults.CaseSensitive = True


                Me.panExtra.Visible = False

                'Dim count1 As Short
                '''''''''''''console.writeline("Start tblAnalysisResults columns...")
                'For count1 = 0 To Me.tblAnalysisResults.Columns.Count - 1
                '    ''''''''''''console.writeline(Me.tblAnalysisResults.Columns(count1).ColumnName)
                'Next
                '''''''''''''console.writeline("End tblAnalysisResults columns...")

                int2 = Me.tblAnalysisResults.Rows.Count 'for debugging

                boolFormLoad = True
                boolFromdgvTable = False

                'Me.dgvAssignedSamples.VirtualMode = True
                Me.rbFilterForAnalyteYes.Checked = True

                Call tblAnalytesConfigure(False)
                Call InitializedgvAnalytes()

                Call FillAnalyticalRuns("")
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try


        Else

            Me.panExtra.Visible = True

            Me.tblHelper1 = tblAssignedSamplesHelper.Copy

            ''debug
            'Dim Count1 As Int16
            'For Count1 = 0 To tblAssignedSamplesHelper.Rows.Count - 1
            '    str1 = tblAssignedSamplesHelper.Rows(Count1).Item("CHARHELPER")
            '    ''Console.WriteLine(str1)
            'Next

            Call InitializeHelper1()
            Call InitializeNomConc()
            'fill cbxSort
            Me.cbxSortAssigneSamples.Items.Add("Original")
            Me.cbxSortAssigneSamples.Items.Add("Level")
            Me.cbxSortAssigneSamples.SelectedIndex = 0

            'Me.lbldgvNomConc.Text = "Assign" & ChrW(10) & "Nom. Conc."
            Me.lbldgvNomConc.Text = "Assign Nom. Conc."

            'position dvgHelper2
            var1 = Me.txtHelper2.Left
            Me.dgvHelper2.Left = var1
            Me.lbldgvHelper2.Left = var1

            Me.tblAnalysisResults = tblAnalysisResultsHome.Copy
            int2 = Me.tblAnalysisResults.Rows.Count 'for debugging

            ''debugging
            'Dim rowsAAA() As DataRow = Me.tblAnalysisResults.Select("RUNID = 4")
            'var1 = rowsAAA.Length


            boolFormLoad = True
            boolFromdgvTable = False

            'Me.dgvAssignedSamples.VirtualMode = True
            Me.rbFilterForAnalyteYes.Checked = True

            'populate cbxStudy
            strM = "Retrieving Studies"
            dt = Now
            '''''''''''''''''''console.writeline(strM & ": " & CStr(dt))
            frm.lblErr.Text = "Opening Assigned Samples..." & strM
            ctP = ctP + 1
            frm.pb1.Value = ctP
            frm.Refresh()
            Cursor.Current = Cursors.WaitCursor
            Call FillcbxStudy()
            Me.cbxStudy.DropDownWidth = Me.cbxStudy.Width ' * 1.5

            'select cbxvalue
            strM = "Retrieving Study Data"
            dt = Now
            ' '''''''''''''''''''console.writeline(strM & ": " & CStr(dt))
            'frm.lblErr.Text = "Opening Assigned Samples..." & strM
            'ctP = ctP + 1
            'frm.pb1.Value = ctP
            'frm.Refresh()
            'Cursor.Current = Cursors.WaitCursor
            'Call ReturnStudyToOriginal()

            'configure tables
            strM = "Populating Tables"
            dt = Now
            '''''''''''''''''''console.writeline(strM & ": " & CStr(dt))
            frm.lblErr.Text = "Opening Assigned Samples..." & strM
            ctP = ctP + 1
            frm.pb1.Value = ctP
            frm.Refresh()
            Cursor.Current = Cursors.WaitCursor

            Call tblAnalytesConfigure(False)
            Call InitializedgvAnalytes()

            frm.lblErr.Text = "Opening Assigned Samples..." & strM
            ctP = ctP + 1
            frm.pb1.Value = ctP
            frm.Refresh()
            Cursor.Current = Cursors.WaitCursor
            Call ReturnStudyToOriginal()

            'configure dgvTables
            strM = "Populating Grids"
            dt = Now
            '''''''''''''''''''console.writeline(strM & ": " & CStr(dt))
            frm.lblErr.Text = "Opening Assigned Samples..." & strM
            ctP = ctP + 1
            frm.pb1.Value = ctP
            frm.Refresh()
            Cursor.Current = Cursors.WaitCursor
            Call FilldgvTables()

            strM = "Initializing Assigned Samples Table"
            dt = Now
            '''''''''''''''''''console.writeline(strM & ": " & CStr(dt))
            frm.lblErr.Text = "Opening Assigned Samples..." & strM
            ctP = ctP + 1
            frm.pb1.Value = ctP
            frm.Refresh()
            Cursor.Current = Cursors.WaitCursor
            Call InitializeAssignedSamples()

            strM = "Populating Assigned Samples Table"
            dt = Now
            '''''''''''''''''''console.writeline(strM & ": " & CStr(dt))
            frm.lblErr.Text = "Opening Assigned Samples..." & strM
            ctP = ctP + 1
            frm.pb1.Value = ctP
            frm.Refresh()
            Cursor.Current = Cursors.WaitCursor
            Call FillAssignedSamples()

            strM = "Applying Visual Aids"
            dt = Now
            '''''''''''''''''''console.writeline(strM & ": " & CStr(dt))
            frm.lblErr.Text = "Opening Assigned Samples..." & strM
            ctP = ctP + 1
            frm.pb1.Value = ctP
            frm.Refresh()
            Cursor.Current = Cursors.WaitCursor

            'do final stuff
            Call AssessSampleAssignment()
            Call FilterHelper1()
            Call ShowColumns()
            'Call InitializeFilterRunID()
            'Call FilterSAMPLETYPE()

            Cursor.Current = Cursors.WaitCursor

            If Me.dgvAnalytes.RowCount = 0 Then
                lastAnalyteDesc = ""
            Else
                lastAnalyteDesc = Me.dgvAnalytes.Rows.Item(0).Cells(0).Value
            End If

            Cursor.Current = Cursors.WaitCursor
            Call ASNum()

            Call NomConcFill(False)


        End If

        Dim intRowA As Short
        Dim boolContA As Boolean = False

        If Me.dgvAnalytes.RowCount = 0 Then
        ElseIf Me.dgvAnalytes.CurrentRow Is Nothing Then
            boolContA = True
            intRowA = 0
        Else
            boolContA = True
            intRowA = Me.dgvAnalytes.CurrentRow.Index
        End If

        If boolContA Then
            Call InitializeFilterRunID(intRowA)
        End If
        Call FilterSAMPLETYPE()


        boolFormLoad = False

        Cursor.Current = Cursors.Default

        frm.Visible = False
        frm.Dispose()

        'lock up form if appropriate
        If boolViewOnly Then
            SetToNonEditMode()
        Else

            boolA = BOOLREPORTTABLECONFIGURATION
            If boolA = 0 Then
                bool = False
            Else
                bool = True
            End If
            If bool = True Then
                Call SetToNonEditMode()
            Else
                Call SetToEditMode()
            End If
        End If


        'SendKeys.Send("%")
        If boolViewOnly Then
            Me.dgvAnalyticalRuns.Focus()
        Else
            Me.dgvTables.Focus()
            Dim id As Int64

            dgv = Me.dgvTables
            If dgv.Rows.Count = 0 Then
            Else
                If dgv.CurrentRow Is Nothing Then
                Else
                    intRow = dgv.CurrentRow.Index
                    id = NZ(dgv("ID_TBLCONFIGREPORTTABLES", CInt(intRow)).Value, 1)
                    Call IncSampleVis(id)

                End If
            End If

            boolFormLoad1 = True
            Call ChangedgvTables()
            boolFormLoad1 = False

        End If

        'clear boolAnalOK
        'arrAnalytes.Clear(arrAnalytes, 0, arrAnalytes.Length)'example

        boolAnalOK.Clear(boolAnalOK, 0, boolAnalOK.Length)

        Call CountSamples()

        Call PlaceControls(boolViewOnly)

        Call LockAssignedSamples(True)

        Call frmAssignSamples_ToolTipSet()

        If gAllowGuWuAccCrit And LAllowGuWuAccCrit Then
        Else
            Me.panAccCrit.Visible = False
        End If

        Call ChooseAnalyte()

        'Me.lblProgress.Left = Me.dgvAnalyticalRuns.Left
        'Me.lblProgress.Top = Me.dgvAnalyticalRuns.Top
        'Me.lblProgress.Width = Me.dgvAnalyticalRuns.Width
        'Me.lblProgress.Height = Me.dgvAnalyticalRuns.Height

        Me.Refresh()

        boolFormLoad = False

        ''pesky
        'boolFormLoad = True
        'Try
        '    Call InitializeAnalyticalRuns()
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try
        'boolFormLoad = False

        Me.lblWait.Visible = False

    End Sub

    Sub InitializeFilterRunID(ByVal intRow As Short)

        'intRow is the dgvAnalyte row
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim Count1 As Short
        Dim var1
        Dim boolT As Boolean
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strF As String
        Dim varVal
        Dim boolIS As Boolean

        If boolFormLoad Then
        Else
            varVal = Me.cbxFilterRunID.SelectedItem
        End If

        'if Analyte is IntStd, then ignore
        str1 = NZ(Me.dgvAnalytes("IsIntStd", intRow).Value, "Yes")
        If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
            boolIS = True
        Else
            boolIS = False
        End If

        Dim id As Int64
        id = GetWStudyID(NZ(CLng(Me.txtStudyID.Text), 0))

        dgv1 = Me.dgvAnalytes
        'str1 = CStr(dgv1("ANALYTEINDEX", intRow).Value)
        'str2 = CStr(dgv1("MASTERASSAYID", intRow).Value)
        'str3 = CStr(dgv1("ANALYTEID", intRow).Value)

        'strF = "ANALYTEINDEX = " & str1 & " AND MASTERASSAYID = " & str2 & " AND STUDYID = " & id & " AND ANALYTEID = " & str3
        'dv = New DataView(Me.tblAnalysisResults, strF, "MASTERASSAYID ASC", DataViewRowState.CurrentRows)

        'strF = "ANALYTEINDEX = " & str1 & " AND STUDYID = " & id & " AND ANALYTEID = " & str3
        'dv = New DataView(Me.tblAnalysisResults, strF, "RUNID ASC", DataViewRowState.CurrentRows)

        'get runs from tblCalStdGroupAssayIDsAll or tblCalStdGroupAssayIDsAcc
        Dim intGroup As Short
        Dim intAnalyteID As Int64
        Dim strAnal As String

        intGroup = NZ(Me.dgvAnalytes("INTGROUP", intRow).Value, -1)
        strAnal = NZ(Me.dgvAnalytes("ANALYTEDESCRIPTION", intRow).Value, -1)

        If boolIS Then
            strF = "INTSTD = '" & CleanText(strAnal) & "'"
        Else
            strF = "INTGROUP = " & intGroup
        End If

        str1 = Me.cbxAccStatus.Text
        If boolFromChangeStudy Or IsStudyChanged() Then
            'get all runs from changed study
            dv = Me.dgvAnalyticalRuns.DataSource

        Else

            '20171218 LEE: The following code is confusing
            'If StrComp(str1, "Not Rejected", CompareMethod.Text) = 0 Then

            '    'add to strf
            '    ''RUNANALYTEREGRESSIONSTATUS =1 or 2 or 3
            '    If StrComp(str1, "Show All", CompareMethod.Text) = 0 Then
            '        strF = str1 ' strF & " AND (RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3)"
            '    ElseIf StrComp(str1, "Not Rejected", CompareMethod.Text) = 0 Then
            '        strF = strF & " AND ((RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3) AND RUNTYPEID <> 3)"
            '    Else
            '        strF = strF & " AND (RUNANALYTEREGRESSIONSTATUS = 3 AND RUNTYPEID <> 3)"
            '    End If
            '    'dv = New DataView(tblCalStdGroupAssayIDsAcc, strF, "RUNID ASC", DataViewRowState.CurrentRows)
            '    dv = New DataView(tblCalStdGroupAssayIDsAll, strF, "RUNID ASC", DataViewRowState.CurrentRows)
            'Else
            '    dv = New DataView(tblCalStdGroupAssayIDsAll, strF, "RUNID ASC", DataViewRowState.CurrentRows)
            'End If

            '20171218 LEE: Cleaned up

            If StrComp(str1, "Show All", CompareMethod.Text) = 0 Then

                '20171219 LEE: tblCalStdGroupAssayIDsAll used to exclude runs with no calibr curve
                'modGroups was modified to ensure tblCalStdGroupAssayIDsAll has ALL analytical runs

                dv = New DataView(tblCalStdGroupAssayIDsAll, strF, "RUNID ASC", DataViewRowState.CurrentRows)

            ElseIf StrComp(str1, "Not Rejected", CompareMethod.Text) = 0 Then
                strF = strF & " AND ((RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3) AND RUNTYPEID <> 3)"
                dv = New DataView(tblCalStdGroupAssayIDsAll, strF, "RUNID ASC", DataViewRowState.CurrentRows)
            ElseIf StrComp(str1, "Accepted", CompareMethod.Text) = 0 Then
                strF = strF & " AND (RUNANALYTEREGRESSIONSTATUS = 3 AND RUNTYPEID <> 3)"
                dv = New DataView(tblCalStdGroupAssayIDsAll, strF, "RUNID ASC", DataViewRowState.CurrentRows)
            End If

        End If


        'dv = Me.dgvAnalyticalRuns.DataSource

        'dgv = Me.dgvAnalyticalRuns
        'dv = dgv.DataSource
        Dim tbl As System.Data.DataTable = dv.ToTable("a", True, "RUNID")
        int1 = tbl.Rows.Count

        boolT = boolHold
        boolHold = True

        Me.cbxFilterRunID.Items.Clear()
        Me.cbxFilterRunID.Items.Add("[None]")
        'NOTE: item "[None]" added at the beginning of formload
        For Count1 = 0 To int1 - 1
            var1 = tbl.Rows.Item(Count1).Item("RUNID")
            Me.cbxFilterRunID.Items.Add(CStr(var1))
        Next

        'select first item
        Me.cbxFilterRunID.SelectedIndex = 0
        If boolFormLoad Then
        Else
            'try to find varVal
            For Count1 = 0 To Me.cbxFilterRunID.Items.Count - 1
                var1 = Me.cbxFilterRunID.Items(Count1)
                If var1 = varVal Then
                    Me.cbxFilterRunID.SelectedIndex = Count1
                    Exit For
                End If
            Next

        End If

        boolHold = boolT


    End Sub

    Sub InitializeNomConc()

        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim intRows As Short
        Dim intCols As Short
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim Count1 As Short
        Dim strS As String
        Dim var1, var2

        'tbl = tblBCQCs.Copy
        'dv = tbl.DefaultView
        dv1 = tblBCQCs.DefaultView

        Me.tblNomConc = dv1.ToTable("a", True, "CONCENTRATION")
        tbl = Me.tblNomConc

        'add column to tbl
        Dim col As New DataColumn
        col.ColumnName = "DEC"
        col.DataType = System.Type.GetType("System.Decimal")
        'col.DataType = System.Type.GetType("System.Double")
        tbl.Columns.Add(col)

        intRows = tbl.Rows.Count
        intCols = tbl.Columns.Count

        'contents will come over as string
        'convert contents to numeric
        For Count1 = 0 To intRows - 1
            var1 = tbl.Rows.Item(Count1).Item("CONCENTRATION")
            var2 = CDec(var1)
            tbl.Rows.Item(Count1).BeginEdit()
            tbl.Rows.Item(Count1).Item("DEC") = var2
            tbl.Rows.Item(Count1).EndEdit()
        Next

        strS = "DEC ASC"

        dgv = Me.dgvNomConc
        dgv.AllowUserToResizeRows = True
        dgv.AllowUserToResizeColumns = True
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.RowHeadersVisible = False
        dgv.ColumnHeadersVisible = False

        dv = New DataView(tbl)

        dv.AllowNew = False
        dv.AllowDelete = False
        dv.Sort = strS
        dgv.DataSource = dv
        intCols = dgv.Columns.Count

        For Count1 = 0 To intCols - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Next

        dgv.Columns.Item("DEC").Visible = True
        dgv.Columns.Item("DEC").DisplayIndex = 0
        var1 = dgv.Width
        dgv.Columns.Item("DEC").Width = var1 * 0.7

    End Sub

    Sub NomConcFill(ByVal boolFromOpt As Boolean)

        'boolFromOpt: comes from optQCConcs or optCalibrConcs

        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim intRows As Short
        Dim intCols As Short
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim strS As String
        Dim strF As String
        Dim var1, var2, var3, var4, var5
        Dim intID As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim dv1 As System.Data.DataView
        Dim boolIS As Boolean

        If boolFormLoad Then
            'Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Try

            'record id of selected table
            int1 = Me.dgvTables.CurrentRow.Index
            intID = Me.dgvTables("ID_TBLCONFIGREPORTTABLES", int1).Value

            'record analyte variables
            If Me.dgvAnalytes.RowCount = 0 Then
                Exit Sub
            End If

            int1 = Me.dgvAnalytes.CurrentRow.Index
            var1 = Me.dgvAnalytes("ANALYTEINDEX", int1).Value
            var2 = Me.dgvAnalytes("MASTERASSAYID", int1).Value
            var3 = Me.dgvAnalytes("AnalyteDescription", int1).Value

            var5 = Me.dgvAnalytes("IsIntStd", int1).Value

            If StrComp(var5, "No", CompareMethod.Text) = 0 Then
                boolIS = False
            Else
                boolIS = True
            End If

            If boolIS Then
                'find intstd in dgv
                Dim dvIS1 As DataView = Me.dgvAnalytes.DataSource
                Dim tblIS1 As DataTable = dvIS1.ToTable
                strF = "IntStd = '" & CleanText(CStr(var3)) & "'"
                Dim rowsIS() As DataRow = tblIS1.Select(strF)
                If rowsIS.Length = 0 Then
                    var4 = Me.dgvAnalytes("ANALYTEID", 0).Value
                Else
                    var4 = rowsIS(0).Item("ANALYTEID")
                End If
            Else
                var4 = Me.dgvAnalytes("ANALYTEID", int1).Value
            End If

            ''debugging
            Dim Count2 As Integer

            'dv1 = New DataView(tblBCStds) 'must assign dv1 first
            'strF = "ANALYTEINDEX > -1"
            'dv1.RowFilter = strF

            If boolFromOpt Then
                If Me.optQCConcs.Checked Then
                    dv1 = New DataView(tblBCQCs)
                    If Me.chkShowAllNomConc.Checked Then
                        strF = "ANALYTEID > -1"
                    Else
                        strF = "ANALYTEID = " & var4
                    End If
                Else
                    dv1 = New DataView(tblBCStds)
                    If Me.chkShowAllNomConc.Checked Then
                        strF = "ANALYTEID > -1"
                    Else
                        strF = "ANALYTEID = " & var4
                    End If
                End If
            Else
                Select Case intID
                    Case 23, 3, 34, 35, 36

                        '3: Summary of Back-Calculated Calibration Std Conc
                        '23: [Period Temp] Spiking Solution Stability Assessment
                        '34: Selectivity in Individual Lots Table v1
                        '35: Carryover in Individual Lots Table v1
                        '36: Method Trial Back-Calculated Calibration Std Conc v1


                        'strF = "ANALYTEINDEX > -1"
                        'dv1 = New DataView(tblBCStds, strF, "ANALYTEINDEX ASC", DataViewRowState.CurrentRows)
                        dv1 = New DataView(tblBCStds)
                        If Me.chkShowAllNomConc.Checked Then
                            strF = "ANALYTEID > -1"
                        Else
                            strF = "ANALYTEID = " & var4
                        End If
                        boolHold = True
                        Me.optCalibrConcs.Checked = True
                        boolHold = False
                    Case Else
                        dv1 = New DataView(tblBCQCs)
                        If Me.chkShowAllNomConc.Checked Then
                            strF = "ANALYTEID > -1"
                        Else
                            strF = "ANALYTEID = " & var4
                        End If
                        boolHold = True
                        Me.optQCConcs.Checked = True
                        boolHold = False
                End Select

            End If

            'If boolFromOpt Then
            '    If Me.optQCConcs.Checked Then
            '        dv1 = New DataView(tblBCQCs)
            '        If Me.chkShowAllNomConc.Checked Then
            '        Else
            '            strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 ' & " AND ANALYTEID = " & var4
            '        End If
            '    Else
            '        dv1 = New DataView(tblBCStds)
            '        If Me.chkShowAllNomConc.Checked Then
            '        Else
            '            strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 & " AND ANALYTEDESCRIPTION = '" & var3 & "'"
            '        End If
            '    End If
            'Else
            '    Select Case intID
            '        Case 23, 3, 34, 35, 36
            '            'strF = "ANALYTEINDEX > -1"
            '            'dv1 = New DataView(tblBCStds, strF, "ANALYTEINDEX ASC", DataViewRowState.CurrentRows)
            '            dv1 = New DataView(tblBCStds)
            '            If Me.chkShowAllNomConc.Checked Then
            '            Else
            '                strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 & " AND ANALYTEDESCRIPTION = '" & var3 & "'"
            '            End If
            '            boolHold = True
            '            Me.optCalibrConcs.Checked = True
            '            boolHold = False
            '        Case Else
            '            dv1 = New DataView(tblBCQCs)
            '            If Me.chkShowAllNomConc.Checked Then
            '            Else
            '                strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 ' & " AND ANALYTEID = " & var4
            '            End If
            '            boolHold = True
            '            Me.optQCConcs.Checked = True
            '            boolHold = False
            '    End Select

            'End If

            int1 = dv1.Count

            dv1.RowFilter = strF
            int2 = dv1.Count 'debugging
            Me.tblNomConc = dv1.ToTable("a", True, "CONCENTRATION")
            tbl = Me.tblNomConc
            'add column to tbl
            If tbl.Columns.Contains("DEC") Then
            Else
                Dim col As New DataColumn
                col.ColumnName = "DEC"
                col.DataType = System.Type.GetType("System.Decimal")
                'col.DataType = System.Type.GetType("System.Double")
                tbl.Columns.Add(col)
            End If

            intRows = tbl.Rows.Count
            intCols = tbl.Columns.Count
            'contents will come over as string
            'convert contents to numeric
            For Count1 = 0 To intRows - 1
                var1 = tbl.Rows.Item(Count1).Item("CONCENTRATION")
                var2 = CDec(var1)
                tbl.Rows.Item(Count1).BeginEdit()
                tbl.Rows.Item(Count1).Item("DEC") = var2
                tbl.Rows.Item(Count1).EndEdit()
            Next
            strS = "DEC ASC"
            dgv = Me.dgvNomConc
            dv = New DataView(tbl)
            dv.AllowNew = False
            dv.AllowDelete = False
            dv.Sort = strS
            dgv.DataSource = dv

        Catch ex As Exception
            var1 = ""
        End Try


    End Sub

    Sub InitializeHelper1()

        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim intRows As Short
        Dim intCols As Short
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim strS As String
        Dim Count2 As Short


        'initializes Helper2 also

        'ID_TBLASSIGNEDSAMPLESHELPER
        'CHARHELPER
        'ID_TBLCONFIGREPORTTABLES
        'NUMCOMPANY

        tbl = Me.tblHelper1
        intRows = tbl.Rows.Count
        intCols = tbl.Columns.Count
        strS = "CHARHELPER ASC"

        For Count2 = 1 To 2
            Select Case Count2
                Case 1
                    dgv = Me.dgvHelper1
                Case 2
                    dgv = Me.dgvHelper2
            End Select
            dgv.AllowUserToResizeRows = True
            dgv.AllowUserToResizeColumns = True
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.RowHeadersVisible = False
            dgv.ColumnHeadersVisible = False

            dv = New DataView(tbl)

            dv.AllowNew = False
            dv.AllowDelete = False
            dv.Sort = strS
            dgv.DataSource = dv
            intCols = dgv.Columns.Count

            For Count1 = 0 To intCols - 1
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            Next

            'ID_TBLASSIGNEDSAMPLESHELPER
            'CHARHELPER
            'ID_TBLCONFIGREPORTTABLES
            'NUMCOMPANY

            dgv.Columns.Item("CHARHELPER").Visible = True
            dgv.Columns.Item("CHARHELPER").DisplayIndex = 0
            'dgv.Columns.item("CHARHELPER").HeaderText = "Term 1"
        Next



    End Sub

    Sub FilterHelper1()

        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim str1 As String
        Dim rows() As DataRow
        Dim intRow As Short
        Dim intID As Long
        'Dim dv as system.data.dataview
        'Dim dv1 as system.data.dataview
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim int1 As Short

        Dim boolHelper1 As Boolean
        Dim boolDHelper2 As Boolean
        Dim boolTHelper2 As Boolean
        Dim boolNomCon As Boolean
        Dim boolOpt As Boolean
        Dim strH1 As String
        Dim strH2 As String
        Dim strNC As String
        Dim strOpt As String
        Dim boolChkUseWat As Boolean
        Dim var1
        Dim tblT1 As New System.Data.DataTable

        Dim idRT As Int32
        Dim intBSNR As Short

        'add CHARANALYTE
        Dim col1 As New DataColumn
        col1.ColumnName = "CHARHELPER"
        col1.Caption = "CHARHELPER"
        tblT1.Columns.Add(col1)

        Dim dvT As system.data.dataview = New DataView(tblBCQCs)
        Dim tblT As System.Data.DataTable = dvT.ToTable("a", True, "ID")
        Dim intLabels As Short
        intLabels = tblT.Rows.Count
        For Count1 = 1 To intLabels
            var1 = tblT.Rows(Count1 - 1).Item("ID")
            Dim nr As DataRow = tblT1.NewRow
            nr.BeginEdit()
            nr.Item("CHARHELPER") = var1
            nr.EndEdit()
            tblT1.Rows.Add(nr)
        Next
        boolChkUseWat = False

        'intRow = Me.dgvTables.CurrentRow.Index

        If Me.dgvTables.CurrentRow Is Nothing Then
            Exit Sub
        End If

        Me.panLabels.Visible = False

        intRow = Me.dgvTables.CurrentRow.Index

        intID = Me.dgvTables("ID_TBLCONFIGREPORTTABLES", intRow).Value
        idRT = Me.dgvTables("ID_TBLREPORTTABLE", intRow).Value


        '20180803 LEE:
        Select Case intID
            Case 17
                'must evaluate for Matrix Effect


                Dim dgvT As DataGridView = Me.dgvTables
                Dim intRowT As Int16
                Dim rowsTP() As DataRow
                Dim intTP As Short

                intRowT = dgvT.CurrentRow.Index
                idRT = NZ(dgvT("ID_TBLREPORTTABLE", intRowT).Value, -1)
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idRT
                rowsTP = tblTableProperties.Select(strF)

                If rowsTP.Length = 0 Then
                Else
                    var1 = NZ(rowsTP(0).Item("BOOLDOINDREC"), 0)
                    'update values
                    Dim strF1 As String
                    strF1 = "ID_TBLCONFIGREPORTTABLES = 17 AND NUMHELPERNUMBER = 1"
                    Dim rowsASH() As DataRow

                    Try
                        rowsASH = Me.tblHelper1.Select(strF1, "ID_TBLASSIGNEDSAMPLESHELPER ASC", DataViewRowState.CurrentRows)
                    Catch ex As Exception
                        Dim var2
                        var2 = ex.Message
                        var2 = var2
                    End Try

                    For Count1 = 0 To rowsASH.Length - 1
                        If var1 = 0 Then
                            str1 = "Lot " & Count1 + 1
                        Else
                            If Count1 = 0 Then
                                str1 = "Solvent"
                            Else
                                str1 = "Lot " & Count1
                            End If
                        End If
                        rowsASH(Count1).BeginEdit()
                        rowsASH(Count1).Item("CHARHELPER") = str1
                        rowsASH(Count1).EndEdit()
                    Next
                End If

            Case 35 'Carryover
                Dim vGo = GetTableProp("boolIncludePSAE")
                If IsNumeric(vGo) Then 'this is ULOQ column in Carryover 35

                    Dim strF1 As String
                    strF1 = "ID_TBLCONFIGREPORTTABLES = 35 AND CHARHELPER = 'ULOQ'"
                    Dim rowsASH() As DataRow

                    Try
                        rowsASH = Me.tblHelper1.Select(strF1, "ID_TBLASSIGNEDSAMPLESHELPER ASC", DataViewRowState.CurrentRows)
                    Catch ex As Exception
                        Dim var2
                        var2 = ex.Message
                        var2 = var2
                    End Try

                    '20180810 LEE:
                    'Remember, boolIncludePSAE in NOT

                    Try
                        If vGo = 0 Then 'this mean ULOQ to be included
                            rowsASH(0).BeginEdit()
                            rowsASH(0).Item("NUMHELPERNUMBER") = 1
                            rowsASH(0).EndEdit()
                        Else
                            rowsASH(0).BeginEdit()
                            rowsASH(0).Item("NUMHELPERNUMBER") = 10
                            rowsASH(0).EndEdit()
                        End If
                    Catch ex As Exception
                        var1 = var1
                    End Try

                End If

        End Select

        dgv = Me.dgvHelper1
        dgv1 = Me.dgvHelper2

        'ID_TBLASSIGNEDSAMPLESHELPER
        'CHARHELPER
        'ID_TBLCONFIGREPORTTABLES
        'NUMCOMPANY

        'boolStopCBX = True


        'retrieve report types - 1
        tbl = Me.tblHelper1
        strF = "ID_TBLCONFIGREPORTTABLES = " & intID & " AND NUMHELPERNUMBER = 1"
        strS = "ID_TBLASSIGNEDSAMPLESHELPER ASC"
        'Dim dv as system.data.dataview = New DataView(tbl)
        Dim dv As system.data.dataview
        If Me.chkUseWatson.Checked Then
            dv = New DataView(tblT1)
            dv.Sort = "CHARHELPER ASC"
        Else
            dv = New DataView(tbl)
            dv.RowFilter = strF
            int1 = dv.Count 'debug
            dv.Sort = strS

        End If
        dgv.DataSource = dv
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'ID_TBLASSIGNEDSAMPLESHELPER
        'CHARHELPER
        'ID_TBLCONFIGREPORTTABLES
        'NUMCOMPANY

        dgv.Columns.Item("CHARHELPER").Visible = True
        dgv.Columns.Item("CHARHELPER").DisplayIndex = 0


        '20181127 LEE:
        'Must do special evaluation for 31 and 32
        Select Case intID
            Case 31, 32
                intBSNR = GetStatsNR(idRT)
                If intBSNR = 8 Or intBSNR = 9 Then
                    dgv.Visible = False
                    Me.lbldgvHelper1.Visible = False
                    Me.cmdHelper1.Visible = False
                    Me.chkUseWatson.Visible = False
                End If
            Case 12, 19, 21
                dgv.Visible = True
                Me.lbldgvHelper1.Visible = True
                Me.cmdHelper1.Visible = True
                Me.chkUseWatson.Visible = True
            Case Else
                intBSNR = 1
        End Select

        If dv.Count = 0 Or (intBSNR = 8 Or intBSNR = 9) Then
            dgv.Visible = False
            Me.lbldgvHelper1.Visible = False
            Me.cmdHelper1.Visible = False
            Me.chkUseWatson.Visible = False
        Else
            dgv.Visible = True
            Me.lbldgvHelper1.Visible = True
            Me.cmdHelper1.Visible = True
            Me.chkUseWatson.Visible = True

            dv.AllowDelete = False
            dv.AllowNew = False
            dgv.DataSource = dv
            dgv.AutoResizeColumns()
        End If

        'retrieve report types - 2
        strF = "ID_TBLCONFIGREPORTTABLES = " & intID & " AND NUMHELPERNUMBER = 2"
        strS = "ID_TBLASSIGNEDSAMPLESHELPER ASC"
        Dim dv1 As system.data.dataview = New DataView(tbl)
        dv1.RowFilter = strF
        dv1.Sort = strS
        If dv1.Count = 0 Then
            dgv1.Visible = False
            Me.lbldgvHelper2.Visible = False
            Me.cmdHelper2.Visible = False
        Else
            dgv1.Visible = True
            Me.lbldgvHelper2.Visible = True
            Me.cmdHelper2.Visible = True

            dv1.AllowDelete = False
            dv1.AllowNew = False
            dgv1.DataSource = dv1
            dgv1.AutoResizeColumns()
        End If

        'do panAccCrit
        If gAllowGuWuAccCrit And LAllowGuWuAccCrit Then
            Select Case intID
                Case 1, 2, 5, 6, 7, 13, 14, 15, 22, 23, 30, 33, 34, 35, 37, 38
                    Me.panAccCrit.Visible = False
                Case Else
                    Me.panAccCrit.Visible = True
            End Select
        Else
            Me.panAccCrit.Visible = False
        End If


        boolHelper1 = False
        boolDHelper2 = False
        boolTHelper2 = False
        boolNomCon = False
        boolOpt = False
        strH1 = ""
        strH2 = ""
        strNC = "Assign Nom. Conc."
        strOpt = ""

        Me.panLabels.Visible = True

        Dim boolMakeVis As Boolean
        boolMakeVis = True
        Select Case intID

            'Case 1 'Summary of Analytical Runs
            Case 1, 2, 5, 30, 33, 38
                boolHelper1 = False
                boolDHelper2 = False
                boolTHelper2 = False
                boolNomCon = False
                boolOpt = False
                Me.panLabels.Visible = False
                boolMakeVis = False

                'strNC = "Assign Nom. Conc."

                'Me.txtHelper2.Visible = False
                'Me.cmdHelper2.Visible = False
                'str1 = "Term 2" & ChrW(10) & "(Optional)"
                'Me.lbltxtHelper2.Text = str1
                'Me.lbltxtHelper2.Visible = False

                'str1 = "Assign Nom. Conc."
                'Me.lbldgvNomConc.Text = str1
                'Me.dgvNomConc.Visible = True
                'Me.lbldgvNomConc.Visible = True
                'Me.cmdNomConc.Visible = True

            Case 3, 28, 36 'Summary of Back-Calculated Calibration Std Conc
                'boolHelper1 = False
                'boolDHelper2 = False
                'boolTHelper2 = False
                'boolNomCon = False
                'boolOpt = False

                'from dilution qc
                boolHelper1 = False
                boolDHelper2 = False
                boolTHelper2 = False
                boolNomCon = True
                boolOpt = False



            Case 4, 37 'Summary of Interpolated QC Std Conc
                boolHelper1 = False
                boolDHelper2 = False
                boolTHelper2 = False
                boolNomCon = True
                boolOpt = False
                strNC = "Assign Nom. Conc."
                boolChkUseWat = True

            Case 11 'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
                boolHelper1 = False
                boolDHelper2 = False
                boolTHelper2 = False
                boolNomCon = True
                boolOpt = False
                strNC = "Assign Nom. Conc."
                boolChkUseWat = True


            Case 12 'Summary of Interpolated Dilution QC Concentrations
                boolHelper1 = False
                boolDHelper2 = False
                boolTHelper2 = False
                boolNomCon = True
                boolOpt = False

                strNC = "Assign Nom. Conc."


            Case 21, 31, 32, 17 '[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
                boolHelper1 = True
                boolDHelper2 = False
                boolTHelper2 = True
                boolNomCon = True
                boolOpt = False
                If intID = 17 Then
                Else
                    boolChkUseWat = True
                End If

                strH2 = "Run Identifier" & ChrW(10) & "(Optional)"

            Case 22 '[Period Temp] Stock Solution Stability Assessment
                boolHelper1 = True
                boolDHelper2 = False
                boolTHelper2 = True
                boolNomCon = False
                boolOpt = False

                Me.txtHelper2.Visible = True
                Me.cmdHelper2.Visible = True
                'strH2 = "Stock Soln." & ChrW(10) & "Conc."
                strH2 = "Stock Soln." & ChrW(10) & "Conc or Label"


            Case 23 '[Period Temp] Spiking Solution Stability Assessment
                boolHelper1 = False
                boolDHelper2 = False
                boolTHelper2 = False
                boolNomCon = True
                boolOpt = False

                strNC = "Assign Std. Conc."


            Case 29, 34 '[Period Temp] Long-Term QC Std Storage Stability
                boolHelper1 = False
                'boolDHelper2 = True
                'boolTHelper2 = False
                boolDHelper2 = True
                boolTHelper2 = False
                boolNomCon = True
                boolOpt = False
                boolChkUseWat = True

                strH2 = "Assign Run" & ChrW(10) & "Identifier"


                strNC = "Assign Nom. Conc."



            Case Else '6, 7, 8, 9, 10, 13, 14, 15, 16, 17, 18, 19, 20, 24, 25, 26, 27 '
                boolHelper1 = False
                boolDHelper2 = False
                boolTHelper2 = False
                boolNomCon = True
                boolOpt = False
                strNC = "Assign Nom. Conc."
                boolChkUseWat = True

        End Select

        Select Case intID
            Case 13, 14, 15, 37
                boolChkUseWat = False
        End Select

        Me.chkUseWatson.Visible = boolChkUseWat

        Me.panAssignTerm2.Visible = boolTHelper2
        Me.txtHelper2.Visible = boolTHelper2
        Me.lbltxtHelper2.Visible = boolTHelper2
        Me.lbltxtHelper2.Text = strH2

        Me.lbldgvHelper2.Visible = boolDHelper2
        If boolDHelper2 Then
            Me.cmdHelper2.Visible = boolDHelper2
        ElseIf boolTHelper2 Then
            Me.cmdHelper2.Visible = boolTHelper2
        Else
            Me.cmdHelper2.Visible = False
        End If
        Me.dgvHelper2.Visible = boolDHelper2
        Me.lbldgvHelper2.Text = strH2

        Me.panAssignNomConc.Visible = boolNomCon
        Me.lbldgvNomConc.Text = strNC
        'Me.dgvNomConc.Visible = boolNomCon
        'Me.chkShowAllNomConc.Visible = boolNomCon
        'Me.lbldgvNomConc.Visible = boolNomCon
        'Me.cmdNomConc.Visible = boolNomCon

        Call FormatPans()

        If boolMakeVis Then
            Me.panLabels.Visible = True
        End If


    End Sub

    Sub FormatPans()

        Dim a, b, c, d

        If Me.dgvHelper2.Visible Then
            Me.dgvHelper2.Left = Me.dgvHelper1.Left
            Me.cmdHelper2.Visible = True
            Me.cmdHelper2.Top = Me.dgvHelper2.Top + Me.dgvHelper2.Height - Me.cmdHelper2.Height

            Me.panAssignTerm2.Top = Me.dgvHelper2.Top + Me.dgvHelper2.Height
            'Me.cmdHelper2.Top = Me.panAssignTerm2.Top + Me.panAssignTerm2.Height - Me.cmdHelper2.Height
            a = Me.panAssignTerm2.Top + Me.panAssignTerm2.Height

            a = Me.dgvHelper2.Top + Me.dgvHelper2.Height
            Me.panAssignNomConc.Top = a
            b = Me.lbldgvHelper1.Top - a
            Me.panAssignNomConc.Height = b
        Else
            Me.panAssignTerm2.Top = Me.lblLabelAssignment.Top + Me.lblLabelAssignment.Height
            Me.cmdHelper2.Top = Me.panAssignTerm2.Top + Me.panAssignTerm2.Height - Me.cmdHelper2.Height
            If Me.panAssignTerm2.Visible Then
                a = Me.panAssignTerm2.Top + Me.panAssignTerm2.Height
            Else
                a = Me.lblLabelAssignment.Top + Me.lblLabelAssignment.Height
            End If
            Me.panAssignNomConc.Top = a
            If Me.dgvHelper1.Visible Then
                b = Me.lbldgvHelper1.Top - a
            Else
                b = (Me.dgvHelper1.Top + Me.dgvHelper1.Height) - a
            End If

            Me.panAssignNomConc.Height = b
        End If

    End Sub

    Sub ASNum()
        Dim dv As system.data.dataview
        Dim int1 As Short
        Dim var2

        dv = Me.dgvAssignedSamples.DataSource
        int1 = dv.Count
        var2 = Format(int1, "#,##0")
        Me.txtASNum.Text = var2
        Me.txtASNum.Refresh()

    End Sub

    Sub SetAnalysisResultsTable(ByVal id As Int64, cn As ADODB.Connection)

        '20180822 LEE:
        'This routine will load external study

        'id is STUDYID

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim rs As New ADODB.Recordset
        'Dim cn As New ADODB.Connection
        Dim fld As ADODB.Field
        Dim drow As DataRow
        Dim var1, var2, var3, var4, var5, var6
        Dim row1() As DataRow
        Dim strF As String
        Dim int1 As Int64
        Dim rowsCheck() As DataRow
        Dim boolAppend As Boolean = False
        Dim strS As String
        Dim Count1 As Int32
        Dim Count2 As Int32

        '20180821 LEE:
        'ignore if original
        If boolOriginal Then
            GoTo end3
        End If

        'cn.Open(constrCur)

        boolANSI = True

        'remember, this SQL Statement is also in cbxStudy validating event


        If boolAccess Then

            '20160225 LEE: had to re-optimize: previous query would non open in design mode for Ricerca 032852, but can open in Frontage
            'str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
            str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, (ANARUNANALYTERESULTS.CONCENTRATION/ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP "
            str2 = "FROM STUDY INNER JOIN ((ASSAYANALYTES INNER JOIN ((ANALYTICALRUNANALYTES INNER JOIN (" & strAnaRunPeak & " INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) ON (ASSAYANALYTES.ANALYTEINDEX = ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (ASSAYANALYTES.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN CONFIGSAMPLETYPES ON ASSAY.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON STUDY.STUDYID = ASSAY.STUDYID "
            str3 = "WHERE(((ASSAY.STUDYID) = " & id & ")) "
            str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

            '20160816 LEE: Had to completely redo. Ricerca PLASMA study was not returning any Recovery samples because no regression was performed in those samples
            str1 = "SELECT " & strAnaRunPeak & ".ANALYTEINDEX, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strAnaRunPeak & ".RECORDTIMESTAMP, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.ASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ([ANARUNANALYTERESULTS].[CONCENTRATION]/[ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]) AS REPORTEDCONC, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, STUDY.STUDYNAME, ASSAY.MASTERASSAYID, ASSAYANALYTES.ANALYTEID "
            str2 = "FROM (STUDY INNER JOIN (CONFIGSAMPLETYPES INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN (" & strAnaRunPeak & " LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON CONFIGSAMPLETYPES.SAMPLETYPEKEY = ASSAY.SAMPLETYPEKEY) ON STUDY.STUDYID = ASSAY.STUDYID) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUNANALYTES.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) "
            str3 = "WHERE (((" & strAnaRunPeak & ".STUDYID)=" & id & ")) "
            str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

            'ANARUNRAWANALYTEPEAK
            'RUNANALYTEREGRESSIONSTATUS

            'Round([ALIQUOTFACTOR],12) AS ALIQUOTFACTOR
            '20171119 LEE: Must account for goofy DilnF like 1/51 or 1/11 and 
            str1 = "SELECT " & strAnaRunPeak & ".ANALYTEINDEX, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strAnaRunPeak & ".RUNID, " & strAnaRunPeak & ".STUDYID, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strAnaRunPeak & ".RECORDTIMESTAMP, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, Round([ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]," & intDFDec & ") AS ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUN.ASSAYID, ANARUNANALYTERESULTS.CONCENTRATION, ([ANARUNANALYTERESULTS].[CONCENTRATION]/[ANALYTICALRUNSAMPLE].[ALIQUOTFACTOR]) AS REPORTEDCONC, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, CONFIGSAMPLETYPES.SAMPLETYPEID, CONFIGSAMPLETYPES.SAMPLETYPEKEY, STUDY.STUDYNAME, ASSAY.MASTERASSAYID, ASSAYANALYTES.ANALYTEID "
            str2 = "FROM (STUDY INNER JOIN (CONFIGSAMPLETYPES INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUNANALYTES INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN (" & strAnaRunPeak & " LEFT JOIN ANARUNANALYTERESULTS ON (" & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strAnaRunPeak & ".ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strAnaRunPeak & ".RUNID = ANARUNANALYTERESULTS.RUNID) AND (" & strAnaRunPeak & ".STUDYID = ANARUNANALYTERESULTS.STUDYID)) ON (ANALYTICALRUNSAMPLE.RUNID = " & strAnaRunPeak & ".RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = " & strAnaRunPeak & ".STUDYID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.RUNID = ANALYTICALRUNSAMPLE.RUNID) AND (ANALYTICALRUN.STUDYID = ANALYTICALRUNSAMPLE.STUDYID)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strAnaRunPeak & ".ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANALYTICALRUN.STUDYID)) ON (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID) AND (ASSAY.STUDYID = ANALYTICALRUN.STUDYID)) ON CONFIGSAMPLETYPES.SAMPLETYPEKEY = ASSAY.SAMPLETYPEKEY) ON STUDY.STUDYID = ASSAY.STUDYID) INNER JOIN ASSAYANALYTES ON (ANALYTICALRUN.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ANALYTICALRUNANALYTES.STUDYID = ASSAYANALYTES.STUDYID) AND (ANALYTICALRUNANALYTES.ANALYTEINDEX = ASSAYANALYTES.ANALYTEINDEX) "
            str3 = "WHERE (((" & strAnaRunPeak & ".STUDYID)=" & id & ")) "
            str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

            'ALIQUOTFACTOR
        Else


            '20160225 LEE: had to re-optimize: previous query would non open in design mode for Ricerca 032852, but can open in Frontage
            'str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP "
            str1 = "SELECT DISTINCT " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANARUNANALYTERESULTS.RECORDTIMESTAMP "
            str2 = "FROM " & strSchema & ".STUDY INNER JOIN ((" & strSchema & ".ASSAYANALYTES INNER JOIN ((" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & "." & strAnaRunPeak & " INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX =" & strSchema & ". ANARUNANALYTERESULTS.ANALYTEINDEX)) LEFT JOIN " & strSchema & ".DESIGNSAMPLE ON " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = " & strSchema & ".DESIGNSAMPLE.DESIGNSAMPLEID) ON (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".CONFIGSAMPLETYPES ON " & strSchema & ".ASSAY.SAMPLETYPEKEY = " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY) ON " & strSchema & ".STUDY.STUDYID = " & strSchema & ".ASSAY.STUDYID "
            str3 = "WHERE(((" & strSchema & ".ASSAY.STUDYID) = " & id & ")) "
            str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

            '20160816 LEE: Had to completely redo. Ricerca PLASMA study was not returning any Recovery samples because no regression was performed in those samples
            str1 = "SELECT " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX, " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, " & strSchema & "." & strAnaRunPeak & ".RUNID, " & strSchema & "." & strAnaRunPeak & ".STUDYID, " & strSchema & "." & strAnaRunPeak & ".ANALYTEHEIGHT, " & strSchema & "." & strAnaRunPeak & ".ANALYTEAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTDNAME, " & strSchema & "." & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, " & strSchema & "." & strAnaRunPeak & ".RECORDTIMESTAMP, " & strSchema & ".ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strSchema & "." & strAnaRunPeak & ".SAMPLENAME, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYDATETIME, " & strSchema & ".ANALYTICALRUN.RUNTYPEID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION, (" & strSchema & ".ANARUNANALYTERESULTS.CONCENTRATION/" & strSchema & ".ANALYTICALRUNSAMPLE.ALIQUOTFACTOR) AS REPORTEDCONC, " & strSchema & ".ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEID, " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".ASSAY.MASTERASSAYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID "
            str2 = "FROM (" & strSchema & ".STUDY INNER JOIN (" & strSchema & ".CONFIGSAMPLETYPES INNER JOIN (" & strSchema & ".ASSAY INNER JOIN (" & strSchema & ".ANALYTICALRUNANALYTES INNER JOIN (" & strSchema & ".ANALYTICALRUN INNER JOIN (" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN (" & strSchema & "." & strAnaRunPeak & " LEFT JOIN " & strSchema & ".ANARUNANALYTERESULTS ON (" & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER = " & strSchema & ".ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (" & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX = " & strSchema & ".ANARUNANALYTERESULTS.ANALYTEINDEX) AND (" & strSchema & "." & strAnaRunPeak & ".RUNID = " & strSchema & ".ANARUNANALYTERESULTS.RUNID) AND (" & strSchema & "." & strAnaRunPeak & ".STUDYID = " & strSchema & ".ANARUNANALYTERESULTS.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & "." & strAnaRunPeak & ".RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & "." & strAnaRunPeak & ".STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = " & strSchema & "." & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER)) ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID)) ON (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & "." & strAnaRunPeak & ".ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNANALYTES.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ANALYTICALRUN.ASSAYID) AND (" & strSchema & ".ASSAY.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) ON " & strSchema & ".CONFIGSAMPLETYPES.SAMPLETYPEKEY = " & strSchema & ".ASSAY.SAMPLETYPEKEY) ON " & strSchema & ".STUDY.STUDYID = " & strSchema & ".ASSAY.STUDYID) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ANALYTICALRUN.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ANALYTICALRUNANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX) "
            str3 = "WHERE (((" & strSchema & "." & strAnaRunPeak & ".STUDYID)=" & id & ")) "
            str4 = "ORDER BY " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"

            'RUNANALYTEREGRESSIONSTATUS

        End If

        'ANALYTEID

        strSQL = str1 & str2 & str3 & str4
        ''console.writeline("tblAnalysisResultsASamples: " & strSQL)
        ''console.writeline(strSQL)
        'Debug.WriteLine(strSQL)
        rs.CursorLocation = CursorLocationEnum.adUseClient
        Try
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            var1 = var1 'DEbug
        End Try

        '''console.writeline(strSQL)
        'Try
        '    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        'Catch ex As Exception
        '    str1 = "SELECT DISTINCT ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, ANARUNRAWANALYTEPEAK_INJECT.SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDHEIGHT, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDAREA, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, ANARUNRAWANALYTEPEAK_INJECT.ANALYTEPEAKRETENTIONTIME, ANARUNRAWANALYTEPEAK_INJECT.INTERNALSTANDARDRETENTIONTIME "
        '    str2 = "FROM (((ANALYTICALRUNANALYTES INNER JOIN (ANARUNRAWANALYTEPEAK_INJECT INNER JOIN (ASSAY INNER JOIN (ANALYTICALRUN INNER JOIN (ANALYTICALRUNSAMPLE INNER JOIN ANARUNANALYTERESULTS ON (ANALYTICALRUNSAMPLE.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNSAMPLE.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER)) ON (ANALYTICALRUN.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANALYTICALRUN.RUNID = ANARUNANALYTERESULTS.RUNID)) ON (ASSAY.STUDYID = ANALYTICALRUN.STUDYID) AND (ASSAY.ASSAYID = ANALYTICALRUN.ASSAYID) AND (ASSAY.RUNID = ANALYTICALRUN.RUNID)) ON (ANARUNRAWANALYTEPEAK_INJECT.STUDYID = ANARUNANALYTERESULTS.STUDYID) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER = ANARUNANALYTERESULTS.RUNSAMPLESEQUENCENUMBER) AND (ANARUNRAWANALYTEPEAK_INJECT.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX)) ON (ANALYTICALRUNANALYTES.ANALYTEINDEX = ANARUNANALYTERESULTS.ANALYTEINDEX) AND (ANALYTICALRUNANALYTES.RUNID = ANARUNANALYTERESULTS.RUNID) AND (ANALYTICALRUNANALYTES.STUDYID = ANARUNANALYTERESULTS.STUDYID)) INNER JOIN STUDY ON ANALYTICALRUNANALYTES.STUDYID = STUDY.STUDYID) LEFT JOIN DESIGNSAMPLE ON ANALYTICALRUNSAMPLE.DESIGNSAMPLEID = DESIGNSAMPLE.DESIGNSAMPLEID) LEFT JOIN CONFIGSAMPLETYPES ON DESIGNSAMPLE.SAMPLETYPEKEY = CONFIGSAMPLETYPES.SAMPLETYPEKEY "
        '    str3 = "WHERE(((ASSAY.STUDYID) = 424)) "
        '    str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANARUNRAWANALYTEPEAK_INJECT.RUNSAMPLESEQUENCENUMBER;"
        '    strSQL = str1 & str2 & str3 & str4
        '    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        'End Try

        rs.ActiveConnection = Nothing
        int1 = rs.RecordCount 'debugging

        ''the following code doesn't make sense anymore
        ''tblAnalysisResults are updated with cbxStudy change
        ''   - boolAppend will always be false

        '20180821 LEE:
        'Not true. tblAnalysisResults is not updated in cbxStudy change (selected index change)
        ''20180821 LEE:
        ''Hmmm. This isn't a consideration any more
        'Try
        '    If Me.tblAnalysisResults.Rows.Count <> 0 Then
        '        'check to see if data should be appended to table
        '        strF = "STUDYID = " & id
        '        rowsCheck = tblAnalysisResultsHomeOutStudy.Select(strF)
        '        If rowsCheck.Length = 0 Then
        '            boolAppend = True
        '        Else
        '            boolAppend = False
        '        End If
        '    Else
        '        boolAppend = True
        '    End If
        'Catch ex As Exception
        '    var1 = ex.Message
        '    boolAppend = True
        'End Try


        boolAppend = True

        If boolAppend Then 'continue
        Else
            GoTo end1
        End If

        If boolOriginal Or boolCancelButton Then
            tblAnalysisResultsHomeOutStudy.Clear()
        End If

        ''debug
        'Console.WriteLine("Start rs")
        'For Count1 = 0 To rs.Fields.Count - 1
        '    Console.WriteLine(rs.Fields(Count1).Name)
        'Next
        'Console.WriteLine("End rs")

        'debug
        'Console.WriteLine("Start tblAnalysisResults1")
        'For Count1 = 0 To tblAnalysisResults.Columns.Count - 1
        '    Console.WriteLine(tblAnalysisResults.Columns(Count1).ColumnName)
        'Next
        'Console.WriteLine("End tblAnalysisResults1")

        tblAnalysisResultsHomeOutStudy.Clear()
        tblAnalysisResultsHomeOutStudy.AcceptChanges()
        tblAnalysisResultsHomeOutStudy.BeginLoadData()
        daDoPr.Fill(tblAnalysisResultsHomeOutStudy, rs)
        tblAnalysisResultsHomeOutStudy.EndLoadData()

        ''debug
        'Console.WriteLine("Start tblAnalysisResultsHomeOutStudy")
        'For Count1 = 0 To tblAnalysisResultsHomeOutStudy.Columns.Count - 1
        '    Console.WriteLine(tblAnalysisResultsHomeOutStudy.Columns(Count1).ColumnName)
        'Next
        'Console.WriteLine("End tblAnalysisResultsHomeOutStudy")

        var1 = var1


        If tblAnalysisResultsHomeOutStudy.Columns.Contains("CHARANALYTE") Then
        Else
            'add CHARANALYTE
            Dim col1 As New DataColumn
            col1.ColumnName = "CHARANALYTE"
            col1.Caption = "Analyte"
            col1.DataType = System.Type.GetType("System.String")
            tblAnalysisResultsHomeOutStudy.Columns.Add(col1)
        End If

        If tblAnalysisResultsHomeOutStudy.Columns.Contains("INTGROUP") Then
        Else
            '20171124 LEE:
            'need intGroup as well
            Dim col2 As New DataColumn
            col2.ColumnName = "INTGROUP"
            col2.Caption = "Group"
            col2.DataType = System.Type.GetType("System.Int16")
            tblAnalysisResultsHomeOutStudy.Columns.Add(col2)
        End If

        'If Me.tblAnalysisResults.Columns.Count > 0 Then
        'Else
        '    For Each fld In rs.Fields
        '        Dim col As New DataColumn
        '        col.ColumnName = fld.Name
        '        If StrComp(col.ColumnName, "CONCENTRATION", CompareMethod.Text) = 0 Then
        '            col.DataType = System.Type.GetType("System.Double")
        '        End If
        '        Me.tblAnalysisResults.Columns.Add(col)
        '    Next
        '    'add CHARANALYTE
        '    Dim col1 As New DataColumn
        '    col1.ColumnName = "CHARANALYTE"
        '    col1.Caption = "Analyte"
        '    col1.DataType = System.Type.GetType("System.String")
        '    Me.tblAnalysisResults.Columns.Add(col1)

        '    '20171124 LEE:
        '    'need intGroup as well
        '    Dim col2 As New DataColumn
        '    col2.ColumnName = "INTGROUP"
        '    col2.Caption = "Group"
        '    col2.DataType = System.Type.GetType("System.Int16")
        '    Me.tblAnalysisResults.Columns.Add(col2)

        'End If

end1:

        '20171124 LEE
        'I don't think this needs to be done anymore
        'it's done in modHomeCode_4.FillAnalysisResultsTable

        '20180821 LEE:
        'Need this. This only gets called now from cbxStudy.selectedindexchange
        'GoTo end2

        If Me.tblAnalytes.Rows.Count = 0 Then
            GoTo end2
        End If

        'filter for new rows
        strF = "STUDYID = " & id
        'strS = "RUNID, RUNSAMPLESEQUENCENUMBER"
        strS = "RUNID, RUNSAMPLEORDERNUMBER"

        Erase rowsCheck
        rowsCheck = tblAnalysisResultsHomeOutStudy.Select(strF)
        Dim intL As Int16
        intL = rowsCheck.Length 'debugging

        Dim rowsA() As DataRow
        Dim intLA As Short


        strF = "IsIntStd = 'No'"

        rowsA = Me.tblAnalytes.Select(strF, "AnalyteDescription ASC")
        intLA = rowsA.Length

        ' Me.cbxAccStatus.Items.Add("Show All")
        str1 = Me.cbxAccStatus.Text
        Dim boolAll As Boolean = True
        If StrComp(str1, "Show All", CompareMethod.Text) = 0 Then
            boolAll = True
        Else
            boolAll = False
        End If

        '20180821 LEE:
        'Note that this function assumes that study has one matrix and one calibration set
        For Count1 = 0 To intLA - 1

        

            'var1 = rowsA(Count1).Item("ANALYTEINDEX")
            'var2 = rowsA(Count1).Item("MASTERASSAYID")
            var3 = rowsA(Count1).Item("AnalyteDescription")
            var5 = rowsA(Count1).Item("MATRIX")
            'strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 & " AND STUDYID = " & id

            var1 = rowsA(Count1).Item("ANALYTEID")
            strF = "ANALYTEID = " & var1 & " AND STUDYID = " & id & " AND SAMPLETYPEID = '" & var5 & "'"

            '20171124 LEE:
            'The previous overwrites if multiple matrix
            'need group

            var4 = rowsA(Count1).Item("INTGROUP")

            '20180821 LEE:
            'GetASSAYIDFilter doesn't work for new study. There are different assayid's
            str2 = "" ' GetASSAYIDFilter(CInt(var4), boolAll, False)
            If Len(str2) = 0 Then
            Else
                strF = strF & " AND (" & str2 & ")"
            End If

            rowsCheck = tblAnalysisResultsHomeOutStudy.Select(strF)
            intL = rowsCheck.Length 'debugging
            For Count2 = 0 To intL - 1
                'must fill CHARANALYTE because it is blank
                'enter CHARANALYTE
                rowsCheck(Count2).BeginEdit()
                rowsCheck(Count2).Item("CHARANALYTE") = var3
                'enter group
                Try
                    rowsCheck(Count2).Item("INTGROUP") = var4
                Catch ex As Exception
                    var1 = var1 'debug
                End Try

                rowsCheck(Count2).EndEdit()
            Next
        Next

end2:

        rs.Close()
        rs = Nothing

end3:


    End Sub

    Sub InitializeAssignedSamples()

        Dim dgv As DataGridView
        Dim strF As String
        Dim strS As String
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim boolV As Boolean
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim var1, var2, var3
        Dim intRow As Short
        Dim tbl As System.Data.DataTable
        Dim boolFormat As Boolean
        Dim str1a As String

        dgv = Me.dgvAssignedSamples
        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        var1 = id_tblStudies
        If Me.dgvTables.RowCount = 0 Then
            intRow = -1
        ElseIf Me.dgvTables.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvTables.CurrentRow.Index
            var2 = intRow
        End If
        If intRow = -1 Then
            var2 = intRow
        Else
            var2 = Me.dgvTables.Rows.Item(intRow).Cells("id_tblconfigreporttables").Value
        End If

        If Me.dgvAnalytes.RowCount = 0 Then
            intRow = -1
        ElseIf Me.dgvAnalytes.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvAnalytes.CurrentRow.Index
        End If
        If intRow = -1 Then
            var3 = intRow
        Else
            var3 = Me.dgvAnalytes.Rows.Item(intRow).Cells(0).Value
        End If

        tbl = tblAssignedSamples

        'for initialization
        Dim dgv2 As DataGridView
        int1 = dgv.Columns.Count

        ''add unbound columns to tbl'ALREADY DONE IN FRMHOME_01_LOAD EVENT
        dgv2 = Me.dgvAnalyticalRuns
        int2 = dgv2.Columns.Count

        var1 = id_tblStudies
        strF = "id_tblStudies = " & var1 & " AND id_tblConfigReportTables = " & var2 & " AND CHARANALYTE = '" & CleanText(CStr(var3)) & "'"
        'strS = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        strSortNonISR = strS
        strSortISR = "DESIGNSAMPLEID ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
        Dim dv3 As system.data.dataview = New DataView(tblAssignedSamples, strF, strS, DataViewRowState.CurrentRows)
        dv3.AllowNew = False
        dv3.AllowDelete = False

        dgv.DataSource = dv3
        Call OrderColumns(dgv, True)

        int1 = dv3.Count 'debuggin

        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        Dim boolFormatDate As Boolean
        Try
            For Count1 = 0 To int2 - 1
                boolFormat = False
                boolFormatDate = False
                str1 = dgv2.Columns.Item(Count1).Name
                If Count1 = 28 Then
                    var1 = var1 'debug
                End If
                If InStr(1, str1, "ALIQUOT", CompareMethod.Text) > 0 Then
                    var1 = var1 'debug
                End If

                str3 = str1
                Select Case str1
                    Case "CONCENTRATION"
                        '20160509
                        boolFormat = True
                    Case "STUDYNAME"
                        str1 = "charStudyName2"
                    Case "ASSAYDATETIME"
                        boolFormatDate = True
                End Select

                'str2 = dgv2.Columns.Item(Count1).HeaderText
                'boolV = dgv2.Columns.Item(Count1).Visible
                'var1 = dgv2.Columns.Item(Count1).Width

                'legend
                'dgv = Me.dgvAssignedSamples
                'dgv2 = Me.dgvAnalyticalRuns

                If dgv.Columns.Contains(str1) Then
                    str2 = dgv2.Columns.Item(str3).HeaderText
                    boolV = dgv2.Columns.Item(str3).Visible
                    var1 = dgv2.Columns.Item(str3).Width

                    Select Case str1
                        Case "ALIQUOTFACTOR"
                            str2 = "Dil." & ChrW(10) & "FactorOld"
                        Case "ALIQUOTFACTOR"
                            str2 = "Dil." & ChrW(10) & "Factor"
                    End Select

                    dgv.Columns.Item(str1).HeaderText = str2
                    dgv.Columns.Item(str1).Visible = boolV
                    dgv.Columns.Item(str1).Width = var1

                    'dgv.Columns.item(str1).SortMode = DataGridViewColumnSortMode.NotSortable

                    dgv.Columns.Item(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.None

                    dgv.Columns.Item(str1).DefaultCellStyle.Alignment = dgv2.Columns.Item(Count1).DefaultCellStyle.Alignment
                    'dgv.Columns.item(str1).Width = var1
                    If boolFormat Then
                        'dgv.Columns.item(Count1).HeaderText = "Conc."
                        'dgv.Columns.item(Count1).Visible = True
                        dgv.Columns.Item(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        str2 = "##0.000"
                        dgv.Columns.Item(str1).DefaultCellStyle.Format = str2
                    ElseIf boolFormatDate Then
                        str2 = "MMM dd, yyyy hh:mm:ss tt"
                        str2 = "MMM dd, yyyy HH:mm:ss"
                        dgv.Columns.Item(str1).DefaultCellStyle.Format = str2
                    End If
                    dgv.Columns.Item(str1).ReadOnly = True
                Else
                    var1 = var1 'debug
                End If


            Next
        Catch ex As Exception
            var1 = var1 'debug
        End Try

        'fix runid
        dgv.Columns.Item("RUNID").HeaderText = "Run ID"
        dgv.Columns.Item("STUDYNAME").Visible = False
        dgv.Columns.Item("charStudyName2").Visible = True
        dgv.Columns.Item("charStudyName2").HeaderText = "Study" & ChrW(10) & "Name"
        dgv.Columns.Item("CHARHELPER1").HeaderText = "Term 1"
        dgv.Columns.Item("CHARHELPER2").HeaderText = "Term 2"
        dgv.Columns.Item("NOMCONC").Visible = True
        dgv.Columns.Item("NOMCONC").HeaderText = "Nom." & ChrW(10) & "Conc."
        dgv.Columns.Item("NOMCONC").AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
        dgv.Columns.Item("NOMCONC").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgv.Columns.Item("RUNSAMPLEKIND").HeaderText = "Sample" & ChrW(10) & "Type"
        'dgv.Columns.Item("BOOLEXCLSAMPLE").Visible = True
        'dgv.Columns.Item("BOOLEXCLSAMPLE").HeaderText = "Excl." & ChrW(10) & "Sample"
        'dgv.Columns.Item("BOOLEXCLSAMPLE").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        If LAllowExclSamples And gAllowExclSamples Then
            dgv.Columns.Item("BOOLEXCLSAMPLECHK").Visible = True
        Else
            dgv.Columns.Item("BOOLEXCLSAMPLECHK").Visible = False
        End If
        dgv.Columns.Item("BOOLEXCLSAMPLECHK").HeaderText = "Excl." & ChrW(10) & "Sample"
        dgv.Columns.Item("BOOLEXCLSAMPLECHK").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns.Item("BOOLUSEGUWUACCCRIT").Visible = False
        dgv.Columns.Item("BOOLUSEGUWUACCCRIT").HeaderText = "Acc." & ChrW(10) & "Crit."
        dgv.Columns.Item("BOOLUSEGUWUACCCRIT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns.Item("NUMMINACCCRIT").Visible = False
        dgv.Columns.Item("NUMMINACCCRIT").HeaderText = "Neg." & ChrW(10) & "Crit."
        dgv.Columns.Item("NUMMINACCCRIT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns.Item("NUMMAXACCCRIT").Visible = False
        dgv.Columns.Item("NUMMAXACCCRIT").HeaderText = "Pos." & ChrW(10) & "Crit."
        dgv.Columns.Item("NUMMAXACCCRIT").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns.Item("DESIGNSAMPLEID").Visible = False
        dgv.Columns.Item("DESIGNSAMPLEID").HeaderText = "Sample" & ChrW(10) & "ID"
        dgv.Columns.Item("DESIGNSAMPLEID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.Columns.Item("INTGROUP").Visible = False

        'dgv.Columns.Item("REPORTEDCONC").Visible = True 'debug

        'add a checkboxcolumn
        'Dim cbxcolumn As New DataGridViewCheckBoxColumn()
        'With cbxcolumn
        '    .HeaderText = "Excl." & ChrW(10) & "Sample"
        '    .Name = "BOOLEXCLSAMPLECHK"
        '    .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        '    .FlatStyle = FlatStyle.Standard
        '    .CellTemplate = New DataGridViewCheckBoxCell()
        '    .CellTemplate.Style.BackColor = Color.Beige
        '    .Visible = True
        'End With

        'dgv.Columns.Insert(int2 - 4, cbxcolumn)


        'dgv.Columns.Item("MASTERASSAYID").Visible = True
        'dgv.Columns.Item("ANALYTEINDEX").Visible = True

        'dgv.Columns.Item("ID_TBLREPORTTABLE").Visible = True

        int2 = dgv.Columns.Count

        'dgv.Columns.Insert(int2 - 4, cbxcolumn)

        int2 = dgv.Columns.Count

        Call OrderColumns(Me.dgvAssignedSamples, True)

        'dgv.Columns.item("CHARANALYTE").Frozen = True

        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        'dgv.AutoResizeColumns()
        'dgv.Refresh()

        ''debug
        'Try
        '    'console.writeline("Start InitializeAssignedSamples")
        '    var1 = ""
        '    For Count1 = 0 To dgv.Columns.Count - 1
        '        var2 = dgv.Columns(Count1).Name
        '        var1 = var1 & ChrW(9) & var2
        '    Next
        '    'console.writeline(var1)
        '    Dim Count2 As Int32
        '    For Count2 = 0 To dgv.RowCount - 1
        '        var1 = ""
        '        For Count1 = 0 To dgv.Columns.Count - 1
        '            var2 = dgv(Count1, Count2).Value
        '            var1 = var1 & ChrW(9) & var2
        '        Next
        '        'console.writeline(var1)
        '    Next
        '    'console.writeline("End InitializeAssignedSamples")
        'Catch ex As Exception
        '    var2 = var2

        'End Try


    End Sub

    Sub AdjustAssignedSamplesWidth()

        'Why?

        Dim int1 As Short
        Dim int2 As Short
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim Count1 As Short
        Dim var1, var2
        Dim str1 As String
        Dim str2 As String
        Dim wid

        'Exit Sub

        dgv = Me.dgvAssignedSamples
        dgv1 = Me.dgvAnalyticalRuns
        dgv2 = Me.dgvAnalytes
        'wid = dgv2.Columns.Item("AnalyteDescription").Width

        'autosize charAnalyte
        dgv1.Columns("CHARANALYTE").AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

        dgv1.Columns.Item("CHARANALYTE").Width = wid
        'dgv1.Refresh()
        'MsgBox("Done")

        int1 = dgv.Columns.Count
        int2 = dgv1.Columns.Count

        For Count1 = 0 To int2 - 1
            Try
                str1 = dgv1.Columns.Item(Count1).Name
                str2 = str1
                Select Case str1
                    Case "STUDYNAME"
                        str2 = "charStudyName2"
                End Select
                var1 = dgv1.Columns.Item(str1).Width
                If dgv.Columns.Contains(str2) Then
                    var2 = dgv.Columns.Item(str2).Width
                    If var1 = var2 Then
                    Else
                        dgv.Columns.Item(str2).Width = var1
                    End If
                End If

            Catch ex As Exception
                var1 = var1 'debug
            End Try

        Next


    End Sub


    Sub FillAssignedSamples()

        Dim dgv As DataGridView
        'Dim tbl1 as System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim int1 As Short
        Dim var1, var2, var3, var4
        'Dim rs As New ADODB.Recordset
        'Dim cn As New ADODB.Connection
        Dim intRow As Short
        Dim dgv2 As DataGridView
        Dim str1 As String
        Dim strIS As String
        Dim intRowA As Short
        Dim idT As Long

        If boolFormLoad Then
        Else
            'Me.lblWait.Visible = True
            'Me.lblWait.Refresh()
        End If

        dgv = Me.dgvAssignedSamples
        dgv2 = Me.dgvAnalyticalRuns

        Dim boolNoAnal As Boolean = False '20190214 LEE:
        If Me.dgvAnalytes.RowCount = 0 Then
            'must clear dgvassignsamples
            intRowA = -1
            boolNoAnal = True
            'Exit Sub
        ElseIf Me.dgvAnalytes.CurrentRow Is Nothing Then
            intRowA = 0
        Else
            intRowA = Me.dgvAnalytes.CurrentRow.Index
        End If


        If boolNoAnal Then
            strIS = "No"
        Else
            strIS = Me.dgvAnalytes.Rows.Item(intRowA).Cells("IsIntStd").Value
        End If


        var1 = id_tblStudies
        idT = 0
        If StrComp(strIS, "Yes", CompareMethod.Text) = 0 Then
            If Me.dgvTables.RowCount = 0 Then
                var2 = 0
                var3 = "a"
                idT = 0
            ElseIf Me.dgvTables.CurrentRow Is Nothing Then
                var2 = 0
                var3 = "a"
                idT = 0
            Else
                intRow = Me.dgvTables.CurrentRow.Index
                var2 = Me.dgvTables.Rows.Item(intRow).Cells("id_tblConfigReportTables").Value
                var3 = Me.dgvAnalytes.Rows.Item(intRowA).Cells("AnalyteDescription").Value
                idT = Me.dgvTables.Rows.Item(intRow).Cells("ID_TBLREPORTTABLE").Value
            End If
            'strF = "id_tblStudies = " & var1 & " AND id_tblConfigReportTables = " & var2 & " AND CHARANALYTE = '" & Trim(var3) & "' AND BOOLINTSTD = -1 AND ID_TBLREPORTTABLE = " & idT
            strF = "id_tblConfigReportTables = " & var2 & " AND CHARANALYTE = '" & Trim(CleanText(CStr(var3))) & "' AND BOOLINTSTD = -1 AND ID_TBLREPORTTABLE = " & idT

        Else
            If Me.dgvTables.RowCount = 0 Then
                var2 = 0
                var3 = 0
                var4 = 0
                idT = 0
            ElseIf Me.dgvTables.CurrentRow Is Nothing Then
                var2 = 0
                var3 = 0
                var4 = 0
                idT = 0
            Else
                intRow = Me.dgvTables.CurrentRow.Index
                var2 = Me.dgvTables.Rows.Item(intRow).Cells("id_tblConfigReportTables").Value
                If boolNoAnal Then '20190214 LEE:
                    var3 = 0
                    var4 = 0
                Else
                    var3 = Me.dgvAnalytes.Rows.Item(intRowA).Cells("ANALYTEINDEX").Value
                    var4 = Me.dgvAnalytes.Rows.Item(intRowA).Cells("MASTERASSAYID").Value
                End If
              
                If boolNoAnal Then
                    var3 = "N"
                Else
                    var3 = Me.dgvAnalytes.Rows.Item(intRowA).Cells("AnalyteDescription").Value
                End If

                idT = Me.dgvTables.Rows.Item(intRow).Cells("ID_TBLREPORTTABLE").Value
            End If
            'strF = "id_tblStudies = " & var1 & " AND id_tblConfigReportTables = " & var2 & " AND ANALYTEINDEX = " & var3 & " AND MASTERASSAYID = " & var4 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT
            strF = "id_tblConfigReportTables = " & var2 & " AND CHARANALYTE = '" & Trim(CleanText(CStr(var3))) & "' AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT

        End If


        '''''''''''''''''''''console.writeline(strF)
        'strS = "id_tblAssignedSamples ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        str1 = Me.cbxSortAssigneSamples.Text
        'strS = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
        If StrComp(str1, "Original", CompareMethod.Text) = 0 Then
            'strS = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
            strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        ElseIf StrComp(str1, "Original", CompareMethod.Text) = 0 Then
            'strS = "ASSAYLEVEL ASC, RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
            strS = "ASSAYLEVEL ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        End If

        Dim dv3 As system.data.dataview = New DataView(tblAssignedSamples, strF, strS, DataViewRowState.CurrentRows)
        'dv3 = tblAssignedSamples.DefaultView
        int1 = tblAssignedSamples.Rows.Count
        dv3.AllowNew = False
        dv3.AllowDelete = False
        dv3.AllowEdit = True

        int1 = dv3.Count

        Dim boolCX As Boolean = boolCont
        boolCont = False

        dgv.DataSource = dv3

        boolCont = boolCX

        Call OrderColumns(dgv, True)

        'Dim cn As New ADODB.Connection
        'Call FillAssignedSamplesDGV(cn)
        'dgv.AutoResizeColumns()

        Call AdjustAssignedSamplesWidth()
        'dgv.AutoResizeRows()

        'If dgv.Columns.item("NOMCONC").Visible Then
        'dgv.AutoResizeColumns()
        'End If


        If boolFormLoad Then
        Else
            'Me.lblWait.Visible = False
            'Me.lblWait.Refresh()
            'Me.Refresh()
        End If

        Call CheckUseGuWuAccCrit()

    End Sub


    Sub AssessSampleAssignment()

        Dim var1
        var1 = wWStudyName 'NZ(Me.txtStudy.Text, "")
        If Len(var1) = 0 Then
            Exit Sub
        End If

        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim int1 As Short
        Dim int2 As Short
        'Dim dgv1 As DataGridView
        'Dim dv1 as system.data.dataview
        Dim int1A As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim strF As String
        Dim bool As Boolean
        Dim colrW, colrS, colrT, colrA
        Dim boolTimer As Boolean
        Dim boolI As Boolean
        Dim boolIS As Boolean
        Dim int3 As Short
        Dim boolA As Boolean
        Dim boolISA As Boolean
        Dim str1 As String
        Dim str2 As String

        Dim tblAnal As System.Data.DataTable
        Dim rowsAnal() As DataRow
        Dim intRowsAnal As Short
        Dim var2, var3, var4
        Dim idT As Long
        Dim idCT As Long
        Dim strMatrix As String

        'Note: tblAnalytes = tblAnalytesHome

        Dim boolCont As Boolean

        strF = "ID_TBLSTUDIES = " & id_tblStudies
        Dim strS As String
        strS = "ID_TBLSTUDIES ASC"
        Dim dvProps As system.data.dataview = New DataView(tblTableProperties, strF, strS, DataViewRowState.CurrentRows)
        Dim intProps As Short

        tblAnal = Me.tblAnalytes
        int1A = tblAnal.Rows.Count

        boolI = False 'for internal standard
        boolTimer = False
        tblA = tblAssignedSamples
        dgv = Me.dgvTables
        dv = dgv.DataSource
        int1 = dv.Count
        'dgv1 = Me.dgvAnalytes
        'dv1 = dgv1.DataSource
        'int1A = dv1.Count

        Dim intOIS As Short = 0
        Dim boolOIS As Boolean = False

        For Count1 = 0 To int1 - 1

            var1 = dv(Count1).Item("CHARTABLENAME")
            bool = dv(Count1).Item("boolRequiresSampleAssignment")
            boolIS = dv(Count1).Item("boolincludeis")
            idT = dv(Count1).Item("ID_TBLREPORTTABLE")
            idCT = dv(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            'further evaluate boolIS because tblReportProperties can override
            strF = "ID_TBLREPORTTABLE = " & idT
            dvProps.RowFilter = strF
            If dvProps.Count = 0 Then
            Else
                intProps = NZ(dvProps(0).Item("BOOLINCLUDEISTBL"), 0)
                intOIS = NZ(dvProps(0).Item("BOOLCUSTOMLEG"), 0)
                If intProps = -1 Then
                    boolIS = True
                    If intOIS = -1 Then
                        boolOIS = True
                    Else
                        boolOIS = False
                    End If
                Else
                    boolIS = False
                    boolOIS = False
                End If
            End If

            colrT = Color.White
            boolI = -1
            boolCont = True
            If boolI = -1 Then
                If bool = -1 Then
                    'loop through Analytes
                    For Count2 = 0 To int1A - 1
                        colrA = Color.White
                        str2 = tblAnal.Rows.Item(Count2).Item("IsIntStd")
                        boolISA = False
                        If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                            boolISA = True
                        Else
                            boolISA = False
                        End If
                        If boolIS And idCT <> 35 Then 'evaluate analytes for int std, ignore 35: Carryover

                            If boolISA Then 'continue
                                str1 = tblAnal.Rows.Item(Count2).Item("AnalyteDescription")
                                var3 = NZ(tblAnal.Rows.Item(Count2).Item("AnalyteIndex"), 0)
                                var4 = NZ(tblAnal.Rows.Item(Count2).Item("MasterAssayID"), 0)
                                strMatrix = NZ(tblAnal.Rows.Item(Count2).Item("MATRIX"), "AA")
                                'determine if analyte is checked
                                'var2 = dv(Count1).Item(str1)
                                'If var2 = -1 Then 'continue
                                'check for table entry for
                                var1 = dv(Count1).Item("id_tblConfigReportTables")
                                strF = "id_tblConfigReportTables = " & var1 & " AND CHARANALYTE = '" & CleanText(str1) & "' AND BOOLINTSTD = -1 AND ID_TBLREPORTTABLE = " & idT ' & " AND MATRIX = '" & strMatrix & "'"

                                Erase rowsA
                                rowsA = tblA.Select(strF)
                                int2 = rowsA.Length
                                If int2 = 0 Then 'color the row
                                    colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                    colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                    boolTimer = True
                                    boolCont = False
                                    Exit For
                                End If
                                'End If
                            Else '

                                If boolOIS Then
                                Else
                                    str1 = tblAnal.Rows.Item(Count2).Item("AnalyteDescription")
                                    var3 = NZ(tblAnal.Rows.Item(Count2).Item("ANALYTEID"), 0)
                                    var4 = NZ(tblAnal.Rows.Item(Count2).Item("MasterAssayID"), 0)
                                    strMatrix = NZ(tblAnal.Rows.Item(Count2).Item("MATRIX"), "AA")
                                    'determine if analyte is checked
                                    var2 = dv(Count1).Item(str1)
                                    If var2 = -1 Then 'continue
                                        'check for table entry for
                                        var1 = dv(Count1).Item("id_tblConfigReportTables")
                                        'strF = "id_tblConfigReportTables = " & var1 & " AND AnalyteIndex = " & var3 & " AND MasterAssayID = " & var4 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT
                                        strF = "id_tblConfigReportTables = " & var1 & " AND ANALYTEID = " & var3 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                        Erase rowsA
                                        rowsA = tblA.Select(strF)
                                        int2 = rowsA.Length
                                        If int2 = 0 Then 'color the row
                                            colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                            colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                            boolTimer = True
                                            boolCont = False
                                            Exit For
                                        End If
                                    Else
                                    End If
                                End If

                            End If
                        Else 'ignore analyte int std
                            If boolISA Then
                            Else
                                str1 = tblAnal.Rows.Item(Count2).Item("AnalyteDescription")
                                var3 = NZ(tblAnal.Rows.Item(Count2).Item("ANALYTEID"), 0)
                                var4 = NZ(tblAnal.Rows.Item(Count2).Item("MasterAssayID"), 0)
                                strMatrix = NZ(tblAnal.Rows.Item(Count2).Item("MATRIX"), "AA")
                                'determine if analyte is checked
                                var2 = dv(Count1).Item(str1)
                                If var2 = -1 Then 'continue
                                    'check for table entry for
                                    var1 = dv(Count1).Item("id_tblConfigReportTables")
                                    strF = "id_tblConfigReportTables = " & var1 & " AND AnalyteIndex = " & var3 & " AND MasterAssayID = " & var4 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT
                                    strF = "id_tblConfigReportTables = " & var1 & " AND ANALYTEID = " & var3 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                    Erase rowsA
                                    rowsA = tblA.Select(strF)
                                    int2 = rowsA.Length
                                    If int2 = 0 Then 'color the row
                                        colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                        colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                        boolTimer = True
                                        boolCont = False
                                        Exit For
                                    End If
                                Else
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            If colrA = dgv.Rows.Item(Count1).DefaultCellStyle.BackColor Then
            Else
                dgv.Rows.Item(Count1).DefaultCellStyle.BackColor = colrA
            End If
        Next

        boolCont = True

        Call AssessSampleAssignmentAnalyte()

    End Sub

    Sub AssessSampleAssignmentAnalyte()

        Dim var1
        var1 = wWStudyName 'NZ(Me.txtStudy.Text, "")
        If Len(var1) = 0 Then
            Exit Sub
        End If

        Dim tblA As System.Data.DataTable
        Dim rowsA() As DataRow
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim dgv1 As DataGridView
        Dim dv1 As System.Data.DataView
        Dim int1A As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim strF As String
        Dim bool As Boolean
        Dim colrW, colrS, colrT, colrA
        Dim boolTimer As Boolean
        Dim boolI As Boolean
        Dim boolIS As Boolean
        Dim int3 As Short
        Dim boolA As Boolean
        Dim boolISA As Boolean
        Dim str1 As String
        Dim str2 As String

        'Dim tblAnal as System.Data.DataTable
        'Dim rowsAnal() As DataRow
        'Dim intRowsAnal As Short
        Dim var2, var3, var4
        Dim intRow As Short

        Dim idT As Long

        strF = "ID_TBLSTUDIES = " & id_tblStudies
        Dim strS As String
        strS = "ID_TBLSTUDIES ASC"
        Dim dvProps As System.Data.DataView = New DataView(tblTableProperties, strF, strS, DataViewRowState.CurrentRows)
        Dim intProps As Short

        Dim strMatrix As String

        'intRow = Me.dgvTables.CurrentRow.Index
        If Me.dgvTables.Rows.Count = 0 Then
            Exit Sub
        End If
        If Me.dgvTables.CurrentRow Is Nothing Then
            Exit Sub
        End If
        intRow = Me.dgvTables.CurrentRow.Index
        Dim boolCont As Boolean

        boolI = False 'for internal standard
        boolTimer = False
        tblA = tblAssignedSamples
        dgv = Me.dgvTables
        dv = dgv.DataSource
        int1 = dv.Count

        dgv1 = Me.dgvAnalytes
        dv1 = dgv1.DataSource
        int1A = dv1.Count


        Dim idCT As Long

        For Count1 = intRow To intRow

            bool = dv(Count1).Item("boolRequiresSampleAssignment")
            boolIS = dv(Count1).Item("boolincludeis")
            idT = dv(Count1).Item("ID_TBLREPORTTABLE")
            idCT = dv(Count1).Item("ID_TBLCONFIGREPORTTABLES")
            'further evaluate boolIS because tblReportProperties can override
            strF = "ID_TBLREPORTTABLE = " & idT
            dvProps.RowFilter = strF
            If dvProps.Count = 0 Then
            Else
                intProps = NZ(dvProps(0).Item("BOOLINCLUDEISTBL"), 0)
                If intProps = -1 Then
                    boolIS = True
                Else
                    boolIS = False
                End If
            End If

            colrT = Color.White
            boolI = -1
            boolCont = True
            If boolI = -1 Then
                If bool = -1 Then
                    'loop through Analytes
                    For Count2 = 0 To int1A - 1
                        colrA = Color.White
                        colrT = Color.White
                        str2 = dv1(Count2).Item("IsIntStd")
                        boolISA = False
                        If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
                            boolISA = True
                        Else
                            boolISA = False
                        End If
                        If boolIS And idCT <> 35 Then 'evaluate analytes for int std, ignore 35: Carryover
                            If boolISA Then 'continue
                                str1 = dv1(Count2).Item("AnalyteDescription")
                                var3 = NZ(dv1(Count2).Item("AnalyteIndex"), 0)
                                var4 = NZ(dv1(Count2).Item("MasterAssayID"), 0)
                                strMatrix = NZ(dv1(Count2).Item("MATRIX"), "AA")
                                'determine if analyte is checked
                                'var2 = dv(Count1).Item(str1)
                                'If var2 = -1 Then 'continue
                                'check for table entry for
                                var1 = dv(Count1).Item("id_tblConfigReportTables")
                                strF = "id_tblConfigReportTables = " & var1 & " AND CHARANALYTE = '" & CleanText(str1) & "' AND BOOLINTSTD = -1 AND ID_TBLREPORTTABLE = " & idT ' & " AND MATRIX = '" & strMatrix & "'"
                                Erase rowsA
                                rowsA = tblA.Select(strF)
                                int2 = rowsA.Length
                                If int2 = 0 Then 'color the row
                                    colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                    colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                    boolTimer = True
                                    boolCont = False
                                End If
                                'End If
                            Else
                                str1 = dv1(Count2).Item("AnalyteDescription")
                                var3 = NZ(dv1(Count2).Item("ANALYTEID"), 0)
                                var4 = NZ(dv1(Count2).Item("MasterAssayID"), 0)
                                strMatrix = NZ(dv1(Count2).Item("MATRIX"), "AA")
                                'determine if analyte is checked
                                var2 = dv(Count1).Item(str1)
                                If var2 = -1 Then 'continue
                                    'check for table entry for
                                    var1 = dv(Count1).Item("id_tblConfigReportTables")
                                    'strF = "id_tblConfigReportTables = " & var1 & " AND AnalyteIndex = " & var3 & " AND MasterAssayID = " & var4 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT
                                    strF = "id_tblConfigReportTables = " & var1 & " AND ANALYTEID = " & var3 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                    Erase rowsA
                                    rowsA = tblA.Select(strF)
                                    int2 = rowsA.Length
                                    If int2 = 0 Then 'color the row
                                        colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                        colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                        boolTimer = True
                                        boolCont = False
                                    End If
                                Else
                                End If
                            End If
                        Else 'ignore analyte int std
                            If boolISA Then
                            Else
                                str1 = dv1(Count2).Item("AnalyteDescription")
                                var3 = NZ(dv1(Count2).Item("ANALYTEID"), 0)
                                var4 = NZ(dv1(Count2).Item("MasterAssayID"), 0)
                                strMatrix = NZ(dv1(Count2).Item("MATRIX"), "AA")
                                'determine if analyte is checked
                                var2 = dv(Count1).Item(str1)
                                If var2 = -1 Then 'continue
                                    'check for table entry for
                                    var1 = dv(Count1).Item("id_tblConfigReportTables")
                                    'strF = "id_tblConfigReportTables = " & var1 & " AND AnalyteIndex = " & var3 & " AND MasterAssayID = " & var4 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT
                                    strF = "id_tblConfigReportTables = " & var1 & " AND ANALYTEID = " & var3 & " AND BOOLINTSTD = 0 AND ID_TBLREPORTTABLE = " & idT & " AND SAMPLETYPEID = '" & strMatrix & "'"
                                    Erase rowsA
                                    rowsA = tblA.Select(strF)
                                    int2 = rowsA.Length
                                    If int2 = 0 Then 'color the row
                                        colrT = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                        colrA = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
                                        boolTimer = True
                                        boolCont = False
                                    End If
                                Else
                                End If
                            End If
                        End If
                        dgv1.Rows.Item(Count2).DefaultCellStyle.BackColor = colrA
                        var1 = dgv1.Rows.Item(Count2).DefaultCellStyle.BackColor
                        'If colrA = dgv1.Rows.Item(Count2).DefaultCellStyle.BackColor Then
                        'Else
                        '    dgv1.Rows.Item(Count2).DefaultCellStyle.BackColor = colrT
                        'End If
                    Next
                End If
            End If
            'If colrT = dgv.Rows.item(Count1).DefaultCellStyle.BackColor Then
            'Else
            '    dgv.Rows.item(Count1).DefaultCellStyle.BackColor = colrT
            'End If
        Next

        boolCont = True

    End Sub

    Sub FillcbxStudy()

        Dim tbl As System.Data.DataTable
        Dim Count1 As Short
        Dim str1 As String
        Dim int1 As Short
        Dim str2 As String
        Dim strF As String
        Dim strS As String
        Dim dv As System.Data.DataView
        Dim var1

        '20190130 LEE:
        'Aack! tblStudiesA is assigned as cbxStudy datasource
        'And cbxStudy datasource seems to persist even when frmAssignSamples is disposed
        'instead, make copy, as was original

        int1 = tblStudies.Rows.Count 'debug
        tbl = tblStudies.Copy

        strS = "CHARWATSONSTUDYNAME ASC"
        strF = "ID_TBLSTUDIES > -1"

        dv = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)

        Me.tblStudiesA = dv.ToTable

        '20190130 LEE:
        'Use new function to load cbxStudies
        Dim tbl2 As DataTable
        tbl2 = CreatedgvwStudiesDatasource(False)
        Me.tblStudiesA = tbl2

        'This should not be tblWStudy anymore
        'It should be StudyDoc Studies
        '20190130 LEE:
        'No! see above
        'tbl = frmH.dgvwStudy.DataSource
        'Me.tblStudiesA = frmH.dgvwStudy.DataSource


        'dv = frmH.dgvwStudy.DataSource

        'Try
        '    Me.tblStudiesA = tblASTUDY.Copy
        'Catch ex As Exception
        '    str1 = ex.Message
        '    str1 = str1
        'End Try

        'tbl = tblwSTUDY

        int1 = Me.tblStudiesA.Rows.Count
        'int1 = dv.Count
        Dim boolH As Boolean = boolHold '20190130 LEE:
        boolHold = True
        Me.cbxStudy.DataSource = Me.tblStudiesA
        Me.cbxStudy.DisplayMember = "STUDYNAME" '"CHARWATSONSTUDYNAME"
        boolHold = boolH

        'Me.cbxStudy.Sorted = True

        'For Count1 = 0 To int1 - 1
        '    str1 = NZ(tbl.Rows.item(Count1).Item("StudyName"), "")
        '    str1 = NZ(tbl.Rows.Item(Count1).Item("CHARWATSONSTUDYNAME"), "")
        '    If Len(str1) = 0 Then
        '    Else
        '        Me.cbxStudy.Items.Add(str1)
        '    End If
        'Next

        'choose txtStudy

        Dim int2 As Int64

        For Count1 = 0 To int1 - 1
            int2 = Me.tblStudiesA.Rows(Count1).Item("STUDYID")
            If int2 = wStudyID Then
                Me.cbxStudy.SelectedIndex = Count1
                Exit For
            End If

        Next

        var1 = var1

        'For Count1 = 0 To int1 - 1
        '    int2 = dv.Item(Count1).Item("STUDYID")
        '    If int2 = wStudyID Then
        '        Me.cbxStudy.SelectedIndex = Count1
        '        Exit For
        '    End If
        'Next

        'str1 = Me.txtStudy.Text
        'For Count1 = 0 To int1 - 1
        '    str2 = tbl.Rows(Count1).Item("CHARWATSONSTUDYNAME")
        '    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
        '        Me.cbxStudy.SelectedIndex = Count1
        '        Exit For
        '    End If

        'Next




    End Sub

    Sub FilldgvTables()

        Dim dv As system.data.dataview
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
        Dim str1 As Short
        Dim rows() As DataRow
        Dim var1, var2, var3

        dv = frmH.dgvReportTableConfiguration.DataSource
        Dim tbl As System.Data.DataTable = dv.ToTable("a")

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
        'col4.ColumnName = "BOOLPLACEHOLDER"
        'col4.DataType = System.Type.GetType("System.Int16")
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
                ' tbl.Rows.Item(Count1).Item("BOOLPLACEHOLDER") = rows(0).Item("BOOLPLACEHOLDER")
                tbl.Rows.Item(Count1).EndEdit()
            Next
        Next

        '20190104 LEE: Git change test

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
        strF = "boolRequiresSampleAssignment = " & True & " AND BOOLINCLUDE = " & True 'Leave as true. Underlying table has boolean
        strF = "boolRequiresSampleAssignment = " & True & " AND BOOLINCLUDE = " & True & " AND BOOLPLACEHOLDER = " & False
        '20181220 LEE:
        'Some custom tables will need sample assignment
        strF = "boolRequiresSampleAssignment = " & True & " AND BOOLINCLUDE = " & True

        'strS = "ID_TBLCONFIGREPORTTABLES ASC"
        'strS = "ORDER ASC"
        'Dim dv1 as system.data.dataview = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)

        strS = "INTORDER ASC"
        Dim dv1 As system.data.dataview = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)

        dv1.AllowDelete = False
        dv1.AllowNew = False
        dgv = Me.dgvTables

        dgv.AllowUserToResizeColumns = True
        dgv.AllowUserToResizeRows = True
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        Dim boolAA As Boolean
        boolAA = boolCont
        boolCont = False
        dgv.DataSource = dv1
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

        '20181206 LEE:
        dgv.Columns("CHARFCID").Visible = True
        dgv.Columns("CHARFCID").HeaderText = "FC ID"

        Dim wd1, wd2
        wd1 = dgv.RowHeadersWidth
        wd2 = dgv.Width - wd1
        'dgv.Columns.Item("CHARHEADINGTEXT").MinimumWidth = wd2 * 0.9 '0.95
        dgv.Columns.Item("CHARHEADINGTEXT").Width = wd2 * 0.85
        dgv.AutoResizeRows()

        Dim idSel As Int64
        Dim idSel1 As Int64
        If boolFormLoad Then
            'find selected row
            intRows_frmh = frmH.dgvReportTableConfiguration.Rows.Count
            If intRows_frmh = 0 Then
                idSel = 0
            ElseIf frmH.dgvReportTableConfiguration.CurrentRow Is Nothing Then
                idSel = 0
            Else
                idSel = frmH.dgvReportTableConfiguration("ID_TBLREPORTTABLE", frmH.dgvReportTableConfiguration.CurrentRow.Index).Value
            End If

            'find introw
            intRow = 0
            For Count1 = 0 To dgv.Rows.Count - 1
                idSel1 = dgv("ID_TBLREPORTTABLE", Count1).Value
                If idSel = idSel1 Then
                    intRow = Count1
                    Exit For
                End If
            Next

            ''now record table id
            'var1 = frmH.dgvReportTableConfiguration.Rows.Item(intRow).Cells("BOOLREQUIRESSAMPLEASSIGNMENT").Value
            ''var2 = frmH.dgvReportTableConfiguration.Rows.Item(intRow).Cells("ID_TBLCONFIGREPORTTABLES").Value
            'var2 = frmH.dgvReportTableConfiguration.Rows.Item(intRow).Cells("ID_TBLREPORTTABLE").Value
            'If var1 = -1 Then
            '    'find var2 in dgv
            '    For Count1 = 0 To dgv.Rows.Count - 1
            '        'var3 = dgv.Item("ID_TBLCONFIGREPORTTABLES", Count1).Value
            '        var3 = dgv.Item("ID_TBLREPORTTABLE", Count1).Value
            '        If var3 = var2 Then
            '            intRow = Count1
            '            Exit For
            '        End If
            '    Next
            'Else
            '    intRow = 0
            'End If
        Else
            'select first row
            intRow = 0
        End If
        dgv.Select()
        If intRows = 0 Or intRow = -1 Then
        Else
            dgv.Rows.Item(intRow).Cells("CHARHEADINGTEXT").Selected = True
        End If

        Call FilldgvAnalytes()

        Call FillAnalyticalRuns("")

        If boolFormLoad Then
        Else
            'Call FillAnalyticalRuns()
            Call FillAssignedSamples()
        End If

    End Sub

    Sub tblAnalytesConfigure(boolFromChange As Boolean)

        Dim lng1 As Int64

        lng1 = Me.txtStudyID.Text

        'NOTE: Cannot use tblAnalyteGroups - this table doesn't have IntStd
        'must use tblAnalytesHome
        If boolFromChange Then
            Me.tblAnalytes = tblAnalytesHome.Copy
            '20161227 LEE: This may need to be addressed when changing study
            intNumMatrices = 1
        Else
            If lng1 = id_tblStudies Then
                Me.tblAnalytes = tblAnalytesHome.Copy
                Dim var1
                var1 = tblAnalytes.Rows.Count
                var1 = var1
            Else 'ignore because tblanalytes was modified in cbxStudyValidating

            End If

            'determine number of matrices in study
            '20161227 LEE: This may need to be addressed when changing study
            Dim tbl1 As DataTable = tblAnalyteGroups
            Dim dv As DataView = New DataView(tbl1)
            Dim tbl2 As DataTable = dv.ToTable("a", True, "MATRIX")
            intNumMatrices = tbl2.Rows.Count

        End If



    End Sub

    Sub InitializedgvAnalytes()

        ''LEGENDS:
        'If StrComp(gSortAnalytes, "Matrix", CompareMethod.Text) = 0 Then
        '    strS = "MATRIX ASC, ANALYTEDESCRIPTION ASC, ANALYTEDESCRIPTION_C ASC, INTGROUP ASC"
        'Else
        '    strS = "ANALYTEDESCRIPTION ASC, ANALYTEDESCRIPTION_C ASC, MATRIX ASC, INTGROUP ASC"
        'End If
        'gSortAnalyteString = strS 'use this in Reassay, Repeat, and Sample Conc tables

        'Dim arrAnalytes(16, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
        '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadminCmpd,14=OriginalAnalyteDescription,15=intGroup,16=MATRIX
        'END LEGENDS

        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim strS As String
        Dim strF As String
        Dim boolAA As Boolean
        Dim var1

        'dgv.AllowUserToResizeRows = True
        'dgv.AllowUserToResizeColumns = True

        dgv = Me.dgvAnalytes
        dgv.AllowUserToResizeRows = True
        dgv.AllowUserToResizeColumns = True
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        strS = ReturnSort(False)

        If boolViewOnly Then
            strF = "IsIntStd = 'No'"
        Else
            strF = "IsIntStd = 'No' or IsIntStd = 'Yes'"
        End If

        Dim dv1 As System.Data.DataView = New DataView(Me.tblAnalytes)
        dv1.AllowDelete = False
        dv1.AllowEdit = False
        dv1.AllowNew = False
        dv1.RowFilter = strF
        dv1.Sort = strS

        int1 = dv1.Count 'debug

        'Try
        '    dgv.Columns("AnalyteDescription").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        '    dgv.Columns("IntStd").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        'Catch ex As Exception
        '    var1 = ex.Message
        '    var1 = var1
        'End Try

        boolAA = boolCont
        boolCont = False
        dgv.DataSource = dv1
        boolCont = boolAA

        int1 = dgv.RowCount

        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns.Item(Count1).Visible = False
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'debug
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        dgv.Columns.Item("AnalyteDescription").Visible = True
        dgv.Columns.Item("AnalyteDescription").HeaderText = "Analyte"
        dgv.Columns.Item("IntStd").HeaderText = "IntStd"

        Dim wd1, wd2
        wd1 = dgv.RowHeadersWidth
        wd2 = dgv.Width - wd1
        dgv.Columns.Item("AnalyteDescription").MinimumWidth = wd2 / 2
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgv.AutoResizeRows()
        'dgv.AutoResizeColumns()

    End Sub

    Sub FilldgvAnalytes()

        Dim tbl As System.Data.DataTable
        Dim Count1 As Short
        Dim bool As Boolean
        Dim dv As system.data.dataview
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
        Dim idRT As Int64
        Dim idT As Int64
        Dim dgvT As DataGridView
        Dim intRowT As Short

        dgv = Me.dgvAnalytes
        tbl = Me.tblAnalytes
        dgvT = Me.dgvTables
        dv = dgvT.DataSource


        'record initial row of dgvAnalytes
        If dgv.RowCount = 0 Then
            intSetRow = 0
        ElseIf dgv.CurrentRow Is Nothing Then
            intSetRow = 0
        Else
            intSetRow = dgv.CurrentRow.Index
        End If

        intRow = 0
        If Me.dgvTables.Rows.Count = 0 Then 'means called from View Analytical Runs
            intRow = -1
            boolIS = False
        Else

            If Me.dgvTables.RowCount = 0 Then
                intRow = 0
            ElseIf Me.dgvTables.CurrentRow Is Nothing Then
                intRow = 0
            Else
                intRow = Me.dgvTables.CurrentRow.Index
            End If
            int1 = NZ(Me.dgvTables("BOOLINCLUDEIS", intRow).Value, 0)

            If int1 = -1 Then
                boolIS = True
            Else
                boolIS = False
            End If
        End If

        'override BOOLINCLUDEIS if desired
        Dim boolOIS As Boolean = False 'only int std
        Dim intOIS As Short

        If Me.dgvTables.Rows.Count = 0 Then
        Else
            If dgvT.CurrentRow Is Nothing Then
            Else
                intRowT = dgvT.CurrentRow.Index

                idT = dgvT("ID_TBLCONFIGREPORTTABLES", intRowT).Value
                Select Case idT
                    Case 13, 14, 15, 22, 23, 31, 32
                        idRT = NZ(dgvT("ID_TBLREPORTTABLE", intRowT).Value, -1)

                        Dim rowsTP() As DataRow
                        Dim intTP As Short

                        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idRT
                        rowsTP = tblTableProperties.Select(strF)
                        'intTP = NZ(rowsTP(0).Item("BOOLINCLUDEISTBL"), 0)
                        If rowsTP.Length = 0 Then
                            intTP = 0
                            intOIS = 0
                        Else
                            intTP = NZ(rowsTP(0).Item("BOOLINCLUDEISTBL"), 0)
                            intOIS = NZ(rowsTP(0).Item("BOOLCUSTOMLEG"), 0)
                        End If
                        If intTP = -1 Then
                            boolIS = True
                        Else
                            boolIS = False
                        End If
                        If intOIS = -1 Then
                            boolOIS = True
                        Else
                            boolOIS = False
                        End If

                    Case 17 'may have Int Std selected

                        idRT = NZ(dgvT("ID_TBLREPORTTABLE", intRowT).Value, -1)

                        Dim rowsTP() As DataRow
                        Dim intTP As Short

                        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idRT
                        rowsTP = tblTableProperties.Select(strF)
                        'intTP = NZ(rowsTP(0).Item("BOOLINCLUDEISTBL"), 0)
                        If rowsTP.Length = 0 Then
                            intTP = 0
                        Else
                            intTP = NZ(rowsTP(0).Item("BOOLINCLUDEISTBL"), 0)
                        End If
                        If intTP = -1 Then
                            boolIS = True
                        Else
                            boolIS = False
                        End If

                    Case Else
                End Select

            End If
        End If

        If intRow = -1 Then 'there are no tables selected to assign samples
        Else
            strF = ""
            Count3 = 0
            For Count1 = 0 To tbl.Rows.Count - 1
                var1 = NZ(tbl.Rows.Item(Count1).Item("AnalyteDescription"), "NA")
                str1 = NZ(tbl.Rows.Item(Count1).Item("IsIntStd"), "")

                If boolIS Then

                    If StrComp(str1, "No", CompareMethod.Text) = 0 Then
                        ' ''20170726 LEE: Why looking boolOIS? This is for custom legend
                        'If boolOIS Then
                        'Else
                        '    bool = dv(intRow).Item(var1)
                        '    If bool Then
                        '        Count3 = Count3 + 1
                        '        If Count3 = 1 Then
                        '            strF = strF & "ANALYTEDESCRIPTION = '" & var1 & "'"
                        '        Else
                        '            strF = strF & " OR ANALYTEDESCRIPTION = '" & var1 & "'"
                        '        End If
                        '    End If
                        'End If

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
                        Count3 = Count3 + 1
                        If Count3 = 1 Then
                            strF = strF & "ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                        Else
                            strF = strF & " OR ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
                        End If

                    End If
                Else
                    'bool = dv(intRow).Item(arrAnalytes(1, Count1))

                    If StrComp(str1, "No", CompareMethod.Text) = 0 Then
                        '20170920 LEE: ?? boolOIS has to do with custom legend
                        'what does custom legend have to do with showing analytes?

                        'If boolOIS Then
                        'Else

                        'End If

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

            If Len(strF) = 0 Then
                strF = "ANALYTEDESCRIPTION = 'AAAABBBBCCCC'"
            End If

            'inspect current dv
            Dim dv2 As System.Data.DataView
            dv2 = Me.dgvAnalytes.DataSource
            str1 = dv2.RowFilter
            If StrComp(str1, strF, CompareMethod.Text) = 0 Then 'ignore
            Else

                If boolIS Then
                    'need to show IntStd column
                    dgv.Columns("IntStd").Visible = True
                Else
                    dgv.Columns("IntStd").Visible = False
                End If

                strS = "IsIntStd ASC, AnalyteDescription ASC"
                strS = "INTORDER ASC"
                Dim dv1 As System.Data.DataView = New DataView(Me.tblAnalytes)
                dv1.RowFilter = strF
                'do not sort!
                dv1.Sort = strS
                dv1.AllowDelete = False
                dv1.AllowNew = False
                boolAA = boolCont
                boolCont = False
                var1 = dv1.Count 'debug
                Dim boolA As Boolean = boolDontChange
                boolDontChange = True
                dgv.DataSource = dv1
                boolDontChange = boolA
                'dgv.AutoResizeColumns()
                boolCont = boolAA
            End If
        End If

        'select first row
        If dgv.RowCount = 0 Then
        Else
            'set initial row
            boolAA = boolCont
            boolCont = False
            If boolFromdgvTable Then
                'str1 = dgv.Rows.item(intSetRow).Cells("AnalyteDescription").Value
                If dgv.RowCount - 1 < intSetRow Then
                    int1 = dgv.RowCount
                    Dim str2 As String
                    Dim boolGo As Boolean
                    boolGo = False
                    str1 = strAnalFromTable
                    For Count1 = 0 To int1 - 1
                        str2 = NZ(dgv.Rows.Item(Count1).Cells("AnalyteDescription").Value, "NA")
                        If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                            boolGo = True
                            Exit For
                        End If
                        If boolGo Then
                            dgv.CurrentCell = dgv.Rows.Item(Count1).Cells(0)
                        Else
                            dgv.CurrentCell = dgv.Rows.Item(0).Cells(0)
                        End If
                    Next
                Else
                    str1 = NZ(dgv.Rows.Item(intSetRow).Cells("AnalyteDescription").Value, "NA")
                    If StrComp(str1, strAnalFromTable, CompareMethod.Text) = 0 Then
                        dgv.CurrentCell = dgv.Rows.Item(intSetRow).Cells("AnalyteDescription")
                    Else 'look for stranalfromtable
                        int1 = dgv.RowCount
                        Dim str2 As String
                        Dim boolGo As Boolean
                        boolGo = False
                        str1 = strAnalFromTable
                        For Count1 = 0 To int1 - 1
                            str2 = NZ(dgv.Rows.Item(Count1).Cells("AnalyteDescription").Value, "NA")
                            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                                boolGo = True
                                Exit For
                            End If
                            If boolGo Then
                                dgv.CurrentCell = dgv.Rows.Item(Count1).Cells(0)
                            Else
                                dgv.CurrentCell = dgv.Rows.Item(0).Cells(0)
                            End If
                        Next
                    End If
                End If
            Else
                dgv.CurrentCell = dgv.Rows.Item(0).Cells(0)
            End If

            boolCont = boolAA
        End If

        Try
            gintGroup = dgv("INTGROUP", 0).Value
        Catch ex As Exception
            gintGroup = -1
        End Try

        gintGroup = gintGroup 'debug

        Call ChooseAnalyte()

    End Sub

    Sub FillAnalyticalRuns(strAccStatus As String)

        Dim tbl As System.Data.DataTable
        Dim strF As String
        Dim strS As String
        Dim dgv As DataGridView
        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9
        Dim intRow As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim boolClear As Boolean
        Dim num As Short
        Dim dt1 As Date
        Dim boolIS As Boolean
        Dim str1 As String
        Dim strName As String
        Dim strRIFilter As String
        Dim strSTFilter As String
        Dim strDFFilter As String
        Dim txtFilter As String
        Dim intSI As Int64
        'Dim strAccStatus As String = "Show All"
        Dim intAnalyteID As Int64
        Dim strMatrix As String
        Dim strFAssayID As String
        Dim strFFF As String

        Dim strChoose As String

        If Me.cbxChooseAnalyte.Items.Count = 0 Then
            strChoose = ""
        Else
            strChoose = Me.cbxChooseAnalyte.Text
        End If

        If boolCont Then
        Else
            Exit Sub
        End If

        If Len(strAccStatus) = 0 Then
            strAccStatus = Me.cbxAccStatus.Text
        End If

        Dim boolSC As Boolean = False

        If boolOriginal Then

        Else
            boolSC = IsStudyChanged() 'is study changed?
        End If



        If boolFormLoad Then

            str1 = ""

            'set dgvAnalyticalRuns datasource as entire tbl, then filter later
            tblAnalysisResults.CaseSensitive = True
            'Note: Sort MUST be applied. FINDSAMPLES uses dv.Find, which requires a sort
            strS = "CHARANALYTE ASC, STUDYNAME ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            Dim dv As DataView = New DataView(tblAnalysisResults, str1, strS, DataViewRowState.CurrentRows)

            dv.AllowDelete = False
            dv.AllowNew = False
            dv.AllowEdit = False

            int1 = dv.Count

            Me.dgvAnalyticalRuns.DataSource = dv

            Call OrderColumns(Me.dgvAnalyticalRuns, False)

        End If


        Dim intGroup As Short
        Dim dgvA As DataGridView = Me.dgvAnalytes
        Dim intARow As Short

        Try

            If dgvA.CurrentRow Is Nothing Or dgvA.RowCount = 0 Then
                intARow = -1
            Else
                intARow = dgvA.CurrentRow.Index
            End If

            If intARow = -1 Then
                intGroup = -1
            Else
                intGroup = dgvA("INTGROUP", intARow).Value
            End If

            intSI = GetWStudyID(NZ(CLng(Me.txtStudyID.Text), 0))

            boolIS = False
            strName = ""
            If boolFormLoad Then  'Loading the form for the first time
                strRIFilter = "[None]"
                strSTFilter = "[None]"
                strDFFilter = "[None]"
                txtFilter = ""
            Else                  'set the Filter variables (for convenience)
                strRIFilter = Me.cbxFilterRunID.Text
                strSTFilter = Me.cbxFilterSampleType.Text
                strDFFilter = Me.cbxFilterDilFactor.Text
                txtFilter = Me.txtFilterSamples.Text
            End If

            'Set state of "Clear Filters" button (disabled if no filters set)
            ''20170729 LEE: This button needs to be always enabled because user can filter in non-edit mode
            'If ((StrComp(strRIFilter, "[None]") <> 0) Or (StrComp(strSTFilter, "[None]") <> 0) Or (StrComp(strDFFilter, "[None]") <> 0) Or (StrComp(txtFilter, "") <> 0)) Then
            '    Me.cmdClearFilters.Enabled = True 'Allow clear if a filter is enabled
            'Else
            '    Me.cmdClearFilters.Enabled = False
            'End If

            str1 = ""
            '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
            '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
            '10=UseIntStd, 11=IntStd, 12=MasterAssayID
            If Me.dgvAnalytes.RowCount = 0 Then
                lastAnalyteDesc = ""
                intAnalyteID = -10
                strMatrix = "NONE"
            Else
                If Me.dgvAnalytes.CurrentRow Is Nothing Then
                    int1 = 0
                Else
                    int1 = Me.dgvAnalytes.CurrentRow.Index
                End If

                strName = Me.dgvAnalytes.Item("ANALYTEDESCRIPTION", int1).Value 'ANALYTEDESCRIPTION
                intAnalyteID = NZ(Me.dgvAnalytes.Item("ANALYTEID", int1).Value, -1)
                strMatrix = NZ(Me.dgvAnalytes.Item("MATRIX", int1).Value, "NONE")

                '20160224 LEE: This should still be called if Report Table is changed
                'commenting it out because it's now allowing Report Table change event go to completion
                'We are now clearing filters with every
                If StrComp(strName, lastAnalyteDesc, CompareMethod.Text) = 0 And boolFromFilter = False And gboolFiltersCleared = False Then
                    'GoTo end1  'NDL - If the same Analyte as last time round, and no Filter has been set, we are done.
                    '20162416 LEE: Not if FillAnalyticalRuns is coming from AutoAssignSamples
                    'tblAnalyticalRuns may have been filtered with AutoAssign parameters that must be cleared
                    '20161227 LEE: Or if from ChangeStudy, must proceed
                    If boolAutoAssign Or boolFromChangeStudy Or boolSC Then
                    Else
                        GoTo end1
                    End If

                Else
                    lastAnalyteDesc = strName
                End If
            End If

            Cursor.Current = Cursors.WaitCursor

            boolClear = False
            If Me.dgvAnalytes.RowCount = 0 Then
                intRow = -1
                boolClear = True
            ElseIf Me.dgvAnalytes.CurrentRow Is Nothing Then
                intRow = 0
            Else
                intRow = Me.dgvAnalytes.CurrentRow.Index
            End If

            If intRow = -1 Then   'NDL - if no rows, then set IS to false
                boolIS = False
            Else                  'Set boolIS according to whether Analyte is an Internal Standard
                strName = Me.dgvAnalytes.Item("ANALYTEDESCRIPTION", intRow).Value 'ANALYTEDESCRIPTION
                str1 = NZ(Me.dgvAnalytes.Item("IsIntStd", intRow).Value, "No") 'IsIntStd
                If StrComp(str1, "Yes", CompareMethod.Text) = 0 Then
                    boolIS = True
                Else
                    boolIS = False
                End If
            End If

            dt1 = Now

            dgv = Me.dgvAnalyticalRuns
            Dim dv As System.Data.DataView

            ''20160224 LEE: This is inconsistent. Should set new datasource every time
            'dv = dgv.DataSource

            'Show All
            'Accepted
            'Rejected

            Dim strF1 As String

            If boolClear Then  'boolClear means there are no rows
                'strF = "ANALYTEINDEX = -1"
                strF = "ANALYTEID = -10"
            Else

                If boolIS Then   'NDL - If Analyte is an Internal Standard, do something special

                    Dim tblAn As System.Data.DataTable = Me.dgvAnalytes.DataSource.ToTable
                    Dim rowsAn() As DataRow

                    If Me.panChooseAnalyte.Visible Then
                        strF1 = "IntStd = '" & CleanText(strName) & "' AND ANALYTEDESCRIPTION = '" & CleanText(strChoose) & "'"
                    Else
                        strF1 = "IntStd = '" & CleanText(strName) & "'"
                    End If


                    rowsAn = tblAn.Select(strF1, "ANALYTEDESCRIPTION ASC")

                    '20160818 LEE:
                    'analytical run filtering is based on two fields:
                    'ANALYTICALRUN.RUNTYPEID:
                    '   1-UNKNOWNS
                    '   2-VALIDATION
                    '   3-PSAE
                    '   4-MANDATORY REPEATS
                    '   5-RECOVERY
                    'and
                    'ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS:
                    '    1-NO Regression Performed
                    '    2-Regression Performed
                    '    3-Accepted
                    '    4-Rejected
                    '    5-Created
                    '    6-Downloaded
                    '    7-Uploaded
                    '    8-Raw Analyte Info Edited
                    '    9-Edited
                    '    10-Deleted
                    '    11-Internal Standard Toggled On
                    '    12-Internal Standard Toggled Off
                    '    13-Cloned from Host Study
                    '    14-Unaccepted
                    '    15-Run Assay Edited
                    '    16-Study Assay Edited
                    '    17-Regression with alternate analyte
                    'Note that there may be more than 17 if users add custom items
                    'The most important of RUNANALYTEREGRESSIONSTATUS is:
                    '1,2,3,4,14
                    'This table includes all RUNTYPEID'S, including PSAE
                    'filter choices should be:
                    '   - Not Rejected
                    '       - RUNANALYTEREGRESSIONSTATUS: Accepted + No Regression Performed + Regression Performed:  3, 1, 2
                    '       - RUNTYPEID: RUNTYPEID <> 3
                    '   - Show All: no filter at all

                    If StrComp(strAccStatus, "Show All", CompareMethod.Binary) = 0 Then
                        var8 = " AND RUNANALYTEREGRESSIONSTATUS > 0 AND RUNTYPE > 0"
                    ElseIf InStr(1, strAccStatus, "Not Rejected", CompareMethod.Binary) > 0 Then
                        var8 = " AND ((RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3) AND RUNTYPEID <> 3)"
                    ElseIf InStr(1, strAccStatus, "Accepted", CompareMethod.Binary) > 0 Then
                        var8 = " AND RUNANALYTEREGRESSIONSTATUS = 3 AND RUNTYPEID <> 3"
                    End If

                    'filter for all IntStds
                    strF = "INTERNALSTDNAME = '" & CleanText(strName) & "'"

                    '20180129 LEE: Added ChooseAnalytes functionality to allow user to choose which Analyte/IS pair wish to be viewed
                    If Me.panChooseAnalyte.Visible Then
                        strF = "INTERNALSTDNAME = '" & strName & "' AND CHARANALYTE = '" & CleanText(strChoose) & "'"
                    Else
                        strF = "INTERNALSTDNAME = '" & CleanText(strName) & "'"
                    End If


                Else  'Analyte is NOT an Internal Standard

                    'get all runs from 'tblCalStdGroupAssayIDsAll

                    Dim boolAll As Boolean = True
                    If StrComp(strAccStatus, "Show All", CompareMethod.Binary) = 0 Then
                        'var8 = ""
                        var8 = " AND RUNANALYTEREGRESSIONSTATUS > 0"
                        boolAll = True
                    ElseIf InStr(1, strAccStatus, "Not Rejected", CompareMethod.Binary) > 0 Then
                        'var8 = " AND RUNANALYTEREGRESSIONSTATUS = 3"
                        var8 = " AND ((RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3) AND RUNTYPEID <> 3)"
                        boolAll = False
                    ElseIf InStr(1, strAccStatus, "Accepted", CompareMethod.Binary) > 0 Then
                        var8 = " AND RUNANALYTEREGRESSIONSTATUS = 3 AND RUNTYPEID <> 3"
                        boolAll = True
                    End If

                    '20161227 LEE: If this comes from ChangeStudy, then AssayID's have to come from somewhere else
                    If boolFromChangeStudy Or boolSC Or boolAll Then
                        strFAssayID = "ANALYTEID = " & intAnalyteID
                        If Len(strF) = 0 Then
                            strF = "ANALYTEID = " & intAnalyteID
                        Else
                            strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID
                        End If

                    Else
                        strFAssayID = GetASSAYIDFilter(intGroup, boolAll, False)
                        If Len(strFAssayID) = 0 Then
                        Else
                            strF = strFAssayID & " AND ANALYTEID = " & intAnalyteID
                        End If

                    End If

                    'ensure RunIDs are selected in anal run summary
                    '20170810 LEE: This logic isn't correct
                    'tblAnalysisResults should show all runs
                    'If boolFromChangeStudy Or boolSC Then
                    'Else
                    '    If Len(strF) = 0 Then
                    '    Else
                    '        strFFF = GetARSRuns(tblAnalysisResults, intAnalyteID, "")
                    '        If Len(strFFF) = 0 Then
                    '        Else
                    '            strFFF = "(" & strFFF & ")"
                    '            strF = strF & " AND " & strFFF
                    '        End If
                    '    End If
                    'End If

                End If

                If Len(strF) = 0 Then  'No samples found
                    strF = "ANALYTEINDEX = -1"
                Else
                    If IsNumeric(strRIFilter) Then 'Set RunID filter to chosen integer value
                        var3 = " AND RUNID = " & CInt(strRIFilter)
                    Else
                        var3 = " AND RUNID > 0"
                    End If
                    var4 = " AND RUNSAMPLEKIND = '" & strSTFilter & "'"
                    var5 = " AND ALIQUOTFACTOR = " & strDFFilter
                    var7 = String.Format(" AND {0} LIKE '%{1}%'", "SAMPLENAME", txtFilter)

                    'Now add filters to string
                    If StrComp(strRIFilter, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        strF = strF & var3
                    End If
                    If StrComp(strSTFilter, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        strF = strF & var4
                    End If
                    If StrComp(strDFFilter, "[None]", CompareMethod.Text) = 0 Then
                    Else
                        strF = strF & var5
                    End If
                    If StrComp(txtFilter, "", CompareMethod.Text) = 0 Then
                    Else
                        strF = strF & var7
                    End If

                    If StrComp(strAccStatus, "Show All", CompareMethod.Binary) = 0 Then
                        var8 = ""
                        var8 = " AND RUNANALYTEREGRESSIONSTATUS > 0"
                    ElseIf InStr(1, strAccStatus, "Not Rejected", CompareMethod.Binary) > 0 Then
                        'var8 = " AND RUNANALYTEREGRESSIONSTATUS = 3"
                        var8 = " AND ((RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3) AND RUNTYPEID <> 3)"
                    ElseIf InStr(1, strAccStatus, "Accepted", CompareMethod.Binary) > 0 Then
                        var8 = " AND RUNANALYTEREGRESSIONSTATUS = 3 AND RUNTYPEID <> 3"
                    End If

                    If Len(var8) = 0 Then
                    Else
                        strF = strF & var8
                    End If

                End If

            End If

            If Len(strF) = 0 Then
            Else
                If boolOriginal Then
                    strF = strF & " AND STUDYID = " & wStudyID
                Else
                    strF = strF & " AND STUDYID = " & intSI
                End If

            End If

            'console.writeline(strF)

            ''now make table
            'tblAnalysisResults.CaseSensitive = True
            'dv = New DataView(tblAnalysisResults, strF, strS, DataViewRowState.CurrentRows)

            'dv.AllowDelete = False
            'dv.AllowNew = False
            'dv.AllowEdit = False

            ''Note: Sort MUST be applied. FINDSAMPLES uses dv.Find, which requires a sort
            'strS = "CHARANALYTE ASC, STUDYNAME ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            'Try
            '    dv.Sort = strS
            'Catch ex As Exception
            '    var1 = ex.Message
            '    var1 = var1
            'End Try

            '20180801 LEE
            'if study has changed, dgv datasource must be reset
            Try
                If boolSC Then

                    'Call FillAnalysisResultsTableOutStudy(intSI)
                    'Call FilterForAnalyte(tblAnalysisResultsHomeOutStudy)

                    '20180821 LEE:
                    'This will re-set dgvAnalyticalRuns datasource
                    Call FilterForAnalyte(tblAnalysisResultsHomeOutStudy)

                    dgv = Me.dgvAnalyticalRuns

                End If
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try
           
            dv = dgv.DataSource
            int1 = dv.Count

            'Console.WriteLine("strF: " & strF)

            Try
                dv.RowFilter = strF
                dv.Sort = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            Catch ex As Exception
                var1 = ex.Message
            End Try

            var1 = dv.Count 'debug
            var1 = var1

            'dgv.DataSource = dv

            'Call OrderColumns(dgv, False)

            num = 1
            'If boolFormLoad And int1 > 0 Then
            If boolFormLoad Then
                Call InitializeAnalyticalRuns()
                Call ShowColumns()
            Else
                dgv.AutoResizeColumns()
                'dgv.AutoResizeRows()
            End If


            Call UpdateAnalRunLabel()

            Try
                '20180327 LEE
                'if dgv has selected rows, must unselect them
                If dgv.RowCount = 0 Then
                Else
                    int1 = 0
                    Do Until dgv.Columns(int1).Visible
                        If dgv.Columns(int1).Visible Then
                            Exit Do
                        End If
                        int1 = int1 + 1
                    Loop
                    dgv.CurrentCell = dgv.Rows(0).Cells(int1)
                    dgv.Rows(0).Selected = True
                    dgv.ClearSelection()
                End If
            Catch ex As Exception

            End Try

        Catch ex As Exception
            var1 = ex.Message 'debug
        End Try


        dgv.AutoResizeColumns()


end1:

        Call CountSamples()

        Cursor.Current = Cursors.Default

    End Sub

    Sub FillAccStatus()
        '20160818 LEE:
        'analytical run filtering is based on two fields:
        'ANALYTICALRUN.RUNTYPEID:
        '   1-UNKNOWNS
        '   2-VALIDATION
        '   3-PSAE
        '   4-MANDATORY REPEATS
        '   5-RECOVERY
        'and
        'ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS:
        '    1-NO Regression Performed
        '    2-Regression Performed
        '    3-Accepted
        '    4-Rejected
        '    5-Created
        '    6-Downloaded
        '    7-Uploaded
        '    8-Raw Analyte Info Edited
        '    9-Edited
        '    10-Deleted
        '    11-Internal Standard Toggled On
        '    12-Internal Standard Toggled Off
        '    13-Cloned from Host Study
        '    14-Unaccepted
        '    15-Run Assay Edited
        '    16-Study Assay Edited
        '    17-Regression with alternate analyte
        'Note that there may be more than 17 if users add custom items
        'The most important of RUNANALYTEREGRESSIONSTATUS is:
        '1,2,3,4,14
        'This table includes all RUNTYPEID'S, including PSAE
        'filter choices should be:
        '   - Not Rejected
        '       - RUNANALYTEREGRESSIONSTATUS: Accepted + No Regression Performed + Regression Performed:  3, 1, 2
        '       - RUNTYPEID: RUNTYPEID <> 3
        '   - Show All: no filter at all


        'now do acceptance status
        Me.cbxAccStatus.Items.Clear()
        Me.cbxAccStatus.Items.Add("Show All")
        Me.cbxAccStatus.Items.Add("Not Rejected") 'RUNANALYTEREGRESSIONSTATUS = 1 or 2 or 3
        Me.cbxAccStatus.Items.Add("Accepted") 'RUNANALYTEREGRESSIONSTATUS = 3
        'Me.cbxAccStatus.Items.Add("Rejected") 'RUNANALYTEREGRESSIONSTATUS = 4 or 14
        'select 2nd item
        Me.cbxAccStatus.SelectedIndex = 1

    End Sub

    Sub FilterSAMPLETYPE()

        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim Count1 As Short
        Dim var1, var2
        Dim str1 As String

        Me.cbxFilterSampleType.Items.Clear()
        tbl = Me.tblAnalysisResults
        dv = New DataView(tbl)
        dv.Sort = "RUNSAMPLEKIND ASC"
        Dim tbl1 As System.Data.DataTable = dv.ToTable("a", True, "RUNSAMPLEKIND")

        int1 = tbl1.Rows.Count
        Me.cbxFilterSampleType.Items.Add("[None]")
        Me.cbxFilterSampleType.Items.Add("QC")
        Me.cbxFilterSampleType.Items.Add("STANDARD")
        For Count1 = 0 To int1 - 1
            var2 = tbl1.Rows.Item(Count1).Item("RUNSAMPLEKIND")
            var1 = NZ(var2, "")
            If Len(var1) = 0 Or StrComp(var1, "QC", CompareMethod.Text) = 0 Or StrComp(var1, "STANDARD", CompareMethod.Text) = 0 Then
            Else
                Me.cbxFilterSampleType.Items.Add(CStr(var1))
            End If
        Next

        'select first item
        Me.cbxFilterSampleType.SelectedIndex = 0

        'now do cbxFilterDilFactor
        Me.cbxFilterDilFactor.Items.Clear()
        Dim tbl2 As System.Data.DataTable = dv.ToTable("b", True, "ALIQUOTFACTOR")
        int1 = tbl2.Rows.Count
        Me.cbxFilterDilFactor.Items.Add("[None]")
        For Count1 = 0 To int1 - 1
            var1 = tbl2.Rows.Item(Count1).Item("ALIQUOTFACTOR")
            If Len(NZ(var1, "")) = 0 Then
            Else
                str1 = CStr(var1)
                Me.cbxFilterDilFactor.Items.Add(str1)
            End If
        Next
        'select first item
        Me.cbxFilterDilFactor.SelectedIndex = 0


    End Sub

    Sub InitializeAnalyticalRuns()

        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dgv As DataGridView

        'If boolFormLoad Then
        'Else
        '    Exit Sub
        'End If

        dgv = Me.dgvAnalyticalRuns
        dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing

        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
            str1 = dgv.Columns.Item(Count1).Name
            If StrComp(str1, "RUNID", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Run ID"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            ElseIf StrComp(str1, "STUDYNAME", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Study Name"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "ELIMINATEDFLAG", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Elim." & ChrW(10) & "Flag"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "SAMPLENAME", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Sample Name"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "ALIQUOTFACTOR", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Dil." & ChrW(10) & "Factor"
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "RUNSAMPLEKIND", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Sample" & ChrW(10) & "Type"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "CONCENTRATION", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Conc.*"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
                str2 = "##0.000"
                '20160509
                dgv.Columns.Item(Count1).DefaultCellStyle.Format = str2

            ElseIf StrComp(str1, "NOMCONC", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Nom. Conc."
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "CHARANALYTE", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Analyte"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.None

            ElseIf StrComp(str1, "ANALYTEHEIGHT", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Analyte" & Chr(10) & "Peak Ht."
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "ANALYTEAREA", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Analyte" & Chr(10) & "Peak Area"
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "INTERNALSTANDARDHEIGHT", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Int. Std." & Chr(10) & "Peak Ht."
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "INTERNALSTANDARDAREA", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Int. Std." & Chr(10) & "Peak Area"
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "INTERNALSTDNAME", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Int. Std." & Chr(10) & "Name"
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "ASSAYID", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "ASSAYID"
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "ASSAYLEVEL", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Assay" & Chr(10) & "Level"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "DESIGNSAMPLEID", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Sample" & Chr(10) & "ID"
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            ElseIf StrComp(str1, "RUNSAMPLESEQUENCENUMBER", CompareMethod.Text) = 0 Then
                'dgv.Columns.Item(Count1).HeaderText = "Seq" & ChrW(10) & "#"
                dgv.Columns.Item(Count1).Visible = False
                'dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.bottomCenter
            ElseIf StrComp(str1, "RUNSAMPLEORDERNUMBER", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Seq" & ChrW(10) & "#"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            ElseIf StrComp(str1, "SAMPLETYPEID", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Matrix"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
            ElseIf StrComp(str1, "ANALYTEID", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "A_ID"
                dgv.Columns.Item(Count1).Visible = False
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            ElseIf StrComp(str1, "ASSAYDATETIME", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Analysis Date"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                str2 = "MMM dd, yyyy hh:mm:ss tt"
                str2 = "MMM dd, yyyy HH:mm:ss"
                dgv.Columns.Item(Count1).DefaultCellStyle.Format = str2

            ElseIf StrComp(str1, "ANALYTEPEAKRETENTIONTIME", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Analyte" & ChrW(10) & "RT"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            ElseIf StrComp(str1, "INTERNALSTANDARDRETENTIONTIME", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).HeaderText = "Int. Std" & ChrW(10) & "RT"
                dgv.Columns.Item(Count1).Visible = True
                dgv.Columns.Item(Count1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            ElseIf StrComp(str1, "intGroup", CompareMethod.Text) = 0 Then
                dgv.Columns.Item(Count1).Visible = False
            Else
                dgv.Columns.Item(Count1).Visible = False
                'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            End If
            'dgv.Columns.item(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
        Next

        Dim var1
        dgv.Columns.Item("CHARANALYTE").Frozen = True

        'dgv.AutoResizeColumns()

    End Sub

    Sub ShowColumnsGroupBox()

        Dim dgv1 As DataGridView = Me.dgvAnalyticalRuns
        Dim dgv2 As DataGridView = Me.dgvAssignedSamples
        Dim boolVis As Boolean
        Dim var1

        'Analysis Date
        If Me.chkAnalysisDate.Checked Then
            boolVis = True
        Else
            boolVis = False
        End If
        Try
            dgv1.Columns("ASSAYDATETIME").Visible = boolVis
        Catch ex As Exception

        End Try
        Try
            dgv2.Columns("ASSAYDATETIME").Visible = boolVis
        Catch ex As Exception

        End Try

        'Assay Level
        If Me.chkAssayLevel.Checked Then
            boolVis = True
        Else
            boolVis = False
        End If
        Try
            dgv1.Columns("ASSAYLEVEL").Visible = boolVis
        Catch ex As Exception

        End Try
        Try
            dgv2.Columns("ASSAYLEVEL").Visible = boolVis
        Catch ex As Exception

        End Try

        'FLAG
        If Me.chkFlag.Checked Then
            boolVis = True
        Else
            boolVis = False
        End If
        Try
            dgv1.Columns("ELIMINATEDFLAG").Visible = boolVis
        Catch ex As Exception

        End Try
        Try
            dgv2.Columns("ELIMINATEDFLAG").Visible = boolVis
        Catch ex As Exception

        End Try

        'chkDilFactor
        If Me.chkDilFactor.Checked Then
            boolVis = True
        Else
            boolVis = False
        End If
        Try
            dgv1.Columns("ALIQUOTFACTOR").Visible = boolVis
        Catch ex As Exception
            var1 = var1 'debug
        End Try

        Try
            dgv2.Columns("ALIQUOTFACTOR").Visible = False
        Catch ex As Exception
            var1 = var1 'debug
        End Try

        Try
            dgv2.Columns("ALIQUOTFACTOR").Visible = boolVis
        Catch ex As Exception
            var1 = var1 'debug
        End Try

        'SAMPLE TYPE
        If Me.chkSampleType.Checked Then
            boolVis = True
        Else
            boolVis = False
        End If
        Try
            dgv1.Columns("RUNSAMPLEKIND").Visible = boolVis
        Catch ex As Exception

        End Try
        Try
            dgv2.Columns("RUNSAMPLEKIND").Visible = boolVis
        Catch ex As Exception

        End Try

        'Analyte RT
        If Me.chkAnalRT.Checked Then
            boolVis = True
        Else
            boolVis = False
        End If
        Try
            dgv1.Columns("ANALYTEPEAKRETENTIONTIME").Visible = boolVis
        Catch ex As Exception
            Dim aaa
            aaa = 1
        End Try
        Try
            dgv2.Columns("ANALYTEPEAKRETENTIONTIME").Visible = boolVis
        Catch ex As Exception
            Dim aaa
            aaa = 1
        End Try

        'is RT
        If Me.chkISRT.Checked Then
            boolVis = True
        Else
            boolVis = False
        End If
        Try
            dgv1.Columns("INTERNALSTANDARDRETENTIONTIME").Visible = boolVis
        Catch ex As Exception

        End Try
        Try
            dgv2.Columns("INTERNALSTANDARDRETENTIONTIME").Visible = boolVis
        Catch ex As Exception

        End Try

        'DesignSampleID
        If Me.chkDesignSampleID.Checked Then
            boolVis = True
        Else
            boolVis = False
        End If
        Try
            dgv1.Columns("DESIGNSAMPLEID").Visible = boolVis
        Catch ex As Exception

        End Try
        Try
            dgv2.Columns("DESIGNSAMPLEID").Visible = boolVis
        Catch ex As Exception

        End Try

        If boolViewOnly Then
        Else
            Call PlaceControls(False) 'to set width of boxes
        End If

    End Sub

    Sub OrderColumns(dgv As DataGridView, boolAssignSamples As Boolean)

        Dim var1
        Dim int1 As Short = 0
        Dim str1 As String
        'do this twice

        Dim Count1 As Short
        Dim p1 As Short
        For Count1 = 1 To 5

            Try

                dgv.Columns("CHARANALYTE").DisplayIndex = int1

                int1 = int1 + 1
                dgv.Columns("INTERNALSTDNAME").DisplayIndex = int1

                int1 = int1 + 1
                dgv.Columns("STUDYNAME").DisplayIndex = int1

                If boolAssignSamples Then
                    int1 = int1 + 1
                    dgv.Columns("CHARSTUDYNAME2").DisplayIndex = int1
                End If

                int1 = int1 + 1
                dgv.Columns("RUNID").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("RUNSAMPLEORDERNUMBER").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("ASSAYLEVEL").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("SAMPLENAME").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("REPORTEDCONC").DisplayIndex = int1

                int1 = int1 + 1
                dgv.Columns("CONCENTRATION").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("ANALYTEAREA").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("INTERNALSTANDARDAREA").DisplayIndex = int1

                int1 = int1 + 1
                dgv.Columns("SAMPLETYPEID").DisplayIndex = int1 'this is matrix

                int1 = int1 + 1
                dgv.Columns("ASSAYDATETIME").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("ANALYTEPEAKRETENTIONTIME").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("INTERNALSTANDARDRETENTIONTIME").DisplayIndex = int1
                int1 = int1 + 1
                If dgv.Columns.Contains("ALIQUOTFACTOR") Then
                    str1 = "ALIQUOTFACTOR"
                ElseIf dgv.Columns.Contains("ALIQUOTFACTOR") Then
                    str1 = "ALIQUOTFACTOR"
                End If
                dgv.Columns(str1).DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("RUNSAMPLEKIND").DisplayIndex = int1 'sample type
                int1 = int1 + 1
                dgv.Columns("ELIMINATEDFLAG").DisplayIndex = int1
                int1 = int1 + 1
                dgv.Columns("DESIGNSAMPLEID").DisplayIndex = int1

                If boolAssignSamples Then

                    int1 = int1 + 1
                    p1 = int1
                    dgv.Columns.Item("BOOLEXCLSAMPLECHK").DisplayIndex = int1
                    int1 = int1 + 1
                    dgv.Columns.Item("NOMCONC").DisplayIndex = int1
                    int1 = int1 + 1
                    dgv.Columns.Item("CHARHELPER1").DisplayIndex = int1
                    int1 = int1 + 1
                    dgv.Columns.Item("CHARHELPER2").DisplayIndex = int1
                    int1 = int1 + 1
                    dgv.Columns.Item("BOOLUSEGUWUACCCRIT").DisplayIndex = int1
                    int1 = int1 + 1
                    dgv.Columns.Item("NUMMINACCCRIT").DisplayIndex = int1
                    int1 = int1 + 1
                    dgv.Columns.Item("NUMMAXACCCRIT").DisplayIndex = int1

                End If


            Catch ex As Exception
                var1 = ex.Message
            End Try

        Next

        'pesky
        var1 = dgv.Columns.Count
        dgv.Columns("CHARANALYTE").DisplayIndex = 0
        If boolAssignSamples Then
            dgv.Columns.Item("BOOLEXCLSAMPLECHK").DisplayIndex = p1
        End If


    End Sub

    Sub ShowColumns()

        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim intRow As Short
        Dim int1 As Short
        Dim Count1 As Short
        Dim boolV As Boolean
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim idTable As Long
        Dim strF As String
        Dim boolTerm As Boolean

        Dim idRT As Int32
        Dim intBSNR As Short = 0 'boolStatsNR


        tbl = tblAssignedSamplesHelper

        dgv = Me.dgvTables
        dgv1 = Me.dgvAnalyticalRuns
        int1 = dgv1.Columns.Count
        Dim int2 As Short
        int2 = dgv.RowCount

        Try
            Call OrderColumns(Me.dgvAnalyticalRuns, False)
        Catch ex As Exception

        End Try

        'For Count1 = 0 To int1 - 1 'debug
        '    ''console.writeline(dgv1.Columns(Count1).Name)
        'Next

        If int2 = 0 Then
            'Exit Sub
            boolTerm = False
            intRow = 0

        Else
            intRow = dgv.CurrentRow.Index

            '20181127 LEE:
            idRT = dgv("ID_TBLREPORTTABLE", intRow).Value
            idTable = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
            strF = "ID_TBLCONFIGREPORTTABLES = " & idTable
            rows = tbl.Select(strF)
            boolTerm = False
            If rows.Length = 0 Then
            Else
                boolTerm = True
            End If

        End If

        If boolViewOnly Then
            dgv1.Columns.Item("CONCENTRATION").Visible = True
            dgv1.Columns.Item("ANALYTEAREA").Visible = True
            dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True
            dgv1.Columns.Item("INTERNALSTDNAME").Visible = True
            'dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv1.AutoResizeColumns()
        Else
            Select Case idTable
                Case 31, 32, 33, 34, 35
                    dgv1.Columns.Item("CONCENTRATION").Visible = True
                    dgv1.Columns.Item("ANALYTEAREA").Visible = True
                    dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True
                    dgv1.Columns.Item("INTERNALSTDNAME").Visible = True
                Case 36, 37, 38
                    dgv1.Columns.Item("CONCENTRATION").Visible = True
                    dgv1.Columns.Item("ANALYTEAREA").Visible = True
                    dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True
                    dgv1.Columns.Item("INTERNALSTDNAME").Visible = False
                Case Else
                    If dgv.Rows.Count = 0 Then
                        dgv1.Columns.Item("CONCENTRATION").Visible = False
                        dgv1.Columns.Item("ANALYTEAREA").Visible = True
                        dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True
                        dgv1.Columns.Item("INTERNALSTDNAME").Visible = True
                    ElseIf dgv("BOOLSHOWAREA", intRow).Value = -1 Then
                        dgv1.Columns.Item("CONCENTRATION").Visible = False
                        dgv1.Columns.Item("ANALYTEAREA").Visible = True
                        dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True
                        dgv1.Columns.Item("INTERNALSTDNAME").Visible = True
                    ElseIf dgv("BOOLSHOWCONC", intRow).Value = -1 Then
                        dgv1.Columns.Item("CONCENTRATION").Visible = True
                        dgv1.Columns.Item("ANALYTEAREA").Visible = False
                        dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = False
                        dgv1.Columns.Item("INTERNALSTDNAME").Visible = False
                    End If
            End Select


            'wtf?? always show peak area columns
            dgv1.Columns.Item("ANALYTEAREA").Visible = True
            dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True

            int1 = dgv1.Columns.Count 'debug

            'now do dgvAssignedSamples
            dgv1 = Me.dgvAssignedSamples

            Try
                Call OrderColumns(Me.dgvAssignedSamples, True)
            Catch ex As Exception

            End Try

            If dgv1.Columns.Count = 0 Then ' Or dgv1.RowCount = 0 Then
            Else
                If intRow = 0 Then
                    dgv1.Columns.Item("CONCENTRATION").Visible = True
                    dgv1.Columns.Item("ANALYTEAREA").Visible = False
                    dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = False
                    dgv1.Columns.Item("INTERNALSTDNAME").Visible = False
                ElseIf dgv("BOOLSHOWAREA", intRow).Value = -1 Then
                    dgv1.Columns.Item("CONCENTRATION").Visible = False
                    dgv1.Columns.Item("ANALYTEAREA").Visible = True
                    dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True
                    dgv1.Columns.Item("INTERNALSTDNAME").Visible = True
                ElseIf dgv("BOOLSHOWCONC", intRow).Value = -1 Then
                    dgv1.Columns.Item("CONCENTRATION").Visible = True
                    dgv1.Columns.Item("ANALYTEAREA").Visible = False
                    dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = False
                    dgv1.Columns.Item("INTERNALSTDNAME").Visible = False
                End If
                'dgv1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
                If boolTerm Then
                    dgv1.Columns.Item("CHARHELPER1").Visible = True
                Else
                    dgv1.Columns.Item("CHARHELPER1").Visible = False
                End If

                Call ShowCritNoDB()

                int2 = dgv1.Columns.Count

                Select Case idTable
                    Case 3, 28
                        dgv1.Columns.Item("BOOLEXCLSAMPLECHK").Visible = False
                    Case Else
                        If LAllowExclSamples And gAllowExclSamples Then
                            dgv1.Columns.Item("BOOLEXCLSAMPLECHK").Visible = True
                        Else
                            dgv1.Columns.Item("BOOLEXCLSAMPLECHK").Visible = False
                        End If
                End Select


                'override for specific table
                If idTable = 17 Then
                    If boolRCConc Then
                    ElseIf boolRCPA Or boolRCPARatio Then
                        dgv1.Columns.Item("CONCENTRATION").Visible = True
                        dgv1.Columns.Item("ANALYTEAREA").Visible = True
                        dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True
                        'dgv1.Columns.Item("INTERNALSTDNAME").Visible = False
                    End If
                End If

                Select Case idTable

                    'Case 21, 17 '20190214 LEE: 21 is now adhocstability 31
                    Case 17
                        strF = "Nom." & ChrW(10) & "Conc."
                        dgv1.Columns.Item("NOMCONC").HeaderText = strF
                        dgv1.Columns.Item("NOMCONC").Visible = True
                        strF = "Run" & ChrW(10) & "Identifier"
                        dgv1.Columns.Item("CHARHELPER2").HeaderText = strF
                        dgv1.Columns.Item("CHARHELPER2").Visible = True
                    Case 22
                        dgv1.Columns.Item("NOMCONC").Visible = False
                        strF = "Stock Soln." & ChrW(10) & "Conc."
                        dgv1.Columns.Item("CHARHELPER2").HeaderText = strF
                        dgv1.Columns.Item("CHARHELPER2").Visible = True
                    Case 23
                        strF = "Std." & ChrW(10) & "Conc."
                        dgv1.Columns.Item("NOMCONC").HeaderText = strF
                        dgv1.Columns.Item("NOMCONC").Visible = True
                        dgv1.Columns.Item("CHARHELPER2").Visible = False
                    Case 3, 28
                        dgv1.Columns.Item("NOMCONC").Visible = True
                        dgv1.Columns.Item("CHARHELPER2").Visible = False
                    Case 29, 34
                        dgv1.Columns.Item("NOMCONC").Visible = True
                        dgv1.Columns.Item("CHARHELPER1").Visible = True
                        strF = "Run" & ChrW(10) & "Identifier"
                        dgv1.Columns.Item("CHARHELPER2").HeaderText = strF
                        dgv1.Columns.Item("CHARHELPER2").Visible = True
                    Case 4, 37
                        dgv1.Columns.Item("NOMCONC").Visible = True
                        dgv1.Columns.Item("CHARHELPER1").Visible = True
                        'strF = "QC Is" & ChrW(10) & "Outlier"
                        'dgv1.Columns.item("CHARHELPER2").HeaderText = strF
                        dgv1.Columns.Item("CHARHELPER2").Visible = False
                    Case 11
                        dgv1.Columns.Item("NOMCONC").Visible = True
                        dgv1.Columns.Item("CHARHELPER1").Visible = True
                        'strF = "QC Is" & ChrW(10) & "Outlier"
                        'dgv1.Columns.item("CHARHELPER2").HeaderText = strF
                        dgv1.Columns.Item("CHARHELPER2").Visible = False

                        'Case 12 '20190214 LEE: 12 is now adhocstability 31
                        'dgv1.Columns.Item("NOMCONC").Visible = True
                        'dgv1.Columns.Item("CHARHELPER1").Visible = False
                        ''strF = "QC Is" & ChrW(10) & "Outlier"
                        ''dgv1.Columns.item("CHARHELPER2").HeaderText = strF
                        'dgv1.Columns.Item("CHARHELPER2").Visible = False

                    Case 5, 30, 38
                        dgv1.Columns.Item("NOMCONC").Visible = False
                        dgv1.Columns.Item("CHARHELPER1").Visible = False
                        dgv1.Columns.Item("CHARHELPER2").Visible = False

                        'Case Is <> 3, 4, 11, 12, 21, 22, 23, 28, 29
                        '    strF = "Nom." & ChrW(10) & "Conc."
                        '    dgv1.Columns.Item("NOMCONC").HeaderText = strF
                        '    dgv1.Columns.Item("NOMCONC").Visible = True
                        '    strF = "Term 2"
                        '    dgv1.Columns.Item("CHARHELPER2").HeaderText = strF
                        '    dgv1.Columns.Item("CHARHELPER2").Visible = False
                    Case 31, 32, 12, 19, 21

                        intBSNR = GetStatsNR(idRT)
                        If intBSNR = 8 Or intBSNR = 9 Then
                            dgv1.Columns.Item("CHARHELPER1").Visible = False
                        Else
                            dgv1.Columns.Item("CHARHELPER1").Visible = True
                        End If
                        strF = "Nom." & ChrW(10) & "Conc."
                        dgv1.Columns.Item("NOMCONC").HeaderText = strF
                        dgv1.Columns.Item("NOMCONC").Visible = True
                        strF = "Run" & ChrW(10) & "Identifier"
                        dgv1.Columns.Item("CHARHELPER2").HeaderText = strF
                        dgv1.Columns.Item("CHARHELPER2").Visible = True
                        dgv1.Columns.Item("ANALYTEAREA").Visible = True
                        dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True

                    Case 33

                        dgv1.Columns.Item("NOMCONC").Visible = False
                        dgv1.Columns.Item("CHARHELPER1").Visible = False
                        dgv1.Columns.Item("CHARHELPER2").Visible = False

                        dgv1.Columns.Item("ANALYTEAREA").Visible = True
                        dgv1.Columns.Item("INTERNALSTANDARDAREA").Visible = True

                    Case Else

                        strF = "Nom." & ChrW(10) & "Conc."
                        dgv1.Columns.Item("NOMCONC").HeaderText = strF
                        dgv1.Columns.Item("NOMCONC").Visible = True
                        strF = "Term 2"
                        dgv1.Columns.Item("CHARHELPER2").HeaderText = strF
                        dgv1.Columns.Item("CHARHELPER2").Visible = False

                End Select

            End If
        End If


        Call ShowColumnsGroupBox()

    End Sub

    Sub DoThis(ByVal cmd As String)

        Dim bool As Boolean
        Dim strF As String
        Dim rows() As DataRow
        Dim boolA As Boolean

        strF = "ID_TBLPERMISSIONS = " & id_tblPermissions
        rows = tblPermissions.Select(strF)
        If rows.Length = 0 Then
            MsgBox("Guest does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
            Exit Sub
        Else
        End If

        boolA = BOOLASSIGNSAMPLES
        If boolA = 0 Then
            bool = False
        Else
            bool = True
        End If
        If bool Then
            'Call LockAssignedSamples(bool)
        Else
            MsgBox("This user does not have Edit privileges.", MsgBoxStyle.Information, "No no...")
            Exit Sub
        End If


        bool = True
        Select Case cmd
            Case "Edit"
                Call SetToEditMode()
                bool = False

            Case "Save", "Cancel"
                Call SetToNonEditMode()
                bool = True
        End Select

        Call LockAssignedSamples(bool)
        Call ShowCrit1()
        Call CheckUseGuWuAccCrit()

    End Sub


    Function assignSamplesVerify(boolFromOK As Boolean) As Boolean

        '20160815 LEE: Need to differentiate the source of assignSamplesVerify

        'Verify that the samples in the current table have the correct information in them.
        '(Previously assignSamplesVerifyA)

        Dim row As DataRowView
        Dim boolRunIDsDiffer As Boolean
        Dim int1 As Int16
        Dim int2 As Int16
        Dim intRunID0 As Int16
        Dim intTableID As Int64
        Dim Count1 As Integer
        Dim strTableName As String
        Dim intRow As Short
        Dim var1, var2

        assignSamplesVerify = True

        If boolAutoAssign Then
            Exit Function
        End If

        'Verify that the current table has samples which fit its criteria

        If boolFromOK Then
            intRow = Me.dgvTables.CurrentRow.Index ' Me.txtdgvReportTableCurrentRow.Text
        Else
            intRow = Me.txtdgvReportTablePreviousRow.Text ' Me.txtdgvReportTableCurrentRow.Text
        End If

        intTableID = Me.dgvTables("ID_TBLCONFIGREPORTTABLES", intRow).Value
        strTableName = Me.dgvTables("CHARHEADINGTEXT", intRow).Value

        '20151209 LEE: intRow is for the changed table. Needs to be intRow of the previous table


        Try

            'Dim intRow As Integer
            'intRow = Me.txtdgvReportTablePreviousRow.Text ' dgvTables.CurrentRow.Index
            'intSID = id_tblStudies
            '
            'idT = Me.dgvTables("ID_TBLREPORTTABLE", intRow).Value
            'strTableName = Me.dgvTables("CHARHEADINGTEXT", intRow).Value

            Dim dv As DataView
            dv = Me.dgvAssignedSamples.DataSource
            int1 = dv.Count
            If int1 = 0 Then
                GoTo end1
            End If

            'Do appropriate checks on those rows
            If int1 > 0 Then
                intRunID0 = dv(0).Item("RUNID")
            End If

            boolRunIDsDiffer = False
            For Count1 = 1 To int1

                int2 = dv(Count1 - 1).Item("RUNID")
                If int2 = intRunID0 Then
                Else
                    boolRunIDsDiffer = True
                    Exit For
                End If

            Next

            Count1 = 0
            var1 = dv.Count 'debug
            Count1 = 0

            Dim boolDo As Boolean
            For Each row In dv

                Select Case intTableID
                    Case 3, 36 'Samples with Nominal Concentration only needed
                        '3	Summary of Back-Calculated Calibration Std Conc
                        '36	Method Trial Back-Calculated Calibration Std Conc v1
                        If (Not (checkNomConc(row, strTableName, Count1 + 1))) Then
                            assignSamplesVerify = False
                        End If
                    Case 4, 11, 13, 14, 15, 17, 18, 19, 21, 23, 31, 32, 12 'Samples with Nominal Concentration and Term1 needed
                        '4	Summary of Interpolated QC Std Conc
                        '11	Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
                        '13	Summary of Combined Recovery
                        '14	Summary of True Recovery
                        '15	Summary of Suppression/Enhancement
                        '17	Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments
                        '18	Summary of [Period Temp] Stability in Matrix
                        '19	Summary of Freeze/Thaw [#Cycles] Stability in Matrix
                        '21	[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
                        '23 [Period Temp] Spiking Solution Stability Assessment
                        '31	Ad Hoc QC Stability Table
                        '32	Ad Hoc QC Stability Comparison Table

                        boolDo = True

                        Select Case intTableID
                            Case 31, 32, 12, 19, 21 'AdHoc Stability and Stability Comparison: if BOOLSTATSLETTER = 8 (Stock Soln Stab) or 9 (Spiking Soln Stab), don't check for NomConc or Term1
                                Try
                                    var1 = row.Item("ID_TBLREPORTTABLE")
                                    var2 = GetStatsNR(var1)
                                    If var2 = 8 Or var2 = 9 Then
                                        boolDo = False
                                    End If
                                Catch ex As Exception
                                    var1 = ex.Message
                                    var1 = var1
                                End Try

                        End Select

                        If boolDo Then
                            If (Not (checkNomConc(row, strTableName, Count1 + 1))) Then
                                assignSamplesVerify = False
                            ElseIf (Not (checkTerm1(row, strTableName, Count1 + 1))) Then
                                assignSamplesVerify = False
                            End If
                        End If

                        'Case 12 'Samples with Nominal Concentration and Dilution needed
                        '    '12	Summary of Interpolated Dilution QC Concentrations
                        '    If (Not (checkNomConc(row, strTableName, Count1 + 1))) Then
                        '        assignSamplesVerify = False
                        '    ElseIf (Not (checkDiluted(row, strTableName, Count1 + 1))) Then
                        '        assignSamplesVerify = False
                        '    End If
                    Case 29 'Samples with Term1 & Term2, and Nominal Concentration needed
                        '29	[Period Temp] Long-Term QC Std Storage Stability

                        If (Not (checkNomConc(row, strTableName, Count1 + 1))) Then
                            assignSamplesVerify = False
                        ElseIf (Not (checkTerm1(row, strTableName, Count1 + 1))) Then
                            assignSamplesVerify = False
                        ElseIf Not (checkTerm2(row, strTableName, Count1 + 1, intTableID)) Then
                            assignSamplesVerify = False
                            GoTo end1
                        End If
                    Case 22 'Samples with Term1 & Term2 needed, but no Nominal Concentration
                        '22	[Period Temp] Stock Solution Stability Assessment
                        If (Not (checkTerm1(row, strTableName, Count1 + 1))) Then
                            assignSamplesVerify = False
                        ElseIf Not (checkTerm2(row, strTableName, Count1 + 1, intTableID)) Then
                            assignSamplesVerify = False
                            GoTo end1
                        End If
                    Case 35 'Samples with Term1 needed, but no Nominal Concentration
                        '35	Carryover in Individual Lots Table v1
                        If (Not (checkTerm1(row, strTableName, Count1 + 1))) Then
                            assignSamplesVerify = False
                            GoTo end1
                        End If
                    Case Else
                        '1	Summary of Analytical Runs
                        '2	Summary of Regression Constants
                        '5	Summary of Samples
                        '6	Summary of Reassayed Samples
                        '7	Summary of Repeat Samples
                        '30 Incurred Samples
                        '33	System Suitability Table v1
                        '34	Selectivity in Individual Lots Table v1
                        '37	Method Trial Control and Fortified QC Samples v1
                        '38	Method Trial Incurred Blinded Samples v1
                End Select
                If assignSamplesVerify = False Then
                    GoTo end1
                End If

                Count1 = Count1 + 1

            Next

            'Now: Check the Tables for Readiness
            Select Case intTableID
                Case 13
                    '13	Summary of Combined Recovery
                    'Requirements for combined recovery table are: 
                    ' (a) Only two types of samples, 'QC' and 'RS - Recovery Standard', should exist, and they should exist in every run.
                    ' [Note: For (b), I am just checking total numbers for now: can put in nominal concentration level matches later]
                    If (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "QC", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    ElseIf (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "RS - Recovery Standard", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    End If
                    If (Not (boolCheckAtMostXUniqueTerm1sInTable(dv, 2, strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    End If


                Case 14
                    '14	Summary of True Recovery
                    'Requirements for true recovery table are: 
                    ' (a) Only two types of samples, 'QC' and 'PES - Post Extraction', should exist, and they should exist in every run.

                    If (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "QC", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    ElseIf (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "PES - Post Extraction Spike", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    End If

                    If (Not (boolCheckAtMostXUniqueTerm1sInTable(dv, 2, strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    End If

                Case 15
                    '15	Summary of Suppression/Enhancement
                    'Requirements for true recovery table are: 
                    ' (a) Only two types of samples, 'RS - Recovery Standard' and 'PES - Post Extraction', should exist, and they should exist in every run.

                    If (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "RS - Recovery Standard", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    ElseIf (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "PES - Post Extraction Spike", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    End If

                    If (Not (boolCheckAtMostXUniqueTerm1sInTable(dv, 2, strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    End If

                    'Check tables where all samples have to be from a single RunID

                Case 22         '22 Stock Solution Stability Assessment
                    'Requirements for true recovery table are: 
                    ' (a) Only two types of samples, 'Old Stock Solution' and 'New Stock Solution', should exist, and they should exist in every run.
                    If (Not (boolSampleWithTerm1ExistsInTable(dv, "Old Stock Solution", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    ElseIf (Not (boolSampleWithTerm1ExistsInTable(dv, "New Stock Solution", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    End If

                Case 23                         '23 [Period Temp] Spiking Solution Stability Assessment
                    If (boolRunIDsDiffer) Then
                        Dim str1
                        str1 = "The subject table requires all Assigned Samples to be from the same Analytical Run." & ChrW(10) & ChrW(10)
                        str1 = str1 & "Please ensure all the samples assigned to this table are from the same Analytical Run before saving."
                        str1 = str1 & ChrW(10) & ChrW(10) & "Table: " & strTableName

                        MsgBox(str1, vbInformation, "Invalid action...")

                        assignSamplesVerify = False
                        GoTo end1
                    End If

                    'Check tables where, in each run, there has to be an equal number of two types of samples

                Case 35                         '35 
                    'Requirements for carryover table are: 
                    '   (a) All samples must be assigned a Term1 (already checked)
                    '   (b) LLOQ, blanks, and 3rd Term1 must all exist
                    '   (c) There must be *only* one 3rd Term1.  NOTE: 13-Jan-2016 currently, it must be ULOQ
                    '   (d) For every 3rd Term1 type, there must be a corresponding blank appearing sequentially right after it.  [not yet covered]
                    '       20180810 LEE: Not true. Sometimes a plate is injected in non-sequential order, making it appear a blank is not run immediately after a ULOQ
                    '   (e) Every blank must follow directly after a 3rd Term (or, in the future, a blank [not yet covered]
                    '       20180810 LEE: Not true. Sometimes a plate is injected in non-sequential order, making it appear a blank is not run immediately after a ULOQ

                 
                    If (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "LLOQ", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    ElseIf (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "Blank", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    Else
                        Dim vGo = GetTableProp("boolIncludePSAE")
                        If IsNumeric(vGo) Then 'this is ULOQ column in Carryover 35
                            '20180810 LEE:
                            'Remember, boolIncludePSAE in NOT
                            If vGo = 0 Then
                                If (Not (boolSampleWithTerm1ExistsInTableInEachRun(dv, "ULOQ", strTableName))) Then
                                    assignSamplesVerify = False
                                    GoTo end1
                                End If
                            End If

                        End If
                    End If

                    ' ''Exactly 3 types of samples allowed
                    ''20180110 LEE: Deprecate. See above
                    'If (Not (boolCheckExactlyXUniqueTerm1sInTable(dv, 3, strTableName))) Then
                    '    assignSamplesVerify = False
                    '    GoTo end1
                    'End If

                    Dim tblTerm1s As System.Data.DataTable = dv.ToTable("Term1's", True, "CHARHELPER1")

                    'Find which one is not LLOQ or Blank
                    Dim Count10 As Short
                    Dim str1, strTerm1HighConc As String
                    For Count10 = 0 To tblTerm1s.Rows.Count - 1
                        str1 = tblTerm1s.Rows(Count10).Item("CHARHELPER1").ToString
                        If (Not (StrComp(str1, "LLOQ") = 0) And Not (StrComp(str1, "Blank") = 0)) Then
                            strTerm1HighConc = str1
                        End If
                    Next

                    'Must be same number of ULOQ's and Blanks in each run.
                    If (Not (boolEqualNumberOfSampleTypeInstancesInEachRunInTable(dv, strTerm1HighConc, "Blank", strTableName))) Then
                        assignSamplesVerify = False
                        GoTo end1
                    End If

                    'boolCarryOverTableBlanksAfterHighConcInEachRunInTable(dv, "ULOQ", "Blank")
            End Select

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
            assignSamplesVerify = True

        End Try

end1:


    End Function

    Function GetTableProp(strP As String)

        Dim dgv As DataGridView
        Dim intRow As Int16
        Dim idC As Int32
        Dim idRT As Int32
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim idT As Int64

        dgv = Me.dgvTables

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        Try
            idC = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
            idRT = dgv("ID_TBLREPORTTABLE", intRow).Value
        Catch ex As Exception
            idC = 0
        End Try

        dtbl = tblTableProperties
        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTTABLE = " & idRT
        rows = dtbl.Select(strF)

        If rows.Length = 0 Then
            GetTableProp = 0
        Else
            GetTableProp = NZ(rows(0).Item(strP), 0)
        End If





    End Function

    Function getSampleReferenceA(row As DataRowView, strTableName As String, intIndex As Short) As String

        Dim str1 As String

        'str1 = "Table:" & ChrW(9) & ChrW(9) & strTableName & ChrW(10)
        'str1 = str1 & "Run ID:" & ChrW(9) & ChrW(9) & row.Item("RUNID") & ChrW(10)
        'str1 = str1 & "Analyte:" & ChrW(9) & ChrW(9) & row.Item("CHARANALYTE") & ChrW(10)
        'str1 = str1 & "Sequence #:" & ChrW(9) & row.Item("RUNSAMPLESEQUENCENUMBER") & ChrW(10)
        'str1 = str1 & "Sample:" & ChrW(9) & ChrW(9) & row.Item("SAMPLENAME") & ChrW(10)
        'str1 = str1 & ChrW(10) & "Please fix this sample before Saving."

        str1 = "Row #:" & ChrW(9) & ChrW(9) & intIndex & ChrW(10)
        str1 = str1 & "Run ID:" & ChrW(9) & ChrW(9) & row.Item("RUNID") & ChrW(10)
        str1 = str1 & "Sequence #:" & ChrW(9) & row.Item("RUNSAMPLESEQUENCENUMBER") & ChrW(10)
        str1 = str1 & "Sample:" & ChrW(9) & ChrW(9) & row.Item("SAMPLENAME") & ChrW(10)
        str1 = str1 & "Table:" & ChrW(9) & ChrW(9) & strTableName & ChrW(10)


        getSampleReferenceA = str1


    End Function

    Private Function getSampleReference(ByRef row As DataRowView, ByRef strTableName As String) As String
        getSampleReference = New String( _
            "Table:" & ChrW(9) & strTableName & vbCrLf & _
                   "Analyte:" & ChrW(9) & row.Item("CHARANALYTE") & vbCrLf & _
                   "Run ID:" & ChrW(9) & row.Item("RUNID") & vbCrLf & " Sequence #: " & row.Item("RUNSAMPLESEQUENCENUMBER") & vbCrLf & _
                   "Sample:" & ChrW(9) & row.Item("SAMPLENAME") & vbCrLf & _
                   "Table:" & ChrW(9) & strTableName & vbCrLf & _
                   ChrW(10) & "Please fix this sample before Saving.")


    End Function

    Private Function checkNomConc(ByRef row As DataRowView, ByRef strTableName As String, intIndex As Short) As Boolean

        checkNomConc = True

        Dim var2
        var2 = row.Item("NOMCONC")
        If IsNothing(var2) Then
            Exit Function
        End If

        If (StrComp(row.Item("NOMCONC").ToString, "") = 0) Then
            checkNomConc = False
            MsgBox("A sample in the Assigned Samples table is missing its NOMINAL CONCENTRATION value." & vbCrLf & vbCrLf _
   & getSampleReferenceA(row, strTableName, intIndex), vbInformation, "Invalid action...")

            Dim var1
            var1 = var1 'debug

        End If

    End Function

    Private Function checkTerm1(ByRef row As DataRowView, ByRef strTableName As String, intIndex As Short) As Boolean
        checkTerm1 = True

        Dim var2
        var2 = row.Item("CHARHELPER1")
        If IsNothing(var2) Then
            Exit Function
        End If

        Dim str1 As String
        str1 = "A sample in the Assigned Samples table is missing its TERM 1 value. " & vbCrLf & vbCrLf & getSampleReferenceA(row, strTableName, intIndex)

        If (StrComp(row.Item("CHARHELPER1").ToString, "") = 0) Then
            checkTerm1 = False
            MsgBox(str1, vbInformation, "Invalid action...")

            Dim var1
            var1 = var1 'debug

        End If
    End Function

    Private Function checkDiluted(ByRef row As DataRowView, ByRef strTableName As String, intIndex As Short) As Boolean
        checkDiluted = True

        Dim var1, var2

        Try
            var2 = row.Item("ALIQUOTFACTOR") 'comes from assigned samples
            If IsNothing(var2) Then
                Exit Function
            End If


            If (row.Item("ALIQUOTFACTOR").Equals(1.0)) Then 'comes from assigned samples
                checkDiluted = False

                MsgBox("One of the Samples in the Assigned Samples table isn't diluted in a table requiring Diluted samples. " & vbCrLf & vbCrLf _
                       & getSampleReferenceA(row, strTableName, intIndex))


                var1 = var1 'debug
            End If
        Catch ex As Exception
            var1 = var1
        End Try


    End Function

    Private Function checkTerm2(ByRef row As DataRowView, ByRef strTableName As String, intIndex As Short, intTableID As Int64) As Boolean
        checkTerm2 = True

        Select Case intTableID
            Case 22 '[Period Temp] Stock Solution Stability Assessment
                '20170810 LEE:
                'This table actually doesn't need charHelper 2
                GoTo end1

        End Select

        Dim var2
        var2 = row.Item("CHARHELPER2")
        If IsNothing(var2) Then
            Exit Function
        End If

        Dim strM As String

        If (StrComp(row.Item("CHARHELPER2").ToString, "") = 0) Then
            checkTerm2 = False

            'Select Case intTableID
            '    Case 22 '[Period Temp] Stock Solution Stability Assessment
            '        '20170810 LEE:
            '        'This table actually doesn't need charHelper 2

            '        'strM = "A sample in the Assigned Samples table is missing its 'Stock Solution Concentration or Label' value (may be the Run Identifier, Stock Sol'n, etc)." & vbCrLf & vbCrLf & getSampleReferenceA(row, strTableName, intIndex)
            '    Case Else
            '        strM = "A sample in the Assigned Samples table is missing its Term 2 value (may be the Run Identifier, Stock Sol'n, etc)." & vbCrLf & vbCrLf & getSampleReferenceA(row, strTableName, intIndex)
            'End Select

            strM = "A sample in the Assigned Samples table is missing its Term 2 value (may be the Run Identifier, Stock Sol'n, etc)." & vbCrLf & vbCrLf & getSampleReferenceA(row, strTableName, intIndex)

            MsgBox(strM, vbInformation, "Invalid action...")
            Dim var1
            var1 = var1 'debug

        End If

end1:


    End Function

    Private Function boolSampleWithTerm1ExistsInTableInEachRun(ByRef dv As DataView, ByVal strTerm1CheckValue As String, _
                                                               ByVal strTableName As String) As Boolean
        'NDL: This function checks that at least one sample with the TERM1 specified exists *for each run* included in this table.

        Dim row As DataRowView
        Dim tbl As DataTable = dv.ToTable(True, "RUNID")
        Dim count As Short
        Dim boolTypeExistsInThisRun As Boolean

        boolSampleWithTerm1ExistsInTableInEachRun = True
        For count = 0 To tbl.Rows.Count - 1 'For each RunID 
            boolTypeExistsInThisRun = False
            For Each row In dv
                If (StrComp(tbl.Rows(count).Item("RunID"), row.Item("RUNID"), CompareMethod.Text) = 0) Then 'If this sample is in this run
                    If (StrComp(row.Item("CHARHELPER1"), strTerm1CheckValue, CompareMethod.Text) = 0) Then  'If the Term1 is the value we're looking for
                        boolTypeExistsInThisRun = True
                        '20160602 LEE: Don't need to loop through entire dv if true
                        Exit For
                    End If
                End If
            Next
            If Not (boolTypeExistsInThisRun) Then  'We've been through all the runs; no sample with the Term1 in this run.
                boolSampleWithTerm1ExistsInTableInEachRun = False

                Dim str1
                str1 = "For at least one of the runs included, no " & strTerm1CheckValue _
                        & " samples are set.  This table requires " _
                        & strTerm1CheckValue & " samples to be assigned." & ChrW(10)
                str1 = str1 & "Please ensure " & strTerm1CheckValue & " values are assigned, and the 'Term 1' for the " _
                        & strTerm1CheckValue & " values is set correctly."
                str1 = str1 & ChrW(10) & ChrW(10) & "Table: " & strTableName
                MsgBox(str1, vbInformation, "Invalid action...")

                Exit Function
            End If
        Next
    End Function

    Private Function boolSampleWithTerm1ExistsInTable(ByRef dv As DataView, ByVal strTerm1CheckValue As String, _
                                                              ByVal strTableName As String) As Boolean
        'NDL: This function checks that at least one sample with the TERM1 specified exists in this table.

        Dim row As DataRowView
        Dim tbl As DataTable = dv.ToTable(True, "RUNID")
        Dim count As Short
        Dim var1

        boolSampleWithTerm1ExistsInTable = False
        '20160321 LEE: This should fire if false, not true

        'For Each row In dv
        '    var1 = NZ(row.Item("CHARHELPER1"), "")
        '    If StrComp(row.Item("CHARHELPER1"), strTerm1CheckValue) = 0 Then  'If the Term1 is the value we're looking for
        '        '20160321 LEE: This should fire if false, not true
        '        boolSampleWithTerm1ExistsInTable = True
        '        Dim str1
        '        str1 = "For this table, no " & strTerm1CheckValue _
        '                & " samples are set.  This table requires " _
        '                & strTerm1CheckValue & " samples to be assigned." & ChrW(10)
        '        str1 = str1 & "Please ensure " & strTerm1CheckValue & " values are assigned, and the 'Term 1' for the " _
        '                & strTerm1CheckValue & " values is set correctly."
        '        str1 = str1 & ChrW(10) & ChrW(10) & "Table: " & strTableName
        '        MsgBox(str1, vbInformation, "Invalid action...")
        '        boolSampleWithTerm1ExistsInTable = False

        '        Exit Function
        '    End If
        'Next

        For Each row In dv
            var1 = row.Item("CHARHELPER1")
            If StrComp(row.Item("CHARHELPER1"), strTerm1CheckValue) = 0 Then  'If the Term1 is the value we're looking for
                boolSampleWithTerm1ExistsInTable = True
                Exit Function
            End If
        Next

        If boolSampleWithTerm1ExistsInTable Then
        Else
            Dim str1
            str1 = "For this table, no " & strTerm1CheckValue _
                    & " samples are set.  This table requires " _
                    & strTerm1CheckValue & " samples to be assigned." & ChrW(10)
            str1 = str1 & "Please ensure " & strTerm1CheckValue & " values are assigned, and the 'Term 1' for the " _
                    & strTerm1CheckValue & " values is set correctly."
            str1 = str1 & ChrW(10) & ChrW(10) & "Table: " & strTableName
            MsgBox(str1, vbInformation, "Invalid action...")
            boolSampleWithTerm1ExistsInTable = False
        End If

    End Function

    Private Function boolCheckExactlyXUniqueTerm1sInTable(ByRef dv As DataView, ByRef intCheck As Integer, _
                                                          ByRef strTableName As String) As Boolean
        boolCheckExactlyXUniqueTerm1sInTable = True
        If Not (boolCheckAtMostXUniqueTerm1sInTable(dv, intCheck, strTableName)) Then
            boolCheckExactlyXUniqueTerm1sInTable = False 'too many types
            Exit Function
        ElseIf Not (boolCheckAtLeastXUniqueTerm1sInTable(dv, intCheck, strTableName)) Then
            boolCheckExactlyXUniqueTerm1sInTable = False 'too few types
            Exit Function
        Else
        End If

    End Function

    Private Function boolCheckAtLeastXUniqueTerm1sInTable(ByRef dv As DataView, ByRef intCheck As Integer, _
                                                      ByRef strTableName As String) As Boolean
        Dim tbl As DataTable = dv.ToTable("temp", True, "CHARHELPER1")
        boolCheckAtLeastXUniqueTerm1sInTable = True
        If (tbl.Rows.Count < intCheck) Then
            boolCheckAtLeastXUniqueTerm1sInTable = False
            Dim str1
            str1 = "This recovery table requires that at least " & intCheck.ToString & " specific types of 'Term 1' be assigned." & ChrW(10)
            str1 = str1 & "Please ensure that only " & intCheck.ToString & " types are shown in 'Term 1' for this table."
            str1 = str1 & ChrW(10) & ChrW(10) & "Table: " & strTableName
            MsgBox(str1, vbInformation, "Invalid action...")
        End If
    End Function

    Private Function boolCheckAtMostXUniqueTerm1sInTable(ByRef dv As DataView, ByRef intCheck As Integer, _
                                                          ByRef strTableName As String) As Boolean
        Dim tbl As DataTable = dv.ToTable("temp", True, "CHARHELPER1")
        boolCheckAtMostXUniqueTerm1sInTable = True
        If (intCheck < tbl.Rows.Count) Then
            boolCheckAtMostXUniqueTerm1sInTable = False
            Dim str1
            str1 = "This recovery table requires that at most " & intCheck.ToString & " specific types of 'Term 1' be assigned." & ChrW(10)
            str1 = str1 & "Please ensure that only " & intCheck.ToString & " types are shown in 'Term 1' for this table."
            str1 = str1 & ChrW(10) & ChrW(10) & "Table: " & strTableName
            MsgBox(str1, vbInformation, "Invalid action...")
        End If
    End Function

    Private Function boolEqualNumberOfSampleTypeInstancesInEachRunInTable(ByRef dv As DataView, _
                ByVal strTerm1CheckValue1 As String, ByVal strTerm1CheckValue2 As String, _
                ByVal strTableName As String) As Boolean
        Dim row As DataRowView
        Dim tblRow As DataRow
        Dim tbl As DataTable = dv.ToTable(True, "RUNID")
        Dim count, countV1, countV2 As Short

        boolEqualNumberOfSampleTypeInstancesInEachRunInTable = True

        For count = 0 To tbl.Rows.Count - 1 'For each RunID
            'Set counters at 0
            countV1 = 0
            countV2 = 0

            For Each row In dv
                'Count the two terms
                If (StrComp(tbl.Rows(count).Item("RunID"), row.Item("RUNID")) = 0) Then
                    If (StrComp(row.Item("CHARHELPER1"), strTerm1CheckValue1) = 0) Then
                        countV1 = countV1 + 1
                    ElseIf (StrComp(row.Item("CHARHELPER1"), strTerm1CheckValue2) = 0) Then
                        countV2 = countV2 + 1
                    End If
                End If
            Next
            If (countV1 <> countV2) And (countV2 < countV1) Then '20160823 LEE: must allow for multiple blanks
                boolEqualNumberOfSampleTypeInstancesInEachRunInTable = False
                Dim str1
                str1 = "This table requires that for each Run, for each '" & strTerm1CheckValue1 & _
                    "', there is a matching '" & strTerm1CheckValue2 & "' (and vice versa)." & ChrW(10)
                str1 = str1 & "Please check the 'Term 1' values of all the runs in the table for these assignments."
                str1 = str1 & ChrW(10) & ChrW(10) & "Table: " & strTableName
                MsgBox(str1, vbInformation, "Invalid action...")

                Exit Function
            End If

        Next

        boolEqualNumberOfSampleTypeInstancesInEachRunInTable = True

    End Function

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click

        'Note: this button is actually cmdSave

        Cursor.Current = Cursors.WaitCursor

        '****

        Dim tUserID As String
        Dim tUserName As String
        Dim strM As String
        Dim intR As Short

        tUserID = gUserID
        tUserName = gUserName

        strRFC = GetDefaultRFC()
        strMOS = GetDefaultMOS()

        gATAdds = 0
        gATDeletes = 0
        gATMods = 0

        If gboolAuditTrail And gboolESig Then

            Dim frm As New frmESig

            frm.ShowDialog()

            If frm.boolCancel Then
                frm.Dispose()
                GoTo end1
            End If

            gUserID = frm.tUserID
            gUserName = frm.tUserName

            frm.Dispose()

        End If

        Dim dt1 As DateTime
        dt1 = Now
        '****

        boolCancel = False
        boolX = True

        Me.dgvAssignedSamples.EndEdit(True)

        Dim intRow As Short
        'record recorded row in dgvTables
        If Me.dgvTables.Rows.Count = 0 Then
            intRow = -1
        ElseIf Me.dgvTables.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = Me.dgvTables.CurrentRow.Index
        End If

        'Dim dvCheck as system.data.dataview = New DataView(tblAssignedSamples)
        'dvCheck.RowStateFilter = DataViewRowState.ModifiedCurrent
        'If dvCheck.Count = 0 Then
        'Else

        'clear audittrailtemp
        tblAuditTrailTemp.Clear()
        idSE = 0

        Call FillAuditTrailTemp(tblAssignedSamples)

        If boolGuWuOracle Then
            Try
                ta_tblAssignedSamples.Update(tblAssignedSamples)
            Catch ex As DBConcurrencyException
                ds2005.TBLASSIGNEDSAMPLES.Merge(ds2005.TBLASSIGNEDSAMPLES, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblAssignedSamplesAcc.Update(tblAssignedSamples)
            Catch ex As DBConcurrencyException
                ds2005Acc.TBLASSIGNEDSAMPLES.Merge(ds2005Acc.TBLASSIGNEDSAMPLES, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblAssignedSamplesSQLServer.Update(tblAssignedSamples)
            Catch ex As DBConcurrencyException
                ds2005Acc.TBLASSIGNEDSAMPLES.Merge(ds2005Acc.TBLASSIGNEDSAMPLES, True)
            End Try
        End If
        'End If
        'ta_tblAssignedSamples.Fill(tblAssignedSamples)

        If intRow = -1 Then
        Else
            Me.dgvTables.CurrentCell = Me.dgvTables.Rows.Item(intRow).Cells("CHARHEADINGTEXT")
        End If

        'Call FillAssignedSamplesDGV()

        ''first reset
        'Dim boolDo As Boolean = False
        'Dim lng1 As Int64
        'lng1 = Me.txtStudyID.Text
        'If lng1 = id_tblStudies Then 'ignore
        'Else
        '    strM = "The study has been changed from original:"
        '    strM = strM & ChrW(10) & ChrW(10) & ChrW(9) & "'" & wWStudyName & "'"
        '    strM = strM & ChrW(10) & ChrW(10) & "to" & ChrW(10) & ChrW(10) & ChrW(9) & "'" & nStudyName & "'."
        '    strM = strM & ChrW(10) & ChrW(10) & "The study will be set back to the original '" & wWStudyName & "'."
        '    MsgBox(strM, vbInformation, "Study will be set to original...")
        '    boolDo = True
        '    'Call ReturnStudyToOriginal()
        'End If

        'record tblaudittrailtemp
        Call RecordAuditTrail(False, dt1)

        Call DoThis("Save")

        'first reset
        Dim boolDo As Boolean = False
        Dim lng1 As Int64
        lng1 = Me.txtStudyID.Text
        If lng1 = id_tblStudies Then 'ignore
        Else
            strM = "The study has been changed from original:"
            strM = strM & ChrW(10) & ChrW(10) & ChrW(9) & "'" & wWStudyName & "'"
            strM = strM & ChrW(10) & ChrW(10) & "to" & ChrW(10) & ChrW(10) & ChrW(9) & "'" & nStudyName & "'."
            strM = strM & ChrW(10) & ChrW(10) & "The study will be set back to the original '" & wWStudyName & "'."
            MsgBox(strM, vbInformation, "Study will be set to original...")
            boolDo = True
            'Call ReturnStudyToOriginal()
        End If


        If boolDo Then
            Call ReturnClick()
        End If

        Call AssessSampleAssignment()
        'Call AssessSampleAssignmentAnalyte()

        'Verify Sample Assignments on Current Table - but continue saving after giving error message.
        'This is also called when changing tables, so that all tables are checked during the user's session.
        Call assignSamplesVerify(True)
end1:

        Cursor.Current = Cursors.Default


    End Sub

    Sub LockAssignedSamples(ByVal bool)

        'Me.cmdOK.Enabled = Not (bool)
        'Me.cmdReset.Enabled = Not (bool)

        Me.cmdRemove.Enabled = Not (bool)
        Me.cmdRemove.Enabled = Not (bool)
        Me.cmdAddRows.Enabled = Not (bool)
        Me.cmdCopy.Enabled = Not (bool)
        Me.cmdHelper1.Enabled = Not (bool)
        Me.cmdHelper2.Enabled = Not (bool)
        Me.cmdNomConc.Enabled = Not (bool)
        Me.txtHelper2.Enabled = Not (bool)

        Me.cmdReturn.Enabled = Not (bool)
        Me.cbxStudy.Enabled = Not (bool)
        'Me.cmdIncurred.Enabled = Not (bool)
        Me.cmdIncurred.Visible = False

        Me.chkUSEGUWUACCCRIT.Enabled = Not (bool)
        Me.chkAsynch.Enabled = Not (bool)
        Me.cmdFillDown.Enabled = Not (bool)
        Me.cmdClearAll.Enabled = Not (bool)
        Me.cmdClearDown.Enabled = Not (bool)

        Dim dgv As DataGridView
        Dim Count1 As Short

        dgv = Me.dgvAssignedSamples
        dgv.ReadOnly = bool
        For Count1 = 0 To dgv.Columns.Count - 1
            dgv.Columns(Count1).ReadOnly = True
        Next
        Try
            Me.dgvAssignedSamples.Columns("BOOLEXCLSAMPLECHK").ReadOnly = bool
        Catch ex As Exception

        End Try
        Try
            Me.dgvAssignedSamples.Columns("NUMMINACCCRIT").ReadOnly = bool
        Catch ex As Exception

        End Try
        Try
            Me.dgvAssignedSamples.Columns("NUMMAXACCCRIT").ReadOnly = bool
        Catch ex As Exception

        End Try

        'Me.dgvAssignedSamples.Columns.item("CHARHELPER1").ReadOnly = bool
        'Me.dgvAssignedSamples.Columns.item("CHARHELPER2").ReadOnly = bool


    End Sub


    Sub ChangedgvTables()

        Dim idCRT As Int64
        Dim idTR As Int64
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim boolDo As Boolean
        Dim str1 As String
        Dim str2 As String

        dgv = Me.dgvTables

        Dim boolViewCBX As Boolean = False

        If dgv.Rows.Count = 0 Then
            boolDo = False
        Else
            If dgv.CurrentRow Is Nothing Then
                boolDo = False
            Else
                intRow = dgv.CurrentRow.Index
                idCRT = NZ(dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value, 1)
                idTR = NZ(dgv("ID_TBLREPORTTABLE", intRow).Value, 1)

                'If idCRT = 29 Then 'this is just testing
                '    boolViewCBX = True
                'End If

                If boolAccess Then
                Else
                    boolViewCBX = True
                End If
                boolDo = True

            End If
        End If

        Try
            If boolDo Then
                Call SetTablePropertiesBool(idTR, idCRT)
            End If
        Catch ex As Exception

        End Try

        Call EnableCBXStudy(boolViewCBX)

        If boolFormLoad Then
            Exit Sub
        End If

        If boolCont Then
        Else
            Exit Sub
        End If

        Dim int1 As Short

        If Me.dgvAnalytes.RowCount = 0 Then
            strAnalFromTable = ""

        ElseIf Me.dgvAnalytes.CurrentRow Is Nothing Then
            int1 = 0
            strAnalFromTable = Me.dgvAnalytes("AnalyteDescription", int1).Value
        Else
            int1 = Me.dgvAnalytes.CurrentRow.Index
            strAnalFromTable = Me.dgvAnalytes("AnalyteDescription", int1).Value
        End If

        boolFromdgvTable = True
        'strAnalFromTable = Me.dgvAnalytes("AnalyteDescription", Me.dgvAnalytes.CurrentRow.Index).Value

        Cursor.Current = Cursors.WaitCursor

        'Me.lblWait.Visible = True
        Me.lblWait.BringToFront()

        Me.lblWait.Refresh()

        Call FilterHelper1()

        'record analyte

        If IsNothing(Me.dgvAnalytes.CurrentRow) Then
            str1 = ""
        Else
            str1 = Me.dgvAnalytes("AnalyteDescription", Me.dgvAnalytes.CurrentRow.Index).Value
        End If

        Call FilldgvAnalytes()

        If IsNothing(Me.dgvAnalytes.CurrentRow) Then
            '20190214 LEE:
            'must clear dgvAssignedSamples and dgvAnalyticalRuns
            Cursor.Current = Cursors.WaitCursor
            Call FillAnalyticalRuns("")
            Cursor.Current = Cursors.WaitCursor
            Call FillAssignedSamples()
            GoTo end1
        End If

        str2 = Me.dgvAnalytes("AnalyteDescription", Me.dgvAnalytes.CurrentRow.Index).Value
        Dim boolUT As Boolean = False
        If StrComp(str1, str2, CompareMethod.Text) = 0 Then
        Else
            boolUT = True
        End If

        '****

        ''only call this if needed
        If gboolFiltersCleared Or boolFormLoad1 Or boolUT Then
            Call FillAnalyticalRuns("")
        End If


        Me.Refresh()
        Cursor.Current = Cursors.WaitCursor

        'causes AnalyticalRunFlicker

        Call FillAssignedSamples()

        Cursor.Current = Cursors.WaitCursor
        'Me.Refresh()

        'Call AdjustAssignedSamplesWidth()

        Call ASNum()

        Call NomConcFill(False)

        Call AssessSampleAssignmentAnalyte()
        '****

        'Call FillAssignedSamples()
        Cursor.Current = Cursors.WaitCursor

        Call ShowColumns()


        '20181206 LEE:
        'Need to call this to update runid
        Try
            Call InitializeFilterRunID(Me.dgvAnalytes.CurrentRow.Index)
        Catch ex As Exception

        End Try

        '
        'Call ASNum()

        'scroll dgvAssignedSamples to the right
        'Me.dgvAssignedSamples.HorizontalScrollingOffset

        Dim var1

        var1 = Me.dgvAssignedSamples.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
        Me.dgvAssignedSamples.HorizontalScrollingOffset = var1

        Me.lblWait.Visible = False
        Me.lblWait.Refresh()

        boolFromdgvTable = False
        strAnalFromTable = ""

        If dgv.Rows.Count = 0 Then
        Else
            If dgv.CurrentRow Is Nothing Then
            Else
                Call IncSampleVis(idCRT)
            End If
        End If

end1:

        Cursor.Current = Cursors.Default

        'Call PlaceControls(boolViewOnly)'called earlier

    End Sub


    Private Sub dgvTables_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTables.MouseEnter
        'Me.dgvTables.Focus()

    End Sub

    Private Sub dgvTables_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvTables.CellContentClick

    End Sub

    Private Sub dgvTables_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvTables.CellValidating

        'If booldgvReportTableCancel And Me.cmdEdit.Enabled = False Then
        '    e.Cancel = True

        '    'now save contents of dgvTables
        '    Dim dgv As DataGridView = Me.dgvTables

        '    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)

        '    'move focus out
        '    Me.dgvAssignedSamples.Focus()

        '    booldgvReportTableCancel = False
        'End If

    End Sub

    Sub ClearFilters(boolClearAccStatus As Boolean)

        gboolFiltersCleared = False

        Dim boolFLT As Boolean = boolFormLoad
        boolFormLoad = True
        Dim str1 As String
        Dim int1 As Short
        str1 = Me.txtFilterSamples.Text
        If Len(str1) = 0 Then
        Else
            Me.txtFilterSamples.Text = ""
            gboolFiltersCleared = True
        End If

        int1 = Me.cbxFilterDilFactor.SelectedIndex
        If int1 > 0 Then
            Me.cbxFilterDilFactor.SelectedIndex = 0
            gboolFiltersCleared = True
        End If

        int1 = Me.cbxFilterRunID.SelectedIndex
        If int1 > 0 Then
            Me.cbxFilterRunID.SelectedIndex = 0
            gboolFiltersCleared = True
        End If

        int1 = Me.cbxFilterSampleType.SelectedIndex
        If int1 > 0 Then
            Me.cbxFilterSampleType.SelectedIndex = 0
            gboolFiltersCleared = True
        End If

        If boolClearAccStatus Then
            int1 = Me.cbxAccStatus.SelectedIndex
            If int1 > 0 Then
                Me.cbxAccStatus.SelectedIndex = 1
                gboolFiltersCleared = True
            End If
        End If

        boolFormLoad = boolFLT

    End Sub

    Sub dgvTables_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTables.SelectionChanged

        '20151221 LEE:
        'clear some filters when changing tables
        If boolFormLoad Then
        Else
            Call ClearFilters(False)
        End If

        If Me.cmdEdit.Enabled Then

            Try
                Me.txtdgvReportTableCurrentRow.Text = dgvTables.CurrentRow.Index
            Catch ex As Exception

            End Try

            Call ChangedgvTables()


        Else

            If booldgvReportTableCancel Then

            Else

                If (Not (assignSamplesVerify(False))) Then

                    Dim intPRow As Short 'previous row
                    Dim intRow As Short 'current row
                    Dim dgv As DataGridView = Me.dgvTables

                    Try
                        intPRow = NZ(Me.txtdgvReportTablePreviousRow.Text, 0)
                        intRow = NZ(dgv.CurrentRow.Index, 0)
                    Catch ex As Exception

                    End Try

                    booldgvReportTableCancel = True

                    dgv.Rows(intPRow).Selected = True

                    dgv.CurrentCell = dgv.Rows(intPRow).Cells("CHARHEADINGTEXT")

                    booldgvReportTableCancel = False

                    GoTo end1

                Else

                    Try
                        Me.txtdgvReportTableCurrentRow.Text = dgvTables.CurrentRow.Index
                    Catch ex As Exception

                    End Try

                    Call ChangedgvTables()


                End If

            End If

        End If


end1:

        Me.lblWait.Visible = False

    End Sub


    Private Sub dgvTables_RowLeave(sender As Object, e As DataGridViewCellEventArgs) Handles dgvTables.RowLeave

        'dgvTables_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        'If Me.cmdEdit.Enabled Then
        'Else

        '    If booldgvReportTableCancel Then
        '    Else
        '        If (Not (assignSamplesVerify())) Then

        '            'booldgvReportTableCancel = True


        '            'Try
        '            '    dgv.Rows(intPRow).Selected = True
        '            '    dgv.CurrentCell = dgv.Rows(intPRow).Cells("CHARHEADINGTEXT")
        '            'Catch ex As Exception

        '            'End Try


        '            'booldgvReportTableCancel = False

        '            'place focus on something else
        '            'Me.dgvAssignedSamples.Focus()



        '            'booldgvReportTableCancel = False

        '            GoTo end1

        '            'DataGridView1.CurrentCell = DataGridView1.Rows(1).Cells(0)

        '        End If
        '    End If

        'End If

        Try
            Me.txtdgvReportTablePreviousRow.Text = dgvTables.CurrentRow.Index
        Catch ex As Exception

        End Try

end1:


    End Sub


    Sub EnableCBXStudy(ByVal bool As Boolean)

        Me.panCbxStudy.Enabled = bool

        Dim var1
        var1 = Me.txtStudyID.Text
        If var1 = id_tblStudies Then 'ignore
        Else
            Call ReturnStudyToOriginal()
        End If


    End Sub

    Private Sub dgvAnalytes_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAnalytes.CellContentClick

    End Sub

    Private Sub dgvAnalytes_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAnalytes.MouseEnter
        'Me.dgvAnalytes.Focus()

    End Sub

    Private Sub dgvAnalytes_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAnalytes.SelectionChanged

        If boolDontChange Then 'triggered during autoassignsamples
        Else
            Try
                Call dgvAnalyteSelectionChange("")
            Catch ex As Exception

            End Try

        End If


    End Sub

    Sub dgvAnalyteSelectionChange(strAccStatus As String)

        'If boolFormLoad Then
        '    Exit Sub
        'End If

        'If boolCont Then
        'Else
        '    Exit Sub
        'End If

        Cursor.Current = Cursors.WaitCursor

        Dim intRow As Short
        Dim var1

        Call ChooseAnalyte()

        Try
            intRow = Me.dgvAnalytes.CurrentRow.Index
            gintGroup = NZ(Me.dgvAnalytes("INTGROUP", intRow).Value, -1) 'null if IntStd
        Catch ex As Exception
            gintGroup = -1
        End Try

        If Me.cmdEdit.Enabled Then

            If boolFormLoad Then
                Exit Sub
            End If

            If boolCont Then
            Else
                Exit Sub
            End If

            Try
                Me.txtdgvAnalyteCurrentRow.Text = dgvAnalytes.CurrentRow.Index
            Catch ex As Exception

            End Try

            If boolAutoAssign Then
            Else
                Me.lblWait.BringToFront()
                Me.lblWait.Visible = True
                Me.lblWait.Refresh()
            End If
           

            Call FillAnalyticalRuns(strAccStatus)
            Me.Refresh()
            Cursor.Current = Cursors.WaitCursor

            Try
                Call InitializeFilterRunID(Me.dgvAnalytes.CurrentRow.Index)
            Catch ex As Exception
                var1 = ex.Message
            End Try


            Call FillAssignedSamples()
            Cursor.Current = Cursors.WaitCursor
            'Me.Refresh()

            Call ASNum()

            Call NomConcFill(False)

            Call ShowColumns()

            Call sortAssignedSamples()

        Else

            If booldgvReportTableCancel Then

            Else
                If (dgvAnalytes.RowCount > 0) Then  'NDL:  7-Dec-2015 This is a temporary fix to prevent crashes.  Issue 
                    If (Not (assignSamplesVerify(False))) And boolAutoAssign = False Then

                        Dim intPRow As Short 'previous row
                        'Dim intRow As Short 'current row
                        Dim dgv As DataGridView = Me.dgvAnalytes

                        Try
                            intPRow = NZ(Me.txtdgvAnalytePreviousRow.Text, 0)
                            intRow = NZ(dgv.CurrentRow.Index, 0)
                        Catch ex As Exception

                        End Try

                        booldgvReportTableCancel = True

                        dgv.CurrentCell = dgv.Rows(intPRow).Cells("AnalyteDescription")
                        dgv.Rows(intPRow).Selected = True

                        booldgvReportTableCancel = False

                        GoTo end1

                    Else

                        If boolFormLoad Then
                            Exit Sub
                        End If

                        If boolCont Then
                        Else
                            Exit Sub
                        End If

                        Try
                            Me.txtdgvAnalyteCurrentRow.Text = dgvAnalytes.CurrentRow.Index
                        Catch ex As Exception

                        End Try

                        If boolAutoAssign Then
                        Else
                            Me.lblWait.BringToFront()
                            Me.lblWait.Visible = True
                            Me.lblWait.Refresh()
                        End If
                        
                        Call FillAnalyticalRuns(strAccStatus)
                        Me.Refresh()
                        Cursor.Current = Cursors.WaitCursor

                        Call InitializeFilterRunID(Me.dgvAnalytes.CurrentRow.Index)

                        Call FillAssignedSamples()
                        Cursor.Current = Cursors.WaitCursor
                        'Me.Refresh()

                        Call ASNum()

                        Call NomConcFill(False)

                        Call ShowColumns()

                        Call sortAssignedSamples()

                    End If

                End If

            End If
        End If


end1:


        Try
            Call CountSamples()

            Call AssessSampleAssignment()
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Me.lblWait.Visible = False
        Me.lblWait.Refresh()

        Cursor.Current = Cursors.Default

        'set focus to dgvAnalyticalRuns
        Me.dgvAnalyticalRuns.Focus()

    End Sub

    Sub ChooseAnalyte()

        '20180129 LEE: Added ChooseAnalytes functionality to allow user to choose which Analyte/IS pair wish to be viewed

        Dim intRow As Short
        Dim var1, var2, var3, var4
        Dim intG As Short
        Dim boolIsIS As Boolean = False
        Dim str1 As String
        Dim dgv As DataGridView = Me.dgvAnalytes
        Dim strIS As String
        Dim tbl1 As DataTable = tblAnalytesHome
        Dim strF As String
        Dim strS As String
        Dim Count1 As Int16
        Dim Count2 As Int16

        var1 = dgv.Rows.Count

        Try
            intRow = dgv.CurrentRow.Index
            intG = NZ(dgv("INTGROUP", intRow).Value, -1) 'null if IntStd
        Catch ex As Exception
            Me.panChooseAnalyte.Visible = False
            GoTo end1
        End Try

        var1 = NZ(dgv("IsIntStd", intRow).Value, "No")
        var2 = NZ(dgv("ANALYTEDESCRIPTION", intRow).Value, "No")

        If StrComp(var1, "Yes", CompareMethod.Text) = 0 Then
            boolIsIS = True
        Else
            boolIsIS = False
        End If

        If boolIsIS Then
        Else
            Me.panChooseAnalyte.Visible = False
            GoTo end1
        End If

        'find analytes with this IS
        strIS = dgv("ANALYTEDESCRIPTION", intRow).Value

        strF = "UseIntStd = '" & strIS & "'"
        strF = "UseIntStd = 'Yes' AND INTSTD = '" & CleanText(strIS) & "'"
        strS = "ANALYTEDESCRIPTION ASC"
        Dim rows1() As DataRow = tbl1.Select(strF, strS)

        Dim cbx As ComboBox = Me.cbxChooseAnalyte

        cbx.Items.Clear()

        If rows1.Length = 0 Then
            GoTo end1
        End If

        For Count1 = 0 To rows1.Length - 1
            var1 = rows1(Count1).Item("ANALYTEDESCRIPTION")
            cbx.Items.Add(var1)
        Next

        cbx.SelectedIndex = 0

        Me.panChooseAnalyte.Visible = True


end1:

    End Sub

    Sub ReturnStudyToOriginal()

        '20180821 LEE:
        'Simply put dgv back to tblAnalysisResults

        Call ReturnClick()

        Exit Sub

        Dim Count1 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim tbl As System.Data.DataTable
        Dim int2 As Int64
        Dim var1, var2, var3

        tbl = Me.cbxStudy.DataSource

        str1 = wStudyID 'wWStudyName ' Me.txtStudy.Text
        'int1 = Me.cbxStudy.Items.Count
        int1 = tbl.Rows.Count

        For Count1 = 0 To int1 - 1
            int2 = tbl.Rows(Count1).Item("STUDYID")
            If int2 = wStudyID Then
                Dim boolFL As Boolean = boolFormLoad
                boolFormLoad = True
                Me.cbxStudy.SelectedIndex = Count1
                boolFormLoad = boolFL
                var1 = var1
                Exit For
            End If
        Next

        '*****

        Dim strF As String
        Dim rows() As DataRow
        Dim intRow As Int64

        'var1 = wWStudyName 'Me.txtStudy.Text ' Me.cbxStudy.Text

        intRow = Me.cbxStudy.SelectedIndex
        tbl = Me.tblStudiesA

        Me.txtStudyID.Text = CStr(id_tblStudies)

        Dim cn As New ADODB.Connection
        cn.Open(constrCur)

        Call SetAnalysisResultsTable(wStudyID, cn)


        Try
            Call ConcLevels(wStudyID, cn)
        Catch ex As Exception
            var1 = var1
        End Try

        If cn.State = ADODB.ObjectStateEnum.adStateOpen Then
            cn.Close()
        End If

        Try
            cn = Nothing
        Catch ex As Exception

        End Try


        Call FillAnalyticalRuns("")



    End Sub

    Private Sub cmdReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReturn.Click

        Call ReturnClick()

    End Sub

    Sub ReturnClick()

        Dim strM As String
        Dim CountM As Short
        Dim str1 As String
        Dim str2 As String
        Dim Count1 As Int16
        Dim id As Int32
        Dim var1

        If boolArchiveSource Then

            strM = "This feature not available when the Watson data source is an archived .mdb"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        'Call ReturnStudyToOriginal()

        '20180821 LEE:
        'return dgvAnalyticalRuns to tblAnalyterResults

        'Call FilterForAnalyte(tblAnalysisResultsHome)

        boolOriginal = True

        If boolAutoAssign Then
        Else
            Me.lblWait.Visible = True
            Me.lblWait.Refresh()
        End If

        Me.txtStudyID.Text = id_tblStudies

        Call FilterForAnalyte(tblAnalysisResults)
        Try
            Call FillAnalyticalRuns("")
        Catch ex As Exception
            var1 = var1
        End Try

        Call CountSamples()

        Try
            Call InitializeFilterRunID(Me.dgvAnalytes.CurrentRow.Index)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        'need to re-establish Conc Levels datatables
        Dim cn As New ADODB.Connection
        cn.Open(constrCur)

        Try
            Call ConcLevels(wStudyID, cn)
        Catch ex As Exception
            var1 = var1
        End Try

        If cn.State = ADODB.ObjectStateEnum.adStateOpen Then
            cn.Close()
        End If


        'return cbxStudy to original index
        For Count1 = 0 To Me.tblStudiesA.Rows.Count - 1
            id = Me.tblStudiesA.Rows(Count1).Item("STUDYID")
            If id = wStudyID Then
                Dim boolFL As Boolean = boolFormLoad
                boolFormLoad = True
                Me.cbxStudy.SelectedIndex = Count1
                boolFormLoad = boolFL
                Exit For
            End If
        Next

        Me.lblWait.Visible = False

    End Sub

    Public Sub AutoComplete(ByRef cb As ComboBox, ByVal e As System.Windows.Forms.KeyPressEventArgs, Optional ByVal blnLimitToList As Boolean = False)

        Dim strFindStr As String

        If e.KeyChar = Chr(8) Then 'Check For Backspace 

            If cb.SelectionStart <= 1 Then

                cb.Text = ""

                Exit Sub

            End If

            If cb.SelectionLength = 0 Then

                strFindStr = cb.Text.Substring(0, cb.Text.Length - 1)

            Else

                strFindStr = cb.Text.Substring(0, cb.SelectionStart - 1)

            End If

        Else

            If cb.SelectionLength = 0 Then

                strFindStr = cb.Text & e.KeyChar

            Else

                strFindStr = cb.Text.Substring(0, cb.SelectionStart) & e.KeyChar

            End If

        End If



        Dim intIdx As Integer = -1

        'Search the string in the ComboBox List.

        intIdx = cb.FindString(strFindStr)

        If intIdx <> -1 Then ' String found in the List.

            cb.SelectedText = ""

            cb.SelectedIndex = intIdx

            cb.SelectionStart = strFindStr.Length

            cb.SelectionLength = cb.Text.Length

            e.Handled = True

        Else

            If blnLimitToList = True Then

                e.Handled = True

            Else

                e.Handled = False

            End If

        End If

    End Sub

    Function cbxStudyValidating(cn As ADODB.Connection) As Boolean

        Cursor.Current = Cursors.WaitCursor

        cbxStudyValidating = True

        Dim tbl As System.Data.DataTable
        Dim intRow As Int64
        Dim var1, var2, var3, var4
        Dim intRows As Short
        Dim strF As String
        Dim rows() As DataRow
        Dim strM As String

        tbl = Me.tblStudiesA
        intRow = Me.cbxStudy.SelectedIndex

        'var3 = tbl.Rows(intRow).Item("STUDYID")
        var3 = Me.cbxStudy.Text

        If var3 = wWStudyName Then
            cbxStudyValidating = True
            Exit Function
        End If

        'var4 = tbl.Rows(intRow).Item("STUDYNAME")

        'find var3 in tblStudies
        'strF = "INT_WATSONSTUDYID = " & var3
        strF = "CHARWATSONSTUDYNAME = '" & var3 & "'"
        'strF = ""
        rows = tblStudies.Select(strF)
        intRows = rows.Length
        If intRows = 0 Then
            strM = "Watson Study '" & var3 & "' has not been configured in StudyDoc and cannot be chosen here."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            cbxStudyValidating = False
            Exit Function
        End If

        'now check to ensure analytes are in study
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim rs As New ADODB.Recordset
        'Dim cn As New ADODB.Connection
        Dim fld As ADODB.Field
        Dim drow As DataRow
        Dim row1() As DataRow
        Dim int1 As Int64
        Dim id As Int64

        var2 = rows(0).Item("INT_WATSONSTUDYID")
        id = var2

        'cn.Open(constrCur)

        boolANSI = True

        'If boolANSI Then
        '    str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, GLOBALANALYTES.ANALYTEDESCRIPTION, GLOBALANALYTES.PROJECTID, ASSAY.MASTERASSAYID "
        '    str2 = "FROM ANARUNANALYTERESULTS INNER JOIN (ASSAY INNER JOIN ((ASSAYANALYTES INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID) INNER JOIN STUDY ON ASSAYANALYTES.STUDYID = STUDY.STUDYID) ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID)) ON ANARUNANALYTERESULTS.RUNID = ASSAY.RUNID "
        '    str3 = "WHERE(((ASSAYANALYTES.STUDYID) = " & id & ") And ((GLOBALANALYTES.ACTIVE) = -1)) "
        '    str4 = "ORDER BY GLOBALANALYTES.ANALYTEDESCRIPTION;"
        'Else
        '    str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, GLOBALANALYTES.ANALYTEDESCRIPTION, GLOBALANALYTES.PROJECTID, ASSAY.MASTERASSAYID "
        '    str2 = "FROM ANARUNANALYTERESULTS, ASSAY, ASSAYANALYTES, GLOBALANALYTES, STUDY "
        '    str2 = str2 & "WHERE (((ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID) AND ASSAYANALYTES.STUDYID = STUDY.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID)) AND ANARUNANALYTERESULTS.RUNID = ASSAY.RUNID "
        '    str3 = "AND (((ASSAYANALYTES.STUDYID) = " & id & ") And ((GLOBALANALYTES.ACTIVE) = -1)) "
        '    str4 = "ORDER BY GLOBALANALYTES.ANALYTEDESCRIPTION;"
        'End If

        If boolAccess Then
            str1 = "SELECT DISTINCT ASSAYANALYTES.STUDYID, ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, GLOBALANALYTES.ANALYTEDESCRIPTION, GLOBALANALYTES.PROJECTID, ASSAY.MASTERASSAYID "
            str2 = "FROM ANARUNANALYTERESULTS INNER JOIN (ASSAY INNER JOIN ((ASSAYANALYTES INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID) INNER JOIN STUDY ON ASSAYANALYTES.STUDYID = STUDY.STUDYID) ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID)) ON ANARUNANALYTERESULTS.RUNID = ASSAY.RUNID "
            str3 = "WHERE(((ASSAYANALYTES.STUDYID) = " & id & ") And ((GLOBALANALYTES.ACTIVE) = -1)) "
            str4 = "ORDER BY GLOBALANALYTES.ANALYTEDESCRIPTION;"
        Else
            str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.STUDYID, " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".GLOBALANALYTES.PROJECTID, " & strSchema & ".ASSAY.MASTERASSAYID "
            str2 = "FROM " & strSchema & ".ANARUNANALYTERESULTS INNER JOIN (" & strSchema & ".ASSAY INNER JOIN ((" & strSchema & ".ASSAYANALYTES INNER JOIN " & strSchema & ".GLOBALANALYTES ON " & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID) INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".STUDY.STUDYID) ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID)) ON " & strSchema & ".ANARUNANALYTERESULTS.RUNID = " & strSchema & ".ASSAY.RUNID "
            str3 = "WHERE(((" & strSchema & ".ASSAYANALYTES.STUDYID) = " & id & ") And ((" & strSchema & ".GLOBALANALYTES.ACTIVE) = -1)) "
            str4 = "ORDER BY " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION;"
        End If

        strSQL = str1 & str2 & str3 & str4
        ''''''''''''console.writeline(strSQL)
        rs.CursorLocation = CursorLocationEnum.adUseClient
        rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        rs.ActiveConnection = Nothing

        int1 = rs.RecordCount 'debugging

        Dim Count1 As Short
        '0=TRUE/FALSE,1=analytedescription, 2=analyteindex, 3=masterassayid, 4=ANALYTEID, 5=wStudyID, 6=StudyName
        Dim boolAnalOKAll As Boolean = False
        Dim dgv As DataGridView

        dgv = Me.dgvAnalytes

        Dim strF1 As String
        Dim rowsA() As DataRow

        For Count1 = 0 To dgv.RowCount - 1
            var1 = Me.dgvAnalytes("ANALYTEDESCRIPTION", Count1).Value
            strF = "ANALYTEDESCRIPTION = '" & CleanText(CStr(var1)) & "'"
            rs.Filter = ""
            rs.Filter = strF

            ctAnalOK = ctAnalOK + 1
            If rs.EOF And rs.BOF Then
                boolAnalOK(0, ctAnalOK) = False
            Else
                boolAnalOK(0, ctAnalOK) = True
                boolAnalOKAll = True
                'modify tblanalytes data
                rowsA = tblAnalytes.Select(strF)
                rowsA(0).BeginEdit()
                rowsA(0).Item("ANALYTEINDEX") = rs.Fields("ANALYTEINDEX").Value
                rowsA(0).Item("MASTERASSAYID") = rs.Fields("MASTERASSAYID").Value
                rowsA(0).Item("ANALYTEID") = rs.Fields("ANALYTEID").Value
                rowsA(0).EndEdit()

            End If
            boolAnalOK(1, ctAnalOK) = dgv("ANALYTEDESCRIPTION", Count1).Value
            boolAnalOK(2, ctAnalOK) = dgv("ANALYTEINDEX", Count1).Value
            boolAnalOK(3, ctAnalOK) = dgv("MASTERASSAYID", Count1).Value
            boolAnalOK(4, ctAnalOK) = dgv("ANALYTEID", Count1).Value
            boolAnalOK(5, ctAnalOK) = rows(0).Item("ID_TBLSTUDIES")
            boolAnalOK(6, ctAnalOK) = var4 'Me.cbxStudy.SelectedItem
        Next

        rs.Close()
        rs = Nothing

        If boolAnalOKAll Then 'continue
        Else
            strM = "Watson Study '" & var3 & "' does not have the same analytes as the parent study '" & wWStudyName & "'."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            cbxStudyValidating = False
            Cursor.Current = Cursors.Default
            Exit Function
        End If

        Cursor.Current = Cursors.Default

    End Function

    Sub ReDoAnalytes()

    End Sub

    Private Sub cbxStudy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxStudy.SelectedIndexChanged

        If boolHold Then
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        booldgvReportTableCancel = True

        Call ChangeStudy()

        booldgvReportTableCancel = False

        Cursor.Current = Cursors.Default

    End Sub


    Sub ChangeStudy()

        'If boolFormLoad Or boolCancelButton Then
        If boolFormLoad Then
            Exit Sub
        End If

        Dim strM As String

        If boolArchiveSource Then

            'str1 = wWStudyName 'Me.txtStudy.Text
            'For CountM = 0 To Me.cbxStudy.Items.Count - 1
            '    str2 = Me.cbxStudy.Items(CountM)
            '    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
            '        Exit For
            '    End If
            'Next
            'If CountM = Me.cbxStudy.Items.Count Then
            '    CountM = 0
            'End If
            'boolFormLoad = True
            'Me.cbxStudy.SelectedIndex = CountM
            'boolFormLoad = False

            strM = "This feature not available when the Watson data source is an archived .mdb"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")

            Exit Sub

        End If

        boolFromChangeStudy = True

        Dim Count1 As Short
        Dim id As Int64


        Dim tbl As System.Data.DataTable
        Dim intRow As Int64
        Dim var1, var2, var3, var4
        Dim intRows As Short
        Dim strF As String
        Dim rows() As DataRow

        'legend:
        'Me.tblStudiesA = frmH.dgvwStudy.DataSource
        tbl = Me.tblStudiesA

        '20180825 LEE:
        'must change runid to [NONE]
        Dim boolFL As Boolean = boolFormLoad
        boolFormLoad = True
        Me.cbxFilterRunID.SelectedIndex = 0
        boolFormLoad = boolFL

        'intRow = Me.cbxStudy.SelectedIndex

        'var3 = tbl.Rows(intRow).Item("STUDYID")

        var3 = Me.cbxStudy.Text
        nStudyName = var3

        If var3 = wWStudyName Then

            Call ReturnStudyToOriginal()

            ''20180821 LEE:
            ''No
            ''Exit Sub

            ''20180821 LEE:
            ''return dgvAnalyticalRuns to tblAnalyterResults

            'Me.lblWait.Visible = True
            'Me.lblWait.Refresh()

            'Me.txtStudyID.Text = id_tblStudies

            ''Call FilterForAnalyte(tblAnalysisResultsHome)
            'Call FilterForAnalyte(tblAnalysisResults)

            'Try
            '    Call FillAnalyticalRuns("")
            'Catch ex As Exception
            '    var1 = var1
            'End Try


            'Call CountSamples()

            'Try
            '    Call InitializeFilterRunID(Me.dgvAnalytes.CurrentRow.Index)
            'Catch ex As Exception
            '    var1 = ex.Message
            '    var1 = var1
            'End Try


            ''return cbxStudy to original index
            'For Count1 = 0 To Me.tblStudiesA.Rows.Count - 1
            '    id = Me.tblStudiesA.Rows(Count1).Item("STUDYID")
            '    If id = wStudyID Then
            '        boolFormLoad = True
            '        Me.cbxStudy.SelectedIndex = Count1
            '        boolFormLoad = False
            '        Exit For
            '    End If
            'Next

            ''Call ReturnStudyToOriginal()

            'Me.lblWait.Visible = False

            Exit Sub
        End If

        Dim cn As New ADODB.Connection

        cn.Open(constrCur)

        'first validate
        If cbxStudyValidating(cn) Then 'continue
        Else
            '20190130 LEE:
            'Hmmm. If studies are filtered, tblStudiesA has more rows than cbxStudyItems because cbxStudyItems is filtered
            'Try looking for id in cbxstudy instead of index

            For Count1 = 0 To Me.tblStudiesA.Rows.Count - 1
                id = Me.tblStudiesA.Rows(Count1).Item("STUDYID")
                If id = wStudyID Then
                    boolFormLoad = True
                    Me.cbxStudy.SelectedIndex = Count1
                    boolFormLoad = False
                    Exit For
                End If
            Next

            If cn.State = ADODB.ObjectStateEnum.adStateOpen Then
                cn.Close()
            End If


            Try
                cn = Nothing
            Catch ex As Exception

            End Try

            Exit Sub

        End If

        Dim CountM As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String

        If boolFormLoad Or boolAutoAssign Then
        Else
            Me.lblWait.BringToFront()
            Me.lblWait.Visible = True
            Me.lblWait.Refresh()
            Me.Refresh()
        End If

        Dim intRowA As Short
        Dim wID As Int64


        Cursor.Current = Cursors.WaitCursor

        'var1 = Me.cbxStudy.Text
        'tbl = tblStudies
        'strF = "charWatsonStudyName = '" & var1 & "'"
        'rows = tbl.Select(strF)

        'var3 = tbl.Rows(intRow).Item("STUDYID")
        'var4 = tbl.Rows(intRow).Item("STUDYNAME")

        'find var3 in tblStudies
        'strF = "INT_WATSONSTUDYID = " & var3
        strF = "CHARWATSONSTUDYNAME = '" & nStudyName & "'"
        rows = tblStudies.Select(strF)
        var2 = rows(0).Item("ID_TBLSTUDIES")
        nSDStudyID = var2

        Me.txtStudyID.Text = CStr(var2)

        var1 = Me.txtStudyID.Text
        wID = rows(0).Item("INT_WATSONSTUDYID") 'GetWStudyID(var1)
        nWStudyID = wID


        If wID = CLng(wStudyID) Then 'booloriginal is global
            boolOriginal = True
        Else
            boolOriginal = False
        End If


        Try
            intRowA = Me.dgvAnalytes.CurrentRow.Index
        Catch ex As Exception
            intRowA = 0
        End Try

        Dim boolDC As Boolean = boolDontChange
        'boolDontChange = True


        strM = "" 'for debugging
        Try
            strM = "1"
            Call tblAnalytesConfigure(True)
            Try
                'Me.dgvAnalytes.Rows(intRowA).Selected = True
            Catch ex As Exception

            End Try
            strM = "2"
            Call InitializedgvAnalytes()
            Try
                'Me.dgvAnalytes.Rows(intRowA).Selected = True
            Catch ex As Exception

            End Try
            strM = "3"
            Call FilldgvAnalytes()
            Try
                ' Me.dgvAnalytes.Rows(intRowA).Selected = True
            Catch ex As Exception

            End Try
            strM = "4"
            Call SetAnalysisResultsTable(wID, cn)
            Try
                'Me.dgvAnalytes.Rows(intRowA).Selected = True
            Catch ex As Exception

            End Try

            Try
                Call ConcLevels(wID, cn)
            Catch ex As Exception
                var1 = var1
            End Try




            ''20160419 LEE:
            ''need to add data to tblAnalyteConcLevelsForAssay
            'If boolAccess Then
            '    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ASSAYANALYTEKNOWN.LEVELNUMBER, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAYANALYTEKNOWN.STUDYID "
            '    str2 = "FROM ((((ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ANALYTICALRUNSAMPLE.ASSAYLEVEL = ASSAYANALYTEKNOWN.LEVELNUMBER)) INNER JOIN GLOBALANALYTES ON (ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID)"
            '    str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wID & ")) "
            '    str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"
            'Else
            '    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
            '    str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) INNER JOIN " & strSchema & ".GLOBALANALYTES ON (" & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID)"
            '    str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wID & ")) "
            '    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

            'End If

            'strSQL = str1 & str2 & str3 & str4
            ' ''console.writeline(strSQL)
            ' ''''''''''''''''''''''''''''''''''''''''console.writeline(strSQL)
            'Dim rs As New ADODB.Recordset
            'If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            '    rs.Close()
            'End If
            'rs.CursorLocation = CursorLocationEnum.adUseClient
            'Try
            '    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'Catch ex As Exception
            '    var1 = ex.Message
            '    var1 = var1
            'End Try

            'rs.ActiveConnection = Nothing

            ''rsspeciesmatrix
            'tblAnalyteConcLevelsForAssay.Clear()
            'tblAnalyteConcLevelsForAssay.AcceptChanges()
            'tblAnalyteConcLevelsForAssay.BeginLoadData()
            'daDoPr.Fill(tblAnalyteConcLevelsForAssay, rs)
            'tblAnalyteConcLevelsForAssay.EndLoadData()


            ''20161227 LEE: must do tblConcLevelsForAssayIDs also

            'If boolAccess Then

            '    '20160816 LEE
            '    'this query was not returning custom sample types
            '    'ASSAYANALYTEKNOWN.KNOWNTYPE: limited sample types to QC, STANDARD, STABILITY
            '    'ANALYTICALRUNSAMPLE.RUNSAMPLEKIND: KNOWNTYPE + UNKNOWN + custom sample types
            '    'Solution: remove join between KNOWNTYPE and RUNSAMPLEKIND
            '    'add additional parameters to FindNomConc
            '    str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ASSAYANALYTEKNOWN.LEVELNUMBER, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.STUDYID "
            '    str2 = "FROM ((((ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ASSAYANALYTES ON (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ANALYTICALRUNSAMPLE.ASSAYLEVEL = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID "
            '    str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wID & ")) "
            '    str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"


            'Else

            '    '20160816 LEE
            '    'this query was not returning custom sample types
            '    'ASSAYANALYTEKNOWN.KNOWNTYPE: limited sample types to QC, STANDARD, STABILITY
            '    'ANALYTICALRUNSAMPLE.RUNSAMPLEKIND: KNOWNTYPE + UNKNOWN + custom sample types
            '    'Solution: remove join between KNOWNTYPE and RUNSAMPLEKIND
            '    'add additional parameters to FindNomConc
            '    str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
            '    str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN " & strSchema & ".GLOBALANALYTES ON " & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID "
            '    str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wID & ")) "
            '    str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

            'End If

            'strSQL = str1 & str2 & str3 & str4

            ''Console.WriteLine("tblConcLevelsForAssayIDs: " & strSQL)

            'If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            '    rs.Close()
            'End If
            'rs.CursorLocation = CursorLocationEnum.adUseClient
            'rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

            'tblConcLevelsForAssayIDs.Clear()
            'tblConcLevelsForAssayIDs.AcceptChanges()
            'tblConcLevelsForAssayIDs.BeginLoadData()
            'daDoPr.Fill(tblConcLevelsForAssayIDs, rs)
            'tblConcLevelsForAssayIDs.EndLoadData()

            ''20161227 LEE: must do tblASSAYREPS also

            'If boolAccess Then
            '    str1 = "SELECT ASSAYREPS.* "
            '    str2 = "FROM ASSAYREPS "
            '    str3 = "WHERE (((ASSAYREPS.STUDYID)=" & wID & "));"
            'Else
            '    str1 = "SELECT " & strSchema & ".ASSAYREPS.* "
            '    str2 = "FROM " & strSchema & ".ASSAYREPS "
            '    str3 = "WHERE (((" & strSchema & ".ASSAYREPS.STUDYID)=" & wID & "));"
            'End If

            'strSQL = str1 & str2 & str3
            ' ''console.writeline(strSQL)
            ' ''''''''''''''''''''''''''''''''''''''''console.writeline(strSQL)
            'If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            '    rs.Close()
            'End If
            'rs.CursorLocation = CursorLocationEnum.adUseClient
            'Try
            '    rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'Catch ex As Exception
            '    var1 = var1
            'End Try

            'rs.ActiveConnection = Nothing

            ''rsspeciesmatrix
            'tblASSAYREPS.Clear()
            'tblASSAYREPS.AcceptChanges()
            'tblASSAYREPS.BeginLoadData()
            'daDoPr.Fill(tblASSAYREPS, rs)
            'tblASSAYREPS.EndLoadData()

            'If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            '    rs.Close()
            'End If

            'Try
            '    rs = Nothing
            'Catch ex As Exception

            'End Try

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        Try
            Me.dgvAnalytes.Rows(intRowA).Selected = True
            '_2018DataGridView.CurrentCell = _2018DataGridView.Item(0, i)
            Me.dgvAnalytes.CurrentCell = Me.dgvAnalytes.Rows.Item(intRowA).Cells("AnalyteDescription")
        Catch ex As Exception

        End Try

        Try
            Call FillAnalyticalRuns("")
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


        Try
            Call InitializeFilterRunID(intRowA)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        Try
            Me.dgvAnalytes.Rows(intRowA).Selected = True
        Catch ex As Exception

        End Try

        'intRow = Me.dgvAnalytes.CurrentRow.Index
        gintGroup = NZ(Me.dgvAnalytes("INTGROUP", intRowA).Value, -1) 'null if IntStd

        boolDontChange = boolDC

        If boolFormLoad Then
        Else
            Me.lblWait.Visible = False
            Me.lblWait.Refresh()
            Me.Refresh()
        End If

        Cursor.Current = Cursors.Default


        Call CountSamples()

end1:

        If cn.State = ADODB.ObjectStateEnum.adStateOpen Then
            cn.Close()
        End If


        Try
            cn = Nothing
        Catch ex As Exception

        End Try

        boolFromChangeStudy = False

    End Sub

    Sub ConcLevels(wID As Int32, cn As ADODB.Connection)

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim var1, var2, var3, var4

        '20160419 LEE:
        'need to add data to tblAnalyteConcLevelsForAssay
        If boolAccess Then
            str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ASSAYANALYTEKNOWN.LEVELNUMBER, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAYANALYTEKNOWN.STUDYID "
            str2 = "FROM ((((ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ANALYTICALRUNSAMPLE.ASSAYLEVEL = ASSAYANALYTEKNOWN.LEVELNUMBER)) INNER JOIN GLOBALANALYTES ON (ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID)"
            str3 = "WHERE (((ASSAYANALYTEKNOWN.STUDYID)=" & wID & ")) "
            str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"
        Else
            str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
            str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID) AND (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID) AND (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER)) INNER JOIN " & strSchema & ".GLOBALANALYTES ON (" & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID)"
            str3 = "WHERE (((" & strSchema & ".ASSAYANALYTEKNOWN.STUDYID)=" & wID & ")) "
            str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

        End If

        strSQL = str1 & str2 & str3 & str4
        ''console.writeline(strSQL)
        ''''''''''''''''''''''''''''''''''''''''console.writeline(strSQL)
        Dim rs As New ADODB.Recordset
        If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            rs.Close()
        End If
        rs.CursorLocation = CursorLocationEnum.adUseClient
        Try
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        rs.ActiveConnection = Nothing

        'rsspeciesmatrix
        tblAnalyteConcLevelsForAssay.Clear()
        tblAnalyteConcLevelsForAssay.AcceptChanges()
        tblAnalyteConcLevelsForAssay.BeginLoadData()
        daDoPr.Fill(tblAnalyteConcLevelsForAssay, rs)
        tblAnalyteConcLevelsForAssay.EndLoadData()


        '20161227 LEE: must do tblConcLevelsForAssayIDs also

        If boolAccess Then

            '20160816 LEE
            'this query was not returning custom sample types
            'ASSAYANALYTEKNOWN.KNOWNTYPE: limited sample types to QC, STANDARD, STABILITY
            'ANALYTICALRUNSAMPLE.RUNSAMPLEKIND: KNOWNTYPE + UNKNOWN + custom sample types
            'Solution: remove join between KNOWNTYPE and RUNSAMPLEKIND
            'add additional parameters to FindNomConc
            str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ASSAYANALYTEKNOWN.LEVELNUMBER, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAYANALYTEKNOWN.KNOWNTYPE, ASSAYANALYTEKNOWN.STUDYID "
            str2 = "FROM ((((ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID) AND (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID)) INNER JOIN ASSAY ON (ANALYTICALRUN.STUDYID = ASSAY.STUDYID) AND (ANALYTICALRUN.RUNID = ASSAY.RUNID)) INNER JOIN ASSAYANALYTES ON (ASSAY.STUDYID = ASSAYANALYTES.STUDYID) AND (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ANALYTICALRUNSAMPLE.ASSAYLEVEL = ASSAYANALYTEKNOWN.LEVELNUMBER) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN GLOBALANALYTES ON ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID "
            str3 = "WHERE(((ANALYTICALRUN.STUDYID) = " & wID & ")) "
            str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"


        Else

            '20160816 LEE
            'this query was not returning custom sample types
            'ASSAYANALYTEKNOWN.KNOWNTYPE: limited sample types to QC, STANDARD, STABILITY
            'ANALYTICALRUNSAMPLE.RUNSAMPLEKIND: KNOWNTYPE + UNKNOWN + custom sample types
            'Solution: remove join between KNOWNTYPE and RUNSAMPLEKIND
            'add additional parameters to FindNomConc
            str1 = "SELECT DISTINCT " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL, " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER, " & strSchema & ".GLOBALANALYTES.ANALYTEDESCRIPTION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNID, " & strSchema & ".ANALYTICALRUN.ASSAYID, " & strSchema & ".ASSAYANALYTEKNOWN.CONCENTRATION, " & strSchema & ".ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, " & strSchema & ".ASSAYANALYTEKNOWN.KNOWNTYPE, " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID "
            str2 = "FROM ((((" & strSchema & ".ANALYTICALRUNSAMPLE INNER JOIN " & strSchema & ".ANALYTICALRUN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.STUDYID = " & strSchema & ".ANALYTICALRUN.STUDYID) AND (" & strSchema & ".ANALYTICALRUNSAMPLE.RUNID = " & strSchema & ".ANALYTICALRUN.RUNID)) INNER JOIN " & strSchema & ".ASSAY ON (" & strSchema & ".ANALYTICALRUN.STUDYID = " & strSchema & ".ASSAY.STUDYID) AND (" & strSchema & ".ANALYTICALRUN.RUNID = " & strSchema & ".ASSAY.RUNID)) INNER JOIN " & strSchema & ".ASSAYANALYTES ON (" & strSchema & ".ASSAY.STUDYID = " & strSchema & ".ASSAYANALYTES.STUDYID) AND (" & strSchema & ".ASSAY.ASSAYID = " & strSchema & ".ASSAYANALYTES.ASSAYID)) INNER JOIN " & strSchema & ".ASSAYANALYTEKNOWN ON (" & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL = " & strSchema & ".ASSAYANALYTEKNOWN.LEVELNUMBER) AND (" & strSchema & ".ASSAYANALYTES.ANALYTEINDEX = " & strSchema & ".ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (" & strSchema & ".ASSAYANALYTES.STUDYID = " & strSchema & ".ASSAYANALYTEKNOWN.STUDYID) AND (" & strSchema & ".ASSAYANALYTES.ASSAYID = " & strSchema & ".ASSAYANALYTEKNOWN.ASSAYID)) INNER JOIN " & strSchema & ".GLOBALANALYTES ON " & strSchema & ".ASSAYANALYTES.ANALYTEID = " & strSchema & ".GLOBALANALYTES.GLOBALANALYTEID "
            str3 = "WHERE(((" & strSchema & ".ANALYTICALRUN.STUDYID) = " & wID & ")) "
            str4 = "ORDER BY " & strSchema & ".ASSAYANALYTES.ANALYTEID, " & strSchema & ".ASSAYANALYTES.ANALYTEINDEX, " & strSchema & ".ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

        End If

        strSQL = str1 & str2 & str3 & str4

        'Console.WriteLine("tblConcLevelsForAssayIDs: " & strSQL)

        If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            rs.Close()
        End If
        rs.CursorLocation = CursorLocationEnum.adUseClient
        rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)

        tblConcLevelsForAssayIDs.Clear()
        tblConcLevelsForAssayIDs.AcceptChanges()
        tblConcLevelsForAssayIDs.BeginLoadData()
        daDoPr.Fill(tblConcLevelsForAssayIDs, rs)
        tblConcLevelsForAssayIDs.EndLoadData()

        '20161227 LEE: must do tblASSAYREPS also

        If boolAccess Then
            str1 = "SELECT ASSAYREPS.* "
            str2 = "FROM ASSAYREPS "
            str3 = "WHERE (((ASSAYREPS.STUDYID)=" & wID & "));"
        Else
            str1 = "SELECT " & strSchema & ".ASSAYREPS.* "
            str2 = "FROM " & strSchema & ".ASSAYREPS "
            str3 = "WHERE (((" & strSchema & ".ASSAYREPS.STUDYID)=" & wID & "));"
        End If

        strSQL = str1 & str2 & str3
        ''console.writeline(strSQL)
        ''''''''''''''''''''''''''''''''''''''''console.writeline(strSQL)
        If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            rs.Close()
        End If
        rs.CursorLocation = CursorLocationEnum.adUseClient
        Try
            rs.Open(strSQL, cn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Catch ex As Exception
            var1 = var1
        End Try

        rs.ActiveConnection = Nothing

        'rsspeciesmatrix
        tblASSAYREPS.Clear()
        tblASSAYREPS.AcceptChanges()
        tblASSAYREPS.BeginLoadData()
        daDoPr.Fill(tblASSAYREPS, rs)
        tblASSAYREPS.EndLoadData()

        If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            rs.Close()
        End If

        Try
            rs = Nothing
        Catch ex As Exception

        End Try


    End Sub

    Function IsStudyChanged() As Boolean

        IsStudyChanged = False

        Dim var1, var2

        var1 = Me.cbxStudy.Text
        '20180801 LEE:
        If Len(var1) = 0 Then
        Else
            If StrComp(var1, wWStudyName, CompareMethod.Text) = 0 Then
            Else
                IsStudyChanged = True
            End If
        End If

    End Function

    Function idCGet() As Int32

        Dim dgv As DataGridView
        Dim intRow As Int16
        Dim idC As Int32

        dgv = Me.dgvTables

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        Try
            idC = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
        Catch ex As Exception
            idC = 0
        End Try

        idCGet = idC

    End Function

    Private Sub cmdAddRows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddRows.Click

        Call AddRows()

    End Sub

    Sub AddRows()

        Dim dgvAR As DataGridView
        Dim dgvAS As DataGridView
        Dim row As DataGridViewRow
        Dim intASRows As Short
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim ctRunCols As Short
        Dim var1, var2
        Dim str1 As String
        Dim dvAS As System.Data.DataView
        Dim intTable As Short
        Dim intAnalyte As Short
        Dim boolA As Boolean
        Dim maxID
        Dim maxID1
        'Dim tblMaxID As System.Data.DataTable
        Dim drowsmaxid() As DataRow
        Dim intIS As Short
        Dim str2 As String
        Dim bool As Boolean
        Dim rowSel(10, 1)
        Dim intSel As Short
        Dim ctAssignedSamples As Short
        Dim idTS As Long
        Dim idConfigRT As Long
        Dim strAnalyte As String
        Dim int1 As Short
        Dim idT As Long
        Dim intRunID As Int32

        Dim dtNow As Date = Now

        Dim boolIS As Boolean

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If
        bool = boolCont
        boolCont = False 'do this to stop dgvAssignedSamples selectionchanged event

        Dim intAnalyteID As Int32
        Dim dgvA As DataGridView = Me.dgvAnalytes

        Dim intColSN As Short = 0 'column of samplename

        Cursor.Current = Cursors.WaitCursor


        ''find maxID for tblReportTable
        'str1 = "charTable = 'tblAssignedSamples'"
        'If boolGuWuOracle Then
        '    ta_tblMaxID.Fill(tblMaxID)
        'ElseIf boolGuWuAccess Then
        '    ta_tblMaxIDAcc.Fill(tblMaxID)
        'ElseIf boolGuWuSQLServer Then
        '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        'End If
        'tblMaxID = tblMaxID
        'drowsmaxid = tblMaxID.Select(str1)
        'maxID = drowsmaxid(0).Item("numMaxID")

        maxID = GetMaxID("tblAssignedSamples", 1, True)
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'maxID1 = maxID

        'Assign DataGridViews to local Variables
        dgvAR = Me.dgvAnalyticalRuns
        dgvAS = Me.dgvAssignedSamples
        intASRows = dgvAS.RowCount

        'deselect any selected rows in dgvAS
        dgvAS.ClearSelection()

        If dgvAR.SelectedRows.Count = 0 Then
            Exit Sub
        End If

        'Add direct access to the Assigned Samples table
        dvAS = dgvAS.DataSource
        dvAS.AllowNew = True

        If Me.dgvTables.CurrentRow Is Nothing Then
            MsgBox("Please select a table.", MsgBoxStyle.Information, "Select a table...")
            Me.dgvTables.Select()
            Exit Sub
        End If

        'Set up indexes for current selections and associated tables
        intTable = Me.dgvTables.CurrentRow.Index
        intAnalyte = Me.dgvAnalytes.CurrentRow.Index
        ctRunCols = Me.dgvAnalyticalRuns.Columns.Count

        idTS = id_tblStudies
        idConfigRT = Me.dgvTables.Rows.Item(intTable).Cells("id_tblConfigReportTables").Value
        idT = Me.dgvTables.Rows.Item(intTable).Cells("ID_TBLREPORTTABLE").Value
        strAnalyte = Me.dgvAnalytes.Rows.Item(intAnalyte).Cells(0).Value


        'determine if selected analyte is internal standard
        str2 = Me.dgvAnalytes("IsIntStd", intAnalyte).Value
        If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
            intIS = -1
            boolIS = True
        Else
            intIS = 0
            boolIS = False
        End If

        intAnalyteID = NZ(dgvA("ANALYTEID", intAnalyte).Value, -1) 'if IntStd, dgv value is null

        intSel = dgvAR.SelectedRows.Count
        ReDim rowSel(10, intSel)

        Dim strF As String
        Dim tblNew As System.Data.DataTable = dvAS.ToTable("a")
        Dim rowsNew() As DataRow

        ctAssignedSamples = 0
        Dim ctColSel As Short
        Dim strM As String

        Dim boolGotConcLevels As Boolean = False
        Dim boolGotSomeLevels As Boolean = False

        '20151104 LEE:
        'first find unique studyids in selection
        'and save in array
        '20160419 LEE:
        'this logic isn't true
        'if new study has been configured,
        'all id_tblStudies will be the same
        Dim id1 As Int64
        Dim id2 As Int64
        Dim intID As Short
        Dim arrID(100)
        Dim boolHit As Boolean
        intID = 0
        int1 = 0

        '20151104 LEE:
        'now establish boolGotConcLevels
        boolGotConcLevels = False
        boolGotSomeLevels = False


        Dim strISIntStd As String
        Dim boolIsIntStd As Boolean
        strISIntStd = dgvA("IsIntStd", intAnalyte).Value
        If StrComp(strIsIntStd, "Yes", CompareMethod.Text) = 0 Then
            boolIsIntStd = True
        Else
            boolIsIntStd = False
        End If


        Select Case idConfigRT
            Case 13, 14, 15

                If boolIsIntStd And BOOLISCOMBINELEVELS Then
                    boolGotConcLevels = True
                    boolGotSomeLevels = False
                Else
                    For Each row In dgvAR.SelectedRows
                        var1 = row.Cells("ASSAYLEVEL").Value
                        If IsDBNull(var1) Then
                            boolGotSomeLevels = True
                        Else
                            boolGotConcLevels = True
                            If boolGotConcLevels And boolGotSomeLevels Then
                                Exit For
                            End If
                        End If
                    Next
                End If

            Case Else
                For Each row In dgvAR.SelectedRows
                    var1 = row.Cells("ASSAYLEVEL").Value
                    If IsDBNull(var1) Then
                        boolGotSomeLevels = True
                    Else
                        boolGotConcLevels = True
                        If boolGotConcLevels And boolGotSomeLevels Then
                            Exit For
                        End If
                    End If
                Next
        End Select

        'These tables do not need nominal concentrations (so no need for Assay Levels)
        '22 Stock Solution Stability Assessment
        '30 Incurred Samples
        '33 System Suitability Table
        '38 Method Trial Incurred Blinded Samples

        Select Case idConfigRT
            Case 22, 30, 33, 35, 38 'No warning needed
            Case Else

                If boolGotConcLevels Then
                    If boolGotSomeLevels And boolAutoAssign = False Then
                        strM = "Please note that some added samples do not contain an Assay Level in the Assay Level column."
                        strM = strM & ChrW(10) & ChrW(10)
                        strM = strM & "User must manually assign Nominal Concentrations in the Assigned Samples table."
                        MsgBox(strM, vbInformation, "No Assay Levels...")
                    End If
                Else
                    strM = "Please note that the added samples do not have an Assay Level in the Assay Level column."
                    strM = strM & ChrW(10) & ChrW(10)
                    strM = strM & "User must manually assign Nominal Concentrations in the Assigned Samples table."
                    MsgBox(strM, vbInformation, "No Assay Levels...")
                End If

        End Select

        'MsgBox("Here1")

        Dim rowsDGVAR As DataGridViewSelectedRowCollection = dgvAR.SelectedRows 'dgvAnalyticalRuns

        '20180312 LEE:
        'https://stackoverflow.com/questions/17573894/reversing-the-way-rows-get-entered-into-a-data-table
        'the selectedrows collection lists the actual selected rows in reverse order
        'in some instances (Ad Hoc Stability Comparison - T0 samples), we want the selected rows to be entered in the actual order
        'must change loop


        int1 = 0
        'For Each row In dgvAR.SelectedRows
        Dim intAR As Int32 = dgvAR.SelectedRows.Count

        Dim boolFL As Boolean = boolFormLoad
        boolFormLoad = True 'need to stop some selection actions

        For i As Int32 = intAR - 1 To 0 Step -1

            row = dgvAR.SelectedRows(i)

            If i = 158 Then
                var1 = var1 'debug
            End If

            int1 = int1 + 1

            'first determine if data already exists

            id1 = Me.txtStudyID.Text ' row.Cells("StudyID").Value

            str1 = "id_tblConfigReportTables = " & idConfigRT
            str1 = str1 & " AND ID_TBLSTUDIES = " & id1 ' id_tblStudies
            '20151102 LEE
            'remove these next two lines for now. Inspect later
            str1 = str1 & " AND ID_TBLSTUDIES2 = " & CLng(Me.txtStudyID.Text) 'this may be external study
            str1 = str1 & " AND CHARSTUDYNAME2 = '" & Me.cbxStudy.Text & "'"

            str1 = str1 & " AND CHARANALYTE = '" & CleanText(strAnalyte) & "'"
            str1 = str1 & " AND RUNID = " & row.Cells("RUNID").Value
            str1 = str1 & " AND ANALYTEINDEX = " & row.Cells("ANALYTEINDEX").Value
            str1 = str1 & " AND MASTERASSAYID = " & row.Cells("MASTERASSAYID").Value
            str1 = str1 & " AND ASSAYID = " & row.Cells("ASSAYID").Value
            'str1 = str1 & " AND RUNSAMPLESEQUENCENUMBER = " & row.Cells("RUNSAMPLESEQUENCENUMBER").Value
            str1 = str1 & " AND RUNSAMPLEORDERNUMBER = " & row.Cells("RUNSAMPLEORDERNUMBER").Value
            str1 = str1 & " AND ID_TBLREPORTTABLE = " & idT
            str1 = str1 & " AND BOOLINTSTD = " & intIS

            '20160227 LEE: For some reason, if ANALYTEID and INTGROUP are added to the query, 0 rows are returned
            'str1 = str1 & " AND INTANALYTEID = " & row.Cells("ANALYTEID").Value
            'str1 = str1 & " AND INTGROUP = " & gintGroup

            strF = str1

            Erase rowsNew
            var1 = tblNew.Rows.Count

            ''debug
            'If int1 = 1 Then
            '    'console.writeline(strF)
            '    var1 = ""
            '    For Count1 = 0 To tblNew.Columns.Count - 1
            '        var2 = tblNew.Columns(Count1).ColumnName
            '        var1 = var1 & ChrW(9) & var2
            '    Next
            '    'console.writeline(var1)
            '    For Count2 = 0 To tblNew.Rows.Count - 1
            '        var1 = ""
            '        For Count1 = 0 To tblNew.Columns.Count - 1
            '            var2 = tblNew.Rows(Count2).Item(Count1)
            '            var1 = var1 & ChrW(9) & var2
            '        Next
            '        'console.writeline(var1)
            '    Next
            'End If

            rowsNew = tblNew.Select(strF)

            ''''''''''''''''Console.WriteLine(str1)

            If rowsNew.Length = 0 Then 'data doesn't exist, continue

                ctAssignedSamples = ctAssignedSamples + 1

                '* 1st, copy explicit data into the row
                maxID = maxID + 1
                intASRows = intASRows + 1
                boolCont = False 'to disable FindSamples

                Dim dvASRow As DataRowView = dvAS.AddNew

                boolCont = True
                var1 = Me.dgvTables.Rows.Item(intTable).Cells("id_tblConfigReportTables").Value
                dvASRow("id_tblAssignedSamples") = maxID
                dvASRow("id_tblConfigReportTables") = idConfigRT 'var1
                dvASRow("id_tblStudies") = id_tblStudies
                dvASRow("id_tblStudies2") = CLng(Me.txtStudyID.Text)
                dvASRow("charStudyName2") = Me.cbxStudy.Text
                dvASRow("BOOLINTSTD") = intIS
                dvASRow("ID_TBLREPORTTABLE") = idT
                dvASRow("BOOLOUTLIER") = 0
                dvASRow("BOOLINCURRED") = 0

                dvASRow("BOOLEXCLSAMPLE") = 0

                If gAllowGuWuAccCrit And LAllowGuWuAccCrit And Me.chkUSEGUWUACCCRIT.Checked Then
                    dvASRow("BOOLUSEGUWUACCCRIT") = -1
                Else
                    dvASRow("BOOLUSEGUWUACCCRIT") = 0
                End If
                dvASRow("BOOLINCURRED") = 0

                '* 2nd, copy standard data into the row
                ctColSel = 0

                For Count1 = 0 To ctRunCols - 1
                    str1 = Me.dgvAnalyticalRuns.Columns.Item(Count1).Name
                    var1 = row.Cells(str1).Value 'from dgvAnalyticalRuns
                    If StrComp(str1, "ALIQUOTFACTOR", CompareMethod.Text) = 0 Then
                        str1 = "ALIQUOTFACTOR"
                    End If

                    If InStr(1, str1, "GROUP", CompareMethod.Text) > 0 Then
                        var1 = var1
                    End If

                    dvASRow(str1) = var1
                    Select Case str1
                        Case "ANALYTEID"
                            dvASRow("INTANALYTEID") = var1
                    End Select

                    'Also, build selection criteria
                    boolA = False
                    Select Case str1
                        Case "RUNID" '2
                            boolA = True
                        Case "ANALYTEINDEX" '6
                            boolA = True
                        Case "MASTERASSAYID" '5
                            boolA = True
                        Case "ASSAYID" '4
                            boolA = True
                        Case "RUNSAMPLEORDERNUMBER" '3
                            boolA = True
                        Case "ANALYTEID" '1
                            boolA = True
                        Case "SAMPLENAME"
                            intColSN = ctColSel + 1
                            boolA = True
                    End Select
                    If boolA Then
                        ctColSel = ctColSel + 1
                        rowSel(ctColSel, ctAssignedSamples) = var1
                    End If

                    var1 = dvASRow("INTGROUP")
                    var1 = var1

                Next


                Try
                    dvASRow("UPSIZE_TS") = dtNow
                Catch ex As Exception
                    var1 = var1
                End Try

                dvASRow("CHARANALYTE") = strAnalyte 'MUST DO THIS AFTER LOOP!!!
                '20180126 LEE: This too!!!
                dvASRow("INTGROUP") = gintGroup

                'Now set the Nominal Mass of the Analyte if it has a level
                Dim strRSK = dvASRow("RUNSAMPLEKIND")

                intIS = NZ(dvASRow("BOOLINTSTD"), 0)


                'If ((StrComp(strRSK, "STANDARD", CompareMethod.Text) = 0) Or (StrComp(strRSK, "QC", CompareMethod.Text) = 0)) And (intIS = 0) Then
                '20150511 LEE: Note that 'QC' is a hand-entered value in Watson. An actual QC may not be assigned a sample type called 'QC'
                'Look at some of WIL Meth Val studies

                '20160227 LEE: made a slight modification to this logic so that IntStd samples now return a nomconc
                'Changed:
                'intAnalyteID = dgvA("ANALYTEID", dgvA.CurrentRow.Index).Value
                '   to
                'intAnalyteID = NZ(dvASRow("ANALYTEID"), -1) 

                'If intIS = 0 Then

                Dim dblNominalConcentration As Decimal
                Dim strRSK1 As String = NZ(dvASRow("RUNSAMPLEKIND"), "NA")
                Dim intStudyID, intAID, intAIn, intAL, intMAID As Int32

                'Compose the DataTable of Concentration Levels (matching Analyte Level with Level Number) for all AssayIDs in this study AnalyteID

                intAnalyteID = NZ(dvASRow("ANALYTEID"), -1) ' dgvA("ANALYTEID", dgvA.CurrentRow.Index).Value
                intAID = NZ(dvASRow("ASSAYID"), -1)
                intAIn = NZ(dvASRow("ANALYTEINDEX"), -1)
                intAL = NZ(dvASRow("ASSAYLEVEL"), -1)
                intMAID = NZ(dvASRow("MASTERASSAYID"), -1)
                If (intAID = -1) Or (intAIn = -1) Or (intAL = -1) Then
                    dblNominalConcentration = -1
                Else
                    dblNominalConcentration = FindNomConc(tblConcLevelsForAssayIDs, intAID, strRSK1, intAIn, intAL, intMAID)
                    dvASRow("NOMCONC") = dblNominalConcentration
                End If

                intRunID = dvASRow("RUNID")

                Dim strLabel As String
                'HELPER1
                '20151106 LEE:
                'Not all QC's have 'QC' as sample type
                'Instead, filter out Standards
                If StrComp(strRSK, "STANDARD", CompareMethod.Text) = 0 Then
                Else
                    str1 = FindLabelHelper1(intAID, strRSK1, intAIn, intAL, intAnalyteID, intRunID, intAL)
                    'Note: BOOLUSESTDCOLLABELS is evaluated in ReturnStdQC
                    strLabel = ReturnStdQC(NZ(str1, ""))
                    dvASRow("CHARHELPER1") = strLabel
                End If

                'End If

                var1 = dvASRow.Item("INTGROUP")

                dvASRow.EndEdit()

                'rowSel(ctAssignedSamples) = intASRows - 1
                'record rowSel info
                '1=runid, 2=analyteindex, 3=masterassayid, 4=RUNSAMPLESEQUENCENUMBER



            Else
                'ctAssignedSamples = ctAssignedSamples + 1
                'rowSel(ctAssignedSamples) = row.Cells("id_tblAssignedSamples").Value
                'rowSel(ctAssignedSamples) = row.Index
            End If
        Next

        boolFormLoad = boolFL

        Select Case idConfigRT
            Case 13, 14, 15
                Dim strIS As String
                Dim numLo As Decimal
                Dim numLoO As Decimal
                'strIS = NZ(dgvA("IsIntStd", dgvA.CurrentRow.Index).Value, "No")
                strIS = NZ(dgvA("IsIntStd", dgvA.CurrentRow.Index).Value, "No")
                If StrComp(strIS, "Yes", CompareMethod.Text) = 0 And BOOLISCOMBINELEVELS Then

                    'find lowest conc level and assign it to all levels
                    Dim dv As DataView = dgvAS.DataSource

                    numLo = 99999
                    numLoO = numLo
                    For Count1 = 0 To dv.Count - 1
                        var1 = NZ(dv(Count1).Item("NOMCONC"), numLo)
                        If var1 < numLo Then
                            numLo = var1
                        End If
                    Next
                    If numLo = numLoO Then
                    Else
                        'assign lowest value to all rows
                        For Count1 = 0 To dv.Count - 1
                            dv(Count1).BeginEdit()
                            dv(Count1).Item("NOMCONC") = numLo
                            dv(Count1).EndEdit()
                        Next
                    End If
                End If


        End Select

        'the next bunch of lines will cause a dgvAS selectionchange
        'need to disable further actions
        Dim boolC As Boolean = boolCont
        boolCont = False

        'select all dgvAR rows again
        Dim intARCell As Short
        For Count1 = 0 To dgvAR.Columns.Count - 1
            If dgvAR.Columns(Count1).Visible Then
                intARCell = Count1
                Exit For
            End If
        Next
        int1 = 0
        For Each row In rowsDGVAR
            int1 = int1 + 1
            If int1 = 1 Then
                dgvAR.CurrentCell = dgvAR.Rows(row.Index).Cells(intARCell)
            End If
            dgvAR.Rows(row.Index).Selected = True
        Next


        'need to ensure all rows that have been added are now selected
        'unselect all dgvAS rows
        'find visible cell
        Dim intASCell As Short
        For Count1 = 0 To dgvAS.Columns.Count - 1
            If dgvAS.Columns(Count1).Visible Then
                intASCell = Count1
                Exit For
            End If
        Next


        dgvAS.ClearSelection()
        dgvAS.CurrentCell = Nothing


        int1 = 0
        For Each row In dgvAR.SelectedRows
            str1 = row.Cells("SAMPLENAME").Value
            For Count2 = 0 To dgvAS.Rows.Count - 1
                str2 = dgvAS("SAMPLENAME", Count2).Value
                If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                    int1 = int1 + 1
                    If int1 = 1 Then
                        dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                    End If
                    dgvAS.Rows(Count2).Selected = True
                    Exit For
                End If
            Next
        Next

        boolCont = boolC

        'check for additional input required of user
        Dim intRecSel As Short = -1
        Dim boolDoH1 As Boolean = False
        Dim boolDoH2 As Boolean = False
        Dim boolDoH2Text As Boolean = False
        Dim strRunID As String = ""

        If DoHelper1(idConfigRT) And boolAutoAssign = False Then

            Dim frm As New frmAddRowsChoice
            frm.idCT = idConfigRT
            frm.idRT = idT
            frm.frm = Me
            frm.ShowDialog()

            If frm.boolCancel Or frm.chkIgnore.Checked Then
            Else
                Select Case idConfigRT

                    Case Is = 1 'Summary of Analytical Runs
                    Case Is = 2 'Summary of Regression Constants
                    Case Is = 3 'Summary of Back-Calculated Calibration Std Conc
                    Case Is = 4 'Summary of Interpolated QC Std Conc
                    Case Is = 5 'Summary of Samples
                    Case Is = 6 'Summary of Reassayed Samples
                    Case Is = 7 'Summary of Repeat Samples
                    Case Is = 11 'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
                        'Case Is = 12 'Summary of Interpolated Dilution QC Concentrations
                    Case Is = 13 'Summary of Combined Recovery

                        If frm.rbQC.Checked Then
                            'select dgvHelper1 first row
                            intRecSel = 0
                        ElseIf frm.rbRS.Checked Then
                            intRecSel = 1
                        End If

                        boolDoH1 = True

                    Case Is = 14 'Summary of True Recovery

                        If frm.rbQC.Checked Then
                            'select dgvHelper1 first row
                            intRecSel = 0
                        ElseIf frm.rbPES.Checked Then
                            intRecSel = 1
                        End If

                        boolDoH1 = True

                    Case Is = 15 'Summary of Suppression/Enhancement

                        If frm.rbRS.Checked Then
                            'select dgvHelper1 first row
                            intRecSel = 0
                        ElseIf frm.rbPES.Checked Then
                            intRecSel = 1
                        End If

                        boolDoH1 = True

                    Case Is = 17 'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments

                        intRecSel = frm.intRBSel
                        boolDoH1 = True

                    Case Is = 18 'Summary of [Period Temp] Stability in Matrix
                        'Case Is = 19 'Summary of Freeze/Thaw [#Cycles] Stability in Matrix
                        'Case Is = 21 '[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
                    Case Is = 22 '[Period Temp] Stock Solution Stability Assessment

                        intRecSel = frm.intRBSel
                        boolDoH1 = True

                    Case Is = 23 '[Period Temp] Spiking Solution Stability Assessment

                        intRecSel = frm.intRBSel
                        boolDoH1 = True

                    Case Is = 29 '[Period Temp] Long-Term QC Std Storage Stability

                        intRecSel = frm.intRBSel
                        boolDoH2 = True

                    Case Is = 30 'Incurred Samples
                    Case Is = 31, 12, 19, 21 'Ad Hoc QC Stability Table
                    Case Is = 32 'Ad Hoc QC Stability Comparison Table

                        intRecSel = 0 'set to 0 to trigger later code
                        strRunID = frm.txtAdHocStabComp.Text
                        boolDoH2Text = True

                    Case Is = 33 'System Suitability Table v1
                    Case Is = 34 'Selectivity in Individual Lots Table v1
                    Case Is = 35 'Carryover in Individual Lots Table v1

                        intRecSel = frm.intRBSel
                        If frm.boolrbULOQVis Or intRecSel = 0 Then
                        Else
                            intRecSel = intRecSel - 1
                        End If
                        boolDoH1 = True

                    Case Is = 36 'Method Trial Back-Calculated Calibration Std Conc v1
                    Case Is = 37 'Method Trial Control and Fortified QC Samples v1
                    Case Is = 38 'Method Trial Incurred Blinded Samples v1

                End Select

                If intRecSel = -1 Then
                Else

                    If boolDoH1 Then
                        Try
                            Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                            Me.dgvHelper1.Rows(intRecSel).Selected = True
                            Call Helper1Click()
                        Catch ex As Exception

                        End Try
                    ElseIf boolDoH2 Then
                        Me.dgvHelper2.CurrentCell = dgvHelper2.Rows(intRecSel).Cells(1)
                        Me.dgvHelper2.Rows(intRecSel).Selected = True
                        Call Helper2Click()
                    ElseIf boolDoH2Text Then
                        If Len(strRunID) = 0 Then
                        Else
                            Me.txtHelper2.Text = strRunID
                            Call Helper2Click()
                        End If
                    End If

                End If

            End If

            frm.Close()
            frm.Dispose()
        End If

        If maxID = maxID1 Then
        Else

            Call PutMaxID("tblAssignedSamples", maxID)

        End If



        Call ASNum()

        'GoTo end1


        '4th, Ensure appropriate rows are selected in both Analytical Runs and Assigned Samples

        Dim boolT As Boolean
        boolT = boolFormLoad
        boolFormLoad = True
        'deselect any selected rows
        For Each row In dgvAS.Rows
            row.Selected = False
        Next

        Dim var3, var4, var5, var6
        Dim var1a, var2a, var3a, var4a, var5a, var6a
        Dim minRow As Short = 0
        '20160817 LEE: new rows may be interspersed
        'set minrow to 0
        minRow = 0 ' dgvAS.RowCount

        Try

            dgvAS.ClearSelection()
            dgvAS.CurrentCell = Nothing
            int1 = 0
            For Count1 = 1 To ctAssignedSamples
                str1 = rowSel(intColSN, Count1)
                For Count2 = 0 To dgvAS.Rows.Count - 1
                    str2 = dgvAS("SAMPLENAME", Count2).Value
                    If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                        int1 = int1 + 1
                        If int1 = 1 Then
                            dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                        End If
                        dgvAS.Rows(Count2).Selected = True
                        Exit For
                    End If
                Next
            Next

            'find indext of first added row
            Dim rowsAA As DataGridViewSelectedRowCollection = dgvAS.SelectedRows
            Try
                minRow = rowsAA(0).Index
            Catch ex As Exception
                minRow = 0
            End Try



        Catch ex As Exception

        End Try
        boolFormLoad = boolT

skip1:

        'ensure first selected row is near the top of the grid
        If ctAssignedSamples = 0 Then
        Else
            Try
                If minRow = 0 Then
                    dgvAS.FirstDisplayedScrollingRowIndex = 0
                Else
                    dgvAS.FirstDisplayedScrollingRowIndex = minRow ' - 1
                End If
            Catch ex As Exception

            End Try
        End If

        'scroll dgvAS all the way to the right
        var1 = dgvAS.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
        dgvAS.HorizontalScrollingOffset = var1



        ''need to ensure all rows that have been added are now selected
        ''unselect all dgvAS rows
        ''find visible cell

        'dgvAS.ClearSelection()
        'dgvAS.CurrentCell = Nothing
        'int1 = 0
        'For Each row In dgvAR.SelectedRows
        '    str1 = row.Cells("SAMPLENAME").Value
        '    For Count2 = 0 To dgvAS.Rows.Count - 1
        '        str2 = dgvAS("SAMPLENAME", Count2).Value
        '        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
        '            int1 = int1 + 1
        '            If int1 = 1 Then
        '                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
        '            End If
        '            dgvAS.Rows(Count2).Selected = True
        '            Exit For
        '        End If
        '    Next
        'Next



end1:
        boolCont = bool

        Cursor.Current = Cursors.Default


        'Me.dgvAssignedSamples.AutoResizeColumns()
    End Sub


    Function DoHelper1(idCT As Int64) As Boolean

        DoHelper1 = False

        Select Case idCT

            Case 13, 14, 15, 17, 22, 23, 29, 32, 35
                DoHelper1 = True

        End Select

    End Function

    Function FindLabelHelper1(intAssayID As Int32, strKnownType As String, intAnalyteIndex As Int32, intAssayLevel As Int32, intAnalyteID As Int32, intRunID As Int32, intLevelNumber As Short) As String

        FindLabelHelper1 = ""

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strF As String

        '20151106 LEE:
        'don't want to return labels for certain tables
        Dim intRow As Short
        Dim dgv As DataGridView = Me.dgvTables
        intRow = dgv.CurrentRow.Index
        Dim id As Int64
        id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value

        Select Case id
            Case 13, 14, 15, 17, 22, 23
                FindLabelHelper1 = ""
            Case Else
                'strF = " ASSAYID = " & intAssayID & " AND KNOWNTYPE = '" & strKnownType & "' AND ANALYTEINDEX = " & intAnalyteIndex & " AND ANALYTEID = " & intAnalyteID & " AND RUNID = " & intRunID & " AND LEVELNUMBER = " & intLevelNumber
                'with ASSAYREPS, there is no ANALYTEINDEX
                'ID is the same for all ANANALYTEINDEX's
                strF = " ASSAYID = " & intAssayID & " AND KNOWNTYPE = '" & strKnownType & "' AND LEVELNUMBER = " & intLevelNumber

                'Dim rows() As DataRow = tblQCRunIDs.Select(strF)
                Dim rows() As DataRow = tblASSAYREPS.Select(strF)

                If rows.Length = 0 Then
                    FindLabelHelper1 = ""
                Else
                    FindLabelHelper1 = rows(0).Item("ID")
                End If
        End Select

    End Function



    'Function getTblAnalyteConcLevelsForAssay(ByVal intStudyID As Int32)
    Function getTblAnalyteConcLevelsForAssay(arrID1) As DataTable

        '20160419 LEE: Deprecated
        'placed in doPrepare


        'Creates a table with all concentration levels for all Assays, all Analytes, all Sample Types
        Dim con As New ADODB.Connection
        Dim rs As New ADODB.Recordset
        Dim str1, str2, str3, str4, strSQL As String
        Dim id1 As Int64
        Dim Count1 As Short
        Dim intIDs As Short = UBound(arrID1)

        getTblAnalyteConcLevelsForAssay = New DataTable

        str1 = "SELECT DISTINCT ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ASSAYANALYTEKNOWN.LEVELNUMBER, GLOBALANALYTES.ANALYTEDESCRIPTION, ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUN.ASSAYID, ASSAYANALYTEKNOWN.CONCENTRATION, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ASSAYANALYTEKNOWN.STUDYID "
        str2 = "FROM ((((ANALYTICALRUNSAMPLE INNER JOIN ANALYTICALRUN ON (ANALYTICALRUNSAMPLE.RUNID = ANALYTICALRUN.RUNID) AND (ANALYTICALRUNSAMPLE.STUDYID = ANALYTICALRUN.STUDYID)) INNER JOIN ASSAY ON (ANALYTICALRUN.RUNID = ASSAY.RUNID) AND (ANALYTICALRUN.STUDYID = ASSAY.STUDYID)) INNER JOIN ASSAYANALYTES ON (ASSAY.ASSAYID = ASSAYANALYTES.ASSAYID) AND (ASSAY.STUDYID = ASSAYANALYTES.STUDYID)) INNER JOIN ASSAYANALYTEKNOWN ON (ASSAYANALYTES.ASSAYID = ASSAYANALYTEKNOWN.ASSAYID) AND (ASSAYANALYTES.STUDYID = ASSAYANALYTEKNOWN.STUDYID) AND (ASSAYANALYTES.ANALYTEINDEX = ASSAYANALYTEKNOWN.ANALYTEINDEX) AND (ANALYTICALRUNSAMPLE.RUNSAMPLEKIND = ASSAYANALYTEKNOWN.KNOWNTYPE) AND (ANALYTICALRUNSAMPLE.ASSAYLEVEL = ASSAYANALYTEKNOWN.LEVELNUMBER)) INNER JOIN GLOBALANALYTES ON (ASSAYANALYTES.ANALYTEID = GLOBALANALYTES.GLOBALANALYTEID)"

        If intIDs = 1 Then
            id1 = arrID1(1)
            str3 = "WHERE((ASSAYANALYTEKNOWN.STUDYID) = " & id1 & ") "
        Else
            'WHERE (((TBLSTUDIES.ID_TBLSTUDIES) = 44 Or (TBLSTUDIES.ID_TBLSTUDIES) = 45));
            str3 = "WHERE (("
            For Count1 = 1 To intIDs
                id1 = arrID1(Count1)
                If Count1 = intIDs Then
                    str1 = str1 & "(TBLSTUDIES.ID_TBLSTUDIES) = " & id1 & "))"
                Else
                    str1 = "(TBLSTUDIES.ID_TBLSTUDIES) = " & id1 & " OR "
                End If
            Next
            str3 = str3 & str1
        End If

        str4 = "ORDER BY ASSAYANALYTES.ANALYTEID, ASSAYANALYTES.ANALYTEINDEX, ANALYTICALRUNSAMPLE.ASSAYLEVEL;"

        strSQL = str1 & str2 & str3 & str4

        ''console.writeline("getTblAnalyteConcLevelsForAssay: " & strSQL)

        con.Open(constrCur)

        rs.CursorLocation = CursorLocationEnum.adUseClient
        rs.Open(strSQL, con, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        getTblAnalyteConcLevelsForAssay.Clear()
        getTblAnalyteConcLevelsForAssay.AcceptChanges()
        getTblAnalyteConcLevelsForAssay.BeginLoadData()
        daDoPr.Fill(getTblAnalyteConcLevelsForAssay, rs)
        getTblAnalyteConcLevelsForAssay.EndLoadData()

        rs.Close()
        rs = Nothing

        con.Close()
        con = Nothing

    End Function

    Function FindNomConc(ByRef tblConcLevelsForAssayIDs As DataTable, intAssayID As Int32, strRunSampleKind As String, intAnalyteIndex As Int32, _
                         intAssayLevel As Int32, intMasterAssayID As Int32) As Decimal

        'Finds the nominal concentration of this level
        'tblConcLevelsForAssayIDs is a DataTable with nominal concentrations for all levels of all AnalyteIndexes for all AssayIfor this study

        Dim str1 As String
        Dim rows() As DataRow

        'Extract Specific AssayID and AnalyteIndex from the table
        'ASSAYANALYTEKNOWN.KNOWNTYPE: limited sample types to QC, STANDARD, STABILITY
        'ANALYTICALRUNSAMPLE.RUNSAMPLEKIND: KNOWNTYPE + UNKNOWN + custom sample types

        Dim strKnownType As String

        '20160816 LEE:
        'first query for RUNSAMPLEKIND
        Select Case strRunSampleKind
            Case "STANDARD"
                str1 = "ASSAYID = " & intAssayID & " AND ANALYTEINDEX = " & intAnalyteIndex & " AND RUNSAMPLEKIND ='" & strRunSampleKind & "' AND  ASSAYLEVEL = " & intAssayLevel & " AND KNOWNTYPE = 'STANDARD'"
            Case Else
                str1 = "ASSAYID = " & intAssayID & " AND ANALYTEINDEX = " & intAnalyteIndex & " AND RUNSAMPLEKIND ='" & strRunSampleKind & "' AND ASSAYLEVEL = " & intAssayLevel & " AND (KNOWNTYPE = 'QC' OR KNOWNTYPE = 'STABILITY')"
                'If IsStudyChanged() Then
                '    str1 = "ANALYTEINDEX = " & intAnalyteIndex & " AND RUNSAMPLEKIND ='" & strRunSampleKind & "' AND ASSAYLEVEL = " & intAssayLevel & " AND (KNOWNTYPE = 'QC' OR KNOWNTYPE = 'STABILITY')"
                'Else
                '    str1 = "ASSAYID = " & intAssayID & " AND ANALYTEINDEX = " & intAnalyteIndex & " AND RUNSAMPLEKIND ='" & strRunSampleKind & "' AND ASSAYLEVEL = " & intAssayLevel & " AND (KNOWNTYPE = 'QC' OR KNOWNTYPE = 'STABILITY')"
                'End If
        End Select

        '''Console.WriteLine(str1)
        rows = tblConcLevelsForAssayIDs.Select(str1)

        If NZ(rows.Length, 0) = 0 Then
            FindNomConc = -1
        Else
            FindNomConc = rows(0).Item("CONCENTRATION")
        End If

    End Function

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim dgvAR As DataGridView = Me.dgvAnalyticalRuns

        Dim dv As DataView = dgvAR.DataSource

        Dim strF As String

        strF = "((ASSAYID = 196720 OR ASSAYID = 196739 OR ASSAYID = 196767 OR ASSAYID = 196793 OR ASSAYID = 196802 OR ASSAYID = 196803 OR ASSAYID = 197320 OR ASSAYID = 197351 OR ASSAYID = 197487 OR ASSAYID = 201003 OR ASSAYID = 201044 OR ASSAYID = 201045 OR ASSAYID = 201056 OR ASSAYID = 201255) AND ANALYTEID = 1314 AND ((RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3) AND RUNTYPEID <> 3) AND STUDYID = 3951) AND ((SAMPLENAME LIKE '*SelDB 1*' OR SAMPLENAME LIKE '*SelDB 2*' OR SAMPLENAME LIKE '*SelDB 3*' OR SAMPLENAME LIKE '*SelDB 4*' OR SAMPLENAME LIKE '*SelDB 5*' OR SAMPLENAME LIKE '*SelDB 6*' OR SAMPLENAME LIKE '*SelDB 7*' OR SAMPLENAME LIKE '*SelDB 8*' OR SAMPLENAME LIKE '*SelDB 9*' OR SAMPLENAME LIKE '*SelDB 10*' OR SAMPLENAME LIKE '*SelBLK 1*' OR SAMPLENAME LIKE '*SelBLK 2*' OR SAMPLENAME LIKE '*SelBLK 3*' OR SAMPLENAME LIKE '*SelBLK 4*' OR SAMPLENAME LIKE '*SelBLK 5*' OR SAMPLENAME LIKE '*SelBLK 6*' OR SAMPLENAME LIKE '*SelBLK 7*' OR SAMPLENAME LIKE '*SelBLK 8*' OR SAMPLENAME LIKE '*SelBLK 9*' OR SAMPLENAME LIKE '*SelBLK 10*' OR SAMPLENAME LIKE '*LLOQ*')) AND ((RUNID < 0))"

        strF = "((ASSAYID = 196720 OR ASSAYID = 196739 OR ASSAYID = 196767 OR ASSAYID = 196793 OR ASSAYID = 196802 OR ASSAYID = 196803 OR ASSAYID = 197320 OR ASSAYID = 197351 OR ASSAYID = 197487 OR ASSAYID = 201003 OR ASSAYID = 201044 OR ASSAYID = 201045 OR ASSAYID = 201056 OR ASSAYID = 201255) AND ANALYTEID = 1314 AND ((RUNANALYTEREGRESSIONSTATUS = 1 OR RUNANALYTEREGRESSIONSTATUS = 2 OR RUNANALYTEREGRESSIONSTATUS = 3) AND RUNTYPEID <> 3) AND STUDYID = 3951) AND ((SAMPLENAME LIKE '*SelDB 1*' OR SAMPLENAME LIKE '*SelDB 2*' OR SAMPLENAME LIKE '*SelDB 3*' OR SAMPLENAME LIKE '*SelDB 4*' OR SAMPLENAME LIKE '*SelDB 5*' OR SAMPLENAME LIKE '*SelDB 6*' OR SAMPLENAME LIKE '*SelDB 7*' OR SAMPLENAME LIKE '*SelDB 8*' OR SAMPLENAME LIKE '*SelDB 9*' OR SAMPLENAME LIKE '*SelDB 10*' OR SAMPLENAME LIKE '*SelBLK 1*' OR SAMPLENAME LIKE '*SelBLK 2*' OR SAMPLENAME LIKE '*SelBLK 3*' OR SAMPLENAME LIKE '*SelBLK 4*' OR SAMPLENAME LIKE '*SelBLK 5*' OR SAMPLENAME LIKE '*SelBLK 6*' OR SAMPLENAME LIKE '*SelBLK 7*' OR SAMPLENAME LIKE '*SelBLK 8*' OR SAMPLENAME LIKE '*SelBLK 9*' OR SAMPLENAME LIKE '*SelBLK 10*' OR SAMPLENAME LIKE '*LLOQ*')) AND ((RUNID < 0))"

        strF = dv.RowFilter


        Dim int1 As Int16

        dv.RowFilter = strF

        int1 = dv.Count

        MsgBox(strF)

    End Sub

    Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click

        boolCancelButton = True
        Dim var1

        'This is actually cmdCancel
        Cursor.Current = Cursors.WaitCursor

        If boolAutoAssign Then
        Else
            Me.lblWait.Visible = True
            Me.lblWait.Refresh()
        End If

        tblAssignedSamples.RejectChanges()

        'Call FillAssignedSamplesDGV()

        Cursor.Current = Cursors.WaitCursor

        'first reset
        Dim lng1 As Int64
        lng1 = Me.txtStudyID.Text
        If lng1 = id_tblStudies Then 'ignore
        Else
            Call ReturnStudyToOriginal()
        End If

        Try
            Call InitializeFilterRunID(Me.dgvAnalytes.CurrentRow.Index)
        Catch ex As Exception
            var1 = ex.Message
        End Try

        Call DoThis("Cancel")

        Call AssessSampleAssignment()
        'Call AssessSampleAssignmentAnalyte()

        Call ASNum()

        Me.lblWait.Visible = False
        Me.lblWait.Refresh()
        Me.Refresh()

        Cursor.Current = Cursors.Default

        boolCancelButton = False

    End Sub

    Private Sub cmdRemove_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRemove.Click

        Call RemoveRows()

    End Sub

    Sub RemoveRows()

        Dim dgv2 As DataGridView
        Dim row As DataGridViewRow
        Dim intRows As Short
        Dim Count1 As Short
        Dim var1
        Dim dv As System.Data.DataView
        Dim intA As Short
        Dim tbl1 As New System.Data.DataTable
        Dim strS As String

        Dim col1 As New DataColumn
        col1.ColumnName = "ID"
        col1.DataType = System.Type.GetType("System.Int16")
        tbl1.Columns.Add(col1)

        dgv2 = Me.dgvAssignedSamples

        If dgv2.SelectedRows.Count = 0 Then
            Exit Sub
        End If

        dv = dgv2.DataSource
        intRows = dv.Count

        Dim arr1(intRows) As Object


        Count1 = 0
        For Each row In dgv2.SelectedRows
            Count1 = Count1 + 1
            var1 = row.Index
            Dim r As DataRow = tbl1.NewRow
            r.BeginEdit()
            r("ID") = var1
            r.EndEdit()
            tbl1.Rows.Add(r)
        Next
        intA = Count1

        'sort tbl1
        Dim dv1 As System.Data.DataView = New DataView(tbl1)
        strS = "ID DESC"
        dv1.Sort = strS
        intA = dv1.Count

        Dim boolContT As Boolean = boolCont
        boolCont = False

        dv.AllowDelete = True

        For Count1 = 0 To intA - 1
            var1 = dv1(Count1).Item("ID")
            dv(var1).Delete()
        Next

        dv.AllowDelete = False

        Call ASNum()

        boolCont = boolContT

    End Sub

    Sub FilterForAnalyte(ByRef tbl As System.Data.DataTable)

        Dim strF As String
        Dim dgv As DataGridView
        'Dim tbl As System.Data.DataTable
        Dim int1 As Short
        Dim intRow As Short
        Dim var1, var2
        Dim strS As String

        'tbl = Me.tblAnalysisResults

        strF = ""
        If Me.rbFilterForAnalyteYes.Checked Then
            If Me.dgvAnalytes.CurrentRow Is Nothing Then
                intRow = 0
            Else
                intRow = Me.dgvAnalytes.CurrentRow.Index
            End If

            var1 = Me.dgvAnalytes.Item(2, intRow).Value
            var2 = Me.dgvAnalytes.Item(11, intRow).Value
            If boolUseGroups Then

            Else
                strF = "ANALYTEINDEX = " & var1 & " AND MASTERASSAYID = " & var2 ' & " AND ASSAYLEVEL = " & var2 & " AND NOMCONC = " & var3

            End If
            'dv.RowFilter = strF
        Else
        End If
        strS = "MASTERASSAYID ASC, ANALYTEINDEX ASC"

        ''debug
        'Dim Count1 As Short
        'Console.WriteLine("Start OutStudy")
        'For Count1 = 0 To tbl.Columns.Count - 1
        '    Console.WriteLine(tbl.Columns(Count1).ColumnName)
        'Next
        'Console.WriteLine("End OutStudy")

        ''debug
        'Console.WriteLine("Start tblAnalysisResultsHomeOutStudy")
        'For Count1 = 0 To tblAnalysisResultsHomeOutStudy.Columns.Count - 1
        '    Console.WriteLine(tblAnalysisResultsHomeOutStudy.Columns(Count1).ColumnName)
        'Next
        'Console.WriteLine("End tblAnalysisResultsHomeOutStudy")

        tbl.CaseSensitive = True
        Dim dv As System.Data.DataView = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)
        dv.AllowDelete = False
        dv.AllowNew = False

        int1 = tbl.Rows.Count 'for debugging

        'dgv.AllowUserToResizeColumns = True
        'dgv.AllowUserToResizeRows = True
        'dgv.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)

        dgv = Me.dgvAnalyticalRuns
        dgv.DataSource = dv

        Call OrderColumns(dgv, False)

    End Sub

    Private Sub rbFilterForAnalyteYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFilterForAnalyteYes.CheckedChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Call FilterForAnalyte(Me.tblAnalysisResults)

    End Sub

    Private Sub cmdEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdEdit.Click

        Dim str1 As String
        If Me.dgvTables.RowCount = 0 Then
            str1 = "There are no tables to edit."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid event...")
            Exit Sub
        End If

        Call DoThis("Edit")

    End Sub

    Sub GetAnalIS(ByVal dgv As DataGridView, ByVal strIS As String)

        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim intRows As Short
        Dim var1, var2
        Dim bool As Boolean

        'Legend:
        'Dim arrAnalytes(14, 51) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
        ''10=UseIntStd, 11=IntStd, 12=MasterAssayID,13=IsCoadministeredCmpd,14=Original Analyte Description

        AnalIS(1) = "NA" 'AnalyteDescription
        AnalIS(2) = 0 'analyteindex
        AnalIS(3) = 0 'masterassayid
        intRows = dgv.RowCount
        bool = True 'debugging
        For Count1 = 0 To intRows - 1
            str1 = NZ(dgv("IntStd", Count1).Value, "")
            If StrComp(str1, strIS, CompareMethod.Text) = 0 Then
                str2 = dgv("AnalyteDescription", Count1).Value
                var1 = dgv("AnalyteIndex", Count1).Value
                var2 = dgv("MasterAssayID", Count1).Value
                AnalIS(1) = str2
                AnalIS(2) = var1
                AnalIS(3) = var2
                bool = False
                Exit For
            End If
        Next

        If bool Then
            var1 = "Hmmm" 'hmm. a problem
        End If

    End Sub

    Private Sub cmdCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCopy.Click

        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim Count3 As Int16
        Dim intRow As Short
        Dim intRows As Short
        Dim intRows1 As Short
        Dim intRows2 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim dgv As DataGridView
        Dim frm As New frmDuplicateAssignment
        Dim bool As Boolean
        Dim tbl As New System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim dv2 As System.Data.DataView
        Dim str1 As String
        Dim var1, var2, var3, var4, var5, var10, var11, var12
        Dim strM As String
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim dgv1 As DataGridView
        Dim sr As DataGridViewSelectedRowCollection
        Dim intTableID As Short
        Dim rowsA() As DataRow
        Dim row As DataGridViewRow
        Dim dgv2 As DataGridView

        Dim tblARH As System.Data.DataTable
        Dim rowsARH() As DataRow
        Dim rowsStudies() As DataRow

        Dim rowsS() As DataRow 'Source
        Dim rowsD() As DataRow 'Destination

        Dim rowsNew() As DataRow
        Dim intT1 As Short
        Dim intT2 As Short
        Dim boolA As Boolean
        Dim maxID
        Dim maxID1
        'Dim tblMaxID As System.Data.DataTable
        'Dim drowsmaxid() As DataRow
        Dim intIS As Short
        Dim intIS1 As Short
        Dim idT As Int64
        Dim intWSID As Int64
        Dim intSID As Int64
        Dim intRowsARH As Short
        Dim boolU As Boolean = False
        Dim boolGoU As Boolean = True
        Dim intGroup As Short

        Dim intAnalyteID As Int64 ' = NZ(dvASRow("ANALYTEID"), -1) ' dgvA("ANALYTEID", dgvA.CurrentRow.Index).Value
        Dim intAID As Int64 ' = NZ(dvASRow("ASSAYID"), -1)
        Dim intAIn As Int64 ' = NZ(dvASRow("ANALYTEINDEX"), -1)
        Dim intAL As Int16 ' = NZ(dvASRow("ASSAYLEVEL"), -1)
        Dim intMAID As Int64 ' = NZ(dvASRow("MASTERASSAYID"), -1)
        Dim dblNominalConcentration As Double
        Dim strRSK1 As String
        Dim strRSK As String
        Dim strHelper1 As String
        Dim intRunID As Int16
        Dim strKnownType As String

        Dim boolIsIS As Boolean = False 'Is / Isn't Int Std

        Dim var1a, var2a

        If Me.dgvAssignedSamples.RowCount = 0 Then
            MsgBox("There are no samples to replicate.", MsgBoxStyle.Information, "No No...")
            Exit Sub
        End If

        tblARH = Me.tblAnalysisResults
        'first ensure tblarh has this analyte
        '0=TRUE/FALSE,1=analytedescription, 2=analyteindex, 3=masterassayid, 4=ANALYTEID, 5=wStudyID,6=StudyName
        'get unique ids in dgvAssignedSamples
        Dim dvU As System.Data.DataView = Me.dgvAssignedSamples.DataSource
        Dim tblU As System.Data.DataTable = dvU.ToTable("a", True, "ID_TBLSTUDIES2")

        'Note: the previous code will work only if user has toggled cbxStudies
        'Solution: don't allow if AssignedSamples has two studies
        'If tblU.Rows.Count > 1 Then
        '    'str1 = "The study '" & boolAnalOK(6, Count1) & "' does not contain the analyte '" & boolAnalOK(1, Count1) & "'." & ChrW(10) & ChrW(10)
        '    str1 = "The assigned samples for this study contain samples from a different study." & ChrW(10) & ChrW(10)
        '    str1 = str1 & "The analyte cannot be replicated with this funtion. The samples for the other analyte(s) must be assigned by hand."
        '    MsgBox(str1, MsgBoxStyle.Critical, "Invalid action...")
        '    Exit Sub

        'End If


        'find maxID for tblReportTable
        str1 = "charTable = 'tblAssignedSamples'"
        If boolGuWuOracle Then
            ta_tblMaxID.Fill(tblMaxID)
        ElseIf boolGuWuAccess Then
            ta_tblMaxIDAcc.Fill(tblMaxID)
        ElseIf boolGuWuSQLServer Then
            ta_tblMaxIDSQLServer.Fill(tblMaxID)
        End If

        'drowsmaxid = tblMaxID.Select(str1)
        'maxID = drowsmaxid(0).Item("numMaxID")
        maxID = GetMaxID("TBLASSIGNEDSAMPLES", 1, False)
        maxID1 = maxID

        intSID = NZ(CLng(Me.txtStudyID.Text), 0)
        intWSID = GetWStudyID(intSID)

        'get distinct items for id_tblStudies
        Dim dvStudies As System.Data.DataView = New DataView(tblAssignedSamples)
        Dim tblDStudies As System.Data.DataTable = dvStudies.ToTable("a", True, "ID_TBLSTUDIES2", "CHARSTUDYNAME2")
        Dim intDRows As Short
        intDRows = tblDStudies.Rows.Count

        dgv = Me.dgvAnalytes
        intRows = dgv.RowCount
        intRow = dgv.CurrentRow.Index
        If intRows = 0 Then
            MsgBox("There are no analytes to replicate to.", MsgBoxStyle.Information, "No No...")
            Exit Sub
        End If

        'get intTableID
        int1 = Me.dgvTables.CurrentRow.Index
        intTableID = Me.dgvTables("ID_TBLCONFIGREPORTTABLES", int1).Value
        idT = Me.dgvTables("ID_TBLREPORTTABLE", int1).Value

        dv2 = Me.dgvAssignedSamples.DataSource
        Dim tbl2 As System.Data.DataTable = dv2.ToTable

        'hide frm rows if selected locally
        'populate frm.dgv
        dv = Me.dgvAnalytes.DataSource
        Dim tbl1 As System.Data.DataTable = dv.ToTable()
        tbl = tbl1.Clone
        For Count1 = 0 To intRows - 1
            If Count1 = intRow Then 'ignore
            Else 'add to table
                Dim rowb As DataRow = tbl.NewRow
                str1 = dgv("AnalyteDescription", Count1).Value
                For Count2 = 0 To dgv.Columns.Count - 1
                    var1 = dgv(Count2, Count1).Value
                    rowb.Item(Count2) = var1
                Next
                tbl.Rows.Add(rowb)
            End If
        Next
        'dv1 = tbl.DefaultView
        Dim dgvF As DataGridView = frm.dgvAnalytes
        dv1 = New DataView(tbl)
        dgvF.DataSource = dv1
        intRows1 = dgv.Columns.Count
        For Count1 = 0 To intRows1 - 1
            dgvF.Columns.Item(Count1).Visible = False
        Next
        dgvF.DefaultCellStyle.Font = New Font(frm.dgvAnalytes.Font, Font.Size = 10)
        dgvF.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvF.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
        dgvF.AllowUserToResizeRows = False
        dgvF.AllowUserToResizeColumns = True

        dgvF.Columns.Item("AnalyteDescription").Visible = True
        'dgvF.Columns.Item("AnalyteID").Visible = True
        'dgvF.Columns.Item("AnalyteIndex").Visible = True
        'dgvF.Columns.Item("MasterAssayID").Visible = True
        frm.ShowDialog()
        Me.Refresh()
        If frm.boolCancel Then
            GoTo end1
        End If

        Cursor.Current = Cursors.WaitCursor

        'evaluate some stuff
        'if data already exists for chosen replicates, send error msg
        dgv2 = Me.dgvAnalyticalRuns
        dgv1 = frm.dgvAnalytes
        sr = dgv1.SelectedRows
        intRows2 = sr.Count

        Dim intA1 As Short

        Dim dt As Date = Now

        'strF is for rowsa=tblassignedsamples.select
        'strF1 is for tblAnalysisResults
        Dim varID
        For Count1 = 0 To intRows2 - 1

            boolIsIS = False

            intA1 = sr.Item(Count1).Index

            var3 = dgv1("AnalyteDescription", intA1).Value
            Call GetAnalIS(Me.dgvAnalytes, var3) 'call this to set analis

            var4 = NZ(dgv1("IsIntStd", intA1).Value, "No")


            'generate filter string
            var1 = dgv1("AnalyteIndex", intA1).Value
            var1 = NZ(var1, AnalIS(2))
            var2 = dgv1("MasterAssayID", intA1).Value
            var2 = NZ(var2, AnalIS(3))
            varID = dgv1("AnalyteID", intA1).Value

            Dim intWSID2 As Int32

            For Count3 = 0 To intDRows - 1

                intSID = tblDStudies.Rows(Count3).Item("ID_TBLSTUDIES2")
                intWSID2 = GetWStudyID(intSID)

                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLSTUDIES2 = " & intSID & " AND ID_TBLREPORTTABLE = " & idT & " AND "
                strF = strF & "ID_TBLCONFIGREPORTTABLES = " & intTableID & " AND "
                If StrComp(var4, "Yes", CompareMethod.Text) = 0 Then

                    boolIsIS = True

                    'Don't use analyteindex because it may be different if data is from a different study
                    intIS1 = -1

                    'get analytid from assignedsamples
                    varID = Me.dgvAssignedSamples("ANALYTEID", 0).Value

                    'strF = strF & "INTERNALSTANDARDNAME = '" & var3 & "' AND "
                    strF = strF & "INTERNALSTDNAME = '" & CleanText(CStr(var3)) & "' AND "
                    strF = strF & "CHARANALYTE = '" & CleanText(CStr(var3)) & "' AND "
                    intGroup = -1
                    'strF = strF1
                Else
                    intIS1 = 0
                    'strF1 = strF & "CHARANALYTE = '" & CleanText(cstr(var3)) & "' AND "
                    strF = strF & "CHARANALYTE = '" & CleanText(CStr(var3)) & "' AND "
                    'strF1 = "ANALYTEINDEX = " & var1 & " AND "
                    'strF1 = strF1 & "MASTERASSAYID = " & var2
                    intGroup = dgv1("INTGROUP", intA1).Value
                End If
                'strF1 = strF1 & "BOOLINTSTD = " & intIS1
                strF = strF & "BOOLINTSTD = " & intIS1
                'strF1 = "CHARANALYTE = '" & CleanText(cstr(var3)) & "'"
                strF = strF & " AND ANALYTEID = " & varID

                Erase rowsA
                Try
                    rowsA = tblAssignedSamples.Select(strF)
                Catch ex As Exception
                    var1 = ex.Message
                    var1 = var1
                End Try

                '20190228 LEE:
                'Aack! this logic is backwards! should be <>, not >
                'If rowsA.Length > 0 Then 'data already exists
                If rowsA.Length <> 0 Then 'data already exists
                    str1 = "Data already exists for " & var3 & "." & Chr(10) & Chr(10)
                    str1 = str1 & "First delete Assigned Sample rows associated with this table and " & var3 & ", then re-attempt this action."
                    MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
                    GoTo end1
                Else 'add  data to table

                    int1 = Me.dgvAssignedSamples.Rows.Count

                    'Legend:
                    'dv2 = Me.dgvAssignedSamples.DataSource
                    'Dim tbl2 As System.Data.DataTable = dv2.ToTable
                    'tblARH = Me.tblAnalysisResults

                    If boolIsIS Then
                        Select Case intTableID
                            Case 13, 14, 15
                                If BOOLISCOMBINELEVELS Then

                                End If
                        End Select
                    End If

                    Dim intFirstLevel As Int32
                    Dim numFirstNomConc As Decimal

                    Dim vFirstLevel
                    Dim vFirstNomConc

                    For Count2 = 0 To int1 - 1

                        strF2 = "ANALYTEID = " & varID
                        strF2 = strF2 & " AND RUNID = " & tbl2.Rows.Item(Count2).Item("RUNID")
                        'strF2 = strF2 & " AND RUNSAMPLESEQUENCENUMBER = " & tbl2.Rows.Item(Count2).Item("RUNSAMPLESEQUENCENUMBER")
                        strF2 = strF2 & " AND RUNSAMPLEORDERNUMBER = " & tbl2.Rows.Item(Count2).Item("RUNSAMPLEORDERNUMBER")
                        strF2 = strF2 & " AND STUDYID = " & intWSID2

                        Erase rowsARH
                        'get this info from filtered rows
                        'tblARH = Me.tblAnalysisResults 'for reference
                        'rowsARH = tblARH.Select(strF2)
                        ''console.writeline(strF2)
                        Try
                            rowsARH = tblARH.Select(strF2)
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try

                        intRowsARH = rowsARH.Length

                        If intRowsARH = 0 Then '
                            var1 = var2
                            GoTo NextCount2
                        End If

                        maxID = maxID + 1
                        Dim rowN As DataRow = tblAssignedSamples.NewRow
                        rowN.BeginEdit()

                        Select Case intTableID
                            Case 13, 14, 15
                                If Count2 = 0 Then
                                    vFirstLevel = NZ(rowsARH(0).Item("ASSAYLEVEL"), DBNull.Value) '
                                    vFirstNomConc = NZ(tbl2.Rows.Item(Count2).Item("NOMCONC"), DBNull.Value) '
                                End If
                        End Select

                        'get this info from current tblAssignedSamples
                        'dv2 = Me.dgvAssignedSamples.DataSource 'for reference
                        'Dim tbl2 as System.Data.DataTable = dv2.ToTable 'for reference
                        rowN.Item("ID_TBLASSIGNEDSAMPLES") = maxID
                        rowN.Item("ID_TBLCONFIGREPORTTABLES") = intTableID
                        rowN.Item("ID_TBLSTUDIES") = id_tblStudies
                        var10 = CLng(Me.txtStudyID.Text) 'DEBUG
                        rowN.Item("ID_TBLSTUDIES2") = intSID ' CLng(Me.txtStudyID.Text)
                        rowN.Item("CHARSTUDYNAME2") = tbl2.Rows.Item(Count2).Item("CHARSTUDYNAME2") 'wWStudyName 'Me.txtStudy.Text
                        rowN.Item("CHARANALYTE") = var3
                        var10 = tbl2.Rows.Item(Count2).Item("CHARSTUDYNAME2")
                        rowN.Item("RUNID") = tbl2.Rows.Item(Count2).Item("RUNID")
                        rowN.Item("ANALYTEINDEX") = var1
                        rowN.Item("MASTERASSAYID") = tbl2.Rows.Item(Count2).Item("MASTERASSAYID") ' var2
                        rowN.Item("ANALYTEID") = varID
                        rowN.Item("RUNSAMPLESEQUENCENUMBER") = tbl2.Rows.Item(Count2).Item("RUNSAMPLESEQUENCENUMBER")
                        rowN.Item("RUNSAMPLEORDERNUMBER") = tbl2.Rows.Item(Count2).Item("RUNSAMPLEORDERNUMBER")
                        rowN.Item("UPSIZE_TS") = dt 'DBNull.Value
                        rowN.Item("BOOLINTSTD") = intIS1
                        var1a = tbl2.Rows.Item(Count2).Item("CHARHELPER1") 'debug
                        rowN.Item("CHARHELPER1") = tbl2.Rows.Item(Count2).Item("CHARHELPER1")
                        var2a = tbl2.Rows.Item(Count2).Item("CHARHELPER2") 'debug
                        rowN.Item("CHARHELPER2") = tbl2.Rows.Item(Count2).Item("CHARHELPER2")

                        rowN.Item("BOOLEXCLSAMPLE") = 0 ' tbl2.Rows.Item(Count2).Item("BOOLEXCLSAMPLE")
                        rowN.Item("BOOLUSEGUWUACCCRIT") = tbl2.Rows.Item(Count2).Item("BOOLUSEGUWUACCCRIT")
                        rowN.Item("NUMMINACCCRIT") = tbl2.Rows.Item(Count2).Item("NUMMINACCCRIT")
                        rowN.Item("NUMMAXACCCRIT") = tbl2.Rows.Item(Count2).Item("NUMMAXACCCRIT")

                        rowN.Item("STUDYID") = rowsARH(0).Item("STUDYID")
                        rowN.Item("STUDYNAME") = rowsARH(0).Item("STUDYNAME")
                        rowN.Item("RUNTYPEID") = rowsARH(0).Item("RUNTYPEID")
                        var1a = rowsARH(0).Item("ASSAYLEVEL") 'debug
                        Select Case intTableID
                            Case 13, 14, 15
                                If BOOLISCOMBINELEVELS And boolIsIS Then
                                    'intFirstLevel = rowsARH(0).Item("ASSAYLEVEL") '
                                    'numFirstNomConc = tbl2.Rows.Item(Count2).Item("NOMCONC") '
                                    rowN.Item("ASSAYLEVEL") = vFirstLevel 'intFirstLevel
                                Else
                                    rowN.Item("ASSAYLEVEL") = rowsARH(0).Item("ASSAYLEVEL")
                                End If
                            Case Else
                                rowN.Item("ASSAYLEVEL") = rowsARH(0).Item("ASSAYLEVEL")
                        End Select

                        rowN.Item("ELIMINATEDFLAG") = rowsARH(0).Item("ELIMINATEDFLAG")
                        rowN.Item("SAMPLENAME") = rowsARH(0).Item("SAMPLENAME")
                        '20171119 LEE
                        Try
                            var11 = NZ(rowsARH(0).Item("ALIQUOTFACTOR"), 1) 'debug
                        Catch ex As Exception
                            var1 = var1 'debug
                        End Try

                        If var11 < 1 Then
                            var12 = Math.Round(var11, intDFDec)
                        Else
                            var12 = var11
                        End If

                        'rowN is assignedsamples
                        rowN.Item("ALIQUOTFACTOR") = var12 ' rowsARH(0).Item("ALIQUOTFACTOR")
                        rowN.Item("RUNSAMPLEKIND") = rowsARH(0).Item("RUNSAMPLEKIND")
                        rowN.Item("ASSAYID") = rowsARH(0).Item("ASSAYID")
                        rowN.Item("CONCENTRATION") = rowsARH(0).Item("CONCENTRATION")
                        rowN.Item("RUNANALYTEREGRESSIONSTATUS") = rowsARH(0).Item("RUNANALYTEREGRESSIONSTATUS")
                        rowN.Item("ANALYTEHEIGHT") = rowsARH(0).Item("ANALYTEHEIGHT")
                        rowN.Item("ANALYTEAREA") = rowsARH(0).Item("ANALYTEAREA")
                        rowN.Item("INTERNALSTANDARDHEIGHT") = rowsARH(0).Item("INTERNALSTANDARDHEIGHT")
                        rowN.Item("INTERNALSTANDARDAREA") = rowsARH(0).Item("INTERNALSTANDARDAREA")
                        rowN.Item("INTERNALSTDNAME") = rowsARH(0).Item("INTERNALSTDNAME")
                        rowN.Item("ANALYTEID") = rowsARH(0).Item("ANALYTEID")

                        rowN.Item("BOOLOUTLIER") = 0
                        rowN.Item("BOOLINCURRED") = 0

                        rowN.Item("ID_TBLREPORTTABLE") = idT
                        Try
                            rowN.Item("SAMPLETYPEID") = rowsARH(0).Item("SAMPLETYPEID")
                        Catch ex As Exception
                            var1 = ex.Message
                            var1 = var1
                        End Try

                        rowN.Item("INTANALYTEID") = rowsARH(0).Item("ANALYTEID")
                        rowN.Item("INTGROUP") = intGroup

                        Select Case intTableID
                            Case 13, 14, 15
                                var1a = tbl2.Rows.Item(Count2).Item("NOMCONC") 'debug
                                'rowN.Item("NOMCONC") = tbl2.Rows.Item(Count2).Item("NOMCONC")
                                If BOOLISCOMBINELEVELS And boolIsIS Then
                                    'intFirstLevel = rowsARH(0).Item("ASSAYLEVEL") '
                                    'numFirstNomConc = tbl2.Rows.Item(Count2).Item("NOMCONC") '
                                    rowN.Item("NOMCONC") = vFirstNomConc 'numFirstNomConc
                                Else
                                    rowN.Item("NOMCONC") = tbl2.Rows.Item(Count2).Item("NOMCONC")
                                End If

                            Case 23
                                rowN.Item("NOMCONC") = tbl2.Rows.Item(Count2).Item("NOMCONC")

                            Case 17, 22, 34, 35, 37

                                'use function to get nominal concentration
                                '*****
                                intAnalyteID = NZ(rowsARH(0).Item("ANALYTEID"), -1) ' dgvA("ANALYTEID", dgvA.CurrentRow.Index).Value
                                intAID = NZ(rowsARH(0).Item("ASSAYID"), -1)
                                intAIn = NZ(rowsARH(0).Item("ANALYTEINDEX"), -1)
                                intAL = NZ(rowsARH(0).Item("ASSAYLEVEL"), -1)
                                intMAID = NZ(rowsARH(0).Item("MASTERASSAYID"), -1)
                                strRSK1 = NZ(rowsARH(0).Item("RUNSAMPLEKIND"), "NA")
                                If (intAID = -1) Or (intAIn = -1) Or (intAL = -1) Then
                                Else
                                    dblNominalConcentration = FindNomConc(tblConcLevelsForAssayIDs, intAID, strRSK1, intAIn, intAL, intMAID)
                                    rowN.Item("NOMCONC") = dblNominalConcentration
                                End If

                            Case Else

                                'use function to get nominal concentration
                                '*****
                                intAnalyteID = NZ(rowsARH(0).Item("ANALYTEID"), -1) ' dgvA("ANALYTEID", dgvA.CurrentRow.Index).Value
                                intAID = NZ(rowsARH(0).Item("ASSAYID"), -1)
                                intAIn = NZ(rowsARH(0).Item("ANALYTEINDEX"), -1)
                                intAL = NZ(rowsARH(0).Item("ASSAYLEVEL"), -1)
                                intMAID = NZ(rowsARH(0).Item("MASTERASSAYID"), -1)
                                strRSK1 = NZ(rowsARH(0).Item("RUNSAMPLEKIND"), "NA")
                                If (intAID = -1) Or (intAIn = -1) Or (intAL = -1) Then
                                Else
                                    dblNominalConcentration = FindNomConc(tblConcLevelsForAssayIDs, intAID, strRSK1, intAIn, intAL, intMAID)
                                    rowN.Item("NOMCONC") = dblNominalConcentration
                                End If

                                'use function to get QC/Calibr level
                                intRunID = rowsARH(0).Item("RUNID")
                                Dim strLabel As String
                                'HELPER1
                                '20151106 LEE:
                                'Not all QC's have 'QC' as sample type
                                'Instead, filter out Standards
                                If StrComp(strRSK1, "STANDARD", CompareMethod.Text) = 0 Then
                                Else
                                    strLabel = FindLabelHelper1(intAID, strRSK1, intAIn, intAL, intAnalyteID, intRunID, intAL)
                                    rowN.Item("CHARHELPER1") = strLabel
                                End If
                        End Select


                        '*****
                        'rowN.Item("NOMCONC") = tbl2.Rows.Item(Count2).Item("NOMCONC")

                        rowN.EndEdit()
                        tblAssignedSamples.Rows.Add(rowN)

NextCount2:
                    Next Count2

                End If

            Next Count3

        Next Count1

        If maxID = maxID1 Then
        Else
            Call PutMaxID("TBLASSIGNEDSAMPLES", maxID)

        End If

        Cursor.Current = Cursors.WaitCursor

        Call AssessSampleAssignment()

        str1 = "Data replication action completed." ' & ChrW(10) & ChrW(10)
        'str1 = str1 & "IMPORTANT!! Remember to inspect the assigned Nominal Concentrations of the replicated data as they may be different for individual analytes."
        MsgBox(str1, MsgBoxStyle.Information, "Action completed...")

end1:
        Me.Refresh()
        frm.Dispose()
        tbl2.Dispose()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmdExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        Call CloseWindow()

    End Sub

    Sub CloseWindow()

        '20180822 LEE:

        Dim var1

        Try
            frmAnalyticalRunSummary.Close()
        Catch ex As Exception
            var1 = var1
        End Try

        boolCancel = True
        boolX = True

        'determine if external study data has been added
        Dim strF As String = "ID_TBLSTUDIES = " & wStudyID
        Dim dv1 As New DataView(tblAssignedSamples, strF, "", DataViewRowState.ModifiedOriginal)
        Dim dv2 As New DataView(tblAssignedSamples, strF, "", DataViewRowState.Added)
        Dim tbl1 As DataTable = dv2.ToTable("b", True, "ID_TBLSTUDIES", "ID_TBLSTUDIES2") 'this should account for matrix and calibrlevels
        Dim tbl2 As DataTable = dv1.ToTable("a", True, "ID_TBLSTUDIES", "ID_TBLSTUDIES2") 'this should account for matrix and calibrlevels

        Dim intRows1 As Int32 = tbl1.Rows.Count
        Dim intRows2 As Int32 = tbl2.Rows.Count

        Dim boolDo As Boolean = False
        Dim boolT As Boolean = False

        Dim Count1 As Int16
        Dim Count2 As Int16

        Dim int1 As Int32
        Dim int2 As Int32

        Dim rows1() As DataRow
        Dim rows2() As DataRow

        For Count1 = 0 To tbl2.Rows.Count - 1
            int1 = tbl2.Rows(Count1).Item("ID_TLBSTUDIES2")
            strF = "ID_TBLSTUDIES2 = " & int1
            rows1 = tbl1.Select(strF)
            If rows1.Length = 0 Then
                boolDo = True
                Exit For
            End If
        Next

        If boolDo Then

            'must reload tblAnalysisResultsHome
            Dim cn As New ADODB.Connection
            cn.Open(constrCur)
            Call FillAnalysisResultsTable(cn)
            cn.Close()
            cn = Nothing

        End If


        Me.Visible = False

    End Sub


    Sub Helper1Click()

        Dim dgv As DataGridView
        Dim sr As DataGridViewSelectedRowCollection
        Dim dgv1 As DataGridView
        Dim sr1 As DataGridViewSelectedRowCollection
        Dim intRow As Short
        Dim intRows As Short
        Dim strX As String
        Dim Count1 As Short
        Dim int1 As Short
        Dim dv As System.Data.DataView

        dgv = Me.dgvAssignedSamples
        sr = dgv.SelectedRows
        If dgv.RowCount = 0 Or sr.Count = 0 Then
            Exit Sub
        End If

        dgv1 = Me.dgvHelper1
        sr1 = dgv.SelectedRows
        If dgv1.RowCount = 0 Or sr1.Count = 0 Then
            Exit Sub
        End If

        intRow = dgv1.CurrentRow.Index
        strX = dgv1("CHARHELPER", intRow).Value
        dv = dgv.DataSource

        intRows = sr.Count
        For Count1 = 0 To intRows - 1
            int1 = sr.Item(Count1).Index
            dv(int1).BeginEdit()
            dv(int1).Item("CHARHELPER1") = strX
            dv(int1).EndEdit()
        Next

    End Sub

    Sub Helper2Click()
        Dim dgv As DataGridView
        Dim sr As DataGridViewSelectedRowCollection
        Dim dgv1 As DataGridView
        Dim sr1 As DataGridViewSelectedRowCollection
        Dim intRow As Short
        Dim intRows As Short
        Dim strX As String
        Dim Count1 As Short
        Dim int1 As Short
        Dim dv As System.Data.DataView


        dgv = Me.dgvAssignedSamples
        sr = dgv.SelectedRows
        If dgv.RowCount = 0 Or sr.Count = 0 Then
            Exit Sub
        End If

        If Me.dgvHelper2.Visible Then
            dgv1 = Me.dgvHelper2
            sr1 = dgv.SelectedRows
            If dgv1.RowCount = 0 Or sr1.Count = 0 Then
                Exit Sub
            End If
            intRow = dgv1.CurrentRow.Index
            strX = dgv1("CHARHELPER", intRow).Value
            'If StrComp(strX, "[Clear]", CompareMethod.Text) = 0 Then
            '    strX = DBNull.Value
            'End If
        ElseIf Me.txtHelper2.Visible Then
            strX = Me.txtHelper2.Text
        End If

        dv = dgv.DataSource
        intRows = sr.Count
        For Count1 = 0 To intRows - 1
            int1 = sr.Item(Count1).Index
            dv(int1).BeginEdit()
            If StrComp(strX, "[Clear]", CompareMethod.Text) = 0 Then
                dv(int1).Item("CHARHELPER2") = DBNull.Value
            Else
                dv(int1).Item("CHARHELPER2") = strX
            End If
            dv(int1).EndEdit()
        Next


    End Sub

    Sub NomConcClick()

        Dim dgv As DataGridView
        Dim sr As DataGridViewSelectedRowCollection
        Dim dgv1 As DataGridView
        Dim sr1 As DataGridViewSelectedRowCollection
        Dim intRow As Short
        Dim intRows As Short
        Dim strX As String
        Dim Count1 As Short
        Dim int1 As Short
        Dim dv As System.Data.DataView

        dgv = Me.dgvAssignedSamples
        sr = dgv.SelectedRows
        If dgv.RowCount = 0 Or sr.Count = 0 Then
            Exit Sub
        End If

        dgv1 = Me.dgvNomConc
        sr1 = dgv.SelectedRows
        If dgv1.RowCount = 0 Or sr1.Count = 0 Then
            Exit Sub
        End If

        intRow = dgv1.CurrentRow.Index
        strX = dgv1("DEC", intRow).Value
        dv = dgv.DataSource

        intRows = sr.Count
        For Count1 = 0 To intRows - 1
            int1 = sr.Item(Count1).Index
            dv(int1).BeginEdit()
            dv(int1).Item("NOMCONC") = strX
            dv(int1).EndEdit()
        Next

    End Sub



    Private Sub cbxSortAssigneSamples_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxSortAssigneSamples.SelectedIndexChanged

        Call sortAssignedSamples()

    End Sub

    Private Sub sortAssignedSamples()

        Dim dgv As DataGridView
        Dim strS As String
        Dim str1 As String
        Dim dv As System.Data.DataView

        If boolFormLoad Then
            Exit Sub
        End If

        dgv = Me.dgvAssignedSamples
        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        'strS = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        str1 = Me.cbxSortAssigneSamples.Text
        If StrComp(str1, "Original", CompareMethod.Text) = 0 Then
            'strS = "RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
            strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        ElseIf StrComp(str1, "Level", CompareMethod.Text) = 0 Then
            'strS = "ASSAYLEVEL ASC, NOMCONC ASC, RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
            'strS = "ASSAYLEVEL ASC, RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
            strS = "ASSAYLEVEL ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC" 'id_tblStudies ASC, id_tblStudies2 ASC, RUNSAMPLESEQUENCENUMBER ASC"
        End If

        dv = dgv.DataSource
        dv.Sort = strS
        'dgv.DataSource = dv

    End Sub

    Sub FindSamples()

        If boolFormLoad Then
            Exit Sub
        End If

        If boolCont Then
        Else
            Exit Sub
        End If

        If boolAutoAssign Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim intRow As Short
        Dim str1 As String
        Dim strF As String
        Dim dv1 As System.Data.DataView
        Dim intFindRow As Short
        'Dim arrF(3) As String
        '20180326 LEE:
        'dgv Sort has only two variables
        'see InitializeAssignedSamples, strS = ...
        Dim arrF(1) As String
        Dim var1, var2, var3, var4, var5

        dgv = Me.dgvAssignedSamples
        dgv1 = Me.dgvAnalyticalRuns
        dv1 = dgv1.DataSource

        If dgv.RowCount = 0 Then
            Exit Sub
        End If
        If dgv.CurrentRow Is Nothing Then
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index

        'arrF(0) = dgv("CHARANALYTE", intRow).Value
        'Try
        '    arrF(1) = NZ(dgv("STUDYNAME", intRow).Value, "") ' NZ(dgv("CHARSTUDYNAME2", intRow).Value, "")
        'Catch ex As Exception
        '    var1 = ex.Message
        'End Try

        'arrF(2) = dgv("RUNID", intRow).Value
        ''arrF(3) = dgv("RUNSAMPLESEQUENCENUMBER", intRow).Value
        'arrF(3) = NZ(dgv("RUNSAMPLEORDERNUMBER", intRow).Value, 0)

        'var5 = dv1.Sort 'debug

        'var1 = arrF(0)
        'var2 = arrF(1)
        'var3 = arrF(2)
        'var4 = arrF(3)

        Try
            '20180326 LEE:
            'dgv Sort has only two variables
            'see InitializeAssignedSamples, strS = ...
            arrF(0) = dgv("RUNID", intRow).Value
            'arrF(3) = dgv("RUNSAMPLESEQUENCENUMBER", intRow).Value
            arrF(1) = NZ(dgv("RUNSAMPLEORDERNUMBER", intRow).Value, 0)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        var1 = arrF(0)
        var2 = arrF(1)

        'Console.WriteLine("CHARANALYTE: " & var1)
        'Console.WriteLine("STUDYNAME: " & var2)
        'Console.WriteLine("RUNID: " & var3)
        'Console.WriteLine("RUNSAMPLEORDERNUMBER: " & var4)


        Try
            intFindRow = dv1.Find(arrF)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


        If intFindRow = -1 Then 'ignore
        Else
            If intFindRow - 3 < 0 Then
                dgv1.FirstDisplayedScrollingRowIndex = intFindRow
            Else
                dgv1.FirstDisplayedScrollingRowIndex = intFindRow - 3
            End If
            dgv1.CurrentCell = dgv1.Rows.Item(intFindRow).Cells("CHARANALYTE")

        End If

    End Sub


    Private Sub dgvAssignedSamples_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAssignedSamples.CellContentClick

    End Sub

    Private Sub dgvAssignedSamples_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvAssignedSamples.CellValidating

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim intCol As Short
        Dim strCol As String
        Dim var1, var2
        Dim boolC As Boolean = True

        If boolAutoAssign Then
            Exit Sub
        End If

        dgv = Me.dgvAssignedSamples

        intRow = dgv.CurrentRow.Index
        intCol = dgv.CurrentCell.ColumnIndex

        strCol = dgv.Columns(intCol).Name
        If StrComp(strCol, "NUMMINACCCRIT", CompareMethod.Text) = 0 Or StrComp(strCol, "NUMMAXACCCRIT", CompareMethod.Text) = 0 Then

            boolC = True
            var1 = e.FormattedValue

            If IsDBNull(var1) Then
                boolC = False
                GoTo end1
            End If

            If Len(var1) = 0 Then
                boolC = False
                GoTo end1
            End If

            If IsNumeric(var1) Then
            Else
                GoTo end1
            End If

            If var1 >= 0 Then
            Else
                GoTo end1
            End If

            boolC = False

end1:
            If boolC Then
                e.Cancel = True
                Dim strM As String
                strM = "Entry must be number >0" & ChrW(10)
                strM = strM & "Note that the (-) sign for the 'Neg.' column will be automatically accounted for during table generation"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
            End If

        End If


    End Sub


    Private Sub dgvAssignedSamples_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAssignedSamples.Click

        Call FindSamples()

    End Sub

    Private Sub dgvAssignedSamples_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAssignedSamples.MouseEnter
        'Me.dgvAssignedSamples.Focus()

    End Sub

    Private Sub dgvAssignedSamples_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAssignedSamples.SelectionChanged

        Call FindSamples()

    End Sub

    Private Sub dgvAnalyticalRuns_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAnalyticalRuns.CellDoubleClick


        Try
            Call AddRows()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgvAnalyticalRuns_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgvAnalyticalRuns.KeyDown
        'If e.KeyCode = Keys.Control AndAlso (e.Alt OrElse e.Control OrElse e.Shift) Then

        If e.Control And e.KeyCode = Keys.C Then
            Call fgKeyDown(Me.dgvAnalyticalRuns)
        End If


    End Sub

    Private Sub dgvAnalyticalRuns_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAnalyticalRuns.MouseEnter
        'Me.dgvAnalyticalRuns.Focus()

    End Sub

    Private Sub cbxFilterRunID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxFilterRunID.SelectedIndexChanged

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        boolFromFilter = True

        Cursor.Current = Cursors.WaitCursor
        Call FillAnalyticalRuns("")
        Cursor.Current = Cursors.Default

        Call CountSamples()

        boolFromFilter = False

        'set focus on dgvAnalRuns
        'users don't remember that the dropdownbox is the still selected
        'they use the scroll wheel thinking that they are scrolling dgvAnalyticalRuns
        'but are actually scrolling the dropdownbox
        Me.dgvAnalyticalRuns.Focus()

    End Sub

    Private Sub txtASNum_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtASNum.Click
        'Me.dgvAssignedSamples.Focus()
    End Sub

    Private Sub cmdClearFilters_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClearFilters.Click

        Call ClearFilters()

    End Sub

    Sub ClearFilters()

        'Clear Filters
        Cursor.Current = Cursors.WaitCursor

        'Me.cbxFilterDilFactor.Text = "[None]"
        'Me.cbxFilterRunID.Text = "[None]"
        'Me.cbxFilterSampleType.Text = "[None]"
        'Me.txtFilterSamples.Text = ""
        'Me.cbxAccStatus.SelectedIndex = 0

        '20160224 LEE: added sub to do this because gets called elsewhere as well
        Call ClearFilters(True)

        boolFromFilter = True
        Call FillAnalyticalRuns("")
        Cursor.Current = Cursors.Default
        boolFromFilter = False
        Call CountSamples()

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cbxFilterSampleType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxFilterSampleType.SelectedIndexChanged
        If boolFormLoad Then
            Exit Sub
        End If

        'NOTE: Leave the bool filters as RI
        boolFromFilter = True

        Cursor.Current = Cursors.WaitCursor
        Call FillAnalyticalRuns("")
        Cursor.Current = Cursors.Default

        Call CountSamples()

        boolFromFilter = False

        'set focus on dgvAnalRuns
        'users don't remember that the dropdownbox is the still selected
        'they use the scroll wheel thinking that they are scrolling dgvAnalyticalRuns
        'but are actually scrolling the dropdownbox
        Me.dgvAnalyticalRuns.Focus()

    End Sub

    Private Sub cbxFilterDilFactor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxFilterDilFactor.SelectedIndexChanged
        If boolFormLoad Then
            Exit Sub
        End If

        'NOTE: Leave the bool filters as RI
        boolFromFilter = True

        Cursor.Current = Cursors.WaitCursor
        Call FillAnalyticalRuns("")
        Cursor.Current = Cursors.Default

        Call CountSamples()

        boolFromFilter = True

        'set focus on dgvAnalRuns
        'users don't remember that the dropdownbox is the still selected
        'they use the scroll wheel thinking that they are scrolling dgvAnalyticalRuns
        'but are actually scrolling the dropdownbox
        Me.dgvAnalyticalRuns.Focus()

    End Sub


    Private Sub txtFilterSamples_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFilterSamples.TextChanged

    End Sub

    Private Sub cmdHelper1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdHelper1.Click

        Call Helper1Click()

    End Sub

    Private Sub cmdHelper2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdHelper2.Click

        Call Helper2Click()

    End Sub

    Private Sub cmdNomConc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdNomConc.Click

        Call NomConcClick()

    End Sub

    Private Sub dgvNomConc_DoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvNomConc.CellDoubleClick
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Call NomConcClick()

    End Sub

    Private Sub dgvHelper1_DoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHelper1.CellDoubleClick
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Call Helper1Click()
    End Sub

    Private Sub dgvHelper1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvHelper1.MouseEnter

        'Me.dgvHelper1.Focus()

    End Sub

    Private Sub dgvHelper2_DoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvHelper2.CellDoubleClick
        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Call Helper2Click()
    End Sub

    Private Sub dgvNomConc_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvNomConc.MouseEnter

        'Me.dgvNomConc.Focus()

    End Sub

    Private Sub cmdIncurred_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdIncurred.Click

        'Call GetIncSamples()

        Dim frm As New frmIncSamplesAssignSamples

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strAnal As String
        Dim idRT As Int64

        str1 = Me.dgvAnalytes(0, Me.dgvAnalytes.CurrentRow.Index).Value

        str2 = frm.Text
        str3 = str2 & " for " & str1
        frm.Text = str3

        Dim int1 As Int32
        Dim intRow As Short
        Dim dgv As DataGridView

        Dim strM As String '
        Dim boolE As Boolean = False

        dgv = Me.dgvAnalytes
        Try
            intRow = dgv.CurrentRow.Index

            int1 = dgv("AnalyteIndex", intRow).Value
            frm.gAnalyteIndex = int1

            int1 = dgv("AnalyteID", intRow).Value
            frm.gAnalyteID = int1

            int1 = dgv("MasterAssayID", intRow).Value
            frm.gMasterAssayID = int1



            strAnal = dgv("AnalyteDescription", intRow).Value

            Try
                idRT = Me.dgvTables.Rows.Item(dgvTables.CurrentRow.Index).Cells("ID_TBLREPORTTABLE").Value()
                frm.gidRT = int1

                Call frm.CreateTableISRAS()

                'now load tlbISRAssSamples
                Call LoadISRSamples(frm.tblISRAssSamples, frm.gAnalyteID, frm.gAnalyteIndex, frm.gMasterAssayID, strAnal, idRT)

                Call frm.FormLoad() 'must call this before ConfigdgvIncSamplesOrig

                Call frm.ConfigdgvIncSamplesOrig()

                Call frm.LoadISRSource()

                frm.ShowDialog()

                If frm.boolCancel Then
                Else
                    Try

                        Call AssignISRSamples(frm.tblISRAssSamples, frm.gAnalyteID, frm.gAnalyteIndex, frm.gMasterAssayID, strAnal, idRT)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try

                End If

                frm.Dispose()
            Catch ex As Exception
                strM = "A Table must be selected." & ChrW(10) & ex.Message
                boolE = True
            End Try

        Catch ex As Exception

            strM = "An Analyte must be selected."
            boolE = True

        End Try

        If boolE Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        End If

    End Sub

    Sub LoadISRSamples(ByVal tblISR As System.Data.DataTable, ByVal gAnalyteID As Int64, ByVal gAnalyteIndex As Int64, ByVal gMasterAssayID As Int64, ByVal strAnalyte As String, ByVal idRT As Int64)

        Dim rows() As DataRow
        Dim strF As String
        Dim Count1 As Short
        Dim Count2 As Int32
        Dim strType As String
        Dim intISR As Short
        Dim bool As Boolean
        Dim var1, var2

        bool = boolCont


        For Count1 = 1 To 2

            Select Case Count1
                Case 1 'Original
                    strType = "O"
                    intISR = 0
                Case 2 'ISR
                    strType = "ISR"
                    intISR = -1
            End Select

            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ANALYTEINDEX = " & gAnalyteIndex & " AND ANALYTEID = " & gAnalyteID & " AND MASTERASSAYID = " & gMasterAssayID & " AND ID_TBLREPORTTABLE = " & idRT & " AND BOOLINCURRED = " & intISR
            Erase rows
            rows = tblAssignedSamples.Select(strF)

            For Count2 = 0 To rows.Length - 1
                boolCont = False 'to disable FindSamples
                Dim dvRow As DataRow = tblISR.NewRow

                dvRow.BeginEdit()

                Try
                    dvRow("SAMPLENAME") = rows(Count2).Item("SAMPLENAME")
                    dvRow("DESIGNSAMPLEID") = rows(Count2).Item("DESIGNSAMPLEID")
                    dvRow("RUNID") = rows(Count2).Item("RUNID")
                    dvRow("RUNSAMPLEORDERNUMBER") = rows(Count2).Item("RUNSAMPLEORDERNUMBER")
                    '20171119 LEE:
                    Try
                        var1 = NZ(rows(Count2).Item("ALIQUOTFACTOR"), 1) 'rows is assignedsamples
                    Catch ex As Exception
                        var1 = var1 'debug
                    End Try

                    If var1 < 1 Then
                        var2 = Math.Round(var1, intDFDec)
                    Else
                        var2 = var1
                    End If
                    Try
                        dvRow("ALIQUOTFACTOR") = var2
                    Catch ex As Exception
                        var1 = var1 'debug
                    End Try

                    '20160509 LEE: If this code gets reinstated, change the CONCENTRATION NZ
                    dvRow("REPORTEDCONC") = RoundToDecimalRAFZ(NZ(rows(Count2).Item("CONCENTRATION"), 0), 3)
                    dvRow("ANALYTEINDEX") = gAnalyteIndex
                    dvRow("ANALYTEID") = gAnalyteID
                    dvRow("MASTERASSAYID") = gMasterAssayID
                    dvRow("CHARTYPE") = strType
                    dvRow("ID_TBLREPORTTABLE") = idRT

                Catch ex As Exception
                    MsgBox(ex.Message)
                    Exit For
                End Try

                dvRow.EndEdit()
                tblISR.Rows.Add(dvRow)
            Next


        Next

        boolCont = bool


    End Sub

    Sub AssignISRSamples(ByVal tblISR As System.Data.DataTable, ByVal gAnalyteID As Int64, ByVal gAnalyteIndex As Int64, ByVal gMasterAssayID As Int64, ByVal strAnalyte As String, ByVal idRT As Int64)

        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim rows() As DataRow
        Dim strF As String
        Dim idDS As Int64
        Dim strSN As String
        Dim strS As String
        Dim maxID As Int64
        Dim maxIDo As Int64
        Dim var1, var2, var3, var10
        Dim idCRT As Int64 = 30
        Dim str1 As String
        Dim str2 As String
        Dim intIncS As Short

        maxID = GetMaxID("TBLASSIGNEDSAMPLES", 1, False)
        maxIDo = maxID

        Dim dvISR As System.Data.DataView
        strS = "DESIGNSAMPLEID ASC, SAMPLENAME ASC"
        strF = "DESIGNSAMPLEID > 0"
        strF = strF & " AND CHARANALYTE = '" & CleanText(strAnalyte) & "'"
        strF = strF & " AND ANALYTEINDEX = " & gAnalyteIndex
        strF = strF & " AND MASTERASSAYID = " & gMasterAssayID
        'strF = strF & " AND ID_TBLSTUDIES = " & id_tblStudies

        Dim dvARuns As System.Data.DataView
        Try
            Dim dv1a As System.Data.DataView = Me.dgvAnalyticalRuns.DataSource
            Dim tbl1a As System.Data.DataTable = dv1a.ToTable
            dvARuns = New DataView(tbl1a, strF, strS, DataViewRowState.CurrentRows)
            var1 = dvARuns.Sort
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        Dim ct1 As Int16 = 0
        Dim strF1 As String
        'first do added
        dvISR = New DataView(tblISR, "DESIGNSAMPLEID > 0", "DESIGNSAMPLEID ASC", DataViewRowState.Added)
        For Count1 = 0 To dvISR.Count - 1

            idDS = dvISR(Count1).Item("DESIGNSAMPLEID")
            strSN = dvISR(Count1).Item("SAMPLENAME")
            strF1 = "DESIGNSAMPLEID = " & idDS & " AND SAMPLENAME = '" & strSN & "'"
            strF1 = strF1 & " AND ID_TBLSTUDIES = " & id_tblStudies
            strF1 = strF1 & " AND ANALYTEINDEX = " & gAnalyteIndex
            strF1 = strF1 & " AND MASTERASSAYID = " & gMasterAssayID
            strF = strF1 & " AND ID_TBLREPORTTABLE = " & idRT
            rows = tblAssignedSamples.Select(strF)
            If rows.Length = 0 Then 'add


                ct1 = ct1 + 1

                maxID = maxID + 1
                boolCont = False 'to disable FindSamples

                str1 = dvISR(Count1).Item("CHARTYPE")
                If StrComp(str1, "O", CompareMethod.Text) = 0 Then
                    intIncS = 0
                Else 'ISR
                    intIncS = -1
                End If

                Dim dvRow As DataRow = tblAssignedSamples.NewRow

                dvRow.BeginEdit()

                dvRow("id_tblAssignedSamples") = maxID
                dvRow("id_tblConfigReportTables") = idCRT ' idConfigRT 'var1
                dvRow("id_tblStudies") = id_tblStudies
                var10 = CLng(Me.txtStudyID.Text) 'debug
                dvRow("id_tblStudies2") = CLng(Me.txtStudyID.Text)
                dvRow("charStudyName2") = Me.cbxStudy.Text
                dvRow("BOOLINTSTD") = 0 ' intIS
                dvRow("ID_TBLREPORTTABLE") = idRT
                dvRow("BOOLOUTLIER") = 0

                dvRow("BOOLEXCLSAMPLE") = 0
                dvRow("BOOLUSEGUWUACCCRIT") = 0

                Dim ctRowSel As Int16
                Dim boolA As Boolean = False

                'find item in dgvAnalyticalRuns
                Dim vals(1) As Object 'CHARANALYTE,STUDYNAME,RUNID,RUNSAMPLEORDERNUMBER
                vals(0) = idDS 'strAnalyte
                vals(1) = strSN 'wWStudyName
                'vals(2) = dvISR(Count1).Item("RUNID")
                'vals(3) = dvISR(Count1).Item("RUNSAMPLEORDERNUMBER")

                Dim i As Int32 = dvARuns.Find(vals)

                'enter column info
                For Count2 = 0 To Me.dgvAnalyticalRuns.Columns.Count - 1
                    str1 = Me.dgvAnalyticalRuns.Columns.Item(Count2).Name
                    var1 = dvARuns(i).Item(str1)
                    boolA = False
                    Select Case str1
                        Case "RUNID"
                            boolA = True
                        Case "ANALYTEINDEX"
                            boolA = True
                        Case "MASTERASSAYID"
                            boolA = True
                            'Case "RUNSAMPLESEQUENCENUMBER"
                            '    boolA = True
                        Case "RUNSAMPLEORDERNUMBER"
                            boolA = True
                    End Select
                    'If boolA Then
                    'Else
                    '    'Me.dgvAssignedSamples.Item(str1, intRows - 1).Value = var1
                    '    dvRow(str1) = var1
                    'End If
                    Try
                        dvRow(str1) = var1
                    Catch ex As Exception
                        MsgBox("dvrow(str1): " & ex.Message)
                    End Try

                Next
                dvRow("CHARANALYTE") = strAnalyte ' idAnalyte 'MUST DO THIS AFTER LOOP!!!
                dvRow("BOOLINCURRED") = intIncS 'MUST DO THIS AFTER LOOP!!!

                dvRow.EndEdit()

                tblAssignedSamples.Rows.Add(dvRow)

            End If

        Next


        If maxID > maxIDo Then
            Call PutMaxID("TBLASSIGNEDSAMPLES", maxID)
        End If

        'now do deleted
        dvISR = New DataView(tblISR, "DESIGNSAMPLEID > 0", "DESIGNSAMPLEID ASC", DataViewRowState.Deleted)

        For Count1 = 0 To dvISR.Count - 1
            idDS = dvISR(Count1).Item("DESIGNSAMPLEID")
            strSN = dvISR(Count1).Item("SAMPLENAME")
            strF1 = "DESIGNSAMPLEID = " & idDS & " AND SAMPLENAME = '" & strSN & "'"
            strF1 = strF1 & " AND ID_TBLSTUDIES = " & id_tblStudies
            strF1 = strF1 & " AND ANALYTEINDEX = " & gAnalyteIndex
            strF1 = strF1 & " AND MASTERASSAYID = " & gMasterAssayID
            strF = strF1 & " AND ID_TBLREPORTTABLE = " & idRT
            rows = tblAssignedSamples.Select(strF)
            If rows.Length = 0 Then
            Else
                For Count2 = 0 To rows.Length - 1
                    rows(Count2).BeginEdit()
                    rows(Count2).Delete()
                    rows(Count2).EndEdit()
                Next
            End If

        Next

        boolCont = True

        Call FindSamples()

        var2 = Format(Me.dgvAssignedSamples.RowCount, "#,##0")
        Me.txtASNum.Text = var2


    End Sub

    Sub GetIncSamples()

        MsgBox("Under construction", MsgBoxStyle.Information, "Under construction...")

    End Sub

    Sub IncSampleVis(ByVal id)

        Dim dv As System.Data.DataView = Me.dgvAssignedSamples.DataSource

        If id = 30 Then
            Me.cmdIncurred.Visible = False 'True
            Me.cmdAddRows.Visible = False
            Me.cmdRemove.Visible = False
            Me.chkDesignSampleID.Checked = True
            dv.Sort = strSortISR
        Else
            Me.cmdIncurred.Visible = False
            Me.cmdAddRows.Visible = True
            Me.cmdRemove.Visible = True
            Me.chkDesignSampleID.Checked = False
            dv.Sort = strSortNonISR
        End If

    End Sub

    Private Sub chkUseWatson_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseWatson.CheckedChanged

        Call FilterHelper1()

    End Sub

    Private Sub chkShowAllNomConc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAllNomConc.CheckedChanged
        Call NomConcFill(False)
    End Sub

    Private Sub cmdShowRunSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowRunSummary.Click

        'MsgBox("Under construction", MsgBoxStyle.Information, "Under construction...")

        Dim frm As New frmAnalyticalRunSummary

        frm.Show(Me)

    End Sub

    Private Sub cmdPrintReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintReport.Click

        'MsgBox("Under construction", MsgBoxStyle.Information, "Under construction...")

        'Exit Sub

        Dim strF As String
        Dim rows() As DataRow
        Dim boolA As Boolean
        Dim bool As Boolean

        If AllowPrint() Then
        Else
            Exit Sub
        End If

        'first ensure appropriate table is chosen in parent window
        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim intRowS As Short
        Dim intRowD As Short
        Dim idS As Int16
        Dim idD As Int16
        Dim Count1 As Short
        Dim Count2 As Short
        Dim strM As String '20190805 LEE

        If Me.cmdEdit.Enabled Then
        Else
            MsgBox("Please Save any changes", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        'remove focus from button
        Me.dgvTables.Focus()

        dgvS = Me.dgvTables
        dgvD = frmH.dgvReportTableConfiguration

        intRowS = dgvS.CurrentRow.Index
        idS = dgvS("ID_TBLREPORTTABLE", intRowS).Value

        For Count1 = 0 To dgvD.RowCount - 1
            idD = dgvD("ID_TBLREPORTTABLE", Count1).Value
            If idS = idD Then
                dgvD.CurrentCell = dgvD.Rows.Item(Count1).Cells("CHARHEADINGTEXT")
                dgvD.Rows.Item(Count1).Cells("CHARHEADINGTEXT").Selected = True
                strM = dgvD.Rows.Item(Count1).Cells("CHARHEADINGTEXT").Value
                Exit For
            End If
        Next

        Call PositionProgress()
        strM = "Creating table " & strM
        Me.lblProgress.Text = strM
        Me.lblProgress.Visible = True
        Call ExampleSection("AssignSamples")
        Me.lblProgress.Visible = False


        Try
            frmWordStatement.Activate()
        Catch ex As Exception

        End Try

        'Me.WindowState = FormWindowState.Minimized


    End Sub

    Private Sub cmdAdvTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdvTable.Click

        '20181206 LEE: declare later
        'Dim frm As New frmReportTableConfig

        Dim dv As System.Data.DataView
        Dim dgv As DataGridView
        Dim strF As String
        Dim strFo As String
        Dim boolLoad As Boolean

        'first ensure appropriate table is chosen in parent window
        Dim dgvS As DataGridView
        Dim dgvD As DataGridView
        Dim intRowS As Short
        Dim intRowD As Short
        Dim idS As Int16
        Dim idD As Int16
        Dim Count1 As Short
        Dim Count2 As Short

        dgvS = Me.dgvTables
        dgvD = frmH.dgvReportTableConfiguration

        '20181206 LEE:
        'Hmmm. Why is this being done?
        'intRowS = dgvS.CurrentRow.Index
        'idS = dgvS("ID_TBLREPORTTABLE", intRowS).Value

        'For Count1 = 0 To dgvD.RowCount - 1
        '    idD = dgvD("ID_TBLREPORTTABLE", Count1).Value
        '    If idS = idD Then
        '        dgvD.CurrentCell = dgvD.Rows.Item(Count1).Cells("CHARHEADINGTEXT")
        '        dgvD.Rows.Item(Count1).Cells("CHARHEADINGTEXT").Selected = True
        '        Exit For
        '    End If
        'Next

        '*****

        Dim intRow As Short
        '20181206 LEE
        'Do this stuff here
        Dim dgv1 As DataGridView
        dgv1 = Me.dgvTables
        If dgv1.RowCount = 0 Then
            Exit Sub
        End If

        If dgv1.CurrentRow Is Nothing Then
            intRow = 1
        Else
            intRow = dgv1.CurrentRow.Index
        End If

        Dim frm As New frmReportTableConfig


        'set frm.dgv
        dv = Me.dgvTables.DataSource
        dgv = frm.dgvReportTables

        'frm.intORow = dgv.CurrentRow.Index

        If Me.dgvTables.RowCount = 0 Then
            frm.intORow = -1
            frm.idSel = 0
        ElseIf Me.dgvTables.CurrentRow Is Nothing Then
            frm.intORow = -1
            frm.idSel = Me.dgvTables("ID_TBLREPORTTABLE", 0).Value
        Else
            frm.intORow = Me.dgvTables.CurrentRow.Index
            frm.idSel = Me.dgvTables("ID_TBLREPORTTABLE", frm.intORow).Value
        End If


        '*****

        'Exit Sub

        'set frm.dgv
        'dv = Me.dgvReportTableConfiguration.DataSource
        'dv = Me.dgvTables.DataSource
        dv = dgvD.DataSource
        dgv = frm.dgvReportTables

        strF = "BOOLINCLUDE = TRUE"

        boolLoad = True
        strFo = dv.RowFilter
        'dv.RowFilter = strF
        boolLoad = False
        dgv.DataSource = dv
        Call frm.InsertDefault(-1)

        'If Me.dgvTables.RowCount = 0 Then
        '    frm.intORow = -1
        'ElseIf Me.dgvTables.CurrentRow Is Nothing Then
        '    frm.intORow = -1
        'Else
        '    frm.intORow = dgvD.CurrentRow.Index
        'End If

        Call frm.FormLoad()

        frm.ShowDialog()

        frm.Dispose()

        Call ChangedgvTables()

    End Sub

    Private Sub chkAnalysisDate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAnalysisDate.CheckedChanged

        If boolHold Then
        Else
            Call ShowColumnsGroupBox()
        End If

    End Sub

    Private Sub chkAssayLevel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAssayLevel.CheckedChanged

        If boolHold Then
        Else
            Call ShowColumnsGroupBox()
        End If

    End Sub

    Private Sub chkFlag_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFlag.CheckedChanged

        If boolHold Then
        Else
            Call ShowColumnsGroupBox()
        End If

    End Sub

    Private Sub chkDilFactor_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDilFactor.CheckedChanged

        If boolHold Then
        Else
            Call ShowColumnsGroupBox()
        End If

    End Sub

    Private Sub chkSampleType_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSampleType.CheckedChanged

        If boolHold Then
        Else
            Call ShowColumnsGroupBox()
        End If

    End Sub

    Private Sub dgvAssignedSamples_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAssignedSamples.CellValueChanged

        Dim intCol As Short
        Dim intCol1 As Short
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim bool1 As Boolean
        Dim bool2 As Boolean
        Dim val1

        dgv = Me.dgvAssignedSamples

        If dgv.ReadOnly Then
        Else
            Try

                intCol = dgv.CurrentCell.ColumnIndex
                intRow = dgv.CurrentCell.RowIndex
                intCol1 = dgv.Columns("BOOLEXCLSAMPLE").Index

                If StrComp(dgv.Columns(intCol).Name, "BOOLEXCLSAMPLECHK", CompareMethod.Text) = 0 Then
                    bool1 = dgv(intCol, intRow).Value
                    dgv.CurrentCell = dgv.Rows(intRow).Cells("CHARANALYTE")
                    'MsgBox(bool1)
                    bool2 = Not (bool1)

                    If bool1 Then
                        dgv("BOOLEXCLSAMPLE", intRow).Value = -1
                    Else
                        dgv("BOOLEXCLSAMPLE", intRow).Value = 0
                    End If

                ElseIf StrComp(dgv.Columns(intCol).Name, "NUMMINACCCRIT", CompareMethod.Text) = 0 Then
                    If Me.chkAsynch.Checked Then
                    Else
                        val1 = dgv(intCol, intRow).Value
                        dgv("NUMMAXACCCRIT", intRow).Value = val1
                    End If
                ElseIf StrComp(dgv.Columns(intCol).Name, "NUMMAXACCCRIT", CompareMethod.Text) = 0 Then
                    If Me.chkAsynch.Checked Then
                    Else
                        val1 = dgv(intCol, intRow).Value
                        dgv("NUMMINACCCRIT", intRow).Value = val1
                    End If
                End If

            Catch ex As Exception

            End Try

        End If


    End Sub

    Private Sub dgvAssignedSamples_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAssignedSamples.CurrentCellDirtyStateChanged

        Dim intCol As Short
        Dim intRow As Short
        Dim dgv As DataGridView
        Dim bool1 As Boolean
        Dim bool2 As Boolean

        dgv = Me.dgvAssignedSamples

        If dgv.ReadOnly Then
        Else
            intCol = dgv.CurrentCell.ColumnIndex
            intRow = dgv.CurrentCell.RowIndex

            If StrComp(dgv.Columns(intCol).Name, "BOOLEXCLSAMPLECHK", CompareMethod.Text) = 0 Then
                dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                bool1 = dgv(intCol, intRow).Value
            End If
        End If


    End Sub

    Private Sub chkUSEGUWUACCCRIT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUSEGUWUACCCRIT.CheckedChanged

        Call ShowCrit1()
        boolDoAccCrit = True
        Call ShowCrit()
        boolDoAccCrit = False
        Call ShowColumnsGroupBox()

    End Sub

    Sub ShowCrit1()

        If Me.chkUSEGUWUACCCRIT.Checked And Me.cmdEdit.Enabled = False Then
            Me.chkAsynch.Enabled = True
            Me.cmdFillDown.Enabled = True
            Me.cmdClearAll.Enabled = True
            Me.cmdClearDown.Enabled = True
        Else
            Me.chkAsynch.Enabled = False
            Me.cmdFillDown.Enabled = False
            Me.cmdClearAll.Enabled = False
            Me.cmdClearDown.Enabled = False
        End If

    End Sub

    Sub ShowCrit()

        If boolDoAccCrit Then
        Else
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim dv As System.Data.DataView
        Dim int2 As Short

        dgv = Me.dgvAssignedSamples
        dv = dgv.DataSource

        If Me.chkUSEGUWUACCCRIT.Checked And LAllowGuWuAccCrit And gAllowGuWuAccCrit And Me.panAccCrit.Visible Then
            dgv.Columns("BOOLUSEGUWUACCCRIT").Visible = False
            dgv.Columns("NUMMINACCCRIT").Visible = True
            dgv.Columns("NUMMAXACCCRIT").Visible = True

            For Count1 = 0 To dv.Count - 1
                dv(Count1).BeginEdit()
                dv(Count1).Item("BOOLUSEGUWUACCCRIT") = -1
                dv(Count1).EndEdit()
            Next
        Else
            dgv.Columns("BOOLUSEGUWUACCCRIT").Visible = False
            dgv.Columns("NUMMINACCCRIT").Visible = False
            dgv.Columns("NUMMAXACCCRIT").Visible = False

            For Count1 = 0 To dv.Count - 1
                dv(Count1).BeginEdit()
                dv(Count1).Item("BOOLUSEGUWUACCCRIT") = 0
                dv(Count1).EndEdit()
            Next
        End If

        int2 = dgv.Columns.Count


    End Sub

    Sub ShowCritNoDB()

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim dv As System.Data.DataView
        Dim int2 As Short
        Dim boolYes As Boolean = False
        Dim int1 As Short

        dgv = Me.dgvAssignedSamples
        dv = dgv.DataSource

        If dv.Count = 0 Then
            boolYes = False
        Else
            int1 = NZ(dv(0).Item("BOOLUSEGUWUACCCRIT"), 0)
            If int1 = -1 Then
                boolYes = True
            Else
                boolYes = False
            End If
        End If

        If boolYes And LAllowGuWuAccCrit And gAllowGuWuAccCrit And Me.panAccCrit.Visible Then
            dgv.Columns("BOOLUSEGUWUACCCRIT").Visible = False
            dgv.Columns("NUMMINACCCRIT").Visible = True
            dgv.Columns("NUMMAXACCCRIT").Visible = True

        Else
            dgv.Columns("BOOLUSEGUWUACCCRIT").Visible = False
            dgv.Columns("NUMMINACCCRIT").Visible = False
            dgv.Columns("NUMMAXACCCRIT").Visible = False
        End If

    End Sub

    Private Sub cmdFillDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFillDown.Click

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim intCol As Short
        Dim Count1 As Short
        Dim val1, val2
        Dim dv As System.Data.DataView

        dgv = Me.dgvAssignedSamples
        'intRow = dgv.CurrentRow.Index

        Try
            intRow = dgv.CurrentRow.Index
            val1 = dgv("NUMMINACCCRIT", intRow).Value
            val2 = dgv("NUMMAXACCCRIT", intRow).Value

            dv = dgv.DataSource
            For Count1 = intRow + 1 To dv.Count - 1

                dv(Count1).BeginEdit()
                dv(Count1).Item("NUMMINACCCRIT") = val1
                dv(Count1).Item("NUMMAXACCCRIT") = val2
                dv(Count1).EndEdit()

            Next


        Catch ex As Exception

        End Try


    End Sub

    Sub CheckUseGuWuAccCrit()

        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim boolU As Boolean = True
        Dim var1

        dv = Me.dgvAssignedSamples.DataSource
        boolU = False

        If dv.Count = 0 Then
        Else
            For Count1 = 0 To dv.Count - 1
                var1 = dv(Count1).Item("BOOLUSEGUWUACCCRIT")
                If IsDBNull(var1) Then
                    boolU = False
                    Exit For
                ElseIf Len(var1) = 0 Then
                    boolU = False
                    Exit For
                ElseIf var1 = 0 Then
                    boolU = False
                    Exit For
                ElseIf var1 = -1 Then
                    boolU = True
                    Exit For
                Else
                    boolU = False
                    Exit For
                End If
            Next

            If boolU Then
                Me.chkUSEGUWUACCCRIT.Checked = True
            Else
                Me.chkUSEGUWUACCCRIT.Checked = False
            End If

        End If

    End Sub


    Private Sub cmdClearDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearDown.Click

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim Count1 As Short

        Dim dv As System.Data.DataView

        dgv = Me.dgvAssignedSamples
        'intRow = dgv.CurrentRow.Index
        dv = dgv.DataSource

        Try
            intRow = dgv.CurrentRow.Index
            For Count1 = intRow To dv.Count - 1

                dv(Count1).BeginEdit()
                dv(Count1).Item("NUMMINACCCRIT") = DBNull.Value
                dv(Count1).Item("NUMMAXACCCRIT") = DBNull.Value
                dv(Count1).EndEdit()

            Next
        Catch ex As Exception

        End Try


    End Sub

    Private Sub cmdClearAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClearAll.Click

        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim dv As System.Data.DataView

        dgv = Me.dgvAssignedSamples
        dv = dgv.DataSource
        For Count1 = 0 To dv.Count - 1

            dv(Count1).BeginEdit()
            dv(Count1).Item("NUMMINACCCRIT") = DBNull.Value
            dv(Count1).Item("NUMMAXACCCRIT") = DBNull.Value
            dv(Count1).EndEdit()

        Next


    End Sub


    Private Sub chkAnalRT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAnalRT.CheckedChanged

        If boolHold Then
        Else
            Call ShowColumnsGroupBox()
        End If

    End Sub

    Private Sub chkISRT_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkISRT.CheckedChanged

        If boolHold Then
        Else
            Call ShowColumnsGroupBox()
        End If

    End Sub


    Private Sub optQCConcs_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optQCConcs.CheckedChanged
        Call NomConcFill(True)
    End Sub

    Private Sub optCalibrConcs_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCalibrConcs.CheckedChanged
        Call NomConcFill(True)
    End Sub

    Private Sub chkDesignSampleID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDesignSampleID.CheckedChanged

        If boolHold Then
        Else
            Call ShowColumnsGroupBox()
        End If

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles lblT1.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub txtFilterSamples_Validated(sender As Object, e As EventArgs) Handles txtFilterSamples.Validated

        Call DoFilter()

    End Sub

    Private Sub txtFilterSamples_KeyUp(sender As Object, e As KeyEventArgs) Handles txtFilterSamples.KeyUp

        If e.KeyCode = 13 Then
            'have the cursor move to the next item
            'this will call the txtfiltersamples_validated event
            Me.dgvAnalyticalRuns.Focus()
        End If

    End Sub

    Sub DoFilter()

        If boolFormLoad Then
            Exit Sub
        End If

        boolFromFilter = True
        Cursor.Current = Cursors.WaitCursor
        Call FillAnalyticalRuns("")
        Cursor.Current = Cursors.Default
        Call CountSamples()
        boolFromFilter = False

    End Sub

    Sub frmAssignSamples_ToolTipSet()

        Dim str1 As String

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

        Try

            'Set mode buttons
            toolTip1.SetToolTip(Me.cmdEdit, "Change to editing mode")
            toolTip1.SetToolTip(Me.cmdOK, "Save all changes")
            toolTip1.SetToolTip(Me.cmdReset, "Cancel unsaved changes")
            toolTip1.SetToolTip(Me.cmdExit, "Go Back")

            'Other buttons
            toolTip1.SetToolTip(Me.cmdShowRunSummary, "Show Watson run summary (all compounds)")
            toolTip1.SetToolTip(Me.cmdPrintReport, "Generate report for selected table only")
            toolTip1.SetToolTip(Me.cmdAdvTable, "Select options (statistical, etc.) for each table")
            toolTip1.SetToolTip(Me.chkUSEGUWUACCCRIT, "Add columns for user-entered acceptance range (+/- %)")
            toolTip1.SetToolTip(Me.chkAsynch, "Allows +/- acceptance range  % to be different.")
            toolTip1.SetToolTip(Me.cmdFillDown, "Fill down acceptance ranges")
            toolTip1.SetToolTip(Me.cmdClearAll, "Clear All acceptance range values")
            toolTip1.SetToolTip(Me.optQCConcs, "Show QC concentrations")
            toolTip1.SetToolTip(Me.optCalibrConcs, "Show Calibration Standard concentrations")
            toolTip1.SetToolTip(Me.cmdNomConc, "Click to add nominal concentration to selected samples")
            toolTip1.SetToolTip(Me.cmdHelper1, "Click to add labels to selected rows")
            toolTip1.SetToolTip(Me.cmdCopy, "Copy Run assignments to another analyte")
            toolTip1.SetToolTip(Me.cmdRemove, "Remove selected samples from table")
            toolTip1.SetToolTip(Me.cmdAddRows, "Add selected samples (above) to table below")
            toolTip1.SetToolTip(Me.txtFilterSamples, "Enter filter text (case-sensitive)")
            toolTip1.SetToolTip(Me.lblAnalRuns, "Double-click row(s) in ""Assign Samples..."" to add samples to selected table")
            toolTip1.SetToolTip(Me.lblTables, "Choose table to assign samples to")

            toolTip1.SetToolTip(Me.chkAnalysisDate, "Show Date of Analysis")
            toolTip1.SetToolTip(Me.chkAssayLevel, "Show Assay Concentration Level Number")
            toolTip1.SetToolTip(Me.chkAnalRT, "Show Analyte Retention Time")
            toolTip1.SetToolTip(Me.chkDilFactor, "Show Dilution Factor")
            toolTip1.SetToolTip(Me.chkSampleType, "Show Sample Type")
            toolTip1.SetToolTip(Me.chkFlag, "Show ""Eliminated"" Flag (outliers that have been flagged in Watson)")
            toolTip1.SetToolTip(Me.chkISRT, "Show Internal Standard Retention Time")
            toolTip1.SetToolTip(Me.chkDesignSampleID, "Show Sample ID")

            str1 = "'Not Rejected' means analytical runs with Regression Status of:" & ChrW(10) & ChrW(9) & "'Accepted'" & ChrW(10) & ChrW(9) & "'Regression Performed'" & ChrW(10) & ChrW(9) & "'NO Regression Performed'"
            toolTip1.SetToolTip(Me.cbxAccStatus, str1)

            'Grids
            Me.dgvAnalyticalRuns.Columns("CHARANALYTE").ToolTipText = "Analyte Measured"
            Me.dgvAnalyticalRuns.Columns("STUDYNAME").ToolTipText = "Watson Study Name"
            Me.dgvAnalyticalRuns.Columns("DESIGNSAMPLEID").ToolTipText = "Watson Sample ID"
            Me.dgvAnalyticalRuns.Columns("RUNID").ToolTipText = "Watson Run ID"
            Me.dgvAnalyticalRuns.Columns("RUNSAMPLEORDERNUMBER").ToolTipText = "Sequence Number (from Watson)"
            Me.dgvAnalyticalRuns.Columns("ASSAYDATETIME").ToolTipText = "Date of Analysis (from Watson)"
            Me.dgvAnalyticalRuns.Columns("ASSAYLEVEL").ToolTipText = "Assay Concentration Level Number"
            Me.dgvAnalyticalRuns.Columns("ELIMINATEDFLAG").ToolTipText = "Whether Watson ""Eliminated"" Flag has been set." _
                & vbCrLf & "Eliminated samples (usually outliers) are reported on, footnoted," _
                & vbCrLf & "and excluded from the StudyDoc statistical summary tables."
            Me.dgvAnalyticalRuns.Columns("SAMPLENAME").ToolTipText = "Watson Sample Name"
            Me.dgvAnalyticalRuns.Columns("ALIQUOTFACTOR").ToolTipText = "Dilution factor"
            Me.dgvAnalyticalRuns.Columns("RUNSAMPLEKIND").ToolTipText = "Sample Type (from Watson)"
            Me.dgvAnalyticalRuns.Columns("CONCENTRATION").ToolTipText = "Concentration (rounded to 3 decimal places on this display)"
            Me.dgvAnalyticalRuns.Columns("ANALYTEAREA").ToolTipText = "Analyte Peak Area"
            Me.dgvAnalyticalRuns.Columns("INTERNALSTANDARDAREA").ToolTipText = "Internal Standard Peak Area"
            Me.dgvAnalyticalRuns.Columns("SAMPLETYPEID").ToolTipText = "Sample Matrix"
            Me.dgvAnalyticalRuns.Columns("ANALYTEPEAKRETENTIONTIME").ToolTipText = "Analyte Retention Time"
            Me.dgvAnalyticalRuns.Columns("INTERNALSTANDARDRETENTIONTIME").ToolTipText = "Internal Standard Retention Time"

            If (Not (IsNothing(Me.dgvAssignedSamples))) Then
                If (Me.dgvAssignedSamples.Columns.Count > 0) Then  'Don't want it to assign on View Analytical Runs...
                    Me.dgvAssignedSamples.Columns("CHARSTUDYNAME2").ToolTipText = "Watson Study Name"
                    Me.dgvAssignedSamples.Columns("DESIGNSAMPLEID").ToolTipText = "Watson Sample ID"
                    Me.dgvAssignedSamples.Columns("CHARANALYTE").ToolTipText = "Analyte Measured"
                    Me.dgvAssignedSamples.Columns("RUNID").ToolTipText = "Watson Run ID"
                    Me.dgvAssignedSamples.Columns("ASSAYDATETIME").ToolTipText = "Date of Analysis (from Watson)"
                    Me.dgvAssignedSamples.Columns("RUNSAMPLEORDERNUMBER").ToolTipText = "Sequence Number (from Watson)"
                    Me.dgvAssignedSamples.Columns("ASSAYLEVEL").ToolTipText = "Assay Concentration Level Number"
                    Me.dgvAssignedSamples.Columns("ELIMINATEDFLAG").ToolTipText = "Whether Watson ""Eliminated"" Flag has been set." _
                        & vbCrLf & "Eliminated samples (usually outliers) are reported on, footnoted," _
                        & vbCrLf & "and excluded from the StudyDoc statistical summary tables."
                    Me.dgvAssignedSamples.Columns("SAMPLENAME").ToolTipText = "Watson Sample Name"
                    Me.dgvAssignedSamples.Columns("ALIQUOTFACTOR").ToolTipText = "Dilution factor"
                    Me.dgvAssignedSamples.Columns("RUNSAMPLEKIND").ToolTipText = "Sample Type (from Watson)"
                    Me.dgvAssignedSamples.Columns("CONCENTRATION").ToolTipText = "Concentration (rounded to 3 decimal places on this display)"
                    Me.dgvAssignedSamples.Columns("SAMPLETYPEID").ToolTipText = "Sample Matrix"
                    Me.dgvAssignedSamples.Columns("NOMCONC").ToolTipText = "Nominal Concentration of Analyte in this sample (must be entered by user)"
                    Me.dgvAssignedSamples.Columns("CHARHELPER1").ToolTipText = "Sample Label  (inputted by user)"
                    Me.dgvAssignedSamples.Columns("CHARHELPER2").ToolTipText = "Sample Label  (inputted by user)"
                    Me.dgvAssignedSamples.Columns("BOOLEXCLSAMPLECHK").ToolTipText = "Exclude Sample: Mark as statistical outlier in report, " _
                        & vbCrLf & "do not include in summary statistics.  " _
                        & vbCrLf & "[Note: Exclusions in Watson (also known as samples with the" _
                        & vbCrLf & """Eliminated"" flag) are automatically marked]"
                    Me.dgvAssignedSamples.Columns("ANALYTEPEAKRETENTIONTIME").ToolTipText = "Analyte Retention Time "
                    Me.dgvAssignedSamples.Columns("INTERNALSTANDARDRETENTIONTIME").ToolTipText = "Internal Standard Retention Time"
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub lbl1_Click(sender As Object, e As EventArgs) Handles lbl1.Click

    End Sub

    Private Sub cbxAccStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxAccStatus.SelectedIndexChanged


        If boolFormLoad Then
            Exit Sub
        End If

        Try

            'NOTE: Leave the bool filters as RI
            boolFromFilter = True

            Cursor.Current = Cursors.WaitCursor
            Call FillAnalyticalRuns("")
            Cursor.Current = Cursors.Default

            Call CountSamples()

            Call UpdateAnalRunLabel()

            Try
                Call InitializeFilterRunID(Me.dgvAnalytes.CurrentRow.Index)
            Catch ex As Exception

            End Try


        Catch ex As Exception

        End Try

        boolFromFilter = True

        'set focus on dgvAnalRuns
        'users don't remember that the dropdownbox is the still selected
        'they use the scroll wheel thinking that they are scrolling dgvAnalyticalRuns
        'but are actually scrolling the dropdownbox
        Me.dgvAnalyticalRuns.Focus()

    End Sub

    Sub UpdateAnalRunLabel()

        ''Legend:
        'Me.cbxAccStatus.Items.Add("Show All")
        'Me.cbxAccStatus.Items.Add("Not Rejected") 'RUNANALYTEREGRESSIONSTATUS = 3
        'Me.cbxAccStatus.Items.Add("Rejected") 'RUNANALYTEREGRESSIONSTATUS = 4

        Dim str1 As String = Me.cbxAccStatus.Text
        Dim str2 As String
        Dim str3 As String
        Dim strIS As String
        Dim strName As String
        Dim int1 As Short
        Dim boolIS As Boolean = False

        Try

            If Me.dgvAnalytes.CurrentRow Is Nothing Then
                int1 = 0
            Else
                int1 = Me.dgvAnalytes.CurrentRow.Index
            End If

            strName = Me.dgvAnalytes.Item("ANALYTEDESCRIPTION", int1).Value 'ANALYTEDESCRIPTION
            strIS = NZ(Me.dgvAnalytes.Item("IsIntStd", int1).Value, "No") 'IsIntStd
            If StrComp(strIS, "Yes", CompareMethod.Text) = 0 Then
                boolIS = True
            Else
                boolIS = False
            End If


            If StrComp(str1, "Show All", CompareMethod.Text) = 0 Then
                str3 = "All"
            ElseIf StrComp(str1, "Not Rejected", CompareMethod.Text) = 0 Then
                str3 = "Not Rejected"
            ElseIf StrComp(str1, "Accepted", CompareMethod.Text) = 0 Then
                str3 = "Accepted"
            End If

            If Me.rbFilterForAnalyteYes.Checked Then
                If boolIS Then
                    str2 = str3 & " Analytical Runs for Internal Standard " & strName
                Else
                    str2 = str3 & " Analytical Runs for " & strName
                End If
            Else
                str2 = str3 & " Analytical Runs for all analytes"
            End If


        Catch ex As Exception
            str2 = "All Analytical Runs"
        End Try

        Me.lblAnalRuns.Text = str2

    End Sub

    Private Sub dgvAnalytes_RowLeave(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAnalytes.RowLeave

        Try
            Me.txtdgvAnalytePreviousRow.Text = dgvAnalytes.CurrentRow.Index
        Catch ex As Exception

        End Try

    End Sub


    Private Sub dgvAssignedSamples_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles dgvAssignedSamples.Validating


    End Sub

    Private Sub dgvHelper1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvHelper1.CellContentClick

    End Sub

    Private Sub dgvNomConc_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvNomConc.CellContentClick

    End Sub

    Private Sub cmdAuto_Click(sender As Object, e As EventArgs) Handles cmdAuto.Click


        Dim boolF As Boolean = False
        Dim boolST As Boolean = False 'Selected Table
        Dim frm As New frmAutoAssignBegin

        frm.ShowDialog()

        Dim boolCancel As Boolean = frm.boolCancel

        If boolCancel Then
            GoTo end1
        End If

        If frm.rbOverwrite.Checked Then
            boolF = True
        ElseIf frm.rbOnlyEmpty.Checked Then
            boolF = False
        ElseIf frm.rbSelectedTable.Checked Then
            boolF = True
            boolST = True
        End If

        frm.Dispose()


        'Dim strM As String
        'Dim intR As Short

        'strM = "This action will cycle through the displayed Tables and Analytes and automatically assign samples based on the settings defined for each table in the Advanced Table Configuration window."
        'strM = strM & ChrW(10) & ChrW(10)
        'strM = strM & "If a table and analyte already has samples assigned, the table/analyte will be ignored."
        'strM = strM & ChrW(10) & ChrW(10)
        'strM = strM & "If it is desired to re-assign samples to one or a few table/analytes, first select those tables and remove the currently assigned samples (if they exist)."
        'strM = strM & ChrW(10) & ChrW(10)
        'strM = strM & "If it is desired to overwrite all existing assigned samples, click 'No'."
        'strM = strM & ChrW(10) & ChrW(10)
        'strM = strM & "Otherwise, click 'Yes' to assign samples to only empty table/analytes."
        'strM = strM & ChrW(10) & ChrW(10)
        'strM = strM & "Or click 'Cancel' to back out of this action."

        'intR = MsgBox(strM, vbYesNoCancel, "Continue?")

        'If intR = 6 Then 'yes
        '    boolF = False
        'ElseIf intR = 7 Then 'no
        '    boolF = True
        'Else
        '    GoTo end1
        'End If

        Call AutoAssignSamples(boolF, boolST)

end1:

    End Sub

   

    Sub AutoAssignSamples(boolFresh As Boolean, boolST As Boolean)

        boolAutoAssign = True

        Dim dgvT As DataGridView = Me.dgvTables
        Dim dgvA As DataGridView = Me.dgvAnalytes
        Dim dgvAR As DataGridView = Me.dgvAnalyticalRuns
        Dim dgvAS As DataGridView = Me.dgvAssignedSamples

        Dim CountT As Int16
        Dim CountA As Int16
        Dim CountAR As Int16
        Dim CountAS As Int16

        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short

        Dim intRowsT As Int16 = dgvT.Rows.Count
        Dim intRowsA As Int16 = dgvA.Rows.Count
        Dim intRowsAR As Int16 = dgvAR.Rows.Count
        Dim intRowsAS As Int16 = dgvAS.Rows.Count

        Dim int1 As Short
        Dim int2 As Short

        Dim AnalyteID As Int64
        Dim id1 As Int64

        Dim strF As String
        Dim strS As String
        Dim strFT As String
        Dim strFA As String
        Dim strFAR As String 'analytical run
        Dim strFAS As String
        Dim strFASP As String 'Filter Sample Name

        Dim strM As String
        Dim strM1 As String
        Dim strM2 As String

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strMatrix As String

        Dim var1, var2, var3, var4

        Dim tblAR As DataTable = tblAnalyticalRunSummary

        Dim intTCell As Short
        Dim intACell As Short
        Dim intARCell As Short
        Dim intASCell As Short

        Dim tblAARuns As DataTable = tblAllAnalRuns.Copy 'use copy because must make table case sensitive
        tblAARuns.CaseSensitive = True


        Dim boolIsIntStd As Boolean = False

        'find visible cell
        For Count1 = 0 To dgvT.Columns.Count - 1
            If dgvT.Columns(Count1).Visible Then
                intTCell = Count1
                Exit For
            End If
        Next

        'find visible cell
        For Count1 = 0 To dgvA.Columns.Count - 1
            If dgvA.Columns(Count1).Visible Then
                intACell = Count1
                Exit For
            End If
        Next

        'find visible cell
        For Count1 = 0 To dgvAR.Columns.Count - 1
            If dgvAR.Columns(Count1).Visible Then
                intARCell = Count1
                Exit For
            End If
        Next

        'find visible cell
        For Count1 = 0 To dgvAS.Columns.Count - 1
            If dgvAS.Columns(Count1).Visible Then
                intASCell = Count1
                Exit For
            End If
        Next

        'first clear all filters
        Call ClearFilters()

        'record original setting of cbxAccStatus
        Dim intAccStatusIndex As Short = Me.cbxAccStatus.SelectedIndex
        'set Acceptance status to 'Not Rejected'
        Me.cbxAccStatus.SelectedIndex = 1 '0=Show All, 1=Not Rejected, 3=Accepted

        Dim idRT As Int64
        Dim idCT As Int64
        Dim idCTLast As Int64

        Dim boolDo As Boolean

        Dim intSR As Short = 0
        Dim intER As Short = intRowsT - 1

        If boolST Then
            intSR = dgvT.CurrentRow.Index
            intER = intSR
        End If

        Me.lblProgress.Visible = True
        For CountT = intSR To intER

            strM1 = "Evaluating table " & CountT + 1 & " of " & intER + 1
            Me.lblProgress.Text = strM1
            Me.lblProgress.Refresh()

            Dim intP1 As Short
            Dim intP2 As Short
            Dim boolHasNOT As Boolean = False
            Dim strNOT1 As String = ""
            Dim strNOT2 As String = ""

            Dim strFST As String = "" 'Filter Sample Type
            Dim strFSN As String = "" 'Filter Sample Name
            Dim strFRD1 As String = "" 'Filter Run Description
            Dim strFRD2 As String = "" 'Filter Run Description Run IDs
            Dim strFDF As String = "" 'Filter for Diln Factor
            Dim strFAAR As String = "" 'Filter for Only Accepted Analytical Runs

            Dim boolDoFST As Boolean = False
            Dim boolDoFSN As Boolean = False
            Dim boolDoFRD As Boolean = False
            Dim boolDoFDF As Boolean = False

            Dim boolRec As Boolean = False

            idRT = dgvT("ID_TBLREPORTTABLE", CountT).Value
            idCT = dgvT("ID_TBLCONFIGREPORTTABLES", CountT).Value

            strFASP = "ID_TBLREPORTTABLE = " & idRT
            Dim rowsASP() As DataRow = tblAutoAssignSamples.Select(strFASP)

            Dim arrASP(100)
            Dim intASP As Short = 5
            Dim intASPM As Short = intASP + 1

            Dim strST As String = "" 'Sample Type
            Dim strAR1 As String = "" 'Run Descr 1
            Dim strAR2 As String = "" 'Run Descr 2

            Dim strAR1WON As String = "" ' w/o NOT portion
            Dim strAR2WON As String = "" ' w/o NOT portion

            Dim strAR1N As String = "" ' NOT portion
            Dim strAR2N As String = "" ' NOT portion
            Dim strAAR As String = ""

            Dim strFilterDGVSamples As String = ""

            arrASP(1) = "BOOLUSESTDCOLLABELS"
            arrASP(2) = "CHARSAMPLETYPE"
            arrASP(3) = "CHARRUNDESCR1"
            arrASP(4) = "CHARRUNDESCR2"
            arrASP(5) = "BOOLACCEPTEDONLY"

            boolDo = True

            Select Case idCT
                Case 3, 2 'Summary of Back-Calculated Calibration Std Conc.  20190305 LEE: Start evaluating Regr Constant table (2)
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARCALSTD"
                Case 4 'Summary of Interpolated QC Std Conc
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNONCOREQC"
                Case 11 'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARCOREQC"
                Case 12 'Summary of Interpolated Dilution QC Concentrations
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARDILN"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARDILNFACTOR"
                Case 13 'Summary of Combined Recovery
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRECRS"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRECQC"
                Case 14 'Summary of True Recovery
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRECPES"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRECQC"
                Case 15 'Summary of Suppression/Enhancement/MatrixFactor
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRECRS"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRECPES"
                Case 17 'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments
                    '20180803 LEE:
                    'Check for Matrix Effect
                    For Count1 = 1 To 10
                        intASP = intASP + 1
                        arrASP(intASP) = "CHARLOT" & Count1
                    Next
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER1"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER2"
                Case 18 'Summary of [Period Temp] Stability in Matrix
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNONCOREQC"
                Case 19 'Summary of Freeze/Thaw [#Cycles] Stability in Matrix
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNONCOREQC"
                Case 21 '[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNONCOREQC"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER1"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER2"
                Case 22 '[Period Temp] Stock Solution Stability Assessment
                    intASP = intASP + 1
                    arrASP(intASP) = "CHAROLD"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNEW"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARSTOCKSOLNCONC"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER1"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER2"
                Case 23 '[Period Temp] Spiking Solution Stability Assessment
                    intASP = intASP + 1
                    arrASP(intASP) = "CHAROLD"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNEW"
                Case 29 '[Period Temp] Long-Term QC Std Storage Stability
                    intASP = intASP + 1
                    arrASP(intASP) = "CHAROLD"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNEW"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER1"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER2"
                Case 31 'Ad Hoc QC Stability Table
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNONCOREQC"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER1"
                Case 32 'Ad Hoc QC Stability Comparison Table
                    intASP = intASP + 1
                    arrASP(intASP) = "CHAROLD"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNEW"

                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNEW2"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARNEW3"

                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER1"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER2"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER3"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER4"

                Case 34 'Selectivity of Lots
                    '20180803 LEE:
                    'Check for Matrix Effect
                    For Count1 = 1 To 10
                        intASP = intASP + 1
                        arrASP(intASP) = "CHARLOT" & Count1
                    Next
                    For Count1 = 1 To 10
                        intASP = intASP + 1
                        arrASP(intASP) = "CHARLOTWOIS" & Count1
                    Next
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARCALSTD"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER1"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARRUNIDENTIFIER2"

                Case 35 'Carryover
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARLLOQ"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARULOQ"
                    intASP = intASP + 1
                    arrASP(intASP) = "CHARBLANK"
                Case Else

                    boolDo = False


            End Select

            If boolDo Then

                'do accepted analytical runs

                Dim intAAR As Short = -1

                Dim intNumLots As Short = 0
                Dim intRec As Short = 0
                Dim intElse As Short = 0
                ''legend
                'Dim strFST As String 'Filter Sample Type
                'Dim strFSN As String 'Filter Sample Name
                'Dim strFRD As String 'Filter Run Description
                'Dim strFDF As String = "" 'Filter for Diln Factor

                'first evaluate 3 and 4 for analytical run specific stuff
                str1 = arrASP(3)
                strAR1 = NZ(rowsASP(0).Item(str1), "")
                If Len(strAR1) = 0 Then
                Else
                    intP1 = InStr(1, strAR1, "(", CompareMethod.Text)
                    If intP1 > 0 Then
                        strAR1WON = Mid(strAR1, 1, intP1 - 1)
                        strAR1N = ReturnNOT(strAR1, "RUNDESCRIPTION")
                    Else
                        strAR1WON = strAR1
                    End If
                End If

                '20161024 LEE: This 2nd Analytical Run description is depricated. Leave code here for now because is doesn't affect anything
                str2 = arrASP(4)
                strAR2 = NZ(rowsASP(0).Item(str2), "")
                If Len(strAR2) = 0 Then
                Else
                    intP1 = InStr(1, strAR2, "(", CompareMethod.Text)
                    If intP1 > 0 Then
                        strAR2WON = Mid(strAR2, 1, intP1 - 1)
                        strAR2N = ReturnNOT(strAR2, "RUNDESCRIPTION")
                    Else
                        strAR2WON = strAR2
                    End If
                End If

                str3 = arrASP(5) 'boolacceptedonly
                intAAR = NZ(rowsASP(0).Item(str3), -1)

                Dim boolDoAR1 As Boolean = False 'Run Descr 1
                Dim boolDoAR2 As Boolean = False 'Run Descr 2
                Dim boolDoAAR As Boolean = True 'Only Accepted Anal Runs, default is true

                Dim rowsAR() As DataRow


                If intAAR = -1 Then
                    boolDoAAR = True
                Else
                    boolDoAAR = False
                End If

                If boolDoAAR Then
                    'need to add a filter such that RUNANALYTEREGRESSIONSTATUS = 3
                    strFAAR = "RUNANALYTEREGRESSIONSTATUS = 3"
                End If

                If Len(strAR1N) > 0 Or Len(strAR1WON) > 0 Then
                    boolDoAR1 = True
                End If
                If Len(strAR2N) > 0 Or Len(strAR2WON) > 0 Then
                    boolDoAR2 = True
                End If

                If boolDoAR1 Then
                    'find runid with this fragment in analytical run description
                    If Len(strAR1WON) = 0 Then
                        If Len(strAR1N) = 0 Then
                        Else
                            strFRD1 = strAR1N
                        End If
                    Else
                        strFRD1 = "RUNDESCRIPTION LIKE '*" & strAR1WON & "*'"
                        If Len(strAR1N) = 0 Then
                        Else
                            strFRD1 = strFRD1 & " AND " & strAR1N
                        End If
                    End If

                Else
                    If boolDoAR2 Then

                        If Len(strAR2WON) = 0 Then
                            If Len(strAR2N) = 0 Then
                            Else
                                strFRD2 = strAR2N
                            End If
                        Else
                            strFRD2 = "RUNDESCRIPTION LIKE '*" & strAR2WON & "*'"
                            If Len(strAR2N) = 0 Then
                            Else
                                strFRD2 = strFRD2 & " AND " & strAR2N
                            End If
                        End If

                    End If
                End If

                If boolDoAR1 Or boolDoAR2 Then '
                    boolDoFRD = True
                End If

                'now do sample type
                str3 = arrASP(2)
                strST = NZ(rowsASP(0).Item(str3), "")
                If Len(strST) > 0 Then
                    boolDoFST = True
                End If
                If boolDoFST Then
                    strFST = "RUNSAMPLEKIND = '" & strST & "'"
                End If

                'wait to do this evaluation within dgvAnalyte loop

                strFASP = ""

                'now do individual table stuff
                int1 = 0

                strNOT1 = ""
                strNOT2 = ""

                For Count1 = intASPM To intASP

                    boolHasNOT = False

                    str1 = NZ(arrASP(Count1), "") 'this returns column name
                    If Len(str1) = 0 Then
                    Else

                        var2 = NZ(rowsASP(0).Item(str1), "")
                        var3 = StripNOT(var2.ToString)
                        If Len(var3) = 0 Then
                        Else

                            intP1 = 0
                            intP2 = 0

                            'check to see if var2 has NOT string
                            intP1 = InStr(1, var2, "(", CompareMethod.Text)
                            If intP1 = 0 Then
                            Else
                                intP2 = InStr(intP1 + 1, var2, ")", CompareMethod.Text)
                                If intP2 = 0 Then
                                Else
                                    If intP2 - intP1 = 1 Then
                                    Else
                                        boolHasNOT = True
                                        strNOT1 = ReturnNOT(var2.ToString, "SAMPLENAME") 'this includes parentheses
                                        If Len(strNOT2) = 0 Then
                                            strNOT2 = "(" & strNOT1
                                        Else
                                            strNOT2 = strNOT2 & " OR " & strNOT1
                                        End If

                                    End If

                                End If
                            End If

                            If intP1 = 0 Then
                            Else
                                'strip off parentheses from var2
                                var2 = Mid(var2.ToString, 1, intP1 - 1)
                            End If


                            Select Case str1

                                Case "CHARDILNFACTOR"
                                    intElse = intElse + 1
                                    If intElse = 1 Then
                                        strFSN = "(ALIQUOTFACTOR = " & var2
                                    Else
                                        strFSN = strFSN & " AND ALIQUOTFACTOR = " & var2
                                    End If

                                Case "CHARRECPES", "CHARRECRS", "CHARRECQC"
                                    intRec = intRec + 1
                                    If intRec = 1 Then
                                        strFSN = "(SAMPLENAME LIKE '*" & var2 & "*'"
                                    Else
                                        strFSN = strFSN & " OR SAMPLENAME LIKE '*" & var2 & "*'"
                                        If Count1 = intASP Then
                                            strFSN = strFSN & ")"
                                        End If
                                    End If

                                    boolRec = True

                                Case "CHAROLD", "CHARNEW", "CHARNEW2", "CHARNEW3"
                                    intRec = intRec + 1
                                    If intRec = 1 Then
                                        strFSN = "(SAMPLENAME LIKE '*" & var2 & "*'"
                                    Else
                                        strFSN = strFSN & " OR SAMPLENAME LIKE '*" & var2 & "*'"
                                        If Count1 = intASP Then
                                            strFSN = strFSN & ")"
                                        End If
                                    End If

                                Case "CHARSTOCKSOLNCONC", "CHARRUNIDENTIFIER1", "CHARRUNIDENTIFIER2", "CHARRUNIDENTIFIER3", "CHARRUNIDENTIFIER4" 'skip these

                                Case "CHARLLOQ", "CHARULOQ", "CHARBLANK"
                                    intRec = intRec + 1
                                    If intRec = 1 Then
                                        strFSN = "(SAMPLENAME LIKE '*" & var2 & "*'"
                                    Else
                                        strFSN = strFSN & " OR SAMPLENAME LIKE '*" & var2 & "*'"
                                        If Count1 = intASP Then
                                            strFSN = strFSN & ")"
                                        End If
                                    End If

                                Case "CHARCALSTD"
                                    intRec = intRec + 1
                                    If intRec = 1 Then
                                        strFSN = "(SAMPLENAME LIKE '*" & var2 & "*'"
                                    Else
                                        strFSN = strFSN & " OR SAMPLENAME LIKE '*" & var2 & "*'"
                                        If Count1 = intASP Then
                                            strFSN = strFSN & ")"
                                        End If
                                    End If

                                Case Else

                                    Select Case True

                                        '20181217 LEE:
                                        Case str1.Contains("CHARLOTWOIS")
                                            intRec = intRec + 1
                                            If intRec = 1 Then
                                                strFSN = "(SAMPLENAME LIKE '*" & var2 & "*'"
                                            Else
                                                strFSN = strFSN & " OR SAMPLENAME LIKE '*" & var2 & "*'"
                                                If Count1 = intASP Then
                                                    strFSN = strFSN & ")"
                                                End If
                                            End If

                                        Case str1.Contains("CHARLOT")
                                            intRec = intRec + 1
                                            If intRec = 1 Then
                                                strFSN = "(SAMPLENAME LIKE '*" & var2 & "*'"
                                            Else
                                                strFSN = strFSN & " OR SAMPLENAME LIKE '*" & var2 & "*'"
                                                If Count1 = intASP Then
                                                    strFSN = strFSN & ")"
                                                End If
                                            End If

                                        Case Else
                                            intElse = intElse + 1
                                            If intElse = 1 Then
                                                strFSN = "(SAMPLENAME LIKE '*" & var2 & "*'"
                                            Else
                                                strFSN = strFSN & " AND SAMPLENAME LIKE '*" & var2 & "*'"
                                            End If

                                    End Select

                            End Select

                        End If
                    End If

                Next Count1

                If Len(strFSN) = 0 And Len(strNOT2) = 0 Then
                Else

                    If Len(strFSN) = 0 Then
                        'add NOT filters
                        If Len(strNOT2) = 0 Then
                        Else
                            strNOT2 = strNOT2 & ")"
                            strFSN = strNOT2
                        End If
                        boolDoFSN = True
                    Else
                        'check for last parenthesis
                        str1 = Mid(strFSN, Len(strFSN), 1)
                        If StrComp(str1, ")", CompareMethod.Text) = 0 Then
                        Else
                            strFSN = strFSN & ")"
                        End If

                        'add NOT filters
                        If Len(strNOT2) = 0 Then
                        Else
                            strNOT2 = strNOT2 & ")"
                            strFSN = strFSN & " AND " & strNOT2
                        End If
                        boolDoFSN = True
                    End If


                End If

                ''legend
                'Dim strFST As String 'Filter Sample Type
                'Dim strFSN As String 'Filter Sample Name
                'Dim strFRD As String 'Filter Run Description

                If boolDoFST Or boolDoFSN Or boolDoFRD Then 'do not evaluate boolDoAAR here
                Else
                    GoTo nextCountT
                End If

                'now assign samples per analyte/matrix

                'select this table in dgv
                dgvT.CurrentCell = dgvT.Rows(CountT).Cells(intTCell)
                dgvT.Rows(CountT).Selected = True
                'check to see if row is visible

                'MsgBox("Done changing table row")

                'fill dgvAnalytes
                Call FilldgvAnalytes()

                intRowsA = dgvA.Rows.Count

                'loop through analyte
                For CountA = 0 To intRowsA - 1

                    strM2 = "Analyte " & CountA + 1 & " of " & intRowsA
                    Me.lblProgress.Text = strM1 & ChrW(10) & ChrW(10) & strM2
                    Me.lblProgress.Refresh()

                    boolIsIntStd = False
                    Dim strIsIntStd As String = ""

                    'select this table in dgv
                    boolDontChange = True
                    dgvA.CurrentCell = dgvA.Rows(CountA).Cells(intACell)
                    dgvA.Rows(CountA).Selected = True
                    boolDontChange = False

                    'do Analyte change action
                    If boolDoAAR Then
                        Call dgvAnalyteSelectionChange("Accepted")
                    Else
                        Call dgvAnalyteSelectionChange("")
                    End If

                    Dim dvAR As DataView = dgvAR.DataSource
                    strFilterDGVSamples = dvAR.RowFilter


                    intRowsAS = dgvAS.Rows.Count
                    If intRowsAS > 0 Then
                        If boolFresh Then

                            'Not needed
                            'Already in edit mode
                            'If Me.cmdEdit.Enabled Then
                            '    Call DoThis("Edit")
                            'End If

                            'delete these rows
                            'select all rows
                            boolCont = False
                            dgvAS.CurrentCell = dgvAS.Rows(0).Cells(intASCell)
                            dgvAS.Rows(0).Selected = True
                            For Count1 = 0 To dgvAS.Rows.Count - 1
                                boolCont = False
                                dgvAS.Rows(Count1).Selected = True
                            Next
                            boolCont = True
                            'now remove rows
                            Call RemoveRows()
                        Else
                            GoTo nextCountA
                        End If
                    End If

                    strIsIntStd = dgvA("IsIntStd", CountA).Value
                    If StrComp(strIsIntStd, "Yes", CompareMethod.Text) = 0 Then
                        boolIsIntStd = True
                    Else
                        boolIsIntStd = False
                    End If
                    If boolIsIntStd Then
                        'find an analyte whose IntStd is this one
                        AnalyteID = 0
                        strMatrix = ""
                        str2 = dgvA("OriginalAnalyteDescription", CountA).Value 'this is intstd name
                        For Count1 = 0 To dgvA.Rows.Count - 1
                            AnalyteID = dgvA("ANALYTEID", Count1).Value
                            str1 = NZ(dgvA("IntStd", Count1).Value, "")
                            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                                AnalyteID = dgvA("ANALYTEID", Count1).Value
                                strMatrix = dgvA("MATRIX", Count1).Value
                                Exit For
                            End If
                        Next
                        If AnalyteID = 0 Then 'this IntStd isn't used
                            GoTo nextCountA
                        End If
                    Else
                        AnalyteID = dgvA("ANALYTEID", CountA).Value
                        strMatrix = dgvA("MATRIX", CountA).Value
                    End If



                    'look for RunID
                    strF = "(ANALYTEID = " & AnalyteID & " AND SAMPLETYPEID = '" & strMatrix & "')"

                    ''legend
                    'Dim strFST As String = "" 'Filter Sample Type
                    'Dim strFSN As String = "" 'Filter Sample Name
                    'Dim strFRD1 As String = "" 'Filter Run Description
                    'Dim strFRD2 As String = "" 'Filter Run Description Run IDs

                    If boolDoFRD Then

                        'need to make a filter with RunID's
                        Dim strF1 As String
                        strF1 = strF & " AND " & strFRD1

                        rowsAR = tblAARuns.Select(strF1, "RUNID ASC") '
                        'this possibly will return more than one row because we haven't filtered for analyte id

                        For Count1 = 0 To rowsAR.Length - 1
                            var1 = rowsAR(Count1).Item("RUNID")
                            If Count1 = 0 Then
                                strFRD2 = "(RUNID = " & var1
                            Else
                                strFRD2 = strFRD2 & " OR RUNID = " & var1
                                If Count1 = rowsAR.Length - 1 Then
                                    strFRD2 = strFRD2 & ")"
                                End If
                            End If
                        Next
                        If rowsAR.Length = 0 Then
                        Else
                            'check for ending ')'
                            If Len(strFRD2) = 0 Then
                            Else
                                str1 = Mid(strFRD2, Len(strFRD2), 1)
                                If StrComp(str1, ")", CompareMethod.Text) = 0 Then
                                Else
                                    strFRD2 = strFRD2 & ")"
                                End If
                            End If
                        End If

                        If Len(strFRD2) = 0 Then
                            'this means there are no matching records
                            strFRD2 = "(RUNID < 0)"
                        End If

                    End If

                    ''legend
                    'Dim strFST As String = "" 'Filter Sample Type
                    'Dim strFSN As String = "" 'Filter Sample Name
                    'Dim strFRD1 As String = "" 'Filter Run Description
                    'Dim strFRD2 As String = "" 'Filter Run Description Run IDs

                    dvAR = dgvAR.DataSource
                    Dim strCRF As String = dvAR.RowFilter

                    strFASP = "(" & strCRF & ")"
                    If boolDoFST Then
                        strFASP = strFASP & " AND (" & strFST & ")"
                    End If
                    If boolDoFSN Then
                        strFASP = strFASP & " AND (" & strFSN & ")"
                    End If
                    If boolDoFRD Then
                        strFASP = strFASP & " AND (" & strFRD2 & ")"
                    End If
                    If boolDoAAR Then
                        'don't need to do this because previous 'Call dgvAnalyteSelectionChange("Accepted")' accounts for RunAnalRegression status
                        'strFASP = strFASP & " AND (" & strFAAR & ")"
                    End If

                    'Console.WriteLine("idCT: " & idCT)
                    'Console.WriteLine(strFASP)



                    Try
                        dvAR.RowFilter = strFASP
                    Catch ex As Exception
                        var1 = var1 'debug
                    End Try

                    'console.writeline(strFASP)

                    Dim boolDoHelper2 As Boolean = False

                    If dvAR.Count = 0 Then
                    Else

                        'select all rows
                        'select first row
                        dgvAR.CurrentCell = dgvAR.Rows(0).Cells(intARCell)
                        dgvAR.Rows(0).Selected = True
                        For Count1 = 0 To dgvAR.Rows.Count - 1
                            dgvAR.Rows(Count1).Selected = True
                        Next

                        'MsgBox("Begin Adding Rows")
                        'now assign samples
                        Call AddRows()

                        idCTLast = idCT

                        'MsgBox("Done Adding Rows")

                        'now do special things, if needed
                        Dim intRecSel As Short
                        Select Case idCT
                            Case Is = 1 'Summary of Analytical Runs
                            Case Is = 2 'Summary of Regression Constants
                            Case Is = 3 'Summary of Back-Calculated Calibration Std Conc
                            Case Is = 4 'Summary of Interpolated QC Std Conc
                            Case Is = 5 'Summary of Samples
                            Case Is = 6 'Summary of Reassayed Samples
                            Case Is = 7 'Summary of Repeat Samples
                            Case Is = 11 'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
                            Case Is = 12 'Summary of Interpolated Dilution QC Concentrations

                            Case 13 'Summary of Combined Recovery

                                'must assign Term1
                                For Count1 = 1 To 2
                                    Select Case Count1
                                        Case 1
                                            var2 = arrASP(intASP - 1) ' "CHARRECRS"'this returns column name
                                            intRecSel = 1
                                        Case 2
                                            var2 = arrASP(intASP) ' "CHARRECQC"'this returns column name
                                            intRecSel = 0
                                    End Select

                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True
                                        End If
                                    Next
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()
                                    End If

                                Next
                            Case 14 'Summary of True Recovery
                                'str1 = "CHARRECPES"
                                'str1 = "CHARRECQC"

                                'must assign Term1
                                For Count1 = 1 To 2
                                    Select Case Count1
                                        Case 1
                                            var2 = arrASP(intASP - 1) ' "CHARRECPES"'this returns column name
                                            intRecSel = 1
                                        Case 2
                                            var2 = arrASP(intASP) ' "CHARRECQC"'this returns column name
                                            intRecSel = 0
                                    End Select

                                    'str1 = NZ(rowsASP(0).Item(var2), "")
                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True
                                        End If
                                    Next
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()
                                    End If

                                Next

                            Case 15 'Summary of Suppression/Enhancement/MatrixFactor

                                'str1 = "CHARRECRS"
                                'str1 = "CHARRECPES"

                                'must assign Term1
                                For Count1 = 1 To 2
                                    Select Case Count1
                                        Case 1
                                            var2 = arrASP(intASP - 1) ' "CHARRECRS"'this returns column name
                                            intRecSel = 0
                                        Case 2
                                            var2 = arrASP(intASP) ' "CHARRECPES"'this returns column name
                                            intRecSel = 1
                                    End Select

                                    'str1 = NZ(rowsASP(0).Item(var2), "")
                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    'MsgBox("Start selecting AS")
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True
                                        End If
                                    Next
                                    'MsgBox("End selecting AS")
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()
                                    End If

                                Next

                            Case Is = 17 'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments
                                'must assign Term1

                                For Count1 = 1 To 10

                                    var2 = arrASP(intASPM + Count1 - 1)
                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    If Len(str1) = 0 Then
                                        Exit For
                                    End If

                                    intRecSel = Count1 - 1

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    'MsgBox("Start selecting AS")
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True

                                            Exit For

                                        End If
                                    Next
                                    'MsgBox("End selecting AS")
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()
                                    End If

                                Next

                            Case Is = 34 'Selectivity in Individual Lots Table v1
                                '20181217 LEE
                                'need to add Term1 (Lot 1, Lot 2, etc.)
                                'need to add run identifier, which dgvHelper2
                                'There are 21 items in Selectivity col

                                Dim intAA As Short = 0

                                For Count3 = 1 To 3

                                    'assign Term1
                                    Me.dgvHelper2.CurrentCell = dgvHelper2.Rows(Count3 - 1).Cells(1)
                                    Me.dgvHelper2.Rows(Count3 - 1).Selected = True

                                    For Count1 = 1 To 10

                                        intAA = intAA + 1
                                        var2 = arrASP(intASPM + intAA - 1)
                                        str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                        If Len(str1) = 0 Then
                                            Exit For
                                        End If

                                        intRecSel = Count1 - 1

                                        'unselect all dgvAS rows
                                        dgvAS.ClearSelection()
                                        dgvAS.CurrentCell = Nothing
                                        int1 = 0
                                        'MsgBox("Start selecting AS")
                                        For Count2 = 0 To dgvAS.Rows.Count - 1
                                            str2 = dgvAS("SAMPLENAME", Count2).Value
                                            If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                                int1 = int1 + 1
                                                If int1 = 1 Then
                                                    dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                                End If
                                                dgvAS.Rows(Count2).Selected = True

                                                'Exit For

                                            End If
                                        Next
                                        'MsgBox("End selecting AS")
                                        If int1 = 0 Then 'nothing selected
                                        Else
                                            Call Helper2Click()
                                        End If

                                        'now do Helper1Click
                                        If int1 = 0 Or Count3 = 3 Then 'nothing selected
                                        Else
                                            'assign Term1
                                            Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                            Me.dgvHelper1.Rows(intRecSel).Selected = True
                                            Call Helper1Click()
                                        End If

                                    Next Count1

                                Next Count3

                            Case Is = 18 'Summary of [Period Temp] Stability in Matrix
                                'Case Is = 19 'Summary of Freeze/Thaw [#Cycles] Stability in Matrix
                                'Case Is = 21 '[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
                            Case Is = 22 '[Period Temp] Stock Solution Stability Assessment

                                'enter Stock Solution
                                var2 = arrASP(intASPM + 2) ' StockSolutionConc
                                str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))
                                If Len(str1) = 0 Then
                                    boolDoHelper2 = False
                                Else
                                    boolDoHelper2 = True
                                    Me.txtHelper2.Text = str1
                                End If

                                'must assign Term1
                                For Count1 = 1 To 2
                                    Select Case Count1
                                        Case 1
                                            var2 = arrASP(intASPM) ' Old Value
                                            intRecSel = 0
                                        Case 2
                                            var2 = arrASP(intASPM + 1) ' New Value
                                            intRecSel = 1
                                    End Select

                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True
                                        End If
                                    Next
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()
                                        If boolDoHelper2 Then
                                            Call Helper2Click() 'assign stock solution conc
                                        End If


                                    End If

                                Next

                            Case Is = 23 '[Period Temp] Spiking Solution Stability Assessment

                                'must assign Term1
                                For Count1 = 1 To 2
                                    Select Case Count1
                                        Case 1
                                            var2 = arrASP(intASPM) ' Old Value
                                            intRecSel = 0

                                        Case 2
                                            var2 = arrASP(intASPM + 1) ' New Value
                                            intRecSel = 1

                                    End Select

                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True
                                        End If
                                    Next
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()
                                    End If

                                Next

                            Case Is = 29 '[Period Temp] Long-Term QC Std Storage Stability

                                'must assign Term1
                                For Count1 = 1 To 2
                                    Select Case Count1
                                        Case 1
                                            var2 = arrASP(intASPM) ' Old Value
                                            intRecSel = 0

                                            'enter Run Identifier
                                            var3 = arrASP(intASPM + 2) '
                                            str1 = StripNOT(NZ(rowsASP(0).Item(var3), ""))
                                            If Len(str1) = 0 Then
                                                boolDoHelper2 = False
                                            Else
                                                boolDoHelper2 = True
                                                Me.txtHelper2.Text = str1
                                            End If

                                        Case 2
                                            var2 = arrASP(intASPM + 1) ' New Value
                                            intRecSel = 1

                                            'enter Run Identifier
                                            var3 = arrASP(intASPM + 3) ' 
                                            str1 = StripNOT(NZ(rowsASP(0).Item(var3), ""))
                                            If Len(str1) = 0 Then
                                                boolDoHelper2 = False
                                            Else
                                                boolDoHelper2 = True
                                                Me.txtHelper2.Text = str1
                                            End If
                                    End Select

                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True
                                        End If
                                    Next
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()

                                        If boolDoHelper2 Then
                                            Call Helper2Click()
                                        End If
                                    End If

                                Next

                            Case Is = 30 'Incurred Samples
                            Case Is = 31  'Ad Hoc QC Stability Table

                                '
                                var2 = arrASP(intASPM)
                                str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                'enter Run Identifier
                                var3 = arrASP(intASPM + 1) '
                                str3 = StripNOT(NZ(rowsASP(0).Item(var3), ""))
                                If Len(str1) = 0 Then
                                    boolDoHelper2 = False
                                Else
                                    boolDoHelper2 = True
                                    Me.txtHelper2.Text = str3
                                End If

                                'unselect all dgvAS rows
                                If boolDoHelper2 Then
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True
                                        End If
                                    Next
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        Call Helper2Click()
                                    End If
                                End If


                            Case Is = 32 'Ad Hoc QC Stability Comparison Table

                                'must assign Term1
                                For Count1 = 1 To 4

                                    Select Case Count1
                                        Case 1
                                            var2 = arrASP(intASPM) ' Old Value
                                            intRecSel = 0

                                            'enter Run Identifier
                                            var3 = arrASP(intASPM + 4) '
                                            str1 = StripNOT(NZ(rowsASP(0).Item(var3), ""))
                                            If Len(str1) = 0 Then
                                                boolDoHelper2 = False
                                            Else
                                                boolDoHelper2 = True
                                                Me.txtHelper2.Text = str1
                                            End If

                                        Case 2
                                            var2 = arrASP(intASPM + 1) ' New Value
                                            intRecSel = 1

                                            'enter Run Identifier
                                            var3 = arrASP(intASPM + 5) ' 
                                            str1 = StripNOT(NZ(rowsASP(0).Item(var3), ""))
                                            If Len(str1) = 0 Then
                                                boolDoHelper2 = False
                                            Else
                                                boolDoHelper2 = True
                                                Me.txtHelper2.Text = str1
                                            End If

                                        Case 3 '20190305 LEE:
                                            var2 = arrASP(intASPM + 2) ' New Value 2
                                            intRecSel = 1

                                            'enter Run Identifier
                                            var3 = arrASP(intASPM + 6) ' 
                                            str1 = StripNOT(NZ(rowsASP(0).Item(var3), ""))
                                            If Len(str1) = 0 Then
                                                boolDoHelper2 = False
                                            Else
                                                boolDoHelper2 = True
                                                Me.txtHelper2.Text = str1
                                            End If

                                        Case 4 '20190305 LEE:
                                            var2 = arrASP(intASPM + 3) ' New Value 3
                                            intRecSel = 1

                                            'enter Run Identifier
                                            var3 = arrASP(intASPM + 7) ' 
                                            str1 = StripNOT(NZ(rowsASP(0).Item(var3), ""))
                                            If Len(str1) = 0 Then
                                                boolDoHelper2 = False
                                            Else
                                                boolDoHelper2 = True
                                                Me.txtHelper2.Text = str1
                                            End If
                                    End Select

                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    If Len(str1) = 0 Then
                                    Else
                                        For Count2 = 0 To dgvAS.Rows.Count - 1
                                            str2 = dgvAS("SAMPLENAME", Count2).Value
                                            If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                                int1 = int1 + 1
                                                If int1 = 1 Then
                                                    dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                                End If
                                                dgvAS.Rows(Count2).Selected = True
                                            End If
                                        Next
                                    End If

                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()

                                        If boolDoHelper2 Then
                                            Call Helper2Click()
                                        End If
                                    End If

                                Next

                            Case Is = 33 'System Suitability Table v1

                            Case Is = 35 'Carryover in Individual Lots Table v1

                                'must assign Term1
                                Dim intE As Short
                                If dgvHelper1.Rows.Count = 2 Then
                                    intE = 2
                                Else
                                    intE = 3
                                End If

                                For Count1 = 1 To intE
                                    Select Case Count1
                                        Case 1
                                            var2 = arrASP(intASPM) ' LLOQ
                                            intRecSel = 0
                                        Case 2
                                            If intE = 2 Then
                                                var2 = arrASP(intASPM + 2) ' Blank
                                            Else
                                                var2 = arrASP(intASPM + 1) ' ULOQ
                                            End If
                                            intRecSel = 1
                                        Case 3
                                            var2 = arrASP(intASPM + 2) ' Blank
                                            intRecSel = 2
                                    End Select

                                    str1 = StripNOT(NZ(rowsASP(0).Item(var2), ""))

                                    'unselect all dgvAS rows
                                    dgvAS.ClearSelection()
                                    dgvAS.CurrentCell = Nothing
                                    int1 = 0
                                    For Count2 = 0 To dgvAS.Rows.Count - 1
                                        str2 = dgvAS("SAMPLENAME", Count2).Value
                                        If InStr(1, str2, str1, CompareMethod.Text) > 0 Then
                                            int1 = int1 + 1
                                            If int1 = 1 Then
                                                dgvAS.CurrentCell = dgvAS.Rows(Count2).Cells(intASCell)
                                            End If
                                            dgvAS.Rows(Count2).Selected = True
                                        End If
                                    Next
                                    If int1 = 0 Then 'nothing selected
                                    Else
                                        'assign Term1
                                        Me.dgvHelper1.CurrentCell = dgvHelper1.Rows(intRecSel).Cells(1)
                                        Me.dgvHelper1.Rows(intRecSel).Selected = True
                                        Call Helper1Click()

                                    End If

                                Next

                            Case Is = 36 'Method Trial Back-Calculated Calibration Std Conc v1
                            Case Is = 37 'Method Trial Control and Fortified QC Samples v1
                            Case Is = 38 'Method Trial Incurred Blinded Samples v1

                        End Select

                    End If

nextCountA:

                Next CountA


            End If

nextCountT:

        Next CountT


        'clear all filters
        Call ClearFilters()

        '20160811 LEE: hmm. Leave the last table and analyte selected
        Me.txtdgvReportTablePreviousRow.Text = Me.dgvTables.CurrentRow.Index
        'the 'AssessSampleAssignment' function triggers incorrectly if table and analyte are set back to first

        ''select first table
        'dgvT.CurrentCell = dgvT.Rows(0).Cells(intTCell)
        'dgvT.Rows(0).Selected = True

        ''select first analyte
        'If dgvA.Rows.Count = 0 Then
        'Else
        '    dgvA.CurrentCell = dgvA.Rows(0).Cells(intACell)
        '    dgvA.Rows(0).Selected = True
        '    'execute change
        '    Call dgvAnalyteSelectionChange()
        'End If


        '*****



        'These tables do not need nominal concentrations (so no need for Assay Levels)
        '22 Stock Solution Stability Assessment
        '30 Incurred Samples
        '33 System Suitability Table
        '38 Method Trial Incurred Blinded Samples

        Select Case idCTLast
            Case 22, 30, 33, 35, 38 'No warning needed
            Case Else

                'do a conc levels check
                '20151104 LEE:
                'now establish boolGotConcLevels
                Dim boolGotConcLevels As Boolean
                Dim boolGotSomeLevels As Boolean
                Dim row As DataGridViewRow
                boolGotConcLevels = False
                boolGotSomeLevels = False

                Select Case idCTLast
                    Case 13, 14, 15

                        If boolIsIntStd And BOOLISCOMBINELEVELS Then
                            Dim dv As DataView = dgvAS.DataSource
                            For Count1 = 0 To dv.Count - 1
                                var1 = dv(Count1).Item("NOMCONC")
                                If IsDBNull(var1) Then
                                    boolGotSomeLevels = True
                                Else
                                    boolGotConcLevels = True
                                    If boolGotConcLevels And boolGotSomeLevels Then
                                        Exit For
                                    End If
                                End If
                            Next
                        Else
                            For Each row In dgvAS.Rows
                                var1 = row.Cells("ASSAYLEVEL").Value
                                If IsDBNull(var1) Then
                                    boolGotSomeLevels = True
                                Else
                                    boolGotConcLevels = True
                                    If boolGotConcLevels And boolGotSomeLevels Then
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                    Case Else
                        For Each row In dgvAS.Rows
                            var1 = row.Cells("ASSAYLEVEL").Value
                            If IsDBNull(var1) Then
                                boolGotSomeLevels = True
                            Else
                                boolGotConcLevels = True
                                If boolGotConcLevels And boolGotSomeLevels Then
                                    Exit For
                                End If
                            End If
                        Next
                End Select

                If dgvAS.RowCount = 0 Then
                Else
                    If boolGotConcLevels Then
                        If boolGotSomeLevels Then
                            strM = "Please note that some added samples do not contain an Assay Level in the Assay Level column."
                            strM = strM & ChrW(10) & ChrW(10)
                            strM = strM & "User must manually assign Nominal Concentrations in the Assigned Samples table."
                            MsgBox(strM, vbInformation, "No Assay Levels...")
                        End If

                    Else
                        strM = "Please note that the added samples do not have an Assay Level in the Assay Level column."
                        strM = strM & ChrW(10) & ChrW(10)
                        strM = strM & "User must manually assign Nominal Concentrations in the Assigned Samples table."
                        MsgBox(strM, vbInformation, "No Assay Levels...")
                    End If
                End If
                
        End Select

        '*****


        ''Don't save!
        ''Let user save
        'If Me.cmdEdit.Enabled Then
        'Else
        '    Call DoThis("Save")
        'End If

        boolAutoAssign = False

        Me.lblProgress.Visible = False


end1:

        strM = "Auto-Assign Samples completed."
        MsgBox(strM, vbInformation, "Completed...")


    End Sub

    Private Sub dgvAnalyticalRuns_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAnalyticalRuns.CellContentClick

    End Sub

    Private Sub cbxChooseAnalyte_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxChooseAnalyte.SelectedIndexChanged

        If boolFormLoad Then
            GoTo end1
        End If

        Call FillAnalyticalRuns("")

end1:

    End Sub

    Private Sub dgvTables_MultiSelectChanged(sender As Object, e As EventArgs) Handles dgvTables.MultiSelectChanged

    End Sub
End Class