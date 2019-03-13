Option Compare Text

Public Class frmIncSamplesAssignSamples

    Public gAnalyteIndex As Int32
    Public gAnalyteID As Int32
    Public gMasterAssayID As Int32
    Public gidRT As Int32
    Public arrID(0)
    Public boolHold As Boolean = False
    Public boolCancel As Boolean = True
    Public tblISRAssSamples As New System.Data.DataTable

    Sub PlaceControls()

        Dim dgv1 As Panel = Me.panDesignSample
        Dim dgv2 As Panel = Me.panO
        Dim dgv3 As Panel = Me.panISRdata
        Dim dgv4 As Panel = Me.panISRi


        Dim w, h
        Dim w1, h1
        Dim tOff, woff

        w = Me.panISR.Width
        h = Me.panISR.Height

        tOff = Me.panSave.Height
        woff = 20

        w1 = (w - woff) / 2
        'h1 = ((h / 2) - tOff) - Me.panISRbuttons.Height
        'h1 = ((h - Me.panISRbuttons.Height) / 2) - tOff
        h1 = (h / 2) - tOff
        h1 = (h - tOff) / 2

        dgv1.Top = tOff
        dgv1.Left = 0
        dgv1.Width = w1
        dgv1.Height = h1

        'Me.panObuttons.Top = dgv1.Top + dgv1.Height
        'Me.panObuttons.Left = 0
        'Me.panObuttons.Width = w1

        dgv2.Top = dgv1.Top + dgv1.Height
        dgv2.Left = 0
        dgv2.Width = w1
        dgv2.Height = h1

        dgv3.Top = dgv1.Top
        dgv3.Left = dgv1.Left + dgv1.Width + woff
        dgv3.Width = w1
        dgv3.Height = h1

        'Me.panISRbuttons.Top = Me.panObuttons.Top
        'Me.panISRbuttons.Left = dgv1.Left + dgv1.Width + woff
        'Me.panISRbuttons.Width = w1

        dgv4.Top = dgv2.Top
        dgv4.Left = dgv1.Left + dgv1.Width + woff
        dgv4.Width = w1
        dgv4.Height = h1

        Call AutoSizeColumns()

        'place buttons
        Me.panOButtons.Top = 0
        Me.panOButtons.Left = (dgv2.Width / 2) - (Me.panOButtons.Width / 2)

        Me.panISRButtons.Top = 0
        Me.panISRButtons.Left = (dgv4.Width / 2) - (Me.panISRButtons.Width / 2)

        Me.lbltxtNS3.Left = Me.lbltxtNS1.Left
        Me.lbltxtNS3.Top = Me.lbltxtNS1.Top
        Me.lbltxtNS3.BringToFront()

    End Sub

    Sub AutoSizeColumns()

        Me.dgvDesignSample.AutoResizeColumns()
        Me.dgvIncSamplesOrig.AutoResizeColumns()
        Me.dgvIncSamplesISR.AutoResizeColumns()
        Me.dgvAllInjections.AutoResizeColumns()


    End Sub

    Private Sub frmIncSamplesAssignSamples_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

        Try
            Call OrderDGV1(Me.dgvDesignSample)
        Catch ex As Exception

        End Try

        Try
            Call OrderDGV1(Me.dgvIncSamplesOrig)
        Catch ex As Exception

        End Try

        Try
            Call OrderDGV1(Me.dgvIncSamplesISR)
        Catch ex As Exception

        End Try

        Try
            Call OrderDGV1(Me.dgvAllInjections)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub frmIncSamplesAssignSamples_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.cmdAddOrig.Text = ChrW(8595) 'down
        Me.cmdAddISR.Text = ChrW(8595) 'down

        Me.cmdRemoveO.Text = ChrW(8593) 'up
        Me.cmdRemoveISR.Text = ChrW(8593) 'up


        Dim w, h

        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        Me.Left = w - (0.95 * w)
        Me.Top = h - (0.95 * h)
        Me.Width = w * 0.9
        Me.Height = h * 0.9

        'Call CreateTableISRAS()


    End Sub

    Private Sub frmIncSamplesAssignSamples_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize

        'some crap won't size properly
        Dim dgv1 As DataGridView = Me.dgvIncSamplesOrig
        Dim dgv2 As DataGridView = Me.dgvIncSamplesISR

        Dim h1, h2

        h1 = Me.panO.Height

        h2 = h1 - dgv1.Top - 10

        dgv1.Height = h2
        dgv2.Height = h2


        Call PlaceControls()

    End Sub

    Sub FormLoad()
        'load grids

        Cursor.Current = Cursors.WaitCursor

        Try

            'do Original Design Sample
            Call LoadDesignSampleOrig()
            Call LoadISRSource()

            Call AutoSizeColumns()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        Cursor.Current = Cursors.Default

    End Sub

    Sub LoadISRSource()


        Dim dgv1 As DataGridView = Me.dgvAllInjections

        Dim strF As String
        Dim strS As String
        Dim Count1 As Int16
        Dim str1 As String
        Dim str2 As String
        Dim var1, var2, var3
        Dim boolDo As Boolean
        Dim rows() As DataRow
        Dim dgv2 As DataGridView


        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

        Cursor.Current = Cursors.WaitCursor

        Dim dv As System.Data.DataView

        If Me.rbSourceAll.Checked Then

            Dim tbl As System.Data.DataTable = tblAnalysisResultsHome

            strF = "AnalyteID = " & gAnalyteID & " AND ANALYTEINDEX  = " & gAnalyteIndex & " AND MASTERASSAYID  = " & gMasterAssayID & " AND ELIMINATEDFLAG = 'N' AND DESIGNSAMPLEID IS NOT NULL"
            strS = "DESIGNSAMPLEID ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            'strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            boolDo = False

            dv = New DataView(tbl, strF, strS, DataViewRowState.CurrentRows)

        ElseIf Me.rbSourceMatched.Checked Then

            Try
                Call FillISR1()
            Catch ex As Exception

            End Try

            Dim tbl As System.Data.DataTable = tblISRUnique

            Dim idDS As Int64
            Try
                idDS = Me.dgvIncSamplesOrig("DESIGNSAMPLEID", Me.dgvIncSamplesOrig.CurrentRow.Index).Value

            Catch ex As Exception
                idDS = 0
            End Try

            strF = "AnalyteID = " & gAnalyteID & " AND DESIGNSAMPLEID = " & idDS & " AND ANALYTEINDEX  = " & gAnalyteIndex & " AND MASTERASSAYID  = " & gMasterAssayID
            strS = "DESIGNSAMPLEID ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

            'get runid and sequence from dgv
            Try
                var1 = Me.dgvIncSamplesOrig("RUNID", Me.dgvIncSamplesOrig.CurrentRow.Index).Value
                var2 = Me.dgvIncSamplesOrig("RUNSAMPLEORDERNUMBER", Me.dgvIncSamplesOrig.CurrentRow.Index).Value
                str1 = " AND (RUNID <> " & var1 & " AND RUNSAMPLEORDERNUMBER <> " & var2 & ")"
                strF = strF & str1
            Catch ex As Exception

            End Try


            ''''''''console.writeline(strF)
            dv = New DataView(tblISR, strF, strS, DataViewRowState.CurrentRows)

        ElseIf Me.rbSourceRepeated.Checked Then
            Call FillISR1()
            Dim tbl As System.Data.DataTable = tblISRUnique
            strF = "AnalyteID = " & gAnalyteID & " AND ANALYTEINDEX  = " & gAnalyteIndex & " AND MASTERASSAYID  = " & gMasterAssayID
            strS = "DESIGNSAMPLEID ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            'strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            boolDo = False
            For Count1 = 1 To UBound(arrID)
                boolDo = True
                var1 = arrID(Count1) 'tblISRUnique.Rows(Count1).Item("DESIGNSAMPLEID")
                str1 = "DESIGNSAMPLEID = " & var1
                If Count1 = 1 Then
                    strF = strF & " AND (" & str1
                Else
                    strF = strF & " OR " & str1
                End If

            Next
            If boolDo Then
                strF = strF & ")"
            End If
            ''''''''console.writeline(strF)
            dv = New DataView(tblISR, strF, strS, DataViewRowState.CurrentRows)

        ElseIf Me.rbAllMatchedSamples.Checked Then
            'Dim tbl as System.Data.Datatable = tblISRUnique
            dgv2 = Me.dgvIncSamplesOrig
            strS = "DESIGNSAMPLEID ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            If dgv2.RowCount = 0 Then
                strF = "DESIGNSAMPLEID <0"
            Else
                Call FillISR1()
                strF = "AnalyteID = " & gAnalyteID & " AND ANALYTEINDEX  = " & gAnalyteIndex & " AND MASTERASSAYID  = " & gMasterAssayID
                boolDo = False
                For Count1 = 0 To dgv2.RowCount - 1
                    boolDo = True
                    var1 = dgv2("RUNID", Count1).Value
                    var2 = dgv2("RUNSAMPLEORDERNUMBER", Count1).Value
                    var3 = dgv2("DESIGNSAMPLEID", Count1).Value

                    str1 = "(DESIGNSAMPLEID = " & var3 & " AND (RUNID <> " & var1 & " AND RUNSAMPLEORDERNUMBER <> " & var2 & "))"
                    If Count1 = 0 Then
                        strF = strF & " AND (" & str1
                    Else
                        strF = strF & " OR " & str1
                    End If
                Next
                If boolDo Then
                    strF = strF & ")"
                End If
            End If

            ''''''''console.writeline(strF)
            dv = New DataView(tblISR, strF, strS, DataViewRowState.CurrentRows)
            ''''''''console.writeline(strF)
        End If

        dv.AllowEdit = False
        dv.AllowDelete = False
        dv.AllowNew = False
        dgv1.DataSource = dv

        Me.txtNS3.Text = Format(dv.Count, "#,##0")

        'Dim boolGo As Boolean = True
        ''check for any visible rows
        'For Count1 = 0 To dgv1.Columns.Count - 1
        '    If dgv1.Columns(Count1).Visible = False Then
        '        boolGo = False
        '        Exit For
        '    End If
        'Next

        'If boolGo Then
        'Else
        '    GoTo end1
        'End If

        For Count1 = 0 To dgv1.Columns.Count - 1
            dgv1.Columns(Count1).Visible = False
        Next

        Dim intIndex As Short = -1

        Try
            str1 = "DESIGNSAMPLEID"
            str2 = "ID"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            str1 = "SAMPLENAME"
            str2 = "Sample Name"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            str1 = "RUNID"
            str2 = "Run ID"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            str1 = "RUNSAMPLEORDERNUMBER"
            str2 = "Inj.#"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            str1 = "ALIQUOTFACTOR"
            str2 = "DilF"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            'dgv1.Columns(str1).DefaultCellStyle.Format = "0.00"

            'str1 = "CONCENTRATION"
            str1 = "REPORTEDCONC"
            str2 = "Corrected Conc."
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            dgv1.Columns(str1).DefaultCellStyle.Format = "0.000"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv1.AutoResizeColumns()
            dgv1.RowHeadersWidth = 20

            'run displayindex again
            intIndex = -1

            str1 = "DESIGNSAMPLEID"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "SAMPLENAME"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "RUNID"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "RUNSAMPLEORDERNUMBER"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "ALIQUOTFACTOR"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "REPORTEDCONC"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

end1:

            Call OrderDGV1(dgv1)

            Call SyncRows()

            dgv1.AutoResizeColumns()
        Catch ex As Exception

        End Try

        Cursor.Current = Cursors.Default


    End Sub

    Sub LoadDesignSampleOrig()

        Dim dgv1 As DataGridView = Me.dgvDesignSample
        Dim strF As String
        Dim strS As String
        Dim Count1 As Int16
        Dim str1 As String
        Dim str2 As String
        Dim var1, var2
        Dim boolDo As Boolean

        strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

        Cursor.Current = Cursors.WaitCursor

        Dim dv1 As System.Data.DataView '= New DataView(tblSampleDesign, strF, strS, DataViewRowState.CurrentRows)

        If Me.rbAllO.Checked Then
            strF = "AnalyteID = " & gAnalyteID
            strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            dv1 = New DataView(tblSampleDesign, strF, strS, DataViewRowState.CurrentRows)
        Else
            Call FillISR1()
            Dim tbl As System.Data.DataTable = tblISRUnique
            strF = "AnalyteID = " & gAnalyteID ' & " AND ANALYTEINDEX  = " & gAnalyteIndex & " AND MASTERASSAYID  = " & gMasterAssayID
            'strS = "DESIGNSAMPLEID ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            strS = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"
            boolDo = False
            For Count1 = 1 To UBound(arrID)
                boolDo = True
                var1 = arrID(Count1) 'tblISRUnique.Rows(Count1).Item("DESIGNSAMPLEID")
                str1 = "DESIGNSAMPLEID = " & var1
                If Count1 = 1 Then
                    strF = strF & " AND (" & str1
                Else
                    strF = strF & " OR " & str1
                End If
            Next
            If boolDo Then
                strF = strF & ")"
            End If
            ''''''''console.writeline(strF)
            dv1 = New DataView(tblSampleDesign, strF, strS, DataViewRowState.CurrentRows)
        End If
        dv1.AllowDelete = False
        dv1.AllowEdit = False
        dv1.AllowNew = False
        dgv1.DataSource = dv1

        Dim num1 As Int16
        num1 = dv1.Count
        Me.txtNS1.Text = Format(num1, "#,##0")

        Dim boolGo As Boolean = True
        'check for any visible rows
        For Count1 = 0 To dgv1.Columns.Count - 1
            If dgv1.Columns(Count1).Visible = False Then
                boolGo = False
                Exit For
            End If
        Next

        If boolGo Then
        Else
            Call OrderDGV1(dgv1)
            GoTo end1
        End If

        For Count1 = 0 To dgv1.ColumnCount - 1
            dgv1.Columns(Count1).Visible = False
        Next

        Dim intIndex As Short = -1

        Try
            str1 = "DESIGNSAMPLEID"
            str2 = "ID"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            str1 = "SAMPLENAME"
            str2 = "Sample Name"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            str1 = "RUNID"
            str2 = "Run ID"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            str1 = "RUNSAMPLEORDERNUMBER"
            str2 = "Inj.#"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            str1 = "ALIQUOTFACTOR"
            str2 = "DilF"
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
            'dgv1.Columns(str1).DefaultCellStyle.Format = "0.00"

            'str1 = "CONCENTRATION"
            str1 = "REPORTEDCONC"
            str2 = "Corrected Conc."
            dgv1.Columns(str1).HeaderText = str2
            dgv1.Columns(str1).Visible = True
            dgv1.Columns(str1).DefaultCellStyle.Format = "0.000"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex
            dgv1.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
            dgv1.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            dgv1.AutoResizeColumns()
            dgv1.RowHeadersWidth = 20

            'run displayindex again
            intIndex = -1

            str1 = "DESIGNSAMPLEID"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "SAMPLENAME"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "RUNID"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "RUNSAMPLEORDERNUMBER"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "ALIQUOTFACTOR"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex

            str1 = "REPORTEDCONC"
            intIndex = intIndex + 1
            dgv1.Columns(str1).DisplayIndex = intIndex



            Call OrderDGV1(dgv1)
            Call OrderDGV1(dgv1)
            Call OrderDGV1(dgv1)



        Catch ex As Exception

        End Try

end1:

        Try
            Call OrderDGV1(dgv1)
            Call SyncRows()
        Catch ex As Exception

        End Try
        Cursor.Current = Cursors.Default

        'GENDERID()
        'ANALYTEID()
        'DESIGNSUBJECTID()
        'DESIGNSUBJECTTAG()
        'SUBJECTGROUPNAME()
        'ENDDAY()
        'ENDHOUR()
        'ENDMINUTE()
        'ENDSECOND()
        'CONCENTRATION()
        'RUNID()
        'DESIGNSAMPLEID()
        'ALIQUOTFACTOR()
        'TREATMENTID()
        'TREATMENTDESC()
        'TIMETEXT()
        'STARTDAY()
        'STARTHOUR()
        'STARTMINUTE()
        'STARTSECOND()
        'SAMPLETYPEID()
        'SAMPLETYPEKEY()
        'AnalyteDescription()
        'SERIALENDTIME()
        'SERIALSTARTTIME()
        'RUNSAMPLEORDERNUMBER
        'SAMPLENAME
        'STUDYID
        'REPORTEDCONC

    End Sub

    Sub OrderDGV1(ByVal dgv1 As DataGridView)

        Dim strF As String
        Dim strS As String
        Dim Count1 As Int16
        Dim str1 As String
        Dim str2 As String
        Dim var1, var2
        Dim boolDo As Boolean
        Dim intIndex As Short = -1

        str1 = "DESIGNSAMPLEID"
        intIndex = intIndex + 1
        dgv1.Columns(str1).DisplayIndex = intIndex

        str1 = "SAMPLENAME"
        intIndex = intIndex + 1
        dgv1.Columns(str1).DisplayIndex = intIndex

        str1 = "RUNID"
        intIndex = intIndex + 1
        dgv1.Columns(str1).DisplayIndex = intIndex

        str1 = "RUNSAMPLEORDERNUMBER"
        intIndex = intIndex + 1
        dgv1.Columns(str1).DisplayIndex = intIndex

        str1 = "ALIQUOTFACTOR"
        intIndex = intIndex + 1
        dgv1.Columns(str1).DisplayIndex = intIndex

        str1 = "REPORTEDCONC"
        intIndex = intIndex + 1
        dgv1.Columns(str1).DisplayIndex = intIndex

    End Sub

    Sub FillISR1()

        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim Count3 As Int16
        Dim Count4 As Int16
        Dim ctSampleDesign As Int16
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS As String
        Dim var1, var2, var3, var4, var5
        Dim ctSamples As Int16 = 0
        Dim int1 As Int16
        Dim int2 As Int16
        Dim intCol As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int123 As Short
        Dim arrOrder(1, 1)
        Dim ctCols As Short
        'arrID

        Dim rowsISR() As DataRow

        ReDim arrOrder(10, 10)
        'arrOrder
        ' 1=ColumnLabel, 2=Order, 3=id_tblConfigReportTables, 4=UserLabel, 5=id_tblConfigHeaderLookup
        '6=CHARWATSONFIELD

        strF = "ANALYTEID = " & gAnalyteID
        strF = strF & " AND ANALYTEINDEX = " & gAnalyteIndex
        strF = strF & " AND MASTERASSAYID = " & gMasterAssayID
        strS = "DESIGNSUBJECTTREATMENTKEY ASC"

        Dim dvISR As System.Data.DataView = New DataView(tblISR, strF, strS, DataViewRowState.CurrentRows)

        ctSampleDesign = tblISRUnique.Rows.Count

        ReDim arrID(ctSampleDesign)

        For Count2 = 0 To ctSampleDesign - 1

            Dim numConcO As Double
            Dim numConcR As Double
            Dim numConcOAF As Single
            Dim numConcRAF As Single

            Dim intDSTK As Int32

            var1 = tblISRUnique.Rows(Count2).Item("DESIGNSAMPLEID")
            intDSTK = var1

            Erase rowsISR
            'strF1 = strF & " AND DESIGNSAMPLEID = " & intDSTK & " AND ELIMINATEDFLAG = 'N'"
            strF1 = strF & " AND DESIGNSAMPLEID = " & intDSTK

            rowsISR = tblISR.Select(strF1, "RUNID DESC") 'last row is ISR

            If rowsISR.Length < 2 Then
                GoTo nextCount2
            End If

            'get DESIGNSAMPLEID
            var1 = rowsISR(0).Item("DESIGNSAMPLEID")
            var2 = rowsISR(0).Item("RUNSAMPLESEQUENCENUMBER")
            var3 = rowsISR(0).Item("RUNID")
            var4 = rowsISR(0).Item("ANALYTEID")

            If var1 = 569 Then
                var1 = var1
            End If

            Dim rowsConflict() As DataRow
            Dim boolCHit As Boolean = False
            For Count3 = 0 To rowsISR.Length - 1
                var1 = rowsISR(Count3).Item("DESIGNSAMPLEID")
                var2 = rowsISR(Count3).Item("RUNSAMPLESEQUENCENUMBER")
                var3 = rowsISR(Count3).Item("RUNID")
                var4 = rowsISR(Count3).Item("ANALYTEID")
                strF1 = "DESIGNSAMPLEID = " & var1 & " AND RUNSAMPLESEQUENCENUMBER = " & var2 & " AND RUNID = " & var3 & " AND ANALYTEID = " & var4
                Erase rowsConflict
                rowsConflict = tblSAMPLERESULTSCONFLICT.Select(strF1)
                If rowsConflict.Length > 0 Then
                    boolCHit = True
                    Exit For
                End If
            Next

            If boolCHit Then
            Else
                'check for eliminatedflag
                'get DESIGNSAMPLEID
                var1 = rowsISR(0).Item("DESIGNSAMPLEID")
                var2 = rowsISR(0).Item("RUNSAMPLESEQUENCENUMBER")
                var3 = rowsISR(0).Item("RUNID")
                var4 = rowsISR(0).Item("ANALYTEID")
                strF1 = "RUNSAMPLESEQUENCENUMBER = " & var2 & " AND RUNID = " & var3 & " AND ANALYTEID = " & var4
                strF1 = strF & " AND RUNSAMPLESEQUENCENUMBER = " & var2 & " AND RUNID = " & var3
                Dim rowsTT() As DataRow
                rowsTT = tblAnalysisResultsHome.Select(strF1)
                var5 = rowsTT.Length
                var5 = rowsTT(0).Item("ELIMINATEDFLAG")
                If StrComp(var5, "Y", CompareMethod.Text) = 0 Then
                    boolCHit = True
                End If

            End If

            Dim rowsSD() As DataRow
            'If rowsConflict.Length = 0 Then ' use
            If boolCHit Then
            Else ' use
                ctSamples = ctSamples + 1
                'get original value
                var1 = rowsISR(0).Item("DESIGNSAMPLEID")
                arrID(ctSamples) = var1
            End If

nextCount2:

        Next Count2

        ReDim Preserve arrID(ctSamples)

    End Sub

    Private Sub rbAllO_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAllO.CheckedChanged

        Call LoadDesignSampleOrig()

    End Sub

    Private Sub cmdAddOrig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddOrig.Click

        Call AddRows(True)

    End Sub

    Sub RemoveRows(ByVal boolFromO As Boolean)

        Dim dgv1 As DataGridView
        Dim row As DataGridViewRow
        Dim str1 As String
        Dim strF As String
        Dim arrD(1) As String
        Dim intRows As Short
        Dim int1 As Short = 0
        Dim Count1 As Short
        Dim Count2 As Short
        Dim bool As Boolean

        bool = boolCont

        boolCont = False

        Dim rowsNew() As DataRow
        Dim strType As String = "O"

        If boolFromO Then
            dgv1 = Me.dgvIncSamplesOrig
            strType = "O"
        Else
            dgv1 = Me.dgvIncSamplesISR
            strType = "ISR"
        End If

        intRows = dgv1.SelectedRows.Count

        ReDim arrD(intRows)

        For Each row In dgv1.SelectedRows
            int1 = int1 + 1
            str1 = "DESIGNSAMPLEID = " & row.Cells("DESIGNSAMPLEID").Value
            str1 = str1 & " AND ANALYTEINDEX = " & gAnalyteIndex
            str1 = str1 & " AND ANALYTEID = " & gAnalyteID
            str1 = str1 & " AND MASTERASSAYID = " & gMasterAssayID
            'remove rows from both tables
            If boolFromO Then
            Else
                str1 = str1 & " AND SAMPLENAME = '" & row.Cells("SAMPLENAME").Value & "'"
                str1 = str1 & " AND RUNSAMPLEORDERNUMBER = " & row.Cells("RUNSAMPLEORDERNUMBER").Value
                str1 = str1 & " AND RUNID = " & row.Cells("RUNID").Value
                str1 = str1 & " AND CHARTYPE = '" & strType & "'"
            End If
            strF = str1

            arrD(int1) = strF
            'Erase rowsNew
            'rowsNew = tblISRAssSamples.Select(strF)

        Next

        For Count1 = 1 To intRows
            strF = arrD(Count1)
            Erase rowsNew
            rowsNew = tblISRAssSamples.Select(strF)

            For Count2 = 0 To rowsNew.Length - 1
                Try
                    rowsNew(Count2).Delete()
                Catch ex As Exception

                End Try
            Next

        Next

        Dim int2 As Int32
        int2 = dgv1.Rows.Count
        If boolFromO Then
            Me.txtNS2.Text = Format(dgv1.Rows.Count, "#,##0")
            Call LoadISRSource()
        Else
            Me.txtNS4.Text = Format(dgv1.Rows.Count, "#,##0")
        End If

        Call ConfigdgvIncSamplesISR()

        Call SyncRows()

        boolCont = True

    End Sub

    Sub AddRows(ByVal boolFromO As Boolean)

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim row As DataGridViewRow
        Dim intRows As Short
        Dim Count1 As Short
        Dim int2 As Short
        Dim var1, var2
        Dim str1 As String
        Dim dv As System.Data.DataView
        Dim intT1 As Short
        Dim intT2 As Short
        Dim boolA As Boolean
        Dim maxID
        Dim maxID1

        Dim drowsmaxid() As DataRow
        Dim intIS As Short
        Dim str2 As String
        Dim bool As Boolean
        Dim rowSel(10, 1)
        Dim intSel As Short
        Dim ct1 As Short
        Dim idRTS As Long
        Dim idConfigRT As Long
        Dim idAnalyte As String
        Dim int1 As Short
        Dim idRT As Long

        bool = boolCont
        boolCont = False 'do this to stop dgvAssignedSamples selectionchanged event

        Dim strType As String = "O"

        If boolFromO Then
            dgv1 = Me.dgvDesignSample
            dgv2 = Me.dgvIncSamplesOrig
            strType = "O"
        Else
            dgv1 = Me.dgvAllInjections
            dgv2 = Me.dgvIncSamplesISR
            strType = "ISR"
        End If

        intRows = dgv1.RowCount

        dv = dgv2.DataSource
        dv.AllowNew = True

        ''MsgBox(dv.RowFilter)
        'str1 = dv.RowFilter

        idRTS = id_tblStudies
        idConfigRT = 30 ' Me.dgvTables.Rows.Item(intT1).Cells("id_tblConfigReportTables").Value
        idRT = gidRT ' Me.dgvTables.Rows.Item(intT1).Cells("ID_TBLREPORTTABLE").Value
        idAnalyte = gAnalyteID ' Me.dgvAnalytes.Rows.Item(intT2).Cells(0).Value

        ''determine if selected analyte is internal standard
        'str2 = Me.dgvAnalytes("IsIntStd", intT2).Value
        'If StrComp(str2, "Yes", CompareMethod.Text) = 0 Then
        '    intIS = -1
        'Else
        '    intIS = 0
        'End If

        intSel = dgv1.SelectedRows.Count
        ReDim rowSel(10, intSel)

        Dim strF As String
        Dim rowsNew() As DataRow
        Dim strS As String = "RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

        ct1 = 0
        Dim ctRowSel As Short

        Cursor.Current = Cursors.WaitCursor

        Dim strM As String = ""

        For Each row In dgv1.SelectedRows

            'first determine if data already exists
            str1 = "DESIGNSAMPLEID = " & row.Cells("DESIGNSAMPLEID").Value
            str1 = str1 & " AND SAMPLENAME = '" & row.Cells("SAMPLENAME").Value & "'"
            str1 = str1 & " AND RUNID = " & row.Cells("RUNID").Value
            str1 = str1 & " AND RUNSAMPLEORDERNUMBER = " & row.Cells("RUNSAMPLEORDERNUMBER").Value
            str1 = str1 & " AND ANALYTEINDEX = " & gAnalyteIndex
            str1 = str1 & " AND ANALYTEID = " & gAnalyteID
            str1 = str1 & " AND MASTERASSAYID = " & gMasterAssayID
            str1 = str1 & " AND CHARTYPE = '" & strType & "'"
            strF = str1
            Erase rowsNew
            rowsNew = tblISRAssSamples.Select(strF)


            Dim boolGo As Boolean = True
            If rowsNew.Length = 0 Then 'data doesn't exist, continue

                If boolFromO Then
                Else 'make sure matching item exists i in original
                    Dim dvO As System.Data.DataView = Me.dgvIncSamplesOrig.DataSource
                    Dim tblO As System.Data.DataTable = dvO.ToTable
                    Dim strFo As String
                    strFo = "DESIGNSAMPLEID = " & row.Cells("DESIGNSAMPLEID").Value
                    Dim ro() As DataRow = tblO.Select(strFo)
                    If ro.Length = 0 Then 'don't allow
                        strM = "Warning!! The chosen sample does not "
                        boolGo = False
                    Else
                        'make sure exact match doesn't exist item exists i in original
                        Erase ro
                        ro = tblO.Select(strF)
                        If ro.Length = 0 Then 'allow
                        Else
                            boolGo = False
                        End If
                    End If

                End If

                If boolGo Then
                    ct1 = ct1 + 1

                    intRows = intRows + 1
                    boolCont = False 'to disable FindSamples
                    Dim dvRow As DataRow = tblISRAssSamples.NewRow

                    dvRow.BeginEdit()

                    dvRow("DESIGNSAMPLEID") = row.Cells("DESIGNSAMPLEID").Value
                    dvRow("SAMPLENAME") = row.Cells("SAMPLENAME").Value
                    dvRow("RUNID") = row.Cells("RUNID").Value
                    dvRow("RUNSAMPLEORDERNUMBER") = row.Cells("RUNSAMPLEORDERNUMBER").Value
                    dvRow("ALIQUOTFACTOR") = row.Cells("ALIQUOTFACTOR").Value
                    dvRow("REPORTEDCONC") = RoundToDecimalRAFZ(NZ(row.Cells("REPORTEDCONC").Value, 0), 3)
                    dvRow("ANALYTEINDEX") = gAnalyteIndex
                    dvRow("ANALYTEID") = gAnalyteID
                    dvRow("MASTERASSAYID") = gMasterAssayID
                    dvRow("CHARTYPE") = strType
                    dvRow("ID_TBLREPORTTABLE") = gidRT

                    dvRow.EndEdit()
                    tblISRAssSamples.Rows.Add(dvRow)
                End If

            Else
            End If
        Next

        If boolFromO Then
        Else
            'dgv seems not to be 
        End If

        dgv2.Refresh()
        dgv2.AutoResizeColumns()

        Dim num1 As Int32
        num1 = dgv2.Rows.Count ' tblISRAssSamples.Rows.Count
        If boolFromO Then
            Me.txtNS2.Text = Format(num1, "#,##0")
            Call LoadISRSource()
        Else
            Me.txtNS4.Text = Format(num1, "#,##0")
        End If


        'ensure first selected row is near the top of the grid
        If ct1 = 0 Then
        Else

            If intRows = 0 Then
            Else
                Try
                    dgv2.FirstDisplayedScrollingRowIndex = intRows - 1
                Catch ex As Exception

                End Try
            End If
        End If


end1:
        boolCont = bool

        dv.AllowEdit = False

        Call ConfigdgvIncSamplesISR()

        Call SyncRows()


        Cursor.Current = Cursors.Default


        'Me.dgvAssignedSamples.AutoResizeColumns()
    End Sub

    Sub CreateTableISRAS()


        'tblISRAssSamples
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1
        Dim Count1 As Short
        Dim Count2 As Short

        str1 = "DESIGNSAMPLEID"

        Dim dtbl As System.Data.DataTable
        dtbl = tblISRAssSamples

        dtbl.Clear()

        If dtbl.Columns.Contains(str1) Then
        Else

            For Count1 = 1 To 11
                Select Case Count1
                    Case 1
                        str1 = "DESIGNSAMPLEID"
                        str2 = "System.Int64"
                    Case 2
                        str1 = "SAMPLENAME"
                        str2 = "System.String"
                    Case 3
                        str1 = "RUNID"
                        str2 = "System.Int64"
                    Case 4
                        str1 = "RUNSAMPLEORDERNUMBER"
                        str2 = "System.Int64"
                    Case 5
                        str1 = "ALIQUOTFACTOR"
                        str2 = "System.Single"
                    Case 6
                        str1 = "REPORTEDCONC"
                        str2 = "System.Double"
                    Case 7
                        str1 = "ANALYTEID"
                        str2 = "System.Int64"
                    Case 8
                        str1 = "ANALYTEINDEX"
                        str2 = "System.Int64"
                    Case 9
                        str1 = "MASTERASSAYID"
                        str2 = "System.Int64"
                    Case 10
                        str1 = "CHARTYPE"
                        str2 = "System.String"
                    Case 11
                        str1 = "ID_TBLREPORTTABLE"
                        str2 = "System.Int64"
                End Select
                Dim col1 As New DataColumn
                col1.ColumnName = str1
                col1.DataType = System.Type.GetType(str2)
                dtbl.Columns.Add(col1)

            Next

        End If

        Try
            Call ConfigdgvIncSamplesOrig()
        Catch ex As Exception

        End Try


    End Sub

    Sub ConfigdgvIncSamplesOrig()

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1
        Dim Count1 As Short
        Dim Count2 As Short
        Dim dgv As DataGridView
        Dim strF As String
        Dim strS As String

        strF = "ANALYTEID = " & gAnalyteID
        strF = strF & " AND ANALYTEINDEX = " & gAnalyteIndex
        strF = strF & " AND MASTERASSAYID = " & gMasterAssayID
        strS = "ANALYTEID ASC"

        Dim dv1 As System.Data.DataView
        Dim dv2 As System.Data.DataView

        Dim dvA As System.Data.DataView = Me.dgvDesignSample.DataSource


        For Count2 = 1 To 2

            strF = "ANALYTEID = " & gAnalyteID
            strF = strF & " AND ANALYTEINDEX = " & gAnalyteIndex
            strF = strF & " AND MASTERASSAYID = " & gMasterAssayID
            strS = "ANALYTEID ASC"
            strS = dvA.Sort

            Select Case Count2
                Case 1
                    strF = strF & " AND CHARTYPE = 'O'"
                    dgv = dgvIncSamplesOrig
                    dv1 = New DataView(tblISRAssSamples, strF, strS, DataViewRowState.CurrentRows)
                    dv1.AllowEdit = False
                    dv1.AllowNew = False
                    dv1.AllowDelete = False
                    dgv.DataSource = dv1
                Case 2
                    strF = strF & " AND CHARTYPE = 'ISR'"
                    dgv = dgvIncSamplesISR
                    dv2 = New DataView(tblISRAssSamples, strF, strS, DataViewRowState.CurrentRows)
                    dv2.AllowEdit = False
                    dv2.AllowNew = False
                    dv2.AllowDelete = False
                    dgv.DataSource = dv2

                    var1 = dgv.RowCount
                    var1 = var1

            End Select


            Dim boolVis As Boolean
            For Count1 = 1 To 11
                boolVis = True
                var1 = DataGridViewContentAlignment.BottomCenter
                'Note: str2 for reference
                Select Case Count1
                    Case 1
                        str1 = "DESIGNSAMPLEID"
                        str3 = "ID"
                        str2 = "System.Int64"
                    Case 2
                        str1 = "SAMPLENAME"
                        str3 = "Sample Name"
                        str2 = "System.String"
                        var1 = DataGridViewContentAlignment.BottomLeft
                    Case 3
                        str1 = "RUNID"
                        str3 = "Run ID"
                        str2 = "System.Int64"
                    Case 4
                        str1 = "RUNSAMPLEORDERNUMBER"
                        str3 = "Inj.#"
                        str2 = "System.Int64"
                    Case 5
                        str1 = "ALIQUOTFACTOR"
                        str3 = "DilF"
                        str2 = "System.Single"
                    Case 6
                        str1 = "REPORTEDCONC"
                        str3 = "Corrected Conc."
                        str2 = "System.Double"
                        var1 = DataGridViewContentAlignment.BottomRight
                    Case 7
                        str1 = "ANALYTEID"
                        str3 = str1
                        str2 = "System.Int64"
                        boolVis = False
                    Case 8
                        str1 = "ANALYTEINDEX"
                        str3 = str1
                        str2 = "System.Int64"
                        boolVis = False
                    Case 9
                        str1 = "MASTERASSAYID"
                        str3 = str1
                        str2 = "System.Int64"
                        boolVis = False
                    Case 10
                        str1 = "CHARTYPE"
                        str3 = str1
                        str2 = "System.String"
                        boolVis = False
                    Case 11
                        str1 = "ID_TBLREPORTTABLE"
                        str3 = str1
                        str2 = "System.Int64"
                        boolVis = False

                End Select

                dgv.Columns(str1).Visible = boolVis
                dgv.Columns(str1).HeaderText = str3
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                dgv.Columns(str1).DefaultCellStyle.Alignment = var1
                dgv.Columns(str1).AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells

            Next

            'special
            dgv.Columns("REPORTEDCONC").DefaultCellStyle.Format = "0.000"

            dgv.AutoResizeColumns()
            dgv.RowHeadersWidth = 20


            dgv.AutoResizeColumns()

            Select Case Count2
                Case 1
                    Me.txtNS2.Text = Format(dv1.Count, "#,##0")
                Case 2
                    Me.txtNS4.Text = Format(dv2.Count, "#,##0")
            End Select

            Call OrderDGV1(dgv)

        Next

    End Sub

    Private Sub cmdRemoveO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveO.Click

        Call RemoveRows(True)

    End Sub


    Private Sub rbSourceRepeated_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSourceRepeated.CheckedChanged

        If boolHold Then
            Exit Sub
        End If
        boolHold = True
        Call LoadISRSource()
        boolHold = False

    End Sub

    Private Sub rbSourceMatched_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSourceMatched.CheckedChanged

        If boolHold Then
            Exit Sub
        End If
        boolHold = True
        Call LoadISRSource()
        boolHold = False

    End Sub

    Private Sub rbSourceAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSourceAll.CheckedChanged

        If boolHold Then
            Exit Sub
        End If
        boolHold = True
        Call LoadISRSource()
        boolHold = False

    End Sub


    Private Sub dgvIncSamplesOrig_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvIncSamplesOrig.SelectionChanged


        If boolCont Then
        Else
            Exit Sub
        End If

        If Me.rbSourceMatched.Checked Then
            Call LoadISRSource()
        End If

        Call ConfigdgvIncSamplesISR()

        'now select appropriate rows
        Call SyncRows()


    End Sub

    Sub SyncRows()


        Dim intRow As Int32
        Dim idDS As Int64
        Dim idDS1 As Int64
        Dim Count1 As Int32

        Dim dgv1 As DataGridView = Me.dgvDesignSample
        Dim dgv2 As DataGridView = Me.dgvIncSamplesOrig
        Dim dgv3 As DataGridView = Me.dgvAllInjections
        Dim dgv4 As DataGridView = Me.dgvIncSamplesISR

        Try
            idDS = dgv2("DESIGNSAMPLEID", dgv2.CurrentRow.Index).Value
        Catch ex As Exception
            idDS = 0
        End Try

        Cursor.Current = Cursors.WaitCursor

        'SYNC DGV1
        For Count1 = 0 To dgv1.Rows.Count - 1
            idDS1 = dgv1("DESIGNSAMPLEID", Count1).Value
            If idDS1 = idDS Then
                dgv1.CurrentCell = dgv1.Rows.Item(Count1).Cells("DESIGNSAMPLEID")
                dgv1.Rows.Item(Count1).Cells("DESIGNSAMPLEID").Selected = True
                'dgv1.FirstDisplayedScrollingRowIndex = Count1
                Exit For
            End If
        Next

        'SYNC DGV3
        For Count1 = 0 To dgv3.Rows.Count - 1
            idDS1 = dgv3("DESIGNSAMPLEID", Count1).Value
            If idDS1 = idDS Then
                dgv3.CurrentCell = dgv3.Rows.Item(Count1).Cells("DESIGNSAMPLEID")
                dgv3.Rows.Item(Count1).Cells("DESIGNSAMPLEID").Selected = True
                'dgv3.FirstDisplayedScrollingRowIndex = Count1
                Exit For
            End If
        Next

        'SYNC DGV4
        For Count1 = 0 To dgv4.Rows.Count - 1
            idDS1 = dgv4("DESIGNSAMPLEID", Count1).Value
            If idDS1 = idDS Then
                dgv4.CurrentCell = dgv4.Rows.Item(Count1).Cells("DESIGNSAMPLEID")
                dgv4.Rows.Item(Count1).Cells("DESIGNSAMPLEID").Selected = True
                'dgv4.FirstDisplayedScrollingRowIndex = Count1
                Exit For
            End If
        Next

        Cursor.Current = Cursors.Default


    End Sub

    Private Sub cmdAddISR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddISR.Click

        Call AddRows(False)

    End Sub

    Private Sub cmdRemoveISR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveISR.Click

        Call RemoveRows(False)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        'tblISRAssSamples

        Dim Count1 As Short
        Dim Count2 As Short

        Dim var1, var2, var3
        Dim tbl As System.Data.DataTable = tblISRAssSamples

        For Count1 = 0 To tbl.Rows.Count - 1
            var1 = ""
            For Count2 = 0 To tbl.Columns.Count - 1
                var2 = tbl.Rows(Count1).Item(Count2)
                var1 = var1 & ChrW(9) & var2
            Next
            var3 = var3 & ChrW(10) & var1
        Next

        MsgBox(var3)

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        boolCancel = True
        Me.Visible = False

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        'validate data
        If ValidateData() Then 'OK

            boolCancel = False
            Me.Visible = False

        End If


    End Sub

    Function ValidateData() As Boolean

        ValidateData = False

        Dim numO As Int32 = CInt(NZ(Me.txtNS2.Text, 0))
        Dim numISR As Int32 = CInt(NZ(Me.txtNS4.Text, 0))

        Dim strM As String = ""
        Dim boolE As Boolean = True

        Me.rbAllAssigned.Checked = True

        'If numO = numISR Then
        'Else
        '    boolE = True
        '    strM = "The number of Original Observations samples must equal the number of Matched ISR samples"
        'End If

        Dim dv1 As System.Data.DataView = Me.dgvIncSamplesOrig.DataSource
        Dim dv2 As System.Data.DataView = Me.dgvIncSamplesISR.DataSource
        Dim Count1 As Int32
        Dim Count2 As Int32
        Dim idDS1 As Int64
        Dim idDS2 As Int64
        Dim boolHit As Boolean
        Dim arrHits(5, dv1.Count)
        Dim intHits As Int32 = 0
        Dim var1, var2, var3
        Dim dgv1 As DataGridView = Me.dgvIncSamplesOrig

        For Count1 = 0 To dv1.Count - 1
            idDS1 = dv1(Count1).Item("DESIGNSAMPLEID")
            var1 = dv1(Count1).Item("SAMPLENAME")
            boolHit = False
            For Count2 = 0 To dv2.Count - 1
                idDS2 = dv2(Count2).Item("DESIGNSAMPLEID")

                If idDS1 = idDS2 Then
                    boolHit = True
                    Exit For
                End If
            Next
            If boolHit Then
            Else
                intHits = intHits + 1
                arrHits(1, intHits) = idDS1
                arrHits(2, intHits) = NZ(var1, "NA")
                arrHits(3, intHits) = Count1 'row number
            End If
        Next

        If intHits = 0 Then
            'look for stuff in dgv2
            'this actually can't happen because delete rows will delete both O and ISR
            For Count1 = 0 To dv2.Count - 1
                idDS1 = dv2(Count1).Item("DESIGNSAMPLEID")
                var1 = dv2(Count1).Item("SAMPLENAME")
                boolHit = False
                For Count2 = 0 To dv1.Count - 1
                    idDS2 = dv1(Count2).Item("DESIGNSAMPLEID")

                    If idDS1 = idDS2 Then
                        boolHit = True
                        Exit For
                    End If
                Next
                If boolHit Then
                Else
                    intHits = intHits + 1
                    arrHits(1, intHits) = idDS1
                    arrHits(2, intHits) = NZ(var1, "NA")
                End If
            Next
            If intHits = 0 Then
            Else
                strM = "Each Matched ISR samples must have at least one Original Observation sample." & ChrW(10)
                strM = strM & "The following Matched ISR samples do not have matching Original Observation samples:" & ChrW(10) & ChrW(10)
                For Count1 = 1 To intHits
                    var1 = "ID: " & arrHits(1, Count1)
                    var2 = "Sample Name: " & arrHits(2, Count1)
                    If Count1 = 1 Then
                        var3 = var1 & " => " & var2
                    Else
                        var3 = var3 & ChrW(10) & var1 & " => " & var2
                    End If
                Next

                'select offending row
                dgv1.CurrentCell = dgv1("DESIGNSAMPLEID", arrHits(3, 1))

                strM = strM & var3
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                boolE = True
                GoTo end1
            End If

        Else
            strM = "Each Original Observation sample must have at least one Matched ISR sample." & ChrW(10)
            strM = strM & "The following Original Observation samples do not have matching ISR samples:" & ChrW(10) & ChrW(10)
            For Count1 = 1 To intHits
                var1 = "ID: " & arrHits(1, Count1)
                var2 = "Sample Name: " & arrHits(2, Count1)
                If Count1 = 1 Then
                    var3 = var1 & " => " & var2
                Else
                    var3 = var3 & ChrW(10) & var1 & " => " & var2
                End If
            Next
            strM = strM & var3

            'select offending row
            dgv1.CurrentCell = dgv1("DESIGNSAMPLEID", arrHits(3, 1))

            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            boolE = True
            GoTo end1
        End If

        'now look to ensure that


        boolE = False
end1:

        If boolE Then

            ValidateData = False
        Else
            ValidateData = True
        End If


    End Function

    Private Sub rbAssignedMatch_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAssignedMatch.CheckedChanged

        Call ConfigdgvIncSamplesISR()

    End Sub


    Sub ConfigdgvIncSamplesISR()

        Dim dgv1 As DataGridView = Me.dgvIncSamplesOrig
        Dim dgv2 As DataGridView = Me.dgvIncSamplesISR

        Dim dv2 As System.Data.DataView

        Dim strf As String
        Dim strS As String
        Dim dvA As System.Data.DataView = Me.dgvIncSamplesOrig.DataSource
        Dim dvISR As System.Data.DataView = Me.dgvIncSamplesISR.DataSource

        Dim var1, var2, var3

        strf = "ANALYTEID = " & gAnalyteID
        strf = strf & " AND ANALYTEINDEX = " & gAnalyteIndex
        strf = strf & " AND MASTERASSAYID = " & gMasterAssayID
        strS = dvA.Sort

        strf = strf & " AND CHARTYPE = 'ISR'"
        If Me.rbAssignedMatch.Checked Then 'match to selected dgvIncSamplesO
            var1 = dgv1("DESIGNSAMPLEID", dgv1.CurrentRow.Index).Value
            strf = strf & " AND DESIGNSAMPLEID = " & var1
        Else 'show all
        End If

        dgv2 = dgvIncSamplesISR
        dv2 = New DataView(tblISRAssSamples, strf, strS, DataViewRowState.CurrentRows)
        dv2.AllowEdit = False
        dv2.AllowNew = False
        dv2.AllowDelete = False
        dgv2.DataSource = dv2

        var1 = dgv2.RowCount
        Me.txtNS4.Text = var1


    End Sub

End Class