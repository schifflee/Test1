Option Compare Text

Public Class frmAnalyteDetails
    Public boolGoTables As Boolean = False
    Private Sub frmAnalyteDetails_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

        Call DoubleBufferControl(Me, "dgv")

        Dim dgv As DataGridView
        Dim dv1 As System.Data.DataView
        Dim dv2 As System.Data.DataView
        Dim dv3 As System.Data.DataView
        Dim dv4 As System.Data.DataView
        Dim strS As String
        Dim tbl As System.Data.DataTable
        Dim strF As String

        dgv = Me.dgvAnalytes
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        tbl = tblAnalytesHome

        strS = "IsIntStd ASC, AnalyteDescription ASC"
        dv1 = New DataView(tbl, "", strS, DataViewRowState.CurrentRows)
        'strS = "AnalyteDescription ASC"
        'dv1.Sort = strS
        dv1.AllowDelete = False
        dv1.AllowNew = False
        dgv.DataSource = dv1
        Dim col As DataGridViewColumn
        For Each col In dgv.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        dgv = Me.dgvQCs
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        tbl = tblQCStds
        dv2 = New DataView(tbl)
        strS = "AnalyteDescription ASC, LevelNumber ASC, Concentration ASC"
        dv2.Sort = strS
        dv2.AllowDelete = False
        dv2.AllowNew = False
        dgv.DataSource = dv2
        For Each col In dgv.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        'dgv = Me.dgvQCs
        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'tbl = tblQCConcs
        'dv4 = New DataView(tbl)
        'strS = "AnalyteDescription ASC, RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
        'dv4.Sort = strS
        'dv4.AllowDelete = False
        'dv4.AllowNew = False
        'dgv.DataSource = dv4
        'For Each col In dgv.Columns
        '    col.SortMode = DataGridViewColumnSortMode.NotSortable
        'Next

        dgv = Me.dgvQCConcs
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        tbl = tblBCQCConcs
        dv4 = New DataView(tbl)
        strS = "AnalyteDescription ASC, RUNID ASC, RUNSAMPLESEQUENCENUMBER ASC"
        dv4.Sort = strS
        dv4.AllowDelete = False
        dv4.AllowNew = False
        dgv.DataSource = dv4
        dgv.Columns("AnalyteDescription").DisplayIndex = 0
        For Each col In dgv.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        dgv = Me.dgvCalibr
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        tbl = tblBCStds
        dv3 = New DataView(tbl)
        'strS = "LevelNumber ASC"
        'dv3.Sort = strS
        'strS = "AnalyteDescription ASC, LevelNumber ASC"
        'strS = "ANALYTEID, MASTERASSAYID, ANALYTEINDEX, LEVELNUMBER;"

        'dv3.Sort = strS
        dv3.AllowDelete = False
        dv3.AllowNew = False
        dgv.DataSource = dv3
        dgv.Columns("AnalyteDescription").DisplayIndex = 0
        For Each col In dgv.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        dgv = Me.dgvCalibrConcs
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.RowHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgv.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        tbl = tblBCStdConcs
        dv3 = New DataView(tbl)
        strS = "ANALYTEID ASC, RUNID ASC, ASSAYLEVEL ASC, RUNSAMPLESEQUENCENUMBER ASC"

        'strS = "LevelNumber ASC"
        dv3.Sort = strS
        'strS = "AnalyteDescription ASC, LevelNumber ASC"
        'strS = "ANALYTEID, MASTERASSAYID, ANALYTEINDEX, LEVELNUMBER;"

        'dv3.Sort = strS
        dv3.AllowDelete = False
        dv3.AllowNew = False
        dgv.DataSource = dv3
        'dgv.Columns("AnalyteDescription").DisplayIndex = 0
        For Each col In dgv.Columns
            col.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        Call Configure_dgvQCs()
        Call Configure_dgvAnalyte()
        Call Configure_dgvQCConcs()
        Call Configure_dgvCalibr()
        Call Configure_dgvCalibrConcs()
        Call FormatRTF()

    End Sub

    Sub FormatRTF()
        Dim str1 As String

        str1 = "The QC/Calibration Standard feature summarizes the Calibration and QC configurations for the chosen Watson™ study."
        str1 = str1 & ChrW(13) & ChrW(13)
        str1 = str1 & "StudyDoc™ expects a certain level of consistency and conformity in Analyte, Calibration Standard, and QC Standard configuration in the underlying Watson study. For example, StudyDoc expects Analytical Run QC Standard levels and names to be consistent throughout the study. If this is not the case, StudyDoc cannot automatically retrieve the correct QC information from the study, resulting in the inability of StudyDoc to generate, in this case, an Interpolated QC Standard Concentrations table."
        str1 = str1 & ChrW(13) & ChrW(13)
        str1 = str1 & "If while attempting to generate a report the user has received an error message stating that the Sample Results table, Back-Calculated Calibration Standard Concentration table, or Interpolated QC Standard Concentration table cannot be prepared, then the user must:"
        str1 = str1 & ChrW(13) & ChrW(13)
        str1 = str1 & ChrW(9) & "• Activate the Report Table Configuration tab (or click on the 'Go to Report Table Configuration' link above)." & ChrW(13)
        str1 = str1 & ChrW(9) & "• Check the 'A*' column in the appropriate table row." & ChrW(13)
        str1 = str1 & ChrW(9) & "• Click on the ‘Assign Samples’ button." & ChrW(13)
        str1 = str1 & ChrW(9) & "• Manually assign samples for the appropriate table."
        Me.rtb1.Text = str1

    End Sub

    Sub Configure_dgvQCs()
        Dim dgv As DataGridView
        Dim str1 As String

        dgv = Me.dgvQCs

        str1 = "AnalyteDescription"
        dgv.Columns(str1).HeaderText = "Analyte"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        str1 = "LevelNumber"
        dgv.Columns(str1).HeaderText = "Level"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "Concentration"
        dgv.Columns(str1).HeaderText = "Nom. Conc."
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "QCName"
        dgv.Columns(str1).HeaderText = "QC Name"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        str1 = "NumReps"
        dgv.Columns(str1).HeaderText = "# Reps"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "AssayID"
        dgv.Columns(str1).HeaderText = "Run" & ChrW(10) & "Assay ID"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "MasterAssayID"
        dgv.Columns(str1).HeaderText = "Master" & ChrW(10) & "Assay ID"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "ID"
        dgv.Columns(str1).HeaderText = "Analyte" & ChrW(10) & "ID"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        str1 = "Index"
        dgv.Columns(str1).HeaderText = "Analyte" & ChrW(10) & "Index"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        str1 = "FlagPercent"
        dgv.Columns(str1).HeaderText = "Flag" & ChrW(10) & "Percent"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        Call Order_dgvQCs()

    End Sub

    Sub Order_dgvQCs()
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        dgv = Me.dgvQCs
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns(Count1).DisplayIndex = int1 - 1
        Next

        str1 = "AnalyteDescription"
        dgv.Columns(str1).DisplayIndex = 0

        str1 = "LevelNumber"
        dgv.Columns(str1).DisplayIndex = 1

        str1 = "Concentration"
        dgv.Columns(str1).DisplayIndex = 2

        str1 = "QCName"
        dgv.Columns(str1).DisplayIndex = 3

        str1 = "NumReps"
        dgv.Columns(str1).DisplayIndex = 4

        str1 = "AssayID"
        dgv.Columns(str1).DisplayIndex = 6

        str1 = "MasterAssayID"
        dgv.Columns(str1).DisplayIndex = 7

        str1 = "ID"
        dgv.Columns(str1).DisplayIndex = 8

        str1 = "Index"
        dgv.Columns(str1).DisplayIndex = 9

        str1 = "FlagPercent"
        dgv.Columns(str1).DisplayIndex = 5

    End Sub

    Sub Configure_dgvAnalyte()
        Dim dgv As DataGridView
        Dim str1 As String


        dgv = Me.dgvAnalytes

        Dim int1 As Short
        Dim Count1 As Short
        int1 = dgv.ColumnCount
        For Count1 = 0 To int1 - 1
            dgv.Columns(Count1).Visible = False
        Next

        str1 = "AnalyteDescription"
        dgv.Columns(str1).HeaderText = "Analyte"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv.Columns(str1).Visible = True

        str1 = "ORIGINALANALYTEDESCRIPTION"
        dgv.Columns(str1).HeaderText = "Original Anal Descr"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns(str1).Visible = True

        str1 = "AnalyteID"
        dgv.Columns(str1).HeaderText = "ID #"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgv.Columns(str1).Visible = True

        str1 = "AnalyteIndex"
        dgv.Columns(str1).HeaderText = "Index #"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns(str1).Visible = True

        str1 = "BQL"
        dgv.Columns(str1).Visible = False

        str1 = "AQL"
        dgv.Columns(str1).Visible = False

        'str1 = "ConcUnits"
        'dgv.Columns(str1).Visible = False
        str1 = "ConcUnits"
        dgv.Columns(str1).Visible = False

        str1 = "AcceptedRuns"
        dgv.Columns(str1).Visible = False

        str1 = "IsReplicate"
        dgv.Columns(str1).HeaderText = "Is" & ChrW(10) & "Replicate?"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns(str1).Visible = True

        str1 = "IsIntStd"
        dgv.Columns(str1).HeaderText = "Is" & ChrW(10) & "Int. Std.?"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns(str1).Visible = True

        str1 = "UseIntStd"
        dgv.Columns(str1).HeaderText = "Use" & ChrW(10) & "Int. Std.?"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv.Columns(str1).Visible = True

        str1 = "IntStd"
        dgv.Columns(str1).HeaderText = "Int. Std."
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        dgv.Columns(str1).Visible = True

        str1 = "MasterAssayID"
        dgv.Columns(str1).HeaderText = "Master" & ChrW(10) & "Assay ID"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgv.Columns(str1).Visible = True

        str1 = "IsCoadminCmpd"
        'dgv.Columns(str1).Visible = False
        dgv.Columns(str1).Visible = True

        Call Order_dgvAnalytes()

    End Sub

    Sub Order_dgvAnalytes()

        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        dgv = Me.dgvAnalytes
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns(Count1).DisplayIndex = int1 - 1
        Next

        str1 = "AnalyteDescription"
        dgv.Columns(str1).DisplayIndex = 0

        str1 = "AnalyteID"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "AnalyteIndex"
        dgv.Columns(str1).DisplayIndex = 11

        str1 = "BQL"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "AQL"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        'str1 = "ConcUnits"
        'dgv.Columns(str1).DisplayIndex = int1 - 1
        str1 = "ConcUnits"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "AcceptedRuns"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "IsReplicate"
        dgv.Columns(str1).DisplayIndex = 1

        str1 = "IsIntStd"
        dgv.Columns(str1).DisplayIndex = 2

        str1 = "UseIntStd"
        dgv.Columns(str1).DisplayIndex = 3

        str1 = "IntStd"
        dgv.Columns(str1).DisplayIndex = 4

        str1 = "MasterAssayID"
        dgv.Columns(str1).DisplayIndex = 10

        str1 = "IsCoadminCmpd"
        dgv.Columns(str1).DisplayIndex = int1 - 1

    End Sub

    Sub Configure_dgvQCConcs()
        Dim dgv As DataGridView
        Dim str1 As String

        dgv = Me.dgvQCConcs

        str1 = "AnalyteDescription"
        dgv.Columns(str1).HeaderText = "Analyte"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        str1 = "NomConc"
        dgv.Columns(str1).HeaderText = "Nom." & ChrW(10) & "Conc."
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "StudyID"
        dgv.Columns(str1).Visible = False

        str1 = "RunTypeID"
        dgv.Columns(str1).Visible = False

        str1 = "RUNID"
        dgv.Columns(str1).HeaderText = "Watson" & ChrW(10) & "Run ID"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        str1 = "RUNSAMPLESEQUENCENUMBER"
        dgv.Columns(str1).Visible = False

        str1 = "ASSAYLEVEL"
        dgv.Columns(str1).HeaderText = "Assay" & ChrW(10) & "Level"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        str1 = "ELIMINATEDFLAG"
        dgv.Columns(str1).HeaderText = "Elim." & ChrW(10) & "Flag"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        str1 = "SAMPLENAME"
        dgv.Columns(str1).HeaderText = "Sample Name"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        str1 = "ALIQUOTFACTOR"
        dgv.Columns(str1).HeaderText = "Dilution" & ChrW(10) & "Factor"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        str1 = "RUNSAMPLEKIND"
        dgv.Columns(str1).Visible = False

        str1 = "ASSAYID"
        dgv.Columns(str1).Visible = False

        str1 = "MasterAssayID"
        dgv.Columns(str1).Visible = False

        'str1 = "ANALYTEINDEX"
        'dgv.Columns(str1).Visible = False
        str1 = "ANALYTEINDEX"
        dgv.Columns(str1).Visible = False

        str1 = "CONCENTRATION"
        dgv.Columns(str1).HeaderText = "Concentration"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "RUNANALYTEREGRESSIONSTATUS"
        dgv.Columns(str1).Visible = False

        Call Order_dgvQCConcs()

    End Sub

    Sub Order_dgvQCConcs()
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        dgv = Me.dgvQCConcs
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns(Count1).DisplayIndex = int1 - 1
        Next


        str1 = "AnalyteDescription"
        dgv.Columns(str1).DisplayIndex = 0

        str1 = "NomConc"
        dgv.Columns(str1).DisplayIndex = 3

        str1 = "StudyID"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "RunTypeID"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "RUNID"
        dgv.Columns(str1).DisplayIndex = 1

        str1 = "RUNSAMPLESEQUENCENUMBER"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "ASSAYLEVEL"
        dgv.Columns(str1).DisplayIndex = 3

        str1 = "ELIMINATEDFLAG"
        dgv.Columns(str1).DisplayIndex = 7

        str1 = "SAMPLENAME"
        dgv.Columns(str1).DisplayIndex = 2

        str1 = "ALIQUOTFACTOR"
        dgv.Columns(str1).DisplayIndex = 6

        str1 = "RUNSAMPLEKIND"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "ASSAYID"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "MasterAssayID"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "ANALYTEINDEX"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "CONCENTRATION"
        dgv.Columns(str1).DisplayIndex = 4

        str1 = "RUNANALYTEREGRESSIONSTATUS"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        'do this three more times
        For Count1 = 1 To 3
            str1 = "AnalyteDescription"
            dgv.Columns(str1).DisplayIndex = 0

            str1 = "NomConc"
            dgv.Columns(str1).DisplayIndex = 3

            str1 = "RUNID"
            dgv.Columns(str1).DisplayIndex = 1

            str1 = "ELIMINATEDFLAG"
            dgv.Columns(str1).DisplayIndex = 7

            str1 = "SAMPLENAME"
            dgv.Columns(str1).DisplayIndex = 2

            str1 = "ALIQUOTFACTOR"
            dgv.Columns(str1).DisplayIndex = 6

            str1 = "CONCENTRATION"
            dgv.Columns(str1).DisplayIndex = 4


            str1 = "ASSAYLEVEL"
            dgv.Columns(str1).DisplayIndex = 3

        Next

    End Sub

    Sub Configure_dgvCalibrConcs()
        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String

        dgv = Me.dgvCalibrConcs
        Dim int1 As Short
        Dim Count1 As Short
        int1 = dgv.ColumnCount
        For Count1 = 0 To int1 - 1
            'dgv.Columns(Count1).Visible = False
        Next

        'str1 = "AnalyteDescription"
        'dgv.Columns(str1).Visible = True
        'dgv.Columns(str1).HeaderText = "Analyte"
        'dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        str1 = "ANALYTEID"
        str2 = "ID"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = str1
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        str1 = "ANALYTEINDEX"
        str2 = "Index"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = str1
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        str1 = "MasterAssayID"
        str2 = str1
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = str1
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        str1 = "RunID"
        str2 = str1
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = str1
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        str1 = "RUNSAMPLESEQUENCENUMBER"
        str2 = "Seq#"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = str1
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        str1 = "AssayLevel"
        str2 = str1
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = str1
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        'str1 = "Concentration"
        'dgv.Columns(str1).Visible = True
        'dgv.Columns(str1).HeaderText = "Nom." & ChrW(10) & "Conc."
        'dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "STUDYID"


        'Call Order_dgvCalibr()

    End Sub

    Sub Configure_dgvCalibr()
        Dim dgv As DataGridView
        Dim str1 As String

        dgv = Me.dgvCalibr
        Dim int1 As Short
        Dim Count1 As Short
        int1 = dgv.ColumnCount
        For Count1 = 0 To int1 - 1
            dgv.Columns(Count1).Visible = False
        Next

        str1 = "AnalyteDescription"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Analyte"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

        str1 = "MasterAssayID"

        str1 = "ANALYTEINDEX"

        str1 = "LevelNumber"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Level"
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "Concentration"
        dgv.Columns(str1).Visible = True
        dgv.Columns(str1).HeaderText = "Nom." & ChrW(10) & "Conc."
        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        str1 = "STUDYID"


        Call Order_dgvCalibr()

    End Sub

    Sub Order_dgvCalibr()
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String

        dgv = Me.dgvCalibr
        int1 = dgv.Columns.Count
        For Count1 = 0 To int1 - 1
            dgv.Columns(Count1).DisplayIndex = int1 - 1
        Next

        str1 = "AnalyteDescription"
        dgv.Columns(str1).DisplayIndex = 0

        str1 = "MasterAssayID"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "ANALYTEINDEX"
        dgv.Columns(str1).DisplayIndex = int1 - 1

        str1 = "LevelNumber"
        dgv.Columns(str1).DisplayIndex = 1

        str1 = "Concentration"
        dgv.Columns(str1).DisplayIndex = 2

        str1 = "STUDYID"
        dgv.Columns(str1).DisplayIndex = int1 - 1

    End Sub


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Me.Visible = False
    End Sub

    Private Sub dgvAnalytes_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAnalytes.CellBeginEdit
        e.Cancel = True
    End Sub

    Private Sub dgvCalibr_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCalibr.CellBeginEdit
        e.Cancel = True
    End Sub

    Private Sub dgvQCConcs_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvQCConcs.CellBeginEdit
        e.Cancel = True
    End Sub

    Private Sub dgvQCs_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvQCs.CellBeginEdit
        e.Cancel = True
    End Sub


    Private Sub llblAssignedSamples_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblAssignedSamples.LinkClicked
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim Count1 As Short

        Me.boolGoTables = True
        Me.Visible = False

        'str1 = "Report Table Configuration"
        'int1 = frmH.lbxTab1.Items.Count
        'For Count1 = 0 To int1 - 1
        '    str2 = frmH.lbxTab1.Items(Count1).ToString
        '    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
        '        frmH.lbxTab1.SelectedIndex = Count1
        '        frmH.tab1.SelectedIndex = Count1
        '        Exit For
        '    End If
        'Next

        ''Me.Visible = False
    End Sub

    Private Sub dgvAnalytes_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAnalytes.MouseEnter
        Me.dgvAnalytes.Focus()

    End Sub

    Private Sub dgvQCs_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvQCs.MouseEnter
        Me.dgvQCs.Focus()

    End Sub

    Private Sub dgvQCConcs_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvQCConcs.MouseEnter
        Me.dgvQCConcs.Focus()

    End Sub

    Private Sub dgvCalibr_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvCalibr.MouseEnter
        Me.dgvCalibr.Focus()

    End Sub

    Private Sub cmdAnalRuns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAnalRuns.Click

        'Call ViewAnalRuns()
        Call OpenAssignedSamples(True)

    End Sub
End Class