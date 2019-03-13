<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReportTableConfig
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReportTableConfig))
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.dgvReportTables = New System.Windows.Forms.DataGridView()
        Me.cmdResize = New System.Windows.Forms.Button()
        Me.gbRTC_Samples = New System.Windows.Forms.GroupBox()
        Me.chkBQLLEGEND = New System.Windows.Forms.CheckBox()
        Me.rbDontShowBQL = New System.Windows.Forms.RadioButton()
        Me.rbShowBQL = New System.Windows.Forms.RadioButton()
        Me.gbRTC_QC = New System.Windows.Forms.GroupBox()
        Me.rbRTC_QC_All = New System.Windows.Forms.RadioButton()
        Me.rbRTC_QC_Acc = New System.Windows.Forms.RadioButton()
        Me.gbRTC_CalStd = New System.Windows.Forms.GroupBox()
        Me.chkRegr = New System.Windows.Forms.CheckBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.gbCalStdValues = New System.Windows.Forms.GroupBox()
        Me.gbxSuper = New System.Windows.Forms.GroupBox()
        Me.rbOutier = New System.Windows.Forms.RadioButton()
        Me.rbNR = New System.Windows.Forms.RadioButton()
        Me.rbShowRejectedValues = New System.Windows.Forms.RadioButton()
        Me.rbDontShowRejected = New System.Windows.Forms.RadioButton()
        Me.rbRTC_CalStd_Acc = New System.Windows.Forms.RadioButton()
        Me.rbRTC_CalStd_All = New System.Windows.Forms.RadioButton()
        Me.gbStats = New System.Windows.Forms.GroupBox()
        Me.chkIncludeIS_Single = New System.Windows.Forms.CheckBox()
        Me.panOptions = New System.Windows.Forms.Panel()
        Me.lblAccuracy = New System.Windows.Forms.Label()
        Me.lblPrecision = New System.Windows.Forms.Label()
        Me.lblDiff = New System.Windows.Forms.Label()
        Me.chkMean = New System.Windows.Forms.CheckBox()
        Me.chkSD = New System.Windows.Forms.CheckBox()
        Me.chkCV = New System.Windows.Forms.CheckBox()
        Me.chkRE = New System.Windows.Forms.CheckBox()
        Me.chkBias = New System.Windows.Forms.CheckBox()
        Me.chkN = New System.Windows.Forms.CheckBox()
        Me.chkDiff = New System.Windows.Forms.CheckBox()
        Me.chkTheoretical = New System.Windows.Forms.CheckBox()
        Me.chkDiffCol = New System.Windows.Forms.CheckBox()
        Me.chkIncludeWatsonLabel = New System.Windows.Forms.CheckBox()
        Me.chkIncludeDate = New System.Windows.Forms.CheckBox()
        Me.lblDivider = New System.Windows.Forms.Label()
        Me.gbAdditional = New System.Windows.Forms.GroupBox()
        Me.lblRemember = New System.Windows.Forms.Label()
        Me.cmdInsert = New System.Windows.Forms.Button()
        Me.lblCHARSTABILITYPERIOD = New System.Windows.Forms.Label()
        Me.cmdBuild = New System.Windows.Forms.Button()
        Me.CHARSTABILITYPERIOD = New System.Windows.Forms.TextBox()
        Me.panTP = New System.Windows.Forms.Panel()
        Me.lblPeriodTemp = New System.Windows.Forms.Label()
        Me.chkCONVERTTEMP = New System.Windows.Forms.CheckBox()
        Me.chkCONVERTTIME = New System.Windows.Forms.CheckBox()
        Me.CHARPERIODTEMP = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CHARTIMEFRAME = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CHARTIMEPERIOD = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.panCycles = New System.Windows.Forms.Panel()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.INTNUMBEROFCYCLES = New System.Windows.Forms.TextBox()
        Me.lblCycles = New System.Windows.Forms.Label()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.cmsFieldCodes = New System.Windows.Forms.ToolStripMenuItem()
        Me.cmdIncSamples = New System.Windows.Forms.Button()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.gbAnovaStats = New System.Windows.Forms.GroupBox()
        Me.chkIntraRunSumStats = New System.Windows.Forms.CheckBox()
        Me.chkIncludeAnovaSumStats = New System.Windows.Forms.CheckBox()
        Me.chkIncludeAnova = New System.Windows.Forms.CheckBox()
        Me.gbSampleSort = New System.Windows.Forms.GroupBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.cbxSampleSAD6 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleS6 = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cbxSampleSAD5 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleS5 = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbxSampleSAD4 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleS4 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleSAD3 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleS3 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleSAD2 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleS2 = New System.Windows.Forms.ComboBox()
        Me.lblA = New System.Windows.Forms.Label()
        Me.cbxSampleSAD1 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleS1 = New System.Windows.Forms.ComboBox()
        Me.lblLevel = New System.Windows.Forms.Label()
        Me.lblSort = New System.Windows.Forms.Label()
        Me.gbSampleGroup = New System.Windows.Forms.GroupBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.cbxSampleGAD4 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleG4 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleGAD3 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleG3 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleGAD2 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleG2 = New System.Windows.Forms.ComboBox()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cbxSampleGAD1 = New System.Windows.Forms.ComboBox()
        Me.cbxSampleG1 = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.chkIncludePSAE = New System.Windows.Forms.CheckBox()
        Me.gbPSAE = New System.Windows.Forms.GroupBox()
        Me.chkInjCol = New System.Windows.Forms.CheckBox()
        Me.gbResultsChoice = New System.Windows.Forms.GroupBox()
        Me.rbUseISPeakArea = New System.Windows.Forms.RadioButton()
        Me.chkBOOLADHOCSTABCOMPCOLUMNS = New System.Windows.Forms.CheckBox()
        Me.panIS = New System.Windows.Forms.Panel()
        Me.CHARISCONC = New System.Windows.Forms.TextBox()
        Me.chkBOOLISCOMBINELEVELS = New System.Windows.Forms.CheckBox()
        Me.chkIncludeIS = New System.Windows.Forms.CheckBox()
        Me.lblCHARISCONC = New System.Windows.Forms.Label()
        Me.chkCustomLeg = New System.Windows.Forms.CheckBox()
        Me.rbUsePeakAreaRatio = New System.Windows.Forms.RadioButton()
        Me.rbUsePeakArea = New System.Windows.Forms.RadioButton()
        Me.rbConc = New System.Windows.Forms.RadioButton()
        Me.chkBOOLDOINDREC = New System.Windows.Forms.CheckBox()
        Me.gbTableLegend = New System.Windows.Forms.GroupBox()
        Me.chkNoneLeg = New System.Windows.Forms.CheckBox()
        Me.panTitleLegends = New System.Windows.Forms.Panel()
        Me.CHARTITLELEG = New System.Windows.Forms.TextBox()
        Me.CHARNUMLEG = New System.Windows.Forms.TextBox()
        Me.lblCHARDENLEG = New System.Windows.Forms.Label()
        Me.lblCHARTITLELEG = New System.Windows.Forms.Label()
        Me.CHARDENLEG = New System.Windows.Forms.TextBox()
        Me.lblCHARNUMLEG = New System.Windows.Forms.Label()
        Me.panNomDenomCalcs = New System.Windows.Forms.Panel()
        Me.chkRTC_CalStd_Acc = New System.Windows.Forms.CheckBox()
        Me.panNomDenom = New System.Windows.Forms.Panel()
        Me.gbDenom = New System.Windows.Forms.GroupBox()
        Me.rbOld = New System.Windows.Forms.RadioButton()
        Me.rbNew = New System.Windows.Forms.RadioButton()
        Me.gbNumerator = New System.Windows.Forms.GroupBox()
        Me.rbPosLeg = New System.Windows.Forms.RadioButton()
        Me.rbNegLeg = New System.Windows.Forms.RadioButton()
        Me.gbCalcs = New System.Windows.Forms.GroupBox()
        Me.lblMeanAccuracy = New System.Windows.Forms.Label()
        Me.lblRecovery = New System.Windows.Forms.Label()
        Me.chkCalcIntStdNMF = New System.Windows.Forms.CheckBox()
        Me.lblPercDiff = New System.Windows.Forms.Label()
        Me.rbMeanAcc = New System.Windows.Forms.RadioButton()
        Me.rbRecovery = New System.Windows.Forms.RadioButton()
        Me.rbDifference = New System.Windows.Forms.RadioButton()
        Me.tabRTC = New System.Windows.Forms.TabControl()
        Me.tpFormat = New System.Windows.Forms.TabPage()
        Me.panFormat = New System.Windows.Forms.Panel()
        Me.gbMatrixFactor = New System.Windows.Forms.GroupBox()
        Me.panMF1 = New System.Windows.Forms.Panel()
        Me.lblMF01 = New System.Windows.Forms.Label()
        Me.chkInclIntStdNMF = New System.Windows.Forms.CheckBox()
        Me.chkInclMFCols = New System.Windows.Forms.CheckBox()
        Me.chkMFTable = New System.Windows.Forms.CheckBox()
        Me.gbGroupSort = New System.Windows.Forms.GroupBox()
        Me.lblGroupSort = New System.Windows.Forms.Label()
        Me.gbSAT = New System.Windows.Forms.GroupBox()
        Me.chkBOOLCONCCOMMENTS = New System.Windows.Forms.CheckBox()
        Me.gbQCGroup = New System.Windows.Forms.GroupBox()
        Me.lblQCGroup = New System.Windows.Forms.Label()
        Me.rbINTQCLEVELGROUPQCLabel = New System.Windows.Forms.RadioButton()
        Me.rbINTQCLEVELGROUPNomConc = New System.Windows.Forms.RadioButton()
        Me.rbINTQCLEVELGROUPLevel = New System.Windows.Forms.RadioButton()
        Me.gbRegrULOQ = New System.Windows.Forms.GroupBox()
        Me.chkBOOLREGRULOQ = New System.Windows.Forms.CheckBox()
        Me.gbCriteria = New System.Windows.Forms.GroupBox()
        Me.NUMPRECCRITLOTS = New System.Windows.Forms.TextBox()
        Me.lblPrecisionCrit = New System.Windows.Forms.Label()
        Me.gbCarryover = New System.Windows.Forms.GroupBox()
        Me.lblCarryover = New System.Windows.Forms.Label()
        Me.CHARCARRYOVERLABEL = New System.Windows.Forms.TextBox()
        Me.gbLegendFormat = New System.Windows.Forms.GroupBox()
        Me.chkBOOLREASSAYREASLETTERS = New System.Windows.Forms.CheckBox()
        Me.gbIncSampleCriteria = New System.Windows.Forms.GroupBox()
        Me.dgvAnalytes = New System.Windows.Forms.DataGridView()
        Me.tpPeriodTemp = New System.Windows.Forms.TabPage()
        Me.panFDARef = New System.Windows.Forms.Panel()
        Me.txtRef = New System.Windows.Forms.TextBox()
        Me.lblFDA = New System.Windows.Forms.Label()
        Me.gbStabilityType = New System.Windows.Forms.GroupBox()
        Me.rbDilution = New System.Windows.Forms.RadioButton()
        Me.rbAutosampler = New System.Windows.Forms.RadioButton()
        Me.rbBatchReinjection = New System.Windows.Forms.RadioButton()
        Me.rbSpiking = New System.Windows.Forms.RadioButton()
        Me.rbReinjection = New System.Windows.Forms.RadioButton()
        Me.rbStockSolution = New System.Windows.Forms.RadioButton()
        Me.rbBlood = New System.Windows.Forms.RadioButton()
        Me.rbLT = New System.Windows.Forms.RadioButton()
        Me.rbFT = New System.Windows.Forms.RadioButton()
        Me.rbBenchTop = New System.Windows.Forms.RadioButton()
        Me.rbProcess = New System.Windows.Forms.RadioButton()
        Me.rbNA = New System.Windows.Forms.RadioButton()
        Me.txtStabilityNotes = New System.Windows.Forms.TextBox()
        Me.lblStabilityNotes = New System.Windows.Forms.Label()
        Me.tpAutoAssignment = New System.Windows.Forms.TabPage()
        Me.lblRunIdentifier = New System.Windows.Forms.Label()
        Me.lblCase = New System.Windows.Forms.Label()
        Me.lblLogic = New System.Windows.Forms.Label()
        Me.dgvSAS = New System.Windows.Forms.DataGridView()
        Me.cmdAnalRuns = New System.Windows.Forms.Button()
        Me.dgvASP = New System.Windows.Forms.DataGridView()
        Me.panTableGraphicExamples = New System.Windows.Forms.Panel()
        Me.pbxTableGraphicExamples = New System.Windows.Forms.PictureBox()
        Me.lblTableGraphicExamplesLabel = New System.Windows.Forms.Label()
        Me.cmdShowRunSummary = New System.Windows.Forms.Button()
        Me.lblOptional = New System.Windows.Forms.Label()
        Me.lblWatsonE = New System.Windows.Forms.Label()
        Me.lblAccepted = New System.Windows.Forms.Label()
        Me.panEdit = New System.Windows.Forms.Panel()
        Me.cmdPasteConditions = New System.Windows.Forms.Button()
        Me.txtTitle = New System.Windows.Forms.TextBox()
        Me.cmdSymbol = New System.Windows.Forms.Button()
        Me.cmdTest = New System.Windows.Forms.Button()
        Me.lblClose = New System.Windows.Forms.Label()
        CType(Me.dgvReportTables, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbRTC_Samples.SuspendLayout()
        Me.gbRTC_QC.SuspendLayout()
        Me.gbRTC_CalStd.SuspendLayout()
        Me.gbCalStdValues.SuspendLayout()
        Me.gbxSuper.SuspendLayout()
        Me.gbStats.SuspendLayout()
        Me.panOptions.SuspendLayout()
        Me.gbAdditional.SuspendLayout()
        Me.panTP.SuspendLayout()
        Me.panCycles.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.gbAnovaStats.SuspendLayout()
        Me.gbSampleSort.SuspendLayout()
        Me.gbSampleGroup.SuspendLayout()
        Me.gbPSAE.SuspendLayout()
        Me.gbResultsChoice.SuspendLayout()
        Me.panIS.SuspendLayout()
        Me.gbTableLegend.SuspendLayout()
        Me.panTitleLegends.SuspendLayout()
        Me.panNomDenomCalcs.SuspendLayout()
        Me.panNomDenom.SuspendLayout()
        Me.gbDenom.SuspendLayout()
        Me.gbNumerator.SuspendLayout()
        Me.gbCalcs.SuspendLayout()
        Me.tabRTC.SuspendLayout()
        Me.tpFormat.SuspendLayout()
        Me.panFormat.SuspendLayout()
        Me.gbMatrixFactor.SuspendLayout()
        Me.panMF1.SuspendLayout()
        Me.gbGroupSort.SuspendLayout()
        Me.gbSAT.SuspendLayout()
        Me.gbQCGroup.SuspendLayout()
        Me.gbRegrULOQ.SuspendLayout()
        Me.gbCriteria.SuspendLayout()
        Me.gbCarryover.SuspendLayout()
        Me.gbLegendFormat.SuspendLayout()
        Me.gbIncSampleCriteria.SuspendLayout()
        CType(Me.dgvAnalytes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tpPeriodTemp.SuspendLayout()
        Me.panFDARef.SuspendLayout()
        Me.gbStabilityType.SuspendLayout()
        Me.tpAutoAssignment.SuspendLayout()
        CType(Me.dgvSAS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvASP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panTableGraphicExamples.SuspendLayout()
        CType(Me.pbxTableGraphicExamples, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panEdit.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.Gray
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.ForestGreen
        Me.cmdSave.Location = New System.Drawing.Point(83, 0)
        Me.cmdSave.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(79, 33)
        Me.cmdSave.TabIndex = 3
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdEdit
        '
        Me.cmdEdit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdEdit.CausesValidation = False
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdEdit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEdit.ForeColor = System.Drawing.Color.Blue
        Me.cmdEdit.Location = New System.Drawing.Point(1, 0)
        Me.cmdEdit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(79, 33)
        Me.cmdEdit.TabIndex = 2
        Me.cmdEdit.Text = "&Edit"
        Me.cmdEdit.UseVisualStyleBackColor = False
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.CausesValidation = False
        Me.cmdExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(247, 0)
        Me.cmdExit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(79, 33)
        Me.cmdExit.TabIndex = 5
        Me.cmdExit.Text = "G&o Back"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gray
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(165, 0)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(79, 33)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'dgvReportTables
        '
        Me.dgvReportTables.AllowUserToAddRows = False
        Me.dgvReportTables.AllowUserToDeleteRows = False
        Me.dgvReportTables.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvReportTables.BackgroundColor = System.Drawing.Color.White
        Me.dgvReportTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.Padding = New System.Windows.Forms.Padding(0, 10, 0, 10)
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvReportTables.DefaultCellStyle = DataGridViewCellStyle1
        Me.dgvReportTables.Location = New System.Drawing.Point(14, 110)
        Me.dgvReportTables.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvReportTables.MultiSelect = False
        Me.dgvReportTables.Name = "dgvReportTables"
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTables.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.dgvReportTables.Size = New System.Drawing.Size(152, 661)
        Me.dgvReportTables.TabIndex = 1
        '
        'cmdResize
        '
        Me.cmdResize.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResize.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdResize.Location = New System.Drawing.Point(720, 0)
        Me.cmdResize.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdResize.Name = "cmdResize"
        Me.cmdResize.Size = New System.Drawing.Size(169, 33)
        Me.cmdResize.TabIndex = 140
        Me.cmdResize.Text = "&Resize Rows"
        Me.cmdResize.UseVisualStyleBackColor = True
        Me.cmdResize.Visible = False
        '
        'gbRTC_Samples
        '
        Me.gbRTC_Samples.Controls.Add(Me.chkBQLLEGEND)
        Me.gbRTC_Samples.Controls.Add(Me.rbDontShowBQL)
        Me.gbRTC_Samples.Controls.Add(Me.rbShowBQL)
        Me.gbRTC_Samples.Enabled = False
        Me.gbRTC_Samples.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbRTC_Samples.Location = New System.Drawing.Point(12, 10)
        Me.gbRTC_Samples.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbRTC_Samples.Name = "gbRTC_Samples"
        Me.gbRTC_Samples.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbRTC_Samples.Size = New System.Drawing.Size(307, 94)
        Me.gbRTC_Samples.TabIndex = 139
        Me.gbRTC_Samples.TabStop = False
        Me.gbRTC_Samples.Text = "Sample BQL Reporting Options"
        '
        'chkBQLLEGEND
        '
        Me.chkBQLLEGEND.AutoSize = True
        Me.chkBQLLEGEND.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBQLLEGEND.Location = New System.Drawing.Point(12, 70)
        Me.chkBQLLEGEND.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkBQLLEGEND.Name = "chkBQLLEGEND"
        Me.chkBQLLEGEND.Size = New System.Drawing.Size(163, 17)
        Me.chkBQLLEGEND.TabIndex = 2
        Me.chkBQLLEGEND.Text = "Display BQL value in Legend"
        Me.chkBQLLEGEND.UseVisualStyleBackColor = True
        '
        'rbDontShowBQL
        '
        Me.rbDontShowBQL.AutoSize = True
        Me.rbDontShowBQL.BackColor = System.Drawing.Color.Transparent
        Me.rbDontShowBQL.Checked = True
        Me.rbDontShowBQL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbDontShowBQL.Location = New System.Drawing.Point(12, 46)
        Me.rbDontShowBQL.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbDontShowBQL.Name = "rbDontShowBQL"
        Me.rbDontShowBQL.Size = New System.Drawing.Size(243, 19)
        Me.rbDontShowBQL.TabIndex = 1
        Me.rbDontShowBQL.TabStop = True
        Me.rbDontShowBQL.Text = "Don't Show Conc. Values that are < BQL"
        Me.rbDontShowBQL.UseVisualStyleBackColor = False
        '
        'rbShowBQL
        '
        Me.rbShowBQL.AutoSize = True
        Me.rbShowBQL.BackColor = System.Drawing.Color.Transparent
        Me.rbShowBQL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbShowBQL.Location = New System.Drawing.Point(12, 22)
        Me.rbShowBQL.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowBQL.Name = "rbShowBQL"
        Me.rbShowBQL.Size = New System.Drawing.Size(211, 19)
        Me.rbShowBQL.TabIndex = 0
        Me.rbShowBQL.Text = "Show Conc. Values that are < BQL"
        Me.rbShowBQL.UseVisualStyleBackColor = False
        '
        'gbRTC_QC
        '
        Me.gbRTC_QC.Controls.Add(Me.rbRTC_QC_All)
        Me.gbRTC_QC.Controls.Add(Me.rbRTC_QC_Acc)
        Me.gbRTC_QC.Enabled = False
        Me.gbRTC_QC.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbRTC_QC.Location = New System.Drawing.Point(12, 109)
        Me.gbRTC_QC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbRTC_QC.Name = "gbRTC_QC"
        Me.gbRTC_QC.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbRTC_QC.Size = New System.Drawing.Size(308, 70)
        Me.gbRTC_QC.TabIndex = 138
        Me.gbRTC_QC.TabStop = False
        Me.gbRTC_QC.Text = "QC Std. Summary Statistics Options"
        '
        'rbRTC_QC_All
        '
        Me.rbRTC_QC_All.AutoSize = True
        Me.rbRTC_QC_All.BackColor = System.Drawing.Color.Transparent
        Me.rbRTC_QC_All.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbRTC_QC_All.Location = New System.Drawing.Point(12, 43)
        Me.rbRTC_QC_All.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbRTC_QC_All.Name = "rbRTC_QC_All"
        Me.rbRTC_QC_All.Size = New System.Drawing.Size(218, 19)
        Me.rbRTC_QC_All.TabIndex = 1
        Me.rbRTC_QC_All.Text = "Report Accepted and Outlier Values"
        Me.rbRTC_QC_All.UseVisualStyleBackColor = False
        '
        'rbRTC_QC_Acc
        '
        Me.rbRTC_QC_Acc.AutoSize = True
        Me.rbRTC_QC_Acc.BackColor = System.Drawing.Color.Transparent
        Me.rbRTC_QC_Acc.Checked = True
        Me.rbRTC_QC_Acc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbRTC_QC_Acc.Location = New System.Drawing.Point(12, 22)
        Me.rbRTC_QC_Acc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbRTC_QC_Acc.Name = "rbRTC_QC_Acc"
        Me.rbRTC_QC_Acc.Size = New System.Drawing.Size(182, 19)
        Me.rbRTC_QC_Acc.TabIndex = 0
        Me.rbRTC_QC_Acc.TabStop = True
        Me.rbRTC_QC_Acc.Text = "Report Only Accepted Values"
        Me.rbRTC_QC_Acc.UseVisualStyleBackColor = False
        '
        'gbRTC_CalStd
        '
        Me.gbRTC_CalStd.Controls.Add(Me.chkRegr)
        Me.gbRTC_CalStd.Controls.Add(Me.GroupBox1)
        Me.gbRTC_CalStd.Controls.Add(Me.gbCalStdValues)
        Me.gbRTC_CalStd.Enabled = False
        Me.gbRTC_CalStd.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbRTC_CalStd.Location = New System.Drawing.Point(339, 10)
        Me.gbRTC_CalStd.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbRTC_CalStd.Name = "gbRTC_CalStd"
        Me.gbRTC_CalStd.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbRTC_CalStd.Size = New System.Drawing.Size(324, 118)
        Me.gbRTC_CalStd.TabIndex = 137
        Me.gbRTC_CalStd.TabStop = False
        Me.gbRTC_CalStd.Text = "Calibr. Std. Summary Statistics Options"
        '
        'chkRegr
        '
        Me.chkRegr.AutoSize = True
        Me.chkRegr.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkRegr.Location = New System.Drawing.Point(15, 93)
        Me.chkRegr.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkRegr.Name = "chkRegr"
        Me.chkRegr.Size = New System.Drawing.Size(245, 21)
        Me.chkRegr.TabIndex = 4
        Me.chkRegr.Text = "Include Regression Constants in table"
        Me.chkRegr.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.ForeColor = System.Drawing.Color.Red
        Me.GroupBox1.Location = New System.Drawing.Point(12, 179)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(306, 75)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Statistics"
        Me.GroupBox1.Visible = False
        '
        'gbCalStdValues
        '
        Me.gbCalStdValues.Controls.Add(Me.gbxSuper)
        Me.gbCalStdValues.Controls.Add(Me.rbShowRejectedValues)
        Me.gbCalStdValues.Controls.Add(Me.rbDontShowRejected)
        Me.gbCalStdValues.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.gbCalStdValues.Location = New System.Drawing.Point(9, 21)
        Me.gbCalStdValues.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCalStdValues.Name = "gbCalStdValues"
        Me.gbCalStdValues.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCalStdValues.Size = New System.Drawing.Size(307, 70)
        Me.gbCalStdValues.TabIndex = 2
        Me.gbCalStdValues.TabStop = False
        Me.gbCalStdValues.Text = "Values"
        '
        'gbxSuper
        '
        Me.gbxSuper.Controls.Add(Me.rbOutier)
        Me.gbxSuper.Controls.Add(Me.rbNR)
        Me.gbxSuper.Location = New System.Drawing.Point(208, 17)
        Me.gbxSuper.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxSuper.Name = "gbxSuper"
        Me.gbxSuper.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbxSuper.Size = New System.Drawing.Size(75, 52)
        Me.gbxSuper.TabIndex = 2
        Me.gbxSuper.TabStop = False
        Me.gbxSuper.Text = "Superscript Type"
        Me.gbxSuper.Visible = False
        '
        'rbOutier
        '
        Me.rbOutier.AutoSize = True
        Me.rbOutier.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbOutier.ForeColor = System.Drawing.Color.Black
        Me.rbOutier.Location = New System.Drawing.Point(65, 25)
        Me.rbOutier.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbOutier.Name = "rbOutier"
        Me.rbOutier.Size = New System.Drawing.Size(173, 17)
        Me.rbOutier.TabIndex = 1
        Me.rbOutier.Text = "Letter referencing outlier criteria"
        Me.rbOutier.UseVisualStyleBackColor = True
        '
        'rbNR
        '
        Me.rbNR.AutoSize = True
        Me.rbNR.Checked = True
        Me.rbNR.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbNR.ForeColor = System.Drawing.Color.Black
        Me.rbNR.Location = New System.Drawing.Point(7, 25)
        Me.rbNR.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbNR.Name = "rbNR"
        Me.rbNR.Size = New System.Drawing.Size(41, 17)
        Me.rbNR.TabIndex = 0
        Me.rbNR.TabStop = True
        Me.rbNR.Text = "NR"
        Me.rbNR.UseVisualStyleBackColor = True
        '
        'rbShowRejectedValues
        '
        Me.rbShowRejectedValues.AutoSize = True
        Me.rbShowRejectedValues.BackColor = System.Drawing.Color.Transparent
        Me.rbShowRejectedValues.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbShowRejectedValues.ForeColor = System.Drawing.Color.Black
        Me.rbShowRejectedValues.Location = New System.Drawing.Point(12, 22)
        Me.rbShowRejectedValues.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbShowRejectedValues.Name = "rbShowRejectedValues"
        Me.rbShowRejectedValues.Size = New System.Drawing.Size(153, 21)
        Me.rbShowRejectedValues.TabIndex = 0
        Me.rbShowRejectedValues.Text = "Show Rejected Values"
        Me.rbShowRejectedValues.UseVisualStyleBackColor = False
        '
        'rbDontShowRejected
        '
        Me.rbDontShowRejected.AutoSize = True
        Me.rbDontShowRejected.BackColor = System.Drawing.Color.Transparent
        Me.rbDontShowRejected.Checked = True
        Me.rbDontShowRejected.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbDontShowRejected.ForeColor = System.Drawing.Color.Black
        Me.rbDontShowRejected.Location = New System.Drawing.Point(12, 42)
        Me.rbDontShowRejected.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbDontShowRejected.Name = "rbDontShowRejected"
        Me.rbDontShowRejected.Size = New System.Drawing.Size(188, 21)
        Me.rbDontShowRejected.TabIndex = 1
        Me.rbDontShowRejected.TabStop = True
        Me.rbDontShowRejected.Text = "Don't Show Rejected Values"
        Me.rbDontShowRejected.UseVisualStyleBackColor = False
        '
        'rbRTC_CalStd_Acc
        '
        Me.rbRTC_CalStd_Acc.BackColor = System.Drawing.Color.Transparent
        Me.rbRTC_CalStd_Acc.Checked = True
        Me.rbRTC_CalStd_Acc.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbRTC_CalStd_Acc.ForeColor = System.Drawing.Color.Black
        Me.rbRTC_CalStd_Acc.Location = New System.Drawing.Point(365, 453)
        Me.rbRTC_CalStd_Acc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbRTC_CalStd_Acc.Name = "rbRTC_CalStd_Acc"
        Me.rbRTC_CalStd_Acc.Size = New System.Drawing.Size(280, 34)
        Me.rbRTC_CalStd_Acc.TabIndex = 1
        Me.rbRTC_CalStd_Acc.TabStop = True
        Me.rbRTC_CalStd_Acc.Text = "Report Only Accepted Values"
        Me.rbRTC_CalStd_Acc.UseVisualStyleBackColor = False
        Me.rbRTC_CalStd_Acc.Visible = False
        '
        'rbRTC_CalStd_All
        '
        Me.rbRTC_CalStd_All.BackColor = System.Drawing.Color.Transparent
        Me.rbRTC_CalStd_All.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbRTC_CalStd_All.ForeColor = System.Drawing.Color.Black
        Me.rbRTC_CalStd_All.Location = New System.Drawing.Point(365, 428)
        Me.rbRTC_CalStd_All.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbRTC_CalStd_All.Name = "rbRTC_CalStd_All"
        Me.rbRTC_CalStd_All.Size = New System.Drawing.Size(293, 31)
        Me.rbRTC_CalStd_All.TabIndex = 0
        Me.rbRTC_CalStd_All.Text = "Report Accepted and Rejected Values"
        Me.rbRTC_CalStd_All.UseVisualStyleBackColor = False
        Me.rbRTC_CalStd_All.Visible = False
        '
        'gbStats
        '
        Me.gbStats.Controls.Add(Me.chkIncludeIS_Single)
        Me.gbStats.Controls.Add(Me.panOptions)
        Me.gbStats.Controls.Add(Me.chkIncludeWatsonLabel)
        Me.gbStats.Controls.Add(Me.chkIncludeDate)
        Me.gbStats.Controls.Add(Me.lblDivider)
        Me.gbStats.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbStats.Location = New System.Drawing.Point(339, 130)
        Me.gbStats.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStats.Name = "gbStats"
        Me.gbStats.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStats.Size = New System.Drawing.Size(324, 296)
        Me.gbStats.TabIndex = 141
        Me.gbStats.TabStop = False
        Me.gbStats.Text = "Statistics Options and Additional Content"
        '
        'chkIncludeIS_Single
        '
        Me.chkIncludeIS_Single.AutoSize = True
        Me.chkIncludeIS_Single.Checked = True
        Me.chkIncludeIS_Single.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncludeIS_Single.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkIncludeIS_Single.Location = New System.Drawing.Point(172, 267)
        Me.chkIncludeIS_Single.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkIncludeIS_Single.Name = "chkIncludeIS_Single"
        Me.chkIncludeIS_Single.Size = New System.Drawing.Size(157, 21)
        Me.chkIncludeIS_Single.TabIndex = 158
        Me.chkIncludeIS_Single.Text = "Include Int Std Column"
        Me.chkIncludeIS_Single.UseVisualStyleBackColor = True
        '
        'panOptions
        '
        Me.panOptions.Controls.Add(Me.lblAccuracy)
        Me.panOptions.Controls.Add(Me.lblPrecision)
        Me.panOptions.Controls.Add(Me.lblDiff)
        Me.panOptions.Controls.Add(Me.chkMean)
        Me.panOptions.Controls.Add(Me.chkSD)
        Me.panOptions.Controls.Add(Me.chkCV)
        Me.panOptions.Controls.Add(Me.chkRE)
        Me.panOptions.Controls.Add(Me.chkBias)
        Me.panOptions.Controls.Add(Me.chkN)
        Me.panOptions.Controls.Add(Me.chkDiff)
        Me.panOptions.Controls.Add(Me.chkTheoretical)
        Me.panOptions.Controls.Add(Me.chkDiffCol)
        Me.panOptions.Location = New System.Drawing.Point(23, 25)
        Me.panOptions.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panOptions.Name = "panOptions"
        Me.panOptions.Size = New System.Drawing.Size(290, 216)
        Me.panOptions.TabIndex = 157
        '
        'lblAccuracy
        '
        Me.lblAccuracy.AutoSize = True
        Me.lblAccuracy.Location = New System.Drawing.Point(3, 86)
        Me.lblAccuracy.Name = "lblAccuracy"
        Me.lblAccuracy.Size = New System.Drawing.Size(229, 17)
        Me.lblAccuracy.TabIndex = 15
        Me.lblAccuracy.Text = "Accuracy: ((Mean/NomConc)-1)*100"
        '
        'lblPrecision
        '
        Me.lblPrecision.AutoSize = True
        Me.lblPrecision.Location = New System.Drawing.Point(3, 44)
        Me.lblPrecision.Name = "lblPrecision"
        Me.lblPrecision.Size = New System.Drawing.Size(197, 17)
        Me.lblPrecision.TabIndex = 14
        Me.lblPrecision.Text = "Precision:  (StdDev/Mean)*100"
        '
        'lblDiff
        '
        Me.lblDiff.AutoSize = True
        Me.lblDiff.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDiff.Location = New System.Drawing.Point(90, 134)
        Me.lblDiff.Name = "lblDiff"
        Me.lblDiff.Size = New System.Drawing.Size(17, 25)
        Me.lblDiff.TabIndex = 13
        Me.lblDiff.Text = "l"
        Me.lblDiff.Visible = False
        '
        'chkMean
        '
        Me.chkMean.AutoSize = True
        Me.chkMean.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkMean.Location = New System.Drawing.Point(3, 4)
        Me.chkMean.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkMean.Name = "chkMean"
        Me.chkMean.Size = New System.Drawing.Size(60, 21)
        Me.chkMean.TabIndex = 0
        Me.chkMean.Text = "Mean"
        Me.chkMean.UseVisualStyleBackColor = True
        '
        'chkSD
        '
        Me.chkSD.AutoSize = True
        Me.chkSD.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkSD.Location = New System.Drawing.Point(3, 24)
        Me.chkSD.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkSD.Name = "chkSD"
        Me.chkSD.Size = New System.Drawing.Size(43, 21)
        Me.chkSD.TabIndex = 1
        Me.chkSD.Text = "SD"
        Me.chkSD.UseVisualStyleBackColor = True
        '
        'chkCV
        '
        Me.chkCV.AutoSize = True
        Me.chkCV.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkCV.Location = New System.Drawing.Point(30, 64)
        Me.chkCV.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkCV.Name = "chkCV"
        Me.chkCV.Size = New System.Drawing.Size(54, 21)
        Me.chkCV.TabIndex = 2
        Me.chkCV.Text = "%CV"
        Me.chkCV.UseVisualStyleBackColor = True
        '
        'chkRE
        '
        Me.chkRE.AutoSize = True
        Me.chkRE.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkRE.Location = New System.Drawing.Point(30, 147)
        Me.chkRE.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkRE.Name = "chkRE"
        Me.chkRE.Size = New System.Drawing.Size(53, 21)
        Me.chkRE.TabIndex = 10
        Me.chkRE.Text = "%RE"
        Me.chkRE.UseVisualStyleBackColor = True
        '
        'chkBias
        '
        Me.chkBias.AutoSize = True
        Me.chkBias.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkBias.Location = New System.Drawing.Point(30, 105)
        Me.chkBias.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkBias.Name = "chkBias"
        Me.chkBias.Size = New System.Drawing.Size(61, 21)
        Me.chkBias.TabIndex = 3
        Me.chkBias.Text = "%Bias"
        Me.chkBias.UseVisualStyleBackColor = True
        '
        'chkN
        '
        Me.chkN.AutoSize = True
        Me.chkN.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkN.Location = New System.Drawing.Point(3, 194)
        Me.chkN.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkN.Name = "chkN"
        Me.chkN.Size = New System.Drawing.Size(34, 21)
        Me.chkN.TabIndex = 7
        Me.chkN.Text = "n"
        Me.chkN.UseVisualStyleBackColor = True
        '
        'chkDiff
        '
        Me.chkDiff.AutoSize = True
        Me.chkDiff.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkDiff.Location = New System.Drawing.Point(30, 126)
        Me.chkDiff.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkDiff.Name = "chkDiff"
        Me.chkDiff.Size = New System.Drawing.Size(58, 21)
        Me.chkDiff.TabIndex = 5
        Me.chkDiff.Text = "%Diff"
        Me.chkDiff.UseVisualStyleBackColor = True
        '
        'chkTheoretical
        '
        Me.chkTheoretical.AutoSize = True
        Me.chkTheoretical.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkTheoretical.Location = New System.Drawing.Point(30, 168)
        Me.chkTheoretical.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkTheoretical.Name = "chkTheoretical"
        Me.chkTheoretical.Size = New System.Drawing.Size(203, 21)
        Me.chkTheoretical.TabIndex = 4
        Me.chkTheoretical.Text = "%Theoretical (100 + Accuracy)"
        Me.chkTheoretical.UseVisualStyleBackColor = True
        '
        'chkDiffCol
        '
        Me.chkDiffCol.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkDiffCol.Location = New System.Drawing.Point(117, 117)
        Me.chkDiffCol.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkDiffCol.Name = "chkDiffCol"
        Me.chkDiffCol.Size = New System.Drawing.Size(162, 39)
        Me.chkDiffCol.TabIndex = 6
        Me.chkDiffCol.Text = "Include Individual Calculation Column"
        Me.chkDiffCol.UseVisualStyleBackColor = True
        '
        'chkIncludeWatsonLabel
        '
        Me.chkIncludeWatsonLabel.AutoSize = True
        Me.chkIncludeWatsonLabel.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkIncludeWatsonLabel.Location = New System.Drawing.Point(27, 267)
        Me.chkIncludeWatsonLabel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkIncludeWatsonLabel.Name = "chkIncludeWatsonLabel"
        Me.chkIncludeWatsonLabel.Size = New System.Drawing.Size(157, 21)
        Me.chkIncludeWatsonLabel.TabIndex = 9
        Me.chkIncludeWatsonLabel.Text = "Include Watson Labels"
        Me.chkIncludeWatsonLabel.UseVisualStyleBackColor = True
        Me.chkIncludeWatsonLabel.Visible = False
        '
        'chkIncludeDate
        '
        Me.chkIncludeDate.AutoSize = True
        Me.chkIncludeDate.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkIncludeDate.Location = New System.Drawing.Point(27, 247)
        Me.chkIncludeDate.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkIncludeDate.Name = "chkIncludeDate"
        Me.chkIncludeDate.Size = New System.Drawing.Size(149, 21)
        Me.chkIncludeDate.TabIndex = 6
        Me.chkIncludeDate.Text = "Include Analysis Date"
        Me.chkIncludeDate.UseVisualStyleBackColor = True
        '
        'lblDivider
        '
        Me.lblDivider.AutoSize = True
        Me.lblDivider.Location = New System.Drawing.Point(20, 234)
        Me.lblDivider.Name = "lblDivider"
        Me.lblDivider.Size = New System.Drawing.Size(133, 17)
        Me.lblDivider.TabIndex = 8
        Me.lblDivider.Text = "-------------------------"
        '
        'gbAdditional
        '
        Me.gbAdditional.Controls.Add(Me.lblRemember)
        Me.gbAdditional.Controls.Add(Me.cmdInsert)
        Me.gbAdditional.Controls.Add(Me.lblCHARSTABILITYPERIOD)
        Me.gbAdditional.Controls.Add(Me.cmdBuild)
        Me.gbAdditional.Controls.Add(Me.CHARSTABILITYPERIOD)
        Me.gbAdditional.Controls.Add(Me.panTP)
        Me.gbAdditional.Controls.Add(Me.panCycles)
        Me.gbAdditional.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.gbAdditional.Location = New System.Drawing.Point(3, 5)
        Me.gbAdditional.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbAdditional.Name = "gbAdditional"
        Me.gbAdditional.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbAdditional.Size = New System.Drawing.Size(451, 340)
        Me.gbAdditional.TabIndex = 142
        Me.gbAdditional.TabStop = False
        Me.gbAdditional.Text = "Stability Condition Information"
        '
        'lblRemember
        '
        Me.lblRemember.AutoSize = True
        Me.lblRemember.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRemember.ForeColor = System.Drawing.Color.Red
        Me.lblRemember.Location = New System.Drawing.Point(70, 263)
        Me.lblRemember.Name = "lblRemember"
        Me.lblRemember.Size = New System.Drawing.Size(378, 17)
        Me.lblRemember.TabIndex = 11
        Me.lblRemember.Text = "Remember to click Build to modify Stability conditions summary"
        Me.lblRemember.Visible = False
        '
        'cmdInsert
        '
        Me.cmdInsert.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.cmdInsert.ForeColor = System.Drawing.Color.Blue
        Me.cmdInsert.Location = New System.Drawing.Point(5, 300)
        Me.cmdInsert.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdInsert.Name = "cmdInsert"
        Me.cmdInsert.Size = New System.Drawing.Size(62, 27)
        Me.cmdInsert.TabIndex = 10
        Me.cmdInsert.Text = "&Insert"
        Me.cmdInsert.UseVisualStyleBackColor = True
        '
        'lblCHARSTABILITYPERIOD
        '
        Me.lblCHARSTABILITYPERIOD.AutoSize = True
        Me.lblCHARSTABILITYPERIOD.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblCHARSTABILITYPERIOD.Location = New System.Drawing.Point(70, 280)
        Me.lblCHARSTABILITYPERIOD.Name = "lblCHARSTABILITYPERIOD"
        Me.lblCHARSTABILITYPERIOD.Size = New System.Drawing.Size(177, 17)
        Me.lblCHARSTABILITYPERIOD.TabIndex = 9
        Me.lblCHARSTABILITYPERIOD.Text = "Stability conditions summary:"
        '
        'cmdBuild
        '
        Me.cmdBuild.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.cmdBuild.ForeColor = System.Drawing.Color.Blue
        Me.cmdBuild.Location = New System.Drawing.Point(5, 270)
        Me.cmdBuild.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdBuild.Name = "cmdBuild"
        Me.cmdBuild.Size = New System.Drawing.Size(62, 27)
        Me.cmdBuild.TabIndex = 8
        Me.cmdBuild.Text = "&Build"
        Me.cmdBuild.UseVisualStyleBackColor = True
        '
        'CHARSTABILITYPERIOD
        '
        Me.CHARSTABILITYPERIOD.BackColor = System.Drawing.Color.White
        Me.CHARSTABILITYPERIOD.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARSTABILITYPERIOD.Location = New System.Drawing.Point(73, 301)
        Me.CHARSTABILITYPERIOD.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARSTABILITYPERIOD.Name = "CHARSTABILITYPERIOD"
        Me.CHARSTABILITYPERIOD.Size = New System.Drawing.Size(250, 25)
        Me.CHARSTABILITYPERIOD.TabIndex = 7
        Me.CHARSTABILITYPERIOD.TabStop = False
        Me.CHARSTABILITYPERIOD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'panTP
        '
        Me.panTP.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panTP.Controls.Add(Me.lblPeriodTemp)
        Me.panTP.Controls.Add(Me.chkCONVERTTEMP)
        Me.panTP.Controls.Add(Me.chkCONVERTTIME)
        Me.panTP.Controls.Add(Me.CHARPERIODTEMP)
        Me.panTP.Controls.Add(Me.Label3)
        Me.panTP.Controls.Add(Me.CHARTIMEFRAME)
        Me.panTP.Controls.Add(Me.Label2)
        Me.panTP.Controls.Add(Me.CHARTIMEPERIOD)
        Me.panTP.Controls.Add(Me.Label1)
        Me.panTP.Location = New System.Drawing.Point(5, 98)
        Me.panTP.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panTP.Name = "panTP"
        Me.panTP.Size = New System.Drawing.Size(432, 163)
        Me.panTP.TabIndex = 3
        '
        'lblPeriodTemp
        '
        Me.lblPeriodTemp.AutoSize = True
        Me.lblPeriodTemp.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPeriodTemp.Location = New System.Drawing.Point(3, 0)
        Me.lblPeriodTemp.Name = "lblPeriodTemp"
        Me.lblPeriodTemp.Size = New System.Drawing.Size(174, 15)
        Me.lblPeriodTemp.TabIndex = 10
        Me.lblPeriodTemp.Text = "[Period Temp] information"
        '
        'chkCONVERTTEMP
        '
        Me.chkCONVERTTEMP.AutoSize = True
        Me.chkCONVERTTEMP.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkCONVERTTEMP.Location = New System.Drawing.Point(27, 135)
        Me.chkCONVERTTEMP.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkCONVERTTEMP.Name = "chkCONVERTTEMP"
        Me.chkCONVERTTEMP.Size = New System.Drawing.Size(133, 21)
        Me.chkCONVERTTEMP.TabIndex = 5
        Me.chkCONVERTTEMP.TabStop = False
        Me.chkCONVERTTEMP.Text = "Convert 'deg C' to"
        Me.chkCONVERTTEMP.UseVisualStyleBackColor = True
        '
        'chkCONVERTTIME
        '
        Me.chkCONVERTTIME.AutoSize = True
        Me.chkCONVERTTIME.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkCONVERTTIME.Location = New System.Drawing.Point(27, 47)
        Me.chkCONVERTTIME.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkCONVERTTIME.Name = "chkCONVERTTIME"
        Me.chkCONVERTTIME.Size = New System.Drawing.Size(142, 21)
        Me.chkCONVERTTIME.TabIndex = 2
        Me.chkCONVERTTIME.TabStop = False
        Me.chkCONVERTTIME.Text = "Convert time to text"
        Me.chkCONVERTTIME.UseVisualStyleBackColor = True
        '
        'CHARPERIODTEMP
        '
        Me.CHARPERIODTEMP.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARPERIODTEMP.Location = New System.Drawing.Point(248, 105)
        Me.CHARPERIODTEMP.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARPERIODTEMP.Multiline = True
        Me.CHARPERIODTEMP.Name = "CHARPERIODTEMP"
        Me.CHARPERIODTEMP.Size = New System.Drawing.Size(179, 50)
        Me.CHARPERIODTEMP.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label3.Location = New System.Drawing.Point(3, 97)
        Me.Label3.MaximumSize = New System.Drawing.Size(222, 34)
        Me.Label3.MinimumSize = New System.Drawing.Size(187, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(212, 34)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Enter temperature (e.g. -70 deg C, Room Temp, Refrigerated, etc.):"
        '
        'CHARTIMEFRAME
        '
        Me.CHARTIMEFRAME.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARTIMEFRAME.Location = New System.Drawing.Point(248, 74)
        Me.CHARTIMEFRAME.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARTIMEFRAME.Name = "CHARTIMEFRAME"
        Me.CHARTIMEFRAME.Size = New System.Drawing.Size(81, 25)
        Me.CHARTIMEFRAME.TabIndex = 3
        Me.CHARTIMEFRAME.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label2.Location = New System.Drawing.Point(3, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(242, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Enter time frame (e.g. Hours, Days, etc.):"
        '
        'CHARTIMEPERIOD
        '
        Me.CHARTIMEPERIOD.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARTIMEPERIOD.Location = New System.Drawing.Point(248, 24)
        Me.CHARTIMEPERIOD.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARTIMEPERIOD.Name = "CHARTIMEPERIOD"
        Me.CHARTIMEPERIOD.Size = New System.Drawing.Size(81, 25)
        Me.CHARTIMEPERIOD.TabIndex = 1
        Me.CHARTIMEPERIOD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.Label1.Location = New System.Drawing.Point(3, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Enter time period (e.g. 1, 5, 10, etc):"
        '
        'panCycles
        '
        Me.panCycles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panCycles.Controls.Add(Me.Label15)
        Me.panCycles.Controls.Add(Me.INTNUMBEROFCYCLES)
        Me.panCycles.Controls.Add(Me.lblCycles)
        Me.panCycles.Location = New System.Drawing.Point(5, 25)
        Me.panCycles.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panCycles.Name = "panCycles"
        Me.panCycles.Size = New System.Drawing.Size(220, 70)
        Me.panCycles.TabIndex = 2
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(3, 1)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(141, 15)
        Me.Label15.TabIndex = 11
        Me.Label15.Text = "[#Cycles] information"
        '
        'INTNUMBEROFCYCLES
        '
        Me.INTNUMBEROFCYCLES.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.INTNUMBEROFCYCLES.Location = New System.Drawing.Point(150, 29)
        Me.INTNUMBEROFCYCLES.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.INTNUMBEROFCYCLES.Name = "INTNUMBEROFCYCLES"
        Me.INTNUMBEROFCYCLES.Size = New System.Drawing.Size(34, 25)
        Me.INTNUMBEROFCYCLES.TabIndex = 1
        Me.INTNUMBEROFCYCLES.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblCycles
        '
        Me.lblCycles.AutoSize = True
        Me.lblCycles.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblCycles.Location = New System.Drawing.Point(0, 37)
        Me.lblCycles.Name = "lblCycles"
        Me.lblCycles.Size = New System.Drawing.Size(144, 17)
        Me.lblCycles.TabIndex = 0
        Me.lblCycles.Text = "Enter number of cycles:"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmsFieldCodes})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(172, 26)
        '
        'cmsFieldCodes
        '
        Me.cmsFieldCodes.Name = "cmsFieldCodes"
        Me.cmsFieldCodes.Size = New System.Drawing.Size(171, 22)
        Me.cmsFieldCodes.Text = "Insert Field Code..."
        '
        'cmdIncSamples
        '
        Me.cmdIncSamples.Enabled = False
        Me.cmdIncSamples.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdIncSamples.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdIncSamples.Location = New System.Drawing.Point(692, 229)
        Me.cmdIncSamples.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdIncSamples.Name = "cmdIncSamples"
        Me.cmdIncSamples.Size = New System.Drawing.Size(175, 55)
        Me.cmdIncSamples.TabIndex = 144
        Me.cmdIncSamples.Text = "Configure &Incurred Samples Criteria..."
        Me.cmdIncSamples.UseVisualStyleBackColor = True
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblTitle.Location = New System.Drawing.Point(14, 24)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(49, 16)
        Me.lblTitle.TabIndex = 145
        Me.lblTitle.Text = "Label4"
        '
        'gbAnovaStats
        '
        Me.gbAnovaStats.Controls.Add(Me.chkIntraRunSumStats)
        Me.gbAnovaStats.Controls.Add(Me.chkIncludeAnovaSumStats)
        Me.gbAnovaStats.Controls.Add(Me.chkIncludeAnova)
        Me.gbAnovaStats.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbAnovaStats.Location = New System.Drawing.Point(15, 184)
        Me.gbAnovaStats.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbAnovaStats.Name = "gbAnovaStats"
        Me.gbAnovaStats.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbAnovaStats.Size = New System.Drawing.Size(307, 88)
        Me.gbAnovaStats.TabIndex = 150
        Me.gbAnovaStats.TabStop = False
        Me.gbAnovaStats.Text = "Summary Statistics Options"
        '
        'chkIntraRunSumStats
        '
        Me.chkIntraRunSumStats.AutoSize = True
        Me.chkIntraRunSumStats.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIntraRunSumStats.Location = New System.Drawing.Point(12, 43)
        Me.chkIntraRunSumStats.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkIntraRunSumStats.Name = "chkIntraRunSumStats"
        Me.chkIntraRunSumStats.Size = New System.Drawing.Size(250, 19)
        Me.chkIntraRunSumStats.TabIndex = 2
        Me.chkIntraRunSumStats.Text = "Include Intra-Run Summary Stats Section"
        Me.chkIntraRunSumStats.UseVisualStyleBackColor = True
        '
        'chkIncludeAnovaSumStats
        '
        Me.chkIncludeAnovaSumStats.AutoSize = True
        Me.chkIncludeAnovaSumStats.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIncludeAnovaSumStats.Location = New System.Drawing.Point(12, 22)
        Me.chkIncludeAnovaSumStats.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkIncludeAnovaSumStats.Name = "chkIncludeAnovaSumStats"
        Me.chkIncludeAnovaSumStats.Size = New System.Drawing.Size(250, 19)
        Me.chkIncludeAnovaSumStats.TabIndex = 1
        Me.chkIncludeAnovaSumStats.Text = "Include Inter-Run Summary Stats Section"
        Me.chkIncludeAnovaSumStats.UseVisualStyleBackColor = True
        '
        'chkIncludeAnova
        '
        Me.chkIncludeAnova.AutoSize = True
        Me.chkIncludeAnova.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIncludeAnova.Location = New System.Drawing.Point(12, 64)
        Me.chkIncludeAnova.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkIncludeAnova.Name = "chkIncludeAnova"
        Me.chkIncludeAnova.Size = New System.Drawing.Size(152, 19)
        Me.chkIncludeAnova.TabIndex = 0
        Me.chkIncludeAnova.Text = "Include ANOVA Section"
        Me.chkIncludeAnova.UseVisualStyleBackColor = True
        '
        'gbSampleSort
        '
        Me.gbSampleSort.Controls.Add(Me.Label17)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleSAD6)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleS6)
        Me.gbSampleSort.Controls.Add(Me.Label16)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleSAD5)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleS5)
        Me.gbSampleSort.Controls.Add(Me.Label7)
        Me.gbSampleSort.Controls.Add(Me.Label6)
        Me.gbSampleSort.Controls.Add(Me.Label5)
        Me.gbSampleSort.Controls.Add(Me.Label4)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleSAD4)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleS4)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleSAD3)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleS3)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleSAD2)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleS2)
        Me.gbSampleSort.Controls.Add(Me.lblA)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleSAD1)
        Me.gbSampleSort.Controls.Add(Me.cbxSampleS1)
        Me.gbSampleSort.Controls.Add(Me.lblLevel)
        Me.gbSampleSort.Controls.Add(Me.lblSort)
        Me.gbSampleSort.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSampleSort.Location = New System.Drawing.Point(12, 443)
        Me.gbSampleSort.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSampleSort.Name = "gbSampleSort"
        Me.gbSampleSort.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSampleSort.Size = New System.Drawing.Size(307, 216)
        Me.gbSampleSort.TabIndex = 151
        Me.gbSampleSort.TabStop = False
        Me.gbSampleSort.Text = "Sort Rules"
        Me.gbSampleSort.Visible = False
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(24, 187)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(15, 17)
        Me.Label17.TabIndex = 21
        Me.Label17.Text = "6"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbxSampleSAD6
        '
        Me.cbxSampleSAD6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleSAD6.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleSAD6.FormattingEnabled = True
        Me.cbxSampleSAD6.Location = New System.Drawing.Point(212, 184)
        Me.cbxSampleSAD6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleSAD6.Name = "cbxSampleSAD6"
        Me.cbxSampleSAD6.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleSAD6.TabIndex = 20
        '
        'cbxSampleS6
        '
        Me.cbxSampleS6.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleS6.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleS6.FormattingEnabled = True
        Me.cbxSampleS6.Location = New System.Drawing.Point(62, 184)
        Me.cbxSampleS6.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleS6.Name = "cbxSampleS6"
        Me.cbxSampleS6.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleS6.TabIndex = 19
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Location = New System.Drawing.Point(24, 158)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(15, 17)
        Me.Label16.TabIndex = 18
        Me.Label16.Text = "5"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbxSampleSAD5
        '
        Me.cbxSampleSAD5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleSAD5.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleSAD5.FormattingEnabled = True
        Me.cbxSampleSAD5.Location = New System.Drawing.Point(212, 155)
        Me.cbxSampleSAD5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleSAD5.Name = "cbxSampleSAD5"
        Me.cbxSampleSAD5.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleSAD5.TabIndex = 17
        '
        'cbxSampleS5
        '
        Me.cbxSampleS5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleS5.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleS5.FormattingEnabled = True
        Me.cbxSampleS5.Location = New System.Drawing.Point(62, 155)
        Me.cbxSampleS5.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleS5.Name = "cbxSampleS5"
        Me.cbxSampleS5.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleS5.TabIndex = 16
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(24, 130)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(15, 17)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "4"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(24, 101)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(15, 17)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "3"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(24, 72)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(15, 17)
        Me.Label5.TabIndex = 13
        Me.Label5.Text = "2"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(24, 43)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(15, 17)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "1"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbxSampleSAD4
        '
        Me.cbxSampleSAD4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleSAD4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleSAD4.FormattingEnabled = True
        Me.cbxSampleSAD4.Location = New System.Drawing.Point(212, 126)
        Me.cbxSampleSAD4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleSAD4.Name = "cbxSampleSAD4"
        Me.cbxSampleSAD4.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleSAD4.TabIndex = 11
        '
        'cbxSampleS4
        '
        Me.cbxSampleS4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleS4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleS4.FormattingEnabled = True
        Me.cbxSampleS4.Location = New System.Drawing.Point(62, 126)
        Me.cbxSampleS4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleS4.Name = "cbxSampleS4"
        Me.cbxSampleS4.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleS4.TabIndex = 10
        '
        'cbxSampleSAD3
        '
        Me.cbxSampleSAD3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleSAD3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleSAD3.FormattingEnabled = True
        Me.cbxSampleSAD3.Location = New System.Drawing.Point(212, 97)
        Me.cbxSampleSAD3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleSAD3.Name = "cbxSampleSAD3"
        Me.cbxSampleSAD3.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleSAD3.TabIndex = 9
        '
        'cbxSampleS3
        '
        Me.cbxSampleS3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleS3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleS3.FormattingEnabled = True
        Me.cbxSampleS3.Location = New System.Drawing.Point(62, 97)
        Me.cbxSampleS3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleS3.Name = "cbxSampleS3"
        Me.cbxSampleS3.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleS3.TabIndex = 8
        '
        'cbxSampleSAD2
        '
        Me.cbxSampleSAD2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleSAD2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleSAD2.FormattingEnabled = True
        Me.cbxSampleSAD2.Location = New System.Drawing.Point(212, 68)
        Me.cbxSampleSAD2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleSAD2.Name = "cbxSampleSAD2"
        Me.cbxSampleSAD2.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleSAD2.TabIndex = 7
        '
        'cbxSampleS2
        '
        Me.cbxSampleS2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleS2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleS2.FormattingEnabled = True
        Me.cbxSampleS2.Location = New System.Drawing.Point(62, 68)
        Me.cbxSampleS2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleS2.Name = "cbxSampleS2"
        Me.cbxSampleS2.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleS2.TabIndex = 6
        '
        'lblA
        '
        Me.lblA.Location = New System.Drawing.Point(212, 20)
        Me.lblA.Name = "lblA"
        Me.lblA.Size = New System.Drawing.Size(83, 18)
        Me.lblA.TabIndex = 5
        Me.lblA.Text = "Asc/Desc"
        Me.lblA.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'cbxSampleSAD1
        '
        Me.cbxSampleSAD1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleSAD1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleSAD1.FormattingEnabled = True
        Me.cbxSampleSAD1.Location = New System.Drawing.Point(212, 39)
        Me.cbxSampleSAD1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleSAD1.Name = "cbxSampleSAD1"
        Me.cbxSampleSAD1.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleSAD1.TabIndex = 4
        '
        'cbxSampleS1
        '
        Me.cbxSampleS1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleS1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleS1.FormattingEnabled = True
        Me.cbxSampleS1.Location = New System.Drawing.Point(62, 39)
        Me.cbxSampleS1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleS1.Name = "cbxSampleS1"
        Me.cbxSampleS1.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleS1.TabIndex = 3
        '
        'lblLevel
        '
        Me.lblLevel.AutoSize = True
        Me.lblLevel.Location = New System.Drawing.Point(13, 20)
        Me.lblLevel.Name = "lblLevel"
        Me.lblLevel.Size = New System.Drawing.Size(40, 17)
        Me.lblLevel.TabIndex = 2
        Me.lblLevel.Text = "Level"
        '
        'lblSort
        '
        Me.lblSort.Location = New System.Drawing.Point(62, 20)
        Me.lblSort.Name = "lblSort"
        Me.lblSort.Size = New System.Drawing.Size(143, 17)
        Me.lblSort.TabIndex = 0
        Me.lblSort.Text = "Sort Column"
        Me.lblSort.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'gbSampleGroup
        '
        Me.gbSampleGroup.Controls.Add(Me.Label8)
        Me.gbSampleGroup.Controls.Add(Me.Label9)
        Me.gbSampleGroup.Controls.Add(Me.Label10)
        Me.gbSampleGroup.Controls.Add(Me.Label11)
        Me.gbSampleGroup.Controls.Add(Me.cbxSampleGAD4)
        Me.gbSampleGroup.Controls.Add(Me.cbxSampleG4)
        Me.gbSampleGroup.Controls.Add(Me.cbxSampleGAD3)
        Me.gbSampleGroup.Controls.Add(Me.cbxSampleG3)
        Me.gbSampleGroup.Controls.Add(Me.cbxSampleGAD2)
        Me.gbSampleGroup.Controls.Add(Me.cbxSampleG2)
        Me.gbSampleGroup.Controls.Add(Me.Label12)
        Me.gbSampleGroup.Controls.Add(Me.cbxSampleGAD1)
        Me.gbSampleGroup.Controls.Add(Me.cbxSampleG1)
        Me.gbSampleGroup.Controls.Add(Me.Label13)
        Me.gbSampleGroup.Controls.Add(Me.Label14)
        Me.gbSampleGroup.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSampleGroup.Location = New System.Drawing.Point(14, 277)
        Me.gbSampleGroup.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSampleGroup.Name = "gbSampleGroup"
        Me.gbSampleGroup.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSampleGroup.Size = New System.Drawing.Size(307, 160)
        Me.gbSampleGroup.TabIndex = 152
        Me.gbSampleGroup.TabStop = False
        Me.gbSampleGroup.Text = "Group Rules"
        Me.gbSampleGroup.Visible = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(24, 130)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(15, 17)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "4"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(24, 101)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(15, 17)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "3"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(24, 72)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(15, 17)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "2"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(24, 43)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(15, 17)
        Me.Label11.TabIndex = 12
        Me.Label11.Text = "1"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cbxSampleGAD4
        '
        Me.cbxSampleGAD4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleGAD4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleGAD4.FormattingEnabled = True
        Me.cbxSampleGAD4.Location = New System.Drawing.Point(212, 126)
        Me.cbxSampleGAD4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleGAD4.Name = "cbxSampleGAD4"
        Me.cbxSampleGAD4.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleGAD4.TabIndex = 11
        '
        'cbxSampleG4
        '
        Me.cbxSampleG4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleG4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleG4.FormattingEnabled = True
        Me.cbxSampleG4.Location = New System.Drawing.Point(62, 126)
        Me.cbxSampleG4.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleG4.Name = "cbxSampleG4"
        Me.cbxSampleG4.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleG4.TabIndex = 10
        '
        'cbxSampleGAD3
        '
        Me.cbxSampleGAD3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleGAD3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleGAD3.FormattingEnabled = True
        Me.cbxSampleGAD3.Location = New System.Drawing.Point(212, 97)
        Me.cbxSampleGAD3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleGAD3.Name = "cbxSampleGAD3"
        Me.cbxSampleGAD3.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleGAD3.TabIndex = 9
        '
        'cbxSampleG3
        '
        Me.cbxSampleG3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleG3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleG3.FormattingEnabled = True
        Me.cbxSampleG3.Location = New System.Drawing.Point(62, 97)
        Me.cbxSampleG3.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleG3.Name = "cbxSampleG3"
        Me.cbxSampleG3.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleG3.TabIndex = 8
        '
        'cbxSampleGAD2
        '
        Me.cbxSampleGAD2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleGAD2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleGAD2.FormattingEnabled = True
        Me.cbxSampleGAD2.Location = New System.Drawing.Point(212, 68)
        Me.cbxSampleGAD2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleGAD2.Name = "cbxSampleGAD2"
        Me.cbxSampleGAD2.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleGAD2.TabIndex = 7
        '
        'cbxSampleG2
        '
        Me.cbxSampleG2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleG2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleG2.FormattingEnabled = True
        Me.cbxSampleG2.Location = New System.Drawing.Point(62, 68)
        Me.cbxSampleG2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleG2.Name = "cbxSampleG2"
        Me.cbxSampleG2.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleG2.TabIndex = 6
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(212, 19)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(83, 18)
        Me.Label12.TabIndex = 5
        Me.Label12.Text = "Asc/Desc"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'cbxSampleGAD1
        '
        Me.cbxSampleGAD1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleGAD1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleGAD1.FormattingEnabled = True
        Me.cbxSampleGAD1.Location = New System.Drawing.Point(212, 39)
        Me.cbxSampleGAD1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleGAD1.Name = "cbxSampleGAD1"
        Me.cbxSampleGAD1.Size = New System.Drawing.Size(69, 25)
        Me.cbxSampleGAD1.TabIndex = 4
        '
        'cbxSampleG1
        '
        Me.cbxSampleG1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleG1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSampleG1.FormattingEnabled = True
        Me.cbxSampleG1.Location = New System.Drawing.Point(62, 39)
        Me.cbxSampleG1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSampleG1.Name = "cbxSampleG1"
        Me.cbxSampleG1.Size = New System.Drawing.Size(143, 25)
        Me.cbxSampleG1.TabIndex = 3
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(13, 20)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(40, 17)
        Me.Label13.TabIndex = 2
        Me.Label13.Text = "Level"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(62, 19)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(143, 17)
        Me.Label14.TabIndex = 0
        Me.Label14.Text = "Group Column"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'chkIncludePSAE
        '
        Me.chkIncludePSAE.AutoSize = True
        Me.chkIncludePSAE.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkIncludePSAE.Location = New System.Drawing.Point(13, 22)
        Me.chkIncludePSAE.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkIncludePSAE.Name = "chkIncludePSAE"
        Me.chkIncludePSAE.Size = New System.Drawing.Size(154, 21)
        Me.chkIncludePSAE.TabIndex = 8
        Me.chkIncludePSAE.Text = "Include PSAE Samples"
        Me.chkIncludePSAE.UseVisualStyleBackColor = True
        '
        'gbPSAE
        '
        Me.gbPSAE.Controls.Add(Me.chkInjCol)
        Me.gbPSAE.Controls.Add(Me.chkIncludePSAE)
        Me.gbPSAE.Enabled = False
        Me.gbPSAE.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.gbPSAE.Location = New System.Drawing.Point(368, 699)
        Me.gbPSAE.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbPSAE.Name = "gbPSAE"
        Me.gbPSAE.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbPSAE.Size = New System.Drawing.Size(324, 75)
        Me.gbPSAE.TabIndex = 153
        Me.gbPSAE.TabStop = False
        Me.gbPSAE.Text = "Include PSAE data"
        '
        'chkInjCol
        '
        Me.chkInjCol.AutoSize = True
        Me.chkInjCol.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkInjCol.Location = New System.Drawing.Point(13, 49)
        Me.chkInjCol.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkInjCol.Name = "chkInjCol"
        Me.chkInjCol.Size = New System.Drawing.Size(183, 21)
        Me.chkInjCol.TabIndex = 9
        Me.chkInjCol.Text = "Exclude Injection # Column"
        Me.chkInjCol.UseVisualStyleBackColor = True
        '
        'gbResultsChoice
        '
        Me.gbResultsChoice.Controls.Add(Me.rbUseISPeakArea)
        Me.gbResultsChoice.Controls.Add(Me.chkBOOLADHOCSTABCOMPCOLUMNS)
        Me.gbResultsChoice.Controls.Add(Me.panIS)
        Me.gbResultsChoice.Controls.Add(Me.rbUsePeakAreaRatio)
        Me.gbResultsChoice.Controls.Add(Me.rbUsePeakArea)
        Me.gbResultsChoice.Controls.Add(Me.rbConc)
        Me.gbResultsChoice.Controls.Add(Me.chkBOOLDOINDREC)
        Me.gbResultsChoice.Enabled = False
        Me.gbResultsChoice.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbResultsChoice.Location = New System.Drawing.Point(339, 485)
        Me.gbResultsChoice.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbResultsChoice.Name = "gbResultsChoice"
        Me.gbResultsChoice.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbResultsChoice.Size = New System.Drawing.Size(324, 203)
        Me.gbResultsChoice.TabIndex = 154
        Me.gbResultsChoice.TabStop = False
        Me.gbResultsChoice.Text = "Choose Results Type"
        Me.gbResultsChoice.Visible = False
        '
        'rbUseISPeakArea
        '
        Me.rbUseISPeakArea.AutoSize = True
        Me.rbUseISPeakArea.BackColor = System.Drawing.Color.Transparent
        Me.rbUseISPeakArea.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbUseISPeakArea.ForeColor = System.Drawing.Color.Black
        Me.rbUseISPeakArea.Location = New System.Drawing.Point(12, 64)
        Me.rbUseISPeakArea.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbUseISPeakArea.Name = "rbUseISPeakArea"
        Me.rbUseISPeakArea.Size = New System.Drawing.Size(151, 21)
        Me.rbUseISPeakArea.TabIndex = 11
        Me.rbUseISPeakArea.Text = "Use Int Std Peak Area"
        Me.rbUseISPeakArea.UseVisualStyleBackColor = False
        '
        'chkBOOLADHOCSTABCOMPCOLUMNS
        '
        Me.chkBOOLADHOCSTABCOMPCOLUMNS.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBOOLADHOCSTABCOMPCOLUMNS.Location = New System.Drawing.Point(41, 33)
        Me.chkBOOLADHOCSTABCOMPCOLUMNS.Name = "chkBOOLADHOCSTABCOMPCOLUMNS"
        Me.chkBOOLADHOCSTABCOMPCOLUMNS.Size = New System.Drawing.Size(287, 43)
        Me.chkBOOLADHOCSTABCOMPCOLUMNS.TabIndex = 10
        Me.chkBOOLADHOCSTABCOMPCOLUMNS.Text = "Format table as a grid with comparison point values in different columns"
        Me.chkBOOLADHOCSTABCOMPCOLUMNS.UseVisualStyleBackColor = True
        Me.chkBOOLADHOCSTABCOMPCOLUMNS.Visible = False
        '
        'panIS
        '
        Me.panIS.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panIS.Controls.Add(Me.CHARISCONC)
        Me.panIS.Controls.Add(Me.chkBOOLISCOMBINELEVELS)
        Me.panIS.Controls.Add(Me.chkIncludeIS)
        Me.panIS.Controls.Add(Me.lblCHARISCONC)
        Me.panIS.Controls.Add(Me.chkCustomLeg)
        Me.panIS.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.panIS.Location = New System.Drawing.Point(0, 108)
        Me.panIS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panIS.Name = "panIS"
        Me.panIS.Size = New System.Drawing.Size(317, 89)
        Me.panIS.TabIndex = 8
        '
        'CHARISCONC
        '
        Me.CHARISCONC.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARISCONC.Location = New System.Drawing.Point(240, 49)
        Me.CHARISCONC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARISCONC.Name = "CHARISCONC"
        Me.CHARISCONC.Size = New System.Drawing.Size(73, 25)
        Me.CHARISCONC.TabIndex = 6
        '
        'chkBOOLISCOMBINELEVELS
        '
        Me.chkBOOLISCOMBINELEVELS.AutoSize = True
        Me.chkBOOLISCOMBINELEVELS.Checked = True
        Me.chkBOOLISCOMBINELEVELS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBOOLISCOMBINELEVELS.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkBOOLISCOMBINELEVELS.Location = New System.Drawing.Point(10, 29)
        Me.chkBOOLISCOMBINELEVELS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkBOOLISCOMBINELEVELS.Name = "chkBOOLISCOMBINELEVELS"
        Me.chkBOOLISCOMBINELEVELS.Size = New System.Drawing.Size(249, 21)
        Me.chkBOOLISCOMBINELEVELS.TabIndex = 8
        Me.chkBOOLISCOMBINELEVELS.Text = "Combine all Int Std data into one level"
        Me.chkBOOLISCOMBINELEVELS.UseVisualStyleBackColor = True
        '
        'chkIncludeIS
        '
        Me.chkIncludeIS.AutoSize = True
        Me.chkIncludeIS.Checked = True
        Me.chkIncludeIS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkIncludeIS.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkIncludeIS.Location = New System.Drawing.Point(10, 5)
        Me.chkIncludeIS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkIncludeIS.Name = "chkIncludeIS"
        Me.chkIncludeIS.Size = New System.Drawing.Size(145, 21)
        Me.chkIncludeIS.TabIndex = 5
        Me.chkIncludeIS.Text = "Include Int Std Table"
        Me.chkIncludeIS.UseVisualStyleBackColor = True
        '
        'lblCHARISCONC
        '
        Me.lblCHARISCONC.AutoSize = True
        Me.lblCHARISCONC.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblCHARISCONC.Location = New System.Drawing.Point(10, 54)
        Me.lblCHARISCONC.Name = "lblCHARISCONC"
        Me.lblCHARISCONC.Size = New System.Drawing.Size(222, 17)
        Me.lblCHARISCONC.TabIndex = 7
        Me.lblCHARISCONC.Text = "Int Std Conc (Optional: include units):"
        Me.lblCHARISCONC.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkCustomLeg
        '
        Me.chkCustomLeg.AutoSize = True
        Me.chkCustomLeg.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCustomLeg.Location = New System.Drawing.Point(161, 5)
        Me.chkCustomLeg.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkCustomLeg.Name = "chkCustomLeg"
        Me.chkCustomLeg.Size = New System.Drawing.Size(113, 21)
        Me.chkCustomLeg.TabIndex = 2
        Me.chkCustomLeg.Text = "Do only Int Std"
        Me.chkCustomLeg.UseVisualStyleBackColor = True
        '
        'rbUsePeakAreaRatio
        '
        Me.rbUsePeakAreaRatio.AutoSize = True
        Me.rbUsePeakAreaRatio.BackColor = System.Drawing.Color.Transparent
        Me.rbUsePeakAreaRatio.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbUsePeakAreaRatio.ForeColor = System.Drawing.Color.Black
        Me.rbUsePeakAreaRatio.Location = New System.Drawing.Point(12, 85)
        Me.rbUsePeakAreaRatio.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbUsePeakAreaRatio.Name = "rbUsePeakAreaRatio"
        Me.rbUsePeakAreaRatio.Size = New System.Drawing.Size(227, 21)
        Me.rbUsePeakAreaRatio.TabIndex = 3
        Me.rbUsePeakAreaRatio.Text = "Use Peak Area Ratio (if applicable)"
        Me.rbUsePeakAreaRatio.UseVisualStyleBackColor = False
        '
        'rbUsePeakArea
        '
        Me.rbUsePeakArea.AutoSize = True
        Me.rbUsePeakArea.BackColor = System.Drawing.Color.Transparent
        Me.rbUsePeakArea.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbUsePeakArea.ForeColor = System.Drawing.Color.Black
        Me.rbUsePeakArea.Location = New System.Drawing.Point(12, 43)
        Me.rbUsePeakArea.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbUsePeakArea.Name = "rbUsePeakArea"
        Me.rbUsePeakArea.Size = New System.Drawing.Size(156, 21)
        Me.rbUsePeakArea.TabIndex = 2
        Me.rbUsePeakArea.Text = "Use Analyte Peak Area"
        Me.rbUsePeakArea.UseVisualStyleBackColor = False
        '
        'rbConc
        '
        Me.rbConc.AutoSize = True
        Me.rbConc.BackColor = System.Drawing.Color.Transparent
        Me.rbConc.Checked = True
        Me.rbConc.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbConc.ForeColor = System.Drawing.Color.Black
        Me.rbConc.Location = New System.Drawing.Point(12, 22)
        Me.rbConc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbConc.Name = "rbConc"
        Me.rbConc.Size = New System.Drawing.Size(133, 21)
        Me.rbConc.TabIndex = 1
        Me.rbConc.TabStop = True
        Me.rbConc.Text = "Use Concentration"
        Me.rbConc.UseVisualStyleBackColor = False
        '
        'chkBOOLDOINDREC
        '
        Me.chkBOOLDOINDREC.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBOOLDOINDREC.Location = New System.Drawing.Point(161, 11)
        Me.chkBOOLDOINDREC.Name = "chkBOOLDOINDREC"
        Me.chkBOOLDOINDREC.Size = New System.Drawing.Size(167, 58)
        Me.chkBOOLDOINDREC.TabIndex = 9
        Me.chkBOOLDOINDREC.Text = "Calculate individual" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Recovery/MatrixFactor" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "values"
        Me.chkBOOLDOINDREC.UseVisualStyleBackColor = True
        '
        'gbTableLegend
        '
        Me.gbTableLegend.Controls.Add(Me.chkNoneLeg)
        Me.gbTableLegend.Controls.Add(Me.panTitleLegends)
        Me.gbTableLegend.Controls.Add(Me.panNomDenomCalcs)
        Me.gbTableLegend.Enabled = False
        Me.gbTableLegend.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbTableLegend.Location = New System.Drawing.Point(799, 614)
        Me.gbTableLegend.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbTableLegend.Name = "gbTableLegend"
        Me.gbTableLegend.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbTableLegend.Size = New System.Drawing.Size(327, 328)
        Me.gbTableLegend.TabIndex = 155
        Me.gbTableLegend.TabStop = False
        Me.gbTableLegend.Text = "Table Legends"
        '
        'chkNoneLeg
        '
        Me.chkNoneLeg.AutoSize = True
        Me.chkNoneLeg.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNoneLeg.Location = New System.Drawing.Point(13, 19)
        Me.chkNoneLeg.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkNoneLeg.Name = "chkNoneLeg"
        Me.chkNoneLeg.Size = New System.Drawing.Size(92, 21)
        Me.chkNoneLeg.TabIndex = 12
        Me.chkNoneLeg.Text = "No Legend"
        Me.chkNoneLeg.UseVisualStyleBackColor = True
        '
        'panTitleLegends
        '
        Me.panTitleLegends.Controls.Add(Me.CHARTITLELEG)
        Me.panTitleLegends.Controls.Add(Me.CHARNUMLEG)
        Me.panTitleLegends.Controls.Add(Me.lblCHARDENLEG)
        Me.panTitleLegends.Controls.Add(Me.lblCHARTITLELEG)
        Me.panTitleLegends.Controls.Add(Me.CHARDENLEG)
        Me.panTitleLegends.Controls.Add(Me.lblCHARNUMLEG)
        Me.panTitleLegends.Location = New System.Drawing.Point(12, 189)
        Me.panTitleLegends.Name = "panTitleLegends"
        Me.panTitleLegends.Size = New System.Drawing.Size(305, 129)
        Me.panTitleLegends.TabIndex = 164
        '
        'CHARTITLELEG
        '
        Me.CHARTITLELEG.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CHARTITLELEG.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARTITLELEG.Location = New System.Drawing.Point(65, 0)
        Me.CHARTITLELEG.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARTITLELEG.Name = "CHARTITLELEG"
        Me.CHARTITLELEG.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.CHARTITLELEG.Size = New System.Drawing.Size(237, 25)
        Me.CHARTITLELEG.TabIndex = 5
        '
        'CHARNUMLEG
        '
        Me.CHARNUMLEG.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CHARNUMLEG.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARNUMLEG.Location = New System.Drawing.Point(65, 29)
        Me.CHARNUMLEG.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARNUMLEG.Multiline = True
        Me.CHARNUMLEG.Name = "CHARNUMLEG"
        Me.CHARNUMLEG.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.CHARNUMLEG.Size = New System.Drawing.Size(237, 47)
        Me.CHARNUMLEG.TabIndex = 7
        '
        'lblCHARDENLEG
        '
        Me.lblCHARDENLEG.AutoSize = True
        Me.lblCHARDENLEG.Location = New System.Drawing.Point(0, 84)
        Me.lblCHARDENLEG.Name = "lblCHARDENLEG"
        Me.lblCHARDENLEG.Size = New System.Drawing.Size(61, 17)
        Me.lblCHARDENLEG.TabIndex = 10
        Me.lblCHARDENLEG.Text = "Denom.:"
        '
        'lblCHARTITLELEG
        '
        Me.lblCHARTITLELEG.AutoSize = True
        Me.lblCHARTITLELEG.Location = New System.Drawing.Point(0, 4)
        Me.lblCHARTITLELEG.Name = "lblCHARTITLELEG"
        Me.lblCHARTITLELEG.Size = New System.Drawing.Size(40, 17)
        Me.lblCHARTITLELEG.TabIndex = 6
        Me.lblCHARTITLELEG.Text = "Title:"
        '
        'CHARDENLEG
        '
        Me.CHARDENLEG.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CHARDENLEG.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARDENLEG.Location = New System.Drawing.Point(65, 80)
        Me.CHARDENLEG.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CHARDENLEG.Multiline = True
        Me.CHARDENLEG.Name = "CHARDENLEG"
        Me.CHARDENLEG.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.CHARDENLEG.Size = New System.Drawing.Size(237, 45)
        Me.CHARDENLEG.TabIndex = 9
        '
        'lblCHARNUMLEG
        '
        Me.lblCHARNUMLEG.AutoSize = True
        Me.lblCHARNUMLEG.Location = New System.Drawing.Point(0, 33)
        Me.lblCHARNUMLEG.Name = "lblCHARNUMLEG"
        Me.lblCHARNUMLEG.Size = New System.Drawing.Size(58, 17)
        Me.lblCHARNUMLEG.TabIndex = 8
        Me.lblCHARNUMLEG.Text = "Numer.:"
        '
        'panNomDenomCalcs
        '
        Me.panNomDenomCalcs.Controls.Add(Me.chkRTC_CalStd_Acc)
        Me.panNomDenomCalcs.Controls.Add(Me.panNomDenom)
        Me.panNomDenomCalcs.Controls.Add(Me.gbCalcs)
        Me.panNomDenomCalcs.Location = New System.Drawing.Point(6, 14)
        Me.panNomDenomCalcs.Name = "panNomDenomCalcs"
        Me.panNomDenomCalcs.Size = New System.Drawing.Size(315, 174)
        Me.panNomDenomCalcs.TabIndex = 165
        '
        'chkRTC_CalStd_Acc
        '
        Me.chkRTC_CalStd_Acc.AutoSize = True
        Me.chkRTC_CalStd_Acc.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkRTC_CalStd_Acc.Location = New System.Drawing.Point(7, 30)
        Me.chkRTC_CalStd_Acc.Name = "chkRTC_CalStd_Acc"
        Me.chkRTC_CalStd_Acc.Size = New System.Drawing.Size(118, 21)
        Me.chkRTC_CalStd_Acc.TabIndex = 165
        Me.chkRTC_CalStd_Acc.Text = "No Calculations"
        Me.chkRTC_CalStd_Acc.UseVisualStyleBackColor = True
        '
        'panNomDenom
        '
        Me.panNomDenom.Controls.Add(Me.gbDenom)
        Me.panNomDenom.Controls.Add(Me.gbNumerator)
        Me.panNomDenom.Location = New System.Drawing.Point(4, 47)
        Me.panNomDenom.Name = "panNomDenom"
        Me.panNomDenom.Size = New System.Drawing.Size(137, 124)
        Me.panNomDenom.TabIndex = 165
        '
        'gbDenom
        '
        Me.gbDenom.Controls.Add(Me.rbOld)
        Me.gbDenom.Controls.Add(Me.rbNew)
        Me.gbDenom.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.gbDenom.Location = New System.Drawing.Point(1, 65)
        Me.gbDenom.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbDenom.Name = "gbDenom"
        Me.gbDenom.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbDenom.Size = New System.Drawing.Size(133, 59)
        Me.gbDenom.TabIndex = 166
        Me.gbDenom.TabStop = False
        Me.gbDenom.Text = "Denominator"
        '
        'rbOld
        '
        Me.rbOld.AutoSize = True
        Me.rbOld.BackColor = System.Drawing.Color.Transparent
        Me.rbOld.Checked = True
        Me.rbOld.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbOld.Location = New System.Drawing.Point(24, 16)
        Me.rbOld.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbOld.Name = "rbOld"
        Me.rbOld.Size = New System.Drawing.Size(47, 21)
        Me.rbOld.TabIndex = 3
        Me.rbOld.TabStop = True
        Me.rbOld.Text = "Old"
        Me.rbOld.UseVisualStyleBackColor = False
        '
        'rbNew
        '
        Me.rbNew.AutoSize = True
        Me.rbNew.BackColor = System.Drawing.Color.Transparent
        Me.rbNew.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbNew.Location = New System.Drawing.Point(24, 36)
        Me.rbNew.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbNew.Name = "rbNew"
        Me.rbNew.Size = New System.Drawing.Size(52, 21)
        Me.rbNew.TabIndex = 4
        Me.rbNew.TabStop = True
        Me.rbNew.Text = "New"
        Me.rbNew.UseVisualStyleBackColor = False
        '
        'gbNumerator
        '
        Me.gbNumerator.Controls.Add(Me.rbPosLeg)
        Me.gbNumerator.Controls.Add(Me.rbNegLeg)
        Me.gbNumerator.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.gbNumerator.Location = New System.Drawing.Point(1, 4)
        Me.gbNumerator.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbNumerator.Name = "gbNumerator"
        Me.gbNumerator.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbNumerator.Size = New System.Drawing.Size(133, 59)
        Me.gbNumerator.TabIndex = 165
        Me.gbNumerator.TabStop = False
        Me.gbNumerator.Text = "Numerator"
        '
        'rbPosLeg
        '
        Me.rbPosLeg.AutoSize = True
        Me.rbPosLeg.BackColor = System.Drawing.Color.Transparent
        Me.rbPosLeg.Checked = True
        Me.rbPosLeg.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbPosLeg.Location = New System.Drawing.Point(24, 16)
        Me.rbPosLeg.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbPosLeg.Name = "rbPosLeg"
        Me.rbPosLeg.Size = New System.Drawing.Size(94, 21)
        Me.rbPosLeg.TabIndex = 3
        Me.rbPosLeg.TabStop = True
        Me.rbPosLeg.Text = "(Old - New)"
        Me.rbPosLeg.UseVisualStyleBackColor = False
        '
        'rbNegLeg
        '
        Me.rbNegLeg.AutoSize = True
        Me.rbNegLeg.BackColor = System.Drawing.Color.Transparent
        Me.rbNegLeg.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbNegLeg.Location = New System.Drawing.Point(24, 36)
        Me.rbNegLeg.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbNegLeg.Name = "rbNegLeg"
        Me.rbNegLeg.Size = New System.Drawing.Size(94, 21)
        Me.rbNegLeg.TabIndex = 4
        Me.rbNegLeg.TabStop = True
        Me.rbNegLeg.Text = "(New - Old)"
        Me.rbNegLeg.UseVisualStyleBackColor = False
        '
        'gbCalcs
        '
        Me.gbCalcs.Controls.Add(Me.lblMeanAccuracy)
        Me.gbCalcs.Controls.Add(Me.lblRecovery)
        Me.gbCalcs.Controls.Add(Me.chkCalcIntStdNMF)
        Me.gbCalcs.Controls.Add(Me.lblPercDiff)
        Me.gbCalcs.Controls.Add(Me.rbMeanAcc)
        Me.gbCalcs.Controls.Add(Me.rbRecovery)
        Me.gbCalcs.Controls.Add(Me.rbDifference)
        Me.gbCalcs.Location = New System.Drawing.Point(141, 0)
        Me.gbCalcs.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCalcs.Name = "gbCalcs"
        Me.gbCalcs.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCalcs.Size = New System.Drawing.Size(170, 169)
        Me.gbCalcs.TabIndex = 156
        Me.gbCalcs.TabStop = False
        Me.gbCalcs.Text = "Calculations"
        '
        'lblMeanAccuracy
        '
        Me.lblMeanAccuracy.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMeanAccuracy.Location = New System.Drawing.Point(24, 121)
        Me.lblMeanAccuracy.Name = "lblMeanAccuracy"
        Me.lblMeanAccuracy.Size = New System.Drawing.Size(137, 33)
        Me.lblMeanAccuracy.TabIndex = 9
        Me.lblMeanAccuracy.Text = "(Old - New)/" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "((Old + New)/2) * 100"
        '
        'lblRecovery
        '
        Me.lblRecovery.AutoSize = True
        Me.lblRecovery.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRecovery.Location = New System.Drawing.Point(24, 79)
        Me.lblRecovery.Name = "lblRecovery"
        Me.lblRecovery.Size = New System.Drawing.Size(79, 13)
        Me.lblRecovery.TabIndex = 8
        Me.lblRecovery.Text = "Old/New *100"
        '
        'chkCalcIntStdNMF
        '
        Me.chkCalcIntStdNMF.Checked = True
        Me.chkCalcIntStdNMF.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCalcIntStdNMF.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkCalcIntStdNMF.Location = New System.Drawing.Point(98, 55)
        Me.chkCalcIntStdNMF.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkCalcIntStdNMF.Name = "chkCalcIntStdNMF"
        Me.chkCalcIntStdNMF.Size = New System.Drawing.Size(67, 22)
        Me.chkCalcIntStdNMF.TabIndex = 7
        Me.chkCalcIntStdNMF.Text = "boolOld"
        Me.chkCalcIntStdNMF.UseVisualStyleBackColor = True
        Me.chkCalcIntStdNMF.Visible = False
        '
        'lblPercDiff
        '
        Me.lblPercDiff.AutoSize = True
        Me.lblPercDiff.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPercDiff.Location = New System.Drawing.Point(25, 38)
        Me.lblPercDiff.Name = "lblPercDiff"
        Me.lblPercDiff.Size = New System.Drawing.Size(117, 13)
        Me.lblPercDiff.TabIndex = 7
        Me.lblPercDiff.Text = "(Old - New)/Old * 100"
        '
        'rbMeanAcc
        '
        Me.rbMeanAcc.AutoSize = True
        Me.rbMeanAcc.BackColor = System.Drawing.Color.Transparent
        Me.rbMeanAcc.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbMeanAcc.Location = New System.Drawing.Point(7, 100)
        Me.rbMeanAcc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbMeanAcc.Name = "rbMeanAcc"
        Me.rbMeanAcc.Size = New System.Drawing.Size(114, 21)
        Me.rbMeanAcc.TabIndex = 6
        Me.rbMeanAcc.Text = "Mean Accuracy"
        Me.rbMeanAcc.UseVisualStyleBackColor = False
        '
        'rbRecovery
        '
        Me.rbRecovery.AutoSize = True
        Me.rbRecovery.BackColor = System.Drawing.Color.Transparent
        Me.rbRecovery.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbRecovery.Location = New System.Drawing.Point(7, 58)
        Me.rbRecovery.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbRecovery.Name = "rbRecovery"
        Me.rbRecovery.Size = New System.Drawing.Size(79, 21)
        Me.rbRecovery.TabIndex = 5
        Me.rbRecovery.Text = "Recovery"
        Me.rbRecovery.UseVisualStyleBackColor = False
        '
        'rbDifference
        '
        Me.rbDifference.AutoSize = True
        Me.rbDifference.BackColor = System.Drawing.Color.Transparent
        Me.rbDifference.Checked = True
        Me.rbDifference.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbDifference.Location = New System.Drawing.Point(7, 17)
        Me.rbDifference.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbDifference.Name = "rbDifference"
        Me.rbDifference.Size = New System.Drawing.Size(85, 21)
        Me.rbDifference.TabIndex = 4
        Me.rbDifference.TabStop = True
        Me.rbDifference.Text = "Difference"
        Me.rbDifference.UseVisualStyleBackColor = False
        '
        'tabRTC
        '
        Me.tabRTC.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabRTC.Controls.Add(Me.tpFormat)
        Me.tabRTC.Controls.Add(Me.tpPeriodTemp)
        Me.tabRTC.Controls.Add(Me.tpAutoAssignment)
        Me.tabRTC.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabRTC.Location = New System.Drawing.Point(193, 110)
        Me.tabRTC.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tabRTC.Name = "tabRTC"
        Me.tabRTC.SelectedIndex = 0
        Me.tabRTC.Size = New System.Drawing.Size(1356, 661)
        Me.tabRTC.TabIndex = 156
        '
        'tpFormat
        '
        Me.tpFormat.Controls.Add(Me.panFormat)
        Me.tpFormat.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tpFormat.Location = New System.Drawing.Point(4, 26)
        Me.tpFormat.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tpFormat.Name = "tpFormat"
        Me.tpFormat.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tpFormat.Size = New System.Drawing.Size(1348, 631)
        Me.tpFormat.TabIndex = 0
        Me.tpFormat.Text = "Format"
        Me.tpFormat.UseVisualStyleBackColor = True
        '
        'panFormat
        '
        Me.panFormat.AutoScroll = True
        Me.panFormat.Controls.Add(Me.gbMatrixFactor)
        Me.panFormat.Controls.Add(Me.gbGroupSort)
        Me.panFormat.Controls.Add(Me.gbSAT)
        Me.panFormat.Controls.Add(Me.gbQCGroup)
        Me.panFormat.Controls.Add(Me.gbRegrULOQ)
        Me.panFormat.Controls.Add(Me.gbCriteria)
        Me.panFormat.Controls.Add(Me.gbCarryover)
        Me.panFormat.Controls.Add(Me.gbLegendFormat)
        Me.panFormat.Controls.Add(Me.gbIncSampleCriteria)
        Me.panFormat.Controls.Add(Me.cmdIncSamples)
        Me.panFormat.Controls.Add(Me.rbRTC_CalStd_All)
        Me.panFormat.Controls.Add(Me.rbRTC_CalStd_Acc)
        Me.panFormat.Controls.Add(Me.gbResultsChoice)
        Me.panFormat.Controls.Add(Me.gbStats)
        Me.panFormat.Controls.Add(Me.gbTableLegend)
        Me.panFormat.Controls.Add(Me.gbSampleGroup)
        Me.panFormat.Controls.Add(Me.gbRTC_CalStd)
        Me.panFormat.Controls.Add(Me.gbSampleSort)
        Me.panFormat.Controls.Add(Me.gbPSAE)
        Me.panFormat.Controls.Add(Me.gbAnovaStats)
        Me.panFormat.Controls.Add(Me.gbRTC_Samples)
        Me.panFormat.Controls.Add(Me.gbRTC_QC)
        Me.panFormat.Dock = System.Windows.Forms.DockStyle.Fill
        Me.panFormat.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.panFormat.Location = New System.Drawing.Point(3, 4)
        Me.panFormat.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panFormat.Name = "panFormat"
        Me.panFormat.Size = New System.Drawing.Size(1342, 623)
        Me.panFormat.TabIndex = 157
        '
        'gbMatrixFactor
        '
        Me.gbMatrixFactor.Controls.Add(Me.panMF1)
        Me.gbMatrixFactor.Controls.Add(Me.chkMFTable)
        Me.gbMatrixFactor.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbMatrixFactor.Location = New System.Drawing.Point(737, 374)
        Me.gbMatrixFactor.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbMatrixFactor.Name = "gbMatrixFactor"
        Me.gbMatrixFactor.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbMatrixFactor.Size = New System.Drawing.Size(352, 225)
        Me.gbMatrixFactor.TabIndex = 158
        Me.gbMatrixFactor.TabStop = False
        Me.gbMatrixFactor.Text = "Matrix Factor"
        '
        'panMF1
        '
        Me.panMF1.Controls.Add(Me.lblMF01)
        Me.panMF1.Controls.Add(Me.chkInclIntStdNMF)
        Me.panMF1.Controls.Add(Me.chkInclMFCols)
        Me.panMF1.Location = New System.Drawing.Point(37, 46)
        Me.panMF1.Name = "panMF1"
        Me.panMF1.Size = New System.Drawing.Size(309, 170)
        Me.panMF1.TabIndex = 160
        '
        'lblMF01
        '
        Me.lblMF01.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblMF01.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblMF01.Location = New System.Drawing.Point(22, 67)
        Me.lblMF01.Name = "lblMF01"
        Me.lblMF01.Size = New System.Drawing.Size(284, 103)
        Me.lblMF01.TabIndex = 9
        Me.lblMF01.Text = "MF Calc"
        '
        'chkInclIntStdNMF
        '
        Me.chkInclIntStdNMF.Checked = True
        Me.chkInclIntStdNMF.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkInclIntStdNMF.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkInclIntStdNMF.Location = New System.Drawing.Point(0, 40)
        Me.chkInclIntStdNMF.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkInclIntStdNMF.Name = "chkInclIntStdNMF"
        Me.chkInclIntStdNMF.Size = New System.Drawing.Size(306, 29)
        Me.chkInclIntStdNMF.TabIndex = 8
        Me.chkInclIntStdNMF.Text = "Include IntStd-Normalized Matrix Factor column"
        Me.chkInclIntStdNMF.UseVisualStyleBackColor = True
        '
        'chkInclMFCols
        '
        Me.chkInclMFCols.Checked = True
        Me.chkInclMFCols.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkInclMFCols.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkInclMFCols.Location = New System.Drawing.Point(0, 0)
        Me.chkInclMFCols.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkInclMFCols.Name = "chkInclMFCols"
        Me.chkInclMFCols.Size = New System.Drawing.Size(306, 39)
        Me.chkInclMFCols.TabIndex = 6
        Me.chkInclMFCols.Text = "Include Matrix Factor columns for Analyte and Int Std"
        Me.chkInclMFCols.UseVisualStyleBackColor = True
        '
        'chkMFTable
        '
        Me.chkMFTable.AutoSize = True
        Me.chkMFTable.Checked = True
        Me.chkMFTable.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkMFTable.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkMFTable.Location = New System.Drawing.Point(15, 21)
        Me.chkMFTable.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkMFTable.Name = "chkMFTable"
        Me.chkMFTable.Size = New System.Drawing.Size(189, 21)
        Me.chkMFTable.TabIndex = 9
        Me.chkMFTable.Text = "Show table as Matrix Factor"
        Me.chkMFTable.UseVisualStyleBackColor = True
        '
        'gbGroupSort
        '
        Me.gbGroupSort.Controls.Add(Me.lblGroupSort)
        Me.gbGroupSort.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbGroupSort.Location = New System.Drawing.Point(1118, 407)
        Me.gbGroupSort.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbGroupSort.Name = "gbGroupSort"
        Me.gbGroupSort.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbGroupSort.Size = New System.Drawing.Size(324, 100)
        Me.gbGroupSort.TabIndex = 164
        Me.gbGroupSort.TabStop = False
        Me.gbGroupSort.Text = "Grouping/Sorting"
        Me.gbGroupSort.Visible = False
        '
        'lblGroupSort
        '
        Me.lblGroupSort.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblGroupSort.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGroupSort.Location = New System.Drawing.Point(3, 22)
        Me.lblGroupSort.Name = "lblGroupSort"
        Me.lblGroupSort.Size = New System.Drawing.Size(318, 74)
        Me.lblGroupSort.TabIndex = 0
        Me.lblGroupSort.Text = "Grouping and Sorting defined in Sample Concentrations window"
        '
        'gbSAT
        '
        Me.gbSAT.Controls.Add(Me.chkBOOLCONCCOMMENTS)
        Me.gbSAT.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSAT.Location = New System.Drawing.Point(1027, 216)
        Me.gbSAT.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSAT.Name = "gbSAT"
        Me.gbSAT.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbSAT.Size = New System.Drawing.Size(327, 65)
        Me.gbSAT.TabIndex = 163
        Me.gbSAT.TabStop = False
        Me.gbSAT.Text = "Additional Sample Analysis Parameters"
        Me.gbSAT.Visible = False
        '
        'chkBOOLCONCCOMMENTS
        '
        Me.chkBOOLCONCCOMMENTS.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkBOOLCONCCOMMENTS.Location = New System.Drawing.Point(12, 19)
        Me.chkBOOLCONCCOMMENTS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkBOOLCONCCOMMENTS.Name = "chkBOOLCONCCOMMENTS"
        Me.chkBOOLCONCCOMMENTS.Size = New System.Drawing.Size(309, 41)
        Me.chkBOOLCONCCOMMENTS.TabIndex = 4
        Me.chkBOOLCONCCOMMENTS.Text = "Add a superscript letter to concentration value if sample contains a Watson comme" & _
    "nt"
        Me.chkBOOLCONCCOMMENTS.UseVisualStyleBackColor = True
        '
        'gbQCGroup
        '
        Me.gbQCGroup.Controls.Add(Me.lblQCGroup)
        Me.gbQCGroup.Controls.Add(Me.rbINTQCLEVELGROUPQCLabel)
        Me.gbQCGroup.Controls.Add(Me.rbINTQCLEVELGROUPNomConc)
        Me.gbQCGroup.Controls.Add(Me.rbINTQCLEVELGROUPLevel)
        Me.gbQCGroup.Enabled = False
        Me.gbQCGroup.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbQCGroup.Location = New System.Drawing.Point(678, 297)
        Me.gbQCGroup.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbQCGroup.Name = "gbQCGroup"
        Me.gbQCGroup.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbQCGroup.Size = New System.Drawing.Size(327, 108)
        Me.gbQCGroup.TabIndex = 162
        Me.gbQCGroup.TabStop = False
        Me.gbQCGroup.Text = "Group QC Levels By (ignored if Assigned Samples):"
        Me.gbQCGroup.Visible = False
        '
        'lblQCGroup
        '
        Me.lblQCGroup.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQCGroup.Location = New System.Drawing.Point(194, 26)
        Me.lblQCGroup.Name = "lblQCGroup"
        Me.lblQCGroup.Size = New System.Drawing.Size(92, 73)
        Me.lblQCGroup.TabIndex = 4
        Me.lblQCGroup.Text = "Database:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "0" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "1" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "2"
        Me.lblQCGroup.Visible = False
        '
        'rbINTQCLEVELGROUPQCLabel
        '
        Me.rbINTQCLEVELGROUPQCLabel.AutoSize = True
        Me.rbINTQCLEVELGROUPQCLabel.BackColor = System.Drawing.Color.Transparent
        Me.rbINTQCLEVELGROUPQCLabel.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbINTQCLEVELGROUPQCLabel.ForeColor = System.Drawing.Color.Black
        Me.rbINTQCLEVELGROUPQCLabel.Location = New System.Drawing.Point(12, 79)
        Me.rbINTQCLEVELGROUPQCLabel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbINTQCLEVELGROUPQCLabel.Name = "rbINTQCLEVELGROUPQCLabel"
        Me.rbINTQCLEVELGROUPQCLabel.Size = New System.Drawing.Size(79, 21)
        Me.rbINTQCLEVELGROUPQCLabel.TabIndex = 3
        Me.rbINTQCLEVELGROUPQCLabel.Text = "QC Label"
        Me.rbINTQCLEVELGROUPQCLabel.UseVisualStyleBackColor = False
        '
        'rbINTQCLEVELGROUPNomConc
        '
        Me.rbINTQCLEVELGROUPNomConc.AutoSize = True
        Me.rbINTQCLEVELGROUPNomConc.BackColor = System.Drawing.Color.Transparent
        Me.rbINTQCLEVELGROUPNomConc.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbINTQCLEVELGROUPNomConc.ForeColor = System.Drawing.Color.Black
        Me.rbINTQCLEVELGROUPNomConc.Location = New System.Drawing.Point(12, 58)
        Me.rbINTQCLEVELGROUPNomConc.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbINTQCLEVELGROUPNomConc.Name = "rbINTQCLEVELGROUPNomConc"
        Me.rbINTQCLEVELGROUPNomConc.Size = New System.Drawing.Size(160, 21)
        Me.rbINTQCLEVELGROUPNomConc.TabIndex = 2
        Me.rbINTQCLEVELGROUPNomConc.Text = "Nominal Concentration"
        Me.rbINTQCLEVELGROUPNomConc.UseVisualStyleBackColor = False
        '
        'rbINTQCLEVELGROUPLevel
        '
        Me.rbINTQCLEVELGROUPLevel.AutoSize = True
        Me.rbINTQCLEVELGROUPLevel.BackColor = System.Drawing.Color.Transparent
        Me.rbINTQCLEVELGROUPLevel.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.rbINTQCLEVELGROUPLevel.ForeColor = System.Drawing.Color.Black
        Me.rbINTQCLEVELGROUPLevel.Location = New System.Drawing.Point(12, 37)
        Me.rbINTQCLEVELGROUPLevel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbINTQCLEVELGROUPLevel.Name = "rbINTQCLEVELGROUPLevel"
        Me.rbINTQCLEVELGROUPLevel.Size = New System.Drawing.Size(145, 21)
        Me.rbINTQCLEVELGROUPLevel.TabIndex = 1
        Me.rbINTQCLEVELGROUPLevel.Text = "Assay Level (Default)"
        Me.rbINTQCLEVELGROUPLevel.UseVisualStyleBackColor = False
        '
        'gbRegrULOQ
        '
        Me.gbRegrULOQ.Controls.Add(Me.chkBOOLREGRULOQ)
        Me.gbRegrULOQ.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbRegrULOQ.Location = New System.Drawing.Point(692, 151)
        Me.gbRegrULOQ.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbRegrULOQ.Name = "gbRegrULOQ"
        Me.gbRegrULOQ.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbRegrULOQ.Size = New System.Drawing.Size(327, 65)
        Me.gbRegrULOQ.TabIndex = 161
        Me.gbRegrULOQ.TabStop = False
        Me.gbRegrULOQ.Text = "Additional Columns"
        Me.gbRegrULOQ.Visible = False
        '
        'chkBOOLREGRULOQ
        '
        Me.chkBOOLREGRULOQ.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkBOOLREGRULOQ.Location = New System.Drawing.Point(12, 19)
        Me.chkBOOLREGRULOQ.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkBOOLREGRULOQ.Name = "chkBOOLREGRULOQ"
        Me.chkBOOLREGRULOQ.Size = New System.Drawing.Size(309, 41)
        Me.chkBOOLREGRULOQ.TabIndex = 4
        Me.chkBOOLREGRULOQ.Text = "Include LLOQ and ULOQ columns"
        Me.chkBOOLREGRULOQ.UseVisualStyleBackColor = True
        '
        'gbCriteria
        '
        Me.gbCriteria.Controls.Add(Me.NUMPRECCRITLOTS)
        Me.gbCriteria.Controls.Add(Me.lblPrecisionCrit)
        Me.gbCriteria.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbCriteria.Location = New System.Drawing.Point(692, 10)
        Me.gbCriteria.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCriteria.Name = "gbCriteria"
        Me.gbCriteria.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCriteria.Size = New System.Drawing.Size(327, 65)
        Me.gbCriteria.TabIndex = 160
        Me.gbCriteria.TabStop = False
        Me.gbCriteria.Text = "Additional Acceptance Criteria"
        Me.gbCriteria.Visible = False
        '
        'NUMPRECCRITLOTS
        '
        Me.NUMPRECCRITLOTS.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NUMPRECCRITLOTS.Location = New System.Drawing.Point(155, 25)
        Me.NUMPRECCRITLOTS.Name = "NUMPRECCRITLOTS"
        Me.NUMPRECCRITLOTS.Size = New System.Drawing.Size(99, 25)
        Me.NUMPRECCRITLOTS.TabIndex = 1
        Me.NUMPRECCRITLOTS.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblPrecisionCrit
        '
        Me.lblPrecisionCrit.AutoSize = True
        Me.lblPrecisionCrit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrecisionCrit.Location = New System.Drawing.Point(27, 22)
        Me.lblPrecisionCrit.Name = "lblPrecisionCrit"
        Me.lblPrecisionCrit.Size = New System.Drawing.Size(115, 34)
        Me.lblPrecisionCrit.TabIndex = 0
        Me.lblPrecisionCrit.Text = "Precision Criteria:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Enter 0 to ignore)"
        '
        'gbCarryover
        '
        Me.gbCarryover.Controls.Add(Me.lblCarryover)
        Me.gbCarryover.Controls.Add(Me.CHARCARRYOVERLABEL)
        Me.gbCarryover.Enabled = False
        Me.gbCarryover.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbCarryover.Location = New System.Drawing.Point(366, 883)
        Me.gbCarryover.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCarryover.Name = "gbCarryover"
        Me.gbCarryover.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbCarryover.Size = New System.Drawing.Size(324, 88)
        Me.gbCarryover.TabIndex = 159
        Me.gbCarryover.TabStop = False
        Me.gbCarryover.Text = "Carryover"
        '
        'lblCarryover
        '
        Me.lblCarryover.AutoSize = True
        Me.lblCarryover.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCarryover.Location = New System.Drawing.Point(15, 22)
        Me.lblCarryover.Name = "lblCarryover"
        Me.lblCarryover.Size = New System.Drawing.Size(224, 17)
        Me.lblCarryover.TabIndex = 1
        Me.lblCarryover.Text = "Label Carryover Injection columns as:"
        '
        'CHARCARRYOVERLABEL
        '
        Me.CHARCARRYOVERLABEL.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.CHARCARRYOVERLABEL.Location = New System.Drawing.Point(18, 49)
        Me.CHARCARRYOVERLABEL.Name = "CHARCARRYOVERLABEL"
        Me.CHARCARRYOVERLABEL.Size = New System.Drawing.Size(212, 25)
        Me.CHARCARRYOVERLABEL.TabIndex = 0
        '
        'gbLegendFormat
        '
        Me.gbLegendFormat.Controls.Add(Me.chkBOOLREASSAYREASLETTERS)
        Me.gbLegendFormat.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbLegendFormat.Location = New System.Drawing.Point(692, 80)
        Me.gbLegendFormat.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbLegendFormat.Name = "gbLegendFormat"
        Me.gbLegendFormat.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbLegendFormat.Size = New System.Drawing.Size(327, 65)
        Me.gbLegendFormat.TabIndex = 157
        Me.gbLegendFormat.TabStop = False
        Me.gbLegendFormat.Text = "Legend Format"
        Me.gbLegendFormat.Visible = False
        '
        'chkBOOLREASSAYREASLETTERS
        '
        Me.chkBOOLREASSAYREASLETTERS.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.chkBOOLREASSAYREASLETTERS.Location = New System.Drawing.Point(12, 19)
        Me.chkBOOLREASSAYREASLETTERS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.chkBOOLREASSAYREASLETTERS.Name = "chkBOOLREASSAYREASLETTERS"
        Me.chkBOOLREASSAYREASLETTERS.Size = New System.Drawing.Size(309, 41)
        Me.chkBOOLREASSAYREASLETTERS.TabIndex = 4
        Me.chkBOOLREASSAYREASLETTERS.Text = "Use letters for [Reason for Reported Conc.] column and legend"
        Me.chkBOOLREASSAYREASLETTERS.UseVisualStyleBackColor = True
        '
        'gbIncSampleCriteria
        '
        Me.gbIncSampleCriteria.Controls.Add(Me.dgvAnalytes)
        Me.gbIncSampleCriteria.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbIncSampleCriteria.Location = New System.Drawing.Point(884, 235)
        Me.gbIncSampleCriteria.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbIncSampleCriteria.Name = "gbIncSampleCriteria"
        Me.gbIncSampleCriteria.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbIncSampleCriteria.Size = New System.Drawing.Size(324, 160)
        Me.gbIncSampleCriteria.TabIndex = 156
        Me.gbIncSampleCriteria.TabStop = False
        Me.gbIncSampleCriteria.Text = "Acceptance Criteria"
        '
        'dgvAnalytes
        '
        Me.dgvAnalytes.AllowUserToAddRows = False
        Me.dgvAnalytes.AllowUserToDeleteRows = False
        Me.dgvAnalytes.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAnalytes.BackgroundColor = System.Drawing.Color.White
        Me.dgvAnalytes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAnalytes.Location = New System.Drawing.Point(7, 25)
        Me.dgvAnalytes.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvAnalytes.Name = "dgvAnalytes"
        Me.dgvAnalytes.ReadOnly = True
        Me.dgvAnalytes.Size = New System.Drawing.Size(310, 127)
        Me.dgvAnalytes.TabIndex = 0
        '
        'tpPeriodTemp
        '
        Me.tpPeriodTemp.Controls.Add(Me.panFDARef)
        Me.tpPeriodTemp.Controls.Add(Me.gbStabilityType)
        Me.tpPeriodTemp.Controls.Add(Me.gbAdditional)
        Me.tpPeriodTemp.Location = New System.Drawing.Point(4, 26)
        Me.tpPeriodTemp.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tpPeriodTemp.Name = "tpPeriodTemp"
        Me.tpPeriodTemp.Size = New System.Drawing.Size(1348, 631)
        Me.tpPeriodTemp.TabIndex = 1
        Me.tpPeriodTemp.Text = "Stability Conditions"
        Me.tpPeriodTemp.UseVisualStyleBackColor = True
        '
        'panFDARef
        '
        Me.panFDARef.Controls.Add(Me.txtRef)
        Me.panFDARef.Controls.Add(Me.lblFDA)
        Me.panFDARef.Location = New System.Drawing.Point(3, 352)
        Me.panFDARef.Name = "panFDARef"
        Me.panFDARef.Size = New System.Drawing.Size(695, 184)
        Me.panFDARef.TabIndex = 167
        '
        'txtRef
        '
        Me.txtRef.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtRef.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.txtRef.Location = New System.Drawing.Point(9, 30)
        Me.txtRef.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtRef.Multiline = True
        Me.txtRef.Name = "txtRef"
        Me.txtRef.Size = New System.Drawing.Size(680, 150)
        Me.txtRef.TabIndex = 11
        '
        'lblFDA
        '
        Me.lblFDA.AutoSize = True
        Me.lblFDA.Font = New System.Drawing.Font("Segoe UI", 9.75!)
        Me.lblFDA.Location = New System.Drawing.Point(6, 9)
        Me.lblFDA.Name = "lblFDA"
        Me.lblFDA.Size = New System.Drawing.Size(399, 17)
        Me.lblFDA.TabIndex = 10
        Me.lblFDA.Text = "Bioanalytical Method Validation - Guidance for Industry - May 2018" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'gbStabilityType
        '
        Me.gbStabilityType.Controls.Add(Me.rbDilution)
        Me.gbStabilityType.Controls.Add(Me.rbAutosampler)
        Me.gbStabilityType.Controls.Add(Me.rbBatchReinjection)
        Me.gbStabilityType.Controls.Add(Me.rbSpiking)
        Me.gbStabilityType.Controls.Add(Me.rbReinjection)
        Me.gbStabilityType.Controls.Add(Me.rbStockSolution)
        Me.gbStabilityType.Controls.Add(Me.rbBlood)
        Me.gbStabilityType.Controls.Add(Me.rbLT)
        Me.gbStabilityType.Controls.Add(Me.rbFT)
        Me.gbStabilityType.Controls.Add(Me.rbBenchTop)
        Me.gbStabilityType.Controls.Add(Me.rbProcess)
        Me.gbStabilityType.Controls.Add(Me.rbNA)
        Me.gbStabilityType.Controls.Add(Me.txtStabilityNotes)
        Me.gbStabilityType.Controls.Add(Me.lblStabilityNotes)
        Me.gbStabilityType.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbStabilityType.Location = New System.Drawing.Point(460, 6)
        Me.gbStabilityType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStabilityType.Name = "gbStabilityType"
        Me.gbStabilityType.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.gbStabilityType.Size = New System.Drawing.Size(238, 339)
        Me.gbStabilityType.TabIndex = 166
        Me.gbStabilityType.TabStop = False
        Me.gbStabilityType.Text = "Stability Type"
        '
        'rbDilution
        '
        Me.rbDilution.AutoSize = True
        Me.rbDilution.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbDilution.ForeColor = System.Drawing.Color.Black
        Me.rbDilution.Location = New System.Drawing.Point(29, 43)
        Me.rbDilution.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbDilution.Name = "rbDilution"
        Me.rbDilution.Size = New System.Drawing.Size(70, 21)
        Me.rbDilution.TabIndex = 1
        Me.rbDilution.Text = "Dilution"
        Me.rbDilution.UseVisualStyleBackColor = True
        '
        'rbAutosampler
        '
        Me.rbAutosampler.AutoSize = True
        Me.rbAutosampler.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbAutosampler.ForeColor = System.Drawing.Color.Black
        Me.rbAutosampler.Location = New System.Drawing.Point(29, 66)
        Me.rbAutosampler.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbAutosampler.Name = "rbAutosampler"
        Me.rbAutosampler.Size = New System.Drawing.Size(100, 21)
        Me.rbAutosampler.TabIndex = 2
        Me.rbAutosampler.Text = "Autosampler"
        Me.rbAutosampler.UseVisualStyleBackColor = True
        '
        'rbBatchReinjection
        '
        Me.rbBatchReinjection.AutoSize = True
        Me.rbBatchReinjection.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbBatchReinjection.ForeColor = System.Drawing.Color.Black
        Me.rbBatchReinjection.Location = New System.Drawing.Point(29, 204)
        Me.rbBatchReinjection.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbBatchReinjection.Name = "rbBatchReinjection"
        Me.rbBatchReinjection.Size = New System.Drawing.Size(124, 21)
        Me.rbBatchReinjection.TabIndex = 8
        Me.rbBatchReinjection.Text = "Batch Reinjection"
        Me.rbBatchReinjection.UseVisualStyleBackColor = True
        '
        'rbSpiking
        '
        Me.rbSpiking.AutoSize = True
        Me.rbSpiking.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbSpiking.ForeColor = System.Drawing.Color.Black
        Me.rbSpiking.Location = New System.Drawing.Point(29, 273)
        Me.rbSpiking.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbSpiking.Name = "rbSpiking"
        Me.rbSpiking.Size = New System.Drawing.Size(118, 21)
        Me.rbSpiking.TabIndex = 11
        Me.rbSpiking.Text = "Spiking solution"
        Me.rbSpiking.UseVisualStyleBackColor = True
        '
        'rbReinjection
        '
        Me.rbReinjection.AutoSize = True
        Me.rbReinjection.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbReinjection.ForeColor = System.Drawing.Color.Black
        Me.rbReinjection.Location = New System.Drawing.Point(29, 181)
        Me.rbReinjection.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbReinjection.Name = "rbReinjection"
        Me.rbReinjection.Size = New System.Drawing.Size(89, 21)
        Me.rbReinjection.TabIndex = 7
        Me.rbReinjection.Text = "Reinjection"
        Me.rbReinjection.UseVisualStyleBackColor = True
        '
        'rbStockSolution
        '
        Me.rbStockSolution.AutoSize = True
        Me.rbStockSolution.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbStockSolution.ForeColor = System.Drawing.Color.Black
        Me.rbStockSolution.Location = New System.Drawing.Point(29, 250)
        Me.rbStockSolution.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbStockSolution.Name = "rbStockSolution"
        Me.rbStockSolution.Size = New System.Drawing.Size(107, 21)
        Me.rbStockSolution.TabIndex = 10
        Me.rbStockSolution.Text = "Stock solution"
        Me.rbStockSolution.UseVisualStyleBackColor = True
        '
        'rbBlood
        '
        Me.rbBlood.AutoSize = True
        Me.rbBlood.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbBlood.ForeColor = System.Drawing.Color.Black
        Me.rbBlood.Location = New System.Drawing.Point(29, 227)
        Me.rbBlood.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbBlood.Name = "rbBlood"
        Me.rbBlood.Size = New System.Drawing.Size(60, 21)
        Me.rbBlood.TabIndex = 9
        Me.rbBlood.Text = "Blood"
        Me.rbBlood.UseVisualStyleBackColor = True
        '
        'rbLT
        '
        Me.rbLT.AutoSize = True
        Me.rbLT.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbLT.ForeColor = System.Drawing.Color.Black
        Me.rbLT.Location = New System.Drawing.Point(29, 158)
        Me.rbLT.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbLT.Name = "rbLT"
        Me.rbLT.Size = New System.Drawing.Size(87, 21)
        Me.rbLT.TabIndex = 6
        Me.rbLT.Text = "Long-term"
        Me.rbLT.UseVisualStyleBackColor = True
        '
        'rbFT
        '
        Me.rbFT.AutoSize = True
        Me.rbFT.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbFT.ForeColor = System.Drawing.Color.Black
        Me.rbFT.Location = New System.Drawing.Point(29, 135)
        Me.rbFT.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbFT.Name = "rbFT"
        Me.rbFT.Size = New System.Drawing.Size(96, 21)
        Me.rbFT.TabIndex = 5
        Me.rbFT.Text = "Freeze-thaw"
        Me.rbFT.UseVisualStyleBackColor = True
        '
        'rbBenchTop
        '
        Me.rbBenchTop.AutoSize = True
        Me.rbBenchTop.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbBenchTop.ForeColor = System.Drawing.Color.Black
        Me.rbBenchTop.Location = New System.Drawing.Point(29, 89)
        Me.rbBenchTop.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbBenchTop.Name = "rbBenchTop"
        Me.rbBenchTop.Size = New System.Drawing.Size(85, 21)
        Me.rbBenchTop.TabIndex = 3
        Me.rbBenchTop.Text = "Bench-top"
        Me.rbBenchTop.UseVisualStyleBackColor = True
        '
        'rbProcess
        '
        Me.rbProcess.AutoSize = True
        Me.rbProcess.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbProcess.ForeColor = System.Drawing.Color.Black
        Me.rbProcess.Location = New System.Drawing.Point(29, 112)
        Me.rbProcess.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbProcess.Name = "rbProcess"
        Me.rbProcess.Size = New System.Drawing.Size(137, 21)
        Me.rbProcess.TabIndex = 4
        Me.rbProcess.Text = "Extract (Processed)"
        Me.rbProcess.UseVisualStyleBackColor = True
        '
        'rbNA
        '
        Me.rbNA.AutoSize = True
        Me.rbNA.Checked = True
        Me.rbNA.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbNA.ForeColor = System.Drawing.Color.Black
        Me.rbNA.Location = New System.Drawing.Point(29, 20)
        Me.rbNA.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.rbNA.Name = "rbNA"
        Me.rbNA.Size = New System.Drawing.Size(113, 21)
        Me.rbNA.TabIndex = 0
        Me.rbNA.TabStop = True
        Me.rbNA.Text = "Not Applicable"
        Me.rbNA.UseVisualStyleBackColor = True
        '
        'txtStabilityNotes
        '
        Me.txtStabilityNotes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStabilityNotes.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStabilityNotes.Location = New System.Drawing.Point(73, 307)
        Me.txtStabilityNotes.Name = "txtStabilityNotes"
        Me.txtStabilityNotes.Size = New System.Drawing.Size(159, 25)
        Me.txtStabilityNotes.TabIndex = 12
        '
        'lblStabilityNotes
        '
        Me.lblStabilityNotes.AutoSize = True
        Me.lblStabilityNotes.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStabilityNotes.Location = New System.Drawing.Point(26, 310)
        Me.lblStabilityNotes.Name = "lblStabilityNotes"
        Me.lblStabilityNotes.Size = New System.Drawing.Size(46, 17)
        Me.lblStabilityNotes.TabIndex = 10
        Me.lblStabilityNotes.Text = "Notes:"
        '
        'tpAutoAssignment
        '
        Me.tpAutoAssignment.Controls.Add(Me.lblRunIdentifier)
        Me.tpAutoAssignment.Controls.Add(Me.lblCase)
        Me.tpAutoAssignment.Controls.Add(Me.lblLogic)
        Me.tpAutoAssignment.Controls.Add(Me.dgvSAS)
        Me.tpAutoAssignment.Controls.Add(Me.cmdAnalRuns)
        Me.tpAutoAssignment.Controls.Add(Me.dgvASP)
        Me.tpAutoAssignment.Controls.Add(Me.panTableGraphicExamples)
        Me.tpAutoAssignment.Controls.Add(Me.lblTableGraphicExamplesLabel)
        Me.tpAutoAssignment.Controls.Add(Me.cmdShowRunSummary)
        Me.tpAutoAssignment.Controls.Add(Me.lblOptional)
        Me.tpAutoAssignment.Controls.Add(Me.lblWatsonE)
        Me.tpAutoAssignment.Controls.Add(Me.lblAccepted)
        Me.tpAutoAssignment.Location = New System.Drawing.Point(4, 26)
        Me.tpAutoAssignment.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tpAutoAssignment.Name = "tpAutoAssignment"
        Me.tpAutoAssignment.Size = New System.Drawing.Size(1348, 631)
        Me.tpAutoAssignment.TabIndex = 2
        Me.tpAutoAssignment.Text = "Sample Auto-Assignment"
        Me.tpAutoAssignment.UseVisualStyleBackColor = True
        '
        'lblRunIdentifier
        '
        Me.lblRunIdentifier.AutoSize = True
        Me.lblRunIdentifier.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblRunIdentifier.Location = New System.Drawing.Point(11, 103)
        Me.lblRunIdentifier.Name = "lblRunIdentifier"
        Me.lblRunIdentifier.Size = New System.Drawing.Size(55, 17)
        Me.lblRunIdentifier.TabIndex = 167
        Me.lblRunIdentifier.Text = "Label16"
        '
        'lblCase
        '
        Me.lblCase.AutoSize = True
        Me.lblCase.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblCase.Location = New System.Drawing.Point(11, 86)
        Me.lblCase.Name = "lblCase"
        Me.lblCase.Size = New System.Drawing.Size(55, 17)
        Me.lblCase.TabIndex = 166
        Me.lblCase.Text = "Label16"
        '
        'lblLogic
        '
        Me.lblLogic.AutoSize = True
        Me.lblLogic.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblLogic.Location = New System.Drawing.Point(11, 69)
        Me.lblLogic.Name = "lblLogic"
        Me.lblLogic.Size = New System.Drawing.Size(55, 17)
        Me.lblLogic.TabIndex = 165
        Me.lblLogic.Text = "Label16"
        '
        'dgvSAS
        '
        Me.dgvSAS.AllowUserToAddRows = False
        Me.dgvSAS.AllowUserToDeleteRows = False
        DataGridViewCellStyle3.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.dgvSAS.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvSAS.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSAS.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvSAS.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSAS.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvSAS.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.Padding = New System.Windows.Forms.Padding(0, 6, 0, 6)
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSAS.DefaultCellStyle = DataGridViewCellStyle5
        Me.dgvSAS.Location = New System.Drawing.Point(16, 124)
        Me.dgvSAS.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvSAS.Name = "dgvSAS"
        Me.dgvSAS.Size = New System.Drawing.Size(613, 449)
        Me.dgvSAS.TabIndex = 2
        '
        'cmdAnalRuns
        '
        Me.cmdAnalRuns.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdAnalRuns.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAnalRuns.FlatAppearance.BorderSize = 0
        Me.cmdAnalRuns.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdAnalRuns.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdAnalRuns.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdAnalRuns.Location = New System.Drawing.Point(1006, 4)
        Me.cmdAnalRuns.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdAnalRuns.Name = "cmdAnalRuns"
        Me.cmdAnalRuns.Size = New System.Drawing.Size(233, 31)
        Me.cmdAnalRuns.TabIndex = 164
        Me.cmdAnalRuns.Text = "&View Analytical Run Samples..."
        Me.cmdAnalRuns.UseVisualStyleBackColor = True
        '
        'dgvASP
        '
        Me.dgvASP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvASP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvASP.Location = New System.Drawing.Point(768, -1)
        Me.dgvASP.Name = "dgvASP"
        Me.dgvASP.Size = New System.Drawing.Size(201, 36)
        Me.dgvASP.TabIndex = 3
        Me.dgvASP.Visible = False
        '
        'panTableGraphicExamples
        '
        Me.panTableGraphicExamples.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panTableGraphicExamples.AutoScroll = True
        Me.panTableGraphicExamples.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.panTableGraphicExamples.Controls.Add(Me.pbxTableGraphicExamples)
        Me.panTableGraphicExamples.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.panTableGraphicExamples.ForeColor = System.Drawing.Color.Transparent
        Me.panTableGraphicExamples.Location = New System.Drawing.Point(635, 124)
        Me.panTableGraphicExamples.Name = "panTableGraphicExamples"
        Me.panTableGraphicExamples.Size = New System.Drawing.Size(604, 449)
        Me.panTableGraphicExamples.TabIndex = 163
        '
        'pbxTableGraphicExamples
        '
        Me.pbxTableGraphicExamples.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbxTableGraphicExamples.BackColor = System.Drawing.Color.White
        Me.pbxTableGraphicExamples.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pbxTableGraphicExamples.Location = New System.Drawing.Point(5, 4)
        Me.pbxTableGraphicExamples.Name = "pbxTableGraphicExamples"
        Me.pbxTableGraphicExamples.Size = New System.Drawing.Size(594, 440)
        Me.pbxTableGraphicExamples.TabIndex = 150
        Me.pbxTableGraphicExamples.TabStop = False
        '
        'lblTableGraphicExamplesLabel
        '
        Me.lblTableGraphicExamplesLabel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblTableGraphicExamplesLabel.AutoSize = True
        Me.lblTableGraphicExamplesLabel.BackColor = System.Drawing.Color.White
        Me.lblTableGraphicExamplesLabel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTableGraphicExamplesLabel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblTableGraphicExamplesLabel.Location = New System.Drawing.Point(632, 103)
        Me.lblTableGraphicExamplesLabel.Name = "lblTableGraphicExamplesLabel"
        Me.lblTableGraphicExamplesLabel.Size = New System.Drawing.Size(102, 17)
        Me.lblTableGraphicExamplesLabel.TabIndex = 151
        Me.lblTableGraphicExamplesLabel.Text = "Example Table:"
        Me.lblTableGraphicExamplesLabel.Visible = False
        '
        'cmdShowRunSummary
        '
        Me.cmdShowRunSummary.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdShowRunSummary.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdShowRunSummary.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdShowRunSummary.FlatAppearance.BorderSize = 0
        Me.cmdShowRunSummary.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdShowRunSummary.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowRunSummary.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdShowRunSummary.Location = New System.Drawing.Point(1006, 38)
        Me.cmdShowRunSummary.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdShowRunSummary.Name = "cmdShowRunSummary"
        Me.cmdShowRunSummary.Size = New System.Drawing.Size(233, 31)
        Me.cmdShowRunSummary.TabIndex = 162
        Me.cmdShowRunSummary.Text = "Sho&w Run Summary..."
        Me.cmdShowRunSummary.UseVisualStyleBackColor = True
        '
        'lblOptional
        '
        Me.lblOptional.AutoSize = True
        Me.lblOptional.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblOptional.Location = New System.Drawing.Point(11, 52)
        Me.lblOptional.Name = "lblOptional"
        Me.lblOptional.Size = New System.Drawing.Size(55, 17)
        Me.lblOptional.TabIndex = 6
        Me.lblOptional.Text = "Label16"
        '
        'lblWatsonE
        '
        Me.lblWatsonE.AutoSize = True
        Me.lblWatsonE.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblWatsonE.Location = New System.Drawing.Point(11, 35)
        Me.lblWatsonE.Name = "lblWatsonE"
        Me.lblWatsonE.Size = New System.Drawing.Size(55, 17)
        Me.lblWatsonE.TabIndex = 5
        Me.lblWatsonE.Text = "Label16"
        '
        'lblAccepted
        '
        Me.lblAccepted.AutoSize = True
        Me.lblAccepted.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblAccepted.Location = New System.Drawing.Point(11, 18)
        Me.lblAccepted.Name = "lblAccepted"
        Me.lblAccepted.Size = New System.Drawing.Size(55, 17)
        Me.lblAccepted.TabIndex = 4
        Me.lblAccepted.Text = "Label16"
        '
        'panEdit
        '
        Me.panEdit.CausesValidation = False
        Me.panEdit.Controls.Add(Me.cmdPasteConditions)
        Me.panEdit.Controls.Add(Me.cmdExit)
        Me.panEdit.Controls.Add(Me.cmdCancel)
        Me.panEdit.Controls.Add(Me.cmdEdit)
        Me.panEdit.Controls.Add(Me.cmdSave)
        Me.panEdit.Location = New System.Drawing.Point(678, 36)
        Me.panEdit.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.panEdit.Name = "panEdit"
        Me.panEdit.Size = New System.Drawing.Size(601, 45)
        Me.panEdit.TabIndex = 158
        '
        'cmdPasteConditions
        '
        Me.cmdPasteConditions.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPasteConditions.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPasteConditions.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdPasteConditions.Location = New System.Drawing.Point(339, 0)
        Me.cmdPasteConditions.Name = "cmdPasteConditions"
        Me.cmdPasteConditions.Size = New System.Drawing.Size(252, 44)
        Me.cmdPasteConditions.TabIndex = 161
        Me.cmdPasteConditions.Text = "&Paste Table Names, Conditions and Auto-Assignments..."
        Me.cmdPasteConditions.UseVisualStyleBackColor = True
        '
        'txtTitle
        '
        Me.txtTitle.Location = New System.Drawing.Point(14, 48)
        Me.txtTitle.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.txtTitle.Multiline = True
        Me.txtTitle.Name = "txtTitle"
        Me.txtTitle.Size = New System.Drawing.Size(727, 58)
        Me.txtTitle.TabIndex = 159
        '
        'cmdSymbol
        '
        Me.cmdSymbol.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSymbol.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSymbol.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdSymbol.Location = New System.Drawing.Point(1404, 5)
        Me.cmdSymbol.Name = "cmdSymbol"
        Me.cmdSymbol.Size = New System.Drawing.Size(150, 33)
        Me.cmdSymbol.TabIndex = 160
        Me.cmdSymbol.Text = "Show Symbol Copy"
        Me.cmdSymbol.UseVisualStyleBackColor = True
        '
        'cmdTest
        '
        Me.cmdTest.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdTest.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdTest.Location = New System.Drawing.Point(1285, 56)
        Me.cmdTest.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cmdTest.Name = "cmdTest"
        Me.cmdTest.Size = New System.Drawing.Size(169, 33)
        Me.cmdTest.TabIndex = 161
        Me.cmdTest.Text = "Test"
        Me.cmdTest.UseVisualStyleBackColor = True
        Me.cmdTest.Visible = False
        '
        'lblClose
        '
        Me.lblClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lblClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClose.ForeColor = System.Drawing.Color.Black
        Me.lblClose.Location = New System.Drawing.Point(460, 18)
        Me.lblClose.Name = "lblClose"
        Me.lblClose.Size = New System.Drawing.Size(99, 22)
        Me.lblClose.TabIndex = 162
        Me.lblClose.Text = "Closing..."
        Me.lblClose.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblClose.Visible = False
        '
        'frmReportTableConfig
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1561, 779)
        Me.Controls.Add(Me.lblClose)
        Me.Controls.Add(Me.cmdTest)
        Me.Controls.Add(Me.cmdSymbol)
        Me.Controls.Add(Me.txtTitle)
        Me.Controls.Add(Me.panEdit)
        Me.Controls.Add(Me.tabRTC)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.cmdResize)
        Me.Controls.Add(Me.dgvReportTables)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReportTableConfig"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Advanced Table Configuration Window"
        CType(Me.dgvReportTables, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbRTC_Samples.ResumeLayout(False)
        Me.gbRTC_Samples.PerformLayout()
        Me.gbRTC_QC.ResumeLayout(False)
        Me.gbRTC_QC.PerformLayout()
        Me.gbRTC_CalStd.ResumeLayout(False)
        Me.gbRTC_CalStd.PerformLayout()
        Me.gbCalStdValues.ResumeLayout(False)
        Me.gbCalStdValues.PerformLayout()
        Me.gbxSuper.ResumeLayout(False)
        Me.gbxSuper.PerformLayout()
        Me.gbStats.ResumeLayout(False)
        Me.gbStats.PerformLayout()
        Me.panOptions.ResumeLayout(False)
        Me.panOptions.PerformLayout()
        Me.gbAdditional.ResumeLayout(False)
        Me.gbAdditional.PerformLayout()
        Me.panTP.ResumeLayout(False)
        Me.panTP.PerformLayout()
        Me.panCycles.ResumeLayout(False)
        Me.panCycles.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.gbAnovaStats.ResumeLayout(False)
        Me.gbAnovaStats.PerformLayout()
        Me.gbSampleSort.ResumeLayout(False)
        Me.gbSampleSort.PerformLayout()
        Me.gbSampleGroup.ResumeLayout(False)
        Me.gbSampleGroup.PerformLayout()
        Me.gbPSAE.ResumeLayout(False)
        Me.gbPSAE.PerformLayout()
        Me.gbResultsChoice.ResumeLayout(False)
        Me.gbResultsChoice.PerformLayout()
        Me.panIS.ResumeLayout(False)
        Me.panIS.PerformLayout()
        Me.gbTableLegend.ResumeLayout(False)
        Me.gbTableLegend.PerformLayout()
        Me.panTitleLegends.ResumeLayout(False)
        Me.panTitleLegends.PerformLayout()
        Me.panNomDenomCalcs.ResumeLayout(False)
        Me.panNomDenomCalcs.PerformLayout()
        Me.panNomDenom.ResumeLayout(False)
        Me.gbDenom.ResumeLayout(False)
        Me.gbDenom.PerformLayout()
        Me.gbNumerator.ResumeLayout(False)
        Me.gbNumerator.PerformLayout()
        Me.gbCalcs.ResumeLayout(False)
        Me.gbCalcs.PerformLayout()
        Me.tabRTC.ResumeLayout(False)
        Me.tpFormat.ResumeLayout(False)
        Me.panFormat.ResumeLayout(False)
        Me.gbMatrixFactor.ResumeLayout(False)
        Me.gbMatrixFactor.PerformLayout()
        Me.panMF1.ResumeLayout(False)
        Me.gbGroupSort.ResumeLayout(False)
        Me.gbSAT.ResumeLayout(False)
        Me.gbQCGroup.ResumeLayout(False)
        Me.gbQCGroup.PerformLayout()
        Me.gbRegrULOQ.ResumeLayout(False)
        Me.gbCriteria.ResumeLayout(False)
        Me.gbCriteria.PerformLayout()
        Me.gbCarryover.ResumeLayout(False)
        Me.gbCarryover.PerformLayout()
        Me.gbLegendFormat.ResumeLayout(False)
        Me.gbIncSampleCriteria.ResumeLayout(False)
        CType(Me.dgvAnalytes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tpPeriodTemp.ResumeLayout(False)
        Me.panFDARef.ResumeLayout(False)
        Me.panFDARef.PerformLayout()
        Me.gbStabilityType.ResumeLayout(False)
        Me.gbStabilityType.PerformLayout()
        Me.tpAutoAssignment.ResumeLayout(False)
        Me.tpAutoAssignment.PerformLayout()
        CType(Me.dgvSAS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvASP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panTableGraphicExamples.ResumeLayout(False)
        CType(Me.pbxTableGraphicExamples, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panEdit.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents dgvReportTables As System.Windows.Forms.DataGridView
    Friend WithEvents cmdResize As System.Windows.Forms.Button
    Friend WithEvents gbRTC_Samples As System.Windows.Forms.GroupBox
    Friend WithEvents rbDontShowBQL As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowBQL As System.Windows.Forms.RadioButton
    Friend WithEvents gbRTC_QC As System.Windows.Forms.GroupBox
    Friend WithEvents rbRTC_QC_All As System.Windows.Forms.RadioButton
    Friend WithEvents rbRTC_QC_Acc As System.Windows.Forms.RadioButton
    Friend WithEvents gbRTC_CalStd As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rbRTC_CalStd_Acc As System.Windows.Forms.RadioButton
    Friend WithEvents rbRTC_CalStd_All As System.Windows.Forms.RadioButton
    Friend WithEvents gbCalStdValues As System.Windows.Forms.GroupBox
    Friend WithEvents rbDontShowRejected As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowRejectedValues As System.Windows.Forms.RadioButton
    Friend WithEvents gbStats As System.Windows.Forms.GroupBox
    Friend WithEvents chkMean As System.Windows.Forms.CheckBox
    Friend WithEvents chkN As System.Windows.Forms.CheckBox
    Friend WithEvents chkBias As System.Windows.Forms.CheckBox
    Friend WithEvents chkCV As System.Windows.Forms.CheckBox
    Friend WithEvents chkSD As System.Windows.Forms.CheckBox
    Friend WithEvents gbAdditional As System.Windows.Forms.GroupBox
    Friend WithEvents panCycles As System.Windows.Forms.Panel
    Friend WithEvents INTNUMBEROFCYCLES As System.Windows.Forms.TextBox
    Friend WithEvents lblCycles As System.Windows.Forms.Label
    Friend WithEvents panTP As System.Windows.Forms.Panel
    Friend WithEvents CHARPERIODTEMP As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents CHARTIMEFRAME As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CHARTIMEPERIOD As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmdIncSamples As System.Windows.Forms.Button
    Friend WithEvents CHARSTABILITYPERIOD As System.Windows.Forms.TextBox
    Friend WithEvents cmdBuild As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents cmsFieldCodes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents chkDiff As System.Windows.Forms.CheckBox
    Friend WithEvents chkDiffCol As System.Windows.Forms.CheckBox
    Friend WithEvents chkRegr As System.Windows.Forms.CheckBox
    Friend WithEvents gbxSuper As System.Windows.Forms.GroupBox
    Friend WithEvents rbNR As System.Windows.Forms.RadioButton
    Friend WithEvents rbOutier As System.Windows.Forms.RadioButton
    Friend WithEvents chkTheoretical As System.Windows.Forms.CheckBox
    Friend WithEvents gbAnovaStats As System.Windows.Forms.GroupBox
    Friend WithEvents chkIncludeAnova As System.Windows.Forms.CheckBox
    Friend WithEvents chkBQLLEGEND As System.Windows.Forms.CheckBox
    Friend WithEvents gbSampleSort As System.Windows.Forms.GroupBox
    Friend WithEvents cbxSampleS1 As System.Windows.Forms.ComboBox
    Friend WithEvents lblLevel As System.Windows.Forms.Label
    Friend WithEvents lblSort As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbxSampleSAD4 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleS4 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleSAD3 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleS3 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleSAD2 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleS2 As System.Windows.Forms.ComboBox
    Friend WithEvents lblA As System.Windows.Forms.Label
    Friend WithEvents cbxSampleSAD1 As System.Windows.Forms.ComboBox
    Friend WithEvents gbSampleGroup As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cbxSampleGAD4 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleG4 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleGAD3 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleG3 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleGAD2 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleG2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cbxSampleGAD1 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleG1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents chkIncludePSAE As System.Windows.Forms.CheckBox
    Friend WithEvents gbPSAE As System.Windows.Forms.GroupBox
    Friend WithEvents gbResultsChoice As System.Windows.Forms.GroupBox
    Friend WithEvents rbUsePeakAreaRatio As System.Windows.Forms.RadioButton
    Friend WithEvents rbUsePeakArea As System.Windows.Forms.RadioButton
    Friend WithEvents rbConc As System.Windows.Forms.RadioButton
    Friend WithEvents chkIncludeIS As System.Windows.Forms.CheckBox
    Friend WithEvents gbTableLegend As System.Windows.Forms.GroupBox
    Friend WithEvents rbPosLeg As System.Windows.Forms.RadioButton
    Friend WithEvents rbNegLeg As System.Windows.Forms.RadioButton
    Friend WithEvents chkCustomLeg As System.Windows.Forms.CheckBox
    Friend WithEvents lblCHARDENLEG As System.Windows.Forms.Label
    Friend WithEvents CHARDENLEG As System.Windows.Forms.TextBox
    Friend WithEvents lblCHARNUMLEG As System.Windows.Forms.Label
    Friend WithEvents CHARNUMLEG As System.Windows.Forms.TextBox
    Friend WithEvents lblCHARTITLELEG As System.Windows.Forms.Label
    Friend WithEvents CHARTITLELEG As System.Windows.Forms.TextBox
    Friend WithEvents chkNoneLeg As System.Windows.Forms.CheckBox
    Friend WithEvents chkIncludeDate As System.Windows.Forms.CheckBox
    Friend WithEvents lblDivider As System.Windows.Forms.Label
    Friend WithEvents gbCalcs As System.Windows.Forms.GroupBox
    Friend WithEvents rbMeanAcc As System.Windows.Forms.RadioButton
    Friend WithEvents rbRecovery As System.Windows.Forms.RadioButton
    Friend WithEvents rbDifference As System.Windows.Forms.RadioButton
    Friend WithEvents chkIncludeWatsonLabel As System.Windows.Forms.CheckBox
    Friend WithEvents tabRTC As System.Windows.Forms.TabControl
    Friend WithEvents tpFormat As System.Windows.Forms.TabPage
    Friend WithEvents chkIncludeAnovaSumStats As System.Windows.Forms.CheckBox
    Friend WithEvents tpPeriodTemp As System.Windows.Forms.TabPage
    Friend WithEvents chkRE As System.Windows.Forms.CheckBox
    'Friend WithEvents lblPrecision As DevExpress.XtraEditors.LabelControl
    'Friend WithEvents lblAccuracy As DevExpress.XtraEditors.LabelControl
    Friend WithEvents lblCHARSTABILITYPERIOD As System.Windows.Forms.Label
    Friend WithEvents cmdInsert As System.Windows.Forms.Button
    Friend WithEvents chkCONVERTTEMP As System.Windows.Forms.CheckBox
    Friend WithEvents chkCONVERTTIME As System.Windows.Forms.CheckBox
    Friend WithEvents lblPeriodTemp As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents panIS As System.Windows.Forms.Panel
    Friend WithEvents CHARISCONC As System.Windows.Forms.TextBox
    Friend WithEvents lblCHARISCONC As System.Windows.Forms.Label
    Friend WithEvents lblDiff As System.Windows.Forms.Label
    Friend WithEvents panOptions As System.Windows.Forms.Panel
    Friend WithEvents chkIntraRunSumStats As System.Windows.Forms.CheckBox
    Friend WithEvents gbIncSampleCriteria As System.Windows.Forms.GroupBox
    Friend WithEvents dgvAnalytes As System.Windows.Forms.DataGridView
    Friend WithEvents lblPrecision As System.Windows.Forms.Label
    Friend WithEvents lblAccuracy As System.Windows.Forms.Label
    Friend WithEvents panFormat As System.Windows.Forms.Panel
    Friend WithEvents tpAutoAssignment As System.Windows.Forms.TabPage
    Friend WithEvents panEdit As System.Windows.Forms.Panel
    Friend WithEvents txtTitle As System.Windows.Forms.TextBox
    Friend WithEvents dgvSAS As System.Windows.Forms.DataGridView
    Friend WithEvents cmdSymbol As System.Windows.Forms.Button
    Friend WithEvents chkBOOLDOINDREC As System.Windows.Forms.CheckBox
    Friend WithEvents dgvASP As System.Windows.Forms.DataGridView
    Friend WithEvents lblAccepted As System.Windows.Forms.Label
    Friend WithEvents lblWatsonE As System.Windows.Forms.Label
    Friend WithEvents lblOptional As System.Windows.Forms.Label
    Friend WithEvents cmdShowRunSummary As System.Windows.Forms.Button
    Friend WithEvents panTableGraphicExamples As System.Windows.Forms.Panel
    Friend WithEvents pbxTableGraphicExamples As System.Windows.Forms.PictureBox
    Friend WithEvents lblTableGraphicExamplesLabel As System.Windows.Forms.Label
    Friend WithEvents cmdAnalRuns As System.Windows.Forms.Button
    Friend WithEvents chkBOOLISCOMBINELEVELS As System.Windows.Forms.CheckBox
    Friend WithEvents gbLegendFormat As System.Windows.Forms.GroupBox
    Friend WithEvents chkBOOLREASSAYREASLETTERS As System.Windows.Forms.CheckBox
    Friend WithEvents lblLogic As System.Windows.Forms.Label
    Friend WithEvents cmdTest As System.Windows.Forms.Button
    Friend WithEvents lblCase As System.Windows.Forms.Label
    Friend WithEvents lblRunIdentifier As System.Windows.Forms.Label
    Friend WithEvents chkIncludeIS_Single As System.Windows.Forms.CheckBox
    Friend WithEvents cmdPasteConditions As System.Windows.Forms.Button
    Friend WithEvents gbCarryover As System.Windows.Forms.GroupBox
    Friend WithEvents lblCarryover As System.Windows.Forms.Label
    Friend WithEvents CHARCARRYOVERLABEL As System.Windows.Forms.TextBox
    Friend WithEvents gbMatrixFactor As System.Windows.Forms.GroupBox
    Friend WithEvents chkInclIntStdNMF As System.Windows.Forms.CheckBox
    Friend WithEvents chkCalcIntStdNMF As System.Windows.Forms.CheckBox
    Friend WithEvents chkInclMFCols As System.Windows.Forms.CheckBox
    Friend WithEvents panMF1 As System.Windows.Forms.Panel
    Friend WithEvents chkMFTable As System.Windows.Forms.CheckBox
    Friend WithEvents lblMF01 As System.Windows.Forms.Label
    Friend WithEvents gbCriteria As System.Windows.Forms.GroupBox
    Friend WithEvents NUMPRECCRITLOTS As System.Windows.Forms.TextBox
    Friend WithEvents lblPrecisionCrit As System.Windows.Forms.Label
    Friend WithEvents gbRegrULOQ As System.Windows.Forms.GroupBox
    Friend WithEvents chkBOOLREGRULOQ As System.Windows.Forms.CheckBox
    Friend WithEvents gbQCGroup As System.Windows.Forms.GroupBox
    Friend WithEvents rbINTQCLEVELGROUPQCLabel As System.Windows.Forms.RadioButton
    Friend WithEvents rbINTQCLEVELGROUPNomConc As System.Windows.Forms.RadioButton
    Friend WithEvents rbINTQCLEVELGROUPLevel As System.Windows.Forms.RadioButton
    Friend WithEvents lblQCGroup As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cbxSampleSAD6 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleS6 As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cbxSampleSAD5 As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleS5 As System.Windows.Forms.ComboBox
    Friend WithEvents gbSAT As System.Windows.Forms.GroupBox
    Friend WithEvents chkBOOLCONCCOMMENTS As System.Windows.Forms.CheckBox
    Friend WithEvents panTitleLegends As System.Windows.Forms.Panel
    Friend WithEvents gbGroupSort As System.Windows.Forms.GroupBox
    Friend WithEvents lblGroupSort As System.Windows.Forms.Label
    Friend WithEvents chkBOOLADHOCSTABCOMPCOLUMNS As System.Windows.Forms.CheckBox
    Friend WithEvents lblClose As System.Windows.Forms.Label
    Friend WithEvents gbDenom As System.Windows.Forms.GroupBox
    Friend WithEvents rbOld As System.Windows.Forms.RadioButton
    Friend WithEvents rbNew As System.Windows.Forms.RadioButton
    Friend WithEvents gbNumerator As System.Windows.Forms.GroupBox
    Friend WithEvents lblRecovery As System.Windows.Forms.Label
    Friend WithEvents lblPercDiff As System.Windows.Forms.Label
    Friend WithEvents lblMeanAccuracy As System.Windows.Forms.Label
    Friend WithEvents panNomDenom As System.Windows.Forms.Panel
    Friend WithEvents chkInjCol As System.Windows.Forms.CheckBox
    Friend WithEvents gbStabilityType As System.Windows.Forms.GroupBox
    Friend WithEvents rbStockSolution As System.Windows.Forms.RadioButton
    Friend WithEvents rbBlood As System.Windows.Forms.RadioButton
    Friend WithEvents rbLT As System.Windows.Forms.RadioButton
    Friend WithEvents rbFT As System.Windows.Forms.RadioButton
    Friend WithEvents rbBenchTop As System.Windows.Forms.RadioButton
    Friend WithEvents rbProcess As System.Windows.Forms.RadioButton
    Friend WithEvents rbNA As System.Windows.Forms.RadioButton
    Friend WithEvents txtStabilityNotes As System.Windows.Forms.TextBox
    Friend WithEvents lblStabilityNotes As System.Windows.Forms.Label
    Friend WithEvents rbReinjection As System.Windows.Forms.RadioButton
    Friend WithEvents lblRemember As System.Windows.Forms.Label
    Friend WithEvents rbSpiking As System.Windows.Forms.RadioButton
    Friend WithEvents panNomDenomCalcs As System.Windows.Forms.Panel
    Friend WithEvents chkRTC_CalStd_Acc As System.Windows.Forms.CheckBox
    Friend WithEvents rbAutosampler As System.Windows.Forms.RadioButton
    Friend WithEvents rbBatchReinjection As System.Windows.Forms.RadioButton
    Friend WithEvents panFDARef As System.Windows.Forms.Panel
    Friend WithEvents txtRef As System.Windows.Forms.TextBox
    Friend WithEvents lblFDA As System.Windows.Forms.Label
    Friend WithEvents rbDilution As System.Windows.Forms.RadioButton
    Friend WithEvents rbUseISPeakArea As System.Windows.Forms.RadioButton
End Class
