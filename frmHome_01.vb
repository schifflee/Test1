Option Compare Text

Imports System
Imports System.IO
Imports System.IO.FileSystemInfo
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports System.Data.Odbc

Imports System.Windows.Forms
Imports System.ComponentModel.PropertyDescriptorCollection
Imports Word = Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic
Imports Microsoft.Office.Interop.Word
Imports ta = GooWoo.ds_2005_GuWu_01TableAdapters
Imports taOra = GooWoo.ds_GuWuOra_01TableAdapters
'Imports taAccess = GooWoo.GuWu_01DataSetTableAdapters
Imports taAccess = GooWoo.StudyDoc_01DataSet1TableAdapters

Imports taSQLServer = GooWoo.StudyDoc_01SQLDataSetTableAdapters

Imports System.Text
Imports System.Drawing

Imports System.Linq.Expressions
Imports System.Linq
Imports System.Data.DataTableExtensions
Imports System.Data.DataRowExtensions

Imports ExtensionMethods


Public Class frmHome_01

    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        frmH = Me 'set this for later module reference

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents tab1 As System.Windows.Forms.TabControl
    Friend WithEvents tp1 As System.Windows.Forms.TabPage
    Friend WithEvents tp2 As System.Windows.Forms.TabPage
    Friend WithEvents tp3 As System.Windows.Forms.TabPage
    Friend WithEvents tp4 As System.Windows.Forms.TabPage
    Friend WithEvents tp6 As System.Windows.Forms.TabPage
    Friend WithEvents lbxTab1 As System.Windows.Forms.ListBox
    Friend WithEvents cmdUpdateProject As System.Windows.Forms.Button
    Friend WithEvents lblcbxStudies As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblTOC As System.Windows.Forms.Label
    Friend WithEvents lblHome As System.Windows.Forms.Label
    Friend WithEvents lblData As System.Windows.Forms.Label
    Friend WithEvents lblAnalyticalRunSummary As System.Windows.Forms.Label
    Friend WithEvents lblSummaryTable As System.Windows.Forms.Label
    Friend WithEvents lblReportTableConfiguration As System.Windows.Forms.Label
    Friend WithEvents dgHome As System.Windows.Forms.DataGrid
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lblGlobalConfiguration As System.Windows.Forms.Label
    Friend WithEvents lblWAR As System.Windows.Forms.Label
    Friend WithEvents lblWatsonStudy As System.Windows.Forms.Label
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdAddRepAnalyte As System.Windows.Forms.Button
    Friend WithEvents lblARST As System.Windows.Forms.Label
    Friend WithEvents rbUseWatsonComments As System.Windows.Forms.RadioButton
    Friend WithEvents rbUseUserComments As System.Windows.Forms.RadioButton
    Friend WithEvents cmdCPAdd As System.Windows.Forms.Button
    Friend WithEvents cmdConfigureReport As System.Windows.Forms.Button
    Friend WithEvents dgStudies As System.Windows.Forms.DataGrid
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblReportTitle As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents cbxAssayTechnique As System.Windows.Forms.ComboBox
    Friend WithEvents cbxAssayTechniqueAcronym As System.Windows.Forms.ComboBox
    Friend WithEvents cbxAnticoagulant As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSubmittedTo As System.Windows.Forms.ComboBox
    Friend WithEvents cbxInSupportOf As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSubmittedBy As System.Windows.Forms.ComboBox
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents Label43 As System.Windows.Forms.Label
    Friend WithEvents txtSubmittedTo As System.Windows.Forms.TextBox
    Friend WithEvents Label44 As System.Windows.Forms.Label
    Friend WithEvents txtSubmittedBy As System.Windows.Forms.TextBox
    Friend WithEvents txtInSupportOf As System.Windows.Forms.TextBox
    Friend WithEvents cmdCreateReportTitle As System.Windows.Forms.Button
    Friend WithEvents lblIncludeInTitle As System.Windows.Forms.Label
    Friend WithEvents cmdCPDelete As System.Windows.Forms.Button
    Friend WithEvents cbxSampleSizeUnits As System.Windows.Forms.ComboBox
    Friend WithEvents cbxSampleStorageTemp As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblDCP As System.Windows.Forms.Label
    Friend WithEvents cmdCPCancel As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdHomeCancel As System.Windows.Forms.Button
    Friend WithEvents cmdDataCancel As System.Windows.Forms.Button
    Friend WithEvents panAnalRunSum As System.Windows.Forms.Panel
    Friend WithEvents cmdAnaRunSumCancel As System.Windows.Forms.Button
    Friend WithEvents cmdRTConfigCancel As System.Windows.Forms.Button
    Friend WithEvents cmdAnalRefCancel As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteRepAnalyte As System.Windows.Forms.Button
    Friend WithEvents lblReportStatement As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents chkMethodValMultiple As System.Windows.Forms.CheckBox
    Friend WithEvents rbMethValAnalyte As System.Windows.Forms.RadioButton
    Friend WithEvents rbMethValMultiple As System.Windows.Forms.RadioButton
    Friend WithEvents txtMethValMultiple As System.Windows.Forms.TextBox
    Friend WithEvents lblMethValMultiple As System.Windows.Forms.Label
    Friend WithEvents gbMethodValMultiple As System.Windows.Forms.GroupBox
    Friend WithEvents cmdMethValReset As System.Windows.Forms.Button
    Friend WithEvents gbMethValApplyGuWu As System.Windows.Forms.GroupBox
    Friend WithEvents cbxMethValExistingGuWu As System.Windows.Forms.ComboBox
    Friend WithEvents lblM2 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents lblProgress As System.Windows.Forms.Label
    Friend WithEvents tp11 As System.Windows.Forms.TabPage
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents cmdRTHeaderConfigCancel As System.Windows.Forms.Button
    Friend WithEvents OleDbSelectCommand32 As System.Data.OleDb.OleDbCommand
    Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmHome_01))
    Friend WithEvents tp7 As System.Windows.Forms.TabPage
    Friend WithEvents tp8 As System.Windows.Forms.TabPage
    Friend WithEvents tp9 As System.Windows.Forms.TabPage
    Friend WithEvents tp5 As System.Windows.Forms.TabPage
    Friend WithEvents tp10 As System.Windows.Forms.TabPage

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdCancelReportStatements As System.Windows.Forms.Button
    Friend WithEvents cmdOpenReportStatements As System.Windows.Forms.Button
    Friend WithEvents cbxStudy As System.Windows.Forms.ComboBox
    Friend WithEvents lblTCH_01 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents dgQATable As System.Windows.Forms.DataGrid
    Friend WithEvents cmdInsertQAEvent As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteQAEvent As System.Windows.Forms.Button
    Friend WithEvents cmdQACancel As System.Windows.Forms.Button
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents tp12 As System.Windows.Forms.TabPage
    Friend WithEvents tp13 As System.Windows.Forms.TabPage
    Friend WithEvents cmdApplyTemplate As System.Windows.Forms.Button
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents cmdCopyRepAnalyte As System.Windows.Forms.Button
    Friend WithEvents dgvAnalyticalRunSummary As System.Windows.Forms.DataGridView
    Friend WithEvents dgvReportTableConfiguration As System.Windows.Forms.DataGridView
    Friend WithEvents lblRTC As System.Windows.Forms.Label
    Friend WithEvents dgvWatsonAnalRef As System.Windows.Forms.DataGridView
    Friend WithEvents dgvReportStatements As System.Windows.Forms.DataGridView
    Friend WithEvents lblQAHyperlink As System.Windows.Forms.LinkLabel
    Friend WithEvents pb1 As System.Windows.Forms.ProgressBar
    Friend WithEvents tp14 As System.Windows.Forms.TabPage
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents cmdDeletSRec As System.Windows.Forms.Button
    Friend WithEvents cmdInsertSRec As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cmdAdministration As System.Windows.Forms.Button
    Friend WithEvents dgvSampleReceipt As System.Windows.Forms.DataGridView
    Friend WithEvents dgvUser As System.Windows.Forms.DataGridView
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents cmdSRecCancel As System.Windows.Forms.Button
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents txtSRecTotal As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents Label41 As System.Windows.Forms.Label
    Friend WithEvents txtSRecTotalReport As System.Windows.Forms.TextBox
    Friend WithEvents chkManualSampleNumber As System.Windows.Forms.CheckBox
    Friend WithEvents Label46 As System.Windows.Forms.Label
    Friend WithEvents dgvSampleReceiptWatson As System.Windows.Forms.DataGridView
    Friend WithEvents Label48 As System.Windows.Forms.Label
    Friend WithEvents chkUseWatsonSampleNumber As System.Windows.Forms.CheckBox
    Friend WithEvents txtSRecTotalReportWatson As System.Windows.Forms.TextBox
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents cmdMethValExecute As System.Windows.Forms.Button
    Friend WithEvents cmdShowOutstanding As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents gbRTC As System.Windows.Forms.GroupBox
    Friend WithEvents rbShowAllRTConfig As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowIncludedRTConfig As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowAllRBody As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowIncludedRBody As System.Windows.Forms.RadioButton
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents cmdOrderReportBodySection As System.Windows.Forms.Button
    Friend WithEvents cmdOrderReportTableConfig As System.Windows.Forms.Button
    Friend WithEvents lblARS_A As System.Windows.Forms.Label
    Friend WithEvents cmdResetSummaryTable As System.Windows.Forms.Button
    Friend WithEvents cmdOrderSummaryTable As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents grbShowSummaryTable As System.Windows.Forms.GroupBox
    Friend WithEvents rbShowIncludedSummaryTable As System.Windows.Forms.RadioButton
    Friend WithEvents rbShowAllSummaryTable As System.Windows.Forms.RadioButton
    Friend WithEvents cmdAssignSamples As System.Windows.Forms.Button
    Friend WithEvents lblColoredRows As System.Windows.Forms.Label
    Friend WithEvents TimerRTC As System.Windows.Forms.Timer
    Friend WithEvents llblAssignedSamples As System.Windows.Forms.LinkLabel
    Friend WithEvents chkQCShowExcludedBatch As System.Windows.Forms.CheckBox
    Friend WithEvents cmdAddAnalyte As System.Windows.Forms.Button
    Friend WithEvents cbxFilter As System.Windows.Forms.ComboBox
    Friend WithEvents lblARS As System.Windows.Forms.Label
    Friend WithEvents dgvContributingPersonnel As System.Windows.Forms.DataGridView
    Friend WithEvents Label53 As System.Windows.Forms.Label
    Friend WithEvents dgvDataCompany As System.Windows.Forms.DataGridView
    Friend WithEvents TimerSplash As System.Windows.Forms.Timer
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents dgvReports As System.Windows.Forms.DataGridView
    Friend WithEvents lblSelectCell As System.Windows.Forms.Label
    Friend WithEvents cbxRBSFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cbxRBSTypeFilter As System.Windows.Forms.ComboBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents cmdRBSAll As System.Windows.Forms.Button
    Friend WithEvents cmdRefreshStatements As System.Windows.Forms.Button
    Friend WithEvents llblSummaryTable As System.Windows.Forms.LinkLabel
    Friend WithEvents lblWordStatements As System.Windows.Forms.Label
    Friend WithEvents dgvReportStatementWord As System.Windows.Forms.DataGridView
    Friend WithEvents dgvwStudy As System.Windows.Forms.DataGridView
    Friend WithEvents cmdHook As System.Windows.Forms.Button
    Friend WithEvents cmdUpdateSummaryInfo As System.Windows.Forms.Button
    Friend WithEvents cbxExampleReport As System.Windows.Forms.ComboBox
    Friend WithEvents cmdAppFig As System.Windows.Forms.Button
    Friend WithEvents grpRBS As System.Windows.Forms.GroupBox
    Friend WithEvents rbRBS_Section As System.Windows.Forms.RadioButton
    Friend WithEvents rbRBS_Col As System.Windows.Forms.RadioButton
    Friend WithEvents lblRBS As System.Windows.Forms.Label
    Friend WithEvents tp15 As System.Windows.Forms.TabPage
    Friend WithEvents cmdAnalDetails As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents panFilterStudy As System.Windows.Forms.Panel
    Friend WithEvents cbxFilterStudy As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents cmdFilterStudy As System.Windows.Forms.Button
    Friend WithEvents txtFilterStudy As System.Windows.Forms.TextBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents txtcbxMDBSelIndex As System.Windows.Forms.TextBox
    Friend WithEvents txtFilterIndex As System.Windows.Forms.TextBox
    Friend WithEvents ms1 As System.Windows.Forms.MenuStrip
    Friend WithEvents mnuAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdViewAnalRuns As System.Windows.Forms.Button
    Friend WithEvents gbSource As System.Windows.Forms.GroupBox
    Friend WithEvents rbArchive As System.Windows.Forms.RadioButton
    Friend WithEvents rbOracle As System.Windows.Forms.RadioButton
    Friend WithEvents txtArchivePath As System.Windows.Forms.TextBox
    Friend WithEvents cmdResize As System.Windows.Forms.Button
    Friend WithEvents cmdViewAnalyticalRuns1 As System.Windows.Forms.Button
    Friend WithEvents cmdCreateReportTitle2 As System.Windows.Forms.Button
    Friend WithEvents dgvMethodValData As System.Windows.Forms.DataGridView
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents cmdDuplicateTables As System.Windows.Forms.Button
    Friend WithEvents cmdAdvancedTable As System.Windows.Forms.Button
    Friend WithEvents chkTableName As System.Windows.Forms.CheckBox
    Friend WithEvents cmdHeader As System.Windows.Forms.Button
    Friend WithEvents cmsHome As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents cmiHomeFieldCode As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents panRBSwb As System.Windows.Forms.Panel
    Friend WithEvents wbRBS As System.Windows.Forms.WebBrowser
    Friend WithEvents cmdReportHistory As System.Windows.Forms.Button
    Friend WithEvents cmdOutliers As System.Windows.Forms.Button
    Friend WithEvents gbSectionStyle As System.Windows.Forms.GroupBox
    Friend WithEvents rbEntireReport As System.Windows.Forms.RadioButton
    Friend WithEvents rbSections As System.Windows.Forms.RadioButton
    Friend WithEvents panSections As System.Windows.Forms.Panel
    Friend WithEvents lbldgvReportTables As System.Windows.Forms.Label
    Friend WithEvents dgvReportTables As System.Windows.Forms.DataGridView
    Friend WithEvents dgvReportTableHeaderConfig As System.Windows.Forms.DataGridView
    Friend WithEvents lbldgvReportTableHeaderConfig As System.Windows.Forms.Label
    Friend WithEvents pb2 As System.Windows.Forms.ProgressBar
    Friend WithEvents mnuMenuAbout As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMenuGenFC As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cbxArchivedMDB As System.Windows.Forms.ComboBox
    Friend WithEvents cmdBrowseMDB As System.Windows.Forms.Button
    Friend WithEvents cmdMethValUpdate As System.Windows.Forms.Button
    Friend WithEvents gbxMultVal As System.Windows.Forms.GroupBox
    Friend WithEvents rbMultValNo As System.Windows.Forms.RadioButton
    Friend WithEvents lblRTCUpDown As System.Windows.Forms.Label
    Friend WithEvents cmdRTCDown As System.Windows.Forms.Button
    Friend WithEvents cmdRTCUp As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents chkQAEventBorder As System.Windows.Forms.CheckBox
    Friend WithEvents lblARS_B As System.Windows.Forms.Label
    Friend WithEvents lblARS_B_Note As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents tp16 As System.Windows.Forms.TabPage
    Friend WithEvents cmdAuditTrail As System.Windows.Forms.Button
    Friend WithEvents lblAuditTrail As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents gbReportOptions As System.Windows.Forms.GroupBox
    Friend WithEvents cmdSelect As System.Windows.Forms.Button
    Friend WithEvents cmdReplacePersonnel As System.Windows.Forms.Button
    Friend WithEvents dtp1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkReadOnlyTables As System.Windows.Forms.CheckBox
    Friend WithEvents tabData As System.Windows.Forms.TabControl
    Friend WithEvents tabData1 As System.Windows.Forms.TabPage
    Friend WithEvents tabData2 As System.Windows.Forms.TabPage
    Friend WithEvents dgvStudyConfig As System.Windows.Forms.DataGridView
    Friend WithEvents tabData3 As System.Windows.Forms.TabPage
    Friend WithEvents dgvFC As System.Windows.Forms.DataGridView
    Friend WithEvents mnuShowFC As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents cmdImportTables As System.Windows.Forms.Button
    Friend WithEvents pbxWord As System.Windows.Forms.PictureBox
    Friend WithEvents cmdCreateTable As System.Windows.Forms.Button
    Friend WithEvents cmdSymbol As System.Windows.Forms.Button
    Friend WithEvents gbxlblReviewAnalyticalRuns As System.Windows.Forms.GroupBox
    Friend WithEvents panEdit As System.Windows.Forms.Panel
    Friend WithEvents panChoose As System.Windows.Forms.Panel
    Friend WithEvents panSumTable As System.Windows.Forms.Panel
    Friend WithEvents panAnalRuns As System.Windows.Forms.Panel
    Friend WithEvents panTopLevel As System.Windows.Forms.Panel
    Friend WithEvents panSampleRec As System.Windows.Forms.Panel
    Friend WithEvents panQAEvent As System.Windows.Forms.Panel
    Friend WithEvents panMethVal As System.Windows.Forms.Panel
    Friend WithEvents panContr As System.Windows.Forms.Panel
    Friend WithEvents panAnalRefStds As System.Windows.Forms.Panel
    Friend WithEvents panColHeadings As System.Windows.Forms.Panel
    Friend WithEvents panRepTables As System.Windows.Forms.Panel
    Friend WithEvents panWordTemp As System.Windows.Forms.Panel
    Friend WithEvents gbStudyFilter As System.Windows.Forms.GroupBox
    Friend WithEvents optStudyDocClosed As System.Windows.Forms.RadioButton
    Friend WithEvents optStudyDocOpen As System.Windows.Forms.RadioButton
    Friend WithEvents lblWarning As System.Windows.Forms.Label
    Friend WithEvents TimerWarning As System.Windows.Forms.Timer
    Friend WithEvents lblWatsonWarning As System.Windows.Forms.Label
    Friend WithEvents dgvCompanyAnalRef As System.Windows.Forms.DataGridView
    Friend WithEvents optStudyDocStudies As System.Windows.Forms.RadioButton
    Friend WithEvents rbMultValYes As System.Windows.Forms.RadioButton
    Friend WithEvents panDot As System.Windows.Forms.Panel
    Friend WithEvents lbl2 As System.Windows.Forms.Label
    Friend WithEvents lbl1 As System.Windows.Forms.Label
    Friend WithEvents lblBlack As System.Windows.Forms.Label
    Friend WithEvents mCal1 As System.Windows.Forms.MonthCalendar
    Friend WithEvents cmdEnterCal As System.Windows.Forms.Button
    Friend WithEvents cmdCalCancel As System.Windows.Forms.Button
    Friend WithEvents panCal As System.Windows.Forms.Panel
    Friend WithEvents lblAT As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents panStudyFilter As System.Windows.Forms.Panel
    Friend WithEvents lblWatsonData As System.Windows.Forms.Label
    Friend WithEvents panWatsonData As System.Windows.Forms.Panel
    Friend WithEvents lblWatsonDataTitle As System.Windows.Forms.Label
    Friend WithEvents txtFilterSamples As System.Windows.Forms.TextBox
    Friend WithEvents cmdClearFilters As System.Windows.Forms.Button
    Friend WithEvents gbFilters As System.Windows.Forms.GroupBox
    Friend WithEvents tabData4 As System.Windows.Forms.TabPage
    Friend WithEvents gbMeanComp As System.Windows.Forms.GroupBox
    Friend WithEvents rbMeanRounded As System.Windows.Forms.RadioButton
    Friend WithEvents rbMeanFullPrec As System.Windows.Forms.RadioButton
    Friend WithEvents gbCritPrecision As System.Windows.Forms.GroupBox
    Friend WithEvents rbCritRounded As System.Windows.Forms.RadioButton
    Friend WithEvents rbCritFullPrec As System.Windows.Forms.RadioButton
    Friend WithEvents gbRound5 As System.Windows.Forms.GroupBox
    Friend WithEvents rbRoundFiveAway As System.Windows.Forms.RadioButton
    Friend WithEvents rbRoundFiveEven As System.Windows.Forms.RadioButton
    Friend WithEvents lblTableGraphicExamplesLabel As System.Windows.Forms.Label
    Friend WithEvents pbxTableGraphicExamples As System.Windows.Forms.PictureBox
    Friend WithEvents chkTableGraphicExamples As System.Windows.Forms.CheckBox
    Friend WithEvents panTableGraphicExamples As System.Windows.Forms.Panel
    Friend WithEvents lblTableGraphicExamplesText As System.Windows.Forms.Label
    Friend WithEvents gbxlblMethodValidation As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlblAddEditContributors As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlblConfigureReportTables1 As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlblAnalyticalReferenceStd As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlblSampleReceiptRecords1 As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlblSampleReceiptRecords2 As System.Windows.Forms.GroupBox
    Friend WithEvents panPrepareReportInside As System.Windows.Forms.Panel
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents MenuPrepareReport As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrepareEntireReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrepareOnlySelectedTableToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrepareOnlyReportBodyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrepareOnlyReportTablesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents panPrepareReportOutside As System.Windows.Forms.Panel
    Friend WithEvents gbxlblConfigureColumnHeadings1 As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlblChooseEditWordTemplate As System.Windows.Forms.GroupBox
    Friend WithEvents gbxlblReviewValidatedMethod As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdShowGroups As System.Windows.Forms.Button
    Friend WithEvents dgvGroups As System.Windows.Forms.DataGridView
    Friend WithEvents mnuTroubleshooting As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents dgvMethValExistingGuWu As System.Windows.Forms.DataGridView
    Friend WithEvents dgvSummaryData As System.Windows.Forms.DataGridView
    Friend WithEvents dgvFieldCodes As System.Windows.Forms.DataGridView
    Friend WithEvents lblActions As System.Windows.Forms.Label
    Friend WithEvents panAnalRunChoices As System.Windows.Forms.Panel
    Friend WithEvents chkPSAE As System.Windows.Forms.CheckBox
    Friend WithEvents chkNoRegrPerformed As System.Windows.Forms.CheckBox
    Friend WithEvents chkRegrPerformed As System.Windows.Forms.CheckBox
    Friend WithEvents chkRejected As System.Windows.Forms.CheckBox
    Friend WithEvents chkAccepted As System.Windows.Forms.CheckBox
    Friend WithEvents chkAll As System.Windows.Forms.CheckBox
    Friend WithEvents lblAnalRunReportOptions As System.Windows.Forms.Label
    Friend WithEvents gbInclude As System.Windows.Forms.GroupBox
    Friend WithEvents rbAll As System.Windows.Forms.RadioButton
    Friend WithEvents rbInclude As System.Windows.Forms.RadioButton
    Friend WithEvents panProgress As System.Windows.Forms.Panel
    Friend WithEvents lblFinalReportLockedDate As System.Windows.Forms.Label
    Friend WithEvents lblFinalReportLockedDateLabel As System.Windows.Forms.Label
    Friend WithEvents chkLockFinalReport As System.Windows.Forms.CheckBox
    Friend WithEvents panLockFinalReport As System.Windows.Forms.Panel
    Friend WithEvents tabData5 As System.Windows.Forms.TabPage
    Friend WithEvents dgvAnalyteGroups As System.Windows.Forms.DataGridView
    Friend WithEvents lblSortAnalyte As System.Windows.Forms.Label
    Friend WithEvents cmdDownA As System.Windows.Forms.Button
    Friend WithEvents cmdUpA As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents cmdDownCF As System.Windows.Forms.Button
    Friend WithEvents cmdUpCF As System.Windows.Forms.Button
    Friend WithEvents lblCFAuditTrail As System.Windows.Forms.Label
    Friend WithEvents cmdApplyTables As System.Windows.Forms.Button
    Friend WithEvents dgvDataWatson As System.Windows.Forms.DataGridView
    Friend WithEvents cmdClearStudy As System.Windows.Forms.Button
    Friend WithEvents lblNotes2 As System.Windows.Forms.Label
    Friend WithEvents lblNotes1 As System.Windows.Forms.Label
    Friend WithEvents panActions As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle162 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle161 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle163 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle164 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle165 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle166 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmHome_01))
        Dim DataGridViewCellStyle167 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle168 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle169 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle170 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle171 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle172 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle173 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle174 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle175 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle176 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle177 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle178 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle179 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle180 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle181 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle182 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle183 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle184 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle185 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle186 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle187 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle188 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle189 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle190 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle191 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle192 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.tab1 = New System.Windows.Forms.TabControl()
        Me.tp1 = New System.Windows.Forms.TabPage()
        Me.dgvReports = New System.Windows.Forms.DataGridView()
        Me.panLockFinalReport = New System.Windows.Forms.Panel()
        Me.chkLockFinalReport = New System.Windows.Forms.CheckBox()
        Me.lblFinalReportLockedDate = New System.Windows.Forms.Label()
        Me.lblFinalReportLockedDateLabel = New System.Windows.Forms.Label()
        Me.cmdHeader = New System.Windows.Forms.Button()
        Me.dgStudies = New System.Windows.Forms.DataGrid()
        Me.dgHome = New System.Windows.Forms.DataGrid()
        Me.cmdConfigureReport = New System.Windows.Forms.Button()
        Me.dgvwStudy = New System.Windows.Forms.DataGridView()
        Me.lblSelectCell = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.cmdCreateReportTitle = New System.Windows.Forms.Button()
        Me.lblHome = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblcbxStudies = New System.Windows.Forms.Label()
        Me.tp2 = New System.Windows.Forms.TabPage()
        Me.dgvDataWatson = New System.Windows.Forms.DataGridView()
        Me.cbxAssayTechnique = New System.Windows.Forms.ComboBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.tabData = New System.Windows.Forms.TabControl()
        Me.tabData1 = New System.Windows.Forms.TabPage()
        Me.dgvDataCompany = New System.Windows.Forms.DataGridView()
        Me.tabData2 = New System.Windows.Forms.TabPage()
        Me.dgvStudyConfig = New System.Windows.Forms.DataGridView()
        Me.tabData3 = New System.Windows.Forms.TabPage()
        Me.lblCFAuditTrail = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.cmdDownCF = New System.Windows.Forms.Button()
        Me.cmdUpCF = New System.Windows.Forms.Button()
        Me.gbInclude = New System.Windows.Forms.GroupBox()
        Me.rbAll = New System.Windows.Forms.RadioButton()
        Me.rbInclude = New System.Windows.Forms.RadioButton()
        Me.dgvFC = New System.Windows.Forms.DataGridView()
        Me.tabData4 = New System.Windows.Forms.TabPage()
        Me.gbMeanComp = New System.Windows.Forms.GroupBox()
        Me.rbMeanRounded = New System.Windows.Forms.RadioButton()
        Me.rbMeanFullPrec = New System.Windows.Forms.RadioButton()
        Me.gbCritPrecision = New System.Windows.Forms.GroupBox()
        Me.rbCritRounded = New System.Windows.Forms.RadioButton()
        Me.rbCritFullPrec = New System.Windows.Forms.RadioButton()
        Me.gbRound5 = New System.Windows.Forms.GroupBox()
        Me.rbRoundFiveAway = New System.Windows.Forms.RadioButton()
        Me.rbRoundFiveEven = New System.Windows.Forms.RadioButton()
        Me.tabData5 = New System.Windows.Forms.TabPage()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lblSortAnalyte = New System.Windows.Forms.Label()
        Me.cmdDownA = New System.Windows.Forms.Button()
        Me.cmdUpA = New System.Windows.Forms.Button()
        Me.dgvAnalyteGroups = New System.Windows.Forms.DataGridView()
        Me.cbxSubmittedTo = New System.Windows.Forms.ComboBox()
        Me.lblDCP = New System.Windows.Forms.Label()
        Me.cbxSampleStorageTemp = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblIncludeInTitle = New System.Windows.Forms.Label()
        Me.Label44 = New System.Windows.Forms.Label()
        Me.txtSubmittedTo = New System.Windows.Forms.TextBox()
        Me.Label43 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtSubmittedBy = New System.Windows.Forms.TextBox()
        Me.txtInSupportOf = New System.Windows.Forms.TextBox()
        Me.cbxSubmittedBy = New System.Windows.Forms.ComboBox()
        Me.cbxInSupportOf = New System.Windows.Forms.ComboBox()
        Me.cbxSampleSizeUnits = New System.Windows.Forms.ComboBox()
        Me.cbxAnticoagulant = New System.Windows.Forms.ComboBox()
        Me.cbxAssayTechniqueAcronym = New System.Windows.Forms.ComboBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.lblData = New System.Windows.Forms.Label()
        Me.tp3 = New System.Windows.Forms.TabPage()
        Me.lblARS_B_Note = New System.Windows.Forms.Label()
        Me.dgvAnalyticalRunSummary = New System.Windows.Forms.DataGridView()
        Me.lblAnalyticalRunSummary = New System.Windows.Forms.Label()
        Me.gbxlblReviewAnalyticalRuns = New System.Windows.Forms.GroupBox()
        Me.lblARS_B = New System.Windows.Forms.Label()
        Me.lblARS_A = New System.Windows.Forms.Label()
        Me.gbReportOptions = New System.Windows.Forms.GroupBox()
        Me.panAnalRunSum = New System.Windows.Forms.Panel()
        Me.rbUseWatsonComments = New System.Windows.Forms.RadioButton()
        Me.rbUseUserComments = New System.Windows.Forms.RadioButton()
        Me.panAnalRunChoices = New System.Windows.Forms.Panel()
        Me.lblAnalRunReportOptions = New System.Windows.Forms.Label()
        Me.chkPSAE = New System.Windows.Forms.CheckBox()
        Me.chkNoRegrPerformed = New System.Windows.Forms.CheckBox()
        Me.chkRegrPerformed = New System.Windows.Forms.CheckBox()
        Me.chkRejected = New System.Windows.Forms.CheckBox()
        Me.chkAccepted = New System.Windows.Forms.CheckBox()
        Me.chkAll = New System.Windows.Forms.CheckBox()
        Me.tp4 = New System.Windows.Forms.TabPage()
        Me.dgvSummaryData = New System.Windows.Forms.DataGridView()
        Me.llblSummaryTable = New System.Windows.Forms.LinkLabel()
        Me.gbxlblMethodValidation = New System.Windows.Forms.GroupBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmdOrderSummaryTable = New System.Windows.Forms.Button()
        Me.lblSummaryTable = New System.Windows.Forms.Label()
        Me.tp5 = New System.Windows.Forms.TabPage()
        Me.gbxlblChooseEditWordTemplate = New System.Windows.Forms.GroupBox()
        Me.lblWordStatements = New System.Windows.Forms.Label()
        Me.panSections = New System.Windows.Forms.Panel()
        Me.cbxRBSFilter = New System.Windows.Forms.ComboBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.cbxRBSTypeFilter = New System.Windows.Forms.ComboBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.rbShowAllRBody = New System.Windows.Forms.RadioButton()
        Me.rbShowIncludedRBody = New System.Windows.Forms.RadioButton()
        Me.gbSectionStyle = New System.Windows.Forms.GroupBox()
        Me.rbEntireReport = New System.Windows.Forms.RadioButton()
        Me.rbSections = New System.Windows.Forms.RadioButton()
        Me.panRBSwb = New System.Windows.Forms.Panel()
        Me.cmsHome = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.cmiHomeFieldCode = New System.Windows.Forms.ToolStripMenuItem()
        Me.wbRBS = New System.Windows.Forms.WebBrowser()
        Me.cmdRBSAll = New System.Windows.Forms.Button()
        Me.grpRBS = New System.Windows.Forms.GroupBox()
        Me.rbRBS_Section = New System.Windows.Forms.RadioButton()
        Me.rbRBS_Col = New System.Windows.Forms.RadioButton()
        Me.cmdOrderReportBodySection = New System.Windows.Forms.Button()
        Me.lblRBS = New System.Windows.Forms.Label()
        Me.dgvReportStatementWord = New System.Windows.Forms.DataGridView()
        Me.dgvReportStatements = New System.Windows.Forms.DataGridView()
        Me.lblReportStatement = New System.Windows.Forms.Label()
        Me.tp6 = New System.Windows.Forms.TabPage()
        Me.cmdShowGroups = New System.Windows.Forms.Button()
        Me.dgvGroups = New System.Windows.Forms.DataGridView()
        Me.lblReportTableConfiguration = New System.Windows.Forms.Label()
        Me.panTableGraphicExamples = New System.Windows.Forms.Panel()
        Me.lblTableGraphicExamplesText = New System.Windows.Forms.Label()
        Me.pbxTableGraphicExamples = New System.Windows.Forms.PictureBox()
        Me.lblTableGraphicExamplesLabel = New System.Windows.Forms.Label()
        Me.chkReadOnlyTables = New System.Windows.Forms.CheckBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.lblRTCUpDown = New System.Windows.Forms.Label()
        Me.cmdRTCDown = New System.Windows.Forms.Button()
        Me.cmdRTCUp = New System.Windows.Forms.Button()
        Me.chkTableName = New System.Windows.Forms.CheckBox()
        Me.cmdResize = New System.Windows.Forms.Button()
        Me.chkQCShowExcludedBatch = New System.Windows.Forms.CheckBox()
        Me.lblColoredRows = New System.Windows.Forms.Label()
        Me.cmdOrderReportTableConfig = New System.Windows.Forms.Button()
        Me.chkTableGraphicExamples = New System.Windows.Forms.CheckBox()
        Me.dgvReportTableConfiguration = New System.Windows.Forms.DataGridView()
        Me.gbxlblConfigureReportTables1 = New System.Windows.Forms.GroupBox()
        Me.lblRTC = New System.Windows.Forms.Label()
        Me.tp7 = New System.Windows.Forms.TabPage()
        Me.lbldgvReportTableHeaderConfig = New System.Windows.Forms.Label()
        Me.lblNotes2 = New System.Windows.Forms.Label()
        Me.lblNotes1 = New System.Windows.Forms.Label()
        Me.gbxlblConfigureColumnHeadings1 = New System.Windows.Forms.GroupBox()
        Me.lblTCH_01 = New System.Windows.Forms.Label()
        Me.dgvReportTableHeaderConfig = New System.Windows.Forms.DataGridView()
        Me.lbldgvReportTables = New System.Windows.Forms.Label()
        Me.dgvReportTables = New System.Windows.Forms.DataGridView()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.tp8 = New System.Windows.Forms.TabPage()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.gbxlblAnalyticalReferenceStd = New System.Windows.Forms.GroupBox()
        Me.lblARS = New System.Windows.Forms.Label()
        Me.dgvCompanyAnalRef = New System.Windows.Forms.DataGridView()
        Me.dgvWatsonAnalRef = New System.Windows.Forms.DataGridView()
        Me.lblWAR = New System.Windows.Forms.Label()
        Me.lblARST = New System.Windows.Forms.Label()
        Me.tp9 = New System.Windows.Forms.TabPage()
        Me.lblGlobalConfiguration = New System.Windows.Forms.Label()
        Me.gbxlblAddEditContributors = New System.Windows.Forms.GroupBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label53 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.dgvContributingPersonnel = New System.Windows.Forms.DataGridView()
        Me.tp10 = New System.Windows.Forms.TabPage()
        Me.gbxlblReviewValidatedMethod = New System.Windows.Forms.GroupBox()
        Me.lbl1 = New System.Windows.Forms.Label()
        Me.lbl2 = New System.Windows.Forms.Label()
        Me.dgvMethodValData = New System.Windows.Forms.DataGridView()
        Me.gbMethValApplyGuWu = New System.Windows.Forms.GroupBox()
        Me.dgvMethValExistingGuWu = New System.Windows.Forms.DataGridView()
        Me.cbxArchivedMDB = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmdMethValExecute = New System.Windows.Forms.Button()
        Me.cbxMethValExistingGuWu = New System.Windows.Forms.ComboBox()
        Me.lblM2 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.cmdBrowseMDB = New System.Windows.Forms.Button()
        Me.gbMethodValMultiple = New System.Windows.Forms.GroupBox()
        Me.lblMethValMultiple = New System.Windows.Forms.Label()
        Me.txtMethValMultiple = New System.Windows.Forms.TextBox()
        Me.rbMethValMultiple = New System.Windows.Forms.RadioButton()
        Me.rbMethValAnalyte = New System.Windows.Forms.RadioButton()
        Me.chkMethodValMultiple = New System.Windows.Forms.CheckBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.tp11 = New System.Windows.Forms.TabPage()
        Me.chkQAEventBorder = New System.Windows.Forms.CheckBox()
        Me.lblQAHyperlink = New System.Windows.Forms.LinkLabel()
        Me.dgQATable = New System.Windows.Forms.DataGrid()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.tp12 = New System.Windows.Forms.TabPage()
        Me.gbxlblSampleReceiptRecords2 = New System.Windows.Forms.GroupBox()
        Me.Label48 = New System.Windows.Forms.Label()
        Me.gbxlblSampleReceiptRecords1 = New System.Windows.Forms.GroupBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.txtSRecTotalReportWatson = New System.Windows.Forms.TextBox()
        Me.chkUseWatsonSampleNumber = New System.Windows.Forms.CheckBox()
        Me.chkManualSampleNumber = New System.Windows.Forms.CheckBox()
        Me.Label46 = New System.Windows.Forms.Label()
        Me.dgvSampleReceiptWatson = New System.Windows.Forms.DataGridView()
        Me.Label41 = New System.Windows.Forms.Label()
        Me.txtSRecTotalReport = New System.Windows.Forms.TextBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.txtSRecTotal = New System.Windows.Forms.TextBox()
        Me.dgvSampleReceipt = New System.Windows.Forms.DataGridView()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.tp13 = New System.Windows.Forms.TabPage()
        Me.pbxWord = New System.Windows.Forms.PictureBox()
        Me.cmdAppFig = New System.Windows.Forms.Button()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.tp14 = New System.Windows.Forms.TabPage()
        Me.cmdAdministration = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tp15 = New System.Windows.Forms.TabPage()
        Me.cmdAnalDetails = New System.Windows.Forms.Button()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tp16 = New System.Windows.Forms.TabPage()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lblAT = New System.Windows.Forms.Label()
        Me.cmdAuditTrail = New System.Windows.Forms.Button()
        Me.lblAuditTrail = New System.Windows.Forms.Label()
        Me.gbFilters = New System.Windows.Forms.GroupBox()
        Me.txtFilterSamples = New System.Windows.Forms.TextBox()
        Me.cmdClearFilters = New System.Windows.Forms.Button()
        Me.cmdCreateReportTitle2 = New System.Windows.Forms.Button()
        Me.gbSource = New System.Windows.Forms.GroupBox()
        Me.rbOracle = New System.Windows.Forms.RadioButton()
        Me.txtArchivePath = New System.Windows.Forms.TextBox()
        Me.rbArchive = New System.Windows.Forms.RadioButton()
        Me.panFilterStudy = New System.Windows.Forms.Panel()
        Me.cmdFilterStudy = New System.Windows.Forms.Button()
        Me.txtFilterStudy = New System.Windows.Forms.TextBox()
        Me.cbxFilterStudy = New System.Windows.Forms.ComboBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.gbxMultVal = New System.Windows.Forms.GroupBox()
        Me.rbMultValNo = New System.Windows.Forms.RadioButton()
        Me.rbMultValYes = New System.Windows.Forms.RadioButton()
        Me.grbShowSummaryTable = New System.Windows.Forms.GroupBox()
        Me.rbShowIncludedSummaryTable = New System.Windows.Forms.RadioButton()
        Me.rbShowAllSummaryTable = New System.Windows.Forms.RadioButton()
        Me.gbRTC = New System.Windows.Forms.GroupBox()
        Me.rbShowAllRTConfig = New System.Windows.Forms.RadioButton()
        Me.rbShowIncludedRTConfig = New System.Windows.Forms.RadioButton()
        Me.cmdHomeCancel = New System.Windows.Forms.Button()
        Me.cmdDataCancel = New System.Windows.Forms.Button()
        Me.cmdViewAnalyticalRuns1 = New System.Windows.Forms.Button()
        Me.cmdAnaRunSumCancel = New System.Windows.Forms.Button()
        Me.cmdUpdateSummaryInfo = New System.Windows.Forms.Button()
        Me.cmdResetSummaryTable = New System.Windows.Forms.Button()
        Me.cmdRefreshStatements = New System.Windows.Forms.Button()
        Me.cmdOpenReportStatements = New System.Windows.Forms.Button()
        Me.cmdCancelReportStatements = New System.Windows.Forms.Button()
        Me.cmdCreateTable = New System.Windows.Forms.Button()
        Me.cmdImportTables = New System.Windows.Forms.Button()
        Me.cmdSelect = New System.Windows.Forms.Button()
        Me.cmdOutliers = New System.Windows.Forms.Button()
        Me.cmdAdvancedTable = New System.Windows.Forms.Button()
        Me.cmdDuplicateTables = New System.Windows.Forms.Button()
        Me.cmdViewAnalRuns = New System.Windows.Forms.Button()
        Me.cmdAssignSamples = New System.Windows.Forms.Button()
        Me.cmdRTConfigCancel = New System.Windows.Forms.Button()
        Me.cmdRTHeaderConfigCancel = New System.Windows.Forms.Button()
        Me.cmdAddAnalyte = New System.Windows.Forms.Button()
        Me.cmdCopyRepAnalyte = New System.Windows.Forms.Button()
        Me.cmdDeleteRepAnalyte = New System.Windows.Forms.Button()
        Me.cmdAnalRefCancel = New System.Windows.Forms.Button()
        Me.cmdAddRepAnalyte = New System.Windows.Forms.Button()
        Me.cmdReplacePersonnel = New System.Windows.Forms.Button()
        Me.cmdCPDelete = New System.Windows.Forms.Button()
        Me.cmdCPAdd = New System.Windows.Forms.Button()
        Me.cmdCPCancel = New System.Windows.Forms.Button()
        Me.cmdMethValUpdate = New System.Windows.Forms.Button()
        Me.cmdMethValReset = New System.Windows.Forms.Button()
        Me.cmdQACancel = New System.Windows.Forms.Button()
        Me.cmdDeleteQAEvent = New System.Windows.Forms.Button()
        Me.cmdInsertQAEvent = New System.Windows.Forms.Button()
        Me.cmdSRecCancel = New System.Windows.Forms.Button()
        Me.cmdDeletSRec = New System.Windows.Forms.Button()
        Me.cmdInsertSRec = New System.Windows.Forms.Button()
        Me.cmdReportHistory = New System.Windows.Forms.Button()
        Me.cmdShowOutstanding = New System.Windows.Forms.Button()
        Me.cmdApplyTemplate = New System.Windows.Forms.Button()
        Me.cmdUpdateProject = New System.Windows.Forms.Button()
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.pb2 = New System.Windows.Forms.ProgressBar()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.llblAssignedSamples = New System.Windows.Forms.LinkLabel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.lbxTab1 = New System.Windows.Forms.ListBox()
        Me.lblTOC = New System.Windows.Forms.Label()
        Me.lblWatsonStudy = New System.Windows.Forms.Label()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lblReportTitle = New System.Windows.Forms.Label()
        Me.cmdEdit = New System.Windows.Forms.Button()
        Me.cmdSave = New System.Windows.Forms.Button()
        Me.OleDbSelectCommand32 = New System.Data.OleDb.OleDbCommand()
        Me.cbxStudy = New System.Windows.Forms.ComboBox()
        Me.cmdRefresh = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.dgvUser = New System.Windows.Forms.DataGridView()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.TimerRTC = New System.Windows.Forms.Timer(Me.components)
        Me.cbxFilter = New System.Windows.Forms.ComboBox()
        Me.TimerSplash = New System.Windows.Forms.Timer(Me.components)
        Me.cmdHook = New System.Windows.Forms.Button()
        Me.cbxExampleReport = New System.Windows.Forms.ComboBox()
        Me.txtcbxMDBSelIndex = New System.Windows.Forms.TextBox()
        Me.txtFilterIndex = New System.Windows.Forms.TextBox()
        Me.ms1 = New System.Windows.Forms.MenuStrip()
        Me.mnuAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMenuAbout = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuShowFC = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuMenuGenFC = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuTroubleshooting = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.dtp1 = New System.Windows.Forms.DateTimePicker()
        Me.cmdSymbol = New System.Windows.Forms.Button()
        Me.panEdit = New System.Windows.Forms.Panel()
        Me.panChoose = New System.Windows.Forms.Panel()
        Me.cmdClearStudy = New System.Windows.Forms.Button()
        Me.panWatsonData = New System.Windows.Forms.Panel()
        Me.lblWatsonData = New System.Windows.Forms.Label()
        Me.lblWatsonDataTitle = New System.Windows.Forms.Label()
        Me.panStudyFilter = New System.Windows.Forms.Panel()
        Me.gbStudyFilter = New System.Windows.Forms.GroupBox()
        Me.optStudyDocStudies = New System.Windows.Forms.RadioButton()
        Me.optStudyDocClosed = New System.Windows.Forms.RadioButton()
        Me.optStudyDocOpen = New System.Windows.Forms.RadioButton()
        Me.panSampleRec = New System.Windows.Forms.Panel()
        Me.panQAEvent = New System.Windows.Forms.Panel()
        Me.panMethVal = New System.Windows.Forms.Panel()
        Me.panContr = New System.Windows.Forms.Panel()
        Me.panWordTemp = New System.Windows.Forms.Panel()
        Me.panRepTables = New System.Windows.Forms.Panel()
        Me.cmdApplyTables = New System.Windows.Forms.Button()
        Me.panSumTable = New System.Windows.Forms.Panel()
        Me.panAnalRuns = New System.Windows.Forms.Panel()
        Me.panColHeadings = New System.Windows.Forms.Panel()
        Me.panAnalRefStds = New System.Windows.Forms.Panel()
        Me.panTopLevel = New System.Windows.Forms.Panel()
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.TimerWarning = New System.Windows.Forms.Timer(Me.components)
        Me.lblWatsonWarning = New System.Windows.Forms.Label()
        Me.panDot = New System.Windows.Forms.Panel()
        Me.lblBlack = New System.Windows.Forms.Label()
        Me.mCal1 = New System.Windows.Forms.MonthCalendar()
        Me.cmdEnterCal = New System.Windows.Forms.Button()
        Me.cmdCalCancel = New System.Windows.Forms.Button()
        Me.panCal = New System.Windows.Forms.Panel()
        Me.panPrepareReportInside = New System.Windows.Forms.Panel()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MenuPrepareReport = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrepareEntireReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrepareOnlySelectedTableToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrepareOnlyReportBodyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrepareOnlyReportTablesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.panPrepareReportOutside = New System.Windows.Forms.Panel()
        Me.panActions = New System.Windows.Forms.Panel()
        Me.dgvFieldCodes = New System.Windows.Forms.DataGridView()
        Me.lblActions = New System.Windows.Forms.Label()
        Me.panProgress = New System.Windows.Forms.Panel()
        Me.tab1.SuspendLayout()
        Me.tp1.SuspendLayout()
        CType(Me.dgvReports, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panLockFinalReport.SuspendLayout()
        CType(Me.dgStudies, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgHome, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvwStudy, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp2.SuspendLayout()
        CType(Me.dgvDataWatson, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabData.SuspendLayout()
        Me.tabData1.SuspendLayout()
        CType(Me.dgvDataCompany, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabData2.SuspendLayout()
        CType(Me.dgvStudyConfig, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabData3.SuspendLayout()
        Me.gbInclude.SuspendLayout()
        CType(Me.dgvFC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabData4.SuspendLayout()
        Me.gbMeanComp.SuspendLayout()
        Me.gbCritPrecision.SuspendLayout()
        Me.gbRound5.SuspendLayout()
        Me.tabData5.SuspendLayout()
        CType(Me.dgvAnalyteGroups, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp3.SuspendLayout()
        CType(Me.dgvAnalyticalRunSummary, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbxlblReviewAnalyticalRuns.SuspendLayout()
        Me.gbReportOptions.SuspendLayout()
        Me.panAnalRunSum.SuspendLayout()
        Me.panAnalRunChoices.SuspendLayout()
        Me.tp4.SuspendLayout()
        CType(Me.dgvSummaryData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbxlblMethodValidation.SuspendLayout()
        Me.tp5.SuspendLayout()
        Me.gbxlblChooseEditWordTemplate.SuspendLayout()
        Me.panSections.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.gbSectionStyle.SuspendLayout()
        Me.panRBSwb.SuspendLayout()
        Me.cmsHome.SuspendLayout()
        Me.grpRBS.SuspendLayout()
        CType(Me.dgvReportStatementWord, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp6.SuspendLayout()
        CType(Me.dgvGroups, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panTableGraphicExamples.SuspendLayout()
        CType(Me.pbxTableGraphicExamples, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvReportTableConfiguration, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbxlblConfigureReportTables1.SuspendLayout()
        Me.tp7.SuspendLayout()
        Me.gbxlblConfigureColumnHeadings1.SuspendLayout()
        CType(Me.dgvReportTableHeaderConfig, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvReportTables, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp8.SuspendLayout()
        Me.gbxlblAnalyticalReferenceStd.SuspendLayout()
        CType(Me.dgvCompanyAnalRef, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvWatsonAnalRef, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp9.SuspendLayout()
        Me.gbxlblAddEditContributors.SuspendLayout()
        CType(Me.dgvContributingPersonnel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp10.SuspendLayout()
        Me.gbxlblReviewValidatedMethod.SuspendLayout()
        CType(Me.dgvMethodValData, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbMethValApplyGuWu.SuspendLayout()
        CType(Me.dgvMethValExistingGuWu, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.gbMethodValMultiple.SuspendLayout()
        Me.tp11.SuspendLayout()
        CType(Me.dgQATable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp12.SuspendLayout()
        Me.gbxlblSampleReceiptRecords2.SuspendLayout()
        Me.gbxlblSampleReceiptRecords1.SuspendLayout()
        CType(Me.dgvSampleReceiptWatson, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSampleReceipt, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp13.SuspendLayout()
        CType(Me.pbxWord, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tp14.SuspendLayout()
        Me.tp15.SuspendLayout()
        Me.tp16.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.gbFilters.SuspendLayout()
        Me.gbSource.SuspendLayout()
        Me.panFilterStudy.SuspendLayout()
        Me.gbxMultVal.SuspendLayout()
        Me.grbShowSummaryTable.SuspendLayout()
        Me.gbRTC.SuspendLayout()
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ms1.SuspendLayout()
        Me.panEdit.SuspendLayout()
        Me.panChoose.SuspendLayout()
        Me.panWatsonData.SuspendLayout()
        Me.panStudyFilter.SuspendLayout()
        Me.gbStudyFilter.SuspendLayout()
        Me.panSampleRec.SuspendLayout()
        Me.panQAEvent.SuspendLayout()
        Me.panMethVal.SuspendLayout()
        Me.panContr.SuspendLayout()
        Me.panWordTemp.SuspendLayout()
        Me.panRepTables.SuspendLayout()
        Me.panSumTable.SuspendLayout()
        Me.panAnalRuns.SuspendLayout()
        Me.panColHeadings.SuspendLayout()
        Me.panAnalRefStds.SuspendLayout()
        Me.panTopLevel.SuspendLayout()
        Me.panCal.SuspendLayout()
        Me.panPrepareReportInside.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.panPrepareReportOutside.SuspendLayout()
        Me.panActions.SuspendLayout()
        CType(Me.dgvFieldCodes, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panProgress.SuspendLayout()
        Me.SuspendLayout()
        '
        'tab1
        '
        Me.tab1.CausesValidation = False
        Me.tab1.Controls.Add(Me.tp1)
        Me.tab1.Controls.Add(Me.tp2)
        Me.tab1.Controls.Add(Me.tp3)
        Me.tab1.Controls.Add(Me.tp4)
        Me.tab1.Controls.Add(Me.tp5)
        Me.tab1.Controls.Add(Me.tp6)
        Me.tab1.Controls.Add(Me.tp7)
        Me.tab1.Controls.Add(Me.tp8)
        Me.tab1.Controls.Add(Me.tp9)
        Me.tab1.Controls.Add(Me.tp10)
        Me.tab1.Controls.Add(Me.tp11)
        Me.tab1.Controls.Add(Me.tp12)
        Me.tab1.Controls.Add(Me.tp13)
        Me.tab1.Controls.Add(Me.tp14)
        Me.tab1.Controls.Add(Me.tp15)
        Me.tab1.Controls.Add(Me.tp16)
        Me.tab1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tab1.ItemSize = New System.Drawing.Size(50, 20)
        Me.tab1.Location = New System.Drawing.Point(356, 97)
        Me.tab1.Name = "tab1"
        Me.tab1.Padding = New System.Drawing.Point(0, 0)
        Me.tab1.SelectedIndex = 0
        Me.tab1.Size = New System.Drawing.Size(916, 631)
        Me.tab1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
        Me.tab1.TabIndex = 2
        '
        'tp1
        '
        Me.tp1.AutoScroll = True
        Me.tp1.BackColor = System.Drawing.Color.Ivory
        Me.tp1.Controls.Add(Me.dgvReports)
        Me.tp1.Controls.Add(Me.panLockFinalReport)
        Me.tp1.Controls.Add(Me.cmdHeader)
        Me.tp1.Controls.Add(Me.dgStudies)
        Me.tp1.Controls.Add(Me.dgHome)
        Me.tp1.Controls.Add(Me.cmdConfigureReport)
        Me.tp1.Controls.Add(Me.dgvwStudy)
        Me.tp1.Controls.Add(Me.lblSelectCell)
        Me.tp1.Controls.Add(Me.Label14)
        Me.tp1.Controls.Add(Me.cmdCreateReportTitle)
        Me.tp1.Controls.Add(Me.lblHome)
        Me.tp1.Controls.Add(Me.Label2)
        Me.tp1.Controls.Add(Me.Label1)
        Me.tp1.Controls.Add(Me.lblcbxStudies)
        Me.tp1.Location = New System.Drawing.Point(4, 24)
        Me.tp1.Margin = New System.Windows.Forms.Padding(0)
        Me.tp1.Name = "tp1"
        Me.tp1.Size = New System.Drawing.Size(908, 603)
        Me.tp1.TabIndex = 0
        Me.tp1.Text = "1"
        '
        'dgvReports
        '
        Me.dgvReports.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvReports.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvReports.BackgroundColor = System.Drawing.Color.White
        Me.dgvReports.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.dgvReports.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvReports.Location = New System.Drawing.Point(5, 439)
        Me.dgvReports.MultiSelect = False
        Me.dgvReports.Name = "dgvReports"
        Me.dgvReports.ReadOnly = True
        Me.dgvReports.Size = New System.Drawing.Size(901, 158)
        Me.dgvReports.TabIndex = 92
        '
        'panLockFinalReport
        '
        Me.panLockFinalReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.panLockFinalReport.Controls.Add(Me.chkLockFinalReport)
        Me.panLockFinalReport.Controls.Add(Me.lblFinalReportLockedDate)
        Me.panLockFinalReport.Controls.Add(Me.lblFinalReportLockedDateLabel)
        Me.panLockFinalReport.Location = New System.Drawing.Point(153, 416)
        Me.panLockFinalReport.Name = "panLockFinalReport"
        Me.panLockFinalReport.Size = New System.Drawing.Size(502, 21)
        Me.panLockFinalReport.TabIndex = 122
        '
        'chkLockFinalReport
        '
        Me.chkLockFinalReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkLockFinalReport.AutoSize = True
        Me.chkLockFinalReport.Location = New System.Drawing.Point(3, 2)
        Me.chkLockFinalReport.Name = "chkLockFinalReport"
        Me.chkLockFinalReport.Size = New System.Drawing.Size(127, 21)
        Me.chkLockFinalReport.TabIndex = 119
        Me.chkLockFinalReport.Text = "Lock Final Report"
        Me.chkLockFinalReport.UseVisualStyleBackColor = True
        '
        'lblFinalReportLockedDate
        '
        Me.lblFinalReportLockedDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblFinalReportLockedDate.AutoSize = True
        Me.lblFinalReportLockedDate.Location = New System.Drawing.Point(292, 2)
        Me.lblFinalReportLockedDate.Name = "lblFinalReportLockedDate"
        Me.lblFinalReportLockedDate.Size = New System.Drawing.Size(26, 17)
        Me.lblFinalReportLockedDate.TabIndex = 121
        Me.lblFinalReportLockedDate.Text = "NA"
        '
        'lblFinalReportLockedDateLabel
        '
        Me.lblFinalReportLockedDateLabel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblFinalReportLockedDateLabel.AutoSize = True
        Me.lblFinalReportLockedDateLabel.Location = New System.Drawing.Point(138, 2)
        Me.lblFinalReportLockedDateLabel.Name = "lblFinalReportLockedDateLabel"
        Me.lblFinalReportLockedDateLabel.Size = New System.Drawing.Size(155, 17)
        Me.lblFinalReportLockedDateLabel.TabIndex = 120
        Me.lblFinalReportLockedDateLabel.Text = "Date of Last Final Report:"
        '
        'cmdHeader
        '
        Me.cmdHeader.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdHeader.Enabled = False
        Me.cmdHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdHeader.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdHeader.Location = New System.Drawing.Point(502, 373)
        Me.cmdHeader.Name = "cmdHeader"
        Me.cmdHeader.Size = New System.Drawing.Size(135, 38)
        Me.cmdHeader.TabIndex = 118
        Me.cmdHeader.Text = "Configure &Header/ Footer (Optional).."
        Me.cmdHeader.UseVisualStyleBackColor = False
        Me.cmdHeader.Visible = False
        '
        'dgStudies
        '
        Me.dgStudies.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.dgStudies.DataMember = ""
        Me.dgStudies.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgStudies.Location = New System.Drawing.Point(700, 13)
        Me.dgStudies.Name = "dgStudies"
        Me.dgStudies.Size = New System.Drawing.Size(92, 50)
        Me.dgStudies.TabIndex = 34
        Me.dgStudies.Visible = False
        '
        'dgHome
        '
        Me.dgHome.DataMember = ""
        Me.dgHome.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgHome.Location = New System.Drawing.Point(675, 34)
        Me.dgHome.Name = "dgHome"
        Me.dgHome.PreferredColumnWidth = 150
        Me.dgHome.Size = New System.Drawing.Size(68, 45)
        Me.dgHome.TabIndex = 27
        Me.dgHome.Visible = False
        '
        'cmdConfigureReport
        '
        Me.cmdConfigureReport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdConfigureReport.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdConfigureReport.Enabled = False
        Me.cmdConfigureReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdConfigureReport.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdConfigureReport.Location = New System.Drawing.Point(661, 398)
        Me.cmdConfigureReport.Name = "cmdConfigureReport"
        Me.cmdConfigureReport.Size = New System.Drawing.Size(99, 38)
        Me.cmdConfigureReport.TabIndex = 33
        Me.cmdConfigureReport.Text = "&Add Report..."
        Me.cmdConfigureReport.UseVisualStyleBackColor = False
        Me.cmdConfigureReport.Visible = False
        '
        'dgvwStudy
        '
        Me.dgvwStudy.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvwStudy.BackgroundColor = System.Drawing.Color.White
        Me.dgvwStudy.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.dgvwStudy.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle162.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle162.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle162.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle162.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle162.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle162.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle162.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle162.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvwStudy.DefaultCellStyle = DataGridViewCellStyle162
        Me.dgvwStudy.Location = New System.Drawing.Point(5, 50)
        Me.dgvwStudy.MultiSelect = False
        Me.dgvwStudy.Name = "dgvwStudy"
        Me.dgvwStudy.ReadOnly = True
        Me.dgvwStudy.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvwStudy.Size = New System.Drawing.Size(901, 345)
        Me.dgvwStudy.TabIndex = 1
        Me.dgvwStudy.TabStop = False
        '
        'lblSelectCell
        '
        Me.lblSelectCell.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSelectCell.BackColor = System.Drawing.Color.Transparent
        Me.lblSelectCell.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectCell.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblSelectCell.Location = New System.Drawing.Point(750, 390)
        Me.lblSelectCell.MaximumSize = New System.Drawing.Size(175, 0)
        Me.lblSelectCell.MinimumSize = New System.Drawing.Size(0, 46)
        Me.lblSelectCell.Name = "lblSelectCell"
        Me.lblSelectCell.Size = New System.Drawing.Size(156, 46)
        Me.lblSelectCell.TabIndex = 94
        Me.lblSelectCell.Text = "* Select cell, then choose " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "   from dropdown box"
        Me.lblSelectCell.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label14
        '
        Me.Label14.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = True
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(5, 421)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(124, 16)
        Me.Label14.TabIndex = 93
        Me.Label14.Text = "Configured Report"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'cmdCreateReportTitle
        '
        Me.cmdCreateReportTitle.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCreateReportTitle.Enabled = False
        Me.cmdCreateReportTitle.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCreateReportTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdCreateReportTitle.Location = New System.Drawing.Point(5, 0)
        Me.cmdCreateReportTitle.Name = "cmdCreateReportTitle"
        Me.cmdCreateReportTitle.Size = New System.Drawing.Size(99, 38)
        Me.cmdCreateReportTitle.TabIndex = 37
        Me.cmdCreateReportTitle.Text = "&Create Report Title"
        Me.cmdCreateReportTitle.UseVisualStyleBackColor = False
        Me.cmdCreateReportTitle.Visible = False
        '
        'lblHome
        '
        Me.lblHome.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblHome.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblHome.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHome.ForeColor = System.Drawing.Color.White
        Me.lblHome.Location = New System.Drawing.Point(0, 0)
        Me.lblHome.Name = "lblHome"
        Me.lblHome.Size = New System.Drawing.Size(908, 21)
        Me.lblHome.TabIndex = 21
        Me.lblHome.Text = "Choose Study && Report"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Location = New System.Drawing.Point(455, 377)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(218, 20)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Date of Last Project Update:"
        Me.Label2.Visible = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(749, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(103, 45)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Report Generation History"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.Label1.Visible = False
        '
        'lblcbxStudies
        '
        Me.lblcbxStudies.AutoSize = True
        Me.lblcbxStudies.BackColor = System.Drawing.Color.Transparent
        Me.lblcbxStudies.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblcbxStudies.Location = New System.Drawing.Point(5, 32)
        Me.lblcbxStudies.Name = "lblcbxStudies"
        Me.lblcbxStudies.Size = New System.Drawing.Size(260, 16)
        Me.lblcbxStudies.TabIndex = 11
        Me.lblcbxStudies.Text = "Watson Studies Configured in StudyDoc"
        '
        'tp2
        '
        Me.tp2.AutoScroll = True
        Me.tp2.BackColor = System.Drawing.Color.Ivory
        Me.tp2.Controls.Add(Me.dgvDataWatson)
        Me.tp2.Controls.Add(Me.cbxAssayTechnique)
        Me.tp2.Controls.Add(Me.Label30)
        Me.tp2.Controls.Add(Me.tabData)
        Me.tp2.Controls.Add(Me.cbxSubmittedTo)
        Me.tp2.Controls.Add(Me.lblDCP)
        Me.tp2.Controls.Add(Me.cbxSampleStorageTemp)
        Me.tp2.Controls.Add(Me.Label5)
        Me.tp2.Controls.Add(Me.lblIncludeInTitle)
        Me.tp2.Controls.Add(Me.Label44)
        Me.tp2.Controls.Add(Me.txtSubmittedTo)
        Me.tp2.Controls.Add(Me.Label43)
        Me.tp2.Controls.Add(Me.Label42)
        Me.tp2.Controls.Add(Me.txtSubmittedBy)
        Me.tp2.Controls.Add(Me.txtInSupportOf)
        Me.tp2.Controls.Add(Me.cbxSubmittedBy)
        Me.tp2.Controls.Add(Me.cbxInSupportOf)
        Me.tp2.Controls.Add(Me.cbxSampleSizeUnits)
        Me.tp2.Controls.Add(Me.cbxAnticoagulant)
        Me.tp2.Controls.Add(Me.cbxAssayTechniqueAcronym)
        Me.tp2.Controls.Add(Me.Label29)
        Me.tp2.Controls.Add(Me.Label28)
        Me.tp2.Controls.Add(Me.Label27)
        Me.tp2.Controls.Add(Me.Label7)
        Me.tp2.Controls.Add(Me.lblData)
        Me.tp2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tp2.Location = New System.Drawing.Point(4, 24)
        Me.tp2.Name = "tp2"
        Me.tp2.Size = New System.Drawing.Size(908, 603)
        Me.tp2.TabIndex = 1
        Me.tp2.Text = "2"
        '
        'dgvDataWatson
        '
        Me.dgvDataWatson.AllowUserToAddRows = False
        Me.dgvDataWatson.AllowUserToDeleteRows = False
        Me.dgvDataWatson.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDataWatson.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle161.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle161.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle161.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle161.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle161.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle161.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle161.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDataWatson.DefaultCellStyle = DataGridViewCellStyle161
        Me.dgvDataWatson.Location = New System.Drawing.Point(548, 360)
        Me.dgvDataWatson.Name = "dgvDataWatson"
        Me.dgvDataWatson.Size = New System.Drawing.Size(353, 231)
        Me.dgvDataWatson.TabIndex = 98
        '
        'cbxAssayTechnique
        '
        Me.cbxAssayTechnique.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxAssayTechnique.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxAssayTechnique.Enabled = False
        Me.cbxAssayTechnique.IntegralHeight = False
        Me.cbxAssayTechnique.Location = New System.Drawing.Point(177, 26)
        Me.cbxAssayTechnique.Name = "cbxAssayTechnique"
        Me.cbxAssayTechnique.Size = New System.Drawing.Size(348, 25)
        Me.cbxAssayTechnique.TabIndex = 44
        '
        'Label30
        '
        Me.Label30.AutoSize = True
        Me.Label30.BackColor = System.Drawing.Color.Transparent
        Me.Label30.Location = New System.Drawing.Point(9, 26)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(107, 17)
        Me.Label30.TabIndex = 34
        Me.Label30.Text = "Assay Technique:"
        '
        'tabData
        '
        Me.tabData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabData.Controls.Add(Me.tabData1)
        Me.tabData.Controls.Add(Me.tabData2)
        Me.tabData.Controls.Add(Me.tabData3)
        Me.tabData.Controls.Add(Me.tabData4)
        Me.tabData.Controls.Add(Me.tabData5)
        Me.tabData.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
        Me.tabData.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabData.ItemSize = New System.Drawing.Size(90, 30)
        Me.tabData.Location = New System.Drawing.Point(9, 153)
        Me.tabData.Name = "tabData"
        Me.tabData.SelectedIndex = 0
        Me.tabData.Size = New System.Drawing.Size(516, 439)
        Me.tabData.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
        Me.tabData.TabIndex = 97
        '
        'tabData1
        '
        Me.tabData1.Controls.Add(Me.dgvDataCompany)
        Me.tabData1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.tabData1.Location = New System.Drawing.Point(4, 34)
        Me.tabData1.Name = "tabData1"
        Me.tabData1.Padding = New System.Windows.Forms.Padding(3)
        Me.tabData1.Size = New System.Drawing.Size(508, 401)
        Me.tabData1.TabIndex = 0
        Me.tabData1.Text = "Study" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Information"
        Me.tabData1.UseVisualStyleBackColor = True
        '
        'dgvDataCompany
        '
        Me.dgvDataCompany.AllowUserToAddRows = False
        Me.dgvDataCompany.AllowUserToDeleteRows = False
        Me.dgvDataCompany.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvDataCompany.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader
        Me.dgvDataCompany.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvDataCompany.BackgroundColor = System.Drawing.Color.White
        Me.dgvDataCompany.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle163.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle163.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle163.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle163.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle163.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle163.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle163.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDataCompany.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle163
        Me.dgvDataCompany.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle164.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle164.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle164.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle164.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle164.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle164.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle164.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle164.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvDataCompany.DefaultCellStyle = DataGridViewCellStyle164
        Me.dgvDataCompany.Location = New System.Drawing.Point(3, 3)
        Me.dgvDataCompany.Name = "dgvDataCompany"
        Me.dgvDataCompany.ReadOnly = True
        Me.dgvDataCompany.RowHeadersWidth = 25
        Me.dgvDataCompany.Size = New System.Drawing.Size(499, 395)
        Me.dgvDataCompany.TabIndex = 95
        '
        'tabData2
        '
        Me.tabData2.Controls.Add(Me.dgvStudyConfig)
        Me.tabData2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.tabData2.Location = New System.Drawing.Point(4, 34)
        Me.tabData2.Name = "tabData2"
        Me.tabData2.Padding = New System.Windows.Forms.Padding(3)
        Me.tabData2.Size = New System.Drawing.Size(508, 401)
        Me.tabData2.TabIndex = 1
        Me.tabData2.Text = "Study " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Configuration"
        Me.tabData2.UseVisualStyleBackColor = True
        '
        'dgvStudyConfig
        '
        Me.dgvStudyConfig.AllowUserToAddRows = False
        Me.dgvStudyConfig.AllowUserToDeleteRows = False
        Me.dgvStudyConfig.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvStudyConfig.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvStudyConfig.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvStudyConfig.BackgroundColor = System.Drawing.Color.White
        Me.dgvStudyConfig.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle165.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle165.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle165.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle165.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle165.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle165.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle165.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvStudyConfig.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle165
        Me.dgvStudyConfig.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle166.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle166.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle166.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle166.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle166.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle166.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle166.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle166.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvStudyConfig.DefaultCellStyle = DataGridViewCellStyle166
        Me.dgvStudyConfig.Location = New System.Drawing.Point(3, 3)
        Me.dgvStudyConfig.Name = "dgvStudyConfig"
        Me.dgvStudyConfig.ReadOnly = True
        Me.dgvStudyConfig.RowHeadersWidth = 25
        Me.dgvStudyConfig.Size = New System.Drawing.Size(499, 395)
        Me.dgvStudyConfig.TabIndex = 96
        '
        'tabData3
        '
        Me.tabData3.BackColor = System.Drawing.Color.Ivory
        Me.tabData3.Controls.Add(Me.lblCFAuditTrail)
        Me.tabData3.Controls.Add(Me.Label12)
        Me.tabData3.Controls.Add(Me.cmdDownCF)
        Me.tabData3.Controls.Add(Me.cmdUpCF)
        Me.tabData3.Controls.Add(Me.gbInclude)
        Me.tabData3.Controls.Add(Me.dgvFC)
        Me.tabData3.Location = New System.Drawing.Point(4, 34)
        Me.tabData3.Name = "tabData3"
        Me.tabData3.Size = New System.Drawing.Size(508, 401)
        Me.tabData3.TabIndex = 2
        Me.tabData3.Text = "Custom Field" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Codes"
        '
        'lblCFAuditTrail
        '
        Me.lblCFAuditTrail.AutoSize = True
        Me.lblCFAuditTrail.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCFAuditTrail.Location = New System.Drawing.Point(166, 40)
        Me.lblCFAuditTrail.Name = "lblCFAuditTrail"
        Me.lblCFAuditTrail.Size = New System.Drawing.Size(299, 17)
        Me.lblCFAuditTrail.TabIndex = 151
        Me.lblCFAuditTrail.Text = "Field Code ordering not recorded in Audit Trail"
        Me.lblCFAuditTrail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.White
        Me.Label12.Location = New System.Drawing.Point(8, 63)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(55, 37)
        Me.Label12.TabIndex = 150
        Me.Label12.Text = "Move" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Row"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdDownCF
        '
        Me.cmdDownCF.BackColor = System.Drawing.Color.Transparent
        Me.cmdDownCF.BackgroundImage = CType(resources.GetObject("cmdDownCF.BackgroundImage"), System.Drawing.Image)
        Me.cmdDownCF.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdDownCF.Enabled = False
        Me.cmdDownCF.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cmdDownCF.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdDownCF.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdDownCF.Location = New System.Drawing.Point(7, 149)
        Me.cmdDownCF.Name = "cmdDownCF"
        Me.cmdDownCF.Size = New System.Drawing.Size(55, 40)
        Me.cmdDownCF.TabIndex = 149
        Me.cmdDownCF.UseVisualStyleBackColor = False
        '
        'cmdUpCF
        '
        Me.cmdUpCF.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdUpCF.BackgroundImage = CType(resources.GetObject("cmdUpCF.BackgroundImage"), System.Drawing.Image)
        Me.cmdUpCF.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdUpCF.Enabled = False
        Me.cmdUpCF.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cmdUpCF.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpCF.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdUpCF.Location = New System.Drawing.Point(7, 103)
        Me.cmdUpCF.Name = "cmdUpCF"
        Me.cmdUpCF.Size = New System.Drawing.Size(55, 40)
        Me.cmdUpCF.TabIndex = 148
        Me.cmdUpCF.UseVisualStyleBackColor = False
        '
        'gbInclude
        '
        Me.gbInclude.Controls.Add(Me.rbAll)
        Me.gbInclude.Controls.Add(Me.rbInclude)
        Me.gbInclude.Location = New System.Drawing.Point(3, 0)
        Me.gbInclude.Name = "gbInclude"
        Me.gbInclude.Size = New System.Drawing.Size(158, 57)
        Me.gbInclude.TabIndex = 1
        Me.gbInclude.TabStop = False
        Me.gbInclude.Text = "Show"
        '
        'rbAll
        '
        Me.rbAll.Location = New System.Drawing.Point(95, 21)
        Me.rbAll.Name = "rbAll"
        Me.rbAll.Size = New System.Drawing.Size(57, 25)
        Me.rbAll.TabIndex = 1
        Me.rbAll.Text = "All"
        Me.rbAll.UseVisualStyleBackColor = True
        '
        'rbInclude
        '
        Me.rbInclude.Checked = True
        Me.rbInclude.Location = New System.Drawing.Point(6, 21)
        Me.rbInclude.Name = "rbInclude"
        Me.rbInclude.Size = New System.Drawing.Size(83, 25)
        Me.rbInclude.TabIndex = 0
        Me.rbInclude.TabStop = True
        Me.rbInclude.Text = "Included"
        Me.rbInclude.UseVisualStyleBackColor = True
        '
        'dgvFC
        '
        Me.dgvFC.AllowUserToDeleteRows = False
        Me.dgvFC.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvFC.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader
        Me.dgvFC.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle167.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle167.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle167.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle167.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle167.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle167.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle167.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvFC.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle167
        Me.dgvFC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle168.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle168.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle168.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle168.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle168.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle168.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle168.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle168.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvFC.DefaultCellStyle = DataGridViewCellStyle168
        Me.dgvFC.Location = New System.Drawing.Point(68, 64)
        Me.dgvFC.Name = "dgvFC"
        Me.dgvFC.ReadOnly = True
        Me.dgvFC.Size = New System.Drawing.Size(437, 334)
        Me.dgvFC.TabIndex = 0
        '
        'tabData4
        '
        Me.tabData4.BackColor = System.Drawing.Color.Ivory
        Me.tabData4.Controls.Add(Me.gbMeanComp)
        Me.tabData4.Controls.Add(Me.gbCritPrecision)
        Me.tabData4.Controls.Add(Me.gbRound5)
        Me.tabData4.Location = New System.Drawing.Point(4, 34)
        Me.tabData4.Name = "tabData4"
        Me.tabData4.Size = New System.Drawing.Size(508, 401)
        Me.tabData4.TabIndex = 3
        Me.tabData4.Text = "Rounding" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Conventions"
        '
        'gbMeanComp
        '
        Me.gbMeanComp.BackColor = System.Drawing.Color.White
        Me.gbMeanComp.Controls.Add(Me.rbMeanRounded)
        Me.gbMeanComp.Controls.Add(Me.rbMeanFullPrec)
        Me.gbMeanComp.Location = New System.Drawing.Point(22, 133)
        Me.gbMeanComp.Name = "gbMeanComp"
        Me.gbMeanComp.Size = New System.Drawing.Size(472, 81)
        Me.gbMeanComp.TabIndex = 2
        Me.gbMeanComp.TabStop = False
        Me.gbMeanComp.Text = "Mean Comparison Convention"
        Me.gbMeanComp.Visible = False
        '
        'rbMeanRounded
        '
        Me.rbMeanRounded.AutoSize = True
        Me.rbMeanRounded.Checked = True
        Me.rbMeanRounded.Location = New System.Drawing.Point(32, 47)
        Me.rbMeanRounded.Name = "rbMeanRounded"
        Me.rbMeanRounded.Size = New System.Drawing.Size(389, 21)
        Me.rbMeanRounded.TabIndex = 3
        Me.rbMeanRounded.TabStop = True
        Me.rbMeanRounded.Text = "Use rounded precision mean in mean comparison calculations"
        Me.rbMeanRounded.UseVisualStyleBackColor = True
        '
        'rbMeanFullPrec
        '
        Me.rbMeanFullPrec.AutoSize = True
        Me.rbMeanFullPrec.Location = New System.Drawing.Point(32, 21)
        Me.rbMeanFullPrec.Name = "rbMeanFullPrec"
        Me.rbMeanFullPrec.Size = New System.Drawing.Size(356, 21)
        Me.rbMeanFullPrec.TabIndex = 2
        Me.rbMeanFullPrec.Text = "Use full precision mean in mean comparison calculations"
        Me.rbMeanFullPrec.UseVisualStyleBackColor = True
        '
        'gbCritPrecision
        '
        Me.gbCritPrecision.BackColor = System.Drawing.Color.White
        Me.gbCritPrecision.Controls.Add(Me.rbCritRounded)
        Me.gbCritPrecision.Controls.Add(Me.rbCritFullPrec)
        Me.gbCritPrecision.Location = New System.Drawing.Point(22, 225)
        Me.gbCritPrecision.Name = "gbCritPrecision"
        Me.gbCritPrecision.Size = New System.Drawing.Size(472, 81)
        Me.gbCritPrecision.TabIndex = 1
        Me.gbCritPrecision.TabStop = False
        Me.gbCritPrecision.Text = "Criteria Precision Convention"
        Me.gbCritPrecision.Visible = False
        '
        'rbCritRounded
        '
        Me.rbCritRounded.AutoSize = True
        Me.rbCritRounded.Checked = True
        Me.rbCritRounded.Location = New System.Drawing.Point(32, 48)
        Me.rbCritRounded.Name = "rbCritRounded"
        Me.rbCritRounded.Size = New System.Drawing.Size(203, 21)
        Me.rbCritRounded.TabIndex = 3
        Me.rbCritRounded.TabStop = True
        Me.rbCritRounded.Text = "Use rounded precision criteria"
        Me.rbCritRounded.UseVisualStyleBackColor = True
        '
        'rbCritFullPrec
        '
        Me.rbCritFullPrec.AutoSize = True
        Me.rbCritFullPrec.Location = New System.Drawing.Point(32, 22)
        Me.rbCritFullPrec.Name = "rbCritFullPrec"
        Me.rbCritFullPrec.Size = New System.Drawing.Size(170, 21)
        Me.rbCritFullPrec.TabIndex = 2
        Me.rbCritFullPrec.Text = "Use full precision criteria"
        Me.rbCritFullPrec.UseVisualStyleBackColor = True
        '
        'gbRound5
        '
        Me.gbRound5.BackColor = System.Drawing.Color.White
        Me.gbRound5.Controls.Add(Me.rbRoundFiveAway)
        Me.gbRound5.Controls.Add(Me.rbRoundFiveEven)
        Me.gbRound5.Location = New System.Drawing.Point(22, 22)
        Me.gbRound5.Name = "gbRound5"
        Me.gbRound5.Size = New System.Drawing.Size(472, 100)
        Me.gbRound5.TabIndex = 0
        Me.gbRound5.TabStop = False
        Me.gbRound5.Text = "Round 5 Convention"
        '
        'rbRoundFiveAway
        '
        Me.rbRoundFiveAway.Checked = True
        Me.rbRoundFiveAway.Location = New System.Drawing.Point(32, 45)
        Me.rbRoundFiveAway.Name = "rbRoundFiveAway"
        Me.rbRoundFiveAway.Size = New System.Drawing.Size(354, 40)
        Me.rbRoundFiveAway.TabIndex = 1
        Me.rbRoundFiveAway.TabStop = True
        Me.rbRoundFiveAway.Text = "Round 5 away from zero (mimics Watson and Excel ROUND function)"
        Me.rbRoundFiveAway.UseVisualStyleBackColor = True
        '
        'rbRoundFiveEven
        '
        Me.rbRoundFiveEven.AutoSize = True
        Me.rbRoundFiveEven.Location = New System.Drawing.Point(32, 19)
        Me.rbRoundFiveEven.Name = "rbRoundFiveEven"
        Me.rbRoundFiveEven.Size = New System.Drawing.Size(242, 21)
        Me.rbRoundFiveEven.TabIndex = 0
        Me.rbRoundFiveEven.Text = "Round 5 to even (Bankers' Rounding)"
        Me.rbRoundFiveEven.UseVisualStyleBackColor = True
        '
        'tabData5
        '
        Me.tabData5.BackColor = System.Drawing.Color.Ivory
        Me.tabData5.Controls.Add(Me.Label4)
        Me.tabData5.Controls.Add(Me.lblSortAnalyte)
        Me.tabData5.Controls.Add(Me.cmdDownA)
        Me.tabData5.Controls.Add(Me.cmdUpA)
        Me.tabData5.Controls.Add(Me.dgvAnalyteGroups)
        Me.tabData5.Location = New System.Drawing.Point(4, 34)
        Me.tabData5.Name = "tabData5"
        Me.tabData5.Size = New System.Drawing.Size(508, 401)
        Me.tabData5.TabIndex = 4
        Me.tabData5.Text = "Analyte Sort Order"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(8, 63)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 37)
        Me.Label4.TabIndex = 147
        Me.Label4.Text = "Move" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Row"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblSortAnalyte
        '
        Me.lblSortAnalyte.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblSortAnalyte.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSortAnalyte.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblSortAnalyte.Location = New System.Drawing.Point(68, 13)
        Me.lblSortAnalyte.Name = "lblSortAnalyte"
        Me.lblSortAnalyte.Size = New System.Drawing.Size(417, 45)
        Me.lblSortAnalyte.TabIndex = 146
        Me.lblSortAnalyte.Text = "Order the Analytes to be displayed in the report" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Analytes can be excluded in th" & _
    "e Configure Report Tables window)"
        Me.lblSortAnalyte.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdDownA
        '
        Me.cmdDownA.BackColor = System.Drawing.Color.Transparent
        Me.cmdDownA.BackgroundImage = CType(resources.GetObject("cmdDownA.BackgroundImage"), System.Drawing.Image)
        Me.cmdDownA.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdDownA.Enabled = False
        Me.cmdDownA.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cmdDownA.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdDownA.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdDownA.Location = New System.Drawing.Point(7, 149)
        Me.cmdDownA.Name = "cmdDownA"
        Me.cmdDownA.Size = New System.Drawing.Size(55, 40)
        Me.cmdDownA.TabIndex = 145
        Me.cmdDownA.UseVisualStyleBackColor = False
        '
        'cmdUpA
        '
        Me.cmdUpA.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdUpA.BackgroundImage = CType(resources.GetObject("cmdUpA.BackgroundImage"), System.Drawing.Image)
        Me.cmdUpA.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdUpA.Enabled = False
        Me.cmdUpA.FlatAppearance.BorderColor = System.Drawing.Color.Black
        Me.cmdUpA.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpA.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdUpA.Location = New System.Drawing.Point(7, 103)
        Me.cmdUpA.Name = "cmdUpA"
        Me.cmdUpA.Size = New System.Drawing.Size(55, 40)
        Me.cmdUpA.TabIndex = 144
        Me.cmdUpA.UseVisualStyleBackColor = False
        '
        'dgvAnalyteGroups
        '
        Me.dgvAnalyteGroups.AllowUserToAddRows = False
        Me.dgvAnalyteGroups.AllowUserToDeleteRows = False
        Me.dgvAnalyteGroups.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAnalyteGroups.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle169.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle169.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle169.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle169.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle169.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle169.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle169.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAnalyteGroups.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle169
        Me.dgvAnalyteGroups.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle170.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle170.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle170.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle170.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle170.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle170.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle170.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle170.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvAnalyteGroups.DefaultCellStyle = DataGridViewCellStyle170
        Me.dgvAnalyteGroups.Location = New System.Drawing.Point(68, 64)
        Me.dgvAnalyteGroups.MultiSelect = False
        Me.dgvAnalyteGroups.Name = "dgvAnalyteGroups"
        Me.dgvAnalyteGroups.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvAnalyteGroups.Size = New System.Drawing.Size(417, 314)
        Me.dgvAnalyteGroups.TabIndex = 0
        '
        'cbxSubmittedTo
        '
        Me.cbxSubmittedTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxSubmittedTo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSubmittedTo.Enabled = False
        Me.cbxSubmittedTo.IntegralHeight = False
        Me.cbxSubmittedTo.Location = New System.Drawing.Point(548, 43)
        Me.cbxSubmittedTo.Name = "cbxSubmittedTo"
        Me.cbxSubmittedTo.Size = New System.Drawing.Size(226, 25)
        Me.cbxSubmittedTo.TabIndex = 48
        '
        'lblDCP
        '
        Me.lblDCP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDCP.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDCP.Location = New System.Drawing.Point(6, 132)
        Me.lblDCP.Name = "lblDCP"
        Me.lblDCP.Size = New System.Drawing.Size(519, 19)
        Me.lblDCP.TabIndex = 89
        Me.lblDCP.Text = "Data and Configuration Parameters Entered by User"
        Me.lblDCP.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'cbxSampleStorageTemp
        '
        Me.cbxSampleStorageTemp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxSampleStorageTemp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleStorageTemp.Enabled = False
        Me.cbxSampleStorageTemp.IntegralHeight = False
        Me.cbxSampleStorageTemp.Location = New System.Drawing.Point(165, 17)
        Me.cbxSampleStorageTemp.Name = "cbxSampleStorageTemp"
        Me.cbxSampleStorageTemp.Size = New System.Drawing.Size(360, 25)
        Me.cbxSampleStorageTemp.TabIndex = 87
        Me.cbxSampleStorageTemp.Visible = False
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(15, 21)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(153, 14)
        Me.Label5.TabIndex = 86
        Me.Label5.Text = "Sample Storage Temp:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label5.Visible = False
        '
        'lblIncludeInTitle
        '
        Me.lblIncludeInTitle.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblIncludeInTitle.AutoSize = True
        Me.lblIncludeInTitle.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblIncludeInTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblIncludeInTitle.Location = New System.Drawing.Point(549, 23)
        Me.lblIncludeInTitle.Name = "lblIncludeInTitle"
        Me.lblIncludeInTitle.Size = New System.Drawing.Size(210, 17)
        Me.lblIncludeInTitle.TabIndex = 80
        Me.lblIncludeInTitle.Text = "* Will be included in report title."
        '
        'Label44
        '
        Me.Label44.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label44.AutoSize = True
        Me.Label44.BackColor = System.Drawing.Color.Transparent
        Me.Label44.Location = New System.Drawing.Point(775, 246)
        Me.Label44.Name = "Label44"
        Me.Label44.Size = New System.Drawing.Size(84, 17)
        Me.Label44.TabIndex = 79
        Me.Label44.Text = "Submitted By"
        '
        'txtSubmittedTo
        '
        Me.txtSubmittedTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubmittedTo.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSubmittedTo.Location = New System.Drawing.Point(549, 69)
        Me.txtSubmittedTo.Multiline = True
        Me.txtSubmittedTo.Name = "txtSubmittedTo"
        Me.txtSubmittedTo.ReadOnly = True
        Me.txtSubmittedTo.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSubmittedTo.Size = New System.Drawing.Size(353, 65)
        Me.txtSubmittedTo.TabIndex = 78
        '
        'Label43
        '
        Me.Label43.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label43.AutoSize = True
        Me.Label43.BackColor = System.Drawing.Color.Transparent
        Me.Label43.Location = New System.Drawing.Point(777, 148)
        Me.Label43.Name = "Label43"
        Me.Label43.Size = New System.Drawing.Size(87, 17)
        Me.Label43.TabIndex = 77
        Me.Label43.Text = "In Support Of"
        '
        'Label42
        '
        Me.Label42.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label42.AutoSize = True
        Me.Label42.BackColor = System.Drawing.Color.Transparent
        Me.Label42.Location = New System.Drawing.Point(775, 49)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(86, 17)
        Me.Label42.TabIndex = 76
        Me.Label42.Text = "Submitted To"
        '
        'txtSubmittedBy
        '
        Me.txtSubmittedBy.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSubmittedBy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSubmittedBy.Location = New System.Drawing.Point(549, 266)
        Me.txtSubmittedBy.Multiline = True
        Me.txtSubmittedBy.Name = "txtSubmittedBy"
        Me.txtSubmittedBy.ReadOnly = True
        Me.txtSubmittedBy.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSubmittedBy.Size = New System.Drawing.Size(353, 65)
        Me.txtSubmittedBy.TabIndex = 75
        '
        'txtInSupportOf
        '
        Me.txtInSupportOf.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInSupportOf.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtInSupportOf.Location = New System.Drawing.Point(549, 168)
        Me.txtInSupportOf.Multiline = True
        Me.txtInSupportOf.Name = "txtInSupportOf"
        Me.txtInSupportOf.ReadOnly = True
        Me.txtInSupportOf.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtInSupportOf.Size = New System.Drawing.Size(353, 65)
        Me.txtInSupportOf.TabIndex = 74
        '
        'cbxSubmittedBy
        '
        Me.cbxSubmittedBy.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxSubmittedBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSubmittedBy.Enabled = False
        Me.cbxSubmittedBy.IntegralHeight = False
        Me.cbxSubmittedBy.Location = New System.Drawing.Point(549, 240)
        Me.cbxSubmittedBy.Name = "cbxSubmittedBy"
        Me.cbxSubmittedBy.Size = New System.Drawing.Size(226, 25)
        Me.cbxSubmittedBy.TabIndex = 50
        '
        'cbxInSupportOf
        '
        Me.cbxInSupportOf.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxInSupportOf.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxInSupportOf.Enabled = False
        Me.cbxInSupportOf.IntegralHeight = False
        Me.cbxInSupportOf.Location = New System.Drawing.Point(549, 142)
        Me.cbxInSupportOf.Name = "cbxInSupportOf"
        Me.cbxInSupportOf.Size = New System.Drawing.Size(228, 25)
        Me.cbxInSupportOf.TabIndex = 49
        '
        'cbxSampleSizeUnits
        '
        Me.cbxSampleSizeUnits.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxSampleSizeUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSampleSizeUnits.Enabled = False
        Me.cbxSampleSizeUnits.IntegralHeight = False
        Me.cbxSampleSizeUnits.Location = New System.Drawing.Point(165, 11)
        Me.cbxSampleSizeUnits.Name = "cbxSampleSizeUnits"
        Me.cbxSampleSizeUnits.Size = New System.Drawing.Size(360, 25)
        Me.cbxSampleSizeUnits.TabIndex = 47
        Me.cbxSampleSizeUnits.Visible = False
        '
        'cbxAnticoagulant
        '
        Me.cbxAnticoagulant.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxAnticoagulant.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxAnticoagulant.Enabled = False
        Me.cbxAnticoagulant.IntegralHeight = False
        Me.cbxAnticoagulant.Location = New System.Drawing.Point(177, 90)
        Me.cbxAnticoagulant.Name = "cbxAnticoagulant"
        Me.cbxAnticoagulant.Size = New System.Drawing.Size(348, 25)
        Me.cbxAnticoagulant.Sorted = True
        Me.cbxAnticoagulant.TabIndex = 46
        '
        'cbxAssayTechniqueAcronym
        '
        Me.cbxAssayTechniqueAcronym.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbxAssayTechniqueAcronym.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxAssayTechniqueAcronym.Enabled = False
        Me.cbxAssayTechniqueAcronym.IntegralHeight = False
        Me.cbxAssayTechniqueAcronym.Location = New System.Drawing.Point(177, 58)
        Me.cbxAssayTechniqueAcronym.Name = "cbxAssayTechniqueAcronym"
        Me.cbxAssayTechniqueAcronym.Size = New System.Drawing.Size(348, 25)
        Me.cbxAssayTechniqueAcronym.TabIndex = 45
        '
        'Label29
        '
        Me.Label29.AutoSize = True
        Me.Label29.BackColor = System.Drawing.Color.Transparent
        Me.Label29.Location = New System.Drawing.Point(9, 60)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(162, 17)
        Me.Label29.TabIndex = 33
        Me.Label29.Text = "Assay Technique Acronym:"
        '
        'Label28
        '
        Me.Label28.AutoSize = True
        Me.Label28.BackColor = System.Drawing.Color.Transparent
        Me.Label28.Location = New System.Drawing.Point(9, 94)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(90, 17)
        Me.Label28.TabIndex = 32
        Me.Label28.Text = "Anticoagulant:"
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(66, 12)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(102, 20)
        Me.Label27.TabIndex = 31
        Me.Label27.Text = "Sample Size Units:"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label27.Visible = False
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(549, 338)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(353, 16)
        Me.Label7.TabIndex = 25
        Me.Label7.Text = "Data From Watson (read-only)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblData
        '
        Me.lblData.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblData.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblData.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblData.ForeColor = System.Drawing.Color.White
        Me.lblData.Location = New System.Drawing.Point(0, 0)
        Me.lblData.Name = "lblData"
        Me.lblData.Size = New System.Drawing.Size(908, 21)
        Me.lblData.TabIndex = 22
        Me.lblData.Text = "Add/Edit Top Level Data"
        '
        'tp3
        '
        Me.tp3.AutoScroll = True
        Me.tp3.BackColor = System.Drawing.Color.Ivory
        Me.tp3.Controls.Add(Me.lblARS_B_Note)
        Me.tp3.Controls.Add(Me.dgvAnalyticalRunSummary)
        Me.tp3.Controls.Add(Me.lblAnalyticalRunSummary)
        Me.tp3.Controls.Add(Me.gbxlblReviewAnalyticalRuns)
        Me.tp3.Controls.Add(Me.gbReportOptions)
        Me.tp3.Location = New System.Drawing.Point(4, 24)
        Me.tp3.Name = "tp3"
        Me.tp3.Size = New System.Drawing.Size(908, 603)
        Me.tp3.TabIndex = 2
        Me.tp3.Text = "3"
        '
        'lblARS_B_Note
        '
        Me.lblARS_B_Note.BackColor = System.Drawing.Color.Transparent
        Me.lblARS_B_Note.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblARS_B_Note.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblARS_B_Note.Location = New System.Drawing.Point(14, 70)
        Me.lblARS_B_Note.Name = "lblARS_B_Note"
        Me.lblARS_B_Note.Size = New System.Drawing.Size(280, 57)
        Me.lblARS_B_Note.TabIndex = 138
        Me.lblARS_B_Note.Text = "B NOTE: If the analytical run has no calibration" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "information, that run will not " & _
    "be reported," & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "even if checked"
        '
        'dgvAnalyticalRunSummary
        '
        Me.dgvAnalyticalRunSummary.AllowUserToAddRows = False
        Me.dgvAnalyticalRunSummary.AllowUserToDeleteRows = False
        Me.dgvAnalyticalRunSummary.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAnalyticalRunSummary.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvAnalyticalRunSummary.BackgroundColor = System.Drawing.Color.White
        Me.dgvAnalyticalRunSummary.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle171.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle171.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle171.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle171.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle171.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle171.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle171.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvAnalyticalRunSummary.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle171
        Me.dgvAnalyticalRunSummary.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle172.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle172.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle172.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle172.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle172.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle172.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle172.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle172.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvAnalyticalRunSummary.DefaultCellStyle = DataGridViewCellStyle172
        Me.dgvAnalyticalRunSummary.Location = New System.Drawing.Point(9, 135)
        Me.dgvAnalyticalRunSummary.Name = "dgvAnalyticalRunSummary"
        Me.dgvAnalyticalRunSummary.ReadOnly = True
        Me.dgvAnalyticalRunSummary.RowHeadersWidth = 25
        Me.dgvAnalyticalRunSummary.Size = New System.Drawing.Size(889, 465)
        Me.dgvAnalyticalRunSummary.TabIndex = 94
        '
        'lblAnalyticalRunSummary
        '
        Me.lblAnalyticalRunSummary.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblAnalyticalRunSummary.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblAnalyticalRunSummary.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblAnalyticalRunSummary.ForeColor = System.Drawing.Color.White
        Me.lblAnalyticalRunSummary.Location = New System.Drawing.Point(0, 0)
        Me.lblAnalyticalRunSummary.Name = "lblAnalyticalRunSummary"
        Me.lblAnalyticalRunSummary.Size = New System.Drawing.Size(908, 21)
        Me.lblAnalyticalRunSummary.TabIndex = 22
        Me.lblAnalyticalRunSummary.Text = "Review Analytical Runs"
        '
        'gbxlblReviewAnalyticalRuns
        '
        Me.gbxlblReviewAnalyticalRuns.Controls.Add(Me.lblARS_B)
        Me.gbxlblReviewAnalyticalRuns.Controls.Add(Me.lblARS_A)
        Me.gbxlblReviewAnalyticalRuns.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblReviewAnalyticalRuns.Location = New System.Drawing.Point(9, 20)
        Me.gbxlblReviewAnalyticalRuns.Name = "gbxlblReviewAnalyticalRuns"
        Me.gbxlblReviewAnalyticalRuns.Size = New System.Drawing.Size(294, 111)
        Me.gbxlblReviewAnalyticalRuns.TabIndex = 0
        Me.gbxlblReviewAnalyticalRuns.TabStop = False
        '
        'lblARS_B
        '
        Me.lblARS_B.AutoSize = True
        Me.lblARS_B.BackColor = System.Drawing.Color.Transparent
        Me.lblARS_B.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblARS_B.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblARS_B.Location = New System.Drawing.Point(0, 30)
        Me.lblARS_B.Name = "lblARS_B"
        Me.lblARS_B.Size = New System.Drawing.Size(295, 17)
        Me.lblARS_B.TabIndex = 137
        Me.lblARS_B.Text = "*B = Include in Regression, Calibr, and QC Tables"
        '
        'lblARS_A
        '
        Me.lblARS_A.AutoSize = True
        Me.lblARS_A.BackColor = System.Drawing.Color.Transparent
        Me.lblARS_A.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblARS_A.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblARS_A.Location = New System.Drawing.Point(0, 10)
        Me.lblARS_A.Name = "lblARS_A"
        Me.lblARS_A.Size = New System.Drawing.Size(213, 17)
        Me.lblARS_A.TabIndex = 95
        Me.lblARS_A.Text = "*A = Include in Run Summary Table"
        '
        'gbReportOptions
        '
        Me.gbReportOptions.Controls.Add(Me.panAnalRunSum)
        Me.gbReportOptions.Controls.Add(Me.panAnalRunChoices)
        Me.gbReportOptions.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbReportOptions.Location = New System.Drawing.Point(309, 21)
        Me.gbReportOptions.Name = "gbReportOptions"
        Me.gbReportOptions.Size = New System.Drawing.Size(552, 110)
        Me.gbReportOptions.TabIndex = 139
        Me.gbReportOptions.TabStop = False
        Me.gbReportOptions.Text = "Analytical Run Summary Table Options"
        '
        'panAnalRunSum
        '
        Me.panAnalRunSum.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panAnalRunSum.Controls.Add(Me.rbUseWatsonComments)
        Me.panAnalRunSum.Controls.Add(Me.rbUseUserComments)
        Me.panAnalRunSum.Enabled = False
        Me.panAnalRunSum.Location = New System.Drawing.Point(366, 17)
        Me.panAnalRunSum.Name = "panAnalRunSum"
        Me.panAnalRunSum.Size = New System.Drawing.Size(177, 87)
        Me.panAnalRunSum.TabIndex = 26
        '
        'rbUseWatsonComments
        '
        Me.rbUseWatsonComments.AutoSize = True
        Me.rbUseWatsonComments.BackColor = System.Drawing.Color.Transparent
        Me.rbUseWatsonComments.Checked = True
        Me.rbUseWatsonComments.Location = New System.Drawing.Point(10, 17)
        Me.rbUseWatsonComments.Name = "rbUseWatsonComments"
        Me.rbUseWatsonComments.Size = New System.Drawing.Size(162, 21)
        Me.rbUseWatsonComments.TabIndex = 0
        Me.rbUseWatsonComments.TabStop = True
        Me.rbUseWatsonComments.Text = "Use Watson Comments"
        Me.rbUseWatsonComments.UseVisualStyleBackColor = False
        '
        'rbUseUserComments
        '
        Me.rbUseUserComments.AutoSize = True
        Me.rbUseUserComments.BackColor = System.Drawing.Color.Transparent
        Me.rbUseUserComments.Location = New System.Drawing.Point(10, 46)
        Me.rbUseUserComments.Name = "rbUseUserComments"
        Me.rbUseUserComments.Size = New System.Drawing.Size(145, 21)
        Me.rbUseUserComments.TabIndex = 1
        Me.rbUseUserComments.Text = "Use User Comments"
        Me.rbUseUserComments.UseVisualStyleBackColor = False
        '
        'panAnalRunChoices
        '
        Me.panAnalRunChoices.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panAnalRunChoices.Controls.Add(Me.lblAnalRunReportOptions)
        Me.panAnalRunChoices.Controls.Add(Me.chkPSAE)
        Me.panAnalRunChoices.Controls.Add(Me.chkNoRegrPerformed)
        Me.panAnalRunChoices.Controls.Add(Me.chkRegrPerformed)
        Me.panAnalRunChoices.Controls.Add(Me.chkRejected)
        Me.panAnalRunChoices.Controls.Add(Me.chkAccepted)
        Me.panAnalRunChoices.Controls.Add(Me.chkAll)
        Me.panAnalRunChoices.Location = New System.Drawing.Point(6, 17)
        Me.panAnalRunChoices.Name = "panAnalRunChoices"
        Me.panAnalRunChoices.Size = New System.Drawing.Size(354, 87)
        Me.panAnalRunChoices.TabIndex = 159
        '
        'lblAnalRunReportOptions
        '
        Me.lblAnalRunReportOptions.AutoSize = True
        Me.lblAnalRunReportOptions.Location = New System.Drawing.Point(9, 3)
        Me.lblAnalRunReportOptions.Name = "lblAnalRunReportOptions"
        Me.lblAnalRunReportOptions.Size = New System.Drawing.Size(256, 17)
        Me.lblAnalRunReportOptions.TabIndex = 6
        Me.lblAnalRunReportOptions.Text = "Include the following run/regression types:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'chkPSAE
        '
        Me.chkPSAE.AutoSize = True
        Me.chkPSAE.Location = New System.Drawing.Point(172, 61)
        Me.chkPSAE.Name = "chkPSAE"
        Me.chkPSAE.Size = New System.Drawing.Size(133, 21)
        Me.chkPSAE.TabIndex = 5
        Me.chkPSAE.Text = "Include PSAE Runs"
        Me.chkPSAE.UseVisualStyleBackColor = True
        '
        'chkNoRegrPerformed
        '
        Me.chkNoRegrPerformed.AutoSize = True
        Me.chkNoRegrPerformed.Location = New System.Drawing.Point(172, 42)
        Me.chkNoRegrPerformed.Name = "chkNoRegrPerformed"
        Me.chkNoRegrPerformed.Size = New System.Drawing.Size(182, 21)
        Me.chkNoRegrPerformed.TabIndex = 4
        Me.chkNoRegrPerformed.Text = "NO Regression Performed"
        Me.chkNoRegrPerformed.UseVisualStyleBackColor = True
        '
        'chkRegrPerformed
        '
        Me.chkRegrPerformed.AutoSize = True
        Me.chkRegrPerformed.Location = New System.Drawing.Point(172, 22)
        Me.chkRegrPerformed.Name = "chkRegrPerformed"
        Me.chkRegrPerformed.Size = New System.Drawing.Size(158, 21)
        Me.chkRegrPerformed.TabIndex = 3
        Me.chkRegrPerformed.Text = "Regression Performed"
        Me.chkRegrPerformed.UseVisualStyleBackColor = True
        '
        'chkRejected
        '
        Me.chkRejected.AutoSize = True
        Me.chkRejected.Location = New System.Drawing.Point(12, 61)
        Me.chkRejected.Name = "chkRejected"
        Me.chkRejected.Size = New System.Drawing.Size(146, 21)
        Me.chkRejected.TabIndex = 2
        Me.chkRejected.Text = "Rejected Regression"
        Me.chkRejected.UseVisualStyleBackColor = True
        '
        'chkAccepted
        '
        Me.chkAccepted.AutoSize = True
        Me.chkAccepted.Location = New System.Drawing.Point(12, 42)
        Me.chkAccepted.Name = "chkAccepted"
        Me.chkAccepted.Size = New System.Drawing.Size(150, 21)
        Me.chkAccepted.TabIndex = 1
        Me.chkAccepted.Text = "Accepted Regression"
        Me.chkAccepted.UseVisualStyleBackColor = True
        '
        'chkAll
        '
        Me.chkAll.AutoSize = True
        Me.chkAll.Location = New System.Drawing.Point(12, 22)
        Me.chkAll.Name = "chkAll"
        Me.chkAll.Size = New System.Drawing.Size(131, 21)
        Me.chkAll.TabIndex = 0
        Me.chkAll.Text = "All Analytical Runs"
        Me.chkAll.UseVisualStyleBackColor = True
        '
        'tp4
        '
        Me.tp4.AutoScroll = True
        Me.tp4.BackColor = System.Drawing.Color.Ivory
        Me.tp4.Controls.Add(Me.dgvSummaryData)
        Me.tp4.Controls.Add(Me.llblSummaryTable)
        Me.tp4.Controls.Add(Me.gbxlblMethodValidation)
        Me.tp4.Controls.Add(Me.cmdOrderSummaryTable)
        Me.tp4.Controls.Add(Me.lblSummaryTable)
        Me.tp4.Location = New System.Drawing.Point(4, 24)
        Me.tp4.Name = "tp4"
        Me.tp4.Size = New System.Drawing.Size(908, 603)
        Me.tp4.TabIndex = 3
        Me.tp4.Text = "4"
        '
        'dgvSummaryData
        '
        Me.dgvSummaryData.AllowUserToAddRows = False
        Me.dgvSummaryData.AllowUserToDeleteRows = False
        Me.dgvSummaryData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSummaryData.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSummaryData.BackgroundColor = System.Drawing.Color.White
        Me.dgvSummaryData.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.dgvSummaryData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSummaryData.Location = New System.Drawing.Point(10, 91)
        Me.dgvSummaryData.Name = "dgvSummaryData"
        Me.dgvSummaryData.ReadOnly = True
        Me.dgvSummaryData.RowHeadersWidth = 25
        DataGridViewCellStyle173.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle173.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSummaryData.RowsDefaultCellStyle = DataGridViewCellStyle173
        Me.dgvSummaryData.Size = New System.Drawing.Size(892, 509)
        Me.dgvSummaryData.TabIndex = 137
        '
        'llblSummaryTable
        '
        Me.llblSummaryTable.AutoSize = True
        Me.llblSummaryTable.BackColor = System.Drawing.Color.Transparent
        Me.llblSummaryTable.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.llblSummaryTable.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.llblSummaryTable.LinkColor = System.Drawing.Color.Blue
        Me.llblSummaryTable.Location = New System.Drawing.Point(7, 22)
        Me.llblSummaryTable.Name = "llblSummaryTable"
        Me.llblSummaryTable.Size = New System.Drawing.Size(349, 15)
        Me.llblSummaryTable.TabIndex = 135
        Me.llblSummaryTable.TabStop = True
        Me.llblSummaryTable.Text = "(Click here to Configure as an Appendix in the Appendices Tab)"
        '
        'gbxlblMethodValidation
        '
        Me.gbxlblMethodValidation.Controls.Add(Me.Label6)
        Me.gbxlblMethodValidation.Location = New System.Drawing.Point(9, 33)
        Me.gbxlblMethodValidation.Margin = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.gbxlblMethodValidation.Name = "gbxlblMethodValidation"
        Me.gbxlblMethodValidation.Padding = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.gbxlblMethodValidation.Size = New System.Drawing.Size(216, 28)
        Me.gbxlblMethodValidation.TabIndex = 136
        Me.gbxlblMethodValidation.TabStop = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label6.Location = New System.Drawing.Point(0, 9)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(122, 13)
        Me.Label6.TabIndex = 102
        Me.Label6.Text = "* A = Include in report"
        '
        'cmdOrderSummaryTable
        '
        Me.cmdOrderSummaryTable.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOrderSummaryTable.BackgroundImage = CType(resources.GetObject("cmdOrderSummaryTable.BackgroundImage"), System.Drawing.Image)
        Me.cmdOrderSummaryTable.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdOrderSummaryTable.Enabled = False
        Me.cmdOrderSummaryTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOrderSummaryTable.Location = New System.Drawing.Point(42, 63)
        Me.cmdOrderSummaryTable.Name = "cmdOrderSummaryTable"
        Me.cmdOrderSummaryTable.Size = New System.Drawing.Size(50, 26)
        Me.cmdOrderSummaryTable.TabIndex = 133
        Me.cmdOrderSummaryTable.Text = "Re-#"
        Me.cmdOrderSummaryTable.UseVisualStyleBackColor = False
        '
        'lblSummaryTable
        '
        Me.lblSummaryTable.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblSummaryTable.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblSummaryTable.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblSummaryTable.ForeColor = System.Drawing.Color.White
        Me.lblSummaryTable.Location = New System.Drawing.Point(0, 0)
        Me.lblSummaryTable.Name = "lblSummaryTable"
        Me.lblSummaryTable.Size = New System.Drawing.Size(908, 21)
        Me.lblSummaryTable.TabIndex = 22
        Me.lblSummaryTable.Text = "Summary Table -  Method Validation and Study Information"
        '
        'tp5
        '
        Me.tp5.AutoScroll = True
        Me.tp5.BackColor = System.Drawing.Color.Ivory
        Me.tp5.Controls.Add(Me.gbxlblChooseEditWordTemplate)
        Me.tp5.Controls.Add(Me.panSections)
        Me.tp5.Controls.Add(Me.gbSectionStyle)
        Me.tp5.Controls.Add(Me.panRBSwb)
        Me.tp5.Controls.Add(Me.cmdRBSAll)
        Me.tp5.Controls.Add(Me.grpRBS)
        Me.tp5.Controls.Add(Me.cmdOrderReportBodySection)
        Me.tp5.Controls.Add(Me.lblRBS)
        Me.tp5.Controls.Add(Me.dgvReportStatementWord)
        Me.tp5.Controls.Add(Me.dgvReportStatements)
        Me.tp5.Controls.Add(Me.lblReportStatement)
        Me.tp5.Location = New System.Drawing.Point(4, 24)
        Me.tp5.Margin = New System.Windows.Forms.Padding(0)
        Me.tp5.Name = "tp5"
        Me.tp5.Size = New System.Drawing.Size(908, 603)
        Me.tp5.TabIndex = 7
        Me.tp5.Text = "5"
        '
        'gbxlblChooseEditWordTemplate
        '
        Me.gbxlblChooseEditWordTemplate.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.gbxlblChooseEditWordTemplate.Controls.Add(Me.lblWordStatements)
        Me.gbxlblChooseEditWordTemplate.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblChooseEditWordTemplate.Location = New System.Drawing.Point(541, 64)
        Me.gbxlblChooseEditWordTemplate.Name = "gbxlblChooseEditWordTemplate"
        Me.gbxlblChooseEditWordTemplate.Padding = New System.Windows.Forms.Padding(0)
        Me.gbxlblChooseEditWordTemplate.Size = New System.Drawing.Size(323, 29)
        Me.gbxlblChooseEditWordTemplate.TabIndex = 148
        Me.gbxlblChooseEditWordTemplate.TabStop = False
        '
        'lblWordStatements
        '
        Me.lblWordStatements.AutoSize = True
        Me.lblWordStatements.BackColor = System.Drawing.Color.White
        Me.lblWordStatements.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWordStatements.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lblWordStatements.Location = New System.Drawing.Point(5, 7)
        Me.lblWordStatements.Margin = New System.Windows.Forms.Padding(0)
        Me.lblWordStatements.Name = "lblWordStatements"
        Me.lblWordStatements.Size = New System.Drawing.Size(307, 17)
        Me.lblWordStatements.TabIndex = 140
        Me.lblWordStatements.Text = "<< Doubleclick to assign Word Report Template"
        '
        'panSections
        '
        Me.panSections.Controls.Add(Me.cbxRBSFilter)
        Me.panSections.Controls.Add(Me.Label16)
        Me.panSections.Controls.Add(Me.cbxRBSTypeFilter)
        Me.panSections.Controls.Add(Me.Label19)
        Me.panSections.Controls.Add(Me.GroupBox2)
        Me.panSections.Location = New System.Drawing.Point(742, 14)
        Me.panSections.Name = "panSections"
        Me.panSections.Size = New System.Drawing.Size(134, 103)
        Me.panSections.TabIndex = 147
        Me.panSections.Visible = False
        '
        'cbxRBSFilter
        '
        Me.cbxRBSFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxRBSFilter.FormattingEnabled = True
        Me.cbxRBSFilter.IntegralHeight = False
        Me.cbxRBSFilter.Location = New System.Drawing.Point(3, 15)
        Me.cbxRBSFilter.Name = "cbxRBSFilter"
        Me.cbxRBSFilter.Size = New System.Drawing.Size(147, 25)
        Me.cbxRBSFilter.TabIndex = 135
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.Color.Transparent
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Blue
        Me.Label16.Location = New System.Drawing.Point(0, -1)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(154, 15)
        Me.Label16.TabIndex = 134
        Me.Label16.Text = "Filter Content for Company:"
        '
        'cbxRBSTypeFilter
        '
        Me.cbxRBSTypeFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxRBSTypeFilter.FormattingEnabled = True
        Me.cbxRBSTypeFilter.Location = New System.Drawing.Point(163, 14)
        Me.cbxRBSTypeFilter.MaxDropDownItems = 20
        Me.cbxRBSTypeFilter.Name = "cbxRBSTypeFilter"
        Me.cbxRBSTypeFilter.Size = New System.Drawing.Size(147, 25)
        Me.cbxRBSTypeFilter.TabIndex = 137
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.BackColor = System.Drawing.Color.Transparent
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Blue
        Me.Label19.Location = New System.Drawing.Point(160, -2)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(165, 15)
        Me.Label19.TabIndex = 136
        Me.Label19.Text = "Filter Content for Report Type"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.rbShowAllRBody)
        Me.GroupBox2.Controls.Add(Me.rbShowIncludedRBody)
        Me.GroupBox2.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(198, 42)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(200, 57)
        Me.GroupBox2.TabIndex = 130
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Show Rows"
        '
        'rbShowAllRBody
        '
        Me.rbShowAllRBody.AutoSize = True
        Me.rbShowAllRBody.BackColor = System.Drawing.Color.Transparent
        Me.rbShowAllRBody.Checked = True
        Me.rbShowAllRBody.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbShowAllRBody.Location = New System.Drawing.Point(6, 31)
        Me.rbShowAllRBody.Name = "rbShowAllRBody"
        Me.rbShowAllRBody.Size = New System.Drawing.Size(70, 17)
        Me.rbShowAllRBody.TabIndex = 1
        Me.rbShowAllRBody.TabStop = True
        Me.rbShowAllRBody.Text = "Show All"
        Me.rbShowAllRBody.UseVisualStyleBackColor = False
        '
        'rbShowIncludedRBody
        '
        Me.rbShowIncludedRBody.AutoSize = True
        Me.rbShowIncludedRBody.BackColor = System.Drawing.Color.Transparent
        Me.rbShowIncludedRBody.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbShowIncludedRBody.Location = New System.Drawing.Point(6, 12)
        Me.rbShowIncludedRBody.Name = "rbShowIncludedRBody"
        Me.rbShowIncludedRBody.Size = New System.Drawing.Size(102, 17)
        Me.rbShowIncludedRBody.TabIndex = 0
        Me.rbShowIncludedRBody.Text = "Show Included"
        Me.rbShowIncludedRBody.UseVisualStyleBackColor = False
        '
        'gbSectionStyle
        '
        Me.gbSectionStyle.Controls.Add(Me.rbEntireReport)
        Me.gbSectionStyle.Controls.Add(Me.rbSections)
        Me.gbSectionStyle.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSectionStyle.Location = New System.Drawing.Point(551, 4)
        Me.gbSectionStyle.Name = "gbSectionStyle"
        Me.gbSectionStyle.Size = New System.Drawing.Size(129, 56)
        Me.gbSectionStyle.TabIndex = 146
        Me.gbSectionStyle.TabStop = False
        Me.gbSectionStyle.Text = "Configuration Style"
        Me.gbSectionStyle.Visible = False
        '
        'rbEntireReport
        '
        Me.rbEntireReport.AutoSize = True
        Me.rbEntireReport.Checked = True
        Me.rbEntireReport.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbEntireReport.Location = New System.Drawing.Point(6, 33)
        Me.rbEntireReport.Name = "rbEntireReport"
        Me.rbEntireReport.Size = New System.Drawing.Size(93, 17)
        Me.rbEntireReport.TabIndex = 1
        Me.rbEntireReport.TabStop = True
        Me.rbEntireReport.Text = "Entire Report"
        Me.rbEntireReport.UseVisualStyleBackColor = True
        '
        'rbSections
        '
        Me.rbSections.AutoSize = True
        Me.rbSections.Enabled = False
        Me.rbSections.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbSections.Location = New System.Drawing.Point(6, 14)
        Me.rbSections.Name = "rbSections"
        Me.rbSections.Size = New System.Drawing.Size(122, 17)
        Me.rbSections.TabIndex = 0
        Me.rbSections.Text = "Individual Sections"
        Me.rbSections.UseVisualStyleBackColor = True
        '
        'panRBSwb
        '
        Me.panRBSwb.ContextMenuStrip = Me.cmsHome
        Me.panRBSwb.Controls.Add(Me.wbRBS)
        Me.panRBSwb.Location = New System.Drawing.Point(611, 14)
        Me.panRBSwb.Name = "panRBSwb"
        Me.panRBSwb.Size = New System.Drawing.Size(92, 40)
        Me.panRBSwb.TabIndex = 145
        Me.panRBSwb.Visible = False
        '
        'cmsHome
        '
        Me.cmsHome.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.cmiHomeFieldCode})
        Me.cmsHome.Name = "cmsHome"
        Me.cmsHome.Size = New System.Drawing.Size(170, 26)
        '
        'cmiHomeFieldCode
        '
        Me.cmiHomeFieldCode.Name = "cmiHomeFieldCode"
        Me.cmiHomeFieldCode.Size = New System.Drawing.Size(169, 22)
        Me.cmiHomeFieldCode.Text = "Enter Field Code..."
        '
        'wbRBS
        '
        Me.wbRBS.AllowWebBrowserDrop = False
        Me.wbRBS.ContextMenuStrip = Me.cmsHome
        Me.wbRBS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.wbRBS.IsWebBrowserContextMenuEnabled = False
        Me.wbRBS.Location = New System.Drawing.Point(0, 0)
        Me.wbRBS.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbRBS.Name = "wbRBS"
        Me.wbRBS.Size = New System.Drawing.Size(92, 40)
        Me.wbRBS.TabIndex = 0
        Me.wbRBS.WebBrowserShortcutsEnabled = False
        '
        'cmdRBSAll
        '
        Me.cmdRBSAll.BackColor = System.Drawing.Color.Transparent
        Me.cmdRBSAll.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdRBSAll.Enabled = False
        Me.cmdRBSAll.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRBSAll.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdRBSAll.Location = New System.Drawing.Point(237, 44)
        Me.cmdRBSAll.Name = "cmdRBSAll"
        Me.cmdRBSAll.Size = New System.Drawing.Size(61, 21)
        Me.cmdRBSAll.TabIndex = 138
        Me.cmdRBSAll.Text = "Select &All"
        Me.cmdRBSAll.UseVisualStyleBackColor = False
        '
        'grpRBS
        '
        Me.grpRBS.BackColor = System.Drawing.Color.Transparent
        Me.grpRBS.Controls.Add(Me.rbRBS_Section)
        Me.grpRBS.Controls.Add(Me.rbRBS_Col)
        Me.grpRBS.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpRBS.Location = New System.Drawing.Point(416, 3)
        Me.grpRBS.Name = "grpRBS"
        Me.grpRBS.Size = New System.Drawing.Size(195, 57)
        Me.grpRBS.TabIndex = 142
        Me.grpRBS.TabStop = False
        Me.grpRBS.Text = "Show Columns"
        Me.grpRBS.Visible = False
        '
        'rbRBS_Section
        '
        Me.rbRBS_Section.AutoSize = True
        Me.rbRBS_Section.BackColor = System.Drawing.Color.Transparent
        Me.rbRBS_Section.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbRBS_Section.Location = New System.Drawing.Point(6, 15)
        Me.rbRBS_Section.Name = "rbRBS_Section"
        Me.rbRBS_Section.Size = New System.Drawing.Size(192, 17)
        Me.rbRBS_Section.TabIndex = 1
        Me.rbRBS_Section.Text = "Bring Section Name Text to front"
        Me.rbRBS_Section.UseVisualStyleBackColor = False
        '
        'rbRBS_Col
        '
        Me.rbRBS_Col.AutoSize = True
        Me.rbRBS_Col.BackColor = System.Drawing.Color.Transparent
        Me.rbRBS_Col.Checked = True
        Me.rbRBS_Col.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbRBS_Col.Location = New System.Drawing.Point(6, 34)
        Me.rbRBS_Col.Name = "rbRBS_Col"
        Me.rbRBS_Col.Size = New System.Drawing.Size(166, 17)
        Me.rbRBS_Col.TabIndex = 0
        Me.rbRBS_Col.TabStop = True
        Me.rbRBS_Col.Text = "Bring Heading Text to front"
        Me.rbRBS_Col.UseVisualStyleBackColor = False
        '
        'cmdOrderReportBodySection
        '
        Me.cmdOrderReportBodySection.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOrderReportBodySection.BackgroundImage = CType(resources.GetObject("cmdOrderReportBodySection.BackgroundImage"), System.Drawing.Image)
        Me.cmdOrderReportBodySection.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdOrderReportBodySection.Enabled = False
        Me.cmdOrderReportBodySection.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOrderReportBodySection.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdOrderReportBodySection.Location = New System.Drawing.Point(304, 44)
        Me.cmdOrderReportBodySection.Name = "cmdOrderReportBodySection"
        Me.cmdOrderReportBodySection.Size = New System.Drawing.Size(50, 21)
        Me.cmdOrderReportBodySection.TabIndex = 131
        Me.cmdOrderReportBodySection.Text = "Re-#"
        Me.cmdOrderReportBodySection.UseVisualStyleBackColor = False
        '
        'lblRBS
        '
        Me.lblRBS.AutoSize = True
        Me.lblRBS.BackColor = System.Drawing.Color.Transparent
        Me.lblRBS.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRBS.Location = New System.Drawing.Point(5, 25)
        Me.lblRBS.Name = "lblRBS"
        Me.lblRBS.Size = New System.Drawing.Size(45, 13)
        Me.lblRBS.TabIndex = 143
        Me.lblRBS.Text = "Label8"
        '
        'dgvReportStatementWord
        '
        Me.dgvReportStatementWord.AllowUserToAddRows = False
        Me.dgvReportStatementWord.AllowUserToDeleteRows = False
        Me.dgvReportStatementWord.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvReportStatementWord.BackgroundColor = System.Drawing.Color.White
        Me.dgvReportStatementWord.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle174.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle174.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle174.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle174.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle174.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle174.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle174.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatementWord.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle174
        Me.dgvReportStatementWord.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvReportStatementWord.Location = New System.Drawing.Point(540, 96)
        Me.dgvReportStatementWord.MultiSelect = False
        Me.dgvReportStatementWord.Name = "dgvReportStatementWord"
        Me.dgvReportStatementWord.ReadOnly = True
        Me.dgvReportStatementWord.RowHeadersWidth = 25
        Me.dgvReportStatementWord.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvReportStatementWord.Size = New System.Drawing.Size(364, 501)
        Me.dgvReportStatementWord.TabIndex = 141
        '
        'dgvReportStatements
        '
        Me.dgvReportStatements.AllowUserToAddRows = False
        Me.dgvReportStatements.AllowUserToDeleteRows = False
        Me.dgvReportStatements.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvReportStatements.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader
        Me.dgvReportStatements.BackgroundColor = System.Drawing.Color.White
        Me.dgvReportStatements.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle175.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle175.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle175.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle175.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle175.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle175.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle175.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportStatements.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle175
        Me.dgvReportStatements.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvReportStatements.Location = New System.Drawing.Point(5, 66)
        Me.dgvReportStatements.Name = "dgvReportStatements"
        Me.dgvReportStatements.ReadOnly = True
        Me.dgvReportStatements.RowHeadersWidth = 25
        Me.dgvReportStatements.Size = New System.Drawing.Size(532, 531)
        Me.dgvReportStatements.TabIndex = 129
        '
        'lblReportStatement
        '
        Me.lblReportStatement.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblReportStatement.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblReportStatement.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblReportStatement.ForeColor = System.Drawing.Color.White
        Me.lblReportStatement.Location = New System.Drawing.Point(0, 0)
        Me.lblReportStatement.Name = "lblReportStatement"
        Me.lblReportStatement.Size = New System.Drawing.Size(908, 21)
        Me.lblReportStatement.TabIndex = 26
        Me.lblReportStatement.Text = "Choose/Edit Word Template"
        '
        'tp6
        '
        Me.tp6.AutoScroll = True
        Me.tp6.BackColor = System.Drawing.Color.Ivory
        Me.tp6.Controls.Add(Me.cmdShowGroups)
        Me.tp6.Controls.Add(Me.dgvGroups)
        Me.tp6.Controls.Add(Me.lblReportTableConfiguration)
        Me.tp6.Controls.Add(Me.panTableGraphicExamples)
        Me.tp6.Controls.Add(Me.chkReadOnlyTables)
        Me.tp6.Controls.Add(Me.Button2)
        Me.tp6.Controls.Add(Me.cmdDelete)
        Me.tp6.Controls.Add(Me.lblRTCUpDown)
        Me.tp6.Controls.Add(Me.cmdRTCDown)
        Me.tp6.Controls.Add(Me.cmdRTCUp)
        Me.tp6.Controls.Add(Me.chkTableName)
        Me.tp6.Controls.Add(Me.cmdResize)
        Me.tp6.Controls.Add(Me.chkQCShowExcludedBatch)
        Me.tp6.Controls.Add(Me.lblColoredRows)
        Me.tp6.Controls.Add(Me.cmdOrderReportTableConfig)
        Me.tp6.Controls.Add(Me.chkTableGraphicExamples)
        Me.tp6.Controls.Add(Me.dgvReportTableConfiguration)
        Me.tp6.Controls.Add(Me.gbxlblConfigureReportTables1)
        Me.tp6.Location = New System.Drawing.Point(4, 24)
        Me.tp6.Name = "tp6"
        Me.tp6.Size = New System.Drawing.Size(908, 603)
        Me.tp6.TabIndex = 4
        Me.tp6.Text = "6"
        '
        'cmdShowGroups
        '
        Me.cmdShowGroups.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdShowGroups.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowGroups.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdShowGroups.Location = New System.Drawing.Point(413, 144)
        Me.cmdShowGroups.Name = "cmdShowGroups"
        Me.cmdShowGroups.Size = New System.Drawing.Size(124, 26)
        Me.cmdShowGroups.TabIndex = 155
        Me.cmdShowGroups.Text = "Show _C[n] Groups"
        Me.cmdShowGroups.UseVisualStyleBackColor = False
        Me.cmdShowGroups.Visible = False
        '
        'dgvGroups
        '
        Me.dgvGroups.AllowUserToAddRows = False
        Me.dgvGroups.AllowUserToDeleteRows = False
        Me.dgvGroups.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvGroups.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvGroups.BackgroundColor = System.Drawing.Color.White
        Me.dgvGroups.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle176.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle176.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle176.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle176.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle176.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle176.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle176.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGroups.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle176
        Me.dgvGroups.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGroups.Location = New System.Drawing.Point(782, 18)
        Me.dgvGroups.MultiSelect = False
        Me.dgvGroups.Name = "dgvGroups"
        Me.dgvGroups.ReadOnly = True
        DataGridViewCellStyle177.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle177.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle177.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle177.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle177.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle177.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle177.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGroups.RowHeadersDefaultCellStyle = DataGridViewCellStyle177
        Me.dgvGroups.RowHeadersWidth = 25
        DataGridViewCellStyle178.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle178.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvGroups.RowsDefaultCellStyle = DataGridViewCellStyle178
        Me.dgvGroups.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvGroups.Size = New System.Drawing.Size(117, 49)
        Me.dgvGroups.TabIndex = 154
        Me.dgvGroups.Visible = False
        '
        'lblReportTableConfiguration
        '
        Me.lblReportTableConfiguration.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblReportTableConfiguration.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblReportTableConfiguration.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblReportTableConfiguration.ForeColor = System.Drawing.Color.White
        Me.lblReportTableConfiguration.Location = New System.Drawing.Point(0, 0)
        Me.lblReportTableConfiguration.Name = "lblReportTableConfiguration"
        Me.lblReportTableConfiguration.Size = New System.Drawing.Size(908, 21)
        Me.lblReportTableConfiguration.TabIndex = 22
        Me.lblReportTableConfiguration.Text = "Configure Report Tables"
        '
        'panTableGraphicExamples
        '
        Me.panTableGraphicExamples.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panTableGraphicExamples.AutoScroll = True
        Me.panTableGraphicExamples.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.panTableGraphicExamples.Controls.Add(Me.lblTableGraphicExamplesText)
        Me.panTableGraphicExamples.Controls.Add(Me.pbxTableGraphicExamples)
        Me.panTableGraphicExamples.Controls.Add(Me.lblTableGraphicExamplesLabel)
        Me.panTableGraphicExamples.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.panTableGraphicExamples.ForeColor = System.Drawing.Color.Transparent
        Me.panTableGraphicExamples.Location = New System.Drawing.Point(4, 219)
        Me.panTableGraphicExamples.Name = "panTableGraphicExamples"
        Me.panTableGraphicExamples.Size = New System.Drawing.Size(840, 381)
        Me.panTableGraphicExamples.TabIndex = 153
        Me.panTableGraphicExamples.Visible = False
        '
        'lblTableGraphicExamplesText
        '
        Me.lblTableGraphicExamplesText.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTableGraphicExamplesText.ForeColor = System.Drawing.Color.White
        Me.lblTableGraphicExamplesText.Location = New System.Drawing.Point(111, 4)
        Me.lblTableGraphicExamplesText.Name = "lblTableGraphicExamplesText"
        Me.lblTableGraphicExamplesText.Size = New System.Drawing.Size(100, 18)
        Me.lblTableGraphicExamplesText.TabIndex = 152
        Me.lblTableGraphicExamplesText.Text = "<None>"
        '
        'pbxTableGraphicExamples
        '
        Me.pbxTableGraphicExamples.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbxTableGraphicExamples.BackColor = System.Drawing.Color.White
        Me.pbxTableGraphicExamples.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pbxTableGraphicExamples.Location = New System.Drawing.Point(5, 25)
        Me.pbxTableGraphicExamples.Name = "pbxTableGraphicExamples"
        Me.pbxTableGraphicExamples.Size = New System.Drawing.Size(806, 351)
        Me.pbxTableGraphicExamples.TabIndex = 150
        Me.pbxTableGraphicExamples.TabStop = False
        '
        'lblTableGraphicExamplesLabel
        '
        Me.lblTableGraphicExamplesLabel.AutoSize = True
        Me.lblTableGraphicExamplesLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblTableGraphicExamplesLabel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTableGraphicExamplesLabel.ForeColor = System.Drawing.Color.White
        Me.lblTableGraphicExamplesLabel.Location = New System.Drawing.Point(5, 4)
        Me.lblTableGraphicExamplesLabel.Name = "lblTableGraphicExamplesLabel"
        Me.lblTableGraphicExamplesLabel.Size = New System.Drawing.Size(102, 17)
        Me.lblTableGraphicExamplesLabel.TabIndex = 151
        Me.lblTableGraphicExamplesLabel.Text = "Example Table:"
        '
        'chkReadOnlyTables
        '
        Me.chkReadOnlyTables.AutoSize = True
        Me.chkReadOnlyTables.BackColor = System.Drawing.Color.Transparent
        Me.chkReadOnlyTables.Location = New System.Drawing.Point(660, 14)
        Me.chkReadOnlyTables.Name = "chkReadOnlyTables"
        Me.chkReadOnlyTables.Size = New System.Drawing.Size(172, 21)
        Me.chkReadOnlyTables.TabIndex = 149
        Me.chkReadOnlyTables.Text = "Create Read-Only Tables"
        Me.chkReadOnlyTables.UseVisualStyleBackColor = False
        Me.chkReadOnlyTables.Visible = False
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(211, 34)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(38, 21)
        Me.Button2.TabIndex = 147
        Me.Button2.Text = "Button2"
        Me.Button2.UseVisualStyleBackColor = True
        Me.Button2.Visible = False
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdDelete.Enabled = False
        Me.cmdDelete.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDelete.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdDelete.Location = New System.Drawing.Point(851, 339)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(48, 29)
        Me.cmdDelete.TabIndex = 146
        Me.cmdDelete.Text = "D&elete"
        Me.cmdDelete.UseVisualStyleBackColor = True
        Me.cmdDelete.Visible = False
        '
        'lblRTCUpDown
        '
        Me.lblRTCUpDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblRTCUpDown.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblRTCUpDown.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRTCUpDown.ForeColor = System.Drawing.Color.White
        Me.lblRTCUpDown.Location = New System.Drawing.Point(848, 170)
        Me.lblRTCUpDown.Name = "lblRTCUpDown"
        Me.lblRTCUpDown.Size = New System.Drawing.Size(55, 37)
        Me.lblRTCUpDown.TabIndex = 144
        Me.lblRTCUpDown.Text = "Move" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Row"
        Me.lblRTCUpDown.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdRTCDown
        '
        Me.cmdRTCDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRTCDown.BackColor = System.Drawing.Color.Transparent
        Me.cmdRTCDown.BackgroundImage = CType(resources.GetObject("cmdRTCDown.BackgroundImage"), System.Drawing.Image)
        Me.cmdRTCDown.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdRTCDown.Enabled = False
        Me.cmdRTCDown.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdRTCDown.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdRTCDown.Location = New System.Drawing.Point(848, 254)
        Me.cmdRTCDown.Name = "cmdRTCDown"
        Me.cmdRTCDown.Size = New System.Drawing.Size(55, 40)
        Me.cmdRTCDown.TabIndex = 143
        Me.cmdRTCDown.UseVisualStyleBackColor = False
        '
        'cmdRTCUp
        '
        Me.cmdRTCUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdRTCUp.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRTCUp.BackgroundImage = CType(resources.GetObject("cmdRTCUp.BackgroundImage"), System.Drawing.Image)
        Me.cmdRTCUp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdRTCUp.Enabled = False
        Me.cmdRTCUp.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRTCUp.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdRTCUp.Location = New System.Drawing.Point(848, 210)
        Me.cmdRTCUp.Name = "cmdRTCUp"
        Me.cmdRTCUp.Size = New System.Drawing.Size(55, 40)
        Me.cmdRTCUp.TabIndex = 142
        Me.cmdRTCUp.UseVisualStyleBackColor = False
        '
        'chkTableName
        '
        Me.chkTableName.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkTableName.AutoSize = True
        Me.chkTableName.BackColor = System.Drawing.Color.Transparent
        Me.chkTableName.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkTableName.Location = New System.Drawing.Point(452, 99)
        Me.chkTableName.Name = "chkTableName"
        Me.chkTableName.Size = New System.Drawing.Size(192, 21)
        Me.chkTableName.TabIndex = 139
        Me.chkTableName.Text = "Show StudyDoc Table Name"
        Me.chkTableName.UseVisualStyleBackColor = False
        '
        'cmdResize
        '
        Me.cmdResize.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResize.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResize.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdResize.Location = New System.Drawing.Point(630, 104)
        Me.cmdResize.Name = "cmdResize"
        Me.cmdResize.Size = New System.Drawing.Size(113, 26)
        Me.cmdResize.TabIndex = 136
        Me.cmdResize.Text = "&Reset Row Size"
        Me.cmdResize.UseVisualStyleBackColor = False
        '
        'chkQCShowExcludedBatch
        '
        Me.chkQCShowExcludedBatch.AutoSize = True
        Me.chkQCShowExcludedBatch.BackColor = System.Drawing.Color.Transparent
        Me.chkQCShowExcludedBatch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkQCShowExcludedBatch.Location = New System.Drawing.Point(660, 38)
        Me.chkQCShowExcludedBatch.Name = "chkQCShowExcludedBatch"
        Me.chkQCShowExcludedBatch.Size = New System.Drawing.Size(187, 17)
        Me.chkQCShowExcludedBatch.TabIndex = 2
        Me.chkQCShowExcludedBatch.Text = "Show data from excluded batches"
        Me.chkQCShowExcludedBatch.UseVisualStyleBackColor = False
        Me.chkQCShowExcludedBatch.Visible = False
        '
        'lblColoredRows
        '
        Me.lblColoredRows.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblColoredRows.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
        Me.lblColoredRows.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblColoredRows.Location = New System.Drawing.Point(725, 70)
        Me.lblColoredRows.Name = "lblColoredRows"
        Me.lblColoredRows.Size = New System.Drawing.Size(174, 42)
        Me.lblColoredRows.TabIndex = 134
        Me.lblColoredRows.Text = "Colored rows = " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "samples need assigning"
        Me.lblColoredRows.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdOrderReportTableConfig
        '
        Me.cmdOrderReportTableConfig.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOrderReportTableConfig.BackgroundImage = CType(resources.GetObject("cmdOrderReportTableConfig.BackgroundImage"), System.Drawing.Image)
        Me.cmdOrderReportTableConfig.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdOrderReportTableConfig.Enabled = False
        Me.cmdOrderReportTableConfig.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOrderReportTableConfig.Location = New System.Drawing.Point(357, 144)
        Me.cmdOrderReportTableConfig.Name = "cmdOrderReportTableConfig"
        Me.cmdOrderReportTableConfig.Size = New System.Drawing.Size(50, 26)
        Me.cmdOrderReportTableConfig.TabIndex = 132
        Me.cmdOrderReportTableConfig.Text = "Re-#"
        Me.cmdOrderReportTableConfig.UseVisualStyleBackColor = False
        '
        'chkTableGraphicExamples
        '
        Me.chkTableGraphicExamples.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkTableGraphicExamples.AutoSize = True
        Me.chkTableGraphicExamples.BackColor = System.Drawing.Color.Transparent
        Me.chkTableGraphicExamples.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkTableGraphicExamples.Location = New System.Drawing.Point(452, 117)
        Me.chkTableGraphicExamples.Name = "chkTableGraphicExamples"
        Me.chkTableGraphicExamples.Size = New System.Drawing.Size(153, 21)
        Me.chkTableGraphicExamples.TabIndex = 152
        Me.chkTableGraphicExamples.Text = "Show Table Examples"
        Me.chkTableGraphicExamples.UseVisualStyleBackColor = False
        '
        'dgvReportTableConfiguration
        '
        Me.dgvReportTableConfiguration.AllowUserToAddRows = False
        Me.dgvReportTableConfiguration.AllowUserToDeleteRows = False
        Me.dgvReportTableConfiguration.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvReportTableConfiguration.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvReportTableConfiguration.BackgroundColor = System.Drawing.Color.White
        Me.dgvReportTableConfiguration.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
        DataGridViewCellStyle179.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle179.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle179.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle179.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle179.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle179.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle179.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTableConfiguration.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle179
        Me.dgvReportTableConfiguration.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle180.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle180.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle180.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle180.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle180.Padding = New System.Windows.Forms.Padding(0, 10, 0, 10)
        DataGridViewCellStyle180.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle180.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle180.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvReportTableConfiguration.DefaultCellStyle = DataGridViewCellStyle180
        Me.dgvReportTableConfiguration.Location = New System.Drawing.Point(3, 170)
        Me.dgvReportTableConfiguration.Name = "dgvReportTableConfiguration"
        Me.dgvReportTableConfiguration.ReadOnly = True
        Me.dgvReportTableConfiguration.RowHeadersWidth = 25
        DataGridViewCellStyle181.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        DataGridViewCellStyle181.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTableConfiguration.RowsDefaultCellStyle = DataGridViewCellStyle181
        Me.dgvReportTableConfiguration.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTableConfiguration.Size = New System.Drawing.Size(841, 430)
        Me.dgvReportTableConfiguration.TabIndex = 99
        '
        'gbxlblConfigureReportTables1
        '
        Me.gbxlblConfigureReportTables1.Controls.Add(Me.lblRTC)
        Me.gbxlblConfigureReportTables1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblConfigureReportTables1.Location = New System.Drawing.Point(3, 14)
        Me.gbxlblConfigureReportTables1.Margin = New System.Windows.Forms.Padding(0)
        Me.gbxlblConfigureReportTables1.Name = "gbxlblConfigureReportTables1"
        Me.gbxlblConfigureReportTables1.Padding = New System.Windows.Forms.Padding(3, 0, 3, 0)
        Me.gbxlblConfigureReportTables1.Size = New System.Drawing.Size(338, 98)
        Me.gbxlblConfigureReportTables1.TabIndex = 150
        Me.gbxlblConfigureReportTables1.TabStop = False
        '
        'lblRTC
        '
        Me.lblRTC.BackColor = System.Drawing.Color.White
        Me.lblRTC.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblRTC.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRTC.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblRTC.Location = New System.Drawing.Point(3, 18)
        Me.lblRTC.Name = "lblRTC"
        Me.lblRTC.Size = New System.Drawing.Size(332, 74)
        Me.lblRTC.TabIndex = 100
        Me.lblRTC.Text = "* A = If checked, requires sample assignment" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "* B = If checked, table placeholder" & _
    " only will be created" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "* P=Portrait, L=Landscape" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "* Optional FC ID: Used to crea" & _
    "te Field Code"
        '
        'tp7
        '
        Me.tp7.AutoScroll = True
        Me.tp7.BackColor = System.Drawing.Color.Ivory
        Me.tp7.Controls.Add(Me.lbldgvReportTableHeaderConfig)
        Me.tp7.Controls.Add(Me.lblNotes2)
        Me.tp7.Controls.Add(Me.lblNotes1)
        Me.tp7.Controls.Add(Me.gbxlblConfigureColumnHeadings1)
        Me.tp7.Controls.Add(Me.dgvReportTableHeaderConfig)
        Me.tp7.Controls.Add(Me.lbldgvReportTables)
        Me.tp7.Controls.Add(Me.dgvReportTables)
        Me.tp7.Controls.Add(Me.Label33)
        Me.tp7.Controls.Add(Me.Label35)
        Me.tp7.Location = New System.Drawing.Point(4, 24)
        Me.tp7.Name = "tp7"
        Me.tp7.Size = New System.Drawing.Size(908, 603)
        Me.tp7.TabIndex = 9
        Me.tp7.Text = "7"
        '
        'lbldgvReportTableHeaderConfig
        '
        Me.lbldgvReportTableHeaderConfig.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbldgvReportTableHeaderConfig.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lbldgvReportTableHeaderConfig.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbldgvReportTableHeaderConfig.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbldgvReportTableHeaderConfig.ForeColor = System.Drawing.Color.White
        Me.lbldgvReportTableHeaderConfig.Location = New System.Drawing.Point(427, 207)
        Me.lbldgvReportTableHeaderConfig.Name = "lbldgvReportTableHeaderConfig"
        Me.lbldgvReportTableHeaderConfig.Size = New System.Drawing.Size(476, 18)
        Me.lbldgvReportTableHeaderConfig.TabIndex = 100
        Me.lbldgvReportTableHeaderConfig.Text = "Table Column Headers "
        '
        'lblNotes2
        '
        Me.lblNotes2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNotes2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNotes2.Location = New System.Drawing.Point(427, 133)
        Me.lblNotes2.Name = "lblNotes2"
        Me.lblNotes2.Size = New System.Drawing.Size(476, 74)
        Me.lblNotes2.TabIndex = 104
        Me.lblNotes2.Text = "None"
        '
        'lblNotes1
        '
        Me.lblNotes1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblNotes1.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblNotes1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNotes1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotes1.ForeColor = System.Drawing.Color.White
        Me.lblNotes1.Location = New System.Drawing.Point(427, 115)
        Me.lblNotes1.Name = "lblNotes1"
        Me.lblNotes1.Size = New System.Drawing.Size(476, 18)
        Me.lblNotes1.TabIndex = 103
        Me.lblNotes1.Text = "Notes"
        '
        'gbxlblConfigureColumnHeadings1
        '
        Me.gbxlblConfigureColumnHeadings1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.gbxlblConfigureColumnHeadings1.Controls.Add(Me.lblTCH_01)
        Me.gbxlblConfigureColumnHeadings1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblConfigureColumnHeadings1.Location = New System.Drawing.Point(427, 24)
        Me.gbxlblConfigureColumnHeadings1.Name = "gbxlblConfigureColumnHeadings1"
        Me.gbxlblConfigureColumnHeadings1.Padding = New System.Windows.Forms.Padding(0)
        Me.gbxlblConfigureColumnHeadings1.Size = New System.Drawing.Size(464, 85)
        Me.gbxlblConfigureColumnHeadings1.TabIndex = 102
        Me.gbxlblConfigureColumnHeadings1.TabStop = False
        '
        'lblTCH_01
        '
        Me.lblTCH_01.BackColor = System.Drawing.Color.Transparent
        Me.lblTCH_01.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTCH_01.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblTCH_01.Location = New System.Drawing.Point(6, 12)
        Me.lblTCH_01.Name = "lblTCH_01"
        Me.lblTCH_01.Size = New System.Drawing.Size(452, 70)
        Me.lblTCH_01.TabIndex = 96
        Me.lblTCH_01.Text = resources.GetString("lblTCH_01.Text")
        Me.lblTCH_01.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dgvReportTableHeaderConfig
        '
        Me.dgvReportTableHeaderConfig.AllowUserToAddRows = False
        Me.dgvReportTableHeaderConfig.AllowUserToDeleteRows = False
        Me.dgvReportTableHeaderConfig.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvReportTableHeaderConfig.BackgroundColor = System.Drawing.Color.White
        Me.dgvReportTableHeaderConfig.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle182.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle182.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle182.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle182.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle182.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle182.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle182.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTableHeaderConfig.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle182
        Me.dgvReportTableHeaderConfig.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle183.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle183.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle183.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle183.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle183.Padding = New System.Windows.Forms.Padding(0, 6, 0, 6)
        DataGridViewCellStyle183.SelectionBackColor = System.Drawing.Color.DodgerBlue
        DataGridViewCellStyle183.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle183.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvReportTableHeaderConfig.DefaultCellStyle = DataGridViewCellStyle183
        Me.dgvReportTableHeaderConfig.Location = New System.Drawing.Point(427, 225)
        Me.dgvReportTableHeaderConfig.Name = "dgvReportTableHeaderConfig"
        Me.dgvReportTableHeaderConfig.ReadOnly = True
        Me.dgvReportTableHeaderConfig.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvReportTableHeaderConfig.Size = New System.Drawing.Size(476, 375)
        Me.dgvReportTableHeaderConfig.TabIndex = 101
        '
        'lbldgvReportTables
        '
        Me.lbldgvReportTables.AutoSize = True
        Me.lbldgvReportTables.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lbldgvReportTables.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbldgvReportTables.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbldgvReportTables.ForeColor = System.Drawing.Color.White
        Me.lbldgvReportTables.Location = New System.Drawing.Point(3, 115)
        Me.lbldgvReportTables.MinimumSize = New System.Drawing.Size(422, 0)
        Me.lbldgvReportTables.Name = "lbldgvReportTables"
        Me.lbldgvReportTables.Size = New System.Drawing.Size(422, 18)
        Me.lbldgvReportTables.TabIndex = 99
        Me.lbldgvReportTables.Text = "Report Tables *"
        '
        'dgvReportTables
        '
        Me.dgvReportTables.AllowUserToAddRows = False
        Me.dgvReportTables.AllowUserToDeleteRows = False
        Me.dgvReportTables.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.dgvReportTables.BackgroundColor = System.Drawing.Color.White
        Me.dgvReportTables.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle184.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle184.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle184.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle184.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle184.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle184.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle184.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvReportTables.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle184
        Me.dgvReportTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle185.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle185.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle185.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle185.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle185.Padding = New System.Windows.Forms.Padding(0, 6, 0, 6)
        DataGridViewCellStyle185.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle185.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle185.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvReportTables.DefaultCellStyle = DataGridViewCellStyle185
        Me.dgvReportTables.Location = New System.Drawing.Point(3, 133)
        Me.dgvReportTables.Name = "dgvReportTables"
        Me.dgvReportTables.ReadOnly = True
        Me.dgvReportTables.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvReportTables.Size = New System.Drawing.Size(422, 467)
        Me.dgvReportTables.TabIndex = 98
        '
        'Label33
        '
        Me.Label33.BackColor = System.Drawing.Color.Transparent
        Me.Label33.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label33.Location = New System.Drawing.Point(3, 27)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(363, 46)
        Me.Label33.TabIndex = 97
        Me.Label33.Text = "Please note that the QA Events Table Rows (Critical Phases) must have configured " & _
    "a User Label named 'Final Report'"
        Me.Label33.Visible = False
        '
        'Label35
        '
        Me.Label35.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label35.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label35.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label35.ForeColor = System.Drawing.Color.White
        Me.Label35.Location = New System.Drawing.Point(0, 0)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(908, 21)
        Me.Label35.TabIndex = 28
        Me.Label35.Text = "Configure Column Headings"
        '
        'tp8
        '
        Me.tp8.AutoScroll = True
        Me.tp8.BackColor = System.Drawing.Color.Ivory
        Me.tp8.Controls.Add(Me.Label9)
        Me.tp8.Controls.Add(Me.gbxlblAnalyticalReferenceStd)
        Me.tp8.Controls.Add(Me.dgvCompanyAnalRef)
        Me.tp8.Controls.Add(Me.dgvWatsonAnalRef)
        Me.tp8.Controls.Add(Me.lblWAR)
        Me.tp8.Controls.Add(Me.lblARST)
        Me.tp8.Location = New System.Drawing.Point(4, 24)
        Me.tp8.Name = "tp8"
        Me.tp8.Size = New System.Drawing.Size(908, 603)
        Me.tp8.TabIndex = 5
        Me.tp8.Text = "8"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label9.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label9.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label9.ForeColor = System.Drawing.Color.White
        Me.Label9.Location = New System.Drawing.Point(0, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(908, 21)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Analytical Reference Std"
        '
        'gbxlblAnalyticalReferenceStd
        '
        Me.gbxlblAnalyticalReferenceStd.Controls.Add(Me.lblARS)
        Me.gbxlblAnalyticalReferenceStd.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblAnalyticalReferenceStd.Location = New System.Drawing.Point(8, 21)
        Me.gbxlblAnalyticalReferenceStd.Margin = New System.Windows.Forms.Padding(0)
        Me.gbxlblAnalyticalReferenceStd.Name = "gbxlblAnalyticalReferenceStd"
        Me.gbxlblAnalyticalReferenceStd.Size = New System.Drawing.Size(201, 43)
        Me.gbxlblAnalyticalReferenceStd.TabIndex = 137
        Me.gbxlblAnalyticalReferenceStd.TabStop = False
        '
        'lblARS
        '
        Me.lblARS.AutoSize = True
        Me.lblARS.BackColor = System.Drawing.Color.White
        Me.lblARS.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblARS.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblARS.Location = New System.Drawing.Point(3, 6)
        Me.lblARS.Margin = New System.Windows.Forms.Padding(0)
        Me.lblARS.Name = "lblARS"
        Me.lblARS.Size = New System.Drawing.Size(188, 34)
        Me.lblARS.TabIndex = 135
        Me.lblARS.Text = "? = Enter Yes or No " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "A* = Check to include in report"
        Me.lblARS.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'dgvCompanyAnalRef
        '
        Me.dgvCompanyAnalRef.AllowUserToAddRows = False
        Me.dgvCompanyAnalRef.AllowUserToDeleteRows = False
        Me.dgvCompanyAnalRef.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvCompanyAnalRef.BackgroundColor = System.Drawing.Color.White
        Me.dgvCompanyAnalRef.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCompanyAnalRef.Location = New System.Drawing.Point(8, 85)
        Me.dgvCompanyAnalRef.Name = "dgvCompanyAnalRef"
        Me.dgvCompanyAnalRef.Size = New System.Drawing.Size(886, 232)
        Me.dgvCompanyAnalRef.TabIndex = 136
        '
        'dgvWatsonAnalRef
        '
        Me.dgvWatsonAnalRef.AllowUserToAddRows = False
        Me.dgvWatsonAnalRef.AllowUserToDeleteRows = False
        Me.dgvWatsonAnalRef.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvWatsonAnalRef.BackgroundColor = System.Drawing.Color.White
        Me.dgvWatsonAnalRef.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle186.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle186.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle186.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle186.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle186.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle186.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle186.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvWatsonAnalRef.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle186
        Me.dgvWatsonAnalRef.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvWatsonAnalRef.Location = New System.Drawing.Point(9, 345)
        Me.dgvWatsonAnalRef.Name = "dgvWatsonAnalRef"
        Me.dgvWatsonAnalRef.ReadOnly = True
        Me.dgvWatsonAnalRef.RowHeadersWidth = 25
        DataGridViewCellStyle187.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvWatsonAnalRef.RowsDefaultCellStyle = DataGridViewCellStyle187
        Me.dgvWatsonAnalRef.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvWatsonAnalRef.Size = New System.Drawing.Size(885, 241)
        Me.dgvWatsonAnalRef.TabIndex = 104
        '
        'lblWAR
        '
        Me.lblWAR.AutoSize = True
        Me.lblWAR.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblWAR.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWAR.ForeColor = System.Drawing.Color.White
        Me.lblWAR.Location = New System.Drawing.Point(9, 326)
        Me.lblWAR.Name = "lblWAR"
        Me.lblWAR.Size = New System.Drawing.Size(751, 17)
        Me.lblWAR.TabIndex = 29
        Me.lblWAR.Text = "Watson Analytical Reference Standard Table (Note: Accuracies and Precisions perti" & _
    "nent to Sample Analysis studies only)"
        Me.lblWAR.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblARST
        '
        Me.lblARST.AutoSize = True
        Me.lblARST.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblARST.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblARST.ForeColor = System.Drawing.Color.White
        Me.lblARST.Location = New System.Drawing.Point(9, 66)
        Me.lblARST.Name = "lblARST"
        Me.lblARST.Size = New System.Drawing.Size(234, 17)
        Me.lblARST.TabIndex = 31
        Me.lblARST.Text = "Analytical Reference Standard Table "
        Me.lblARST.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tp9
        '
        Me.tp9.AutoScroll = True
        Me.tp9.BackColor = System.Drawing.Color.Ivory
        Me.tp9.Controls.Add(Me.lblGlobalConfiguration)
        Me.tp9.Controls.Add(Me.gbxlblAddEditContributors)
        Me.tp9.Controls.Add(Me.dgvContributingPersonnel)
        Me.tp9.Location = New System.Drawing.Point(4, 24)
        Me.tp9.Name = "tp9"
        Me.tp9.Size = New System.Drawing.Size(908, 603)
        Me.tp9.TabIndex = 6
        Me.tp9.Text = "9"
        '
        'lblGlobalConfiguration
        '
        Me.lblGlobalConfiguration.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblGlobalConfiguration.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblGlobalConfiguration.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblGlobalConfiguration.ForeColor = System.Drawing.Color.White
        Me.lblGlobalConfiguration.Location = New System.Drawing.Point(0, 0)
        Me.lblGlobalConfiguration.Name = "lblGlobalConfiguration"
        Me.lblGlobalConfiguration.Size = New System.Drawing.Size(908, 21)
        Me.lblGlobalConfiguration.TabIndex = 23
        Me.lblGlobalConfiguration.Text = "Add/Edit Contributors "
        '
        'gbxlblAddEditContributors
        '
        Me.gbxlblAddEditContributors.Controls.Add(Me.Label34)
        Me.gbxlblAddEditContributors.Controls.Add(Me.Label53)
        Me.gbxlblAddEditContributors.Controls.Add(Me.Label18)
        Me.gbxlblAddEditContributors.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblAddEditContributors.Location = New System.Drawing.Point(5, 18)
        Me.gbxlblAddEditContributors.Name = "gbxlblAddEditContributors"
        Me.gbxlblAddEditContributors.Size = New System.Drawing.Size(392, 67)
        Me.gbxlblAddEditContributors.TabIndex = 107
        Me.gbxlblAddEditContributors.TabStop = False
        '
        'Label34
        '
        Me.Label34.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label34.AutoSize = True
        Me.Label34.BackColor = System.Drawing.Color.Transparent
        Me.Label34.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label34.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label34.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label34.Location = New System.Drawing.Point(0, 26)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(366, 17)
        Me.Label34.TabIndex = 53
        Me.Label34.Text = "* B: Order in which to display on Contributing Personnel Page"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label53
        '
        Me.Label53.AutoSize = True
        Me.Label53.BackColor = System.Drawing.Color.Transparent
        Me.Label53.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label53.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label53.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label53.Location = New System.Drawing.Point(0, 45)
        Me.Label53.Name = "Label53"
        Me.Label53.Size = New System.Drawing.Size(106, 17)
        Me.Label53.TabIndex = 105
        Me.Label53.Text = "** Required Field"
        Me.Label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label18
        '
        Me.Label18.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.Color.Transparent
        Me.Label18.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label18.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.Label18.Location = New System.Drawing.Point(0, 7)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(268, 17)
        Me.Label18.TabIndex = 106
        Me.Label18.Text = "* A: Include on Contributing Personnel Page?"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'dgvContributingPersonnel
        '
        Me.dgvContributingPersonnel.AllowUserToAddRows = False
        Me.dgvContributingPersonnel.AllowUserToDeleteRows = False
        Me.dgvContributingPersonnel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvContributingPersonnel.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvContributingPersonnel.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders
        Me.dgvContributingPersonnel.BackgroundColor = System.Drawing.Color.White
        Me.dgvContributingPersonnel.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle188.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle188.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle188.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle188.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle188.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle188.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle188.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvContributingPersonnel.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle188
        Me.dgvContributingPersonnel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle189.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle189.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle189.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle189.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle189.Padding = New System.Windows.Forms.Padding(0, 6, 0, 6)
        DataGridViewCellStyle189.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle189.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle189.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvContributingPersonnel.DefaultCellStyle = DataGridViewCellStyle189
        Me.dgvContributingPersonnel.Location = New System.Drawing.Point(5, 88)
        Me.dgvContributingPersonnel.Name = "dgvContributingPersonnel"
        Me.dgvContributingPersonnel.ReadOnly = True
        Me.dgvContributingPersonnel.RowHeadersWidth = 25
        DataGridViewCellStyle190.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvContributingPersonnel.RowsDefaultCellStyle = DataGridViewCellStyle190
        Me.dgvContributingPersonnel.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvContributingPersonnel.Size = New System.Drawing.Size(900, 488)
        Me.dgvContributingPersonnel.TabIndex = 104
        '
        'tp10
        '
        Me.tp10.AutoScroll = True
        Me.tp10.BackColor = System.Drawing.Color.Ivory
        Me.tp10.Controls.Add(Me.gbxlblReviewValidatedMethod)
        Me.tp10.Controls.Add(Me.dgvMethodValData)
        Me.tp10.Controls.Add(Me.gbMethValApplyGuWu)
        Me.tp10.Controls.Add(Me.Label17)
        Me.tp10.Location = New System.Drawing.Point(4, 24)
        Me.tp10.Name = "tp10"
        Me.tp10.Size = New System.Drawing.Size(908, 603)
        Me.tp10.TabIndex = 8
        Me.tp10.Text = "10"
        '
        'gbxlblReviewValidatedMethod
        '
        Me.gbxlblReviewValidatedMethod.AutoSize = True
        Me.gbxlblReviewValidatedMethod.Controls.Add(Me.lbl1)
        Me.gbxlblReviewValidatedMethod.Controls.Add(Me.lbl2)
        Me.gbxlblReviewValidatedMethod.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblReviewValidatedMethod.Location = New System.Drawing.Point(9, 21)
        Me.gbxlblReviewValidatedMethod.Name = "gbxlblReviewValidatedMethod"
        Me.gbxlblReviewValidatedMethod.Padding = New System.Windows.Forms.Padding(0)
        Me.gbxlblReviewValidatedMethod.Size = New System.Drawing.Size(531, 92)
        Me.gbxlblReviewValidatedMethod.TabIndex = 46
        Me.gbxlblReviewValidatedMethod.TabStop = False
        '
        'lbl1
        '
        Me.lbl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lbl1.Location = New System.Drawing.Point(3, 37)
        Me.lbl1.Name = "lbl1"
        Me.lbl1.Size = New System.Drawing.Size(525, 37)
        Me.lbl1.TabIndex = 44
        Me.lbl1.Text = "MethVal"
        Me.lbl1.Visible = False
        '
        'lbl2
        '
        Me.lbl2.AutoSize = True
        Me.lbl2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.lbl2.Location = New System.Drawing.Point(3, 13)
        Me.lbl2.Name = "lbl2"
        Me.lbl2.Size = New System.Drawing.Size(120, 16)
        Me.lbl2.TabIndex = 45
        Me.lbl2.Text = "SampleAnalysis"
        Me.lbl2.Visible = False
        '
        'dgvMethodValData
        '
        Me.dgvMethodValData.AllowUserToAddRows = False
        Me.dgvMethodValData.AllowUserToDeleteRows = False
        Me.dgvMethodValData.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMethodValData.BackgroundColor = System.Drawing.Color.White
        Me.dgvMethodValData.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.dgvMethodValData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMethodValData.Location = New System.Drawing.Point(9, 118)
        Me.dgvMethodValData.Name = "dgvMethodValData"
        Me.dgvMethodValData.Size = New System.Drawing.Size(851, 317)
        Me.dgvMethodValData.TabIndex = 43
        '
        'gbMethValApplyGuWu
        '
        Me.gbMethValApplyGuWu.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbMethValApplyGuWu.BackColor = System.Drawing.Color.Transparent
        Me.gbMethValApplyGuWu.Controls.Add(Me.dgvMethValExistingGuWu)
        Me.gbMethValApplyGuWu.Controls.Add(Me.cbxArchivedMDB)
        Me.gbMethValApplyGuWu.Controls.Add(Me.GroupBox1)
        Me.gbMethValApplyGuWu.Controls.Add(Me.cmdBrowseMDB)
        Me.gbMethValApplyGuWu.Controls.Add(Me.gbMethodValMultiple)
        Me.gbMethValApplyGuWu.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbMethValApplyGuWu.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.gbMethValApplyGuWu.Location = New System.Drawing.Point(11, 439)
        Me.gbMethValApplyGuWu.Name = "gbMethValApplyGuWu"
        Me.gbMethValApplyGuWu.Size = New System.Drawing.Size(849, 158)
        Me.gbMethValApplyGuWu.TabIndex = 42
        Me.gbMethValApplyGuWu.TabStop = False
        Me.gbMethValApplyGuWu.Text = "Apply Existing StudyDoc Information to columns above, if desired"
        '
        'dgvMethValExistingGuWu
        '
        Me.dgvMethValExistingGuWu.AllowUserToAddRows = False
        Me.dgvMethValExistingGuWu.AllowUserToDeleteRows = False
        Me.dgvMethValExistingGuWu.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvMethValExistingGuWu.BackgroundColor = System.Drawing.Color.White
        Me.dgvMethValExistingGuWu.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Me.dgvMethValExistingGuWu.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvMethValExistingGuWu.Location = New System.Drawing.Point(344, 20)
        Me.dgvMethValExistingGuWu.Name = "dgvMethValExistingGuWu"
        Me.dgvMethValExistingGuWu.ReadOnly = True
        Me.dgvMethValExistingGuWu.Size = New System.Drawing.Size(499, 133)
        Me.dgvMethValExistingGuWu.TabIndex = 47
        '
        'cbxArchivedMDB
        '
        Me.cbxArchivedMDB.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxArchivedMDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxArchivedMDB.IntegralHeight = False
        Me.cbxArchivedMDB.Location = New System.Drawing.Point(203, 47)
        Me.cbxArchivedMDB.Name = "cbxArchivedMDB"
        Me.cbxArchivedMDB.Size = New System.Drawing.Size(58, 21)
        Me.cbxArchivedMDB.Sorted = True
        Me.cbxArchivedMDB.TabIndex = 43
        Me.cbxArchivedMDB.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmdMethValExecute)
        Me.GroupBox1.Controls.Add(Me.cbxMethValExistingGuWu)
        Me.GroupBox1.Controls.Add(Me.lblM2)
        Me.GroupBox1.Controls.Add(Me.Label31)
        Me.GroupBox1.Controls.Add(Me.Label21)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.GroupBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.GroupBox1.Location = New System.Drawing.Point(0, 22)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(341, 136)
        Me.GroupBox1.TabIndex = 46
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Follow these steps:"
        '
        'cmdMethValExecute
        '
        Me.cmdMethValExecute.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdMethValExecute.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdMethValExecute.Enabled = False
        Me.cmdMethValExecute.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdMethValExecute.Location = New System.Drawing.Point(255, 15)
        Me.cmdMethValExecute.Name = "cmdMethValExecute"
        Me.cmdMethValExecute.Size = New System.Drawing.Size(80, 25)
        Me.cmdMethValExecute.TabIndex = 6
        Me.cmdMethValExecute.Text = "E&xecute"
        Me.cmdMethValExecute.UseVisualStyleBackColor = False
        '
        'cbxMethValExistingGuWu
        '
        Me.cbxMethValExistingGuWu.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxMethValExistingGuWu.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxMethValExistingGuWu.IntegralHeight = False
        Me.cbxMethValExistingGuWu.Location = New System.Drawing.Point(7, 80)
        Me.cbxMethValExistingGuWu.Name = "cbxMethValExistingGuWu"
        Me.cbxMethValExistingGuWu.Size = New System.Drawing.Size(328, 21)
        Me.cbxMethValExistingGuWu.TabIndex = 2
        '
        'lblM2
        '
        Me.lblM2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblM2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblM2.Location = New System.Drawing.Point(6, 43)
        Me.lblM2.Name = "lblM2"
        Me.lblM2.Size = New System.Drawing.Size(329, 37)
        Me.lblM2.TabIndex = 4
        Me.lblM2.Text = "2. Retrieve information from an Existing StudyDoc-configured Watson Study by choo" & _
    "sing a study from the dropdown box below"
        '
        'Label31
        '
        Me.Label31.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label31.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Label31.Location = New System.Drawing.Point(6, 19)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(235, 18)
        Me.Label31.TabIndex = 5
        Me.Label31.Text = "1. Choose a table item(s) --------------->"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Label21.Location = New System.Drawing.Point(6, 106)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(173, 16)
        Me.Label21.TabIndex = 42
        Me.Label21.Text = "3. Click 'Execute'"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'cmdBrowseMDB
        '
        Me.cmdBrowseMDB.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdBrowseMDB.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseMDB.Enabled = False
        Me.cmdBrowseMDB.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdBrowseMDB.Location = New System.Drawing.Point(212, 63)
        Me.cmdBrowseMDB.Name = "cmdBrowseMDB"
        Me.cmdBrowseMDB.Size = New System.Drawing.Size(68, 25)
        Me.cmdBrowseMDB.TabIndex = 44
        Me.cmdBrowseMDB.Text = "&Browse..."
        Me.cmdBrowseMDB.UseVisualStyleBackColor = False
        Me.cmdBrowseMDB.Visible = False
        '
        'gbMethodValMultiple
        '
        Me.gbMethodValMultiple.BackColor = System.Drawing.Color.Transparent
        Me.gbMethodValMultiple.Controls.Add(Me.lblMethValMultiple)
        Me.gbMethodValMultiple.Controls.Add(Me.txtMethValMultiple)
        Me.gbMethodValMultiple.Controls.Add(Me.rbMethValMultiple)
        Me.gbMethodValMultiple.Controls.Add(Me.rbMethValAnalyte)
        Me.gbMethodValMultiple.Controls.Add(Me.chkMethodValMultiple)
        Me.gbMethodValMultiple.Enabled = False
        Me.gbMethodValMultiple.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbMethodValMultiple.Location = New System.Drawing.Point(516, 0)
        Me.gbMethodValMultiple.Name = "gbMethodValMultiple"
        Me.gbMethodValMultiple.Size = New System.Drawing.Size(252, 155)
        Me.gbMethodValMultiple.TabIndex = 40
        Me.gbMethodValMultiple.TabStop = False
        Me.gbMethodValMultiple.Text = "Configure Multiple Method Validation References"
        Me.gbMethodValMultiple.Visible = False
        '
        'lblMethValMultiple
        '
        Me.lblMethValMultiple.Location = New System.Drawing.Point(70, 114)
        Me.lblMethValMultiple.Name = "lblMethValMultiple"
        Me.lblMethValMultiple.Size = New System.Drawing.Size(238, 20)
        Me.lblMethValMultiple.TabIndex = 4
        Me.lblMethValMultiple.Text = "Enter number of method validation references"
        Me.lblMethValMultiple.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtMethValMultiple
        '
        Me.txtMethValMultiple.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtMethValMultiple.Enabled = False
        Me.txtMethValMultiple.Location = New System.Drawing.Point(35, 114)
        Me.txtMethValMultiple.Name = "txtMethValMultiple"
        Me.txtMethValMultiple.Size = New System.Drawing.Size(33, 20)
        Me.txtMethValMultiple.TabIndex = 3
        Me.txtMethValMultiple.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'rbMethValMultiple
        '
        Me.rbMethValMultiple.Enabled = False
        Me.rbMethValMultiple.Location = New System.Drawing.Point(15, 89)
        Me.rbMethValMultiple.Name = "rbMethValMultiple"
        Me.rbMethValMultiple.Size = New System.Drawing.Size(403, 20)
        Me.rbMethValMultiple.TabIndex = 2
        Me.rbMethValMultiple.Text = "Configure multiple method validation references not related to analytes"
        '
        'rbMethValAnalyte
        '
        Me.rbMethValAnalyte.Checked = True
        Me.rbMethValAnalyte.Enabled = False
        Me.rbMethValAnalyte.Location = New System.Drawing.Point(15, 69)
        Me.rbMethValAnalyte.Name = "rbMethValAnalyte"
        Me.rbMethValAnalyte.Size = New System.Drawing.Size(318, 20)
        Me.rbMethValAnalyte.TabIndex = 1
        Me.rbMethValAnalyte.TabStop = True
        Me.rbMethValAnalyte.Text = "Configure one method reference for each analyte"
        '
        'chkMethodValMultiple
        '
        Me.chkMethodValMultiple.Location = New System.Drawing.Point(15, 44)
        Me.chkMethodValMultiple.Name = "chkMethodValMultiple"
        Me.chkMethodValMultiple.Size = New System.Drawing.Size(318, 20)
        Me.chkMethodValMultiple.TabIndex = 0
        Me.chkMethodValMultiple.Text = "Configure multiple method validation references"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label17.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label17.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label17.ForeColor = System.Drawing.Color.White
        Me.Label17.Location = New System.Drawing.Point(0, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(908, 21)
        Me.Label17.TabIndex = 27
        Me.Label17.Text = "Review Validated Method"
        '
        'tp11
        '
        Me.tp11.AutoScroll = True
        Me.tp11.BackColor = System.Drawing.Color.Ivory
        Me.tp11.Controls.Add(Me.chkQAEventBorder)
        Me.tp11.Controls.Add(Me.lblQAHyperlink)
        Me.tp11.Controls.Add(Me.dgQATable)
        Me.tp11.Controls.Add(Me.Label32)
        Me.tp11.Location = New System.Drawing.Point(4, 24)
        Me.tp11.Name = "tp11"
        Me.tp11.Size = New System.Drawing.Size(908, 603)
        Me.tp11.TabIndex = 10
        Me.tp11.Text = "11"
        '
        'chkQAEventBorder
        '
        Me.chkQAEventBorder.AutoSize = True
        Me.chkQAEventBorder.BackColor = System.Drawing.Color.Transparent
        Me.chkQAEventBorder.Location = New System.Drawing.Point(280, 72)
        Me.chkQAEventBorder.Name = "chkQAEventBorder"
        Me.chkQAEventBorder.Size = New System.Drawing.Size(262, 21)
        Me.chkQAEventBorder.TabIndex = 104
        Me.chkQAEventBorder.Text = "Border the QA Event Table in the Report"
        Me.chkQAEventBorder.UseVisualStyleBackColor = False
        '
        'lblQAHyperlink
        '
        Me.lblQAHyperlink.BackColor = System.Drawing.Color.Transparent
        Me.lblQAHyperlink.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblQAHyperlink.Location = New System.Drawing.Point(277, 24)
        Me.lblQAHyperlink.Name = "lblQAHyperlink"
        Me.lblQAHyperlink.Size = New System.Drawing.Size(356, 36)
        Me.lblQAHyperlink.TabIndex = 103
        Me.lblQAHyperlink.TabStop = True
        Me.lblQAHyperlink.Text = "Click to Configure Column Heading and Critical Phase text using the Report Table " & _
    "Header Configuration Page"
        Me.lblQAHyperlink.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'dgQATable
        '
        Me.dgQATable.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgQATable.BackgroundColor = System.Drawing.Color.White
        Me.dgQATable.CaptionBackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.dgQATable.CaptionForeColor = System.Drawing.Color.White
        Me.dgQATable.DataMember = ""
        Me.dgQATable.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgQATable.Location = New System.Drawing.Point(9, 97)
        Me.dgQATable.Name = "dgQATable"
        Me.dgQATable.ReadOnly = True
        Me.dgQATable.Size = New System.Drawing.Size(896, 500)
        Me.dgQATable.TabIndex = 28
        '
        'Label32
        '
        Me.Label32.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label32.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label32.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label32.ForeColor = System.Drawing.Color.White
        Me.Label32.Location = New System.Drawing.Point(0, 0)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(908, 21)
        Me.Label32.TabIndex = 27
        Me.Label32.Text = "QA Event Table"
        '
        'tp12
        '
        Me.tp12.AutoScroll = True
        Me.tp12.BackColor = System.Drawing.Color.Ivory
        Me.tp12.Controls.Add(Me.gbxlblSampleReceiptRecords2)
        Me.tp12.Controls.Add(Me.gbxlblSampleReceiptRecords1)
        Me.tp12.Controls.Add(Me.Label24)
        Me.tp12.Controls.Add(Me.Label47)
        Me.tp12.Controls.Add(Me.txtSRecTotalReportWatson)
        Me.tp12.Controls.Add(Me.chkUseWatsonSampleNumber)
        Me.tp12.Controls.Add(Me.chkManualSampleNumber)
        Me.tp12.Controls.Add(Me.Label46)
        Me.tp12.Controls.Add(Me.dgvSampleReceiptWatson)
        Me.tp12.Controls.Add(Me.Label41)
        Me.tp12.Controls.Add(Me.txtSRecTotalReport)
        Me.tp12.Controls.Add(Me.Label40)
        Me.tp12.Controls.Add(Me.txtSRecTotal)
        Me.tp12.Controls.Add(Me.dgvSampleReceipt)
        Me.tp12.Controls.Add(Me.Label37)
        Me.tp12.Location = New System.Drawing.Point(4, 24)
        Me.tp12.Name = "tp12"
        Me.tp12.Size = New System.Drawing.Size(908, 603)
        Me.tp12.TabIndex = 12
        Me.tp12.Text = "12"
        '
        'gbxlblSampleReceiptRecords2
        '
        Me.gbxlblSampleReceiptRecords2.Controls.Add(Me.Label48)
        Me.gbxlblSampleReceiptRecords2.Font = New System.Drawing.Font("Segoe UI", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblSampleReceiptRecords2.Location = New System.Drawing.Point(565, 51)
        Me.gbxlblSampleReceiptRecords2.Name = "gbxlblSampleReceiptRecords2"
        Me.gbxlblSampleReceiptRecords2.Size = New System.Drawing.Size(251, 49)
        Me.gbxlblSampleReceiptRecords2.TabIndex = 114
        Me.gbxlblSampleReceiptRecords2.TabStop = False
        '
        'Label48
        '
        Me.Label48.AutoSize = True
        Me.Label48.BackColor = System.Drawing.Color.Transparent
        Me.Label48.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label48.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label48.Location = New System.Drawing.Point(6, 9)
        Me.Label48.Name = "Label48"
        Me.Label48.Size = New System.Drawing.Size(235, 34)
        Me.Label48.TabIndex = 108
        Me.Label48.Text = "** NOTE: If checked, this number takes " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "precedence over Watson total"
        Me.Label48.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'gbxlblSampleReceiptRecords1
        '
        Me.gbxlblSampleReceiptRecords1.Controls.Add(Me.Label39)
        Me.gbxlblSampleReceiptRecords1.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxlblSampleReceiptRecords1.Location = New System.Drawing.Point(3, 24)
        Me.gbxlblSampleReceiptRecords1.Name = "gbxlblSampleReceiptRecords1"
        Me.gbxlblSampleReceiptRecords1.Size = New System.Drawing.Size(263, 65)
        Me.gbxlblSampleReceiptRecords1.TabIndex = 113
        Me.gbxlblSampleReceiptRecords1.TabStop = False
        '
        'Label39
        '
        Me.Label39.BackColor = System.Drawing.Color.Transparent
        Me.Label39.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label39.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label39.Location = New System.Drawing.Point(2, 5)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(255, 57)
        Me.Label39.TabIndex = 97
        Me.Label39.Text = "* A:  Use these sample counts to calculate total number of samples received to be" & _
    " used in Report."
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label24.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.White
        Me.Label24.Location = New System.Drawing.Point(3, 101)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(266, 16)
        Me.Label24.TabIndex = 112
        Me.Label24.Text = "StudyDoc Sample Receipt Records"
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.BackColor = System.Drawing.Color.Transparent
        Me.Label47.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label47.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Label47.Location = New System.Drawing.Point(347, 344)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(431, 16)
        Me.Label47.TabIndex = 111
        Me.Label47.Text = ":Total number of samples received to be used in Report (read only)"
        '
        'txtSRecTotalReportWatson
        '
        Me.txtSRecTotalReportWatson.BackColor = System.Drawing.Color.White
        Me.txtSRecTotalReportWatson.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSRecTotalReportWatson.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSRecTotalReportWatson.Location = New System.Drawing.Point(271, 344)
        Me.txtSRecTotalReportWatson.Multiline = True
        Me.txtSRecTotalReportWatson.Name = "txtSRecTotalReportWatson"
        Me.txtSRecTotalReportWatson.ReadOnly = True
        Me.txtSRecTotalReportWatson.Size = New System.Drawing.Size(72, 36)
        Me.txtSRecTotalReportWatson.TabIndex = 110
        Me.txtSRecTotalReportWatson.Text = "0"
        Me.txtSRecTotalReportWatson.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'chkUseWatsonSampleNumber
        '
        Me.chkUseWatsonSampleNumber.AutoSize = True
        Me.chkUseWatsonSampleNumber.BackColor = System.Drawing.Color.Transparent
        Me.chkUseWatsonSampleNumber.Enabled = False
        Me.chkUseWatsonSampleNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUseWatsonSampleNumber.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.chkUseWatsonSampleNumber.Location = New System.Drawing.Point(349, 360)
        Me.chkUseWatsonSampleNumber.Name = "chkUseWatsonSampleNumber"
        Me.chkUseWatsonSampleNumber.Size = New System.Drawing.Size(233, 20)
        Me.chkUseWatsonSampleNumber.TabIndex = 109
        Me.chkUseWatsonSampleNumber.Text = "Use Watson Sample Receipt Data"
        Me.chkUseWatsonSampleNumber.UseVisualStyleBackColor = False
        '
        'chkManualSampleNumber
        '
        Me.chkManualSampleNumber.AutoSize = True
        Me.chkManualSampleNumber.BackColor = System.Drawing.Color.Transparent
        Me.chkManualSampleNumber.Enabled = False
        Me.chkManualSampleNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkManualSampleNumber.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.chkManualSampleNumber.Location = New System.Drawing.Point(565, 29)
        Me.chkManualSampleNumber.Name = "chkManualSampleNumber"
        Me.chkManualSampleNumber.Size = New System.Drawing.Size(224, 20)
        Me.chkManualSampleNumber.TabIndex = 104
        Me.chkManualSampleNumber.Text = "Enter sample number manually **"
        Me.chkManualSampleNumber.UseVisualStyleBackColor = False
        '
        'Label46
        '
        Me.Label46.AutoSize = True
        Me.Label46.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label46.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label46.ForeColor = System.Drawing.Color.White
        Me.Label46.Location = New System.Drawing.Point(3, 365)
        Me.Label46.Name = "Label46"
        Me.Label46.Size = New System.Drawing.Size(215, 16)
        Me.Label46.TabIndex = 103
        Me.Label46.Text = "Watson Sample Receipt Records"
        '
        'dgvSampleReceiptWatson
        '
        Me.dgvSampleReceiptWatson.AllowUserToAddRows = False
        Me.dgvSampleReceiptWatson.AllowUserToDeleteRows = False
        Me.dgvSampleReceiptWatson.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSampleReceiptWatson.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSampleReceiptWatson.BackgroundColor = System.Drawing.Color.White
        Me.dgvSampleReceiptWatson.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle191.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle191.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle191.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle191.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle191.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle191.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle191.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSampleReceiptWatson.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle191
        Me.dgvSampleReceiptWatson.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSampleReceiptWatson.Location = New System.Drawing.Point(3, 381)
        Me.dgvSampleReceiptWatson.Name = "dgvSampleReceiptWatson"
        Me.dgvSampleReceiptWatson.ReadOnly = True
        Me.dgvSampleReceiptWatson.RowHeadersWidth = 25
        Me.dgvSampleReceiptWatson.Size = New System.Drawing.Size(901, 219)
        Me.dgvSampleReceiptWatson.TabIndex = 102
        '
        'Label41
        '
        Me.Label41.BackColor = System.Drawing.Color.Transparent
        Me.Label41.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label41.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Label41.Location = New System.Drawing.Point(332, 57)
        Me.Label41.Name = "Label41"
        Me.Label41.Size = New System.Drawing.Size(253, 19)
        Me.Label41.TabIndex = 101
        Me.Label41.Text = ":Total number to be used in Report"
        '
        'txtSRecTotalReport
        '
        Me.txtSRecTotalReport.BackColor = System.Drawing.Color.White
        Me.txtSRecTotalReport.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSRecTotalReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSRecTotalReport.Location = New System.Drawing.Point(271, 54)
        Me.txtSRecTotalReport.Multiline = True
        Me.txtSRecTotalReport.Name = "txtSRecTotalReport"
        Me.txtSRecTotalReport.ReadOnly = True
        Me.txtSRecTotalReport.Size = New System.Drawing.Size(59, 24)
        Me.txtSRecTotalReport.TabIndex = 100
        Me.txtSRecTotalReport.Text = "0"
        Me.txtSRecTotalReport.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label40
        '
        Me.Label40.BackColor = System.Drawing.Color.Transparent
        Me.Label40.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label40.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Label40.Location = New System.Drawing.Point(332, 30)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(253, 19)
        Me.Label40.TabIndex = 99
        Me.Label40.Text = ":Total number of samples received"
        '
        'txtSRecTotal
        '
        Me.txtSRecTotal.BackColor = System.Drawing.Color.White
        Me.txtSRecTotal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSRecTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSRecTotal.Location = New System.Drawing.Point(271, 27)
        Me.txtSRecTotal.Name = "txtSRecTotal"
        Me.txtSRecTotal.ReadOnly = True
        Me.txtSRecTotal.Size = New System.Drawing.Size(59, 24)
        Me.txtSRecTotal.TabIndex = 98
        Me.txtSRecTotal.Text = "0"
        Me.txtSRecTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'dgvSampleReceipt
        '
        Me.dgvSampleReceipt.AllowUserToAddRows = False
        Me.dgvSampleReceipt.AllowUserToDeleteRows = False
        Me.dgvSampleReceipt.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSampleReceipt.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvSampleReceipt.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvSampleReceipt.BackgroundColor = System.Drawing.Color.White
        Me.dgvSampleReceipt.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        DataGridViewCellStyle192.Alignment = System.Windows.Forms.DataGridViewContentAlignment.BottomCenter
        DataGridViewCellStyle192.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle192.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle192.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle192.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle192.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle192.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvSampleReceipt.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle192
        Me.dgvSampleReceipt.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSampleReceipt.Location = New System.Drawing.Point(3, 117)
        Me.dgvSampleReceipt.Name = "dgvSampleReceipt"
        Me.dgvSampleReceipt.ReadOnly = True
        Me.dgvSampleReceipt.RowHeadersWidth = 25
        Me.dgvSampleReceipt.Size = New System.Drawing.Size(901, 219)
        Me.dgvSampleReceipt.TabIndex = 95
        '
        'Label37
        '
        Me.Label37.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label37.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label37.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label37.ForeColor = System.Drawing.Color.White
        Me.Label37.Location = New System.Drawing.Point(0, 0)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(908, 21)
        Me.Label37.TabIndex = 27
        Me.Label37.Text = "Sample Receipt Records"
        '
        'tp13
        '
        Me.tp13.AutoScroll = True
        Me.tp13.BackColor = System.Drawing.Color.Ivory
        Me.tp13.Controls.Add(Me.pbxWord)
        Me.tp13.Controls.Add(Me.cmdAppFig)
        Me.tp13.Controls.Add(Me.Label38)
        Me.tp13.Location = New System.Drawing.Point(4, 24)
        Me.tp13.Name = "tp13"
        Me.tp13.Size = New System.Drawing.Size(908, 603)
        Me.tp13.TabIndex = 13
        Me.tp13.Text = "13"
        '
        'pbxWord
        '
        Me.pbxWord.Location = New System.Drawing.Point(104, 176)
        Me.pbxWord.Name = "pbxWord"
        Me.pbxWord.Size = New System.Drawing.Size(638, 335)
        Me.pbxWord.TabIndex = 103
        Me.pbxWord.TabStop = False
        '
        'cmdAppFig
        '
        Me.cmdAppFig.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cmdAppFig.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAppFig.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAppFig.Location = New System.Drawing.Point(9, 47)
        Me.cmdAppFig.Name = "cmdAppFig"
        Me.cmdAppFig.Size = New System.Drawing.Size(173, 72)
        Me.cmdAppFig.TabIndex = 102
        Me.cmdAppFig.Text = "&Open Appendices and Figures..."
        Me.cmdAppFig.UseVisualStyleBackColor = False
        '
        'Label38
        '
        Me.Label38.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label38.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label38.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label38.ForeColor = System.Drawing.Color.White
        Me.Label38.Location = New System.Drawing.Point(0, 0)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(908, 21)
        Me.Label38.TabIndex = 28
        Me.Label38.Text = "Edit Appendices && Figures"
        '
        'tp14
        '
        Me.tp14.AutoScroll = True
        Me.tp14.BackColor = System.Drawing.Color.Ivory
        Me.tp14.Controls.Add(Me.cmdAdministration)
        Me.tp14.Controls.Add(Me.Label3)
        Me.tp14.Location = New System.Drawing.Point(4, 24)
        Me.tp14.Name = "tp14"
        Me.tp14.Size = New System.Drawing.Size(908, 603)
        Me.tp14.TabIndex = 14
        Me.tp14.Text = "14"
        '
        'cmdAdministration
        '
        Me.cmdAdministration.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cmdAdministration.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdministration.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAdministration.Location = New System.Drawing.Point(9, 47)
        Me.cmdAdministration.Name = "cmdAdministration"
        Me.cmdAdministration.Size = New System.Drawing.Size(173, 72)
        Me.cmdAdministration.TabIndex = 30
        Me.cmdAdministration.Text = "Open &Administration Window..."
        Me.cmdAdministration.UseVisualStyleBackColor = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(0, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(908, 21)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "Report Writer Administration"
        '
        'tp15
        '
        Me.tp15.AutoScroll = True
        Me.tp15.BackColor = System.Drawing.Color.Ivory
        Me.tp15.Controls.Add(Me.cmdAnalDetails)
        Me.tp15.Controls.Add(Me.Label8)
        Me.tp15.Location = New System.Drawing.Point(4, 24)
        Me.tp15.Name = "tp15"
        Me.tp15.Size = New System.Drawing.Size(908, 603)
        Me.tp15.TabIndex = 15
        Me.tp15.Text = "15"
        '
        'cmdAnalDetails
        '
        Me.cmdAnalDetails.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cmdAnalDetails.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAnalDetails.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAnalDetails.Location = New System.Drawing.Point(9, 47)
        Me.cmdAnalDetails.Name = "cmdAnalDetails"
        Me.cmdAnalDetails.Size = New System.Drawing.Size(173, 72)
        Me.cmdAnalDetails.TabIndex = 136
        Me.cmdAnalDetails.Text = "&View Details..."
        Me.cmdAnalDetails.UseVisualStyleBackColor = False
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label8.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label8.ForeColor = System.Drawing.Color.White
        Me.Label8.Location = New System.Drawing.Point(0, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(908, 21)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Calibration and QC Standard Details"
        '
        'tp16
        '
        Me.tp16.AutoScroll = True
        Me.tp16.BackColor = System.Drawing.Color.Ivory
        Me.tp16.Controls.Add(Me.GroupBox3)
        Me.tp16.Controls.Add(Me.cmdAuditTrail)
        Me.tp16.Controls.Add(Me.lblAuditTrail)
        Me.tp16.Location = New System.Drawing.Point(4, 24)
        Me.tp16.Name = "tp16"
        Me.tp16.Size = New System.Drawing.Size(908, 603)
        Me.tp16.TabIndex = 16
        Me.tp16.Text = "16"
        '
        'GroupBox3
        '
        Me.GroupBox3.AutoSize = True
        Me.GroupBox3.Controls.Add(Me.lblAT)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 3.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(9, 133)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(419, 52)
        Me.GroupBox3.TabIndex = 140
        Me.GroupBox3.TabStop = False
        '
        'lblAT
        '
        Me.lblAT.AutoSize = True
        Me.lblAT.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAT.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblAT.Location = New System.Drawing.Point(6, 15)
        Me.lblAT.Name = "lblAT"
        Me.lblAT.Size = New System.Drawing.Size(405, 21)
        Me.lblAT.TabIndex = 139
        Me.lblAT.Text = "Study must be in Saved mode in order to view Audit Trail"
        '
        'cmdAuditTrail
        '
        Me.cmdAuditTrail.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.cmdAuditTrail.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAuditTrail.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAuditTrail.Location = New System.Drawing.Point(9, 47)
        Me.cmdAuditTrail.Name = "cmdAuditTrail"
        Me.cmdAuditTrail.Size = New System.Drawing.Size(173, 72)
        Me.cmdAuditTrail.TabIndex = 138
        Me.cmdAuditTrail.Text = "&View Audit Trail"
        Me.cmdAuditTrail.UseVisualStyleBackColor = False
        '
        'lblAuditTrail
        '
        Me.lblAuditTrail.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblAuditTrail.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblAuditTrail.Font = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblAuditTrail.ForeColor = System.Drawing.Color.White
        Me.lblAuditTrail.Location = New System.Drawing.Point(0, 0)
        Me.lblAuditTrail.Name = "lblAuditTrail"
        Me.lblAuditTrail.Size = New System.Drawing.Size(908, 21)
        Me.lblAuditTrail.TabIndex = 137
        Me.lblAuditTrail.Text = "Report Writer Audit Trail"
        '
        'gbFilters
        '
        Me.gbFilters.Controls.Add(Me.txtFilterSamples)
        Me.gbFilters.Controls.Add(Me.cmdClearFilters)
        Me.gbFilters.Location = New System.Drawing.Point(2, 91)
        Me.gbFilters.Name = "gbFilters"
        Me.gbFilters.Size = New System.Drawing.Size(122, 78)
        Me.gbFilters.TabIndex = 172
        Me.gbFilters.TabStop = False
        Me.gbFilters.Text = "Filter Tables"
        '
        'txtFilterSamples
        '
        Me.txtFilterSamples.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilterSamples.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFilterSamples.Location = New System.Drawing.Point(6, 19)
        Me.txtFilterSamples.Name = "txtFilterSamples"
        Me.txtFilterSamples.Size = New System.Drawing.Size(107, 25)
        Me.txtFilterSamples.TabIndex = 156
        '
        'cmdClearFilters
        '
        Me.cmdClearFilters.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClearFilters.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdClearFilters.Enabled = False
        Me.cmdClearFilters.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearFilters.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdClearFilters.Location = New System.Drawing.Point(6, 46)
        Me.cmdClearFilters.Name = "cmdClearFilters"
        Me.cmdClearFilters.Size = New System.Drawing.Size(107, 27)
        Me.cmdClearFilters.TabIndex = 158
        Me.cmdClearFilters.Text = "&Clear Filter"
        Me.cmdClearFilters.UseVisualStyleBackColor = False
        '
        'cmdCreateReportTitle2
        '
        Me.cmdCreateReportTitle2.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCreateReportTitle2.Enabled = False
        Me.cmdCreateReportTitle2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCreateReportTitle2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdCreateReportTitle2.Location = New System.Drawing.Point(0, 164)
        Me.cmdCreateReportTitle2.Name = "cmdCreateReportTitle2"
        Me.cmdCreateReportTitle2.Size = New System.Drawing.Size(111, 44)
        Me.cmdCreateReportTitle2.TabIndex = 96
        Me.cmdCreateReportTitle2.Text = "&Create Report Title"
        Me.cmdCreateReportTitle2.UseVisualStyleBackColor = False
        '
        'gbSource
        '
        Me.gbSource.Controls.Add(Me.rbOracle)
        Me.gbSource.Controls.Add(Me.txtArchivePath)
        Me.gbSource.Controls.Add(Me.rbArchive)
        Me.gbSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSource.Location = New System.Drawing.Point(59, 9)
        Me.gbSource.Name = "gbSource"
        Me.gbSource.Size = New System.Drawing.Size(46, 36)
        Me.gbSource.TabIndex = 117
        Me.gbSource.TabStop = False
        Me.gbSource.Text = "Watson Data Source"
        Me.gbSource.Visible = False
        '
        'rbOracle
        '
        Me.rbOracle.AutoSize = True
        Me.rbOracle.Checked = True
        Me.rbOracle.Location = New System.Drawing.Point(43, 38)
        Me.rbOracle.Name = "rbOracle"
        Me.rbOracle.Size = New System.Drawing.Size(61, 19)
        Me.rbOracle.TabIndex = 0
        Me.rbOracle.TabStop = True
        Me.rbOracle.Text = "Oracle"
        Me.rbOracle.UseVisualStyleBackColor = True
        Me.rbOracle.Visible = False
        '
        'txtArchivePath
        '
        Me.txtArchivePath.Location = New System.Drawing.Point(46, 17)
        Me.txtArchivePath.Name = "txtArchivePath"
        Me.txtArchivePath.Size = New System.Drawing.Size(47, 21)
        Me.txtArchivePath.TabIndex = 3
        Me.txtArchivePath.Visible = False
        '
        'rbArchive
        '
        Me.rbArchive.AutoSize = True
        Me.rbArchive.Location = New System.Drawing.Point(5, 38)
        Me.rbArchive.Name = "rbArchive"
        Me.rbArchive.Size = New System.Drawing.Size(102, 19)
        Me.rbArchive.TabIndex = 1
        Me.rbArchive.Text = "Archived MDB"
        Me.rbArchive.UseVisualStyleBackColor = True
        Me.rbArchive.Visible = False
        '
        'panFilterStudy
        '
        Me.panFilterStudy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panFilterStudy.Controls.Add(Me.cmdFilterStudy)
        Me.panFilterStudy.Controls.Add(Me.txtFilterStudy)
        Me.panFilterStudy.Controls.Add(Me.cbxFilterStudy)
        Me.panFilterStudy.Controls.Add(Me.Label13)
        Me.panFilterStudy.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.panFilterStudy.Location = New System.Drawing.Point(0, 434)
        Me.panFilterStudy.Name = "panFilterStudy"
        Me.panFilterStudy.Size = New System.Drawing.Size(112, 118)
        Me.panFilterStudy.TabIndex = 116
        '
        'cmdFilterStudy
        '
        Me.cmdFilterStudy.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFilterStudy.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdFilterStudy.Location = New System.Drawing.Point(3, 76)
        Me.cmdFilterStudy.Name = "cmdFilterStudy"
        Me.cmdFilterStudy.Size = New System.Drawing.Size(105, 25)
        Me.cmdFilterStudy.TabIndex = 119
        Me.cmdFilterStudy.Text = "Execute &Filter"
        Me.cmdFilterStudy.UseVisualStyleBackColor = True
        '
        'txtFilterStudy
        '
        Me.txtFilterStudy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFilterStudy.Location = New System.Drawing.Point(3, 48)
        Me.txtFilterStudy.Name = "txtFilterStudy"
        Me.txtFilterStudy.Size = New System.Drawing.Size(105, 21)
        Me.txtFilterStudy.TabIndex = 118
        '
        'cbxFilterStudy
        '
        Me.cbxFilterStudy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxFilterStudy.FormattingEnabled = True
        Me.cbxFilterStudy.Location = New System.Drawing.Point(3, 18)
        Me.cbxFilterStudy.Name = "cbxFilterStudy"
        Me.cbxFilterStudy.Size = New System.Drawing.Size(105, 23)
        Me.cbxFilterStudy.TabIndex = 117
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label13.Location = New System.Drawing.Point(3, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(102, 15)
        Me.Label13.TabIndex = 116
        Me.Label13.Text = "Filter Studies for: "
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'Label10
        '
        Me.Label10.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label10.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.Label10.Location = New System.Drawing.Point(0, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(110, 45)
        Me.Label10.TabIndex = 106
        Me.Label10.Text = "Filter Configured Studies:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'gbxMultVal
        '
        Me.gbxMultVal.Controls.Add(Me.rbMultValNo)
        Me.gbxMultVal.Controls.Add(Me.rbMultValYes)
        Me.gbxMultVal.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbxMultVal.Location = New System.Drawing.Point(4, 86)
        Me.gbxMultVal.Name = "gbxMultVal"
        Me.gbxMultVal.Size = New System.Drawing.Size(118, 131)
        Me.gbxMultVal.TabIndex = 137
        Me.gbxMultVal.TabStop = False
        Me.gbxMultVal.Text = "Report Summary Table Type"
        '
        'rbMultValNo
        '
        Me.rbMultValNo.AutoSize = True
        Me.rbMultValNo.BackColor = System.Drawing.Color.Transparent
        Me.rbMultValNo.Location = New System.Drawing.Point(3, 98)
        Me.rbMultValNo.Name = "rbMultValNo"
        Me.rbMultValNo.Size = New System.Drawing.Size(97, 21)
        Me.rbMultValNo.TabIndex = 1
        Me.rbMultValNo.Text = "Single Table"
        Me.rbMultValNo.UseVisualStyleBackColor = False
        '
        'rbMultValYes
        '
        Me.rbMultValYes.BackColor = System.Drawing.Color.Transparent
        Me.rbMultValYes.Checked = True
        Me.rbMultValYes.Location = New System.Drawing.Point(3, 53)
        Me.rbMultValYes.Name = "rbMultValYes"
        Me.rbMultValYes.Size = New System.Drawing.Size(103, 38)
        Me.rbMultValYes.TabIndex = 0
        Me.rbMultValYes.TabStop = True
        Me.rbMultValYes.Text = "Table for each Analyte"
        Me.rbMultValYes.UseVisualStyleBackColor = False
        '
        'grbShowSummaryTable
        '
        Me.grbShowSummaryTable.Controls.Add(Me.rbShowIncludedSummaryTable)
        Me.grbShowSummaryTable.Controls.Add(Me.rbShowAllSummaryTable)
        Me.grbShowSummaryTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grbShowSummaryTable.Location = New System.Drawing.Point(4, 218)
        Me.grbShowSummaryTable.Name = "grbShowSummaryTable"
        Me.grbShowSummaryTable.Size = New System.Drawing.Size(119, 67)
        Me.grbShowSummaryTable.TabIndex = 134
        Me.grbShowSummaryTable.TabStop = False
        Me.grbShowSummaryTable.Text = "Show Tables"
        '
        'rbShowIncludedSummaryTable
        '
        Me.rbShowIncludedSummaryTable.AutoSize = True
        Me.rbShowIncludedSummaryTable.BackColor = System.Drawing.Color.Transparent
        Me.rbShowIncludedSummaryTable.Location = New System.Drawing.Point(3, 40)
        Me.rbShowIncludedSummaryTable.Name = "rbShowIncludedSummaryTable"
        Me.rbShowIncludedSummaryTable.Size = New System.Drawing.Size(106, 19)
        Me.rbShowIncludedSummaryTable.TabIndex = 1
        Me.rbShowIncludedSummaryTable.Text = "Show Included"
        Me.rbShowIncludedSummaryTable.UseVisualStyleBackColor = False
        '
        'rbShowAllSummaryTable
        '
        Me.rbShowAllSummaryTable.AutoSize = True
        Me.rbShowAllSummaryTable.BackColor = System.Drawing.Color.Transparent
        Me.rbShowAllSummaryTable.Checked = True
        Me.rbShowAllSummaryTable.Location = New System.Drawing.Point(3, 18)
        Me.rbShowAllSummaryTable.Name = "rbShowAllSummaryTable"
        Me.rbShowAllSummaryTable.Size = New System.Drawing.Size(72, 19)
        Me.rbShowAllSummaryTable.TabIndex = 0
        Me.rbShowAllSummaryTable.TabStop = True
        Me.rbShowAllSummaryTable.Text = "Show All"
        Me.rbShowAllSummaryTable.UseVisualStyleBackColor = False
        '
        'gbRTC
        '
        Me.gbRTC.Controls.Add(Me.rbShowAllRTConfig)
        Me.gbRTC.Controls.Add(Me.rbShowIncludedRTConfig)
        Me.gbRTC.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.gbRTC.Location = New System.Drawing.Point(2, 34)
        Me.gbRTC.Name = "gbRTC"
        Me.gbRTC.Size = New System.Drawing.Size(122, 58)
        Me.gbRTC.TabIndex = 131
        Me.gbRTC.TabStop = False
        Me.gbRTC.Text = "Show Tables"
        '
        'rbShowAllRTConfig
        '
        Me.rbShowAllRTConfig.AutoSize = True
        Me.rbShowAllRTConfig.BackColor = System.Drawing.Color.Transparent
        Me.rbShowAllRTConfig.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.rbShowAllRTConfig.Location = New System.Drawing.Point(3, 35)
        Me.rbShowAllRTConfig.Name = "rbShowAllRTConfig"
        Me.rbShowAllRTConfig.Size = New System.Drawing.Size(71, 19)
        Me.rbShowAllRTConfig.TabIndex = 1
        Me.rbShowAllRTConfig.Text = "Show All"
        Me.rbShowAllRTConfig.UseVisualStyleBackColor = False
        '
        'rbShowIncludedRTConfig
        '
        Me.rbShowIncludedRTConfig.AutoSize = True
        Me.rbShowIncludedRTConfig.BackColor = System.Drawing.Color.Transparent
        Me.rbShowIncludedRTConfig.Checked = True
        Me.rbShowIncludedRTConfig.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.rbShowIncludedRTConfig.Location = New System.Drawing.Point(3, 16)
        Me.rbShowIncludedRTConfig.Name = "rbShowIncludedRTConfig"
        Me.rbShowIncludedRTConfig.Size = New System.Drawing.Size(103, 19)
        Me.rbShowIncludedRTConfig.TabIndex = 0
        Me.rbShowIncludedRTConfig.TabStop = True
        Me.rbShowIncludedRTConfig.Text = "Show Included"
        Me.rbShowIncludedRTConfig.UseVisualStyleBackColor = False
        '
        'cmdHomeCancel
        '
        Me.cmdHomeCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdHomeCancel.CausesValidation = False
        Me.cmdHomeCancel.Enabled = False
        Me.cmdHomeCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdHomeCancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdHomeCancel.Location = New System.Drawing.Point(0, 3)
        Me.cmdHomeCancel.Name = "cmdHomeCancel"
        Me.cmdHomeCancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdHomeCancel.TabIndex = 86
        Me.cmdHomeCancel.Text = "&Reset"
        Me.cmdHomeCancel.UseVisualStyleBackColor = False
        '
        'cmdDataCancel
        '
        Me.cmdDataCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDataCancel.CausesValidation = False
        Me.cmdDataCancel.Enabled = False
        Me.cmdDataCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDataCancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdDataCancel.Location = New System.Drawing.Point(3, 3)
        Me.cmdDataCancel.Name = "cmdDataCancel"
        Me.cmdDataCancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdDataCancel.TabIndex = 90
        Me.cmdDataCancel.Text = "Reset"
        Me.cmdDataCancel.UseVisualStyleBackColor = False
        '
        'cmdViewAnalyticalRuns1
        '
        Me.cmdViewAnalyticalRuns1.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdViewAnalyticalRuns1.CausesValidation = False
        Me.cmdViewAnalyticalRuns1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdViewAnalyticalRuns1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdViewAnalyticalRuns1.Location = New System.Drawing.Point(3, 34)
        Me.cmdViewAnalyticalRuns1.Name = "cmdViewAnalyticalRuns1"
        Me.cmdViewAnalyticalRuns1.Size = New System.Drawing.Size(105, 64)
        Me.cmdViewAnalyticalRuns1.TabIndex = 136
        Me.cmdViewAnalyticalRuns1.Text = "&View Analytical Runs"
        Me.cmdViewAnalyticalRuns1.UseVisualStyleBackColor = False
        '
        'cmdAnaRunSumCancel
        '
        Me.cmdAnaRunSumCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAnaRunSumCancel.CausesValidation = False
        Me.cmdAnaRunSumCancel.Enabled = False
        Me.cmdAnaRunSumCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdAnaRunSumCancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdAnaRunSumCancel.Location = New System.Drawing.Point(3, 3)
        Me.cmdAnaRunSumCancel.Name = "cmdAnaRunSumCancel"
        Me.cmdAnaRunSumCancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdAnaRunSumCancel.TabIndex = 91
        Me.cmdAnaRunSumCancel.Text = "Reset"
        Me.cmdAnaRunSumCancel.UseVisualStyleBackColor = False
        '
        'cmdUpdateSummaryInfo
        '
        Me.cmdUpdateSummaryInfo.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdUpdateSummaryInfo.CausesValidation = False
        Me.cmdUpdateSummaryInfo.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdUpdateSummaryInfo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateSummaryInfo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdUpdateSummaryInfo.Location = New System.Drawing.Point(4, 33)
        Me.cmdUpdateSummaryInfo.Name = "cmdUpdateSummaryInfo"
        Me.cmdUpdateSummaryInfo.Size = New System.Drawing.Size(74, 44)
        Me.cmdUpdateSummaryInfo.TabIndex = 136
        Me.cmdUpdateSummaryInfo.Text = "&Refresh Info"
        Me.cmdUpdateSummaryInfo.UseVisualStyleBackColor = False
        '
        'cmdResetSummaryTable
        '
        Me.cmdResetSummaryTable.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdResetSummaryTable.CausesValidation = False
        Me.cmdResetSummaryTable.Enabled = False
        Me.cmdResetSummaryTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetSummaryTable.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdResetSummaryTable.Location = New System.Drawing.Point(3, 3)
        Me.cmdResetSummaryTable.Name = "cmdResetSummaryTable"
        Me.cmdResetSummaryTable.Size = New System.Drawing.Size(73, 25)
        Me.cmdResetSummaryTable.TabIndex = 101
        Me.cmdResetSummaryTable.Text = "Reset"
        Me.cmdResetSummaryTable.UseVisualStyleBackColor = False
        '
        'cmdRefreshStatements
        '
        Me.cmdRefreshStatements.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRefreshStatements.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdRefreshStatements.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdRefreshStatements.Location = New System.Drawing.Point(3, 34)
        Me.cmdRefreshStatements.Name = "cmdRefreshStatements"
        Me.cmdRefreshStatements.Size = New System.Drawing.Size(97, 66)
        Me.cmdRefreshStatements.TabIndex = 139
        Me.cmdRefreshStatements.Text = "Sho&w Templates"
        Me.cmdRefreshStatements.UseVisualStyleBackColor = False
        '
        'cmdOpenReportStatements
        '
        Me.cmdOpenReportStatements.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOpenReportStatements.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdOpenReportStatements.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdOpenReportStatements.Location = New System.Drawing.Point(3, 106)
        Me.cmdOpenReportStatements.Name = "cmdOpenReportStatements"
        Me.cmdOpenReportStatements.Size = New System.Drawing.Size(97, 66)
        Me.cmdOpenReportStatements.TabIndex = 126
        Me.cmdOpenReportStatements.Text = "&Edit Templates"
        Me.cmdOpenReportStatements.UseVisualStyleBackColor = False
        '
        'cmdCancelReportStatements
        '
        Me.cmdCancelReportStatements.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCancelReportStatements.CausesValidation = False
        Me.cmdCancelReportStatements.Enabled = False
        Me.cmdCancelReportStatements.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdCancelReportStatements.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdCancelReportStatements.Location = New System.Drawing.Point(3, 3)
        Me.cmdCancelReportStatements.Name = "cmdCancelReportStatements"
        Me.cmdCancelReportStatements.Size = New System.Drawing.Size(73, 25)
        Me.cmdCancelReportStatements.TabIndex = 123
        Me.cmdCancelReportStatements.Text = "Reset"
        Me.cmdCancelReportStatements.UseVisualStyleBackColor = False
        '
        'cmdCreateTable
        '
        Me.cmdCreateTable.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCreateTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdCreateTable.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdCreateTable.Location = New System.Drawing.Point(0, 170)
        Me.cmdCreateTable.Name = "cmdCreateTable"
        Me.cmdCreateTable.Size = New System.Drawing.Size(115, 45)
        Me.cmdCreateTable.TabIndex = 151
        Me.cmdCreateTable.Text = "Create Re&port Table"
        Me.cmdCreateTable.UseVisualStyleBackColor = False
        '
        'cmdImportTables
        '
        Me.cmdImportTables.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdImportTables.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdImportTables.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdImportTables.Location = New System.Drawing.Point(0, 550)
        Me.cmdImportTables.Name = "cmdImportTables"
        Me.cmdImportTables.Size = New System.Drawing.Size(115, 45)
        Me.cmdImportTables.TabIndex = 150
        Me.cmdImportTables.Text = "&Import Tables"
        Me.cmdImportTables.UseVisualStyleBackColor = False
        '
        'cmdSelect
        '
        Me.cmdSelect.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSelect.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdSelect.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdSelect.Location = New System.Drawing.Point(0, 503)
        Me.cmdSelect.Name = "cmdSelect"
        Me.cmdSelect.Size = New System.Drawing.Size(115, 45)
        Me.cmdSelect.TabIndex = 148
        Me.cmdSelect.Text = "Select/ &Deselect All"
        Me.cmdSelect.UseVisualStyleBackColor = False
        '
        'cmdOutliers
        '
        Me.cmdOutliers.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdOutliers.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdOutliers.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdOutliers.Location = New System.Drawing.Point(0, 392)
        Me.cmdOutliers.Name = "cmdOutliers"
        Me.cmdOutliers.Size = New System.Drawing.Size(115, 45)
        Me.cmdOutliers.TabIndex = 141
        Me.cmdOutliers.Text = "Eval &Outliers"
        Me.cmdOutliers.UseVisualStyleBackColor = False
        '
        'cmdAdvancedTable
        '
        Me.cmdAdvancedTable.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAdvancedTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdAdvancedTable.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdAdvancedTable.Location = New System.Drawing.Point(0, 264)
        Me.cmdAdvancedTable.Name = "cmdAdvancedTable"
        Me.cmdAdvancedTable.Size = New System.Drawing.Size(115, 62)
        Me.cmdAdvancedTable.TabIndex = 138
        Me.cmdAdvancedTable.Text = "Advanced &Table Configuration"
        Me.cmdAdvancedTable.UseVisualStyleBackColor = False
        '
        'cmdDuplicateTables
        '
        Me.cmdDuplicateTables.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDuplicateTables.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdDuplicateTables.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdDuplicateTables.Location = New System.Drawing.Point(0, 328)
        Me.cmdDuplicateTables.Name = "cmdDuplicateTables"
        Me.cmdDuplicateTables.Size = New System.Drawing.Size(115, 62)
        Me.cmdDuplicateTables.TabIndex = 137
        Me.cmdDuplicateTables.Text = "Create a &Replicate Table"
        Me.cmdDuplicateTables.UseVisualStyleBackColor = False
        '
        'cmdViewAnalRuns
        '
        Me.cmdViewAnalRuns.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdViewAnalRuns.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdViewAnalRuns.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdViewAnalRuns.Location = New System.Drawing.Point(0, 439)
        Me.cmdViewAnalRuns.Name = "cmdViewAnalRuns"
        Me.cmdViewAnalRuns.Size = New System.Drawing.Size(115, 62)
        Me.cmdViewAnalRuns.TabIndex = 139
        Me.cmdViewAnalRuns.Text = "&View Analytical Runs"
        Me.cmdViewAnalRuns.UseVisualStyleBackColor = False
        '
        'cmdAssignSamples
        '
        Me.cmdAssignSamples.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAssignSamples.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdAssignSamples.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAssignSamples.Location = New System.Drawing.Point(0, 217)
        Me.cmdAssignSamples.Name = "cmdAssignSamples"
        Me.cmdAssignSamples.Size = New System.Drawing.Size(115, 45)
        Me.cmdAssignSamples.TabIndex = 140
        Me.cmdAssignSamples.Text = "&Assign Samples"
        Me.cmdAssignSamples.UseVisualStyleBackColor = False
        '
        'cmdRTConfigCancel
        '
        Me.cmdRTConfigCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRTConfigCancel.CausesValidation = False
        Me.cmdRTConfigCancel.Enabled = False
        Me.cmdRTConfigCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdRTConfigCancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdRTConfigCancel.Location = New System.Drawing.Point(0, 3)
        Me.cmdRTConfigCancel.Name = "cmdRTConfigCancel"
        Me.cmdRTConfigCancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdRTConfigCancel.TabIndex = 92
        Me.cmdRTConfigCancel.Text = "Reset"
        Me.cmdRTConfigCancel.UseVisualStyleBackColor = False
        '
        'cmdRTHeaderConfigCancel
        '
        Me.cmdRTHeaderConfigCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdRTHeaderConfigCancel.CausesValidation = False
        Me.cmdRTHeaderConfigCancel.Enabled = False
        Me.cmdRTHeaderConfigCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRTHeaderConfigCancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdRTHeaderConfigCancel.Location = New System.Drawing.Point(3, 3)
        Me.cmdRTHeaderConfigCancel.Name = "cmdRTHeaderConfigCancel"
        Me.cmdRTHeaderConfigCancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdRTHeaderConfigCancel.TabIndex = 93
        Me.cmdRTHeaderConfigCancel.Text = "&Reset"
        Me.cmdRTHeaderConfigCancel.UseVisualStyleBackColor = False
        '
        'cmdAddAnalyte
        '
        Me.cmdAddAnalyte.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddAnalyte.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdAddAnalyte.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdAddAnalyte.Location = New System.Drawing.Point(3, 31)
        Me.cmdAddAnalyte.Name = "cmdAddAnalyte"
        Me.cmdAddAnalyte.Size = New System.Drawing.Size(102, 52)
        Me.cmdAddAnalyte.TabIndex = 105
        Me.cmdAddAnalyte.Text = "&Add Standard..."
        Me.cmdAddAnalyte.UseVisualStyleBackColor = False
        '
        'cmdCopyRepAnalyte
        '
        Me.cmdCopyRepAnalyte.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCopyRepAnalyte.Enabled = False
        Me.cmdCopyRepAnalyte.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdCopyRepAnalyte.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdCopyRepAnalyte.Location = New System.Drawing.Point(3, 147)
        Me.cmdCopyRepAnalyte.Name = "cmdCopyRepAnalyte"
        Me.cmdCopyRepAnalyte.Size = New System.Drawing.Size(102, 52)
        Me.cmdCopyRepAnalyte.TabIndex = 102
        Me.cmdCopyRepAnalyte.Text = "&Copy Data..."
        Me.cmdCopyRepAnalyte.UseVisualStyleBackColor = False
        '
        'cmdDeleteRepAnalyte
        '
        Me.cmdDeleteRepAnalyte.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDeleteRepAnalyte.Enabled = False
        Me.cmdDeleteRepAnalyte.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdDeleteRepAnalyte.ForeColor = System.Drawing.Color.Red
        Me.cmdDeleteRepAnalyte.Location = New System.Drawing.Point(3, 205)
        Me.cmdDeleteRepAnalyte.Name = "cmdDeleteRepAnalyte"
        Me.cmdDeleteRepAnalyte.Size = New System.Drawing.Size(102, 52)
        Me.cmdDeleteRepAnalyte.TabIndex = 93
        Me.cmdDeleteRepAnalyte.Text = "&Delete Standard..."
        Me.cmdDeleteRepAnalyte.UseVisualStyleBackColor = False
        '
        'cmdAnalRefCancel
        '
        Me.cmdAnalRefCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAnalRefCancel.CausesValidation = False
        Me.cmdAnalRefCancel.Enabled = False
        Me.cmdAnalRefCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAnalRefCancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdAnalRefCancel.Location = New System.Drawing.Point(3, 3)
        Me.cmdAnalRefCancel.Name = "cmdAnalRefCancel"
        Me.cmdAnalRefCancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdAnalRefCancel.TabIndex = 92
        Me.cmdAnalRefCancel.Text = "Reset"
        Me.cmdAnalRefCancel.UseVisualStyleBackColor = False
        '
        'cmdAddRepAnalyte
        '
        Me.cmdAddRepAnalyte.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdAddRepAnalyte.Enabled = False
        Me.cmdAddRepAnalyte.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdAddRepAnalyte.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdAddRepAnalyte.Location = New System.Drawing.Point(3, 89)
        Me.cmdAddRepAnalyte.Name = "cmdAddRepAnalyte"
        Me.cmdAddRepAnalyte.Size = New System.Drawing.Size(102, 52)
        Me.cmdAddRepAnalyte.TabIndex = 30
        Me.cmdAddRepAnalyte.Text = "Add &Replicate..."
        Me.cmdAddRepAnalyte.UseVisualStyleBackColor = False
        '
        'cmdReplacePersonnel
        '
        Me.cmdReplacePersonnel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdReplacePersonnel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdReplacePersonnel.Enabled = False
        Me.cmdReplacePersonnel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdReplacePersonnel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdReplacePersonnel.Location = New System.Drawing.Point(4, 96)
        Me.cmdReplacePersonnel.Name = "cmdReplacePersonnel"
        Me.cmdReplacePersonnel.Size = New System.Drawing.Size(121, 106)
        Me.cmdReplacePersonnel.TabIndex = 107
        Me.cmdReplacePersonnel.Text = "&Import and Replace Personnel from Another Study"
        Me.cmdReplacePersonnel.UseVisualStyleBackColor = False
        '
        'cmdCPDelete
        '
        Me.cmdCPDelete.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCPDelete.Enabled = False
        Me.cmdCPDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCPDelete.ForeColor = System.Drawing.Color.Red
        Me.cmdCPDelete.Location = New System.Drawing.Point(4, 65)
        Me.cmdCPDelete.Name = "cmdCPDelete"
        Me.cmdCPDelete.Size = New System.Drawing.Size(121, 25)
        Me.cmdCPDelete.TabIndex = 41
        Me.cmdCPDelete.Text = "&Delete"
        Me.cmdCPDelete.UseVisualStyleBackColor = False
        '
        'cmdCPAdd
        '
        Me.cmdCPAdd.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCPAdd.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCPAdd.Enabled = False
        Me.cmdCPAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCPAdd.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdCPAdd.Location = New System.Drawing.Point(4, 34)
        Me.cmdCPAdd.Name = "cmdCPAdd"
        Me.cmdCPAdd.Size = New System.Drawing.Size(121, 25)
        Me.cmdCPAdd.TabIndex = 40
        Me.cmdCPAdd.Text = "&Add"
        Me.cmdCPAdd.UseVisualStyleBackColor = False
        '
        'cmdCPCancel
        '
        Me.cmdCPCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdCPCancel.CausesValidation = False
        Me.cmdCPCancel.Enabled = False
        Me.cmdCPCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCPCancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdCPCancel.Location = New System.Drawing.Point(3, 3)
        Me.cmdCPCancel.Name = "cmdCPCancel"
        Me.cmdCPCancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdCPCancel.TabIndex = 39
        Me.cmdCPCancel.Text = "Reset"
        Me.cmdCPCancel.UseVisualStyleBackColor = False
        '
        'cmdMethValUpdate
        '
        Me.cmdMethValUpdate.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdMethValUpdate.CausesValidation = False
        Me.cmdMethValUpdate.Enabled = False
        Me.cmdMethValUpdate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMethValUpdate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdMethValUpdate.Location = New System.Drawing.Point(3, 34)
        Me.cmdMethValUpdate.Name = "cmdMethValUpdate"
        Me.cmdMethValUpdate.Size = New System.Drawing.Size(110, 25)
        Me.cmdMethValUpdate.TabIndex = 44
        Me.cmdMethValUpdate.Text = "&Update Info"
        Me.cmdMethValUpdate.UseVisualStyleBackColor = False
        Me.cmdMethValUpdate.Visible = False
        '
        'cmdMethValReset
        '
        Me.cmdMethValReset.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdMethValReset.CausesValidation = False
        Me.cmdMethValReset.Enabled = False
        Me.cmdMethValReset.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdMethValReset.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdMethValReset.Location = New System.Drawing.Point(3, 3)
        Me.cmdMethValReset.Name = "cmdMethValReset"
        Me.cmdMethValReset.Size = New System.Drawing.Size(73, 25)
        Me.cmdMethValReset.TabIndex = 41
        Me.cmdMethValReset.Text = "&Reset"
        Me.cmdMethValReset.UseVisualStyleBackColor = False
        '
        'cmdQACancel
        '
        Me.cmdQACancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdQACancel.CausesValidation = False
        Me.cmdQACancel.Enabled = False
        Me.cmdQACancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdQACancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdQACancel.Location = New System.Drawing.Point(3, 3)
        Me.cmdQACancel.Name = "cmdQACancel"
        Me.cmdQACancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdQACancel.TabIndex = 42
        Me.cmdQACancel.Text = "Reset"
        Me.cmdQACancel.UseVisualStyleBackColor = False
        '
        'cmdDeleteQAEvent
        '
        Me.cmdDeleteQAEvent.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDeleteQAEvent.Enabled = False
        Me.cmdDeleteQAEvent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdDeleteQAEvent.ForeColor = System.Drawing.Color.Red
        Me.cmdDeleteQAEvent.Location = New System.Drawing.Point(3, 80)
        Me.cmdDeleteQAEvent.Name = "cmdDeleteQAEvent"
        Me.cmdDeleteQAEvent.Size = New System.Drawing.Size(123, 40)
        Me.cmdDeleteQAEvent.TabIndex = 31
        Me.cmdDeleteQAEvent.Text = "&Delete Row"
        Me.cmdDeleteQAEvent.UseVisualStyleBackColor = False
        '
        'cmdInsertQAEvent
        '
        Me.cmdInsertQAEvent.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdInsertQAEvent.Enabled = False
        Me.cmdInsertQAEvent.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdInsertQAEvent.ForeColor = System.Drawing.Color.Blue
        Me.cmdInsertQAEvent.Location = New System.Drawing.Point(3, 34)
        Me.cmdInsertQAEvent.Name = "cmdInsertQAEvent"
        Me.cmdInsertQAEvent.Size = New System.Drawing.Size(123, 40)
        Me.cmdInsertQAEvent.TabIndex = 29
        Me.cmdInsertQAEvent.Text = "&Insert Row Below Selection"
        Me.cmdInsertQAEvent.UseVisualStyleBackColor = False
        '
        'cmdSRecCancel
        '
        Me.cmdSRecCancel.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdSRecCancel.CausesValidation = False
        Me.cmdSRecCancel.Enabled = False
        Me.cmdSRecCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSRecCancel.ForeColor = System.Drawing.Color.Firebrick
        Me.cmdSRecCancel.Location = New System.Drawing.Point(3, 3)
        Me.cmdSRecCancel.Name = "cmdSRecCancel"
        Me.cmdSRecCancel.Size = New System.Drawing.Size(73, 25)
        Me.cmdSRecCancel.TabIndex = 96
        Me.cmdSRecCancel.Text = "Reset"
        Me.cmdSRecCancel.UseVisualStyleBackColor = False
        '
        'cmdDeletSRec
        '
        Me.cmdDeletSRec.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdDeletSRec.Enabled = False
        Me.cmdDeletSRec.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdDeletSRec.ForeColor = System.Drawing.Color.Red
        Me.cmdDeletSRec.Location = New System.Drawing.Point(3, 80)
        Me.cmdDeletSRec.Name = "cmdDeletSRec"
        Me.cmdDeletSRec.Size = New System.Drawing.Size(123, 40)
        Me.cmdDeletSRec.TabIndex = 34
        Me.cmdDeletSRec.Text = "Delete Row"
        Me.cmdDeletSRec.UseVisualStyleBackColor = False
        '
        'cmdInsertSRec
        '
        Me.cmdInsertSRec.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdInsertSRec.Enabled = False
        Me.cmdInsertSRec.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdInsertSRec.ForeColor = System.Drawing.Color.Blue
        Me.cmdInsertSRec.Location = New System.Drawing.Point(3, 34)
        Me.cmdInsertSRec.Name = "cmdInsertSRec"
        Me.cmdInsertSRec.Size = New System.Drawing.Size(123, 40)
        Me.cmdInsertSRec.TabIndex = 33
        Me.cmdInsertSRec.Text = "Insert Row Below Selection"
        Me.cmdInsertSRec.UseVisualStyleBackColor = False
        '
        'cmdReportHistory
        '
        Me.cmdReportHistory.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdReportHistory.CausesValidation = False
        Me.cmdReportHistory.Enabled = False
        Me.cmdReportHistory.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReportHistory.ForeColor = System.Drawing.Color.FromArgb(CType(CType(231, Byte), Integer), CType(CType(87, Byte), Integer), CType(CType(56, Byte), Integer))
        Me.cmdReportHistory.Location = New System.Drawing.Point(0, 281)
        Me.cmdReportHistory.Name = "cmdReportHistory"
        Me.cmdReportHistory.Size = New System.Drawing.Size(111, 44)
        Me.cmdReportHistory.TabIndex = 119
        Me.cmdReportHistory.Text = "&View Report History"
        Me.cmdReportHistory.UseVisualStyleBackColor = False
        '
        'cmdShowOutstanding
        '
        Me.cmdShowOutstanding.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdShowOutstanding.CausesValidation = False
        Me.cmdShowOutstanding.Enabled = False
        Me.cmdShowOutstanding.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShowOutstanding.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdShowOutstanding.Location = New System.Drawing.Point(0, 214)
        Me.cmdShowOutstanding.Name = "cmdShowOutstanding"
        Me.cmdShowOutstanding.Size = New System.Drawing.Size(111, 61)
        Me.cmdShowOutstanding.TabIndex = 91
        Me.cmdShowOutstanding.Text = "Show &Outstanding Report Items"
        Me.cmdShowOutstanding.UseVisualStyleBackColor = False
        '
        'cmdApplyTemplate
        '
        Me.cmdApplyTemplate.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdApplyTemplate.CausesValidation = False
        Me.cmdApplyTemplate.Enabled = False
        Me.cmdApplyTemplate.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdApplyTemplate.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdApplyTemplate.Location = New System.Drawing.Point(0, 331)
        Me.cmdApplyTemplate.Name = "cmdApplyTemplate"
        Me.cmdApplyTemplate.Size = New System.Drawing.Size(111, 44)
        Me.cmdApplyTemplate.TabIndex = 90
        Me.cmdApplyTemplate.Text = "&Apply/View Study Template"
        Me.cmdApplyTemplate.UseVisualStyleBackColor = False
        '
        'cmdUpdateProject
        '
        Me.cmdUpdateProject.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdUpdateProject.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpdateProject.ForeColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(227, Byte), Integer))
        Me.cmdUpdateProject.Location = New System.Drawing.Point(0, 97)
        Me.cmdUpdateProject.Name = "cmdUpdateProject"
        Me.cmdUpdateProject.Size = New System.Drawing.Size(111, 61)
        Me.cmdUpdateProject.TabIndex = 12
        Me.cmdUpdateProject.Text = "Retrieve Watson Study"
        Me.cmdUpdateProject.UseVisualStyleBackColor = False
        '
        'pb1
        '
        Me.pb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.pb1.Cursor = System.Windows.Forms.Cursors.Default
        Me.pb1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pb1.Location = New System.Drawing.Point(0, 257)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(68, 26)
        Me.pb1.TabIndex = 101
        '
        'pb2
        '
        Me.pb2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb2.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.pb2.Cursor = System.Windows.Forms.Cursors.Default
        Me.pb2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.pb2.Location = New System.Drawing.Point(0, 283)
        Me.pb2.Name = "pb2"
        Me.pb2.Size = New System.Drawing.Size(68, 26)
        Me.pb2.TabIndex = 116
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.BackColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblProgress.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblProgress.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!)
        Me.lblProgress.ForeColor = System.Drawing.Color.White
        Me.lblProgress.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblProgress.Location = New System.Drawing.Point(0, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(68, 257)
        Me.lblProgress.TabIndex = 86
        Me.lblProgress.Text = "lblProgress"
        Me.lblProgress.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'llblAssignedSamples
        '
        Me.llblAssignedSamples.ActiveLinkColor = System.Drawing.Color.Black
        Me.llblAssignedSamples.AutoSize = True
        Me.llblAssignedSamples.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
        Me.llblAssignedSamples.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.llblAssignedSamples.LinkColor = System.Drawing.Color.Black
        Me.llblAssignedSamples.Location = New System.Drawing.Point(1052, 76)
        Me.llblAssignedSamples.Name = "llblAssignedSamples"
        Me.llblAssignedSamples.Size = New System.Drawing.Size(183, 16)
        Me.llblAssignedSamples.TabIndex = 92
        Me.llblAssignedSamples.TabStop = True
        Me.llblAssignedSamples.Text = "Samples Need Assigning"
        Me.llblAssignedSamples.Visible = False
        Me.llblAssignedSamples.VisitedLinkColor = System.Drawing.Color.Black
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.Color.Gray
        Me.cmdCancel.CausesValidation = False
        Me.cmdCancel.Enabled = False
        Me.cmdCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCancel.Location = New System.Drawing.Point(147, 0)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(68, 25)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "&Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'lbxTab1
        '
        Me.lbxTab1.BackColor = System.Drawing.Color.Ivory
        Me.lbxTab1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lbxTab1.Font = New System.Drawing.Font("Segoe UI", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbxTab1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lbxTab1.ItemHeight = 20
        Me.lbxTab1.Location = New System.Drawing.Point(8, 119)
        Me.lbxTab1.Name = "lbxTab1"
        Me.lbxTab1.Size = New System.Drawing.Size(226, 360)
        Me.lbxTab1.TabIndex = 0
        '
        'lblTOC
        '
        Me.lblTOC.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTOC.Location = New System.Drawing.Point(8, 96)
        Me.lblTOC.Name = "lblTOC"
        Me.lblTOC.Size = New System.Drawing.Size(226, 21)
        Me.lblTOC.TabIndex = 2
        Me.lblTOC.Text = "Table of Contents"
        Me.lblTOC.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblWatsonStudy
        '
        Me.lblWatsonStudy.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWatsonStudy.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.lblWatsonStudy.Location = New System.Drawing.Point(9, 27)
        Me.lblWatsonStudy.Name = "lblWatsonStudy"
        Me.lblWatsonStudy.Size = New System.Drawing.Size(152, 20)
        Me.lblWatsonStudy.TabIndex = 28
        Me.lblWatsonStudy.Text = "WatsonTM Study:"
        Me.lblWatsonStudy.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmdExit
        '
        Me.cmdExit.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdExit.CausesValidation = False
        Me.cmdExit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdExit.ForeColor = System.Drawing.Color.Red
        Me.cmdExit.Location = New System.Drawing.Point(218, 0)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(68, 25)
        Me.cmdExit.TabIndex = 6
        Me.cmdExit.Text = "E&xit"
        Me.cmdExit.UseVisualStyleBackColor = False
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.Label11.Location = New System.Drawing.Point(9, 52)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(118, 20)
        Me.Label11.TabIndex = 34
        Me.Label11.Text = "Report Title:"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblReportTitle
        '
        Me.lblReportTitle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblReportTitle.Enabled = False
        Me.lblReportTitle.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReportTitle.ForeColor = System.Drawing.Color.Black
        Me.lblReportTitle.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.lblReportTitle.Location = New System.Drawing.Point(155, 54)
        Me.lblReportTitle.Name = "lblReportTitle"
        Me.lblReportTitle.Size = New System.Drawing.Size(890, 40)
        Me.lblReportTitle.TabIndex = 35
        Me.lblReportTitle.Text = "Report Title:"
        '
        'cmdEdit
        '
        Me.cmdEdit.BackColor = System.Drawing.Color.Gray
        Me.cmdEdit.CausesValidation = False
        Me.cmdEdit.Enabled = False
        Me.cmdEdit.FlatAppearance.BorderColor = System.Drawing.Color.White
        Me.cmdEdit.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Navy
        Me.cmdEdit.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEdit.ForeColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(60, Byte), Integer), CType(CType(227, Byte), Integer))
        Me.cmdEdit.Location = New System.Drawing.Point(7, 0)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(68, 25)
        Me.cmdEdit.TabIndex = 3
        Me.cmdEdit.Text = "&Edit"
        Me.cmdEdit.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.BackColor = System.Drawing.Color.Gray
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.Color.ForestGreen
        Me.cmdSave.Location = New System.Drawing.Point(77, 0)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(68, 25)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'OleDbSelectCommand32
        '
        Me.OleDbSelectCommand32.CommandText = "SELECT boolDefault, charColumnLabel, id_tblHeaderLookup, id_tblReportTables, intO" & _
    "rder FROM tblConfigHeaderLookup"
        '
        'cbxStudy
        '
        Me.cbxStudy.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cbxStudy.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cbxStudy.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxStudy.Location = New System.Drawing.Point(155, 27)
        Me.cbxStudy.MaxDropDownItems = 20
        Me.cbxStudy.Name = "cbxStudy"
        Me.cbxStudy.Size = New System.Drawing.Size(386, 25)
        Me.cbxStudy.TabIndex = 0
        '
        'cmdRefresh
        '
        Me.cmdRefresh.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRefresh.ForeColor = System.Drawing.Color.Maroon
        Me.cmdRefresh.Location = New System.Drawing.Point(12, 69)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(117, 28)
        Me.cmdRefresh.TabIndex = 91
        Me.cmdRefresh.Text = "Refresh Study"
        Me.cmdRefresh.Visible = False
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'dgvUser
        '
        Me.dgvUser.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvUser.Location = New System.Drawing.Point(147, 69)
        Me.dgvUser.Name = "dgvUser"
        Me.dgvUser.Size = New System.Drawing.Size(102, 22)
        Me.dgvUser.TabIndex = 103
        Me.dgvUser.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(927, 57)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(70, 29)
        Me.Button1.TabIndex = 104
        Me.Button1.Text = "Test"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 500
        '
        'TimerRTC
        '
        Me.TimerRTC.Interval = 500
        '
        'cbxFilter
        '
        Me.cbxFilter.FormattingEnabled = True
        Me.cbxFilter.Location = New System.Drawing.Point(1021, 63)
        Me.cbxFilter.Name = "cbxFilter"
        Me.cbxFilter.Size = New System.Drawing.Size(83, 25)
        Me.cbxFilter.TabIndex = 110
        Me.cbxFilter.Visible = False
        '
        'cmdHook
        '
        Me.cmdHook.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdHook.CausesValidation = False
        Me.cmdHook.Enabled = False
        Me.cmdHook.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdHook.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.cmdHook.Location = New System.Drawing.Point(742, 46)
        Me.cmdHook.Name = "cmdHook"
        Me.cmdHook.Size = New System.Drawing.Size(73, 25)
        Me.cmdHook.TabIndex = 111
        Me.cmdHook.Text = "&Run Hook"
        Me.cmdHook.UseVisualStyleBackColor = False
        Me.cmdHook.Visible = False
        '
        'cbxExampleReport
        '
        Me.cbxExampleReport.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxExampleReport.Enabled = False
        Me.cbxExampleReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxExampleReport.ForeColor = System.Drawing.Color.Blue
        Me.cbxExampleReport.FormattingEnabled = True
        Me.cbxExampleReport.Location = New System.Drawing.Point(480, 26)
        Me.cbxExampleReport.MaxDropDownItems = 20
        Me.cbxExampleReport.Name = "cbxExampleReport"
        Me.cbxExampleReport.Size = New System.Drawing.Size(207, 23)
        Me.cbxExampleReport.TabIndex = 112
        Me.cbxExampleReport.Visible = False
        '
        'txtcbxMDBSelIndex
        '
        Me.txtcbxMDBSelIndex.Location = New System.Drawing.Point(860, 49)
        Me.txtcbxMDBSelIndex.Name = "txtcbxMDBSelIndex"
        Me.txtcbxMDBSelIndex.ReadOnly = True
        Me.txtcbxMDBSelIndex.Size = New System.Drawing.Size(33, 25)
        Me.txtcbxMDBSelIndex.TabIndex = 113
        Me.txtcbxMDBSelIndex.Text = "-10"
        Me.txtcbxMDBSelIndex.Visible = False
        Me.txtcbxMDBSelIndex.WordWrap = False
        '
        'txtFilterIndex
        '
        Me.txtFilterIndex.Location = New System.Drawing.Point(825, 49)
        Me.txtFilterIndex.Name = "txtFilterIndex"
        Me.txtFilterIndex.Size = New System.Drawing.Size(33, 25)
        Me.txtFilterIndex.TabIndex = 114
        Me.txtFilterIndex.Text = "-10"
        Me.txtFilterIndex.Visible = False
        Me.txtFilterIndex.WordWrap = False
        '
        'ms1
        '
        Me.ms1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ms1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAbout})
        Me.ms1.Location = New System.Drawing.Point(0, 0)
        Me.ms1.Name = "ms1"
        Me.ms1.Size = New System.Drawing.Size(1350, 25)
        Me.ms1.TabIndex = 115
        '
        'mnuAbout
        '
        Me.mnuAbout.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuMenuAbout, Me.mnuShowFC, Me.mnuMenuGenFC, Me.mnuTroubleshooting, Me.mnuHelp})
        Me.mnuAbout.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.mnuAbout.Name = "mnuAbout"
        Me.mnuAbout.Size = New System.Drawing.Size(53, 21)
        Me.mnuAbout.Text = "&Menu"
        '
        'mnuMenuAbout
        '
        Me.mnuMenuAbout.Name = "mnuMenuAbout"
        Me.mnuMenuAbout.Size = New System.Drawing.Size(248, 22)
        Me.mnuMenuAbout.Text = "&About..."
        '
        'mnuShowFC
        '
        Me.mnuShowFC.Name = "mnuShowFC"
        Me.mnuShowFC.Size = New System.Drawing.Size(248, 22)
        Me.mnuShowFC.Text = "Show Field Code Window..."
        '
        'mnuMenuGenFC
        '
        Me.mnuMenuGenFC.Name = "mnuMenuGenFC"
        Me.mnuMenuGenFC.Size = New System.Drawing.Size(248, 22)
        Me.mnuMenuGenFC.Text = "&Generate Field Code Report..."
        '
        'mnuTroubleshooting
        '
        Me.mnuTroubleshooting.Name = "mnuTroubleshooting"
        Me.mnuTroubleshooting.Size = New System.Drawing.Size(248, 22)
        Me.mnuTroubleshooting.Text = "Troubleshooting..."
        '
        'mnuHelp
        '
        Me.mnuHelp.Name = "mnuHelp"
        Me.mnuHelp.Size = New System.Drawing.Size(248, 22)
        Me.mnuHelp.Text = "Help..."
        '
        'dtp1
        '
        Me.dtp1.Location = New System.Drawing.Point(8, 715)
        Me.dtp1.Name = "dtp1"
        Me.dtp1.Size = New System.Drawing.Size(197, 25)
        Me.dtp1.TabIndex = 142
        Me.dtp1.Visible = False
        '
        'cmdSymbol
        '
        Me.cmdSymbol.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSymbol.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSymbol.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdSymbol.Location = New System.Drawing.Point(1195, 28)
        Me.cmdSymbol.Name = "cmdSymbol"
        Me.cmdSymbol.Size = New System.Drawing.Size(150, 33)
        Me.cmdSymbol.TabIndex = 144
        Me.cmdSymbol.Text = "Show Symbol Copy"
        Me.cmdSymbol.UseVisualStyleBackColor = True
        '
        'panEdit
        '
        Me.panEdit.CausesValidation = False
        Me.panEdit.Controls.Add(Me.cmdExit)
        Me.panEdit.Controls.Add(Me.cmdCancel)
        Me.panEdit.Controls.Add(Me.cmdEdit)
        Me.panEdit.Controls.Add(Me.cmdSave)
        Me.panEdit.Location = New System.Drawing.Point(693, 27)
        Me.panEdit.Name = "panEdit"
        Me.panEdit.Size = New System.Drawing.Size(286, 25)
        Me.panEdit.TabIndex = 145
        '
        'panChoose
        '
        Me.panChoose.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panChoose.AutoScroll = True
        Me.panChoose.Controls.Add(Me.cmdClearStudy)
        Me.panChoose.Controls.Add(Me.panFilterStudy)
        Me.panChoose.Controls.Add(Me.panWatsonData)
        Me.panChoose.Controls.Add(Me.panStudyFilter)
        Me.panChoose.Controls.Add(Me.cmdHomeCancel)
        Me.panChoose.Controls.Add(Me.cmdCreateReportTitle2)
        Me.panChoose.Controls.Add(Me.cmdReportHistory)
        Me.panChoose.Controls.Add(Me.cmdUpdateProject)
        Me.panChoose.Controls.Add(Me.cmdApplyTemplate)
        Me.panChoose.Controls.Add(Me.cmdShowOutstanding)
        Me.panChoose.Location = New System.Drawing.Point(-131, 13)
        Me.panChoose.Name = "panChoose"
        Me.panChoose.Size = New System.Drawing.Size(129, 604)
        Me.panChoose.TabIndex = 147
        '
        'cmdClearStudy
        '
        Me.cmdClearStudy.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdClearStudy.CausesValidation = False
        Me.cmdClearStudy.Enabled = False
        Me.cmdClearStudy.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdClearStudy.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdClearStudy.Location = New System.Drawing.Point(0, 381)
        Me.cmdClearStudy.Name = "cmdClearStudy"
        Me.cmdClearStudy.Size = New System.Drawing.Size(111, 44)
        Me.cmdClearStudy.TabIndex = 123
        Me.cmdClearStudy.Text = "&Clear Study"
        Me.cmdClearStudy.UseVisualStyleBackColor = False
        '
        'panWatsonData
        '
        Me.panWatsonData.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panWatsonData.Controls.Add(Me.gbSource)
        Me.panWatsonData.Controls.Add(Me.lblWatsonData)
        Me.panWatsonData.Controls.Add(Me.lblWatsonDataTitle)
        Me.panWatsonData.Location = New System.Drawing.Point(0, 34)
        Me.panWatsonData.Name = "panWatsonData"
        Me.panWatsonData.Size = New System.Drawing.Size(112, 58)
        Me.panWatsonData.TabIndex = 122
        '
        'lblWatsonData
        '
        Me.lblWatsonData.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblWatsonData.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWatsonData.Location = New System.Drawing.Point(2, 33)
        Me.lblWatsonData.Name = "lblWatsonData"
        Me.lblWatsonData.Size = New System.Drawing.Size(106, 22)
        Me.lblWatsonData.TabIndex = 4
        Me.lblWatsonData.Text = "Archived MDB"
        Me.lblWatsonData.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblWatsonDataTitle
        '
        Me.lblWatsonDataTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblWatsonDataTitle.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWatsonDataTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblWatsonDataTitle.Location = New System.Drawing.Point(0, 0)
        Me.lblWatsonDataTitle.Name = "lblWatsonDataTitle"
        Me.lblWatsonDataTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblWatsonDataTitle.Size = New System.Drawing.Size(110, 34)
        Me.lblWatsonDataTitle.TabIndex = 107
        Me.lblWatsonDataTitle.Text = "Watson Data Source"
        Me.lblWatsonDataTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'panStudyFilter
        '
        Me.panStudyFilter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panStudyFilter.Controls.Add(Me.gbStudyFilter)
        Me.panStudyFilter.Controls.Add(Me.Label10)
        Me.panStudyFilter.Location = New System.Drawing.Point(115, 472)
        Me.panStudyFilter.Name = "panStudyFilter"
        Me.panStudyFilter.Size = New System.Drawing.Size(112, 107)
        Me.panStudyFilter.TabIndex = 121
        '
        'gbStudyFilter
        '
        Me.gbStudyFilter.Controls.Add(Me.optStudyDocStudies)
        Me.gbStudyFilter.Controls.Add(Me.optStudyDocClosed)
        Me.gbStudyFilter.Controls.Add(Me.optStudyDocOpen)
        Me.gbStudyFilter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbStudyFilter.Location = New System.Drawing.Point(3, 32)
        Me.gbStudyFilter.Name = "gbStudyFilter"
        Me.gbStudyFilter.Size = New System.Drawing.Size(117, 69)
        Me.gbStudyFilter.TabIndex = 120
        Me.gbStudyFilter.TabStop = False
        '
        'optStudyDocStudies
        '
        Me.optStudyDocStudies.Checked = True
        Me.optStudyDocStudies.ForeColor = System.Drawing.Color.Black
        Me.optStudyDocStudies.Location = New System.Drawing.Point(0, 3)
        Me.optStudyDocStudies.Name = "optStudyDocStudies"
        Me.optStudyDocStudies.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optStudyDocStudies.Size = New System.Drawing.Size(118, 20)
        Me.optStudyDocStudies.TabIndex = 1
        Me.optStudyDocStudies.TabStop = True
        Me.optStudyDocStudies.Text = "All Studies"
        Me.optStudyDocStudies.UseVisualStyleBackColor = True
        '
        'optStudyDocClosed
        '
        Me.optStudyDocClosed.ForeColor = System.Drawing.Color.Black
        Me.optStudyDocClosed.Location = New System.Drawing.Point(0, 51)
        Me.optStudyDocClosed.Name = "optStudyDocClosed"
        Me.optStudyDocClosed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optStudyDocClosed.Size = New System.Drawing.Size(118, 20)
        Me.optStudyDocClosed.TabIndex = 3
        Me.optStudyDocClosed.Text = "Closed Studies"
        Me.optStudyDocClosed.UseVisualStyleBackColor = True
        '
        'optStudyDocOpen
        '
        Me.optStudyDocOpen.ForeColor = System.Drawing.Color.Black
        Me.optStudyDocOpen.Location = New System.Drawing.Point(0, 27)
        Me.optStudyDocOpen.Name = "optStudyDocOpen"
        Me.optStudyDocOpen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.optStudyDocOpen.Size = New System.Drawing.Size(118, 20)
        Me.optStudyDocOpen.TabIndex = 2
        Me.optStudyDocOpen.Text = "Open Studies"
        Me.optStudyDocOpen.UseVisualStyleBackColor = True
        '
        'panSampleRec
        '
        Me.panSampleRec.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panSampleRec.Controls.Add(Me.cmdSRecCancel)
        Me.panSampleRec.Controls.Add(Me.cmdInsertSRec)
        Me.panSampleRec.Controls.Add(Me.cmdDeletSRec)
        Me.panSampleRec.Location = New System.Drawing.Point(258, 353)
        Me.panSampleRec.Name = "panSampleRec"
        Me.panSampleRec.Size = New System.Drawing.Size(67, 120)
        Me.panSampleRec.TabIndex = 149
        '
        'panQAEvent
        '
        Me.panQAEvent.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panQAEvent.Controls.Add(Me.cmdQACancel)
        Me.panQAEvent.Controls.Add(Me.cmdInsertQAEvent)
        Me.panQAEvent.Controls.Add(Me.cmdDeleteQAEvent)
        Me.panQAEvent.Location = New System.Drawing.Point(96, 13)
        Me.panQAEvent.Name = "panQAEvent"
        Me.panQAEvent.Size = New System.Drawing.Size(74, 116)
        Me.panQAEvent.TabIndex = 149
        '
        'panMethVal
        '
        Me.panMethVal.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panMethVal.Controls.Add(Me.cmdMethValUpdate)
        Me.panMethVal.Controls.Add(Me.cmdMethValReset)
        Me.panMethVal.Location = New System.Drawing.Point(105, 252)
        Me.panMethVal.Name = "panMethVal"
        Me.panMethVal.Size = New System.Drawing.Size(84, 85)
        Me.panMethVal.TabIndex = 149
        '
        'panContr
        '
        Me.panContr.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panContr.Controls.Add(Me.cmdReplacePersonnel)
        Me.panContr.Controls.Add(Me.cmdCPCancel)
        Me.panContr.Controls.Add(Me.cmdCPAdd)
        Me.panContr.Controls.Add(Me.cmdCPDelete)
        Me.panContr.Location = New System.Drawing.Point(226, 171)
        Me.panContr.Name = "panContr"
        Me.panContr.Size = New System.Drawing.Size(85, 134)
        Me.panContr.TabIndex = 149
        '
        'panWordTemp
        '
        Me.panWordTemp.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panWordTemp.Controls.Add(Me.cmdRefreshStatements)
        Me.panWordTemp.Controls.Add(Me.cmdCancelReportStatements)
        Me.panWordTemp.Controls.Add(Me.cmdOpenReportStatements)
        Me.panWordTemp.Location = New System.Drawing.Point(-4, 11)
        Me.panWordTemp.Name = "panWordTemp"
        Me.panWordTemp.Size = New System.Drawing.Size(79, 270)
        Me.panWordTemp.TabIndex = 149
        '
        'panRepTables
        '
        Me.panRepTables.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panRepTables.AutoScroll = True
        Me.panRepTables.Controls.Add(Me.cmdApplyTables)
        Me.panRepTables.Controls.Add(Me.gbFilters)
        Me.panRepTables.Controls.Add(Me.gbRTC)
        Me.panRepTables.Controls.Add(Me.cmdCreateTable)
        Me.panRepTables.Controls.Add(Me.cmdRTConfigCancel)
        Me.panRepTables.Controls.Add(Me.cmdImportTables)
        Me.panRepTables.Controls.Add(Me.cmdAssignSamples)
        Me.panRepTables.Controls.Add(Me.cmdOutliers)
        Me.panRepTables.Controls.Add(Me.cmdDuplicateTables)
        Me.panRepTables.Controls.Add(Me.cmdSelect)
        Me.panRepTables.Controls.Add(Me.cmdAdvancedTable)
        Me.panRepTables.Controls.Add(Me.cmdViewAnalRuns)
        Me.panRepTables.Location = New System.Drawing.Point(-270, 10)
        Me.panRepTables.Name = "panRepTables"
        Me.panRepTables.Size = New System.Drawing.Size(133, 658)
        Me.panRepTables.TabIndex = 149
        '
        'cmdApplyTables
        '
        Me.cmdApplyTables.BackColor = System.Drawing.Color.Gainsboro
        Me.cmdApplyTables.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.cmdApplyTables.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdApplyTables.Location = New System.Drawing.Point(0, 597)
        Me.cmdApplyTables.Name = "cmdApplyTables"
        Me.cmdApplyTables.Size = New System.Drawing.Size(115, 62)
        Me.cmdApplyTables.TabIndex = 173
        Me.cmdApplyTables.Text = "&Apply Template Tables"
        Me.cmdApplyTables.UseVisualStyleBackColor = False
        '
        'panSumTable
        '
        Me.panSumTable.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panSumTable.Controls.Add(Me.cmdResetSummaryTable)
        Me.panSumTable.Controls.Add(Me.grbShowSummaryTable)
        Me.panSumTable.Controls.Add(Me.gbxMultVal)
        Me.panSumTable.Controls.Add(Me.cmdUpdateSummaryInfo)
        Me.panSumTable.Location = New System.Drawing.Point(84, 350)
        Me.panSumTable.Name = "panSumTable"
        Me.panSumTable.Size = New System.Drawing.Size(140, 330)
        Me.panSumTable.TabIndex = 149
        '
        'panAnalRuns
        '
        Me.panAnalRuns.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panAnalRuns.Controls.Add(Me.cmdAnaRunSumCancel)
        Me.panAnalRuns.Controls.Add(Me.cmdViewAnalyticalRuns1)
        Me.panAnalRuns.Controls.Add(Me.panColHeadings)
        Me.panAnalRuns.Location = New System.Drawing.Point(200, 49)
        Me.panAnalRuns.Name = "panAnalRuns"
        Me.panAnalRuns.Size = New System.Drawing.Size(74, 101)
        Me.panAnalRuns.TabIndex = 149
        '
        'panColHeadings
        '
        Me.panColHeadings.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panColHeadings.Controls.Add(Me.cmdRTHeaderConfigCancel)
        Me.panColHeadings.Controls.Add(Me.panAnalRefStds)
        Me.panColHeadings.Location = New System.Drawing.Point(0, 31)
        Me.panColHeadings.Name = "panColHeadings"
        Me.panColHeadings.Size = New System.Drawing.Size(88, 51)
        Me.panColHeadings.TabIndex = 149
        '
        'panAnalRefStds
        '
        Me.panAnalRefStds.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panAnalRefStds.Controls.Add(Me.cmdAnalRefCancel)
        Me.panAnalRefStds.Controls.Add(Me.cmdAddAnalyte)
        Me.panAnalRefStds.Controls.Add(Me.cmdAddRepAnalyte)
        Me.panAnalRefStds.Controls.Add(Me.cmdDeleteRepAnalyte)
        Me.panAnalRefStds.Controls.Add(Me.cmdCopyRepAnalyte)
        Me.panAnalRefStds.Location = New System.Drawing.Point(0, 26)
        Me.panAnalRefStds.Name = "panAnalRefStds"
        Me.panAnalRefStds.Size = New System.Drawing.Size(105, 277)
        Me.panAnalRefStds.TabIndex = 149
        '
        'panTopLevel
        '
        Me.panTopLevel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.panTopLevel.Controls.Add(Me.cmdDataCancel)
        Me.panTopLevel.Location = New System.Drawing.Point(102, 160)
        Me.panTopLevel.Name = "panTopLevel"
        Me.panTopLevel.Size = New System.Drawing.Size(67, 66)
        Me.panTopLevel.TabIndex = 149
        '
        'lblWarning
        '
        Me.lblWarning.BackColor = System.Drawing.Color.White
        Me.lblWarning.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblWarning.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWarning.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.lblWarning.Location = New System.Drawing.Point(8, 512)
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.Size = New System.Drawing.Size(226, 186)
        Me.lblWarning.TabIndex = 121
        Me.lblWarning.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TimerWarning
        '
        Me.TimerWarning.Interval = 1000
        '
        'lblWatsonWarning
        '
        Me.lblWatsonWarning.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWatsonWarning.Location = New System.Drawing.Point(8, 490)
        Me.lblWatsonWarning.Name = "lblWatsonWarning"
        Me.lblWatsonWarning.Size = New System.Drawing.Size(226, 22)
        Me.lblWatsonWarning.TabIndex = 150
        Me.lblWatsonWarning.Text = "Watson Check"
        Me.lblWatsonWarning.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panDot
        '
        Me.panDot.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panDot.Location = New System.Drawing.Point(210, 703)
        Me.panDot.Name = "panDot"
        Me.panDot.Size = New System.Drawing.Size(32, 19)
        Me.panDot.TabIndex = 151
        Me.panDot.Visible = False
        '
        'lblBlack
        '
        Me.lblBlack.BackColor = System.Drawing.Color.Black
        Me.lblBlack.Location = New System.Drawing.Point(1124, 25)
        Me.lblBlack.Name = "lblBlack"
        Me.lblBlack.Size = New System.Drawing.Size(44, 34)
        Me.lblBlack.TabIndex = 153
        Me.lblBlack.Text = "Label"
        '
        'mCal1
        '
        Me.mCal1.Location = New System.Drawing.Point(3, 3)
        Me.mCal1.Name = "mCal1"
        Me.mCal1.ScrollChange = 1
        Me.mCal1.TabIndex = 156
        '
        'cmdEnterCal
        '
        Me.cmdEnterCal.Font = New System.Drawing.Font("Segoe UI", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEnterCal.ForeColor = System.Drawing.Color.FromArgb(CType(CType(49, Byte), Integer), CType(CType(112, Byte), Integer), CType(CType(193, Byte), Integer))
        Me.cmdEnterCal.Location = New System.Drawing.Point(3, 166)
        Me.cmdEnterCal.Name = "cmdEnterCal"
        Me.cmdEnterCal.Size = New System.Drawing.Size(107, 25)
        Me.cmdEnterCal.TabIndex = 157
        Me.cmdEnterCal.Text = "Enter Date"
        Me.cmdEnterCal.UseVisualStyleBackColor = True
        '
        'cmdCalCancel
        '
        Me.cmdCalCancel.Font = New System.Drawing.Font("Segoe UI", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle))
        Me.cmdCalCancel.ForeColor = System.Drawing.Color.Red
        Me.cmdCalCancel.Location = New System.Drawing.Point(123, 166)
        Me.cmdCalCancel.Name = "cmdCalCancel"
        Me.cmdCalCancel.Size = New System.Drawing.Size(107, 25)
        Me.cmdCalCancel.TabIndex = 158
        Me.cmdCalCancel.Text = "Cancel"
        Me.cmdCalCancel.UseVisualStyleBackColor = True
        '
        'panCal
        '
        Me.panCal.Controls.Add(Me.mCal1)
        Me.panCal.Controls.Add(Me.cmdCalCancel)
        Me.panCal.Controls.Add(Me.cmdEnterCal)
        Me.panCal.Location = New System.Drawing.Point(8, 469)
        Me.panCal.Name = "panCal"
        Me.panCal.Size = New System.Drawing.Size(234, 196)
        Me.panCal.TabIndex = 150
        Me.panCal.Visible = False
        '
        'panPrepareReportInside
        '
        Me.panPrepareReportInside.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.panPrepareReportInside.Controls.Add(Me.MenuStrip1)
        Me.panPrepareReportInside.Location = New System.Drawing.Point(2, 2)
        Me.panPrepareReportInside.Name = "panPrepareReportInside"
        Me.panPrepareReportInside.Size = New System.Drawing.Size(126, 19)
        Me.panPrepareReportInside.TabIndex = 155
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.Gainsboro
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.MenuStrip1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.GripMargin = New System.Windows.Forms.Padding(0)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuPrepareReport})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Padding = New System.Windows.Forms.Padding(0)
        Me.MenuStrip1.Size = New System.Drawing.Size(126, 19)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MenuPrepareReport
        '
        Me.MenuPrepareReport.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PrepareEntireReportToolStripMenuItem, Me.PrepareOnlySelectedTableToolStripMenuItem, Me.PrepareOnlyReportBodyToolStripMenuItem, Me.PrepareOnlyReportTablesToolStripMenuItem})
        Me.MenuPrepareReport.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuPrepareReport.ForeColor = System.Drawing.Color.Blue
        Me.MenuPrepareReport.Name = "MenuPrepareReport"
        Me.MenuPrepareReport.Size = New System.Drawing.Size(123, 19)
        Me.MenuPrepareReport.Text = "Prepare a Report..."
        '
        'PrepareEntireReportToolStripMenuItem
        '
        Me.PrepareEntireReportToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.PrepareEntireReportToolStripMenuItem.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.PrepareEntireReportToolStripMenuItem.Name = "PrepareEntireReportToolStripMenuItem"
        Me.PrepareEntireReportToolStripMenuItem.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.PrepareEntireReportToolStripMenuItem.Size = New System.Drawing.Size(304, 28)
        Me.PrepareEntireReportToolStripMenuItem.Text = "Prepare Entire Report..."
        Me.PrepareEntireReportToolStripMenuItem.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'PrepareOnlySelectedTableToolStripMenuItem
        '
        Me.PrepareOnlySelectedTableToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.PrepareOnlySelectedTableToolStripMenuItem.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.PrepareOnlySelectedTableToolStripMenuItem.Name = "PrepareOnlySelectedTableToolStripMenuItem"
        Me.PrepareOnlySelectedTableToolStripMenuItem.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.PrepareOnlySelectedTableToolStripMenuItem.Size = New System.Drawing.Size(304, 28)
        Me.PrepareOnlySelectedTableToolStripMenuItem.Text = "Prepare Only Selected Section/Table..."
        '
        'PrepareOnlyReportBodyToolStripMenuItem
        '
        Me.PrepareOnlyReportBodyToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.PrepareOnlyReportBodyToolStripMenuItem.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.PrepareOnlyReportBodyToolStripMenuItem.Name = "PrepareOnlyReportBodyToolStripMenuItem"
        Me.PrepareOnlyReportBodyToolStripMenuItem.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.PrepareOnlyReportBodyToolStripMenuItem.Size = New System.Drawing.Size(304, 28)
        Me.PrepareOnlyReportBodyToolStripMenuItem.Text = "Prepare Only Report Body..."
        '
        'PrepareOnlyReportTablesToolStripMenuItem
        '
        Me.PrepareOnlyReportTablesToolStripMenuItem.Font = New System.Drawing.Font("Segoe UI", 10.0!)
        Me.PrepareOnlyReportTablesToolStripMenuItem.ForeColor = System.Drawing.Color.FromArgb(CType(CType(21, Byte), Integer), CType(CType(35, Byte), Integer), CType(CType(79, Byte), Integer))
        Me.PrepareOnlyReportTablesToolStripMenuItem.Name = "PrepareOnlyReportTablesToolStripMenuItem"
        Me.PrepareOnlyReportTablesToolStripMenuItem.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.PrepareOnlyReportTablesToolStripMenuItem.Size = New System.Drawing.Size(304, 28)
        Me.PrepareOnlyReportTablesToolStripMenuItem.Text = "Prepare Only Report Tables..."
        '
        'panPrepareReportOutside
        '
        Me.panPrepareReportOutside.BackColor = System.Drawing.Color.White
        Me.panPrepareReportOutside.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panPrepareReportOutside.Controls.Add(Me.panPrepareReportInside)
        Me.panPrepareReportOutside.Location = New System.Drawing.Point(555, 27)
        Me.panPrepareReportOutside.Name = "panPrepareReportOutside"
        Me.panPrepareReportOutside.Size = New System.Drawing.Size(132, 25)
        Me.panPrepareReportOutside.TabIndex = 156
        '
        'panActions
        '
        Me.panActions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.panActions.Controls.Add(Me.panRepTables)
        Me.panActions.Controls.Add(Me.panSumTable)
        Me.panActions.Controls.Add(Me.panSampleRec)
        Me.panActions.Controls.Add(Me.panChoose)
        Me.panActions.Controls.Add(Me.panQAEvent)
        Me.panActions.Controls.Add(Me.panTopLevel)
        Me.panActions.Controls.Add(Me.panMethVal)
        Me.panActions.Controls.Add(Me.panWordTemp)
        Me.panActions.Controls.Add(Me.panContr)
        Me.panActions.Controls.Add(Me.panAnalRuns)
        Me.panActions.Location = New System.Drawing.Point(240, 119)
        Me.panActions.Name = "panActions"
        Me.panActions.Size = New System.Drawing.Size(114, 605)
        Me.panActions.TabIndex = 157
        '
        'dgvFieldCodes
        '
        Me.dgvFieldCodes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvFieldCodes.Location = New System.Drawing.Point(1274, 69)
        Me.dgvFieldCodes.Name = "dgvFieldCodes"
        Me.dgvFieldCodes.Size = New System.Drawing.Size(69, 51)
        Me.dgvFieldCodes.TabIndex = 143
        Me.dgvFieldCodes.Visible = False
        '
        'lblActions
        '
        Me.lblActions.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblActions.Location = New System.Drawing.Point(236, 96)
        Me.lblActions.Name = "lblActions"
        Me.lblActions.Size = New System.Drawing.Size(118, 21)
        Me.lblActions.TabIndex = 158
        Me.lblActions.Text = "Actions"
        Me.lblActions.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panProgress
        '
        Me.panProgress.Controls.Add(Me.pb2)
        Me.panProgress.Controls.Add(Me.pb1)
        Me.panProgress.Controls.Add(Me.lblProgress)
        Me.panProgress.Location = New System.Drawing.Point(1274, 237)
        Me.panProgress.Name = "panProgress"
        Me.panProgress.Size = New System.Drawing.Size(68, 309)
        Me.panProgress.TabIndex = 159
        Me.panProgress.Visible = False
        '
        'frmHome_01
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 18)
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(229, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(249, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1350, 741)
        Me.Controls.Add(Me.panProgress)
        Me.Controls.Add(Me.lblActions)
        Me.Controls.Add(Me.panActions)
        Me.Controls.Add(Me.panPrepareReportOutside)
        Me.Controls.Add(Me.dgvFieldCodes)
        Me.Controls.Add(Me.cmdSymbol)
        Me.Controls.Add(Me.panCal)
        Me.Controls.Add(Me.lblBlack)
        Me.Controls.Add(Me.panDot)
        Me.Controls.Add(Me.lblWatsonWarning)
        Me.Controls.Add(Me.lblWarning)
        Me.Controls.Add(Me.tab1)
        Me.Controls.Add(Me.panEdit)
        Me.Controls.Add(Me.dtp1)
        Me.Controls.Add(Me.txtFilterIndex)
        Me.Controls.Add(Me.txtcbxMDBSelIndex)
        Me.Controls.Add(Me.llblAssignedSamples)
        Me.Controls.Add(Me.cbxExampleReport)
        Me.Controls.Add(Me.cmdHook)
        Me.Controls.Add(Me.cbxFilter)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.dgvUser)
        Me.Controls.Add(Me.cmdRefresh)
        Me.Controls.Add(Me.cbxStudy)
        Me.Controls.Add(Me.lblReportTitle)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.lblTOC)
        Me.Controls.Add(Me.lbxTab1)
        Me.Controls.Add(Me.lblWatsonStudy)
        Me.Controls.Add(Me.ms1)
        Me.DoubleBuffered = True
        Me.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.ms1
        Me.Name = "frmHome_01"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "frmout"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.tab1.ResumeLayout(False)
        Me.tp1.ResumeLayout(False)
        Me.tp1.PerformLayout()
        CType(Me.dgvReports, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panLockFinalReport.ResumeLayout(False)
        Me.panLockFinalReport.PerformLayout()
        CType(Me.dgStudies, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgHome, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvwStudy, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp2.ResumeLayout(False)
        Me.tp2.PerformLayout()
        CType(Me.dgvDataWatson, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabData.ResumeLayout(False)
        Me.tabData1.ResumeLayout(False)
        CType(Me.dgvDataCompany, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabData2.ResumeLayout(False)
        CType(Me.dgvStudyConfig, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabData3.ResumeLayout(False)
        Me.tabData3.PerformLayout()
        Me.gbInclude.ResumeLayout(False)
        CType(Me.dgvFC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabData4.ResumeLayout(False)
        Me.gbMeanComp.ResumeLayout(False)
        Me.gbMeanComp.PerformLayout()
        Me.gbCritPrecision.ResumeLayout(False)
        Me.gbCritPrecision.PerformLayout()
        Me.gbRound5.ResumeLayout(False)
        Me.gbRound5.PerformLayout()
        Me.tabData5.ResumeLayout(False)
        CType(Me.dgvAnalyteGroups, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp3.ResumeLayout(False)
        CType(Me.dgvAnalyticalRunSummary, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbxlblReviewAnalyticalRuns.ResumeLayout(False)
        Me.gbxlblReviewAnalyticalRuns.PerformLayout()
        Me.gbReportOptions.ResumeLayout(False)
        Me.panAnalRunSum.ResumeLayout(False)
        Me.panAnalRunSum.PerformLayout()
        Me.panAnalRunChoices.ResumeLayout(False)
        Me.panAnalRunChoices.PerformLayout()
        Me.tp4.ResumeLayout(False)
        Me.tp4.PerformLayout()
        CType(Me.dgvSummaryData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbxlblMethodValidation.ResumeLayout(False)
        Me.gbxlblMethodValidation.PerformLayout()
        Me.tp5.ResumeLayout(False)
        Me.tp5.PerformLayout()
        Me.gbxlblChooseEditWordTemplate.ResumeLayout(False)
        Me.gbxlblChooseEditWordTemplate.PerformLayout()
        Me.panSections.ResumeLayout(False)
        Me.panSections.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.gbSectionStyle.ResumeLayout(False)
        Me.gbSectionStyle.PerformLayout()
        Me.panRBSwb.ResumeLayout(False)
        Me.cmsHome.ResumeLayout(False)
        Me.grpRBS.ResumeLayout(False)
        Me.grpRBS.PerformLayout()
        CType(Me.dgvReportStatementWord, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvReportStatements, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp6.ResumeLayout(False)
        Me.tp6.PerformLayout()
        CType(Me.dgvGroups, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panTableGraphicExamples.ResumeLayout(False)
        Me.panTableGraphicExamples.PerformLayout()
        CType(Me.pbxTableGraphicExamples, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvReportTableConfiguration, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbxlblConfigureReportTables1.ResumeLayout(False)
        Me.tp7.ResumeLayout(False)
        Me.tp7.PerformLayout()
        Me.gbxlblConfigureColumnHeadings1.ResumeLayout(False)
        CType(Me.dgvReportTableHeaderConfig, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvReportTables, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp8.ResumeLayout(False)
        Me.tp8.PerformLayout()
        Me.gbxlblAnalyticalReferenceStd.ResumeLayout(False)
        Me.gbxlblAnalyticalReferenceStd.PerformLayout()
        CType(Me.dgvCompanyAnalRef, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvWatsonAnalRef, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp9.ResumeLayout(False)
        Me.gbxlblAddEditContributors.ResumeLayout(False)
        Me.gbxlblAddEditContributors.PerformLayout()
        CType(Me.dgvContributingPersonnel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp10.ResumeLayout(False)
        Me.tp10.PerformLayout()
        Me.gbxlblReviewValidatedMethod.ResumeLayout(False)
        Me.gbxlblReviewValidatedMethod.PerformLayout()
        CType(Me.dgvMethodValData, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbMethValApplyGuWu.ResumeLayout(False)
        CType(Me.dgvMethValExistingGuWu, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.gbMethodValMultiple.ResumeLayout(False)
        Me.gbMethodValMultiple.PerformLayout()
        Me.tp11.ResumeLayout(False)
        Me.tp11.PerformLayout()
        CType(Me.dgQATable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp12.ResumeLayout(False)
        Me.tp12.PerformLayout()
        Me.gbxlblSampleReceiptRecords2.ResumeLayout(False)
        Me.gbxlblSampleReceiptRecords2.PerformLayout()
        Me.gbxlblSampleReceiptRecords1.ResumeLayout(False)
        CType(Me.dgvSampleReceiptWatson, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSampleReceipt, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp13.ResumeLayout(False)
        CType(Me.pbxWord, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tp14.ResumeLayout(False)
        Me.tp15.ResumeLayout(False)
        Me.tp16.ResumeLayout(False)
        Me.tp16.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.gbFilters.ResumeLayout(False)
        Me.gbFilters.PerformLayout()
        Me.gbSource.ResumeLayout(False)
        Me.gbSource.PerformLayout()
        Me.panFilterStudy.ResumeLayout(False)
        Me.panFilterStudy.PerformLayout()
        Me.gbxMultVal.ResumeLayout(False)
        Me.gbxMultVal.PerformLayout()
        Me.grbShowSummaryTable.ResumeLayout(False)
        Me.grbShowSummaryTable.PerformLayout()
        Me.gbRTC.ResumeLayout(False)
        Me.gbRTC.PerformLayout()
        CType(Me.dgvUser, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ms1.ResumeLayout(False)
        Me.ms1.PerformLayout()
        Me.panEdit.ResumeLayout(False)
        Me.panChoose.ResumeLayout(False)
        Me.panWatsonData.ResumeLayout(False)
        Me.panStudyFilter.ResumeLayout(False)
        Me.gbStudyFilter.ResumeLayout(False)
        Me.panSampleRec.ResumeLayout(False)
        Me.panQAEvent.ResumeLayout(False)
        Me.panMethVal.ResumeLayout(False)
        Me.panContr.ResumeLayout(False)
        Me.panWordTemp.ResumeLayout(False)
        Me.panRepTables.ResumeLayout(False)
        Me.panSumTable.ResumeLayout(False)
        Me.panAnalRuns.ResumeLayout(False)
        Me.panColHeadings.ResumeLayout(False)
        Me.panAnalRefStds.ResumeLayout(False)
        Me.panTopLevel.ResumeLayout(False)
        Me.panCal.ResumeLayout(False)
        Me.panPrepareReportInside.ResumeLayout(False)
        Me.panPrepareReportInside.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.panPrepareReportOutside.ResumeLayout(False)
        Me.panActions.ResumeLayout(False)
        CType(Me.dgvFieldCodes, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panProgress.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region



    Private boolHold As Boolean
    Public boolStudyFired As Boolean = False

    'disable windows form close button
    Protected Overrides ReadOnly Property CreateParams() As CreateParams

        Get
            Dim param As CreateParams = MyBase.CreateParams
            param.ClassStyle = param.ClassStyle Or &H200
            Return param
        End Get

    End Property

    Public boolFromDataTab As Boolean = False

    Public boolStudyClick As Boolean = False

    Public gCalGrid As DataGridView
    Public boolCalGrid As Boolean = False

    Public boolOpened As Boolean = False


    'Private Sub frmHome_01_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

    '    '20161102 LEE: For some reason, me.refresh
    '    'Me.Refresh()

    '    'Call PositionProgress()

    '    'pesky

    '    Exit Sub

    '    Call SetPanAction()

    '    Me.panActions.BringToFront()

    '    Me.Refresh()

    'End Sub

    Private Sub frmHome_01_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Try
            'Me.afrRBS.Close()
            Try
                'Me.afrRBS.Close()
                frmH.wbRBS.Navigate("about:blank")

            Catch ex As Exception

            End Try
            'Me.wbRBS.Dispose()
        Catch ex As Exception
            Dim var1
            var1 = "Done" 'debuggins
        End Try

        Try
            Call DoExit()
        Catch ex As Exception

        End Try
    End Sub


    Sub FormLoad2()

        'Me.DoubleBuffered = True

        ''see modExtensionMethods
        'Call DoubleBufferControl(Me, "dgv")
        'Call modExtensionMethods.DoubleBufferedControl(Me.tab1, True)
        'Call modExtensionMethods.DoubleBufferedControl(Me, True)

        'Call SetButtonStylesParent(Me, "cmd", 1)

        Call ControlDefaults(Me)

        Cursor.Current = Cursors.AppStarting

        Call SetFormPos(Me)

        Dim x, y
        'x = Me.cmdLogin.Location.X
        'y = Me.cmdLogin.Location.Y
        'y = Me.cmdLogin.Location.Y ' + Me.cmdLogin.Top + Me.cmdLogin.Height

        'Cursor.Current = Cursors.Default

        'Me.Cursor = New Cursor(Cursor.Current.Handle)
        ''Cursor.Position = new system.drawing.point(Cursor.Position.X - 50, Cursor.Position.Y - 50)
        'Cursor.Position = New System.Drawing.Point(x + 10, y + 10)

        Dim str1 As String
        str1 = "Round 5 to even (Bankers' Rounding)" ' "Round 5 to even (mimics Watson LIMS" & ChrW(8482) & " rounding)"
        Me.rbRoundFiveEven.Text = str1

        str1 = "Round 5 away from zero (mimics  Watson LIMS" & ChrW(8482) & " and Excel ROUND function)"
        Me.rbRoundFiveAway.Text = str1

        'frmH.txtcbxMDBSelIndex.Text = 0

        boolMeRefresh = False
        boolFormLoad = False 'do this in action event

        Call SetPanPos()

        Call SizeCompanyAnalRef()

        Call SizeDot()

        Dim boolF As Boolean = boolFormLoad
        boolFormLoad = True
        Call ClearSelection(Me.dgvwStudy)
        boolFormLoad = boolF

        Me.AutoScroll = False
        Pause(0.25)
        Me.AutoScroll = True

        Me.lbl2.Size = Me.lbl1.Size
        Me.lbl2.Location = Me.lbl1.Location

        Dim strM As String
        strM = "A Method Validation study has been assigned to this Analyte." & ChrW(10) & "Therefore, the table entries associated with this Analyte are read-only."
        Me.lbl1.Text = strM

        strM = "The colored items are read-only for a Method Validation report." & ChrW(10) & "Click on the item to find where the value can be changed."
        Me.lbl2.Text = strM

        'frmH.cmdEdit.Enabled = False
        'frmH.cmdEdit.BackColor = System.Drawing.Color.Gray

        str1 = "&Retrieve" & ChrW(10) & "Watson" & ChrW(10) & "Study"
        Me.cmdUpdateProject.Text = str1

        Call UpdateRSW()

        Call LockSectionCheck()

        'pesky
        Call SetPanAction()

        Call SetStudyCount() '20190130 LEE:

    End Sub


    Private Sub lbxTab1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles lbxTab1.DrawItem

        'https://social.msdn.microsoft.com/Forums/en-US/aa2ba97a-5e93-4e7d-ab06-f7919939092a/listbox-items-line-spacing?forum=Vsexpressvb

        'e.Graphics.DrawString(lbxTab1.Items(e.Index).ToString, lbxTab1.Font, Brushes.Black, e.Bounds.Left, ((e.Bounds.Height - lbxTab1.Font.Height) \ 2) + e.Bounds.Top)

        Dim var1
        Try
            Dim drawBrush As New SolidBrush(Me.lbxTab1.ForeColor)
            e.Graphics.DrawString(lbxTab1.Items(e.Index).ToString, lbxTab1.Font, drawBrush, e.Bounds.Left, ((e.Bounds.Height - lbxTab1.Font.Height) \ 2) + e.Bounds.Top)
        Catch ex As Exception
            var1 = ex.Message
        End Try
     

    End Sub

    Private Sub lbxTab1_MeasureItem(sender As Object, e As MeasureItemEventArgs) Handles lbxTab1.MeasureItem

        'https://social.msdn.microsoft.com/Forums/en-US/aa2ba97a-5e93-4e7d-ab06-f7919939092a/listbox-items-line-spacing?forum=Vsexpressvb

        'itemheight at the default font settings is 20

        e.ItemHeight = 22


    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'https://social.msdn.microsoft.com/Forums/en-US/aa2ba97a-5e93-4e7d-ab06-f7919939092a/listbox-items-line-spacing?forum=Vsexpressvb
        '20170729 LEE: tried to add some padding between listbox items, but solution deactivates all other properties of the listbox, like selected-background-color
        'could probably do some more digging to address problem, but too much work for little benefit

        'Me.lbxTab1.DrawMode = DrawMode.OwnerDrawVariable

        '20160620 LEE: The code below moved to FormLoad2 and called in frmConsole just before frmh.visible=true
        'trying to address slow form loading



    End Sub

    Sub SizeDot()

        Dim a, b, c, d

        a = Me.lblWatsonWarning.Top
        Me.panDot.Top = a + 4
        Me.panDot.Height = Me.lblWatsonWarning.Height - 8
        Me.panDot.Width = Me.panDot.Height
        Me.panDot.Left = Me.lblWatsonWarning.Left + Me.lblWatsonWarning.Width - Me.panDot.Width - 4

        Me.panDot.BorderStyle = BorderStyle.None


    End Sub


    Sub SetPanPos()

        Dim var1

        Try
            Me.tab1.Appearance = TabAppearance.Buttons
            Me.tab1.ItemSize = New Size(0, 1)
        Catch ex As Exception
            var1 = ex.Message
        End Try
      
        Dim a, b, c, d
        Dim l, t, w, h

        'first set panActions width
        Me.panActions.Left = Me.lbxTab1.Left + Me.lbxTab1.Width + 5 ' Me.gbActions.Left
        Me.panActions.Top = Me.lbxTab1.Top ' Me.gbActions.Top

        l = Me.panActions.Left ' grA.Left
        t = Me.tab1.Left
        'grA.Width = t - l - 5
        Me.panActions.Width = t - l - 2.5

        a = Me.tab1.Top + Me.tab1.Height
        b = Me.panActions.Top
        c = a - b
        Me.panActions.Height = c

        l = 0 ' 2
        t = 0 ' 22
        'w = gbActions.Width - l
        w = Me.panActions.Width  ' - l

        h = Me.panActions.Height ' grA.Height 'NDL   -t

        Dim pan As Panel
        Dim btn As Button
        Dim ctl As Control

        Dim Count1 As Integer

        For Count1 = 1 To 12

            w = panActions.Width - 1

            Select Case Count1
                Case 1
                    pan = Me.panChoose
                Case 2
                    pan = Me.panTopLevel
                Case 3
                    pan = Me.panAnalRuns
                Case 4
                    pan = Me.panSumTable
                Case 5
                    pan = Me.panWordTemp
                Case 6
                    pan = Me.panRepTables
                Case 7
                    pan = Me.panColHeadings
                Case 8
                    pan = Me.panAnalRefStds
                Case 9
                    pan = Me.panContr
                Case 10
                    pan = Me.panMethVal
                Case 11
                    pan = Me.panQAEvent
                Case 12
                    pan = Me.panSampleRec
            End Select

            pan.Top = t
            pan.Left = l
            pan.Width = w ' - 5 '* 0.98
            pan.Height = h ' - 30 '* 0.9

            pan.Anchor = AnchorStyles.Bottom
            pan.Anchor = AnchorStyles.Top

            'size each button in pan
            w = pan.Width
            For Each ctl In pan.Controls
                If InStr(1, ctl.Name, "cmd", CompareMethod.Text) > 0 Or InStr(1, ctl.Name, "lbl", CompareMethod.Text) > 0 Then
                    If InStr(1, ctl.Name, "ClearFilter") > 0 Then
                        var1 = var1
                    Else
                        ctl.Left = 0 ' 2
                        ctl.Width = w ' - 4
                        ctl.BringToFront()
                    End If
                End If
            Next

            If Count1 = 1 Then

                'do the groupbox
                'Me.gbStudyFilter.Left = 1
                'Me.gbStudyFilter.Width = w - 2

                Me.panWatsonData.Left = 0
                Me.panWatsonData.Width = w

                Me.panStudyFilter.Left = 0 ' 1
                Me.panStudyFilter.Width = w ' - 2

                Me.panFilterStudy.Left = 0 ' 1
                Me.panFilterStudy.Width = w ' - 2

                Dim num1 As Single
                Dim num2 As Single

                num1 = Me.cmdReportHistory.Top - (Me.cmdShowOutstanding.Top + Me.cmdShowOutstanding.Height)

                Me.panStudyFilter.Top = Me.panFilterStudy.Top + Me.panFilterStudy.Height + num1

            ElseIf Count1 = 4 Then

                Me.gbxMultVal.Left = l
                Me.gbxMultVal.Width = w

                Me.grbShowSummaryTable.Left = 1
                Me.grbShowSummaryTable.Width = w

            ElseIf Count1 = 6 Then

                Me.gbRTC.Left = l
                Me.gbRTC.Width = w

                Me.gbFilters.Left = l
                Me.gbFilters.Width = w

            End If

        Next

        Call SetBlack()

    End Sub


    Sub SetBlack()

        Dim t, l, w, h

        t = Me.lbxTab1.Top
        l = Me.lbxTab1.Left
        w = Me.lbxTab1.Width
        h = Me.lbxTab1.Height

        Me.lblBlack.Top = t - 1
        Me.lblBlack.Left = l - 1
        Me.lblBlack.Height = h + 2
        Me.lblBlack.Width = w + 2

        Me.lblBlack.SendToBack()

        Me.lblBlack.Visible = False

    End Sub

    Sub SelectActionPan(intIndex As Short)

        Dim pan As Panel

        Dim Count1 As Integer = 12

        For Count1 = 1 To 12

            Select Case Count1
                Case 1
                    pan = Me.panChoose
                Case 2
                    pan = Me.panTopLevel
                Case 3
                    pan = Me.panAnalRuns
                Case 4
                    pan = Me.panSumTable
                Case 5
                    pan = Me.panWordTemp
                Case 6
                    pan = Me.panRepTables
                Case 7
                    pan = Me.panColHeadings
                Case 8
                    pan = Me.panAnalRefStds
                Case 9
                    pan = Me.panContr
                Case 10
                    pan = Me.panMethVal
                Case 11
                    pan = Me.panQAEvent
                Case 12
                    pan = Me.panSampleRec
            End Select

            If Count1 = intIndex Then
                pan.Height = Me.panActions.Height ' Me.gbActions.Height - 30 '* 0.9
                pan.Visible = True
            Else
                'pan.Height = 0
                pan.Visible = False
            End If

        Next

    End Sub


    Sub FormLoad()

        Call PositionProgress()

        'see modExtensionMethods
        Call modExtensionMethods.DoubleBufferedControl(Me, True)
        Call modExtensionMethods.DoubleBufferedControl(Me.tab1, True)
        Call DoubleBufferControl(Me, "dgv")

        Call CreateReturnLabelPosition()

        'if eval, then make boolEval true

        frmH = Me 'set this for later module reference

        Dim strM As String

        Dim var1, var2 As Object
        Dim Count1 As Short
        Dim int1 As Short
        Dim str1 As String
        Dim str1a As String
        Dim str2 As String
        Dim str3 As String
        Dim strSQL As String
        Dim ct1 As Short
        'Dim cn As New ADODB.Connection
        Dim wcn As New ADODB.Connection
        Dim str4 As String
        Dim rs As New ADODB.Recordset
        'Dim rs1 As New ADODB.Recordset
        Dim boolRO As Boolean
        Dim dv As System.Data.DataView
        Dim fld As ADODB.Field
        'Dim tblReports As System.Data.DataTable
        Dim conn As ADODB.Recordset
        'Dim boolAccess As Boolean
        'Dim bool64 As Boolean
        Dim strXSDpath As String
        Dim WatsonUserID, WatsonPswd, GuWuUserID, GuWuPswd
        'Dim fso As New Scripting.FileSystemObject
        Dim frm As New frmSplash1
        Dim pbmax As Short
        Dim pb As Short
        Dim dgv As DataGridView
        Dim drow As DataRow
        Dim boolW As Boolean


        'added stuff here

        'Exit Sub

        Dim w, h
        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        ''Me.Top = 0
        ''Me.Left = 0
        ''Me.Width = w
        ''Me.Height = h

        'Me.Top = h * 0.1
        'Me.Left = w * 0.1
        'Me.Width = w * 0.8
        'Me.Height = h * 0.8



        boolMsg = False
        Me.txtFilterIndex.Text = "Not Filtered"

        boolHold = False

        pb = 0
        pbmax = 15
        frm.pb1.Maximum = pbmax
        '1
        pb = pb + 1
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Value = pb
        frm.pb1.Visible = True
        frm.pb1.Refresh()
        frm.Show()
        frm.Refresh()
        'frm.Refresh()

        'record some stuff
        Sw = frm.pan1.Width
        Sh = frm.pan1.Height
        St = frm.pan1.Top
        Sl = frm.pan1.Left

        str1 = ChrW(8730) & " All"
        cmdRBSAll.Text = str1

        str1 = "Watson" & ChrW(8482) & " Study:"
        Me.lblWatsonStudy.Text = str1

        str1 = "? = Enter Yes or No" & ChrW(10)
        str1 = str1 & "A* = Check to include in report"
        Me.lblARS.Text = str1

        Me.rbOracle.Checked = True

        'investigate Command$ to get pathini
        Dim pathINI As String
        Dim boolCommand As Boolean
        Dim fso As FileInfo
        Dim strP As String ' = My.Computer.FileSystem.CurrentDirectory


        boolCommand = False
        pathINI = ""
        var1 = Command$()
        If var1 = "" Then
            'find connectionstrings
            'strP = "C:\Labintegrity\StudyDoc\Ini\GuWu.ini"
            strP = "C:\LabIntegrity\StudyDoc\Ini\StudyDoc.ini"
            If My.Computer.FileSystem.FileExists(strP) Then
            Else

                str1 = "The configured .ini file:" & ChrW(10) & ChrW(10)
                str1 = str1 & strP & ChrW(10) & ChrW(10)
                str1 = str1 & "does not seem to exist." & ChrW(10) & ChrW(10)
                str1 = str1 & "This startup will be terminated."
                MsgBox(str1, MsgBoxStyle.Critical, "Invalid .ini file...")
                End

                'strP = GetAppPath() & "GuWu.ini"
                'If My.Computer.FileSystem.FileExists(strP) Then
                'Else
                '    str1 = "The configured .ini file:" & ChrW(10) & ChrW(10)
                '    str1 = str1 & strP & ChrW(10) & ChrW(10)
                '    str1 = str1 & "does not seem to exist." & ChrW(10) & ChrW(10)
                '    str1 = str1 & "This startup will be terminated."
                '    MsgBox(str1, MsgBoxStyle.Critical, "Invalid .ini file...")
                '    End

                'End If
            End If

        Else
            boolCommand = True
            pathINI = Command$()
            'remove any quotations marks from pathini
            Dim str1b As String
            Dim str2b As String
            str1b = Mid(pathINI, 1, 1)
            If StrComp(str1b, Chr(34), vbTextCompare) = 0 Then
                str2b = Mid(pathINI, 2, Len(pathINI) - 1)
                pathINI = str2b
            End If
            str1b = Mid(pathINI, Len(pathINI), 1)
            If StrComp(str1b, Chr(34), vbTextCompare) = 0 Then
                str2b = Mid(pathINI, 1, Len(pathINI) - 1)
                pathINI = str2b
            End If

            strP = pathINI

        End If

        Cursor.Current = Cursors.AppStarting

        'str1 = strP & " \GuWu.ini"

        Dim objReader ' As New StreamReader(strP)

        Try
            objReader = New StreamReader(strP)
        Catch ex As Exception
            str1 = "The configured .ini file:" & ChrW(10) & ChrW(10)
            str1 = str1 & strP & ChrW(10) & ChrW(10)
            str1 = str1 & "does not seem to exist." & ChrW(10) & ChrW(10)
            str1 = str1 & "This startup will be terminated."
            MsgBox(str1, MsgBoxStyle.Critical, "Invalid .ini file...")
            End

            'str1 = "Hmmm. There was an error trying to access" & Chr(10) & Chr(10) & & strP & Chr(10) & Chr(10) & "Please contact your StudyDoc Administrator."
            'str2 = "Error finding GuWu.ini..."
            ''MsgBox(str1, MsgBoxStyle.Information, "Error finding GuWu.ini...")
            'GoTo end3
        End Try

        Dim sLine As String = ""
        Dim arrText As New ArrayList()
        Dim connectionstringGuWuAccess As String
        Dim connectionstringGuWu As String
        Dim connectionstringGuWuODBC As String
        Dim connectionstringWatson As String
        Dim ConnectionStringWatsonANSI As String
        Dim ConnectionStringWatsonAccess As String

        connectionstringGuWuAccess = ""
        connectionstringGuWu = ""
        connectionstringGuWuODBC = ""
        connectionstringWatson = ""
        ConnectionStringWatsonANSI = ""
        ConnectionStringWatsonAccess = ""
        boolAccess = False 'If Watson is an Access database
        bool64 = False 'If Watson is v6.4 or earlier
        boolGuWuAccess = False
        boolGuWuSQLServer = False
        boolGuWuOracle = False
        strXSDpath = ""


        Do
            sLine = objReader.ReadLine()
            If Not sLine Is Nothing And Len(sLine) <> 0 Then
                arrText.Add(sLine)
                int1 = InStr(1, sLine, Chr(9), CompareMethod.Text)
                If int1 = 0 Then
                Else
                    var1 = Mid(sLine, 1, int1 - 1)
                    var2 = Mid(sLine, int1 + 1, Len(sLine) - int1)
                    If StrComp(var1, "connectionstringGuWuAccess", CompareMethod.Text) = 0 Then
                        connectionstringGuWuAccess = var2
                    ElseIf StrComp(var1, "connectionstringGuWu", CompareMethod.Text) = 0 Then
                        connectionstringGuWu = var2 'str1a
                    ElseIf StrComp(var1, "connectionstringGuWuODBC", CompareMethod.Text) = 0 Then
                        connectionstringGuWuODBC = var2 'str1a
                    ElseIf StrComp(var1, "connectionstringWatson", CompareMethod.Text) = 0 Then
                        connectionstringWatson = var2 ' & ";uid=watson;pwd=476blue",str2
                    ElseIf StrComp(var1, "ConnectionStringWatsonANSI", CompareMethod.Text) = 0 Then
                        ConnectionStringWatsonANSI = var2 ' & ";uid=watson;pwd=gubbs",str3
                    ElseIf StrComp(var1, "ConnectionStringWatsonAccess", CompareMethod.Text) = 0 Then
                        ConnectionStringWatsonAccess = var2 'str4
                    ElseIf StrComp(var1, "boolAccess", CompareMethod.Text) = 0 Then
                        boolAccess = var2
                    ElseIf StrComp(var1, "bool64", CompareMethod.Text) = 0 Then
                        bool64 = var2
                    ElseIf StrComp(var1, "boolGuWuAccess", CompareMethod.Text) = 0 Then
                        boolGuWuAccess = var2
                    ElseIf StrComp(var1, "boolGuWuSQLServer", CompareMethod.Text) = 0 Then
                        boolGuWuSQLServer = var2
                    ElseIf StrComp(var1, "boolGuWuOracle", CompareMethod.Text) = 0 Then
                        boolGuWuOracle = var2
                    ElseIf StrComp(var1, "xsdpath", CompareMethod.Text) = 0 Then
                        strXSDpath = var2
                    ElseIf StrComp(var1, "GuWuUserID", CompareMethod.Text) = 0 Then
                        GuWuUserID = var2
                    ElseIf StrComp(var1, "GuWuPassword", CompareMethod.Text) = 0 Then
                        GuWuPswd = var2
                        Try
                            GuWuPswd = PasswordUnEncrypt(var2.ToString) ' Coding(Decode(var2, True), False)
                        Catch ex As Exception
                            strM = "There was a problem de-crypting the 'GuWuPassword' in the StudyDoc.ini file."
                            strM = strM & ChrW(10) & ChrW(10)
                            strM = strM & "Please inspect C:\LabIntegrity\StudyDoc\Ini\StudyDoc.ini."
                            strM = strM & ChrW(10) & ChrW(10)
                            strM = strM & "StudyDoc will close."
                            MsgBox(strM, vbInformation, "Fatal error...")
                            End
                        End Try

                    ElseIf StrComp(var1, "WATSONSCHEMAOWNER", CompareMethod.Text) = 0 Then
                        strSchema = var2
                    ElseIf StrComp(var1, "WatsonUserID", CompareMethod.Text) = 0 Then
                        WatsonUserID = var2
                    ElseIf StrComp(var1, "WatsonPassword", CompareMethod.Text) = 0 Then
                        WatsonPswd = var2
                        Try
                            WatsonPswd = PasswordUnEncrypt(var2.ToString) ' Coding(Decode(var2, True), False)
                        Catch ex As Exception

                            strM = "There was a problem de-crypting the 'GuWuPassword' in the StudyDoc.ini file."
                            strM = strM & ChrW(10) & ChrW(10)
                            strM = strM & "Please inspect C:\LabIntegrity\StudyDoc\Ini\StudyDoc.ini."
                            strM = strM & ChrW(10) & ChrW(10)
                            strM = strM & "StudyDoc will close."
                            MsgBox(strM, vbInformation, "Fatal error...")
                            End
                        End Try

                    End If

                    'WatsonID, WatsonPswd, GuWuID, GuWuPswd
                End If
            End If
        Loop Until sLine Is Nothing

errIni01:

        objReader.Close()

        'oConn.Open("Driver={Oracle ODBC Driver};" & _
        '   "Dbq=myDBName;" & _
        '   "Uid=myUsername;" & _
        '   "Pwd=myPassword")


        'bool64 = False
        'boolAccess = True
        'initiate GuWu database connection string
        'str1a = str1a & ";uid=" & GuWuUserID & ";pwd=" & GuWuPswd

        If boolGuWuAccess Then
            If Len(connectionstringGuWuAccess) = 0 Then
                str1 = "In the configured .ini file:" & ChrW(10) & ChrW(10)
                str1 = str1 & strP & ChrW(10) & ChrW(10)
                str1 = str1 & "the boolGuWuAccess line has been configured as TRUE; however, the connectionstringGuWuAccess line is null." & ChrW(10) & ChrW(10)
                str1 = str1 & "The connectionstringGuWuAccess line must contain a connection string."
                MsgBox(str1, MsgBoxStyle.Critical, "Invalid .ini file...")
                GoTo end1
            End If
        ElseIf boolGuWuSQLServer Then
            If Len(connectionstringGuWu) = 0 Then
                str1 = "In the configured .ini file:" & ChrW(10) & ChrW(10)
                str1 = str1 & strP & ChrW(10) & ChrW(10)
                str1 = str1 & "the boolGuWuSQLServer line has been configured as TRUE; however, the connectionstringGuWu line is null." & ChrW(10) & ChrW(10)
                str1 = str1 & "The connectionstringGuWu line must contain a connection string."
                MsgBox(str1, MsgBoxStyle.Critical, "Invalid .ini file...")
                GoTo end1
            End If
        ElseIf boolGuWuOracle Then
            If Len(connectionstringGuWu) = 0 Then
                str1 = "In the configured .ini file:" & ChrW(10) & ChrW(10)
                str1 = str1 & strP & ChrW(10) & ChrW(10)
                str1 = str1 & "the boolGuWuOracle line has been configured as TRUE; however, the connectionstringGuWu line is null." & ChrW(10) & ChrW(10)
                str1 = str1 & "The connectionstringGuWu line must contain a connection string."
                MsgBox(str1, MsgBoxStyle.Critical, "Invalid .ini file...")
                GoTo end1
            End If

            If Len(connectionstringGuWuODBC) = 0 Then
                str1 = "In the configured .ini file:" & ChrW(10) & ChrW(10)
                str1 = str1 & strP & ChrW(10) & ChrW(10)
                str1 = str1 & "the boolGuWuOracle line has been configured as TRUE; however, the connectionstringGuWuODBC line is null." & ChrW(10) & ChrW(10)
                str1 = str1 & "The connectionstringGuWuODBC line must contain a connection string."
                MsgBox(str1, MsgBoxStyle.Critical, "Invalid .ini file...")
                GoTo end1
            End If

        Else
            str1 = "The configured .ini file:" & ChrW(10) & ChrW(10)
            str1 = str1 & strP & ChrW(10) & ChrW(10)
            str1 = str1 & "does not seem to be a StudyDoc .ini file." & ChrW(10) & ChrW(10)
            str1 = str1 & "This startup will be terminated."
            MsgBox(str1, MsgBoxStyle.Critical, "Invalid .ini file...")
            End

        End If

        'If StrComp(Mid(connectionstringGuWu, Len(connectionstringGuWu), 1), ";", CompareMethod.Text) = 0 Then
        '    connectionstringGuWu = connectionstringGuWu & "User ID=" & GuWuUserID & ";password=" & GuWuPswd & ";"
        'Else
        '    connectionstringGuWu = connectionstringGuWu & ";User ID=" & GuWuUserID & ";password=" & GuWuPswd & ";"
        'End If

        If StrComp(Mid(connectionstringGuWu, Len(connectionstringGuWu), 1), ";", CompareMethod.Text) = 0 Then
            connectionstringGuWu = connectionstringGuWu & "UID=" & GuWuUserID & ";PWD=" & GuWuPswd & ";"
        Else
            connectionstringGuWu = connectionstringGuWu & ";UID=" & GuWuUserID & ";PWD=" & GuWuPswd & ";"
        End If

        'If StrComp(Mid(connectionstringGuWuODBC, Len(connectionstringGuWuODBC), 1), ";", CompareMethod.Text) = 0 Then
        '    connectionstringGuWuODBC = connectionstringGuWuODBC & "UID=" & GuWuUserID & ";PWD=" & GuWuPswd & ";"
        'Else
        '    connectionstringGuWuODBC = connectionstringGuWuODBC & ";UID=" & GuWuUserID & ";PWD=" & GuWuPswd & ";"
        'End If

        If StrComp(Mid(connectionstringGuWuODBC, Len(connectionstringGuWuODBC), 1), ";", CompareMethod.Text) = 0 Then
            connectionstringGuWuODBC = connectionstringGuWuODBC & "User ID=" & GuWuUserID & ";password=" & GuWuPswd & ";"
        Else
            connectionstringGuWuODBC = connectionstringGuWuODBC & ";User ID=" & GuWuUserID & ";password=" & GuWuPswd & ";"
        End If


        constrIni = connectionstringGuWu

        If boolGuWuAccess Then
            constrIni = connectionstringGuWuAccess
            ''''''''console.writeline(constrIni)
            Call ConfigAccess()
        ElseIf boolGuWuSQLServer Then
            constrIni = connectionstringGuWu
            ''20160527 LEE: set boolGuWuAccess to true because all calls are the same
            '20160531 LEE: Not true for dates
            'boolGuWuAccess = True
            Call ConfigSQLServer()
        ElseIf boolGuWuOracle Then
            constrIni = connectionstringGuWu
            constrIniGuWuODBC = connectionstringGuWuODBC
            Call ConfigOra()
        End If

        'initiate Watson database connection string

        If boolAccess Then
            constrWatson = ConnectionStringWatsonAccess
            ''''''''console.writeline(constrWatson)
            boolANSI = True
            boolW = False
            Me.rbArchive.Checked = True
            'Me.cmdBrowse.Visible = True
            Me.rbOracle.Enabled = False
            Me.panFilterStudy.Visible = False
            Me.panStudyFilter.Visible = False
        Else
            Me.rbOracle.Enabled = True
            Me.rbArchive.Enabled = False
            Me.panFilterStudy.Visible = True
            Me.panStudyFilter.Visible = True
            If bool64 Then 'Watson64
                'constrWatson = str2 & ";uid=" & WatsonUserID & ";pwd=" & WatsonPswd
                If StrComp(Mid(connectionstringWatson, Len(connectionstringWatson), 1), ";", CompareMethod.Text) = 0 Then
                    constrWatson = connectionstringWatson & "User ID=" & WatsonUserID & ";password=" & WatsonPswd & ";"
                Else
                    constrWatson = connectionstringWatson & ";User ID=" & WatsonUserID & ";password=" & WatsonPswd & ";"
                End If
                boolANSI = False
            Else 'Watson72

                'check for LI testing
                'if 
                'constrWatson = str3 & ";uid=" & WatsonUserID & ";pwd=" & WatsonPswd
                If StrComp(Mid(ConnectionStringWatsonANSI, Len(ConnectionStringWatsonANSI), 1), ";", CompareMethod.Text) = 0 Then
                    constrWatson = ConnectionStringWatsonANSI & "User ID=" & WatsonUserID & ";password=" & WatsonPswd & ";"
                Else
                    constrWatson = ConnectionStringWatsonANSI & ";User ID=" & WatsonUserID & ";password=" & WatsonPswd & ";"
                End If
                '''console.writeline(constrWatson)

                boolANSI = True
            End If
        End If

        '''''console.writeline(constrWatson) 'DEBUG

        constrCur = constrWatson
        conAccess97 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="


        'check for updates
        strUpdateMsg = ""
        boolUpdateCheckBad = True

        '20190131 LEE:
        'initialize strInstall
        intInstall = 0
        ReDim strInstall(2000000)

        Call UpdateCheck(boolGuWuAccess, boolGuWuOracle, boolGuWuSQLServer)

        If boolUpdateCheckBad Then 'continue
            End
        End If

        '20190131 LEE:
        Call WriteInstallDir()

        frm.Refresh()

        'for boolANSI testing

        'keep boolansi as false
        'bool64 = True
        'boolANSI = False

        boolRTCEnter = False

        'constr = "dsn=GuWu_01"
        'constrWatson = "dsn=Watson_01"
        ctAnalytes = 0
        ctAnalytes_IS = 0
        boolFromRTC = True

        'id_tblPersonnel = 0
        'id_tblUserAccounts = 0

        'initialize some button defaults
        rbShowIncludedRTConfig.Checked = True
        rbShowIncludedRBody.Checked = False
        'rbBodySectionFilterYes.Checked = True

        'rbShowIncludedSummaryTable.Checked = True
        rbShowAllSummaryTable.Checked = True

        Cursor.Current = Cursors.AppStarting

        boolW = True
        If boolAccess Then
            boolW = False
        Else

            frm.lblErr.Text = "...Establishing communication with the Watson" & ChrW(8482) & " Oracle" & ChrW(8482) & " database..."
            frm.lblErr.Refresh()
            Try
                'this is taking too long!
                wcn.Open(constrCur)
                frm.lblErr.Text = "...Communication with the Watson" & ChrW(8482) & " Oracle" & ChrW(8482) & " database established..."
                frm.lblErr.Refresh()
            Catch ex As Exception
                str1 = "Hmmm. There seems to be a problem connecting to the Watson datatabase."
                'str1 = str1 & Chr(10) & Chr(10) & "This is not a critical error. Users may still interact with archived .mdb files."
                str1 = str1 & Chr(10) & Chr(10) & ex.Message
                str1 = str1 & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
                str1 = str1 & Chr(10) & Chr(10) & "StudyDoc will close."
                str2 = "Oracle communication error..."

                MsgBox(str1, vbInformation, str2)
                End

                frm.lblErr.Text = str1
                frm.lblErr.Refresh()
                boolW = False
                Me.rbArchive.Checked = True
                'Me.rbArchive.Enabled = True
                'Me.cmdBrowse.Visible = True
                Me.rbOracle.Enabled = False
                Me.panFilterStudy.Enabled = False
                boolANSI = True
                Pause(1.5)
                'GoTo end3
            End Try

        End If

        Try
            If boolAccess Then
            Else
                Call GetWatsonUsers(wcn)
                Call GetWatsonStudyRoles(wcn)
            End If
        Catch ex As Exception

        End Try

        lblReportTitle.Text = ""
        str1 = "** If applicable. Enter numbers as numeric (e.g. 7 Days Refrigerated)." & ChrW(10)
        str1 = ""
        str1 = str1 & "* SA = If checked, requires sample assignment" & ChrW(10)
        str1 = str1 & "* B = If checked, table placeholder only will be created" & ChrW(10)
        str1 = str1 & "* P=Portrait, L=Landscape" & ChrW(10)
        str1 = str1 & "* Optional FC ID: Used to create Field Code" & ChrW(10)
        lblRTC.Text = str1

        'now size components of report table config page
        'goofy lblRTC seems to be moving itself

        'Me.gbxlblConfigureReportTables1.Top = Me.lblReportTableConfiguration.Top + Me.lblReportTableConfiguration.Height  ' Me.gbRTC.Top + Me.gbRTC.Height + 20

        Me.dgvGroups.Height = (Me.gbxlblConfigureReportTables1.Top + Me.gbxlblConfigureReportTables1.Height) + Me.lblReportTableConfiguration.Top - Me.dgvGroups.Top

        'Me.dgvReportTableConfiguration.Top = Me.gbxlblConfigureReportTables1.Top + Me.gbxlblConfiggureReportTables1.Height + 4
        Dim intHt As Int16
        intHt = cmdOrderReportTableConfig.Height + 4
        Me.dgvReportTableConfiguration.Top = Me.gbxlblConfigureReportTables1.Top + Me.gbxlblConfigureReportTables1.Height + 4 + intHt
        Me.dgvReportTableConfiguration.Height = Me.tp6.Height - Me.dgvReportTableConfiguration.Top - 1
        Me.cmdOrderReportTableConfig.Top = Me.dgvReportTableConfiguration.Top - Me.cmdOrderReportTableConfig.Height ' - 1
        'Me.cmdResize.Top = Me.dgvReportTableConfiguration.Top - Me.cmdResize.Height ' - 1
        Me.lblColoredRows.Top = Me.lblReportTableConfiguration.Top + Me.lblReportTableConfiguration.Height + 1 ' 20 ' - 1
        Me.chkTableName.Top = Me.lblColoredRows.Top '20 ' - 1
        Me.chkTableGraphicExamples.Top = Me.chkTableName.Top + Me.chkTableName.Height + 1
        Me.cmdResize.Left = Me.gbxlblConfigureReportTables1.Left + Me.gbxlblConfigureReportTables1.Width + 5
        Me.cmdResize.Top = Me.cmdOrderReportTableConfig.Top '(Me.lblRTC.Top + Me.lblRTC.Height) - Me.cmdResize.Height

        Dim intGraphicExamplesMargin As Short
        intGraphicExamplesMargin = 5

        'Configure example Tables option
        'Line up panel with table of reports
        Me.panTableGraphicExamples.Left = Me.dgvReportTableConfiguration.Left
        'Add margin for labels & picturebox
        Me.lblTableGraphicExamplesLabel.Left = intGraphicExamplesMargin
        Me.pbxTableGraphicExamples.Left = intGraphicExamplesMargin 'same as label
        'Make width of panel correct
        Me.panTableGraphicExamples.Width = dgvReportTableConfiguration.Width - 7 'Width is wider than shown for some reason, so -7
        'Make width of picturebox correct (with margin)
        Me.pbxTableGraphicExamples.Width = Me.panTableGraphicExamples.Width - (2 * intGraphicExamplesMargin)
        'Put text label beside "Example Table:" text
        Me.lblTableGraphicExamplesText.Left = lblTableGraphicExamplesLabel.Left + lblTableGraphicExamplesLabel.Width + 5
        Me.lblTableGraphicExamplesText.Width = Me.pbxTableGraphicExamples.Width - Me.lblTableGraphicExamplesText.Left


        ''set position of lblprogress
        Dim tp, lf, ht, wd, var3
  

        var1 = lblProgress.Height
        var2 = lblProgress.Top
        var3 = lblProgress.Width
        Cursor.Current = Cursors.AppStarting
        'frm.Refresh()

        Dim x1, x2
        Dim tp1, tp2
        Dim b1

        'lblProgress.Size = dgvwStudy.Size
        'Call PositionProgress()
        x1 = tab1.Left
        x2 = dgvwStudy.Left
        b1 = 0 ' 4
  

        Cursor.Current = Cursors.AppStarting
        'frm.Refresh()

        Call CreateQCTables()

        Dim boolEnd5 As Boolean
        boolEnd5 = True
        If boolGuWuOracle Then
            'initialize data adapters
            Call DAConnect(frm)
            'frm.Refresh()

            If DAsRefresh(frm) Then
            Else
                boolEnd5 = False
                GoTo end5
            End If
            frm.Refresh()

        ElseIf boolGuWuAccess Then
            'initialize data adapters
            Call DAConnectAcc(frm)
            'frm.Refresh()

            If DAsRefreshAcc(frm) Then
            Else
                boolEnd5 = False
                GoTo end5
            End If
            frm.Refresh()

        ElseIf boolGuWuSQLServer Then
            'initialize data adapters
            Call DAConnectSQLServer(frm)
            'frm.Refresh()

            If DAsRefreshSQLServer(frm) Then
            Else
                boolEnd5 = False
                GoTo end5
            End If
            frm.Refresh()
        End If

        Call CorrectActive() 'updates tblWordStatements

        Call DatabaseCorrections() 'updates database stuff

        ''debug
        'Dim Count11 As Int16
        'For Count11 = 0 To tblAssignedSamplesHelper.Rows.Count - 1
        '    str1 = tblAssignedSamplesHelper.Rows(Count11).Item("CHARHELPER")
        '    ''Console.WriteLine(str1)
        'Next

        Call FillDoPrepareTables()
        Cursor.Current = Cursors.AppStarting

        Call Create_tblQCStds()

        'fill a cbx
        Me.cbxFilterStudy.Items.Add("[None]")
        Me.cbxFilterStudy.Items.Add("Project ID")
        '20190124 LEE: Don't show species anymore - not useful
        'Me.cbxFilterStudy.Items.Add("Species")
        '20190124 LEE: add Study Type
        Me.cbxFilterStudy.Items.Add("Study Type")
        Me.cbxFilterStudy.Items.Add("Study Name")
        '20190124 LEE remove Study Number - it isn't displayed
        'Me.cbxFilterStudy.Items.Add("Study Number")
        Me.cbxFilterStudy.Items.Add("Study Title")

        Me.cbxFilterStudy.SelectedIndex = 0


        'record GDateFormat
        Dim tbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        tbl1 = tblConfiguration
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Table Date Format'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            GDateFormat = "MM/dd/yyyy"
        Else
            GDateFormat = var1
        End If
        LDateFormat = GDateFormat

        'record GTextDateFormat
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Text Date Format'"
        Erase rows1
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            GTextDateFormat = "MM/dd/yyyy"
        Else
            GTextDateFormat = var1
        End If
        LTextDateFormat = GTextDateFormat
        'seems that a YYYY gets in there somehow
        LTextDateFormat = Replace(LTextDateFormat, "YYYY", "yyyy", 1, -1, CompareMethod.Binary)

        'record boolUseHyperlinks
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Use Hyperlink Feature'"
        Erase rows1
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            boolUseHyperlinks = True
        Else
            If StrComp(var1, "FALSE", CompareMethod.Text) = 0 Then
                boolUseHyperlinks = False
            Else
                boolUseHyperlinks = True
            End If
        End If

        'record gintQCDec
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Default # of Decimals for QC Stats'"
        Erase rows1
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "")
        If Len(var1) = 0 Then
            gintQCDec = 1
        Else
            gintQCDec = CInt(var1)
        End If

        'gAllowExclSamples
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE LIKE 'Allow users to exclude data in StudyDoc*'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "FALSE")
        If Len(var1) = 0 Then
            gAllowExclSamples = False
        Else
            If StrComp(var1, "False", CompareMethod.Text) = 0 Then
                gAllowExclSamples = False
            Else
                gAllowExclSamples = True
            End If
        End If

        'gAllowGuWuAccCrit
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE LIKE 'Allow users to set QC and Calibr Std Acceptance Criteria*'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "FALSE")
        If Len(var1) = 0 Then
            gAllowGuWuAccCrit = False
        Else
            If StrComp(var1, "False", CompareMethod.Text) = 0 Then
                gAllowGuWuAccCrit = False
            Else
                gAllowGuWuAccCrit = True
            End If
        End If

        ''gGoToWord
        ''20160713 LEE: gGoToWord is deprecated
        'Erase rows1
        'str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE LIKE 'Go directly to Word" & ChrW(8482) & " after report generation.*'"
        'rows1 = tbl1.Select(str1)
        'var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "FALSE")
        'If Len(var1) = 0 Then
        '    gGoToWord = False
        'Else
        '    If StrComp(var1, "False", CompareMethod.Text) = 0 Then
        '        gGoToWord = False
        '    Else
        '        gGoToWord = True
        '    End If
        'End If
        '20160713 LEE: gGoToWord is deprecated
        gGoToWord = False

        'gboolET
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE LIKE 'Enable Word" & ChrW(8482) & " template management.*'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "FALSE")
        If Len(var1) = 0 Then
            gboolET = False
        Else
            If StrComp(var1, "False", CompareMethod.Text) = 0 Then
                gboolET = False
            Else
                gboolET = True
            End If
        End If

        'gboolER
        'Enable Generated Report management.
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE LIKE 'Enable Generated Report management.*'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "FALSE")

        If Len(var1) = 0 Then
            gboolER = False
        Else
            If StrComp(var1, "False", CompareMethod.Text) = 0 Then
                gboolER = False
            Else
                gboolER = True
            End If
        End If

        'boolReportGenAdvPrompt
        'Report Generation Advanced Prompt
        Erase rows1
        str1 = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE LIKE 'Report Generation Advanced Prompt*'"
        rows1 = tbl1.Select(str1)
        var1 = NZ(rows1(0).Item("CHARCONFIGVALUE"), "FALSE")
        If Len(var1) = 0 Then
            boolReportGenAdvPrompt = False
        Else
            If StrComp(var1, "False", CompareMethod.Text) = 0 Then
                boolReportGenAdvPrompt = False
            Else
                boolReportGenAdvPrompt = True
            End If
        End If

        'configure cbxDateFormat
        Dim dvDF As System.Data.DataView = New DataView(tblDateFormats, "ID_TBLDATEFORMATS > 0", "INTORDER ASC", DataViewRowState.CurrentRows)
        cbxDateFormat.DataSource = dvDF ' tblDateFormats.Select("ID_TBLDATEFORMATS > 0", "INTORDER ASC")
        cbxDateFormat.DisplayMember = tblDateFormats.Columns.Item("CHARFORMAT").ColumnName

        'fill cbxExampleReport
        For Count1 = 1 To 5
            Select Case Count1
                Case 1
                    str1 = "Prepare a Report..."
                Case 2
                    str1 = "Prepare Entire Report..."
                Case 3
                    str1 = "Prepare Only Selected Section/Table..."
                Case 4
                    str1 = "Prepare Only Report Body Section..."
                Case 5
                    str1 = "Prepare Only Report Table Section..."
                    'Case 6
                    '    str1 = "Prepare Report Body Showing Field Codes"
            End Select
            Me.cbxExampleReport.Items.Add(str1)
            'If Count1 < 5 Then
            Me.cbxExampleReport.Items.Add("")
            'Else
            'End If
        Next

        'select first item
        Me.cbxExampleReport.SelectedIndex = 0
        Me.cbxExampleReport.DropDownWidth = cbxExampleReport.Width * 1.5

        Cursor.Current = Cursors.AppStarting

        tblStudiesL = tblStudies

        'assign GSigFig
        Dim strF As String
        Dim rowsG() As DataRow
        strF = "CHARCONFIGGROUP = 'Global' AND CHARCONFIGCATEGORY = 'Global Settings' AND CHARCONFIGTITLE = 'Default Significant Figures for Conc Data'"
        rowsG = tblConfiguration.Select(strF)
        GSigFig = CInt(rowsG(0).Item("CHARCONFIGVALUE"))


        '2
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()
        Cursor.Current = Cursors.AppStarting

        boolCont = True

        'configure column mappings for dgvReports
        Call ReportsHomeInitialize()


        '3
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()
        Cursor.Current = Cursors.AppStarting

        If boolW Then
            Try
                Call Configure_dgvwStudy(boolW, wcn, boolANSI)
                If boolW Then
                    Try
                        Call ConfigStudyTable(True, True)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If

            Catch ex As Exception

            End Try
        End If


        '4
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()

        'configure cbxMethValExisting
        Me.cbxMethValExistingGuWu.Items.Clear()
        Dim intRows As Short
        Dim intRowsR As Short
        Dim rowsS() As DataRow
        Dim rowsR() As DataRow
        Dim strS As String
        Dim id As Int64
        strF = "ID_TBLCONFIGREPORTTYPE > 1 AND ID_TBLCONFIGREPORTTYPE < 5"
        Try
            rowsR = tblReports.Select(strF)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        intRowsR = rowsR.Length
        If intRowsR = 0 Then
        Else
            strF = ""
            For Count1 = 0 To intRowsR - 1
                id = rowsR(Count1).Item("ID_TBLSTUDIES")
                str1 = "id_tblStudies = " & id
                If Count1 = 0 Then
                    strF = str1
                Else
                    strF = strF & " OR " & str1
                End If
            Next
            'strF = "id_tblStudies > 0"
            strS = "CHARWATSONSTUDYNAME ASC"
            rowsS = tblStudies.Select(strF, strS)
            intRows = rowsS.Length
            Me.cbxMethValExistingGuWu.Items.Add("[NONE]")
            For Count1 = 0 To intRows - 1
                Me.cbxMethValExistingGuWu.Items.Add(rowsS(Count1).Item("charWatsonStudyName"))
            Next
        End If

        'first clear tab text
        int1 = tab1.TabPages.Count
        For Count1 = 0 To int1 - 1
            tab1.TabPages(Count1).Text = ""
        Next

        'fill lbxTab1

        str1 = "intForm = 1"
        str2 = "intOrder ASC"
        Dim drows() As DataRow
        drows = tblTab1.Select(str1, str2)
        ct1 = drows.Length
        For Count1 = 0 To ct1 - 1
            str1 = drows(Count1).Item("charItem").ToString
            Me.lbxTab1.Items.Add(str1)
        Next

        'initialize comboboxes in Contributing Personnel tab
        Call CPTab_Initialize()
        Cursor.Current = Cursors.AppStarting

        'create tables and fill the first column of all tables
        Call CreateTables()
        Cursor.Current = Cursors.AppStarting

        '5
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()
        Cursor.Current = Cursors.AppStarting


        '6
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()
        Cursor.Current = Cursors.AppStarting

        'fill some comboboxes
        Call FillStuff()
        Call FillHomeDropdownBoxes()
        If boolMsg Then
            'frm.Refresh()
            boolMsg = False
        End If
        Cursor.Current = Cursors.AppStarting

        'configure the following table
        'tblMethValExistingGuWu
        'add columns
        Dim ncol1 As New DataColumn
        str1 = "Analyte Name"
        str2 = "ColumnName"
        str3 = "System.String"
        ncol1.ColumnName = str2
        ncol1.DataType = System.Type.GetType(str3)
        ncol1.AllowDBNull = True
        tblMethValExistingGuWu.Columns.Add(ncol1)

        Dim ncol2 As New DataColumn
        str1 = "Watson Study"
        str2 = "WatsonStudy"
        str3 = "System.String"
        ncol2.ColumnName = str2
        ncol2.DataType = System.Type.GetType(str3)
        ncol2.AllowDBNull = True
        tblMethValExistingGuWu.Columns.Add(ncol2)

        Dim ncol3 As New DataColumn
        str1 = "id_tblStudies"
        str2 = str1
        str3 = "System.Int64"
        ncol3.ColumnName = str2
        ncol3.DataType = System.Type.GetType(str3)
        ncol3.AllowDBNull = True
        tblMethValExistingGuWu.Columns.Add(ncol3)

        Dim ncol4 As New DataColumn
        str1 = "CHARARCHIVEPATH"
        str2 = str1
        str3 = "System.String"
        ncol4.ColumnName = str2
        ncol4.DataType = System.Type.GetType(str3)
        ncol4.AllowDBNull = True
        tblMethValExistingGuWu.Columns.Add(ncol4)

        'add a row to tblMethValExistingGuWu

        drow = tblMethValExistingGuWu.NewRow
        drow.BeginEdit()
        drow("ColumnName") = "Value"
        drow.EndEdit()
        tblMethValExistingGuWu.Rows.Add(drow)
        dv = New DataView(tblMethValExistingGuWu)
        dv.AllowNew = False
        dgvMethValExistingGuWu.DataSource = dv
        For Count1 = 0 To 3
            Select Case Count1
                Case 0
                    str1 = "Analyte Name"
                    str2 = "ColumnName"
                Case 1
                    str1 = "Watson Study"
                    str2 = "WatsonStudy"
                Case 2
                    str1 = "id_tblStudies"
                    str2 = str1
                Case 3
                    str1 = "CHARARCHIVEPATH"
                    str2 = str1
            End Select
            dgvMethValExistingGuWu.Columns(Count1).HeaderText = str1
        Next
        Call ConfigMethValExistingGuWu()
        dgvMethValExistingGuWu.AutoResizeColumns()

        '7
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()
        Cursor.Current = Cursors.AppStarting

        'create tblTableN for recording table numbers in reports
        Call CreateTableN()

        'create tblAppendix others for recording appendix letters in reports
        Call CreatetblAppendix()
        Call CreatetblFigures()
        Call CreatetblAttachment()

        Cursor.Current = Cursors.AppStarting

        'initialize lbxReporthistory
        'Call ReportHistoryInitialize()'no longer exists
        Cursor.Current = Cursors.AppStarting

        'initialize ReportStatements tab

        Call ReportStatementGetStatementTitlesFromWord(frm)
        'frm.Refresh()

        '8
        pb = pb + 1
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Value = pb
        frm.pb1.Refresh()
        'frm.refresh()

        Call ReportStatmentInitialize()

        '9
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()
        Cursor.Current = Cursors.AppStarting

        'Dim dt As Date'for debugging
        'dt = Now
        ''''''''''''''''''console.writeline("Start ReportStatementsFill: " & dt)

        Call ReportStatementsFill() 'fill tblReportStatementsGuWu
        Cursor.Current = Cursors.AppStarting

        'dt = Now
        ''''''''''''''''''console.writeline("Start ReportStatementsFill: " & dt)

        '10
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()
        Cursor.Current = Cursors.AppStarting



        'initialize ReportTableHeader tab
        Call ReportTableHeaderConfig()

        '11
        pb = pb + 1
        frm.pb1.Value = pb
        frm.lblC.Text = pb & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        'frm.refresh()
        Cursor.Current = Cursors.AppStarting

        'initialize QATable tab
        Call QATableInitialize()
        Cursor.Current = Cursors.AppStarting

        'initialize SampleReceipt
        Call SampleReceiptInitialize()
        Cursor.Current = Cursors.AppStarting

        'initialize Summary Data
        Call InitializeSummaryData()
        Cursor.Current = Cursors.AppStarting

        'fill Data tab cbxs
        Call FillDataCbx()

        'fill cbxFilter and cbxRBSFilter
        Call cbxFilterPopulate()
        Call cbxRBSFilterPopulate()
        Call cbxRBSTypeFilterPopulate()
        Cursor.Current = Cursors.AppStarting
        'frm.Refresh()

        'record guest user
        str1 = GetStudyDocHeader(False)

        str1 = str1 & " v" & GetVersion() & gUserLabel ' " - User: Guest"
        Text = str1

        'select first item in lbxTab1
        Me.lbxTab1.SelectedIndex = 0

        'fill any hooks
        Dim tbl As System.Data.DataTable
        tbl = tblHooks
        int1 = tbl.Rows.Count
        'int1 = 0
        Dim int2 As Short
        For Count1 = 0 To int1 - 1
            str1 = NZ(tbl.Rows.Item(Count1).Item("CHARHOOK"), "")
            int2 = tbl.Rows.Item(Count1).Item("BOOLINCLUDE")
            If int2 = -1 Then 'continue
                Select Case str1
                    Case "CRLWor_AnalRefStandard"

                        'frmE.cmdOK.Visible = False
                        frm.lblErr.Text = "...Establishing communication with the " & str1 & " hook..."
                        'frmH.Text = "   Connecting..."
                        'frmH.pb1.Visible = False
                        'frmH.Show()
                        frm.lblErr.Visible = True
                        frm.lblErr.Refresh()
                        'frmE.TimerE.Start()
                        'Call frmE.RunTimer()
                        'frmE.Refresh()

                        Cursor.Current = Cursors.AppStarting

                        Call HookFill_CRL_AnalRefStandard()

                        Cursor.Current = Cursors.AppStarting

                        frm.lblErr.Text = ""
                        frm.lblErr.Refresh()
                        'frm.Refresh()

                        If boolHook1 Then
                            'configure cbxCompanyID
                            cbxCompanyID.DataSource = tblHook1

                            Try
                                cbxCompanyID.DisplayMember = tblHook1.Columns.Item("BottleID").ColumnName
                                tbl.Rows.Item(Count1).BeginEdit()
                                tbl.Rows.Item(Count1).Item("BOOLERROR") = 0
                            Catch ex As Exception
                                tbl.Rows.Item(Count1).BeginEdit()
                                tbl.Rows.Item(Count1).Item("BOOLERROR") = -1
                            End Try

                            tbl.Rows.Item(Count1).EndEdit()

                            If boolGuWuOracle Then
                                Try
                                    ta_tblHooks.Update(tblHooks)
                                Catch ex As Exception
                                    ds2005.TBLHOOKS.Merge(ds2005.TBLHOOKS, True)
                                End Try

                            ElseIf boolGuWuAccess Then
                                Try
                                    ta_tblHooksAcc.Update(tblHooks)
                                Catch ex As Exception
                                    ds2005Acc.TBLHOOKS.Merge(ds2005Acc.TBLHOOKS, True)
                                End Try
                            ElseIf boolGuWuSQLServer Then
                                Try
                                    ta_tblHooksSQLServer.Update(tblHooks)
                                Catch ex As Exception
                                    ds2005Acc.TBLHOOKS.Merge(ds2005Acc.TBLHOOKS, True)
                                End Try

                            End If

                        Else

                        End If
                End Select
            End If

            'var1 = tbl.Rows.item(0).Item("BOOLERROR")'for debugging

            Cursor.Current = Cursors.AppStarting
            'frm.Refresh()

            '12
            pb = pb + 1
            If pb > frm.pb1.Maximum Then
                pbmax = frm.pb1.Maximum + 5
                frm.pb1.Maximum = pbmax
            End If
            frm.pb1.Value = pb
            frm.lblC.Text = pb & " of " & pbmax
            frm.lblC.Refresh()
            frm.pb1.Refresh()

        Next
        Cursor.Current = Cursors.AppStarting

        ''position a label
        'frmH.gbxlblConfigureReportTables1.Top = frmH.dgvReportTableConfiguration.Top - frmH.gbxlblConfigureReportTables1.Height

        'pesky
        Call ResizeRows(Me.dgvCompanyAnalRef)
        Call ResizeRows(Me.dgvWatsonAnalRef)

        'do tblfieldcodes
        Call ResetFieldCodes(True)

        'set dgvFC
        Call FillFCRW()

        Call SizeCompanyAnalRef()

        Call CheckMaxID()

        Call ConfigLockFinalReport()

        '11
        pb = pb + 1
        frm.pb1.Value = pbmax
        frm.lblC.Text = pbmax & " of " & pbmax
        frm.lblC.Refresh()
        frm.pb1.Refresh()
        Cursor.Current = Cursors.AppStarting

        'ShowInTaskbar = True
        Call ToolTipSet()

end1:
        If rs.State = ADODB.ObjectStateEnum.adStateOpen Then
            rs.Close()
        End If

end2:
        rs = Nothing

        'record initial rbs values
        arrRBSColumns(4, 0) = 595 'Me.dgvReportStatements.Width
        arrRBSColumns(5, 0) = 2 'Me.dgvReportStatements.Left
        arrRBSColumns(6, 0) = 210 'Me.dgvReportStatementWord.Width
        arrRBSColumns(7, 0) = 600 'Me.dgvReportStatementWord.Left

        If Me.rbEntireReport.Checked Then
            Call ViewSections(False)
        Else
            Call ViewSections(True)
        End If

        boolLoad = False
        'boolFormLoad = False 'do this in action event
        boolFromRTC = False

        If boolAuditTrail Then
            Me.lblCFAuditTrail.Visible = True
        Else
            Me.lblCFAuditTrail.Visible = False
        End If

        'pesky
        Call SetPanAction()

        'lock all stuff
        Call DoThis("Logoff")

        Call Login()

        'close splash
        frm.Close()
        'frm.Visible = False
        'frm.Dispose()
        'frmE.Dispose()

        'clean up directories
        Call CleanUpDirs()

        Me.Refresh()
        Cursor.Current = Cursors.Default

        Me.cbxStudy.SelectedIndex = -1
        Me.dgvwStudy.ClearSelection()
        Me.dgvwStudy.CurrentCell = Nothing

        'misspelling in tblDropdownboxContent
        strF = "CHARVALUE = 'High Performance Liquid Chromatograhy - Mass Spectrometry'"
        str1 = "High Performance Liquid Chromatography - Mass Spectrometry"
        Dim rowsDB() As DataRow
        rowsDB = tblDropdownBoxContent.Select(strF)
        If rowsDB.Length = 0 Then
        Else
            rowsDB(0).BeginEdit()
            rowsDB(0).Item("CHARVALUE") = str1
            rowsDB(0).EndEdit()

            If boolGuWuOracle Then
                Try
                    ta_tblDropdownBoxContent.Update(tblDropdownBoxContent)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaMeth Val Tabl: " & ex.Message)
                    'ds2005.TBLMETHODVALIDATIONDATA.Merge('ds2005.TBLMETHODVALIDATIONDATA, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblDropdownBoxContentAcc.Update(tblDropdownBoxContent)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaMeth Val Tabl: " & ex.Message)
                    'ds2005Acc.TBLMETHODVALIDATIONDATA.Merge('ds2005Acc.TBLMETHODVALIDATIONDATA, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblDropdownBoxContentSQLServer.Update(tblDropdownBoxContent)
                Catch ex As DBConcurrencyException
                    ''msgbox("aaMeth Val Tabl: " & ex.Message)
                    'ds2005Acc.TBLMETHODVALIDATIONDATA.Merge('ds2005Acc.TBLMETHODVALIDATIONDATA, True)
                End Try
            End If

        End If


        'final control focus and mouse position determined at form_load


        GoTo end4

end3:

        MsgBox(str1, MsgBoxStyle.Critical, str2)

        Dim dt2 As Date
        Dim dt1 As Date
        dt2 = Now
        dt1 = DateAdd(DateInterval.Second, 1, dt2)
        Do Until dt2 > dt1
            dt2 = Now
        Loop
        End

        'lock all stuff
        Call DoThis("Logoff")

        'Cursor.Current = Cursors.Default

end4:
        Call ModMethChoice()
        Call FillArchivedMDB()

        Call LockAll(True, True)

        Call PositionProgress()

        Me.Visible = False

end5:

        'for some reason, this isn't firing in rb event
        'evaluate again
        If Me.rbArchive.Checked Then
            Me.lblWatsonData.Text = "Archived MDB"
        Else
            Me.lblWatsonData.Text = "Oracle"
        End If

        Me.dgvwStudy.ClearSelection()
        Me.dgvwStudy.CurrentCell = Nothing


        If wcn.State = ADODB.ObjectStateEnum.adStateOpen Then
            wcn.Close()

        End If
        wcn = Nothing

        'pesky
        Me.cbxStudy.SelectedIndex = -1
        Me.dgvwStudy.ClearSelection()
        Me.dgvwStudy.CurrentCell = Nothing

        'debug
        'MsgBox(Me.dgvwStudy.Rows.Count)

        If boolEnd5 Then
        Else
            End
        End If

    End Sub


    '    Sub DAConnect(ByVal frm As Form)

    '        On Error GoTo end1

    '        ta_tblSampleReceipt.Connection.ConnectionString = constrIni
    '        ta_tblData.Connection.ConnectionString = constrIni
    '        ta_tblTab1.Connection.ConnectionString = constrIni
    '        ta_tblConfiguration.Connection.ConnectionString = constrIni
    '        ta_tblOutstandingItems.Connection.ConnectionString = constrIni
    '        ta_tblPermissions.Connection.ConnectionString = constrIni
    '        ta_tblPersonnel.Connection.ConnectionString = constrIni
    '        ta_tblUserAccounts.Connection.ConnectionString = constrIni
    '        ta_tblAnalRefStandards.Connection.ConnectionString = constrIni
    '        ta_tblAnalyticalRunSummary.Connection.ConnectionString = constrIni
    '        ta_tblConfigBodySections.Connection.ConnectionString = constrIni
    '        ta_tblConfigHeaderLookup.Connection.ConnectionString = constrIni
    '        ta_tblConfigReportType.Connection.ConnectionString = constrIni
    '        ta_tblContributingPersonnel.Connection.ConnectionString = constrIni
    '        ta_tblCorporateAddresses.Connection.ConnectionString = constrIni
    '        ta_tblDataTableRowTitles.Connection.ConnectionString = constrIni
    '        ta_tblMaxID.Connection.ConnectionString = constrIni
    '        ta_tblMethodValidationData.Connection.ConnectionString = constrIni
    '        ta_tblQATables.Connection.ConnectionString = constrIni
    '        ta_tblReportHistory.Connection.ConnectionString = constrIni
    '        ta_tblReports.Connection.ConnectionString = constrIni
    '        ta_tblReportStatements.Connection.ConnectionString = constrIni
    '        ta_tblReportTable.Connection.ConnectionString = constrIni
    '        ta_tblReportTableAnalytes.Connection.ConnectionString = constrIni
    '        ta_tblReportTableHeaderConfig.Connection.ConnectionString = constrIni
    '        ta_tblStudies.Connection.ConnectionString = constrIni
    '        ta_tblTemplates.Connection.ConnectionString = constrIni
    '        ta_tblTemplateAttributes.Connection.ConnectionString = constrIni
    '        ta_tblConfigReportTables.Connection.ConnectionString = constrIni
    '        ta_tblAddressLabels.Connection.ConnectionString = constrIni
    '        ta_tblCorporateNickNames.Connection.ConnectionString = constrIni
    '        ta_tblDropdownBoxContent.Connection.ConnectionString = constrIni
    '        ta_tblDropdownBoxName.Connection.ConnectionString = constrIni
    '        ta_tblPasswordHistory.Connection.ConnectionString = constrIni
    '        ta_tblSummaryData.Connection.ConnectionString = constrIni
    '        ta_tblHooks.Connection.ConnectionString = constrIni
    '        ta_tblAssignedSamples.Connection.ConnectionString = constrIni
    '        ta_tblDateFormats.Connection.ConnectionString = constrIni
    '        ta_tblAssignedSamplesHelper.Connection.ConnectionString = constrIni
    '        ta_tblIncludedRows.Connection.ConnectionString = constrIni
    '        ta_tblConfigAppFigs.Connection.ConnectionString = constrIni
    '        ta_tblAppFigs.Connection.ConnectionString = constrIni

    '        ta_tblTableProperties.Connection.ConnectionString = constrIni
    '        ta_tblTableLegends.Connection.ConnectionString = constrIni

    '        ta_tblFieldCodes.Connection.ConnectionString = constrIni
    '        ta_tblReportHeaders.Connection.ConnectionString = constrIni
    '        ta_tblWordStatements.Connection.ConnectionString = constrIni

    '        'ta_tblWorddocs.Connection.ConnectionString = constrIni

    '        ta_tblReasonForChange.Connection.ConnectionString = constrIni
    '        ta_tblMeaningOfSig.Connection.ConnectionString = constrIni
    '        ta_tblSaveEvent.Connection.ConnectionString = constrIni
    '        ta_tblDataSystem.Connection.ConnectionString = constrIni
    '        ta_tblConfigCompliance.Connection.ConnectionString = constrIni
    '        '02218:
    '        ta_tblCustomFieldCodes.Connection.ConnectionString = constrIni
    '        '030008
    '        ta_TBLWORDSTATEMENTSVERSIONS.Connection.ConnectionString = constrIni
    '        '03000901
    '        'come back to this later
    '        'ta_TBLSECTIONTEMPLATES.Connection.ConnectionString = constrIni

    '        On Error GoTo 0

    '        Exit Sub

    'end1:
    '        Dim str1 As String
    '        Dim str2 As String
    '        If Err.Number <> 0 Then
    '            str1 = "Hmmm." & Chr(10) & "There seems to be a problem connecting to the StudyDoc datatabase."
    '            str1 = str1 & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
    '            str2 = "Critical communication error..."
    '            If boolFormLoad Then

    '                frm.Controls("lblErr").Text = str1
    '                frm.Controls("lblErr").Refresh()
    '            Else

    '                Call PositionProgress()
    '                Me.lblProgress.Text = str1
    '                Me.lblProgress.Visible = True
    '                Me.lblProgress.Refresh()
    '            End If

    '            MsgBox(str1, MsgBoxStyle.Critical, str2)
    '            Dim dt As Date
    '            Dim dt1 As Date
    '            dt = Now
    '            dt1 = DateAdd(DateInterval.Second, 1, dt)
    '            Do Until dt > dt1
    '                dt = Now
    '            Loop

    '            End

    '        End If
    '        On Error GoTo 0

    '    End Sub

    '    Sub DAConnectAcc(ByVal frm As Form)

    '        Try
    '            ''console.writeline(constrIni)
    '            ta_tblSampleReceiptAcc.Connection.ConnectionString = constrIni
    '            ta_tblDataAcc.Connection.ConnectionString = constrIni
    '            ta_tblTab1Acc.Connection.ConnectionString = constrIni
    '            ta_tblConfigurationAcc.Connection.ConnectionString = constrIni
    '            ta_tblOutstandingItemsAcc.Connection.ConnectionString = constrIni
    '            ta_tblPermissionsAcc.Connection.ConnectionString = constrIni
    '            ta_tblPersonnelAcc.Connection.ConnectionString = constrIni
    '            ta_tblUserAccountsAcc.Connection.ConnectionString = constrIni
    '            ta_tblAnalRefStandardsAcc.Connection.ConnectionString = constrIni
    '            ta_tblAnalyticalRunSummaryAcc.Connection.ConnectionString = constrIni
    '            ta_tblConfigBodySectionsAcc.Connection.ConnectionString = constrIni
    '            ta_tblConfigHeaderLookupAcc.Connection.ConnectionString = constrIni
    '            ta_tblConfigReportTypeAcc.Connection.ConnectionString = constrIni
    '            ta_tblContributingPersonnelAcc.Connection.ConnectionString = constrIni
    '            ta_tblCorporateAddressesAcc.Connection.ConnectionString = constrIni
    '            ta_tblDataTableRowTitlesAcc.Connection.ConnectionString = constrIni
    '            ta_tblMaxIDAcc.Connection.ConnectionString = constrIni
    '            ta_tblMethodValidationDataAcc.Connection.ConnectionString = constrIni
    '            ta_tblQATablesAcc.Connection.ConnectionString = constrIni
    '            ta_tblReportHistoryAcc.Connection.ConnectionString = constrIni
    '            ta_tblReportsAcc.Connection.ConnectionString = constrIni
    '            ta_tblReportStatementsAcc.Connection.ConnectionString = constrIni
    '            ta_tblReportTableAcc.Connection.ConnectionString = constrIni
    '            ta_tblReportTableAnalytesAcc.Connection.ConnectionString = constrIni
    '            ta_tblReportTableHeaderConfigAcc.Connection.ConnectionString = constrIni
    '            ta_tblStudiesAcc.Connection.ConnectionString = constrIni
    '            ta_tblTemplatesAcc.Connection.ConnectionString = constrIni
    '            ta_tblTemplateAttributesAcc.Connection.ConnectionString = constrIni
    '            ta_tblConfigReportTablesAcc.Connection.ConnectionString = constrIni
    '            ta_tblAddressLabelsAcc.Connection.ConnectionString = constrIni
    '            ta_tblCorporateNickNamesAcc.Connection.ConnectionString = constrIni
    '            ta_tblDropdownBoxContentAcc.Connection.ConnectionString = constrIni
    '            ta_tblDropdownBoxNameAcc.Connection.ConnectionString = constrIni
    '            ta_tblPasswordHistoryAcc.Connection.ConnectionString = constrIni
    '            ta_tblSummaryDataAcc.Connection.ConnectionString = constrIni
    '            ta_tblHooksAcc.Connection.ConnectionString = constrIni
    '            ta_tblAssignedSamplesAcc.Connection.ConnectionString = constrIni
    '            ta_tblDateFormatsAcc.Connection.ConnectionString = constrIni
    '            ta_tblAssignedSamplesHelperAcc.Connection.ConnectionString = constrIni
    '            ta_tblIncludedRowsAcc.Connection.ConnectionString = constrIni
    '            ta_tblConfigAppFigsAcc.Connection.ConnectionString = constrIni
    '            ta_tblAppFigsAcc.Connection.ConnectionString = constrIni

    '            ta_tblTablePropertiesAcc.Connection.ConnectionString = constrIni
    '            ta_tblTableLegendsAcc.Connection.ConnectionString = constrIni

    '            ta_tblFieldCodesAcc.Connection.ConnectionString = constrIni
    '            ta_tblReportHeadersAcc.Connection.ConnectionString = constrIni
    '            ta_tblWordStatementsAcc.Connection.ConnectionString = constrIni

    '            'ta_tblWorddocsAcc.Connection.ConnectionString = constrIni
    '            ta_tblAuditTrailAcc.Connection.ConnectionString = constrIni

    '            ta_tblReasonForChangeAcc.Connection.ConnectionString = constrIni
    '            ta_tblMeaningOfSigAcc.Connection.ConnectionString = constrIni
    '            ta_tblSaveEventAcc.Connection.ConnectionString = constrIni
    '            ta_tblDataSystemAcc.Connection.ConnectionString = constrIni
    '            ta_tblConfigComplianceAcc.Connection.ConnectionString = constrIni
    '            '02218:
    '            ta_tblCustomFieldCodesAcc.Connection.ConnectionString = constrIni
    '            '030008
    '            ta_TBLWORDSTATEMENTSVERSIONSAcc.Connection.ConnectionString = constrIni
    '            '03000901
    '            ta_TBLSECTIONTEMPLATESAcc.Connection.ConnectionString = constrIni
    '            '030030_01
    '            ta_TBLFINALREPORTAcc.Connection.ConnectionString = constrIni
    '            ta_TBLFINALREPORTWORDDOCSAcc.Connection.ConnectionString = constrIni


    '            'start Study Design
    '            ta_tblModulesAcc.Connection.ConnectionString = constrIni
    '            ta_TBLVERSIONAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUANIMALRECEIPTAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUCOMPOUNDSAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUCOMPOUNDSINDAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUCOMPOUNDTYPEAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUPROJECTSAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSPECIESAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSPECIESAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSTUDYDESIGNTYPEAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSTUDYSPECIESAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSTUDYSTATAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUASSAYPERSAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUASSAYAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSPECIESSTRAINAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUDOSEUNITSAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUPKGROUPSAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUPKROUTESAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUPKSUBJECTSAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWURTTIMEPOINTSAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUASSIGNEDCMPDAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUASSIGNEDCMPDLOTAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSTUDYSCHEDULINGAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUTPCONFIGAcc.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUTPNAMESCONFIGAcc.Connection.ConnectionString = constrIni
    '            'ta_QRYGUWUCALENDARAcc.Connection.ConnectionString = constrIni

    '            ta_TBLGUWUSTUDIESAcc.Connection.ConnectionString = constrIni

    '        Catch ex As Exception

    '            Dim str1 As String
    '            Dim str2 As String
    '            str1 = "Hmmm." & ChrW(10) & "There seems to be a problem connecting to the StudyDoc datatabase."
    '            str1 = str1 & ChrW(10) & ChrW(10) & "Please contact your StudyDoc system administrator."
    '            str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
    '            str2 = "Critical communication error..."
    '            If boolFormLoad Then

    '                frm.Controls("lblErr").Text = str1
    '                frm.Controls("lblErr").Refresh()
    '            Else
    '                Call PositionProgress()
    '                Me.lblProgress.Text = str1
    '                Me.lblProgress.Visible = True
    '                Me.lblProgress.Refresh()
    '            End If

    '            MsgBox(str1, MsgBoxStyle.Critical, str2)
    '            Dim dt As Date
    '            Dim dt1 As Date
    '            dt = Now
    '            dt1 = DateAdd(DateInterval.Second, 1, dt)
    '            Do Until dt > dt1
    '                dt = Now
    '            Loop

    '            End

    '        End Try



    '    End Sub

    '    Sub DAConnectSQLServer(ByVal frm As Form)

    '        Try
    '            ''console.writeline(constrIni)
    '            ta_tblSampleReceiptSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblDataSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblTab1SQLServer.Connection.ConnectionString = constrIni
    '            ta_tblConfigurationSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblOutstandingItemsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblPermissionsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblPersonnelSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblUserAccountsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblAnalRefStandardsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblAnalyticalRunSummarySQLServer.Connection.ConnectionString = constrIni
    '            ta_tblConfigBodySectionsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblConfigHeaderLookupSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblConfigReportTypeSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblContributingPersonnelSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblCorporateAddressesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblDataTableRowTitlesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblMaxIDSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblMethodValidationDataSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblQATablesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblReportHistorySQLServer.Connection.ConnectionString = constrIni
    '            ta_tblReportsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblReportStatementsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblReportTableSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblReportTableAnalytesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblReportTableHeaderConfigSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblStudiesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblTemplatesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblTemplateAttributesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblConfigReportTablesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblAddressLabelsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblCorporateNickNamesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblDropdownBoxContentSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblDropdownBoxNameSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblPasswordHistorySQLServer.Connection.ConnectionString = constrIni
    '            ta_tblSummaryDataSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblHooksSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblAssignedSamplesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblDateFormatsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblAssignedSamplesHelperSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblIncludedRowsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblConfigAppFigsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblAppFigsSQLServer.Connection.ConnectionString = constrIni

    '            ta_tblTablePropertiesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblTableLegendsSQLServer.Connection.ConnectionString = constrIni

    '            ta_tblFieldCodesSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblReportHeadersSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblWordStatementsSQLServer.Connection.ConnectionString = constrIni

    '            'ta_tblWorddocsSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblAuditTrailSQLServer.Connection.ConnectionString = constrIni

    '            ta_tblReasonForChangeSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblMeaningOfSigSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblSaveEventSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblDataSystemSQLServer.Connection.ConnectionString = constrIni
    '            ta_tblConfigComplianceSQLServer.Connection.ConnectionString = constrIni
    '            '02218:
    '            ta_tblCustomFieldCodesSQLServer.Connection.ConnectionString = constrIni
    '            '030008
    '            ta_TBLWORDSTATEMENTSVERSIONSSQLServer.Connection.ConnectionString = constrIni
    '            '03000901
    '            ta_TBLSECTIONTEMPLATESSQLServer.Connection.ConnectionString = constrIni
    '            '030030_01
    '            ta_TBLFINALREPORTSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLFINALREPORTWORDDOCSSQLServer.Connection.ConnectionString = constrIni


    '            'start Study Design
    '            ta_tblModulesSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLVERSIONSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUANIMALRECEIPTSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUCOMPOUNDSSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUCOMPOUNDSINDSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUCOMPOUNDTYPESQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUPROJECTSSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSPECIESSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSPECIESSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSTUDYDESIGNTYPESQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSTUDYSPECIESSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSTUDYSTATSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUASSAYPERSSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUASSAYSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSPECIESSTRAINSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUDOSEUNITSSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUPKGROUPSSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUPKROUTESSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUPKSUBJECTSSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWURTTIMEPOINTSSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUASSIGNEDCMPDSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUASSIGNEDCMPDLOTSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUSTUDYSCHEDULINGSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUTPCONFIGSQLServer.Connection.ConnectionString = constrIni
    '            ta_TBLGUWUTPNAMESCONFIGSQLServer.Connection.ConnectionString = constrIni
    '            'ta_QRYGUWUCALENDARSQLServer.Connection.ConnectionString = constrIni

    '            ta_TBLGUWUSTUDIESSQLServer.Connection.ConnectionString = constrIni

    '        Catch ex As Exception

    '            Dim str1 As String
    '            Dim str2 As String
    '            str1 = "Hmmm." & ChrW(10) & "There seems to be a problem connecting to the StudyDoc datatabase."
    '            str1 = str1 & ChrW(10) & ChrW(10) & "Please contact your StudyDoc system administrator."
    '            str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
    '            str2 = "Critical communication error..."
    '            If boolFormLoad Then

    '                frm.Controls("lblErr").Text = str1
    '                frm.Controls("lblErr").Refresh()
    '            Else
    '                Call PositionProgress()
    '                Me.lblProgress.Text = str1
    '                Me.lblProgress.Visible = True
    '                Me.lblProgress.Refresh()
    '            End If

    '            MsgBox(str1, MsgBoxStyle.Critical, str2)
    '            Dim dt As Date
    '            Dim dt1 As Date
    '            dt = Now
    '            dt1 = DateAdd(DateInterval.Second, 1, dt)
    '            Do Until dt > dt1
    '                dt = Now
    '            Loop

    '            End

    '        End Try



    '    End Sub

   

    Private Sub lbxTab1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lbxTab1.SelectedIndexChanged

        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim var1

        'MsgBox("Change")
        Dim strL As String

        If boolChooseReportWindow Then
            boolChooseReportWindow = False
            Exit Sub
        End If


        strL = Me.lbxTab1.Text

        Cursor.Current = Cursors.WaitCursor

        If StrComp(strL, "Audit Trail", CompareMethod.Text) = 0 Then

            Try
                Me.lbxTab1.SelectedIndex = intLBXTabPos
            Catch ex As Exception

            End Try
            Call ShowAuditTrail()

        Else

            intLBXTabPos = Me.lbxTab1.SelectedIndex

            'annoying!! must do the following
            Cursor.Current = Cursors.WaitCursor
            Call ViewSections(False)

            Cursor.Current = Cursors.WaitCursor
            'record selected row number
            int1 = lbxTab1.SelectedIndex
            int2 = tab1.TabPages.Count
            'select appropriate tab
            If int1 > int2 - 1 Then
            Else
                'boolHold = True
                'tab1.SelectedTab = tab1.TabPages.Item(int1)
                'boolHold = False

                Try
                    boolHold = True
                    tab1.SelectedTab = tab1.TabPages.Item(int1)
                    boolHold = False
                Catch ex As Exception
                    boolHold = True
                    tab1.SelectedTab = tab1.TabPages.Item(int1)
                    boolHold = False

                End Try
            End If

            If StrComp(strL, "Configure Report Tables", CompareMethod.Text) = 0 Then '.NET 4.6.1 thing
                Dim dgv As DataGridView
                dgv = Me.dgvReportTableConfiguration

                Cursor.Current = Cursors.WaitCursor

                Try
                    dgv.AutoResizeRows()
                Catch ex As Exception
                    var1 = ex.Message
                End Try

                Cursor.Current = Cursors.WaitCursor

                Try
                    dgv.AutoResizeColumns()
                Catch ex As Exception
                    var1 = ex.Message
                End Try

                Cursor.Current = Cursors.WaitCursor

                'pesky
                Cursor.Current = Cursors.WaitCursor
                Try
                    Dim nP2 As New Padding(0, 10, 0, 10)
                    Me.dgvReportTableConfiguration.DefaultCellStyle.Padding = nP2
                Catch ex As Exception
                    var1 = ex.Message
                End Try

                Cursor.Current = Cursors.WaitCursor
                Call OrderReportTableConfig()
                Cursor.Current = Cursors.WaitCursor

                Call SetComboCell(Me.dgvReportTableConfiguration, "CHARPAGEORIENTATION")
                Cursor.Current = Cursors.WaitCursor

            End If

            If InStr(1, strL, "Word Template", CompareMethod.Text) > 0 Then '.NET 4.6.1 thing
                Dim dgv As DataGridView
                dgv = Me.dgvReportStatementWord

                '.NET 4.6.1 thing
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            End If

            Cursor.Current = Cursors.WaitCursor
            If InStr(1, strL, "Validated Method", CompareMethod.Text) > 0 Then '.NET 4.6.1 thing
                Call ColorMethodValRows()
            End If

            Cursor.Current = Cursors.WaitCursor
            Call SelectActionPan(intLBXTabPos + 1)

        End If

        'Call ReportOptionChecks()

        'pesky
        Cursor.Current = Cursors.WaitCursor
        If StrComp(strL, "Review Analytical Runs", CompareMethod.Text) = 0 Then
            Call ReportOptionChecks(Me.chkAll.Checked)
        End If
        Cursor.Current = Cursors.WaitCursor
        Call ResizeFC()
        Cursor.Current = Cursors.WaitCursor
        If StrComp(strL, "Configure Report Tables", CompareMethod.Text) = 0 Then '.NET 4.6.1 thing
            Call AssessSampleAssignment()
        End If


        Cursor.Current = Cursors.WaitCursor
        Try
            Dim boolA As Boolean = BOOLASSIGNSAMPLES
            If boolA Then
                If frmH.cmdEdit.Enabled Then
                    frmH.cmdAssignSamples.Enabled = True
                Else
                    frmH.cmdAssignSamples.Enabled = False
                End If
            Else
                frmH.cmdAssignSamples.Enabled = False
            End If

        Catch ex As Exception

        End Try
        Cursor.Current = Cursors.WaitCursor
        Call HideWatsonRows()
        Cursor.Current = Cursors.WaitCursor
        Call UpdateRSW()
        Cursor.Current = Cursors.WaitCursor
        Call ConfigDropDowDGVs()
        Cursor.Current = Cursors.WaitCursor
        Call NumberAnalSumRows(Me.dgvAnalyticalRunSummary)
        Cursor.Current = Cursors.WaitCursor

        'pesky
        If StrComp(strL, "Configure Report Tables", CompareMethod.Text) = 0 Then '.NET 4.6.1 thing
            Try
                Dim nP1 As New Padding(0, 10, 0, 10)
                Me.dgvReportTableConfiguration.DefaultCellStyle.Padding = nP1
            Catch ex As Exception
                var1 = ex.Message
            End Try
        End If

        If StrComp(strL, "Review Validated Method", CompareMethod.Text) = 0 Then
            Try
                dgvMethodValData.AutoResizeColumns()
            Catch ex As Exception

            End Try
        End If

        Cursor.Current = Cursors.WaitCursor
        'pesky
        Call SetPanAction()

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub tab1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tab1.Click

        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short

        Cursor.Current = Cursors.WaitCursor

        int1 = tab1.SelectedIndex
        int2 = lbxTab1.Items.Count
        'select appropriate lbxTab1 item
        If int1 > int2 - 1 Then
        Else
            lbxTab1.SelectedIndex = int1
            Call SelectActionPan(int1 + 1)
        End If

        Cursor.Current = Cursors.Default

        Call setTableGraphicExample()

    End Sub

    Sub OpenOracleStudy()

        Dim boolCancel As Boolean
        Dim frmW As New frmBrowseWatson
        frmW.boolGetOracle = True
        frmW.boolArchive = False
        frmW.Text = "Retrieve a Study from the Watson" & ChrW(8482) & " Oracle database"
        frmW.ShowDialog()

        If frmW.boolCancel Then
            GoTo end1
        End If

end1:

    End Sub


    Private Sub cmdUpdateProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdateProject.Click

        Me.lblProgress.Text = "Opening a study..."
        Me.panProgress.Visible = True
        Me.panProgress.Refresh()

        'NDL Disable button until finished to fix SD-67 issue
        cmdUpdateProject.Enabled = False
        Call UpdateProjectClick_01()

        'NDL Then re-enable button.
        cmdUpdateProject.Enabled = True

        '20160226 LEE: Show/Hide Report Table Config StudyDoc table names
        Call ShowTableName()

        Call SetStudyCount() '20190130 LEE:

        Me.lblProgress.Text = ""
        Me.panProgress.Visible = False
        Me.panProgress.Refresh()

    End Sub


    Sub UpdateProjectClick_01()


        'check to see if tblwStudy
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Int64
        Dim int2 As Short

end3:

        'establish some global variables
        gWID = 0
        gWPID = 0
        boolNewOracle = False

        If frmH.rbOracle.Checked Then
            Call frmH.OpenOracleStudy()
        ElseIf Me.rbArchive.Checked Then
            Call frmH.OpenArchiveMDB(True)
        End If

        'now record this as a silent audit trail entry
        Call RecordStudyOpenAuditTrail()


        GoTo end2

end1:

end2:

    End Sub

    Sub RecordStudyOpenAuditTrail()

        Try


            ''Legend
            'ID_TBLAUDITTRAIL()
            'ID_TBLSAVEEVENT()
            'CHAROLDVALUE()
            'CHARNEWVALUE()
            'CHARTABLE()
            'CHARCOLUMN()
            'CHARACTION()
            'ID_SOURCETABLE()
            'CHARLINK1()
            'CHARLINK2()
            'CHARACTUALITEM()
            'DTSAVEDATE()
            'CHARUSERNAME()
            'CHARUSERID()
            'CHARTBLREASONFORCHANGE()
            'CHARTBLCHARMEANINGOFSIG()
            'id_tblStudies()
            'CHARWORKSTATION()
            'CHARTABLEDESCRIPTION()
            'CHARLINK1VALUE()
            'CHARLINK2VALUE()
            'CHARSTANDARDTIMEZONE()
            'CHARDAYLIGHTSAVINGZONE()
            'CHARDAYLIGHTSAVINGTIME()
            'CHARCOORUNIVTIME()
            'CHARUTCOFFSET()
            'CHARWATSONSTUDYNAME()
            'DTCOORUNIVTIME()
            'CHARAUDITTYPE()

            'get ID
            Dim strNA As String = "Not Applicable"
            Dim dt1 As Date = Now
            Dim maxid As Int64
            maxid = GetMaxID("TBLAUDITTRAIL", 1, True)
            maxid = maxid + 1
            '20190219 LEE: Don't need anymore. Used GetMaxID
            'Call PutMaxID("TBLAUDITTRAIL", CLng(maxid))

            Dim tblS As System.Data.DataTable = tblStudies
            Dim strStudyName As String = strNA
            Dim rowsS() As DataRow
            rowsS = tblS.Select("ID_TBLSTUDIES = " & id_tblStudies)
            If Len(gConfigStudy) = 0 Then
                If rowsS.Length = 0 Then
                    'Exit Sub
                    strStudyName = "Not Applicable"
                Else
                    strStudyName = NZ(rowsS(0).Item("CHARWATSONSTUDYNAME"), "Not Applicable")
                End If
            Else
                strStudyName = gConfigStudy
            End If

            ' Get the local time zone and the current local time and year.
            Dim localZone As TimeZone = TimeZone.CurrentTimeZone
            Dim currentDate As DateTime = dt1 'DateTime.Now
            Dim currentYear As Integer = currentDate.Year

            Dim strTimeZoneName As String
            Dim strDaylightName As String
            Dim boolDST As Boolean = False
            Dim strCUT As String
            Dim strOffset As String

            ' Display the names for standard time and daylight saving 
            ' time for the local time zone.
            ''''''''console.writeline(dataFmt, "Standard time name:", localZone.StandardName)
            strTimeZoneName = NZ(localZone.StandardName, "NA")
            ''''''''console.writeline(dataFmt, "Daylight saving time name:", localZone.DaylightName)
            strDaylightName = NZ(localZone.DaylightName, "NA")

            ' Display the current date and time and show if they occur 
            ' in daylight saving time.
            ''''''''console.writeline(vbCrLf & timeFmt, "Current date and time:", currentDate)
            ''''''''console.writeline(dataFmt, "Daylight saving time?", localZone.IsDaylightSavingTime(currentDate))
            boolDST = localZone.IsDaylightSavingTime(currentDate)
            ' Get the current Coordinated Universal Time (UTC) and UTC 
            ' offset.
            Dim currentUTC As DateTime = localZone.ToUniversalTime(currentDate)
            Dim currentOffset As TimeSpan = localZone.GetUtcOffset(currentDate)

            strCUT = Format(currentUTC, "MMM dd, yyyy HH:mm:ss tt")
            strOffset = currentOffset.ToString

            Dim dtbl As System.Data.DataTable = tblAuditTrail
            Dim nr As DataRow = dtbl.NewRow
            nr.BeginEdit()

            nr("ID_TBLAUDITTRAIL") = maxid
            nr("ID_TBLSAVEEVENT") = 1
            nr("CHAROLDVALUE") = "NA" ' "Old: User has opened this StudyDoc study"
            nr("CHARNEWVALUE") = "User has opened this StudyDoc study" '"New: User has opened this StudyDoc study"
            nr("CHARTABLE") = strNA
            nr("CHARCOLUMN") = strNA
            nr("CHARACTION") = "Study Doc study opened"
            nr("ID_SOURCETABLE") = 0
            nr("CHARLINK1") = "Where study name =" ' strNA
            nr("CHARLINK2") = strNA
            nr("CHARACTUALITEM") = strNA ' strStudyName
            nr("DTSAVEDATE") = dt1
            nr("CHARUSERNAME") = gUserName
            If gboolLDAP Then
                nr("CHARUSERID") = gUserID & " (Logged in from Network User ID " & gNetAcct & ")"
            Else
                nr("CHARUSERID") = gUserID
            End If
            nr("CHARTBLREASONFORCHANGE") = strNA
            nr("CHARTBLCHARMEANINGOFSIG") = strNA
            nr("ID_TBLSTUDIES") = id_tblStudies
            nr("CHARWORKSTATION") = gWorkstation
            nr("CHARTABLEDESCRIPTION") = strNA
            nr("CHARLINK1VALUE") = strStudyName 'strNA
            nr("CHARLINK2VALUE") = strNA
            nr("CHARSTANDARDTIMEZONE") = strTimeZoneName
            nr("CHARDAYLIGHTSAVINGZONE") = strDaylightName
            nr("CHARDAYLIGHTSAVINGTIME") = boolDST.ToString
            nr("CHARCOORUNIVTIME") = strCUT
            nr("CHARUTCOFFSET") = strOffset
            nr("CHARWATSONSTUDYNAME") = strStudyName
            nr("DTCOORUNIVTIME") = currentUTC
            nr("CHARAUDITTYPE") = "StudyDoc study opened"

            nr.EndEdit()

            tblAuditTrail.Rows.Add(nr)

            Call SaveAuditTrail()

        Catch ex As Exception

        End Try


    End Sub


    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        Call DoExit()

    End Sub

    Sub DoExit()

        Try
            Call DoThis("cmdExit")

        Catch ex As Exception

        End Try


        Me.Visible = False

        frmC.Visible = True

    End Sub

    Private Sub cmdAddRepAnalyte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddRepAnalyte.Click

        Dim frm As New frmAddReplicate
        Dim Count1 As Short
        Dim Count2 As Short
        Dim ct1 As Short
        Dim ct2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dt As System.Data.DataTable
        Dim dtC As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim rw As DataRow
        Dim col As DataColumn
        Dim strName As String
        Dim var1
        Dim int1 As Short
        Dim int2 As Short
        Dim boolC As Boolean
        Dim colRep As Short
        Dim dgv As DataGridView
        Dim boolF As Boolean

        dgv = Me.dgvCompanyAnalRef

        'Check for datatable existance
        'On Error Resume Next
        'strName = CType(dgCompanyAnalRef.DataSource, DataTable).TableName.ToString

        If ctAnalytes = 0 Then
            MsgBox("There are no analytes listed in the Analytical Reference Standard Table.", MsgBoxStyle.Information, "No analytes...")
            Err.Clear()
            On Error GoTo 0
            GoTo end1
        End If

        'get analyte info from frm.dgWatsonAnalRef
        'dv = dgCompanyAnalRef.DataSource
        dv = Me.dgvCompanyAnalRef.DataSource

        dtC = tblCompanyAnalRefTable
        'dtC = dgWatsonAnalRef.DataSource
        ct1 = dtC.Columns.Count
        'start with column 1
        ct2 = FindRowDVByCol("Is Replicate?", dv, "Item")
        For Count1 = 0 To ct1 - 1
            str1 = dtC.Columns.Item(Count1).ColumnName
            If StrComp(str1, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'ignore
            ElseIf StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'ignore
            ElseIf StrComp(str1, "Item", CompareMethod.Text) = 0 Then 'ignore
            Else
                'determine if column is replicate
                str2 = dv.Item(ct2).Item(str1)
                If StrComp(str2, "No", CompareMethod.Text) = 0 Then
                    'add to lbx
                    frm.lbxAnalytes.Items.Add(str1)
                End If
            End If
        Next
        frm.lblAddRep.Text = "Choose analyte for which a new lot will be added:"
        frm.Text = "Add Replicate"
        frm.ShowDialog(Me) 'show as modal
        int1 = frm.lbxAnalytes.SelectedIndex
        Dim strAnal As String
        strAnal = frm.lbxAnalytes.SelectedItem
        If int1 = -1 Or frm.boolStop Then 'cancel or no selection
            GoTo end1
        Else 'add an analyte

            'find column number of replicated analyte
            For Count1 = 0 To ct1 - 1
                str1 = dtC.Columns.Item(Count1).ColumnName
                If StrComp(str1, strAnal, CompareMethod.Text) = 0 Then
                    colRep = Count1
                    Exit For
                End If
            Next

            boolF = dgv.Columns.Item("Item").Frozen
            dgv.Columns.Item("Item").Frozen = False

            Dim col1 As New DataColumn
            'Dim gc1 As New DataGridTextBoxColumn
            'Dim ts1 As DataGridTableStyle
            'must add incrementer to distinguish
            boolC = False
            Count1 = 0
            int1 = dtC.Columns.Count
            Do Until boolC
                Count1 = Count1 + 1
                str1 = strAnal & "(" & Count1 & ")"
                boolC = True
                For Count2 = 0 To int1 - 1
                    str2 = dtC.Columns.Item(Count2).ColumnName
                    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                        boolC = False
                        Exit For
                    End If
                Next
            Loop
            'ts1 = dgCompanyAnalRef.TableStyles(0)
            col1.DataType = System.Type.GetType("System.String")
            col1.ColumnName = str1
            col1.Caption = str1 'strAnal
            dtC.Columns.Add(col1)


            'enter Is Replicate info in dtC
            int2 = FindRowDVByCol("Is Replicate?", dv, "Item")
            dtC.Rows.Item(int2).BeginEdit()
            dtC.Rows.Item(int2).Item(ct1) = "Yes"
            dtC.Rows.Item(int2).EndEdit()
            'enter Analyte Name info in dtC
            int2 = FindRowDVByCol("Analyte Name", dv, "Item")
            dtC.Rows.Item(int2).BeginEdit()
            dtC.Rows.Item(int2).Item(ct1) = strAnal
            dtC.Rows.Item(int2).EndEdit()
            'enter Is Coadministered Cmpd info in dtC
            int2 = FindRowDVByCol("Is Coadministered Cmpd?", dv, "Item")
            dtC.Rows.Item(int2).BeginEdit()
            dtC.Rows.Item(int2).Item(ct1) = dtC.Rows.Item(int2).Item(colRep)
            dtC.Rows.Item(int2).EndEdit()
            'enter Is Configured in Watson info in dtC
            int2 = FindRowDVByCol("Is Configured in Watson?", dv, "Item")
            dtC.Rows.Item(int2).BeginEdit()
            dtC.Rows.Item(int2).Item(ct1) = "No"
            dtC.Rows.Item(int2).EndEdit()
            'enter Analyte Parent info in dtC
            int2 = FindRowDVByCol("Analyte Parent", dv, "Item")
            dtC.Rows.Item(int2).BeginEdit()
            dtC.Rows.Item(int2).Item(ct1) = strAnal
            dtC.Rows.Item(int2).EndEdit()
            'enter Is Int Std info in dtC
            int2 = FindRowDVByCol("Is Internal Standard?", dv, "Item")
            dtC.Rows.Item(int2).BeginEdit()
            dtC.Rows.Item(int2).Item(ct1) = dtC.Rows.Item(int2).Item(colRep)
            dtC.Rows.Item(int2).EndEdit()

            dv = New DataView(dtC)
            Me.dgvCompanyAnalRef.DataSource = dv
            Me.dgvCompanyAnalRef.Refresh()
            Me.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
            Me.dgvCompanyAnalRef.AutoResizeColumns()

        End If

        str2 = AnalRefHook()
        If Len(str2) > 0 Then
            Select Case str2
                Case Is = "CRLWor_AnalRefStandard"
                    Call ComboBoxCRLAnalRefFill()
            End Select
        End If

        Call HideAnalRefRows()

        Me.dgvCompanyAnalRef.Columns.Item(ct1).Name = str1 'dtC.Columns.item(ct1).ColumnName
        Me.dgvCompanyAnalRef.Columns.Item(ct1).HeaderText = str1 'dtC.Columns.item(ct1).Caption

        'MsgBox("name: " & dtC.Columns.item(ct1).ColumnName & "  caption: " & dtC.Columns.item(ct1).Caption)
        'MsgBox("dgv name: " & me.dgvCompanyAnalRef.Columns.item(ct1).Name & "  headertext: " & me.dgvCompanyAnalRef.Columns.item(ct1).HeaderText)

        int1 = Me.dgvCompanyAnalRef.Columns.Count
        Me.dgvCompanyAnalRef.Columns.Item(int1 - 1).SortMode = DataGridViewColumnSortMode.NotSortable

        Call SyncCols(Me.dgvWatsonAnalRef, Me.dgvCompanyAnalRef)

        Call ResizeRows(Me.dgvCompanyAnalRef)
        Call ResizeRows(Me.dgvWatsonAnalRef)

end1:
        On Error GoTo 0
        dt = Nothing
        dtC = Nothing

        dgv.Columns.Item("Item").Frozen = boolF


    End Sub


    Private Sub cmdCPCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCPCancel.Click

        Call doCPCancel()

    End Sub


    Private Sub cbxSubmittedTo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSubmittedTo.SelectedIndexChanged

        Call FindNickname(cbxSubmittedTo, txtSubmittedTo)

    End Sub

    Private Sub cbxInSupportOf_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxInSupportOf.SelectedIndexChanged

        Call FindNickname(cbxInSupportOf, txtInSupportOf)

    End Sub

    Private Sub cbxSubmittedBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSubmittedBy.SelectedIndexChanged

        Call FindNickname(cbxSubmittedBy, txtSubmittedBy)

    End Sub

    Sub CreateReportTitle()

        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2, var3
        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strTitle As String
        Dim strcTitle As String
        Dim tbl As System.Data.DataTable
        Dim dg As DataGridView
        Dim arr1(100)
        Dim dgv As DataGridView
        Dim tblNick As System.Data.DataTable
        Dim rowsNick() As DataRow
        Dim tblR As System.Data.DataTable
        Dim rowsR() As DataRow
        Dim strF As String
        Dim dv As system.data.dataview
        Dim intR As Short
        Dim intRow As Short

        dv = dgvReports.DataSource
        int1 = dv.Count
        If int1 = 0 Then
            MsgBox("A report must be added.", MsgBoxStyle.Information, "Add a report...")
            'cmdConfigureReport.Select()
            Exit Sub
        End If

        'get record from tblReports
        strF = "id_tblstudies = " & id_tblStudies
        tblR = tblReports
        rowsR = tblR.Select(strF)
        intR = rowsR(0).Item("id_tblConfigReportType")

        'retrieve current Report Title
        int1 = dgvReports.CurrentRow.Index
        intRow = int1
        str1 = NZ(dgvReports.Item("charReportTitle", intRow).Value, "")

        'find technique adjective
        tbl = tblDropdownBoxContent
        'Dim dRows() As DataRow = tblStudiesL.Select(str1, "int_WatsonStudyID ASC")
        str1 = cbxAssayTechnique.Text
        If StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
            'MsgBox("An assay technique must be chosen", MsgBoxStyle.Information, "Choose an Assay Technique...")
            'GoTo end1
        End If
        str1 = "id_tblDropdownboxName = 3 AND charValue = '" & str1 & "'"

        Dim dRows() As DataRow = tbl.Select(str1)
        str2 = NZ(dRows(0).Item("CHARADJECTIVE"), "NO ADJECTIVE")

        Select Case intR
            Case 1 'Sample Analysis
                strTitle = str2 & " Analysis of "
            Case 2 'Method Validation
                strTitle = "Validation of a " & str2 & " Method for the Analysis of "
            Case 3 'Partial Method Validation
                strTitle = "Partial Validtion of a " & str2 & " Method for the Analysis of "
            Case 4 'Discovery
                strTitle = str2 & " Analysis of "
            Case 5 'Dose Form
                strTitle = str2 & " Analysis of "
        End Select

        'find analytes
        'dg = dgWatsonAnalRef
        dgv = Me.dgvWatsonAnalRef
        tbl = tblWatsonAnalRefTable
        int1 = tbl.Columns.Count
        int2 = FindRowDV("Is Internal Standard?", dgv.DataSource)
        Count2 = 0
        For Count1 = 1 To int1 - 1
            'str1 = dg.TableStyles(0).GridColumnStyles(Count1).HeaderText
            str1 = dgv.Columns.Item(Count1).HeaderText
            str2 = dgv.Item(Count1, int2).Value
            If StrComp(str2, "No", CompareMethod.Text) = 0 Then
                Count2 = Count2 + 1
                If Count2 > UBound(arr1) Then
                    ReDim Preserve arr1(UBound(arr1) + 10)
                End If
                arr1(Count2) = str1
            End If
        Next
        For Count1 = 1 To Count2
            str1 = arr1(Count1)
            strTitle = strTitle & str1 & " "
            If Count2 = Count1 Then
            Else
                strTitle = strTitle & "And "
            End If
        Next
        strTitle = strTitle & "in "

        'find anticoagulant
        str1 = cbxAnticoagulant.Text
        If StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
            'MsgBox("An anticoagulant must be chosen.", MsgBoxStyle.Information, "Choose an Anticoagulant...")
            'GoTo end1
        Else
            strTitle = strTitle & str1 & " buffered "
        End If

        'find species
        dg = dgvDataWatson
        int1 = FindRow("Species", tblWatsonData, "Item")
        str1 = dg(int1, 1).Value
        strTitle = strTitle & str1 & " "

        'find matrix
        int1 = FindRow("Matrix", tblWatsonData, "Item")
        str1 = dg(int1, 1).Value
        str1 = LowerCase(str1)
        str2 = Capit(str1)
        strTitle = strTitle & str2

        Select Case intR
            Case 1 'Sample Analysis
                strTitle = strTitle & " Samples"
                'find sponsor
                str1 = cbxSubmittedTo.Text
                If StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                Else
                    str2 = "charNickName = '" & str1 & "'"
                    tblNick = tblCorporateNickNames
                    rowsNick = tblNick.Select(str2)
                    var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                    str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = -1'
                    'str2 = "charNickname = '" & str1 & "' AND boolIncludeInTitle = -1 AND boolInclude = -1"
                    dRows = tblCorporateAddresses.Select(str2)
                    int1 = dRows.Length
                    strTitle = strTitle & " for "
                    For Count1 = 0 To int1 - 1
                        str1 = dRows(Count1).Item("charValue")
                        If Count1 = int1 - 1 Then
                            strTitle = strTitle & str1
                        Else
                            strTitle = strTitle & str1 & " "
                        End If
                    Next
                End If

                'find in support of
                str1 = cbxInSupportOf.Text
                If StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                Else
                    str2 = "charNickName = '" & str1 & "'"
                    tblNick = tblCorporateNickNames
                    rowsNick = tblNick.Select(str2)
                    var1 = rowsNick(0).Item("id_tblCorporateNickNames")
                    str2 = "id_tblCorporateNickNames = " & var1 & " AND boolIncludeInTitle = -1" ' & " AND boolInclude = -1"
                    'str2 = "charNickname = '" & str1 & "' AND boolIncludeInTitle = -1 AND boolInclude = -1"
                    dRows = tblCorporateAddresses.Select(str2)
                    int1 = dRows.Length
                    strTitle = strTitle & " In Support Of "
                    For Count1 = 0 To int1 - 1
                        str1 = dRows(Count1).Item("charValue")
                        If Count1 = int1 - 1 Then
                            strTitle = strTitle & str1
                        Else
                            strTitle = strTitle & str1 & " "
                        End If

                        'str1 = dRows(Count1).Item("charValue")
                        'strTitle = strTitle & str1
                    Next
                End If

                'find Study
                int1 = FindRow("Sponsor Study Number", tblCompanyData, "Item")
                str1 = NZ(tblCompanyData.Rows.Item(int1).Item(1), "")
                'str1 = "1"
                If Len(str1) = 0 Then
                    'MsgBox("Please note that a Sponsor Study Number has not been entered.", MsgBoxStyle.Information, "No Sponsor Study Number...")
                ElseIf StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
                Else
                    strTitle = strTitle & " Study " & str1
                End If
            Case 2 'Method Validation
            Case 3 'Partial Method Validation
            Case 4 'Discovery
            Case 5 'Dose Form
        End Select

        'check string length
        Dim strMod As String = "Choose Study & Template - Configured Reports Table"
        Dim strSource As String = "Report Title cell"

        If CheckColLenEx(strTitle, 255, strMod, strSource) Then
            GoTo end1
        End If

        strcTitle = NZ(dgvReports.Item("charReportTitle", dgvReports.CurrentRow.Index).Value, "")
        If Len(strcTitle) = 0 Then
            int1 = 1
        Else
            str1 = "The current Report Title is:" & Chr(13) & Chr(13) & strcTitle & Chr(13) & Chr(13)
            str1 = str1 & "Do you wish to replace this title with:" & Chr(13) & Chr(13)
            str1 = str1 & strTitle
            int1 = MsgBox(str1, MsgBoxStyle.OkCancel, "Record Report Title")
        End If
        If int1 = 1 Then
            lblReportTitle.Text = Replace(strTitle, "&", "&&", 1, -1)
            'dgvReports.ReadOnly = False
            dgv = dgvReports
            dgv("CHARREPORTTITLE", intRow).Value = strTitle
            dgv.AutoResizeRows()

            'rowsR(0).BeginEdit()
            'rowsR(0).Item("charReportTitle") = strTitle
            'rowsR(0).EndEdit()
            ''dgvReports.DataSource = dv
            'dv = tblReports.DefaultView
            'dv.RowFilter = strF
            'dv.AllowEdit = True
            'dv.AllowNew = False
            'dv.AllowDelete = False
            'dgvReports.DataSource = dv

        End If

end1:

    End Sub

    Private Sub cmdCreateReportTitle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreateReportTitle.Click
        Call CreateReportTitle()


    End Sub

    Private Sub cmdEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEdit.Click

        'ensure study is configured
        If id_tblStudies = 0 Then
            MsgBox("A study must first be loaded.", MsgBoxStyle.Information, "Load study first...")
            Exit Sub
        End If


        Call DoThis("Edit")
        'Me.Button1.Visible = True
    End Sub

    Function boolIsReport() As Boolean
        Dim dv As system.data.dataview
        dv = frmH.dgvReports.DataSource

        If dv.Count = 0 Then
            boolIsReport = False
        Else
            boolIsReport = True
        End If

    End Function

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim str1 As String
        Dim var1
        Dim Count1 As Short
        Dim int1 As Short

        'a report must be configured before a save action can occur
        If boolIsReport() Then
        Else
            str1 = "A report on the Home tab must be configured before a study can be saved."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            int1 = frmH.lbxTab1.Items.Count
            For Count1 = 0 To int1 - 1
                var1 = frmH.lbxTab1.Items(Count1).ToString
                If StrComp(var1, "Choose Study & Report", CompareMethod.Text) = 0 Then
                    Exit For
                End If
            Next
            frmH.lbxTab1.SelectedIndex = Count1
            Exit Sub
        End If


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

        Call DoThis("Save")

end1:

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Call PositionProgress()
        lblProgress.Text = "...cancelling..."
        'lblProgress.Visible = True
        lblProgress.Refresh()
        Me.panProgress.Visible = True
        Me.panProgress.Refresh()


        Call DoThis("Cancel")

        If Me.rbEntireReport.Checked Then
            Call ViewSections(False)
        Else
            Call ViewSections(True)
        End If

        'lblProgress.Visible = False
        'lblProgress.Refresh()
        Me.panProgress.Visible = False
        Me.panProgress.Refresh()

    End Sub

    Private Sub cmdConfigureReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConfigureReport.Click

        Dim dv As system.data.dataview
        Dim str1 As String
        Dim frm As New frmChooseStudy
        Dim strR As String

        frm.ShowDialog()
        Refresh()
        'SendKeys.Send("%")

        If frm.boolCancel Then
            MsgBox("Action canceled.", MsgBoxStyle.Information, "Action canceled...")
            Exit Sub
        End If
        Dim intRows As Short
        Dim Count1 As Short
        Dim tbl As System.Data.DataTable
        Dim varID
        tbl = tblConfigReportType
        intRows = frm.lvStudyType.Items.Count
        For Count1 = 0 To intRows - 1
            If frm.lvStudyType.Items(Count1).Checked Then
                Exit For
            End If
        Next
        varID = tbl.Rows.Item(Count1).Item("id_tblConfigReportType")
        strR = NZ(tbl.Rows.Item(Count1).Item("charReportType"), "Sample Analysis")

        frm.Dispose()

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID

        maxID = 1
        maxID = GetMaxID("tblReports", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
        'Call PutMaxID(tblReports, maxID)

        'If boolGuWuOracle Then
        '    ta_tblMaxID.Fill(tblMaxID)
        'ElseIf boolGuWuAccess Then
        '    ta_tblMaxIDAcc.Fill(tblMaxID)
        'ElseIf boolGuWuSQLServer Then
        '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        'End If
        'strFMax = "charTable = 'tblReports'"
        'tblMax = tblMaxID
        'rowsMax = tblMax.Select(strFMax)
        'maxID = rowsMax(0).Item("nummaxid")
        'maxID = maxID + 1
        'rowsMax(0).BeginEdit()
        'rowsMax(0).Item("nummaxid") = maxID
        'rowsMax(0).EndEdit()
        'If boolGuWuOracle Then
        '    Try
        '        ta_tblMaxID.Update(tblMaxID)
        '    Catch ex As DBConcurrencyException
        '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
        '    End Try
        'ElseIf boolGuWuAccess Then
        '    Try
        '        ta_tblMaxIDAcc.Update(tblMaxID)
        '    Catch ex As DBConcurrencyException
        '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '    End Try
        'ElseIf boolGuWuSQLServer Then
        '    Try
        '        ta_tblMaxIDSQLServer.Update(tblMaxID)
        '    Catch ex As DBConcurrencyException
        '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '    End Try
        'End If


        Dim tbl1 As System.Data.DataTable
        tbl1 = tblReports
        Dim r As DataRow = tbl1.NewRow
        r.BeginEdit()
        r.Item("id_tblReports") = maxID
        r.Item("id_tblStudies") = id_tblStudies
        r.Item("id_tblConfigReportType") = varID
        r.Item("charReportType") = strR
        r.EndEdit()
        tbl1.Rows.Add(r)

        'dv = dgvReports.DataSource
        'dv = tblReports.DefaultView
        dv = New DataView(tbl1)

        str1 = "id_tblStudies = " & id_tblStudies
        dv.RowFilter = str1
        dv.AllowNew = True
        'Dim drv As DataRowView
        'drv = dv.AddNew
        'drv.Item("id_tblReports") = maxID
        'drv.Item("id_tblStudies") = id_tblStudies
        'drv.Item("id_tblConfigReportType") = varID
        'drv.Item("charReportType") = strR
        dv.AllowNew = False
        dgvReports.DataSource = dv

        cmdConfigureReport.Enabled = False

        'fill HeaderFooter table
        Call FillHeaderFooterTable()


    End Sub

    Private Sub cmdHomeCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHomeCancel.Click

        Dim dv As System.Data.DataView

        Call DoHomeCancel()
        dv = dgvReports.DataSource
        If dv.Count = 0 Then
            cmdConfigureReport.Enabled = True
        Else
            cmdConfigureReport.Enabled = False
        End If

    End Sub

    Private Sub cmdDataCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDataCancel.Click
        Call DoDataCancel(True)

    End Sub

    Private Sub cmdAnaRunSumCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAnaRunSumCancel.Click
        DoAnalRunSumCancel(True)

    End Sub

    Private Sub cmdRTConfigCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRTConfigCancel.Click

        Cursor.Current = Cursors.WaitCursor
        Call DoRTConfigCancel()
        Call RTFilter()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmdAnalRefCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAnalRefCancel.Click
        Call DoAnalRefCancel()
    End Sub

    Private Sub cmdDeleteRepAnalyte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteRepAnalyte.Click

        Dim frm As New frmAddReplicate
        Dim Count1 As Short
        Dim Count2 As Short
        Dim ct1 As Short
        Dim ct2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dt As System.Data.DataTable
        Dim dtW As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim rw As DataRow
        Dim col As DataColumn
        Dim strName As String
        Dim var1
        Dim int1 As Short
        Dim int2 As Short
        Dim boolC As Boolean
        Dim dgv As DataGridView
        Dim boolF As Boolean

        dgv = Me.dgvCompanyAnalRef

        If ctAnalytes = 0 Then
            MsgBox("There are no analytes listed in the Analytical Reference Standard Table.", MsgBoxStyle.Information, "No analytes...")
            Err.Clear()
            On Error GoTo 0
            GoTo end1
        End If

        'get analyte info from frm.dgWatsonAnalRef
        'dv = dgCompanyAnalRef.DataSource
        dv = Me.dgvCompanyAnalRef.DataSource

        dtW = tblCompanyAnalRefTable
        'dtW = dgWatsonAnalRef.DataSource
        ct1 = dtW.Columns.Count
        'start with column 1
        ct2 = FindRowDVByCol("Is Configured in Watson?", dv, "Item")
        frm.lbxAnalytes.Items.Clear()
        For Count1 = 0 To ct1 - 1
            str1 = dtW.Columns.Item(Count1).ColumnName
            'determine if column is configured in Watson
            str2 = dv.Item(ct2).Item(str1)
            If StrComp(str2, "No", CompareMethod.Text) = 0 Then
                'add to lbx
                frm.lbxAnalytes.Items.Add(str1)
            End If
            'End If
        Next
        If frm.lbxAnalytes.Items.Count = 0 Then
            str1 = "All analytes are configured in Watson and may not be deleted from this grid."
            str1 = str1 & ChrW(10) & "Only replicates or coadministered compounds may be deleted."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        frm.lblAddRep.Text = "Choose analyte to delete." & ChrW(10) & "Note: not applicable to Watson-configured analytes."
        frm.Text = "Delete Replicate..."
        frm.ShowDialog(Me) 'show as modal
        If frm.boolStop Then 'stop
            GoTo end1
        End If
        int1 = frm.lbxAnalytes.SelectedIndex
        If int1 = -1 Then 'cancel or no selection
        Else 'delete an analyte replicate

            boolF = dgv.Columns.Item("Item").Frozen
            dgv.Columns.Item("Item").Frozen = False

            Dim col1 As New DataColumn
            Dim strAnal As String
            strAnal = frm.lbxAnalytes.SelectedItem
            'remove column
            Dim cols As DataColumnCollection
            cols = dtW.Columns
            cols.Remove(strAnal)
            dv = New DataView(dtW)
            Me.dgvCompanyAnalRef.DataSource = dv
            Me.dgvCompanyAnalRef.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
            Me.dgvCompanyAnalRef.AutoResizeColumns()
            dgv.Columns.Item("Item").Frozen = boolF

        End If

        str1 = AnalRefHook()
        If Len(str1) > 0 Then
            Select Case str1
                Case Is = "CRLWor_AnalRefStandard"
                    Call ComboBoxCRLAnalRefFill()
            End Select
        End If
        Call HideAnalRefRows()

        dgv.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
        Call SyncCols(Me.dgvWatsonAnalRef, Me.dgvCompanyAnalRef)


end1:
        On Error GoTo 0
        dt = Nothing
        dtW = Nothing
        frm.Dispose()


    End Sub

    Private Sub cmdCPAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCPAdd.Click
        Dim dv As system.data.dataview
        Dim strF As String
        Dim ct1 As Short
        Dim intMax As Short
        Dim Count1 As Short
        Dim int1 As Short
        Dim tbl As System.Data.DataTable
        Dim ct2 As Short
        Dim ct3 As Short
        Dim ct4 As Short
        Dim dtbl As System.Data.DataTable

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID

        maxID = 1
        maxID = GetMaxID("tblContributingPersonnel", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
        'Call PutMaxID("tblContributingPersonnel", maxID)

        'If boolGuWuOracle Then
        '    ta_tblMaxID.Fill(tblMaxID)
        'ElseIf boolGuWuAccess Then
        '    ta_tblMaxIDAcc.Fill(tblMaxID)
        'ElseIf boolGuWuSQLServer Then
        '    ta_tblMaxIDSQLServer.Fill(tblMaxID)
        'End If

        'strFMax = "charTable = 'tblContributingPersonnel'"
        ''"
        'tblMax = tblMaxID
        'rowsMax = tblMax.Select(strFMax)
        'maxID = rowsMax(0).Item("nummaxid")
        'maxID = maxID + 1
        'rowsMax(0).BeginEdit()
        'rowsMax(0).Item("nummaxid") = maxID
        'rowsMax(0).EndEdit()
        'If boolGuWuOracle Then
        '    Try
        '        ta_tblMaxID.Update(tblMaxID)
        '    Catch ex As DBConcurrencyException
        '        'ds2005.TBLMAXID.Merge('ds2005.TBLMAXID, True)
        '    End Try
        'ElseIf boolGuWuAccess Then
        '    Try
        '        ta_tblMaxIDAcc.Update(tblMaxID)
        '    Catch ex As DBConcurrencyException
        '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '    End Try
        'ElseIf boolGuWuSQLServer Then
        '    Try
        '        ta_tblMaxIDSQLServer.Update(tblMaxID)
        '    Catch ex As DBConcurrencyException
        '        'ds2005Acc.TBLMAXID.Merge('ds2005Acc.TBLMAXID, True)
        '    End Try
        'End If


        tbl = tblContributingPersonnel ' tblCP
        ct1 = tbl.Rows.Count

        dv = dgvContributingPersonnel.DataSource
        ct2 = dv.Count

        dv.AllowNew = True
        Dim dvr As DataRowView = dv.AddNew
        dvr.Item("id_tblStudies") = id_tblStudies
        'dvr.Item("id_tblContributingPersonnel") = 0
        'for some reason, bools start as null, even though the gridstyle = disallow null
        'dvr.Item("boolIncludeSigOnCompStatement") = 0

        'dvr.Item("boolIncludeSOP") = False
        dvr.Item("boolIncludeSOTP") = False
        'dvr.Item("boolIncludeSOCS") = False

        dvr.Item("id_tblContributingPersonnel") = maxID
        'find intOrder Max
        intMax = 0
        For Count1 = 0 To ct2 - 1
            int1 = dv(Count1).Item("intOrder")
            'int1 = dv(Count1).Item("intOrder")
            If int1 > intMax Then
                intMax = int1
            End If
        Next
        intMax = intMax + 1
        dvr.Item("intOrder") = intMax
        dvr.EndEdit()
        dv.AllowNew = False

        dgvContributingPersonnel.CurrentCell = dgvContributingPersonnel.Rows.Item(ct2).Cells("CHARCPPREFIX")

    End Sub


    Private Sub cmdCPDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCPDelete.Click
        Dim int1 As Short
        Dim dv As system.data.dataview
        Dim tbl As System.Data.DataTable
        Dim ct1 As Short
        Dim ct2 As Short
        Dim Count1 As Short
        Dim r As DataRow

        tbl = tblCP
        ct1 = tbl.Rows.Count
        ''debugWriteLine("Beginning...")
        'For Each r In tbl.Rows
        '    'debugWriteLine(r.RowState)
        'Next
        If dgvContributingPersonnel.CurrentRow Is Nothing Then
            int1 = -1
        Else
            int1 = dgvContributingPersonnel.CurrentRow.Index
        End If
        If int1 = -1 Then
            Exit Sub
        End If

        dv = dgvContributingPersonnel.DataSource
        dv.AllowDelete = True
        dv(int1).Delete()
        dv.AllowDelete = False

    End Sub

    Private Sub chkMethodValMultiple_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMethodValMultiple.CheckedChanged
        Dim bool
        Dim tbl As System.Data.DataTable


        bool = chkMethodValMultiple.Checked
        rbMethValAnalyte.Enabled = bool
        rbMethValMultiple.Enabled = bool

        If rbMethValAnalyte.Checked Then
            txtMethValMultiple.Enabled = Not (bool)
            cmdMethValExecute.Enabled = Not (bool)
        Else
            txtMethValMultiple.Enabled = bool
            cmdMethValExecute.Enabled = bool
        End If

        If bool Then
            If rbMethValAnalyte.Checked Then
                Call MethValMultipleAddColumns(1)
            Else
                Call MethValMultipleAddColumns(2)
            End If
        Else
            Dim Count1 As Short
            Dim intRows1 As Short
            Dim drow As DataRow
            Dim var1

            Call MethValMultipleAddColumns(0)
            'initialize tblMethValExistingGuWu
            tbl = tblMethValExistingGuWu
            tbl.Rows.Clear()
            For Count1 = 1 To 1
                drow = tbl.NewRow
                var1 = tbl.Columns.Item(Count1).ColumnName
                'var1 = dgvMethodValData.TableStyles(0).GridColumnStyles(Count1).HeaderText
                var1 = Me.dgvMethodValData.Columns(Count1).HeaderText
                drow("ColumnName") = var1
                'leave last two columns blank
                drow.EndEdit()
                tbl.Rows.Add(drow)
            Next

        End If

    End Sub


    Private Sub rbMethValMultiple_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMethValMultiple.CheckedChanged
        Exit Sub

        If rbMethValAnalyte.Checked Then
            txtMethValMultiple.Enabled = False
            cmdMethValExecute.Enabled = False
        Else
            txtMethValMultiple.Enabled = True
            cmdMethValExecute.Enabled = True
        End If
    End Sub

    Private Sub rbMethValAnalyte_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMethValAnalyte.CheckedChanged

        '20190208 LEE: This was deprecated long ago
        Exit Sub

        Dim bool
        Dim Count1 As Short
        Dim intCols As Short
        Dim dtbl As System.Data.DataTable
        Dim gc As GridColumnStylesCollection
        Dim dv As System.Data.DataView
        Dim ts1 As DataGridTableStyle

        bool = rbMethValAnalyte.Checked

        If rbMethValAnalyte.Checked Then
            txtMethValMultiple.Enabled = Not (bool)
            cmdMethValExecute.Enabled = Not (bool)
            MethValMultipleAddColumns(1)
        Else
            txtMethValMultiple.Enabled = bool
            cmdMethValExecute.Enabled = bool
            'delete all but two tbl columns and gridstyles
            dtbl = tblMethodValData
            'ts1 = dgvMethodValData.TableStyles(0)
            'gc = ts1.GridColumnStyles
            intCols = dtbl.Columns.Count
            For Count1 = intCols - 1 To 2 Step -1
                dtbl.Columns.Remove(dtbl.Columns.Item(Count1))
                'gc.RemoveAt(Count1)
            Next
            Me.dgvMethodValData.Columns(1).HeaderText = "Value"
            'gc(1).HeaderText = "Value"
            dv = dtbl.DefaultView
            'dgvMethodValData.TableStyles.Clear()
            'dgvMethodValData.TableStyles.Add(ts1)
            dgvMethodValData.DataSource = dv
        End If

        Call FillMethValExistingGuWu()


    End Sub


    Private Sub cmdMethValReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMethValReset.Click

        Call doMethValCancel()

    End Sub

    Private Sub cbxMethValExistingGuWu_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxMethValExistingGuWu.SelectedIndexChanged

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim str1 As String
        Dim str2 As String
        Dim tbl As System.Data.DataTable
        Dim drows() As DataRow
        Dim drows1() As DataRow
        Dim var1, var2, var3, var4
        Dim dv As system.data.dataview
        Dim int1 As Short
        Dim Count1 As Short
        Dim boolS As Boolean
        Dim dtbl As System.Data.DataTable
        Dim rowsR() As DataRow
        Dim strF As String

        Dim boolClear As Boolean = False

        If cmdEdit.Enabled Then
            Exit Sub
        End If

        If cmdEdit.Enabled = False And cmdSave.Enabled = False Then
            Exit Sub
        End If

        dtbl = tblMethodValidationData

        dgv = dgvMethValExistingGuWu
        'determine if a row has been selected
        If dgv.RowCount = 0 Then
            GoTo end1
        End If
        If dgv.CurrentRow Is Nothing Then
            MsgBox("Please select an item from the table.", MsgBoxStyle.Information, "Select an item...")
            GoTo end1
        End If
        intRow = dgv.CurrentRow.Index

        str1 = "charWatsonStudyName = '" & cbxMethValExistingGuWu.SelectedItem & "'"
        'tbl = tblStudies
        tbl = tblStudies
        drows = tbl.Select(str1)

        If drows.Length = 0 Then
            'probably because '[NONE]' was chosen
            boolClear = True
        End If

        var1 = cbxMethValExistingGuWu.SelectedItem
        If StrComp(var1, "[NONE]", CompareMethod.Text) = 0 Then
            var1 = DBNull.Value
            var2 = DBNull.Value
        Else
            var2 = drows(0).Item("id_tblStudies")
        End If

        'debug
        For Count1 = 0 To dgv.Columns.Count - 1
            var3 = dgv.Columns(Count1).Name
            var3 = var3
        Next

        'enter info
        dv = dgv.DataSource
        int1 = dv.Count
        boolS = False
        For Count1 = 0 To int1 - 1
            If dgv.Rows(Count1).Selected Then
                boolS = True
                dv.Item(Count1).BeginEdit()
                dv.Item(Count1).Item("WatsonStudy") = var1
                dv.Item(Count1).Item("ID_TBLSTUDIES") = var2
                'find archive path
                strF = "ID_TBLSTUDIES = " & NZ(var2, -98736) & " AND INTCOLUMNNUMBER = " & Count1 + 1
                rowsR = dtbl.Select(strF)
                If rowsR.Length = 0 Then
                    var3 = ""
                Else
                    var3 = NZ(rowsR(0).Item("CHARARCHIVEPATH"), "")
                End If
                dv.Item(Count1).Item("CHARARCHIVEPATH") = var3
                dv(Count1).EndEdit()
            End If
        Next

        If boolS Then 'continue
        Else
            str1 = "Remember to select one or more rows in order to fill Watson Study data."
            MsgBox(str1, MsgBoxStyle.Information, "No rows selected...")
            dgvMethValExistingGuWu.Select()
        End If

        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()

end1:

    End Sub

    Private Sub cmdCancelReportStatements_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancelReportStatements.Click

        Cursor.Current = Cursors.WaitCursor

        Call DoReportStatementsCancel()

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmdOpenReportStatements_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenReportStatements.Click

        Dim strM As String
        Dim boolA As Boolean = BOOLEDITWORDTEMPLATE
        If boolA Then
        Else
            strM = "User does not have permission to edit Report Templates."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end2
        End If

        Dim strPath As String
        Dim dgv As DataGridView
        Dim dgv1 As DataGridView
        Dim intRow As Short
        Dim intRow1 As Short
        Dim boolM As Boolean
        Dim id As Int64
        Dim strTitle As String
        Dim strSection As String

        Cursor.Current = Cursors.WaitCursor

        boolM = False
        dgv = Me.dgvReportStatementWord
        If dgv.Rows.Count = 0 Then
            boolM = True
            strM = "A Report Statement must be chosen."
            GoTo end1
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        dgv1 = Me.dgvReportStatements
        intRow1 = dgv1.CurrentRow.Index
        strSection = dgv1("CHARHEADINGTEXT", intRow1).Value

        strPath = dgv("CHARWORDSTATEMENT", intRow).Value
        id = dgv("ID_TBLWORDSTATEMENTS", intRow).Value
        strTitle = dgv("CHARTITLE", intRow).Value

        Call OpenfrmWord(strPath, id, strTitle, strSection, True)

end1:

        Cursor.Current = Cursors.Default

        If boolM Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        End If

end2:

    End Sub

    Sub OpenfrmWord(ByVal strPath As String, ByVal id As Int64, ByVal strTitle As String, ByVal strSection As String, boolEdit As Boolean)

        'This comes from Edit Templates

        Dim frm As New frmWordStatement
        Dim str1 As String
        Dim strLbl As String
        Dim strpathT As String
        Dim var1

        If Me.rbEntireReport.Checked Then

            strpathT = CreatexmlHome(Me.dgvReportStatementWord)

        Else
            strpathT = strPath

        End If

        'frm.boolSTB = boolSTB

        str1 = Me.cmdOpenReportStatements.Text
        If InStr(1, str1, "Templates", CompareMethod.Text) > 0 Then
            frm.boolReport = True
            frm.Text = " Template Editor"
        Else
            frm.boolReport = False
            frm.Text = " Template Statement Editor"
        End If

        'must always be true if coming from Report Template
        frm.boolReport = True

        frm.boolRO = False
        'frm.cmdShow.Visible = False
        frm.strPath = strPath
        frm.strReport = strpathT
        frm.id = id
        'frm.Text = " Report Statement Editor"
        frm.strSection = strSection
        'frm.CHARTITLE.Text = strTitle
        'frm.MdiParent = Me
        frm.boolEdit = boolEdit
        Call frm.PlaceControls()

        Call frm.FormLoad()

        Call frm.DoReadOnly()

        If Me.cmdEdit.Enabled Or id_tblStudies = 0 Then
            'frm.lblEdit.Visible = True
            'frm.cmdAddStatement.Enabled = False
            frm.cmdEditStatements.Enabled = False
            'frm.cmdFieldCode.Enabled = False
            'frm.cmdSave.Visible = False
            'If boolEdit Then
            '    frm.lblEdit.Visible = False
            'Else
            '    frm.lblEdit.Visible = True
            'End If

        Else
            'frm.lblEdit.Visible = False
        End If

        'Me.Visible = False

        Call frm.RefreshAFR()

        Call frm.DoReadOnly()
        'frm.ShowDialog(Me)

        frm.Show(Me)
        'Me.Visible = False


    End Sub

    Private Sub cmdRTHeaderConfigCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRTHeaderConfigCancel.Click

        Call DoCancelRTHConfig()

    End Sub

    Private Sub cmdInsertQAEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertQAEvent.Click

        Dim var1
        Dim int1 As Short
        Dim int2 As Short
        Dim dg As DataGrid
        Dim tbl As System.Data.DataTable
        Dim tbl1 As System.Data.DataTable
        Dim row1() As DataRow
        Dim dv As system.data.dataview
        Dim row As DataRow
        Dim str1 As String
        Dim rowIndex As Short
        Dim numID As Int64
        Dim intOrder As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim ct1 As Short
        Dim dv1 As system.data.dataview
        Dim frm As New frmAddQARow

        dg = dgQATable
        tbl = tblQATableTemp
        dv = dg.DataSource
        ct1 = dv.Count ' tbl.Rows.Count

        int1 = dg.CurrentRowIndex
        If int1 = -1 Then
            'Exit Sub
        Else
            'record critical phase and other parameters
            intOrder = dv(int1).Item("intOrder")
            'before adding new row, re-order the remaining rows
            For Count1 = intOrder To ct1 - 1
                int2 = dv(Count1).Item("intOrder")
                dv(Count1).Item("intOrder") = int2 + 1
                dv(Count1).EndEdit()
            Next
        End If

        'call form
        'populate frmlbx
        tbl1 = tblReportTableHeaderConfig
        str1 = "id_tblStudies = " & id_tblStudies & " AND id_tblConfigReportTables = 2000 AND boolInclude = -1"
        row1 = tbl1.Select(str1, "intOrder ASC")
        ct1 = row1.Length
        frm.lbxQACriticalPhase.Items.Clear()
        frm.lbxID.Items.Clear()
        For Count1 = 0 To ct1 - 1
            var1 = row1(Count1).Item("charUserLabel")
            frm.lbxQACriticalPhase.Items.Add(var1)
            numID = row1(Count1).Item("id_tblReportTableHeaderConfig")
            frm.lbxID.Items.Add(numID)
        Next
        frm.ShowDialog(Me)
        If frm.boolCancel Or frm.numID = 0 Then 'cancel
            intOrder = intOrder - 1
        Else

            'row = tbl.NewRow
            ''row.Item("charUserLabel") = str1
            'row.Item("id_tblStudies") = id_tblStudies
            'row.Item("id_tblReportTableHeaderConfig") = frm.numID
            'row.Item("charUserLabel") = frm.char1
            'row.Item("intOrder") = intOrder + 1
            'row.Item("id_tblReports") = 0
            'row.Item("id_tblQATables") = 0
            'tbl.Rows.Add(row)

            'dv = tbl.DefaultView
            'str1 = "intOrder ASC"
            'dv.Sort = str1

            dv.AllowNew = True
            dv.AllowEdit = True
            Dim rowView As DataRowView = dv.AddNew

            ' Change values in the DataRow.
            ID_QATEMPID = ID_QATEMPID + 1
            rowView.Item("ID_QATEMPID") = ID_QATEMPID
            rowView.Item("id_tblStudies") = id_tblStudies
            rowView.Item("id_tblReportTableHeaderConfig") = frm.numID
            rowView.Item("charUserLabel") = frm.char1
            rowView.Item("intOrder") = intOrder + 1
            rowView.Item("id_tblReports") = 0
            rowView.Item("id_tblQATables") = 0
            rowView.EndEdit()

            dv.Sort = "INTORDER ASC"
            dv.AllowNew = False
            dv.AllowEdit = True
            dv.AllowDelete = False
            'dg.DataSource = dv
            dg.Refresh()

        End If
        'select inserted row
        dg.CurrentRowIndex = intOrder

    End Sub

    Private Sub cmdDeleteQAEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteQAEvent.Click

        Dim var1
        Dim int1 As Short
        Dim int2 As Short
        Dim dg As DataGrid
        Dim tbl As System.Data.DataTable
        Dim dv As system.data.dataview
        Dim row As DataRow
        Dim str1 As String
        Dim rowIndex As Short
        Dim numID As Int64
        Dim intOrder As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim ct1 As Short
        Dim dv1 As system.data.dataview
        Dim intRows As Short
        Dim arr1(1) As Short


        dg = dgQATable
        tbl = tblQATableTemp
        dv = dg.DataSource
        ct1 = dv.Count ' tbl.Rows.Count

        int1 = dg.CurrentRowIndex
        If int1 = -1 Then
            Exit Sub
        End If

        intRows = dv.Count
        ReDim arr1(intRows)

        'record selected rows
        Count2 = 0
        For Count1 = 0 To intRows - 1
            If dg.IsSelected(Count1) Then
                Count2 = Count2 + 1
                arr1(Count2) = Count1
            End If
        Next

        'delete selected rows
        dv.AllowDelete = True
        For Count1 = Count2 To 1 Step -1
            'tbl.Rows.item(arr1(Count1)).Delete()
            dv(arr1(Count1)).Delete()
        Next
        dv.AllowDelete = False

        intRows = dv.Count

        're-order the rows
        'For Count1 = intOrder - 1 To ct1 - 2
        For Count1 = 0 To intRows - 1
            int2 = dv(Count1).Item("intOrder")
            dv(Count1).BeginEdit()
            dv(Count1).Item("intOrder") = Count1 + 1 'int2 - 1
            dv(Count1).EndEdit()
        Next
        dv.Sort = "INTORDER ASC"

        'dv = tbl.DefaultView
        'str1 = "intOrder ASC"
        'dv.Sort = str1
        'dv.AllowNew = False
        'dv.AllowEdit = True
        'dv.AllowDelete = False
        'dg.DataSource = dv
        'dg.Refresh()

    End Sub

    Private Sub cmdQACancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdQACancel.Click
        Call DoCancelQATable()
    End Sub



    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        boolRefresh = True
        Call ActivateStudyChange()
        boolRefresh = False
    End Sub

    Private Sub cmdCopyRepAnalyte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopyRepAnalyte.Click
        Dim frm As New frmCopyReplicate
        Dim Count1 As Short
        Dim Count2 As Short
        Dim ct1 As Short
        Dim ct2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dt As System.Data.DataTable
        Dim dtC As System.Data.DataTable
        Dim dv As system.data.dataview
        Dim rw As DataRow
        Dim col As DataColumn
        Dim strName As String
        Dim var1
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim boolC As Boolean
        Dim strAnalTo As String
        Dim strAnalFrom As String

        'Check for datatable existance
        'On Error Resume Next
        'strName = CType(dgCompanyAnalRef.DataSource, DataTable).TableName.ToString

        If ctAnalytes = 0 Then
            MsgBox("There are no analytes listed in the Analytical Reference Standard Table.", MsgBoxStyle.Information, "No analytes...")
            Err.Clear()
            On Error GoTo 0
            GoTo end1
        End If

        'get analyte info from frm.dgWatsonAnalRef
        'dv = dgCompanyAnalRef.DataSource
        dv = Me.dgvCompanyAnalRef.DataSource

        dtC = tblCompanyAnalRefTable
        'dtC = dgWatsonAnalRef.DataSource
        ct1 = dtC.Columns.Count
        'start with column 1
        ct2 = FindRowDVByCol("Is Replicate?", dv, "Item")
        For Count1 = 0 To ct1 - 1
            str1 = dtC.Columns.Item(Count1).ColumnName
            If StrComp(str1, "ID_TBLDATATABLEROWTITLES", CompareMethod.Text) = 0 Then 'ignore
            ElseIf StrComp(str1, "BOOLINCLUDE", CompareMethod.Text) = 0 Then 'ignore
            ElseIf StrComp(str1, "Item", CompareMethod.Text) = 0 Then 'ignore
            Else
                'add to lbx
                frm.lbxAnalytesCopiedTo.Items.Add(str1)
                frm.lbxAnalytesCopiedFrom.Items.Add(str1)
            End If
        Next
        'frm.lblAddRep.Text = "Choose analyte for which a new lot will be added:"
        'frm.Text = "Add Replicate"
        frm.ShowDialog(Me) 'show as modal
        If frm.boolStop Then 'stop
            GoTo end1
        Else
        End If
        int1 = frm.lbxAnalytesCopiedFrom.SelectedIndex
        int3 = frm.lbxAnalytesCopiedTo.SelectedIndex

        Dim strI1 As String
        Dim strI2 As String
        strI1 = frm.lbxAnalytesCopiedFrom.Text
        strI2 = frm.lbxAnalytesCopiedTo.Text

        If int1 = -1 Or int3 = -1 Then 'cancel or no selection
        ElseIf int1 = int3 Then 'not allowed
            strAnalFrom = frm.lbxAnalytesCopiedFrom.SelectedItem
            strAnalTo = frm.lbxAnalytesCopiedTo.SelectedItem
            MsgBox("You've chosen to copy data from " & strAnalFrom & " to " & strAnalTo & ". There is no need to copy data to itself.", MsgBoxStyle.Information, "No need...")

        Else 'copy data from an analyte
            strAnalFrom = frm.lbxAnalytesCopiedFrom.SelectedItem
            strAnalTo = frm.lbxAnalytesCopiedTo.SelectedItem

            'begin copy row data
            dv.AllowEdit = True
            int2 = dv.Count
            For Count1 = 0 To int2 - 2 'do not copy some stuff
                str1 = dv.Item(Count1).Item("Item")
                str1 = dv(Count1).Item("Item")
                If StrComp(str1, "ID", CompareMethod.Text) = 0 Then 'do not copy
                ElseIf StrComp(str1, "Analyte Parent", CompareMethod.Text) = 0 Then 'do not copy
                ElseIf StrComp(str1, "Is Coadministered Cmpd?", CompareMethod.Text) = 0 Then 'do not copy
                ElseIf StrComp(str1, "Analyte Name", CompareMethod.Text) = 0 Then 'do not copy
                ElseIf StrComp(str1, "Is Replicate?", CompareMethod.Text) = 0 Then 'do not copy
                ElseIf StrComp(str1, "Is Configured in Watson?", CompareMethod.Text) = 0 Then 'do not copy
                ElseIf StrComp(str1, "Is Internal Standard?", CompareMethod.Text) = 0 Then 'do not copy
                    'var1 = "No"
                    'dv(Count1).BeginEdit()
                    'dv(Count1).Item(int3 + 1) = var1
                    'dv(Count1).EndEdit()
                Else
                    var1 = dv.Item(Count1).Item(strI1)
                    dv(Count1).BeginEdit()
                    dv(Count1).Item(strI2) = var1
                    dv(Count1).EndEdit()
                End If
                'dv.Item(Count1).Item(int3 + 1) = var1

            Next
            'dv.AllowEdit = False
            'dgCompanyAnalRef.DataSource = dv
            'dgCompanyAnalRef.Refresh()

            'me.dgvCompanyAnalRef.DataSource = dv

            str1 = AnalRefHook()
            If Len(str1) > 0 Then
                Select Case str1
                    Case Is = "CRLWor_AnalRefStandard"
                        Call ComboBoxCRLAnalRefFill()
                End Select
            End If

            Call HideAnalRefRows()

        End If
end1:
        On Error GoTo 0
        dt = Nothing
        dtC = Nothing
    End Sub

    Private Sub dgvReportTableConfiguration_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportTableConfiguration.CellValidated

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim str1 As String

        dgv = dgvReportTableConfiguration
        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        If StrComp(str1, "INTORDER", CompareMethod.Text) = 0 Then
            dgv.AutoResizeRows()
        End If
        If boolRTCEnter Then
            dgv.CurrentCell.Value = varR
            dgv.Refresh()
        End If


    End Sub

    Private Sub dgvReportTableConfiguration_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvReportTableConfiguration.CellValidating

        Dim boolErr As Boolean
        Dim strErr As String
        Dim dgv As DataGridView
        Dim value As System.Drawing.Point
        Dim intRow ' As Short
        Dim intCol ' As Short
        Dim int1 As Short
        Dim var1, var2, var3, var4
        Dim boolCC As Boolean
        Dim str1 As String
        Dim boolG As Boolean
        Dim boolG1 As Boolean
        Dim boolG2 As Boolean
        Dim Count1 As Short
        Dim Count2 As Short
        Dim str2 As String
        Dim boolFCID As Boolean = False

        If boolLoad Then
            Exit Sub
        End If
        If boolFromRTC Then
            Exit Sub
        End If
        If cmdEdit.Enabled Then
            Exit Sub
        End If

        dgv = dgvReportTableConfiguration
        value = dgv.CurrentCellAddress
        If value.X = -1 Or value.Y = -1 Then
            Exit Sub
        End If

        intRow = e.RowIndex
        intCol = e.ColumnIndex

        var1 = e.FormattedValue
        newCurrentCellRTC = var1
        int1 = dgv.Rows.Count
        If int1 < 1 Then
            Exit Sub
        End If

        boolRTCEnter = False
        boolOKtoVal = False
        boolG = False
        boolG1 = False
        boolFCID = False

        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        If StrComp(str1, "CHARPAGEORIENTATION", CompareMethod.Text) = 0 Then
            boolOKtoVal = True
            boolG1 = True
        ElseIf StrComp(str1, "PeriodTemp", CompareMethod.Text) = 0 Then
            'boolG = True
        ElseIf StrComp(str1, "CHARFCID", CompareMethod.Text) = 0 Then
            boolFCID = True
        Else
            GoTo end1
        End If


        If boolG1 Then
            'check to see that entry is only P or L
            boolErr = False
            If Len(var1) = 0 Or Len(var1) > 1 Then
                boolErr = True
                strErr = "This entry must be 'P' or 'L'"
                boolOKtoVal = False
            Else
                var2 = Asc(var1)
                If var2 = 76 Or var2 = 80 Then 'acceptable
                    boolErr = False
                    boolOKtoVal = True
                ElseIf var2 = 108 Then 'make capital L
                    varR = "L"
                    boolRTCEnter = True
                    'dgv.Item(oldCurrentColRTC, oldCurrentRowRTC).Value = var3
                    'dgv.CurrentCell.Value = varR
                    boolErr = False
                    boolOKtoVal = True
                ElseIf var2 = 112 Then 'make capital P
                    varR = "P"
                    boolRTCEnter = True
                    'dgv.Item(oldCurrentColRTC, oldCurrentRowRTC).Value = var3
                    'dgv.CurrentCell.Value = varR
                    boolErr = False
                    boolOKtoVal = True
                Else
                    boolErr = True
                    strErr = "This entry must be 'P' or 'L'"
                    boolOKtoVal = False
                End If
            End If
            If boolOKtoVal Then
                e.Cancel = False
            Else
                e.Cancel = True
                MsgBox(strErr, MsgBoxStyle.Information, "Entry must be 'P' or 'L'...")
                GoTo end1
            End If
        End If

        If boolFCID Then

            Dim varVal
            Dim strM As String
            Dim dv As system.data.dataview
            Dim varV

            varVal = e.FormattedValue

            If Len(varVal) = 0 Then
                boolErr = False
                GoTo err1
            End If

            'If IsNumeric(varVal) Then
            '    strM = "Entry cannot be pure numeric. Entry must be text or mixture of text and numbers."
            '    boolErr = True
            '    GoTo err1
            'End If

            If HasSpecialCharacters(CStr(varVal)) Then
                strM = ""
                boolErr = True
                GoTo err1
            End If

            'first character cannot be numeric
            str1 = Mid(varVal, 1, 1)
            If IsNumeric(str1) Then
                strM = "Table FC ID's cannot start with a numeric character." & ChrW(10) & ChrW(10)
                'strM = strM & dv(Count1).Item("CHARHEADINGTEXT") & ChrW(10) & ChrW(10)
                'strM = strM & "Though this is allowed, the Field Code for this entry will refer only to the first table labeled with this FC ID."
                boolErr = True
                GoTo err1
            End If

            'must be unique in table
            dv = dgv.DataSource
            For Count1 = 0 To dv.Count - 1
                If Count1 = intRow Then 'ignore
                Else
                    varV = NZ(dv(Count1).Item("CHARFCID"), "")
                    If StrComp(CStr(varVal), CStr(varV), CompareMethod.Text) = 0 Then
                        strM = "Please note that this FC ID entry already exists in this table:" & ChrW(10) & ChrW(10)
                        strM = strM & dv(Count1).Item("CHARHEADINGTEXT") & ChrW(10) & ChrW(10)
                        strM = strM & "Though this is allowed, the Field Code for this entry will refer only to the first table labeled with this FC ID."
                        'boolErr = True
                        MsgBox(strM, MsgBoxStyle.Information, "Note...")
                        GoTo err1
                    End If
                End If

            Next

err1:

            If boolErr Then
                e.Cancel = True
                If Len(strM) = 0 Then
                Else
                    MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
                End If

            End If

        End If

        Dim strCol As String
        Dim strTempInfo
        Dim strM1 As String
        Dim strM2 As String
        Dim strM3 As String
        Dim strM4 As String


end1:

        'ensure table can be displayed
        Dim id As Long
        id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
        If StrComp(dgv.Columns(intCol).Name, "BOOLREQUIRESSAMPLEASSIGNMENT", CompareMethod.Text) = 0 Then
            'If id = 1 Or id = 2 Then
            If id = 1 Then
                '20190304 LEE: 2 - Regr Constant table - now allows sample assignment
                var1 = dgv("boolRequiresSampleAssignment", intRow).Value
                If var1 = -1 Then
                    e.Cancel = True
                    str1 = "The table '" & dgv("CHARHEADINGTEXT", intRow).Value & "' cannot have samples assigned to it."
                    MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
                    'dgv("boolRequiresSampleAssignment", intRow).Value = 0
                End If
            End If
        End If



    End Sub

    Function HasSpecialCharacters(ByVal strVal As String) As Boolean

        HasSpecialCharacters = False
        Dim intL As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim varC
        Dim boolGo1 As Boolean = False
        Dim boolGo2 As Boolean = False
        Dim boolGo3 As Boolean = False
        Dim boolGo4 As Boolean = False

        intL = Len(strVal)
        For Count1 = 1 To intL
            str1 = Mid(strVal, Count1, 1)
            varC = AscW(str1)
            boolGo1 = False
            boolGo2 = False
            boolGo3 = False
            boolGo4 = False
            If (varC > 64 And varC < 91) Or (varC > 60 And varC < 123) Then 'letters OK
                boolGo1 = True
            End If

            If (varC > 47 And varC < 58) Then 'numbers OK
                boolGo2 = True
            End If

            If varC = 92 Or varC = 45 Or varC = 32 Then '_,-,space
                boolGo3 = True
            End If

            If boolGo1 Or boolGo2 Or boolGo3 Then
                HasSpecialCharacters = False
            Else
                HasSpecialCharacters = True
                Exit For
            End If

        Next

        If HasSpecialCharacters Then
            Dim strM As String
            strM = "Special characters are not allowed." & ChrW(10) & ChrW(10) & "' " & str1 & " ' is considered a special character."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        End If


    End Function

    Private Sub dgvReportTableConfiguration_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvReportTableConfiguration.DataError
        Dim var1
        Dim int1 As Short

        int1 = e.ColumnIndex
        If int1 = 0 Then
        Else
            var1 = e.Exception.Message
            MsgBox("Data Error: " & CStr(var1))
        End If
    End Sub

    Private Sub dgvReportStatements_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportStatements.CellContentClick

        Dim dgv As DataGridView
        Dim bool As Boolean
        Dim str1 As String

        str1 = NZ(cbxStudy.Text, "")
        If Len(str1) = 0 Then
            Exit Sub
        End If
        If cmdEdit.Enabled Then
            Exit Sub
        End If
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        dgv = dgvReportStatements
        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        bool = False

        If StrComp(str1, "boolI", CompareMethod.Text) = 0 Then
            bool = True
        ElseIf StrComp(str1, "boolUStatements", CompareMethod.Text) = 0 Then
            bool = True
        ElseIf StrComp(str1, "boolGW", CompareMethod.Text) = 0 Then
            bool = True
        ElseIf StrComp(str1, "boolPB", CompareMethod.Text) = 0 Then
            bool = True
        End If
        Dim var1

        var1 = dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value
        If bool Then
            'first commit
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            Call dgvReportStatementsCellContentClick("CellContentClick")
        End If

    End Sub


    Private Sub dgvReportStatements_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportStatements.CellValidated

        If boolFormLoad Then
            Exit Sub
        End If
        If cmdEdit.Enabled Then
            Exit Sub
        End If
        If boolStopRBS Then
            Exit Sub
        End If
        If cbxStudy.SelectedIndex = -1 Then
            Exit Sub
        End If
        If boolHold Then
            Exit Sub
        End If

        Dim intRow As Short
        Dim intCol As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim intA As Short
        Dim intB As Short
        Dim intS As Short
        Dim boolGo As Boolean
        Dim dgv As DataGridView
        Dim var1, var2

        'intRow = e.RowIndex
        intRow = dgvReportStatements.CurrentRow.Index
        intCol = dgvReportStatements.CurrentCell.ColumnIndex



        Dim int1 As Short

        'intCol = e.ColumnIndex
        boolGo = False
        intA = -1
        intB = -1
        int1 = dgvReportStatements.Columns.Count
        For Count1 = 0 To int1 - 1
            str1 = dgvReportStatements.Columns.Item(Count1).Name
            If StrComp(str1, "boolGuWu", CompareMethod.Text) = 0 Then
                intB = Count1
            ElseIf StrComp(str1, "boolUseStatements", CompareMethod.Text) = 0 Then
                intA = Count1
            ElseIf StrComp(str1, "charStatement", CompareMethod.Text) = 0 Then
                intS = Count1
            End If
        Next
        str1 = dgvReportStatements.Columns.Item(intCol).Name
        If StrComp(str1, "boolUseStatements", CompareMethod.Text) = 0 Then
            boolGo = True
            str2 = "boolGuWu"
        ElseIf StrComp(str1, "boolGuWu", CompareMethod.Text) = 0 Then
            boolGo = True
            str2 = "boolUseStatements"
        End If

        dgv = dgvReportStatements
        var1 = dgv.Item(intB, intRow).Value 'StudyDoc
        var2 = dgv.Item(intA, intRow).Value 'report
        If var1 = var2 Then 're-assess
            If intCol = intB Then 'StudyDoc then
                dgv.Item(intB, intRow).Value = Not (var1)
                dgv.Item(intA, intRow).Value = var1
            ElseIf intCol = intA Then
                dgv.Item(intA, intRow).Value = Not (var2)
                dgv.Item(intB, intRow).Value = var2
            End If
            dgvReportStatements.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
        var1 = dgv.Item(intB, intRow).Value 'guwu
        var2 = dgv.Item(intA, intRow).Value 'report
        'MsgBox("guwu: " & var1.ToString & ", report: " & var2.ToString)


    End Sub


    Private Sub dgvReportStatements_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvReportStatements.CellValidating
        Dim var1
        Dim boolErr As Boolean
        Dim dgv As DataGridView
        Dim str1 As String
        Dim boolGo As Boolean
        Dim boolGo1 As Boolean
        Dim boolGo2 As Boolean
        Dim intCol As Short
        Dim intRow As Short

        If boolFormLoad Then
            Exit Sub
        End If
        If cmdEdit.Enabled Then
            Exit Sub
        End If
        If boolStopRBS Then
            Exit Sub
        End If
        If boolHold Then
            Exit Sub
        End If

        If cbxStudy.SelectedIndex = -1 Then
            Exit Sub
        End If

        dgv = dgvReportStatements
        intRow = dgv.CurrentRow.Index
        intCol = dgv.CurrentCell.ColumnIndex
        str1 = dgv.Columns.Item(intCol).Name
        boolGo = False
        boolGo1 = False
        boolGo2 = False
        If StrComp(str1, "intOrder", CompareMethod.Text) = 0 Then
            boolGo = True
        ElseIf StrComp(str1, "NUMHEADINGLEVEL", CompareMethod.Text) = 0 Then
            boolGo1 = True
        End If

        boolErr = False
        If boolGo Then
            var1 = e.FormattedValue
            'number must be numeric
            If IsNumeric(var1) Then
                'number must be integer > 0
                If IsInt(var1) Then
                    'must be > 0
                    If CInt(var1) > 0 Then
                    Else
                        boolErr = True
                    End If
                Else
                    boolErr = True
                End If
            Else
                boolErr = True
            End If
        End If
        If boolErr Then
            MsgBox("Entry must be integer > 0.", MsgBoxStyle.Information, "Integer > 0...")
            e.Cancel = True
            Exit Sub
        End If

        boolErr = False
        If boolGo1 Then
            var1 = e.FormattedValue
            'number must be numeric
            If IsNumeric(var1) Then
                'number must be integer >= 0 and < 4
                If IsInt(var1) Then
                    'must be > 0
                    If CInt(var1) >= 0 And CInt(var1) <= 6 Then
                    Else
                        boolErr = True
                    End If
                Else
                    boolErr = True
                End If
            Else
                boolErr = True
            End If
        End If
        If boolErr Then
            MsgBox("Entry must be integer >= 0 and <= 6.", MsgBoxStyle.Information, "Integer >= 0 and <= 6...")
            e.Cancel = True
        End If


    End Sub


    Private Sub dgvReportStatements_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvReportStatements.DataError

        Dim dgv As DataGridView
        Dim introw As Short
        Dim intcol As Short
        Dim str1 As String
        Dim boolGo As Boolean
        Dim var1
        Dim boolErr As Boolean

        dgv = dgvReportStatements
        introw = e.RowIndex
        intcol = e.ColumnIndex
        str1 = dgv.Columns.Item(intcol).Name
        boolGo = False
        If StrComp(str1, "intOrder", CompareMethod.Text) = 0 Then
            MsgBox("Entry must be integer > 0.", MsgBoxStyle.Information, "Integer > 0...")
        End If

    End Sub


    Private Sub lblQAHyperlink_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblQAHyperlink.LinkClicked
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim boolGo As Boolean

        str1 = "Configure Column Headings" ' "Report Table Header Configuration"
        int1 = lbxTab1.Items.Count
        boolGo = False
        For Count1 = 0 To int1 - 1
            str2 = lbxTab1.Items(Count1).ToString
            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                lbxTab1.SelectedIndex = Count1
                boolGo = True
                Exit For
            End If
        Next

        'now find QA Rows
        Dim int2 As Short

        If boolGo Then
            Dim dv As system.data.dataview
            dv = dgvReportTables.DataSource
            int1 = FindRowDVByCol("QA Events Table Columns (Events/Management)", dv, "charTableName")
            'find first visible cell
            int2 = 0
            For Count1 = 0 To Me.dgvReportTables.ColumnCount - 1
                If Me.dgvReportTables.Columns(Count1).Visible Then
                    int2 = Count1
                    Exit For
                End If
            Next
            Me.dgvReportTables.CurrentCell = Me.dgvReportTables.Rows(int1).Cells(int2)
            Me.dgvReportTables.Rows(int1).Selected = True

        End If
    End Sub

    Private Sub cmdInsertSRec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdInsertSRec.Click
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim int2 As Short
        Dim ct1 As Short
        Dim dv As system.data.dataview
        Dim strS As String

        dgv = dgvSampleReceipt
        dv = dgv.DataSource
        ct1 = dv.Count ' tbl.Rows.Count

        If dgv.SelectedRows.Count = 0 And ct1 = 0 Then
            int1 = 0
        ElseIf dgv.SelectedRows.Count = 0 And ct1 > 0 Then
            int1 = ct1
        Else
            int1 = dgv.CurrentRow.Index + 1
        End If

        Call GenericDGVRowInsert(dgvSampleReceipt, tblSampleReceipt, "tblSampleReceipt", "id_tblSampleReceipt")

        'select new row
        dgv.CurrentCell = dgv("dtShipmentReceived", int1)
        'do some default values
        dgv("boolUse", int1).Value = -1
        dgv("boolU", int1).Value = True
        If chkUseWatsonSampleNumber.CheckState = CheckState.Checked Then
            dgv("boolUseWatson", int1).Value = -1 'True
        Else
            dgv("boolUseWatson", int1).Value = 0 'False
        End If
        If chkManualSampleNumber.CheckState = CheckState.Checked Then
            dgv("boolUseManual", int1).Value = -1 'True
        Else
            dgv("boolUseManual", int1).Value = 0 'False
        End If
        dgv.Update()
        dgv.AutoResizeColumns()



        Call ReorderSRec() 'funny that this has to be called


    End Sub

    Private Sub cmdDeletSRec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeletSRec.Click
        Dim int1 As Short
        Dim ct1 As Short
        Dim dgv As DataGridView

        dgv = dgvSampleReceipt

        'int1 = dgv.CurrentRow.Index

        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        int1 = dgv.CurrentRow.Index

        Call GenericDGVRowDelete(dgvSampleReceipt)

        ct1 = dgv.Rows.Count
        'select row before deleted row
        If ct1 = 0 Then
        ElseIf int1 = 0 Then
            dgv.CurrentCell = dgv("dtShipmentReceived", int1)
            dgv.Rows.Item(int1).Selected = True
        ElseIf int1 > ct1 - 1 Then
            dgv.CurrentCell = dgv("dtShipmentReceived", ct1 - 1)
            dgv.Rows.Item(ct1 - 1).Selected = True
        Else
            dgv.CurrentCell = dgv("dtShipmentReceived", int1 - 1)
            dgv.Rows.Item(int1 - 1).Selected = True
        End If

        Call CalcSampleCount()

        dgv.AutoResizeColumns()


    End Sub

    Private Sub cmdAdministration_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdministration.Click

        Dim strM As String
        Dim boolA As Boolean = BOOLADMINISTRATION
        If boolA Then
        Else
            strM = "User does not have permissions to access the 'Report Writer Administration' window."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        Dim frm As New frmAdministration

        frm.frmName = Me.Name

        frm.lblGlobalParameters.Text = "Report Writer Global Parameters"

        frm.cmdExit.Text = "G&o Back"

        'Call frm.FormLoad()

        'now disable cbxModule
        frm.cbxModules.Visible = False
        frm.lblcbxModules.Visible = False

        frm.ShowDialog()

        frm.Dispose()

end1:

    End Sub


    Private Sub tab1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tab1.SelectedIndexChanged
        Dim int1 As Short
        Dim int2 As Short

        Cursor.Current = Cursors.WaitCursor

        int1 = Me.tab1.SelectedIndex
        int2 = Me.lbxTab1.Items.Count
        'select appropriate lbxTab1 item
        If int1 > int2 - 1 Then
        Else
            Me.lbxTab1.SelectedIndex = int1
        End If

        Cursor.Current = Cursors.WaitCursor

        Call SelectedRefresh()

        Cursor.Current = Cursors.Default

    End Sub


    Private Sub cmdSRecCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSRecCancel.Click

        Call DoCancelSampleReceipt()

    End Sub

    Private Sub dgvSampleReceipt_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSampleReceipt.CellClick

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled Then
        Else
            Exit Sub
        End If

        Dim dgv As DataGridView

        dgv = Me.dgvSampleReceipt

        Dim intRow As Short
        Dim intCol As Short
        Dim strName As String
        Dim var1

        intRow = e.RowIndex
        intCol = e.ColumnIndex
        'strName = dgv.Columns(intCol).Name

        If intCol < 0 Then
            Exit Sub
        End If

        strName = dgv.Columns(intCol).Name

        Dim locX, locY

        If InStr(1, strName, "dt", CompareMethod.Text) > 0 Then 'show calendar

            var1 = dgv(intCol, intRow).Value

            'boolFromTab = True
            'locX = Me.tab1.Left + dgv.Left + dgv.RowHeadersWidth + (dgv.Columns(intCol).Width * 2)
            Dim ld As Single
            Dim Count1 As Short
            ld = 0
            For Count1 = 0 To intCol
                If dgv.Columns(Count1).Visible Then
                    ld = ld + dgv.Columns(intCol).Width
                End If
            Next
            locX = Me.tab1.Left + dgv.Left + dgv.RowHeadersWidth + ld
            locY = Me.tab1.Top + dgv.Location.Y + (dgv.Rows(intRow).Height * intRow) + dgv.ColumnHeadersHeight

            gCalGrid = dgv

            Dim dt As Date
            If IsDate(var1) Then
                dt = CDate(var1)
            Else
                dt = Now
            End If

            Call MakeCalVis(locX, locY, dt, True)

        Else

            Call MakeCalVis(0, 0, Now, False)

        End If

    End Sub

    Private Sub dgvSampleReceipt_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSampleReceipt.CellContentClick

        If boolFormLoad Then
            Exit Sub
        End If
        If cmdEdit.Enabled Then
            Exit Sub
        End If
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        Dim int1 As Short
        int1 = dgvSampleReceipt.CurrentCell.ColumnIndex
        If StrComp(dgvSampleReceipt.Columns.Item(int1).Name, "boolUse", CompareMethod.Text) = 0 Then
            dgvSampleReceipt.CommitEdit(DataGridViewDataErrorContexts.Commit)
            If chkManualSampleNumber.CheckState = CheckState.Checked Then
            Else
                Call CalcSampleCount()
            End If
        End If

    End Sub



    Private Sub dgvSampleReceipt_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSampleReceipt.CellEndEdit
        If chkManualSampleNumber.CheckState = CheckState.Checked Then
        Else
            Call CalcSampleCount()
        End If
    End Sub

    Private Sub dgvSampleReceipt_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSampleReceipt.CurrentCellDirtyStateChanged
        Dim dgv As DataGridView
        Dim dv As system.data.dataview
        Dim str1 As String
        Dim intRow As Short
        Dim intCol As Short
        Dim bool As Boolean

        dgv = dgvSampleReceipt
        intCol = dgv.CurrentCell.ColumnIndex
        intRow = dgv.CurrentRow.Index
        str1 = dgv.Columns.Item(intCol).Name
        dv = dgv.DataSource
        If StrComp(str1, "boolU", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            bool = dgv.Rows.Item(intRow).Cells(intCol).Value
            dv(intRow).BeginEdit()
            If bool Then
                dv(intRow).Item("boolUse") = -1
            Else
                dv(intRow).Item("boolUse") = 0
            End If
            dv(intRow).EndEdit()
        End If
    End Sub

    Private Sub dgvSampleReceipt_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvSampleReceipt.DataError
        Dim var1

        var1 = e.Exception.Message

        MsgBox("Data Error: " & CStr(var1))



    End Sub

    Private Sub chkManualSampleNumber_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkManualSampleNumber.CheckedChanged
        If chkManualSampleNumber.CheckState = CheckState.Checked Then
            txtSRecTotalReport.ReadOnly = False
            txtSRecTotalReport.Text = 0
            'chkUseWatsonSampleNumber.CheckState = CheckState.Unchecked
        Else
            txtSRecTotalReport.ReadOnly = True
            txtSRecTotalReport.Text = 0
            Call CalcSampleCount()
        End If
        Call CorrectSampleReceipt(False, True)

    End Sub


    Private Sub txtSRecTotalReport_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSRecTotalReport.Validating
        'entry must be integer
        Dim var1
        Dim str1 As String

        var1 = txtSRecTotalReport.Text
        If IsInt(var1) Then
        Else
            str1 = "Entry must be integer >= 0."
            MsgBox(str1, MsgBoxStyle.Information, "Must be integer >= 0...")
            txtSRecTotalReport.Text = 0
            e.Cancel = True

        End If

    End Sub


    Private Sub chkUseWatsonSampleNumber_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUseWatsonSampleNumber.CheckedChanged

        Call CorrectSampleReceipt(True, False)

    End Sub

    Sub ClearSelection(dgv As DataGridView)

        dgv.ClearSelection()

    End Sub


    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        MsgBox(Me.lbl2.Top)

    End Sub

    Sub MyHandler(ByVal sender As Object, ByVal args As UnhandledExceptionEventArgs)

        Dim e As Exception = DirectCast(args.ExceptionObject, Exception)
        MsgBox("MyHandler caught: " & e.Message)

    End Sub

    Private Sub cmdMethValExecute_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMethValExecute.Click

        Dim intR As Short
        Dim strM As String
        Dim boolAll As Boolean

        strM = "Do you wish to update the selected analyte (Yes)"
        strM = strM & ChrW(10) & ChrW(10)
        strM = strM & "Or"
        strM = strM & ChrW(10) & ChrW(10)
        strM = strM & "all the analytes (No)?"

        intR = MsgBox(strM, vbYesNoCancel, "Choose...")
        If intR = 6 Then 'yes
            boolAll = False
        ElseIf intR = 7 Then 'no
            boolAll = True
        Else
            GoTo end1
        End If

        Call MethValExecute(False, boolAll)

end1:

    End Sub

    Sub MethValExecute(boolStabOnly As Boolean, boolAll As Boolean)

        Dim dtbl As System.Data.DataTable
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim intRows As Short
        Dim Count1 As Short
        Dim var1, var2, var3
        Dim str1 As String
        Dim strM As String
        Dim intR As Short
        Dim boolM As Boolean = True

        If cmdEdit.Enabled Then
            Exit Sub
        End If

        'If Me.gbMethValApplyGuWu.Visible Then
        'Else
        '    strM = "This action intended for Sample Analysis studies."
        '    MsgBox(strM, vbInformation, "Invalid action...")
        '    GoTo end1
        'End If

        dtbl = tblMethodValidationData

        'ensure there is a value in study field
        dgv = Me.dgvMethValExistingGuWu

        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        dv = dgv.DataSource
        intRows = dv.Count
        If intRows = 0 Then
            Exit Sub
        End If
        For Count1 = 0 To intRows - 1
            'var1 = NZ(d(Count1, 1), "")
            var1 = NZ(dgv(1, Count1).Value, "")
            var2 = dgv(0, Count1).Value
            If Len(var1) = 0 Then
                If dgv.Rows(Count1).Selected Then

                    '20171106 LEE:
                    strM = "A Watson Study has not been chosen for " & var2 & "." & ChrW(10) & ChrW(10)
                    strM = strM & "If you wish to continue, an existing method validation study information will be cleared from the table." & ChrW(10) & ChrW(10)
                    strM = strM & "Do you wish to continue?"
                    intR = MsgBox(strM, vbYesNo, "Continue?")
                    If intR = 6 Then 'Yes
                        boolM = False
                    Else
                        GoTo end1
                    End If
                    'MsgBox("Please ensure a Watson study is configured.", MsgBoxStyle.Information, "Configure a study...")
                    'Exit Sub
                End If
            End If
        Next

        str1 = "This action will replace all existing information in the above table."
        str1 = str1 & ChrW(10) & ChrW(10)
        str1 = str1 & "Do you wish to continue?"
        If boolM Then
            intR = MsgBox(str1, vbYesNo, "Continue?")
        End If


        If intR = 6 Then
            Call RealMethValExecute(boolAll)
        Else
            Exit Sub
        End If

end1:

    End Sub


    Private Sub cmdShowOutstanding_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowOutstanding.Click

        If Len(NZ(cbxStudy.Text, "")) = 0 Then
            MsgBox("Please choose a study.", MsgBoxStyle.Information, "Choose a study...")
            dgvwStudy.Select()
            Exit Sub
        End If

        Dim Count1 As Short
        Dim Count2 As Short
        Dim str1 As String
        Dim dv As system.data.dataview
        Dim strS As String

        str1 = "id_tblStudies = " & id_tblStudies
        dv = New DataView(tblOutstandingItems)
        dv.RowFilter = str1
        Dim tbl As System.Data.DataTable
        tbl = dv.ToTable("tbl", True, "charSectionName", "charTabName", "charLocation", "charValue", "CHARFIELDCODE")
        'tbl = dv.ToTable("tbl", True, "charTabName", "charLocation", "charValue")
        Dim dv1 As system.data.dataview
        dv1 = tbl.DefaultView
        'strS = "charSectionName ASC, charTabName ASC"
        strS = "charTabName ASC, charLocation ASC, charValue ASC"
        dv1.Sort = strS

        If dv1.Count = 0 Then
            str1 = "There are no outstanding items for this report."
            MsgBox(str1, MsgBoxStyle.Information, "No outstanding items...")

        Else
            Dim frm As New frmReportErrMsg

            frm.dgvReportErrMsg.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            frm.dgvReportErrMsg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            frm.dgvReportErrMsg.DataSource = dv1
            'configure dgv
            'frm.dgvReportErrMsg.Columns.item(0).Visible = False
            frm.dgvReportErrMsg.Columns.Item("charSectionName").Visible = False
            frm.dgvReportErrMsg.Columns.Item("charSectionName").HeaderText = "Report Section Name"
            frm.dgvReportErrMsg.Columns.Item("charTabName").HeaderText = "StudyDoc Tab Name"
            frm.dgvReportErrMsg.Columns.Item("charLocation").HeaderText = "Report Item"
            frm.dgvReportErrMsg.Columns.Item("charValue").HeaderText = "Item Within Tab"
            frm.dgvReportErrMsg.Columns.Item("CHARFIELDCODE").HeaderText = "Field Code"

            For Count1 = 0 To frm.dgvReportErrMsg.ColumnCount - 1
                frm.dgvReportErrMsg.Columns.Item(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
                frm.dgvReportErrMsg.Columns.Item(Count1).DisplayIndex = frm.dgvReportErrMsg.ColumnCount - 1
            Next

            frm.dgvReportErrMsg.Columns.Item("charSectionName").DisplayIndex = 0 'HeaderText = "Report Section Name"
            frm.dgvReportErrMsg.Columns.Item("charLocation").DisplayIndex = 1 '.HeaderText = "Report Item"
            frm.dgvReportErrMsg.Columns.Item("charTabName").DisplayIndex = 2 '.HeaderText = "StudyDoc Tab Name"
            frm.dgvReportErrMsg.Columns.Item("charValue").DisplayIndex = 3 '.HeaderText = "Item Within Tab"

            frm.dgvReportErrMsg.AutoResizeColumns()
            frm.dgvReportErrMsg.AutoResizeRows()

            frm.lblEnd.Visible = False
            frm.lblStart.Visible = False
            frm.lblTotal.Visible = False

            frm.txtReportTitle.Text = lblReportTitle.Text

            ctArrReportNA = frm.dgvReportErrMsg.Rows.Count

            frm.ShowDialog()
            frm.Dispose()
            Refresh()
            'SendKeys.Send("%")

            ctArrReportNA = 0

        End If


    End Sub


    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick


    End Sub



    Private Sub cmdOrderReportBodySection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOrderReportBodySection.Click

        boolStopRBS = True

        Call OrderDGV(dgvReportStatements, "intOrder", "CHARSECTIONNAME")

        boolStopRBS = False

    End Sub

    Private Sub cmdOrderReportTableConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOrderReportTableConfig.Click

        'Re-# button is clicked
        'NDL 3-Dec-2015
        'This function will first put the non-included reports at the bottom (maintaining their order, however).  Then it re-orders all
        'tables according to number.  The re-ordering is done on all tables; this means that the function works when the tables
        'are filtered as well.

        '20170522 LEE: Doing all this causes the grid to reset, which is undesirable. Sacrifice move non-included reports to end for less cumbersome re-selecting of original row that has been re-positioned

        ''*****

        'Dim strFilterOrig, strSortOrig As String
        'Dim dv As DataView
        'Dim intCurrentTbl As Integer
        'dv = dgvReportTableConfiguration.DataSource

        ''Stop updates from occuring while re-ordering is being performed.
        'boolFormLoad = True

        ''Save original filter & sort
        'strFilterOrig = dv.RowFilter
        'strSortOrig = dv.Sort
        'intCurrentTbl = dgvReportTableConfiguration.CurrentRow.Cells.Item("ID_TBLREPORTTABLE").Value

        ''Change Filter & Sort to include all
        'dv.RowFilter = [String].Empty
        'dv.Sort = [String].Empty

        ''Determine highest number of table that is to be included
        'Dim intRow, intHighestIncludedRow, intRowOrderNum As Integer
        'intHighestIncludedRow = 0
        'For intRow = 0 To dgvReportTableConfiguration.RowCount - 1
        '    If (dv(intRow).Item("BOOLINCLUDE") = True) Then
        '        intRowOrderNum = dv(intRow).Item("INTORDER")
        '        If (intHighestIncludedRow < intRowOrderNum) Then
        '            intHighestIncludedRow = intRowOrderNum
        '        End If
        '    End If
        'Next

        ''Assign non-included tables numbers below the highest included table (while maintaining their order)
        'For intRow = 0 To dv.Count - 1
        '    If (dv(intRow).Item("BOOLINCLUDE") = False) Then
        '        intRowOrderNum = dv(intRow).Item("INTORDER")
        '        dv(intRow).BeginEdit()
        '        dv(intRow).Item("INTORDER") = intRowOrderNum + intHighestIncludedRow + 1
        '        dv(intRow).EndEdit()
        '    End If
        'Next

        ''Redo Sort in order to move non-included tables to bottom of the list
        'dv.Sort = "INTORDER ASC, BOOLPLACEHOLDER ASC"
        'boolFormLoad = False

        ''*****


        'Re-Number the tables based on their positions
        Call OrderDGV(dgvReportTableConfiguration, "INTORDER", "ID_TBLREPORTTABLE")

        ''Reset original filter & sort
        'dgvReportTableConfiguration.DataSource.RowFilter = strFilterOrig
        'dgvReportTableConfiguration.DataSource.Sort = strSortOrig

        'Rest numbers based on order
        Call AssessSampleAssignment()

        ''Rest current row back to what it was before (if there)
        'For intRow = 0 To dv.Count - 1
        '    If (dv(intRow).Item("ID_TBLREPORTTABLE") = intCurrentTbl) Then
        '        dgvReportTableConfiguration.CurrentCell = dgvReportTableConfiguration.Rows(intRow).Cells.Item("CHARHEADINGTEXT")
        '    End If
        'Next

        'Ensure that the table is re-drawn
        dgvReportTableConfiguration.AutoResizeRows()


    End Sub

    Private Sub dgvReportTableConfiguration_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportTableConfiguration.CellValueChanged

        Dim intRow As Short
        Dim intCol As Short
        Dim str1 As String
        Dim dgv As DataGridView
        Dim bool As Boolean
        Dim id As Long
        Dim var1, var2

        If boolFormLoad Then
            Exit Sub
        End If

        dgv = Me.dgvReportTableConfiguration

        If dgv.RowCount = 0 Then
            Exit Sub
        ElseIf dgv.CurrentRow Is Nothing Then
            Exit Sub
        End If
        intRow = dgv.CurrentRow.Index ' e.RowIndex
        intCol = dgv.CurrentCell.ColumnIndex 'e.ColumnIndex


        'ensure table can be displayed
        'id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
        'If StrComp(dgv.Columns(intCol).Name, "BOOLREQUIRESSAMPLEASSIGNMENT", CompareMethod.Text) = 0 Then
        '    If id = 1 Or id = 2 Then
        '        var1 = dgv("boolRequiresSampleAssignment", intRow).Value
        '        If var1 = -1 Then
        '            str1 = "This table cannot have samples assigned to it."
        '            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
        '            dgv("boolRequiresSampleAssignment", intRow).Value = 0
        '        End If
        '    End If
        'End If


        'Try
        '    bool = dgv.Item(intCol, intRow).Value
        '    Try
        '        If bool Then
        '            dgv.Item("BOOLREQUIRESSAMPLEASSIGNMENT", intRow).Value = -1
        '        Else
        '            dgv.Item("BOOLREQUIRESSAMPLEASSIGNMENT", intRow).Value = 0
        '        End If

        '    Catch ex As StackOverflowException

        '    End Try

        'Catch ex As Exception

        'End Try

        dgvReportTableConfiguration.AutoResizeRows()

        Call AssessSampleAssignment()

end1:


    End Sub


    Private Sub rbShowIncludedRTConfig_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbShowIncludedRTConfig.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Call RTFilter()

    End Sub


    Private Sub txtFilterSamples_Validated(sender As Object, e As EventArgs) Handles txtFilterSamples.Validated

        'Call RTFilter()

    End Sub

    Private Sub txtFilterSamples_KeyUp(sender As Object, e As KeyEventArgs) Handles txtFilterSamples.KeyUp

        If e.KeyCode = 13 Then
            'have the cursor move to the next item
            'this will call the txtfiltersamples_validated event
            Me.cmdClearFilters.Enabled = True
            Me.cmdClearFilters.Focus()
        End If
    End Sub

    Private Sub txtFilterSamples_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFilterSamples.TextChanged

        Call RTFilter()

    End Sub

    Private Sub cmdClearFilters_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClearFilters.Click
        'Clear Filters
        Cursor.Current = Cursors.WaitCursor
        Me.txtFilterSamples.Text = ""

        '20151104 LEE: commented out. RTFilter gets called automatically when .text is set to ""
        'Call RTFilter()

        Cursor.Current = Cursors.Default
    End Sub

    Private Sub rbShowIncludedRBody_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbShowIncludedRBody.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Call RBFilter()
    End Sub

    Private Sub cmdApplyTemplate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdApplyTemplate.Click

        'Dim strM As String
        'Dim intR As Short

        ''ensure study is configured
        'If id_tblStudies = 0 Then
        '    MsgBox("The study must first be configured before a template can be applied.", MsgBoxStyle.Information, "Configure study first...")
        '    GoTo end1
        'End If

        'strM = "Warning! If a new Study Template is applied, all changes to the current study will be reset to Study Template Defaults."
        'strM = strM & ChrW(10) & ChrW(10) & "Do you wish to continue?"

        'intR = MsgBox(strM, vbOKCancel, "Continue?")

        'If intR = 1 Then 'continue
        'Else
        '    GoTo end1
        'End If

        Call ApplyTemplateMasterHome() 'this is not the same as modConserved.ApplyTemplateMaster

end1:

    End Sub

    Sub ApplyTemplateMasterHome()

        Dim frm As New frmAssignTemplate
        Dim strM As String = ""
        Dim intR As Short
        Dim var1

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled = False Then
            strM = "No study has been loaded."
            strM = strM & ChrW(10) & "'Apply' action not allowed."
            frm.boolAllowApply = False
        Else
            If Me.cmdEdit.Enabled = False Then
                frm.boolAllowApply = True
                strM = "From the option group below:"
                strM = strM & ChrW(10) & "    Choose 'View studies...' to simply view the studies assigned to Study Templates."
                strM = strM & ChrW(10) & "    Choose 'Apply Study Template...' to apply a Study Template to the underlying study."
            Else
                frm.boolAllowApply = False
                strM = "Underlying study must be in Edit mode."
                strM = strM & ChrW(10) & "'Apply' action not allowed."
            End If
        End If

        frm.gLabel = strM
        If InStr(1, strM, "Choose", CompareMethod.Text) > 0 Then
            'frm.lbl1.Visible = False
            frm.rbApply.Enabled = True
        Else
            'frm.lbl1.Visible = True
            frm.rbApply.Enabled = False
        End If
        Call frm.SetLabel()

        frm.ShowDialog()
        If frm.boolCancel Then
            Exit Sub
        End If

        'ensure study is configured
        If id_tblStudies = 0 Then
            MsgBox("The study must first be configured before a template can be applied.", MsgBoxStyle.Information, "Configure study first...")
            GoTo end1
        End If

        strM = "Warning! If a new Study Template is applied, all changes to the current study will be reset to Study Template Defaults."
        strM = strM & ChrW(10) & ChrW(10) & "Do you wish to continue?"

        intR = MsgBox(strM, vbOKCancel, "Continue?")

        If intR = 1 Then 'continue
        Else
            GoTo end1
        End If

        Dim var2
        'var2 = frm.lbxTemplates.SelectedItem

        Dim dgv As DataGridView = frm.dgvTemplates
        Dim dgr As DataGridViewRow = dgv.SelectedRows(0)

        var2 = dgr.Cells("charStudyTemplate").Value

        frm.Dispose()
        Refresh()
        'SendKeys.Send("%")

        Try
            Call ApplyTemplate(var2)
        Catch ex As Exception
            var1 = var1
        End Try


        'annoying!
        Call ViewSections(False)

        'pesky
        Call RTFilter()

end1:

    End Sub


    Private Sub cmdResetSummaryTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResetSummaryTable.Click

        Call DoCancelSummaryTab()

    End Sub

    Private Sub cmdOrderSummaryTable_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOrderSummaryTable.Click

        Call OrderDGV(dgvSummaryData, "INTORDER", "CHARROWNAME")

    End Sub

    Private Sub rbShowAllSummaryTable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbShowAllSummaryTable.CheckedChanged
        Call ShowSummaryTable()

    End Sub


    Private Sub rbShowIncludedSummaryTable_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbShowIncludedSummaryTable.CheckedChanged
        Call ShowSummaryTable()
    End Sub


    Private Sub TimerRTC_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimerRTC.Tick
        Dim col1, col2, col3, col4, col5

        col1 = Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(207, Byte), Integer), CType(CType(176, Byte), Integer))
        col2 = Color.Gainsboro
        col3 = cmdAssignSamples.BackColor
        col4 = Color.White
        col5 = llblAssignedSamples.BackColor
        If col3 = col1 Then
            col3 = col2
        Else
            col3 = col1
        End If

        cmdAssignSamples.BackColor = col3
        If col5 = col4 Then
            col5 = col1
        Else
            col5 = col4
        End If
        llblAssignedSamples.BackColor = col5

    End Sub

    Private Sub cmdAssignSamples_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAssignSamples.Click

        Call OpenAssignedSamples(False)


    End Sub

    Private Sub llblAssignedSamples_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblAssignedSamples.LinkClicked

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim Count1 As Short

        str1 = "Configure Report Tables"
        int1 = lbxTab1.Items.Count
        For Count1 = 0 To int1 - 1
            str2 = lbxTab1.Items(Count1).ToString
            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                lbxTab1.SelectedIndex = Count1
                Exit For
            End If
        Next
    End Sub

    Sub LoginBad()

        Dim str2 As String
        Dim boolDo As Boolean


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
            'cmdLogin.Text = "&Log Off"
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
            Dim strPerm As String

            tblP = tblPersonnel
            tblU = tblUserAccounts
            tblPerm = tblPermissions

            'find user account
            strF = "id_tblUserAccounts = " & id_tblUserAccounts
            rowU = tblU.Select(strF)
            str1 = rowU(0).Item("charUserID")

            strF = "ID_TBLPERMISSIONS = " & rowU(0).Item("ID_TBLPERMISSIONS")
            rowPerm = tblPerm.Select(strF)
            strPerm = rowPerm(0).Item("CHARPERMISSIONSNAME")

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
                str2 = str2 & " v" & GetVersion() & " - Network User: " & str5 & " logged in as " & str3
            Else
                str2 = str2 & " v" & GetVersion() & " - StudyDoc User: " & str5 & " logged in as " & str1
            End If
            str2 = str2 & " assigned to Permissions Group: " & strPerm

            Text = str2

            gUserName = str5
            gUserID = str1

            'set permissions
            Call SetPermissions(True)

            Text = GetCaption("ReportWriter")

            MeCaption = Text

            Me.Text = MeCaption


            Dim boolB As Boolean
            If BOOLALLOWPDFREPORT = False And BOOLALLOWREPORTGENERATION = False Then
                boolB = False
            Else
                boolB = True
            End If
            Call LockReportGeneration(Not (boolB))

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

    End Sub

    Sub Login()

        Dim strT As String
        strT = "&Log In" ' cmdLogin.Text
        Dim str2 As String
        Dim boolDo As Boolean

        boolDo = True
        If StrComp(strT, "&Log In", CompareMethod.Text) = 0 Then 'log in
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
                boolInitLogIn = False
            Else

                boolInitLogIn = True
                'cmdLogin.Text = "&Log Off"
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
                Dim strPerm As String



                tblP = tblPersonnel
                tblU = tblUserAccounts
                tblPerm = tblPermissions

                'find user account
                strF = "id_tblUserAccounts = " & id_tblUserAccounts
                rowU = tblU.Select(strF)
                str1 = rowU(0).Item("charUserID")

                strF = "ID_TBLPERMISSIONS = " & rowU(0).Item("ID_TBLPERMISSIONS")
                rowPerm = tblPerm.Select(strF)
                strPerm = rowPerm(0).Item("CHARPERMISSIONSNAME")

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
                    str2 = str2 & " v" & GetVersion() & " - Network User: " & str5 & " logged in as " & str3
                Else
                    str2 = str2 & " v" & GetVersion() & " - StudyDoc User: " & str5 & " logged in as " & str1
                End If
                'str2 = str2 & " v" & GetVersion() & " - User: " & str5 & " logged in as " & str1
                str2 = str2 & " assigned to Permissions Group: " & strPerm

                Text = str2

                gUserName = str5
                gUserID = str1

                'set permissions
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
            intInput = MsgBox("Are you sure you wish to log off?", MsgBoxStyle.YesNo, "Log off check...")
            If intInput = 6 Then 'continue

                gUserID = ""
                gUserName = ""

                Cursor.Current = Cursors.WaitCursor

                id_tblPersonnel = 0
                id_tblUserAccounts = 0

                Call DoThis("Logoff")

                'cmdOpenReportStatements.Enabled = False

                Call SetPermissions(False)


                str2 = GetStudyDocHeader(False)
                str2 = str2 & " - User: Guest"
                Text = str2

                'cmdLogin.Text = "&Log In"
                ''select cmdLogin
                'cmdLogin.Select()
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
    End Sub


    Sub ConfigAnalRef()

        Dim Count1 As Integer
        Dim dgv As DataGridView

        dgv = Me.dgvCompanyAnalRef

        For Count1 = 0 To dgv.Columns.Count - 1

            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable

        Next

        dgv = Me.dgvWatsonAnalRef

        For Count1 = 0 To dgv.Columns.Count - 1

            dgv.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable

        Next


    End Sub

    Private Sub cmdAddAnalyte_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddAnalyte.Click

        Dim var1, var2
        Dim strRet As String
        Dim boolGo As Boolean
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim frm As New frmAddCompound
        Dim boolCancel As Boolean
        Dim strCo As String
        Dim dgv As DataGridView
        Dim boolF As Boolean

        dgv = Me.dgvCompanyAnalRef
        boolF = dgv.Columns.Item("Item").Frozen
        dgv.Columns.Item("Item").Frozen = False

        boolGo = False
        tbl = tblAnalRefStandards


        Do Until boolGo
            frm.txtName.Text = ""
            frm.rbYes.Checked = True

            frm.ShowDialog()

            boolCancel = frm.boolCancel
            If boolCancel Then
                GoTo end1
            End If
            strRet = frm.txtName.Text
            If frm.rbYes.Checked Then
                strCo = "Yes"
            Else
                strCo = "No"
            End If

            'strRet = InputBox("Enter name of new Analytical Reference Standard:", "Enter new Analytical Reference Standard...")

            'check to ensure name is unique
            strF = "id_tblStudies = " & id_tblStudies & " AND charAnalyteName = '" & strRet & "'"
            rows = tbl.Select(strF)
            If rows.Length = 0 Then
                boolGo = True
            Else
                MsgBox(strRet & " is already configured for this study.", MsgBoxStyle.Information, strRet & " already configured...")
            End If
        Loop
        frm.Dispose()
        Refresh()
        'SendKeys.Send("%")

        If boolGo = False Then
            GoTo end1
        End If


        Dim col1 As New DataColumn
        Dim strAnal As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim ct1 As Short
        Dim ct2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dt As System.Data.DataTable
        Dim dtC As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim rw As DataRow
        Dim col As DataColumn
        Dim strName As String
        Dim int1 As Short
        Dim int2 As Short
        Dim boolC As Boolean

        dtC = tblCompanyAnalRefTable
        dv = Me.dgvCompanyAnalRef.DataSource
        ct1 = dtC.Columns.Count
        strAnal = strRet
        'ts1 = dgCompanyAnalRef.TableStyles(0)
        col1.DataType = System.Type.GetType("System.String")
        col1.ColumnName = strAnal
        col1.Caption = strAnal
        dtC.Columns.Add(col1)
        'enter Analyte Name info in dtC
        int2 = FindRowDVByCol("Analyte Name", dv, "Item")
        dtC.Rows.Item(int2).BeginEdit()
        dtC.Rows.Item(int2).Item(ct1) = strAnal
        dtC.Rows.Item(int2).EndEdit()
        'enter Is Replicate info in dtC
        int2 = FindRowDVByCol("Is Replicate?", dv, "Item")
        dtC.Rows.Item(int2).BeginEdit()
        dtC.Rows.Item(int2).Item(ct1) = "No"
        dtC.Rows.Item(int2).EndEdit()
        'enter Is Coadministered Cmpd info in dtC
        int2 = FindRowDVByCol("Is Coadministered Cmpd?", dv, "Item")
        dtC.Rows.Item(int2).BeginEdit()
        dtC.Rows.Item(int2).Item(ct1) = strCo
        dtC.Rows.Item(int2).EndEdit()
        'enter Is Configured in Watson info in dtC
        int2 = FindRowDVByCol("Is Configured in Watson?", dv, "Item")
        dtC.Rows.Item(int2).BeginEdit()
        dtC.Rows.Item(int2).Item(ct1) = "No"
        dtC.Rows.Item(int2).EndEdit()
        'enter Analyte Parent info in dtC
        int2 = FindRowDVByCol("Analyte Parent", dv, "Item")
        dtC.Rows.Item(int2).BeginEdit()
        dtC.Rows.Item(int2).Item(ct1) = strAnal
        dtC.Rows.Item(int2).EndEdit()
        int2 = FindRowDVByCol("Is Internal Standard?", dv, "Item")
        dtC.Rows.Item(int2).BeginEdit()
        dtC.Rows.Item(int2).Item(ct1) = "No"
        dtC.Rows.Item(int2).EndEdit()

        dv = New DataView(dtC)
        dv.AllowDelete = False
        dv.AllowNew = False
        dgv.DataSource = dv
        dgv.Refresh()
        int1 = dgv.Columns.Count
        int2 = Me.dgvCompanyAnalRef.Columns.Count
        For Count1 = 0 To int1 - 1
            str1 = dgv.Columns.Item(Count1).Name
            var1 = str1
        Next
        int2 = dtC.Columns.Count

        dgv.Columns.Item("BOOLINCLUDE").HeaderText = "A*"
        dgv.AutoResizeColumns()


        str1 = AnalRefHook()
        If Len(str1) > 0 Then
            Select Case str1
                Case Is = "CRLWor_AnalRefStandard"
                    Call ComboBoxCRLAnalRefFill()
            End Select
        End If

        Call HideAnalRefRows()


        int1 = Me.dgvCompanyAnalRef.Columns.Count
        Me.dgvCompanyAnalRef.Columns.Item(int1 - 1).SortMode = DataGridViewColumnSortMode.NotSortable

        Call SyncCols(Me.dgvWatsonAnalRef, Me.dgvCompanyAnalRef)

        Call ResizeRows(Me.dgvCompanyAnalRef)
        Call ResizeRows(Me.dgvWatsonAnalRef)

        dgv.Columns.Item("Item").Frozen = boolF


end1:


    End Sub

    Private Sub cbxFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxFilter.SelectedIndexChanged

        Dim dv As system.data.dataview
        Dim dv1 As system.data.dataview
        Dim strF As String
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim int1 As Short
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String

        str2 = cbxFilter.Text
        dv = dgvwStudy.DataSource
        tbl = tblStudies
        If StrComp(str2, "All", CompareMethod.Text) = 0 Then
            str1 = ""
        Else
            strF = "charCust = '" & str2 & "'"
            rows = tbl.Select(strF)
            int1 = rows.Length
            str1 = "StudyID = '" & rows(0).Item("int_WatsonStudyID") & "'"
            For Count1 = 1 To int1 - 1
                str1 = str1 & " OR StudyID = '" & rows(Count1).Item("int_WatsonStudyID") & "'"
            Next
        End If
        dv.RowFilter = str1
        '20161102 LEE: Don't need to do this
        'dv.AllowDelete = False
        'dv.AllowNew = False
        'dgvwStudy.DataSource = dv


    End Sub

    Private Sub dgvContributingPersonnel_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvContributingPersonnel.CellClick

        Call ContrPersDropdowns(e.RowIndex, e.ColumnIndex)

    End Sub

    Sub ContrPersDropdowns(ByVal intRow As Short, ByVal intCol As Short)

        Dim str1 As String
        Dim str2 As String
        Dim dgv As DataGridView
        Dim boolGo As Boolean
        Dim var1
        Dim boolSort As Boolean

        If boolFormLoad Then
            Exit Sub
        End If

        If cmdEdit.Enabled Or Len(cbxStudy.Text) = 0 Then
            Exit Sub
        End If

        If intRow < 0 Or intCol < 0 Then
            Exit Sub
        End If


        dgv = dgvContributingPersonnel
        str1 = dgv.Columns.Item(intCol).Name
        If InStr(1, str1, "bool", CompareMethod.Text) > 0 Then 'ignore
            'ElseIf InStr(1, str1, "int", CompareMethod.Text) > 0 Then 'ignore
        Else
            str2 = NZ(dgv.Rows.Item(intRow).Cells(intCol).EditType.FullName, "")
            'If InStr(1, str2, "combobox", CompareMethod.Text) > 0 And boolHomeCBox = False Then
            If InStr(1, str2, "combobox", CompareMethod.Text) > 0 And boolHomeCBox Then
            Else
                Dim cbx As New DataGridViewComboBoxCell
                boolGo = False
                boolSort = True
                Select Case str1
                    Case "CHARCPPREFIX"
                        cbx = cbxxCPPrefix.Clone
                        boolGo = True
                    Case "CHARCPNAME"
                        cbx.Sorted = False
                        cbxxCPName.Sorted = False
                        cbx = cbxxCPName.Clone
                        boolGo = True
                        boolSort = False
                    Case "CHARCPSUFFIX"
                        cbx = cbxxCPSuffix.Clone
                        boolGo = True
                    Case "CHARCPDEGREE"
                        cbx = cbxxCPDegree.Clone
                        boolGo = True
                    Case "CHARCPTITLE"
                        cbx = cbxxCPTitle.Clone
                        boolGo = True
                    Case "CHARCPROLE"
                        cbx = cbxxCPRole.Clone
                        boolGo = True
                End Select
                If boolGo Then
                    Dim var2
                    cbx.Sorted = boolSort
                    var1 = dgv.Columns.Item(intCol).Width
                    var2 = var1 * 1.75
                    'cbx.DropDownWidth = var2
                    'if data doesn't exist in dropdown list
                    'data error will be called that inserts unlisted value into dropdown box
                    Try
                        dgv(intCol, intRow) = cbx
                    Catch ex As Exception

                    End Try
                End If

            End If
        End If

    End Sub

    Private Sub dgvContributingPersonnel_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvContributingPersonnel.CellContentClick

        Dim str1 As String
        Dim dgv As DataGridView
        Dim boolG As Boolean
        Dim boolV As Boolean
        Dim str2 As String
        Dim int1 As Short

        If e.RowIndex < 0 Then
            Exit Sub
        End If
        dgv = dgvContributingPersonnel
        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        boolG = False
        Select Case str1
            Case "boolIncludeSOTP"
                boolG = True
                str2 = "boolIncludeSigOnTablePage"
        End Select
        If boolG Then
            Dim dv As system.data.dataview
            dv = dgv.DataSource

            'dgv.EndEdit(True)
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            boolV = dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value
            If boolV Then
                int1 = -1
            Else
                int1 = 0
            End If
            dv(e.RowIndex).BeginEdit()
            dv(e.RowIndex).Item(str2) = int1
            'dv(e.RowIndex).Item(str1) = Not (boolV)
            dv(e.RowIndex).EndEdit()

        End If


    End Sub


    Private Sub cmdHook_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdHook.Click
        If id_tblPersonnel = 0 Then
            'MsgBox("Guest not allowed to run hook", MsgBoxStyle.Information, "Guest not allowed...")
            'GoTo end1
        End If

        'Call Hook_CRL_AnalRef() 'for testing purposes

end1:

    End Sub


    Sub SizeCompanyAnalRef()


        'always give grids half the space
        'anchors just don't seem to be working for this panel
        'mysterious

        'only do heights

        Dim dgv As DataGridView
        Dim dgvW As DataGridView

        dgv = Me.dgvCompanyAnalRef
        dgvW = Me.dgvWatsonAnalRef

        'dgv.Anchor = AnchorStyles.None
        'dgvW.Anchor = AnchorStyles.None

        Dim a, b, c, d, e, f, g, h, ah, l1, w, w1, t1


        'find 1/2 of tp8
        a = Me.tp8.Height
        w = Me.tp8.Width
        l1 = Me.lblARST.Left
        t1 = Me.lblARST.Top + Me.lblARST.Height
        w1 = w - l1 - l1
        a = a - t1

        ah = a
        b = (a / 2) + Me.lblARST.Top + Me.lblARST.Height
        Me.lblWAR.Top = b
        'Me.lblWAR.Left = l1

        dgv.Top = t1 'a ' Me.lblARST.Top + Me.lblARST.Height
        'dgv.Left = l1
        'dgv.Width = w1

        dgv.Height = Me.lblWAR.Top - t1 - 20

        e = Me.lblWAR.Top
        f = Me.lblWAR.Height
        g = e + f
        dgvW.Top = g
        h = a - g
        dgvW.Height = dgv.Height ' h - 5
        'dgvW.Width = dgv.Width
        ' dgvW.Left = dgv.Left
        'dgvW.Left = l1


    End Sub

    Private Sub dgvDataCompany_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDataCompany.CellClick

        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strF As String


        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim locX, locY
        Dim var1, var2, var3
        Dim intRow As Short

        If e.ColumnIndex = 1 Then
        Else
            Exit Sub
        End If

        Dim dgv As DataGridView
        dgv = dgvDataCompany
        Dim dv As System.Data.DataView
        dv = dgv.DataSource

        If dgv.ReadOnly Then
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        'int1 = FindRowDV("Table Date Format", dv)
        'int2 = FindRowDV("Text Date Format", dv)
        int3 = FindRowDV("Study Start Date", dv)
        int4 = FindRowDV("Study End Date", dv)

        Me.panCal.Visible = False

        var1 = NZ(dgv.Rows(intRow).Cells(1).Value, "")

        If int3 = e.RowIndex Or int4 = e.RowIndex Then 'show calendar

            boolFromDataTab = True
            locX = Me.tab1.Left + dgv.Left + dgv.RowHeadersWidth + dgv.Columns(0).Width + dgv.Columns(1).Width

            locY = Me.tab1.Top + Me.tabData.Top + dgv.Location.Y + (dgv.Rows(intRow).Height * intRow) + dgv.ColumnHeadersHeight

            Dim dt As Date

            gCalGrid = dgv

            If IsDate(var1) Then
                dt = var1
            Else
                dt = Now
            End If

            Call MakeCalVis(locX, locY, dt, True)

        Else
            Call MakeCalVis(0, 0, Now, False)
        End If

        'If int3 = e.RowIndex Or int4 = e.RowIndex Then 'show calendar

        '    boolFromDataTab = True
        '    locX = Me.tab1.Left + dgv.Left + dgv.RowHeadersWidth + (dgv.Columns(1).Width * 2)
        '    locY = Me.tab1.Top + Me.tabData.Top + dgv.Location.Y + (dgv.Rows(intRow).Height * intRow) + dgv.ColumnHeadersHeight

        '    If IsDate(var1) Then
        '        Me.dtp1.Value = var1
        '    Else
        '        Me.dtp1.Value = Now
        '    End If
        '    Me.dtp1.Location = New System.Drawing.Point(locX, locY)
        '    'Me.dtp1.ScrollChange = 1
        '    'Me.dtp1.MaxSelectionCount = 1

        '    Me.dtp1.Visible = True

        '    Me.dtp1.BringToFront()

        'End If


    End Sub

    Sub MakeCalVis(locX As Single, locY As Single, dtS As Date, boolVis As Boolean)

        frmH.boolHold = True

        If boolVis Then

            'panCal must not go outside of viewable area or Enter button will be hidden
            'this is especially true for Draft date and Report date on Home tab

            Dim a, b, c, d
            Dim w, h

            w = My.Computer.Screen.WorkingArea.Width
            h = My.Computer.Screen.WorkingArea.Height

            Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
            Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

            c = h - (tbh * 2)

            a = locY + Me.panCal.Height
            If a > c Then
                b = h - Me.panCal.Height - tbh
                locY = b
            End If

            Me.panCal.Location = New System.Drawing.Point(locX, locY)

            Me.mCal1.SelectionStart = dtS
            Me.mCal1.SelectionEnd = dtS

            Me.mCal1.ScrollChange = 1
            Me.mCal1.MaxSelectionCount = 1

            Me.mCal1.BringToFront()
            Me.cmdCalCancel.BringToFront()
            Me.cmdEnterCal.BringToFront()
            Me.panCal.BringToFront()

            Me.panCal.Visible = True

            Me.mCal1.Focus()

        Else

            Me.panCal.Visible = False

        End If

        frmH.boolHold = False

    End Sub


    Private Sub cmdEnterCal_Click(sender As System.Object, e As System.EventArgs) Handles cmdEnterCal.Click


        Dim dt As Date
        Dim intRow As Short
        Dim intCol As Short
        Dim dgv As DataGridView = gCalGrid
        Dim strM As String
        Dim var1

        Try
            var1 = Me.mCal1.SelectionRange.Start
            dt = CDate(var1)

            intRow = gCalGrid.CurrentRow.Index

            Select Case gCalGrid.Name
                Case "dgvDataCompany"
                    intCol = dgv.CurrentCell.ColumnIndex
                Case "dgvReports"
                    intCol = dgv.CurrentCell.ColumnIndex
                Case "dgvSampleReceipt"
                    intCol = dgv.Columns("DTSHIPMENTRECEIVED").Index
            End Select

            If dgv.CurrentRow Is Nothing Then
                'find row of current cell
                intRow = dgv.CurrentCell.RowIndex
            Else
                intRow = dgv.CurrentRow.Index
            End If

            Try
                dgv.Rows(intRow).Cells(intCol).Value = Format(dt, LDateFormat)
            Catch ex As Exception
                dgv.Rows(intRow).Cells(intCol).Value = Format(dt, GDateFormat)
            End Try

        Catch ex As Exception
            strM = "There was a problem entering the selected date."
            strM = strM & ChrW(10) & ChrW(10) & ex.Message
            MsgBox(strM, vbInformation, "Problem...")
        End Try

        Me.panCal.Visible = False

    End Sub

    Private Sub cmdCalCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCalCancel.Click

        Me.panCal.Visible = False

    End Sub


    Private Sub dgvDataCompany_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvDataCompany.CellValidating

        Dim var1, var2
        Dim dgv As DataGridView
        Dim intRowDate1 As Short
        Dim intRow As Short
        Dim intCol As Short
        Dim varNull As System.DBNull
        Dim str1 As String
        Dim dt As Date
        Dim dv As System.Data.DataView
        Dim strM As String

        If e.ColumnIndex = 1 Then
        Else
            Exit Sub
        End If

        If frmH.boolHold Then
            Exit Sub
        End If

        dgv = Me.dgvStudyConfig
        str1 = dgv.Rows.Item(e.RowIndex).Cells(0).Value

        'check for sigfig value
        Dim boolCancel As Boolean
        intRow = e.RowIndex
        intCol = e.ColumnIndex
        If dgv(intCol, intRow).ReadOnly = True Then
            Exit Sub
        End If

        'now check for column limit
        Dim strMod As String = "Add/Edit Top Level Data - Study Information"
        Dim strSource As String = str1

        If boolCLExceeded(str1, "TBLDATA", e.FormattedValue, False, strMod, strSource) Then
            e.Cancel = True
            GoTo end1
        End If


        'check for dates
        dgv = dgvDataCompany
        dv = dgv.DataSource

        Dim int3 As Short
        Dim int4 As Short
        int3 = FindRowDV("Study Start Date", dv)
        int4 = FindRowDV("Study End Date", dv)

        Me.panCal.Visible = False

        var1 = NZ(dgv.Rows(intRow).Cells(1).Value, "")

        'use e.FormattedValue, not value
        var2 = e.FormattedValue

        If Len(var2) = 0 Then

        Else
            If int3 = e.RowIndex Or int4 = e.RowIndex Then 'show calendar

                If IsDate(var2) Then
                Else

                    strM = "Entry must be date."
                    strM = strM & ChrW(10) & ChrW(10) & "Press the Escape key to cancel action."
                    MsgBox(strM, vbInformation, "Invalid entry...")
                    e.Cancel = True

                    GoTo end1

                End If


            End If
        End If



end1:

    End Sub


    Private Sub dgvContributingPersonnel_CellEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvContributingPersonnel.CellEnter

        'Call ContrPersDropdowns(e.RowIndex, e.ColumnIndex)
    End Sub

    Private Sub dgvContributingPersonnel_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvContributingPersonnel.DataError

        Dim var1
        Dim var2
        Dim dgv As DataGridView
        Dim str1 As String
        Dim boolGo As Boolean
        Dim cbx As DataGridViewComboBoxCell
        Dim cbx1 As New DataGridViewComboBoxCell

        dgv = dgvContributingPersonnel
        var1 = dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value
        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        boolGo = False
        Select Case str1
            Case "CHARCPPREFIX"
                cbx = cbxxCPPrefix
                boolGo = True
            Case "CHARCPNAME"
                cbx = cbxxCPName
                boolGo = True
            Case "CHARCPSUFFIX"
                cbx = cbxxCPSuffix
                boolGo = True
            Case "CHARCPDEGREE"
                cbx = cbxxCPDegree
                boolGo = True
            Case "CHARCPTITLE"
                cbx = cbxxCPTitle
                boolGo = True
            Case "CHARCPROLE"
                cbx = cbxxCPRole
                boolGo = True
        End Select
        If boolGo Then
            cbx.Items.Add(var1)
            cbx1 = cbx.Clone
            dgv(e.ColumnIndex, e.RowIndex) = cbx1
            'dgv(e.ColumnIndex, e.RowIndex).Value = var1
        End If


    End Sub




    Private Sub dgvDataCompany_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvDataCompany.DataError
        e.Cancel = True
    End Sub


    Private Sub dgvReports_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReports.CellClick

        Dim str1 As String
        Dim str2 As String
        Dim dgv As DataGridView
        Dim var1, var2
        Dim boolDt As Boolean
        Dim boolR As Boolean
        Dim boolRTemp As Boolean
        Dim boolRType As Boolean

        Dim boolDate As Boolean

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled Then
        Else
            Exit Sub
        End If

        If cmdEdit.Enabled Or Len(cbxStudy.Text) = 0 Then
            Exit Sub
        End If

        If e.RowIndex < 0 Or e.ColumnIndex < 0 Then
            Exit Sub
        End If

        boolDt = False
        boolR = False
        boolRTemp = False
        boolRType = False
        boolDate = False

        dgv = dgvReports
        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        If StrComp(str1, "CHARREPORTTEMPLATE", CompareMethod.Text) = 0 Then
            boolR = True
            boolDt = False
            boolRTemp = True
            boolRType = False

            boolRTemp = False 'not doing this anymore
            boolR = False 'not doing this anymore
            boolDate = False

        ElseIf StrComp(str1, "CHARREPORTTYPE", CompareMethod.Text) = 0 Then
            boolR = True
            boolDt = False
            boolRTemp = False
            boolRType = True
            boolDate = False

        ElseIf StrComp(str1, "DTREPORTDRAFTISSUEDATE", CompareMethod.Text) = 0 Then
            boolDate = True
        ElseIf StrComp(str1, "DTREPORTFINALISSUEDATE", CompareMethod.Text) = 0 Then
            boolDate = True
        End If

        If boolDate Then

            dgv = Me.dgvReports

            Dim intRow As Short
            Dim intCol As Short
            Dim strName As String

            intRow = e.RowIndex
            intCol = e.ColumnIndex
            strName = dgv.Columns(intCol).Name

            Dim locX, locY

            If InStr(1, strName, "dt", CompareMethod.Text) > 0 Then 'show calendar

                var1 = dgv(intCol, intRow).Value

                'boolFromTab = True
                Dim ld As Single
                Dim Count1 As Short
                ld = 0
                For Count1 = 0 To intCol
                    If dgv.Columns(Count1).Visible Then
                        ld = ld + dgv.Columns(intCol).Width
                    End If
                Next
                locX = Me.tab1.Left + dgv.Left + dgv.RowHeadersWidth + ld
                'want this calendar position to be a bit higher
                'locY = Me.tab1.Top + dgv.Location.Y + (dgv.Rows(intRow).Height * intRow) + dgv.ColumnHeadersHeight
                locY = Me.tab1.Top + Me.tp1.Top + dgv.Location.Y ' + (dgv.Rows(intRow).Height * intRow) + dgv.ColumnHeadersHeight

                Dim dt As Date

                If IsDate(var1) Then
                    dt = CDate(var1)
                Else
                    dt = Now
                End If

                gCalGrid = dgvReports

                Call MakeCalVis(locX, locY, dt, True)

            End If

        Else

            Call MakeCalVis(0, 0, Now, False)

        End If

        If boolR Then
            str2 = NZ(dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).EditType.FullName, "")
            If InStr(1, str2, "combobox", CompareMethod.Text) > 0 Then
            Else
                Select Case str1
                    Case "CHARREPORTTEMPLATE"
                        'var1 = dgv.Columns.item(e.ColumnIndex).Width
                        'var2 = var1 * 2
                        'cbxxReportTemplates.DropDownWidth = var2
                        'On Error Resume Next
                        'dgv(e.ColumnIndex, e.RowIndex) = cbxxReportTemplates
                        'If Err.Number <> 0 Then
                        '    Err.Clear()
                        'End If
                        'On Error GoTo 0



                        Dim cbx As New DataGridViewComboBoxCell
                        cbx = cbxxReportTemplates.Clone
                        var1 = dgv.Columns.Item(e.ColumnIndex).Width
                        var2 = var1 * 2
                        cbx.DropDownWidth = var2
                        'if data doesn't exist in dropdown list
                        'data error will be called that inserts unlisted value into dropdown box
                        On Error Resume Next
                        dgv(e.ColumnIndex, e.RowIndex) = cbx
                        If Err.Number <> 0 Then
                            Err.Clear()
                        End If
                        On Error GoTo 0
                    Case "CHARREPORTTYPE"
                        'var1 = dgv.Columns.item(e.ColumnIndex).Width
                        'var2 = var1 * 2
                        'cbxxReportTypes.DropDownWidth = var2
                        ''On Error Resume Next
                        'dgv(e.ColumnIndex, e.RowIndex) = cbxxReportTypes
                        'If Err.Number <> 0 Then
                        '    Err.Clear()
                        'End If
                        'On Error GoTo 0

                        Dim cbx1 As New DataGridViewComboBoxCell
                        cbx1 = cbxxReportTypes.Clone
                        var1 = dgv.Columns.Item(e.ColumnIndex).Width
                        var2 = var1 * 2
                        cbx1.DropDownWidth = var2
                        'if data doesn't exist in dropdown list
                        'data error will be called that inserts unlisted value into dropdown box
                        On Error Resume Next
                        dgv(e.ColumnIndex, e.RowIndex) = cbx1
                        If Err.Number <> 0 Then
                            Err.Clear()
                        End If
                        On Error GoTo 0
                End Select
            End If
        End If

    End Sub


    Private Sub dgvReports_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvReports.CellValidating

        If boolFormLoad Then
            Exit Sub
        End If

        If cmdEdit.Enabled Or Len(cbxStudy.Text) = 0 Then
            Exit Sub
        End If

        If e.RowIndex < 0 Or e.ColumnIndex < 0 Then
            Exit Sub
        End If

        Dim var1
        Dim var2
        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String
        Dim intCL As Short
        Dim boolStop As Boolean

        Dim strMod As String = "Choose Study & Template - Configure Reports Table"
        Dim strSource As String = ""


        dgv = dgvReports
        var1 = dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value
        str1 = dgv.Columns.Item(e.ColumnIndex).Name

        boolStop = False
        Select Case str1

            Case "CHARREPORTNUMBER"
                strSource = "Report Number"
                intCL = 255
                boolStop = CheckColLenEx(e.FormattedValue, intCL, strMod, strSource)

            Case "CHARREPORTTITLE"
                strSource = "Report Title"
                intCL = 255
                boolStop = CheckColLenEx(e.FormattedValue, intCL, strMod, strSource)

        End Select

        If boolStop Then
            e.Cancel = True
            GoTo end1
        End If

        If StrComp(Mid(str1, 1, 2), "DT", CompareMethod.Text) = 0 Then
            'format date for ldateformat
            If Len(e.FormattedValue) = 0 Then 'ignore
            ElseIf IsDate(e.FormattedValue) Then 'continue
                Dim dv As system.data.dataview
                dv = dgv.DataSource
                str1 = dgv.Columns.Item(e.ColumnIndex).Name
                'dgv(e.ColumnIndex, e.RowIndex).Value = var1
                var1 = Format(CDate(e.FormattedValue), LDateFormat)
                'dgv(e.ColumnIndex, e.RowIndex).Value = var1
                'dv(e.RowIndex).Item(str1) = var1
                'dgv.Update()
            Else
                MsgBox("Entry must be in an appropriate date format.", MsgBoxStyle.Information, "Invalid entry...")
                e.Cancel = True
            End If
        ElseIf StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then
            Call SetReportConfigType()
        End If

end1:

    End Sub

    Private Sub dgvReports_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReports.CellValueChanged

        Dim dgv As DataGridView
        Dim str1 As String
        Dim var1, var2
        Dim int1 As Short

        dgv = dgvReports
        str1 = dgv.Columns.Item(e.ColumnIndex).Name

        If StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then
            Call SetReportConfigType()
        End If

    End Sub

    Private Sub dgvReports_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReports.Click
        Call ReportsSelection(False)

        Call Set_idtblReports()
        'Call SetReportHistory()
    End Sub


    Private Sub dgvReports_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvReports.DataError
        Dim var1
        Dim var2, var3, var4
        Dim dgv As DataGridView
        Dim str1 As String
        Dim boolGo As Boolean
        'Dim cbx As DataGridViewComboBoxCell
        Dim cbx1 As New DataGridViewComboBoxCell
        Dim cbx2 As New DataGridViewComboBoxCell
        Dim boolRTemp As Boolean
        Dim boolRType As Boolean
        Dim boolCont As Boolean
        Dim Count1 As Short
        Dim int1 As Short

        Exit Sub

        boolRTemp = False
        boolRType = False

        'dgv = dgvReports
        'var1 = dgv.Rows.item(e.RowIndex).Cells(e.ColumnIndex).Value
        dgv = dgvReports
        var1 = dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value
        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        boolGo = False
        Select Case str1
            Case "CHARREPORTTEMPLATE"
                'cbx = cbxxReportTemplates
                boolGo = True
                boolRTemp = True
                boolRType = False

            Case "CHARREPORTTYPE"
                'cbx = cbxxReportTypes
                boolGo = True
                boolRTemp = False
                boolRType = True

        End Select
        Dim varE
        varE = e.Exception.ToString
        '''''''''''''console.writeline(varE)
        'debug.writeline(varE)
        If boolGo Then
            If boolRTemp Then
                'first ensure anything has to be done
                int1 = cbxxReportTemplates.Items.Count
                'boolCont = True
                'For Count1 = 0 To int1 - 1
                '    var2 = cbxxReportTemplates.Items(Count1)
                '    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                '        boolCont = False
                '        Exit For
                '    End If
                'Next
                'If boolCont Then
                cbxxReportTemplates.Items.Add(var1)
                cbx1 = cbxxReportTemplates.Clone
                var3 = dgv.Columns.Item(e.ColumnIndex).Width
                var4 = var3 * 2
                cbx1.DropDownWidth = var4
                dgv(e.ColumnIndex, e.RowIndex) = cbx1
                'dgv(e.ColumnIndex, e.RowIndex).Value = var1
                'End If

            ElseIf boolRType Then
                'first ensure anything has to be done
                int1 = cbxxReportTypes.Items.Count
                'boolCont = True
                'For Count1 = 0 To int1 - 1
                '    var2 = cbxxReportTypes.Items(Count1)
                '    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                '        boolCont = False
                '        Exit For
                '    End If
                'Next
                'boolCont = True
                'If boolCont Then
                cbxxReportTypes.Items.Add(var1)
                cbx2 = cbxxReportTypes.Clone
                var3 = dgv.Columns.Item(e.ColumnIndex).Width
                var4 = var3 * 2
                cbx2.DropDownWidth = var4
                dgv(e.ColumnIndex, e.RowIndex) = cbx2
                'dgv(e.ColumnIndex, e.RowIndex).Value = var1
                'End If

            End If
        End If
        'MsgBox(e.Exception.ToString)


    End Sub

    Private Sub dgvReports_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles dgvReports.EditingControlShowing

        Dim var1
        Dim intCol As Short
        Dim str1 As String
        Dim str2 As String


        'Exit Sub

        intCol = dgvReports.CurrentCell.ColumnIndex
        str1 = dgvReports.Columns.Item(intCol).Name
        If StrComp(str1, "CHARREPORTTYPE", CompareMethod.Text) = 0 Then
            str2 = e.Control.GetType.Name.ToString
            If InStr(1, str2, "Combobox", CompareMethod.Text) > 0 Then

                Dim combo As ComboBox = CType(e.Control, ComboBox)
                If (combo IsNot Nothing) Then

                    ' Remove an existing event-handler, if present, to avoid 
                    ' adding multiple handlers when the editing control is reused.
                    RemoveHandler combo.SelectionChangeCommitted, New EventHandler(AddressOf ComboBox_SelectionChangeCommitted)
                    boolEventAdd = True
                    If boolEventAdd Then
                        ' Add the event handler. 
                        AddHandler combo.SelectionChangeCommitted, New EventHandler(AddressOf ComboBox_SelectionChangeCommitted)
                        boolEventAdd = False
                    End If
                End If
            End If
        End If
    End Sub


    Private Sub ComboBox_SelectionChangeCommitted(ByVal sender As Object, ByVal e As EventArgs)

        Dim var1
        Dim str1 As String
        Dim intRow As Short
        Dim intCol As Short
        Dim varE

        intCol = dgvReports.CurrentCell.ColumnIndex
        str1 = dgvReports.Columns.Item(intCol).Name
        If StrComp(str1, "CHARREPORTTYPE", CompareMethod.Text) = 0 Then

            Dim comboBox1 As ComboBox = CType(sender, ComboBox)
            Dim dgv As DataGridView

            dgv = dgvReports
            intRow = dgv.CurrentRow.Index
            'record id_tblconfigreporttype
            str1 = NZ(comboBox1.Text, "Sample Analysis")
            Dim tbl As System.Data.DataTable
            Dim tbl1 As System.Data.DataTable
            Dim strF As String
            Dim dr() As DataRow
            Dim dr1() As DataRow
            Dim varID
            tbl = tblConfigReportType
            strF = "CHARREPORTTYPE = '" & str1 & "'"
            dr = tbl.Select(strF)
            varID = dr(0).Item("id_tblConfigReportType")
            dgv("id_tblConfigReportType", intRow).Value = varID
            dgv("CHARREPORTTYPE", intRow).Value = str1

        End If

    End Sub

    Private Sub cbxRBSFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxRBSFilter.SelectedIndexChanged
        If boolFormLoad Then
            Exit Sub
        End If

        Call RBFilter()

        'set focus to something else
        dgvReportStatements.Focus()


    End Sub

    Private Sub cbxRBSTypeFilter_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxRBSTypeFilter.SelectedIndexChanged
        If boolFormLoad Then
            Exit Sub
        End If

        Call RBFilter()

        'set focus to something else
        dgvReportStatements.Focus()

    End Sub



    Private Sub dgvAnalyticalRunSummary_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvAnalyticalRunSummary.CellBeginEdit

        'if row is the blank row, don't allow edit
        Dim var1

        var1 = NZ(dgvAnalyticalRunSummary("Analyte", e.RowIndex).Value, "")
        If Len(var1) = 0 Then
            e.Cancel = True
        End If

    End Sub

    Private Sub dgQATable_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgQATable.MouseEnter
        'dgQATable.Focus()

        Try
            dgQATable.Focus()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub dgStudies_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgStudies.MouseEnter
        dgStudies.Focus()

    End Sub

    Private Sub dgvAnalyticalRunSummary_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvAnalyticalRunSummary.MouseEnter
        dgvAnalyticalRunSummary.Focus()

    End Sub



    Private Sub dgvContributingPersonnel_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvContributingPersonnel.MouseEnter
        Me.dgvContributingPersonnel.Focus()

    End Sub

    Private Sub dgvDataCompany_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvDataCompany.MouseEnter
        dgvDataCompany.Focus()

    End Sub

    Private Sub dgvReports_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReports.MouseEnter
        dgvReports.Focus()

    End Sub

    Private Sub dgvReportTableConfiguration_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportTableConfiguration.MouseEnter

        dgvReportTableConfiguration.Focus()

    End Sub

    Private Sub dgvSampleReceipt_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSampleReceipt.MouseEnter
        dgvSampleReceipt.Focus()

    End Sub

    Private Sub dgvSampleReceiptWatson_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSampleReceiptWatson.MouseEnter
        dgvSampleReceiptWatson.Focus()

    End Sub

    Private Sub dgvUser_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvUser.MouseEnter
        dgvUser.Focus()

    End Sub

    Private Sub dgvWatsonAnalRef_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvWatsonAnalRef.MouseEnter
        Me.dgvWatsonAnalRef.Focus()

    End Sub

    Private Sub lblProgress_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblProgress.TextChanged

        'Try
        '    Dim frm As New frmAssignSamples
        '    frmAssignSamples.lblProgress.Text = Me.lblProgress.Text
        '    frmAssignSamples.lblProgress.Refresh()
        '    frm.Dispose()
        'Catch ex As Exception

        'End Try

    End Sub

    Private Sub lblProgress_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblProgress.VisibleChanged

        Dim d1, d2

        d1 = lblProgress.Size
        d2 = dgvwStudy.Size
        If d1 = d2 Then 'no need to do anything
            Exit Sub
        End If

        Dim x1, x2
        Dim tp1, tp2
        Dim b1

        Call PositionProgress()


    End Sub

    Private Sub cmdRefreshStatements_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRefreshStatements.Click

        Dim strM As String
        Dim boolA As Boolean = BOOLVIEWWORDTEMPLATE
        If boolA Then
        Else
            strM = "User does not have permission to view Report Templates."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If


        Dim bool As Boolean
        Dim str1 As String
        Dim boolExpand As Boolean = True
        Dim strLbl As String

        Cursor.Current = Cursors.WaitCursor

        str1 = Me.cmdRefreshStatements.Text
        If InStr(1, str1, "Sho&w", CompareMethod.Text) > 0 Then
            bool = True
            If Me.rbEntireReport.Checked Then
                'str1 = "&Hide Reports"
                boolExpand = False
            Else
                str1 = "&Hide Statements"
            End If
        Else
            bool = False
            If Me.rbEntireReport.Checked Then
                str1 = "Sho&w Reports"
                boolExpand = False
            Else
                str1 = "Sho&w Statements"
            End If
        End If

        Me.cmdRefreshStatements.Text = str1

        If boolExpand Then 'expand statements
            Call ExpandRBS(bool)

            Call UpdateWB_RBS()

            'check for toprow
            Dim dgv1 As DataGridView
            Dim dgv2 As DataGridView
            Dim intRow1 As Short
            Dim intRow2 As Short

            dgv1 = Me.dgvReportStatements
            dgv2 = Me.dgvReportStatementWord

            Try
                intRow1 = dgv1.CurrentRow.Index
                dgv1.FirstDisplayedScrollingRowIndex = intRow1

            Catch ex As Exception
                intRow2 = dgv2.CurrentRow.Index
                dgv2.FirstDisplayedScrollingRowIndex = intRow2

            End Try

        Else 'for entire report, show in frmWord

            Dim strpathT As String

            strpathT = CreatexmlHome(Me.dgvReportStatementWord)

            Dim dgv As DataGridView
            Dim intRow As Short

            dgv = Me.dgvReportStatementWord
            If dgv.CurrentRow Is Nothing Then
                intRow = 0
            Else
                intRow = dgv.CurrentRow.Index
            End If
            strLbl = dgv("CHARTITLE", intRow).Value

            Call OpenAFR(strpathT, strLbl, True, False, False, True)


        End If

        Cursor.Current = Cursors.Default

end1:

    End Sub


    Private Sub dgvReportStatements_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportStatements.MouseEnter
        dgvReportStatements.Focus()

    End Sub

    Private Sub dgvReportStatements_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportStatements.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        'If boolHold Then
        '    Exit Sub
        'End If

        Cursor.Current = Cursors.WaitCursor

        Call UpdateWord_dgv()

        Cursor.Current = Cursors.Default

        'Call UpdateWB_RBS()

    End Sub

    Private Sub llblSummaryTable_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblSummaryTable.LinkClicked
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim Count1 As Short

        str1 = "Edit Appendices & Figures"
        int1 = lbxTab1.Items.Count
        For Count1 = 0 To int1 - 1
            str2 = lbxTab1.Items(Count1).ToString
            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                lbxTab1.SelectedIndex = Count1
                Exit For
            End If
        Next

    End Sub


    Private Sub dgvReportStatementWord_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportStatementWord.DoubleClick

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim str1 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim dv As system.data.dataview
        Dim id As Int64

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If Len(Me.cbxStudy.Text) = 0 Then
            Exit Sub
        End If

        dgv1 = Me.dgvReportStatements
        dgv2 = Me.dgvReportStatementWord

        int1 = dgv2.CurrentRow.Index
        str1 = dgv2.Rows.Item(int1).Cells("CHARTITLE").Value
        id = dgv2.Rows.Item(int1).Cells("ID_TBLWORDSTATEMENTS").Value

        int2 = dgv1.CurrentRow.Index
        dv = dgv1.DataSource
        dv(int2).BeginEdit()
        dv(int2).Item("CHARSTATEMENT") = str1
        dv(int2).Item("ID_TBLWORDSTATEMENTS") = id
        dv(int2).EndEdit()


    End Sub

    Private Sub dgvReportStatementWord_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportStatementWord.MouseEnter
        dgvReportStatementWord.Focus()

    End Sub

    Sub StudyDo()

        Dim int1 As Short
        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String

        If boolFormLoad Then
            Exit Sub
        End If

        dgv = Me.dgvwStudy
        If dgv.Rows.Count > 1 And boolNewOracle = False Then
            Exit Sub
        End If

        Call ActivateStudyChange()


        If dgv.CurrentRow Is Nothing Then
        Else
            If dgv.CurrentRow.Displayed Then
            Else
                'ensure selected row is displayed
                dgv.FirstDisplayedScrollingRowIndex = int1
            End If
        End If

        'Me.lblProgress.Visible = False
        'Me.pb1.Visible = False

        'Me.panProgress.Visible = False
        'Me.panProgress.Refresh()


        boolFormLoad = False
        boolFromdgvwStudy = False
        boolcbxExample = False
        dgv.Focus()

    End Sub



    Private Sub dgvwStudy_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvwStudy.MouseEnter
        dgvwStudy.Focus()

    End Sub


    Private Sub dgvwStudy_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvwStudy.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Try
            Me.Button1.Text = Me.dgvwStudy.CurrentRow.Index
        Catch ex As Exception

        End Try

        Call dgvwStudyChange()

end1:

    End Sub

    Sub dgvwStudyChange()

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.rbArchive.Checked Then
            Exit Sub
        End If

        If StudyAllowed() Then
        Else
            Exit Sub
        End If

        Call dgvwStudySelCh()

        'now record this as a silent audit trail entry
        Call RecordStudyOpenAuditTrail()

        'Me.pb1.Visible = False
        'Me.pb2.Visible = False
        'Me.lblProgress.Visible = False

        Me.panProgress.Visible = False
        Me.panProgress.Refresh()

        boolStudyFired = True

        'pesky
        Call SetPanAction()

    End Sub

    Sub dgvwStudySelCh()

        If boolFormLoad Then
            Exit Sub
        End If

        'If Me.rbArchive.Checked Then
        '    Exit Sub
        'End If

        'NOTE: cbxStudy.datasource is linked to dgvwStudy.datasource. Therefore, changes in either selection will change the other.
        'However, if the first row is clicked the first time, then cbxStudy change won't fire

        Dim int1 As Short
        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String
        Dim int2 As Short


        dgv = Me.dgvwStudy

        boolFromdgvwStudy = True
        If dgv.Rows.Count = 0 Then
            Exit Sub
        End If
        int1 = dgvwStudy.CurrentRow.Index 'debug
        cbxStudy.Refresh()

        'str1 = NZ(Me.cbxStudy.Text, "")
        'str2 = NZ(dgv("STUDYNAME", int1).Value, "")
        'If StrComp(str1, str2, CompareMethod.Text) = 0 Then
        'Else
        '    Me.cbxStudy.SelectedIndex = int1
        'End If

        Call ActivateStudyChange()

        'now assign id_tblconfigreporttype
        Call SetReportConfigType()

        If dgv.CurrentRow.Displayed Then
        Else
            'ensure selected row is displayed
            dgv.FirstDisplayedScrollingRowIndex = int1
        End If

        'Me.lblProgress.Visible = False
        'Me.pb1.Visible = False

        Me.panProgress.Visible = False
        Me.panProgress.Refresh()

        boolFormLoad = False
        boolFromdgvwStudy = False
        boolcbxExample = False
        dgv.Focus()

    End Sub

    Private Sub cmdUpdateSummaryInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdUpdateSummaryInfo.Click

        Call UpdateValueSummaryTable()

    End Sub

    Private Sub cbxExampleReport_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxExampleReport.MouseEnter
        cbxExampleReport.ForeColor = Color.Black
    End Sub

    Private Sub cbxExampleReport_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxExampleReport.MouseLeave
        cbxExampleReport.ForeColor = Color.Blue
    End Sub

    Private Sub cbxExampleReport_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxExampleReport.SelectedIndexChanged

        '*****

        Dim strM As String

        Dim bool As Boolean
        bool = False
        Dim int1 As Short

        gdtReportDate = Now
        gboolTableSection = True
        gboolTSDone = True

        gDoPDF = False

        If boolFormLoad Then
            Exit Sub
        End If

        If boolcbxExample Then
            Exit Sub
        End If

        If Me.cbxExampleReport.SelectedIndex = 0 Then
            Exit Sub
        End If

        If BOOLVIEWFINALREPORT Then
        Else
            strM = "User is not allowed to view reports. By extension, this means user also does not have permission to generate a report."
            MsgBox(strM, vbInformation, "Invalid action...")
            Me.cbxExampleReport.SelectedIndex = 0
            Exit Sub
        End If

        If BOOLALLOWREPORTGENERATION Then
        Else
            strM = "User does not have permission to generate a report."
            MsgBox(strM, vbInformation, "Invalid action...")
            Me.cbxExampleReport.SelectedIndex = 0
            Exit Sub
        End If

        tPswd = "" 'document pswd if doc is saved


        'first check default Report Template

        Dim dv As DataView = frmH.dgvReportStatementWord.DataSource
        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String

        '20160324 LEE: Check not needed anymore

        Dim boolHit As Boolean = False

        ''determine if default is available
        'str2 = NZ(frmH.dgvReportStatements("CHARSTATEMENT", 0).Value, "NA")
        '
        'For Count1 = 0 To dv.Count - 1
        '    str1 = dv(Count1).Item("CHARTITLE").ToString
        '    If StrComp(str2, str1, CompareMethod.Text) = 0 Then
        '        boolHit = True
        '        Exit For
        '    End If
        'Next

        'If boolHit Then
        'Else
        '    strM = "The default Report Template for this study" & ChrW(10) & ChrW(10) & str2 & ChrW(10) & ChrW(10)
        '    strM = strM & "Does not exist in the list of available Report Templates." & ChrW(10) & ChrW(10)
        '    'strM = strM & "It is probable that the Report Template was renamed or deleted." & ChrW(10) & ChrW(10)
        '    strM = strM & "It is probable that the Report Template was deactivated." & ChrW(10) & ChrW(10)
        '    strM = strM & "Please select the Report Template Configuration page and assign a new default Report Template for this study."
        '    MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
        '    boolFormLoad = True
        '    GoTo end2
        'End If

        'str1 = NZ(cbxStudy.Text, "")
        'If Len(str1) = 0 Then
        '    bool = True
        '    boolcbxExample = True
        '    GoTo end1
        'End If

        boolcbxExample = True
        boolhit = True

        'recommend saving data
        If Me.cmdEdit.Enabled Then
        Else
            If Me.dgvwStudy.RowCount = 0 Then
                str1 = "Please load a study before generating a report."
            Else
                str1 = "Please save the study before generating a report."
            End If

            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        If AllowPrint() Then
        Else
            GoTo end1
        End If

        intTTot = 0
        intTCur = 0


        'set gboolDisplayAttachments global variable
        Call SetDisplayAttachment()

        'annoying!! must do the following
        Call ViewSections(False)

        Dim intCurrentI As Short = frmH.lbxTab1.SelectedIndex

        Cursor.Current = Cursors.WaitCursor

        'show panprogress
        Me.lblProgress.Text = ""
        Me.pb1.Value = 0
        Me.pb2.Value = 0
        Me.panProgress.Visible = True
        Me.panProgress.Refresh()

        str1 = cbxExampleReport.Text
        int1 = Me.cbxExampleReport.SelectedIndex
        Try
            Select Case int1
                Case 0 '"Prepare a Report..."
                    'do nothing
                Case 2 '"Prepare Entire Report..."
                    If BOOLFINALREPORTLOCKED Then
                        strM = "Final Reports for this study have been locked."
                        strM = strM & ChrW(10) & "User may still prepare tables and sections."
                        MsgBox(strM, vbInformation, "Invalid action...")
                        GoTo end2
                    End If
                    Call ChooseReportWindow()
                    Call PrepareReport()
                Case 4 '"Prepare Only Selected Section..."
                    Call ExampleSection("Home")
                Case 6 '"Prepare Only Report Body Section..."
                    'Call ExampleReportBody(False)
                    '20190222 LEE: should be tru
                    Call ExampleReportBody(True)
                Case 8 '"Prepare Only Report Table Section..."
                    Call ChooseReportWindow()
                    Call ExampleTablesSection()
                    'Case 10 'Prepare Report Body Showing Field Codes 'depricated
                    '    Call ExampleReportBody(True)
            End Select
        Catch ex As Exception
            strM = "There was a problem preparing this report." & ChrW(10) & ChrW(10)
            strM = strM & "Please contact your StudyDoc Administrator." & ChrW(10) & ChrW(10) & ex.Message
            Try
                'Me.lblProgress.Visible = False
                'Me.pb1.Visible = False
                'Me.pb2.Visible = False

                Me.panProgress.Visible = False
                Me.panProgress.Refresh()
            Catch ex1 As Exception

            End Try
            MsgBox(strM, MsgBoxStyle.Information, "Problem...")
        End Try


end1:
        cbxExampleReport.SelectedIndex = 0
        If bool And boolcbxExample Then
            str1 = "A study must be selected."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            boolcbxExample = False
        End If

        'reset variables
        boolDoFormulas = False
        boolDoHyperlinks = True

        Call AssessSampleAssignment()

        Cursor.Current = Cursors.Default

        'now set focus to a different control
        'dgvwStudy.Focus()
        'NDL - commented this out as it was covering the report window
        'with the main window, and I'm not sure why it was there in the 
        'first place.

        gDoPDF = False

end2:
        If boolHit Then
        Else
            cbxExampleReport.SelectedIndex = 0
            boolFormLoad = False
        End If

        cbxExampleReport.SelectedIndex = 0
        boolcbxExample = False

        Try
            frmH.lbxTab1.SelectedIndex = intCurrentI
        Catch ex As Exception

        End Try

        'Me.lblProgress.Visible = False
        'Me.pb1.Visible = False
        'Me.pb2.Visible = False

        Me.panProgress.Visible = False
        Me.panProgress.Refresh()

end3:

    End Sub

    Private Sub cmdAppFig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAppFig.Click

        Call OpenAppFig()

    End Sub

    Sub OpenAppFig()

        Dim strM As String
        Dim boolA As Boolean = BOOLAPPENDICES
        Dim var1

        If boolA Then
        Else
            strM = "User does not have permission to access the 'Edit Appendices & Figures' window."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        Dim frm As New frmAppFigs
        Try
            frm.ShowDialog()
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try




        Refresh()

        Cursor.Current = Cursors.WaitCursor
        frm.Dispose()
        Cursor.Current = Cursors.Default

end1:

    End Sub

    Private Sub cbxAnticoagulant_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxAnticoagulant.SelectedIndexChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dgv As DataGridView
        Dim dgv2 As DataGridView
        Dim dvMVD As System.Data.DataView
        Dim strCol As String
        Dim boolVal As Boolean
        Dim idR As Int64

        'determine if method validation
        boolVal = False
        Dim dgvR As DataGridView
        dgvR = Me.dgvReports
        If dgvR.Rows.Count = 0 Then
        Else
            idR = NZ(dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value, -1)
            If idR > 1 And idR < 1000 Then
                boolVal = True
            Else
                boolVal = False
            End If
        End If

        dgv = dgvMethodValData
        dgv2 = dgvMethValExistingGuWu
        dvMVD = dgvMethValExistingGuWu.DataSource
        int3 = dvMVD.Count
        str1 = cbxAnticoagulant.Text
        dv = dgvMethodValData.DataSource
        Try
            int1 = dv.Count
            int2 = FindRowDV("Anticoagulant/Preservative", dv)

            If boolVal Then
                Try
                    If dgv.Rows.Count = 0 Then
                    Else
                        For Count1 = 0 To int3 - 1
                            strCol = dgv2(0, Count1).Value
                            dv(int2).BeginEdit()
                            dv(int2).Item(strCol) = str1
                            dv(int2).EndEdit()
                        Next
                    End If
                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub rbRBS_Col_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbRBS_Col.CheckedChanged
        If boolFormLoad Then
            Exit Sub
        End If
        Call OrderReportStatementCol()
    End Sub

    Private Sub cbxStudy_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxStudy.Enter

        If boolFormLoad Then
            Exit Sub
        End If

        'select entire text
        'SendKeys.Send("{HOME}")
        'SendKeys.Send("+{END}")

        Me.cbxStudy.SelectAll()

    End Sub

    Private Sub cbxStudy_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxStudy.GotFocus
        If boolFormLoad Then
            Exit Sub
        End If

        If boolQuickFind Then
            boolQuickFind = False
        Else
            'select entire text
            'SendKeys.Send("{HOME}")
            'SendKeys.Send("+{END}")

            Me.cbxStudy.SelectAll()

        End If

    End Sub

    Private Sub cbxStudy_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cbxStudy.KeyDown
        If boolFormLoad Then
            Exit Sub
        End If

        If e.KeyCode = Keys.Enter Then
            Me.dgvwStudy.Focus()
        End If
    End Sub

    Private Sub cbxStudy_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxStudy.LostFocus

        If CInt(txtcbxMDBSelIndex.Text) = -10 Then
            'Exit Sub
        End If
        If boolFormLoad Then
            Exit Sub
        End If

        Dim str1 As String
        str1 = Me.cbxStudy.Text
        If Len(str1) = 0 Then
            Exit Sub
        End If

        boolStudyFired = False

        Me.dgvwStudy.Focus()

        If boolStudyFired Then

        Else
            'Call dgvwStudySelCh()

        End If

        'Call cbxStudyCorrect()

    End Sub

    Private Sub cbxStudy_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxStudy.MouseEnter
        If boolFormLoad Then
            Exit Sub
        End If
        boolQuickFind = True
        cbxStudy.Focus()
        ''select entire text
        'SendKeys.Send("{HOME}")
        'SendKeys.Send("+{END}")

    End Sub

    Private Sub cbxStudy_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxStudy.MouseLeave
        If boolFormLoad Then
            Exit Sub
        End If
        If CInt(Me.txtcbxMDBSelIndex.Text) = -10 Then
            Exit Sub
        End If

        'the next code will cause GuWu to hang if Watson database can't be found
        'Call cbxStudyCorrect()

    End Sub


    Private Sub cmdRBSAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRBSAll.Click
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim Count1 As Short
        Dim strS As String
        Dim dtbl As System.Data.DataTable
        Dim strF As String
        Dim dv1 As System.Data.DataView
        Dim rows1() As DataRow

        boolStopRBS = True

        Cursor.Current = Cursors.WaitCursor

        dgv = frmH.dgvReportStatements
        dv1 = dgv.DataSource
        strF = dv1.RowFilter
        strS = dv1.Sort
        int1 = dgv.Rows.Count
        dtbl = tblReportstatements
        rows1 = dtbl.Select(strF, strS)
        For Count1 = 0 To int1 - 1
            'dv1(Count1).BeginEdit()
            'dv1(Count1).Item("BOOLI") = -1
            'dv1(Count1).EndEdit()
            'dgv("BOOLI", Count1).Value = -1
            rows1(Count1).BeginEdit()
            rows1(Count1).Item("BOOLI") = -1
            rows1(Count1).Item("BOOLINCLUDE") = -1
            rows1(Count1).EndEdit()
        Next
        'dgv.DataSource = dv1

        strS = "intOrder ASC"
        'Dim dv as system.data.dataview = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        'dv.AllowNew = False
        'dv.AllowDelete = False
        'dv.AllowEdit = True
        'dgv.DataSource = dv
        Dim dv As System.Data.DataView
        dv = dgv.DataSource
        dv.RowFilter = strF
        dv.Sort = strS

        Call OrderReportStatementCol() 'pesky

        boolStopRBS = False

        Cursor.Current = Cursors.Default
    End Sub

    Private Sub cmdAnalDetails_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAnalDetails.Click

        Call OpenAnalDetails()

    End Sub

    Sub OpenAnalDetails()

        Dim var1

        Dim strM As String
        Dim boolA As Boolean = BOOLSAMPLEDETAILS
        If boolA Then
        Else
            strM = "User does not have permission to access the 'Sample/QC/Calibr Std Details' window."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If


        var1 = NZ(Me.cbxStudy.Text, "")
        If Len(var1) = 0 Then
            MsgBox("A study must be chosen.", MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        Dim frm As New frmAnalyteDetails

        'If Me.cmdUpdateProject.Enabled Then
        '    MsgBox("This study is not configured within StudyDoc.", MsgBoxStyle.Information, "Invalid action...")
        '    Exit Sub
        'End If

        frm.ShowDialog()
        If frm.boolGoTables Then
            Dim str1 As String
            Dim str2 As String
            Dim Count1 As Short
            Dim int1 As Short

            str1 = "Configure Report Tables"
            int1 = frmH.lbxTab1.Items.Count
            For Count1 = 0 To int1 - 1
                str2 = frmH.lbxTab1.Items(Count1).ToString
                If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                    frmH.lbxTab1.SelectedIndex = Count1
                    'frmH.tab1.SelectedIndex = Count1
                    Exit For
                End If
            Next

        End If

        Me.Refresh()

        frm.Close()

end1:

    End Sub

    Sub ExecuteFilter()

        Dim dgv As DataGridView
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strF As String
        Dim strF1 As String
        Dim strS As String
        Dim dv As System.Data.DataView
        Dim dtbl As System.Data.DataTable
        Dim strI As String
        Dim boolStop As Boolean

        If boolFormLoad Then
            Exit Sub
        End If

        boolStop = False
        str1 = Me.cbxFilterStudy.Text
        str2 = NZ(Me.txtFilterStudy.Text, "")
        strI = Me.txtFilterIndex.Text


        dgv = Me.dgvwStudy

        Try
            dv = dgv.DataSource
        Catch ex As Exception
            'dgv probably has table as datasource
            dtbl = dgv.DataSource
            dv = dtbl.DefaultView
        End Try

        If StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
            If StrComp(strI, "Not Filtered", CompareMethod.Text) = 0 Then
                boolStop = True
            Else
                Me.txtFilterIndex.Text = "Not Filtered"
            End If
            strF = ""
        ElseIf Len(str2) = 0 Then

            'Exit Sub
            If StrComp(strI, "Not Filtered", CompareMethod.Text) = 0 Then
                boolStop = True
                GoTo end1
            Else
                Me.txtFilterIndex.Text = "Not Filtered"
            End If

        Else
            Select Case str1
                Case "Project ID"
                    str3 = "PROJECTIDTEXT"
                    'Case "Species"
                    '    str3 = "SPECIES"
                    '20190124 LEE
                Case "Study Type"
                    str3 = "CHARREPORTTYPE"
                Case "Study Name"
                    str3 = "STUDYNAME"
                    'Case "Study Number"
                    '    str3 = "STUDYNUMBER"
                Case "Study Title"
                    str3 = "STUDYTITLE"
            End Select
            If StrComp(strI, "Filtered", CompareMethod.Text) = 0 Then
                strF1 = dv.RowFilter
                'strF = "(" & strF1 & ") AND (" & str3 & " LIKE '" & str2 & "')"
                strF = str3 & " LIKE '" & str2 & "*'"
            Else
                strF = str3 & " LIKE '" & str2 & "*'"
            End If
            Me.txtFilterIndex.Text = "Filtered"
        End If


        If boolStop Then
        Else
            Me.Cursor.Current = Cursors.WaitCursor
            Call ResetStudyRecord() 'Me.txtcbxMDBSelIndex.Text = -10
            boolFormLoad = True
            dv.RowFilter = strF
            boolFormLoad = False

            'select first row
            'NO! Don't select first row
            'If dv.Count = 0 Then
            'Else
            '    Me.txtcbxMDBSelIndex.Text = 1
            '    Me.txtcbxMDBSelIndex.Refresh()
            '    Me.dgvwStudy.CurrentCell = Me.dgvwStudy.Rows.Item(0).Cells("STUDYNAME")
            '    Me.txtcbxMDBSelIndex.Text = 0
            'End If
            Me.dgvwStudy.AutoResizeRows()

            'clear data
            Call ClearData()
            Dim boolFL As Boolean = boolFormLoad
            boolFormLoad = True
            dgv.ClearSelection()
            boolFormLoad = boolFL
            Me.Cursor.Current = Cursors.Default

            Me.txtFilterStudy.Focus()

        End If

        Call SetStudyCount() '20190130 LEE:

end1:

    End Sub

    Private Sub cmdFilterStudy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdFilterStudy.Click

        Call ExecuteFilter()

    End Sub


    Private Sub cbxFilterStudy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxFilterStudy.SelectedIndexChanged

        Dim str1 As String

        If boolFormLoad Then
            Exit Sub
        End If

        Dim var1, var2
        var1 = Me.cbxFilterStudy.SelectedIndex

        str1 = Me.cbxFilterStudy.Text
        If StrComp(str1, "[None]", CompareMethod.Text) = 0 Then
            Me.txtFilterStudy.Text = ""
            Call ExecuteFilter()
        Else
            Me.txtFilterStudy.Text = ""
            Me.txtFilterStudy.Select()
        End If


    End Sub

    Private Sub txtFilterStudy_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFilterStudy.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call ExecuteFilter()
        End If
    End Sub

    Private Sub dgvReportTableConfiguration_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportTableConfiguration.RowLeave
        If boolFormLoad Then
            Exit Sub
        End If

        Dim dgv As DataGridView
        Dim str1 As String

        dgv = dgvReportTableConfiguration
        str1 = dgv.Columns.Item(e.ColumnIndex).Name
        If StrComp(str1, "INTORDER", CompareMethod.Text) = 0 Then
            dgv.AutoResizeRows()
        End If

    End Sub

    Private Sub cmdViewAnalRuns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdViewAnalRuns.Click

        If BOOLREPORTTABLECONFIGURATION Then
        Else
            MsgBox("User not allowed to execute items in Table Configuration window.", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        'Call ViewAnalRuns()
        Call OpenAssignedSamples(True)


    End Sub


    Private Sub rbOracle_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbOracle.CheckedChanged

        If Me.rbArchive.Checked Then
            Me.lblWatsonData.Text = "Archived MDB"
        Else
            Me.lblWatsonData.Text = "Oracle"
        End If

        If boolHold Then
            boolHold = False
            Exit Sub
        End If

        Call DataSourceChecked()

    End Sub

    Private Sub rbArchive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbArchive.CheckedChanged

        If Me.rbArchive.Checked Then
            Me.lblWatsonData.Text = "Archived MDB"
        Else
            Me.lblWatsonData.Text = "Oracle"
        End If
        If boolHold Then
            boolHold = False
            Exit Sub
        End If

        Call DataSourceChecked()

        Call ModMethChoice()


    End Sub

    Sub DataSourceChecked()

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView

        If boolFormLoad Then
            Exit Sub
        End If

        pArchivePath = ""

        boolHold = True

        dgv = Me.dgvwStudy

        Try
            If frmH.rbOracle.Checked Then
                Me.panFilterStudy.Enabled = True
                constrCur = constrWatson 'Oracle connection string
                boolANSI = boolTempANSI
                'dv = New DataView(tblwSTUDY)
                'dv = New DataView(Me.dgvwStudy.DataSource)
                boolArchiveSource = False
            Else
                Me.panFilterStudy.Enabled = False
                'frmH.cmdBrowse.Visible = True
                'constrCur = constrA 'constra is set in DoArchiveMDB
                boolTempANSI = boolANSI
                boolANSI = True
                'dv = New DataView(tblASTUDY)
                'dv = New DataView(Me.dgvwStudy.DataSource)
                boolArchiveSource = True
            End If

            Dim int1 As Short
            Dim int2 As Short

            int1 = tblwSTUDY.Rows.Count
            int2 = tblASTUDY.Rows.Count

            'dv.AllowDelete = False
            'dv.AllowEdit = False
            'dv.AllowNew = False

            'assign this datasource to cbxstudy
            'set selection to nothing 
            cbxStudy.DataSource = Me.dgvwStudy.DataSource ' dv
            Me.cbxStudy.DisplayMember = "STUDYNAME"
            'Me.cbxStudy.ValueMember = "ID_TBLSTUDIES"
            Me.cbxStudy.SelectedIndex = -1
            Me.cbxStudy.AutoCompleteMode = AutoCompleteMode.SuggestAppend

            If Me.cbxStudy.Items.Count = 0 Then
                Me.cbxStudy.Text = ""
            End If

            '20141027_Takeout
            'Try
            '    dgv.DataSource = dv

            'Catch ex As Exception
            '    dv = New DataView(tblASTUDY)
            '    dv.AllowDelete = False
            '    dv.AllowEdit = False
            '    dv.AllowNew = False

            '    cbxStudy.DataSource = dv
            '    Me.cbxStudy.DisplayMember = "STUDYNAME"
            '    'Me.cbxStudy.ValueMember = "ID_TBLSTUDIES"
            '    Me.cbxStudy.SelectedIndex = -1
            '    Me.cbxStudy.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            '    If Me.cbxStudy.Items.Count = 0 Then
            '        Me.cbxStudy.Text = ""
            '    End If

            '    boolFormLoad = True
            '    dgv.DataSource = dv
            '    boolFormLoad = False

            'End Try
            '20141027_Takeout End

            dgv.AutoResizeRows()

        Catch ex As Exception

        End Try

        Call ClearData()

        intOTables = 0

        If frmH.rbOracle.Checked Then
            frmH.txtcbxMDBSelIndex.Text = -10
            Call GetStudyInfo()
            Call SetDecs()
            Call Set_idtblReports()

            Call MethodValColor()

        End If



    End Sub

    Sub OpenArchiveMDB(boolAccess As Boolean)

        'Call PositionProgress()

        Dim strPath As String
        Dim strFilter As String
        Dim strFileName As String
        Dim str1 As String
        Dim str2 As String
        Dim strP As String
        Dim boolGo As Boolean
        Dim var1, var2
        Dim strF As String

        'Me.txtArchivePath.Text = "C:\Labintegrity\StudyDoc\ArchivedMDBs\"

        If boolAccess Then

            str1 = "C:\LabIntegrity\StudyDoc\ArchivedMDBs\"
            strP = Me.txtArchivePath.Text

            'get default
            strF = "ID_TBLCONFIGURATION = 32"
            Dim rows() As DataRow = tblConfiguration.Select(strF)
            str2 = rows(0).Item("CHARCONFIGVALUE")

            If Directory.Exists(strP) Then
                strPath = strP
            Else
                If Directory.Exists(str2) Then
                    strPath = str2
                Else
                    If Directory.Exists(str1) Then
                        strPath = str1
                    Else
                        strPath = "C:\"
                    End If
                End If
            End If

            strFilter = ".MDB files (*.MDB*)|*.MDB"
            strFileName = "*.MDB"

            str1 = ReturnDirectoryBrowse(True, strPath, strFilter, strFileName, True) 'true = looking for file

            If Len(str1) = 0 Then
                boolGo = False
                GoTo end1
            End If

            strPath = Path.GetDirectoryName(str1)
            Me.txtArchivePath.Text = strPath

        Else
            str1 = ""
        End If

        boolGo = DoArchiveMDB(str1, boolAccess)

        If boolGo Then
            boolRefresh = True

            Call StudyDo()


            Try

            Catch ex As Exception
                var1 = ex.Message
                var2 = var1

            End Try

            boolRefresh = False
        End If

        Call SetReportConfigType()


end1:

        'Me.pb1.Visible = False
        'Me.pb2.Visible = False
        'Me.lblProgress.Visible = False

        Me.panProgress.Visible = False
        Me.panProgress.Refresh()

    End Sub

    Function DoArchiveMDB(ByVal strPath As String, boolAccess As Boolean) As Boolean

        Dim cnn As New ADODB.Connection
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim rs As New ADODB.Recordset
        Dim int1 As Integer
        Dim strM As String
        Dim var1

        If boolAccess Then
            constrA = "Provider=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & strPath
            constrCur = constrA
            conAccess97 = "Provider=Microsoft.Jet.OLEDB.3.5;Data Source="
        Else
            constrCur = constrWatson
        End If

        pArchivePath = ""

        Try
            cnn.Open(constrCur)

        Catch ex As Exception
            DoArchiveMDB = False

            If boolAccess Then
                strM = "Access file could not be loaded." & ChrW(10) & ex.Message
            Else
                strM = "Watson Oracle database could not be accessed." & ChrW(10) & ex.Message
            End If
            MsgBox(strM, MsgBoxStyle.Information, "Stopping database load...")
            GoTo end2
        End Try


        If boolAccess Then
            str1 = "SELECT PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME, STUDY.STUDYNUMBER, CONFIGSPECIES.SPECIES, STUDY.STUDYTITLE, PROJECT.PROJECTID, STUDY.STUDYID, STUDY.SPECIESID "
            str2 = "FROM (PROJECT INNER JOIN STUDY ON PROJECT.PROJECTID = STUDY.PROJECTID) LEFT JOIN CONFIGSPECIES ON STUDY.SPECIESID = CONFIGSPECIES.SPECIESID "
            str3 = "ORDER BY PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME;"
            str4 = ""


            'str1 = "SELECT PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME, STUDY.STUDYNUMBER, STUDY.STUDYTITLE, PROJECT.PROJECTID, STUDY.STUDYID, STUDY.SPECIESID "
            'str2 = "FROM PROJECT INNER JOIN STUDY ON PROJECT.PROJECTID = STUDY.PROJECTID "
            'str3 = "ORDER BY PROJECT.PROJECTIDTEXT, STUDY.STUDYNAME; "
            'str4 = ""

        Else
            'if Oracle, this SQL must contain a filter for study ID
            'NO! Need everything
            'str1 = "SELECT " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".STUDY.STUDYNUMBER, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".STUDY.STUDYTITLE, " & strSchema & ".PROJECT.PROJECTID, " & strSchema & ".STUDY.STUDYID, " & strSchema & ".STUDY.SPECIESID "
            'str2 = "FROM (" & strSchema & ".PROJECT INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".PROJECT.PROJECTID = " & strSchema & ".STUDY.PROJECTID) LEFT JOIN " & strSchema & ".CONFIGSPECIES ON " & strSchema & ".STUDY.SPECIESID = " & strSchema & ".CONFIGSPECIES.SPECIESID "
            'str4 = "ORDER BY " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".STUDY.STUDYNAME;"

            'str1 = "SELECT " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".STUDY.STUDYNUMBER, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".STUDY.STUDYTITLE, " & strSchema & ".PROJECT.PROJECTID, " & strSchema & ".STUDY.STUDYID, " & strSchema & ".STUDY.SPECIESID "
            'str2 = "FROM " & strSchema & ".CONFIGSPECIES INNER JOIN (" & strSchema & ".PROJECT INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".PROJECT.PROJECTID = " & strSchema & ".STUDY.PROJECTID) ON " & strSchema & ".CONFIGSPECIES.SPECIESID = " & strSchema & ".STUDY.SPECIESID "
            'str3 = "ORDER BY UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".STUDY.STUDYNAME);"
            'str4 = ""

            '20170901 LEE: modified joins to return all records
            str1 = "SELECT " & strSchema & ".STUDY.STUDYNAME, " & strSchema & ".PROJECT.PROJECTIDTEXT, " & strSchema & ".STUDY.STUDYNUMBER, " & strSchema & ".CONFIGSPECIES.SPECIES, " & strSchema & ".STUDY.STUDYTITLE, " & strSchema & ".PROJECT.PROJECTID, " & strSchema & ".STUDY.STUDYID, " & strSchema & ".STUDY.SPECIESID  "
            str2 = "FROM (" & strSchema & ".PROJECT INNER JOIN " & strSchema & ".STUDY ON " & strSchema & ".PROJECT.PROJECTID = " & strSchema & ".STUDY.PROJECTID) LEFT JOIN " & strSchema & ".CONFIGSPECIES ON " & strSchema & ".STUDY.SPECIESID = " & strSchema & ".CONFIGSPECIES.SPECIESID  "
            str3 = "ORDER BY UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".PROJECT.PROJECTIDTEXT), UPPER(" & strSchema & ".STUDY.STUDYNAME);"
            str4 = ""

        End If

        strSQL = str1 & str2 & str3 & str4
        'Console.WriteLine(strSQL)

        DoArchiveMDB = True

        Try
            rs.CursorLocation = CursorLocationEnum.adUseClient
            Try
                'rs.Open(strSQL, wcn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                rs.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly) ' ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            Catch ex1 As Exception

                cnn.Close()

                If boolAccess Then
                    constrA = conAccess97 & strPath
                    constrCur = constrA
                    Try
                        cnn.Open(constrCur)
                    Catch ex3 As Exception
                        If boolAccess Then
                            strM = "Access file could not be loaded." & ChrW(10) & ex3.Message
                        Else
                            strM = "Watson Oracle database could not be accessed." & ChrW(10) & ex3.Message
                        End If

                        MsgBox(strM, MsgBoxStyle.Information, "Stopping database load...")
                        DoArchiveMDB = False
                        GoTo end2
                    End Try

                    Try
                        'rs.Open(strSQL, wcn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        rs.Open(strSQL, cnn, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly) ' ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    Catch ex2 As Exception
                        DoArchiveMDB = False
                        MsgBox("Cannot open .mdb", MsgBoxStyle.Information, "Stopping database opening...")
                        GoTo end2
                    End Try

                Else

                End If


            End Try

            pArchivePath = strPath

            rs.ActiveConnection = Nothing
            'int1 = rs.RecordCount
            int1 = rs.RecordCount

            'added new

            tblwSTUDY.Clear()
            tblwSTUDY.AcceptChanges()
            tblwSTUDY.BeginLoadData()
            daDoPr.Fill(tblwSTUDY, rs)
            tblwSTUDY.EndLoadData()

            '20190124 LEE:
            'add tblReports.CHARREPORTTYPE
            If tblwSTUDY.Columns.Contains("CHARREPORTTYPE") Then
            Else
                Try
                    Dim col1 As New DataColumn
                    col1.ColumnName = "CHARREPORTTYPE"
                    col1.Caption = "Study Type"
                    col1.DataType = System.Type.GetType("System.String")
                    tblwSTUDY.Columns.Add(col1)
                Catch ex As Exception
                    var1 = var1
                End Try
            End If

            Dim intAAA As Int64 'debug
            intAAA = tblwSTUDY.Rows.Count
            intAAA = intAAA

            'end added new

            tblASTUDY.Clear()
            tblASTUDY.BeginLoadData()
            daDoPr.Fill(tblASTUDY, rs)
            tblASTUDY.EndLoadData()

            Call ConfigStudyTable(True, True)

            rs.Close()

            cnn.Close()
        Catch ex As Exception
            DoArchiveMDB = False
            MsgBox("Cannot open .mdb", MsgBoxStyle.Information, "Stopping database opening...")
            GoTo end2
        End Try



end2:

        rs = Nothing

end1:
        cnn = Nothing

        'set connection string



    End Function

    Private Sub cbxStudy_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxStudy.SelectedIndexChanged

        If boolFormLoad Then
            Exit Sub
        End If

        'If ChangeStudy() Then
        'Else
        '    Exit Sub
        'End If

        Dim dgv As DataGridView
        dgv = Me.dgvwStudy

        If Me.rbArchive.Checked Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled = False Then 'this means this is the first selection of this session, so continue
            If dgv.RowCount = 0 Then
                Exit Sub
            Else
                'change row in dgvwStudy
                'Call dgvwStudySelCh()
            End If
        Else
            Exit Sub
        End If

        Exit Sub

        Dim int1 As Integer
        Dim int2 As Integer
        Dim Count1 As Integer
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim dv As System.Data.DataView
        Dim rows() As DataRow
        Dim bool As Boolean

        Cursor.Current = Cursors.WaitCursor


        dv = dgv.DataSource

        str1 = NZ(Me.cbxStudy.Text, "zzzzzzz")
        Dim tbl As System.Data.DataTable = dv.ToTable()
        strF = "STUDYNAME = '" & str1 & "'"
        rows = tbl.Select(strF)
        If rows.Length = 0 Then
            GoTo end1
        End If
        'find row in dgv
        int1 = tbl.Rows.Count
        bool = False
        For Count1 = 0 To int1 - 1
            str2 = tbl.Rows.Item(Count1).Item("STUDYNAME")
            If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                bool = True
                Exit For
            End If

        Next
        If bool Then
            Try
                dgv.CurrentCell = dgv.Rows.Item(Count1).Cells("STUDYNAME")
                dgv.Rows.Item(Count1).Selected = True

            Catch ex As Exception

            End Try

        End If

end1:

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmdResize_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdResize.Click
        If Me.dgvReportTableConfiguration.Rows.Count = 0 Then
            Exit Sub
        End If
        Me.dgvReportTableConfiguration.AutoResizeRows()

    End Sub

    Private Sub dgvReportTableConfiguration_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportTableConfiguration.SelectionChanged

        If boolLoad Then
            Exit Sub
        End If
        If boolFromRTC Then
            Exit Sub
        End If


        'Update graphic in all edit modes
        If (chkTableGraphicExamples.Checked = True) Then
            Call setTableGraphicExample()
        End If

        If cmdEdit.Enabled Then
            Exit Sub
        End If

        Call UpdateWord_dgv()

    End Sub

    Private Sub setTableGraphicExample()

        Dim dgv As DataGridView = Me.dgvReportTableConfiguration
        Dim intRow As Short
        Dim intTableNum As Long
        Dim strLabel As String

        If (IsNothing(dgv.CurrentRow)) Then
        Else
            intRow = dgv.CurrentRow.Index
            intTableNum = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
            strLabel = dgv("CHARHEADINGTEXT", intRow).Value

            'Set Label correctly
            Me.lblTableGraphicExamplesText.Text = "<no example available>"

            'Then set graphic
            pbxTableGraphicExamples.Visible = True

            'Set relative Path
            Dim strDirPath As String
            strDirPath = "C:\LabIntegrity\StudyDoc\Figs\"
            Select Case intTableNum
                Case 1
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "1-SummaryOfAnalyticalRuns.png"
                Case 2
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "2-SummaryOfRegressionConstants.png"
                Case 3
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "3-SummaryOfBackCalculatedCalibrationStandards.png"
                Case 4
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "4-SummaryOfInterpolatedQCStandard Concentrations.png"
                Case 5
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "5-SummaryOfConcentrationsInSamples.png"
                Case 6
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "6-SummaryOfReassayedSamples.png"
                Case 7
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "7-SummaryOfRepeatSamples.png"
                Case 11
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "11-SummaryOfInterpolatedQCStdConcIntra-andInter-RunPrecision.png"
                Case 12
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "12-InterpolatedDilutionQCConcentrations.png"
                Case 13
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "13-SummaryOfCombinedRecovery.png"
                Case 14
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "14-SummaryOfTrueRecovery.png"
                Case 15
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "15-SummaryOfSuppressionEnhancement.png"
                Case 17
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "17-SummaryOfInterpolatedUniqueQCLowForMatrixEffects.png"
                Case 18
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "18-SummaryOfStabilityInMatrix.png"
                Case 19
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "19-SummaryOfFreezeThawStabilityInMatrix.png"
                Case 21
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "21-FinalExtractStabilityOfInterpolatedQCStdConcentrations.png"
                Case 22
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "22-StockSolutionStabilityAssessment.png"
                Case 29
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "29-Long-TermQCStdStorageStability.png"
                Case 31
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "31-AdHocQCStabilityTable.png"
                Case 32
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "32-AdHocQCStabilityComparisonTable.png"
                Case 34
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "34-SelectivityInIndividualLotsTable.png"
                Case 35
                    pbxTableGraphicExamples.ImageLocation = strDirPath & "35-CarryoverInIndividualLotsTable.png"
                Case Else
                    pbxTableGraphicExamples.Visible = False
            End Select

            'since this routine is called before pbx is visible, must set before IF statement
            Me.lblTableGraphicExamplesText.Text = strLabel

            If (pbxTableGraphicExamples.Visible) Then
                Try
                    pbxTableGraphicExamples.Load()
                Catch ex As Exception
                    'must clear select
                    Me.lblTableGraphicExamplesText.Text = "<no example available>"
                    Me.pbxTableGraphicExamples.ImageLocation = ""
                    'MsgBox(ex.Message)
                End Try
            End If

        End If

    End Sub

    Private Sub cmdViewAnalyticalRuns1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdViewAnalyticalRuns1.Click

        'Call ViewAnalRuns()
        Call OpenAssignedSamples(True)

    End Sub

    Private Sub cmdCreateReportTitle2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCreateReportTitle2.Click

        Call CreateReportTitle()

    End Sub

    Private Sub dgvReportTableConfiguration_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportTableConfiguration.CellDoubleClick

        Dim str1 As String
        Dim str2 As String
        Dim intCol As Short
        Dim intRow As Short
        Dim int1 As Short
        Dim var1, var2, var3
        Dim strCol As String
        Dim dgv As DataGridView
        Dim intColTable As Short
        Dim Count1 As Integer

        If frmH.cmdEdit.Enabled Then
            Exit Sub
        End If

        Exit Sub


        dgv = frmH.dgvReportTableConfiguration

        intCol = e.ColumnIndex
        intRow = e.RowIndex
        strCol = dgv.Columns(intCol).HeaderText
        If InStr(1, strCol, "Period Temp", CompareMethod.Text) > 0 Then 'continue
        Else
            Exit Sub
        End If

        int1 = dgv.Columns.Count
        intColTable = -1
        'For Count1 = 0 To int1 - 1
        '    str1 = dgv.Columns(Count1).HeaderText
        '    If StrComp(str1, "Table Title", CompareMethod.Text) = 0 Then
        '        intColTable = Count1
        '        Exit For
        '    End If
        'Next

        For Count1 = 0 To int1 - 1
            str1 = dgv.Columns(Count1).HeaderText
            If StrComp(str1, "Table Title", CompareMethod.Text) = 0 Then
                intColTable = Count1
                Exit For
            End If
        Next

        If intColTable = -1 Then
            Exit Sub
        End If



        str1 = dgv.Item(intColTable, intRow).Value
        Dim frm As New frmPeriodTemp
        If InStr(1, str1, "[Period Temp]", CompareMethod.Text) > 0 Then
            frm.panCycles.Enabled = False
            frm.panTP.Enabled = True
            frm.txtTP.Focus()
        ElseIf InStr(1, str1, "[#Cycles]", CompareMethod.Text) > 0 Then
            frm.panCycles.Enabled = True
            frm.panTP.Enabled = False
            frm.txtCycles.Focus()
        Else
            GoTo end1
        End If
        frm.ShowDialog()

        If frm.boolCancel Then

        Else
            If InStr(1, str1, "[Period Temp]", CompareMethod.Text) > 0 Then
                var1 = frm.txtTP.Text
                var2 = frm.txtTF.Text
                var3 = frm.txtTemp.Text

                str1 = NZ(var1, "0") & " " & NZ(var2, "Days") & " " & NZ(var3, "Room Temp")


            ElseIf InStr(1, str1, "[#Cycles]", CompareMethod.Text) > 0 Then
                var1 = frm.txtCycles.Text

                str1 = "(" & NZ(var1, "0") & " Cycles)"

            End If
            dgv.Item(intCol, intRow).Value = str1
        End If

end1:

        frm.Dispose()


    End Sub

    Private Sub dgvMethodValData_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMethodValData.CellClick

        Dim str1 As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim int2 As Long
        Dim var1, var2
        Dim str2 As String
        Dim Count1 As Short
        Dim strM As String

        Dim boolVal As Boolean
        boolVal = False
        Dim dgvR As DataGridView
        Dim idR As Int64
        Dim intCol As Short

        dgv = Me.dgvMethodValData
        int1 = e.RowIndex
        int2 = e.ColumnIndex
        If e.RowIndex < 0 Then
            Exit Sub
        End If

        idR = 1 '20190213 LEE:
        '20190213 LEE:
        'This logic isn't correct.
        'It should only apply to method validation studies: idR = 2
        dgvR = Me.dgvReports
        If dgvR.Rows.Count = 0 Then
        Else
            idR = dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value
            If idR = 2 Then
                boolVal = True
            Else
                boolVal = False
            End If
        End If

        If int2 = 0 Then

            Try
                'dgvR = Me.dgvReports
                'If dgvR.Rows.Count = 0 Then
                'Else
                '    idR = dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value
                '    If idR > 1 And idR < 1000 Then
                '        boolVal = True
                '    Else
                '        boolVal = False
                '    End If
                'End If

                '20190213 LEE:
                'This logic isn't correct.
                'It should only apply to method validation studies: idR = 2
                If idR = 2 Then
                    boolVal = True
                Else
                    boolVal = False
                End If

                If boolVal Then 'check for readonly
                    var1 = dgv(0, e.RowIndex).Value

                    strM = ""
                    Select Case var1
                        Case "Validation Corporate Study/Project Number"
                            strM = "To change '" & var1 & "', change the 'Corporate Study/Project Number' text box on the 'Add/Edit Top Level Data' page."
                        Case "Validation Protocol Number"
                            strM = "To change '" & var1 & "', change the 'Protocol Number' text box on the 'Add/Edit Top Level Data' page."
                        Case "Validation Report Title"
                            strM = "To change '" & var1 & "', change the 'Report Title' text box of the 'Configured Reports' table on the 'Choose Study & Report' page."
                        Case "Validation Report Number"
                            strM = "To change '" & var1 & "', change the 'Report Number' text box of the 'Configured Reports' table on the  'Choose Study & Report' page."
                            'Case "Analytical Method Type"
                            '    strM = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."
                        Case "Assay Technique" '20190212 LEE:
                            strM = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."

                            '20190220 LEE: Add some more warnings for stability
                        Case "Freeze/Thaw Stability", "Bench-top Stability", "Process Stability", "Reinjection Stability", "Batch Reinjection Stability", "Long-term Storage Stability", "Whole Blood Stability", "Stock Solution Stability", "Spiking Solution Stability", "Autosampler Stability" '20190212 LEE: 
                            strM = "To change '" & var1 & "', corresponding Stability Conditions Summary cell in the Advanced Table Configuration window - Stability Tab."
                        Case "Maximum # of Freeze/thaw Cycles" '20190212 LEE: 
                            strM = "To change '" & var1 & "', corresponding [#Cylces] Information cell in the Advanced Table Configuration window - Stability Tab."
                        Case "Anticoagulant/Preservative" '20190220 LEE: 
                            strM = "To change '" & var1 & "', change the 'Anticoagulant' dropdown box on the 'Add/Edit Top Level Data' page."
                    End Select

                    If Len(strM) = 0 Then
                    Else
                        MsgBox(strM, vbInformation, "Invalid action...")
                    End If


                End If
            Catch ex As Exception

            End Try

            Exit Sub
        End If

        '20190219 LEE: Moved this out of an IF-THEN. Should be run all the time
        Try

            'dgvR = Me.dgvReports
            'If dgvR.Rows.Count = 0 Then
            'Else
            '    idR = dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value
            '    If idR > 1 And idR < 1000 Then
            '        boolVal = True
            '    Else
            '        boolVal = False
            '    End If
            'End If

            If boolVal Then 'check for readonly
                var1 = dgv(0, e.RowIndex).Value
                strM = ""
                Select Case var1
                    Case "Validation Corporate Study/Project Number"
                        strM = "To change '" & var1 & "', change the 'Corporate Study/Project Number' text box on the 'Add/Edit Top Level Data' page."
                    Case "Validation Protocol Number"
                        strM = "To change '" & var1 & "', change the 'Protocol Number' text box on the 'Add/Edit Top Level Data' page."
                    Case "Validation Report Title"
                        strM = "To change '" & var1 & "', change the 'Report Title' text box of the 'Configured Reports' table on the 'Choose Study & Report' page."
                    Case "Validation Report Number"
                        strM = "To change '" & var1 & "', change the 'Report Number' text box of the 'Configured Reports' table on the  'Choose Study & Report' page."
                        'Case "Analytical Method Type"
                        '    strM = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."
                    Case "Assay Technique" '20190212 LEE:
                        strM = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."

                        '20190220 LEE: Add some more warnings for stability
                    Case "Freeze/Thaw Stability", "Bench-top Stability", "Process Stability", "Reinjection Stability", "Batch Reinjection Stability", "Long-term Storage Stability", "Whole Blood Stability", "Stock Solution Stability", "Spiking Solution Stability", "Autosampler Stability" '20190212 LEE: 
                        strM = "To change '" & var1 & "', corresponding Stability Conditions Summary cell in the Advanced Table Configuration window - Stability Tab."
                    Case "Maximum # of Freeze/thaw Cycles" '20190212 LEE: 
                        strM = "To change '" & var1 & "', corresponding [#Cylces] Information cell in the Advanced Table Configuration window - Stability Tab."
                    Case "Anticoagulant/Preservative" '20190220 LEE: 
                        strM = "To change '" & var1 & "', change the 'Anticoagulant' dropdown box on the 'Add/Edit Top Level Data' page."
                End Select

                If Len(strM) = 0 Then
                Else
                    MsgBox(strM, vbInformation, "Invalid action...")
                    Exit Sub
                End If

            End If
        Catch ex As Exception

        End Try

        '****
        If boolMethodValReadOnly(e.ColumnIndex) Then

            strM = "A Method Validation study has been assigned to this Analyte." & ChrW(10) & ChrW(10) & "Therefore, the table entries associated with this Analyte are read-only."
            'strM = strM & ChrW(10) & ChrW(10) & "Hit the [ESC] key to restore old value."
            If e.ColumnIndex < 0 Then

            Else
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            End If
            'Me.dgvMethodValData.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvMethodValData.Item(e.ColumnIndex, e.RowIndex).Value
            Try
                Me.dgvMethodValData.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvMethodValData.Item(e.ColumnIndex, e.RowIndex).Value

            Catch ex As Exception

            End Try
            GoTo end1
        End If
        '****

        var1 = NZ(dgv(0, e.RowIndex).Value, "")
        If Len(var1) = 0 Then 'give up
        Else
            If StrComp(var1, "Extraction Procedure Description", CompareMethod.Text) = 0 Then
                'determine if cell ia dropdownbox
                str2 = NZ(dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).EditType.FullName, "")
                'If InStr(1, str2, "combobox", CompareMethod.Text) > 0 Then
                'Else

                Try
                    Dim cbx As New DataGridViewComboBoxCell
                    cbx = cbxxAssayDescr.Clone
                    'cbx.DropDownWidth = dgv.Columns(e.ColumnIndex).Width * 3
                    cbx.DropDownWidth = dgv.Width * 0.5
                    var1 = dgv.Columns.Item(e.ColumnIndex).Width
                    var2 = var1 * 2
                    'cbx.DropDownWidth = var2
                    'if data doesn't exist in dropdown list
                    'data error will be called that inserts unlisted value into dropdown box
                    dgv(e.ColumnIndex, e.RowIndex) = cbx

                Catch ex As Exception

                End Try

                'End If

                'ElseIf StrComp(var1, "Analytical Method Type", CompareMethod.Text) = 0 Then '20190212 LEE: deprecated
            ElseIf StrComp(var1, "Assay Technique", CompareMethod.Text) = 0 Then '20190212 LEE


                Try
                    Dim cbx As New DataGridViewComboBoxCell
                    cbx = cbxxAnalMethType.Clone

                    cbx.DropDownWidth = dgv.Width * 0.5
                    var1 = dgv.Columns.Item(e.ColumnIndex).Width
                    var2 = var1 * 2
                    'cbx.DropDownWidth = var2
                    'if data doesn't exist in dropdown list
                    'data error will be called that inserts unlisted value into dropdown box
                    dgv(e.ColumnIndex, e.RowIndex) = cbx

                    'cbxxAnalMethType.Items.Clear()
                    'For Count1 = 0 To frmH.cbxAssayTechniqueAcronym.Items.Count - 1
                    '    str1 = frmH.cbxAssayTechniqueAcronym.Items(Count1)
                    '    cbxxAnalMethType.Items.Add(str1)
                    'Next
                    'cbxxAnalMethType.AutoComplete = True
                    'cbxxAnalMethType.MaxDropDownItems = 20
                    'cbxxAnalMethType.Sorted = True
                    'cbxxAnalMethType.DisplayStyleForCurrentCellOnly = True
                    ''cbxxAnalMethType.DropDownWidth = cbxxCPRole.DropDownWidth * 1.5
                    'cbxxAnalMethType.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton


                    'Dim cbx As New DataGridViewComboBoxCell
                    'cbx = cbxxAnalMethType.Clone

                    'cbx.DropDownWidth = dgv.Columns(e.ColumnIndex).Width * 3
                    'var1 = dgv.Columns.Item(e.ColumnIndex).Width
                    'var2 = var1 * 2
                    'cbx.DropDownWidth = var2
                    'if data doesn't exist in dropdown list
                    'data error will be called that inserts unlisted value into dropdown box
                    'dgv(e.ColumnIndex, e.RowIndex) = cbx

                Catch ex As Exception

                End Try

            Else

                Try
                    dgvR = Me.dgvReports
                    If dgvR.Rows.Count = 0 Then
                    Else
                        idR = dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value
                        If idR > 1 And idR < 1000 Then
                            boolVal = True
                        Else
                            boolVal = False
                        End If
                    End If

                    If boolVal Then 'check for readonly
                        var1 = dgv(0, e.RowIndex).Value

                        strM = ""
                        Select Case var1
                            Case "Validation Corporate Study/Project Number"
                                strM = "To change '" & var1 & "', change the 'Corporate Study/Project Number' text box on the 'Add/Edit Top Level Data' page."
                            Case "Validation Protocol Number"
                                strM = "To change '" & var1 & "', change the 'Protocol Number' text box on the 'Add/Edit Top Level Data' page."
                            Case "Validation Report Title"
                                strM = "To change '" & var1 & "', change the 'Report Title' text box of the 'Configured Reports' table on the 'Choose Study & Report' page."
                            Case "Validation Report Number"
                                strM = "To change '" & var1 & "', change the 'Report Number' text box of the 'Configured Reports' table on the  'Choose Study & Report' page."
                                'Case "Analytical Method Type"
                                '    strM = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."
                            Case "Assay Technique" '20190212 LEE:
                                strM = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."

                                '20190220 LEE: Add some more warnings for stability
                            Case "Freeze/Thaw Stability", "Bench-top Stability", "Process Stability", "Reinjection Stability", "Batch Reinjection Stability", "Long-term Storage Stability", "Whole Blood Stability", "Stock Solution Stability", "Spiking Solution Stability", "Autosampler Stability" '20190212 LEE: 
                                strM = "To change '" & var1 & "', corresponding Stability Conditions Summary cell in the Advanced Table Configuration window - Stability Tab."
                            Case "Maximum # of Freeze/thaw Cycles" '20190212 LEE: 
                                strM = "To change '" & var1 & "', corresponding [#Cylces] Information cell in the Advanced Table Configuration window - Stability Tab."
                            Case "Anticoagulant/Preservative" '20190220 LEE: 
                                strM = "To change '" & var1 & "', change the 'Anticoagulant' dropdown box on the 'Add/Edit Top Level Data' page."
                        End Select

                        If Len(strM) = 0 Then
                        Else
                            MsgBox(strM, vbInformation, "Invalid action...")
                        End If

                    End If
                Catch ex As Exception

                End Try

            End If
        End If

end1:

    End Sub

    Function boolMethodValReadOnly(ByVal intCol As Short) As Boolean

        boolMethodValReadOnly = False

        If intCol < 0 Then
            boolMethodValReadOnly = True
            Exit Function
        End If

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView

        dgv1 = Me.dgvMethodValData
        dgv2 = Me.dgvMethValExistingGuWu

        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1, var2, var3, var4
        Dim bool As Boolean

        Dim dgvR As DataGridView
        dgvR = Me.dgvReports
        If dgvR.RowCount = 0 Then
        Else
            var1 = NZ(dgvR.Item("CHARREPORTTYPE", 0).Value, "Sample Analysis")
            If StrComp(var1, "Sample Analysis", CompareMethod.Text) = 0 Then
                var1 = NZ(dgv2("WatsonStudy", intCol - 1).Value, "")
                If Len(var1) = 0 Then
                    boolMethodValReadOnly = False
                Else 'read-only dgv1 column
                    boolMethodValReadOnly = True
                End If
            Else
                boolMethodValReadOnly = False
            End If
        End If

    End Function

    Private Sub dgvMethodValData_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvMethodValData.CellValidating

        If e.ColumnIndex = 0 Then
            Exit Sub
        End If

        If Me.cmdExit.Enabled Then
            Exit Sub
        End If

        If id_tblStudies < 1 Then
            Exit Sub
        End If

        If boolFormLoad Then
            Exit Sub
        End If

        If id_tblPersonnel < 1 Then
            Exit Sub
        End If

        Dim strM As String

        Dim strMod As String = ""
        Dim strSource As String = ""


        If boolMethodValReadOnly(e.ColumnIndex) Then
            If e.FormattedValue = NZ(Me.dgvMethodValData.Item(e.ColumnIndex, e.RowIndex).Value, "") Then
            Else
                strM = "A Method Validation study has been assigned to this Analyte." & ChrW(10) & ChrW(10) & "Therefore, the table entries associated with this Analyte are read-only."
                strM = strM & ChrW(10) & ChrW(10) & "Hit the [ESC] key to restore old value."
                MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
                Me.dgvMethodValData.Item(e.ColumnIndex, e.RowIndex).Value = Me.dgvMethodValData.Item(e.ColumnIndex, e.RowIndex).Value
                e.Cancel = True
            End If
        Else
            Dim boolVal As Boolean
            boolVal = False
            Dim dgvR As DataGridView
            Dim idR As Int64
            Dim var1, var2, var3
            Dim str1 As String
            Dim str2 As String
            Dim intCol As Short

            dgvR = Me.dgvReports
            If dgvR.Rows.Count = 0 Then
            Else
                idR = dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value
                If idR > 1 And idR < 1000 Then
                    boolVal = True
                Else
                    boolVal = False
                End If
            End If

            Dim dgv As DataGridView
            Dim boolGo As Boolean
            Dim boolGoNum As Boolean = False
            dgv = Me.dgvMethodValData

            Dim strM1 As String

            strM = ""
            If boolVal Then 'check for readonly
                var1 = dgv(0, e.RowIndex).Value
                boolGo = False
                boolGoNum = False
                strM1 = ""
                Select Case var1
                    Case "Validation Corporate Study/Project Number"
                        strM1 = "To change '" & var1 & "', change the 'Corporate Study/Project Number' text box on the 'Add/Edit Top Level Data' page."
                        strM1 = strM1 & ChrW(10) & ChrW(10)
                        boolGo = True
                    Case "Validation Protocol Number"
                        strM1 = "To change '" & var1 & "', change the 'Protocol Number' text box on the 'Add/Edit Top Level Data' page."
                        strM1 = strM1 & ChrW(10) & ChrW(10)
                        boolGo = True
                    Case "Validation Report Title"
                        strM1 = "To change '" & var1 & "', change the 'Report Title' text box of the 'Configured Reports' table on the 'Choose Study & Report' page."
                        strM1 = strM1 & ChrW(10) & ChrW(10)
                        boolGo = True
                    Case "Validation Report Number"
                        strM1 = "To change '" & var1 & "', change the 'Report Number' text box of the 'Configured Reports' table on the  'Choose Study & Report' page."
                        strM1 = strM1 & ChrW(10) & ChrW(10)
                        boolGo = True
                        'Case "Analytical Method Type" '20190212 LEE: deprecated
                        '    boolGo = True
                        '    strM1 = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."
                        '    strM1 = strM1 & ChrW(10) & ChrW(10)
                    Case "Assay Technique" '20190212 LEE:
                        boolGo = True
                        strM1 = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the 'Add/Edit Top Level Data' page."
                        strM1 = strM1 & ChrW(10) & ChrW(10)
                    Case "Sample Size"
                        boolGo = False
                        boolGoNum = True
                        'strM1 = "To change '" & var1 & "', change the 'Assay Technique' dropdown box on the Data page."
                        'strM1 = strM1 & ChrW(10) & ChrW(10)

                        '20190220 LEE: Add some more warnings for stability
                    Case "Freeze/Thaw Stability", "Bench-top Stability", "Process Stability", "Reinjection Stability", "Batch Reinjection Stability", "Long-term Storage Stability", "Whole Blood Stability", "Stock Solution Stability", "Spiking Solution Stability", "Autosampler Stability" '20190212 LEE: 
                        strM = "To change '" & var1 & "', corresponding Stability Conditions Summary cell in the Advanced Table Configuration window - Stability Tab."
                        strM1 = strM1 & ChrW(10) & ChrW(10)
                    Case "Maximum # of Freeze/thaw Cycles" '20190212 LEE: 
                        strM = "To change '" & var1 & "', corresponding [#Cylces] Information cell in the Advanced Table Configuration window - Stability Tab."
                        strM1 = strM1 & ChrW(10) & ChrW(10)
                    Case "Anticoagulant/Preservative" '20190220 LEE: 
                        strM = "To change '" & var1 & "', change the 'Anticoagulant' dropdown box on the 'Add/Edit Top Level Data' page."
                        strM1 = strM1 & ChrW(10) & ChrW(10)
                End Select
                If boolGo Then
                    str1 = NZ(dgv(e.ColumnIndex, e.RowIndex).Value, "")
                    str2 = NZ(e.FormattedValue, "")
                    If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                    Else
                        strM = "This item is read-only for a Method Validation report."
                        strM = strM & ChrW(10) & ChrW(10)
                        strM = strM & strM1
                        strM = strM & "Press Esc on the keyboard to return original value."
                        MsgBox(strM, MsgBoxStyle.Information, "Read-only...")
                        e.Cancel = True
                        GoTo end1
                    End If
                End If

                If boolGoNum Then
                    'entry must be numeric
                    var2 = NZ(e.FormattedValue, 0)
                    If IsNumeric(var2) Then
                    Else
                        strM = "This entry must be numeric." & ChrW(10) & "The entry '" & var2 & "' is not numeric."
                        MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")
                        e.Cancel = True
                        GoTo end1
                    End If

                End If

                'now check for column limit
                strMod = "Review Method Validation"
                strSource = var1.ToString
                If boolCLExceeded(var1, "TBLMETHODVALIDATIONDATA", e.FormattedValue, False, strMod, strSource) Then
                    e.Cancel = True
                    GoTo end1
                Else
                    dgvMethodValData.AutoResizeRows()
                End If

            End If


            'now check for anticoagulant
            If StrComp(var1, "Anticoagulant/Preservative", CompareMethod.Text) = 0 And boolVal Then
                Dim strA As String
                strA = frmH.cbxAnticoagulant.Text
                str1 = e.FormattedValue ' NZ(dgv(e.ColumnIndex, e.RowIndex).Value, "")
                If StrComp(strA, str1, CompareMethod.Text) = 0 Then
                Else
                    strM = "This item is read-only for a Method Validation report."
                    strM = strM & ChrW(10) & ChrW(10)
                    strM = strM & "To change 'Anticoagulant', use the Anticoagulant dropdown box on the 'Add/Edit Top Level Data' page."
                    strM = strM & ChrW(10) & ChrW(10)
                    strM = strM & "Press Esc on the keyboard to return original value."
                    dgv(e.ColumnIndex, e.RowIndex).Value = strA
                    'dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                    MsgBox(strM, MsgBoxStyle.Information, "Read-only...")
                    e.Cancel = True
                    GoTo end1
                End If

            End If


        End If

        dgvMethodValData.AutoResizeColumns()

end1:

    End Sub

    Private Sub dgvMethodValData_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvMethodValData.DataError
        Dim var1
        Dim var2
        Dim dgv As DataGridView
        Dim str1 As String
        Dim boolGo As Boolean
        Dim cbx As DataGridViewComboBoxCell
        Dim cbx1 As New DataGridViewComboBoxCell
        Dim int1 As Short
        Dim Count1 As Short

        If e.ColumnIndex = 0 Then
            Exit Sub
        End If
        Try
            dgv = Me.dgvMethodValData
            var1 = NZ(dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).Value, "NA")
            str1 = dgv.Columns.Item(e.ColumnIndex).Name

            str1 = dgv.Item(0, e.RowIndex).Value
            If StrComp(str1, "Extraction Procedure Description", CompareMethod.Text) = 0 Then

                cbx = cbxxAssayDescr
                'look for var1 in cbx
                boolGo = False
                For Count1 = 0 To cbx.Items.Count - 1 'for debugging
                    var2 = cbx.Items(Count1)
                    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                        boolGo = True
                        Exit For
                    End If
                Next
                If boolGo Then
                Else
                    ''this code should work, but it doesn't
                    'cbx.Items.Add(var1)
                    'cbx1 = cbx.Clone
                    'dgv(e.ColumnIndex, e.RowIndex) = cbx1

                    cbx.Items.Add(var1)
                    'cbx1 = cbx.Clone
                    'select added item
                    cbx.Value = var1

                    'cbx1 = cbx.Clone
                End If

                'dgv(e.ColumnIndex, e.RowIndex) = cbx1
                'dgv(e.ColumnIndex, e.RowIndex).Value = var1

                'ElseIf StrComp(str1, "Analytical Method Type", CompareMethod.Text) = 0 Then '20190212 LEE: deprecated
            ElseIf StrComp(str1, "Assay Technique", CompareMethod.Text) = 0 Then '20190212 LEE:

                cbx = cbxxAnalMethType
                'look for var1 in cbx
                boolGo = False
                For Count1 = 0 To cbx.Items.Count - 1 'for debugging
                    var2 = cbx.Items(Count1)
                    If StrComp(var1, var2, CompareMethod.Text) = 0 Then
                        boolGo = True
                        Exit For
                    End If
                Next
                If boolGo Then
                Else
                    ''this code should work, but it doesn't
                    'cbx.Items.Add(var1)
                    'cbx1 = cbx.Clone
                    'dgv(e.ColumnIndex, e.RowIndex) = cbx1

                    cbx.Items.Add(var1)
                    'cbx1 = cbx.Clone
                    'select added item
                    cbx.Value = var1

                    'cbx1 = cbx.Clone
                End If

                'dgv(e.ColumnIndex, e.RowIndex) = cbx1
                'dgv(e.ColumnIndex, e.RowIndex).Value = var1



            Else
                'str1 = e.Exception.Message
                'MsgBox("Data Error: " & CStr(str1))
            End If
        Catch ex As Exception
            MsgBox(NZ(ex.InnerException, "Is Null"))

        End Try



    End Sub

    Private Sub dgvMethodValData_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMethodValData.MouseEnter
        Me.dgvMethodValData.Focus()
    End Sub

    Private Sub dgvReportTableConfiguration_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportTableConfiguration.CurrentCellDirtyStateChanged

        Dim intRow As Short
        Dim intCol As Short
        Dim str1 As String
        Dim dgv As DataGridView
        Dim bool As Boolean
        Dim var1, var2
        Dim id As Long

        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        dgv = Me.dgvReportTableConfiguration
        intRow = dgv.CurrentRow.Index ' e.RowIndex
        intCol = dgv.CurrentCell.ColumnIndex 'e.ColumnIndex

        str1 = dgv.Columns(intCol).Name
        id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value

        If StrComp(str1, "boolRequiresSampleAssignment", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            'If id = 1 Or id = 2 Then
            '    boolFormLoad = True
            '    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            '    var1 = dgv("boolRequiresSampleAssignment", intRow).Value
            '    If var1 = -1 Then
            '        str1 = "This table cannot have samples assigned to it."
            '        MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            '        dgv("boolRequiresSampleAssignment", intRow).Value = 0
            '        dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            '    End If
            '    boolFormLoad = False
            '    'boolFormLoad = True
            '    'dgv.Item("BOOLREQUIRESSAMPLEASSIGNMENT", intRow).Value = 0
            '    'boolFormLoad = False

            'Else
            '    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            'End If
        End If
        var1 = dgv("boolRequiresSampleAssignment", intRow).Value

    End Sub

    Private Sub dgvReportTableConfiguration_CellContentClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportTableConfiguration.CellContentClick

        Dim dgv As DataGridView
        Dim str1 As String
        Dim strCol As String
        Dim intRow As Short
        Dim intCol As Short
        Dim id As Long
        Dim var1, var2

        dgv = Me.dgvReportTableConfiguration


        'intRow = dgv.CurrentRow.Index ' e.RowIndex

        If dgv.CurrentRow Is Nothing Then
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index ' e.RowIndex
        intCol = dgv.CurrentCell.ColumnIndex 'e.ColumnIndex

        'ensure table can be displayed
        id = dgv("ID_TBLCONFIGREPORTTABLES", intRow).Value
        'If StrComp(dgv.Columns(intCol).Name, "BOOLREQUIRESSAMPLEASSIGNMENT", CompareMethod.Text) = 0 Then
        '    If id = 1 Or id = 2 Then
        '        var1 = dgv("boolRequiresSampleAssignment", intRow).Value
        '        If var1 = -1 Then
        '            str1 = "This table cannot have samples assigned to it."
        '            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
        '            dgv("BOOLREQUIRESSAMPLEASSIGNMENT", intRow).Value = 0
        '            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub cmdDuplicateTables_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDuplicateTables.Click

        Dim var1, var2
        Dim Count1 As Short
        Dim intRow As Short
        Dim intRows As Short
        Dim dgv As DataGridView
        Dim strTable As String
        Dim strM As String
        Dim varM

        dgv = Me.dgvReportTableConfiguration

        If id_tblStudies = 0 Then
            MsgBox("A study must be chosen", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            strM = "StudyDoc must be in Edit Mode to continue."
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")

            Exit Sub
        End If

        intRows = dgv.Rows.Count
        If intRows = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        strTable = NZ(dgv("CHARHEADINGTEXT", intRow).Value, "")

        strM = "Are you sure you wish to duplicate the following table?:" & ChrW(10) & ChrW(10)
        strM = strM & strTable & ChrW(10) & ChrW(10)
        strM = strM & "NOTE: Table data will not be duplicated. You must assign samples to the new table."

        varM = MsgBox(strM, MsgBoxStyle.OkCancel, "Are you sure?...")
        If varM = vbOK Then
            Call DuplicateTables(intRow, dgv)
            Call AssessSampleAssignment()
            Me.dgvReportTableConfiguration.AutoResizeRows()

        Else
            Exit Sub
        End If

    End Sub

    Sub DuplicateTables(ByVal intRow As Short, ByVal dgv As DataGridView)

        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim str1 As String
        Dim drowsMaxID() As DataRow
        Dim maxID
        Dim strF As String
        Dim var1
        Dim intCols As Short
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim idO As Int64
        Dim integnum As Short
        Dim strS As String

        dtbl = tblReportTables

        'find maxID for tblReportTable
        maxID = GetMaxID("tblReportTable", 1, True) 'if maxid increment is 1, then getmaxid already does putmaxid
       
        'add a row to tblreporttables
        var1 = dgv("ID_TBLREPORTTABLE", intRow).Value

        idO = var1
        strF = "ID_TBLREPORTTABLE = " & var1
        strS = "INTEGNUM DESC"
        rows = dtbl.Select(strF, strS)
        int1 = rows.Length 'debuggin
        Dim var2
        var2 = rows(0).Item("INTEGNUM")
        integnum = NZ(var2, 0) + 1
        intCols = dtbl.Columns.Count
        Dim nrow As DataRow = dtbl.NewRow
        nrow.BeginEdit()
        For Count1 = 0 To intCols - 1
            str1 = dtbl.Columns(Count1).ColumnName
            If StrComp(str1, "ID_TBLREPORTTABLE", CompareMethod.Text) = 0 Then
                nrow.Item(Count1) = maxID
            Else
                var1 = rows(0).Item(Count1)
                nrow.Item(Count1) = var1
            End If
        Next
        nrow.Item("INTEGNUM") = integnum
        nrow.EndEdit()
        dtbl.Rows.Add(nrow)

        'now select the row
        int2 = -1
        For Count1 = 0 To dgv.Rows.Count - 1
            var1 = dgv("ID_TBLREPORTTABLE", Count1).Value
            If var1 = maxID Then
                If dgv.Columns("CHARHEADINGTEXT").Visible Then
                    int1 = dgv.Columns("CHARHEADINGTEXT").Index
                Else
                    int1 = dgv.Columns("CHARTABLENAME").Index
                End If
                dgv.CurrentCell = dgv.Rows.Item(Count1).Cells(int1)
                dgv.Rows(Count1).Selected = True
                int2 = Count1
                Exit For
            End If

        Next


        Call CheckForTblProperties(int2, idO)

        Call CheckForAutoAssignSamplesTable(idO, maxID, id_tblStudies, id_tblStudies, False, -1)

    End Sub

    Private Sub dgvReportTableConfiguration_ColumnDividerWidthChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewColumnEventArgs) Handles dgvReportTableConfiguration.ColumnDividerWidthChanged

        Call SizecmdOrder(Me.dgvReportTableConfiguration, frmH.cmdOrderReportTableConfig, "INTORDER")

    End Sub

    Private Sub cmdAdvancedTable_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdvancedTable.Click

        Dim intRow As Short
        Dim bool As Short
        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim strM As String
        Dim varR
        Dim str1 As String
        Dim strF As String
        Dim strFo As String

        'Call AssessSampleAssignment()

        Dim dgv1 As DataGridView
        dgv1 = Me.dgvReportTableConfiguration
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
        dv = Me.dgvReportTableConfiguration.DataSource
        dgv = frm.dgvReportTables

        'frm.intORow = dgv.CurrentRow.Index

        If Me.dgvReportTableConfiguration.RowCount = 0 Then
            frm.intORow = -1
            frm.idSel = 0
        ElseIf Me.dgvReportTableConfiguration.CurrentRow Is Nothing Then
            frm.intORow = -1
            frm.idSel = Me.dgvReportTableConfiguration("ID_TBLREPORTTABLE", 0).Value
        Else
            frm.intORow = Me.dgvReportTableConfiguration.CurrentRow.Index
            frm.idSel = Me.dgvReportTableConfiguration("ID_TBLREPORTTABLE", frm.intORow).Value
        End If


        'Dim int1 As Short
        'int1 = Me.dgvReportTableConfiguration.Columns.Count
        'For Count1 = 0 To int1 - 1
        '    varR = Me.dgvReportTableConfiguration.Columns(Count1).Name
        '    varR = 1
        'Next

        strF = "BOOLINCLUDE = TRUE"

        boolLoad = True
        strFo = dv.RowFilter
        dv.RowFilter = strF
        boolLoad = False
        dgv.DataSource = dv
        Call frm.InsertDefault(-1)

        'If Me.dgvReportTableConfiguration.RowCount = 0 Then
        '    frm.intORow = -1
        'ElseIf Me.dgvReportTableConfiguration.CurrentRow Is Nothing Then
        '    frm.intORow = -1
        'Else
        '    frm.intORow = Me.dgvReportTableConfiguration.CurrentRow.Index
        'End If

        Call frm.FormLoad()

        frm.strFilter = strFo

        frm.boolFormLoad = True

        frm.ShowDialog()

        frm.Dispose()



        'Me.dgvReportTableConfiguration.Visible = False

        'dv.RowFilter = strFo

        ''redo colors

        'Me.dgvReportTableConfiguration.AutoResizeRows()

        'Me.dgvReportTableConfiguration.Visible = True

        ''Call ResizeRows(Me.dgvCompanyAnalRef)
        'Call AssessSampleAssignment()

        'Call UpdateTablePropBools()

        'Call AssessQCs()


        'reselect introw
        Try

            dgv1.FirstDisplayedScrollingRowIndex = intRow

        Catch ex As Exception

        End Try

        Try

        Catch ex As Exception

            str1 = "Hmmm. There seems to be a problem opening the Advanced Table Configuration window." & ChrW(10)
            str1 = str1 & "Please try one or all of the following to resolve this situation:" & ChrW(10) & ChrW(10)
            str1 = str1 & "   - If StudyDoc is in Edit mode, try saving (Save button) current data." & ChrW(10)
            str1 = str1 & "   - Re-open the study." & ChrW(10)
            str1 = str1 & "   - Close and re-start StudyDoc."
            MsgBox(str1, MsgBoxStyle.Exclamation, "Hmmm...")

        End Try

        Cursor.Current = Cursors.Default


    End Sub

    Private Sub chkTableName_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTableName.CheckedChanged

        Call ShowTableName()

    End Sub

    Sub ShowTableName()

        Dim dgv As DataGridView
        dgv = Me.dgvReportTableConfiguration

        If Me.chkTableName.Checked Then
            dgv.Columns("CHARTABLENAME").Visible = True
        Else
            dgv.Columns("CHARTABLENAME").Visible = False
        End If

        Call SizecmdOrder(dgv, Me.cmdOrderReportTableConfig, "INTORDER")

    End Sub

    Private Sub cmdHeader_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHeader.Click

        Dim frm As New frmHeaderFooter
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim var1

        dgv = Me.dgvReports
        If dgv.Rows.Count = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            intRow = 0
        Else
            intRow = dgv.CurrentRow.Index
        End If

        var1 = dgv("ID_TBLREPORTS", intRow).Value
        id_tblReports = var1

        Call frm.LoadData()

        frm.ShowDialog()

        frm.Dispose()


    End Sub

    Private Sub dgvReportStatementWord_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportStatementWord.SelectionChanged


        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        Call UpdateWB_RBS()

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub cmiHomeFieldCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmiHomeFieldCode.Click
        Dim boolM As Boolean
        Dim strM As String

        strM = ""
        boolM = False

        If Me.dgvwStudy.Rows.Count = 0 Then
            strM = "A study must be chosen before this feature can be used"
            boolM = True
            Exit Sub
        End If

        If Me.cmdEdit.Enabled Then
            strM = "Please put StudyDoc in Edit mode in order to perform this action"
            boolM = True
            Exit Sub
        End If

        If boolM Then
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1

        End If

        Dim pos As Int64
        Dim strT As String
        Dim str1 As String
        Dim strL As String
        Dim strR As String
        'Dim wb As WebBrowser = frmH.wbRBS
        Dim wd As Microsoft.Office.Interop.Word.Document



        ''record position of cursor in text box
        'wd = wb.Document
        'pos = wd.selection.start

        'Dim frm As New frmFieldCodes

        'Me.Cursor = New Cursor(Cursor.Current.Handle)

        'frm.Location = new system.drawing.point(Cursor.Position.X, Cursor.Position.Y + 10)

        'frm.ShowDialog()

        'If frm.boolCancel Then

        '    wd.Selection.Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)


        'Else

        '    wd.Selection.TypeText(Text:=frm.strFC)

        'End If

        'frm.Dispose()

end1:

    End Sub


    Private Sub cmdReportHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReportHistory.Click

        Dim str1 As String
        Dim str2 As String

        If frmH.dgvReports.RowCount = 0 Then
            str1 = "A study report must be selected."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")

            Exit Sub
        End If

        Dim frm As New frmReportHistory
        Dim intRow As Short
        Dim dgv As DataGridView

        dgv = Me.dgvwStudy

        intRow = dgv.CurrentRow.Index
        str1 = dgv("STUDYNAME", intRow).Value
        str2 = "Report history for study: " & str1
        str2 = str2 & ChrW(10) & "(Initially sorted by date in DESC order)"

        frm.lblTitle.Text = str2

        frm.ShowDialog()

        frm.Dispose()

    End Sub

    Private Sub cmdOutliers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOutliers.Click


        If BOOLREPORTTABLECONFIGURATION Then
        Else
            MsgBox("User not allowed to execute items in Table Configuration window.", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Dim frm As New frmOutliers

        If id_tblStudies = 0 Then
            MsgBox("A study must be chosen", MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If


        frm.boolFormLoad = True

        frm.ShowDialog()

    End Sub

    Private Sub rbEntireReport_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbEntireReport.CheckedChanged

        If Me.rbEntireReport.Checked Then
            Call ViewSections(False)
        Else
            Call ViewSections(True)
        End If

    End Sub

    Public Sub ViewSections(ByVal bool As Boolean)

        'bool: False=show reports, True=show sections
        Dim dgv As DataGridView
        Dim Count1 As Short
        Dim str1 As String
        Dim dv As System.Data.DataView
        Dim strF As String
        Dim var1

        If Me.rbEntireReport.Checked Then
            boolEntireReport = True
        Else
            boolEntireReport = False
        End If


        Try
            dgv.Columns("CHARHEADINGTEXT").Frozen = False

        Catch ex As Exception

        End Try

        Cursor.Current = Cursors.WaitCursor

        Dim a, b

        Try

            dgv = Me.dgvReportStatements
            dv = dgv.DataSource
            If bool Then

                'Call InitializeReportStatements()

                'restore dgv columns
                For Count1 = 0 To dgv.Columns.Count - 1
                    Cursor.Current = Cursors.WaitCursor

                    dgv.Columns(Count1).HeaderText = arrRBSColumns(1, Count1)
                    dgv.Columns(Count1).Visible = arrRBSColumns(0, Count1)
                Next

                Cursor.Current = Cursors.WaitCursor
                var1 = arrRBSColumns(2, 0) 'debuging
                dv.RowFilter = arrRBSColumns(2, 0)

                Cursor.Current = Cursors.WaitCursor

                Me.rbRBS_Col.Checked = arrRBSColumns(3, 0)

                Me.cmdRefreshStatements.Text = "Sho&w Word Report Templates"
                Me.cmdOpenReportStatements.Text = "&Edit Word Report Templates"
                Me.lblWordStatements.Text = "<< Doubleclick to assign Word Report Template"

                a = Me.lblWordStatements.Left
                b = Me.lblWordStatements.Width

                Me.gbxlblChooseEditWordTemplate.Width = a + b + 10

            Else

                'record dgv filter
                'var1 = dv.RowFilter.ToString
                'arrRBSColumns(2, 0) = dv.RowFilter.ToString

                'record rbRBS_Col
                arrRBSColumns(3, 0) = Me.rbRBS_Col.Checked

                For Count1 = 0 To dgv.Columns.Count - 1
                    str1 = dgv.Columns(Count1).Name
                    If StrComp(str1, "CHARHEADINGTEXT", CompareMethod.Text) = 0 Then
                        dgv.Columns(Count1).Visible = False
                        dgv.Columns(Count1).HeaderText = ""
                    ElseIf StrComp(str1, "CHARSTATEMENT", CompareMethod.Text) = 0 Then
                        dgv.Columns(Count1).Visible = True
                        dgv.Columns(Count1).HeaderText = "Assigned Word Report Template"
                        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                    Else
                        dgv.Columns(Count1).Visible = False
                    End If
                Next

                strF = "NUMCOMPANY = 20000 AND ID_TBLSTUDIES = " & id_tblStudies
                dv.RowFilter = strF

                Me.cmdRefreshStatements.Text = "Sho&w Word Report Templates"
                Me.cmdOpenReportStatements.Text = "&Edit Word Report Templates"
            End If

            Me.lblRBS.Visible = bool
            Me.cmdRBSAll.Visible = bool
            Me.cmdOrderReportBodySection.Visible = bool
            Me.panSections.Visible = bool

        Catch ex As Exception

        End Try

        Call SetStatementTitle()

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        'If boolEntireReport Then
        '    Try
        '        dgv.Columns("CHARHEADINGTEXT").Frozen = False


        '    Catch ex As Exception

        '    End Try
        'Else
        '    Try
        '        If frmH.rbRBS_Col.Checked Then
        '            frmH.dgvReportStatements.Columns.Item("CHARHEADINGTEXT").Frozen = True
        '        Else
        '            frmH.dgvReportStatements.Columns.Item("charSectionName").Frozen = True
        '        End If

        '    Catch ex As Exception
        '    End Try
        'End If

        dgv.AutoResizeColumns()

        Call ReportStatementWidth(Not (bool))
        Cursor.Current = Cursors.Default


    End Sub

    Sub ReportStatementWidth(ByVal boolEntire As Boolean)

        Dim dgv As DataGridView = Me.dgvReportStatementWord
        Dim var1

        Try
            dgv.Columns("ID_TBLWORDSTATEMENTS").Visible = False
        Catch ex As Exception
            var1 = var1 'debug
        End Try

        Exit Sub

        'for reference
        'arrRBSColumns(4, 0) = Me.dgvReportStatements.Width
        'arrRBSColumns(5, 0) = Me.dgvReportStatements.Left
        'arrRBSColumns(6, 0) = Me.dgvReportStatementWord.Width
        'arrRBSColumns(7, 0) = Me.dgvReportStatementWord.Left

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim num1 As Decimal
        Dim num2 As Decimal
        Dim Count1 As Short
        Dim str1 As String


        dgv1 = Me.dgvReportStatements
        dgv2 = Me.dgvReportStatementWord

        num1 = 100
        num2 = 100
        Dim num3 As Decimal

        If boolEntire Then
            'find visible column
            For Count1 = 0 To dgv1.Columns.Count - 1
                str1 = dgv1.Columns(Count1).Name
                If StrComp(str1, "CHARHEADINGTEXT", CompareMethod.Text) = 0 Then
                    num1 = dgv1.Columns(Count1).Width
                ElseIf StrComp(str1, "CHARSTATEMENT", CompareMethod.Text) = 0 Then
                    num2 = dgv1.Columns(Count1).Width
                End If
            Next


            Try
                dgv1.Width = arrRBSColumns(4, 0) - 200 ' (num1 + num2) * 1.1
                dgv2.Left = arrRBSColumns(7, 0) - 200 'dgv1.Left + dgv1.Width + 3
                dgv2.Width = arrRBSColumns(6, 0) + 200 ' Me.tp5.Width - dgv2.Left
                Me.lblWordStatements.Left = dgv2.Left
                Me.lblWordStatements.Width = dgv2.Width

                Try
                    dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                Catch ex As Exception

                End Try

                dgv1.Columns("CHARHEADINGTEXT").Width = num1 / 2
                'num3 = (dgv1.Width - (num1 / 2))
                'dgv1.Columns("CHARSTATEMENT").Width = num3 '(dgv1.Width - (num1 / 2)) * 0.9


            Catch ex As Exception

            End Try

        Else

            Try
                dgv1.Width = arrRBSColumns(4, 0)
                dgv2.Left = arrRBSColumns(7, 0)
                dgv2.Width = arrRBSColumns(6, 0)
                Me.lblWordStatements.Left = dgv2.Left
                Me.lblWordStatements.Width = dgv2.Width
                dgv1.Columns("CHARHEADINGTEXT").Width = arrRBSColumns(8, 0)
                dgv1.Columns("CHARSTATEMENT").Width = arrRBSColumns(9, 0)

                dgv1.AutoSizeColumnsMode = arrRBSColumns(10, 0)

            Catch ex As Exception

            End Try

        End If

        Try
            dgv2.Columns("ID_TBLWORDSTATEMENTS").Visible = False
        Catch ex As Exception
            var1 = var1 'debug
        End Try


        dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        gbxlblChooseEditWordTemplate.Width = dgv2.Width

    End Sub

    Private Sub dgvReportTables_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportTables.MouseEnter
        dgvReportTables.Focus()
    End Sub


    Private Sub dgvReportTables_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReportTables.CellContentClick

    End Sub

    Private Sub dgvReportTables_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportTables.SelectionChanged
        If boolFormLoad Then
            Exit Sub
        End If

        Call ReportTableHeaderConfigPopulate()
    End Sub

    Private Sub mnuMenuAbout_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuMenuAbout.Click

        Dim frm As New frmAbout

        frm.ShowDialog()

        frm.Dispose()

    End Sub

    Private Sub mnuTroubleshooting_Click(sender As Object, e As EventArgs) Handles mnuTroubleshooting.Click

        Dim frm As New frmViewAllGroups

        frm.ShowDialog()

        frm.Dispose()

    End Sub

    Private Sub mnuMenuGenFC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuMenuGenFC.Click

        Dim var1
        Dim str1 As String

        str1 = "Do you wish to generate a Field Code report?"
        var1 = MsgBox(str1, MsgBoxStyle.YesNo, "Field Code report...")
        If var1 = 6 Then 'continue
        Else
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        Call GenerateFCReport()

        Cursor.Current = Cursors.Default


    End Sub

    Private Sub mnuHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelp.Click

        Try
            Dim strPath As String
            Dim strFilter As String
            Dim strFileName As String
            Dim str1 As String
            Dim str2 As String
            Dim boolGo As Boolean

            strPath = Me.txtArchivePath.Text
            If Len(strPath) = 0 Then
            Else
                Me.txtArchivePath.Text = "C:\"
            End If
            strPath = Me.txtArchivePath.Text

            strPath = "C:\Labintegrity\StudyDoc\Manuals\"
            strPath = "C:\LabIntegrity\StudyDoc\Manuals\"

            'strFilter = ".PDF files (*.PDF*)|*.PDF|.DOC files (*.DOC*)|*.DOC"
            strFilter = ".PDF files (*.PDF*)|*.PDF"
            strFileName = "*.PDF*"

            str1 = ReturnDirectoryBrowse(True, strPath, strFilter, strFileName, True) 'true = looking for file

            If Len(str1) = 0 Then
                Exit Sub
            End If

            Cursor.Current = Cursors.WaitCursor

            'Dim avCodeFile As CAcroAVDoc
            Dim Acroapp, avCodeFile
            Acroapp = CreateObject("AcroExch.App")
            Acroapp.Show()
            avCodeFile = CreateObject("AcroExch.AVDoc") 'This is the code file
            avCodeFile.Open(str1, "Code File")



        Catch ex As Exception

        End Try

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub dgvReports_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReports.SelectionChanged

        'Call Set_idtblReports()
        'Call RecordValSummary(False)


    End Sub


    Private Sub cbxArchivedMDB_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxArchivedMDB.SelectedIndexChanged

        Dim dgv As DataGridView
        Dim intRow As Short
        Dim str1 As String
        Dim tbl As System.Data.DataTable
        Dim drows() As DataRow
        Dim drows1() As DataRow
        Dim var1, var2
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim Count1 As Short
        Dim boolS As Boolean

        If cmdEdit.Enabled Then
            Exit Sub
        End If

        If cmdEdit.Enabled = False And cmdSave.Enabled = False Then
            Exit Sub
        End If

        dgv = dgvMethValExistingGuWu
        'determine if a row has been selected
        If dgv.RowCount = 0 Then
            Exit Sub
        End If
        If dgv.CurrentRow Is Nothing Then
            MsgBox("Please select an item from the table.", MsgBoxStyle.Information, "Select an item...")
            GoTo end1
        End If
        intRow = dgv.CurrentRow.Index

        'str1 = "charWatsonStudyName = '" & cbxMethValExistingGuWu.SelectedItem & "'"
        'tbl = tblStudies
        'drows = tbl.Select(str1)
        'var1 = cbxMethValExistingGuWu.SelectedItem
        'var2 = drows(0).Item("id_tblStudies")

        'enter info
        dv = dgv.DataSource
        int1 = dv.Count

        boolS = False

        str1 = Me.cbxArchivedMDB.SelectedItem

        For Count1 = 0 To int1 - 1
            If dgv.Rows(Count1).Selected Then
                boolS = True
                dv.Item(Count1).BeginEdit()
                dv.Item(Count1).Item("WatsonStudy") = str1
                dv.Item(Count1).Item("CHARARCHIVEPATH") = str1
                dv.Item(Count1).Item("id_tblStudies") = -1
                dv(Count1).EndEdit()
            End If

        Next

        If boolS Then 'continue
        Else
            str1 = "Remember to select one or more rows in order to fill Watson Study data."
            MsgBox(str1, MsgBoxStyle.Information, "No rows selected...")
            dgvMethValExistingGuWu.Select()
        End If


        dgv.AutoResizeColumns()
        dgv.AutoResizeRows()

end1:

    End Sub

    Sub ConfigMethValExistingGuWu()

        Dim dgv As DataGridView
        Dim intCols As Short
        Dim strM As String = ""
        Dim var1

        Try
            Try
                dgv = Me.dgvMethValExistingGuWu
            Catch ex As Exception
                MsgBox("Sub ConfigMethValExistingGuWu: dgv = Me.dgvMethValExistingGuWu:" & ChrW(10) & ex.Message)
            End Try

            intCols = dgv.ColumnCount
            'MsgBox(" intCols = dgv.ColumnCount: " & intCols)

            dgv.AllowUserToResizeRows = True
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AllowUserToResizeColumns = True
            dgv.AllowUserToResizeRows = True
            dgv.RowHeadersWidth = 20
            dgv.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing
            dgv.DefaultCellStyle.WrapMode = DataGridViewTriState.True

            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            Dim Count1 As Short

            If intCols < 1 Then
            Else
                dgv.Columns(dgv.Columns.Count - 1).Visible = False
                dgv.Columns(dgv.Columns.Count - 2).Visible = False

                For Count1 = 0 To dgv.Columns.Count - 1
                    Try
                        dgv.Columns(Count1).MinimumWidth = 100
                    Catch ex As Exception
                        'MsgBox("Sub ConfigMethValExistingGuWu: dgv.Columns(Count1).MinimumWidth = 100: " & Count1 & " of " & dgv.Columns.Count - 1 & ChrW(10) & ex.Message)
                        strM = "Sub ConfigMethValExistingGuWu: dgv.Columns(Count1).MinimumWidth = 100: " & Count1 & " of " & dgv.Columns.Count - 1 & ChrW(10) & ex.Message
                    End Try
                    dgv.Columns(Count1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                    dgv.Columns(Count1).DefaultCellStyle.WrapMode = DataGridViewTriState.True
                Next

                dgv.Columns(dgv.Columns.Count - 2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            End If

            dgv.AutoResizeColumns()
            dgv.AutoResizeRows()

            Try
                dgv.Columns(0).Width = 100
            Catch ex As Exception
                var1 = ex.Message
            End Try

            'try again. .NET 4.6.1 thing
            Try
                dgv.Columns(0).Width = 100
            Catch ex As Exception
                var1 = ex.Message
            End Try

        Catch ex As Exception
            'MsgBox("Sub ConfigMethValExistingGuWu:" & ChrW(10) & ex.Message)
        End Try

        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None


    End Sub


    Private Sub cbxAssayTechniqueAcronym_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbxAssayTechniqueAcronym.SelectedIndexChanged

        If boolFormLoad Then
            Exit Sub
        End If

        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim dgv As DataGridView
        Dim dgv2 As DataGridView
        Dim dvMVD As System.Data.DataView
        Dim strCol As String
        Dim boolVal As Boolean
        Dim idR As Int64

        'determine if method validation
        boolVal = False
        Dim dgvR As DataGridView
        dgvR = Me.dgvReports
        If dgvR.Rows.Count = 0 Then
        Else
            idR = NZ(dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value, -1)
            If idR > 1 And idR < 1000 Then
                boolVal = True
            Else
                boolVal = False
            End If
        End If

        dgv = dgvMethodValData
        dgv2 = dgvMethValExistingGuWu
        dvMVD = dgvMethValExistingGuWu.DataSource
        int3 = dvMVD.Count
        str1 = Me.cbxAssayTechniqueAcronym.Text
        dv = dgvMethodValData.DataSource

        Try
            int1 = dv.Count
            'int2 = FindRowDV("Analytical Method Type", dv) '20190212 LEE: deprecated
            int2 = FindRowDV("Assay Technique", dv) '20190212 LEE:

            If boolVal Then
                Try
                    If dgv.Rows.Count = 0 Then
                    Else
                        For Count1 = 0 To int3 - 1
                            strCol = dgv2(0, Count1).Value
                            dv(int2).BeginEdit()
                            dv(int2).Item(strCol) = str1
                            dv(int2).EndEdit()
                        Next
                    End If
                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception

        End Try


    End Sub

    Private Sub cmdMethValUpdate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdMethValUpdate.Click

        Dim int1 As Short
        Dim Count1 As Short
        'Dim strM As Short

        'strM = "This action will replace all values in the table below."

        Dim strM1 As String
        Dim boolStabOnly As Boolean

        Dim strM As String
        If Me.gbMethValApplyGuWu.Visible Then
            strM1 = "This action will replace all values in the table below."

            strM1 = "This action may take several seconds to a minute."
            strM1 = strM1 & ChrW(10) & ChrW(10)
            strM1 = strM1 & "Do you wish to continue?"
            strM1 = strM1 & ChrW(10) & ChrW(10)

            int1 = MsgBox(strM1, MsgBoxStyle.OkCancel, "Do you wish to continue...")
            If int1 = 1 Then
            Else
                Exit Sub
            End If
            boolStabOnly = False
        Else
            strM1 = "This action will replace Stability values with those from Advanced Table Configuration - Sample Conditions value."
            strM1 = strM1 & ChrW(10) & ChrW(10)
            strM1 = strM1 & "Do you wish to continue?"
            int1 = MsgBox(strM1, MsgBoxStyle.OkCancel, "Do you wish to continue...")
            If int1 = 1 Then
            Else
                Exit Sub
            End If
            boolStabOnly = True

            Call FillTableStuffMethVal(True)
            Call UpdateValueSummaryTable()
            GoTo end1

        End If


        Cursor.Current = Cursors.WaitCursor

        'select all rows
        For Count1 = 0 To Me.dgvMethValExistingGuWu.RowCount - 1
            Me.dgvMethValExistingGuWu.Rows(Count1).Selected = True
        Next

        Call MethValExecute(boolStabOnly, True)

        'Call FillTableStuffMethVal()

        'Call UpdateValueSummaryTable()

end1:

        Cursor.Current = Cursors.Default

        MsgBox("Action complete.", MsgBoxStyle.OkOnly, "Action complete...") '20190203 LEE:

    End Sub

    Private Sub rbMultValYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbMultValYes.CheckedChanged
        Call RecordValSummary(True)
    End Sub

    Sub RecordValSummary(ByVal boolFromCheck As Boolean)

        Dim dgv1 As DataGridView
        Dim dv As System.Data.DataView
        Dim int1 As Short
        Dim bool As Boolean

        If frmH.boolHold Then
            Exit Sub
        End If

        bool = frmH.boolHold
        frmH.boolHold = True

        dgv1 = frmH.dgvReports
        dv = dgv1.DataSource

        'the following has been deprecated. Keep doing anyway to get a value
        If boolFromCheck Then
            If dgv1.RowCount = 0 Then
            Else
                If frmH.rbMultValYes.Checked Then
                    int1 = -1
                Else
                    int1 = 0
                End If

                dv(0).BeginEdit()
                dv(0).Item("BOOLMULTIVALSUM") = int1
                dv(0).EndEdit()
            End If
        Else

            If dgv1.RowCount = 0 Then
                int1 = -1
            Else
                int1 = NZ(dv(0).Item("BOOLMULTIVALSUM"), -1)
            End If

            If int1 = -1 Then
                frmH.rbMultValYes.Checked = True
                frmH.rbMultValNo.Checked = False
            Else
                frmH.rbMultValYes.Checked = False
                frmH.rbMultValNo.Checked = True
            End If
        End If

        frmH.boolHold = bool

    End Sub

    Private Sub cmdRTCUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRTCUp.Click

        Call UpDown(True)

        Exit Sub

        Dim dgv As DataGridView
        'Dim dv as system.data.dataview
        Dim intRow As Short
        Dim str1 As String
        Dim rowO As Short
        Dim rowN As Short
        Dim varS
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim id1 As Int64
        Dim id2 As Int64
        Dim strF As String
        Dim strS As String

        Dim strTxtFilter As String = frmH.txtFilterSamples.Text
        Dim strM As String
        If StrComp(strTxtFilter, "", CompareMethod.Text) = 0 Then
        Else
            strM = "Move Row function not available when Table rows have been filtered."
            MsgBox(strM, vbInformation, "Invalid action...")
            Exit Sub
        End If

        dgv = Me.dgvReportTableConfiguration

        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            str1 = "Please choose a row."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        If intRow = 0 Then
            Exit Sub
        End If

        Call OrderDGV(dgvReportTableConfiguration, "INTORDER", "ID_TBLREPORTTABLE")

        rowO = intRow + 1
        rowN = rowO - 1


        'record new row
        'Dim dv as system.data.dataview = New DataView(dgv.DataSource)

        Dim dv As System.Data.DataView

        dv = dgv.DataSource

        id1 = dv(intRow).Item("ID_TBLREPORTTABLE") 'debugging
        id2 = dv(intRow - 1).Item("ID_TBLREPORTTABLE") 'debugging

        'record sort
        varS = dv.Sort
        'set sort
        dv.Sort = "ID_TBLREPORTTABLE ASC"

        int1 = dv.Find(id1)
        dv(int1).BeginEdit()
        dv(int1).Item("INTORDER") = rowN
        dv(int1).EndEdit()

        int2 = dv.Find(id2)
        dv(int2).BeginEdit()
        dv(int2).Item("INTORDER") = rowO
        dv(int2).EndEdit()

        'put sort back on
        Try
            dv.Sort = varS
        Catch ex As Exception
            varS = "a"
        End Try

        dgv.DataSource = dv

        Call AssessSampleAssignment()
        dgv.AutoResizeRows()

        'select newly positioned row
        Try
            dgv.CurrentCell = dgv("CHARHEADINGTEXT", intRow - 1)
            dgv.CurrentRow.Selected = True
        Catch ex As Exception

        End Try


    End Sub

    Private Sub cmdRTCDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRTCDown.Click

        Call UpDown(False)

        Exit Sub

        Dim dgv As DataGridView
        'Dim dv as system.data.dataview
        Dim intRow As Short
        Dim str1 As String
        Dim rowO As Short
        Dim rowN As Short
        Dim varS
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim id1 As Int64
        Dim id2 As Int64
        Dim strF As String
        Dim strS As String

        Dim strTxtFilter As String = frmH.txtFilterSamples.Text
        Dim strM As String
        If StrComp(strTxtFilter, "", CompareMethod.Text) = 0 Then
        Else
            strM = "Move Row function not available when Table rows have been filtered."
            MsgBox(strM, vbInformation, "Invalid action...")
            Exit Sub
        End If

        dgv = Me.dgvReportTableConfiguration

        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            str1 = "Please choose a row."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        If intRow = dgv.RowCount - 1 Then
            Exit Sub
        End If

        Call OrderDGV(dgvReportTableConfiguration, "INTORDER", "ID_TBLREPORTTABLE")

        rowO = intRow + 1
        rowN = rowO + 1


        'record new row
        'Dim dv as system.data.dataview = New DataView(dgv.DataSource)

        Dim dv As System.Data.DataView

        dv = dgv.DataSource

        id1 = dv(intRow).Item("ID_TBLREPORTTABLE") 'debugging
        id2 = dv(intRow + 1).Item("ID_TBLREPORTTABLE") 'debugging

        'record sort
        varS = dv.Sort
        'set sort
        dv.Sort = "ID_TBLREPORTTABLE ASC"

        int1 = dv.Find(id1)
        dv(int1).BeginEdit()
        dv(int1).Item("INTORDER") = rowN
        dv(int1).EndEdit()

        int2 = dv.Find(id2)
        dv(int2).BeginEdit()
        dv(int2).Item("INTORDER") = rowO
        dv(int2).EndEdit()

        'put sort back on
        Try
            dv.Sort = varS
        Catch ex As Exception
            varS = "a"
        End Try

        dgv.DataSource = dv

        Call AssessSampleAssignment()
        dgv.AutoResizeRows()

        'select newly positioned row
        Try
            dgv.CurrentCell = dgv("CHARHEADINGTEXT", intRow + 1)
            dgv.CurrentRow.Selected = True
        Catch ex As Exception

        End Try

    End Sub

    Sub UpDown(boolUp As Boolean)

        Dim dgv As DataGridView
        'Dim dv as system.data.dataview
        Dim intRow As Short
        Dim str1 As String
        Dim rowO As Short
        Dim rowN As Short
        Dim varS
        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim id1 As Int64
        Dim id2 As Int64
        Dim strF As String
        Dim strS As String

        Dim strTxtFilter As String = frmH.txtFilterSamples.Text
        Dim strM As String
        If StrComp(strTxtFilter, "", CompareMethod.Text) = 0 Then
        Else
            strM = "Move Row function not available when Table rows have been filtered."
            MsgBox(strM, vbInformation, "Invalid action...")
            Exit Sub
        End If

        boolFormLoad = True


        dgv = Me.dgvReportTableConfiguration

        If dgv.RowCount = 0 Then
            Exit Sub
        End If

        If dgv.CurrentRow Is Nothing Then
            str1 = "Please choose a row."
            MsgBox(str1, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If


        intRow = dgv.CurrentRow.Index
        If boolUp Then
            If intRow = 0 Then
                Exit Sub
            End If
        Else
            If intRow = dgv.RowCount - 1 Then
                Exit Sub
            End If
        End If
     

        Call OrderDGV(dgvReportTableConfiguration, "INTORDER", "ID_TBLREPORTTABLE")


        rowO = intRow ' + 1
        If boolUp Then
            rowN = rowO - 1
        Else
            rowN = rowO + 1
        End If


        Dim dv As System.Data.DataView

        dv = dgv.DataSource

        id1 = dv(intRow).Item("ID_TBLREPORTTABLE") 'debugging
        If boolUp Then
            id2 = dv(intRow - 1).Item("ID_TBLREPORTTABLE") 'debugging
        Else
            id2 = dv(intRow + 1).Item("ID_TBLREPORTTABLE") 'debugging
        End If

        'int1 = dv.Find(id1)
        dv(intRow).BeginEdit()
        dv(intRow).Item("INTORDER") = rowN + 1
        dv(intRow).EndEdit()

        rowN = FindRowInDGV("ID_TBLREPORTTABLE", id2, dgv)
        dv(rowN).BeginEdit()
        dv(rowN).Item("INTORDER") = rowO + 1
        dv(rowN).EndEdit()


end2:

        Call AssessSampleAssignment()
        dgv.AutoResizeRows()

        'select newly positioned row
        Try
            'dgv.CurrentCell = dgv("CHARHEADINGTEXT", intRow)
            If boolUp Then
                dgv.CurrentCell = dgv("CHARHEADINGTEXT", intRow - 1)
            Else
                dgv.CurrentCell = dgv("CHARHEADINGTEXT", intRow + 1)
            End If

            dgv.CurrentRow.Selected = True
        Catch ex As Exception

        End Try

        'Call OrderDGV(dgvReportTableConfiguration, "INTORDER", "ID_TBLREPORTTABLE")

        boolFormLoad = False

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click

        Dim dgv As DataGridView
        Dim Count1 As Short

        dgv = Me.dgvReportTableConfiguration



    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Me.dgvReportTableConfiguration.Columns("ID_TBLREPORTTABLE").Visible = True

    End Sub

    Private Sub dgvDataCompany_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDataCompany.CellContentClick

    End Sub


    Sub ShowAuditTrail()

        Dim strM As String
        Dim boolA As Boolean = BOOLRWAUDITTRAIL
        If boolA Then
        Else
            strM = "User does not have permission to access the 'Report Writer Audit Trail' window."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        Cursor.Current = Cursors.WaitCursor

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled Then
            strM = "Please Save the current study"
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            GoTo end1
        End If

        boolFromRW = True

        Dim frm As New frmAuditTrail

        frm.strForm = Me.Name

        frm.ShowDialog()

        Try
            frm.Dispose()
        Catch ex As Exception

        End Try

        boolFromRW = False

        Cursor.Current = Cursors.Default

END1:

    End Sub

    Private Sub cmdAuditTrail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAuditTrail.Click

        Call ShowAuditTrail()

    End Sub

    Private Sub cbxAssayTechnique_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxAssayTechnique.SelectedIndexChanged

        'fill cbxAssayTechniqueAcronym
        Dim dtbl2 As System.Data.DataTable
        Dim rows() As System.Data.DataRow
        Dim strF As String
        Dim strS As String
        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short

        dtbl2 = tblDropdownBoxContent
        strF = "ID_TBLDROPDOWNBOXNAME = 3"
        strS = "INTORDER ASC"
        rows = dtbl2.Select(strF, strS)
        int1 = Me.cbxAssayTechnique.SelectedIndex
        str2 = NZ(rows(int1).Item("CHARACRONYM"), "NA")

        Try
            Me.cbxAssayTechniqueAcronym.Text = str2
        Catch ex As Exception

        End Try


    End Sub

    Private Sub dgvReports_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReports.CellValidated

        Dim dgv As DataGridView
        Dim str1 As String
        Dim var1, var2
        Dim int1 As Short

        Dim dv As System.Data.DataView
        Dim Count1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str2 As String
        Dim dgv2 As DataGridView
        Dim dvMVD As System.Data.DataView
        Dim strCol As String
        Dim boolVal As Boolean
        Dim idR As Int64
        Dim dgvR As DataGridView

        'determine if method validation
        boolVal = False
        dgvR = Me.dgvReports
        If dgvR.Rows.Count = 0 Then
        Else
            idR = NZ(dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value, -1)
            If idR > 1 And idR < 1000 Then
                boolVal = True
            Else
                boolVal = False
            End If
        End If

        dgv = dgvReports
        str1 = dgv.Columns.Item(e.ColumnIndex).Name

        If StrComp(str1, "ID_TBLCONFIGREPORTTYPE", CompareMethod.Text) = 0 Then
            Call SetReportConfigType()
            'ElseIf StrComp(str1, "CHARREPORTTYPE", CompareMethod.Text) = 0 Then
            '    Try
            '        var1 = dgv.Item(e.ColumnIndex, e.RowIndex).Value
            '        var2 = cbxxReportTypes.Value

            '        int1 = cbxxReportTypes.Items.Count
            '        If StrComp(var1, var2, CompareMethod.Text) = 0 Then
            '        Else
            '            'dgv.Item(e.ColumnIndex, e.RowIndex).Value = var2
            '            'str1 = var1'for debugging
            '        End If

            '    Catch ex As Exception

            '    End Try
            '    'str1 = "a"
        ElseIf StrComp(str1, "CHARREPORTNUMBER", CompareMethod.Text) = 0 Then
            If boolVal And boolFormLoad = False Then
                dgv = dgvMethodValData
                dgv2 = dgvMethValExistingGuWu
                dvMVD = dgvMethValExistingGuWu.DataSource
                int3 = dvMVD.Count
                dv = dgvMethodValData.DataSource
                int1 = dv.Count
                int2 = FindRowDV("Validation Report Number", dv)
                str1 = NZ(dgvR(e.ColumnIndex, e.RowIndex).Value, "")

                Try
                    If dgv.Rows.Count = 0 Then
                    Else
                        For Count1 = 0 To int3 - 1
                            strCol = dgv2(0, Count1).Value
                            dv(int2).BeginEdit()
                            dv(int2).Item(strCol) = str1
                            dv(int2).EndEdit()
                        Next
                    End If
                Catch ex As Exception

                End Try

            End If
        ElseIf StrComp(str1, "CHARREPORTTITLE", CompareMethod.Text) = 0 Then
            If boolVal And boolFormLoad = False Then
                dgv = dgvMethodValData
                dgv2 = dgvMethValExistingGuWu
                dvMVD = dgvMethValExistingGuWu.DataSource
                int3 = dvMVD.Count
                dv = dgvMethodValData.DataSource
                int1 = dv.Count
                int2 = FindRowDV("Validation Report Title", dv)
                str1 = NZ(dgvR(e.ColumnIndex, e.RowIndex).Value, "")

                Try
                    If dgv.Rows.Count = 0 Then
                    Else
                        For Count1 = 0 To int3 - 1
                            strCol = dgv2(0, Count1).Value
                            dv(int2).BeginEdit()
                            dv(int2).Item(strCol) = str1
                            dv(int2).EndEdit()
                        Next
                    End If
                Catch ex As Exception

                End Try

            End If
        End If

    End Sub


    Private Sub dgvDataCompany_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvDataCompany.CellValidated

        Dim var1, var2
        Dim dgv As DataGridView
        Dim intRowDate1 As Short
        Dim tbl As System.Data.DataTable
        Dim intRow As Short
        Dim intCol As Short
        Dim varNull As System.DBNull
        Dim str1 As String
        Dim dt As Date
        Dim dv As System.Data.DataView

        Dim int1 As Short
        Dim Count1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim str2 As String
        Dim dgv2 As DataGridView
        Dim dvMVD As System.Data.DataView
        Dim strCol As String
        Dim boolVal As Boolean
        Dim idR As Int64
        Dim dgvR As DataGridView
        Dim dgvC As DataGridView

        If e.ColumnIndex = 1 Then
        Else
            Exit Sub
        End If

        If frmH.boolHold Then
            Exit Sub
        End If

        tbl = tblCompanyData
        dgv = dgvDataCompany
        dgvC = dgv
        str1 = dgv.Rows.Item(e.RowIndex).Cells(0).Value

        'check for sigfig value
        Dim boolCancel As Boolean
        intRow = e.RowIndex
        intCol = e.ColumnIndex
        If dgv(intCol, intRow).ReadOnly = True Then
            Exit Sub
        End If

        'determine if method validation
        boolVal = False
        dgvR = Me.dgvReports
        If dgvR.Rows.Count = 0 Then
        Else
            idR = NZ(dgvR("ID_TBLCONFIGREPORTTYPE", 0).Value, -1)
            If idR > 1 And idR < 1000 Then
                boolVal = True
            Else
                boolVal = False
            End If
        End If

        boolCancel = False
        If StrComp(str1, "Corporate Study/Project Number", CompareMethod.Text) = 0 Then
            If boolVal And boolFormLoad = False Then
                dgv = dgvMethodValData
                dgv2 = dgvMethValExistingGuWu
                dvMVD = dgvMethValExistingGuWu.DataSource
                int3 = dvMVD.Count
                dv = dgvMethodValData.DataSource
                int1 = dv.Count
                int2 = FindRowDV("Validation Corporate Study/Project Number", dv)
                str1 = dgvC(e.ColumnIndex, e.RowIndex).Value.ToString  'NDL: Added ToString for case in which value is NULL (was crashing previously)

                Try
                    If dgv.Rows.Count = 0 Then
                    Else
                        For Count1 = 0 To int3 - 1
                            strCol = dgv2(0, Count1).Value
                            dv(int2).BeginEdit()
                            dv(int2).Item(strCol) = str1
                            dv(int2).EndEdit()
                        Next
                    End If
                Catch ex As Exception

                End Try

            End If
        ElseIf StrComp(str1, "Protocol Number", CompareMethod.Text) = 0 Then
            If boolVal And boolFormLoad = False Then
                dgv = dgvMethodValData
                dgv2 = dgvMethValExistingGuWu
                dvMVD = dgvMethValExistingGuWu.DataSource
                int3 = dvMVD.Count
                dv = dgvMethodValData.DataSource
                int1 = dv.Count
                int2 = FindRowDV("Validation Protocol Number", dv)
                str1 = NZ(dgvC(e.ColumnIndex, e.RowIndex).Value, "")

                Try
                    If dgv.Rows.Count = 0 Then
                    Else
                        For Count1 = 0 To int3 - 1
                            strCol = dgv2(0, Count1).Value
                            dv(int2).BeginEdit()
                            dv(int2).Item(strCol) = str1
                            dv(int2).EndEdit()
                        Next
                    End If
                Catch ex As Exception

                End Try

            End If
        End If

    End Sub


    Private Sub dgvMethodValData_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvMethodValData.Click
        'Call MethodValReadOnly()
    End Sub

    Private Sub cmdSelect_ChangeUICues(sender As Object, e As UICuesEventArgs) Handles cmdSelect.ChangeUICues

    End Sub

    Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled Then
        Else
            Dim strM As String
            strM = "Please load a study"
            If Me.cmdEdit.Enabled Then
                strM = "Must be in Edit mode"
            Else
                strM = "Please load a study"
            End If
            MsgBox(strM, MsgBoxStyle.Information, "Invalid action...")
            Exit Sub
        End If

        Dim frm As New frmSelectTables

        frm.ShowDialog()

        If frm.boolCancel Then
        Else

            Dim Count1 As Short
            Dim Count2 As Short
            Dim Count3 As Short
            Dim dv As System.Data.DataView
            Dim int1 As Short
            Dim int2 As Short
            Dim tblAnalytes As System.Data.DataTable
            Dim str1 As String
            Dim str2 As String
            Dim str3 As String
            Dim drows() As DataRow
            Dim dgv As DataGridView
            Dim dtbl1 As System.Data.DataTable
            Dim strF As String
            Dim strS As String
            Dim boolCols As Boolean
            Dim boolI As Boolean

            'Dim cs As DataGridColumnStyle
            'Note: a tablestyle already exists

            dgv = frmH.dgvReportTableConfiguration
            dv = dgv.DataSource

            tblAnalytes = tblAnalRefStandards
            strF = "ID_TBLSTUDIES = " & id_tblStudies
            strS = "ID_TBLSTUDIES ASC"
            'drows = tblAnalytes.Select(strF, strS)

            strF = "IsIntStd = 'No'"
            strS = "OriginalAnalyteDescription ASC"
            drows = tblAnalytesHome.Select(strF, strS)

            dgv = Me.dgvReportTableConfiguration
            dv = dgv.DataSource

            If frm.rbColumns.Checked Then
                boolCols = True
            Else
                boolCols = False
            End If

            If frm.rbSelect.Checked Then
                boolI = True
            Else
                boolI = False
            End If

            int1 = frm.lbxAnalytes.SelectedItems.Count

            Dim items As ListBox.SelectedObjectCollection = New ListBox.SelectedObjectCollection(frm.lbxAnalytes)
            Dim itemsI As ListBox.SelectedIndexCollection = New ListBox.SelectedIndexCollection(frm.lbxAnalytes)

            Dim intA As Short
            Dim intB As Short

            intA = items.Count 'debug
            intB = itemsI.Count 'debug

            int1 = items.Count

            If boolCols Then
                For Count1 = 0 To int1 - 1
                    'int2 = itemsI(Count1)
                    str1 = items(Count1).ToString
                    For Count2 = 0 To dv.Count - 1
                        dv(Count2).BeginEdit()
                        dv(Count2).Item(str1) = boolI
                        dv(Count2).EndEdit()
                    Next

                Next
            Else

                For Count1 = 0 To int1 - 1
                    str1 = items(Count1).ToString
                    'find row in dgv
                    For Count3 = 0 To dgv.Rows.Count - 1
                        str2 = dgv("CHARHEADINGTEXT", Count3).Value
                        If StrComp(str1, str2, CompareMethod.Text) = 0 Then
                            'check analytes
                            For Count2 = 0 To drows.Length - 1
                                str3 = drows(Count2).Item("AnalyteDescription")
                                If dgv.Columns.Contains(str3) Then
                                    dv(Count3).BeginEdit()
                                    dv(Count3).Item(str3) = boolI
                                    dv(Count3).EndEdit()
                                End If
                            Next
                            Exit For
                        End If
                    Next

                Next

            End If


        End If

        Call AssessSampleAssignment()

        frm.Dispose()

    End Sub

    Private Function GetItemText(ByVal i As Integer, ByVal cbx As ComboBox) As String
        ' Return the text of the item using the index:
        Return cbx.Items(i).ToString
    End Function


    Private Sub cmdReplacePersonnel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReplacePersonnel.Click

        Dim frm As New frmReplacePersonnel

        frm.ShowDialog()

        frm.Dispose()


    End Sub

    Private Sub dtp1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtp1.TextChanged


        Dim intRow As Short
        Dim dgv As DataGridView

        If boolFromDataTab Then

            dgv = Me.dgvDataCompany
            intRow = dgv.CurrentRow.Index

            Try
                dgv.Rows(intRow).Cells(1).Value = Format(Me.dtp1.Value, LDateFormat)
            Catch ex As Exception
                dgv.Rows(intRow).Cells(1).Value = Format(Me.dtp1.Value, GDateFormat)
            End Try
            Me.dtp1.Visible = False

            boolFromDataTab = False

        End If
    End Sub

    Private Sub chkReadOnlyTables_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReadOnlyTables.CheckedChanged

        Call PutReadOnlyTables()

    End Sub

    Private Sub dgvMethodValData_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvMethodValData.CellContentClick

    End Sub

    Private Sub dgvAnalyticalRunSummary_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvAnalyticalRunSummary.CellContentClick

    End Sub

    Private Sub dgvAnalyticalRunSummary_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvAnalyticalRunSummary.CellValidating


        If e.ColumnIndex = 0 Then
            Exit Sub
        End If

        If Me.cmdExit.Enabled Then
            Exit Sub
        End If

        If id_tblStudies < 1 Then
            Exit Sub
        End If

        If boolFormLoad Then
            Exit Sub
        End If

        If id_tblPersonnel < 1 Then
            Exit Sub
        End If

        Dim strM As String
        Dim dgv As DataGridView
        Dim var1

        dgv = Me.dgvAnalyticalRunSummary

        Dim strM1 As String

        'find column name
        Dim strColName As String

        strColName = dgv.Columns(e.ColumnIndex).Name

        If StrComp(strColName, "User Comments", CompareMethod.Text) = 0 Then
            var1 = e.FormattedValue ' dgv(e.ColumnIndex, e.RowIndex).Value

            strM = ""
            'now check for column limit
            If Len(var1) > 255 Then
                e.Cancel = True

                strM = "This field is limited in length to 255." & ChrW(10) & ChrW(10)
                strM = strM & "The length of the entered text is " & Len(var1) & "." & ChrW(10) & ChrW(10)
                strM = strM & "Please modify the text to conform to the defined text limit."
                strM = strM & ChrW(10) & ChrW(10) & "Analytical Run Summary - User Comments cell"
                MsgBox(strM, MsgBoxStyle.Information, "Invalid entry...")

                GoTo end1

            End If
        End If

end1:

    End Sub

    Private Sub dgvReports_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvReports.CellContentClick

    End Sub

    Sub ConfigDropDowDGVs()


        Dim Count1 As Short
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim int2 As Short
        Dim var1, var2, var3

        'column names:
        'Item
        'Value
        'Example
        'charTab
        'boolIsBool
        'boolDD

        Try
            dgv = Me.dgvStudyConfig
            ''console.writeline(dgv.Name)
            'set dgv bool comboboxes
            For Count1 = 0 To dgv.RowCount - 1
                int2 = NZ(dgv("boolIsBool", Count1).Value, 0)
                If int2 = 0 Then
                    'if 514
                    Try
                        var1 = dgv("ITEM", Count1).Value
                        If StrComp(var1, "Use BQL/AQL vs BLQ/ALQ vs LLOQ/ULOQ", CompareMethod.Text) = 0 Then
                            Dim cbx3 As New DataGridViewComboBoxCell
                            cbx3.Items.Add("BQL/AQL")
                            cbx3.Items.Add("BLQ/ALQ")
                            cbx3.Items.Add("LLOQ/ULOQ")
                            dgv("Value", Count1) = cbx3
                            var1 = dgv("Value", Count1).Value
                            cbx3.Value = CStr(var1)
                        ElseIf StrComp(var1, "Character following table/figure/appendix caption", CompareMethod.Text) = 0 Then
                            Dim cbx3 As New DataGridViewComboBoxCell
                            cbx3.Items.Add("Tab")
                            cbx3.Items.Add("Soft Return")
                            cbx3.Items.Add("Space")
                            dgv("Value", Count1) = cbx3
                            var1 = dgv("Value", Count1).Value
                            cbx3.Value = CStr(var1)
                        End If
                    Catch ex As Exception
                        var2 = ex.Message
                    End Try
                    
                Else
                    Dim cbx3 As New DataGridViewComboBoxCell
                    cbx3.Items.Add("TRUE")
                    cbx3.Items.Add("FALSE")
                    'str1 = "Allow users to exclude data in StudyDoc (Watson overrides StudyDoc)"
                    'int1 = FindRowDVByCol(str1, dv1, "CHARCONFIGTITLE")
                    dgv("Value", Count1) = cbx3
                    'show Example column
                    'dgv.Columns("Example").Visible = False
                    var1 = dgv("Value", Count1).Value
                    cbx3.Value = CStr(var1)
                End If

            Next

            dgv = Me.dgvDataCompany
            ''console.writeline(dgv.Name)
            'set dgv bool comboboxes
            For Count1 = 0 To dgv.RowCount - 1
                int2 = NZ(dgv("boolIsBool", Count1).Value, 0)
                If int2 = 0 Then
                Else
                    Dim cbx3 As New DataGridViewComboBoxCell
                    cbx3.Items.Add("TRUE")
                    cbx3.Items.Add("FALSE")
                    'str1 = "Allow users to exclude data in StudyDoc (Watson overrides StudyDoc)"
                    'int1 = FindRowDVByCol(str1, dv1, "CHARCONFIGTITLE")
                    dgv("Value", Count1) = cbx3
                    'show Example column
                    'dgv.Columns("Example").Visible = False
                    var1 = dgv("Value", Count1).Value
                    cbx3.Value = CStr(var1)
                End If

            Next
        Catch ex As Exception
            var1 = ex.Message
        End Try


    End Sub

    Private Sub dgvStudyConfig_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvStudyConfig.CellClick

        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short
        Dim int4 As Short
        Dim intSTP As Short = -1
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim dv As System.Data.DataView
        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim locX, locY
        Dim var1, var2, var3
        Dim intRow As Short

        If e.ColumnIndex = 1 Then
        Else
            'Exit Sub
        End If

        dgv = Me.dgvStudyConfig
        dv = dgv.DataSource

        If dgv.ReadOnly Then
            'Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        int1 = FindRowDV("Table Date Format", dv)
        int2 = FindRowDV("Text Date Format", dv)
        intSTP = FindRowDV("Table-specific page numbering option", dv)
        'int3 = FindRowDV("Study Start Date", dv)
        'int4 = FindRowDV("Study End Date", dv)

        Me.panCal.Visible = False

        var1 = NZ(dgv.Rows(intRow).Cells(1).Value, "")

        If int1 = e.RowIndex Or int2 = e.RowIndex Or intSTP = e.RowIndex Then

            '***

            If int1 = e.RowIndex Then
                Dim cbx As New DataGridViewComboBoxCell
                cbx = cbxDateFormat.Clone
                cbx.AutoComplete = True
                cbx.MaxDropDownItems = 20
                cbx.DisplayStyleForCurrentCellOnly = True
                cbx.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
                Me.dgvStudyConfig("Value", e.RowIndex) = cbx
            End If

            If int2 = e.RowIndex Then
                Dim cbx1 As New DataGridViewComboBoxCell
                cbx1 = cbxDateFormat.Clone
                cbx1.AutoComplete = True
                cbx1.MaxDropDownItems = 20
                cbx1.DisplayStyleForCurrentCellOnly = True
                cbx1.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
                Me.dgvStudyConfig("Value", e.RowIndex) = cbx1
            End If

            If intSTP = e.RowIndex Then
                Dim cbx2 As New DataGridViewComboBoxCell
                cbx2.Items.Clear()
                cbx2.Items.Add("[None]")
                cbx2.Items.Add("Page x")
                cbx2.Items.Add("Page x of y")
                cbx2.AutoComplete = True
                cbx2.MaxDropDownItems = 20
                cbx2.DisplayStyleForCurrentCellOnly = True
                cbx2.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton
                Me.dgvStudyConfig("Value", e.RowIndex) = cbx2
            End If

            'Me.dgvStudyConfig.Refresh()

            '***
            If intSTP = e.RowIndex Then

                dgv.Columns.Item("Example").Visible = False
                'dgv.HorizontalScrollingOffset = dgv.HorizontalScrollingOffset - 10
                dv(int1).BeginEdit()
                dv(int1).Item("Example") = ""
                dv(int1).EndEdit()
                'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                dgv.AllowUserToResizeColumns = True
                dgv.AutoResizeColumns()

            Else

                If int1 = e.RowIndex Then
                ElseIf int2 = e.RowIndex Then
                    int1 = int2
                End If
                str1 = NZ(dv(int1).Item("Value"), "")
                If Len(str1) = 0 Then
                Else
                    strF = "CHARFORMAT = '" & str1 & "'"
                    tbl = tblDateFormats
                    rows = tbl.Select(strF)
                    If rows.Length = 0 Then 'ignore
                    Else
                        str2 = rows(0).Item("CHARDESCRIPTION") & " for Sep 1, 2006"
                        dv(int1).BeginEdit()
                        dv(int1).Item("Example") = str2
                        dv(int1).EndEdit()
                        dgv.Columns.Item("Example").Visible = True
                        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                        dgv.AutoResizeColumns()
                        'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                        'dgv.AllowUserToResizeColumns = True
                        dgv.HorizontalScrollingOffset = dgv.Width
                    End If
                End If
            End If

        Else
            dgv.Columns.Item("Example").Visible = False
            'dgv.HorizontalScrollingOffset = dgv.HorizontalScrollingOffset - 10
            dv(int1).BeginEdit()
            dv(int1).Item("Example") = ""
            dv(int1).EndEdit()
            'dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            dgv.AllowUserToResizeColumns = True
            dgv.AutoResizeColumns()
            dgv.AllowUserToResizeColumns = True

        End If

    End Sub

    Private Sub dgvStudyConfig_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvStudyConfig.CellValidating

        Dim var1, var2
        Dim dgv As DataGridView
        Dim intRowDate1 As Short
        Dim intRow As Short
        Dim intCol As Short
        Dim varNull As System.DBNull
        Dim str1 As String
        Dim dt As Date
        Dim dv As System.Data.DataView
        Dim boolTemp As Boolean
        Dim intCrit As Short

        If e.ColumnIndex = 1 Then
        Else
            Exit Sub
        End If

        dgv = Me.dgvStudyConfig
        str1 = dgv.Rows.Item(e.RowIndex).Cells(0).Value

        Dim strMod As String = "Add/Edit Top Level Data - Study Configuration"
        Dim strSource As String = str1

        'check for sigfig value
        Dim boolCancel As Boolean
        intRow = e.RowIndex
        intCol = e.ColumnIndex
        If dgv(intCol, intRow).ReadOnly = True Then
            Exit Sub
        End If

        'now check for column limit
        If boolCLExceeded(str1, "TBLDATA", e.FormattedValue, False, strMod, strSource) Then
            e.Cancel = True
            GoTo end1
        End If

        boolCancel = False
        Dim boolFo As Boolean = True
        Select Case str1

            Case "Data Sig Figs/Decimals" '"Data Significant Figures"
            Case "Regr Const Sig Figs/Decimals" '"Regr Const Sig Figs"
            Case "Regr R2 Sig Figs/Decimals" ' "Regr R2 Sig Figs"
            Case "Peak Area Sig Figs/Decimals" ' "Peak Area Significant Figures"
            Case "Peak Area Ratio Sig Figs/Decimals"
                'Case "Peak Area Decimal Places"
            Case Else
                boolFo = False
        End Select
        If boolFo Then
            'ensure data is integer >0
            'var1 = NZ(dgv(1, intRow).Value, "")
            intCrit = 0
            valErr = "Entry must be integer > 0"
            Select Case str1
                Case "Data Sig Figs/Decimals" '"Data Significant Figures"
                    var1 = dgv(intCol, intRow + 1).Value
                    If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                        valErr = "Entry must be integer > 0"
                        intCrit = 0
                    Else
                        valErr = "Entry must be integer >= 0"
                        intCrit = -1
                    End If
                Case "Regr Const Sig Figs/Decimals" '"Regr Const Sig Figs"
                Case "Regr R2 Sig Figs/Decimals" ' "Regr R2 Sig Figs"
                Case "Peak Area Sig Figs/Decimals" ' "Peak Area Significant Figures"
                    var1 = dgv(intCol, intRow + 1).Value
                    If StrComp(var1, "TRUE", CompareMethod.Text) = 0 Then
                        valErr = "Entry must be integer > 0"
                        intCrit = 1
                    Else
                        valErr = "Entry must be integer >= 0"
                        intCrit = 0
                    End If

                Case "Peak Area Ratio Sig Figs/Decimals" ' "Peak Area Significant Figures"


                    'Case "Peak Area Decimal Places"
                Case Else

            End Select

            var1 = e.FormattedValue
            If Len(var1) = 0 Then
                boolCancel = True
            ElseIf IsNumeric(var1) Then
                'evaluate further
                If IsInt(var1) Then
                    'ensure number is greater than crit
                    If CInt(var1) < intCrit Then
                        boolCancel = True
                    Else
                        boolOKtoValCD = True
                    End If
                Else
                    boolCancel = True
                End If
            Else
                boolCancel = True
            End If
            If boolCancel Then
                boolOKtoValCD = False

            End If
        ElseIf StrComp(str1, "Data Decimal Places", CompareMethod.Text) = 0 Or StrComp(str1, "Regression and R2 Decimal Places", CompareMethod.Text) = 0 Then
            'ensure data is integer >= 0
            'var1 = NZ(dgv(1, intRow).Value, "")
            var1 = e.FormattedValue
            If Len(var1) = 0 Then
                boolCancel = True
            ElseIf IsNumeric(var1) Then
                'evaluate further
                If IsInt(var1) Then
                    'ensure number is >= than 0
                    If var1 < 0 Then
                        boolCancel = True
                    Else
                        boolOKtoValCD = True
                    End If
                Else
                    boolCancel = True
                End If
            Else
                boolCancel = True
            End If
            If boolCancel Then
                boolOKtoValCD = False
                valErr = "Entry must be integer >= 0"
            End If
            'ElseIf StrComp(str1, "QC Stats % Decimal Places", CompareMethod.Text) = 0 Or StrComp(str1, "Peak Area Decimal Places", CompareMethod.Text) = 0 Then

        ElseIf StrComp(str1, "Table font size (0 to use Normal style font size)", CompareMethod.Text) = 0 Then
            'ensure data is integer >= 0
            'var1 = NZ(dgv(1, intRow).Value, "")
            var1 = e.FormattedValue
            If Len(var1) = 0 Then
                boolCancel = True
            ElseIf IsNumeric(var1) Then
                'evaluate further
                 If var1 < 0 Then
                    boolCancel = True
                Else
                    boolOKtoValCD = True
                End If
            Else
                boolCancel = True
            End If
            If boolCancel Then
                boolOKtoValCD = False
                valErr = "Entry must be number >= 0"
            End If

        ElseIf StrComp(str1, "QC Stats % Decimal Places", CompareMethod.Text) = 0 Then
            'ensure data is integer >= 0
            'var1 = NZ(dgv(1, intRow).Value, "")
            var1 = e.FormattedValue
            If Len(var1) = 0 Then
                boolCancel = True
            ElseIf IsNumeric(var1) Then
                'evaluate further
                If IsInt(var1) Then
                    'ensure number is >= than 0
                    If var1 < 0 Then
                        boolCancel = True
                    Else
                        boolOKtoValCD = True
                    End If
                Else
                    boolCancel = True
                End If
            Else
                boolCancel = True
            End If
            If boolCancel Then
                boolOKtoValCD = False
                valErr = "Entry must be integer >= 0"
            End If

            '20170106 LEE: Don't need this anymore because entry is now dropdownbox
            'ElseIf StrComp(str1, "Data: Use Sig Figs, not Decimals", CompareMethod.Text) = 0 Or StrComp(str1, "Data: Use Conc Special Rounding", CompareMethod.Text) = 0 Or StrComp(str1, "Enable StudyDoc Exclude Samples feature", CompareMethod.Text) = 0 Or StrComp(str1, "Enable StudyDoc Acceptance Crit. feature", CompareMethod.Text) = 0 Or StrComp(str1, "Peak Areas: Use Sig Figs, not Decimals", CompareMethod.Text) = 0 Or StrComp(str1, "Peak Areas: Use Conc Special Rounding", CompareMethod.Text) = 0 Or StrComp(str1, "Regression and R2: Use Sig Figs, not Decimals", CompareMethod.Text) = 0 Or StrComp(str1, "Regression and R2: Use Sci. Notation", CompareMethod.Text) = 0 Or StrComp(str1, "Place Nominal Concentrations in parentheses", CompareMethod.Text) = 0 Or StrComp(str1, "Add a date/time stamp on tables", CompareMethod.Text) = 0 Or StrComp(str1, "Footnote QC Means that exceed acceptance criteria", CompareMethod.Text) = 0 Or StrComp(str1, "Header/footer in right/left margin on landscape page", CompareMethod.Text) = 0 Or StrComp(str1, "Enter NA for non-entry QC or Calibr Std values", CompareMethod.Text) = 0 Or StrComp(str1, "Use BQL/AQL vs BLQ/ALQ", CompareMethod.Text) = 0 Or StrComp(str1, "Ignore Table-Specific Field Code generation", CompareMethod.Text) = 0 Or StrComp(str1, "Enable Page-Specific Legends", CompareMethod.Text) = 0 Or StrComp(str1, "Peak Area Ratio: Use Sig Figs, not Decimals", CompareMethod.Text) = 0 Or StrComp(str1, "Peak Area Ratio: Use Conc Special Rounding", CompareMethod.Text) = 0 Or StrComp(str1, "Use SigFigs/Decimals for Recovery/MatrixFactor values", CompareMethod.Text) = 0 Then '
            '    'ensure data is TRUE OR FALSE
            '    If StrComp(str1, "Use SigFigs/Decimals for Recovery/MatrixFactor values", CompareMethod.Text) = 0 Then
            '        var1 = var1 'debug
            '    End If
            '    var1 = e.FormattedValue
            '    If Len(var1) = 0 Then
            '        boolCancel = True
            '    ElseIf StrComp(var1, "TRUE", CompareMethod.Text) = 0 Or StrComp(var1, "FALSE", CompareMethod.Text) = 0 Then
            '        boolCancel = False
            '        boolOKtoValCD = True
            '        'ensure value is all caps
            '        str1 = AllCaps(var1)
            '        dv = Me.dgvStudyConfig.DataSource
            '        dv(e.RowIndex).BeginEdit()
            '        dv(e.RowIndex).Item("Value") = str1
            '        dv(e.RowIndex).EndEdit()
            '        Me.dgvStudyConfig.Refresh()
            '    Else
            '        boolCancel = True
            '    End If
            '    If boolCancel Then
            '        boolOKtoValCD = False
            '        valErr = "Entry must be TRUE or FALSE"
            '    End If


        ElseIf StrComp(str1, "Format comma for number >= (enter 0 to ignore)", CompareMethod.Text) = 0 Then
            var1 = e.FormattedValue
            valErr = "Entry must integer >= 1000 or = 0"
            If Len(var1) = 0 Or IsDBNull(var1) Or var1 Is Nothing Then
                boolOKtoValCD = False
            Else
                'check if numeric
                If IsNumeric(var1) Then
                    'ensure is integer
                    If IsInt(var1) Then
                        If var1 >= 1000 Or var1 = 0 Then
                            boolOKtoValCD = True
                        Else
                            boolOKtoValCD = False
                        End If
                    Else
                        boolOKtoValCD = False
                    End If
                Else
                    boolOKtoValCD = False
                End If
            End If

        ElseIf StrComp(str1, "Make hyperlinks and TOC blue font color", CompareMethod.Text) = 0 Then
            'ensure data is TRUE OR FALSE
            var1 = e.FormattedValue
            If Len(var1) = 0 Then
                boolCancel = True
            ElseIf StrComp(var1, "TRUE", CompareMethod.Text) = 0 Or StrComp(var1, "FALSE", CompareMethod.Text) = 0 Then
                boolCancel = False
                boolOKtoValCD = True
                'ensure value is all caps
                str1 = AllCaps(var1)
                dv = Me.dgvStudyConfig.DataSource
                dv(e.RowIndex).BeginEdit()
                dv(e.RowIndex).Item("Value") = str1
                dv(e.RowIndex).EndEdit()
                Me.dgvStudyConfig.Refresh()
            Else
                boolCancel = True
            End If
            If boolCancel Then
                boolOKtoValCD = False
                valErr = "Entry must be TRUE or FALSE"
            End If

        ElseIf StrComp(str1, "Format table anomalies with red bold font", CompareMethod.Text) = 0 Then
            'ensure data is TRUE OR FALSE
            var1 = e.FormattedValue
            If Len(var1) = 0 Then
                boolCancel = True
            ElseIf StrComp(var1, "TRUE", CompareMethod.Text) = 0 Or StrComp(var1, "FALSE", CompareMethod.Text) = 0 Then
                boolCancel = False
                boolOKtoValCD = True
                'ensure value is all caps
                str1 = AllCaps(var1)
                dv = Me.dgvStudyConfig.DataSource
                dv(e.RowIndex).BeginEdit()
                dv(e.RowIndex).Item("Value") = str1
                dv(e.RowIndex).EndEdit()
                Me.dgvStudyConfig.Refresh()
            Else
                boolCancel = True
            End If
            If boolCancel Then
                boolOKtoValCD = False
                valErr = "Entry must be TRUE or FALSE"
            End If

        End If

        If boolOKtoValCD Then
        Else
            e.Cancel = True

            boolFromCD = True
            boolFromCD = False
            MsgBox(valErr, MsgBoxStyle.Information, "Validation Error...")
        End If

end1:

    End Sub

    Private Sub dgvStudyConfig_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvStudyConfig.CurrentCellDirtyStateChanged

        Dim int1 As Short
        Dim int2 As Short
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim dv As System.Data.DataView
        Dim dgv As DataGridView
        Dim tbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim intRow As Short
        Dim intCol As Short

        dgv = Me.dgvStudyConfig
        intCol = dgv.CurrentCell.ColumnIndex
        If intCol = 1 Then
        Else
            Exit Sub
        End If

        intRow = dgv.CurrentRow.Index
        dv = dgv.DataSource
        int1 = FindRowDV("Text Date Format", dv)
        int2 = FindRowDV("Table Date Format", dv)

        If int1 = intRow Or int2 = intRow Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            str1 = dv(intRow).Item("Value")

            strF = "CHARFORMAT = '" & str1 & "'"
            tbl = tblDateFormats
            rows = tbl.Select(strF)
            str2 = rows(0).Item("CHARDESCRIPTION") & " for Sep 1, 2006"
            dv(intRow).BeginEdit()
            dv(intRow).Item("Example") = str2
            dv(intRow).EndEdit()
            dgv.Columns.Item("Example").Visible = True
            dgv.HorizontalScrollingOffset = dgv.Width
            dgv.AllowUserToResizeColumns = True
            dgv.AutoResizeColumns()
            dgv.AllowUserToResizeColumns = True


        Else
            dgv.Columns.Item("Example").Visible = False
            'dgv.HorizontalScrollingOffset = dgv.HorizontalScrollingOffset - 10
            'dv(int1).BeginEdit()
            'dv(int1).Item("Example") = ""
            'dv(int1).EndEdit()
            dgv.AllowUserToResizeColumns = True

        End If

    End Sub

    Private Sub dgvStudyConfig_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles dgvStudyConfig.DataError
        e.Cancel = True
    End Sub

    Private Sub dgvStudyConfig_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvStudyConfig.MouseEnter
        Me.dgvStudyConfig.Focus()
    End Sub

    Private Sub tabData_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles tabData.DrawItem

        Dim g As Graphics = e.Graphics
        Dim _TextBrush As Brush

        ' Use our own font.
        Dim _TabFont As New System.Drawing.Font("Segoe UI", 8.25, FontStyle.Bold, GraphicsUnit.Point)


        ' Get the item from the collection.
        Dim _TabPage As TabPage = Me.tabData.TabPages(e.Index)

        ' Get the real bounds for the tab rectangle.
        Dim _TabBounds As System.Drawing.Rectangle = Me.tabData.GetTabRect(e.Index)
        '_TabBounds.Width = _TabBounds.Width * 1.2

        '_TabPage.Width = _TabBounds.Width

        If (e.State = DrawItemState.Selected) Then
            ' Draw a different background color, and don't paint a focus rectangle.
            _TextBrush = New SolidBrush(Color.Blue)
            g.FillRectangle(Brushes.White, e.Bounds)
        Else
            _TextBrush = New System.Drawing.SolidBrush(e.ForeColor)
            e.DrawBackground()
        End If


        ' Draw string. Center the text.
        Dim _StringFlags As New StringFormat()
        _StringFlags.Alignment = StringAlignment.Center
        _StringFlags.LineAlignment = StringAlignment.Center
        g.DrawString(_TabPage.Text, _TabFont, _TextBrush, _TabBounds, New StringFormat(_StringFlags))


    End Sub


    Private Sub cmdCancelFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Call DoFCCancel()

        Dim dgv As DataGridView
        dgv = Me.dgvFC

        Dim dv As System.Data.DataView
        dv = dgv.DataSource

        dv.AllowEdit = True

        dgv.ReadOnly = False

    End Sub

    Private Sub dgvFC_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvFC.CellValidating

        Dim var1, var2
        Dim dgv As DataGridView
        Dim intRowDate1 As Short
        Dim intRow As Short
        Dim intCol As Short
        Dim varNull As System.DBNull
        Dim str1 As String
        Dim dt As Date
        Dim dv As System.Data.DataView
        Dim strM As String
        Dim boolE As Boolean = False


        If Me.cmdEdit.Enabled Then
            Exit Sub
        End If

        If Me.cmdEdit.Enabled = False And Me.cmdSave.Enabled = False Then
            Exit Sub
        End If

        Dim strCName As String
        dgv = Me.dgvFC
        str1 = dgv.Rows.Item(e.RowIndex).Cells(0).Value
        strCName = dgv.Columns(e.ColumnIndex).Name


        Dim strMod As String = "Administration - Custom Field Codes"
        Dim strSource As String = str1


        If StrComp(strCName, "CHARVALUE", CompareMethod.Text) = 0 Then
            'check for sigfig value
            Dim boolCancel As Boolean
            intRow = e.RowIndex
            intCol = e.ColumnIndex

            'now check for column limit
            If boolCLExceeded(strCName, "TBLCUSTOMFIELDCODE", e.FormattedValue, True, strMod, strSource) Then
                e.Cancel = True
                GoTo end1
            End If
        ElseIf StrComp(strCName, "INTORDER", CompareMethod.Text) = 0 Then
            'value must be integer
            var1 = e.FormattedValue
            strM = "Value must be integer >= 1"

            If IsNumeric(var1) Then
                If var1 >= 1 Then
                    If IsInt(var1) Then
                    Else
                        boolE = True
                    End If
                Else
                    boolE = True
                End If
            Else
                boolE = True
            End If

            If boolE Then
                MsgBox(strM, vbInformation, "Invalid entry...")
                e.Cancel = True
                GoTo end1
            End If

        End If

end1:

    End Sub

    Private Sub mnuShowFC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuShowFC.Click

        Dim frm As New frmFieldCodes

        frm.OK_Button.Enabled = False
        frm.Cancel_Button.Text = "E&xit"
        'frm.gbCopyAll.Visible = False
        frm.boolFromShowFC = True

        frm.ShowDialog()

        frm.Dispose()

end1:

    End Sub

   

    Private Sub cmdImportTables_Click(sender As System.Object, e As System.EventArgs) Handles cmdImportTables.Click

        Dim bool As Boolean

        Dim frm As New frmImportTables

        bool = Me.rbShowIncludedRTConfig.Checked

        Cursor.Current = Cursors.WaitCursor

        frm.ShowDialog()

        frm.Dispose()

        Me.rbShowIncludedRTConfig.Checked = bool
        Call RTFilter()

    End Sub

    Private Sub frmHome_01_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        'Call SizeCompanyAnalRef()

    End Sub

    Private Sub frmHome_01_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize

        'Call PositionProgress()

        Call SizeDumbThings()

    End Sub

    Sub SizeDumbThings()

        'Exit Sub 'deprecated

        'anchors aren't working with these items

        Dim a, b, c, d
        Dim t, l, w, h

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height

        w = Me.Width
        'a = Me.panSymbol.Width
        'l = w - a - 25

        'Me.panSymbol.Left = l

        a = Me.tab1.Left
        b = w - bw ' Me.panSymbol.Left
        c = b - a - 10
        Me.tab1.Width = c

        h = Me.Height
        a = Me.tab1.Top
        b = h - a - 50
        Me.tab1.Height = b


        'Me.panSymbol.Top = Me.tab1.Top
        'Me.panSymbol.Height = Me.tab1.Height

        'Call SetPanPos()



    End Sub

    Private Sub dgvwStudy_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvwStudy.CellContentClick

    End Sub

    Private Sub dgvwStudy_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvwStudy.CellClick

    End Sub

    Private Sub cmdCreateTable_Click(sender As System.Object, e As System.EventArgs) Handles cmdCreateTable.Click

        Dim strM As String

        If BOOLVIEWFINALREPORT Then
        Else
            strM = "User is not allowed to view reports. By extension, this means user also does not have permission to generate a report."
            MsgBox(strM, vbInformation, "Invalid action...")
            Me.cbxExampleReport.SelectedIndex = 0
            GoTo end1
        End If

        If BOOLALLOWREPORTGENERATION Then

            Me.cbxExampleReport.SelectedIndex = 4 '6
        Else
            strM = "User does not have permission to generate a report."
            MsgBox(strM, vbInformation, "Invalid action...")
            Me.cbxExampleReport.SelectedIndex = 0

        End If

end1:

    End Sub

    Private Sub cmdSymbol_Click(sender As System.Object, e As System.EventArgs) Handles cmdSymbol.Click

        Dim frm As New frmShowSymbol
        'place form to right of form

        Dim a, b, c, d

        Dim bw As Int16 = (Me.Width - Me.ClientSize.Width) / 2 'form border width
        Dim tbh As Int16 = Me.Height - Me.ClientSize.Height - 2 * bw 'titlebar height


        a = Me.ClientSize.Width
        a = Me.cmdSymbol.Left
        b = a ' a - frm.Width

        frm.Left = b

        c = Me.cmdSymbol.Top + Me.cmdSymbol.Height + tbh + 2
        frm.Top = c

        frm.Show()


    End Sub

    Private Sub gbActions_Paint(sender As Object, e As System.Windows.Forms.PaintEventArgs) 

        '20160617 LEE:
        'depricate


        ''penType: System.Drawing.Pen()

        ''Pen that determines the color, width, and style of the line. 
        ''x1Type: System.Int32()

        ''The x-coordinate of the first point. 
        ''y1Type: System.Int32()

        ''The y-coordinate of the first point. 
        ''x2Type: System.Int32()

        ''The x-coordinate of the second point. 
        ''y2Type: System.Int32()

        ''The y-coordinate of the second point. 



        'Dim p As Pen


        'p = New Pen(Color.Black, 1)
        'e.Graphics.DrawLine(p, 0, 5, 0, e.ClipRectangle.Height - 2) 'this is bottom border
        'e.Graphics.DrawLine(p, 0, 5, 10, 5)
        'e.Graphics.DrawLine(p, 62, 5, e.ClipRectangle.Width - 2, 5)
        'e.Graphics.DrawLine(p, e.ClipRectangle.Width - 2, 5, e.ClipRectangle.Width - 2, e.ClipRectangle.Height - 2)
        'e.Graphics.DrawLine(p, e.ClipRectangle.Width - 2, e.ClipRectangle.Height - 2, 0, e.ClipRectangle.Height - 2)

    End Sub

    Private Sub optStudyDocStudies_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optStudyDocStudies.CheckedChanged

        Call wStudyClicks(Me.optStudyDocStudies.Name)

    End Sub

    Private Sub optStudyDocOpen_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optStudyDocOpen.CheckedChanged

        Call wStudyClicks(Me.optStudyDocOpen.Name)

    End Sub

    Private Sub optStudyDocClosed_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles optStudyDocClosed.CheckedChanged

        Call wStudyClicks(Me.optStudyDocClosed.Name)

    End Sub

    Private Sub optStudyDocStudies_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles optStudyDocStudies.Validating

        'Call wStudyClicks(Me.optStudyDocStudies.Name)

    End Sub

    Private Sub optStudyDocOpen_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles optStudyDocOpen.Validating

        'Call wStudyClicks(Me.optStudyDocOpen.Name)

    End Sub

    Private Sub optStudyDocClosed_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles optStudyDocClosed.Validating

        'Call wStudyClicks(Me.optStudyDocClosed.Name)

    End Sub



    Private Sub TimerWarning_Tick(sender As System.Object, e As System.EventArgs) Handles TimerWarning.Tick


        If Me.panDot.Visible Then
            Me.panDot.Visible = False
        Else
            Me.panDot.Visible = True
        End If

        'Dim vc
        'vc = Me.lblWarning.BackColor

        'If vc = Color.FromArgb(255, 224, 192) Then
        '    Me.lblWarning.BackColor = Color.FromArgb(192, 255, 192)
        'Else
        '    Me.lblWarning.BackColor = Color.FromArgb(255, 224, 192)
        'End If

    End Sub

    Private Sub frmHome_01_ResizeEnd(sender As Object, e As System.EventArgs) Handles Me.ResizeEnd

        Call SizeCompanyAnalRef()

    End Sub

    Private Sub dgvCompanyAnalRef_CellBeginEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles dgvCompanyAnalRef.CellBeginEdit

        Dim dgv As DataGridView
        Dim str1 As String
        Dim bool As Boolean

        dgv = Me.dgvCompanyAnalRef
        bool = True
        str1 = dgv.Rows.Item(e.RowIndex).Cells("Item").Value
        Select Case str1
            Case "Is Replicate?"
                bool = False
            Case "Is Configured in Watson?"
                bool = False
            Case "ID"
                bool = False
            Case "Analyte Parent"
                bool = False
            Case "Is Internal Standard?"
                bool = False
        End Select
        If bool Then
        Else
            e.Cancel = True
            MsgBox("This cell is read-only", MsgBoxStyle.Information, "Read-only...")
        End If

    End Sub



    Private Sub dgvCompanyAnalRef_CellContentClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvCompanyAnalRef.CellContentClick

        Dim str1 As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim int2 As Long
        Dim var1
        Dim str2 As String

        dgv = Me.dgvCompanyAnalRef
        int1 = e.RowIndex
        If e.RowIndex < 0 Then
            Exit Sub
        End If
        var1 = NZ(dgv("ID_TBLDATATABLEROWTITLES", e.RowIndex).Value, "")
        If Len(var1) = 0 Then 'give up
            If var1 = 1 Then 'Company ID
                'determine if cell ia dropdownbox
                str2 = NZ(dgv.Rows.Item(e.RowIndex).Cells(e.ColumnIndex).EditType.FullName, "")
                If InStr(1, str2, "combobox", CompareMethod.Text) > 0 Then
                    'move to a different cell so that focus is lost
                    dgv.CurrentCell = dgv.Rows.Item(e.RowIndex + 1).Cells(e.ColumnIndex)
                End If
            End If
        End If

    End Sub

    Private Sub dgvCompanyAnalRef_CellValidating(sender As Object, e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles dgvCompanyAnalRef.CellValidating

        Dim dgv As DataGridView
        Dim str1 As String
        Dim bool As Boolean
        Dim bool1 As Boolean
        Dim bool2 As Boolean

        If e.ColumnIndex < 3 Then
            Exit Sub
        End If

        If boolFormLoad Then
            Exit Sub
        End If
        If cmdEdit.Enabled Then
            Exit Sub
        End If
        If Len(NZ(cbxStudy.Text, "")) = 0 Then
            Exit Sub
        End If

        dgv = Me.dgvCompanyAnalRef

        bool = False
        bool2 = False
        str1 = dgv.Rows.Item(e.RowIndex).Cells("Item").Value

        Dim strMod As String = "Analytical Reference Standard"
        Dim strSource As String = str1

        'now check for column limit
        If boolCLExceeded(str1, "tblCompanyAnalRefTable", e.FormattedValue, False, strMod, strSource) Then
            e.Cancel = True
            GoTo end1
        End If

        Select Case str1
            Case "Is Coadministered Cmpd?"
                bool = True
        End Select
        If InStr(1, str1, "Date", CompareMethod.Text) > 0 Then
            bool2 = True
        End If
        bool1 = True
        If bool Then
            'check to ensure value is Yes or No
            str1 = NZ(e.FormattedValue, "")
            If Len(str1) = 0 Then
                bool1 = False
            ElseIf StrComp(str1, "Yes", CompareMethod.Text) = 0 Or StrComp(str1, "No", CompareMethod.Text) = 0 Then
                'allow
            Else
                bool1 = False
            End If
            If bool1 Then
                Dim dv As System.Data.DataView
                dv = dgv.DataSource
                'ensure value is capitalized
                str1 = Capit(str1)
                dv(e.RowIndex).BeginEdit()
                dv(e.RowIndex).Item(e.ColumnIndex) = str1
                dv(e.RowIndex).EndEdit()
            Else
                e.Cancel = True
                MsgBox("This cell must be 'Yes' or 'No'", MsgBoxStyle.Information, "Yes or No...")
            End If
        ElseIf bool2 Then 'ensure value is date
            bool1 = False
            If Len(NZ(e.FormattedValue, "")) = 0 Then 'ignore
            ElseIf e.ColumnIndex < 3 Then 'ignore
                'ElseIf e.FormattedValue = True Or e.FormattedValue = False Then 'ignore
            Else
                bool1 = IsDate(e.FormattedValue)
                If bool1 Then 'continue
                    'format the date according to default
                    str1 = Format(CDate(e.FormattedValue), LDateFormat)
                    dgv(e.ColumnIndex, e.RowIndex).Value = str1
                Else
                    'e.Cancel = True
                    'MsgBox("Dude, value must be an acceptable date format", MsgBoxStyle.Information, "Invalid entry...")
                End If
            End If

        End If

end1:

    End Sub

    Private Sub dgvCompanyAnalRef_CurrentCellDirtyStateChanged(sender As Object, e As System.EventArgs) Handles dgvCompanyAnalRef.CurrentCellDirtyStateChanged

        If boolFormLoad Then
            Exit Sub
        End If
        If cmdEdit.Enabled Then
            Exit Sub
        End If
        If boolDirty Then
            boolDirty = False
            Exit Sub
        End If

        boolDirty = True

        Me.dgvCompanyAnalRef.CommitEdit(DataGridViewDataErrorContexts.Commit)

        Dim dgv As DataGridView
        Dim str1 As String
        Dim bool1 As Boolean
        Dim bool2 As Boolean
        Dim intRow As Short
        Dim intCol As Short

        dgv = Me.dgvCompanyAnalRef
        intRow = dgv.CurrentRow.Index
        intCol = dgv.CurrentCell.ColumnIndex
        If intCol = 0 Or intRow > 0 Then
            Exit Sub
        End If

        str1 = AnalRefHook()
        If Len(str1) > 0 Then
            Select Case str1
                Case "CRLWor_AnalRefStandard"
                    Call PopulateSingleFromCRLAnalRefHook(intCol)
            End Select
        End If

    End Sub

    Private Sub dgvCompanyAnalRef_LostFocus(sender As Object, e As System.EventArgs) Handles dgvCompanyAnalRef.LostFocus

        Me.panCal.Visible = False

    End Sub

    Private Sub dgvCompanyAnalRef_MouseEnter(sender As Object, e As System.EventArgs) Handles dgvCompanyAnalRef.MouseEnter

        Me.dgvCompanyAnalRef.Focus()

    End Sub

    Private Sub dgvCompanyAnalRef_Validated(sender As Object, e As System.EventArgs) Handles dgvCompanyAnalRef.Validated

        Me.dgvCompanyAnalRef.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        Me.dgvCompanyAnalRef.AutoResizeColumns()

    End Sub

    Sub FillDot()



    End Sub

    Private Sub panDot_Paint(sender As System.Object, e As System.Windows.Forms.PaintEventArgs) Handles panDot.Paint
        'Has already been appropriately sized in sizeDot function
        e.Graphics.FillRectangle(Brushes.Red, 0, 0, Me.panDot.Width, Me.panDot.Height)

    End Sub


    Private Sub dgvSampleReceiptWatson_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvSampleReceiptWatson.CellContentClick

    End Sub

    Sub CalLeave()

        If Me.mCal1.Focused Or Me.cmdEnterCal.Focused Or Me.cmdCalCancel.Focused Or Me.panCal.Focused Then
        Else
            Me.panCal.Visible = False
        End If

    End Sub

    Private Sub cmdEnterCal_LostFocus(sender As Object, e As System.EventArgs) Handles cmdEnterCal.LostFocus

        Call CalLeave()

    End Sub

    Private Sub cmdEnterCal_MouseLeave(sender As Object, e As System.EventArgs) Handles cmdEnterCal.MouseLeave

        Call CalLeave()

    End Sub

    Private Sub cmdCalCancel_LostFocus(sender As Object, e As System.EventArgs) Handles cmdCalCancel.LostFocus

        Call CalLeave()

    End Sub

    Private Sub cmdCalCancel_MouseLeave(sender As Object, e As System.EventArgs) Handles cmdCalCancel.MouseLeave

        Call CalLeave()

    End Sub


    Private Sub mCal1_LostFocus(sender As Object, e As System.EventArgs) Handles mCal1.LostFocus

        Call CalLeave()

    End Sub

    Private Sub mCal1_MouseLeave(sender As Object, e As System.EventArgs) Handles mCal1.MouseLeave

        Call CalLeave()

    End Sub


    Private Sub panCal_LostFocus(sender As Object, e As System.EventArgs) Handles panCal.LostFocus

        Call CalLeave()

    End Sub

    Private Sub panCal_MouseLeave(sender As Object, e As System.EventArgs) Handles panCal.MouseLeave

        Call CalLeave()

    End Sub

    Private Sub lblWarning_Click(sender As Object, e As EventArgs) Handles lblWarning.Click

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click

    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click

    End Sub

    Private Sub gbMethValApplyGuWu_Enter(sender As Object, e As EventArgs) Handles gbMethValApplyGuWu.Enter

    End Sub

    Private Sub Label40_Click(sender As Object, e As EventArgs) Handles Label40.Click

    End Sub

    Private Sub tp4_Click(sender As Object, e As EventArgs) Handles tp4.Click

    End Sub

    Private Sub dgQATable_Navigate(sender As Object, ne As NavigateEventArgs) Handles dgQATable.Navigate

    End Sub

    Private Sub dgvStudyConfig_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvStudyConfig.CellContentClick

    End Sub

    Sub dgvStudyConfig_CellMouseEnter(ByVal sender As Object, _
    ByVal e As DataGridViewCellEventArgs) _
    Handles dgvStudyConfig.CellMouseEnter

        If boolFormLoad Then
            Exit Sub
        End If

        Dim str1 As String
        Dim str2 As String
        Dim charConfigTitle
        Try
            If (e.ColumnIndex = Me.dgvStudyConfig.Columns("Item").Index) Then
                If ((e.ColumnIndex > -1) And (e.RowIndex > -1)) Then
                    charConfigTitle = Me.dgvStudyConfig.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                    With Me.dgvStudyConfig.Rows(e.RowIndex).Cells(e.ColumnIndex)
                        If (charConfigTitle.Equals("Data Sig Figs/Decimals")) Then
                            .ToolTipText = "Concentration data/stats: number of significant figures OR " _
                                & vbCrLf & "number of decimal places (see option)"
                        ElseIf (charConfigTitle.Equals("Data: Use Sig Figs, not Decimals")) Then
                            .ToolTipText = "Concentration data/stats: set to TRUE to use significant figure rounding,  " _
                                & vbCrLf & "FALSE to use decimal rounding"
                        ElseIf (charConfigTitle.Equals("Data: Use Conc Special Rounding")) Then
                            .ToolTipText = "Concentration data/stats: use significant figures when < 1000, " _
                                & vbCrLf & " otherwise round to whole number"

                        ElseIf (charConfigTitle.Equals("Peak Area Sig Figs/Decimals")) Then
                            .ToolTipText = "Peak Area data/stats: number of significant figures OR  " _
                                & vbCrLf & "number of decimal places (see option)"
                        ElseIf (charConfigTitle.Equals("Peak Areas: Use Sig Figs, not Decimals")) Then
                            .ToolTipText = "Peak Area data/stats: set to TRUE to use significant figure rounding,  " _
                                & vbCrLf & "FALSE to use decimal rounding"
                        ElseIf (charConfigTitle.Equals("Peak Areas: Use Conc Special Rounding")) Then
                            .ToolTipText = "Peak Area data/stats: use significant figures for numbers < 1000,  " _
                                & vbCrLf & "otherwise round to whole number"

                        ElseIf (charConfigTitle.Equals("Peak Area Ratio Sig Figs/Decimals")) Then
                            .ToolTipText = "Peak Area Ratio data/stats: number of significant figures OR  " _
                                & vbCrLf & "number of decimal places (see option)"
                        ElseIf (charConfigTitle.Equals("Peak Area Ratio: Use Sig Figs, not Decimals")) Then
                            .ToolTipText = "Peak Area Ratio data/stats: set to TRUE to use significant figure rounding,  " _
                                & vbCrLf & "FALSE to use decimal rounding"
                        ElseIf (charConfigTitle.Equals("Peak Area Ratio: Use Conc Special Rounding")) Then
                            .ToolTipText = "Peak Area Ratio data/stats: use significant figures for numbers < 1000,  " _
                                & vbCrLf & "otherwise round to whole number"
                        ElseIf (charConfigTitle.Equals("Regr Const Sig Figs/Decimals")) Then
                            .ToolTipText = "Calibration Standard regression constants: number of significant figures OR " _
                                & vbCrLf & "number of decimal places (depending on option chosen)"
                        ElseIf (charConfigTitle.Equals("Regr R2 Sig Figs/Decimals")) Then
                            .ToolTipText = "R-squared value:  number of significant figures OR  " _
                                & vbCrLf & "number of decimal places (see option)"
                        ElseIf (charConfigTitle.Equals("Regression and R2: Use Sig Figs, not Decimals")) Then
                            .ToolTipText = "Regression Constants & R-squared: set to TRUE to use significant figure rounding, " _
                                & vbCrLf & "FALSE to use decimal rounding"
                        ElseIf (charConfigTitle.Equals("Regression and R2: Use Sci. Notation")) Then
                            .ToolTipText = "Use scientific notation for regression constants and for R-Squared values"
                        ElseIf (charConfigTitle.Equals("Table Date Format")) Then
                            .ToolTipText = "Date format when dates appear on table  " _
                                & vbCrLf & "(e.g   MM/dd/yyyy ="" 06/30/2019"", dd-MMM-yyyy=""30-JUN-2019"")"
                        ElseIf (charConfigTitle.Equals("Text Date Format")) Then
                            .ToolTipText = "Date format when dates appear in report body " _
                                & vbCrLf & "(e.g. dddd  MMMM dd, yyyy=""Friday June 30, 2019"")"
                        ElseIf (charConfigTitle.Equals("Time Zone")) Then
                            .ToolTipText = "Value returned if the [TIMEZONE] field code is used in the report"
                        ElseIf (charConfigTitle.Equals("Alternate Calibr/QC Std Units")) Then
                            .ToolTipText = "If Watson does not provide the choice of calibration/QC standard units desired " _
                                & vbCrLf & "for this study (e.g. ng/g), enter the units here. "
                        ElseIf (charConfigTitle.Equals("QC Stats % Decimal Places")) Then
                            .ToolTipText = "QC tables:  In statistics section,  any % value (e.g. % Bias) is displayed to " _
                                & vbCrLf & "this number of decimal places"
                        ElseIf (charConfigTitle.Equals("Enable StudyDoc Exclude Samples feature")) Then
                            .ToolTipText = "QC Samples: Enable QC’s to be excluded from StudyDoc statistics " _
                                & vbCrLf & "(e.g. statistical outliers not already excluded in Watson).  " _
                                & vbCrLf & "See User Guide for more information. Only works if permitted in " _
                                & vbCrLf & "Report Writer Administration."

                        ElseIf (charConfigTitle.Equals("Enable StudyDoc Acceptance Crit. feature")) Then
                            .ToolTipText = "QC Samples: Allow different upper & lower % acceptance criteria " _
                                & vbCrLf & "(e.g.  +25%/ -20%  instead of   +/-25%)"
                        ElseIf (charConfigTitle.Equals("Format comma for number >= (enter 0 to ignore)")) Then
                            .ToolTipText = "Set to 1 to add comma thousands separator, 0 for no comma thousands separator " _
                                & vbCrLf & "(""11,000 vs. 11000"")"
                        ElseIf (charConfigTitle.Equals("Make hyperlinks and TOC blue font color")) Then
                            .ToolTipText = "Use blue font color for Table of Contents, Figures, Tables, " _
                                & vbCrLf & "Appendices, and all hyperlinked references to them"
                        ElseIf (charConfigTitle.Equals("Format table anomalies with red bold font")) Then
                            .ToolTipText = "Use red bold font for all report table anomalies & their footnotes " _
                                & vbCrLf & "(e.g. QC outside acceptance criteria)"
                        ElseIf (charConfigTitle.Equals("Place Nominal Concentrations in parentheses")) Then
                            .ToolTipText = "Report nominal concentrations in tables with parentheses " _
                                & vbCrLf & "[e.g.  ‘(10)’ vs. ’10’]"
                        ElseIf (charConfigTitle.Equals("Table-specific page numbering option")) Then
                            .ToolTipText = "Add table-specific page numbers at bottom of each page " _
                                & vbCrLf & "in every table (e.g. ""1 of 7"" for a 7-page table in a 40-page report) " _
                                & vbCrLf & "in addition to standard page numbering."
                        ElseIf (charConfigTitle.Equals("Add a date/time stamp on tables")) Then
                            .ToolTipText = "Add a date/time stamp to the bottom of each page of every table"
                        ElseIf (charConfigTitle.Equals("Footnote QC Means that exceed acceptance criteria")) Then
                            .ToolTipText = "QC Tables: add footnote to QC mean value if it exceeds acceptance criteria."
                        ElseIf (charConfigTitle.Equals("Header/footer in right/left margin on landscape page")) Then
                            .ToolTipText = "For tables in landscape mode, put Header in right margin (90 deg rotated)," _
                                & vbCrLf & " and put Footer in left margin"
                        ElseIf (charConfigTitle.Equals("Enter NA for non-entry QC or Calibr Std values")) Then
                            .ToolTipText = "If calibration standard or QC values do not exist " _
                                & vbCrLf & "replace with ""NA"" on report (vs leaving blank)"
                        ElseIf (charConfigTitle.Equals("Use BQL/AQL vs BLQ/ALQ vs LLOQ/ULOQ")) Then
                            .ToolTipText = "Choose the nomenclature for BQL/AQL reporting."
                        ElseIf (charConfigTitle.Equals("Ignore Table-Specific Field Code generation")) Then
                            .ToolTipText = "For very large studies (e.g. >100 analytical runs), " _
                                & vbCrLf & "this option can be set to TRUE to speed up loading of study."
                        ElseIf (charConfigTitle.Equals("Enable Page-Specific Legends")) Then
                            .ToolTipText = "Only include legend terms when those terms occur on the table page " _
                                & vbCrLf & "(e.g. ""BQL = Below Quantitation Limit"")"
                        ElseIf (charConfigTitle.Equals("Appendix/Figure/Table Caption Trailer")) Then
                            .ToolTipText = "Characters to add after Appendix/Figure/Table Caption.  " _
                                & vbCrLf & "E.g. '.' (Table 1.), ':' (Table 1:), '' (Table 1), etc."

                        ElseIf (charConfigTitle.Equals("Allow StdDev calculation if n = 2")) Then
                            .ToolTipText = "Calculate statistics Std Dev if n = 2" _
                                & vbCrLf & "(If FALSE, then n must be >= to 3 in order to have Std Dev calculated)"

                        ElseIf (charConfigTitle.Equals("If %Accuracy Column is displayed, then report average in Statistics section")) Then
                            .ToolTipText = "If %Bias/RE/Diff/Theor column is displayed," _
                               & vbCrLf & "report average of %Bias/RE/Diff/Theor column in Statistics section"

                        ElseIf (charConfigTitle.Equals("Character following table/figure/appendix caption")) Then
                            str1 = "The character following table/figure/appendix captions when the Word report is generated."
                            str1 = str1 & ChrW(10) & "Note that if choice is 'Soft Return', the Table of Contents/Figures/Appendices field codes may not display as expected."
                            .ToolTipText = str1

                            '20181109 LEE:
                        ElseIf (charConfigTitle.Equals("Use SigFigs for Recovery values")) Then
                            str1 = "If TRUE, the Recovery value will be reported to the same number of SigFigs as 'Data Sig Figs/Decimals'."
                            str1 = str1 & ChrW(10) & "If FALSE, the Recover value will be reported to the same number of decimal places as 'QC Stats % Decimal Places'. "
                            .ToolTipText = str1

                        ElseIf (charConfigTitle.Equals("Use %RSD label instead of %CV")) Then
                            str1 = "Use %RSD label instead of %CV for precision values."
                            .ToolTipText = str1

                        ElseIf (charConfigTitle.Equals("Table font size (0 to use Normal style font size)")) Then
                            str1 = "Sets font size of tables. Set to 0 to use Normal Style font size."
                            .ToolTipText = str1

                        ElseIf (charConfigTitle.Equals("Add chapter number to table caption label")) Then
                            str1 = "Add chapter number to table caption label."
                            str1 = str1 & ChrW(10) & "(e.g. 'Table 3-1' for the first table in Section 3)."
                            .ToolTipText = str1

                        ElseIf (charConfigTitle.Equals("Include Calibr Range in table title if multi-calibr range")) Then
                            str1 = "If the study has multiple calibration levels, include calibration level in table title."
                            .ToolTipText = str1

                        End If
                    End With
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Sub dgvDataCompany_CellMouseEnter(ByVal sender As Object, _
      ByVal e As DataGridViewCellEventArgs) _
      Handles dgvDataCompany.CellMouseEnter
        Dim charConfigTitle
        Try
            If ((e.ColumnIndex > -1) And (e.RowIndex > -1)) Then
                charConfigTitle = Me.dgvDataCompany.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
                With Me.dgvDataCompany.Rows(e.RowIndex).Cells(e.ColumnIndex)
                    If (charConfigTitle.Equals("Corporate Study/Project Number")) Then
                        .ToolTipText = "Number assigned to this study or project"
                    ElseIf (charConfigTitle.Equals("Protocol Number")) Then
                        .ToolTipText = "Number assigned to this protocol"
                    ElseIf (charConfigTitle.Equals("Sponsor Study Number")) Then
                        .ToolTipText = "Number assigned by sponsor (ignore if not applicable)"
                    ElseIf (charConfigTitle.Equals("Sponsor Study Title")) Then
                        .ToolTipText = "Study Title assigned by sponsor (ignore if not applicable)"
                    ElseIf (charConfigTitle.Equals("Study Start Date")) Then
                        .ToolTipText = "Enter date study was started (optional)"
                    ElseIf (charConfigTitle.Equals("Study End Date")) Then
                        .ToolTipText = "Enter date study was finished (optional)"
                    ElseIf (charConfigTitle.Equals("Data Archival Location")) Then
                        .ToolTipText = "Where data is archived " _
                            & vbCrLf & "(originally copied from your StudyDoc template)"
                    ElseIf (charConfigTitle.Equals("Outlier Method")) Then
                        .ToolTipText = "Which outlier method was used " _
                            & vbCrLf & "(originally copied from your StudyDoc template )"
                    End If
                End With
            End If
        Catch ex As Exception

        End Try
       
    End Sub

    Sub dgvReportTableHeaderConfig_CellMouseEnter(ByVal sender As Object, _
                                            ByVal e As DataGridViewCellEventArgs) _
                                            Handles dgvReportTableHeaderConfig.CellMouseEnter
        Dim charConfigTitle
        Dim str1 As String
        Dim strCol As String
        Dim dgv As DataGridView = Me.dgvReportTableHeaderConfig
        Dim strCR As String = "'Shift-Enter' to enter carriage return"

        '20160907 LEE: Reformed the IF-THEN-ELSE structure to all strCR tooltip for every CHARUSERLABEL mouseenter

        Try
            If (e.ColumnIndex = Me.dgvReportTableHeaderConfig.Columns("charColumnLabel").Index) Or _
           (e.ColumnIndex = Me.dgvReportTableHeaderConfig.Columns("charUserLabel").Index) Then

                If ((e.ColumnIndex > -1) And (e.RowIndex > -1)) Then

                    strCol = dgv.Columns(e.ColumnIndex).Name

                    charConfigTitle = Me.dgvReportTableHeaderConfig.Rows(e.RowIndex).Cells("charColumnLabel").Value
                    With Me.dgvReportTableHeaderConfig.Rows(e.RowIndex).Cells(e.ColumnIndex)

                        If StrComp(strCol, "CHARUSERLABEL", CompareMethod.Text) = 0 Then
                            strCR = "'Shift-Enter' to enter carriage return"
                        Else
                            strCR = ""
                        End If

                        If (StrComp(dgvReportTables.CurrentCell.Value, "Summary of Samples") = 0) Then
                            '*** Summary of Samples ***
                            If (charConfigTitle.Equals("Group")) Then
                                .ToolTipText = "Subject Group Name for the Design Subject Group" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Subject")) Then
                                .ToolTipText = "Tag representing the Study Design Subject" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Treatment")) Then
                                .ToolTipText = "The Treatment ID undertaken by the subject (e.g. "" A2"")" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Day")) Then
                                .ToolTipText = "Nominal Sampling Day of a collection (in days)" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Time")) Then
                                .ToolTipText = "Nominal Sampling Time of a collection (in hours, minutes, & seconds)" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Concentration")) Then
                                str1 = "Concentration of the Analyte measured in the Sample"
                                str1 = str1 & ChrW(10) & "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR
                                .ToolTipText = str1
                            ElseIf (charConfigTitle.Equals("Watson Run ID")) Then
                                .ToolTipText = "Watson Run ID #" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Analyte")) Then
                                .ToolTipText = "The Analyte Name whose measurement is being reported in the table" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Dil Factor")) Then
                                .ToolTipText = "The Dilution factor of the sample" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Start Day")) Then
                                .ToolTipText = "Nominal Start day of a collection (in days)." _
                                    & vbCrLf & "This is often left blank in Watson." & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Start Time")) Then
                                .ToolTipText = "Nominal Start time  of a collection (in hours, minutes, & seconds)." _
                                    & vbCrLf & "This is often left blank in Watson." & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Time Text")) Then
                                .ToolTipText = "Time Text associated with this sample (added in Watson)" _
                                    & vbCrLf & "(e.g. ""Pre-Dose"")" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Matrix")) Then
                                .ToolTipText = "Study Matrix (if entered into Watson)" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Gender")) Then
                                .ToolTipText = "Gender of subject (if entered into Watson)" & ChrW(10) & strCR
                            End If
                            'End If

                            '*** Summary of Analytical Runs***
                        ElseIf (StrComp(dgvReportTables.CurrentCell.Value, "Summary of Analytical Runs") = 0) Then
                            If (charConfigTitle.Equals("Watson Run ID")) Then
                                .ToolTipText = "Watson Run ID #" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Notebook ID")) Then
                                .ToolTipText = "Notebook Identifier (i.e. Run Identifier within Lab Notebook)" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Extraction Date")) Then
                                .ToolTipText = "Date that the first sample extraction took place" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Analysis Date")) Then
                                .ToolTipText = "Date on which the Analytical Run was started" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Samples")) Then
                                .ToolTipText = "Description of the Run (from Watson)" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Pass/Fail")) Then
                                .ToolTipText = "The Status of the Analyte’s Regression " _
                                    & vbCrLf & "(i.e. did the Analyte analysis pass or fail for the run)" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Comments")) Then
                                .ToolTipText = "Comments from within Watson or StudyDoc (depending on user settings " _
                                    & vbCrLf & "in Review Analytical Runs)" & ChrW(10) & strCR
                            ElseIf (charConfigTitle.Equals("Run Type")) Then
                                str1 = "The Run Type of the analytical run (e.g. UNKNOWNS, PSAE, etc.)" & ChrW(10) & strCR
                                .ToolTipText = str1
                            ElseIf (charConfigTitle.Equals("Matrix")) Then
                                str1 = "The Matrix assigned to this analytical run" & ChrW(10) & strCR
                                .ToolTipText = str1
                            ElseIf (charConfigTitle.Equals("LLOQ")) Then
                                str1 = "The Lower Limit of Quantitation of the analytical run calibration curve" & ChrW(10) & strCR
                                .ToolTipText = str1
                            ElseIf (charConfigTitle.Equals("ULOQ")) Then
                                str1 = "The Upper Limit of Quantitation of the analytical run calibration curve" & ChrW(10) & strCR
                                .ToolTipText = str1

                            End If
                            ' End If

                            '*** Summary of Reassayed Samples ***
                        ElseIf (StrComp(dgvReportTables.CurrentCell.Value, "Summary of Reassayed Samples") = 0) Then
                            Select Case charConfigTitle
                                Case "OriginalConc."
                                    .ToolTipText = "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR
                                Case "ReassayConc."
                                    .ToolTipText = "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR
                                Case "ReportedConc."
                                    .ToolTipText = "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR

                            End Select
                            'End If

                            '*** Summary of Repeat Samples ***
                        ElseIf (StrComp(dgvReportTables.CurrentCell.Value, "Summary of Repeat Samples") = 0) Then
                            Select Case charConfigTitle
                                Case "Dup 1 Result"
                                    .ToolTipText = "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR
                                Case "Dup 2 Result"
                                    .ToolTipText = "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR
                                Case "Mean of Dups"
                                    .ToolTipText = "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR
                                Case "Orig Result"
                                    .ToolTipText = "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR
                                Case "Reported Result"
                                    .ToolTipText = "Concentration Units automatically added to column header by StudyDoc" & ChrW(10) & strCR
                            End Select

                        Else

                            .ToolTipText = strCR

                        End If

                    End With
                End If
            End If
        Catch ex As Exception

        End Try
        
    End Sub


    Private Sub rbRoundFiveEven_CheckedChanged(sender As Object, e As EventArgs) Handles rbRoundFiveEven.CheckedChanged

        ''legend
        'Public gboolRoundFiveEven As Boolean = False
        'Public gboolRoundFiveAway As Boolean = True

        'Public gboolCritFullPrec As Boolean = False
        'Public gboolCritRounded As Boolean = True

        'Public gboolMeanFullPrec As Boolean = False
        'Public gboolMeanRounded As Boolean = True

        Dim boolA As Boolean
        boolA = Me.rbRoundFiveEven.Checked

        gboolRoundFiveEven = boolA
        gboolRoundFiveAway = Not (boolA)


    End Sub

    Private Sub rbRoundFiveAway_CheckedChanged(sender As Object, e As EventArgs) Handles rbRoundFiveAway.CheckedChanged

        ''legend
        'Public gboolRoundFiveEven As Boolean = False
        'Public gboolRoundFiveAway As Boolean = True

        'Public gboolCritFullPrec As Boolean = False
        'Public gboolCritRounded As Boolean = True

        'Public gboolMeanFullPrec As Boolean = False
        'Public gboolMeanRounded As Boolean = True

        Dim boolA As Boolean
        boolA = Me.rbRoundFiveAway.Checked

        gboolRoundFiveAway = boolA
        gboolRoundFiveEven = Not (boolA)

    End Sub

    Private Sub rbCritFullPrec_CheckedChanged(sender As Object, e As EventArgs) Handles rbCritFullPrec.CheckedChanged

        ''legend
        'Public gboolRoundFiveEven As Boolean = False
        'Public gboolRoundFiveAway As Boolean = True

        'Public gboolCritFullPrec As Boolean = False
        'Public gboolCritRounded As Boolean = True

        'Public gboolMeanFullPrec As Boolean = False
        'Public gboolMeanRounded As Boolean = True

        Dim boolA As Boolean
        boolA = Me.rbCritFullPrec.Checked

        gboolCritFullPrec = boolA
        gboolCritRounded = Not (boolA)

    End Sub

    Private Sub rbCritRounded_CheckedChanged(sender As Object, e As EventArgs) Handles rbCritRounded.CheckedChanged

        ''legend
        'Public gboolRoundFiveEven As Boolean = False
        'Public gboolRoundFiveAway As Boolean = True

        'Public gboolCritFullPrec As Boolean = False
        'Public gboolCritRounded As Boolean = True

        'Public gboolMeanFullPrec As Boolean = False
        'Public gboolMeanRounded As Boolean = True

        Dim boolA As Boolean
        boolA = Me.rbCritRounded.Checked

        gboolCritRounded = boolA
        gboolCritFullPrec = Not (boolA)

    End Sub

    Private Sub rbMeanFullPrec_CheckedChanged(sender As Object, e As EventArgs) Handles rbMeanFullPrec.CheckedChanged

        ''legend
        'Public gboolRoundFiveEven As Boolean = False
        'Public gboolRoundFiveAway As Boolean = True

        'Public gboolCritFullPrec As Boolean = False
        'Public gboolCritRounded As Boolean = True

        'Public gboolMeanFullPrec As Boolean = False
        'Public gboolMeanRounded As Boolean = True

        Dim boolA As Boolean
        boolA = Me.rbMeanFullPrec.Checked

        gboolMeanFullPrec = boolA
        gboolMeanRounded = Not (boolA)

    End Sub

    Private Sub rbMeanRounded_CheckedChanged(sender As Object, e As EventArgs) Handles rbMeanRounded.CheckedChanged

        ''legend
        'Public gboolRoundFiveEven As Boolean = False
        'Public gboolRoundFiveAway As Boolean = True

        'Public gboolCritFullPrec As Boolean = False
        'Public gboolCritRounded As Boolean = True

        'Public gboolMeanFullPrec As Boolean = False
        'Public gboolMeanRounded As Boolean = True

        Dim boolA As Boolean
        boolA = Me.rbMeanRounded.Checked

        gboolMeanRounded = boolA
        gboolMeanFullPrec = Not (boolA)

    End Sub


    Private Sub chkTableGraphicExamples_CheckedChanged(sender As Object, e As EventArgs) Handles chkTableGraphicExamples.CheckedChanged

        Dim h1, h2, h3, h4, t1, t2

        Dim dgv As DataGridView = Me.dgvReportTableConfiguration
        Dim pan As Panel = Me.panTableGraphicExamples
        Dim chk As System.Windows.Forms.CheckBox = Me.chkTableGraphicExamples
        Dim intPanTopMargin As Short = 10

        h1 = Me.tp6.Height

        If (chk.Checked) Then

            ''first put pan to sync with dgv bottom
            'h2 = h1 - (dgv.Top + dgv.Height)
            'pan.Top = h2 + pan.Height

            'Note: Setting pan.top isn't working, probably because of anchor conflicts
            'Instead, place bottom of pan manually in design

            ''size pan to grid left and width
            pan.Left = dgv.Left
            pan.Width = dgv.Width

            'now set height of dgv
            t1 = dgv.Top
            t2 = pan.Top

            h4 = t2 - t1
            dgv.Height = h4 - intPanTopMargin

            Call setTableGraphicExample()

            'now make visible
            pan.Visible = True

        Else

            'dgv.Height = dgv.Height + 400  'We need better way to do this - NDL
            pan.Visible = False

            'return dgv.h to original
            h1 = Me.tp6.Height
            dgv.Height = h1 - dgv.Top

        End If

    End Sub

    Private Sub tp8_Click(sender As Object, e As EventArgs) Handles tp8.Click

    End Sub

    Private Sub lblARS_Click(sender As Object, e As EventArgs) Handles lblARS.Click

    End Sub

    Private Sub Label53_Click(sender As Object, e As EventArgs) Handles Label53.Click

    End Sub

    Private Sub lblRTC_Click(sender As Object, e As EventArgs) Handles lblRTC.Click

    End Sub

    Private Sub tp12_Click(sender As Object, e As EventArgs) Handles tp12.Click

    End Sub

    Private Sub PrepareEntireReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrepareEntireReportToolStripMenuItem.Click
        cbxExampleReport.SelectedIndex = 2
    End Sub

    Private Sub PrepareOnlySelectedTableToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrepareOnlySelectedTableToolStripMenuItem.Click
        cbxExampleReport.SelectedIndex = 4
    End Sub


    Private Sub PrepareOnlyReportBodyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrepareOnlyReportBodyToolStripMenuItem.Click
        cbxExampleReport.SelectedIndex = 6
    End Sub


    Private Sub PrepareOnlyReportTablesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrepareOnlyReportTablesToolStripMenuItem.Click
        cbxExampleReport.SelectedIndex = 8
    End Sub

    Private Sub lblTCH_01_Click(sender As Object, e As EventArgs) Handles lblTCH_01.Click

    End Sub

    Private Sub tp7_Click(sender As Object, e As EventArgs) Handles tp7.Click

    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub cbxExampleReport_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbxExampleReport.SelectedValueChanged

    End Sub

    Private Sub cmdShowGroups_Click(sender As Object, e As EventArgs) Handles cmdShowGroups.Click

        Dim str1 As String
        Dim str2 As String
        Dim dgv As DataGridView = Me.dgvGroups

        str1 = Me.cmdShowGroups.Text

        If InStr(1, str1, "Show", CompareMethod.Text) > 0 Then
            str2 = "Hide _C[n] Groups"
            'dgv.Visible = True
            frmGroups.Show(Me)
            frmGroups.BringToFront()
        Else
            str2 = "Show _C[n] Groups"
            'dgv.Visible = False
            frmGroups.Dispose()
        End If

        Me.cmdShowGroups.Text = str2


    End Sub

  
    Private Sub dgvReportTableConfiguration_Scroll(sender As Object, e As ScrollEventArgs) Handles dgvReportTableConfiguration.Scroll

        Dim dgv As DataGridView = Me.dgvReportTableConfiguration


    End Sub

    Private Sub dgvSummaryData_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSummaryData.CurrentCellDirtyStateChanged

        Dim dgv As DataGridView
        Dim dv As System.Data.DataView
        Dim str1 As String
        Dim intRow As Short
        Dim intCol As Short
        Dim bool As Boolean

        dgv = dgvSummaryData
        intCol = dgv.CurrentCell.ColumnIndex
        intRow = dgv.CurrentRow.Index
        str1 = dgv.Columns.Item(intCol).Name
        dv = dgv.DataSource
        If StrComp(str1, "boolI", CompareMethod.Text) = 0 Then
            dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
            'bool = dgv.Rows.Item(intRow).Cells(intCol).Value
            'dv(intRow).BeginEdit()
            'If bool Then
            '    dv(intRow).Item("boolInclude") = -1
            'Else
            '    dv(intRow).Item("boolInclude") = 0
            'End If
            'dv(intRow).EndEdit()

            Try
                bool = dgv.Rows.Item(intRow).Cells(intCol).Value
                dv(intRow).BeginEdit()
                If bool Then
                    dv(intRow).Item("boolInclude") = -1
                Else
                    dv(intRow).Item("boolInclude") = 0
                End If
                dv(intRow).EndEdit()
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub dgvSummaryData_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvSummaryData.MouseEnter

        dgvSummaryData.Focus()

    End Sub





    Private Sub tabData_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabData.SelectedIndexChanged

        Try
            Call ConfigDropDowDGVs()
        Catch ex As Exception

        End Try

        Try
            Call ResizeFC()
        Catch ex As Exception

        End Try


    End Sub

    Private Sub mCal1_DateChanged(sender As Object, e As DateRangeEventArgs) Handles mCal1.DateChanged

    End Sub


    Private Sub chkAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkAll.CheckedChanged

        Call ReportOptionChecks(True)

    End Sub

    Private Sub chkAccepted_CheckedChanged(sender As Object, e As EventArgs) Handles chkAccepted.CheckedChanged

        Call ReportOptionChecks(False)

    End Sub

    Private Sub chkRejected_CheckedChanged(sender As Object, e As EventArgs) Handles chkRejected.CheckedChanged

        Call ReportOptionChecks(False)

    End Sub

    Private Sub chkRegrPerformed_CheckedChanged(sender As Object, e As EventArgs) Handles chkRegrPerformed.CheckedChanged

        Call ReportOptionChecks(False)

    End Sub

    Private Sub chkNoRegrPerformed_CheckedChanged(sender As Object, e As EventArgs) Handles chkNoRegrPerformed.CheckedChanged

        Call ReportOptionChecks(False)

    End Sub

    Private Sub chkPSAE_CheckedChanged(sender As Object, e As EventArgs) Handles chkPSAE.CheckedChanged

        Call ReportOptionChecks(False)

    End Sub

    Sub ReportOptionChecks(boolAll As Boolean)

        If boolFormLoad Then
            Exit Sub
        End If

        If boolStopRBS Then
            Exit Sub
        End If

        Dim boolF As Boolean = boolFormLoad
        boolFormLoad = True

        'If boolAll Then
        '    Me.chkAccepted.Checked = False
        '    Me.chkRejected.Checked = False
        '    Me.chkRegrPerformed.Checked = False
        '    Me.chkNoRegrPerformed.Checked = False
        '    Me.chkPSAE.Checked = False
        'Else
        '    Me.chkAll.Checked = False
        'End If
        boolFormLoad = boolF

        Dim bool As Boolean

        boolFromAnalSum = True
        Call FillAnalRunSum()
        boolFromAnalSum = False

    End Sub

    Private Sub dgvReportTableHeaderConfig_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReportTableHeaderConfig.CellContentClick

    End Sub

    Private Sub dgvReportTableHeaderConfig_CellValidated(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReportTableHeaderConfig.CellValidated


        Dim strCol As String
        Dim dgv As DataGridView = Me.dgvReportTableHeaderConfig

        strCol = dgv.Columns(e.ColumnIndex).Name

        If StrComp(strCol, "CHARUSERLABEL", CompareMethod.Text) = 0 Then

            dgv.AutoResizeRow(e.RowIndex)

        End If

    End Sub


    Function ChangeStudy() As Boolean

        ChangeStudy = False

        Dim intR As Short
        Dim strM As String

        strM = "Do you wish to change studies?"

        If boolOpened Then
            intR = MsgBox(strM, MsgBoxStyle.YesNo)
        Else
            intR = 6
            boolOpened = True
        End If



        If intR = 6 Then
            ChangeStudy = True
        Else
            ChangeStudy = False
        End If


    End Function

    Private Sub dgvwStudy_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvwStudy.RowValidating

        If boolFormLoad Then
            Exit Sub
        End If

        Dim intR As Short
        Dim strM As String

        strM = "Do you wish to change studies?"

        Dim dgv As DataGridView = Me.dgvwStudy

        If dgv.Focused Then
            If ChangeStudy() Then
            Else
                e.Cancel = True
                GoTo end1
            End If
        End If

        Me.Button1.Text = e.RowIndex

        'Dim boolAllowed As Boolean = StudyAllowed()


        'strM = "Cancel row selection?"
        'intR = MsgBox(strM, vbYesNo, "Cancel?")

        'If intR = 6 Then
        '    e.Cancel = True
        'End If

end1:

    End Sub

    Private Sub txtFilterStudy_TextChanged(sender As Object, e As EventArgs) Handles txtFilterStudy.TextChanged

        Call ExecuteFilter()

    End Sub

 
    Private Sub dgvFC_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvFC.CurrentCellDirtyStateChanged

        'this code is for assigning unbound checkbox value to BOOLINCLUDE

        'https://social.msdn.microsoft.com/Forums/windows/en-US/d50d961f-e7f7-48aa-8166-51d01a4b9c62/how-to-capture-a-checkbox-value-in-a-datagridview-vbnet?forum=winformsdatacontrols

        Dim dgv As DataGridView = Me.dgvFC

        If dgv.ReadOnly Then
            Exit Sub
        End If

        Dim str1 As String
        Dim int1 As Short
        Dim intCol As Short
        Dim intRow As Short
        Dim var1, var2, var3
        Dim strBool As String = "BOOLINCLUDE"

        Try
            intCol = dgv.CurrentCell.ColumnIndex
            intRow = dgv.CurrentCell.RowIndex

            str1 = dgv.Columns(intCol).Name

            If StrComp(str1, "CHKINCLUDE", CompareMethod.Text) = 0 Then

                If dgv.IsCurrentCellDirty Then
                    dgv.CommitEdit(DataGridViewDataErrorContexts.Commit)
                End If

            End If

        Catch ex As Exception

            var3 = ex.Message
            'MsgBox("dgvFC_CurrentCellDirtyStateChanged:  " & var3)

        End Try



    End Sub

    Private Sub dgvFC_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvFC.CellValueChanged

        'this code is for assigning unbound checkbox value to BOOLINCLUDE

        'https://social.msdn.microsoft.com/Forums/windows/en-US/d50d961f-e7f7-48aa-8166-51d01a4b9c62/how-to-capture-a-checkbox-value-in-a-datagridview-vbnet?forum=winformsdatacontrols

        Dim dgv As DataGridView = Me.dgvFC

        If dgv.ReadOnly Then
            Exit Sub
        End If

        Dim str1 As String
        Dim int1 As Short
        Dim intCol As Short
        Dim intRow As Short
        Dim var1, var2, var3
        Dim strBool As String = "BOOLINCLUDE"

        Try

            intCol = e.ColumnIndex
            intRow = e.RowIndex

            str1 = dgv.Columns(intCol).Name

            If StrComp(str1, "CHKINCLUDE", CompareMethod.Text) = 0 Then

                'DataGridView1.Rows(i).Cells("CheckboxColumn").Value = DataGridView1.CurrentCell.Value
                var1 = dgv(intCol, intRow).Value

                If var1 Then
                    dgv.Rows(intRow).Cells("BOOLINCLUDE").Value = -1
                Else
                    dgv.Rows(intRow).Cells("BOOLINCLUDE").Value = 0
                End If


            End If

        Catch ex As Exception

            var3 = ex.Message
            'MsgBox("dgvFC_CellValueChanged:  " & var3)

        End Try

    End Sub

    Private Sub rbInclude_CheckedChanged(sender As Object, e As EventArgs) Handles rbInclude.CheckedChanged

        Call FilterFC()

    End Sub

    Private Sub dgvReportStatementWord_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReportStatementWord.CellContentClick

    End Sub

    Private Sub dgvReportTableConfiguration_KeyUp(sender As Object, e As KeyEventArgs) Handles dgvReportTableConfiguration.KeyUp

        Dim var1, var2, var3, var4

        Dim dgv As DataGridView = Me.dgvReportTableConfiguration

        If e.Control And e.KeyCode = Keys.C Then

            'check to see if entire grid is selected
            Dim intRows As Int16
            intRows = dgv.Rows.Count
            If dgv.SelectedRows.Count = intRows Then
                dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
                Clipboard.SetDataObject(dgv.GetClipboardContent())
                dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable
            End If

        End If

        var1 = e.KeyCode
        var2 = e.KeyData
        var3 = e.KeyValue

        var1 = var1

    End Sub

    Private Sub chkLockFinalReport_CheckedChanged(sender As Object, e As EventArgs) Handles chkLockFinalReport.CheckedChanged

        If AllowLockFinalReport() Then

            Dim intL As Short
            If Me.chkLockFinalReport.Checked Then
                BOOLFINALREPORTLOCKED = True
                intL = -1
            Else
                BOOLFINALREPORTLOCKED = False
                intL = 0
            End If

            'need to update tblFinalReport
            Dim strF As String
            Dim strS As String

            strF = "CHARREPORTTYPE = 'Final Report' AND ID_TBLSTUDIES = " & id_tblStudies
            strS = "UPSIZE_TS DESC"

            Dim rows() As DataRow = tblFinalReport.Select(strF, strS, DataViewRowState.CurrentRows)
            If rows.Length = 0 Then
                Me.lblFinalReportLockedDate.Text = "NA"
            Else
                rows(0).BeginEdit()
                rows(0).Item("BOOLLOCKED") = intL
                rows(0).EndEdit()

                'enter label information
                Try
                    Dim dt As Date = rows(0).Item("UPSIZE_TS")
                    Dim str1 As String = LTextDateFormat & " HH:mm:ss tt"
                    Dim strDt As String = Format(dt, str1)
                    Me.lblFinalReportLockedDate.Text = strDt

                Catch ex As Exception
                    Me.lblFinalReportLockedDate.Text = "NA"
                End Try

            End If

        End If


    End Sub

    Sub ConfigAnalyteOrder()

        Dim dgv As DataGridView = Me.dgvAnalyteGroups
        Dim dtbl As System.Data.DataTable = tblAnalytesHome
        Dim strS As String
        Dim strF As String
        strF = "IsIntStd = 'No'"
        strS = "INTORDER ASC"
        Dim dv As DataView = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)

        dv.AllowDelete = False
        dv.AllowEdit = True ' False
        dv.AllowNew = False

        dgv.DataSource = dv

        dgv.AllowUserToOrderColumns = False

        'Legends:
        'tblAnalyteGroups '"ANALYTEDESCRIPTION", "ANALYTEID", "INTSTD", "INTGROUP", "ANALYTEDESCRIPTION_C", "MATRIX", "INTCALSET", "CALIBRSET", "INTORDER"

        'tblAnalytesHome:
        'Select Case Count1
        '    Case 1
        '        str1 = "AnalyteDescription"
        '        var1 = System.Type.GetType("System.String")
        '    Case 2
        '        str1 = "AnalyteID"
        '        var1 = System.Type.GetType("System.Int64")
        '    Case 3
        '        str1 = "AnalyteIndex"
        '        var1 = System.Type.GetType("System.Int64")
        '    Case 4
        '        str1 = "BQL"
        '        var1 = System.Type.GetType("System.Single")
        '    Case 5
        '        str1 = "AQL"
        '        var1 = System.Type.GetType("System.Single")
        '    Case 6
        '        str1 = "ConcUnits"
        '        var1 = System.Type.GetType("System.String")
        '    Case 7
        '        str1 = "AcceptedRuns"
        '        var1 = System.Type.GetType("System.Int64")
        '    Case 8
        '        str1 = "IsReplicate"
        '        var1 = System.Type.GetType("System.String")
        '    Case 9
        '        str1 = "IsIntStd"
        '        var1 = System.Type.GetType("System.String")
        '    Case 10
        '        str1 = "UseIntStd"
        '        var1 = System.Type.GetType("System.String")
        '    Case 11
        '        str1 = "IntStd"
        '        var1 = System.Type.GetType("System.String")
        '    Case 12
        '        str1 = "MasterAssayID"
        '        var1 = System.Type.GetType("System.Int64")
        '    Case 13
        '        str1 = "IsCoadminCmpd"
        '        var1 = System.Type.GetType("System.String")
        '    Case 14
        '        str1 = "ORIGINALANALYTEDESCRIPTION"
        '        var1 = System.Type.GetType("System.String")
        '    Case 15
        '        str1 = "INTGROUP"
        '        var1 = System.Type.GetType("System.Int16")
        '    Case 16
        '        str1 = "MATRIX"
        '        var1 = System.Type.GetType("System.String")
        '    Case 17
        '        str1 = "INTORDER"
        '        var1 = System.Type.GetType("System.Int16")
        '    Case 18
        '        str1 = "CALIBRSET"
        '        var1 = System.Type.GetType("System.String")


        '    Case 19
        '        str1 = "CHARUSERANALYTE"
        '        var1 = System.Type.GetType("System.String")
        '    Case 120
        '        str1 = "CHARUSERIS"
        '        var1 = System.Type.GetType("System.String")

        'End Select

        'make all columns invisible

        Dim Count1 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String

        Dim var1

        Dim intI1 As Short = 0
        Dim intI2 As Short = 0

        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        Dim boolC As Boolean
        Dim boolI As Boolean
        Dim Count2 As Short

        Dim int1 As Short
        Dim int2 As Short
        Dim int3 As Short

        For Count2 = 1 To 2 'must iterate

            Try
                For Count1 = 1 To 7
                    boolC = False
                    boolI = False
                    Select Case Count1
                        Case 1
                            str1 = "ANALYTEDESCRIPTION"
                            str2 = "Analyte"
                            int1 = dgv.Columns(str1).DisplayIndex
                        Case 2
                            str1 = "CHARUSERANALYTE"
                            str2 = "Analyte" & ChrW(10) & "User Name"
                            boolI = True
                            intI1 = int1 + 1
                        Case 3
                            str1 = "INTSTD"
                            str2 = "Assigned" & ChrW(10) & "Int. Std."
                            int2 = dgv.Columns(str1).DisplayIndex
                        Case 4
                            str1 = "CHARUSERIS"
                            str2 = "Int. Std." & ChrW(10) & "User Name"
                            boolI = True
                            intI1 = int2 + 1
                        Case 5
                            str1 = "MATRIX"
                            str2 = "Matrix"
                        Case 6
                            str1 = "INTORDER"
                            str2 = "Order"
                            boolC = True
                            boolI = True
                            intI1 = dgv.Columns(str1).DisplayIndex + 1
                        Case 7
                            'find Calibration Set
                            str1 = "CALIBRSET"
                            str2 = "Calibration Set (Calibration Level)"
                            boolI = True
                            intI2 = intI1 - 1

                    End Select

                    dgv.Columns(str1).HeaderText = str2
                    dgv.Columns(str1).Visible = True
                    dgv.Columns(str1).SortMode = DataGridViewColumnSortMode.NotSortable
                    If boolC Then
                        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                    Else
                        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                    End If

                    If intI1 = 0 Then
                    Else
                        If intI2 = 0 Then
                            If intI1 > dgv.ColumnCount - 1 Then
                            Else
                                dgv.Columns(str1).DisplayIndex = intI1
                            End If
                        Else
                            dgv.Columns(str1).DisplayIndex = intI2
                        End If
                    End If

                    'dgv.Columns("INTGROUP").Visible = True'debug

                Next
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

        Next

        'select first row
        dgv.Rows(0).Selected = True

        'autosize columns
        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AutoResizeColumns()


    End Sub

    Private Sub cmdUpA_Click(sender As Object, e As EventArgs) Handles cmdUpA.Click

        Call DoUpDownClick(True)

    End Sub

    Private Sub cmdDownA_Click(sender As Object, e As EventArgs) Handles cmdDownA.Click

        Call DoUpDownClick(False)

    End Sub

    Sub DoUpDownClick(boolUp As Boolean)

        Dim dgv As DataGridView = Me.dgvAnalyteGroups
        Dim dv As DataView = dgv.DataSource
        Dim dtbl As System.Data.DataTable = tblAnalytesHome
        Dim strF As String

        'must change values in dtbl

        Dim intCrit As Short
        If boolUp Then
            intCrit = 0
        Else
            intCrit = dgv.RowCount - 1
        End If

        Dim intRow As Int16
        Dim intO1 As Int16
        Dim intO2 As Int16
        Try
            intRow = dgv.CurrentRow.Index
            If intRow = intCrit Then
            Else
                intO1 = dgv("INTORDER", intRow).Value
                If boolUp Then
                    intO2 = intO1 - 1
                Else
                    intO2 = intO1 + 1
                End If



                strF = "INTORDER = " & intO1 & " OR INTORDER = " & intO2
                Dim rows() As DataRow = dtbl.Select(strF, "INTORDER ASC")

                If boolUp Then

                    rows(0).BeginEdit()
                    rows(0).Item("INTORDER") = intO1
                    rows(0).EndEdit()

                    rows(1).BeginEdit()
                    rows(1).Item("INTORDER") = intO2
                    rows(1).EndEdit()

                    'select  row
                    dgv.Rows(intRow - 1).Selected = True

                Else

                    rows(0).BeginEdit()
                    rows(0).Item("INTORDER") = intO2
                    rows(0).EndEdit()

                    rows(1).BeginEdit()
                    rows(1).Item("INTORDER") = intO1
                    rows(1).EndEdit()

                    dgv.Rows(intRow + 1).Selected = True

                End If

                Call ReorderAnalytes()

                'now reorder Report Tables
                Cursor.Current = Cursors.WaitCursor
                Call FillTableReports(True)
                Cursor.Current = Cursors.WaitCursor
                Call FillTableReportsAnalytes(True)
                Cursor.Current = Cursors.WaitCursor
                Call FillTableReportDataAnalytes(True)
                Call RTFilter()
                Cursor.Current = Cursors.Default

            End If


        Catch ex As Exception

        End Try


    End Sub

    Private Sub cmdUpCF_Click(sender As Object, e As EventArgs) Handles cmdUpCF.Click

        Call DoUpDownClickCF(True)

    End Sub

    Private Sub cmdDownCF_Click(sender As Object, e As EventArgs) Handles cmdDownCF.Click

        Call DoUpDownClickCF(False)

    End Sub

    Sub DoUpDownClickCF(boolUp As Boolean)

        Dim dgv As DataGridView = Me.dgvFC
        Dim dv As DataView = dgv.DataSource
        Dim dtbl As System.Data.DataTable = tblCustomFieldCodes
        Dim strF As String
        Dim var1

        'must change values in dtbl

        Dim intCrit As Short
        If boolUp Then
            intCrit = 0
        Else
            intCrit = dgv.RowCount - 1
        End If

        Dim intRow As Int16
        Dim intO1 As Int16
        Dim intO2 As Int16
        Try
            intRow = dgv.CurrentRow.Index
            If intRow = intCrit Then
            Else
                intO1 = dgv("INTORDER", intRow).Value
                If boolUp Then
                    intO2 = intO1 - 1
                Else
                    intO2 = intO1 + 1
                End If


                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND (INTORDER = " & intO1 & " OR INTORDER = " & intO2 & ")"
                Dim rows() As DataRow = dtbl.Select(strF, "INTORDER ASC")

                If rows.Length = 0 Then
                    var1 = var1
                End If

                If boolUp Then

                    rows(0).BeginEdit()
                    rows(0).Item("INTORDER") = intO1
                    rows(0).EndEdit()

                    rows(1).BeginEdit()
                    rows(1).Item("INTORDER") = intO2
                    rows(1).EndEdit()

                    'select  row
                    dgv.Rows(intRow - 1).Selected = True

                Else

                    rows(0).BeginEdit()
                    rows(0).Item("INTORDER") = intO2
                    rows(0).EndEdit()

                    rows(1).BeginEdit()
                    rows(1).Item("INTORDER") = intO1
                    rows(1).EndEdit()

                    dgv.Rows(intRow + 1).Selected = True

                End If


            End If


        Catch ex As Exception
            var1 = ex.Message
        End Try


    End Sub

    Private Sub dgvFC_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvFC.CellContentClick

    End Sub

    Private Sub cmdApplyTables_Click(sender As Object, e As EventArgs) Handles cmdApplyTables.Click

        Dim frm As New frmApplyTemplateTables

        frm.ShowDialog()

        frm.Dispose()

    End Sub

    Private Sub cmdClearStudy_Click(sender As Object, e As EventArgs) Handles cmdClearStudy.Click

        Dim frm As New frmClearStudy

        frm.ShowDialog()

        frm.Dispose()

    End Sub



    Private Sub dgvAnalyteGroups_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAnalyteGroups.CellContentClick

    End Sub

    Private Sub dgvAnalyteGroups_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles dgvAnalyteGroups.Validating


    End Sub

End Class









