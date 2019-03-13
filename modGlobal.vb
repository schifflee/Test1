Option Compare Text

Imports Word = Microsoft.Office.Interop.Word
Imports System
Imports System.IO
Imports System.Text


Module modGlobal


    Public boolAuditTrail As Boolean = True

    Public gboolAuditTrail As Boolean = False
    Public gboolESig As Boolean = False
    Public strRFC As String = ""
    Public strMOS As String = ""
    Public strAuditType As String = ""
    Public gATAdds As Int16 = 0
    Public gATMods As Int16 = 0
    Public gATDeletes As Int16 = 0

    Public constrWatson As String
    Public constrIni As String
    Public constrIniGuWuODBC As String
    Public constrA As String 'connection string for MS Access
    Public constrW As String 'connection string for Watson
    Public constrCur As String 'current connection string
    Public GSigFig As Short = 3 'Global Significant Figure value
    Public LSigFig As Short = 3 'Local Significant Figure value for each study
    Public GboolWyethRounding As Boolean = False
    Public LboolWyethRounding As Boolean = False
    Public GRegrSigFigs As Short = 5
    Public LRegrSigFigs As Short = 5
    Public GR2SigFigs As Short = 5
    Public LR2SigFigs As Short = 5
    Public GRegrDec As Short = 5
    Public LRegrDec As Short = 5
    Public GR2Dec As Short = 5
    Public LR2Dec As Short = 5
    Public strRegrDec As String = "0.00000"
    Public strR2Dec As String = "0.00000"
    Public boolGUseRegrSciNot As Boolean = True
    Public boolLUseRegrSciNot As Boolean = True

    Public LboolNomConcParen As Boolean = False
    Public LcharSTPage As String = "[None]"

    Public LBOOLTABLEDTTIMESTAMP As Boolean = False
    Public LTableDateTimeStamp As Date

    Public GDateFormat As String = "MM/dd/yyyy" 'global default date format
    Public LDateFormat As String = "MM/dd/yyyy" 'local date format for each study
    Public LTextDateFormat As String = "MMMM dd, yyyy"
    Public GTextDateFormat As String = "MMMM dd, yyyy"
    Public GDec As Short = 1 'Global Decimal Value
    Public LDec As Short = 1 'Local Decimal Value
    Public GIncDiff As String = "%Difference"
    Public LIncDiff As String = "%Difference"
    Public GTimeZone As String = "Eastern Time Zone"
    Public LTimeZone As String = "Eastern Time Zone"
    Public gReportTitle As String = ""

    Public GSigFigArea As Short = 3
    Public LSigFigArea As Short = 3
    Public GDecArea As Short = 3
    Public LDecArea As Short = 3
    Public GboolWyethRoundingArea As Boolean = False
    Public LboolWyethRoundingArea As Boolean = False
    Public strAreaDec As String = "0.000"

    Public GSigFigAreaRatio As Short = 5
    Public LSigFigAreaRatio As Short = 5
    Public GDecAreaRatio As Short = 5
    Public LDecAreaRatio As Short = 5
    Public GboolWyethRoundingAreaRatio As Boolean = False
    Public LboolWyethRoundingAreaRatio As Boolean = False
    Public strAreaDecAreaRatio As String = "0.00000"

    Public gUserName As String = ""
    Public gUserID As String = ""
    Public gUserLabel As String = " - User: Guest with Read Only permissions"
    Public gCaption As String
    Public MeCaption As String
    Public gWorkstation As String = ""
    Public gPswd As String
    Public tPswd As String = "" 'record doc protected password to save to database later

    Public gAllowExclSamples As Boolean = False
    Public gAllowGuWuAccCrit As Boolean = False
    Public LAllowExclSamples As Boolean = False
    Public LAllowGuWuAccCrit As Boolean = False

    Public boolEval As Boolean = False

    Public gstrAnal As String
    Public gnumAnal As Short = 1

    Public wd_app_RBS As Microsoft.Office.Interop.Word.Application
    Public wd_doc_RBS As Microsoft.Office.Interop.Word.Document

    Public boolGUseSigFigs As Boolean = True 'Global use SigFig or Decimal
    Public boolLUseSigFigs As Boolean = True 'Local use SigFig or Decimal

    Public boolGUseSigFigsArea As Boolean = True 'Global use SigFig or Decimal
    Public boolLUseSigFigsArea As Boolean = True 'Local use SigFig or Decimal

    Public boolGUseSigFigsAreaRatio As Boolean = True 'Global use SigFig or Decimal
    Public boolLUseSigFigsAreaRatio As Boolean = True 'Local use SigFig or Decimal

    Public boolGUseSigFigsRegr As Boolean = True 'Global use SigFig or Decimal
    Public boolLUseSigFigsRegr As Boolean = True 'Local use SigFig or Decimal

    Public boolUseHyperlinks As Boolean = True


    Public tbl_dgHome As New System.Data.DataTable
    Public tblWatsonData As New System.Data.DataTable
    Public tblCompanyAnalRefTable As System.Data.DataTable
    Public tblWatsonAnalRefTable As System.Data.DataTable
    Public tblCompanyData As New System.Data.DataTable
    Public tblAnalRunSum As New System.Data.DataTable
    'Public tblStudies As new System.Data.DataTable
    Public tblStudiesL As New System.Data.DataTable 'tblStudies Local
    Public tblwSTUDY As New System.Data.DataTable
    Public tblASTUDY As New System.Data.DataTable
    Public tblwPROJECTS As New System.Data.DataTable

    'Public tblReports As new System.Data.DataTable
    Public tblMethodValData As New System.Data.DataTable
    Public rsStudies As New ADODB.Recordset
    Public tblReportTables As New System.Data.DataTable
    'Public tblReportStatementsGuWu As New System.Data.DataTable
    Public tblReportStatementsStore As New System.Data.DataTable
    Public tblReportTHeaderConfig As New System.Data.DataTable
    Public tblMethValExistingGuWu As New System.Data.DataTable
    Public tblMethValExistingStore As New System.Data.DataTable
    Public tblAnalysisResultsHome As New System.Data.DataTable 'use in Assign Samples
    Public tblAnalysisResultsHomeOutStudy As New System.Data.DataTable 'use in Assign Samples out study 20180801
    Public tblAnalytesHome As New System.Data.DataTable 'to store study specific analytes
    Public tblAnalyteIDs As New System.Data.DataTable 'to store a list of unique AnalyteIDs in the Study
    Public tblMatrices As New System.Data.DataTable 'to store list of unique Matrices (Matrix) in the Study
    Public tblAssayLabels As New System.Data.DataTable 'to store QC and CalStd Labels, possibly more later
    'Public tblUserAccounts As New System.Data.DataTable 'to find customer account info
    Public tblAnalU As New System.Data.DataTable 'Unique analytes

    Public tblISR As New System.Data.DataTable 'for ISR tables'DEPRECATED
    Public tblISRUnique As New System.Data.DataTable 'for ISR tables
    'Public tblISR01_01 As New System.Data.DataTable 'for ISR tables

    Public tblHeaderLabels As New System.Data.DataTable
    Public tblQATableTemp As New System.Data.DataTable
    Public tblCP As New System.Data.DataTable
    Public tblTableN As New System.Data.DataTable
    Public tblSRecWatson As New System.Data.DataTable

    Public tblAppendix As New System.Data.DataTable
    Public tblFigures As New System.Data.DataTable
    Public tblAttachments As New System.Data.DataTable

    Public tblQCConcs As New System.Data.DataTable 'used for displaying QC data
    Public tblSampleConcs As New System.Data.DataTable 'used for display sample
    Public tblReassay As New System.Data.DataTable
    Public tblReassayReport As New System.Data.DataTable
    Public tblReassayReasons As New System.Data.DataTable
    Public tblSAMPLERESULTSCONFLICT As New System.Data.DataTable
    Public tblSampleMatrix As New System.Data.DataTable
    Public tblAAUnk As New System.Data.DataTable
    Public tblAAUnkRunID As New System.Data.DataTable

    Public tblGetDecRunID As New System.Data.DataTable
    Public tblSAMPRESCONFLICTDEC As New System.Data.DataTable

    Public tblRepeatAllRunSamples As New System.Data.DataTable

    Public tblReportCompanies As New System.Data.DataTable 'to record report body type company ids
    Public tblHook1 As New System.Data.DataTable
    Public tblHook2 As New System.Data.DataTable
    Public tblHook3 As New System.Data.DataTable
    Public tblHook4 As New System.Data.DataTable
    Public tblARHTemp As New System.Data.DataTable

    Public tblConcLevelsForAssayIDs As New DataTable

    Public ctAppendix As Short = 0
    Public ctTableN As Short = 0
    Public ctFigures As Short
    Public ctAttachments As Short = 0

    Public wStudyID As Object
    Public wSpeciesID As Object
    Public wProjectID As Object
    Public wWStudyName As Object
    Public id_tblStudies As Int64 = 0
    Public id_tblReports As Int64 = 0
    Public id_tblConfigReportType As Int64 = -1
    Public id_tblPersonnel As Int64
    Public id_tblUserAccounts As Int64
    Public id_tblPermissions As Int64
    Public ctAnalytes As Short
    Public ctAnalytes_IS As Short
    Public arrAnalytes(15, 100) As Object '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
    '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns, 8=IsReplicate, 9=IsIntStd
    '10=UseIntStd, 11=IntStd, 12=MasterAssayID, 13=IsCoadministeredCmpd 14=Original AnalyteDescription,15=Group
    Public arrAnalytesCB(50) As String 'used to populate datagridview analyte comboboxes
    Public arrctQCs(3, 50) As Object '1=AnalyteIndex, 2=MasterAssayID, 3=ctanalytes
    Public arrQCReps(4, 50) As Object
    '1=AnalyteID(arrAnalyte(2,n)), 2=replicate count of non-dilution QCs, 3=number of non-dilution QC levels, 4=replicate count of dilution QCs
    Public ctAnalyticalRuns As Short
    Public ctCalibrStds As Short
    Public strSponsor As String
    Public strCompany As String
    Public strInSupportOf As String
    Public intILS As Short = 0
    Public boolCont As Boolean
    Public boolLoad As Boolean = True
    Public boolFormLoad As Boolean = True
    Public boolFromCAR As Boolean = False
    Public boolFromCD As Boolean = False
    Public boolFromRTC As Boolean = False
    Public boolFromcbxStudy As Boolean = False
    Public boolFromdgvwStudy As Boolean = False
    Public boolRefresh As Boolean = False
    Public boolReportCont As Boolean
    Public boolShowExample As Boolean = False
    Public boolANSI As Boolean = False
    Public boolStopCBX As Boolean = True
    Public boolGuWuAccess As Boolean = False
    Public boolGuWuSQLServer As Boolean = False
    Public boolGuWuOracle As Boolean = False
    Public boolMsg As Boolean = False
    Public boolDirty As Boolean = False
    Public boolOR As Boolean = False 'ApplyTemplate must be able to save all stuff
    Public boolRSCFill As Boolean = False 'do most of ReportStatementsFillCharSection on GetStudyInfo action
    Public boolHook1 As Boolean = True 'true if hook can continue
    Public boolHook2 As Boolean = True 'true if hook can continue
    Public boolHook3 As Boolean = True 'true if hook can continue
    Public boolHook4 As Boolean = True 'true if hook can continue
    Public boolEventAdd As Boolean = False 'for dgvReports event handler
    Public boolStopRBS As Boolean = False 'for cellvalidating action on dgvReportStatements
    Public boolcbxExample As Boolean = False
    Public boolQuickFind As Boolean = True
    Public boolQFDone As Boolean = True

    Public boolMeRefresh As Boolean = True
    Public boolFromAnalSum As Boolean = False
    Public boolHomeCBox As Boolean = True
    Public boolArchiveSource As Boolean = False

    Public boolBQLSHOWCONC As Boolean = True
    Public boolCSSHOWREJVALUES As Boolean = True
    Public boolCSREPORTACCVALUES As Boolean = True
    Public boolQCREPORTACCVALUES As Boolean = True
    Public boolSTATSMEAN As Boolean = True
    Public boolSTATSSD As Boolean = True
    Public boolSTATSCV As Boolean = True
    Public boolSTATSBIAS As Boolean = True
    Public boolSTATSN As Boolean = True
    Public BOOLSTATSRE As Boolean = False
    Public boolSTATSDIFF As Boolean = False
    Public boolSTATSDIFFCOL As Boolean = False 'this is the checkbox to add a %Accuracy column. From Advanced Table Config window
    Public BOOLDIFFCOLSTATS As Boolean = False 'this is if Reported Stats Section is ave of StatsDiffCol (if shown)
    Public boolSTATSREGR As Boolean = False
    Public boolSTATSNR As Boolean = True
    Public boolSTATSLETTER As Boolean = False
    Public boolTHEORETICAL As Boolean = False
    Public boolINCLANOVA As Boolean = True
    Public boolINCLANOVASUMSTATS As Boolean = True
    Public boolBQLLEGEND As Boolean = False
    Public boolIncludePSAE As Boolean = False
    Public boolRCConc As Boolean = True
    Public boolRCPA As Boolean = True
    Public boolRCPARatio As Boolean = False
    Public boolIncludeISTbl As Boolean = True
    Public boolNONELEG As Boolean = True
    Public boolPOSLEG As Boolean = True
    Public boolNEGLEG As Boolean = False
    Public boolCUSTOMLEG As Boolean = False
    Public boolMEANACCURACY As Boolean = False
    Public boolRECOVERY As Boolean = False
    Public BOOLDIFFERENCE As Boolean = False
    Public BOOLINCLUDEDATE As Boolean = False
    Public BOOLINCLUDEWATSONLABELS As Boolean = False
    Public BOOLINTRARUNSUMSTATS As Boolean = False
    Public BOOLQAEVENTBORDER As Boolean = False
    Public BOOLTEMPLATEFIELDCODESLOADED As Boolean = False  'Report Template Field Codes loaded

    Public RBPosCov As Int64
    Public RBPos1 As Int64
    Public RBPos2 As Int64
    Public EOCOV As Int64
    Public EOTOC As Int64
    Public EOTOT As Int64
    Public EOTOA As Int64
    Public EOTOF As Int64

    Public oldCurrentRowRS As Short = 0

    Public oldCurrentRowCAR As Short = -1
    Public oldCurrentColCAR As Short = -1
    Public newCurrentRowCAR As Short = -1
    Public newCurrentColCAR As Short = -1
    Public oldCurrentCellCAR As Object
    Public newCurrentCellCAR As Object

    Public oldCurrentRowRTC As Short = -1
    Public oldCurrentColRTC As Short = -1
    Public newCurrentRowRTC As Short = -1
    Public newCurrentColRTC As Short = -1
    Public oldCurrentCellRTC As Object
    Public newCurrentCellRTC As Object

    Public boolRTCEnter As Boolean
    Public varR As String

    Public boolOKtoVal As Boolean = True
    Public boolOKtoValCD As Boolean = True
    Public valErr As String = "No error."
    Public valErrCD As String = "No error."
    Public tblBCStds As New System.Data.DataTable
    Public tblBCStdsAll As New System.Data.DataTable
    Public tblBCStdsAssayID As New System.Data.DataTable
    Public tblBCStdsAssayIDAll As New System.Data.DataTable
    Public tblBCQCStdsAll As New System.Data.DataTable
    Public tblBCStdConcs As New System.Data.DataTable
    Public tblBCStdConcsNew As New System.Data.DataTable
    Public tblBCQCStdsAssayID As New System.Data.DataTable
    Public tblQCStds As New System.Data.DataTable
    Public tblBCQCs As New System.Data.DataTable
    Public tblBCQCsAssayID As New System.Data.DataTable
    Public tblBCQCConcs As New System.Data.DataTable
    Public tblBCQCConcsNew As New System.Data.DataTable
    Public tblQCRunIDs As New System.Data.DataTable
    Public tblQCAI As New System.Data.DataTable
    Public tblQCReps As New System.Data.DataTable
    Public tblQCF As New System.Data.DataTable
    Public tblRegCon As New System.Data.DataTable
    Public tblRegConAll As New System.Data.DataTable
    Public tblFindNomConc As New System.Data.DataTable

    Public tblAccAnalRuns As New System.Data.DataTable
    Public tblAllAnalRuns As New System.Data.DataTable

    Public tblASSAYREPS As New System.Data.DataTable


    Public tblSampleDesign As New System.Data.DataTable
    Public arrReportNA(5, 200) As Object '1=SectionName, 2=Value, 3=Tab, 4=Grid, 5=Field Code
    Public ctArrReportNA As Short 'for recording Report NAs
    Public intErrCount As Short 'for counting Report Body errors
    Public strErrMsg As String 'for reporting Report Body errors
    Public boolAppendix As Boolean = False 'T/F if there are appendices to add
    Public ctPB As Short 'progress bar counter
    Public ctPBMax As Short 'progress bar max
    Public intCoverPageStyle As Short
    Public intRefStdStyle As Short
    Public intStartTable As Short = 2 'table# to start TofC when generating a report
    Public cbxDateFormat As New DataGridViewComboBoxCell
    Public cbxCompanyID As New DataGridViewComboBoxCell
    Public cbxxCPPrefix As New DataGridViewComboBoxCell
    Public cbxxCPName As New DataGridViewComboBoxCell
    Public cbxxCPSuffix As New DataGridViewComboBoxCell
    Public cbxxCPDegree As New DataGridViewComboBoxCell
    Public cbxxCPTitle As New DataGridViewComboBoxCell
    Public cbxxCPRole As New DataGridViewComboBoxCell
    Public cbxAnalytes As New DataGridViewComboBoxCell
    Public cbxxReportTemplates As New DataGridViewComboBoxCell
    Public cbxxReportTypes As New DataGridViewComboBoxCell
    Public cbxxAssayDescr As New DataGridViewComboBoxCell
    Public cbxxIncSmplDiff As New DataGridViewComboBoxCell
    Public cbxxAnalMethType As New DataGridViewComboBoxCell
    Public Sw, Sh, Sl, St As Single
    Public strSchema As String = "WATSON"
    Public boolTempANSI As Boolean
    Public conAccess97 As String

    Public boolDoFormulas As Boolean = True
    Public boolDoHyperlinks As Boolean = True

    Public strPathWd As String

    Public NormalFontsize 'records the normal fontsize

    Public arrRBSColumns(50, 50)
    Public boolEntireReport As Boolean = False

    Public strReportTypeApp As String = "" 'for report history section capture
    Public boolDoTables As Boolean = False

    Public intOTables As Short = 0 'the original number of tables in a report template

    Public id_tblGuWuStudies As Int64

    Public pArchivePath As String = ""

    Public boolDemo As Boolean

    Public arrGroups(3, 4)
    '1=Column, 2=Sort, 3=Field
    Public intGroups As Short = 1
    Public arrSort(3, 6)
    '1=Column, 2=Sort, 3=Field
    Public intSort As Short = 1

    Public wdAbort As Microsoft.Office.Interop.Word.Application

    Public intQCDec As Short = 1
    Public strQCDec As String = "0.0"
    Public gintQCDec As Short = 1

    Public tblSpeciesMatrix As New System.Data.DataTable
    Public intNumSpecies As Short
    Public intNumMatrix As Short
    Public tblSpeciesMatrixSV As New System.Data.DataTable 'has STANDARDVOLUME

    'this datatable will store statistics items for QC tables
    Public tblQCTables As New System.Data.DataTable
    Public tblStabilityTables As New System.Data.DataTable
    Public tblANOVATables As New System.Data.DataTable

    Public gdtSave As Date = Now
    Public tblAuditTrailTemp As New System.Data.DataTable

    Public idSE As Double
    Public intLBXTabPos As Short

    Public boolFromRW As Boolean = False
    Public ID_QATEMPID As Int64 = 0

    Public gINTCOMMAFORMAT As Int64 = 0
    Public boolBLUEHYPERLINK As Boolean = True
    Public boolTOC As Boolean = False
    Public boolTOT As Boolean = False
    Public boolTOF As Boolean = False
    Public boolTOA As Boolean = False

    Public gboolDisplayAttachments As Boolean = False
    Public gConfigStudy As String = ""
    Public gDoPDF As Boolean = False
    Public gboolReadOnlyTables As Boolean = False
    Public garrMargins

    Public ftNormal As String = "Times New Roman"
    Public boolRedBoldFont As Boolean = True

    Public boolVerbose As Boolean = False

    Public numQCLevels As Short
    Public numRepDilnQC As Short

    Public boolFootNoteQCMean As Boolean = True
    Public boolFlipHeader As Boolean = False

    Public boolQCNA As Boolean = True
    Public boolBQL As Boolean = True
    Public gstrBQL As String = "BQL/AQL"

    Public ctrsSamples As Object
    Public ctrsRepeat As Object
    Public ctrsReassayed As Object
    Public ctrsISR As Object

    Public boolDisableWarnings As Boolean = False
    Public boolIgnoreFC As Boolean = False

    Public boolSampleName01 As Boolean = False
    Public boolExcludeTableNumbers As Boolean = False
    Public boolExcludeTableTitles As Boolean = False
    Public boolExcludeEntireTableTitle As Boolean = False

    Public boolExcludeCoverPage As Boolean = False
    Public boolExcludeHeaderFooter As Boolean = False
    Public boolIncludeWaterMark As Boolean = False

    'Public xlROT As Microsoft.Office.Interop.Excel.Application

    Public gdtReportDate As Date

    Public strAnaRunPeak As String = "ANARUNRAWANALYTEPEAK"

    Public strWatsonGuWuUser As String = "StudyDoc"
    Public boolAccess As Boolean
    Public bool64 As Boolean

    Public gGoToWord As Boolean = False

    Public boolPlaceHolder As Boolean = False

    Public gboolET As Boolean = False
    Public gboolER As Boolean = False

    Public gSID As Int64 = 0 'SaveEventID

    Public boolPSL As Boolean = False

    Public gWID As Int64 'Watson Study ID when opening study in Oracle database   STUDY.STUDYID
    Public gWPID As Int64 'Watson Project ID when opening study in Oracle database   PROJECT.PROJECTID  and  STUDY.PROJECTID
    Public boolNewOracle As Boolean = False

    Public gboolRoundFiveEven As Boolean = False
    Public gboolRoundFiveAway As Boolean = True

    Public gboolCritFullPrec As Boolean = False
    Public gboolCritRounded As Boolean = True

    Public gboolMeanFullPrec As Boolean = False
    Public gboolMeanRounded As Boolean = True

    Public intTTot As Int16 = 0 'Total number of tables in a print job. Meant to be included in status labels
    Public intTCur As Int16 = 0 'Current table number of a print job. Meant to be included in status labels

    Public boolTableSectionStart As Boolean = True
    Public boolAppFigSectionStart As Boolean = True

    Public lcharCaptionTrailer As String = "."

    Public gNumCalSets As Short = 1

    Public tblAllStdsAssay As New System.Data.DataTable

    Public gSortAnalytes As String = "Matrix"
    Public gSortAnalyteString As String

    Public gNumSpecies As Short = 1 'probably don't need
    Public gNumMatrix As Short = 1 'set in DoPrepare

    Public gCHARREPORTGENERATEDSTATUS As String = ""

    Public id_tblReportHistory As Int64 = 0

    Public boolWatsonWarning As Boolean = False

    Public tblAnalyteConcLevelsForAssay As New System.Data.DataTable

    Public BOOLHOME As Boolean = False
    Public BOOLDATA As Boolean = False
    Public BOOLANALRUNSUMMARYTABLE As Boolean = False
    Public BOOLSUMMARYTABLE As Boolean = False
    Public BOOLREPORTTABLECONFIGURATION As Boolean = False
    Public BOOLREPORTTABLEHEADERCONFIG As Boolean = False
    Public BOOLANALREFSTANDARD As Boolean = False
    Public BOOLCONTRIBUTINGPERSONNEL As Boolean = False
    Public BOOLREPORTBODYSECTIONS As Boolean = False
    Public BOOLMETHODVALIDATIONDATA As Boolean = False
    Public BOOLQAEVENTTABLE As Boolean = False
    Public BOOLSAMPLERECEIPT As Boolean = False
    Public BOOLAPPENDICES As Boolean = False
    Public BOOLADMINISTRATION As Boolean = False
    Public BOOLUSERACCOUNTS As Boolean = False
    Public BOOLDROPDOWNBOXCONFIGURATION As Boolean = False
    Public BOOLCORPORATEADDRESSES As Boolean = False
    Public BOOLREPORTTEMPLATEDEFINITIONS As Boolean = False
    Public BOOLGLOBALPARAMETERS As Boolean = False
    Public BOOLALLOWREPORTGENERATION As Boolean = False
    Public BOOLHOOKS As Boolean = False
    Public BOOLALLOWPDFREPORT As Boolean = False
    Public BOOLADMINISTRATIONADMIN As Boolean = False
    Public BOOLGLOBALPARAMETERSADMIN As Boolean = False
    Public BOOLSDPROJECTS As Boolean = False
    Public BOOLSDSTUDIES As Boolean = False
    Public BOOLSDASSAYS As Boolean = False
    Public BOOLCOMPLIANCEGLOBAL As Boolean = False
    Public BOOLASSIGNSAMPLES As Boolean = False
    Public BOOLADVANCEDTABLE As Boolean = False
    Public BOOLCUSTOMFIELDCODES As Boolean = False
    Public BOOLFORCEWATERMARK As Boolean = False
    Public BOOLPERMISSIONSMANAGER As Boolean = False
    Public BOOLEDITWORDTEMPLATE As Boolean = False
    Public BOOLVIEWWORDTEMPLATE As Boolean = False
    Public BOOLSAMPLEDETAILS As Boolean = False
    Public BOOLRWAUDITTRAIL As Boolean = False
    Public BOOLCONSOLEAUDITTRAIL As Boolean = False
    Public BOOLCONSOLERW As Boolean = False
    Public BOOLALLOWRTEMPLATEPRINT As Boolean = False
    Public BOOLALLOWFINALREPORTPRINT As Boolean = False
    Public BOOLALLOWFINALREPORTWORD As Boolean = False
    Public BOOLALLOWREPORTTEMPLATEWORD As Boolean = False
    Public BOOLVIEWFINALREPORT As Boolean = False
    Public BOOLALLOWREPORTTEMPLATEPDF As Boolean = False
    Public BOOLEDITFINALREPORT As Boolean = False
    Public BOOLFORCEFINALREPORTPDF As Boolean = False

    Public boolInitLogIn As Boolean = False

    Public BOOLDOINDREC As Boolean = False
    Public BOOLRECSIGFIG As Boolean = True

    Public strWatsonVersion As String
    Public tblWatsonDBVersion As New DataTable
    Public strWatsonDBVersion As String
    Public strSDWatsonDBVErsion As String 'StudyDoc convert strWatsonDBVersion to XXYYZZ
    Public vWatsonDB(3)
    Public tblISR01_01 As New DataTable
    Public boolCanDoISR As Boolean

    Public gWatsonCutOffDt As Date

    Public BOOLREASSAYREASLETTERS As Boolean = False
    Public BOOLISCOMBINELEVELS As Boolean = True

    Public BOOLUSESTDCOLLABELS As Boolean = False

    Public strVerOld As String
    Public strVerNew As String '20190201 LEE

    Public BOOLMFTABLE As Boolean = False
    Public BOOLINCLMFCOLS As Boolean = False
    Public BOOLINCLINTSTDNMF As Boolean = False
    Public BOOLCALCINTSTDNMF As Boolean = False

    Public tblWatsonUsers As New DataTable
    Public tblWatsonStudyRoles As New DataTable

    Public gboolUseWatson As Boolean = False
    Public gboolLDAP As Boolean = False
    Public gNetAcct As String = ""
    Public gWatsonAcct As String = ""

    Public NUMPRECCRITLOTS As Decimal = 15
    Public BOOLREGRULOQ As Boolean = False

    Public gidTR As Int64 = 0

    Public BOOLSD2 As Boolean = False
    Public gSDMax As Short = 3



    Public boolSaveAsDocx As Boolean = False

    Public INTQCLEVELGROUP As Short = 0

    Public rsSampleReceiptWatson As New ADODB.Recordset

    Public tblSampleReceiptWatson As New DataTable

    Public boolTableSection As Boolean = False
    Public intTableSection As Int64 = 0

    Public boolCountLogin As Boolean = True
    Public INTWINAUTH As Short = 1

    Public intToolTipDelay As Short = 250

    Public boolReportGenAdvPrompt As Boolean = False

    Public BOOLINTSTDONLY As Boolean = False

    Public BOOLFINALREPORTLOCKED As Boolean = False

    Public gstrCAPTIONFOLLOW As String = "Tab"

    Public intUBAA As Short = 18

    Public BOOLUSERSD As Boolean = False

    Public BOOLTABLELABELSECTION As Boolean = False

    Public NUMTABLEFONTSIZE As Single = 0

    Public BOOLCONCCOMMENTS As Boolean = False

    Public BlueHyperlinkColor = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue '   Microsoft.Office.Interop.Word.WdColor.wdColorBlue

    Public dtblReturnLabelPosition As New Data.DataTable

    Public gIDDeleteStudy As Int64 = 0

    Public intDFDec As Short = 12 '20171119 LEE: Number of decimals to round Diln Factor. Need this because DF of 11-fold or 51-fold is a non-round number
    'needs to be 12 to get correct x-fold value

    Public boolChooseReportWindow As Boolean = False

    Public boolAdHocStabCompColumns As Boolean = False

    Public SpellSetting As Boolean
    Public GrammarSetting As Boolean

    Public gboolTableSection As Boolean = True
    Public arrTS(1)
    Public gboolTSDone As Boolean = False

    Public boolMeanRS As Boolean = False 'If nRS <> nPES, then boolmeanrs=true. For Ind Recovery values in Recovery tables

    Public BOOLCALIBRTABLETITLE As Boolean = False 'If study is multi-calibr level, then True/False show the calibration level in the table title

    Public strInstall() As String
    Public intInstall As Int32 = 0

End Module
