Imports ta = GooWoo.ds_2005_GuWu_01TableAdapters
Imports taOra = GooWoo.ds_GuWuOra_01TableAdapters
'Imports taAccess = GooWoo.GuWu_01DataSetTableAdapters
Imports taAccess = GooWoo.StudyDoc_01DataSet1TableAdapters

Imports taSQLServer = GooWoo.StudyDoc_01SQLDataSetTableAdapters


Module modTables


    Public daDoPr As OleDbDataAdapter = New OleDbDataAdapter()
    Public dsDoPr As DataSet = New DataSet

    Public ds2005 As ds_GuWuOra_01
    'Public ds2005Acc As GuWu_01DataSet
    Public ds2005Acc As StudyDoc_01DataSet1


    'declare oracle stuff
    Public ta_tblSampleReceipt As New taOra.TBLSAMPLERECEIPTTableAdapter
    Public tblSampleReceiptOra As New ds_GuWuOra_01.TBLSAMPLERECEIPTDataTable
    Public ta_tblData As New taOra.TBLDATATableAdapter
    Public tblDataOra As New ds_GuWuOra_01.TBLDATADataTable
    Public ta_tblTab1 As New taOra.TBLTAB1TableAdapter
    Public tblTab1Ora As New ds_GuWuOra_01.TBLTAB1DataTable
    Public ta_tblConfiguration As New taOra.TBLCONFIGURATIONTableAdapter
    Public tblConfigurationOra As New ds_GuWuOra_01.TBLCONFIGURATIONDataTable
    Public ta_tblOutstandingItems As New taOra.TBLOUTSTANDINGITEMSTableAdapter
    Public tblOutstandingItemsOra As New ds_GuWuOra_01.TBLOUTSTANDINGITEMSDataTable
    Public ta_tblPermissions As New taOra.TBLPERMISSIONSTableAdapter
    Public tblPermissionsOra As New ds_GuWuOra_01.TBLPERMISSIONSDataTable
    Public ta_tblPersonnel As New taOra.TBLPERSONNELTableAdapter
    Public tblPersonnelOra As New ds_GuWuOra_01.TBLPERSONNELDataTable
    Public ta_tblUserAccounts As New taOra.TBLUSERACCOUNTSTableAdapter
    Public tblUserAccountsOra As New ds_GuWuOra_01.TBLUSERACCOUNTSDataTable

    Public ta_tblAnalRefStandards As New taOra.TBLANALREFSTANDARDSTableAdapter
    Public tblAnalRefStandardsOra As New ds_GuWuOra_01.TBLANALREFSTANDARDSDataTable

    Public ta_tblAnalyticalRunSummary As New taOra.TBLANALYTICALRUNSUMMARYTableAdapter
    Public tblAnalyticalRunSummaryOra As New ds_GuWuOra_01.TBLANALYTICALRUNSUMMARYDataTable

    Public ta_tblConfigBodySections As New taOra.TBLCONFIGBODYSECTIONSTableAdapter
    Public tblConfigBodySectionsOra As New ds_GuWuOra_01.TBLCONFIGBODYSECTIONSDataTable

    Public ta_tblConfigHeaderLookup As New taOra.TBLCONFIGHEADERLOOKUPTableAdapter
    Public tblConfigHeaderLookupOra As New ds_GuWuOra_01.TBLCONFIGHEADERLOOKUPDataTable
    Public ta_tblConfigReportType As New taOra.TBLCONFIGREPORTTYPETableAdapter
    Public tblConfigReportTypeOra As New ds_GuWuOra_01.TBLCONFIGREPORTTYPEDataTable
    Public ta_tblContributingPersonnel As New taOra.TBLCONTRIBUTINGPERSONNELTableAdapter
    Public tblContributingPersonnelOra As New ds_GuWuOra_01.TBLCONTRIBUTINGPERSONNELDataTable
    Public ta_tblCorporateAddresses As New taOra.TBLCORPORATEADDRESSESTableAdapter
    Public tblCorporateAddressesOra As New ds_GuWuOra_01.TBLCORPORATEADDRESSESDataTable
    Public ta_tblDataTableRowTitles As New taOra.TBLDATATABLEROWTITLESTableAdapter
    Public tblDataTableRowTitlesOra As New ds_GuWuOra_01.TBLDATATABLEROWTITLESDataTable
    Public ta_tblMaxID As New taOra.TBLMAXIDTableAdapter
    Public tblMaxIDOra As New ds_GuWuOra_01.TBLMAXIDDataTable
    Public ta_tblMethodValidationData As New taOra.TBLMETHODVALIDATIONDATATableAdapter
    Public tblMethodValidationDataOra As New ds_GuWuOra_01.TBLMETHODVALIDATIONDATADataTable
    Public ta_tblQATables As New taOra.TBLQATABLESTableAdapter
    Public tblQATablesOra As New ds_GuWuOra_01.TBLQATABLESDataTable
    Public ta_tblReportHistory As New taOra.TBLREPORTHISTORYTableAdapter
    Public tblReportHistoryOra As New ds_GuWuOra_01.TBLREPORTHISTORYDataTable
    Public ta_tblReports As New taOra.TBLREPORTSTableAdapter
    Public tblReportsOra As New ds_GuWuOra_01.TBLREPORTSDataTable
    Public ta_tblReportStatements As New taOra.TBLREPORTSTATEMENTSTableAdapter
    Public tblReportStatementsOra As New ds_GuWuOra_01.TBLREPORTSTATEMENTSDataTable
    Public ta_tblReportTable As New taOra.TBLREPORTTABLETableAdapter
    Public tblReportTableOra As New ds_GuWuOra_01.TBLREPORTTABLEDataTable
    Public ta_tblReportTableAnalytes As New taOra.TBLREPORTTABLEANALYTESTableAdapter
    Public tblReportTableAnalytesOra As New ds_GuWuOra_01.TBLREPORTTABLEANALYTESDataTable
    Public ta_tblReportTableHeaderConfig As New taOra.TBLREPORTTABLEHEADERCONFIGTableAdapter
    Public tblReportTableHeaderConfigOra As New ds_GuWuOra_01.TBLREPORTTABLEHEADERCONFIGDataTable
    Public ta_tblStudies As New taOra.TBLSTUDIESTableAdapter
    Public tblStudiesOra As New ds_GuWuOra_01.TBLSTUDIESDataTable
    Public ta_tblTemplates As New taOra.TBLTEMPLATESTableAdapter
    Public tblTemplatesOra As New ds_GuWuOra_01.TBLTEMPLATESDataTable
    Public ta_tblTemplateAttributes As New taOra.TBLTEMPLATEATTRIBUTESTableAdapter
    Public tblTemplateAttributesOra As New ds_GuWuOra_01.TBLTEMPLATEATTRIBUTESDataTable
    Public ta_tblConfigReportTables As New taOra.TBLCONFIGREPORTTABLESTableAdapter
    Public tblConfigReportTablesOra As New ds_GuWuOra_01.TBLCONFIGREPORTTABLESDataTable
    Public ta_tblAddressLabels As New taOra.TBLADDRESSLABELSTableAdapter
    Public tblAddressLabelsOra As New ds_GuWuOra_01.TBLADDRESSLABELSDataTable
    Public ta_tblCorporateNickNames As New taOra.TBLCORPORATENICKNAMESTableAdapter
    Public tblCorporateNickNamesOra As New ds_GuWuOra_01.TBLCORPORATENICKNAMESDataTable
    Public ta_tblDropdownBoxContent As New taOra.TBLDROPDOWNBOXCONTENTTableAdapter
    Public tblDropdownBoxContentOra As New ds_GuWuOra_01.TBLDROPDOWNBOXCONTENTDataTable
    Public ta_tblDropdownBoxName As New taOra.TBLDROPDOWNBOXNAMETableAdapter
    Public tblDropdownBoxNameOra As New ds_GuWuOra_01.TBLDROPDOWNBOXNAMEDataTable

    Public ta_tblPasswordHistory As New taOra.TBLPASSWORDHISTORYTableAdapter
    Public tblPasswordHistoryOra As New ds_GuWuOra_01.TBLPASSWORDHISTORYDataTable

    Public ta_tblSummaryData As New taOra.TBLSUMMARYDATATableAdapter
    Public tblSummaryDataOra As New ds_GuWuOra_01.TBLSUMMARYDATADataTable

    Public ta_tblHooks As New taOra.TBLHOOKSTableAdapter
    Public tblHooksOra As New ds_GuWuOra_01.TBLHOOKSDataTable

    Public ta_tblAssignedSamples As New taOra.TBLASSIGNEDSAMPLESTableAdapter
    Public tblAssignedSamplesOra As New ds_GuWuOra_01.TBLASSIGNEDSAMPLESDataTable
    Public ta_tblDateFormats As New taOra.TBLDATEFORMATSTableAdapter
    Public tblDateFormatsOra As New ds_GuWuOra_01.TBLDATEFORMATSDataTable
    Public ta_tblAssignedSamplesHelper As New taOra.TBLASSIGNEDSAMPLESHELPERTableAdapter
    Public tblAssignedSamplesHelperOra As New ds_GuWuOra_01.TBLASSIGNEDSAMPLESHELPERDataTable
    Public ta_tblIncludedRows As New taOra.TBLINCLUDEDROWSTableAdapter
    Public tblIncludedRowsOra As New ds_GuWuOra_01.TBLINCLUDEDROWSDataTable
    Public ta_tblAppFigs As New taOra.TBLAPPFIGSTableAdapter
    Public tblAppFigsOra As New ds_GuWuOra_01.TBLAPPFIGSDataTable
    Public ta_tblConfigAppFigs As New taOra.TBLCONFIGAPPFIGSTableAdapter
    Public tblConfigAppFigsOra As New ds_GuWuOra_01.TBLCONFIGAPPFIGSDataTable

    Public ta_tblTableProperties As New taOra.TBLTABLEPROPERTIESTableAdapter
    Public tblTablePropertiesOra As New ds_GuWuOra_01.TBLTABLEPROPERTIESDataTable
    Public ta_tblTableLegends As New taOra.TBLTABLELEGENDSTableAdapter
    Public tblTableLegendsOra As New ds_GuWuOra_01.TBLTABLELEGENDSDataTable

    Public ta_tblFieldCodes As New taOra.TBLFIELDCODESTableAdapter
    Public tblFieldCodesOra As New ds_GuWuOra_01.TBLFIELDCODESDataTable
    Public ta_tblReportHeaders As New taOra.TBLREPORTHEADERSTableAdapter
    Public tblReportHeadersOra As New ds_GuWuOra_01.TBLREPORTHEADERSDataTable
    Public ta_tblWordStatements As New taOra.TBLWORDSTATEMENTSTableAdapter
    Public tblWordStatementsOra As New ds_GuWuOra_01.TBLWORDSTATEMENTSDataTable

    Public ta_tblWorddocs As New taOra.TBLWORDDOCSTableAdapter
    Public tblWorddocsOra As New ds_GuWuOra_01.TBLWORDDOCSDataTable

    Public ta_tblReasonForChange As New taOra.TBLWORDDOCSTableAdapter
    Public tblReasonForChangeOra As New ds_GuWuOra_01.TBLWORDDOCSDataTable

    Public ta_tblMeaningOfSig As New taOra.TBLWORDDOCSTableAdapter
    Public tblMeaningOfSigOra As New ds_GuWuOra_01.TBLWORDDOCSDataTable

    Public ta_tblSaveEvent As New taOra.TBLWORDDOCSTableAdapter
    Public tblSaveEventOra As New ds_GuWuOra_01.TBLWORDDOCSDataTable

    Public ta_tblDataSystem As New taOra.TBLWORDDOCSTableAdapter
    Public tblDataSystemOra As New ds_GuWuOra_01.TBLWORDDOCSDataTable

    Public ta_tblConfigCompliance As New taOra.TBLWORDDOCSTableAdapter
    Public tblConfigComplianceOra As New ds_GuWuOra_01.TBLWORDDOCSDataTable

    Public ta_tblCustomFieldCodes As New taOra.TBLWORDDOCSTableAdapter
    Public tblCustomFieldCodesOra As New ds_GuWuOra_01.TBLWORDDOCSDataTable

    '******

    Public ta_tblAuditTrail As New taOra.TBLAUDITTRAILTableAdapter
    Public tblAuditTrailOra As New ds_GuWuOra_01.TBLAUDITTRAILDataTable




    ' ''20160607 Come back to this
    ''030008
    'Public ta_TBLWORDSTATEMENTSVERSIONS As New taOra.TBLWORDSTATEMENTSVERSIONSTableAdapter
    'Public TBLWORDSTATEMENTSVERSIONSOra As New ds_GuWuOra_01DataSet1.TBLWORDSTATEMENTSVERSIONSDataTable

    ''03000901
    'Public ta_TBLSECTIONTEMPLATES As New taOra.TBLSECTIONTEMPLATESTableAdapter
    'Public TBLSECTIONTEMPLATESOra As New ds_GuWuOra_01DataSet1.TBLSECTIONTEMPLATESDataTable

    ''030030_01
    'Public ta_TBLFINALREPORT As New taOra.TBLFINALREPORTTableAdapter
    'Public TBLFINALREPORTOra As New ds_GuWuOra_01DataSet1.TBLFINALREPORTDataTable

    'Public ta_TBLFINALREPORTWORDDOCS As New taOra.TBLFINALREPORTWORDDOCSTableAdapter
    'Public TBLFINALREPORTWORDDOCSOra As New ds_GuWuOra_01DataSet1.TBLFINALREPORTWORDDOCSDataTable

    'Public ta_TBLAUTOASSIGNSAMPLES As New taOra.TBLAUTOASSIGNSAMPLESTableAdapter
    'Public TBLAUTOASSIGNSAMPLESOra As New ds_GuWuOra_01DataSet1.TBLAUTOASSIGNSAMPLESDataTable

    'Public ta_tblAppFigWordDocs As New taOra.tblAppFigWordDocsTableAdapter
    'Public tblAppFigWordDocs As New ds_GuWuOra_01.TBLAPPFIGWORDDOCSDataTable

    '******

    'end declare oracle stuff

    'declare Access stuff
    Public dd As New StudyDoc_01DataSet1
    Public ta_tblSampleReceiptAcc As New taAccess.TBLSAMPLERECEIPTTableAdapter
    Public tblSampleReceiptAcc As New StudyDoc_01DataSet1.TBLSAMPLERECEIPTDataTable
    Public ta_tblDataAcc As New taAccess.TBLDATATableAdapter
    Public tblDataAcc As New StudyDoc_01DataSet1.TBLDATADataTable
    Public ta_tblTab1Acc As New taAccess.TBLTAB1TableAdapter
    Public tblTab1Acc As New StudyDoc_01DataSet1.TBLTAB1DataTable
    Public ta_tblConfigurationAcc As New taAccess.TBLCONFIGURATIONTableAdapter
    Public tblConfigurationAcc As New StudyDoc_01DataSet1.TBLCONFIGURATIONDataTable
    Public ta_tblOutstandingItemsAcc As New taAccess.TBLOUTSTANDINGITEMSTableAdapter
    Public tblOutstandingItemsAcc As New StudyDoc_01DataSet1.TBLOUTSTANDINGITEMSDataTable
    Public ta_tblPermissionsAcc As New taAccess.TBLPERMISSIONSTableAdapter
    Public tblPermissionsAcc As New StudyDoc_01DataSet1.TBLPERMISSIONSDataTable
    Public ta_tblPersonnelAcc As New taAccess.TBLPERSONNELTableAdapter
    Public tblPersonnelAcc As New StudyDoc_01DataSet1.TBLPERSONNELDataTable
    Public ta_tblUserAccountsAcc As New taAccess.TBLUSERACCOUNTSTableAdapter
    Public tblUserAccountsAcc As New StudyDoc_01DataSet1.TBLUSERACCOUNTSDataTable
    Public ta_tblAnalRefStandardsAcc As New taAccess.TBLANALREFSTANDARDSTableAdapter
    Public tblAnalRefStandardsAcc As New StudyDoc_01DataSet1.TBLANALREFSTANDARDSDataTable
    Public ta_tblAnalyticalRunSummaryAcc As New taAccess.TBLANALYTICALRUNSUMMARYTableAdapter
    Public tblAnalyticalRunSummaryAcc As New StudyDoc_01DataSet1.TBLANALYTICALRUNSUMMARYDataTable
    Public ta_tblConfigBodySectionsAcc As New taAccess.TBLCONFIGBODYSECTIONSTableAdapter
    Public tblConfigBodySectionsAcc As New StudyDoc_01DataSet1.TBLCONFIGBODYSECTIONSDataTable
    Public ta_tblConfigHeaderLookupAcc As New taAccess.TBLCONFIGHEADERLOOKUPTableAdapter
    Public tblConfigHeaderLookupAcc As New StudyDoc_01DataSet1.TBLCONFIGHEADERLOOKUPDataTable
    Public ta_tblConfigReportTypeAcc As New taAccess.TBLCONFIGREPORTTYPETableAdapter
    Public tblConfigReportTypeAcc As New StudyDoc_01DataSet1.TBLCONFIGREPORTTYPEDataTable
    Public ta_tblContributingPersonnelAcc As New taAccess.TBLCONTRIBUTINGPERSONNELTableAdapter
    Public tblContributingPersonnelAcc As New StudyDoc_01DataSet1.TBLCONTRIBUTINGPERSONNELDataTable
    Public ta_tblCorporateAddressesAcc As New taAccess.TBLCORPORATEADDRESSESTableAdapter
    Public tblCorporateAddressesAcc As New StudyDoc_01DataSet1.TBLCORPORATEADDRESSESDataTable
    Public ta_tblDataTableRowTitlesAcc As New taAccess.TBLDATATABLEROWTITLESTableAdapter
    Public tblDataTableRowTitlesAcc As New StudyDoc_01DataSet1.TBLDATATABLEROWTITLESDataTable
    Public ta_tblMaxIDAcc As New taAccess.TBLMAXIDTableAdapter
    Public tblMaxIDAcc As New StudyDoc_01DataSet1.TBLMAXIDDataTable
    Public ta_tblMethodValidationDataAcc As New taAccess.TBLMETHODVALIDATIONDATATableAdapter
    Public tblMethodValidationDataAcc As New StudyDoc_01DataSet1.TBLMETHODVALIDATIONDATADataTable
    Public ta_tblQATablesAcc As New taAccess.TBLQATABLESTableAdapter
    Public tblQATablesAcc As New StudyDoc_01DataSet1.TBLQATABLESDataTable
    Public ta_tblReportHistoryAcc As New taAccess.TBLREPORTHISTORYTableAdapter
    Public tblReportHistoryAcc As New StudyDoc_01DataSet1.TBLREPORTHISTORYDataTable
    Public ta_tblReportsAcc As New taAccess.TBLREPORTSTableAdapter
    Public tblReportsAcc As New StudyDoc_01DataSet1.TBLREPORTSDataTable
    Public ta_tblReportStatementsAcc As New taAccess.TBLREPORTSTATEMENTSTableAdapter
    Public tblReportStatementsAcc As New StudyDoc_01DataSet1.TBLREPORTSTATEMENTSDataTable
    Public ta_tblReportTableAcc As New taAccess.TBLREPORTTABLETableAdapter
    Public tblReportTableAcc As New StudyDoc_01DataSet1.TBLREPORTTABLEDataTable
    Public ta_tblReportTableAnalytesAcc As New taAccess.TBLREPORTTABLEANALYTESTableAdapter
    Public tblReportTableAnalytesAcc As New StudyDoc_01DataSet1.TBLREPORTTABLEANALYTESDataTable
    Public ta_tblReportTableHeaderConfigAcc As New taAccess.TBLREPORTTABLEHEADERCONFIGTableAdapter
    Public tblReportTableHeaderConfigAcc As New StudyDoc_01DataSet1.TBLREPORTTABLEHEADERCONFIGDataTable
    Public ta_tblStudiesAcc As New taAccess.TBLSTUDIESTableAdapter
    Public tblStudiesAcc As New StudyDoc_01DataSet1.TBLSTUDIESDataTable
    Public ta_tblTemplatesAcc As New taAccess.TBLTEMPLATESTableAdapter
    Public tblTemplatesAcc As New StudyDoc_01DataSet1.TBLTEMPLATESDataTable
    Public ta_tblTemplateAttributesAcc As New taAccess.TBLTEMPLATEATTRIBUTESTableAdapter
    Public tblTemplateAttributesAcc As New StudyDoc_01DataSet1.TBLTEMPLATEATTRIBUTESDataTable
    Public ta_tblConfigReportTablesAcc As New taAccess.TBLCONFIGREPORTTABLESTableAdapter
    Public tblConfigReportTablesAcc As New StudyDoc_01DataSet1.TBLCONFIGREPORTTABLESDataTable
    Public ta_tblAddressLabelsAcc As New taAccess.TBLADDRESSLABELSTableAdapter
    Public tblAddressLabelsAcc As New StudyDoc_01DataSet1.TBLADDRESSLABELSDataTable
    Public ta_tblCorporateNickNamesAcc As New taAccess.TBLCORPORATENICKNAMESTableAdapter
    Public tblCorporateNickNamesAcc As New StudyDoc_01DataSet1.TBLCORPORATENICKNAMESDataTable
    Public ta_tblDropdownBoxContentAcc As New taAccess.TBLDROPDOWNBOXCONTENTTableAdapter
    Public tblDropdownBoxContentAcc As New StudyDoc_01DataSet1.TBLDROPDOWNBOXCONTENTDataTable
    Public ta_tblDropdownBoxNameAcc As New taAccess.TBLDROPDOWNBOXNAMETableAdapter
    Public tblDropdownBoxNameAcc As New StudyDoc_01DataSet1.TBLDROPDOWNBOXNAMEDataTable
    Public ta_tblPasswordHistoryAcc As New taAccess.TBLPASSWORDHISTORYTableAdapter
    Public tblPasswordHistoryAcc As New StudyDoc_01DataSet1.TBLPASSWORDHISTORYDataTable
    Public ta_tblSummaryDataAcc As New taAccess.TBLSUMMARYDATATableAdapter
    Public tblSummaryDataAcc As New StudyDoc_01DataSet1.TBLSUMMARYDATADataTable
    Public ta_tblHooksAcc As New taAccess.TBLHOOKSTableAdapter
    Public tblHooksAcc As New StudyDoc_01DataSet1.TBLHOOKSDataTable
    Public ta_tblAssignedSamplesAcc As New taAccess.TBLASSIGNEDSAMPLESTableAdapter
    Public tblAssignedSamplesAcc As New StudyDoc_01DataSet1.TBLASSIGNEDSAMPLESDataTable
    Public ta_tblDateFormatsAcc As New taAccess.TBLDATEFORMATSTableAdapter
    Public tblDateFormatsAcc As New StudyDoc_01DataSet1.TBLDATEFORMATSDataTable
    Public ta_tblAssignedSamplesHelperAcc As New taAccess.TBLASSIGNEDSAMPLESHELPERTableAdapter
    Public tblAssignedSamplesHelperAcc As New StudyDoc_01DataSet1.TBLASSIGNEDSAMPLESHELPERDataTable
    Public ta_tblIncludedRowsAcc As New taAccess.TBLINCLUDEDROWSTableAdapter
    Public tblIncludedRowsAcc As New StudyDoc_01DataSet1.TBLINCLUDEDROWSDataTable
    Public ta_tblAppFigsAcc As New taAccess.TBLAPPFIGSTableAdapter
    Public tblAppFigsAcc As New StudyDoc_01DataSet1.TBLAPPFIGSDataTable
    Public ta_tblConfigAppFigsAcc As New taAccess.TBLCONFIGAPPFIGSTableAdapter
    Public tblConfigAppFigsAcc As New StudyDoc_01DataSet1.TBLCONFIGAPPFIGSDataTable

    Public ta_tblTablePropertiesAcc As New taAccess.TBLTABLEPROPERTIESTableAdapter
    Public tblTablePropertiesAcc As New StudyDoc_01DataSet1.TBLTABLEPROPERTIESDataTable
    Public ta_tblTableLegendsAcc As New taAccess.TBLTABLELEGENDSTableAdapter
    Public tblTableLegendsAcc As New StudyDoc_01DataSet1.TBLTABLELEGENDSDataTable

    Public ta_tblFieldCodesAcc As New taAccess.TBLFIELDCODESTableAdapter
    Public tblFieldCodesAcc As New StudyDoc_01DataSet1.TBLFIELDCODESDataTable
    Public ta_tblReportHeadersAcc As New taAccess.TBLREPORTHEADERSTableAdapter
    Public tblReportHeadersAcc As New StudyDoc_01DataSet1.TBLREPORTHEADERSDataTable
    Public ta_tblWordStatementsAcc As New taAccess.TBLWORDSTATEMENTSTableAdapter
    Public tblWordStatementsAcc As New StudyDoc_01DataSet1.TBLWORDSTATEMENTSDataTable

    'Public ta_tblWorddocsAcc As New taAccess.TBLWORDDOCSTableAdapter
    'Public tblWorddocsAcc As New StudyDoc_01DataSet1.TBLWORDDOCSDataTable

    Public ta_tblAuditTrailAcc As New taAccess.TBLAUDITTRAILTableAdapter
    Public tblAuditTrailAcc As New StudyDoc_01DataSet1.TBLAUDITTRAILDataTable

    Public ta_tblReasonForChangeAcc As New taAccess.TBLREASONFORCHANGETableAdapter
    Public tblReasonForChangeAcc As New StudyDoc_01DataSet1.TBLREASONFORCHANGEDataTable

    Public ta_tblMeaningOfSigAcc As New taAccess.TBLMEANINGOFSIGTableAdapter
    Public tblMeaningOfSigAcc As New StudyDoc_01DataSet1.TBLMEANINGOFSIGDataTable

    Public ta_tblSaveEventAcc As New taAccess.TBLSAVEEVENTTableAdapter
    Public tblSaveEventAcc As New StudyDoc_01DataSet1.TBLSAVEEVENTDataTable

    Public ta_tblDataSystemAcc As New taAccess.TBLDATASYSTEMTableAdapter
    Public tblDataSystemAcc As New StudyDoc_01DataSet1.TBLDATASYSTEMDataTable

    Public ta_tblConfigComplianceAcc As New taAccess.TBLCONFIGCOMPLIANCETableAdapter
    Public tblConfigComplianceAcc As New StudyDoc_01DataSet1.TBLCONFIGCOMPLIANCEDataTable

    '02218
    Public ta_tblCustomFieldCodesAcc As New taAccess.TBLCUSTOMFIELDCODESTableAdapter
    Public tblCustomFieldCodesAcc As New StudyDoc_01DataSet1.TBLCUSTOMFIELDCODESDataTable

    '030008
    Public ta_TBLWORDSTATEMENTSVERSIONSAcc As New taAccess.TBLWORDSTATEMENTSVERSIONSTableAdapter
    Public TBLWORDSTATEMENTSVERSIONSAcc As New StudyDoc_01DataSet1.TBLWORDSTATEMENTSVERSIONSDataTable

    '03000901
    Public ta_TBLSECTIONTEMPLATESAcc As New taAccess.TBLSECTIONTEMPLATESTableAdapter
    Public TBLSECTIONTEMPLATESAcc As New StudyDoc_01DataSet1.TBLSECTIONTEMPLATESDataTable

    '030030_01
    Public ta_TBLFINALREPORTAcc As New taAccess.TBLFINALREPORTTableAdapter
    Public TBLFINALREPORTAcc As New StudyDoc_01DataSet1.TBLFINALREPORTDataTable

    Public ta_TBLFINALREPORTWORDDOCSAcc As New taAccess.TBLFINALREPORTWORDDOCSTableAdapter
    Public TBLFINALREPORTWORDDOCSAcc As New StudyDoc_01DataSet1.TBLFINALREPORTWORDDOCSDataTable

    '030040_04
    Public ta_TBLAUTOASSIGNSAMPLESAcc As New taAccess.TBLAUTOASSIGNSAMPLESTableAdapter
    Public TBLAUTOASSIGNSAMPLESAcc As New StudyDoc_01DataSet1.TBLAUTOASSIGNSAMPLESDataTable
    Public ta_tblAppFigWordDocsAcc As New taAccess.TBLAPPFIGWORDDOCSTableAdapter
    Public tblAppFigWordDocsAcc As New StudyDoc_01DataSet1.TBLAPPFIGWORDDOCSDataTable

    '030066_02
    Public ta_TBLSTUDYDOCANALYTESAcc As New taAccess.TBLSTUDYDOCANALYTESTableAdapter
    Public TBLSTUDYDOCANALYTESAcc As New StudyDoc_01DataSet1.TBLSTUDYDOCANALYTESDataTable

    'TBLAUTOASSIGNSAMPLES

    'start StudyDesigner
    Public ta_tblModulesAcc As New taAccess.TBLMODULESTableAdapter
    Public TBLMODULESAcc As New StudyDoc_01DataSet1.TBLMODULESDataTable

    Public ta_TBLVERSIONAcc As New taAccess.TBLVERSIONTableAdapter
    Public TBLVERSIONAcc As New StudyDoc_01DataSet1.TBLVERSIONDataTable

    Public ta_TBLGUWUANIMALRECEIPTAcc As New taAccess.TBLGUWUANIMALRECEIPTTableAdapter
    Public TBLGUWUANIMALRECEIPTAcc As New StudyDoc_01DataSet1.TBLGUWUANIMALRECEIPTDataTable

    Public ta_TBLGUWUCOMPOUNDSAcc As New taAccess.TBLGUWUCOMPOUNDSTableAdapter
    Public TBLGUWUCOMPOUNDSAcc As New StudyDoc_01DataSet1.TBLGUWUCOMPOUNDSDataTable

    Public ta_TBLGUWUCOMPOUNDSINDAcc As New taAccess.TBLGUWUCOMPOUNDSINDTableAdapter
    Public TBLGUWUCOMPOUNDSINDAcc As New StudyDoc_01DataSet1.TBLGUWUCOMPOUNDSINDDataTable

    Public ta_TBLGUWUCOMPOUNDTYPEAcc As New taAccess.TBLGUWUCOMPOUNDTYPETableAdapter
    Public TBLGUWUCOMPOUNDTYPEAcc As New StudyDoc_01DataSet1.TBLGUWUCOMPOUNDTYPEDataTable

    Public ta_TBLGUWUPROJECTSAcc As New taAccess.TBLGUWUPROJECTSTableAdapter
    Public TBLGUWUPROJECTSAcc As New StudyDoc_01DataSet1.TBLGUWUPROJECTSDataTable

    Public ta_TBLGUWUSPECIESAcc As New taAccess.TBLGUWUSPECIESTableAdapter
    Public TBLGUWUSPECIESAcc As New StudyDoc_01DataSet1.TBLGUWUSPECIESDataTable

    Public ta_TBLGUWUSTUDIESAcc As New taAccess.TBLGUWUSTUDIESTableAdapter
    Public TBLGUWUSTUDIESAcc As New StudyDoc_01DataSet1.TBLGUWUSTUDIESDataTable

    Public ta_TBLGUWUSTUDYDESIGNTYPEAcc As New taAccess.TBLGUWUSTUDYDESIGNTYPETableAdapter
    Public TBLGUWUSTUDYDESIGNTYPEAcc As New StudyDoc_01DataSet1.TBLGUWUSTUDYDESIGNTYPEDataTable

    Public ta_TBLGUWUSTUDYSPECIESAcc As New taAccess.TBLGUWUSTUDYSPECIESTableAdapter
    Public TBLGUWUSTUDYSPECIESAcc As New StudyDoc_01DataSet1.TBLGUWUSTUDYSPECIESDataTable

    Public ta_TBLGUWUSTUDYSTATAcc As New taAccess.TBLGUWUSTUDYSTATTableAdapter
    Public TBLGUWUSTUDYSTATAcc As New StudyDoc_01DataSet1.TBLGUWUSTUDYSTATDataTable

    Public ta_TBLGUWUASSAYPERSAcc As New taAccess.TBLGUWUASSAYPERSTableAdapter
    Public TBLGUWUASSAYPERSAcc As New StudyDoc_01DataSet1.TBLGUWUASSAYPERSDataTable

    Public ta_TBLGUWUASSAYAcc As New taAccess.TBLGUWUASSAYTableAdapter
    Public TBLGUWUASSAYAcc As New StudyDoc_01DataSet1.TBLGUWUASSAYDataTable

    Public ta_TBLGUWUSPECIESSTRAINAcc As New taAccess.TBLGUWUSPECIESSTRAINTableAdapter
    Public TBLGUWUSPECIESSTRAINAcc As New StudyDoc_01DataSet1.TBLGUWUSPECIESSTRAINDataTable

    Public ta_TBLGUWUDOSEUNITSAcc As New taAccess.TBLGUWUDOSEUNITSTableAdapter
    Public TBLGUWUDOSEUNITSAcc As New StudyDoc_01DataSet1.TBLGUWUDOSEUNITSDataTable

    Public ta_TBLGUWUPKGROUPSAcc As New taAccess.TBLGUWUPKGROUPSTableAdapter
    Public TBLGUWUPKGROUPSAcc As New StudyDoc_01DataSet1.TBLGUWUPKGROUPSDataTable

    Public ta_TBLGUWUPKROUTESAcc As New taAccess.TBLGUWUPKROUTESTableAdapter
    Public TBLGUWUPKROUTESAcc As New StudyDoc_01DataSet1.TBLGUWUPKROUTESDataTable

    Public ta_TBLGUWUPKSUBJECTSAcc As New taAccess.TBLGUWUPKSUBJECTSTableAdapter
    Public TBLGUWUPKSUBJECTSAcc As New StudyDoc_01DataSet1.TBLGUWUPKSUBJECTSDataTable

    Public ta_TBLGUWURTTIMEPOINTSAcc As New taAccess.TBLGUWURTTIMEPOINTSTableAdapter
    Public TBLGUWURTTIMEPOINTSAcc As New StudyDoc_01DataSet1.TBLGUWURTTIMEPOINTSDataTable

    Public ta_TBLGUWUASSIGNEDCMPDAcc As New taAccess.TBLGUWUASSIGNEDCMPDTableAdapter
    Public TBLGUWUASSIGNEDCMPDAcc As New StudyDoc_01DataSet1.TBLGUWUASSIGNEDCMPDDataTable

    Public ta_TBLGUWUASSIGNEDCMPDLOTAcc As New taAccess.TBLGUWUASSIGNEDCMPDLOTTableAdapter
    Public TBLGUWUASSIGNEDCMPDLOTAcc As New StudyDoc_01DataSet1.TBLGUWUASSIGNEDCMPDLOTDataTable

    Public ta_TBLGUWUSTUDYSCHEDULINGAcc As New taAccess.TBLGUWUSTUDYSCHEDULINGTableAdapter
    Public TBLGUWUSTUDYSCHEDULINGAcc As New StudyDoc_01DataSet1.TBLGUWUSTUDYSCHEDULINGDataTable

    Public ta_TBLGUWUTPCONFIGAcc As New taAccess.TBLGUWUTPCONFIGTableAdapter
    Public TBLGUWUTPCONFIGAcc As New StudyDoc_01DataSet1.TBLGUWUTPCONFIGDataTable

    Public ta_TBLGUWUTPNAMESCONFIGAcc As New taAccess.TBLGUWUTPNAMESCONFIGTableAdapter
    Public TBLGUWUTPNAMESCONFIGAcc As New StudyDoc_01DataSet1.TBLGUWUTPNAMESCONFIGDataTable

    'Public ta_QRYGUWUCALENDARAcc As New taAccess.QRYGUWUCALENDARTableAdapter
    'Public QRYGUWUCALENDARAcc As New StudyDoc_01DataSet1.QRYGUWUCALENDARDataTable

    'end taAccess


    'start taSQLServer

    Public ta_tblSampleReceiptSQLServer As New taSQLServer.TBLSAMPLERECEIPTTableAdapter
    'Public tblSampleReceiptSQLServer As New StudyDoc_01DataSet1.TBLSAMPLERECEIPTDataTable
    Public tblSampleReceiptSQLServer As New StudyDoc_01SQLDataSet.TBLSAMPLERECEIPTDataTable

    Public ta_tblDataSQLServer As New taSQLServer.TBLDATATableAdapter
    Public tblDataSQLServer As New StudyDoc_01SQLDataSet.TBLDATADataTable
    Public ta_tblTab1SQLServer As New taSQLServer.TBLTAB1TableAdapter
    Public tblTab1SQLServer As New StudyDoc_01SQLDataSet.TBLTAB1DataTable
    Public ta_tblConfigurationSQLServer As New taSQLServer.TBLCONFIGURATIONTableAdapter
    Public tblConfigurationSQLServer As New StudyDoc_01SQLDataSet.TBLCONFIGURATIONDataTable
    Public ta_tblOutstandingItemsSQLServer As New taSQLServer.TBLOUTSTANDINGITEMSTableAdapter
    Public tblOutstandingItemsSQLServer As New StudyDoc_01SQLDataSet.TBLOUTSTANDINGITEMSDataTable
    Public ta_tblPermissionsSQLServer As New taSQLServer.TBLPERMISSIONSTableAdapter
    Public tblPermissionsSQLServer As New StudyDoc_01SQLDataSet.TBLPERMISSIONSDataTable
    Public ta_tblPersonnelSQLServer As New taSQLServer.TBLPERSONNELTableAdapter
    Public tblPersonnelSQLServer As New StudyDoc_01SQLDataSet.TBLPERSONNELDataTable
    Public ta_tblUserAccountsSQLServer As New taSQLServer.TBLUSERACCOUNTSTableAdapter
    Public tblUserAccountsSQLServer As New StudyDoc_01SQLDataSet.TBLUSERACCOUNTSDataTable
    Public ta_tblAnalRefStandardsSQLServer As New taSQLServer.TBLANALREFSTANDARDSTableAdapter
    Public tblAnalRefStandardsSQLServer As New StudyDoc_01SQLDataSet.TBLANALREFSTANDARDSDataTable
    Public ta_tblAnalyticalRunSummarySQLServer As New taSQLServer.TBLANALYTICALRUNSUMMARYTableAdapter
    Public tblAnalyticalRunSummarySQLServer As New StudyDoc_01SQLDataSet.TBLANALYTICALRUNSUMMARYDataTable
    Public ta_tblConfigBodySectionsSQLServer As New taSQLServer.TBLCONFIGBODYSECTIONSTableAdapter
    Public tblConfigBodySectionsSQLServer As New StudyDoc_01SQLDataSet.TBLCONFIGBODYSECTIONSDataTable
    Public ta_tblConfigHeaderLookupSQLServer As New taSQLServer.TBLCONFIGHEADERLOOKUPTableAdapter
    Public tblConfigHeaderLookupSQLServer As New StudyDoc_01SQLDataSet.TBLCONFIGHEADERLOOKUPDataTable
    Public ta_tblConfigReportTypeSQLServer As New taSQLServer.TBLCONFIGREPORTTYPETableAdapter
    Public tblConfigReportTypeSQLServer As New StudyDoc_01SQLDataSet.TBLCONFIGREPORTTYPEDataTable
    Public ta_tblContributingPersonnelSQLServer As New taSQLServer.TBLCONTRIBUTINGPERSONNELTableAdapter
    Public tblContributingPersonnelSQLServer As New StudyDoc_01SQLDataSet.TBLCONTRIBUTINGPERSONNELDataTable
    Public ta_tblCorporateAddressesSQLServer As New taSQLServer.TBLCORPORATEADDRESSESTableAdapter
    Public tblCorporateAddressesSQLServer As New StudyDoc_01SQLDataSet.TBLCORPORATEADDRESSESDataTable
    Public ta_tblDataTableRowTitlesSQLServer As New taSQLServer.TBLDATATABLEROWTITLESTableAdapter
    Public tblDataTableRowTitlesSQLServer As New StudyDoc_01SQLDataSet.TBLDATATABLEROWTITLESDataTable
    Public ta_tblMaxIDSQLServer As New taSQLServer.TBLMAXIDTableAdapter
    Public tblMaxIDSQLServer As New StudyDoc_01SQLDataSet.TBLMAXIDDataTable
    Public ta_tblMethodValidationDataSQLServer As New taSQLServer.TBLMETHODVALIDATIONDATATableAdapter
    Public tblMethodValidationDataSQLServer As New StudyDoc_01SQLDataSet.TBLMETHODVALIDATIONDATADataTable
    Public ta_tblQATablesSQLServer As New taSQLServer.TBLQATABLESTableAdapter
    Public tblQATablesSQLServer As New StudyDoc_01SQLDataSet.TBLQATABLESDataTable
    Public ta_tblReportHistorySQLServer As New taSQLServer.TBLREPORTHISTORYTableAdapter
    Public tblReportHistorySQLServer As New StudyDoc_01SQLDataSet.TBLREPORTHISTORYDataTable
    Public ta_tblReportsSQLServer As New taSQLServer.TBLREPORTSTableAdapter
    Public tblReportsSQLServer As New StudyDoc_01SQLDataSet.TBLREPORTSDataTable
    Public ta_tblReportStatementsSQLServer As New taSQLServer.TBLREPORTSTATEMENTSTableAdapter
    Public tblReportStatementsSQLServer As New StudyDoc_01SQLDataSet.TBLREPORTSTATEMENTSDataTable
    Public ta_tblReportTableSQLServer As New taSQLServer.TBLREPORTTABLETableAdapter
    Public tblReportTableSQLServer As New StudyDoc_01SQLDataSet.TBLREPORTTABLEDataTable
    Public ta_tblReportTableAnalytesSQLServer As New taSQLServer.TBLREPORTTABLEANALYTESTableAdapter
    Public tblReportTableAnalytesSQLServer As New StudyDoc_01SQLDataSet.TBLREPORTTABLEANALYTESDataTable
    Public ta_tblReportTableHeaderConfigSQLServer As New taSQLServer.TBLREPORTTABLEHEADERCONFIGTableAdapter
    Public tblReportTableHeaderConfigSQLServer As New StudyDoc_01SQLDataSet.TBLREPORTTABLEHEADERCONFIGDataTable
    Public ta_tblStudiesSQLServer As New taSQLServer.TBLSTUDIESTableAdapter
    Public tblStudiesSQLServer As New StudyDoc_01SQLDataSet.TBLSTUDIESDataTable
    Public ta_tblTemplatesSQLServer As New taSQLServer.TBLTEMPLATESTableAdapter
    Public tblTemplatesSQLServer As New StudyDoc_01SQLDataSet.TBLTEMPLATESDataTable
    Public ta_tblTemplateAttributesSQLServer As New taSQLServer.TBLTEMPLATEATTRIBUTESTableAdapter
    Public tblTemplateAttributesSQLServer As New StudyDoc_01SQLDataSet.TBLTEMPLATEATTRIBUTESDataTable
    Public ta_tblConfigReportTablesSQLServer As New taSQLServer.TBLCONFIGREPORTTABLESTableAdapter
    Public tblConfigReportTablesSQLServer As New StudyDoc_01SQLDataSet.TBLCONFIGREPORTTABLESDataTable
    Public ta_tblAddressLabelsSQLServer As New taSQLServer.TBLADDRESSLABELSTableAdapter
    Public tblAddressLabelsSQLServer As New StudyDoc_01SQLDataSet.TBLADDRESSLABELSDataTable
    Public ta_tblCorporateNickNamesSQLServer As New taSQLServer.TBLCORPORATENICKNAMESTableAdapter
    Public tblCorporateNickNamesSQLServer As New StudyDoc_01SQLDataSet.TBLCORPORATENICKNAMESDataTable
    Public ta_tblDropdownBoxContentSQLServer As New taSQLServer.TBLDROPDOWNBOXCONTENTTableAdapter
    Public tblDropdownBoxContentSQLServer As New StudyDoc_01SQLDataSet.TBLDROPDOWNBOXCONTENTDataTable
    Public ta_tblDropdownBoxNameSQLServer As New taSQLServer.TBLDROPDOWNBOXNAMETableAdapter
    Public tblDropdownBoxNameSQLServer As New StudyDoc_01SQLDataSet.TBLDROPDOWNBOXNAMEDataTable
    Public ta_tblPasswordHistorySQLServer As New taSQLServer.TBLPASSWORDHISTORYTableAdapter
    Public tblPasswordHistorySQLServer As New StudyDoc_01SQLDataSet.TBLPASSWORDHISTORYDataTable
    Public ta_tblSummaryDataSQLServer As New taSQLServer.TBLSUMMARYDATATableAdapter
    Public tblSummaryDataSQLServer As New StudyDoc_01SQLDataSet.TBLSUMMARYDATADataTable
    Public ta_tblHooksSQLServer As New taSQLServer.TBLHOOKSTableAdapter
    Public tblHooksSQLServer As New StudyDoc_01SQLDataSet.TBLHOOKSDataTable
    Public ta_tblAssignedSamplesSQLServer As New taSQLServer.TBLASSIGNEDSAMPLESTableAdapter
    Public tblAssignedSamplesSQLServer As New StudyDoc_01SQLDataSet.TBLASSIGNEDSAMPLESDataTable
    Public ta_tblDateFormatsSQLServer As New taSQLServer.TBLDATEFORMATSTableAdapter
    Public tblDateFormatsSQLServer As New StudyDoc_01SQLDataSet.TBLDATEFORMATSDataTable
    Public ta_tblAssignedSamplesHelperSQLServer As New taSQLServer.TBLASSIGNEDSAMPLESHELPERTableAdapter
    Public tblAssignedSamplesHelperSQLServer As New StudyDoc_01SQLDataSet.TBLASSIGNEDSAMPLESHELPERDataTable
    Public ta_tblIncludedRowsSQLServer As New taSQLServer.TBLINCLUDEDROWSTableAdapter
    Public tblIncludedRowsSQLServer As New StudyDoc_01SQLDataSet.TBLINCLUDEDROWSDataTable
    Public ta_tblAppFigsSQLServer As New taSQLServer.TBLAPPFIGSTableAdapter
    Public tblAppFigsSQLServer As New StudyDoc_01SQLDataSet.TBLAPPFIGSDataTable
    Public ta_tblConfigAppFigsSQLServer As New taSQLServer.TBLCONFIGAPPFIGSTableAdapter
    Public tblConfigAppFigsSQLServer As New StudyDoc_01SQLDataSet.TBLCONFIGAPPFIGSDataTable

    Public ta_tblTablePropertiesSQLServer As New taSQLServer.TBLTABLEPROPERTIESTableAdapter
    Public tblTablePropertiesSQLServer As New StudyDoc_01SQLDataSet.TBLTABLEPROPERTIESDataTable
    Public ta_tblTableLegendsSQLServer As New taSQLServer.TBLTABLELEGENDSTableAdapter
    Public tblTableLegendsSQLServer As New StudyDoc_01SQLDataSet.TBLTABLELEGENDSDataTable

    Public ta_tblFieldCodesSQLServer As New taSQLServer.TBLFIELDCODESTableAdapter
    Public tblFieldCodesSQLServer As New StudyDoc_01SQLDataSet.TBLFIELDCODESDataTable
    Public ta_tblReportHeadersSQLServer As New taSQLServer.TBLREPORTHEADERSTableAdapter
    Public tblReportHeadersSQLServer As New StudyDoc_01SQLDataSet.TBLREPORTHEADERSDataTable
    Public ta_tblWordStatementsSQLServer As New taSQLServer.TBLWORDSTATEMENTSTableAdapter
    Public tblWordStatementsSQLServer As New StudyDoc_01SQLDataSet.TBLWORDSTATEMENTSDataTable

    'Public ta_tblWorddocsSQLServer As New taSQLServer.TBLWORDDOCSTableAdapter
    'Public tblWorddocsSQLServer As New StudyDoc_01SQLDataSet.TBLWORDDOCSDataTable

    Public ta_tblAuditTrailSQLServer As New taSQLServer.TBLAUDITTRAILTableAdapter
    Public tblAuditTrailSQLServer As New StudyDoc_01SQLDataSet.TBLAUDITTRAILDataTable

    Public ta_tblReasonForChangeSQLServer As New taSQLServer.TBLREASONFORCHANGETableAdapter
    Public tblReasonForChangeSQLServer As New StudyDoc_01SQLDataSet.TBLREASONFORCHANGEDataTable

    Public ta_tblMeaningOfSigSQLServer As New taSQLServer.TBLMEANINGOFSIGTableAdapter
    Public tblMeaningOfSigSQLServer As New StudyDoc_01SQLDataSet.TBLMEANINGOFSIGDataTable

    Public ta_tblSaveEventSQLServer As New taSQLServer.TBLSAVEEVENTTableAdapter
    Public tblSaveEventSQLServer As New StudyDoc_01SQLDataSet.TBLSAVEEVENTDataTable

    Public ta_tblDataSystemSQLServer As New taSQLServer.TBLDATASYSTEMTableAdapter
    Public tblDataSystemSQLServer As New StudyDoc_01SQLDataSet.TBLDATASYSTEMDataTable

    Public ta_tblConfigComplianceSQLServer As New taSQLServer.TBLCONFIGCOMPLIANCETableAdapter
    Public tblConfigComplianceSQLServer As New StudyDoc_01SQLDataSet.TBLCONFIGCOMPLIANCEDataTable

    '02218
    Public ta_tblCustomFieldCodesSQLServer As New taSQLServer.TBLCUSTOMFIELDCODESTableAdapter
    Public tblCustomFieldCodesSQLServer As New StudyDoc_01SQLDataSet.TBLCUSTOMFIELDCODESDataTable

    '030008
    Public ta_TBLWORDSTATEMENTSVERSIONSSQLServer As New taSQLServer.TBLWORDSTATEMENTSVERSIONSTableAdapter
    Public TBLWORDSTATEMENTSVERSIONSSQLServer As New StudyDoc_01SQLDataSet.TBLWORDSTATEMENTSVERSIONSDataTable

    '03000901
    Public ta_TBLSECTIONTEMPLATESSQLServer As New taSQLServer.TBLSECTIONTEMPLATESTableAdapter
    Public TBLSECTIONTEMPLATESSQLServer As New StudyDoc_01SQLDataSet.TBLSECTIONTEMPLATESDataTable

    '030030_01
    Public ta_TBLFINALREPORTSQLServer As New taSQLServer.TBLFINALREPORTTableAdapter
    Public TBLFINALREPORTSQLServer As New StudyDoc_01SQLDataSet.TBLFINALREPORTDataTable

    Public ta_TBLFINALREPORTWORDDOCSSQLServer As New taSQLServer.TBLFINALREPORTWORDDOCSTableAdapter
    Public TBLFINALREPORTWORDDOCSSQLServer As New StudyDoc_01SQLDataSet.TBLFINALREPORTWORDDOCSDataTable

    ''030040_04
    Public ta_TBLAUTOASSIGNSAMPLESSQLServer As New taSQLServer.TBLAUTOASSIGNSAMPLESTableAdapter
    Public TBLAUTOASSIGNSAMPLESSQLServer As New StudyDoc_01SQLDataSet.TBLAUTOASSIGNSAMPLESDataTable
    Public ta_tblAppFigWordDocsSQLSERVER As New taSQLServer.TBLAPPFIGWORDDOCSTableAdapter
    Public tblAppFigWordDocsSQLSERVER As New StudyDoc_01SQLDataSet.TBLAPPFIGWORDDOCSDataTable

    '030066_02
    Public ta_TBLSTUDYDOCANALYTESSQLSERVER As New taSQLServer.TBLSTUDYDOCANALYTESTableAdapter
    Public TBLSTUDYDOCANALYTESSQLSERVER As New StudyDoc_01SQLDataSet.TBLSTUDYDOCANALYTESDataTable

    'start StudyDesigner
    Public ta_tblModulesSQLServer As New taSQLServer.TBLMODULESTableAdapter
    Public TBLMODULESSQLServer As New StudyDoc_01SQLDataSet.TBLMODULESDataTable

    Public ta_TBLVERSIONSQLServer As New taSQLServer.TBLVERSIONTableAdapter
    Public TBLVERSIONSQLServer As New StudyDoc_01SQLDataSet.TBLVERSIONDataTable

    Public ta_TBLGUWUANIMALRECEIPTSQLServer As New taSQLServer.TBLGUWUANIMALRECEIPTTableAdapter
    Public TBLGUWUANIMALRECEIPTSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUANIMALRECEIPTDataTable

    Public ta_TBLGUWUCOMPOUNDSSQLServer As New taSQLServer.TBLGUWUCOMPOUNDSTableAdapter
    Public TBLGUWUCOMPOUNDSSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUCOMPOUNDSDataTable

    Public ta_TBLGUWUCOMPOUNDSINDSQLServer As New taSQLServer.TBLGUWUCOMPOUNDSINDTableAdapter
    Public TBLGUWUCOMPOUNDSINDSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUCOMPOUNDSINDDataTable

    Public ta_TBLGUWUCOMPOUNDTYPESQLServer As New taSQLServer.TBLGUWUCOMPOUNDTYPETableAdapter
    Public TBLGUWUCOMPOUNDTYPESQLServer As New StudyDoc_01SQLDataSet.TBLGUWUCOMPOUNDTYPEDataTable

    Public ta_TBLGUWUPROJECTSSQLServer As New taSQLServer.TBLGUWUPROJECTSTableAdapter
    Public TBLGUWUPROJECTSSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUPROJECTSDataTable

    Public ta_TBLGUWUSPECIESSQLServer As New taSQLServer.TBLGUWUSPECIESTableAdapter
    Public TBLGUWUSPECIESSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUSPECIESDataTable

    Public ta_TBLGUWUSTUDIESSQLServer As New taSQLServer.TBLGUWUSTUDIESTableAdapter
    Public TBLGUWUSTUDIESSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUSTUDIESDataTable

    Public ta_TBLGUWUSTUDYDESIGNTYPESQLServer As New taSQLServer.TBLGUWUSTUDYDESIGNTYPETableAdapter
    Public TBLGUWUSTUDYDESIGNTYPESQLServer As New StudyDoc_01SQLDataSet.TBLGUWUSTUDYDESIGNTYPEDataTable

    Public ta_TBLGUWUSTUDYSPECIESSQLServer As New taSQLServer.TBLGUWUSTUDYSPECIESTableAdapter
    Public TBLGUWUSTUDYSPECIESSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUSTUDYSPECIESDataTable

    Public ta_TBLGUWUSTUDYSTATSQLServer As New taSQLServer.TBLGUWUSTUDYSTATTableAdapter
    Public TBLGUWUSTUDYSTATSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUSTUDYSTATDataTable

    Public ta_TBLGUWUASSAYPERSSQLServer As New taSQLServer.TBLGUWUASSAYPERSTableAdapter
    Public TBLGUWUASSAYPERSSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUASSAYPERSDataTable

    Public ta_TBLGUWUASSAYSQLServer As New taSQLServer.TBLGUWUASSAYTableAdapter
    Public TBLGUWUASSAYSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUASSAYDataTable

    Public ta_TBLGUWUSPECIESSTRAINSQLServer As New taSQLServer.TBLGUWUSPECIESSTRAINTableAdapter
    Public TBLGUWUSPECIESSTRAINSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUSPECIESSTRAINDataTable

    Public ta_TBLGUWUDOSEUNITSSQLServer As New taSQLServer.TBLGUWUDOSEUNITSTableAdapter
    Public TBLGUWUDOSEUNITSSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUDOSEUNITSDataTable

    Public ta_TBLGUWUPKGROUPSSQLServer As New taSQLServer.TBLGUWUPKGROUPSTableAdapter
    Public TBLGUWUPKGROUPSSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUPKGROUPSDataTable

    Public ta_TBLGUWUPKROUTESSQLServer As New taSQLServer.TBLGUWUPKROUTESTableAdapter
    Public TBLGUWUPKROUTESSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUPKROUTESDataTable

    Public ta_TBLGUWUPKSUBJECTSSQLServer As New taSQLServer.TBLGUWUPKSUBJECTSTableAdapter
    Public TBLGUWUPKSUBJECTSSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUPKSUBJECTSDataTable

    Public ta_TBLGUWURTTIMEPOINTSSQLServer As New taSQLServer.TBLGUWURTTIMEPOINTSTableAdapter
    Public TBLGUWURTTIMEPOINTSSQLServer As New StudyDoc_01SQLDataSet.TBLGUWURTTIMEPOINTSDataTable

    Public ta_TBLGUWUASSIGNEDCMPDSQLServer As New taSQLServer.TBLGUWUASSIGNEDCMPDTableAdapter
    Public TBLGUWUASSIGNEDCMPDSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUASSIGNEDCMPDDataTable

    Public ta_TBLGUWUASSIGNEDCMPDLOTSQLServer As New taSQLServer.TBLGUWUASSIGNEDCMPDLOTTableAdapter
    Public TBLGUWUASSIGNEDCMPDLOTSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUASSIGNEDCMPDLOTDataTable

    Public ta_TBLGUWUSTUDYSCHEDULINGSQLServer As New taSQLServer.TBLGUWUSTUDYSCHEDULINGTableAdapter
    Public TBLGUWUSTUDYSCHEDULINGSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUSTUDYSCHEDULINGDataTable

    Public ta_TBLGUWUTPCONFIGSQLServer As New taSQLServer.TBLGUWUTPCONFIGTableAdapter
    Public TBLGUWUTPCONFIGSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUTPCONFIGDataTable

    Public ta_TBLGUWUTPNAMESCONFIGSQLServer As New taSQLServer.TBLGUWUTPNAMESCONFIGTableAdapter
    Public TBLGUWUTPNAMESCONFIGSQLServer As New StudyDoc_01SQLDataSet.TBLGUWUTPNAMESCONFIGDataTable

    'Public ta_QRYGUWUCALENDARSQLServer As New taSQLServer.QRYGUWUCALENDARTableAdapter
    'Public QRYGUWUCALENDARSQLServer As New StudyDoc_01SQLDataSet.QRYGUWUCALENDARDataTable


    'end taSQLServer

    'end StudyDesigner

    'end declare access stuff

    Public tblSampleReceipt As New System.Data.DataTable
    Public tblData As New System.Data.DataTable
    Public tblTab1 As New System.Data.DataTable
    Public tblConfiguration As New System.Data.DataTable
    Public tblOutstandingItems As New System.Data.DataTable
    Public tblPermissions As New System.Data.DataTable
    Public tblPersonnel As New System.Data.DataTable
    Public tblUserAccounts As New System.Data.DataTable
    Public tblAnalRefStandards As New System.Data.DataTable
    Public tblAnalyticalRunSummary As New System.Data.DataTable
    Public tblConfigBodySections As New System.Data.DataTable
    Public tblConfigHeaderLookup As New System.Data.DataTable
    Public tblConfigReportType As New System.Data.DataTable
    Public tblContributingPersonnel As New System.Data.DataTable
    Public tblCorporateAddresses As New System.Data.DataTable
    Public tblDataTableRowTitles As New System.Data.DataTable
    Public tblMaxID As New System.Data.DataTable
    Public tblMethodValidationData As New System.Data.DataTable
    Public tblQATables As New System.Data.DataTable
    Public tblReportHistory As New System.Data.DataTable
    Public tblReports As New System.Data.DataTable
    Public tblReportStatements As New System.Data.DataTable
    Public tblReportTable As New System.Data.DataTable
    Public tblReportTableAnalytes As New System.Data.DataTable
    Public tblReportTableHeaderConfig As New System.Data.DataTable
    Public tblStudies As New System.Data.DataTable
    Public tblTemplates As New System.Data.DataTable
    Public tblTemplateAttributes As New System.Data.DataTable
    Public tblConfigReportTables As New System.Data.DataTable
    Public tblAddressLabels As New System.Data.DataTable
    Public tblCorporateNickNames As New System.Data.DataTable
    Public tblDropdownBoxContent As New System.Data.DataTable
    Public tblDropdownBoxName As New System.Data.DataTable
    Public tblPasswordHistory As New System.Data.DataTable
    Public tblSummaryData As New System.Data.DataTable
    Public tblHooks As New System.Data.DataTable
    Public tblAssignedSamples As New System.Data.DataTable
    Public tblDateFormats As New System.Data.DataTable
    Public tblAssignedSamplesHelper As New System.Data.DataTable
    Public tblIncludedRows As New System.Data.DataTable
    Public tblAppFigs As New System.Data.DataTable
    Public tblConfigAppFigs As New System.Data.DataTable
    Public tblTableProperties As New System.Data.DataTable
    Public tblTableLegends As New System.Data.DataTable
    Public tblFieldCodes As New System.Data.DataTable
    Public tblReportHeaders As New System.Data.DataTable
    Public tblWordStatements As New System.Data.DataTable
    Public tblWorddocs As New System.Data.DataTable
    '
    Public tblReasonForChange As New System.Data.DataTable
    Public tblMeaningOfSig As New System.Data.DataTable
    Public tblSaveEvent As New System.Data.DataTable
    Public tblDataSystem As New System.Data.DataTable
    Public tblConfigCompliance As New System.Data.DataTable
    Public tblCustomFieldCodes As New System.Data.DataTable

    '20160607 Declare problem
    Public tblAuditTrail As New System.Data.DataTable
    Public TBLWORDSTATEMENTSVERSIONS As New System.Data.DataTable
    Public TBLSECTIONTEMPLATES As New System.Data.DataTable
    Public tblFinalReport As New System.Data.DataTable
    Public tblFinalReportWordDocs As New System.Data.DataTable

    Public tblAutoAssignSamples As New System.Data.DataTable
    Public tblAppFigWordDocs As New System.Data.DataTable
    Public TBLSTUDYDOCANALYTES As New System.Data.DataTable


    'START STUDY DESIGN
    Public tblVersion As New System.Data.DataTable
    Public tblModules As New System.Data.DataTable
    Public tblGuWuAnimalReceipt As New System.Data.DataTable
    Public tblGuWuCompounds As New System.Data.DataTable
    Public tblGuWuCompoundsInd As New System.Data.DataTable
    Public tblGuWuCompoundType As New System.Data.DataTable
    Public tblGuWuProjects As New System.Data.DataTable
    Public tblGuWuSpecies As New System.Data.DataTable
    Public tblGuWuStudies As New System.Data.DataTable
    Public tblGuWuStudyDesignType As New System.Data.DataTable
    Public tblGuWuStudySpecies As New System.Data.DataTable
    Public tblGuWuStudyStat As New System.Data.DataTable
    Public TBLGUWUASSAYPERS As New System.Data.DataTable
    Public tblGuWuAssay As New System.Data.DataTable
    Public tblGuWuSpeciesStrain As New System.Data.DataTable
    Public tblGuWuDoseUnits As New System.Data.DataTable
    Public tblGuWuPKGroups As New System.Data.DataTable
    Public tblGuWuPKRoutes As New System.Data.DataTable
    Public tblGuWuPKSubjects As New System.Data.DataTable
    Public tblGuWuRTTimePoints As New System.Data.DataTable
    Public tblGuWuAssignedCmpd As New System.Data.DataTable
    Public tblGuWuAssignedCmpdLot As New System.Data.DataTable
    Public TBLGUWUSTUDYSCHEDULING As New System.Data.DataTable
    Public TBLGUWUTPCONFIG As New System.Data.DataTable
    Public TBLGUWUTPNAMESCONFIG As New System.Data.DataTable
    'Public QRYGUWUCALENDAR As new System.Data.DataTable


    Sub ConfigOra()

        tblSampleReceipt = tblSampleReceiptOra
        tblData = tblDataOra
        tblTab1 = tblTab1Ora
        tblConfiguration = tblConfigurationOra
        tblOutstandingItems = tblOutstandingItemsOra
        tblPermissions = tblPermissionsOra
        tblPersonnel = tblPersonnelOra
        tblUserAccounts = tblUserAccountsOra
        tblAnalRefStandards = tblAnalRefStandardsOra
        tblAnalyticalRunSummary = tblAnalyticalRunSummaryOra
        tblConfigBodySections = tblConfigBodySectionsOra
        tblConfigHeaderLookup = tblConfigHeaderLookupOra
        tblConfigReportType = tblConfigReportTypeOra
        tblContributingPersonnel = tblContributingPersonnelOra
        tblCorporateAddresses = tblCorporateAddressesOra
        tblDataTableRowTitles = tblDataTableRowTitlesOra
        tblMaxID = tblMaxIDOra
        tblMethodValidationData = tblMethodValidationDataOra
        tblQATables = tblQATablesOra
        tblReportHistory = tblReportHistoryOra
        tblReports = tblReportsOra
        tblReportStatements = tblReportStatementsOra
        tblReportTable = tblReportTableOra
        tblReportTableAnalytes = tblReportTableAnalytesOra
        tblReportTableHeaderConfig = tblReportTableHeaderConfigOra
        tblStudies = tblStudiesOra
        tblTemplates = tblTemplatesOra
        tblTemplateAttributes = tblTemplateAttributesOra
        tblConfigReportTables = tblConfigReportTablesOra
        tblAddressLabels = tblAddressLabelsOra
        tblCorporateNickNames = tblCorporateNickNamesOra
        tblDropdownBoxContent = tblDropdownBoxContentOra
        tblDropdownBoxName = tblDropdownBoxNameOra
        tblPasswordHistory = tblPasswordHistoryOra
        tblSummaryData = tblSummaryDataOra
        tblHooks = tblHooksOra
        tblAssignedSamples = tblAssignedSamplesOra
        tblDateFormats = tblDateFormatsOra
        tblAssignedSamplesHelper = tblAssignedSamplesHelperOra
        tblIncludedRows = tblIncludedRowsOra
        tblAppFigs = tblAppFigsOra
        tblConfigAppFigs = tblConfigAppFigsOra
        tblTableProperties = tblTablePropertiesOra
        tblTableLegends = tblTableLegendsOra
        tblFieldCodes = tblFieldCodesOra
        tblReportHeaders = tblReportHeadersOra
        tblWordStatements = tblWordStatementsOra
        tblWorddocs = tblWorddocsOra

        'come back to this later
        'TBLSECTIONTEMPLATES = TBLSECTIONTEMPLATESOra


    End Sub

    Sub ConfigAccess()

        tblSampleReceipt = tblSampleReceiptAcc
        tblData = tblDataAcc
        tblTab1 = tblTab1Acc
        tblConfiguration = tblConfigurationAcc
        tblOutstandingItems = tblOutstandingItemsAcc
        tblPermissions = tblPermissionsAcc
        tblPersonnel = tblPersonnelAcc
        tblUserAccounts = tblUserAccountsAcc
        tblAnalRefStandards = tblAnalRefStandardsAcc
        tblAnalyticalRunSummary = tblAnalyticalRunSummaryAcc
        tblConfigBodySections = tblConfigBodySectionsAcc
        tblConfigHeaderLookup = tblConfigHeaderLookupAcc
        tblConfigReportType = tblConfigReportTypeAcc
        tblContributingPersonnel = tblContributingPersonnelAcc
        tblCorporateAddresses = tblCorporateAddressesAcc
        tblDataTableRowTitles = tblDataTableRowTitlesAcc
        tblMaxID = tblMaxIDAcc
        tblMethodValidationData = tblMethodValidationDataAcc
        tblQATables = tblQATablesAcc
        tblReportHistory = tblReportHistoryAcc
        tblReports = tblReportsAcc
        tblReportStatements = tblReportStatementsAcc
        tblReportTable = tblReportTableAcc
        tblReportTableAnalytes = tblReportTableAnalytesAcc
        tblReportTableHeaderConfig = tblReportTableHeaderConfigAcc
        tblStudies = tblStudiesAcc
        tblTemplates = tblTemplatesAcc
        tblTemplateAttributes = tblTemplateAttributesAcc
        tblConfigReportTables = tblConfigReportTablesAcc
        tblAddressLabels = tblAddressLabelsAcc
        tblCorporateNickNames = tblCorporateNickNamesAcc
        tblDropdownBoxContent = tblDropdownBoxContentAcc
        tblDropdownBoxName = tblDropdownBoxNameAcc
        tblPasswordHistory = tblPasswordHistoryAcc
        tblSummaryData = tblSummaryDataAcc
        tblHooks = tblHooksAcc
        tblAssignedSamples = tblAssignedSamplesAcc
        tblDateFormats = tblDateFormatsAcc
        tblAssignedSamplesHelper = tblAssignedSamplesHelperAcc
        tblIncludedRows = tblIncludedRowsAcc
        tblAppFigs = tblAppFigsAcc
        tblConfigAppFigs = tblConfigAppFigsAcc
        tblTableProperties = tblTablePropertiesAcc
        tblTableLegends = tblTableLegendsAcc
        tblFieldCodes = tblFieldCodesAcc
        tblReportHeaders = tblReportHeadersAcc
        tblWordStatements = tblWordStatementsAcc
        'tblWorddocs = tblWorddocsAcc
        tblAuditTrail = tblAuditTrailAcc
        tblReasonForChange = tblReasonForChangeAcc
        tblMeaningOfSig = tblMeaningOfSigAcc
        tblSaveEvent = tblSaveEventAcc
        tblDataSystem = tblDataSystemAcc
        tblConfigCompliance = tblConfigComplianceAcc
        '02218
        tblCustomFieldCodes = tblCustomFieldCodesAcc
        '030008
        TBLWORDSTATEMENTSVERSIONS = TBLWORDSTATEMENTSVERSIONSAcc
        '03000901
        TBLSECTIONTEMPLATES = TBLSECTIONTEMPLATESAcc

        '030030_01
        tblFinalReport = TBLFINALREPORTAcc
        tblFinalReportWordDocs = TBLFINALREPORTWORDDOCSAcc
        '030040_04
        tblAutoAssignSamples = TBLAUTOASSIGNSAMPLESAcc
        tblAppFigWordDocs = tblAppFigWordDocsAcc
        '030466_02
        TBLSTUDYDOCANALYTES = TBLSTUDYDOCANALYTESAcc

        tblModules = TBLMODULESAcc
        tblVersion = TBLVERSIONAcc
        tblGuWuAnimalReceipt = TBLGUWUANIMALRECEIPTAcc
        tblGuWuCompounds = TBLGUWUCOMPOUNDSAcc
        tblGuWuCompoundsInd = TBLGUWUCOMPOUNDSINDAcc
        tblGuWuCompoundType = TBLGUWUCOMPOUNDTYPEAcc
        tblGuWuProjects = TBLGUWUPROJECTSAcc
        tblGuWuSpecies = TBLGUWUSPECIESAcc
        tblGuWuStudies = TBLGUWUSTUDIESAcc
        tblGuWuStudyDesignType = TBLGUWUSTUDYDESIGNTYPEAcc
        tblGuWuStudySpecies = TBLGUWUSTUDYSPECIESAcc
        tblGuWuStudyStat = TBLGUWUSTUDYSTATAcc
        TBLGUWUASSAYPERS = TBLGUWUASSAYPERSAcc
        tblGuWuAssay = TBLGUWUASSAYAcc
        tblGuWuSpeciesStrain = TBLGUWUSPECIESSTRAINAcc
        tblGuWuDoseUnits = TBLGUWUDOSEUNITSAcc
        tblGuWuPKGroups = TBLGUWUPKGROUPSAcc
        tblGuWuPKRoutes = TBLGUWUPKROUTESAcc
        tblGuWuPKSubjects = TBLGUWUPKSUBJECTSAcc
        tblGuWuRTTimePoints = TBLGUWURTTIMEPOINTSAcc
        tblGuWuAssignedCmpd = TBLGUWUASSIGNEDCMPDAcc
        tblGuWuAssignedCmpdLot = TBLGUWUASSIGNEDCMPDLOTAcc
        TBLGUWUSTUDYSCHEDULING = TBLGUWUSTUDYSCHEDULINGAcc
        TBLGUWUTPCONFIG = TBLGUWUTPCONFIGAcc
        TBLGUWUTPNAMESCONFIG = TBLGUWUTPNAMESCONFIGAcc
        'QRYGUWUCALENDAR = QRYGUWUCALENDARAcc

    End Sub


    Sub ConfigSQLServer()

        tblSampleReceipt = tblSampleReceiptSQLServer
        tblData = tblDataSQLServer
        tblTab1 = tblTab1SQLServer
        tblConfiguration = tblConfigurationSQLServer
        tblOutstandingItems = tblOutstandingItemsSQLServer
        tblPermissions = tblPermissionsSQLServer
        tblPersonnel = tblPersonnelSQLServer
        tblUserAccounts = tblUserAccountsSQLServer
        tblAnalRefStandards = tblAnalRefStandardsSQLServer
        tblAnalyticalRunSummary = tblAnalyticalRunSummarySQLServer
        tblConfigBodySections = tblConfigBodySectionsSQLServer
        tblConfigHeaderLookup = tblConfigHeaderLookupSQLServer
        tblConfigReportType = tblConfigReportTypeSQLServer
        tblContributingPersonnel = tblContributingPersonnelSQLServer
        tblCorporateAddresses = tblCorporateAddressesSQLServer
        tblDataTableRowTitles = tblDataTableRowTitlesSQLServer
        tblMaxID = tblMaxIDSQLServer
        tblMethodValidationData = tblMethodValidationDataSQLServer
        tblQATables = tblQATablesSQLServer
        tblReportHistory = tblReportHistorySQLServer
        tblReports = tblReportsSQLServer
        tblReportStatements = tblReportStatementsSQLServer
        tblReportTable = tblReportTableSQLServer
        tblReportTableAnalytes = tblReportTableAnalytesSQLServer
        tblReportTableHeaderConfig = tblReportTableHeaderConfigSQLServer
        tblStudies = tblStudiesSQLServer
        tblTemplates = tblTemplatesSQLServer
        tblTemplateAttributes = tblTemplateAttributesSQLServer
        tblConfigReportTables = tblConfigReportTablesSQLServer
        tblAddressLabels = tblAddressLabelsSQLServer
        tblCorporateNickNames = tblCorporateNickNamesSQLServer
        tblDropdownBoxContent = tblDropdownBoxContentSQLServer
        tblDropdownBoxName = tblDropdownBoxNameSQLServer
        tblPasswordHistory = tblPasswordHistorySQLServer
        tblSummaryData = tblSummaryDataSQLServer
        tblHooks = tblHooksSQLServer
        tblAssignedSamples = tblAssignedSamplesSQLServer
        tblDateFormats = tblDateFormatsSQLServer
        tblAssignedSamplesHelper = tblAssignedSamplesHelperSQLServer
        tblIncludedRows = tblIncludedRowsSQLServer
        tblAppFigs = tblAppFigsSQLServer
        tblConfigAppFigs = tblConfigAppFigsSQLServer
        tblTableProperties = tblTablePropertiesSQLServer
        tblTableLegends = tblTableLegendsSQLServer
        tblFieldCodes = tblFieldCodesSQLServer
        tblReportHeaders = tblReportHeadersSQLServer
        tblWordStatements = tblWordStatementsSQLServer
        'tblWorddocs = tblWorddocsSQLServer
        tblAuditTrail = tblAuditTrailSQLServer
        tblReasonForChange = tblReasonForChangeSQLServer
        tblMeaningOfSig = tblMeaningOfSigSQLServer
        tblSaveEvent = tblSaveEventSQLServer
        tblDataSystem = tblDataSystemSQLServer
        tblConfigCompliance = tblConfigComplianceSQLServer
        '02218
        tblCustomFieldCodes = tblCustomFieldCodesSQLServer
        '030008
        TBLWORDSTATEMENTSVERSIONS = TBLWORDSTATEMENTSVERSIONSSQLServer
        '03000901
        TBLSECTIONTEMPLATES = TBLSECTIONTEMPLATESSQLServer

        '030030_01
        tblFinalReport = TBLFINALREPORTSQLServer
        tblFinalReportWordDocs = TBLFINALREPORTWORDDOCSSQLServer
        '030040_04
        tblAutoAssignSamples = TBLAUTOASSIGNSAMPLESSQLServer
        tblAppFigWordDocs = tblAppFigWordDocsSQLSERVER
        '030066_02
        TBLSTUDYDOCANALYTES = TBLSTUDYDOCANALYTESSQLSERVER

        tblModules = TBLMODULESSQLServer
        tblVersion = TBLVERSIONSQLServer
        tblGuWuAnimalReceipt = TBLGUWUANIMALRECEIPTSQLServer
        tblGuWuCompounds = TBLGUWUCOMPOUNDSSQLServer
        tblGuWuCompoundsInd = TBLGUWUCOMPOUNDSINDSQLServer
        tblGuWuCompoundType = TBLGUWUCOMPOUNDTYPESQLServer
        tblGuWuProjects = TBLGUWUPROJECTSSQLServer
        tblGuWuSpecies = TBLGUWUSPECIESSQLServer
        tblGuWuStudies = TBLGUWUSTUDIESSQLServer
        tblGuWuStudyDesignType = TBLGUWUSTUDYDESIGNTYPESQLServer
        tblGuWuStudySpecies = TBLGUWUSTUDYSPECIESSQLServer
        tblGuWuStudyStat = TBLGUWUSTUDYSTATSQLServer
        TBLGUWUASSAYPERS = TBLGUWUASSAYPERSSQLServer
        tblGuWuAssay = TBLGUWUASSAYSQLServer
        tblGuWuSpeciesStrain = TBLGUWUSPECIESSTRAINSQLServer
        tblGuWuDoseUnits = TBLGUWUDOSEUNITSSQLServer
        tblGuWuPKGroups = TBLGUWUPKGROUPSSQLServer
        tblGuWuPKRoutes = TBLGUWUPKROUTESSQLServer
        tblGuWuPKSubjects = TBLGUWUPKSUBJECTSSQLServer
        tblGuWuRTTimePoints = TBLGUWURTTIMEPOINTSSQLServer
        tblGuWuAssignedCmpd = TBLGUWUASSIGNEDCMPDSQLServer
        tblGuWuAssignedCmpdLot = TBLGUWUASSIGNEDCMPDLOTSQLServer
        TBLGUWUSTUDYSCHEDULING = TBLGUWUSTUDYSCHEDULINGSQLServer
        TBLGUWUTPCONFIG = TBLGUWUTPCONFIGSQLServer
        TBLGUWUTPNAMESCONFIG = TBLGUWUTPNAMESCONFIGSQLServer
        'QRYGUWUCALENDAR = QRYGUWUCALENDARSQLServer

    End Sub


    Sub DAConnect(ByVal frm As Form)

        On Error GoTo end1

        ta_tblSampleReceipt.Connection.ConnectionString = constrIni
        ta_tblData.Connection.ConnectionString = constrIni
        ta_tblTab1.Connection.ConnectionString = constrIni
        ta_tblConfiguration.Connection.ConnectionString = constrIni
        ta_tblOutstandingItems.Connection.ConnectionString = constrIni
        ta_tblPermissions.Connection.ConnectionString = constrIni
        ta_tblPersonnel.Connection.ConnectionString = constrIni
        ta_tblUserAccounts.Connection.ConnectionString = constrIni
        ta_tblAnalRefStandards.Connection.ConnectionString = constrIni
        ta_tblAnalyticalRunSummary.Connection.ConnectionString = constrIni
        ta_tblConfigBodySections.Connection.ConnectionString = constrIni
        ta_tblConfigHeaderLookup.Connection.ConnectionString = constrIni
        ta_tblConfigReportType.Connection.ConnectionString = constrIni
        ta_tblContributingPersonnel.Connection.ConnectionString = constrIni
        ta_tblCorporateAddresses.Connection.ConnectionString = constrIni
        ta_tblDataTableRowTitles.Connection.ConnectionString = constrIni
        ta_tblMaxID.Connection.ConnectionString = constrIni
        ta_tblMethodValidationData.Connection.ConnectionString = constrIni
        ta_tblQATables.Connection.ConnectionString = constrIni
        ta_tblReportHistory.Connection.ConnectionString = constrIni
        ta_tblReports.Connection.ConnectionString = constrIni
        ta_tblReportStatements.Connection.ConnectionString = constrIni
        ta_tblReportTable.Connection.ConnectionString = constrIni
        ta_tblReportTableAnalytes.Connection.ConnectionString = constrIni
        ta_tblReportTableHeaderConfig.Connection.ConnectionString = constrIni
        ta_tblStudies.Connection.ConnectionString = constrIni
        ta_tblTemplates.Connection.ConnectionString = constrIni
        ta_tblTemplateAttributes.Connection.ConnectionString = constrIni
        ta_tblConfigReportTables.Connection.ConnectionString = constrIni
        ta_tblAddressLabels.Connection.ConnectionString = constrIni
        ta_tblCorporateNickNames.Connection.ConnectionString = constrIni
        ta_tblDropdownBoxContent.Connection.ConnectionString = constrIni
        ta_tblDropdownBoxName.Connection.ConnectionString = constrIni
        ta_tblPasswordHistory.Connection.ConnectionString = constrIni
        ta_tblSummaryData.Connection.ConnectionString = constrIni
        ta_tblHooks.Connection.ConnectionString = constrIni
        ta_tblAssignedSamples.Connection.ConnectionString = constrIni
        ta_tblDateFormats.Connection.ConnectionString = constrIni
        ta_tblAssignedSamplesHelper.Connection.ConnectionString = constrIni
        ta_tblIncludedRows.Connection.ConnectionString = constrIni
        ta_tblConfigAppFigs.Connection.ConnectionString = constrIni
        ta_tblAppFigs.Connection.ConnectionString = constrIni

        ta_tblTableProperties.Connection.ConnectionString = constrIni
        ta_tblTableLegends.Connection.ConnectionString = constrIni

        ta_tblFieldCodes.Connection.ConnectionString = constrIni
        ta_tblReportHeaders.Connection.ConnectionString = constrIni
        ta_tblWordStatements.Connection.ConnectionString = constrIni

        'ta_tblWorddocs.Connection.ConnectionString = constrIni

        ta_tblReasonForChange.Connection.ConnectionString = constrIni
        ta_tblMeaningOfSig.Connection.ConnectionString = constrIni
        ta_tblSaveEvent.Connection.ConnectionString = constrIni
        ta_tblDataSystem.Connection.ConnectionString = constrIni
        ta_tblConfigCompliance.Connection.ConnectionString = constrIni
        '02218:
        ta_tblCustomFieldCodes.Connection.ConnectionString = constrIni
        '030008
        'ta_TBLWORDSTATEMENTSVERSIONS.Connection.ConnectionString = constrIni
        '03000901
        'come back to this later
        'ta_TBLSECTIONTEMPLATES.Connection.ConnectionString = constrIni

        On Error GoTo 0

        Exit Sub

end1:
        Dim str1 As String
        Dim str2 As String
        If Err.Number <> 0 Then
            str1 = "Hmmm." & Chr(10) & "There seems to be a problem connecting to the StudyDoc datatabase."
            str1 = str1 & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
            str2 = "Critical communication error..."
            If boolFormLoad Then

                frm.Controls("lblErr").Text = str1
                frm.Controls("lblErr").Refresh()
            Else

                Call PositionProgress()
                frmH.lblProgress.Text = str1
                frmH.lblProgress.Visible = True
                frmH.lblProgress.Refresh()

                frmH.panProgress.Visible = True
                frmH.panProgress.Refresh()

            End If

            MsgBox(str1, MsgBoxStyle.Critical, str2)
            Dim dt As Date
            Dim dt1 As Date
            dt = Now
            dt1 = DateAdd(DateInterval.Second, 1, dt)
            Do Until dt > dt1
                dt = Now
            Loop

            End

        End If
        On Error GoTo 0

    End Sub

    Sub DAConnectAcc(ByVal frm As Form)

        Try
            ''console.writeline(constrIni)
            ta_tblSampleReceiptAcc.Connection.ConnectionString = constrIni
            ta_tblDataAcc.Connection.ConnectionString = constrIni
            ta_tblTab1Acc.Connection.ConnectionString = constrIni
            ta_tblConfigurationAcc.Connection.ConnectionString = constrIni
            ta_tblOutstandingItemsAcc.Connection.ConnectionString = constrIni
            ta_tblPermissionsAcc.Connection.ConnectionString = constrIni
            ta_tblPersonnelAcc.Connection.ConnectionString = constrIni
            ta_tblUserAccountsAcc.Connection.ConnectionString = constrIni
            ta_tblAnalRefStandardsAcc.Connection.ConnectionString = constrIni
            ta_tblAnalyticalRunSummaryAcc.Connection.ConnectionString = constrIni
            ta_tblConfigBodySectionsAcc.Connection.ConnectionString = constrIni
            ta_tblConfigHeaderLookupAcc.Connection.ConnectionString = constrIni
            ta_tblConfigReportTypeAcc.Connection.ConnectionString = constrIni
            ta_tblContributingPersonnelAcc.Connection.ConnectionString = constrIni
            ta_tblCorporateAddressesAcc.Connection.ConnectionString = constrIni
            ta_tblDataTableRowTitlesAcc.Connection.ConnectionString = constrIni
            ta_tblMaxIDAcc.Connection.ConnectionString = constrIni
            ta_tblMethodValidationDataAcc.Connection.ConnectionString = constrIni
            ta_tblQATablesAcc.Connection.ConnectionString = constrIni
            ta_tblReportHistoryAcc.Connection.ConnectionString = constrIni
            ta_tblReportsAcc.Connection.ConnectionString = constrIni
            ta_tblReportStatementsAcc.Connection.ConnectionString = constrIni
            ta_tblReportTableAcc.Connection.ConnectionString = constrIni
            ta_tblReportTableAnalytesAcc.Connection.ConnectionString = constrIni
            ta_tblReportTableHeaderConfigAcc.Connection.ConnectionString = constrIni
            ta_tblStudiesAcc.Connection.ConnectionString = constrIni
            ta_tblTemplatesAcc.Connection.ConnectionString = constrIni
            ta_tblTemplateAttributesAcc.Connection.ConnectionString = constrIni
            ta_tblConfigReportTablesAcc.Connection.ConnectionString = constrIni
            ta_tblAddressLabelsAcc.Connection.ConnectionString = constrIni
            ta_tblCorporateNickNamesAcc.Connection.ConnectionString = constrIni
            ta_tblDropdownBoxContentAcc.Connection.ConnectionString = constrIni
            ta_tblDropdownBoxNameAcc.Connection.ConnectionString = constrIni
            ta_tblPasswordHistoryAcc.Connection.ConnectionString = constrIni
            ta_tblSummaryDataAcc.Connection.ConnectionString = constrIni
            ta_tblHooksAcc.Connection.ConnectionString = constrIni
            ta_tblAssignedSamplesAcc.Connection.ConnectionString = constrIni
            ta_tblDateFormatsAcc.Connection.ConnectionString = constrIni
            ta_tblAssignedSamplesHelperAcc.Connection.ConnectionString = constrIni
            ta_tblIncludedRowsAcc.Connection.ConnectionString = constrIni
            ta_tblConfigAppFigsAcc.Connection.ConnectionString = constrIni
            ta_tblAppFigsAcc.Connection.ConnectionString = constrIni

            ta_tblTablePropertiesAcc.Connection.ConnectionString = constrIni
            ta_tblTableLegendsAcc.Connection.ConnectionString = constrIni

            ta_tblFieldCodesAcc.Connection.ConnectionString = constrIni
            ta_tblReportHeadersAcc.Connection.ConnectionString = constrIni
            ta_tblWordStatementsAcc.Connection.ConnectionString = constrIni

            'ta_tblWorddocsAcc.Connection.ConnectionString = constrIni
            ta_tblAuditTrailAcc.Connection.ConnectionString = constrIni

            ta_tblReasonForChangeAcc.Connection.ConnectionString = constrIni
            ta_tblMeaningOfSigAcc.Connection.ConnectionString = constrIni
            ta_tblSaveEventAcc.Connection.ConnectionString = constrIni
            ta_tblDataSystemAcc.Connection.ConnectionString = constrIni
            ta_tblConfigComplianceAcc.Connection.ConnectionString = constrIni
            '02218:
            ta_tblCustomFieldCodesAcc.Connection.ConnectionString = constrIni
            '030008
            ta_TBLWORDSTATEMENTSVERSIONSAcc.Connection.ConnectionString = constrIni
            '03000901
            ta_TBLSECTIONTEMPLATESAcc.Connection.ConnectionString = constrIni
            '030030_01
            ta_TBLFINALREPORTAcc.Connection.ConnectionString = constrIni
            ta_TBLFINALREPORTWORDDOCSAcc.Connection.ConnectionString = constrIni
            '030040_04
            ta_TBLAUTOASSIGNSAMPLESAcc.Connection.ConnectionString = constrIni
            ta_tblAppFigWordDocsAcc.Connection.ConnectionString = constrIni
            '030466_02
            ta_TBLSTUDYDOCANALYTESAcc.Connection.ConnectionString = constrIni

            'start Study Design
            ta_tblModulesAcc.Connection.ConnectionString = constrIni
            ta_TBLVERSIONAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUANIMALRECEIPTAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUCOMPOUNDSAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUCOMPOUNDSINDAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUCOMPOUNDTYPEAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUPROJECTSAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUSPECIESAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUSPECIESAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUSTUDYDESIGNTYPEAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUSTUDYSPECIESAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUSTUDYSTATAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUASSAYPERSAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUASSAYAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUSPECIESSTRAINAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUDOSEUNITSAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUPKGROUPSAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUPKROUTESAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUPKSUBJECTSAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWURTTIMEPOINTSAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUASSIGNEDCMPDAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUASSIGNEDCMPDLOTAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUSTUDYSCHEDULINGAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUTPCONFIGAcc.Connection.ConnectionString = constrIni
            ta_TBLGUWUTPNAMESCONFIGAcc.Connection.ConnectionString = constrIni
            'ta_QRYGUWUCALENDARAcc.Connection.ConnectionString = constrIni

            ta_TBLGUWUSTUDIESAcc.Connection.ConnectionString = constrIni

        Catch ex As Exception

            Dim str1 As String
            Dim str2 As String
            str1 = "Hmmm." & ChrW(10) & "There seems to be a problem connecting to the StudyDoc datatabase."
            str1 = str1 & ChrW(10) & ChrW(10) & "Please contact your StudyDoc system administrator."
            str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
            str2 = "Critical communication error..."
            If boolFormLoad Then

                frm.Controls("lblErr").Text = str1
                frm.Controls("lblErr").Refresh()
            Else
                Call PositionProgress()
                frmH.lblProgress.Text = str1
                frmH.lblProgress.Visible = True
                frmH.lblProgress.Refresh()

                frmH.panProgress.Visible = True
                frmH.panProgress.Refresh()

            End If

            MsgBox(str1, MsgBoxStyle.Critical, str2)
            Dim dt As Date
            Dim dt1 As Date
            dt = Now
            dt1 = DateAdd(DateInterval.Second, 1, dt)
            Do Until dt > dt1
                dt = Now
            Loop

            End

        End Try



    End Sub

    Sub DAConnectSQLServer(ByVal frm As Form)

        'Note: If StudyDoc is SQLServer, this is an example of the OLEDB connection string:
        '     Provider=SQLOLEDB;DATA SOURCE=GUBBSLAP07;INITIAL CATALOG=STUDYDOC_00028
        'For some reason, the vb.net connection string fails if 'Provider' is provided
        'Here, must strip if off


        Dim constr As String = GetNETProvider()
        Dim str1 As String
        Dim str2 As String

        '20190218 LEE:
        'testing
        Dim dt1 As DateTime = Now
        Dim dt2 As DateTime

        Try
            ''console.writeline(constr)
            ta_tblSampleReceiptSQLServer.Connection.ConnectionString = constr
            ta_tblDataSQLServer.Connection.ConnectionString = constr
            ta_tblTab1SQLServer.Connection.ConnectionString = constr
            ta_tblConfigurationSQLServer.Connection.ConnectionString = constr
            ta_tblOutstandingItemsSQLServer.Connection.ConnectionString = constr
            ta_tblPermissionsSQLServer.Connection.ConnectionString = constr
            ta_tblPersonnelSQLServer.Connection.ConnectionString = constr
            ta_tblUserAccountsSQLServer.Connection.ConnectionString = constr
            ta_tblAnalRefStandardsSQLServer.Connection.ConnectionString = constr
            ta_tblAnalyticalRunSummarySQLServer.Connection.ConnectionString = constr
            ta_tblConfigBodySectionsSQLServer.Connection.ConnectionString = constr
            ta_tblConfigHeaderLookupSQLServer.Connection.ConnectionString = constr
            ta_tblConfigReportTypeSQLServer.Connection.ConnectionString = constr
            ta_tblContributingPersonnelSQLServer.Connection.ConnectionString = constr
            ta_tblCorporateAddressesSQLServer.Connection.ConnectionString = constr
            ta_tblDataTableRowTitlesSQLServer.Connection.ConnectionString = constr
            ta_tblMaxIDSQLServer.Connection.ConnectionString = constr
            ta_tblMethodValidationDataSQLServer.Connection.ConnectionString = constr
            ta_tblQATablesSQLServer.Connection.ConnectionString = constr
            ta_tblReportHistorySQLServer.Connection.ConnectionString = constr
            ta_tblReportsSQLServer.Connection.ConnectionString = constr
            ta_tblReportStatementsSQLServer.Connection.ConnectionString = constr
            ta_tblReportTableSQLServer.Connection.ConnectionString = constr
            ta_tblReportTableAnalytesSQLServer.Connection.ConnectionString = constr
            ta_tblReportTableHeaderConfigSQLServer.Connection.ConnectionString = constr
            ta_tblStudiesSQLServer.Connection.ConnectionString = constr
            ta_tblTemplatesSQLServer.Connection.ConnectionString = constr
            ta_tblTemplateAttributesSQLServer.Connection.ConnectionString = constr
            ta_tblConfigReportTablesSQLServer.Connection.ConnectionString = constr
            ta_tblAddressLabelsSQLServer.Connection.ConnectionString = constr
            ta_tblCorporateNickNamesSQLServer.Connection.ConnectionString = constr
            ta_tblDropdownBoxContentSQLServer.Connection.ConnectionString = constr
            ta_tblDropdownBoxNameSQLServer.Connection.ConnectionString = constr
            ta_tblPasswordHistorySQLServer.Connection.ConnectionString = constr
            ta_tblSummaryDataSQLServer.Connection.ConnectionString = constr
            ta_tblHooksSQLServer.Connection.ConnectionString = constr
            ta_tblAssignedSamplesSQLServer.Connection.ConnectionString = constr
            ta_tblDateFormatsSQLServer.Connection.ConnectionString = constr
            ta_tblAssignedSamplesHelperSQLServer.Connection.ConnectionString = constr
            ta_tblIncludedRowsSQLServer.Connection.ConnectionString = constr
            ta_tblConfigAppFigsSQLServer.Connection.ConnectionString = constr
            ta_tblAppFigsSQLServer.Connection.ConnectionString = constr

            ta_tblTablePropertiesSQLServer.Connection.ConnectionString = constr
            ta_tblTableLegendsSQLServer.Connection.ConnectionString = constr

            ta_tblFieldCodesSQLServer.Connection.ConnectionString = constr
            ta_tblReportHeadersSQLServer.Connection.ConnectionString = constr
            ta_tblWordStatementsSQLServer.Connection.ConnectionString = constr

            'ta_tblWorddocsSQLServer.Connection.ConnectionString = constr
            ta_tblAuditTrailSQLServer.Connection.ConnectionString = constr

            ta_tblReasonForChangeSQLServer.Connection.ConnectionString = constr
            ta_tblMeaningOfSigSQLServer.Connection.ConnectionString = constr
            ta_tblSaveEventSQLServer.Connection.ConnectionString = constr
            ta_tblDataSystemSQLServer.Connection.ConnectionString = constr
            ta_tblConfigComplianceSQLServer.Connection.ConnectionString = constr
            '02218:
            ta_tblCustomFieldCodesSQLServer.Connection.ConnectionString = constr
            '030008
            ta_TBLWORDSTATEMENTSVERSIONSSQLServer.Connection.ConnectionString = constr
            '03000901
            ta_TBLSECTIONTEMPLATESSQLServer.Connection.ConnectionString = constr
            '030030_01
            ta_TBLFINALREPORTSQLServer.Connection.ConnectionString = constr
            ta_TBLFINALREPORTWORDDOCSSQLServer.Connection.ConnectionString = constr
            '030040_04
            ta_TBLAUTOASSIGNSAMPLESSQLServer.Connection.ConnectionString = constr
            ta_tblAppFigWordDocsSQLSERVER.Connection.ConnectionString = constr
            '030066_02
            ta_TBLSTUDYDOCANALYTESSQLSERVER.Connection.ConnectionString = constr

            'start Study Design
            ta_tblModulesSQLServer.Connection.ConnectionString = constr
            ta_TBLVERSIONSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUANIMALRECEIPTSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUCOMPOUNDSSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUCOMPOUNDSINDSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUCOMPOUNDTYPESQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUPROJECTSSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUSPECIESSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUSPECIESSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUSTUDYDESIGNTYPESQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUSTUDYSPECIESSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUSTUDYSTATSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUASSAYPERSSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUASSAYSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUSPECIESSTRAINSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUDOSEUNITSSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUPKGROUPSSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUPKROUTESSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUPKSUBJECTSSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWURTTIMEPOINTSSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUASSIGNEDCMPDSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUASSIGNEDCMPDLOTSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUSTUDYSCHEDULINGSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUTPCONFIGSQLServer.Connection.ConnectionString = constr
            ta_TBLGUWUTPNAMESCONFIGSQLServer.Connection.ConnectionString = constr
            'ta_QRYGUWUCALENDARSQLServer.Connection.ConnectionString = constr

            ta_TBLGUWUSTUDIESSQLServer.Connection.ConnectionString = constr

        Catch ex As Exception

 
            str1 = "Hmmm." & ChrW(10) & "There seems to be a problem connecting to the StudyDoc datatabase."
            str1 = str1 & ChrW(10) & ChrW(10) & "Please contact your StudyDoc system administrator."
            str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
            str2 = "Critical communication error..."
            If boolFormLoad Then

                frm.Controls("lblErr").Text = str1
                frm.Controls("lblErr").Refresh()
            Else
                Call PositionProgress()
                frmH.lblProgress.Text = str1
                frmH.lblProgress.Visible = True
                frmH.lblProgress.Refresh()

                frmH.panProgress.Visible = True
                frmH.panProgress.Refresh()

            End If

            MsgBox(str1, MsgBoxStyle.Critical, str2)

            End
        

        End Try

        ''debug
        'Dim int1 As Int32
        'dt2 = Now
        'Dim strDt1 As String
        'Dim strDt2 As String

        'strDt1 = Format(dt1, "hh:mm:ss:fff")
        'strDt2 = Format(dt2, "hh:mm:ss:fff")

        'Dim strX As String

        'Try
        '    strX = GetTimeDiff(strDt1, strDt2)
        'Catch ex As Exception
        '    strX = strX
        'End Try

        'MsgBox(strX)



    End Sub


    Function vParse(strX As String, strDel As String)

        '20190218 LEE:

        Dim strParse As String
        Dim int1 As Short

        vParse = Split(strX, strDel)

        'debug
        Dim var1, var2
        var1 = LBound(vParse)
        var2 = UBound(vParse)
        var1 = var1

    End Function


    Function GetTimeDiff(strDate1 As String, strDate2 As String) As String

        '20190218 LEE:

        'Debug.Print GetTimeDiff("08:34:12:744", "08:34:45:734") ' => "0:0:32:990"


        Dim arr1() As String
        Dim arr2() As String

        arr1 = vParse(strDate1, ":")
        arr2 = vParse(strDate2, ":")

        Dim intH As Short
        Dim intM As Short
        Dim intS As Short
        Dim intMS As Short

        intH = arr2(0) - arr1(0)
        intM = arr2(1) - arr1(1)
        intS = arr2(2) - arr1(2)
        intMS = arr2(3) - arr1(3)

        GetTimeDiff = intS & "s  " & intMS & "ms"


    End Function


    Sub DAsRefreshSpecific()

        Dim strF As String

        '' Encloses the keyword in SQL wildcard characters.
        'titleKeyword = "%" & txtTitleKeyword.Text & "%"
        'OleDbDataAdapter1.SelectCommand.Parameters("Title_Keyword").Value = titleKeyword
        'OleDbDataAdapter1.Fill(dsAuthors1)

        strF = "ID_TBLSTUDIES = " & id_tblStudies

        Dim int1 As Int64
        Dim var1
        Dim strM As String

        Try
            If boolGuWuAccess Then

                'Note: .Fill automatically clears existing data

                ta_tblAssignedSamplesAcc.ClearBeforeFill = True
                ta_TBLFINALREPORTAcc.ClearBeforeFill = True
                ta_tblReportTableHeaderConfigAcc.ClearBeforeFill = True
                ta_TBLAUTOASSIGNSAMPLESAcc.ClearBeforeFill = True
                ta_tblReportTableAcc.ClearBeforeFill = True
                ta_tblReportTableAnalytesAcc.ClearBeforeFill = True
                ta_tblAppFigWordDocsAcc.ClearBeforeFill = True


                tblAssignedSamples.Clear()
                tblAssignedSamples.AcceptChanges()
                tblAssignedSamples = ta_tblAssignedSamplesAcc.GetDataByID_TBLSTUDIES(id_tblStudies)
                int1 = tblAssignedSamples.Rows.Count 'DEBUG

                tblFinalReport.Clear()
                tblFinalReport.AcceptChanges()
                tblFinalReport = ta_TBLFINALREPORTAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblFinalReport.Rows.Count 'DEBUG

                tblReportTableHeaderConfig.Clear()
                tblReportTableHeaderConfig.AcceptChanges()
                tblReportTableHeaderConfig = ta_tblReportTableHeaderConfigAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblReportTableHeaderConfig.Rows.Count 'DEBUG

                tblTableProperties.Clear()
                tblTableProperties.AcceptChanges()
                tblTableProperties = ta_tblTablePropertiesAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblTableProperties.Rows.Count 'DEBUG

                tblAutoAssignSamples.Clear()
                tblAutoAssignSamples.AcceptChanges()
                tblAutoAssignSamples = ta_TBLAUTOASSIGNSAMPLESAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblAutoAssignSamples.Rows.Count 'DEBUG


                tblReportTable.Clear()
                tblReportTable.AcceptChanges()
                tblReportTable = ta_tblReportTableAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblReportTable.Rows.Count 'DEBUG

                tblReportTableAnalytes.Clear()
                tblReportTableAnalytes.AcceptChanges()
                tblReportTableAnalytes = ta_tblReportTableAnalytesAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblReportTableAnalytes.Rows.Count 'DEBUG

                tblAppFigWordDocs.Clear()
                tblAppFigWordDocs.AcceptChanges()
                tblAppFigWordDocs = ta_tblAppFigWordDocsAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblAppFigWordDocs.Rows.Count 'DEBUG


            ElseIf boolGuWuSQLServer Then

                ta_tblAssignedSamplesSQLServer.ClearBeforeFill = True
                ta_TBLFINALREPORTSQLServer.ClearBeforeFill = True
                ta_tblReportTableHeaderConfigSQLServer.ClearBeforeFill = True
                ta_TBLAUTOASSIGNSAMPLESSQLServer.ClearBeforeFill = True
                ta_tblReportTableSQLServer.ClearBeforeFill = True
                ta_tblReportTableAnalytesSQLServer.ClearBeforeFill = True
                ta_tblAppFigWordDocsSQLSERVER.ClearBeforeFill = True

                tblAssignedSamples.Clear()
                tblAssignedSamples.AcceptChanges()
                tblAssignedSamples = ta_tblAssignedSamplesSQLServer.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblAssignedSamples.Rows.Count 'DEBUG

                tblFinalReport.Clear()
                tblFinalReport.AcceptChanges()
                tblFinalReport = ta_TBLFINALREPORTSQLServer.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblFinalReport.Rows.Count 'DEBUG

                tblReportTableHeaderConfig.Clear()
                tblReportTableHeaderConfig.AcceptChanges()
                tblReportTableHeaderConfig = ta_tblReportTableHeaderConfigSQLServer.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblReportTableHeaderConfig.Rows.Count 'DEBUG

                tblTableProperties.Clear()
                tblTableProperties.AcceptChanges()
                tblTableProperties = ta_tblTablePropertiesSQLServer.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblTableProperties.Rows.Count 'DEBUG

                tblAutoAssignSamples.Clear()
                tblAutoAssignSamples.AcceptChanges()
                tblAutoAssignSamples = ta_TBLAUTOASSIGNSAMPLESSQLServer.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblAutoAssignSamples.Rows.Count 'DEBUG

                tblReportTable.Clear()
                tblReportTable.AcceptChanges()
                tblReportTable = ta_tblReportTableSQLServer.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblReportTable.Rows.Count 'DEBUG

                tblReportTableAnalytes.Clear()
                tblReportTableAnalytes.AcceptChanges()
                tblReportTableAnalytes = ta_tblReportTableAnalytesSQLServer.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblReportTableAnalytes.Rows.Count 'DEBUG

                tblAppFigWordDocs.Clear()
                tblAppFigWordDocs.AcceptChanges()
                tblAppFigWordDocs = ta_tblAppFigWordDocsSQLSERVER.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                int1 = tblAppFigWordDocs.Rows.Count 'DEBUG

            ElseIf boolGuWuOracle Then

                strM = "Need Oracle stuff here:  DAsRefreshSpecific"
                MsgBox(strM, vbInformation, "Oracle stuff..")

                ''20160607 LEE: come back to this
                'tblAssignedSamples.Clear()
                'tblAssignedSamples.AcceptChanges()
                'tblAssignedSamples = ta_tblAssignedSamples.GetDataByID_TBLSTUDIES(id_tblStudies)
                'int1 = tblAssignedSamples.Rows.Count 'DEBUG

                'tblFinalReport.Clear()
                'tblFinalReport.AcceptChanges()
                'tblFinalReport = ta_TBLFINALREPORT.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                'int1 = tblFinalReport.Rows.Count 'DEBUG

                'tblReportTableHeaderConfig.Clear()
                'tblReportTableHeaderConfig.AcceptChanges()
                'tblReportTableHeaderConfig = ta_tblReportTableHeaderConfigAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                'int1 = tblReportTableHeaderConfig.Rows.Count 'DEBUG

                'tblTableProperties.Clear()
                'tblTableProperties.AcceptChanges()
                'tblTableProperties = ta_tblTablePropertiesAcc.GetDataBy_ID_TBLSTUDIES(id_tblStudies)
                'int1 = tblTableProperties.Rows.Count 'DEBUG

            End If

            'must add columns again
            Call AddCols_tblAss()

        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try




    End Sub

    Function DAsRefresh(ByVal frm As Form) As Boolean

        Dim str1 As String
        '20150528 Larry: This routine doesn't ever seem to be called

        Call RemoveBOOLEXCLSAMPLECHK()

        DAsRefresh = True

        'Dim frmE As New frmErrorMsg
        'Dim frmE As frmSplash1
        Dim str2 As String
        Dim strM As String

        If boolFormLoad Then
            str2 = "...Establishing communication with the Oracle StudyDoc database..."
            frm.Controls("lblErr").Text = str2
            frm.Controls("lblErr").Refresh()
        Else
            Call PositionProgress()
            str2 = "...Refreshing StudyDoc database tables..."
            frmh.lblProgress.Text = str2
            frmh.lblProgress.Visible = True
            frmH.lblProgress.Refresh()

            frmH.panProgress.Visible = True
            frmH.panProgress.Refresh()

        End If
        strM = ""
        'On Error GoTo end1

        'do StudyDoc study-specific queries
        'do this now because columns are added later
        Call DAsRefreshSpecific()

        'use beginloaddata and endloaddata to speed up table filling

        strM = "....tblData"

        Dim ct As Short
        Dim ctMax As Short

        ct = 0
        ctMax = 100

        frmh.pb1.Value = 0
        frmh.pb1.Maximum = ctMax

        Try
            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblData.Clear()
            tblData.AcceptChanges()
            tblData.BeginLoadData()
            ta_tblData.Fill(tblData)
            tblData.EndLoadData()
            strM = "....tblSampleReceipt"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblSampleReceipt.Clear()
            tblSampleReceipt.AcceptChanges()
            tblSampleReceipt.BeginLoadData()
            ta_tblSampleReceipt.Fill(tblSampleReceipt)
            tblSampleReceipt.EndLoadData()
            strM = "....tblTab1"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTab1.Clear()
            tblTab1.AcceptChanges()
            tblTab1.BeginLoadData()
            ta_tblTab1.Fill(tblTab1)
            tblTab1.EndLoadData()
            strM = "....tblConfiguration"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfiguration.Clear()
            tblConfiguration.AcceptChanges()
            tblConfiguration.BeginLoadData()
            ta_tblConfiguration.Fill(tblConfiguration)
            tblConfiguration.EndLoadData()
            strM = "....tblOutstandingItems"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblOutstandingItems.Clear()
            tblOutstandingItems.AcceptChanges()
            tblOutstandingItems.BeginLoadData()
            ta_tblOutstandingItems.Fill(tblOutstandingItems)
            tblOutstandingItems.EndLoadData()
            strM = "....tblPermissions"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPermissions.Clear()
            tblPermissions.AcceptChanges()
            tblPermissions.BeginLoadData()
            ta_tblPermissions.Fill(tblPermissions)
            tblPermissions.EndLoadData()
            strM = "....tblPersonnel"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPersonnel.Clear()
            tblPersonnel.AcceptChanges()
            tblPersonnel.BeginLoadData()
            ta_tblPersonnel.Fill(tblPersonnel)
            tblPersonnel.EndLoadData()
            strM = "....tblUserAccounts"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblUserAccounts.Clear()
            tblUserAccounts.AcceptChanges()
            tblUserAccounts.BeginLoadData()
            ta_tblUserAccounts.Fill(tblUserAccounts)
            tblUserAccounts.EndLoadData()
            strM = "....tblAnalRefStandards"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAnalRefStandards.Clear()
            tblAnalRefStandards.AcceptChanges()
            tblAnalRefStandards.BeginLoadData()
            ta_tblAnalRefStandards.Fill(tblAnalRefStandards)
            tblAnalRefStandards.EndLoadData()
            strM = "....tblAnalyticalRunSummary"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAnalyticalRunSummary.Clear()
            tblAnalyticalRunSummary.AcceptChanges()
            tblAnalyticalRunSummary.BeginLoadData()
            ta_tblAnalyticalRunSummary.Fill(tblAnalyticalRunSummary)
            tblAnalyticalRunSummary.EndLoadData()
            strM = "....tblConfigBodySections"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigBodySections.Clear()
            tblConfigBodySections.AcceptChanges()
            tblConfigBodySections.BeginLoadData()
            ta_tblConfigBodySections.Fill(tblConfigBodySections)
            tblConfigBodySections.EndLoadData()
            strM = "....tblConfigHeaderLookup"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigHeaderLookup.Clear()
            tblConfigHeaderLookup.AcceptChanges()
            tblConfigHeaderLookup.BeginLoadData()
            ta_tblConfigHeaderLookup.Fill(tblConfigHeaderLookup)
            tblConfigHeaderLookup.EndLoadData()
            strM = "....tblConfigReportType"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigReportType.Clear()
            tblConfigReportType.AcceptChanges()
            tblConfigReportType.BeginLoadData()
            ta_tblConfigReportType.Fill(tblConfigReportType)
            tblConfigReportType.EndLoadData()
            strM = "....tblContributingPersonnel"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblContributingPersonnel.Clear()
            tblContributingPersonnel.AcceptChanges()
            tblContributingPersonnel.BeginLoadData()
            ta_tblContributingPersonnel.Fill(tblContributingPersonnel)
            tblContributingPersonnel.EndLoadData()
            strM = "....tblCorporateAddresses"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblCorporateAddresses.Clear()
            tblCorporateAddresses.AcceptChanges()
            tblCorporateAddresses.BeginLoadData()
            ta_tblCorporateAddresses.Fill(tblCorporateAddresses)
            tblCorporateAddresses.EndLoadData()
            strM = "....tblDataTableRowTitles"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDataTableRowTitles.Clear()
            tblDataTableRowTitles.AcceptChanges()
            tblDataTableRowTitles.BeginLoadData()
            ta_tblDataTableRowTitles.Fill(tblDataTableRowTitles)
            tblDataTableRowTitles.EndLoadData()
            strM = "....tblMaxID"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblMaxID.Clear()
            tblMaxID.AcceptChanges()
            tblMaxID.BeginLoadData()
            ta_tblMaxID.Fill(tblMaxID)
            tblMaxID.EndLoadData()
            strM = "....tblMethodValidationData"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblMethodValidationData.Clear()
            tblMethodValidationData.AcceptChanges()
            tblMethodValidationData.BeginLoadData()
            ta_tblMethodValidationData.Fill(tblMethodValidationData)
            tblMethodValidationData.EndLoadData()
            strM = "....tblQATables"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblQATables.Clear()
            tblQATables.AcceptChanges()
            tblQATables.BeginLoadData()
            ta_tblQATables.Fill(tblQATables)
            tblQATables.EndLoadData()
            strM = "....tblReportHistory"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportHistory.Clear()
            tblReportHistory.AcceptChanges()
            tblReportHistory.BeginLoadData()
            ta_tblReportHistory.Fill(tblReportHistory)
            tblReportHistory.EndLoadData()
            strM = "....tblReports"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReports.Clear()
            tblReports.AcceptChanges()
            tblReports.BeginLoadData()
            ta_tblReports.Fill(tblReports)
            tblReports.EndLoadData()
            strM = "....tblReportStatements"

            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblReportStatements.BeginLoadData()
            'ta_tblReportStatements.Fill(tblReportStatements)
            'tblReportStatements.EndLoadData()
            'strM = "....tblReportTable"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportTable.Clear()
            tblReportTable.AcceptChanges()
            tblReportTable.BeginLoadData()
            ta_tblReportTable.Fill(tblReportTable)
            tblReportTable.EndLoadData()
            strM = "....tblReportTableAnalytes"

            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblReportTableAnalytes.BeginLoadData()
            'ta_tblReportTableAnalytes.Fill(tblReportTableAnalytes)
            'tblReportTableAnalytes.EndLoadData()
            'strM = "....tblStudies"

            ct = ct + 1
            'do specific
            frmh.pb1.Value = ct
            frmh.pb1.Refresh()
            'tblReportTableHeaderConfig.BeginLoadData()
            'ta_tblReportTableHeaderConfig.Fill(tblReportTableHeaderConfig)
            'tblReportTableHeaderConfig.EndLoadData()
            'strM = "....tblStudies"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblStudies.Clear()
            tblStudies.AcceptChanges()
            tblStudies.BeginLoadData()
            ta_tblStudies.Fill(tblStudies)
            tblStudies.EndLoadData()
            strM = "....tblTemplates"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTemplates.Clear()
            tblTemplates.AcceptChanges()
            tblTemplates.BeginLoadData()
            ta_tblTemplates.Fill(tblTemplates)
            tblTemplates.EndLoadData()
            strM = "....tblTemplateAttributes"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTemplateAttributes.Clear()
            tblTemplateAttributes.AcceptChanges()
            tblTemplateAttributes.BeginLoadData()
            ta_tblTemplateAttributes.Fill(tblTemplateAttributes)
            tblTemplateAttributes.EndLoadData()
            strM = "....tblConfigReportTables"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigReportTables.Clear()
            tblConfigReportTables.AcceptChanges()
            tblConfigReportTables.BeginLoadData()
            ta_tblConfigReportTables.Fill(tblConfigReportTables)
            tblConfigReportTables.EndLoadData()
            strM = "....tblAddressLabels"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAddressLabels.Clear()
            tblAddressLabels.AcceptChanges()
            tblAddressLabels.BeginLoadData()
            ta_tblAddressLabels.Fill(tblAddressLabels)
            tblAddressLabels.EndLoadData()
            strM = "....tblCorporateNickNames"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblCorporateNickNames.Clear()
            tblCorporateNickNames.AcceptChanges()
            tblCorporateNickNames.BeginLoadData()
            ta_tblCorporateNickNames.Fill(tblCorporateNickNames)
            tblCorporateNickNames.EndLoadData()
            strM = "....tblDropdownBoxContent"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDropdownBoxContent.Clear()
            tblDropdownBoxContent.AcceptChanges()
            tblDropdownBoxContent.BeginLoadData()
            ta_tblDropdownBoxContent.Fill(tblDropdownBoxContent)
            tblDropdownBoxContent.EndLoadData()
            strM = "....tblDropdownBoxName"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDropdownBoxName.Clear()
            tblDropdownBoxName.AcceptChanges()
            tblDropdownBoxName.BeginLoadData()
            ta_tblDropdownBoxName.Fill(tblDropdownBoxName)
            tblDropdownBoxName.EndLoadData()
            strM = "....tblPasswordHistory"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPasswordHistory.Clear()
            tblPasswordHistory.AcceptChanges()
            tblPasswordHistory.BeginLoadData()
            ta_tblPasswordHistory.Fill(tblPasswordHistory)
            tblPasswordHistory.EndLoadData()
            strM = "....tblSummaryData"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblSummaryData.Clear()
            tblSummaryData.AcceptChanges()
            tblSummaryData.BeginLoadData()
            ta_tblSummaryData.Fill(tblSummaryData)
            tblSummaryData.EndLoadData()
            strM = "....tblHooks"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblHooks.Clear()
            tblHooks.AcceptChanges()
            tblHooks.BeginLoadData()
            ta_tblHooks.Fill(tblHooks)
            tblHooks.EndLoadData()
            strM = "....tblAssignedSamples"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmh.pb1.Refresh()
            ''too big. Return specific in DAsRefreshSpecific
            'tblAssignedSamples.BeginLoadData()
            'ta_tblAssignedSamples.Fill(tblAssignedSamples)
            'tblAssignedSamples.EndLoadData()
            strM = "....tblDateFormats"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDateFormats.Clear()
            tblDateFormats.AcceptChanges()
            tblDateFormats.BeginLoadData()
            ta_tblDateFormats.Fill(tblDateFormats)
            tblDateFormats.EndLoadData()
            strM = "....tblAssignedSamplesHelper"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAssignedSamplesHelper.Clear()
            tblAssignedSamplesHelper.AcceptChanges()
            tblAssignedSamplesHelper.BeginLoadData()
            ta_tblAssignedSamplesHelper.Fill(tblAssignedSamplesHelper)
            tblAssignedSamplesHelper.EndLoadData()
            strM = "....tblAppFigs"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmh.pb1.Refresh()
            Try
                tblAppFigs.Clear()
                tblAppFigs.AcceptChanges()
                tblAppFigs.BeginLoadData()
                ta_tblAppFigs.Fill(tblAppFigs)
                tblAppFigs.EndLoadData()
                strM = "....tblConfigAppFigs"

            Catch ex As Exception

            End Try

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigAppFigs.Clear()
            tblConfigAppFigs.AcceptChanges()
            tblConfigAppFigs.BeginLoadData()
            ta_tblConfigAppFigs.Fill(tblConfigAppFigs)
            tblConfigAppFigs.EndLoadData()
            strM = "....tblIncludedRows"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblIncludedRows.Clear()
            tblIncludedRows.AcceptChanges()
            tblIncludedRows.BeginLoadData()
            ta_tblIncludedRows.Fill(tblIncludedRows)
            tblIncludedRows.EndLoadData()
            strM = "....tblTableLegends"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmh.pb1.Refresh()
            'do specific
            'tblTableProperties.BeginLoadData()
            'ta_tblTableProperties.Fill(tblTableProperties)
            'tblTableProperties.EndLoadData()
            'strM = "....tblTableLegends"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTableLegends.Clear()
            tblTableLegends.AcceptChanges()
            tblTableLegends.BeginLoadData()
            ta_tblTableLegends.Fill(tblTableLegends)
            tblTableLegends.EndLoadData()
            strM = "....tblFieldCodes"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblFieldCodes.Clear()
            tblFieldCodes.AcceptChanges()
            tblFieldCodes.BeginLoadData()
            ta_tblFieldCodes.Fill(tblFieldCodes)
            tblFieldCodes.EndLoadData()
            strM = "....tblReportHeaders"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportHeaders.Clear()
            tblReportHeaders.AcceptChanges()
            tblReportHeaders.BeginLoadData()
            ta_tblReportHeaders.Fill(tblReportHeaders)
            tblReportHeaders.EndLoadData()
            strM = "....tblWordStatements"

            ct = ct + 1
            frmh.pb1.Value = ct
            frmH.pb1.Refresh()
            tblWordStatements.Clear()
            tblWordStatements.AcceptChanges()
            tblWordStatements.BeginLoadData()
            ta_tblWordStatements.Fill(tblWordStatements)
            tblWordStatements.EndLoadData()
            strM = "....TBLWORDSTATEMENTSVERSIONS"

            Call AddCols_tblAss() 'add BOOLEXCLSAMPLECHK back

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            'tblAuditTrail.BeginLoadData()
            str1 = "SELECT * FROM TBLAUDITTRAIL WHERE ID_TBLAUDITTRAIL < 0"
            Dim con As New ADODB.Connection
            Try
                con.Open(constrIni)
                Dim rs1 As New ADODB.Recordset
                rs1.CursorLocation = CursorLocationEnum.adUseClient
                rs1.Open(str1, con, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
                rs1.ActiveConnection = Nothing
                tblAuditTrail.Clear()
                tblAuditTrail.AcceptChanges()
                tblAuditTrail.BeginLoadData()
                daDoPr.Fill(tblAuditTrail, rs1)
                tblAuditTrail.EndLoadData()
                rs1.Close()
                rs1 = Nothing
                con.Close()
                con = Nothing
            Catch ex As Exception
                strM = "constrIni needs to be modified for '" & str1 & "' in DAsRefresh"
                MsgBox(strM, vbInformation, "Note...")
            End Try

            strM = "Note that a lot more things need to be added at the end of DAsRefresh"
            MsgBox(str1, vbInformation, "Note...")


            ct = ct + 1
            ''20150501: Don't load tblWordDocs - will get too big
            ''load when needed as a filtered set
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblWorddocs.BeginLoadData()
            'ta_tblWorddocs.Fill(tblWorddocs)
            'tblWorddocs.EndLoadData()
            'strM = "....Done"

            ''030008
            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'TBLWORDSTATEMENTSVERSIONS.BeginLoadData()
            'ta_TBLWORDSTATEMENTSVERSIONS.Fill(TBLWORDSTATEMENTSVERSIONS)
            'TBLWORDSTATEMENTSVERSIONS.EndLoadData()
            'strM = "....TBLSECTIONTEMPLATES"

            '03000901
            ''come back to this later
            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'TBLSECTIONTEMPLATES.BeginLoadData()
            'ta_TBLSECTIONTEMPLATES.Fill(TBLSECTIONTEMPLATES)
            'TBLSECTIONTEMPLATES.EndLoadData()
            'strM = "....tblModules"

            'start Study Design

            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblModules.BeginLoadData()
            'ta_tblModules.Fill(tblModules)
            'tblModules.EndLoadData()
            'strM = "....tblPermissionsStDes"

            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblPermissionsStDes.BeginLoadData()
            'ta_tblPermissionsStDes.Fill(tblPermissionsStDes)
            'tblPermissionsStDes.EndLoadData()
            'strM = "....Done"


            frmh.pb1.Value = frmh.pb1.Maximum
            frmh.pb1.Refresh()
            'On Error GoTo 0

            If boolFormLoad Then
                frm.Controls("lblErr").Text = ""
                frm.Controls("lblErr").Refresh()
                ''SendKeys.Send("%")
                'frmE.Dispose()
            Else
                'frmh.lblProgress.Visible = False
                frmh.lblProgress.Text = ""
                frmh.Refresh()
            End If



        Catch ex As Exception


            str1 = "Hmmm." & Chr(10) & "There seems to be a problem retrieving data from the StudyDoc datatabase."
            str1 = str1 & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
            str1 = str1 & Chr(10) & strM & "...."
            str1 = str1 & Chr(10) & ex.Message
            str1 = str1 & Chr(10) & "DAsRefresh"
            str2 = "Critical communication error..."

            If boolFormLoad Then
                frm.Controls("lblErr").Text = str1
                frm.Controls("lblErr").Refresh()
                'frmE.Visible = False
                'frmE.cmdOK.Visible = True
                'frmE.pb1.Visible = False
                'frmE.TimerE.Enabled = False
                'frmE.ShowDialog()
                'SendKeys.Send("%")
            Else
                Call PositionProgress()
                frmh.lblProgress.Visible = True
                frmh.lblProgress.Text = str1
                frmH.lblProgress.Refresh()

                frmH.panProgress.Visible = True
                frmH.panProgress.Refresh()

            End If
            'Dim frmE As New frmErrorMsg
            MsgBox(str1, MsgBoxStyle.Critical, str2)

            Dim dt As Date
            Dim dt1 As Date
            dt = Now
            dt1 = DateAdd(DateInterval.Second, 1, dt)
            Do Until dt > dt1
                dt = Now
            Loop

            DAsRefresh = False
            Exit Function

        End Try

    End Function

    Function DAsRefreshAcc(ByVal frm As Form) As Boolean

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String

        Dim int1 As Int64

        'MsgBox("Here1")

        Call RemoveBOOLEXCLSAMPLECHK()

        DAsRefreshAcc = True

        'Dim frmE As New frmErrorMsg
        'Dim frmE As frmSplash1
        Dim strM As String

        If boolFormLoad Then
            str2 = "...Establishing communication with the LABIntegrity" & ChrW(8482) & " StudyDoc" & ChrW(8482) & " Microsoft" & ChrW(8482) & " Access database..."
            frm.Controls("lblErr").Text = str2
            frm.Controls("lblErr").Refresh()
        Else
            Call PositionProgress()
            str2 = "...Refreshing StudyDoc database tables..."
            frmh.lblProgress.Text = str2
            frmh.lblProgress.Visible = True
            frmH.lblProgress.Refresh()

            frmH.panProgress.Visible = True
            frmH.panProgress.Refresh()

        End If
        strM = ""
        'On Error GoTo end1

        'do StudyDoc study-specific queries
        'do this now because columns are added later
        Call DAsRefreshSpecific()

        'use beginloaddata and endloaddata to speed up table filling

        strM = "....tblData"

        Dim ct As Short
        Dim ctMax As Short

        ct = 0
        ctMax = 100

        frmh.pb1.Value = 0
        frmh.pb1.Maximum = ctMax

        Try

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblData.Clear()

            tblData.AcceptChanges()
            tblData.BeginLoadData()
            ta_tblDataAcc.ClearBeforeFill = True
            ta_tblDataAcc.Fill(tblData)
            tblData.EndLoadData()
            strM = "....tblSampleReceipt"

            'debug
            Dim intAA As Int32
            intAA = tblData.Rows.Count

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblSampleReceipt.Clear()
            tblSampleReceipt.AcceptChanges()
            tblSampleReceipt.BeginLoadData()
            ta_tblSampleReceiptAcc.ClearBeforeFill = True
            ta_tblSampleReceiptAcc.Fill(tblSampleReceipt)
            tblSampleReceipt.EndLoadData()
            strM = "....tblTab1"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTab1.Clear()
            tblTab1.AcceptChanges()
            tblTab1.BeginLoadData()
            ta_tblTab1Acc.ClearBeforeFill = True
            ta_tblTab1Acc.Fill(tblTab1)
            tblTab1.EndLoadData()
            strM = "....tblConfiguration"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfiguration.Clear()
            tblConfiguration.AcceptChanges()
            tblConfiguration.BeginLoadData()
            ta_tblConfigurationAcc.ClearBeforeFill = True
            ta_tblConfigurationAcc.Fill(tblConfiguration)
            tblConfiguration.EndLoadData()
            strM = "....tblOutstandingItems"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblOutstandingItems.Clear()
            tblOutstandingItems.AcceptChanges()
            tblOutstandingItems.BeginLoadData()
            ta_tblOutstandingItemsAcc.ClearBeforeFill = True
            ta_tblOutstandingItemsAcc.Fill(tblOutstandingItems)
            tblOutstandingItems.EndLoadData()
            strM = "....tblPermissions"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPermissions.Clear()
            tblPermissions.AcceptChanges()
            tblPermissions.BeginLoadData()
            ta_tblPermissionsAcc.ClearBeforeFill = True
            ta_tblPermissionsAcc.Fill(tblPermissions)
            tblPermissions.EndLoadData()
            strM = "....tblPersonnel"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPersonnel.Clear()
            tblPersonnel.AcceptChanges()
            tblPersonnel.BeginLoadData()
            ta_tblPersonnelAcc.ClearBeforeFill = True
            ta_tblPersonnelAcc.Fill(tblPersonnel)
            tblPersonnel.EndLoadData()
            strM = "....tblUserAccounts"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblUserAccounts.Clear()
            tblUserAccounts.AcceptChanges()
            tblUserAccounts.BeginLoadData()
            ta_tblUserAccountsAcc.ClearBeforeFill = True
            ta_tblUserAccountsAcc.Fill(tblUserAccounts)
            tblUserAccounts.EndLoadData()
            strM = "....tblAnalRefStandards"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAnalRefStandards.Clear()
            tblAnalRefStandards.AcceptChanges()
            tblAnalRefStandards.BeginLoadData()
            ta_tblAnalRefStandardsAcc.ClearBeforeFill = True
            ta_tblAnalRefStandardsAcc.Fill(tblAnalRefStandards)
            tblAnalRefStandards.EndLoadData()
            strM = "....tblAnalyticalRunSummary"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAnalyticalRunSummary.Clear()
            tblAnalyticalRunSummary.AcceptChanges()
            tblAnalyticalRunSummary.BeginLoadData()
            ta_tblAnalyticalRunSummaryAcc.ClearBeforeFill = True
            ta_tblAnalyticalRunSummaryAcc.Fill(tblAnalyticalRunSummary)
            tblAnalyticalRunSummary.EndLoadData()
            strM = "....tblConfigBodySections"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigBodySections.Clear()
            tblConfigBodySections.AcceptChanges()
            tblConfigBodySections.BeginLoadData()
            ta_tblConfigBodySectionsAcc.ClearBeforeFill = True
            ta_tblConfigBodySectionsAcc.Fill(tblConfigBodySections)
            tblConfigBodySections.EndLoadData()
            strM = "....tblConfigHeaderLookup"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigHeaderLookup.Clear()
            tblConfigHeaderLookup.AcceptChanges()
            tblConfigHeaderLookup.BeginLoadData()
            ta_tblConfigHeaderLookupAcc.ClearBeforeFill = True
            ta_tblConfigHeaderLookupAcc.Fill(tblConfigHeaderLookup)
            tblConfigHeaderLookup.EndLoadData()
            strM = "....tblConfigReportType"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigReportType.Clear()
            tblConfigReportType.AcceptChanges()
            tblConfigReportType.BeginLoadData()
            ta_tblConfigReportTypeAcc.ClearBeforeFill = True
            ta_tblConfigReportTypeAcc.Fill(tblConfigReportType)
            tblConfigReportType.EndLoadData()
            strM = "....tblContributingPersonnel"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblContributingPersonnel.Clear()
            tblContributingPersonnel.AcceptChanges()
            tblContributingPersonnel.BeginLoadData()
            ta_tblContributingPersonnelAcc.ClearBeforeFill = True
            ta_tblContributingPersonnelAcc.Fill(tblContributingPersonnel)
            tblContributingPersonnel.EndLoadData()
            strM = "....tblCorporateAddresses"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblCorporateAddresses.Clear()
            tblCorporateAddresses.AcceptChanges()
            tblCorporateAddresses.BeginLoadData()
            ta_tblCorporateAddressesAcc.ClearBeforeFill = True
            ta_tblCorporateAddressesAcc.Fill(tblCorporateAddresses)
            tblCorporateAddresses.EndLoadData()
            strM = "....tblDataTableRowTitles"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDataTableRowTitles.Clear()
            tblDataTableRowTitles.AcceptChanges()
            tblDataTableRowTitles.BeginLoadData()
            ta_tblDataTableRowTitlesAcc.ClearBeforeFill = True
            ta_tblDataTableRowTitlesAcc.Fill(tblDataTableRowTitles)
            tblDataTableRowTitles.EndLoadData()
            strM = "....tblMaxID"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblMaxID.Clear()
            tblMaxID.AcceptChanges()
            tblMaxID.BeginLoadData()
            ta_tblMaxIDAcc.ClearBeforeFill = True
            ta_tblMaxIDAcc.Fill(tblMaxID)
            tblMaxID.EndLoadData()
            strM = "....tblMethodValidationData"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblMethodValidationData.Clear()
            tblMethodValidationData.AcceptChanges()
            tblMethodValidationData.BeginLoadData()
            ta_tblMethodValidationDataAcc.ClearBeforeFill = True
            ta_tblMethodValidationDataAcc.Fill(tblMethodValidationData)
            tblMethodValidationData.EndLoadData()
            strM = "....tblQATables"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblQATables.Clear()
            tblQATables.AcceptChanges()
            tblQATables.BeginLoadData()
            ta_tblQATablesAcc.ClearBeforeFill = True
            ta_tblQATablesAcc.Fill(tblQATables)
            tblQATables.EndLoadData()
            strM = "....tblReportHistory"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportHistory.Clear()
            tblReportHistory.AcceptChanges()
            tblReportHistory.BeginLoadData()
            ta_tblReportHistoryAcc.ClearBeforeFill = True
            ta_tblReportHistoryAcc.Fill(tblReportHistory)
            tblReportHistory.EndLoadData()
            strM = "....tblReports"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReports.Clear()
            tblReports.AcceptChanges()
            tblReports.BeginLoadData()
            ta_tblReportsAcc.ClearBeforeFill = True
            ta_tblReportsAcc.Fill(tblReports)
            tblReports.EndLoadData()
            strM = "....tblReportStatements"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportStatements.Clear()
            tblReportStatements.AcceptChanges()
            tblReportStatements.BeginLoadData()
            ta_tblReportStatementsAcc.ClearBeforeFill = True
            ta_tblReportStatementsAcc.Fill(tblReportStatements)
            tblReportStatements.EndLoadData()
            strM = "....tblReportTable"

            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblReportTable.CLEAR()
            'tblReportTable.ACCEPTCHANGES()
            'tblReportTable.BeginLoadData()
            'ta_tblReportTableAcc.ClearBeforeFill = True
            'ta_tblReportTableAcc.Fill(tblReportTable)
            'tblReportTable.EndLoadData()
            'strM = "....tblReportTableAnalytes"

            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblReportTableAnalytes.CLEAR()
            'tblReportTableAnalytes.ACCEPTCHANGES()
            'tblReportTableAnalytes.BeginLoadData()
            'ta_tblReportTableAnalytesAcc.ClearBeforeFill = True
            'ta_tblReportTableAnalytesAcc.Fill(tblReportTableAnalytes)
            'tblReportTableAnalytes.EndLoadData()
            'strM = "....tblStudies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            'do at study load
            'do specific
            'tblReportTableHeaderConfig.CLEAR()
            'tblReportTableHeaderConfig.ACCEPTCHANGES()
            'tblReportTableHeaderConfig.BeginLoadData()
            'ta_tblReportTableHeaderConfigAcc.ClearBeforeFill = True
            'ta_tblReportTableHeaderConfigAcc.Fill(tblReportTableHeaderConfig)
            'tblReportTableHeaderConfig.EndLoadData()
            'strM = "....tblStudies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblStudies.Clear()
            tblStudies.AcceptChanges()
            tblStudies.BeginLoadData()
            ta_tblStudiesAcc.ClearBeforeFill = True
            ta_tblStudiesAcc.Fill(tblStudies)
            tblStudies.EndLoadData()
            strM = "....tblTemplates"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTemplates.Clear()
            tblTemplates.AcceptChanges()
            tblTemplates.BeginLoadData()
            ta_tblTemplatesAcc.ClearBeforeFill = True
            ta_tblTemplatesAcc.Fill(tblTemplates)
            tblTemplates.EndLoadData()
            strM = "....tblTemplateAttributes"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTemplateAttributes.Clear()
            tblTemplateAttributes.AcceptChanges()
            tblTemplateAttributes.BeginLoadData()
            ta_tblTemplateAttributesAcc.ClearBeforeFill = True
            ta_tblTemplateAttributesAcc.Fill(tblTemplateAttributes)
            tblTemplateAttributes.EndLoadData()
            strM = "....tblConfigReportTables"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigReportTables.Clear()
            tblConfigReportTables.AcceptChanges()
            tblConfigReportTables.BeginLoadData()
            ta_tblConfigReportTablesAcc.ClearBeforeFill = True
            ta_tblConfigReportTablesAcc.Fill(tblConfigReportTables)
            tblConfigReportTables.EndLoadData()
            strM = "....tblAddressLabels"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAddressLabels.Clear()
            tblAddressLabels.AcceptChanges()
            tblAddressLabels.BeginLoadData()
            ta_tblAddressLabelsAcc.ClearBeforeFill = True
            ta_tblAddressLabelsAcc.Fill(tblAddressLabels)
            tblAddressLabels.EndLoadData()
            strM = "....tblCorporateNickNames"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblCorporateNickNames.Clear()
            tblCorporateNickNames.AcceptChanges()
            tblCorporateNickNames.BeginLoadData()
            ta_tblCorporateNickNamesAcc.ClearBeforeFill = True
            ta_tblCorporateNickNamesAcc.Fill(tblCorporateNickNames)
            tblCorporateNickNames.EndLoadData()
            strM = "....tblDropdownBoxContent"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDropdownBoxContent.Clear()
            tblDropdownBoxContent.AcceptChanges()
            tblDropdownBoxContent.BeginLoadData()
            ta_tblDropdownBoxContentAcc.ClearBeforeFill = True
            ta_tblDropdownBoxContentAcc.Fill(tblDropdownBoxContent)
            tblDropdownBoxContent.EndLoadData()
            strM = "....tblDropdownBoxName"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDropdownBoxName.Clear()
            tblDropdownBoxName.AcceptChanges()
            tblDropdownBoxName.BeginLoadData()
            ta_tblDropdownBoxNameAcc.ClearBeforeFill = True
            ta_tblDropdownBoxNameAcc.Fill(tblDropdownBoxName)
            tblDropdownBoxName.EndLoadData()
            strM = "....tblPasswordHistory"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPasswordHistory.Clear()
            tblPasswordHistory.AcceptChanges()
            tblPasswordHistory.BeginLoadData()
            ta_tblPasswordHistoryAcc.ClearBeforeFill = True
            ta_tblPasswordHistoryAcc.Fill(tblPasswordHistory)
            tblPasswordHistory.EndLoadData()
            strM = "....tblSummaryData"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblSummaryData.Clear()
            tblSummaryData.AcceptChanges()
            tblSummaryData.BeginLoadData()
            ta_tblSummaryDataAcc.ClearBeforeFill = True
            ta_tblSummaryDataAcc.Fill(tblSummaryData)
            tblSummaryData.EndLoadData()
            strM = "....tblHooks"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblHooks.Clear()
            tblHooks.AcceptChanges()
            tblHooks.BeginLoadData()
            ta_tblHooksAcc.ClearBeforeFill = True
            ta_tblHooksAcc.Fill(tblHooks)
            tblHooks.EndLoadData()
            strM = "....tblAssignedSamples"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            ''too big. Return specific in DAsRefreshSpecific


            'tblAssignedSamples.BeginLoadData()

            'ta_tblAssignedSamplesAcc.Fill(tblAssignedSamples)
            'tblAssignedSamples.EndLoadData()
            strM = "....tblDateFormats"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDateFormats.Clear()
            tblDateFormats.AcceptChanges()
            tblDateFormats.BeginLoadData()
            ta_tblDateFormatsAcc.ClearBeforeFill = True
            ta_tblDateFormatsAcc.Fill(tblDateFormats)
            tblDateFormats.EndLoadData()
            strM = "....tblAssignedSamplesHelper"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAssignedSamplesHelper.Clear()
            tblAssignedSamplesHelper.AcceptChanges()
            tblAssignedSamplesHelper.BeginLoadData()
            ta_tblAssignedSamplesHelperAcc.ClearBeforeFill = True
            ta_tblAssignedSamplesHelperAcc.Fill(tblAssignedSamplesHelper)
            tblAssignedSamplesHelper.EndLoadData()
            strM = "....tblAppFigs"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            Try
                tblAppFigs.Clear()
                tblAppFigs.AcceptChanges()
                tblAppFigs.BeginLoadData()
                ta_tblAppFigsAcc.ClearBeforeFill = True
                ta_tblAppFigsAcc.Fill(tblAppFigs)
                tblAppFigs.EndLoadData()
                strM = "....tblConfigAppFigs"

            Catch ex As Exception

            End Try

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigAppFigs.Clear()
            tblConfigAppFigs.AcceptChanges()
            tblConfigAppFigs.BeginLoadData()
            ta_tblConfigAppFigsAcc.ClearBeforeFill = True
            ta_tblConfigAppFigsAcc.Fill(tblConfigAppFigs)
            tblConfigAppFigs.EndLoadData()
            strM = "....tblIncludedRows"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblIncludedRows.Clear()
            tblIncludedRows.AcceptChanges()
            tblIncludedRows.BeginLoadData()
            ta_tblIncludedRowsAcc.ClearBeforeFill = True
            ta_tblIncludedRowsAcc.Fill(tblIncludedRows)
            tblIncludedRows.EndLoadData()
            strM = "....tblTableLegends"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            'do specific
            'tblTableProperties.CLEAR()
            'tblTableProperties.ACCEPTCHANGES()
            'tblTableProperties.BeginLoadData()
            'ta_tblTablePropertiesAcc.ClearBeforeFill = True
            'ta_tblTablePropertiesAcc.Fill(tblTableProperties)
            'tblTableProperties.EndLoadData()
            'strM = "....tblTableLegends"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTableLegends.Clear()
            tblTableLegends.AcceptChanges()
            tblTableLegends.BeginLoadData()
            ta_tblTableLegendsAcc.ClearBeforeFill = True
            ta_tblTableLegendsAcc.Fill(tblTableLegends)
            tblTableLegends.EndLoadData()
            strM = "....tblFieldCodes"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblFieldCodes.Clear()
            tblFieldCodes.AcceptChanges()
            tblFieldCodes.BeginLoadData()
            ta_tblFieldCodesAcc.ClearBeforeFill = True
            ta_tblFieldCodesAcc.Fill(tblFieldCodes)
            tblFieldCodes.EndLoadData()
            strM = "....tblReportHeaders"

            'debug
            int1 = tblFieldCodes.Rows.Count
            int1 = int1

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportHeaders.Clear()
            tblReportHeaders.AcceptChanges()
            tblReportHeaders.BeginLoadData()
            ta_tblReportHeadersAcc.ClearBeforeFill = True
            ta_tblReportHeadersAcc.Fill(tblReportHeaders)
            tblReportHeaders.EndLoadData()
            strM = "....tblWordStatements"

            'debug
            int1 = tblReportHeaders.Rows.Count
            int1 = int1

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblWordStatements.Clear()
            tblWordStatements.AcceptChanges()
            tblWordStatements.BeginLoadData()
            ta_tblWordStatementsAcc.ClearBeforeFill = True
            ta_tblWordStatementsAcc.Fill(tblWordStatements)
            tblWordStatements.EndLoadData()
            strM = "....tblWordDocs"


            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblWorddocs.CLEAR()
            'tblWorddocs.ACCEPTCHANGES()
            'tblWorddocs.BeginLoadData()
            'ta_tblWorddocsAcc.ClearBeforeFill = True
            'ta_tblWorddocsAcc.Fill(tblWorddocs)
            'tblWorddocs.EndLoadData()
            'strM = "....tblAuditTrail"


            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            'tblAuditTrail.BeginLoadData()
            str1 = "SELECT * FROM TBLAUDITTRAIL WHERE ID_TBLAUDITTRAIL < 0"
            Dim con As New ADODB.Connection
            con.Open(constrIni)
            Dim rs1 As New ADODB.Recordset
            rs1.CursorLocation = CursorLocationEnum.adUseClient
            rs1.Open(str1, con, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
            rs1.ActiveConnection = Nothing
            tblAuditTrail.Clear()
            tblAuditTrail.AcceptChanges()
            tblAuditTrail.BeginLoadData()
            daDoPr.Fill(tblAuditTrail, rs1)
            tblAuditTrail.EndLoadData()
            rs1.Close()
            rs1 = Nothing
            con.Close()
            con = Nothing





            'ta_tblAuditTrailAcc.Fill(tblAuditTrail)
            'tblAuditTrail.EndLoadData()
            strM = "....tblReasonForChange"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReasonForChange.Clear()
            tblReasonForChange.AcceptChanges()
            tblReasonForChange.BeginLoadData()
            ta_tblReasonForChangeAcc.ClearBeforeFill = True
            ta_tblReasonForChangeAcc.Fill(tblReasonForChange)
            tblReasonForChange.EndLoadData()
            strM = "....tblMeaningOfSig"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblMeaningOfSig.Clear()
            tblMeaningOfSig.AcceptChanges()
            tblMeaningOfSig.BeginLoadData()
            ta_tblMeaningOfSigAcc.ClearBeforeFill = True
            ta_tblMeaningOfSigAcc.Fill(tblMeaningOfSig)
            tblMeaningOfSig.EndLoadData()
            strM = "....tblSaveEvent"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblSaveEvent.Clear()
            tblSaveEvent.AcceptChanges()
            tblSaveEvent.BeginLoadData()
            ta_tblSaveEventAcc.ClearBeforeFill = True
            ta_tblSaveEventAcc.Fill(tblSaveEvent)
            tblSaveEvent.EndLoadData()
            strM = "....tblDataSystem"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDataSystem.Clear()
            tblDataSystem.AcceptChanges()
            tblDataSystem.BeginLoadData()
            ta_tblDataSystemAcc.ClearBeforeFill = True
            ta_tblDataSystemAcc.Fill(tblDataSystem)
            tblDataSystem.EndLoadData()
            strM = "....tblCustomFieldCodes"

            '02218
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblCustomFieldCodes.Clear()
            tblCustomFieldCodes.AcceptChanges()
            tblCustomFieldCodes.BeginLoadData()
            ta_tblCustomFieldCodesAcc.ClearBeforeFill = True
            ta_tblCustomFieldCodesAcc.Fill(tblCustomFieldCodes)
            tblCustomFieldCodes.EndLoadData()
            strM = "....tblConfigCompliance"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigCompliance.Clear()
            tblConfigCompliance.AcceptChanges()
            tblConfigCompliance.BeginLoadData()
            ta_tblConfigComplianceAcc.ClearBeforeFill = True
            ta_tblConfigComplianceAcc.Fill(tblConfigCompliance)
            tblConfigCompliance.EndLoadData()
            strM = "....TBLWORDSTATEMENTSVERSIONS"

            '030008
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLWORDSTATEMENTSVERSIONS.Clear()
            TBLWORDSTATEMENTSVERSIONS.AcceptChanges()
            TBLWORDSTATEMENTSVERSIONS.BeginLoadData()
            ta_TBLWORDSTATEMENTSVERSIONSAcc.ClearBeforeFill = True
            ta_TBLWORDSTATEMENTSVERSIONSAcc.Fill(TBLWORDSTATEMENTSVERSIONS)
            TBLWORDSTATEMENTSVERSIONS.EndLoadData()
            strM = "....TBLSECTIONTEMPLATES"


            '03000901
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLSECTIONTEMPLATES.Clear()
            TBLSECTIONTEMPLATES.AcceptChanges()
            TBLSECTIONTEMPLATES.BeginLoadData()
            ta_TBLSECTIONTEMPLATESAcc.ClearBeforeFill = True
            ta_TBLSECTIONTEMPLATESAcc.Fill(TBLSECTIONTEMPLATES)
            TBLSECTIONTEMPLATES.EndLoadData()
            strM = "....TBLSTUDYDOCANALYTES"

            '030046602
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLSTUDYDOCANALYTES.Clear()
            TBLSTUDYDOCANALYTES.AcceptChanges()
            TBLSTUDYDOCANALYTES.BeginLoadData()
            ta_TBLSTUDYDOCANALYTESAcc.ClearBeforeFill = True
            ta_TBLSTUDYDOCANALYTESAcc.Fill(TBLSTUDYDOCANALYTES)
            TBLSTUDYDOCANALYTES.EndLoadData()
            strM = "....tblModules"

            '030030_01
            'Note: do not fill TBLFINALREPORT or TBLFINALREPORTWORDDOCS or TBLREPORTTABLEHEADERCONFIG here
            'fill in WordDoc module when user saves report
            'fill in TBLREPORTTABLEHEADERCONFIG when study is loaded


            'start Study Design

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblModules.Clear()
            tblModules.AcceptChanges()
            tblModules.BeginLoadData()
            ta_tblModulesAcc.ClearBeforeFill = True
            ta_tblModulesAcc.Fill(tblModules)
            tblModules.EndLoadData()
            strM = "....tblVersion"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblVersion.Clear()
            tblVersion.AcceptChanges()
            tblVersion.BeginLoadData()
            ta_TBLVERSIONAcc.ClearBeforeFill = True
            ta_TBLVERSIONAcc.Fill(tblVersion)
            tblVersion.EndLoadData()
            strM = "....tblGuWuAnimalReceipt"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuAnimalReceipt.Clear()
            tblGuWuAnimalReceipt.AcceptChanges()
            tblGuWuAnimalReceipt.BeginLoadData()
            ta_TBLGUWUANIMALRECEIPTAcc.ClearBeforeFill = True
            ta_TBLGUWUANIMALRECEIPTAcc.Fill(tblGuWuAnimalReceipt)
            tblGuWuAnimalReceipt.EndLoadData()
            strM = "....tblGuWuCompounds"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuCompounds.Clear()
            tblGuWuCompounds.AcceptChanges()
            tblGuWuCompounds.BeginLoadData()
            ta_TBLGUWUCOMPOUNDSAcc.ClearBeforeFill = True
            ta_TBLGUWUCOMPOUNDSAcc.Fill(tblGuWuCompounds)
            tblGuWuCompounds.EndLoadData()
            strM = "....tblGuWuCompoundsInd"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuCompoundsInd.Clear()
            tblGuWuCompoundsInd.AcceptChanges()
            tblGuWuCompoundsInd.BeginLoadData()
            ta_TBLGUWUCOMPOUNDSINDAcc.ClearBeforeFill = True
            ta_TBLGUWUCOMPOUNDSINDAcc.Fill(tblGuWuCompoundsInd)
            tblGuWuCompoundsInd.EndLoadData()
            strM = "....tblGuWuCompoundType"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuCompoundType.Clear()
            tblGuWuCompoundType.AcceptChanges()
            tblGuWuCompoundType.BeginLoadData()
            ta_TBLGUWUCOMPOUNDTYPEAcc.ClearBeforeFill = True
            ta_TBLGUWUCOMPOUNDTYPEAcc.Fill(tblGuWuCompoundType)
            tblGuWuCompoundType.EndLoadData()
            strM = "....tblGuWuProjects"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuProjects.Clear()
            tblGuWuProjects.AcceptChanges()
            tblGuWuProjects.BeginLoadData()
            ta_TBLGUWUPROJECTSAcc.ClearBeforeFill = True
            ta_TBLGUWUPROJECTSAcc.Fill(tblGuWuProjects)
            tblGuWuProjects.EndLoadData()
            strM = "....tblGuWuSpecies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuSpecies.Clear()
            tblGuWuSpecies.AcceptChanges()
            tblGuWuSpecies.BeginLoadData()
            ta_TBLGUWUSPECIESAcc.ClearBeforeFill = True
            ta_TBLGUWUSPECIESAcc.Fill(tblGuWuSpecies)
            tblGuWuSpecies.EndLoadData()
            strM = "....tblGuWuStudies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuStudies.Clear()
            tblGuWuStudies.AcceptChanges()
            tblGuWuStudies.BeginLoadData()
            ta_TBLGUWUSTUDIESAcc.ClearBeforeFill = True
            ta_TBLGUWUSTUDIESAcc.Fill(tblGuWuStudies)
            tblGuWuStudies.EndLoadData()
            strM = "....tblGuWuStudyDesignType"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuStudyDesignType.Clear()
            tblGuWuStudyDesignType.AcceptChanges()
            tblGuWuStudyDesignType.BeginLoadData()
            ta_TBLGUWUSTUDYDESIGNTYPEAcc.ClearBeforeFill = True
            ta_TBLGUWUSTUDYDESIGNTYPEAcc.Fill(tblGuWuStudyDesignType)
            tblGuWuStudyDesignType.EndLoadData()
            strM = "....tblGuWuStudySpecies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuStudySpecies.Clear()
            tblGuWuStudySpecies.AcceptChanges()
            tblGuWuStudySpecies.BeginLoadData()
            ta_TBLGUWUSTUDYSPECIESAcc.ClearBeforeFill = True
            ta_TBLGUWUSTUDYSPECIESAcc.Fill(tblGuWuStudySpecies)
            tblGuWuStudySpecies.EndLoadData()
            strM = "....tblGuWuStudyStat"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuStudyStat.Clear()
            tblGuWuStudyStat.AcceptChanges()
            tblGuWuStudyStat.BeginLoadData()
            ta_TBLGUWUSTUDYSTATAcc.ClearBeforeFill = True
            ta_TBLGUWUSTUDYSTATAcc.Fill(tblGuWuStudyStat)
            tblGuWuStudyStat.EndLoadData()
            strM = "....TBLGUWUASSAYPERS"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLGUWUASSAYPERS.Clear()
            TBLGUWUASSAYPERS.AcceptChanges()
            TBLGUWUASSAYPERS.BeginLoadData()
            ta_TBLGUWUASSAYPERSAcc.ClearBeforeFill = True
            ta_TBLGUWUASSAYPERSAcc.Fill(TBLGUWUASSAYPERS)
            TBLGUWUASSAYPERS.EndLoadData()
            strM = "....tblGuWuAssay"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuAssay.Clear()
            tblGuWuAssay.AcceptChanges()
            tblGuWuAssay.BeginLoadData()
            ta_TBLGUWUASSAYAcc.ClearBeforeFill = True
            ta_TBLGUWUASSAYAcc.Fill(tblGuWuAssay)
            tblGuWuAssay.EndLoadData()
            strM = "....TBLGUWUSPECIESSTRAIN"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuSpeciesStrain.Clear()
            tblGuWuSpeciesStrain.AcceptChanges()
            tblGuWuSpeciesStrain.BeginLoadData()
            ta_TBLGUWUSPECIESSTRAINAcc.ClearBeforeFill = True
            ta_TBLGUWUSPECIESSTRAINAcc.Fill(tblGuWuSpeciesStrain)
            tblGuWuSpeciesStrain.EndLoadData()
            strM = "....tblGuWuDoseUnits"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuDoseUnits.Clear()
            tblGuWuDoseUnits.AcceptChanges()
            tblGuWuDoseUnits.BeginLoadData()
            ta_TBLGUWUDOSEUNITSAcc.ClearBeforeFill = True
            ta_TBLGUWUDOSEUNITSAcc.Fill(tblGuWuDoseUnits)
            tblGuWuDoseUnits.EndLoadData()
            strM = "....tblGuWuPKGroups"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuPKGroups.Clear()
            tblGuWuPKGroups.AcceptChanges()
            tblGuWuPKGroups.BeginLoadData()
            ta_TBLGUWUPKGROUPSAcc.ClearBeforeFill = True
            ta_TBLGUWUPKGROUPSAcc.Fill(tblGuWuPKGroups)
            tblGuWuPKGroups.EndLoadData()
            strM = "....tblGuWuPKRoutes"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuPKRoutes.Clear()
            tblGuWuPKRoutes.AcceptChanges()
            tblGuWuPKRoutes.BeginLoadData()
            ta_TBLGUWUPKROUTESAcc.ClearBeforeFill = True
            ta_TBLGUWUPKROUTESAcc.Fill(tblGuWuPKRoutes)
            tblGuWuPKRoutes.EndLoadData()
            strM = "....tblGuWuPKSubjects"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuPKSubjects.Clear()
            tblGuWuPKSubjects.AcceptChanges()
            tblGuWuPKSubjects.BeginLoadData()
            ta_TBLGUWUPKSUBJECTSAcc.ClearBeforeFill = True
            ta_TBLGUWUPKSUBJECTSAcc.Fill(tblGuWuPKSubjects)
            tblGuWuPKSubjects.EndLoadData()
            strM = "....tblGuWuRTTimePoints"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuRTTimePoints.Clear()
            tblGuWuRTTimePoints.AcceptChanges()
            tblGuWuRTTimePoints.BeginLoadData()
            ta_TBLGUWURTTIMEPOINTSAcc.ClearBeforeFill = True
            ta_TBLGUWURTTIMEPOINTSAcc.Fill(tblGuWuRTTimePoints)
            tblGuWuRTTimePoints.EndLoadData()
            strM = "....tblGuWuAssignedCmpd"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuAssignedCmpd.Clear()
            tblGuWuAssignedCmpd.AcceptChanges()
            tblGuWuAssignedCmpd.BeginLoadData()
            ta_TBLGUWUASSIGNEDCMPDAcc.ClearBeforeFill = True
            ta_TBLGUWUASSIGNEDCMPDAcc.Fill(tblGuWuAssignedCmpd)
            tblGuWuAssignedCmpd.EndLoadData()
            strM = "....tblGuWuAssignedCmpdLot"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuAssignedCmpdLot.Clear()
            tblGuWuAssignedCmpdLot.AcceptChanges()
            tblGuWuAssignedCmpdLot.BeginLoadData()
            ta_TBLGUWUASSIGNEDCMPDLOTAcc.ClearBeforeFill = True
            ta_TBLGUWUASSIGNEDCMPDLOTAcc.Fill(tblGuWuAssignedCmpdLot)
            tblGuWuAssignedCmpdLot.EndLoadData()
            strM = "....Done"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLGUWUSTUDYSCHEDULING.Clear()
            TBLGUWUSTUDYSCHEDULING.AcceptChanges()
            TBLGUWUSTUDYSCHEDULING.BeginLoadData()
            ta_TBLGUWUSTUDYSCHEDULINGAcc.ClearBeforeFill = True
            ta_TBLGUWUSTUDYSCHEDULINGAcc.Fill(TBLGUWUSTUDYSCHEDULING)
            TBLGUWUSTUDYSCHEDULING.EndLoadData()
            strM = "....TBLGUWUTPCONFIG"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLGUWUTPCONFIG.Clear()
            TBLGUWUTPCONFIG.AcceptChanges()
            TBLGUWUTPCONFIG.BeginLoadData()
            ta_TBLGUWUTPCONFIGAcc.ClearBeforeFill = True
            ta_TBLGUWUTPCONFIGAcc.Fill(TBLGUWUTPCONFIG)
            TBLGUWUTPCONFIG.EndLoadData()
            strM = "....TBLGUWUTPNAMESCONFIG"

            'TBLGUWUTPNAMESCONFIG
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLGUWUTPNAMESCONFIG.Clear()
            TBLGUWUTPNAMESCONFIG.AcceptChanges()
            TBLGUWUTPNAMESCONFIG.BeginLoadData()
            ta_TBLGUWUTPNAMESCONFIGAcc.ClearBeforeFill = True
            ta_TBLGUWUTPNAMESCONFIGAcc.Fill(TBLGUWUTPNAMESCONFIG)
            TBLGUWUTPNAMESCONFIG.EndLoadData()
            strM = "....Done"

            ''QRYGUWUCALENDAR
            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'QRYGUWUCALENDAR.CLEAR()
            'QRYGUWUCALENDAR.ACCEPTCHANGES()
            'QRYGUWUCALENDAR.BeginLoadData()
            'ta_QRYGUWUCALENDARAcc.ClearBeforeFill = True
            'ta_QRYGUWUCALENDARAcc.Fill(QRYGUWUCALENDAR)
            'QRYGUWUCALENDAR.EndLoadData()
            'strM = "....Done"

            Call AddCols_tblAss() 'add BOOLEXCLSAMPLECHK back

            frmh.pb1.Value = frmh.pb1.Maximum
            frmh.pb1.Refresh()
            'On Error GoTo 0

            If boolFormLoad Then
                frm.Controls("lblErr").Text = ""
                frm.Controls("lblErr").Refresh()
                ''SendKeys.Send("%")
                'frmE.Dispose()
            Else
                'frmh.lblProgress.Visible = False
                frmh.lblProgress.Text = ""
                frmh.Refresh()
            End If

        Catch ex As Exception

            str1 = "Hmmm." & Chr(10) & "There seems to be a problem retrieving data from the StudyDoc datatabase."
            str1 = str1 & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
            str1 = str1 & Chr(10) & strM & "...."
            str1 = str1 & Chr(10) & ex.Message
            str1 = str1 & Chr(10) & "DAsRefreshAcc"
            str2 = "Critical communication error..."
            If boolFormLoad Then
                frm.Controls("lblErr").Text = str1
                frm.Controls("lblErr").Refresh()
                'frmE.Visible = False
                'frmE.cmdOK.Visible = True
                'frmE.pb1.Visible = False
                'frmE.TimerE.Enabled = False
                'frmE.ShowDialog()
                'SendKeys.Send("%")
            Else
                Call PositionProgress()
                frmh.lblProgress.Visible = True
                frmh.lblProgress.Text = str1
                frmH.lblProgress.Refresh()

                frmH.panProgress.Visible = True
                frmH.panProgress.Refresh()

            End If
            'Dim frmE As New frmErrorMsg
            MsgBox(str1, MsgBoxStyle.Critical, str2)

            Dim dt As Date
            Dim dt1 As Date
            dt = Now
            dt1 = DateAdd(DateInterval.Second, 1, dt)
            Do Until dt > dt1
                dt = Now
            Loop

            DAsRefreshAcc = False
            Exit Function

            'End

            'On Error GoTo 0
        End Try

    End Function

    Function DAsRefreshSQLServer(ByVal frm As Form) As Boolean

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String


        Call RemoveBOOLEXCLSAMPLECHK()

        DAsRefreshSQLServer = True

        'Dim frmE As New frmErrorMsg
        'Dim frmE As frmSplash1
        Dim strM As String

        If boolFormLoad Then
            str2 = "...Establishing communication with the LABIntegrity" & ChrW(8482) & " StudyDoc" & ChrW(8482) & " Microsoft" & ChrW(8482) & " SQLServer database..."
            frm.Controls("lblErr").Text = str2
            frm.Controls("lblErr").Refresh()
        Else
            Call PositionProgress()
            str2 = "...Refreshing StudyDoc database tables..."
            frmh.lblProgress.Text = str2
            frmh.lblProgress.Visible = True
            frmH.lblProgress.Refresh()

            frmH.panProgress.Visible = True
            frmH.panProgress.Refresh()

        End If
        strM = ""
        'On Error GoTo end1

        'do StudyDoc study-specific queries
        'do this now because columns are added later
        Call DAsRefreshSpecific()

        'use beginloaddata and endloaddata to speed up table filling

        strM = "....tblData"

        Dim ct As Short
        Dim ctMax As Short

        ct = 0
        ctMax = 100

        frmh.pb1.Value = 0
        frmH.pb1.Maximum = ctMax

        Try
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()

            tblData.Clear()
            tblData.AcceptChanges()
            tblData.BeginLoadData()
            ta_tblDataSQLServer.ClearBeforeFill = True
            ta_tblDataSQLServer.Fill(tblData)
            tblData.EndLoadData()
            strM = "....tblSampleReceipt"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblSampleReceipt.Clear()
            tblSampleReceipt.AcceptChanges()
            tblSampleReceipt.BeginLoadData()
            ta_tblSampleReceiptSQLServer.ClearBeforeFill = True
            ta_tblSampleReceiptSQLServer.Fill(tblSampleReceipt)
            tblSampleReceipt.EndLoadData()
            strM = "....tblTab1"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTab1.Clear()
            tblTab1.AcceptChanges()
            tblTab1.BeginLoadData()
            ta_tblTab1SQLServer.ClearBeforeFill = True
            ta_tblTab1SQLServer.Fill(tblTab1)
            tblTab1.EndLoadData()
            strM = "....tblConfiguration"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfiguration.Clear()
            tblConfiguration.AcceptChanges()
            tblConfiguration.BeginLoadData()
            ta_tblConfigurationSQLServer.ClearBeforeFill = True
            ta_tblConfigurationSQLServer.Fill(tblConfiguration)
            tblConfiguration.EndLoadData()
            strM = "....tblOutstandingItems"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblOutstandingItems.Clear()
            tblOutstandingItems.AcceptChanges()
            tblOutstandingItems.BeginLoadData()
            ta_tblOutstandingItemsSQLServer.ClearBeforeFill = True
            ta_tblOutstandingItemsSQLServer.Fill(tblOutstandingItems)
            tblOutstandingItems.EndLoadData()
            strM = "....tblPermissions"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPermissions.Clear()
            tblPermissions.AcceptChanges()
            tblPermissions.BeginLoadData()
            ta_tblPermissionsSQLServer.ClearBeforeFill = True
            ta_tblPermissionsSQLServer.Fill(tblPermissions)
            tblPermissions.EndLoadData()
            strM = "....tblPersonnel"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPersonnel.Clear()
            tblPersonnel.AcceptChanges()
            tblPersonnel.BeginLoadData()
            ta_tblPersonnelSQLServer.ClearBeforeFill = True
            ta_tblPersonnelSQLServer.Fill(tblPersonnel)
            tblPersonnel.EndLoadData()
            strM = "....tblUserAccounts"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblUserAccounts.Clear()
            tblUserAccounts.AcceptChanges()
            tblUserAccounts.BeginLoadData()
            ta_tblUserAccountsSQLServer.ClearBeforeFill = True
            ta_tblUserAccountsSQLServer.Fill(tblUserAccounts)
            tblUserAccounts.EndLoadData()
            strM = "....tblAnalRefStandards"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAnalRefStandards.Clear()
            tblAnalRefStandards.AcceptChanges()
            tblAnalRefStandards.BeginLoadData()
            ta_tblAnalRefStandardsSQLServer.ClearBeforeFill = True
            ta_tblAnalRefStandardsSQLServer.Fill(tblAnalRefStandards)
            tblAnalRefStandards.EndLoadData()
            strM = "....tblAnalyticalRunSummary"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAnalyticalRunSummary.Clear()
            tblAnalyticalRunSummary.AcceptChanges()
            tblAnalyticalRunSummary.BeginLoadData()
            ta_tblAnalyticalRunSummarySQLServer.ClearBeforeFill = True
            ta_tblAnalyticalRunSummarySQLServer.Fill(tblAnalyticalRunSummary)
            tblAnalyticalRunSummary.EndLoadData()
            strM = "....tblConfigBodySections"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigBodySections.Clear()
            tblConfigBodySections.AcceptChanges()
            tblConfigBodySections.BeginLoadData()
            ta_tblConfigBodySectionsSQLServer.ClearBeforeFill = True
            ta_tblConfigBodySectionsSQLServer.Fill(tblConfigBodySections)
            tblConfigBodySections.EndLoadData()
            strM = "....tblConfigHeaderLookup"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigHeaderLookup.Clear()
            tblConfigHeaderLookup.AcceptChanges()
            tblConfigHeaderLookup.BeginLoadData()
            ta_tblConfigHeaderLookupSQLServer.ClearBeforeFill = True
            ta_tblConfigHeaderLookupSQLServer.Fill(tblConfigHeaderLookup)
            tblConfigHeaderLookup.EndLoadData()
            strM = "....tblConfigReportType"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigReportType.Clear()
            tblConfigReportType.AcceptChanges()
            tblConfigReportType.BeginLoadData()
            ta_tblConfigReportTypeSQLServer.ClearBeforeFill = True
            ta_tblConfigReportTypeSQLServer.Fill(tblConfigReportType)
            tblConfigReportType.EndLoadData()
            strM = "....tblContributingPersonnel"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblContributingPersonnel.Clear()
            tblContributingPersonnel.AcceptChanges()
            tblContributingPersonnel.BeginLoadData()
            ta_tblContributingPersonnelSQLServer.ClearBeforeFill = True
            ta_tblContributingPersonnelSQLServer.Fill(tblContributingPersonnel)
            tblContributingPersonnel.EndLoadData()
            strM = "....tblCorporateAddresses"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblCorporateAddresses.Clear()
            tblCorporateAddresses.AcceptChanges()
            tblCorporateAddresses.BeginLoadData()
            ta_tblCorporateAddressesSQLServer.ClearBeforeFill = True
            ta_tblCorporateAddressesSQLServer.Fill(tblCorporateAddresses)
            tblCorporateAddresses.EndLoadData()
            strM = "....tblDataTableRowTitles"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDataTableRowTitles.Clear()
            tblDataTableRowTitles.AcceptChanges()
            tblDataTableRowTitles.BeginLoadData()
            ta_tblDataTableRowTitlesSQLServer.ClearBeforeFill = True
            ta_tblDataTableRowTitlesSQLServer.Fill(tblDataTableRowTitles)
            tblDataTableRowTitles.EndLoadData()
            strM = "....tblMaxID"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblMaxID.Clear()
            tblMaxID.AcceptChanges()
            tblMaxID.BeginLoadData()
            ta_tblMaxIDSQLServer.ClearBeforeFill = True
            ta_tblMaxIDSQLServer.Fill(tblMaxID)
            tblMaxID.EndLoadData()
            strM = "....tblMethodValidationData"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblMethodValidationData.Clear()
            tblMethodValidationData.AcceptChanges()
            tblMethodValidationData.BeginLoadData()
            ta_tblMethodValidationDataSQLServer.ClearBeforeFill = True
            ta_tblMethodValidationDataSQLServer.Fill(tblMethodValidationData)
            tblMethodValidationData.EndLoadData()
            strM = "....tblQATables"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblQATables.Clear()
            tblQATables.AcceptChanges()
            tblQATables.BeginLoadData()
            ta_tblQATablesSQLServer.ClearBeforeFill = True
            ta_tblQATablesSQLServer.Fill(tblQATables)
            tblQATables.EndLoadData()
            strM = "....tblReportHistory"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportHistory.Clear()
            tblReportHistory.AcceptChanges()
            tblReportHistory.BeginLoadData()
            ta_tblReportHistorySQLServer.ClearBeforeFill = True
            ta_tblReportHistorySQLServer.Fill(tblReportHistory)
            tblReportHistory.EndLoadData()
            strM = "....tblReports"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReports.Clear()
            tblReports.AcceptChanges()
            tblReports.BeginLoadData()
            ta_tblReportsSQLServer.ClearBeforeFill = True
            ta_tblReportsSQLServer.Fill(tblReports)
            tblReports.EndLoadData()
            strM = "....tblReportStatements"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportStatements.Clear()
            tblReportStatements.AcceptChanges()
            tblReportStatements.BeginLoadData()
            ta_tblReportStatementsSQLServer.ClearBeforeFill = True
            ta_tblReportStatementsSQLServer.Fill(tblReportStatements)
            tblReportStatements.EndLoadData()
            strM = "....tblReportTable"

            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblReportTable.CLEAR()
            'tblReportTable.ACCEPTCHANGES()
            'tblReportTable.BeginLoadData()
            'ta_tblReportTableSQLServer.ClearBeforeFill = True
            'ta_tblReportTableSQLServer.Fill(tblReportTable)
            'tblReportTable.EndLoadData()
            'strM = "....tblReportTableAnalytes"

            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblReportTableAnalytes.CLEAR()
            'tblReportTableAnalytes.ACCEPTCHANGES()
            'tblReportTableAnalytes.BeginLoadData()
            'ta_tblReportTableAnalytesSQLServer.ClearBeforeFill = True
            'ta_tblReportTableAnalytesSQLServer.Fill(tblReportTableAnalytes)
            'tblReportTableAnalytes.EndLoadData()
            'strM = "....tblStudies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            'do specific
            'tblReportTableHeaderConfig.CLEAR()
            'tblReportTableHeaderConfig.ACCEPTCHANGES()
            'tblReportTableHeaderConfig.BeginLoadData()
            'ta_tblReportTableHeaderConfigSQLServer.ClearBeforeFill = True
            'ta_tblReportTableHeaderConfigSQLServer.Fill(tblReportTableHeaderConfig)
            'tblReportTableHeaderConfig.EndLoadData()
            'strM = "....tblStudies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblStudies.Clear()
            tblStudies.AcceptChanges()
            tblStudies.BeginLoadData()
            ta_tblStudiesSQLServer.ClearBeforeFill = True
            ta_tblStudiesSQLServer.Fill(tblStudies)
            tblStudies.EndLoadData()
            strM = "....tblTemplates"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTemplates.Clear()
            tblTemplates.AcceptChanges()
            tblTemplates.BeginLoadData()
            ta_tblTemplatesSQLServer.ClearBeforeFill = True
            ta_tblTemplatesSQLServer.Fill(tblTemplates)
            tblTemplates.EndLoadData()
            strM = "....tblTemplateAttributes"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTemplateAttributes.Clear()
            tblTemplateAttributes.AcceptChanges()
            tblTemplateAttributes.BeginLoadData()
            ta_tblTemplateAttributesSQLServer.ClearBeforeFill = True
            ta_tblTemplateAttributesSQLServer.Fill(tblTemplateAttributes)
            tblTemplateAttributes.EndLoadData()
            strM = "....tblConfigReportTables"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigReportTables.Clear()
            tblConfigReportTables.AcceptChanges()
            tblConfigReportTables.BeginLoadData()
            ta_tblConfigReportTablesSQLServer.ClearBeforeFill = True
            ta_tblConfigReportTablesSQLServer.Fill(tblConfigReportTables)
            tblConfigReportTables.EndLoadData()
            strM = "....tblAddressLabels"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAddressLabels.Clear()
            tblAddressLabels.AcceptChanges()
            tblAddressLabels.BeginLoadData()
            ta_tblAddressLabelsSQLServer.ClearBeforeFill = True
            ta_tblAddressLabelsSQLServer.Fill(tblAddressLabels)
            tblAddressLabels.EndLoadData()
            strM = "....tblCorporateNickNames"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblCorporateNickNames.Clear()
            tblCorporateNickNames.AcceptChanges()
            tblCorporateNickNames.BeginLoadData()
            ta_tblCorporateNickNamesSQLServer.ClearBeforeFill = True
            ta_tblCorporateNickNamesSQLServer.Fill(tblCorporateNickNames)
            tblCorporateNickNames.EndLoadData()
            strM = "....tblDropdownBoxContent"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDropdownBoxContent.Clear()
            tblDropdownBoxContent.AcceptChanges()
            tblDropdownBoxContent.BeginLoadData()
            ta_tblDropdownBoxContentSQLServer.ClearBeforeFill = True
            ta_tblDropdownBoxContentSQLServer.Fill(tblDropdownBoxContent)
            tblDropdownBoxContent.EndLoadData()
            strM = "....tblDropdownBoxName"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDropdownBoxName.Clear()
            tblDropdownBoxName.AcceptChanges()
            tblDropdownBoxName.BeginLoadData()
            ta_tblDropdownBoxNameSQLServer.ClearBeforeFill = True
            ta_tblDropdownBoxNameSQLServer.Fill(tblDropdownBoxName)
            tblDropdownBoxName.EndLoadData()
            strM = "....tblPasswordHistory"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblPasswordHistory.Clear()
            tblPasswordHistory.AcceptChanges()
            tblPasswordHistory.BeginLoadData()
            ta_tblPasswordHistorySQLServer.ClearBeforeFill = True
            ta_tblPasswordHistorySQLServer.Fill(tblPasswordHistory)
            tblPasswordHistory.EndLoadData()
            strM = "....tblSummaryData"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblSummaryData.Clear()
            tblSummaryData.AcceptChanges()
            tblSummaryData.BeginLoadData()
            ta_tblSummaryDataSQLServer.ClearBeforeFill = True
            ta_tblSummaryDataSQLServer.Fill(tblSummaryData)
            tblSummaryData.EndLoadData()
            strM = "....tblHooks"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblHooks.Clear()
            tblHooks.AcceptChanges()
            tblHooks.BeginLoadData()
            ta_tblHooksSQLServer.ClearBeforeFill = True
            ta_tblHooksSQLServer.Fill(tblHooks)
            tblHooks.EndLoadData()
            strM = "....tblAssignedSamples"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            ''too big. Return specific in DAsRefreshSpecific
            'tblAssignedSamples.CLEAR()
            'tblAssignedSamples.ACCEPTCHANGES()
            'tblAssignedSamples.BeginLoadData()
            'ta_tblAssignedSamplesSQLServer.ClearBeforeFill = True
            'ta_tblAssignedSamplesSQLServer.Fill(tblAssignedSamples)
            'tblAssignedSamples.EndLoadData()
            strM = "....tblDateFormats"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDateFormats.Clear()
            tblDateFormats.AcceptChanges()
            tblDateFormats.BeginLoadData()
            ta_tblDateFormatsSQLServer.ClearBeforeFill = True
            ta_tblDateFormatsSQLServer.Fill(tblDateFormats)
            tblDateFormats.EndLoadData()
            strM = "....tblAssignedSamplesHelper"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblAssignedSamplesHelper.Clear()
            tblAssignedSamplesHelper.AcceptChanges()
            tblAssignedSamplesHelper.BeginLoadData()
            ta_tblAssignedSamplesHelperSQLServer.ClearBeforeFill = True
            ta_tblAssignedSamplesHelperSQLServer.Fill(tblAssignedSamplesHelper)
            tblAssignedSamplesHelper.EndLoadData()
            strM = "....tblAppFigs"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            Try
                tblAppFigs.Clear()
                tblAppFigs.AcceptChanges()
                tblAppFigs.BeginLoadData()

                ta_tblAppFigsSQLServer.Fill(tblAppFigs)
                tblAppFigs.EndLoadData()
                strM = "....tblConfigAppFigs"

            Catch ex As Exception

            End Try

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigAppFigs.Clear()
            tblConfigAppFigs.AcceptChanges()
            tblConfigAppFigs.BeginLoadData()
            ta_tblConfigAppFigsSQLServer.ClearBeforeFill = True
            ta_tblConfigAppFigsSQLServer.Fill(tblConfigAppFigs)
            tblConfigAppFigs.EndLoadData()
            strM = "....tblIncludedRows"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblIncludedRows.Clear()
            tblIncludedRows.AcceptChanges()
            tblIncludedRows.BeginLoadData()
            ta_tblIncludedRowsSQLServer.ClearBeforeFill = True
            ta_tblIncludedRowsSQLServer.Fill(tblIncludedRows)
            tblIncludedRows.EndLoadData()
            strM = "....tblTableLegends"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            'do specific
            'tblTableProperties.CLEAR()
            'tblTableProperties.ACCEPTCHANGES()
            'tblTableProperties.BeginLoadData()
            'ta_tblTablePropertiesSQLServer.ClearBeforeFill = True
            'ta_tblTablePropertiesSQLServer.Fill(tblTableProperties)
            'tblTableProperties.EndLoadData()
            'strM = "....tblTableLegends"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblTableLegends.Clear()
            tblTableLegends.AcceptChanges()
            tblTableLegends.BeginLoadData()
            ta_tblTableLegendsSQLServer.ClearBeforeFill = True
            ta_tblTableLegendsSQLServer.Fill(tblTableLegends)
            tblTableLegends.EndLoadData()
            strM = "....tblFieldCodes"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblFieldCodes.Clear()
            tblFieldCodes.AcceptChanges()
            tblFieldCodes.BeginLoadData()
            ta_tblFieldCodesSQLServer.ClearBeforeFill = True
            ta_tblFieldCodesSQLServer.Fill(tblFieldCodes)
            tblFieldCodes.EndLoadData()
            strM = "....tblReportHeaders"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReportHeaders.Clear()
            tblReportHeaders.AcceptChanges()
            tblReportHeaders.BeginLoadData()
            ta_tblReportHeadersSQLServer.ClearBeforeFill = True
            ta_tblReportHeadersSQLServer.Fill(tblReportHeaders)
            tblReportHeaders.EndLoadData()
            strM = "....tblWordStatements"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblWordStatements.Clear()
            tblWordStatements.AcceptChanges()
            tblWordStatements.BeginLoadData()
            ta_tblWordStatementsSQLServer.ClearBeforeFill = True
            ta_tblWordStatementsSQLServer.Fill(tblWordStatements)
            tblWordStatements.EndLoadData()
            strM = "....tblWordDocs"


            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'tblWorddocs.CLEAR()
            'tblWorddocs.ACCEPTCHANGES()
            'tblWorddocs.BeginLoadData()
            'ta_tblWorddocsSQLServer.ClearBeforeFill = True
            'ta_tblWorddocsSQLServer.Fill(tblWorddocs)
            'tblWorddocs.EndLoadData()
            'strM = "....tblAuditTrail"


            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            'tblAuditTrail.BeginLoadData()
            str1 = "SELECT * FROM TBLAUDITTRAIL WHERE ID_TBLAUDITTRAIL < 0"
            Dim con As New ADODB.Connection
            'con.Open("Provider=SQLOLEDB;" & constrIni)
            con.Open(constrIni)

            Dim rs1 As New ADODB.Recordset
            rs1.CursorLocation = CursorLocationEnum.adUseClient
            rs1.Open(str1, con, CursorTypeEnum.adOpenStatic, LockTypeEnum.adLockReadOnly)
            rs1.ActiveConnection = Nothing
            tblAuditTrail.Clear()
            tblAuditTrail.AcceptChanges()
            tblAuditTrail.BeginLoadData()
            daDoPr.Fill(tblAuditTrail, rs1)
            tblAuditTrail.EndLoadData()
            rs1.Close()
            rs1 = Nothing
            con.Close()
            con = Nothing




            'ta_tblAuditTrailSQLServer.Fill(tblAuditTrail)
            'tblAuditTrail.EndLoadData()
            strM = "....tblReasonForChange"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblReasonForChange.Clear()
            tblReasonForChange.AcceptChanges()
            tblReasonForChange.BeginLoadData()
            ta_tblReasonForChangeSQLServer.ClearBeforeFill = True
            ta_tblReasonForChangeSQLServer.Fill(tblReasonForChange)
            tblReasonForChange.EndLoadData()
            strM = "....tblMeaningOfSig"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblMeaningOfSig.Clear()
            tblMeaningOfSig.AcceptChanges()
            tblMeaningOfSig.BeginLoadData()
            ta_tblMeaningOfSigSQLServer.ClearBeforeFill = True
            ta_tblMeaningOfSigSQLServer.Fill(tblMeaningOfSig)
            tblMeaningOfSig.EndLoadData()
            strM = "....tblSaveEvent"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblSaveEvent.Clear()
            tblSaveEvent.AcceptChanges()
            tblSaveEvent.BeginLoadData()
            ta_tblSaveEventSQLServer.ClearBeforeFill = True
            ta_tblSaveEventSQLServer.Fill(tblSaveEvent)
            tblSaveEvent.EndLoadData()
            strM = "....tblDataSystem"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblDataSystem.Clear()
            tblDataSystem.AcceptChanges()
            tblDataSystem.BeginLoadData()
            ta_tblDataSystemSQLServer.ClearBeforeFill = True
            ta_tblDataSystemSQLServer.Fill(tblDataSystem)
            tblDataSystem.EndLoadData()
            strM = "....tblCustomFieldCodes"

            '02218
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblCustomFieldCodes.Clear()
            tblCustomFieldCodes.AcceptChanges()
            tblCustomFieldCodes.BeginLoadData()
            ta_tblCustomFieldCodesSQLServer.ClearBeforeFill = True
            ta_tblCustomFieldCodesSQLServer.Fill(tblCustomFieldCodes)
            tblCustomFieldCodes.EndLoadData()
            strM = "....tblConfigCompliance"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblConfigCompliance.Clear()
            tblConfigCompliance.AcceptChanges()
            tblConfigCompliance.BeginLoadData()
            ta_tblConfigComplianceSQLServer.ClearBeforeFill = True
            ta_tblConfigComplianceSQLServer.Fill(tblConfigCompliance)
            tblConfigCompliance.EndLoadData()
            strM = "....TBLWORDSTATEMENTSVERSIONS"

            '030008
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLWORDSTATEMENTSVERSIONS.Clear()
            TBLWORDSTATEMENTSVERSIONS.AcceptChanges()
            TBLWORDSTATEMENTSVERSIONS.BeginLoadData()
            ta_TBLWORDSTATEMENTSVERSIONSSQLServer.ClearBeforeFill = True
            ta_TBLWORDSTATEMENTSVERSIONSSQLServer.Fill(TBLWORDSTATEMENTSVERSIONS)
            TBLWORDSTATEMENTSVERSIONS.EndLoadData()
            strM = "....TBLSECTIONTEMPLATES"


            '03000901
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLSECTIONTEMPLATES.Clear()
            TBLSECTIONTEMPLATES.AcceptChanges()
            TBLSECTIONTEMPLATES.BeginLoadData()
            ta_TBLSECTIONTEMPLATESSQLServer.ClearBeforeFill = True
            ta_TBLSECTIONTEMPLATESSQLServer.Fill(TBLSECTIONTEMPLATES)
            TBLSECTIONTEMPLATES.EndLoadData()
            strM = "....TBLSTUDYDOCANALYTES"

            '03046602
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLSTUDYDOCANALYTES.Clear()
            TBLSTUDYDOCANALYTES.AcceptChanges()
            TBLSTUDYDOCANALYTES.BeginLoadData()
            ta_TBLSTUDYDOCANALYTESSQLSERVER.ClearBeforeFill = True
            ta_TBLSTUDYDOCANALYTESSQLSERVER.Fill(TBLSTUDYDOCANALYTES)
            TBLSTUDYDOCANALYTES.EndLoadData()
            strM = "....tblModules"


            '030030_01
            'Note: do not fill TBLFINALREPORT or TBLFINALREPORTWORDDOCS here
            'fill in WordDoc module when user saves report


            'start Study Design

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblModules.Clear()
            tblModules.AcceptChanges()
            tblModules.BeginLoadData()
            ta_tblModulesSQLServer.ClearBeforeFill = True
            ta_tblModulesSQLServer.Fill(tblModules)
            tblModules.EndLoadData()
            strM = "....tblVersion"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblVersion.Clear()
            tblVersion.AcceptChanges()
            tblVersion.BeginLoadData()
            ta_TBLVERSIONSQLServer.ClearBeforeFill = True
            ta_TBLVERSIONSQLServer.Fill(tblVersion)
            tblVersion.EndLoadData()
            strM = "....tblGuWuAnimalReceipt"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuAnimalReceipt.Clear()
            tblGuWuAnimalReceipt.AcceptChanges()
            tblGuWuAnimalReceipt.BeginLoadData()
            ta_TBLGUWUANIMALRECEIPTSQLServer.ClearBeforeFill = True
            ta_TBLGUWUANIMALRECEIPTSQLServer.Fill(tblGuWuAnimalReceipt)
            tblGuWuAnimalReceipt.EndLoadData()
            strM = "....tblGuWuCompounds"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuCompounds.Clear()
            tblGuWuCompounds.AcceptChanges()
            tblGuWuCompounds.BeginLoadData()
            ta_TBLGUWUCOMPOUNDSSQLServer.ClearBeforeFill = True
            ta_TBLGUWUCOMPOUNDSSQLServer.Fill(tblGuWuCompounds)
            tblGuWuCompounds.EndLoadData()
            strM = "....tblGuWuCompoundsInd"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuCompoundsInd.Clear()
            tblGuWuCompoundsInd.AcceptChanges()
            tblGuWuCompoundsInd.BeginLoadData()
            ta_TBLGUWUCOMPOUNDSINDSQLServer.ClearBeforeFill = True
            ta_TBLGUWUCOMPOUNDSINDSQLServer.Fill(tblGuWuCompoundsInd)
            tblGuWuCompoundsInd.EndLoadData()
            strM = "....tblGuWuCompoundType"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuCompoundType.Clear()
            tblGuWuCompoundType.AcceptChanges()
            tblGuWuCompoundType.BeginLoadData()
            ta_TBLGUWUCOMPOUNDTYPESQLServer.ClearBeforeFill = True
            ta_TBLGUWUCOMPOUNDTYPESQLServer.Fill(tblGuWuCompoundType)
            tblGuWuCompoundType.EndLoadData()
            strM = "....tblGuWuProjects"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuProjects.Clear()
            tblGuWuProjects.AcceptChanges()
            tblGuWuProjects.BeginLoadData()
            ta_TBLGUWUPROJECTSSQLServer.ClearBeforeFill = True
            ta_TBLGUWUPROJECTSSQLServer.Fill(tblGuWuProjects)
            tblGuWuProjects.EndLoadData()
            strM = "....tblGuWuSpecies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuSpecies.Clear()
            tblGuWuSpecies.AcceptChanges()
            tblGuWuSpecies.BeginLoadData()
            ta_TBLGUWUSPECIESSQLServer.ClearBeforeFill = True
            ta_TBLGUWUSPECIESSQLServer.Fill(tblGuWuSpecies)
            tblGuWuSpecies.EndLoadData()
            strM = "....tblGuWuStudies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuStudies.Clear()
            tblGuWuStudies.AcceptChanges()
            tblGuWuStudies.BeginLoadData()
            ta_TBLGUWUSTUDIESSQLServer.ClearBeforeFill = True
            ta_TBLGUWUSTUDIESSQLServer.Fill(tblGuWuStudies)
            tblGuWuStudies.EndLoadData()
            strM = "....tblGuWuStudyDesignType"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuStudyDesignType.Clear()
            tblGuWuStudyDesignType.AcceptChanges()
            tblGuWuStudyDesignType.BeginLoadData()
            ta_TBLGUWUSTUDYDESIGNTYPESQLServer.ClearBeforeFill = True
            ta_TBLGUWUSTUDYDESIGNTYPESQLServer.Fill(tblGuWuStudyDesignType)
            tblGuWuStudyDesignType.EndLoadData()
            strM = "....tblGuWuStudySpecies"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuStudySpecies.Clear()
            tblGuWuStudySpecies.AcceptChanges()
            tblGuWuStudySpecies.BeginLoadData()
            ta_TBLGUWUSTUDYSPECIESSQLServer.ClearBeforeFill = True
            ta_TBLGUWUSTUDYSPECIESSQLServer.Fill(tblGuWuStudySpecies)
            tblGuWuStudySpecies.EndLoadData()
            strM = "....tblGuWuStudyStat"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuStudyStat.Clear()
            tblGuWuStudyStat.AcceptChanges()
            tblGuWuStudyStat.BeginLoadData()
            ta_TBLGUWUSTUDYSTATSQLServer.ClearBeforeFill = True
            ta_TBLGUWUSTUDYSTATSQLServer.Fill(tblGuWuStudyStat)
            tblGuWuStudyStat.EndLoadData()
            strM = "....TBLGUWUASSAYPERS"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLGUWUASSAYPERS.Clear()
            TBLGUWUASSAYPERS.AcceptChanges()
            TBLGUWUASSAYPERS.BeginLoadData()
            ta_TBLGUWUASSAYPERSSQLServer.ClearBeforeFill = True
            ta_TBLGUWUASSAYPERSSQLServer.Fill(TBLGUWUASSAYPERS)
            TBLGUWUASSAYPERS.EndLoadData()
            strM = "....tblGuWuAssay"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuAssay.Clear()
            tblGuWuAssay.AcceptChanges()
            tblGuWuAssay.BeginLoadData()
            ta_TBLGUWUASSAYSQLServer.ClearBeforeFill = True
            ta_TBLGUWUASSAYSQLServer.Fill(tblGuWuAssay)
            tblGuWuAssay.EndLoadData()
            strM = "....TBLGUWUSPECIESSTRAIN"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuSpeciesStrain.Clear()
            tblGuWuSpeciesStrain.AcceptChanges()
            tblGuWuSpeciesStrain.BeginLoadData()
            ta_TBLGUWUSPECIESSTRAINSQLServer.ClearBeforeFill = True
            ta_TBLGUWUSPECIESSTRAINSQLServer.Fill(tblGuWuSpeciesStrain)
            tblGuWuSpeciesStrain.EndLoadData()
            strM = "....tblGuWuDoseUnits"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuDoseUnits.Clear()
            tblGuWuDoseUnits.AcceptChanges()
            tblGuWuDoseUnits.BeginLoadData()
            ta_TBLGUWUDOSEUNITSSQLServer.ClearBeforeFill = True
            ta_TBLGUWUDOSEUNITSSQLServer.Fill(tblGuWuDoseUnits)
            tblGuWuDoseUnits.EndLoadData()
            strM = "....tblGuWuPKGroups"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuPKGroups.Clear()
            tblGuWuPKGroups.AcceptChanges()
            tblGuWuPKGroups.BeginLoadData()
            ta_TBLGUWUPKGROUPSSQLServer.ClearBeforeFill = True
            ta_TBLGUWUPKGROUPSSQLServer.Fill(tblGuWuPKGroups)
            tblGuWuPKGroups.EndLoadData()
            strM = "....tblGuWuPKRoutes"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuPKRoutes.Clear()
            tblGuWuPKRoutes.AcceptChanges()
            tblGuWuPKRoutes.BeginLoadData()
            ta_TBLGUWUPKROUTESSQLServer.ClearBeforeFill = True
            ta_TBLGUWUPKROUTESSQLServer.Fill(tblGuWuPKRoutes)
            tblGuWuPKRoutes.EndLoadData()
            strM = "....tblGuWuPKSubjects"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuPKSubjects.Clear()
            tblGuWuPKSubjects.AcceptChanges()
            tblGuWuPKSubjects.BeginLoadData()
            ta_TBLGUWUPKSUBJECTSSQLServer.ClearBeforeFill = True
            ta_TBLGUWUPKSUBJECTSSQLServer.Fill(tblGuWuPKSubjects)
            tblGuWuPKSubjects.EndLoadData()
            strM = "....tblGuWuRTTimePoints"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuRTTimePoints.Clear()
            tblGuWuRTTimePoints.AcceptChanges()
            tblGuWuRTTimePoints.BeginLoadData()
            ta_TBLGUWURTTIMEPOINTSSQLServer.ClearBeforeFill = True
            ta_TBLGUWURTTIMEPOINTSSQLServer.Fill(tblGuWuRTTimePoints)
            tblGuWuRTTimePoints.EndLoadData()
            strM = "....tblGuWuAssignedCmpd"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuAssignedCmpd.Clear()
            tblGuWuAssignedCmpd.AcceptChanges()
            tblGuWuAssignedCmpd.BeginLoadData()
            ta_TBLGUWUASSIGNEDCMPDSQLServer.ClearBeforeFill = True
            ta_TBLGUWUASSIGNEDCMPDSQLServer.Fill(tblGuWuAssignedCmpd)
            tblGuWuAssignedCmpd.EndLoadData()
            strM = "....tblGuWuAssignedCmpdLot"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            tblGuWuAssignedCmpdLot.Clear()
            tblGuWuAssignedCmpdLot.AcceptChanges()
            tblGuWuAssignedCmpdLot.BeginLoadData()
            ta_TBLGUWUASSIGNEDCMPDLOTSQLServer.ClearBeforeFill = True
            ta_TBLGUWUASSIGNEDCMPDLOTSQLServer.Fill(tblGuWuAssignedCmpdLot)
            tblGuWuAssignedCmpdLot.EndLoadData()
            strM = "....Done"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLGUWUSTUDYSCHEDULING.Clear()
            TBLGUWUSTUDYSCHEDULING.AcceptChanges()
            TBLGUWUSTUDYSCHEDULING.BeginLoadData()
            ta_TBLGUWUSTUDYSCHEDULINGSQLServer.ClearBeforeFill = True
            ta_TBLGUWUSTUDYSCHEDULINGSQLServer.Fill(TBLGUWUSTUDYSCHEDULING)
            TBLGUWUSTUDYSCHEDULING.EndLoadData()
            strM = "....TBLGUWUTPCONFIG"

            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLGUWUTPCONFIG.Clear()
            TBLGUWUTPCONFIG.AcceptChanges()
            TBLGUWUTPCONFIG.BeginLoadData()
            ta_TBLGUWUTPCONFIGSQLServer.ClearBeforeFill = True
            ta_TBLGUWUTPCONFIGSQLServer.Fill(TBLGUWUTPCONFIG)
            TBLGUWUTPCONFIG.EndLoadData()
            strM = "....TBLGUWUTPNAMESCONFIG"

            'TBLGUWUTPNAMESCONFIG
            ct = ct + 1
            frmH.pb1.Value = ct
            frmH.pb1.Refresh()
            TBLGUWUTPNAMESCONFIG.Clear()
            TBLGUWUTPNAMESCONFIG.AcceptChanges()
            TBLGUWUTPNAMESCONFIG.BeginLoadData()
            ta_TBLGUWUTPNAMESCONFIGSQLServer.ClearBeforeFill = True
            ta_TBLGUWUTPNAMESCONFIGSQLServer.Fill(TBLGUWUTPNAMESCONFIG)
            TBLGUWUTPNAMESCONFIG.EndLoadData()
            strM = "....Done"

            Call AddCols_tblAss() 'add BOOLEXCLSAMPLECHK back

            ''QRYGUWUCALENDAR
            'ct = ct + 1
            'frmh.pb1.Value = ct
            'frmh.pb1.Refresh()
            'QRYGUWUCALENDAR.CLEAR()
            'QRYGUWUCALENDAR.ACCEPTCHANGES()
            'QRYGUWUCALENDAR.BeginLoadData()
            'ta_QRYGUWUCALENDARSQLServer.ClearBeforeFill = True
            'ta_QRYGUWUCALENDARSQLServer.Fill(QRYGUWUCALENDAR)
            'QRYGUWUCALENDAR.EndLoadData()
            'strM = "....Done"



            frmH.pb1.Value = frmH.pb1.Maximum
            frmH.pb1.Refresh()
            'On Error GoTo 0

            If boolFormLoad Then
                frm.Controls("lblErr").Text = ""
                frm.Controls("lblErr").Refresh()
                ''SendKeys.Send("%")
                'frmE.Dispose()
            Else
                'frmh.lblProgress.Visible = False
                frmH.lblProgress.Text = ""
                frmH.Refresh()
            End If

        Catch ex As Exception

            str1 = "Hmmm." & Chr(10) & "There seems to be a problem retrieving data from the StudyDoc datatabase."
            str1 = str1 & Chr(10) & Chr(10) & "Please contact your StudyDoc system administrator."
            str1 = str1 & Chr(10) & strM & "...."
            str1 = str1 & Chr(10) & ex.Message
            str1 = str1 & Chr(10) & "DAsRefreshSQLServer"
            str2 = "Critical communication error..."
            If boolFormLoad Then
                frm.Controls("lblErr").Text = str1
                frm.Controls("lblErr").Refresh()
                'frmE.Visible = False
                'frmE.cmdOK.Visible = True
                'frmE.pb1.Visible = False
                'frmE.TimerE.Enabled = False
                'frmE.ShowDialog()
                'SendKeys.Send("%")
            Else
                Call PositionProgress()
                frmH.lblProgress.Visible = True
                frmH.lblProgress.Text = str1
                frmH.lblProgress.Refresh()

                frmH.panProgress.Visible = True
                frmH.panProgress.Refresh()

            End If
            'Dim frmE As New frmErrorMsg
            MsgBox(str1, MsgBoxStyle.Critical, str2)

            Dim dt As Date
            Dim dt1 As Date
            dt = Now
            dt1 = DateAdd(DateInterval.Second, 1, dt)
            Do Until dt > dt1
                dt = Now
            Loop

            DAsRefreshSQLServer = False
            Exit Function

            'End

            'On Error GoTo 0
        End Try

    End Function

End Module
