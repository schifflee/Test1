Option Compare Text

Public Class frmValidateReport

    Public dtCutOff As Date

    Private Sub frmValidateReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Call ControlDefaults(Me)

    End Sub

    Sub PositionForm()

        Dim w, h, w1, h1

        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        Me.Top = h * 0.05
        Me.Left = w * 0.05
        Me.Height = h * 0.9
        Me.Width = w * 0.9

    End Sub

    Sub FormLoad()

        Call PositionForm()

        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim strS As String
        Dim strM As String
        Dim Count1 As Short
        Dim Count2 As Int64
        Dim str1 As String
        Dim str2 As String
        Dim intO As Short
        Dim var1, var2
        Dim int1 As Short

        'get analytes
        For Count1 = 0 To tblAnalytesHome.Rows.Count - 1
            var1 = NZ(tblAnalytesHome.Rows(Count1).Item("ANALYTEDESCRIPTION"), "")
            var2 = NZ(tblAnalytesHome.Rows(Count1).Item("ISINTSTD"), "Yes")
            If Len(var1) = 0 Then
            Else
                If StrComp(var2, "No", CompareMethod.Text) = 0 Then
                    int1 = int1 + 1
                    If int1 = 1 Then
                        strF1 = "(CHARANALYTE = '" & CleanText(CStr(var1)) & "')"
                    Else
                        strF1 = strF1 & " OR (CHARANALYTE = '" & CleanText(CStr(var1)) & "')"
                    End If
                End If
            End If
        Next

        'strF = "RECORDTIMESTAMP > #" & dtCutOff & "#"
        strF = "RECORDTIMESTAMP > " & ReturnDate(dtCutOff)
        strF2 = strF & " AND (" & strF1 & ")"
        'str4 = "ORDER BY ANALYTICALRUNSAMPLE.RUNID, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER;"
        strS = "CHARANALYTE ASC, RUNID ASC, RUNSAMPLEORDERNUMBER ASC"

        'NOTE: tblAnalysisResultsHome has duplicates of you don't filter for analyte
        'must do some screwing around

        Dim dv As DataView = New DataView(tblAnalysisResultsHome, strF2, strS, DataViewRowState.CurrentRows)

        'SELECT DISTINCT ANALYTICALRUNSAMPLE.DESIGNSAMPLEID, ASSAYANALYTES.ANALYTEID, ANALYTICALRUNSAMPLE.STUDYID, STUDY.STUDYNAME, ANALYTICALRUN.RUNTYPEID, ANALYTICALRUNSAMPLE.RUNID, " & strAnaRunPeak & ".RUNSAMPLESEQUENCENUMBER, ANALYTICALRUNSAMPLE.RUNSAMPLEORDERNUMBER, ANALYTICALRUNSAMPLE.ASSAYDATETIME, ANALYTICALRUNSAMPLE.ASSAYLEVEL, ANARUNANALYTERESULTS.ELIMINATEDFLAG, " & strAnaRunPeak & ".SAMPLENAME, ANALYTICALRUNSAMPLE.ALIQUOTFACTOR, ANALYTICALRUNSAMPLE.RUNSAMPLEKIND, ANALYTICALRUN.ASSAYID, ASSAY.MASTERASSAYID, ANARUNANALYTERESULTS.ANALYTEINDEX, ANARUNANALYTERESULTS.CONCENTRATION, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, " & strAnaRunPeak & ".ANALYTEHEIGHT, " & strAnaRunPeak & ".ANALYTEAREA, " & strAnaRunPeak & ".INTERNALSTANDARDHEIGHT, " & strAnaRunPeak & ".INTERNALSTANDARDAREA, " & strAnaRunPeak & ".INTERNALSTDNAME, CONFIGSAMPLETYPES.SAMPLETYPEKEY, CONFIGSAMPLETYPES.SAMPLETYPEID, " & strAnaRunPeak & ".ANALYTEPEAKRETENTIONTIME, " & strAnaRunPeak & ".INTERNALSTANDARDRETENTIONTIME, ANALYTICALRUNANALYTES.RUNANALYTEREGRESSIONSTATUS, [ANARUNANALYTERESULTS.CONCENTRATION]/[ANALYTICALRUNSAMPLE.ALIQUOTFACTOR] AS REPORTEDCONC, ANARUNANALYTERESULTS.RECORDTIMESTAMP, SECUSERACCOUNTS.LOGINNAME, SECUSERACCOUNTS.FIRSTNAME, SECUSERACCOUNTS.MIDDLEINITIAL, SECUSERACCOUNTS.LASTNAME

        'NOTE: CHARANALYTE is added later

        Dim dgv As DataGridView = Me.dgvReport

        dv.AllowDelete = False
        dv.AllowEdit = False
        dv.AllowNew = False

        Dim intRR As Int64
        intRR = dv.Count

        dgv.DataSource = dv

        'first hide all columns
        For Count1 = 0 To dgv.ColumnCount - 1
            dgv.Columns(Count1).Visible = False
        Next

        'now make specific columns visible
        Dim boolFormat As Boolean = False
        Dim strFormat As String = ""
        Dim strAlign As String = "L"
        intO = -1
        For Count1 = 1 To 12
            boolFormat = False
            strAlign = "L"
            Select Case Count1
                Case 1
                    str1 = "CHARANALYTE"
                    str2 = "Analyte"
                Case 2
                    str1 = "RUNID"
                    str2 = "Run ID"
                    strAlign = "C"
                Case 3
                    str1 = "RUNSAMPLEORDERNUMBER"
                    str2 = "Seq Order #"
                    strAlign = "C"
                Case 4
                    str1 = "SAMPLENAME"
                    str2 = "Sample Name"
                Case 5
                    str1 = "RUNSAMPLEKIND"
                    str2 = "Sample Type"
                    strAlign = "C"
                Case 6
                    str1 = "ANALYTEAREA"
                    str2 = "Analyte Peak Area"
                    boolFormat = True
                    strFormat = "0"
                    strAlign = "R"
                Case 7
                    str1 = "REPORTEDCONC"
                    str2 = "Conc."
                    boolFormat = True
                    strFormat = "0.000"
                    strAlign = "R"
                Case 8
                    str1 = "RECORDTIMESTAMP"
                    str2 = "Time Stamp"
                    boolFormat = True
                    strFormat = LTextDateFormat & " HH:mm:ss tt"
                Case 9
                    str1 = "LOGINNAME"
                    str2 = "User ID"
                Case 10
                    str1 = "FIRSTNAME"
                    str2 = "First Name"
                Case 11
                    str1 = "MIDDLEINITIAL"
                    str2 = "Middle Init"
                Case 12
                    str1 = "LASTNAME"
                    str2 = "Last Name"

            End Select

            If dgv.Columns.Contains(str1) Then
                intO = intO + 1
                dgv.Columns(str1).Visible = True
                dgv.Columns(str1).HeaderText = str2
                dgv.Columns(str1).DisplayIndex = intO

                If boolFormat Then
                    dgv.Columns(str1).DefaultCellStyle.Format = strFormat
                End If

                Select Case strAlign
                    Case "R"
                        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
                    Case "C"
                        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
                    Case "L"
                        dgv.Columns(str1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft
                End Select
            End If

        Next

        dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        dgv.AutoResizeColumns()

        'problem here
        LTextDateFormat = Replace(LTextDateFormat, "YYYY", "yyyy", 1, -1, CompareMethod.Binary)


        If gboolER Then
            strM = "Last Report Saved Date:  " & Format(dtCutOff, LTextDateFormat & " HH:mm:ss tt") & ChrW(10)
        Else
            strM = "Last Report Generation Date:  " & Format(dtCutOff, LTextDateFormat & " HH:mm:ss tt") & ChrW(10)
        End If
        If dgv.RowCount = 0 Then
            If gboolER Then
                strM = strM & "The study has no records created or modified after the last final report saved date."
            Else
                strM = strM & "The study has no records created or modified after the last final report generation date."
            End If
        Else
            If gboolER Then
                strM = strM & "The records listed below were created or modified after the last Final Report saved date." & ChrW(10)
                strM = strM & "Select the table, then press control-c on the keyboard to copy to clipboard."
            Else
                strM = strM & "The records listed below were created or modified after the last Final Report generation date." & ChrW(10)
                strM = strM & "Select the table, then press control-c on the keyboard to copy to clipboard."
            End If
            
        End If

        Me.lblTitle.Text = strM

    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click

        Me.Visible = False

    End Sub
End Class