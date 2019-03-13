Option Compare Text

Imports System
Imports System.IO
Imports System.Text
Imports System.Data
Imports System.Linq.Expressions
Imports System.Linq
Imports Microsoft.Office.Interop.Word


Module modReportPrep

    Public strReportTemplateChoice As String = ""
    Public boolHasMacro As Boolean = False

    Sub ExampleReportBody(ByVal boolFieldCodesOnly As Boolean)

        boolHasMacro = False

        Dim var1, var2, var3
        Dim int1 As Short
        Dim int2 As Short
        Dim dv as system.data.dataview
        Dim drow As DataRow
        Dim tbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strSQL As String
        Dim dbPath As String
        Dim BACStudy As String
        Dim strTemplate As String
        'Dim rng1 As Range
        Dim Count1 As Short
        Dim Count2 As Short
        'Dim fi
        Dim ctCols As Short
        'Dim frm As New frmHome_01
        'Dim ''frmp As New ''frmprogress_01
        Dim intPMax As Short
        Dim intPCt As Short

        'check for permission
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short
        Dim drows() As DataRow

        Dim boolTempH As Boolean


        LTableDateTimeStamp = Now

        'first check for configured reports
        Dim dvR as system.data.dataview
        dvR = frmH.dgvReports.DataSource
        If dvR.Count = 0 Then
            str1 = "A report must be configured on the Home Tab in order to use this function."
            MsgBox(str1, MsgBoxStyle.Information, "Please configure a report...")
            Exit Sub
        End If

        If frmH.dgvReports.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = frmH.dgvReports.CurrentRow.Index
        End If

        tblAppendix.Clear()
        tblFigures.Clear()
        tblAttachments.Clear()
        ctAppendix = 0
        ctFigures = 0
        ctAttachments = 0

        boolShowExample = True
        ctPB = 0 'pesky
        intILS = 0
        intStartTable = 2

        If id_tblStudies = 0 Then
            MsgBox("A study must be chosen.", MsgBoxStyle.Information, "A study must be chosen...")
            GoTo end1
        End If

        Cursor.Current = Cursors.WaitCursor

        Dim frm As New frmMsgBox
        frm.lblText.Text = "Do you wish to generate an Example Report Body Section?"
        frm.lblReportTemplate.Visible = True
        frm.cbxReportTemplate.Visible = True
        frm.boolTable = False

        frm.ShowDialog()
        If frm.boolCancel Then
            int1 = 7
        Else
            int1 = 6
        End If
        strReportTemplateChoice = frm.cbxReportTemplate.Text
        frmH.Refresh()
        frm.Close()

        'int1 = MsgBox("Do you wish to generate an Example Report Body Section?", MsgBoxStyle.YesNo, "Generate Example Report Body Section...")

        If int1 = 6 Then
        Else
            GoTo end2
        End If

        'must make Report Body 'show included'
        If frmH.rbShowIncludedRBody.Checked Then
        Else
            frmH.rbShowIncludedRBody.Checked = True
        End If
        Call PositionProgress()
        frmH.lblProgress.Text = "Generating Example Report Body Section..."
        frmH.pb1.Value = 0
        frmH.lblProgress.Visible = True
        frmH.pb1.Visible = True
        frmH.lblProgress.Refresh()
        frmH.pb1.Refresh()

        'frmH.panProgress.Visible = True
        'frmH.panProgress.Refresh()

        'set arrReportNA
        arrReportNA.Clear(arrReportNA, 0, arrReportNA.Length)
        ctArrReportNA = 0
        intILS = 0
        intStartTable = 2

        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim boolSTB As Boolean = True
        Try
            boolSTB = wd.Application.ShowWindowsInTaskbar
        Catch ex As Exception
            var1 = var1 'debug
        End Try

        tbl = tblConfiguration

        Call PositionProgress()
        frmH.lblProgress.Text = ""
        frmH.lblProgress.Visible = True

        'frmH.panProgress.Visible = True
        'frmH.panProgress.Refresh()

        intPCt = intPCt + 1
        intPMax = 10

        ReDim garrMargins(1, 4) 'T,L,B,R for Portrait:  1=P, 2=L

        Dim intCBS As Int64

        intTCur = 0
        intTTot = 1

        Try

            strPathWd = GetNewTempFile(True)
            intCBS = frmH.dgvReportStatements("ID_TBLWORDSTATEMENTS", 0).Value
            'find new intcbs
            Dim intCBSa As Int64
            intCBSa = GetNewCBS()
            If intCBSa = 0 Then
            Else
                intCBS = intCBSa
            End If

            strTemplate = OpenTemplate(intCBS, strPathWd)
            wd.Documents.Open(FileName:=strTemplate, ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", Format:=Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAuto, XMLTransform:="")



            Dim ver
            ver = wd.Version
            Dim bool2007 As Boolean
            bool2007 = True
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If
            Dim strExt As String = ".docx"

            boolHasMacro = False
            If bool2007 Then
                'Note: Word2007 has no compatability mode parameter
                If wd.ActiveDocument.HasVBProject Then
                    strExt = ".docm"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                End If
            Else
                If wd.ActiveDocument.HasVBProject Then
                    strExt = ".docm"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs2(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs2(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                End If
            End If

            If boolHasMacro Then
                boolSaveAsDocx = SaveAsDocx(wd)
            Else
                boolSaveAsDocx = True
            End If

            '20180701 LEE:
            Call SetDocGlobal(wd.ActiveDocument)
            '20180701 LEE
            'Implement time saving trick
            Call SpellingOff(wd.ActiveDocument, False)

            'wdd.visible = True
            Call SetNormal(wd)

            If boolVerbose Then
                wd.Visible = True
                wd.WindowState = Office.Interop.Word.WdWindowState.wdWindowStateMaximize
            End If

            ftNormal = wd.Selection.Font.Name

            'wdd.visible = True

            'record normal font
            NormalFontsize = wd.Selection.Font.Size

            'add BlueHyperlink style
            If boolBLUEHYPERLINK Then
                Call CreateBlueHyperlink(wd)
            End If

        Catch ex As Exception
            wd.Application.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            wd = Nothing
            'str1 = "Hmmm. The document " & Chr(10) & Chr(10) & strTemplate & Chr(10) & Chr(10) & "could not be found."
            'str1 = str1 & str2

            str1 = "There was a problem preparing the Example Report." & Chr(10) & Chr(10) & ex.Message & Chr(10) & Chr(10) & str2

            MsgBox(str1, MsgBoxStyle.Information, "File not found...")
            boolReportCont = False
            GoTo end2
        End Try

        boolTOC = False
        boolTOA = False
        boolTOT = False
        boolTOF = False

        '***Here!!
        'wdd.visible = True

        'get ReportStatements.doc path
        Dim wdSt 'As new Microsoft.Office.Interop.Word.application
        Dim strPath As String
        Dim strPathGuWu As String
        'str1 = "charConfigTitle = 'Report Statements'"
        'Erase drows
        'drows = tbl.Select(str1)
        'strPath = drows(0).Item("charConfigValue")
        Erase drows
        'str1 = "charConfigTitle = 'GuWu Statements'"
        'drows = tbl.Select(str1)
        'strPathGuWu = drows(0).Item("charConfigValue")

        boolTempH = boolUseHyperlinks
        boolUseHyperlinks = False

        strPathWd = ""
        strPathWd = GetNewTempFile(True) 'DON'T SKIP THIS!!! Paste routine uses this

        'wdd.visible = True
        Try

            frmH.lblProgress.Text = "Preparing Report Body..."
            frmH.lblProgress.Refresh()

            Call ReportBody(wd, wdSt, True, boolFieldCodesOnly)

            frmH.lblProgress.Text = "Executing Repeat Section actions..."
            frmH.lblProgress.Refresh()

            '20190222 LEE: Does not need to be run in Example Report Body
            'Call SearchRepeat(wd)

            frmH.lblProgress.Text = "Executing Field Code actions..."
            frmH.lblProgress.Refresh()

            '20190222 LEE: Does not need to be run in Example Report Body
            'Call ReturnAnovaValues(wd)

        Catch ex As Exception
            str1 = "There seems to have been a problem creating the body portion of this report." & ChrW(10) & ChrW(10)
            str1 = str1 & "Try generating the report again." & ChrW(10) & ChrW(10)
            str1 = str1 & ex.Message & ChrW(10) & ChrW(10)
            str1 = str1 & "If the problem persists, please contract your StudyDoc Administrator."
            MsgBox(str1, MsgBoxStyle.Information, "Error in report body...")
            GoTo end1
        End Try

        Dim intSR As Short
        intSR = 1
        Try 'do this search first because table field codes have other field codes embedded in the table names
            Call RunIDSearch(wd)
        Catch ex As Exception
            str1 = "There was a problem completing the RunID Search action: " & intSR & ". Report generation will continue."
            str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
        End Try

        Try
            'wdd.visible = False
            Call SignatureSearch(wd)
        Catch ex As Exception
            str1 = "There was a problem completing the Signature Search/Replace action. Report generation will continue."
            str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
        End Try


        'do searchreplace
        Call DoSearchReplace(wd, True, boolFieldCodesOnly, False, True, False, False)

        boolUseHyperlinks = boolTempH

        Cursor.Current = Cursors.WaitCursor

        ''do headers
        'frmH.lblProgress.Text = "Inserting Headers..."
        'frmH.lblProgress.Refresh()

        Try
            Call EnterHeaders(wd)
        Catch ex As Exception
            str1 = "There was a problem creating headers for this report." & ChrW(10) & ChrW(10)
            str1 = str1 & "Try re-generating the report. If you still encounter this error, please contact your StudyDoc Administrator." & ChrW(10) & ChrW(10)
            MsgBox(str1, MsgBoxStyle.Information, "Report Header error...")
        End Try

        Try
            Call EnterFooters(wd)
        Catch ex As Exception
            str1 = "There was a problem creating footers for this report." & ChrW(10) & ChrW(10)
            str1 = str1 & "Try re-generating the report. If you still encounter this error, please contact your StudyDoc Administrator." & ChrW(10) & ChrW(10)
            MsgBox(str1, MsgBoxStyle.Information, "Report Header error...")
        End Try

        'wdd.visible = True
        'frmH.Activate()

        '20190221 LEE:
        'do searchreplace in footnotes
        'https://www.extendoffice.com/documents/word/5428-word-select-all-footnotes.html
        Dim xDoc As Microsoft.Office.Interop.Word.Document
        Dim xRange As Microsoft.Office.Interop.Word.Range
        xDoc = wd.ActiveDocument
        If xDoc.Footnotes.Count > 0 Then
            xRange = xDoc.Footnotes(1).Range
            xRange.WholeStory()
            xRange.Select()

            Try
                Call SearchReplace(wd, "Report Body", xRange, False, "", intSR, 0, 0, False, True, True, 1)
            Catch ex As Exception
                str1 = "There was a problem completing the Search/Replace action: " & intSR & ". Report generation will continue."
                str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
                MsgBox(str1)
            End Try

        End If


        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        'clear formatting again
        wd.Selection.Find.ClearFormatting()

        'update index table page numbers
        frmH.lblProgress.Text = "Updating Tables of Figures..." & ChrW(10) & "...this may take several minutes..."
        frmH.lblProgress.Refresh()
        Try
            Dim tofLoop As Microsoft.Office.Interop.Word.TableOfFigures
            For Each tofLoop In wd.ActiveDocument.TablesOfFigures
                tofLoop.Update()
                tofLoop.UpdatePageNumbers()
            Next tofLoop

        Catch ex As Exception

        End Try

        frmH.lblProgress.Text = "Updating Table of Contents..."
        frmH.lblProgress.Refresh()
        Try
            Dim tocLoop As Microsoft.Office.Interop.Word.tableOfContents
            For Each tocLoop In wd.ActiveDocument.TablesOfContents
                tocLoop.Update()
                tocLoop.UpdatePageNumbers()
            Next tocLoop
        Catch ex As Exception

        End Try

        'format toc color to blue
        Try
            Call FormatTOCColor(wd)
        Catch ex As Exception

        End Try

        Call ModifyTOC(wd)

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        Try
            wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception

        End Try

        If ctArrReportNA > 0 Then
            'Call UpdateOutstandingItems()
        End If

        ''Watermarks now get performed in ReportBody
        'frmH.lblProgress.Text = "Inserting Watermarks..."
        'frmH.lblProgress.Refresh()

        'Try
        '    Call InsertWatermark(wd)
        'Catch ex As Exception
        '    str1 = "Unfortunately, this version of Microsoft" & ChrW(10) & " Word does not contain the Word watermarking funtion supported by GuWu." & ChrW(10) & ChrW(10)
        '    str1 = str1 & "Word 2002 or higher must be used. " & ChrW(10) & ChrW(10)
        '    str1 = str1 & "The report will be prepared without a watermark."
        '    MsgBox(str1, MsgBoxStyle.Information, "Watermark not supported...")
        'End Try



        'str1 = "Example Report Body Section Completed."
        'MsgBox(str1, MsgBoxStyle.Information, "Action completed...")

end1:

        'save as temp then display in afr
        Dim strP As String
        strP = GetNewTempFile(True)

        If boolHasMacro And boolSaveAsDocx = False Then
            strP = Replace(strP, ".xml", ".docm", 1, -1, CompareMethod.Text)
        Else
            strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)
        End If


        frmH.lblProgress.Text = "Saving document..." ':" & ChrW(10) & ChrW(10) & strP
        frmH.lblProgress.Refresh()

        If boolHasMacro And boolSaveAsDocx = False Then
            wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
        Else
            wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
        End If

        'wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)

        If gDoPDF Then
            If wd.Version < 12 Then
                gDoPDF = False
            Else
                gDoPDF = True
            End If
        End If

        If gDoPDF Then
            Call CreatePDF(wd, strP)
        End If

        'frmH.pb1.Visible = False
        'frmH.lblProgress.Visible = False
        frmH.lblProgress.Text = ""

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()


        str1 = "Example Report Body"
        Call ReportHistoryItem(str1)

        '20180701 LEE
        'Implement time saving trick
        Call SpellingOff(wd.ActiveDocument, True)


        Try
            wd.ActiveDocument.Close(False)
        Catch ex As Exception

        End Try

        Try
            wd.Application.ShowWindowsInTaskbar = boolSTB
        Catch ex As Exception

        End Try

        Try
            wd.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
        Catch ex As Exception

        End Try

        Threading.Thread.Sleep(250)

        'If gboolER Then
        '    frmH.lblProgress.Visible = False
        '    frmH.pb1.Visible = False
        '    frmH.pb2.Visible = False
        '    'frmH.Refresh()
        'End If

        Call OpenAFR(strP, "", False, boolSTB, True, False)

        'Try
        '    wd.Visible = True
        '    wd.Activate()
        'Catch ex As Exception

        'End Try

        'wdSt.Application.Quit()
        wdSt = Nothing

        Cursor.Current = Cursors.Default

        'str1 = "Example Report Body"
        'Call ReportHistoryItem(str1)

        boolShowExample = False

        If gboolER Then
        Else
            frmWordStatement.Activate()
        End If

        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False
        'frmH.pb2.Visible = False

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()

        frmH.Refresh()


        Exit Sub

end2:
        'frmH.pb1.Visible = False
        'frmH.lblProgress.Visible = False
        'frmH.lblProgress.Text = ""

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()

        frmH.Refresh()
        Cursor.Current = Cursors.Default
        boolShowExample = False


    End Sub

    Sub PrepareReport()

        Dim var1, var2, var3
        Dim int1 As Short
        Dim int2 As Short
        Dim dv As system.data.dataview
        Dim drow As DataRow
        Dim tbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strSQL As String
        Dim dbPath As String
        Dim strTemplate As String
        Dim Count1 As Short
        Dim Count2 As Short
        Dim fi
        Dim ctCols As Short
        Dim intPMax As Short
        Dim intPCt As Short

        'check for permission
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short
        Dim boolWatermark As Boolean
        Dim rngA As Microsoft.Office.Interop.Word.Range
        Dim intSR As Short
        Dim boolOKTF As Boolean 'Keep Track of Formatting
        Dim dtStart As DateTime
        Dim dtEnd As DateTime

        Dim posStartTables As Int64

        'Dim fso As FileInfo

        ctAppendix = 0
        ctFigures = 0
        ctAttachments = 0
        tblAppendix.Clear()
        tblFigures.Clear()
        tblAttachments.Clear()


        tPswd = "" 'document pswd if doc is saved

        boolTableSection = False
        intTableSection = 0

        boolWatermark = True
        If BOOLALLOWREPORTGENERATION = False And BOOLALLOWPDFREPORT = False Then
            MsgBox("User does not have privileges to generate a report.", MsgBoxStyle.Information, "No no...")
            Exit Sub
        Else

        End If

        'set arrReportNA
        arrReportNA.Clear(arrReportNA, 0, arrReportNA.Length)
        ctArrReportNA = 0
        intILS = 0
        intStartTable = 2

        dv = frmH.dgvReports.DataSource
        int1 = dv.Count
        'determine ctpbmax
        ctPBMax = dv.Count
        ctPBMax = 30

        If int1 = 0 Then
            MsgBox("A report must be configured.", MsgBoxStyle.Information, "Add a report...")
            frmH.cmdConfigureReport.Select()
            Exit Sub
        End If

        If frmH.dgvReports.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = frmH.dgvReports.CurrentRow.Index
        End If

        'Dim dt As DateTime = DateTifrmh.Now
        Dim dt As DateTime
        dt = Now

        Dim frm As New frmMsgBox
        frm.lblText.Text = "Do you wish to generate an Entire Report?"
        frm.boolTable = False
        If BOOLALLOWREPORTGENERATION Then
            frm.chkWatermark.Visible = True
        Else
            frm.chkWatermark.Visible = False
            frm.chkWatermark.Checked = True
        End If
        'frm.lblReportTemplate.Visible = False
        'frm.cbxReportTemplate.Visible = False
        frm.ShowDialog()
        If frm.boolCancel Then
            GoTo end2
        End If
        strReportTemplateChoice = frm.cbxReportTemplate.Text

        frmH.Refresh()
        If frm.chkWatermark.Checked Then
            boolWatermark = True
        Else
            boolWatermark = False
        End If
        frm.Close()

        'var1 = MsgBox("Click OK to continue...", vbInformation + vbOKCancel, "Generate Report...")
        'If var1 = 1 Then 'continue
        'Else
        '    GoTo end2
        'End If

        Dim intCT As Short
        intCT = 0

        ''create strPathwd
        'strPathWd = ""
        'strPathWd = GetNewTempFile() 'DON'T SKIP THIS!!! Paste routine uses this

        boolDoTables = True

        'must make Report Body 'show included'
        If frmH.rbShowIncludedRBody.Checked Then
        Else
            frmH.rbShowIncludedRBody.Checked = True
        End If

        If frmH.rbShowIncludedRTConfig.Checked Then
        Else
            frmH.rbShowIncludedRTConfig.Checked = True
        End If

        intTCur = 0
        'intTTot = getIntTTot(False)
        Try
            intTTot = getIntTTot(False)
        Catch ex As Exception
            var1 = ex.Message
            intTTot = 0
        End Try

        Call ClearAllQCTables()

        dtStart = Now
        LTableDateTimeStamp = dtStart

        Cursor.Current = Cursors.WaitCursor

        str1 = "Opening Microsoft® Word..." ' & Chr(10) & Chr(10) & "Item " & intPCt & " of " & intPMax

        Dim pb2Max As Short
        Dim ctPB2 As Short
        Call PositionProgress()
        frmH.lblProgress.Text = str1
        'jeez
        frmH.lblProgress.Visible = True
        frmH.lblProgress.Refresh()

        'frmH.panProgress.Visible = True
        'frmH.panProgress.Refresh()

        If ctPBMax < 2 Then
            ctPBMax = 30
        End If

        pb2Max = 10
        ctPB2 = 0

        frmH.pb1.Maximum = ctPBMax
        ctPB = 0
        ctPB = ctPB + 1
        frmH.pb1.Value = ctPB
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        frmH.pb2.Left = frmH.pb1.Left
        frmH.pb2.Top = frmH.pb1.Top + frmH.pb1.Height + 2
        frmH.pb2.Width = frmH.pb1.Width

        frmH.pb2.Maximum = pb2Max
        frmH.pb2.BringToFront()

        ctPB2 = ctPB2 + 1
        frmH.pb2.Value = ctPB2
        frmH.pb2.Visible = True
        frmH.pb2.Refresh()

        'begin preparing tables for for use
        '***Start 1
        boolReportCont = True
        boolAppendix = False

        Cursor.Current = Cursors.WaitCursor

        ctTableN = 0
        'clear contents of tblTableN
        tblTableN.Clear()

        ''clear contents of tblOutstandingItems'DO THIS IN UPDATEOUTSTANDINGREPORT
        intErrCount = 0
        'str1 = "id_tblStudies = " & id_tblStudies
        'Dim drowsO() As DataRow
        'drowsO = tblOutstandingItems.Select(str1)
        'int1 = drowsO.Length
        'For Count1 = 0 To int1 - 1
        '    tblOutstandingItems.Rows.Remove(drowsO(Count1))
        '    'drowsO(Count1).Delete()
        'Next

        'clear contents of tblAppendix
        tblAppendix.Clear()
        ctAppendix = 0
        tblFigures.Clear()
        ctFigures = 0
        tblAttachments.Clear()
        ctAttachments = 0

        ''***End 1

        '***Start 2
        'Dim wd As Object
        Dim drows() As DataRow
        Dim strPath As String
        Dim strPathGuWu As String

        ''find template path
        'str1 = "charConfigTitle = 'Report Templates'"
        tbl = tblConfiguration

        Erase drows

        intPCt = intPCt + 1
        intPMax = 10

        boolTOC = False
        boolTOA = False
        boolTOT = False
        boolTOF = False

        'wd = CreateObject("Word.Application")
        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim boolSTB As Boolean = True
        Try
            boolSTB = wd.Application.ShowWindowsInTaskbar
        Catch ex As Exception
            var1 = var1 'debub
        End Try

        Dim intCBS As Int64

        'create strPathwd
        strPathWd = ""
        strPathWd = GetNewTempFile(True)

        ReDim garrMargins(1, 4) 'T,L,B,R for Portrait:  1=P, 2=L

        Dim boolIsDoGuWuFast As Boolean = False

        Try

            'Instead, open Report Template

            intCBS = frmH.dgvReportStatements("ID_TBLWORDSTATEMENTS", 0).Value
            'find new intcbs
            Dim intCBSa As Int64
            intCBSa = GetNewCBS()
            If intCBSa = 0 Then
            Else
                intCBS = intCBSa
            End If

            strTemplate = strPathWd
            strTemplate = OpenTemplate(intCBS, strPathWd)
            wd.Documents.Open(FileName:=strTemplate, ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", Format:=Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAuto, XMLTransform:="")

            'wd.Documents.Add() '(Template:="Normal", NewTemplate:=False, DocumentType:=0)
            'wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument, LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False)

            Dim ver
            ver = wd.Version
            Dim bool2007 As Boolean
            bool2007 = True
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If
            Dim strExt As String = ".docx"

            boolHasMacro = False
            If bool2007 Then
                'Note: Word2007 has no compatability mode parameter
                If wd.ActiveDocument.HasVBProject Then
                    strExt = ".docm"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                End If
            Else
                If wd.ActiveDocument.HasVBProject Then
                    strExt = ".docm"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs2(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs2(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                End If
            End If

            Cursor.Current = Cursors.WaitCursor

            'check for DoGuWuFast
            boolIsDoGuWuFast = IsGuWuFast(wd)

            If boolHasMacro Then
                boolSaveAsDocx = SaveAsDocx(wd)
            Else
                boolSaveAsDocx = True
            End If

            'Call CallMsgBox(1, wd)

            '20180701 LEE:
            Call SetDocGlobal(wd.ActiveDocument)
            '20180701 LEE
            'Implement time saving trick
            Call SpellingOff(wd.ActiveDocument, False)

            Call SetNormal(wd)

            'Call CallMsgBox(2, wd)

            If boolVerbose Then
                wd.Visible = True
                wd.WindowState = Office.Interop.Word.WdWindowState.wdWindowStateMaximize
            End If

            'set normal font for later font problems in TOC
            Try
                ftNormal = wd.Selection.Font.Name
            Catch ex As Exception
                'MsgBox("ftNormal: " & ex.Message)
                ftNormal = "Arial"
            End Try


            'record and set some defaults
            Try
                boolOKTF = wd.Options.FormatScanning
                wd.Options.FormatScanning = False
            Catch ex As Exception
                'MsgBox("boolOKTF: " & ex.Message)
            End Try

            'record normal font
            Try
                NormalFontsize = wd.Selection.Font.Size
            Catch ex As Exception
                NormalFontsize = 11
                'MsgBox("NormalFontsize: " & ex.Message)
            End Try

            intOTables = wd.ActiveDocument.Tables.Count

            'add BlueHyperlink style
            If boolBLUEHYPERLINK And boolIsDoGuWuFast = False Then
                Call CreateBlueHyperlink(wd)
            End If

            ''immediately add a paragraph return to footer
            ''to compensate when a watermark gets inserted, the useable page space seems to decrease by 1 row
            ''enter footer
            'If wd.ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
            '    wd.ActiveWindow.Panes(2).Close()
            'End If
            'If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow. _
            '    ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
            '    wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            'End If
            'wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter
            'wd.Selection.TypeParagraph()
            'wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

            ''for word 2000
            ''wdd.visible = True
            ''frmH.Activate()

        Catch ex As Exception
            'str1 = "Hmmm. The report template " & Chr(10) & Chr(10) & strTemplate & Chr(10) & Chr(10) & "could not be found."
            str1 = "Hmmm. There was a problem communicating with the Report Template."
            'str1 = str1 & str2 & Chr(10) & Chr(10) & ex.Message
            str1 = str1 & Chr(10) & Chr(10) & ex.Message & Chr(10) & Chr(10) & str2
            MsgBox(str1, MsgBoxStyle.Information, "File not found...")
            boolReportCont = False
            GoTo end3
        End Try

        'wdd.visible = True

        wd.ActiveWindow.View.ShowFieldCodes = False

        'get ReportStatements.doc path
        Dim wdSt As Object ' new Microsoft.Office.Interop.Word.application

        ctPB2 = ctPB2 + 1
        frmH.pb2.Value = ctPB2
        frmH.pb2.Visible = True
        frmH.Refresh()
        frmH.pb2.Refresh()

        ReDim ctrsSamples(5, ctAnalytes)
        ReDim ctrsRepeat(5, ctAnalytes)
        ReDim ctrsReassayed(5, ctAnalytes)
        ReDim ctrsISR(5, ctAnalytes)
        '1=#samples, 2=RunID string


        'skip GuWuDoc opening
        'GoTo skip1

        'str1 = "charConfigTitle = 'Report Statements'"
        'Erase drows
        'drows = tbl.Select(str1)
        'strPath = drows(0).Item("charConfigValue")
        'Erase drows
        'str1 = "charConfigTitle = 'GuWu Statements'"
        'drows = tbl.Select(str1)
        'strPathGuWu = drows(0).Item("charConfigValue")


        'generate Watson tables

        '******

        Cursor.Current = Cursors.WaitCursor

        Dim dv1 As system.data.dataview
        Dim arr1(3, 3)
        '1=dgv row number, 2=header text, 3=ID_TBLCONFIGBODYSECTIONS
        Dim boolGo As Boolean

        intCT = 0
        If boolEntireReport Then
            'prepare dv1 from tblReportStatements
            Dim tblR As System.Data.DataTable
            Dim strFR As String

            tblR = tblReportstatements
            strFR = "ID_TBLSTUDIES = " & id_tblStudies
            dv1 = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)

        Else
            dv1 = frmH.dgvReportStatements.DataSource
        End If


        For Count2 = 1 To 3
            Select Case Count2
                Case 1
                    int1 = 139 'figures
                Case 2
                    int1 = 140 'tables
                Case 3
                    int1 = 141 'appendices
            End Select
            int2 = FindRowDVByCol(CStr(int1), dv1, "ID_TBLCONFIGBODYSECTIONS")
            If int2 = -1 Then
            Else
                intCT = intCT + 1
                arr1(1, intCT) = int2
                arr1(2, intCT) = NZ(dv1(int2).Item("CHARHEADINGTEXT"), "[NO CAPTION]")
                arr1(3, intCT) = int1
            End If
        Next

        If intCT = 0 Then
            MsgBox("Hmmm. Tables, Figures, and Appendices sections have not been configured in the Report Body Section. Please investigate.", MsgBoxStyle.Information, "Nothing to generate")
            GoTo end1
        End If

        'sort asc arr1
        For Count1 = 1 To intCT - 1
            int1 = arr1(1, Count1)
            For Count2 = Count1 + 1 To intCT
                int2 = arr1(1, Count2)
                If int2 < int1 Then
                    var1 = arr1(1, Count1)
                    var2 = arr1(2, Count1)
                    var3 = arr1(3, Count1)

                    arr1(1, Count1) = arr1(1, Count2)
                    arr1(2, Count1) = arr1(2, Count2)
                    arr1(3, Count1) = arr1(3, Count2)

                    arr1(1, Count2) = var1
                    arr1(2, Count2) = var2
                    arr1(3, Count2) = var3

                    int1 = arr1(1, Count1)
                End If
            Next
        Next
        Call PositionProgress()
        frmH.lblProgress.Text = "Creating GuWuTOF Style..."
        frmH.lblProgress.Visible = True
        frmH.lblProgress.Refresh()
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        'frmH.panProgress.Visible = True
        'frmH.panProgress.Refresh()

        'Call CreateGuWuTOF(wd)

        boolMeRefresh = False

        'wdd.visible = True

        Dim fontsize, varAlign

        'wdd.visible = True

        intSR = 1

        'wdd.visible = True

        If boolEntireReport Then 'prepare sections in different order
            boolMeRefresh = True

            'wdd.visible = True

            If boolReportCont Then 'continue

                Cursor.Current = Cursors.WaitCursor
                If boolIsDoGuWuFast Then
                Else
                    Try
                        Call ReportBody(wd, wdSt, boolWatermark, False)
                    Catch ex As Exception
                        str1 = "There was a problem creating the Body section of the report." & ChrW(10) & ChrW(10)
                        str1 = str1 & "Try re-generating the report. If you still encounter this error, please contact your StudyDoc Administrator." & ChrW(10) & ChrW(10)
                        MsgBox(str1, MsgBoxStyle.Information, "Report Body Section error...")
                        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    End Try
                End If


                Cursor.Current = Cursors.WaitCursor
            Else
                MsgBox("There was a problem preparing the Watson tables. Please contact your LABIntegrity StudyDoc administrator.", MsgBoxStyle.Information, "Problem preparing Watson tables...")
                GoTo end1
            End If

            're-establish ctTableN
            'wdd.visible = True

            '20180912 LEE:
            'Hmm. This doesn't work if tables are embedded in sections
            '1st check if report has [TABLESECTION] field code

            ctTableN = 0
            gboolTableSection = False
            gboolTSDone = False
            Dim rng As Microsoft.Office.Interop.Word.Range
            rng = wd.ActiveDocument.Content
            With rng.Find
                .ClearFormatting()
                .Forward = True
                .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                .Format = True
                .MatchCase = True
                .Execute(FindText:="[TABLESECTION]")
                If .Found Then
                    gboolTableSection = True
                End If
                .ClearFormatting()
            End With


            If gboolTableSection Then
                gboolTSDone = True
            Else
                'this means report has individual tables and need to be enumerated according to position in report
                'call FillTS in IncrNextTableNumber
                GoTo skip1
            End If

            '20151027 Larry:
            'funny behavior here. If first character of next line is '[', then a space is placed before it after deleting that first line

            Dim vChar1
            Dim rngChar1 As Microsoft.Office.Interop.Word.Range
            Dim pos1, pos2
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            'record first character this line
            pos1 = wd.Selection.Start
            pos2 = pos1 + 1
            rngChar1 = wd.ActiveDocument.Range(Start:=pos1, End:=pos2)
            vChar1 = rngChar1.Text

            wd.Selection.TypeParagraph()
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            wd.Selection.Style = wd.ActiveDocument.Styles("Normal")


            Dim boolTN As Boolean
            boolTN = False
            ctTableN = 0

            Do Until boolTN = True
                ctTableN = ctTableN + 1
                Try
                    wd.Selection.InsertCrossReference(ReferenceType:="Table", ReferenceKind:=Microsoft.Office.Interop.Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=CStr(ctTableN), InsertAsHyperlink:=True, IncludePosition:=False) ', SeparateNumbers:=False, SeparatorString:=" ")
                Catch ex As Exception
                    boolTN = True
                    ctTableN = ctTableN - 1
                End Try
            Loop


            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            pos1 = wd.Selection.Start
            wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            '20151027 Larry:
            'funny behavior here. If first character of next line is '[', then a space is placed before it after deleting that first line
            If StrComp(vChar1, "[", CompareMethod.Text) = 0 Then
                wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
            Else
                wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
            End If

skip1:

            ctPB2 = ctPB2 + 1
            frmH.pb2.Value = ctPB2
            'frmH.pb2.Visible = True
            frmH.pb2.Refresh()

            'now do searchreplace
            'hmm try running this earlier
            'If the report has a large amount of tables, then Search Replace can take a long time
            'Can't do it now. Something is erroring out when this is run again later

            System.Windows.Forms.Application.DoEvents()

            Cursor.Current = Cursors.WaitCursor

            Call DoSearchReplace(wd, boolWatermark, False, True, True, False, boolIsDoGuWuFast)

            Cursor.Current = Cursors.WaitCursor

            System.Windows.Forms.Application.DoEvents()

            If boolIsDoGuWuFast Then
                GoTo GuWuFast
            End If

            Call InsertTableReferences(wd)
            'now do Table variable returns
            'Call ReturnAnovaValues(wd)

            ''try setting Table Of Figures style formatting early
            'Try
            '    Call FormatIndex(wd, 84, 19)
            'Catch ex As Exception

            'End Try

            'wdd.visible = True

            'Application.DoEvents()

            ''Call CallMsgBox(1, wd)

            Cursor.Current = Cursors.WaitCursor

            Try
                'wdd.visible = False
                Call SignatureSearch(wd)
            Catch ex As Exception
                str1 = "There was a problem completing the Signature Search/Replace action. Report generation will continue."
                str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
            End Try

            'now enter figs, tables and appendices in order

            Cursor.Current = Cursors.WaitCursor

            frmH.lblProgress.Text = "Preparing Figure/Table/Appendix Section..."
            frmH.lblProgress.Refresh()

            For Count1 = 1 To intCT

                'wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                'str1 = arr1(2, Count1)
                str1 = "Report Body Sections"
                int1 = arr1(1, Count1)
                int2 = arr1(3, Count1)

                Cursor.Current = Cursors.WaitCursor

                'rngA = wd.Selection.Range

                Try
                    rngA = wd.Selection.Range
                Catch ex As Exception

                End Try
                intSR = 0

                'jeez\
                'wdd.visible = True

                System.Windows.Forms.Application.DoEvents()

                Select Case int2
                    Case 139 'figures
                        Try
                            frmH.lblProgress.Text = "Inserting Individual Figures..."
                            frmH.lblProgress.Refresh()

                            Call InsertIndividualFigs(wd)
                            Call SearchReplace(wd, "Report Body", rngA, False, "", intSR, 187, 187, False, True, False, 0) '[FIGURESECTION]
                        Catch ex As Exception

                        End Try
                    Case 140 'tables
                        Try
                            frmH.lblProgress.Text = "Preparing Report Body...Tables..."
                            frmH.lblProgress.Refresh()

                            Call SearchReplace(wd, "Report Body", rngA, False, "", intSR, 185, 185, False, True, False, 0) '[TABLESECTION]
                        Catch ex As Exception

                        End Try

                        Try
                            frmH.lblProgress.Text = "Executing Repeat Section actions..."
                            frmH.lblProgress.Refresh()

                            Call SearchRepeat(wd)

                            frmH.lblProgress.Text = "Executing Field Code actions..."
                            frmH.lblProgress.Refresh()

                            Call ReturnAnovaValues(wd)
                        Catch ex As Exception

                        End Try

                    Case 141 'appendices
                        Try
                            frmH.lblProgress.Text = "Preparing Report Body...Appendices..."
                            frmH.lblProgress.Refresh()

                            Call SearchReplace(wd, "Report Body", rngA, False, "", intSR, 186, 186, False, True, False, 0) '[APPENDIXSECTION]
                        Catch ex As Exception

                        End Try

                End Select

            Next

            'Call CallMsgBox(2, wd)

        Else

            'don't do the following because a non-fieldcode report is requested
            'Try
            '    'wdd.visible = False
            '    Call SignatureSearch(wd)
            'Catch ex As Exception
            '    str1 = "There was a problem completing the Signature Search/Replace action. Report generation will continue."
            '    str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
            'End Try

            'now enter figs, tables and appendices in order

            'now do searchreplace
            'hmm try running this earlier
            'If the report has a large amount of tables, then Search Replace can take a long time
            'Can't do it now. Something is erroring out when this is run again later

            System.Windows.Forms.Application.DoEvents()

            Cursor.Current = Cursors.WaitCursor

            Call DoSearchReplace(wd, boolWatermark, False, True, True, False, boolIsDoGuWuFast)

            System.Windows.Forms.Application.DoEvents()

            Cursor.Current = Cursors.WaitCursor

            If boolIsDoGuWuFast Then
                GoTo GuWuFast
            End If

            frmH.lblProgress.Text = "Preparing Figure/Table/Appendix Section..."
            frmH.lblProgress.Refresh()

            For Count1 = 1 To intCT
                'str1 = arr1(2, Count1)
                str1 = "Report Body Sections"
                int1 = arr1(1, Count1)
                int2 = arr1(3, Count1)

                Cursor.Current = Cursors.WaitCursor

                System.Windows.Forms.Application.DoEvents()

                Select Case int2
                    Case 139 'figures
                        Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                        'Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                        boolAppFigSectionStart = True
                        Call InsertGraphics("Figure", wd, False, 0, False)
                    Case 140 'tables
                        Try
                            boolAppFigSectionStart = True
                            Call PrepareWatson(wd)
                        Catch ex As Exception
                            str1 = "There seems to have been a problem creating the figure/table/appendix portion of this report." & ChrW(10) & ChrW(10)
                            str1 = str1 & "Try generating the report again." & ChrW(10) & ChrW(10)
                            str1 = str1 & "If the problem persists, please contract your StudyDoc Administrator."
                            MsgBox(str1, MsgBoxStyle.Information, "Error in report body...")
                            'wdd.visible = True
                            GoTo end1
                        End Try
                    Case 141 'appendices
                        Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                        'Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                        boolAppFigSectionStart = True
                        Call InsertGraphics("Appendix", wd, False, 0, False)
                End Select

            Next

            boolMeRefresh = True

            System.Windows.Forms.Application.DoEvents()

            Cursor.Current = Cursors.WaitCursor

            frmH.lblProgress.Text = "Preparing Report Body..."
            frmH.lblProgress.Refresh()

            If boolReportCont Then 'continue
                Cursor.Current = Cursors.WaitCursor
                Try

                    frmH.lblProgress.Text = "Preparing Report Body..."
                    frmH.lblProgress.Refresh()

                    Call ReportBody(wd, wdSt, boolWatermark, False)
                Catch ex As Exception
                    str1 = "There was a problem creating the Body section of the report." & ChrW(10) & ChrW(10)
                    str1 = str1 & "Try re-generating the report. If you still encounter this error, please contact your StudyDoc Administrator." & ChrW(10) & ChrW(10)
                    MsgBox(str1, MsgBoxStyle.Information, "Report Body Section error...")
                    wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                End Try

                Cursor.Current = Cursors.WaitCursor
            Else
                MsgBox("There was a problem preparing the Watson tables. Please contact your LABIntegrity StudyDoc administrator.", MsgBoxStyle.Information, "Problem preparing Watson tables...")
                GoTo end1
            End If

            Cursor.Current = Cursors.WaitCursor

        End If

        Cursor.Current = Cursors.WaitCursor

        'Call CallMsgBox(3, wd)

        'ctPB2 = frmH.pb2.Maximum - 4
        'frmH.pb2.Value = ctPB2

        If frmH.pb2.Maximum <= 0 Then
            frmH.pb2.Maximum = 100
        End If
        ctPB2 = frmH.pb2.Maximum - 4
        frmH.pb2.Value = ctPB2
        'frmH.pb2.Visible = True
        frmH.pb2.Refresh()

        System.Windows.Forms.Application.DoEvents()

        Cursor.Current = Cursors.WaitCursor

        Try 'do this search first because table field codes have other field codes embedded in the table names
            Call RunIDSearch(wd)
        Catch ex As Exception
            str1 = "There was a problem completing the RunID Search action: " & intSR & ". Report generation will continue."
            str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
        End Try

        'the Word document may be huge by now, size may cause problems
        'try saving, opening, closing

        'Call CallMsgBox(4, wd)

        Cursor.Current = Cursors.WaitCursor

        Try
            Dim boolVis As Boolean = False
            boolVis = wd.Visible

            wd.ActiveDocument.Save()
            Pause(2)
            wd.ActiveDocument.Close()
            Pause(2)
            wd.Documents.Open(FileName:=strTemplate, ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
                    PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", Format:=Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAuto, XMLTransform:="")

            If boolVis Then
            Else
                wd.Visible = False
            End If
            Pause(2)
        Catch ex As Exception
            GoTo end3
        End Try

        Cursor.Current = Cursors.WaitCursor

        'now do searchreplace
        'hmm try running this earlier
        'If the report has a large amount of tables, then Search Replace can take a long time
        'BUT: Tables may have search replace stuff, so do again

        Cursor.Current = Cursors.WaitCursor

        System.Windows.Forms.Application.DoEvents()

        Call DoSearchReplace(wd, boolWatermark, False, False, False, True, boolIsDoGuWuFast)

        'Call CallMsgBox(5, wd)

        Cursor.Current = Cursors.WaitCursor

        Try
            wd.ActiveWindow.View.ShowFieldCodes = False
        Catch ex As Exception
            GoTo end3
        End Try

        'wdd.visible = True

        ctPB2 = frmH.pb2.Maximum - 3
        frmH.pb2.Value = ctPB2
        'frmH.pb2.Visible = True
        frmH.pb2.Refresh()


end1:

        Cursor.Current = Cursors.WaitCursor

        'do headers
        frmH.lblProgress.Text = "Creating headers..."
        frmH.lblProgress.Refresh()

        'look for first page special
        Dim dtblF As System.Data.DataTable
        Dim strFF As String
        Dim rowsF() As DataRow
        Dim boolF As Boolean

        strFF = "ID_TBLSTUDIES = " & id_tblStudies
        dtblF = tblReportHeaders
        rowsF = dtblF.Select(strFF)
        Try
            var1 = rowsF(0).Item("BOOLDIFFFIRSTPAGE")
            If var1 = -1 Then
                boolF = True
            Else
                boolF = False
            End If
        Catch ex As Exception

        End Try

        'wdd.visible = True

        If boolF Then

            'wdd.visible = True

            'insert a nextpagebreak
            'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
            With wd.Selection.PageSetup
                '.DifferentFirstPageHeaderFooter = True
            End With
        End If

        'configure all sections as linktoprevious, except for 1 and 2

        System.Windows.Forms.Application.DoEvents()

        'Call CallMsgBox(6, wd)

        Cursor.Current = Cursors.WaitCursor

        Call LinkPrevious(wd, True, False)

        Cursor.Current = Cursors.WaitCursor

        'try saving the document to head off possible memory errors when document is large
        'wd.ActiveDocument.Save()

        System.Windows.Forms.Application.DoEvents()

        Try
            'wdd.visible = False
            Call EnterHeaders(wd)
        Catch ex As Exception
            str1 = "There was a problem creating headers for this report." & ChrW(10) & ChrW(10)
            str1 = str1 & "Try re-generating the report. If you still encounter this error, please contact your StudyDoc administrator." & ChrW(10) & ChrW(10)
            MsgBox(str1, MsgBoxStyle.Information, "Report Header error...")

        End Try

        Cursor.Current = Cursors.WaitCursor

        'Call CallMsgBox(7, wd)

        ctPB2 = frmH.pb2.Maximum - 2
        frmH.pb2.Value = ctPB2
        frmH.pb2.Refresh()

        'try saving the document to head off possible memory errors when document is large
        'wd.ActiveDocument.Save()

        System.Windows.Forms.Application.DoEvents()

        Try
            'wdd.visible = False
            Call EnterFooters(wd)
        Catch ex As Exception
            str1 = "There was a problem creating footers for this report." & ChrW(10) & ChrW(10)
            str1 = str1 & "Try re-generating the report. If you still encounter this error, please contact your StudyDoc administrator." & ChrW(10) & ChrW(10)
            MsgBox(str1, MsgBoxStyle.Information, "Report Header error...")

        End Try

        Cursor.Current = Cursors.WaitCursor

        'Call CallMsgBox(8, wd)

        'now set margins for individual sections

        System.Windows.Forms.Application.DoEvents()

        Try
            Call LinkPrevious(wd, False, False)

        Catch ex As Exception

        End Try

        Cursor.Current = Cursors.WaitCursor

        '20190221 LEE:
        'do searchreplace in footnotes
        'https://www.extendoffice.com/documents/word/5428-word-select-all-footnotes.html
        Dim xDoc As Microsoft.Office.Interop.Word.Document
        Dim xRange As Microsoft.Office.Interop.Word.Range
        xDoc = wd.ActiveDocument
        If xDoc.Footnotes.Count > 0 Then
            xRange = xDoc.Footnotes(1).Range
            xRange.WholeStory()
            xRange.Select()

            Try
                Call SearchReplace(wd, "Report Body", xRange, False, "", intSR, 0, 0, False, True, True, 1)
            Catch ex As Exception
                str1 = "There was a problem completing the Search/Replace action: " & intSR & ". Report generation will continue."
                str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
                MsgBox(str1)
            End Try

        End If



        Cursor.Current = Cursors.WaitCursor

        ctPB2 = frmH.pb2.Maximum - 1
        frmH.pb2.Value = ctPB2
        frmH.pb2.Refresh()

        'wdd.visible = True
        'frmH.Activate()

        Cursor.Current = Cursors.WaitCursor

        System.Windows.Forms.Application.DoEvents()

        Try
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
        Catch ex As Exception
            'frmH.lblProgress.Visible = False
            'frmH.pb1.Visible = False

            'frmH.panProgress.Visible = False
            'frmH.panProgress.Refresh()

            GoTo end3
        End Try
        'update index table page numbers
        frmH.lblProgress.Text = "Updating Tables of Figures..." & ChrW(10) & "...this may take several minutes..."
        frmH.lblProgress.Refresh()

        'wdd.visible = True

        'update index table page numbers
        'This HAS to be done in case a client TOC has been configured
        Dim ctLoop As Short = 0
        Dim strLoop As String

        System.Windows.Forms.Application.DoEvents()

        'Call CallMsgBox(9, wd)

        Call ModifyTOF(wd)

        Cursor.Current = Cursors.WaitCursor

        Try
            Dim tofLoop As Microsoft.Office.Interop.Word.TableOfFigures
            For Each tofLoop In wd.ActiveDocument.TablesOfFigures

                Cursor.Current = Cursors.WaitCursor

                System.Windows.Forms.Application.DoEvents()

                ctLoop = ctLoop + 1
                'wdd.visible = True
                strLoop = tofLoop.Caption.ToString
                'var1 = tofLoop.UseHeadingStyles
                'tofLoop.Update()

                Select Case strLoop
                    Case "Table"
                        'select table of tables
                        If FindTOF(wd, "of Tables") Then
                            tofLoop.UpdatePageNumbers()
                            wd.Selection.Tables.Item(1).Cell(2, 1).Select()
                            Call ApplyWordTOF(wd)
                        Else
                            tofLoop.Update()
                        End If

                        'Call CallMsgBox(10, wd)

                    Case "Figure"
                        If FindTOF(wd, "of Figures") Then
                            tofLoop.UpdatePageNumbers()
                            wd.Selection.Tables.Item(1).Cell(2, 1).Select()
                            Call ApplyWordTOF(wd)
                        Else
                            tofLoop.Update()
                        End If

                        'Call CallMsgBox(11, wd)

                    Case "Appendix"
                        If FindTOF(wd, "of Appendices") Then
                            tofLoop.UpdatePageNumbers()
                            wd.Selection.Tables.Item(1).Cell(2, 1).Select()
                            Call ApplyWordTOF(wd)
                        Else
                            tofLoop.Update()
                        End If

                        'Call CallMsgBox(12, wd)

                End Select
            Next tofLoop

        Catch ex As Exception

        End Try

        Cursor.Current = Cursors.WaitCursor

        'Call ModifyTOF(wd)
        If boolTOF Or boolTOA Or boolTOT Then
            Call ModifyTOC(wd)
        End If

        'Call CallMsgBox(13, wd)

        'update index table page numbers
        frmH.lblProgress.Text = "Updating Table of Contents..."
        frmH.lblProgress.Refresh()

        'This HAS to be done in case a client TOC has been configured

        Try
            Dim tocLoop As Microsoft.Office.Interop.Word.TableOfContents
            ctLoop = 0
            For Each tocLoop In wd.ActiveDocument.TablesOfContents

                Cursor.Current = Cursors.WaitCursor

                System.Windows.Forms.Application.DoEvents()

                ctLoop = ctLoop + 1
                'tocLoop.Update()
                'tocLoop.UpdatePageNumbers()

                'If FindTOF(wd, "Table of Contents") Then
                '    'do nothing because TOC was last thing to be prepared
                '    'wd.Selection.Tables.Item(1).Cell(2, 1).Select()
                '    'wd.Selection.Fields.Update()
                '    'Call ApplyWordTOF(wd)
                'Else
                '    tocLoop.Update()
                'End If

                tocLoop.Update()

            Next tocLoop
        Catch ex As Exception

        End Try

        'Call CallMsgBox(14, wd)

        'format toc color to blue

        Cursor.Current = Cursors.WaitCursor

        Try
            If boolBLUEHYPERLINK Then
                Call FormatTOCColor(wd)
            End If
        Catch ex As Exception

        End Try

        'look for locked sections
        frmH.lblProgress.Text = "Searching for Locked Sections..."
        frmH.lblProgress.Refresh()

        Cursor.Current = Cursors.WaitCursor

        Try
            Call LockSections(wd)
        Catch ex As Exception

        End Try

        ctPB2 = frmH.pb2.Maximum
        frmH.pb2.Value = ctPB2
        frmH.pb2.Refresh()

        frmH.pb1.Value = 0
        frmH.pb1.Refresh()

        'wdd.visible = True
        'wd.Selection.WholeStory()
        'wd.Selection.Fields.Update()
        'wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        Cursor.Current = Cursors.WaitCursor

        Call ReportHistoryItem("Final")

        'frmH.lblProgress.Text = "Inserting Watermarks..."
        'frmH.lblProgress.Refresh()

        ''The following code must be performed BEFORE replacement of spaces with npsp
        'If boolWatermark Then
        '    Try
        '        Call InsertWatermark(wd)
        '    Catch ex As Exception
        '        str1 = "Unfortunately, this version of Microsoft" & ChrW(10) & " Word does not contain the Word watermarking funtion supported by GuWu." & ChrW(10) & ChrW(10)
        '        str1 = str1 & "Word 2002 or higher must be used. " & ChrW(10) & ChrW(10)
        '        str1 = str1 & "The report will be prepared without a watermark."
        '        MsgBox(str1, MsgBoxStyle.Information, "Watermark not supported...")
        '    End Try
        'End If
        Try
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            wd.Selection.Find.ClearFormatting()

        Catch ex As Exception

        End Try

        Try
            wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
        Catch ex As Exception

        End Try

        'reset some defaults
        Try
            wd.Options.FormatScanning = boolOKTF
        Catch ex As Exception

        End Try

GuWuFast:

        'fso = Nothing


        Cursor.Current = Cursors.Default

        If intErrCount > 0 Then
            MsgBox(strErrMsg, MsgBoxStyle.Information, "Report Sections not configured...")
            frmH.Refresh()
            'SendKeys.Send("%")

        End If

        'frmH.lblProgress.Text = "Preparing Outstanding Report Items list..."
        'frmH.lblProgress.Refresh()
        ''SendKeys.Send("%")

        'dtEnd = Now
        'If ctArrReportNA > 0 Then
        '    Call UpdateOutstandingItems(dtStart, dtEnd)
        'End If

        'frmH.pb1.Value = frmH.pb1.Maximum
        'frmH.pb2.Value = frmH.pb2.Maximum
        'frmH.pb1.Refresh()
        'frmH.pb2.Refresh()

        ''reset cttablen
        'ctTableN = 0

        '20180701 LEE
        'Implement time saving trick
        Call SpellingOff(wd.ActiveDocument, True)

        Try
            'wdd.visible = True
            'wd.Activate()

            'clear formatting again
            Try
                wd.Selection.Find.ClearFormatting()
            Catch ex As Exception

            End Try
            Try
                wd.Selection.Range.Find.ClearFormatting()
            Catch ex As Exception

            End Try
            'wdd.visible = True
            'wd.Activate()

            'first save file 
            'wd.ActiveDocument.Save()

            'save as temp then display in afr
            Dim strP As String
            strP = GetNewTempFile(True)
            'strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)

            If boolHasMacro And boolSaveAsDocx = False Then
                strP = Replace(strP, ".xml", ".docm", 1, -1, CompareMethod.Text)
            Else
                strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)
            End If

            frmH.lblProgress.Text = "Saving document..." ' & ChrW(10) & ChrW(10) & strP
            frmH.lblProgress.Refresh()

            If boolHasMacro And boolSaveAsDocx = False Then
                wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
            Else
                wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
            End If
            ' wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)

            If gDoPDF Then
                If wd.Version < 12 Then
                    gDoPDF = False
                Else
                    gDoPDF = True
                End If
            End If

            Cursor.Current = Cursors.WaitCursor

            Threading.Thread.Sleep(100)

            If gDoPDF Then
                frmH.lblProgress.Text = "Creating .pdf..."
                frmH.lblProgress.Refresh()
                Call CreatePDF(wd, strP)
            End If

            Try
                frmH.lblProgress.Text = "Checking for custom report..."
                frmH.lblProgress.Refresh()
                Call CheckforDoGuWu(wd, strP)
            Catch ex As Exception

            End Try

            'in some systems, getting 'File In Use' window when afr tries to open strP
            'maybe need to close, then quit
            wd.ActiveDocument.Save()
            Threading.Thread.Sleep(250)

            Try
                wd.ActiveDocument.Close(False)
            Catch ex As Exception

            End Try


            Try
                wd.Application.ShowWindowsInTaskbar = boolSTB
            Catch ex As Exception
                var1 = var1 'debug
            End Try

            Try
                wd.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            Catch ex As Exception

            End Try

            Threading.Thread.Sleep(100)

            If boolIsDoGuWuFast Then
            Else
                frmH.lblProgress.Text = "Preparing Outstanding Report Items list..."
                frmH.lblProgress.Refresh()
                'SendKeys.Send("%")

                dtEnd = Now
                If ctArrReportNA > 0 Then
                    'Call UpdateOutstandingItems(dtStart, dtEnd)
                End If
                Call UpdateOutstandingItems(dtStart, dtEnd)
            End If


            frmH.pb1.Value = frmH.pb1.Maximum
            frmH.pb2.Value = frmH.pb2.Maximum
            frmH.pb1.Refresh()
            frmH.pb2.Refresh()

            'reset cttablen
            ctTableN = 0

            frmH.lblProgress.Text = "Opening Word" & ChrW(8482) & " document..."
            frmH.lblProgress.Refresh()

            If gDoPDF Then
            Else
                Call OpenAFR(strP, "", False, boolSTB, True, False)
            End If

            If gboolER Then
            Else
                frmWordStatement.Activate()
            End If

            'frmH.lblProgress.Visible = False
            'frmH.pb1.Visible = False
            'frmH.pb2.Visible = False

            'frmH.panProgress.Visible = False
            'frmH.panProgress.Refresh()

            frmH.Refresh()

            Cursor.Current = Cursors.WaitCursor

        Catch ex As Exception

        End Try

        wd = Nothing

        boolDoTables = False

end2:

        Cursor.Current = Cursors.WaitCursor

        Call CheckWatsonRecords()

        boolTableSection = False

        Exit Sub

end3:
        Try
            wd.Application.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            wd = Nothing

        Catch ex As Exception

        End Try

        boolDoTables = False

        Try
            'frmH.lblProgress.Visible = False
            'frmH.pb1.Visible = False
            'frmH.pb2.Visible = False

            'frmH.panProgress.Visible = False
            'frmH.panProgress.Refresh()

            frmH.Refresh()

        Catch ex As Exception

        End Try

        Cursor.Current = Cursors.WaitCursor

        'If boolAccess Then
        'Else
        '    Call CheckWatsonRecords()
        'End If
        Call CheckWatsonRecords()

        boolTableSection = False

        Cursor.Current = Cursors.Default

    End Sub

    Sub ReportHistoryItem(ByVal strType As String)

        Dim tblMax As System.Data.DataTable
        Dim rowsMax() As DataRow
        Dim strFMax As String
        Dim maxID
        Dim int1 As Short
        Dim var1, var2
        Dim drow As DataRow
        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView
        Dim dt As Date

        dt = gdtReportDate ' Now

        maxID = GetMaxID("tblReportHistory", 1, True)
        id_tblReportHistory = maxID

        'ID_TBLREPORTHISTORY
        'ID_TBLSTUDIES
        'ID_TBLREPORTS
        'CHARREPORTGENERATEDSTATUS
        'UPSIZE_TS
        'DTREPORTGENERATED
        'CHARREPORTTITLE
        'CHARUSERID
        'CHARUSERNAME

        int1 = frmH.dgvReports.CurrentRow.Index
        dv = frmH.dgvReports.DataSource
        var1 = dv.Item(int1).Item("id_tblReports")

        tbl = tblReportHistory
        drow = tbl.NewRow()
        drow.BeginEdit()
        drow.Item("id_tblReportHistory") = maxID
        drow.Item("id_tblStudies") = id_tblStudies
        drow.Item("id_tblReports") = var1

        If gDoPDF Then
            strType = strType & " (as PDF)"
        Else
            'strType = strType & " (as .doc(x))"
            strType = strType & " (as .doc(x))"
        End If

        gCHARREPORTGENERATEDSTATUS = strType

        drow.Item("CHARREPORTGENERATEDSTATUS") = strType
        'var2 = dt 'Format(dt, "mm/dd/yyyy hh:mm:ss AM_PM")
        'str1 = dt.ToString
        drow.Item("DTREPORTGENERATED") = dt
        drow.Item("CHARREPORTTITLE") = dv.Item(int1).Item("CHARREPORTTITLE") 'frmH.lblReportTitle.Text
        drow.Item("CHARUSERID") = gUserID
        drow.Item("CHARUSERNAME") = gUserName

        drow.EndEdit()
        tbl.Rows.Add(drow)

        Try
            If boolGuWuOracle Then
                Try
                    ta_tblReportHistory.Update(tblReportHistory)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLREPORTHISTORY.Merge('ds2005.TBLREPORTHISTORY, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblReportHistoryAcc.Update(tblReportHistory)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTHISTORY.Merge('ds2005Acc.TBLREPORTHISTORY, True)
                    var1 = ex.Message
                    var1 = var1
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblReportHistorySQLServer.Update(tblReportHistory)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLREPORTHISTORY.Merge('ds2005Acc.TBLREPORTHISTORY, True)
                    var1 = ex.Message
                    var1 = var1
                End Try
            End If
        Catch ex1 As Exception
            var1 = ex1.Message
            var1 = var1
        End Try




    End Sub


    Sub BlueHyperlink(ByVal wd As Microsoft.Office.Interop.Word.Application)

        If boolBLUEHYPERLINK Then
        Else
            Exit Sub
        End If

        Dim var1

        Try

            'wd.Selection.Style = wd.ActiveDocument.Styles("BlueHyperlink")

            wd.Selection.Font.ColorIndex = BlueHyperlinkColor

            'Try
            '    wd.ActiveDocument.Styles.Add(Name:="BlueHyperlink Char", Type:=Microsoft.Office.Interop.Word.WdStyleType.wdStyleTypeCharacter)
            '    wd.ActiveDocument.Styles("BlueHyperlink Char").LinkStyle = "BlueHyperlink"
            'Catch ex As Exception

            'End Try
            ''wdd.visible = True
            'wd.Selection.Style = wd.ActiveDocument.Styles("BlueHyperlink Char")
        Catch ex As Exception
            var1 = ex.Message
        End Try
    End Sub

    Sub CreateBlueHyperlink(ByVal wd)

        '20170717 LEE: Deprecated. Only apply color, not style

        Exit Sub

        If boolBLUEHYPERLINK Then
        Else
            Exit Sub
        End If

        Try
            wd.ActiveDocument.Styles.Add(Name:="BlueHyperlink", Type:=Microsoft.Office.Interop.Word.WdStyleType.wdStyleTypeCharacter)
            With wd.ActiveDocument.Styles("BlueHyperlink").Font
                .Name = ""
                .Colorindex = BlueHyperlinkColor '  Microsoft.Office.Interop.Word.WdColor.wdColorBlue
            End With

        Catch ex As Exception

        End Try

    End Sub

    Sub PrepareTable(ByVal int1 As Short, ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal idTR As Int64, ByVal idCRT As Int64, ByVal intRow As Int16, strMM As String)

        strAreaDec = GetRegrDecStr(LSigFigArea)
        strRegrDec = GetRegrDecStr(LRegrSigFigs)
        strR2Dec = GetRegrDecStr(LR2SigFigs)
        strAreaDecAreaRatio = GetAreaRatioDecStr()
        LDec = LSigFig
        LDecArea = LSigFigArea
        LDecAreaRatio = LSigFigAreaRatio
        LR2Dec = LR2SigFigs
        LRegrDec = LRegrSigFigs


        Dim arr1(1) 'blank array for SRSummaryOfBCSC
        Dim dv As System.Data.DataView
        Dim int2 As Short
        Dim var1, var2
        Dim str1 As String
        Dim intIncl As Short
        Dim intReq As Short
        Dim dgvR As DataGridView = frmH.dgvReportTableConfiguration

        dv = dgvR.DataSource

        'set Advanced Table properties
        Call SetTablePropertiesBool(idTR, idCRT)

        '20161206 LEE: Deprecate this ability
        'always use mean rounded
        gboolMeanRounded = True
        gboolMeanFullPrec = False


        System.Windows.Forms.Application.DoEvents()

        'wdd.visible = True

        Cursor.Current = Cursors.WaitCursor

        '20161215 LEE: global parameter that other functions may need
        gidTR = int1


        Try

            Select Case int1
                Case Is = 1 'Summary of Analytical Runs
                    Call SRSummaryOfAnalRuns_1(wd, idTR)
                Case Is = 2 'Summary of Regression Constants

                    '20190304 LEE:
                    'check if samples are to be assigned
                    str1 = "ID_TBLCONFIGREPORTTABLES" 'column name
                    'int2 = FindRowDVNumByCol(int1, dv, str1)

                    int2 = intRow '

                    str1 = "BOOLREQUIRESSAMPLEASSIGNMENT" 'column name
                    intReq = dv(int2).Item(str1)
                    If intReq = 0 Then 'normal
                        'need filler items
                        'Dim tbl1 As System.Data.DataTable
                        'Dim tbl2 As System.Data.DataTable
                        Call SRSummaryOfLSR_2(wd, idTR, False)
                    Else
                        Call SRSummaryOfLSR_2(wd, idTR, True)
                    End If

                    'Call SRSummaryOfLSR(wd, idTR)
                Case Is = 3 'Summary of Back-Calculated Calibration Std Conc
                    'check if samples are to be assigned
                    str1 = "ID_TBLCONFIGREPORTTABLES" 'column name
                    'int2 = FindRowDVNumByCol(int1, dv, str1)

                    int2 = intRow '

                    str1 = "BOOLREQUIRESSAMPLEASSIGNMENT" 'column name
                    intReq = dv(int2).Item(str1)
                    If intReq = 0 Then 'normal
                        'need filler items
                        Dim tbl1 As System.Data.DataTable
                        Dim tbl2 As System.Data.DataTable
                        Call SRSummaryOfBCSC_UseGroups_3(wd, 0, arr1, 0, 0, tbl1, tbl2, idTR)
                    Else
                        Call AssignedBackCalcCalibr_3(wd, idTR)
                    End If
                Case Is = 4 'Summary of Interpolated QC Std Conc
                    'check if samples are to be assigned
                    str1 = "ID_TBLCONFIGREPORTTABLES" 'column name
                    'int2 = FindRowDVNumByCol(int1, dv, str1)

                    int2 = intRow ' 
                    str1 = "BOOLREQUIRESSAMPLEASSIGNMENT" 'column name
                    intReq = dv(int2).Item(str1)

                    If intReq = 0 Then 'normal
                        Call SRSummaryOfIQCCR_UseGroups_4(wd, idTR)
                    Else 'Assign samples
                        Call AssignedQCsTable_4(wd, idTR)
                    End If

                Case Is = 5 'Summary of Samples
                    Call SRSummaryOfSC_5(wd, idTR)
                Case Is = 6 'Summary of Reassayed Samples
                    Call SRSummaryReassaySamplesNew_6(wd, idTR)
                Case Is = 7 'Summary of Repeat Samples
                    Call SRSummaryRepeatSamples_7(wd, idTR)
                Case Is = 8 'Representative Analyte Calibration Curve Figures
                Case Is = 11 'Summary of Interpolated QC Std Conc Intra- and Inter-Run Precision
                    Call MVSummaryOfIQCBetweenRun_11(wd, idTR)
                Case Is = 12 'Summary of Interpolated Dilution QC Concentration
                    'Call MVSummaryDilutionQC_12(wd, idTR)
                    '20191017 LEE: Divert to AdHocStab
                    'must force boolRCConc
                    boolRCConc = True
                    Call MVAdHocQCStability_31(wd, idTR, int1)
                Case Is = 13 'Summary of Combined Recovery
                    Call MVSummaryCombinedRecoveryQC_13(wd, idTR, True, False, False)
                Case Is = 14 'Summary of True Recovery
                    'Call MVSummaryTrueRecoveryQC_14(wd, idTR)
                    Call MVSummaryCombinedRecoveryQC_13(wd, idTR, False, True, False)
                Case Is = 15 'Summary of Suppression/Enhancement
                    'Call MVSummarySuppressionEnhancementQC_15(wd, idTR)
                    Call MVSummaryCombinedRecoveryQC_13(wd, idTR, False, False, True)
                Case Is = 17 'Summary of Interpolated Unique QC Low for Matrix Effects on Quantitation Assessments
                    Call MVSummaryLowQCMatrixEffects_17(wd, idTR)
                Case Is = 18 'Summary of [Temp Descr] Stability in Matrix
                    Call MVSummaryTempStabilityQC_18(wd, idTR)
                Case Is = 19 'Summary of Freeze/Thaw [#Cycles] Stability in Matrix
                    'Call MVSummaryFTStabilityQC_19(wd, idTR)
                    '20191017 LEE: Divert to AdHocStab
                    'must force boolRCConc
                    boolRCConc = True
                    Call MVAdHocQCStability_31(wd, idTR, int1)
                Case Is = 21 '[Period Temp] Final Extract Stability of Interpolated QC Std Concentrations
                    'Call MVSummaryFinalExtractQC_21(wd, idTR)
                    '20191017 LEE: Divert to AdHocStab
                    'must force boolRCConc
                    boolRCConc = True
                    Call MVAdHocQCStability_31(wd, idTR, int1)
                Case Is = 22 '[Period Temp] Stock Solution Stability Assessment
                    'Call MVSummaryStockSolnStability_22(wd, idTR)
                    '20191017 LEE: Divert to AdHocStabComparison
                    Call MVAdHocQCStabilityComparison_32(wd, idTR, int1)
                Case Is = 23 '[Period Temp] Spiking Solution Stability Assessment
                    'Call MVSummarySpikingSolnStability_23(wd, idTR)
                    '20191017 LEE: Divert to AdHocStabComparison
                    Call MVAdHocQCStabilityComparison_32(wd, idTR, int1)
                Case Is = 24 'Summary of Interpolated QC Std Concs Containing Coadministered Compounds
                Case Is = 25 'Summary of Back-Calculated Calibr Std Concs Containing Coadministered Compounds
                Case Is = 27 'PK Curve Figures
                Case Is = 29 '[Period Temp] Long-Term QC Std Storage Stability
                    'Call MVSummaryLongTermQCStability_29(wd, idTR)
                    '20191017 LEE: Divert to AdHocStabComparison
                    'must force boolRCConc
                    boolRCConc = True
                    Call MVAdHocQCStabilityComparison_32(wd, idTR, int1)
                Case Is = 30 'Incurred Samples
                    Call ISR_01_02_30(wd, idTR)
                Case Is = 31 'Ad Hoc QC Stability
                    Call MVAdHocQCStability_31(wd, idTR, int1)
                Case Is = 32 'Ad Hoc QC Stability Comparison
                    Call MVAdHocQCStabilityComparison_32(wd, idTR, int1)
                Case Is = 33 'System Suitability
                    Call MVSystemSuit_v1_33(wd, idTR)
                Case Is = 34
                    Call MVSelectivity_v1_34(wd, idTR)
                Case Is = 35
                    Call MVCarryover_v1_35(wd, idTR)
                Case Is = 36
                    Call MT_CalibrStds_36(wd, idTR)
                Case Is = 37
                    Call MT_QCs_37(wd, idTR)
                Case Is = 38
                    Call MT_IncurredSamples_38(wd, idTR)
            End Select

        Catch ex As Exception


            str1 = "There was a problem preparing table:"
            str1 = str1 & ChrW(10) & ChrW(10) & strMM
            str1 = str1 & ChrW(10) & ChrW(10)
            str1 = str1 & ex.Message

            Dim boolAV As Boolean = wd.Visible
            wd.Visible = True

            MsgBox(str1, vbInformation, "Problem...")

            Try
                'attempt to get out of table
                Call MoveOneCellDown(wd)
            Catch ex1 As Exception
                var1 = ex.Message
                var1 = var1
            End Try

            wd.Visible = boolAV

        End Try


    End Sub

    Sub ExampleSection(strSource As String)

        Dim tblDGV As System.Data.DataTable
        Dim rowsDGV() As DataRow

        'strSource:
        '  - Home
        '  - AssignSamples

        Dim boolHome As Boolean = True
        Dim boolAssSamples As Boolean = False
        Dim lblP As Label = frmH.lblProgress
        Select Case strSource
            Case "Home"
                boolHome = True
                boolAssSamples = False
            Case "AssignSamples"
                boolHome = False
                boolAssSamples = True
                lblP = frmAssignSamples.lblProgress
            Case Else
                boolHome = True
                boolAssSamples = False
        End Select

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim bool As Boolean
        Dim boolGo As Boolean
        Dim var1
        Dim strLBX As String
        Dim str4 As String
        Dim strFilter As String
        Dim dg As DataGrid
        Dim dv As System.Data.DataView
        Dim strDo As String
        Dim boolWD1 As Boolean
        Dim boolWD2 As Boolean
        Dim strSec As String
        Dim strText As String
        Dim boolTables As Boolean
        Dim strTable As String
        Dim Count2 As Short
        Dim strM As String
        Dim intT As Long
        Dim boolARST As Boolean 'Analytical Run Summary Tab
        Dim tbl As System.Data.DataTable
        Dim idTR As Int64
        Dim intRow As Short
        Dim intRowR As Short
        Dim idRT As Int64
        Dim idCRT As Int64
        Dim dt As Date

        strReportTypeApp = ""

        dt = Now

        LTableDateTimeStamp = dt

        tblAppendix.Clear()
        tblFigures.Clear()
        ctAppendix = 0
        ctFigures = 0
        tblAttachments.Clear()
        ctAttachments = 0

        ReDim ctrsSamples(5, ctAnalytes)
        ReDim ctrsRepeat(5, ctAnalytes)
        ReDim ctrsReassayed(5, ctAnalytes)
        ReDim ctrsISR(5, ctAnalytes)

        'first check for configured reports
        '20160823 LEE: This isn't true anymore
        'Dim dvR As System.Data.DataView
        'dvR = frmH.dgvReports.DataSource
        'If dvR.Count = 0 Then
        '    str1 = "A report must be configured on the Home Tab in order to use this function."
        '    MsgBox(str1, MsgBoxStyle.Information, "Please configure a report...")
        '    Exit Sub
        'End If


        boolShowExample = True
        boolARST = False
        ctPB = 0 'pesky
        intILS = 0
        intStartTable = 2

        boolWD1 = False
        boolWD2 = False
        boolTables = False

        ctTableN = 0

        If id_tblStudies = 0 Then
            MsgBox("A study must be chosen.", MsgBoxStyle.Information, "A study must be chosen...")
            GoTo end2
        End If

        Cursor.Current = Cursors.WaitCursor

        'create strPathwd
        Dim Count1 As Short
        Count1 = 0
        strPathWd = ""
        strPathWd = GetNewTempFile(True)

        If boolHome Then
            int1 = frmH.tab1.SelectedIndex
            strLBX = frmH.lbxTab1.Items(int1)
            boolGo = False
            strM = "This does not apply to the chosen Table of Contents selection."
            intRow = 0
            If InStr(1, strLBX, "Review Analytical Runs", CompareMethod.Text) > 0 Then
                boolGo = True
                boolARST = True
                boolTables = True
                'strM = "To show an example Analytical Run Summary example, please choose the Analytical Run Summary entry in the Report Table Configuration Tab."
                'frmh.lbxTab1.SelectedItem = "Report Table Configuration"
            ElseIf InStr(1, strLBX, "Summary Table", CompareMethod.Text) > 0 Then
                boolGo = True
            ElseIf InStr(1, str1, "QA Event Table", CompareMethod.Text) > 0 Or InStr(1, str1, "Contributing Personnel", CompareMethod.Text) > 0 Then
                boolGo = True
            ElseIf InStr(1, strLBX, "Analytical Reference Std", CompareMethod.Text) > 0 Then
                'MsgBox("Analytical Reference Std Table Under Construction...", vbInformation, "Invalid action...")
                'GoTo end3
                boolGo = True
            ElseIf InStr(1, strLBX, "Add/Edit Contributors", CompareMethod.Text) > 0 Then
                boolGo = True
            ElseIf StrComp(strLBX, "Report Body Sections", CompareMethod.Text) = 0 Then
                boolGo = True
                If boolEntireReport Then 'treat as if example report body
                    'Call ExampleReportBody(False)
                    '20190222 LEE: Should be true
                    Call ExampleReportBody(True)
                    GoTo end3
                End If
            ElseIf InStr(1, strLBX, "QA Event Table", CompareMethod.Text) > 0 Then
                boolGo = True
            ElseIf InStr(1, strLBX, "Configure Report Tables", CompareMethod.Text) > 0 Then
                boolGo = True
                boolTables = True
            ElseIf InStr(1, strLBX, "Sample Receipt Records", CompareMethod.Text) > 0 Then
                boolGo = True
            ElseIf InStr(1, strLBX, "Appendices", CompareMethod.Text) > 0 Then
                'clear contents of tblAppendix and others
                tblAppendix.Clear()
                ctAppendix = 0
                tblAttachments.Clear()
                ctAttachments = 0
                tblFigures.Clear()
                ctFigures = 0
                boolGo = True
            End If
        Else
            boolGo = True
            boolTables = True
            strLBX = "Configure Report Tables"
        End If


        If boolGo Then
        Else
            MsgBox(strM, MsgBoxStyle.Information, "Not applicable...")
            boolReportCont = False
            boolGo = False
            boolWD1 = False
            boolWD2 = False

            GoTo end1
        End If

        bool = False
        boolGo = False
        boolWD1 = False
        boolWD2 = False
        strSec = ""
        strText = ""
        If InStr(1, strLBX, "Report Body Sections", CompareMethod.Text) > 0 Then
            'find chosen section name
            If boolEntireReport Then
                'prepare dv1 from tblReportStatements
                Dim tblR As System.Data.DataTable
                Dim strFR As String

                tblR = tblReportStatements
                strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
                dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
                dv.RowFilter = arrRBSColumns(2, 0)

            Else
                dv = frmH.dgvReportStatements.DataSource
            End If
            'dv = frmH.dgvReportStatements.DataSource
            intRow = frmH.dgvReportStatements.CurrentRow.Index
            var1 = dv(intRow).Item("BOOLINCLUDE")
            If var1 = -1 Then 'continue
            Else
                MsgBox("The Include box must be checked before an Example Report Body Section can be prepared.", MsgBoxStyle.Information, "Invalid action...")
                boolReportCont = False
                boolGo = False
                boolWD1 = False
                boolWD2 = False
                GoTo end1
            End If

            'strSec = frmH.dgvReportStatements.Rows.item(int1).Cells("charSectionName").Value
            'strText = frmH.dgvReportStatements.Rows.Item(int1).Cells("CHARHEADINGTEXT").Value

            strSec = dv(intRow).Item("charSectionName")
            strText = dv(intRow).Item("CHARHEADINGTEXT")

            Dim frm As New frmMsgBox
            frm.lblText.Text = "Do you wish to generate an Example Section for " & strText & "?"
            frm.lblReportTemplate.Visible = True
            frm.cbxReportTemplate.Visible = True
            frm.boolTable = False

            frm.ShowDialog()
            If frm.boolCancel Then
                int1 = 7
            Else
                int1 = 6
            End If
            strReportTemplateChoice = frm.cbxReportTemplate.Text

            frmH.Refresh()
            frm.Close()

            strReportTypeApp = strText

            'int1 = MsgBox("Do you wish to generate an Example Section for " & strText & "?", MsgBoxStyle.YesNo, "Generate Example Report...")

        ElseIf InStr(1, strLBX, "Configure Report Tables", CompareMethod.Text) > 0 Then
            'find chosen section name
            dv = frmH.dgvReportTableConfiguration.DataSource
            'intRow = frmH.dgvReportTableConfiguration.CurrentRow.Index

            Try
                intRow = frmH.dgvReportTableConfiguration.CurrentRow.Index
            Catch ex As Exception
                MsgBox("Please select a table", MsgBoxStyle.Information, "Invalid action...")
                boolReportCont = False
                boolGo = False
                boolWD1 = False
                boolWD2 = False
                GoTo end1
            End Try
            strSec = dv(intRow).Item("CHARTABLENAME")
            Dim strTableText As String
            strTableText = dv(intRow).Item("CHARHEADINGTEXT")
            'strSec = frmH.dgvReportTableConfiguration.Rows.Item(int1).Cells("CHARTABLENAME").Value

            Dim frm As New frmMsgBox
            frm.lblText.Text = "Do you wish to generate the table:  " & strTableText & "?"
            frm.lblReportTemplate.Visible = True
            frm.cbxReportTemplate.Visible = True
            frm.boolTable = True

            tblDGV = dv.Table
            rowsDGV = tblDGV.Select("ID_TBLREPORTTABLE = " & frmH.dgvReportTableConfiguration("ID_TBLREPORTTABLE", intRow).Value)
            rowsDGV(0).AcceptChanges()


            frm.rowsDGV = rowsDGV

            frm.ShowDialog()
            If frm.boolCancel Then
                int1 = 7
            Else
                int1 = 6
            End If
            strReportTemplateChoice = frm.cbxReportTemplate.Text

            frmH.Refresh()
            frm.Close()

            strReportTypeApp = strTableText 'strSec

            'int1 = MsgBox("Do you wish to generate an Example Section for " & strSec & "?", MsgBoxStyle.YesNo, "Generate Example Report...")

        Else

            Dim frm As New frmMsgBox
            frm.lblText.Text = "Do you wish to generate an Example Section for " & strLBX & "?"
            frm.lblReportTemplate.Visible = True
            frm.cbxReportTemplate.Visible = True
            frm.boolTable = False

            frm.ShowDialog()
            If frm.boolCancel Then
                int1 = 7
            Else
                int1 = 6
            End If
            strReportTemplateChoice = frm.cbxReportTemplate.Text

            frmH.Refresh()
            frm.Close()

            strReportTypeApp = strLBX

            'int1 = MsgBox("Do you wish to generate an Example Section for " & strLBX & "?", MsgBoxStyle.YesNo, "Generate Example Report...")
        End If
        frmH.Refresh()
        If int1 = 6 Then
        Else
            boolReportCont = False
            boolGo = False
            boolWD1 = False
            boolWD2 = False
            GoTo end1
        End If

        boolGo = True
        Call PositionProgress()

        'frmH.lblProgress.Text = "Opening Microsoft" & ChrW(174) & " Word..." ' data stores..."
        'frmH.lblProgress.Visible = True
        'frmH.lblProgress.Refresh()

        lblP.Text = "Opening Microsoft" & ChrW(174) & " Word..." ' data stores..."
        lblP.Visible = True
        lblP.Refresh()

        'frmH.panProgress.Visible = True
        'frmH.panProgress.Refresh()

        Call PositionAbort(True)


        boolWD1 = True
        boolWD2 = True
        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim boolSTB As Boolean = True
        Try
            boolSTB = wd.Application.ShowWindowsInTaskbar
        Catch ex As Exception

        End Try

        wdAbort = wd

        Dim strTemplate As String
        'Dim tbl as System.Data.DataTable
        Dim drows8() As DataRow
        Dim intCBS As Int64

        tbl = tblConfiguration

        dv = frmH.dgvReports.DataSource
        int1 = frmH.dgvReports.CurrentRow.Index

        ReDim garrMargins(1, 4) 'T,L,B,R for Portrait:  1=P, 2=L

        intTCur = 0
        intTTot = getIntTTot(True)

        Try

            strPathWd = GetNewTempFile(True)

            intCBS = frmH.dgvReportStatements("ID_TBLWORDSTATEMENTS", 0).Value
            'find new intcbs
            Dim intCBSa As Int64
            intCBSa = GetNewCBS()
            If intCBSa = 0 Then
            Else
                intCBS = intCBSa
            End If

            strTemplate = OpenTemplate(intCBS, strPathWd)

            wd.Documents.Open(FileName:=strTemplate, ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
            PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", Format:=Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAuto, XMLTransform:="")

            If wd.ActiveDocument.HasVBProject Then
                boolHasMacro = True
            Else
                boolHasMacro = False
            End If

            If boolHasMacro Then
                boolSaveAsDocx = SaveAsDocx(wd)
            Else
                boolSaveAsDocx = True
            End If

            wd.Selection.WholeStory()
            wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            '20180524 LEE:
            'now save
            wd.ActiveDocument.Save()

            Dim ver
            ver = wd.Version
            Dim bool2007 As Boolean
            bool2007 = True
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If
            Dim strExt As String = ".docx"

            boolHasMacro = False
            If bool2007 Then
                'Note: Word2007 has no compatability mode parameter
                If wd.ActiveDocument.HasVBProject Then
                    strExt = ".docm"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                End If
            Else
                If wd.ActiveDocument.HasVBProject Then
                    strExt = ".docm"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs2(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs2(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                End If
            End If

            '20180701 LEE:
            Call SetDocGlobal(wd.ActiveDocument)
            '20180701 LEE
            'Implement time saving trick
            Call SpellingOff(wd.ActiveDocument, False)

            Call SetNormal(wd)

            If boolVerbose Then
                wd.Visible = True
                wd.WindowState = Office.Interop.Word.WdWindowState.wdWindowStateMaximize
            End If

            ftNormal = wd.Selection.Font.Name

            'record normal font
            NormalFontsize = wd.Selection.Font.Size


            'add BlueHyperlink style
            If boolBLUEHYPERLINK Then
                Call CreateBlueHyperlink(wd)
            End If

            ''immediately add a paragraph return to footer
            ''to compensate when a watermark gets inserted, the useable page space seems to decrease by 1 row
            ''enter footer
            'If wd.ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
            '    wd.ActiveWindow.Panes(2).Close()
            'End If
            'If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow. _
            '    ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
            '    wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            'End If
            'wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter
            'wd.Selection.TypeParagraph()
            'wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

            'for word 2000
            'wdd.visible = True
            'frmH.Activate()
        Catch ex As Exception
            'str1 = "Hmmm. The document " & Chr(10) & Chr(10) & strTemplate & Chr(10) & Chr(10) & "could not be found."
            'str2 = str1 & Chr(10) & Chr(10) & "Reports cannot be generated unless this file is present." & Chr(10) & Chr(10)
            'str2 = str2 & Chr(10) & Chr(10) & "Please make note of this document path and contact a StudyDoc administrator."
            'str2 = str2 & Chr(10) & Chr(10) & "Or better yet, take a screen capture of this message and send it to a StudyDoc administrator."
            'str2 = str2 & Chr(10) & Chr(10) & "Report Generation will be canceled."

            str1 = "There was a problem communicating with the Report Template."
            str1 = str1 & Chr(10) & Chr(10) & "Report Generation will be canceled."
            str1 = str1 & Chr(10) & Chr(10) & ex.Message

            Try
                'frmH.lblProgress.Visible = False
                'frmH.pb1.Visible = False
                'frmH.pb2.Visible = False

                'frmH.panProgress.Visible = False
                'frmH.panProgress.Refresh()

            Catch ex1 As Exception

            End Try

            MsgBox(str1, MsgBoxStyle.Information, "File not found...")
            boolReportCont = False
            boolGo = False
            boolWD1 = False
            boolWD2 = False
            GoTo end1
        End Try

        wd.ActiveWindow.View.ShowFieldCodes = False

        '***Here!!
        'wdd.visible = True

        'frmH.lblProgress.Text = "Preparing Example Report Section..."
        'frmH.lblProgress.Refresh()

        lblP.Text = "Preparing Example Report Section..."
        lblP.Refresh()
        Dim strMM As String
        strMM = ""

        Dim strOrientation As String = "P"

        If boolTables Then

            'set headers and footers before creating table
            Call EnterHeaders(wd)
            Call EnterFooters(wd)

            If InStr(1, strLBX, "Configure Report Tables", CompareMethod.Text) > 0 Or boolARST Then
                'find table
                int1 = frmH.dgvReportTableConfiguration.CurrentRow.Index

                If boolARST Then
                    strTable = "Analytical Run Summary"
                    'must choose appropriate table in dgv
                    For Count1 = 0 To frmH.dgvReportTableConfiguration.RowCount - 1
                        var1 = frmH.dgvReportTableConfiguration("ID_TBLCONFIGREPORTTABLES", Count1).Value
                        If var1 = 1 Then
                            int1 = Count1
                            Exit For
                        End If
                    Next
                End If
                strTable = frmH.dgvReportTableConfiguration.Item("CHARTABLENAME", int1).Value.ToString
                strMM = frmH.dgvReportTableConfiguration.Item("CHARHEADINGTEXT", int1).Value.ToString
                intT = frmH.dgvReportTableConfiguration.Item("id_tblConfigReportTables", int1).Value
                idTR = frmH.dgvReportTableConfiguration.Item("ID_TBLREPORTTABLE", int1).Value
                idCRT = frmH.dgvReportTableConfiguration.Item("ID_TBLCONFIGREPORTTABLES", int1).Value
                boolPlaceHolder = frmH.dgvReportTableConfiguration.Item("BOOLPLACEHOLDER", int1).Value ' dv(Count1).Item("BOOLPLACEHOLDER")
                strOrientation = frmH.dgvReportTableConfiguration.Item("CHARPAGEORIENTATION", int1).Value

                intRowR = int1

                'frmH.lblProgress.Text = ""
                'frmH.lblProgress.Refresh()

                lblP.Text = ""
                lblP.Refresh()

                'add info to document
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                If boolExcludeCoverPage Then
                Else
                    With wd.Selection.ParagraphFormat
                        .LeftIndent = 100 'InchesToPoints(1)
                        '.SpaceBeforeAuto = False
                        '.SpaceAfterAuto = False
                    End With
                    With wd.Selection.ParagraphFormat
                        '.SpaceBeforeAuto = False
                        '.SpaceAfterAuto = False
                        .FirstLineIndent = -100 'InchesToPoints(-1)
                    End With

                    'wdd.visible = True

                    wd.Selection.Font.Bold = True
                    wd.Selection.TypeText(Text:="Study: " & vbTab & frmH.cbxStudy.Text)
                    wd.Selection.TypeParagraph()
                    'wd.Selection.TypeText(Text:="Report Title: " & vbTab & frmH.lblReportTitle.Text)
                    wd.Selection.TypeText(Text:="Report Title: " & vbTab & gReportTitle)
                    wd.Selection.TypeParagraph()
                    strDo = "Example section: " & vbTab & strLBX & ": " & strTable
                    wd.Selection.TypeText(Text:=strDo)
                    wd.Selection.TypeParagraph()
                    'enter date
                    strDo = "Date Prepared: " & vbTab & Format(dt, "MMMM dd, yyyy hh:mm:ss tt")
                    wd.Selection.TypeText(Text:=strDo)
                    wd.Selection.TypeParagraph()
                    With wd.Selection.ParagraphFormat
                        .LeftIndent = 0 'InchesToPoints(1)
                        '.SpaceBeforeAuto = False
                        '.SpaceAfterAuto = False
                    End With
                    With wd.Selection.ParagraphFormat
                        '.SpaceBeforeAuto = False
                        '.SpaceAfterAuto = False
                        .FirstLineIndent = 0 'InchesToPoints(-1)
                    End With

                    wd.Selection.TypeParagraph()
                    wd.Selection.Font.Bold = False

                    Call ExampleCoverPageBreak(wd)

                End If


                'wd.Visible = True

                Try
                    'If gboolReadOnlyTables Then
                    '    xlROT = New Microsoft.Office.Interop.Excel.Application
                    '    xlROT.Workbooks.Add()

                    'End If

                    '20170107 LEE:
                    'must clear tblQCTables before calling
                    Try
                        tblQCTables.Clear()
                        tblQCTables.AcceptChanges()
                    Catch ex As Exception

                    End Try

                    boolTableSectionStart = False
                    Call PrepareTable(intT, wd, idTR, idCRT, intRowR, strMM) 'for debugging
                    boolTableSectionStart = True

                    'Try
                    '    xlROT.ActiveWorkbook.Close(False)
                    '    xlROT.Quit()
                    'Catch ex As Exception

                    'End Try
                Catch ex As Exception

                    If intT = 4 Then
                        str1 = "There was a problem preparing table:"
                        str1 = str1 & ChrW(10) & ChrW(10) & strMM & ChrW(10) & ChrW(10) & "It is possible there is an "
                        str1 = str1 & "inconsistency in QC configuration for this study." & ChrW(10) & ChrW(10)
                        str1 = str1 & "Please activate the 'Sample/QC/Calibr Std Details' tab and inspect the QC Levels section for "
                        str1 = str1 & "QC Level assignment inconsistencies. The user may have to manually assign QC Samples."
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                    Else
                        str1 = "There was a problem preparing table:"
                        str1 = str1 & ChrW(10) & ChrW(10) & strMM
                        str1 = str1 & ChrW(10) & ChrW(10)
                        str1 = str1 & ex.Message
                    End If
                    If boolDisableWarnings Then
                    Else
                        MsgBox(str1, MsgBoxStyle.Information, "Problem preparing table...")
                    End If
                    'wdd.visible = True
                End Try

                Try
                    wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    wd.Selection.WholeStory()
                Catch ex As Exception

                End Try



                Dim intSR As Short

                'Hmmm. Probably don't have to do this
                'try excluding - will result in quicker table production
                Try
                    'Call SearchReplace(wd, "Report Body", wd.Selection.Range, False, "", intSR, 0, 0, False, False, False)

                Catch ex As Exception
                    str1 = "There was a problem completing the Search/Replace action: " & intSR & ". Report generation will continue."
                    str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                    If boolDisableWarnings Then
                    Else
                        MsgBox(str1, MsgBoxStyle.Information, "Problem Search/Replace...")
                    End If
                End Try
                'wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                Try
                    wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                Catch ex As Exception

                End Try

                'linkprevious

            End If
            'ElseIf StrComp(strSec, "Cover Page", CompareMethod.Text) = 0 Then

        Else

            If InStr(1, strLBX, "Appendices", CompareMethod.Text) > 0 Then

                'Call InsertAppendices(wd)

            ElseIf InStr(1, strLBX, "Summary Table", CompareMethod.Text) > 0 Then

                'add info to document
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                If boolExcludeCoverPage Then
                Else
                    With wd.Selection.ParagraphFormat
                        .LeftIndent = 100 'InchesToPoints(1)
                        '.SpaceBeforeAuto = False
                        '.SpaceAfterAuto = False
                    End With
                    With wd.Selection.ParagraphFormat
                        '.SpaceBeforeAuto = False
                        '.SpaceAfterAuto = False
                        .FirstLineIndent = -100 'InchesToPoints(-1)
                    End With
                    wd.Selection.Font.Bold = True
                    wd.Selection.TypeText(Text:="Study: " & vbTab & frmH.cbxStudy.Text)
                    wd.Selection.TypeParagraph()
                    'wd.Selection.TypeText(Text:="Report: " & vbTab & frmH.lblReportTitle.Text)
                    wd.Selection.TypeText(Text:="Report: " & vbTab & gReportTitle)
                    wd.Selection.TypeParagraph()
                    strDo = "Example section: " & vbTab & strLBX
                    wd.Selection.TypeText(Text:=strDo)
                    wd.Selection.TypeParagraph()
                    'enter date
                    strDo = "Date Prepared: " & vbTab & Format(dt, "MMMM dd, yyyy hh:mm:ss tt")
                    wd.Selection.TypeText(Text:=strDo)
                    wd.Selection.TypeParagraph()
                    With wd.Selection.ParagraphFormat
                        .LeftIndent = 0 'InchesToPoints(1)
                        '.SpaceBeforeAuto = False
                        '.SpaceAfterAuto = False
                    End With
                    With wd.Selection.ParagraphFormat
                        '.SpaceBeforeAuto = False
                        '.SpaceAfterAuto = False
                        .FirstLineIndent = 0 'InchesToPoints(-1)
                    End With

                    wd.Selection.TypeParagraph()
                    wd.Selection.Font.Bold = False

                    Call ExampleCoverPageBreak(wd)

                    wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                End If

                Try
                    Call SummaryTableAppendix(wd)
                Catch ex As Exception
                    str1 = "There was a problem creating the Method Summary Table report. Report generation will continue."
                    str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                    str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
                    MsgBox(str1, MsgBoxStyle.Information, "Problem creating the Method Summary Table report...")

                End Try

            Else 'report body section, others

                'open word documents
                'get ReportStatements.doc path
                Dim dtbl As System.Data.DataTable
                Dim strPath As String
                Dim strPathGuWu As String

                dtbl = tblConfiguration
                Dim drows() As DataRow

                '***Here!!

                str2 = "Report Body Sections"

                If InStr(1, strLBX, "Analytical Reference Std", CompareMethod.Text) > 0 Then
                    frmH.lblProgress.Text = ""
                    frmH.lblProgress.Refresh()

                    'add info to document
                    wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    If boolExcludeCoverPage Then
                    Else
                        With wd.Selection.ParagraphFormat
                            .LeftIndent = 100 'InchesToPoints(1)
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                        End With
                        With wd.Selection.ParagraphFormat
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                            .FirstLineIndent = -100 'InchesToPoints(-1)
                        End With
                        wd.Selection.Font.Bold = True
                        wd.Selection.TypeText(Text:="Study: " & vbTab & frmH.cbxStudy.Text)
                        wd.Selection.TypeParagraph()
                        'wd.Selection.TypeText(Text:="Report: " & vbTab & frmH.lblReportTitle.Text)
                        wd.Selection.TypeText(Text:="Report: " & vbTab & gReportTitle)
                        wd.Selection.TypeParagraph()
                        strDo = "Example section: " & vbTab & strLBX
                        wd.Selection.TypeText(Text:=strDo)
                        wd.Selection.TypeParagraph()
                        'enter date
                        strDo = "Date Prepared: " & vbTab & Format(dt, "MMMM dd, yyyy hh:mm:ss tt")
                        wd.Selection.TypeText(Text:=strDo)
                        wd.Selection.TypeParagraph()
                        With wd.Selection.ParagraphFormat
                            .LeftIndent = 0 'InchesToPoints(1)
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                        End With
                        With wd.Selection.ParagraphFormat
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                            .FirstLineIndent = 0 'InchesToPoints(-1)
                        End With

                        wd.Selection.TypeParagraph()
                        wd.Selection.Font.Bold = False

                        Call ExampleCoverPageBreak(wd)

                    End If

                    'do Analytical Ref tables
                    frmH.lblProgress.Text = "Preparing example Analytical Reference Tables..."
                    frmH.lblProgress.Refresh()
                    If boolEntireReport Then
                        'prepare dv1 from tblReportStatements
                        Dim tblR As System.Data.DataTable
                        Dim strFR As String

                        tblR = tblReportStatements
                        strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
                        dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
                        dv.RowFilter = arrRBSColumns(2, 0)

                    Else
                        dv = frmH.dgvReportStatements.DataSource
                    End If
                    'dv = frmH.dgvReportStatements.DataSource

                    int1 = FindRowDVByCol(134, dv, "ID_TBLCONFIGBODYSECTIONS")
                    'int1 = FindRowDVByCol("Analytical Reference Standard Characterization", dv, "charSectionName")
                    'select this row
                    'frmH.dgvReportStatements.CurrentCell = frmH.dgvReportStatements.Rows.Item(int1).Cells("charSectionName")
                    Try
                        boolGo = DoIndReportSections(wd, strLBX, int1, dv)
                    Catch ex As Exception

                        str1 = "There was a problem creating the " & strLBX & " report. Report generation will continue."
                        str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                        MsgBox(str1, MsgBoxStyle.Information, "Problem creating the " & strLBX & " report...")

                    End Try
                    'Note: the previous action will place anal ref tables in document

                ElseIf InStr(1, strLBX, "QA Event Table", CompareMethod.Text) > 0 Then

                    frmH.lblProgress.Text = ""
                    frmH.lblProgress.Refresh()

                    'add info to document
                    wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    If boolExcludeCoverPage Then
                    Else
                        With wd.Selection.ParagraphFormat
                            .LeftIndent = 100 'InchesToPoints(1)
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                        End With
                        With wd.Selection.ParagraphFormat
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                            .FirstLineIndent = -100 'InchesToPoints(-1)
                        End With
                        wd.Selection.Font.Bold = True
                        wd.Selection.TypeText(Text:="Study: " & vbTab & frmH.cbxStudy.Text)
                        wd.Selection.TypeParagraph()
                        'wd.Selection.TypeText(Text:="Report: " & vbTab & frmH.lblReportTitle.Text)
                        wd.Selection.TypeText(Text:="Report: " & vbTab & gReportTitle)
                        wd.Selection.TypeParagraph()
                        strDo = "Example section: " & vbTab & strLBX
                        wd.Selection.TypeText(Text:=strDo)
                        wd.Selection.TypeParagraph()
                        'enter date
                        strDo = "Date Prepared: " & vbTab & Format(dt, "MMMM dd, yyyy hh:mm:ss tt")
                        wd.Selection.TypeText(Text:=strDo)
                        wd.Selection.TypeParagraph()
                        With wd.Selection.ParagraphFormat
                            .LeftIndent = 0 'InchesToPoints(1)
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                        End With
                        With wd.Selection.ParagraphFormat
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                            .FirstLineIndent = 0 'InchesToPoints(-1)
                        End With

                        wd.Selection.TypeParagraph()

                        Call ExampleCoverPageBreak(wd)

                    End If

                    'do Pre-QA table
                    frmH.lblProgress.Text = "Preparing example QA Event Table..."
                    frmH.lblProgress.Refresh()
                    If boolEntireReport Then
                        'prepare dv1 from tblReportStatements
                        Dim tblR As System.Data.DataTable
                        Dim strFR As String

                        tblR = tblReportStatements
                        strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
                        dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
                        dv.RowFilter = arrRBSColumns(2, 0)

                    Else
                        dv = frmH.dgvReportStatements.DataSource
                    End If
                    'dv = frmH.dgvReportStatements.DataSource
                    int1 = FindRowDVByCol(2, dv, "ID_TBLCONFIGBODYSECTIONS")

                    'dv = frmh.dgvReportStatements.DataSource
                    'int1 = FindRowDVByCol("Pre-QA Table Statements", dv, "charSectionName")
                    'select this row
                    'frmH.dgvReportStatements.CurrentCell = frmH.dgvReportStatements.Rows.Item(int1).Cells("charSectionName")
                    boolGo = DoIndReportSections(wd, strLBX, int1, dv)
                    'Try
                    '    boolGo = DoIndReportSections(wd, str2, int1, dv)
                    'Catch ex As Exception
                    '    str1 = "There was a problem creating the " & strLBX & " report. Report generation will continue."
                    '    str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                    '    MsgBox(str1, MsgBoxStyle.Information, "Problem creating the " & strLBX & " report...")
                    'End Try

                ElseIf InStr(1, strLBX, "Contributing Personnel", CompareMethod.Text) > 0 Then

                    frmH.lblProgress.Text = "Preparing example " & strSec & "..."
                    frmH.lblProgress.Refresh()

                    'add info to document
                    wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    If boolExcludeCoverPage Then
                    Else
                        With wd.Selection.ParagraphFormat
                            .LeftIndent = 100 'InchesToPoints(1)
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                        End With
                        With wd.Selection.ParagraphFormat
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                            .FirstLineIndent = -100 'InchesToPoints(-1)
                        End With
                        wd.Selection.Font.Bold = True
                        wd.Selection.TypeText(Text:="Study: " & vbTab & frmH.cbxStudy.Text)
                        wd.Selection.TypeParagraph()
                        'wd.Selection.TypeText(Text:="Report: " & vbTab & frmH.lblReportTitle.Text)
                        wd.Selection.TypeText(Text:="Report: " & vbTab & gReportTitle)
                        wd.Selection.TypeParagraph()
                        strDo = "Example section: " & vbTab & strLBX
                        wd.Selection.TypeText(Text:=strDo)
                        wd.Selection.TypeParagraph()
                        'enter date
                        strDo = "Date Prepared: " & vbTab & Format(dt, "MMMM dd, yyyy hh:mm:ss tt")
                        wd.Selection.TypeText(Text:=strDo)
                        wd.Selection.TypeParagraph()
                        With wd.Selection.ParagraphFormat
                            .LeftIndent = 0 'InchesToPoints(1)
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                        End With
                        With wd.Selection.ParagraphFormat
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                            .FirstLineIndent = 0 'InchesToPoints(-1)
                        End With

                        wd.Selection.TypeParagraph()

                        Call ExampleCoverPageBreak(wd)
                    End If

                    If boolEntireReport Then
                        'prepare dv1 from tblReportStatements
                        Dim tblR As System.Data.DataTable
                        Dim strFR As String

                        tblR = tblReportStatements
                        strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
                        dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
                        dv.RowFilter = arrRBSColumns(2, 0)

                    Else
                        dv = frmH.dgvReportStatements.DataSource
                    End If
                    'dv = frmH.dgvReportStatements.DataSource
                    'int1 = FindRowDVByCol(2, dv, "ID_TBLCONFIGBODYSECTIONS")

                    boolGo = DoIndReportSections(wd, strLBX, 1, dv)


                Else 'report body section

                    frmH.lblProgress.Text = "Preparing example " & strSec & "..."
                    frmH.lblProgress.Refresh()

                    If boolEntireReport Then
                        'prepare dv1 from tblReportStatements
                        Dim tblR As System.Data.DataTable
                        Dim strFR As String

                        tblR = tblReportStatements
                        strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
                        dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
                        dv.RowFilter = arrRBSColumns(2, 0)

                    Else
                        dv = frmH.dgvReportStatements.DataSource
                    End If
                    'dv = frmH.dgvReportStatements.DataSource
                    intRow = frmH.dgvReportStatements.CurrentRow.Index
                    idRT = dv(intRow).Item("ID_TBLWORDSTATEMENTS")

                    If StrComp(strSec, "Cover Page", CompareMethod.Text) = 0 Then
                    Else
                        'add info to document
                        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                        If boolExcludeCoverPage Then
                        Else
                            With wd.Selection.ParagraphFormat
                                .LeftIndent = 100 'InchesToPoints(1)
                                '.SpaceBeforeAuto = False
                                '.SpaceAfterAuto = False
                            End With
                            With wd.Selection.ParagraphFormat
                                '.SpaceBeforeAuto = False
                                '.SpaceAfterAuto = False
                                .FirstLineIndent = -100 'InchesToPoints(-1)
                            End With
                            wd.Selection.Font.Bold = True
                            wd.Selection.TypeText(Text:="Study: " & vbTab & frmH.cbxStudy.Text)
                            wd.Selection.TypeParagraph()
                            'wd.Selection.TypeText(Text:="Report: " & vbTab & frmH.lblReportTitle.Text)
                            wd.Selection.TypeText(Text:="Report: " & vbTab & gReportTitle)
                            wd.Selection.TypeParagraph()
                            strDo = "Example section: " & vbTab & str2 & ": " & strSec
                            wd.Selection.TypeText(Text:=strDo)
                            wd.Selection.TypeParagraph()
                            'enter date
                            strDo = "Date Prepared: " & vbTab & Format(dt, "MMMM dd, yyyy hh:mm:ss tt")
                            wd.Selection.TypeText(Text:=strDo)
                            wd.Selection.TypeParagraph()
                            With wd.Selection.ParagraphFormat
                                .LeftIndent = 0 'InchesToPoints(1)
                                '.SpaceBeforeAuto = False
                                '.SpaceAfterAuto = False
                            End With
                            With wd.Selection.ParagraphFormat
                                '.SpaceBeforeAuto = False
                                '.SpaceAfterAuto = False
                                .FirstLineIndent = 0 'InchesToPoints(-1)
                            End With

                            wd.Selection.TypeParagraph()

                            Call ExampleCoverPageBreak(wd)
                        End If


                    End If

                    Try
                        boolGo = DoIndReportSections(wd, strLBX, intRow, dv)
                    Catch ex As Exception
                        str1 = "There was a problem creating the " & strLBX & " report. Report generation will continue."
                        str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                        MsgBox(str1, MsgBoxStyle.Information, "Problem creating the " & strLBX & " report...")
                    End Try


                End If


                'frmH.lblProgress.Text = "Processing Field Codes..."
                'frmH.lblProgress.Refresh()

                lblP.Text = "Processing Field Codes..."
                lblP.Refresh()

                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                Dim intSR As Short
                intSR = 1


                Try
                    Call SearchReplace(wd, "Report Body", wd.Selection.Range, False, "", intSR, 0, 0, False, False, False, 0)

                Catch ex As Exception
                    str1 = "There was a problem completing the Search/Replace action: " & intSR & ". Report generation will continue."
                    str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."

                End Try

                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                wd.Selection.WholeStory()

                Try
                    'wdd.visible = False
                    Call SignatureSearch(wd)

                Catch ex As Exception
                    str1 = "There was a problem completing the Signature Search/Replace action. Report generation will continue."
                    str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."

                End Try


                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                wd.Selection.WholeStory()
                'wdd.visible = True
                'frmH.Activate()

                If boolDoFormulas Then
                    'frmH.lblProgress.Text = "Formatting chemical formulas..."
                    'frmH.lblProgress.Refresh()

                    lblP.Text = "Formatting chemical formulas..."
                    lblP.Refresh()
                    Try
                        Call ChemFormula(wd.Selection.Range, wd) 'to address any sub/superscripts

                    Catch ex As Exception
                        str1 = "There was a problem completing the Chemical Formula Signature Search/Replace action. Report generation will continue."
                        str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."

                    End Try
                End If
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                wd.Selection.WholeStory()

                'frmH.lblProgress.Text = "Replacing hyphens..."
                'frmH.lblProgress.Refresh()

                lblP.Text = "Replacing hyphens..."
                lblP.Refresh()

                'Call ReplaceDegC(wd, wd.Selection.Range)
                'Call ReplaceHyphens(wd, wd.Selection.Range)
                Try
                    Call ReplaceDegC(wd, wd.Selection.Range)
                Catch ex As Exception
                    'wdd.visible = True
                    str1 = "Problem with converting deg C."
                    MsgBox(str1, MsgBoxStyle.Information, "deg C problem...")
                End Try

                Try
                    Call ReplaceHyphens(wd, wd.Selection.Range)
                Catch ex As Exception
                    'wdd.visible = True
                    str1 = "Problem with converting hyphens to nbh."
                    MsgBox(str1, MsgBoxStyle.Information, "Hyphen problem...")
                End Try

                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)


                wd.ActiveWindow.View.ShowFieldCodes = False

                Try
                    'Call PrintPDF(wd)
                Catch ex As Exception
                    str1 = "The Adobe print driver must be installed in order to present this report as PDF." & ChrW(10) & ChrW(10)
                    str1 = str1 & "The report will be presented as Microsoft" & ChrW(174) & "."
                    MsgBox(str1, MsgBoxStyle.Information, "Adobe print driver not installed...")
                End Try

                Call PositionAbort(False)

                '20180701 LEE
                'Implement time saving trick
                Call SpellingOff(wd.ActiveDocument, True)

                'frmH.pb1.Visible = False
                frmH.Refresh()
end1:
                If boolWD1 Then

                    boolGo = True
                Else
                    Try
                        wd.Application.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)

                    Catch ex As Exception

                    End Try


                    Try
                        wd.Application.ShowWindowsInTaskbar = boolSTB
                    Catch ex As Exception

                    End Try

                    Try
                        wd.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)

                    Catch ex As Exception

                    End Try
                    wd = Nothing
                    boolGo = False
                End If

                If boolWD2 Then
                    boolGo = True
                Else
                    boolGo = False
                End If

            End If
        End If

end2:
        If boolGo Then

            'wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            Try
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            Catch ex As Exception

            End Try

            Dim strDate As String
            Dim strTime As Date
            strDate = Format(dt, "MMM dd, yyyy")
            strTime = Format(dt, "hh:mm:ss tt")
            str1 = "DRAFT" & ChrW(10) & strDate & ChrW(10) & strTime

            Try
                Call InsertWatermark(wd, boolIncludeWaterMark, str1)
            Catch ex As Exception
                str1 = "Unfortunately, this version of Microsoft" & ChrW(174) & " Word does not contain the Word watermarking funtion supported by StudyDoc." & ChrW(10) & ChrW(10)
                str1 = str1 & "Word 2002 or higher must be used. " & ChrW(10) & ChrW(10)
                str1 = str1 & "The report will be prepared without a watermark."
                MsgBox(str1, MsgBoxStyle.Information, "Watermark not supported...")
            End Try

            'wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            Try
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            Catch ex As Exception

            End Try

            'look for section break as first character
            Try
                Call EvalCoverPage(wd)
            Catch ex As Exception

            End Try

            Try
                '20180309 LEE:
                'Note: Leave header in. If table is to mimic that of report, then header needs to be close to identical
                '20180327 LEE: No, need to evaluate boolExcludeHeaderFooter
                If boolIncludeWaterMark Then
                Else
                    If boolExcludeHeaderFooter Then
                        Call ClearHeaderFooter(wd)
                    End If
                End If
            Catch ex As Exception

            End Try

            'now set margins for individual sections

            System.Windows.Forms.Application.DoEvents()

            Try
                Call LinkPrevious(wd, False, True)

            Catch ex As Exception

            End Try

            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            Cursor.Current = Cursors.Default

            'frmH.lblProgress.Visible = False
            Call PositionAbort(False)

            'frmH.pb1.Visible = False
            frmH.Refresh()

            str1 = "Example Section Completed."
            'MsgBox(str1, MsgBoxStyle.Information, "Action completed...")
            Try

                'clear formatting again
                Try
                    wd.Selection.Find.ClearFormatting()
                Catch ex As Exception

                End Try
                Try
                    wd.Selection.Range.Find.ClearFormatting()
                Catch ex As Exception

                End Try

                Try
                    wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                Catch ex As Exception

                End Try

                'save as temp then display in afr
                Dim strP As String
                strP = GetNewTempFile(True)
                'strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)

                If boolHasMacro And boolSaveAsDocx = False Then
                    strP = Replace(strP, ".xml", ".docm", 1, -1, CompareMethod.Text)
                Else
                    strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)
                End If

                'frmH.lblProgress.Text = "Saving DOC:" & ChrW(10) & ChrW(10) & strP
                'frmH.lblProgress.Refresh()

                lblP.Text = "Saving document..." ' & ChrW(10) & ChrW(10) & strP
                lblP.Refresh()

                If boolHasMacro And boolSaveAsDocx = False Then
                    wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
                Else
                    wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                End If
                'wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)

                If gDoPDF Then
                    If wd.Version < 12 Then
                        gDoPDF = False
                    Else
                        gDoPDF = True
                    End If
                End If

                If gDoPDF Then
                    Call CreatePDF(wd, strP)
                End If

                str1 = "Example Section - " & strReportTypeApp
                Call ReportHistoryItem(str1)

                Try
                    wd.ActiveDocument.Close(False)
                Catch ex As Exception

                End Try

                Try
                    wd.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)

                Catch ex As Exception

                End Try

                Threading.Thread.Sleep(250)

                If gDoPDF Then
                Else
                    Call OpenAFR(strP, "", False, boolSTB, True, False)

                End If


                'wdd.visible = True
                'wd.Application.Activate()

            Catch ex As Exception

                var1 = ex.Message
                var1 = var1

            End Try
        Else
            'wd.Quit()
        End If
        wd = Nothing

        Cursor.Current = Cursors.Default

        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()

        Call PositionAbort(False)


        frmH.Refresh()

        boolShowExample = False

        'str1 = "Example Section - " & strReportTypeApp
        'Call ReportHistoryItem(str1)


        If gboolER Then
        Else
            frmWordStatement.Activate()
        End If

        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False
        'frmH.pb2.Visible = False

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()

        frmH.Refresh()

        Try
            rowsDGV(0).RejectChanges()
        Catch ex As Exception

        End Try

        Exit Sub

end3:
        Cursor.Current = Cursors.Default

        Try
            'frmH.lblProgress.Visible = False
            Call PositionAbort(False)

            'frmH.pb1.Visible = False

            'frmH.panProgress.Visible = False
            'frmH.panProgress.Refresh()

            frmH.Refresh()
        Catch ex As Exception

        End Try


    End Sub


    Sub OpenAFR(ByVal strP As String, ByVal strLbl As String, ByVal boolFromReportTemplate As Boolean, boolSTB As Boolean, boolTested As Boolean, boolRO As Boolean)

        'boolFromReportTemplate:
        '  from Report Template: True
        '  from Generate report:  False
        'boolRO = read only

        Dim strE As String = "No E"
        Dim str1 As String
        Dim var1

        If gboolER And boolFromReportTemplate = False Then
            Try

                strE = "Before Dim frm As New frmDocumentCompare"
                Dim frm As New frmDocumentCompare
                strE = "After Dim frm As New frmDocumentCompare"

                frm.boolTemplate = False
                frm.gDoc = "Final Report"
                frm.gReport = strP
                frm.strPrevForm = "Home"
                'frm.gbSaveType.Visible = True
                frm.Text = "Document Control"
                Cursor.Current = Cursors.Default

                If BOOLFORCEFINALREPORTPDF Then
                    'str1 = "User permissions are set to force document as PDF"
                    'str1 = str1 & ChrW(10) & ChrW(10) & "You will be directed to a PDF window."
                    'MsgBox(str1, MsgBoxStyle.Information, "Opening in PDF...")
                    frm.Show(frmH)
                    Call frm.DoPDF()
                    frm.Dispose()
                    'Call frm.DoPDF()
                Else
                    frm.Show(frmH)
                    'frmH.Visible = False

                End If



            Catch ex As Exception

            End Try

        Else

            Try

                strE = "Before Dim frm As New frmWordStatement"
                Dim frm As New frmWordStatement
                strE = "After Dim frm As New frmWordStatement"
                Dim id As Int64

                'MsgBox("frm as new just finished")

                frm.boolReport = boolFromReportTemplate
                frm.strReport = strP
                frm.boolSTB = boolSTB
                frm.boolTested = boolTested
                frm.boolEdit = False
                Call frm.PlaceControls()

                '****
                Dim dgv As DataGridView
                Dim intRow As Short

                dgv = frmH.dgvReportStatementWord

                intRow = 0
                If dgv.CurrentRow Is Nothing Then
                    intRow = 0
                Else
                    intRow = dgv.CurrentRow.Index
                End If

                id = dgv("ID_TBLWORDSTATEMENTS", intRow).Value
                frm.id = id

                '****

                frm.pan2.Visible = True
                frm.pan2.BringToFront()
                If boolFromReportTemplate Then 'show pan
                    'frm.panEditReports.Visible = False
                    frm.panList.Visible = True
                    'frm.panSave.Visible = False
                Else
                    'frm.panEditReports.Visible = False
                    frm.panList.Visible = False
                    'frm.panSave.Visible = False
                End If
                'frm.panEditReports.Visible = False
                'frm.panList.Visible = False
                frm.panSave.Visible = False

                frm.lblBlink.Visible = False
                frm.tSave.Enabled = False

                'frm.pan1a.Visible = False
                'frm.panEdit.Visible = False

                'MsgBox("almost ready to load")

                If boolFromReportTemplate Then
                    strE = "Before Call frm.FormLoad()"
                    Call frm.FormLoad()
                    strE = "After Call frm.FormLoad()"
                End If
                'Call frm.FormLoad()

                'frm.lblEdit.Visible = True

                Pause(0.25)

                frm.Text = "StudyDoc Microsoft" & ChrW(174) & " Word Document"

                strE = "Before Call frm.DoReadOnly()"
                Call frm.DoReadOnly()
                strE = "After Call frm.DoReadOnly()"

                frm.boolReadOnly = boolRO

                If boolFromReportTemplate Then
                Else
                    frm.lblReadOnly.Visible = False
                End If


                'If BOOLFORCEFINALREPORTPDF And boolFromReportTemplate = False Then
                '    str1 = "User permissions are set to force document as PDF"
                '    str1 = str1 & ChrW(10) & ChrW(10) & "You will be directed to a PDF window."
                '    MsgBox(str1, MsgBoxStyle.Information, "Opening in PDF...")

                '    GoTo end1
                'End If

                '20160713 LEE: gGoToWord is deprecated. Default = false
                If gGoToWord Then

                    str1 = "StudyDoc has been configured to open the generated report directly in Word" & ChrW(8482) & "."
                    str1 = str1 & ChrW(10) & ChrW(10) & "You will be directed to a Microsoft Word window."
                    MsgBox(str1, MsgBoxStyle.Information, "Opening in Word...")

                    Call frm.GoToWord(False)
                Else


                    If BOOLFORCEFINALREPORTPDF And boolFromReportTemplate = False Then
                        'str1 = "User permissions are set to force document as PDF"
                        'str1 = str1 & ChrW(10) & ChrW(10) & "You will be directed to a PDF window."
                        'MsgBox(str1, MsgBoxStyle.Information, "Opening in PDF...")
                        frm.Show(frmH)
                        Call frm.DoPDF()
                        frm.Dispose()
                        'Call frm.DoPDF()
                    Else
                        frm.Show(frmH)
                        frmH.Visible = False
                        Try
                            frm.Activate()
                        Catch ex As Exception
                            var1 = ex.Message
                        End Try
                    End If

                    'frm.Show()
                    'Call frm.DoPDF()
                    'frm.Dispose()
                End If

                'frm.Show()

                '20160713 LEE: gGoToWord is deprecated. Default is false
                'If gGoToWord Then

                '    Call frm.GoToWord()

                'End If
            Catch ex As Exception

                'MsgBox(ex.InnerException.ToString & ":  " & ex.Message)
                MsgBox("OpenAFR:" & ChrW(10) & strE & ChrW(10) & ex.Message, vbInformation, "Problem...")

            End Try
        End If

end1:

    End Sub

    Sub FindExistingGuWu(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal docName As String)

        'docName is full path
        'http://www.andreavb.com/tip110001.html

        Dim MyDoc As Microsoft.Office.Interop.Word.Document
        Dim Opened As Boolean

        Opened = False
        For Each MyDoc In wd.Documents
            If MyDoc.FullName = docName Then Opened = True
        Next
        If Opened Then GoTo AlreadyOpen
AlreadyOpen:

    End Sub

    Sub InsertWatermark(ByRef wd As Microsoft.Office.Interop.Word.Application, boolIncludeWaterMark As Boolean, strM As String)


        If boolIncludeWaterMark Then
        Else
            Exit Sub
        End If

        '
        ' Macro1 Macro
        ' Macro recorded 9/7/2006 by Gubbs
        '

        Dim str1 As String
        Dim sec As Microsoft.Office.Interop.Word.Section
        Dim ct1 As Short
        Dim ct2 As Short
        Dim pb1VO As Short
        Dim pb1MO As Short
        Dim intShpCt As Integer

        Dim var1, var2
        Dim intSecs As Int16 = 0
        Dim Count1 As Int16

        Dim numFontSize As Single

        '20190214 LEE:
        Dim boolPortrait As Boolean = True
        Dim boolP1 As Boolean = True
        Dim boolP2 As Boolean = True
        Dim boolHit As Boolean = False

        Dim strDate As String
        Dim strTime As Date
        Dim dt As Date
        dt = Now
        strDate = Format(dt, "MMM dd, yyyy")
        strTime = Format(dt, "hh:mm:ss tt")

        str1 = strM ' "DRAFT" & Chr(10) & strDate & Chr(10) & strTime
        'str1 = "DRAFT"

        'wdd.visible = True

        With wd

            ct2 = .ActiveDocument.Sections.Count
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            Dim wid, ht, lm, rm, tm, bm ', wd1, ht1
            Dim wd1 As Single
            Dim ht1 As Single
            With .ActiveDocument.Sections.Item(1).PageSetup
                wid = .PageWidth
                ht = .PageHeight
                tm = .TopMargin
                bm = .BottomMargin
                lm = .LeftMargin
                rm = .RightMargin

                wd1 = wid - rm - lm
                ht1 = ht - tm - bm

                '20190214 LEE:
                'adjust size a bit
                wd1 = wd1 * 0.8
                ht1 = ht1 * 0.8

            End With

            pb1VO = frmH.pb1.Value
            pb1MO = frmH.pb1.Maximum

            Try


                frmH.pb1.Value = 0
                frmH.pb1.Maximum = ct2
                ct1 = 0
                .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter

                boolP1 = True
                boolP2 = True

                intSecs = .ActiveDocument.Sections.Count

                For Each sec In .ActiveDocument.Sections

                    If sec.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                        boolPortrait = False
                    Else
                        boolPortrait = True
                    End If

                    boolP1 = boolPortrait

                    numFontSize = sec.Range.Font.Size

                    ct1 = ct1 + 1
                    frmH.lblProgress.Text = "Inserting watermark " & ct1 & " of " & ct2 & "..."
                    frmH.lblProgress.Refresh()
                    frmH.pb1.Value = ct1
                    frmH.pb1.Refresh()
                    'wdd.visible = True
                    'sec.Range.Select()
                    '.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader
                    '.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter
                    If ct1 = 1 Then

                        intShpCt = 1
                        'wd.ActiveDocument.Sections(1).Range.Select()
                        'sec.Range.Select()
                        .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter
                        'goto end
                        .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                        '.Selection.TypeParagraph()

                        Call DoWatermark(wd, ct1, str1, ht1, wd1, sec)


                    Else

                        If boolPortrait = boolP2 Then
                        Else
                            'need to evaluate link to previous
                            If wd.ActiveDocument.Sections(ct1).Footers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious Then
                                'unlink
                                'Selection.HeaderFooter.LinkToPrevious = Not Selection.HeaderFooter.LinkToPrevious
                                wd.ActiveDocument.Sections(ct1).Footers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False
                            End If

                            var1 = wd.ActiveDocument.Sections.Count 'debug
                            var1 = var1

                            Try
                                .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToSection, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToAbsolute, Count:=ct1)
                            Catch ex As Exception
                                var1 = var1
                            End Try

                            'now go into footer
                            .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter


                            'goto end
                            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                            intShpCt = intShpCt + 1

                            If boolHit = False Then
                                boolHit = True 'the first time linktoprevious = false is done, there is already a watermark present
                            Else
                                Call DoWatermark(wd, ct1, str1, ht1, wd1, sec)
                            End If



                            'Dim int1 As Short
                            'Try
                            '    'int1 = sec.Footers(ct1).Shapes.Count '20190214 LEE: Hmmm. This returns total shapes in all sections
                            '    .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter
                            '    int1 = .Selection.HeaderFooter.Shapes.Count
                            'Catch ex As Exception
                            '    var1 = var1
                            'End Try


                            'If intShpCt > int1 Then
                            '    With wd.ActiveDocument.Sections(ct1)
                            '        '20190214 LEE:
                            '        'Hmmm. Word 2013: Paste no longer remembers size
                            '        .Footers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Paste()
                            '    End With
                            'End If


                            'Try

                            '    .Selection.HeaderFooter.Shapes(intShpCt).Select() '20190214 LEE
                            'Catch ex As Exception
                            '    var1 = var1
                            'End Try



                            'If sec.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                            '    .Selection.ShapeRange.Height = wd1 'InchesToPoints(2.82)
                            '    .Selection.ShapeRange.Width = ht1 'InchesToPoints(7.05)
                            'Else
                            '    .Selection.ShapeRange.Height = ht1 'InchesToPoints(2.82)
                            '    .Selection.ShapeRange.Width = wd1 'InchesToPoints(7.05)
                            'End If
                            ''If wd.ActiveDocument.Sections(ct1).PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                            ''    'wd.ActiveDocument.Shapes(intShpCt).Height = wd1
                            ''    'wd.ActiveDocument.Shapes(intShpCt).Width = ht1
                            ''Else
                            ''    '.Selection.ShapeRange.Height = ht1 'InchesToPoints(2.82)
                            ''    '.Selection.ShapeRange.Width = wd1 'InchesToPoints(7.05)
                            ''End If

                        End If
                    End If



                    boolP2 = boolPortrait

                Next
            Catch ex As Exception

                var1 = ex.Message
                var1 = var1

            End Try

            '20190214 LEE: clear clipboard
            Try
                Clipboard.Clear()
            Catch ex As Exception
                var1 = var1
            End Try


            '20190214 LEE: Why am I going back to the document earlier?
            'do it here
            wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

            '.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)


        End With

        frmH.pb1.Value = 0
        frmH.pb1.Maximum = pb1MO
        frmH.pb1.Value = pb1VO
        frmH.pb1.Refresh()

        Try
            Clipboard.Clear()
        Catch ex As Exception

        End Try

    End Sub

    Sub DoWatermark(ByRef wd As Microsoft.Office.Interop.Word.Application, ct1 As Short, str1 As String, ht1 As Single, wd1 As Single, sec As Microsoft.Office.Interop.Word.Section)

        '20190210 LEE:
        'It was found by Frontage that Watermark feature has stopped working in Office 2016. It was verified that this was also true for Office 2013.
        'The old code copy/pasted the original watermark - in order to increase performance. In Office 2010, the size of the paste was identical to the size of copy.
        'In Office 2013 and above, paste does not retain size. It has to be resized to that of the copy.
        'It was decided to make each watermark individually because it is hard to select the paste shape.

        Dim var1

        With wd

            .Selection.HeaderFooter.Shapes.AddTextEffect(Office.Core.MsoPresetTextEffect.msoTextEffect1, str1, "Times New Roman", 14, False, False, 0, 0).Select()

            .Selection.ShapeRange.Name = "Watermark" & ct1
            .Selection.ShapeRange.TextEffect.NormalizedHeight = False
            .Selection.ShapeRange.Line.Visible = False
            .Selection.ShapeRange.Fill.Visible = True
            .Selection.ShapeRange.Fill.Solid()
            .Selection.ShapeRange.Fill.ForeColor.RGB = RGB(192, 192, 192)
            '.Selection.ShapeRange.Fill.Transparency = 0.75
            '.Selection.ShapeRange.Rotation = 315
            '.EqualsSelection.ShapeRange.LockAspectRatio = True

            If sec.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                .Selection.ShapeRange.Height = wd1 'InchesToPoints(2.82)
                .Selection.ShapeRange.Width = ht1 'InchesToPoints(7.05)
            Else
                .Selection.ShapeRange.Height = ht1 'InchesToPoints(2.82)
                .Selection.ShapeRange.Width = wd1 'InchesToPoints(7.05)
            End If

            .Selection.ShapeRange.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
            .Selection.ShapeRange.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin
            .Selection.ShapeRange.WrapFormat.AllowOverlap = False
            .Selection.ShapeRange.WrapFormat.Side = Microsoft.Office.Interop.Word.WdWrapType.wdWrapNone ' wdWrapLargest
            .Selection.ShapeRange.WrapFormat.Type = 3 'Word.WdWrapType.wdWrapNone
            .Selection.ShapeRange.Left = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter ' wdShapeCenter
            .Selection.ShapeRange.Top = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter ' wdShapeCenter


            .Selection.ShapeRange.ZOrder(5)

            '20190210 LEE:
            'hmmm. Need to size again
            If sec.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                .Selection.ShapeRange.Height = wd1 'InchesToPoints(2.82)
                .Selection.ShapeRange.Width = ht1 'InchesToPoints(7.05)
            Else
                .Selection.ShapeRange.Height = ht1 'InchesToPoints(2.82)
                .Selection.ShapeRange.Width = wd1 'InchesToPoints(7.05)
            End If

            'must do this again
            .Selection.ShapeRange.Left = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter ' wdShapeCenter
            .Selection.ShapeRange.Top = Microsoft.Office.Interop.Word.WdShapePosition.wdShapeCenter ' wdShapeCenter
            .Selection.ShapeRange.RelativeHorizontalPosition = Microsoft.Office.Interop.Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
            .Selection.ShapeRange.RelativeVerticalPosition = Microsoft.Office.Interop.Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin

            'now must unselect graphic
            .Selection.Collapse()

        End With


    End Sub


    Sub MakePDF(ByVal wd)

        wd.Application.Run(MacroName:="AdobePDFMaker.AutoExec.Main")
        wd.Application.Run(MacroName:="AdobePDFMaker.AutoExec.ConvertToPDF")

    End Sub

    Sub PrintPDF(ByVal wd)

        With wd

            Exit Sub

            If Len(GetPDFDriver) = 0 Then
            Else

                .ActivePrinter = GetPDFDriver()

                .Application.PrintOut(FileName:="", Range:=Microsoft.Office.Interop.Word.WdPrintOutRange.wdPrintAllDocument, Item:= _
                    Microsoft.Office.Interop.Word.WdPrintOutItem.wdPrintDocumentContent, Copies:=1, Pages:="", PageType:=Microsoft.Office.Interop.Word.WdPrintOutPages.wdPrintAllPages)
            End If

        End With


    End Sub


    Sub ExampleTablesSection()

        Dim var1, var2, var3
        Dim int1 As Short
        Dim int2 As Short
        Dim dv As System.Data.DataView
        Dim drow As DataRow
        Dim tbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim strSQL As String
        Dim dbPath As String
        Dim BACStudy As String
        Dim strTemplate As String
        'Dim rng1 As Range
        Dim Count1 As Short
        Dim Count2 As Short
        Dim fi
        Dim ctCols As Short
        'Dim frm As New frmHome_01
        'Dim ''frmp As New ''frmprogress_01
        Dim intPMax As Short
        Dim intPCt As Short
        Dim strDo As String
        'Dim wdGuWu As Object
        Dim wdSt As Object
        Dim pos1, pos2
        Dim wrdselection As Microsoft.Office.Interop.Word.Selection

        'check for permission
        Dim strF As String
        Dim bool As Boolean
        Dim boolA As Short
        Dim boolTemp As Boolean
        Dim drows() As DataRow
        Dim dt As Date

        Dim intRow As Short

        ReDim ctrsSamples(5, ctAnalytes)
        ReDim ctrsRepeat(5, ctAnalytes)
        ReDim ctrsReassayed(5, ctAnalytes)
        ReDim ctrsISR(5, ctAnalytes)

        boolExcludeHeaderFooter = True

        'Dim wd As Object

        dt = Now
        LTableDateTimeStamp = dt

        ctAppendix = 0
        ctFigures = 0
        tblAppendix.Clear()
        tblFigures.Clear()
        tblAttachments.Clear()
        ctAttachments = 0

        'first check for configured reports
        Dim dvR As System.Data.DataView
        dvR = frmH.dgvReports.DataSource
        If dvR.Count = 0 Then
            str1 = "A report must be configured on the Home Tab in order to use this function."
            MsgBox(str1, MsgBoxStyle.Information, "Please configure a report...")
            Exit Sub
        End If

        If frmH.dgvReports.CurrentRow Is Nothing Then
            int1 = 0
        Else
            int1 = frmH.dgvReports.CurrentRow.Index
        End If

        boolShowExample = True
        ctPB = 0 'pesky
        intILS = 0
        intStartTable = 2

        If id_tblStudies = 0 Then
            MsgBox("A study must be chosen.", MsgBoxStyle.Information, "A study must be chosen...")
            GoTo end1
        End If

        Dim frm As New frmMsgBox
        frm.lblText.Text = "Do you wish to generate an Example Report Table Section?"
        frm.lblReportTemplate.Visible = True
        frm.cbxReportTemplate.Visible = True
        frm.boolTable = False

        frm.ShowDialog()
        If frm.boolCancel Then
            int1 = 7
            frm.Close()
            frm.Dispose()
            frmH.Refresh()
            GoTo end3
        Else
            int1 = 6
        End If
        strReportTemplateChoice = frm.cbxReportTemplate.Text

        frm.Close()
        frm.Dispose()
        frmH.Refresh()

        'int1 = MsgBox("Do you wish to generate an Example Report Table Section?", MsgBoxStyle.YesNo, "Generate Example Report Table Section...")
        'frmH.Refresh()
        If int1 = 7 Then
            GoTo end3
        End If

        strPathWd = ""
        strPathWd = GetNewTempFile(True) 'DON'T SKIP THIS!!! Paste routine uses this

        'must make table section 'show included'
        If frmH.rbShowIncludedRTConfig.Checked Then
        Else
            frmH.rbShowIncludedRTConfig.Checked = True
            frmH.Refresh()
        End If

        Cursor.Current = Cursors.WaitCursor

        'find Figures, Tables and Appendices
        Dim dv1 As System.Data.DataView
        Dim intCt As Short
        Dim arr1(3, 3)
        '1=dgv row number, 2=header text, 3=ID_TBLCONFIGBODYSECTIONS
        Dim boolGo As Boolean
        Dim wdobj As Object 'for legacy purposes

        intCt = 0
        If boolEntireReport Then
            'prepare dv1 from tblReportStatements
            Dim tblR As System.Data.DataTable
            Dim strFR As String

            tblR = tblReportStatements
            strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
            dv1 = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
            dv1.RowFilter = arrRBSColumns(2, 0)

        Else
            dv1 = frmH.dgvReportStatements.DataSource
        End If

        For Count2 = 2 To 2
            Select Case Count2
                Case 1
                    int1 = 139 'figures
                Case 2
                    int1 = 140 'tables
                Case 3
                    int1 = 141 'appendices
            End Select
            int2 = FindRowDVByCol(CStr(int1), dv1, "ID_TBLCONFIGBODYSECTIONS")
            If int2 = -1 Then
            Else
                intCt = intCt + 1
                arr1(1, intCt) = int2 'introw
                arr1(2, intCt) = NZ(dv1(int2).Item("CHARHEADINGTEXT"), "[NO CAPTION]")
                arr1(3, intCt) = int1
            End If
        Next

        If intCt = 0 Then
            MsgBox("Hmmm. Tables, Figures, and Appendices sections have not been configured in the Report Body Section. Please investigate.", MsgBoxStyle.Information, "Nothing to generate")
            GoTo end1
        End If

        'sort asc arr1
        For Count1 = 1 To intCt - 1
            int1 = arr1(1, Count1)
            For Count2 = Count1 + 1 To intCt
                int2 = arr1(1, Count2)
                If int2 < int1 Then
                    var1 = arr1(1, Count1)
                    var2 = arr1(2, Count1)
                    var3 = arr1(3, Count1)

                    arr1(1, Count1) = arr1(1, Count2)
                    arr1(2, Count1) = arr1(2, Count2)
                    arr1(3, Count1) = arr1(3, Count2)

                    arr1(1, Count2) = var1
                    arr1(2, Count2) = var2
                    arr1(3, Count2) = var3

                    int1 = arr1(1, Count1)
                End If
            Next
        Next


        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim boolSTB As Boolean = True
        Try
            boolSTB = wd.Application.ShowWindowsInTaskbar
        Catch ex As Exception

        End Try

        'Dim wdguwu As New Microsoft.Office.Interop.Word.Application
        Dim intCBS As Int64

        tbl = tblConfiguration

        ReDim garrMargins(1, 4) 'T,L,B,R for Portrait:  1=P, 2=L

        intTCur = 0
        '20180514 LEE:
        intTTot = getIntTTot(False)
        'intTTot = getIntTTot(True)

        Try

            strPathWd = GetNewTempFile(True)

            intCBS = frmH.dgvReportStatements("ID_TBLWORDSTATEMENTS", 0).Value
            'find new intcbs
            Dim intCBSa As Int64
            intCBSa = GetNewCBS()
            If intCBSa = 0 Then
            Else
                intCBS = intCBSa
            End If

            strTemplate = OpenTemplate(intCBS, strPathWd)
            wd.Documents.Open(FileName:=strTemplate, ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", WritePasswordTemplate:="", Format:=Microsoft.Office.Interop.Word.WdOpenFormat.wdOpenFormatAuto, XMLTransform:="")

            If wd.ActiveDocument.HasVBProject Then
                boolHasMacro = True
            Else
                boolHasMacro = False
            End If

            If boolHasMacro Then
                boolSaveAsDocx = SaveAsDocx(wd)
            Else
                boolSaveAsDocx = True
            End If

            wd.Selection.WholeStory()
            wd.Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

            '20151218 LEE: modification to address the no-page-break for first table logic added earlier
            'add a paragraph return and a sectionbreak
            wd.Selection.TypeParagraph()
            wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

            '20180524 LEE:
            'now save
            wd.ActiveDocument.Save()


            'wd.Documents.Add() '(Template:="Normal", NewTemplate:=False, DocumentType:=0)
            Dim ver
            ver = wd.Version
            Dim bool2007 As Boolean
            bool2007 = True
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If
            Dim strExt As String = ".docx"

            boolHasMacro = False
            If bool2007 Then
                'Note: Word2007 has no compatability mode parameter
                If wd.ActiveDocument.HasVBProject Then
                    strExt = ".docm"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                End If
            Else
                If wd.ActiveDocument.HasVBProject Then
                    strExt = ".docm"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs2(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strTemplate = Replace(strTemplate, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wd.ActiveDocument.SaveAs2(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                End If
            End If

            'wd.ActiveDocument.SaveAs(FileName:=strTemplate, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:=False)
            Call SetNormal(wd)

            '20180701 LEE:
            Call SetDocGlobal(wd.ActiveDocument)

            '20180701 LEE
            'Implement time saving trick
            Call SpellingOff(wd.ActiveDocument, False)

            If boolVerbose Then
                wd.Visible = True
                wd.WindowState = Office.Interop.Word.WdWindowState.wdWindowStateMaximize
            End If

            ftNormal = wd.Selection.Font.Name

            'record normal font
            NormalFontsize = wd.Selection.Font.Size

            intOTables = wd.ActiveDocument.Tables.Count

            'add BlueHyperlink style
            If boolBLUEHYPERLINK Then
                Call CreateBlueHyperlink(wd)
            End If

        Catch ex As Exception
            Try
                wd.Application.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            Catch ex1 As Exception

            End Try
            wd = Nothing

            str1 = "There was a problem preparing the Example Report." & Chr(10) & Chr(10) & ex.Message ' & Chr(10) & Chr(10) & str2


            MsgBox(str1, MsgBoxStyle.Information, "File not found...")
            GoTo end3
        End Try

        wd.ActiveWindow.View.ShowFieldCodes = False

        Call PositionProgress()
        frmH.lblProgress.Text = "Preparing Example Report Table Section..."
        frmH.lblProgress.Visible = True
        frmH.lblProgress.Refresh()
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        'frmH.panProgress.Visible = True
        'frmH.panProgress.Refresh()

        boolMeRefresh = False

        'now enter figs, tables and appendices in order
        Dim dvRBS As System.Data.DataView
        If boolEntireReport Then
            'prepare dv1 from tblReportStatements
            Dim tblR As System.Data.DataTable
            Dim strFR As String

            tblR = tblReportStatements
            strFR = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE = -1 AND NUMCOMPANY < 20000" ' & " AND ID_TBLCONFIGBODYSECTIONS >= 139 AND ID_TBLCONFIGBODYSECTIONS <= 141"
            dvRBS = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)
            dvRBS.RowFilter = arrRBSColumns(2, 0)

        Else
            dvRBS = frmH.dgvReportStatements.DataSource
        End If


        Dim fontsize
        Dim varAlign

        For Count1 = 1 To intCt
            'str1 = arr1(2, Count1)
            str1 = "Report Body Sections"
            int1 = arr1(1, Count1) 'introw
            intRow = int1
            boolTemp = boolFormLoad
            boolFormLoad = True
            'frmH.dgvReportStatements.CurrentCell = frmH.dgvReportStatements.Rows.item(int1).Cells("CHARHEADINGTEXT")
            boolFormLoad = boolTemp
            'boolGo = DoIndReportSections(wd, str1, wdguwu, Count1 - 1, dvRBS)

            If boolEntireReport Then
            Else
                boolGo = DoIndReportSections(wd, str1, intRow, dvRBS)
            End If
            int2 = arr1(3, Count1)

            'wdd.visible = True

            Select Case int2

                Case 139 'figures
                    If boolEntireReport Then
                        wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        'Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                        fontsize = wd.Selection.Font.Size
                        wd.Selection.Font.Size = 18
                        varAlign = wd.Selection.ParagraphFormat.Alignment
                        For Count2 = 1 To 8
                            wd.Selection.TypeParagraph()
                        Next
                        wd.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        wd.Selection.TypeText("FIGURES")
                        wd.Selection.TypeParagraph()
                        wd.Selection.Font.Size = fontsize
                        wd.Selection.ParagraphFormat.Alignment = varAlign
                        wd.Selection.TypeParagraph()
                        wd.Selection.TypeParagraph()

                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Else
                        Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                    End If

                    'Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                    Call InsertGraphics("Figure", wd, False, 0, False)


                Case 140 'tables

                    Try
                        Call PrepareWatson(wd)
                    Catch ex As Exception
                        str1 = "There seems to have been a problem creating the table portion of this report." & ChrW(10) & ChrW(10)
                        str1 = str1 & "Try generating the report again." & ChrW(10) & ChrW(10)
                        str1 = str1 & "If the problem persists, please contract your StudyDoc Administrator."
                        MsgBox(str1, MsgBoxStyle.Information, "Error in report body...")
                        'wdd.visible = True
                        GoTo end1
                    End Try

                Case 141 'appendices

                    If boolEntireReport Then
                        wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                        fontsize = wd.Selection.Font.Size
                        wd.Selection.Font.Size = 18
                        varAlign = wd.Selection.ParagraphFormat.Alignment
                        For Count2 = 1 To 8
                            wd.Selection.TypeParagraph()
                        Next
                        wd.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                        wd.Selection.TypeText("APPENDICES")
                        wd.Selection.TypeParagraph()
                        wd.Selection.Font.Size = fontsize
                        wd.Selection.ParagraphFormat.Alignment = varAlign
                        wd.Selection.TypeParagraph()
                        wd.Selection.TypeParagraph()

                        'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Else
                        Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                    End If

                    'Call PageSetup(wd, "P") 'L=Landscape, P=Portrait
                    Call InsertGraphics("Appendix", wd, False, 0, False)

            End Select


        Next


        Try
            'now do search/replace
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            wd.Selection.WholeStory()
        Catch ex As Exception
            GoTo end3
        End Try


        Dim intSR As Short
        intSR = 1

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        'now enter Table of Tables



        If boolExcludeCoverPage Then
        Else
            wd.Selection.TypeParagraph()
            Call TableofTables_01(wd, 247)
            wd.Selection.TypeParagraph()

            Call ModifyTOC(wd)
        End If

        'look for section break as first character. delete if found
        Try
            Call EvalCoverPage(wd)
        Catch ex As Exception

        End Try

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        'generate Watson tables
        'wdd.visible = True 'for testing

        boolMeRefresh = True

        wrdselection = wd.Selection()
        With wd.ActiveDocument.Bookmarks
            .Add(Range:=wrdselection.Range, Name:="End2")
            .ShowHidden = False
        End With
        pos2 = wd.Selection.Start

        If boolExcludeCoverPage Then
        Else


            'enter report header
            wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            With wd.Selection.ParagraphFormat
                .LeftIndent = 100 'InchesToPoints(1)
                '.SpaceBeforeAuto = False
                '.SpaceAfterAuto = False
            End With
            With wd.Selection.ParagraphFormat
                '.SpaceBeforeAuto = False
                '.SpaceAfterAuto = False
                .FirstLineIndent = -100 'InchesToPoints(-1)
            End With
            wd.Selection.Font.Bold = True
            wd.Selection.TypeText(Text:="Study: " & vbTab & frmH.cbxStudy.Text)
            wd.Selection.TypeParagraph()
            'wd.Selection.TypeText(Text:="Report Title: " & vbTab & frmH.lblReportTitle.Text)
            wd.Selection.TypeText(Text:="Report Title: " & vbTab & gReportTitle)
            wd.Selection.TypeParagraph()
            strDo = "Example Report Tables Section"
            wd.Selection.TypeText(Text:=strDo)
            wd.Selection.TypeParagraph()
            'enter date
            strDo = "Date Prepared: " & vbTab & Format(dt, "MMMM dd, yyyy hh:mm:ss tt")
            wd.Selection.TypeText(Text:=strDo)
            wd.Selection.TypeParagraph()
            With wd.Selection.ParagraphFormat
                .LeftIndent = 0 'InchesToPoints(1)
                '.SpaceBeforeAuto = False
                '.SpaceAfterAuto = False
            End With
            With wd.Selection.ParagraphFormat
                '.SpaceBeforeAuto = False
                '.SpaceAfterAuto = False
                .FirstLineIndent = 0 'InchesToPoints(-1)
            End With

            wd.Selection.TypeParagraph()
            wd.Selection.Font.Bold = False
            wd.Selection.TypeParagraph()

        End If




        'goto beginning of document
        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        'now set margins for individual sections

        System.Windows.Forms.Application.DoEvents()

        Try
            Call LinkPrevious(wd, False, True)

        Catch ex As Exception

        End Try

        'goto beginning of document
        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        Try
            If EvalCoverPage(wd) Then
                'redo End2
                wrdselection = wd.Selection()
                With wd.ActiveDocument.Bookmarks
                    .Add(Range:=wrdselection.Range, Name:="End2")
                    .ShowHidden = False
                End With
                pos2 = wd.Selection.Start
            End If
        Catch ex As Exception

        End Try

        Try

            '20180309 LEE:
            'No. Leave header/footer in so that pages match report if user needs to paste to report
            '20180327 LEE: No, need to evaluate boolExcludeHeaderFooter
            If boolIncludeWaterMark Then
            Else
                If boolExcludeHeaderFooter Then
                    Call ClearHeaderFooter(wd)
                End If
            End If

        Catch ex As Exception

        End Try

        'call searchreplace
        Dim rngA As Microsoft.Office.Interop.Word.Range
        With wd


            'wdd.visible = True

            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="End2")
            pos2 = .Selection.Bookmarks.Item("End2").Start
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            pos1 = .Selection.Start

            'this next line will cause the entire selection to be chosen
            .Selection.SetRange(Start:=0, End:=wd.ActiveDocument.Bookmarks.Item("End2").Start)
            wrdselection = .Selection
            rngA = .Selection.Range


            If boolDoFormulas Then
                frmH.lblProgress.Text = "Formatting chemical formulas..."
                frmH.lblProgress.Refresh()
                Try
                    Call ChemFormula(rngA, wd) 'to address any sub/superscripts
                Catch ex As Exception
                    str1 = "There was a problem completing the Chemical Formula Signature Search/Replace action. Report generation will continue."
                    str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."

                End Try
            End If

            '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

            'remove bookmarks
            Try
                .ActiveDocument.Bookmarks.Item("End2").Delete()
            Catch ex As Exception
            End Try
            Try
                .ActiveDocument.Bookmarks.Item("Temp1").Delete()
            Catch ex As Exception
            End Try
            Try
                .ActiveDocument.Bookmarks.Item("Temp2").Delete()
            Catch ex As Exception
            End Try

        End With

        frmH.lblProgress.Text = "Inserting watermarks..."
        frmH.lblProgress.Refresh()

        Dim strDate As String
        Dim strTime As Date
        strDate = Format(dt, "MMM dd, yyyy")
        strTime = Format(dt, "hh:mm:ss tt")
        str1 = "DRAFT" & ChrW(10) & strDate & ChrW(10) & strTime

        Try
            Call InsertWatermark(wd, boolIncludeWaterMark, str1)
        Catch ex As Exception
            str1 = "Unfortunately, this version of Microsoft" & ChrW(10) & " Word does not contain the Word watermarking funtion supported by StudyDoc." & ChrW(10) & ChrW(10)
            str1 = str1 & "Word 2002 or higher must be used. " & ChrW(10) & ChrW(10)
            str1 = str1 & "The report will be prepared without a watermark."
            MsgBox(str1, MsgBoxStyle.Information, "Watermark not supported...")
        End Try

        str1 = "Example Report Table Section Completed."
        'MsgBox(str1, MsgBoxStyle.Information, "Action completed...")


        '20180701 LEE
        'Implement time saving trick
        Call SpellingOff(wd.ActiveDocument, True)

end1:

        Try

            wd.ActiveWindow.View.ShowFieldCodes = False

            'clear formatting again
            Try
                wd.Selection.Find.ClearFormatting()
            Catch ex As Exception

            End Try
            Try
                wd.Selection.Range.Find.ClearFormatting()
            Catch ex As Exception

            End Try

            Try
                wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            Catch ex1 As Exception

            End Try
            'wdd.visible = True
            'wd.Activate()



            'save as temp then display in afr
            Dim strP As String
            strP = GetNewTempFile(True)

            If boolHasMacro And boolSaveAsDocx = False Then
                strP = Replace(strP, ".xml", ".docm", 1, -1, CompareMethod.Text)
            Else
                strP = Replace(strP, ".xml", ".docx", 1, -1, CompareMethod.Text)
            End If

            frmH.lblProgress.Text = "Saving document..." & ChrW(10) & ChrW(10) & strP
            frmH.lblProgress.Refresh()

            If boolHasMacro And boolSaveAsDocx = False Then
                wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
            Else
                wd.ActiveDocument.SaveAs(FileName:=strP, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
            End If

            If gDoPDF Then
                If wd.Version < 12 Then
                    gDoPDF = False
                Else
                    gDoPDF = True
                End If
            End If

            If gDoPDF Then
                Call CreatePDF(wd, strP)
            End If

            str1 = "Example Table Section"
            Call ReportHistoryItem(str1)

            Try
                wd.ActiveDocument.Close(False)
            Catch ex As Exception

            End Try

            Try
                wd.Application.ShowWindowsInTaskbar = boolSTB
            Catch ex As Exception

            End Try

            Try
                wd.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            Catch ex As Exception

            End Try

            Threading.Thread.Sleep(250)

            If gDoPDF Then
            Else
                Call OpenAFR(strP, "", False, boolSTB, True, False)
            End If


        Catch ex As Exception

        End Try

        wd = Nothing

end2:

        Cursor.Current = Cursors.Default

        'Try
        '    str1 = "Example Table Section"
        '    Call ReportHistoryItem(str1)

        'Catch ex As Exception

        'End Try

        boolShowExample = False



        If gboolER Then
        Else
            frmWordStatement.Activate()
        End If

        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False
        'frmH.pb2.Visible = False

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()

        frmH.Refresh()

        Exit Sub

end3:

        Try
            wd.Application.ShowWindowsInTaskbar = boolSTB
        Catch ex As Exception

        End Try

        Try
            wd.Application.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)
            wd = Nothing
        Catch ex As Exception

        End Try

        'Try
        '    wdguwu.Application.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)

        'Catch ex As Exception

        'End Try

        'wdguwu = Nothing

        Cursor.Current = Cursors.Default
        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()

        frmH.Refresh()

        boolShowExample = False


    End Sub

    Sub ExampleCoverPageBreak(ByVal wd As Microsoft.Office.Interop.Word.Application)

        With wd

            wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
            wd.Selection.TypeParagraph()

            wd.Selection.TypeText("PAGE INTENTIONALLY LEFT BLANK")
            wd.Selection.TypeParagraph()



            '20160920 LEE:
            'The next code is not correct logic
            'An individual table must have the same headers and footers as a normal document to all copy/paste
            GoTo end1

            'delete any contents in the footer
            'enter footer
            If wd.ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
                wd.ActiveWindow.Panes(2).Close()
            End If
            If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow. _
                ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            End If
            wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader

            Try
                .ActiveWindow.ActivePane.View.NextHeaderFooter()
            Catch ex As Exception

            End Try
            .Selection.HeaderFooter.LinkToPrevious = False ' Not .Selection.HeaderFooter.LinkToPrevious

            wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
            wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
            If wd.ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
                wd.ActiveWindow.Panes(2).Close()
            End If
            If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow. _
                ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            End If
            wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader

            Try
                .ActiveWindow.ActivePane.View.NextHeaderFooter()
            Catch ex As Exception

            End Try
            .Selection.WholeStory()
            .Selection.Delete(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)
            .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

end1:

        End With

    End Sub

    Sub InsertIndividualFigs(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim boolFig As Boolean
        Dim boolApp As Boolean
        Dim boolAtt As Boolean
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim tbl As System.Data.DataTable
        Dim tblF As System.Data.DataTable
        Dim str2 As String
        Dim var1
        Dim fontsize
        Dim Count2 As Short
        Dim varAlign
        Dim boolST As Boolean = False
        Dim varReplace
        Dim myRange As Microsoft.Office.Interop.Word.Range
        Dim boolFound As Boolean
        Dim boolF1 As Boolean = False
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim strFC As String
        Dim int1 As Short
        Dim int2 As Short
        Dim strFA As String
        Dim tblA As System.Data.DataTable

        tblF = tblFigures
        tblA = tblAppendix
        boolFig = False
        boolApp = False
        ctFigures = 0
        ctAppendix = 0
        tblAttachments.Clear()
        ctAttachments = 0
        boolAtt = False

        'wdd.visible = True

        Dim str1 As String
        Dim strTitle As String

        'wdd.visible = True

        With wd

            'do individual Figures first for Figure labeling purposes
            Dim strFind As String
            'strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLFIGURE = -1 AND BOOLAPPENDIX = 0"
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARFCID IS NOT NULL AND BOOLINCLUDEINREPORT <> 0"

            Erase rows
            rows = tblAppFigs.Select(strF, strS)
            intRows = rows.Length
            For Count1 = 0 To intRows - 1
                boolFig = False
                boolApp = False
                Dim boolWord As Boolean = False
                Dim int3 As Short

                int1 = rows(Count1).Item("BOOLFIGURE")
                int2 = rows(Count1).Item("BOOLAPPENDIX")
                int3 = rows(Count1).Item("BOOLINSERTWORDDOCS")
                If int1 = -1 Then
                    boolFig = True
                End If
                If int2 = -1 Then
                    boolApp = True
                End If
                If int3 = -1 Then
                    boolWord = True
                End If


                strFC = NZ(rows(Count1).Item("CHARFCID"), "")
                strFind = "[APPFIGINSERT_" & strFC & "]"

                'do search for appfiginserts
                boolF1 = True

                'wdd.visible = True

                Do Until boolF1 = False
                    .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                    'wd.Visible = True 'debug
                    mySel = wd.Selection
                    With mySel.Find
                        .ClearFormatting()
                        .Text = strFind
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Forward = True
                        With .Replacement
                            varReplace = ""
                            .ClearFormatting()
                            .Text = varReplace
                        End With
                        Try
                            Do While .Execute(Replace:=Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne, Format:=True, MatchCase:=False, MatchWholeWord:=True, Forward:=True)
                                boolF1 = .Found
                                If boolF1 Then
                                Else
                                    Exit Do
                                End If

                                If boolWord Then
                                    'don't add items to tables because will happen later
                                Else
                                    If boolFig Then
                                        If gboolDisplayAttachments Then
                                            strFA = "Attachment"
                                            ctAttachments = ctAttachments + 1
                                            'record info in tblFigures
                                            Dim row As DataRow = tblAttachments.NewRow
                                            row.BeginEdit()
                                            row("AttachmentNumber") = ctAttachments
                                            row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                            row("AttachmentName") = strFA '"Figure"
                                            row("RepWatsonID") = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                                            row("NumRow") = ctFigures
                                            str2 = NZ(rows(Count1).Item("CHARFCID"), "")
                                            row("CHARFCID") = str2
                                            row.EndEdit()
                                            tblAttachments.Rows.Add(row)

                                        Else
                                            strFA = "Figure"
                                            ctFigures = ctFigures + 1
                                            'record info in tblFigures
                                            Dim row As DataRow = tblF.NewRow
                                            row.BeginEdit()
                                            row("FigureNumber") = ctFigures
                                            row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                            row("FigureName") = strFA '"Figure"
                                            row("RepWatsonID") = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                                            row("NumRow") = ctFigures
                                            str2 = NZ(rows(Count1).Item("CHARFCID"), "")
                                            row("CHARFCID") = str2
                                            row.EndEdit()
                                            tblF.Rows.Add(row)
                                        End If


                                    Else
                                        If gboolDisplayAttachments Then
                                            strFA = "Attachment"
                                            ctAttachments = ctAttachments + 1
                                            'record info in tblFigures
                                            Dim row As DataRow = tblAttachments.NewRow
                                            row.BeginEdit()
                                            row("AttachmentNumber") = ctAttachments
                                            row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                            row("AttachmentName") = strFA
                                            row("RepWatsonID") = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                                            row("NumRow") = ctFigures
                                            str2 = NZ(rows(Count1).Item("CHARFCID"), "")
                                            row("CHARFCID") = str2
                                            row.EndEdit()
                                            tblAttachments.Rows.Add(row)
                                        Else
                                            strFA = "Appendix"
                                            ctAppendix = ctAppendix + 1
                                            'record info in tblFigures
                                            Dim row As DataRow = tblA.NewRow
                                            row.BeginEdit()
                                            row("AppendixNumber") = ctAppendix 'AppendixLetter(ctAppendix)
                                            str1 = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                            row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                            str1 = rows(Count1).Item("CHARTYPE")
                                            str2 = ""
                                            If StrComp(str1, "RC", CompareMethod.Text) = 0 Then
                                                str2 = "Chromatogram"
                                            ElseIf StrComp(str1, "LM", CompareMethod.Text) = 0 Then
                                                str2 = "LM"
                                            Else
                                                str2 = "Appendix"
                                                If StrComp(str1, "ST", CompareMethod.Text) = 0 Then
                                                    boolST = True
                                                Else
                                                    boolST = False
                                                End If
                                            End If
                                            var1 = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                                            row("AppendixName") = str2
                                            row("RepWatsonID") = var1
                                            row("NumRow") = ctAppendix
                                            row("CHARFCID") = NZ(rows(Count1).Item("CHARFCID"), "")
                                            row.EndEdit()
                                            tblA.Rows.Add(row)
                                        End If


                                    End If
                                End If


                                'add caption
                                If boolWord Then
                                Else
                                    With wd
                                        Try
                                            .Selection.InsertCaption(Label:=strFA, TitleAutoText:="", Title:="", Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionBelow) ', ExcludeLabel:=0)
                                            If BOOLTABLELABELSECTION Then
                                                Call ApplyChapterNumber(wd, strFA)
                                            End If
                                        Catch ex As Exception
                                            'need to add caption

                                            If gboolDisplayAttachments Then
                                                wd.CaptionLabels.Add(Name:=strFA)
                                                With wd.CaptionLabels(strFA)
                                                    .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseRoman
                                                    .IncludeChapterNumber = False
                                                End With

                                                If BOOLTABLELABELSECTION Then
                                                    Call ApplyChapterNumber(wd, strFA)
                                                End If

                                                wd.Selection.InsertCaption(Label:=strFA, TitleAutoText:="", Title:="", Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionBelow) ', ExcludeLabel:=0)

                                            Else
                                                If boolFig Then
                                                    wd.CaptionLabels.Add(Name:=strFA)
                                                    If StrComp(strFA, "Figure", CompareMethod.Text) = 0 Then
                                                        With wd.CaptionLabels(strFA)
                                                            .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleArabic ' wdCaptionNumberStyleUppercaseLetter
                                                            .IncludeChapterNumber = False
                                                        End With
                                                    Else
                                                        With wd.CaptionLabels(strFA)
                                                            .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseLetter ' wdCaptionNumberStyleUppercaseLetter
                                                            .IncludeChapterNumber = False
                                                        End With
                                                    End If

                                                    'wdCaptionNumberStyleUppercaseLetter

                                                    If BOOLTABLELABELSECTION Then
                                                        Call ApplyChapterNumber(wd, strFA)
                                                    End If

                                                    wd.Selection.InsertCaption(Label:=strFA, TitleAutoText:="", Title:="", Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionBelow) ', ExcludeLabel:=0)

                                                Else
                                                    wd.CaptionLabels.Add(Name:=strFA)
                                                    If StrComp(strFA, "Appendix", CompareMethod.Text) = 0 Then
                                                        With wd.CaptionLabels(strFA)
                                                            .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseLetter ' wdCaptionNumberStyleUppercaseLetter
                                                            .IncludeChapterNumber = False
                                                        End With
                                                    Else
                                                        With wd.CaptionLabels(strFA)
                                                            .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleArabic ' wdCaptionNumberStyleUppercaseLetter
                                                            .IncludeChapterNumber = False
                                                        End With
                                                    End If

                                                    If BOOLTABLELABELSECTION Then
                                                        Call ApplyChapterNumber(wd, strFA)
                                                    End If

                                                    wd.Selection.InsertCaption(Label:=strFA, TitleAutoText:="", Title:="", Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionBelow) ', ExcludeLabel:=0)

                                                End If
                                            End If

                                        End Try

                                        'enter nonbreaking space
                                        Call NBSP(wd, False)

                                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                                        .Selection.ParagraphFormat.TabStops.Add(Position:=72, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                                        With .Selection.ParagraphFormat
                                            .LeftIndent = 72
                                            '.SpaceBeforeAuto = False
                                            '.SpaceAfterAuto = False
                                        End With
                                        With .Selection.ParagraphFormat
                                            '.SpaceBeforeAuto = False
                                            '.SpaceAfterAuto = False
                                            .FirstLineIndent = -72
                                        End With
                                        strTitle = NZ(rows(Count1).Item("CHARTITLE"), "")
                                        'replace hyphens with nbh
                                        str1 = Replace(strTitle, "-", NBH, 1, -1, CompareMethod.Text)
                                        strTitle = str1
                                        .Selection.TypeText(Text:=vbTab & strTitle)
                                        .Selection.TypeParagraph()
                                        .Selection.TypeParagraph()
                                        'make keep with next
                                        .Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                                        .Selection.MoveUp(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                                        With .Selection.ParagraphFormat

                                            .WidowControl = True
                                            .KeepWithNext = True
                                            .KeepTogether = True

                                        End With
                                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Count:=1)
                                    End With

                                End If

                                'insert figure
                                'this needs to say if it is fig or app
                                '20151027 Larry: InsertGraphic isn't working
                                'ct[n] is being wiped out
                                Call InsertGraphicInd(wd, strFC)

                            Loop

                            boolF1 = False


                        Catch ex As Exception
                            boolF1 = False
                        End Try

                    End With

                Loop

            Next

        End With
end1:

    End Sub

    Sub InsertGraphicInd(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal strFCID As String)

        Dim boolFig As Boolean
        Dim boolApp As Boolean
        Dim boolWd As Boolean
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim tbl As System.Data.DataTable
        Dim tblF As System.Data.DataTable
        Dim str2 As String
        Dim var1
        Dim fontsize
        Dim Count2 As Short
        Dim varAlign
        Dim boolST As Boolean = False
        Dim strPath As String

        strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND CHARFCID = '" & strFCID & "'"
        rows = tblAppFigs.Select(strF)
        intRows = rows.Length

        If intRows = 0 Then
            GoTo end1
        End If

        Dim str1 As String
        Dim strTitle As String

        'wdd.visible = True

        With wd

            For Count1 = 0 To intRows - 1
                'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)

                'page setup according to configuration
                str1 = rows(Count1).Item("CHARPAGEORIENTATION")
                strPath = rows(Count1).Item("CHARPATH")
                var1 = rows(Count1).Item("BOOLFIGURE")
                If var1 = 0 Then
                    boolFig = False
                Else
                    boolFig = True
                End If
                var1 = rows(Count1).Item("BOOLAPPENDIX")
                If var1 = 0 Then
                    boolApp = False
                Else
                    boolApp = True
                End If
                var1 = NZ(rows(Count1).Item("BOOLINSERTWORDDOCS"), 0)
                If var1 = 0 Then
                    boolWd = False
                Else
                    boolWd = True
                End If
                If Len(strPath) = 0 Then
                Else
                    'insert page break
                    Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                    If boolWd Then
                        Call InsertWordFile(wd, strPath, boolFig, boolApp, gboolDisplayAttachments, rows, Count1)
                    Else
                        Call InsertFigs(wd, Count1, rows)
                    End If

                End If

                'DONT INSERT PAGE BREAK FOR AN INDIVIDUAL FIGURE
                If boolWd Then
                Else
                    If boolApp Then
                        wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    End If
                End If


            Next

        End With
end1:

    End Sub

    Sub InsertWordFile(wd As Microsoft.Office.Interop.Word.Application, strPath As String, boolFig As Boolean, boolApp As Boolean, boolAtt As Boolean, rows() As DataRow, intRow As Short)


        Dim pos1 As Int64
        Dim pos2 As Int64
        Dim pos3 As Int64
        Dim pos4 As Long
        Dim rng1 As Microsoft.Office.Interop.Word.Range
        Dim rng2 As Microsoft.Office.Interop.Word.Range
        Dim rng3 As Microsoft.Office.Interop.Word.Range
        Dim sty As Microsoft.Office.Interop.Word.Style
        Dim strStyle As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strA As String
        Dim strB As String
        Dim strC As String
        Dim var1, var2
        Dim Count1 As Int16

        Dim intCaptions As Int16 = 0

        If Len(NZ(strPath, "")) = 0 Then
        Else
            'does file exist?
            If File.Exists(strPath) Then
                With wd
                    pos1 = .Selection.Start
                    rng1 = .Selection.Range

                    '20170714 LEE:
                    'must first open file and record number of captions
                    'file may have text in addition to or instead of figures
                    'cannot use 'Do Until pos3 >= pos2' logic
                    Try
                        wd.Documents.Open(strPath)
                        Dim doc As Microsoft.Office.Interop.Word.Document
                        doc = wd.ActiveDocument
                        If boolFig Then
                            str1 = "Figure"
                        ElseIf boolApp Then
                            str1 = "Appendix"
                        ElseIf boolAtt Then
                            str1 = "Attachment"
                        Else
                            str1 = "Figure"
                        End If
                        Try
                            Dim arrC = doc.GetCrossReferenceItems(str1)
                            intCaptions = UBound(arrC)
                        Catch ex As Exception
                            intCaptions = 0
                        End Try

                        doc.Close()
                        Pause(0.2)

                    Catch ex As Exception
                        var1 = ex.Message
                        GoTo end1
                    End Try

                    .Selection.InsertFile(FileName:=strPath, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False)

                    Dim intC As Int16

                    If boolFig Then
                        intC = ctFigures
                    ElseIf boolApp Then
                        intC = ctAppendix
                    ElseIf boolAtt Then
                        intC = ctAttachments
                    Else
                        intC = ctFigures
                    End If

                    Dim row As DataRow
                    Dim dtbl As System.Data.DataTable ' DataTable
                    For Count1 = 1 To intCaptions

                        intC = intC + 1
                        If boolFig Then
                            dtbl = tblFigures
                            strA = "FigureNumber"
                            strB = "FigureName"
                            strC = "Figure"
                        ElseIf boolApp Then
                            dtbl = tblAppendix
                            strA = "AppendixNumber"
                            strB = "AppendixName"
                            strC = "Appendix"
                        ElseIf boolAtt Then
                            dtbl = tblAttachments
                            strA = "AttachmentNumber"
                            strB = "AttachmentName"
                            strC = "Attachment"
                        Else
                            dtbl = tblFigures
                            strA = "FigureNumber"
                            strB = "FigureName"
                            strC = "Figure"
                        End If

                        row = dtbl.NewRow
                        row.BeginEdit()
                        row(strA) = intC
                        row("AnalyteName") = NZ(rows(intRow).Item("charAnalyte"), "[NA]")
                        If boolApp Then
                            str1 = rows(intRow).Item("CHARTYPE")
                            If StrComp(str1, "RC", CompareMethod.Text) = 0 Then
                                strC = "Chromatogram"
                            ElseIf StrComp(str1, "LM", CompareMethod.Text) = 0 Then
                                strC = "LM"
                            Else
                                strC = "Appendix"
                            End If
                        End If
                        row(strB) = strC
                        row("RepWatsonID") = NZ(rows(intRow).Item("NUMWATSONRUNNUMBER"), 0)
                        row("NumRow") = intC
                        str2 = NZ(rows(intRow).Item("CHARFCID"), "")
                        row("CHARFCID") = str2
                        row.EndEdit()
                        dtbl.Rows.Add(row)

                    Next

                    If boolFig Then
                        ctFigures = ctFigures + intCaptions
                    ElseIf boolApp Then
                        ctAppendix = ctAppendix + intCaptions
                    ElseIf boolAtt Then
                        ctAttachments = ctAttachments + intCaptions
                    Else
                        ctFigures = ctFigures + intCaptions
                    End If

                    pos2 = .Selection.Start
                    rng2 = .Selection.Range

                    'select range and update
                    rng3 = .ActiveDocument.Range(Start:=pos1, End:=pos2)
                    rng3.Select()
                    .Selection.Fields.Update()


                    'don't do this anymore
                    GoTo skip1

                    'count number of figures between pos1 and pos2
                    rng1.Select()
                    pos3 = .Selection.Start

                    Do Until pos3 >= pos2

                        pos4 = .Selection.Start
                        sty = .Selection.Style
                        strStyle = sty.NameLocal
                        Try
                            If InStr(1, strStyle, "Caption", CompareMethod.Text) > 0 Then
                                If boolAtt Then
                                    ctAttachments = ctAttachments + 1
                                    'record info in tblFigures
                                    'Dim row As DataRow = tblAttachments.NewRow
                                    row.BeginEdit()
                                    row("AttachmentNumber") = ctAttachments
                                    row("AnalyteName") = NZ(rows(intRow).Item("charAnalyte"), "[NA]")
                                    row("AttachmentName") = "Attachment"
                                    row("RepWatsonID") = NZ(rows(intRow).Item("NUMWATSONRUNNUMBER"), 0)
                                    row("NumRow") = ctAttachments
                                    str2 = NZ(rows(intRow).Item("CHARFCID"), "")
                                    row("CHARFCID") = str2
                                    row.EndEdit()
                                    tblAttachments.Rows.Add(row)
                                Else
                                    If boolFig Then
                                        ctFigures = ctFigures + 1
                                        'record info in tblFigures
                                        'Dim row As DataRow = tblFigures.NewRow
                                        row.BeginEdit()
                                        row("FigureNumber") = ctFigures
                                        row("AnalyteName") = NZ(rows(intRow).Item("charAnalyte"), "[NA]")
                                        row("FigureName") = "Figure"
                                        row("RepWatsonID") = NZ(rows(intRow).Item("NUMWATSONRUNNUMBER"), 0)
                                        row("NumRow") = ctFigures
                                        str2 = NZ(rows(intRow).Item("CHARFCID"), "")
                                        row("CHARFCID") = str2
                                        row.EndEdit()
                                        tblFigures.Rows.Add(row)
                                    ElseIf boolApp Then
                                        ctAppendix = ctAppendix + 1
                                        'record info in tblAppendix
                                        'Dim row As DataRow = tblAppendix.NewRow
                                        row.BeginEdit()
                                        row("AppendixNumber") = ctAppendix 'AppendixLetter(ctAppendix)
                                        str1 = NZ(rows(intRow).Item("charAnalyte"), "[NA]")
                                        row("AnalyteName") = NZ(rows(intRow).Item("charAnalyte"), "[NA]")
                                        str1 = rows(intRow).Item("CHARTYPE")
                                        If StrComp(str1, "RC", CompareMethod.Text) = 0 Then
                                            str2 = "Chromatogram"
                                        ElseIf StrComp(str1, "LM", CompareMethod.Text) = 0 Then
                                            str2 = "LM"
                                        Else
                                            str2 = "Appendix"
                                        End If
                                        var1 = NZ(rows(intRow).Item("NUMWATSONRUNNUMBER"), 0)
                                        row("AppendixName") = str2
                                        row("RepWatsonID") = var1
                                        row("NumRow") = ctAppendix
                                        row("CHARFCID") = NZ(rows(intRow).Item("CHARFCID"), "")
                                        row.EndEdit()
                                        tblAppendix.Rows.Add(row)
                                    End If
                                End If

                            End If
                        Catch ex As Exception

                        End Try

                        '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToGraphic, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToNext, Count:=1, Name:="")

                        .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdParagraph, Count:=1)

                        pos3 = .Selection.Start

                        If pos3 = pos4 Or pos3 >= pos2 Then
                            Exit Do
                        End If

                    Loop

skip1:

                    rng2.Select()

                End With

            End If

        End If

end1:

    End Sub

    Sub InsertGraphics(ByVal strType As String, ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal boolSingle As Boolean, ByVal numRow As Short, boolIWD As Boolean)

        Dim boolFig As Boolean
        Dim boolApp As Boolean
        Dim boolAtt As Boolean
        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intRows As Short
        Dim Count1 As Short
        Dim tbl As System.Data.DataTable
        Dim tblF As System.Data.DataTable
        Dim str2 As String
        Dim var1
        Dim fontsize
        Dim Count2 As Short
        Dim varAlign
        Dim boolST As Boolean = False

        Dim strM As String

        tbl = tblAppendix
        tblF = tblFigures
        boolFig = False
        boolApp = False
        ctFigures = 0
        ctAppendix = 0
        tblAttachments.Clear()
        ctAttachments = 0
        boolAtt = False

        'get ctFigures from tblFigures
        ctFigures = tblF.Rows.Count
        ctAppendix = tbl.Rows.Count
        ctAttachments = tblAttachments.Rows.Count

        'wdd.visible = True

        Select Case strType
            Case "Figure"
                boolFig = True
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLFIGURE <> 0 AND BOOLAPPENDIX = 0 AND BOOLINCLUDEINREPORT <> 0"
            Case "Appendix"
                boolApp = True
                strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLFIGURE = 0 AND BOOLAPPENDIX <> 0 AND BOOLINCLUDEINREPORT <> 0"
        End Select

        Dim intS As Short
        Dim intE As Short



        dtbl = tblAppFigs
        strS = "INTORDER ASC"
        rows = dtbl.Select(strF, strS)
        intRows = rows.Length

        If intRows = 0 Then
            GoTo end1
        End If

        If boolSingle Then
            intS = numRow
            intE = numRow
        Else
            intS = 0
            intE = intRows - 1
        End If

        Dim str1 As String
        Dim strTitle As String

        'wdd.visible = True

        Dim boolDidInsert As Boolean
        'Dim boolIWD As Boolean
        Dim sty As Microsoft.Office.Interop.Word.Style
        Dim strStyle As String

        Dim boolWordInsert As Boolean

        With wd

            boolDidInsert = False

            For Count1 = intS To intE

                boolWordInsert = False

                'wd.Visible = True 'debug

                'Call InsertPageBreakAppFig(wd)

                'page setup according to configuration
                str1 = rows(Count1).Item("CHARPAGEORIENTATION")
                var1 = NZ(rows(Count1).Item("BOOLINSERTWORDDOCS"), 0)
                'Legend
                'do not take boolIWD from parameters
                'get from data
                If var1 = 0 Then
                    boolIWD = False
                Else
                    boolIWD = True
                End If
                'insert page break
                Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                If boolIWD Then

                    Dim strPath ' As String
                    Dim pos1 As Int64
                    Dim pos2 As Int64
                    Dim pos3 As Int64
                    Dim pos4 As Long
                    Dim rng1 As Microsoft.Office.Interop.Word.Range
                    Dim rng2 As Microsoft.Office.Interop.Word.Range
                    Dim rng3 As Microsoft.Office.Interop.Word.Range

                    Dim strA As String
                    Dim strB As String
                    Dim strC As String

                    strPath = NZ(rows(Count1).Item("CHARPATH"), "") ' "WIL-776002 BioAC Figures Report for Total Dox_gubbs01.doc"

                    'does file exist?
                    If File.Exists(strPath) Then

                        Call InsertPageBreakAppFig(wd)

                        Call PageSetup(wd, str1) 'L=Landscape, P=Portrait

                        boolWordInsert = True

                        'If boolDidInsert Then 'need to insert a section break
                        '    .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        'End If
                        '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

                        boolDidInsert = True
                        'boolWordInsert = True

                        With wd


                            'If boolWordInsert Then
                            'Else
                            '    .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                            '    boolWordInsert = True
                            'End If
                            '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)


                            pos1 = .Selection.Start
                            rng1 = .Selection.Range

                            '*****

                            '20170714 LEE:
                            'must first open file and record number of captions
                            'file may have text in addition to or instead of figures
                            'cannot use 'Do Until pos3 >= pos2' logic
                            Dim intCaptions As Int16 = 0

                            Try
                                wd.Documents.Open(strPath)
                                Dim doc As Microsoft.Office.Interop.Word.Document
                                doc = wd.ActiveDocument
                                If boolFig Then
                                    str1 = "Figure"
                                ElseIf boolApp Then
                                    str1 = "Appendix"
                                ElseIf boolAtt Then
                                    str1 = "Attachment"
                                Else
                                    str1 = "Figure"
                                End If
                                Try
                                    Dim arrC = doc.GetCrossReferenceItems(str1)
                                    intCaptions = UBound(arrC)
                                Catch ex As Exception
                                    intCaptions = 0
                                End Try

                                doc.Close()
                                Pause(0.2)

                            Catch ex As Exception
                                var1 = ex.Message
                                GoTo end1
                            End Try

                            .Selection.InsertFile(FileName:=strPath, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False)

                            'this creates an extra line
                            'delete it if it is the last inserted document
                            If Count1 = intE Then
                                .Selection.Delete(Unit:=WdUnits.wdCharacter, Count:=1)
                            End If


                            Dim intC As Int16

                            If boolFig Then
                                intC = ctFigures
                            ElseIf boolApp Then
                                intC = ctAppendix
                            ElseIf boolAtt Then
                                intC = ctAttachments
                            Else
                                intC = ctFigures
                            End If

                            Dim row As DataRow

                            For Count2 = 1 To intCaptions

                                intC = intC + 1
                                If boolFig Then
                                    dtbl = tblFigures
                                    strA = "FigureNumber"
                                    strB = "FigureName"
                                    strC = "Figure"
                                ElseIf boolApp Then
                                    dtbl = tblAppendix
                                    strA = "AppendixNumber"
                                    strB = "AppendixName"
                                    strC = "Appendix"
                                ElseIf boolAtt Then
                                    dtbl = tblAttachments
                                    strA = "AttachmentNumber"
                                    strB = "AttachmentName"
                                    strC = "Attachment"
                                Else
                                    dtbl = tblFigures
                                    strA = "FigureNumber"
                                    strB = "FigureName"
                                    strC = "Figure"
                                End If

                                Try
                                    row = dtbl.NewRow
                                    row.BeginEdit()
                                    row(strA) = intC
                                    row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                    If boolApp Then
                                        str1 = rows(Count1).Item("CHARTYPE")
                                        If StrComp(str1, "RC", CompareMethod.Text) = 0 Then
                                            strC = "Chromatogram"
                                        ElseIf StrComp(str1, "LM", CompareMethod.Text) = 0 Then
                                            strC = "LM"
                                        Else
                                            strC = "Appendix"
                                        End If
                                    End If
                                    row(strB) = strC
                                    row("RepWatsonID") = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                                    row("NumRow") = intC
                                    str2 = NZ(rows(Count1).Item("CHARFCID"), "")
                                    row("CHARFCID") = str2
                                    row.EndEdit()
                                    dtbl.Rows.Add(row)
                                Catch ex As Exception
                                    var1 = ex.Message
                                    var1 = var1
                                End Try



                            Next

                            If boolFig Then
                                ctFigures = ctFigures + intCaptions
                            ElseIf boolApp Then
                                ctAppendix = ctAppendix + intCaptions
                            ElseIf boolAtt Then
                                ctAttachments = ctAttachments + intCaptions
                            Else
                                ctFigures = ctFigures + intCaptions
                            End If

                            pos2 = .Selection.Start
                            rng2 = .Selection.Range

                            'select range and update
                            rng3 = .ActiveDocument.Range(Start:=pos1, End:=pos2)
                            rng3.Select()

                            Try
                                If BOOLTABLELABELSECTION Then
                                    Call ApplyChapterNumber(wd, strC)
                                End If
                            Catch ex As Exception
                                var1 = ex.Message
                            End Try

                            .Selection.Fields.Update()

                            GoTo skip1

                            '*****

                            .Selection.InsertFile(FileName:=strPath, Range:="", ConfirmConversions:=False, Link:=False, Attachment:=False)

                            'this creates an extra line
                            'delete it if it is the last inserted document
                            If Count1 = intE Then
                                .Selection.Delete(Unit:=WdUnits.wdCharacter, Count:=1)
                            End If


                            pos2 = .Selection.Start
                            rng2 = .Selection.Range

                            'select range and update
                            rng3 = .ActiveDocument.Range(Start:=pos1, End:=pos2)
                            rng3.Select()
                            .Selection.Fields.Update()

                            'count number of figures between pos1 and pos2
                            rng1.Select()
                            pos3 = .Selection.Start

                            Do Until pos3 >= pos2

                                pos4 = .Selection.Start
                                sty = .Selection.Style
                                strStyle = sty.NameLocal



                                Try
                                    If InStr(1, strStyle, "Caption", CompareMethod.Text) > 0 Then
                                        If boolFig Then
                                            ctFigures = ctFigures + 1
                                            'record info in tblFigures
                                            'Dim row As DataRow = tblF.NewRow
                                            row.BeginEdit()
                                            row("FigureNumber") = ctFigures
                                            row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                            row("FigureName") = "Figure"
                                            row("RepWatsonID") = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                                            row("NumRow") = ctFigures
                                            str2 = NZ(rows(Count1).Item("CHARFCID"), "")
                                            row("CHARFCID") = str2
                                            row.EndEdit()
                                            tblF.Rows.Add(row)
                                        ElseIf boolApp Then
                                            ctAppendix = ctAppendix + 1
                                            'record info in tblAppendix
                                            'Dim row As DataRow = tbl.NewRow
                                            row.BeginEdit()
                                            row("AppendixNumber") = ctAppendix 'AppendixLetter(ctAppendix)
                                            str1 = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                            row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                                            str1 = rows(Count1).Item("CHARTYPE")
                                            If StrComp(str1, "RC", CompareMethod.Text) = 0 Then
                                                str2 = "Chromatogram"
                                            ElseIf StrComp(str1, "LM", CompareMethod.Text) = 0 Then
                                                str2 = "LM"
                                            Else
                                                str2 = "Appendix"
                                                If StrComp(str1, "ST", CompareMethod.Text) = 0 Then
                                                    boolST = True
                                                Else
                                                    boolST = False
                                                End If
                                            End If
                                            var1 = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                                            row("AppendixName") = str2
                                            row("RepWatsonID") = var1
                                            row("NumRow") = ctAppendix
                                            row("CHARFCID") = NZ(rows(Count1).Item("CHARFCID"), "")
                                            row.EndEdit()
                                            tbl.Rows.Add(row)
                                        End If

                                    End If
                                Catch ex As Exception

                                End Try

                                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToGraphic, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToNext, Count:=1, Name:="")

                                .Selection.MoveDown(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdParagraph, Count:=1)

                                pos3 = .Selection.Start

                                If pos3 = pos4 Or pos3 >= pos2 Then
                                    Exit Do
                                End If

                            Loop

skip1:

                            rng2.Select()

                        End With

                    Else

                        'If boolDidInsert Then 'need to insert a section break
                        '    .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                        'End If

                        strM = "Note that the path configured for " & strType & " " & Count1 + 1 & ":" & ChrW(10) & ChrW(10) & NZ(strPath, "NA") & ChrW(10) & ChrW(10) & "does not exist."
                        strM = strM & ChrW(10) & ChrW(10) & "The " & strType & " will be configured as a place holder."
                        MsgBox(strM, vbInformation, "Invalid path...")

                        boolWordInsert = False

                    End If

                End If

                'If boolDidInsert Then 'need to insert a section break
                '    .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                'End If

                If boolWordInsert = False Then

                    If boolDidInsert Then
                        .Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)
                    Else
                        boolDidInsert = True
                        boolAppFigSectionStart = False
                    End If

                    If boolFig Then

                        ctFigures = ctFigures + 1
                        'record info in tblFigures
                        Dim row As DataRow = tblF.NewRow
                        row.BeginEdit()
                        row("FigureNumber") = ctFigures
                        row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                        row("FigureName") = "Figure"
                        row("RepWatsonID") = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                        row("NumRow") = ctFigures
                        row("CHARFCID") = NZ(rows(Count1).Item("CHARFCID"), "")
                        row.EndEdit()
                        tblF.Rows.Add(row)

                        'wdd.visible = True


                        Try
                            'don't do. use style in document
                            'With .CaptionLabels.Item("Figure")
                            '    .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleArabic
                            'End With

                            'need to test to make sure caption exists
                            With .CaptionLabels.Item("Figure")
                                var1 = .NumberStyle ' = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseRoman
                            End With

                            If BOOLTABLELABELSECTION Then
                                Call ApplyChapterNumber(wd, "Figure")
                            End If

                        Catch ex As Exception
                            wd.Application.CaptionLabels.Add(Name:="Figure")
                            With wd.Application.CaptionLabels.Item("Figure")
                                .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleArabic
                            End With

                            If BOOLTABLELABELSECTION Then
                                With wd.CaptionLabels("Figure")
                                    Call ApplyChapterNumber(wd, "Figure")
                                End With
                            End If

                        End Try

                        .Selection.InsertCaption(Label:="Figure", TitleAutoText:="", Title:="", _
                          Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionBelow) ', ExcludeLabel:=0)

                        'enter nonbreaking space
                        Call NBSP(wd, False)

                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                        .Selection.ParagraphFormat.TabStops.Add(Position:=72, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                        With .Selection.ParagraphFormat
                            .LeftIndent = 72
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                        End With
                        With .Selection.ParagraphFormat
                            '.SpaceBeforeAuto = False
                            '.SpaceAfterAuto = False
                            .FirstLineIndent = -72
                        End With
                        strTitle = NZ(rows(Count1).Item("CHARTITLE"), "")
                        'replace hyphens with nbh
                        str1 = Replace(strTitle, "-", NBH, 1, -1, CompareMethod.Text)
                        strTitle = str1
                        .Selection.TypeText(Text:=vbTab & strTitle)

                        .Selection.TypeParagraph()
                        If Count1 = intE Then
                        Else
                            .Selection.TypeParagraph()
                        End If


                    ElseIf boolApp Then

                        ctAppendix = ctAppendix + 1
                        'record info in tblAppendix
                        Dim row As DataRow = tbl.NewRow
                        row.BeginEdit()
                        row("AppendixNumber") = ctAppendix 'AppendixLetter(ctAppendix)
                        str1 = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                        row("AnalyteName") = NZ(rows(Count1).Item("charAnalyte"), "[NA]")
                        str1 = rows(Count1).Item("CHARTYPE")
                        If StrComp(str1, "RC", CompareMethod.Text) = 0 Then
                            str2 = "Chromatogram"
                        ElseIf StrComp(str1, "LM", CompareMethod.Text) = 0 Then
                            str2 = "LM"
                        Else
                            str2 = "Appendix"
                            If StrComp(str1, "ST", CompareMethod.Text) = 0 Then
                                boolST = True
                            Else
                                boolST = False
                            End If
                        End If
                        var1 = NZ(rows(Count1).Item("NUMWATSONRUNNUMBER"), 0)
                        row("AppendixName") = str2
                        row("RepWatsonID") = var1
                        row("NumRow") = ctAppendix
                        row("CHARFCID") = NZ(rows(Count1).Item("CHARFCID"), "")
                        row.EndEdit()
                        tbl.Rows.Add(row)

                        Dim vNS

                        If gboolDisplayAttachments Then
                            Try

                                'don't do. use style in document
                                'With .CaptionLabels.Item("Attachment")
                                '    .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseRoman
                                'End With

                                'need to test to make sure caption exists
                                With .CaptionLabels.Item("Attachment")
                                    var1 = .NumberStyle ' = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseRoman
                                End With


                                If BOOLTABLELABELSECTION Then
                                    Call ApplyChapterNumber(wd, "Attachment")
                                End If

                            Catch ex As Exception
                                wd.Application.CaptionLabels.Add(Name:="Attachment")
                                With wd.Application.CaptionLabels.Item("Attachment")
                                    .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseRoman
                                End With

                                If BOOLTABLELABELSECTION Then
                                    Call ApplyChapterNumber(wd, "Attachment")
                                End If

                            End Try
                        Else
                            Try
                                'don't do. use style in document
                                'With .CaptionLabels.Item("Appendix")
                                '    .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseLetter
                                'End With

                                'need to test to make sure caption exists
                                With .CaptionLabels.Item("Appendix")
                                    var1 = .NumberStyle ' = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseRoman
                                End With

                                If BOOLTABLELABELSECTION Then
                                    Call ApplyChapterNumber(wd, "Appendix")
                                End If

                            Catch ex As Exception
                                wd.Application.CaptionLabels.Add(Name:="Appendix")
                                With wd.Application.CaptionLabels.Item("Appendix")
                                    .NumberStyle = Microsoft.Office.Interop.Word.WdCaptionNumberStyle.wdCaptionNumberStyleUppercaseLetter
                                End With

                                If BOOLTABLELABELSECTION Then
                                    Call ApplyChapterNumber(wd, "Appendix")
                                End If

                            End Try
                        End If


                        'wdd.visible = True

                        If gboolDisplayAttachments Then
                            .Selection.InsertCaption(Label:="Attachment", TitleAutoText:="", Title:="", _
                              Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionBelow) ', ExcludeLabel:=0)
                        Else
                            .Selection.InsertCaption(Label:="Appendix", TitleAutoText:="", Title:="", _
                              Position:=Microsoft.Office.Interop.Word.WdCaptionPosition.wdCaptionPositionBelow) ', ExcludeLabel:=0)
                        End If

                        'enter nonbreaking space
                        Call NBSP(wd, False)

                        Dim numLI As Single 'Left Indent
                        'For appendix, attachment and table, set left indent depending on font size and font type
                        'Ricerca has Arial 12, which results in crowded caption and label
                        'The current selection is 'caption' style

                        If gboolDisplayAttachments Then
                            numLI = ReturnLeftIndent(wd, False, True, False)
                        Else
                            numLI = ReturnLeftIndent(wd, True, False, False)
                        End If

                        .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                        .Selection.ParagraphFormat.TabStops.Add(Position:=numLI, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabLeft, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                        Try
                            With .Selection.ParagraphFormat
                                .LeftIndent = numLI
                                '.SpaceBeforeAuto = False
                                '.SpaceAfterAuto = False
                            End With
                        Catch ex As Exception

                        End Try

                        Try
                            With .Selection.ParagraphFormat
                                '.SpaceBeforeAuto = False
                                '.SpaceAfterAuto = False
                                .FirstLineIndent = -numLI
                            End With
                        Catch ex As Exception

                        End Try


                        strTitle = NZ(rows(Count1).Item("CHARTITLE"), "")
                        str1 = Replace(strTitle, "-", NBH, 1, -1, CompareMethod.Text)
                        strTitle = str1

                        .Selection.TypeText(Text:=vbTab & strTitle)

                        .Selection.TypeParagraph()
                        If Count1 = intE Then
                        Else
                            .Selection.TypeParagraph()
                        End If

                        If boolST Then
                            Call SummaryTableAppendix(wd)
                        End If

                        '.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
                        '.Selection.TypeParagraph()
                        '.Selection.TypeParagraph
                    End If

                    If boolST Then
                        boolST = False
                    Else
                        Call InsertFigs(wd, Count1, rows)
                    End If

                    If boolIWD Then
                    Else

                    End If

                End If

            Next

        End With
end1:

    End Sub

    Sub ReportBody(ByVal wd, ByVal wdSt, ByVal boolWatermark, ByVal boolFieldCodesOnly)

        Dim strPath As String
        Dim strPathGuWu As String
        Dim dtbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim charS As String
        Dim boolInclude As Boolean
        Dim boolGuWu As Boolean
        Dim boolStatement As Boolean
        Dim dv As System.Data.DataView
        Dim dg As DataGrid
        Dim ct As Short
        Dim intRType As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim var1, var2, var3
        Dim rm, lm, rt
        Dim intr As Short
        Dim intc As Short
        Dim boolCont As Boolean
        Dim pos1, pos2
        Dim rngEnd As Microsoft.Office.Interop.Word.Range
        Dim rngStart As Microsoft.Office.Interop.Word.Range
        Dim idWS As Int64

        boolCont = True 'for legacy purposes
        strErrMsg = ""
        intErrCount = 0

        'wdd.visible = True

        EOCOV = 0
        EOTOC = 0
        EOTOT = 0
        EOTOA = 0
        EOTOF = 0

        'now enter blank tables for tables of contents
        'goto beginning of document
        wd.Selection.homeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
        'enter a paragraph return
        'wd.selection.typeparagraph()

        Dim tbl1 As System.Data.DataTable
        Dim tbl2 As System.Data.DataTable
        Dim strF As String
        Dim strF1 As String
        Dim strF2 As String
        Dim dr1() As DataRow
        Dim dr2() As DataRow
        Dim Count1 As Short
        Dim intCBS As Int64
        Dim charSectionName As String
        Dim intTNum As Short
        Dim strS As String
        Dim dgv As DataGridView
        Dim int1 As Short
        Dim charSCPage As String
        Dim boolIncludeCPage As Boolean
        Dim boolA As Short
        Dim HLEVEL As Short
        Dim strText As String
        Dim boolSkip As Boolean
        Dim boolPB As Boolean
        Dim strSectionText As String

        Dim boolFirstReturn As Boolean

        tbl1 = tblConfigBodySections
        tbl2 = tblReportStatements
        dgv = frmH.dgvReportStatements

        'wdd.visible = True

        'If boolEntireReport Then
        '    'prepare dv1 from tblReportStatements
        '    Dim tblR as System.Data.DataTable
        '    Dim strFR As String

        '    tblR = tblReportstatements
        '    strFR = "ID_TBLSTUDIES = " & id_tblStudies
        '    dv = New DataView(tblR, strFR, "INTORDER ASC", DataViewRowState.CurrentRows)

        'Else
        '    dv = frmH.dgvReportStatements.DataSource
        'End If
        dv = dgv.DataSource 'leave this as is. Need only one row for entire report
        ct = dv.Count

        boolFirstReturn = False
        strErrMsg = ""
        charSCPage = "Style 1"
        Call PositionProgress()
        frmH.pb1.Value = 0
        frmH.pb1.Maximum = ct + 1
        frmH.pb1.Visible = True
        frmH.lblProgress.Text = "Preparing Example Report Body..."
        frmH.lblProgress.Visible = True
        frmH.pb1.Refresh()
        frmH.lblProgress.Refresh()

        'frmH.panProgress.Visible = True
        'frmH.panProgress.Refresh()


        '***
        ''''''''''''''console.writeline("Start Counting")
        If boolEntireReport Then
        Else

        End If
        For Count1 = 0 To ct - 1

            frmH.pb1.Value = Count1 + 1
            frmH.pb1.Refresh()
            boolSkip = False

            intCBS = dv(Count1).Item("id_tblConfigBodySections")
            Select Case intCBS
                Case Is = 139 'Figure Section
                    boolSkip = True
                Case Is = 140 'Table Section
                    boolSkip = True
                Case Is = 141 'Appendix Section
                    boolSkip = True
            End Select

            If boolSkip Then
            Else
                strText = dv(Count1).Item("CHARHEADINGTEXT")
                If intCBS = 341 Then 'entire report
                    boolA = -1
                    HLEVEL = 0
                    boolInclude = True
                    boolGuWu = True
                    boolPB = False

                Else
                    boolA = dv(Count1).Item("boolInclude")
                    HLEVEL = dv(Count1).Item("NUMHEADINGLEVEL")
                    If boolA = -1 Then
                        boolInclude = True
                    Else
                        boolInclude = False
                    End If
                    boolA = dv(Count1).Item("boolGuWu")
                    If boolA = -1 Then
                        boolGuWu = True
                    Else
                        boolGuWu = False
                    End If
                    boolA = dv(Count1).Item("boolUseStatements")
                    If boolA = -1 Then
                        boolStatement = True
                    Else
                        boolStatement = False
                    End If
                    boolPB = dv(Count1).Item("boolPB")

                End If

                charS = NZ(dv(Count1).Item("charStatement"), "")
                var1 = dv(Count1).Item("charSectionName")
                idWS = dv(Count1).Item("ID_TBLWORDSTATEMENTS")

                charSectionName = NZ(var1, "NA") 'remember, a nonbound field named 'charSectionName' has been added to tblReportStatements

                If Len(charSectionName) = 0 Or boolInclude = 0 Then
                Else
                    strF2 = "charSectionName = '" & charSectionName & "'"
                    dr2 = tbl1.Select(strF2)
                    intTNum = dr2(0).Item("intWordTableNumber")
                    frmH.lblProgress.Text = "Generating " & strText & " Section..."
                    frmH.lblProgress.Refresh()
                    ctPB = ctPB + 1
                    If ctPB > frmH.pb1.Maximum Then
                        ctPB = 1
                    End If
                    frmH.pb1.Value = ctPB
                    frmH.pb1.Refresh()

                    'wdd.visible = True


                    ''''''''''''''console.writeline(Format(Now, "hh:mm:ss"))
                    Call DoReportBodySection(intCBS, boolInclude, wd, boolFirstReturn, charS, strText, boolGuWu, boolStatement, intTNum, HLEVEL, strText, boolPB, idWS)
                End If
            End If
        Next
        ''''''''''''''console.writeline("End Counting")

        wrdSelection = wd.selection()
        With wd.ActiveDocument.Bookmarks
            .Add(Range:=wrdSelection.Range, Name:="End2")
            .ShowHidden = False
        End With
        pos2 = wd.Selection.Start

        If intErrCount > 0 Then
            If intErrCount = 1 Then
                str1 = "The following section has been configured to use Report Statements. However, a Report Statement has not been assigned to this section. Please make note that this section body has been generated by StudyDoc."
            Else
                str1 = "The following sections have been configured to use Report Statements. However, a Report Statement has not been assigned to these sections. Please make note that these sections have been generated by StudyDoc."
            End If
            str1 = str1 & Chr(10) & Chr(10) & strErrMsg
            strErrMsg = str1
            'display this at end
        End If

        'goto beginning of document
        wd.Selection.homeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        'NO! Do search replace later
        'Call DoSearchReplace(wd, boolWatermark, boolFieldCodesOnly)


end1:
    End Sub

    Sub DoSearchReplace(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal boolWatermark As Boolean, ByVal boolFieldCodesOnly As Boolean, ByVal boolIgnoreTOC As Boolean, ByVal boolIgnoreTables As Boolean, ByVal boolOnlyTableFigs As Boolean, boolIsGuWuFast As Boolean)

        'call searchreplace
        Dim rngA As Microsoft.Office.Interop.Word.Range
        Dim intSR As Short
        Dim strPath As String
        Dim strPathGuWu As String
        Dim dtbl As System.Data.DataTable
        Dim str1 As String
        Dim str2 As String
        Dim charS As String
        Dim boolInclude As Boolean
        Dim boolGuWu As Boolean
        Dim boolStatement As Boolean
        Dim dv As System.Data.DataView
        Dim dg As DataGrid
        Dim ct As Short
        Dim intRType As Short
        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim var1, var2, var3
        Dim rm, lm, rt
        Dim intr As Short
        Dim intc As Short
        Dim boolCont As Boolean
        Dim pos1, pos2
        Dim rngEnd As Microsoft.Office.Interop.Word.Range
        Dim rngStart As Microsoft.Office.Interop.Word.Range
        Dim idWS As Int64
        Dim Count1 As Short

        'wdd.visible = True


        Try
            With wd

                'wdd.visible = True
                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                pos1 = .Selection.Start
                'this next line will cause the entire selection to be chosen
                Try
                    .Selection.SetRange(Start:=0, End:=wd.ActiveDocument.Bookmarks.Item("End2").Start)
                Catch ex As Exception

                End Try
                '.selection.setrange(Start:=0, End:=pos2)
                wrdSelection = .Selection

                rngA = .Selection.Range
                frmH.lblProgress.Text = "Processing Field Codes..."
                frmH.lblProgress.Refresh()
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                wd.Selection.WholeStory()

                intSR = 0

                'wdd.visible = True

                If boolFieldCodesOnly Then
                    'must search replace these:
                    '[FIRSTPAGESPECIAL]'removed
                    Call SearchReplace(wd, "Report Body", rngA, False, "", intSR, 161, 161, False, boolIgnoreTOC, boolIgnoreTables, 0) '[INSERTPAGEBREAK]
                    Call SearchReplace(wd, "Report Body", rngA, False, "", intSR, 160, 160, False, boolIgnoreTOC, boolIgnoreTables, 0) '[SECTIONBREAKNEXTPAGE]
                    Call SearchReplaceCustomFieldCode(wd)
                Else
                    Try
                        Call SearchReplace(wd, "Report Body", rngA, False, "", intSR, 0, 0, False, boolIgnoreTOC, boolIgnoreTables, 0)

                    Catch ex As Exception
                        str1 = "There was a problem completing the Search/Replace action: " & intSR & ". Report generation will continue."
                        str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                        str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
                        MsgBox(str1)
                    End Try

                    Try
                        Call SearchReplaceCustomFieldCode(wd)

                    Catch ex As Exception
                        str1 = "There was a problem completing the Search/Replace Custom Field action: " & intSR & ". Report generation will continue."
                        str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                        str1 = str1 & ChrW(10) & ChrW(10) & ex.Message
                        MsgBox(str1)
                    End Try

                    wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    wd.Selection.WholeStory()


                    'Try
                    '    'wdd.visible = False
                    '    Call SignatureSearch(wd)
                    'Catch ex As Exception
                    '    str1 = "There was a problem completing the Signature Search/Replace action. Report generation will continue."
                    '    str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."
                    'End Try
                End If

                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                If boolOnlyTableFigs Or boolIsGuWuFast Then
                Else

                    wd.Selection.WholeStory()
                    If boolDoFormulas Then
                        frmH.lblProgress.Text = "Formatting chemical formulas..."
                        frmH.lblProgress.Refresh()
                        Try

                            'wdd.visible = True
                            Call ChemFormula(rngA, wd) 'to address any sub/superscripts
                        Catch ex As Exception
                            str1 = "There was a problem completing the Chemical Formula Search/Replace action. Report generation will continue."
                            str1 = str1 & ChrW(10) & ChrW(10) & "Please report this to your StudyDoc Administrator."

                        End Try
                    End If


                    'reset rngA to not include index tables
                    Try
                        .Selection.SetRange(Start:=0, End:=wd.ActiveDocument.Bookmarks.Item("End2").Start)

                    Catch ex As Exception

                    End Try

                    'find pos2
                    pos1 = 0
                    For Count1 = 1 To 5
                        Select Case Count1
                            Case 1
                                Try
                                    var1 = wd.ActiveDocument.Bookmarks.Item("EOCOV").Start 'EOCOV
                                Catch ex As Exception
                                    var1 = 0
                                End Try
                            Case 2
                                Try
                                    var1 = wd.ActiveDocument.Bookmarks.Item("EOTOC").Start 'EOCOV
                                Catch ex As Exception
                                    var1 = 0
                                End Try
                            Case 3
                                Try
                                    var1 = wd.ActiveDocument.Bookmarks.Item("EOTOT").Start 'EOCOV
                                Catch ex As Exception
                                    var1 = 0
                                End Try
                            Case 4
                                Try
                                    var1 = wd.ActiveDocument.Bookmarks.Item("EOTOA").Start 'EOCOV
                                Catch ex As Exception
                                    var1 = 0
                                End Try
                            Case 5
                                Try
                                    var1 = wd.ActiveDocument.Bookmarks.Item("EOTOF").Start 'EOCOV
                                Catch ex As Exception
                                    var1 = 0
                                End Try
                        End Select
                        If var1 > pos1 Then
                            pos1 = var1
                        End If
                    Next

                    'this next line will cause the entire selection to be chosen

                    'wdd.visible = True

                    '.Selection.SetRange(Start:=pos1, End:=wd.ActiveDocument.Bookmarks.Item("End2").Start)

                    Try
                        .Selection.SetRange(Start:=pos1, End:=wd.ActiveDocument.Bookmarks.Item("End2").Start)
                    Catch ex As Exception

                    End Try

                    wrdSelection = .Selection
                    rngA = .Selection.Range

                    'wdd.visible = True

                    frmH.lblProgress.Text = "Configuring non-breaking hyphens..."
                    frmH.lblProgress.Refresh()
                    'wd.Selection.Range
                    'rngA
                    Try
                        Call ReplaceDegC(wd, wd.Selection.Range)
                    Catch ex As Exception
                        'wdd.visible = True
                        str1 = "Problem with converting deg C."
                        MsgBox(str1, MsgBoxStyle.Information, "deg C problem...")
                    End Try

                    'DONOT select tables for replace hyphens
                    Try
                        .Selection.SetRange(Start:=pos1, End:=wd.ActiveDocument.Bookmarks.Item("StartTables").Start)
                    Catch ex As Exception
                        Try
                            .Selection.SetRange(Start:=pos1, End:=wd.ActiveDocument.Bookmarks.Item("End2").Start)
                        Catch ex1 As Exception

                        End Try

                    End Try

                    'Try
                    '    Call ReplaceHyphens(wd, wd.Selection.Range)
                    'Catch ex As Exception
                    '    'wdd.visible = True
                    '    str1 = "Problem with converting hyphens to nbh."
                    '    MsgBox(str1, MsgBoxStyle.Information, "Hyphen problem...")
                    'End Try

                    'The following code must be performed BEFORE replacement of spaces with npsp
                    frmH.lblProgress.Text = "Updating fields..."
                    frmH.lblProgress.Refresh()

                    'update all fields
                    wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    'wd.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    wd.Selection.WholeStory()
                    wd.Selection.Fields.Update()
                    wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                    frmH.lblProgress.Text = "Inserting Watermarks..."
                    frmH.lblProgress.Refresh()

                    'The following code must be performed BEFORE replacement of spaces with npsp
                    If boolWatermark Then

                        Dim strDate As String
                        Dim strTime As Date
                        Dim dt As Date
                        dt = Now
                        strDate = Format(dt, "MMM dd, yyyy")
                        strTime = Format(dt, "hh:mm:ss tt")
                        str1 = "DRAFT" & ChrW(10) & strDate & ChrW(10) & strTime

                        Try
                            Call InsertWatermark(wd, boolIncludeWaterMark, str1)
                        Catch ex As Exception
                            str1 = "Unfortunately, this version of Microsoft" & ChrW(10) & " Word does not contain the Word watermarking funtion supported by StudyDoc." & ChrW(10) & ChrW(10)
                            str1 = str1 & "Word 2002 or higher must be used. " & ChrW(10) & ChrW(10)
                            str1 = str1 & "The report will be prepared without a watermark."
                            MsgBox(str1, MsgBoxStyle.Information, "Watermark not supported...")
                        End Try
                    End If

                End If

                Try
                    Call ReplaceHyphens(wd, wd.Selection.Range)
                Catch ex As Exception
                    'wdd.visible = True
                    str1 = "Problem with converting hyphens to nbh."
                    MsgBox(str1, MsgBoxStyle.Information, "Hyphen problem...")
                End Try

                'wdd.visible = True
                'frmH.Activate()

                'If boolUseHyperlinks And boolIsGuWuFast = False Then
                If boolIsGuWuFast = False Then
                    frmH.lblProgress.Text = "Creating table hyperlinks..."
                    frmH.lblProgress.Refresh()

                    Try
                        Call HyperlinkTables(wd, rngA)
                    Catch ex As Exception
                        'wdd.visible = True
                        str2 = frmH.lblProgress.Text
                        str1 = "Problem with creating table hyperlinks." & ChrW(10) & str2
                        MsgBox(str1, MsgBoxStyle.Information, "Table hyperlink problem...")
                    End Try

                    frmH.lblProgress.Text = "Creating appendix and figure hyperlinks..."
                    frmH.lblProgress.Refresh()

                    'wdd.visible = True

                    Try
                        Call HyperlinkAppendices(wd, rngA)
                    Catch ex As Exception
                        'wdd.visible = True
                        str2 = frmH.lblProgress.Text
                        str1 = "Problem with creating appendix and figure hyperlinks." & ChrW(10) & str2
                        MsgBox(str1, MsgBoxStyle.Information, "Appendix hyperlink problem...")
                    End Try
                    frmH.lblProgress.Text = "Done creating appendix and figure hyperlinks..."
                    frmH.lblProgress.Refresh()

                    'Try
                    '    Call HyperlinkFigures(wd, rngA)
                    'Catch ex As Exception
                    '    'wdd.visible = True
                    '    str1 = "Problem with inserting figure hyperlinks."
                    '    MsgBox(str1, MsgBoxStyle.Information, "Figure hyperlink problem...")
                    'End Try

                End If

                '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="Temp2")

                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                'clear find formatting
                .Selection.Find.ClearFormatting()

                'remove bookmarks
                Try
                    .ActiveDocument.Bookmarks.Item("End2").Delete()
                Catch ex As Exception
                End Try
                Try
                    .ActiveDocument.Bookmarks.Item("Temp1").Delete()
                Catch ex As Exception
                End Try
                Try
                    .ActiveDocument.Bookmarks.Item("Temp2").Delete()
                Catch ex As Exception
                End Try
                Try
                    .ActiveDocument.Bookmarks.Item("EOCOV").Delete()
                Catch ex As Exception
                End Try
                Try
                    .ActiveDocument.Bookmarks.Item("EOTOA").Delete()
                Catch ex As Exception
                End Try
                Try
                    .ActiveDocument.Bookmarks.Item("EOTOT").Delete()
                Catch ex As Exception
                End Try
                Try
                    .ActiveDocument.Bookmarks.Item("EOTOF").Delete()
                Catch ex As Exception
                End Try
                Try
                    .ActiveDocument.Bookmarks.Item("EOTOC").Delete()
                Catch ex As Exception
                End Try

            End With

        Catch ex As Exception

        End Try

    End Sub

    Sub FormatTOCColor(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Try
            With wd.ActiveDocument.Styles("TOC 1").Font
                If boolBLUEHYPERLINK Then
                    '.Color = BlueHyperlinkColor '  Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                    .ColorIndex = BlueHyperlinkColor ' WdColorIndex.wdBlue
                End If

            End With
            With wd.ActiveDocument.Styles("TOC 1")
                .AutomaticallyUpdate = True
            End With

        Catch ex As Exception

        End Try

        Try
            With wd.ActiveDocument.Styles("TOC 2").Font
                If boolBLUEHYPERLINK Then
                    .ColorIndex = BlueHyperlinkColor '  Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                End If
            End With
            With wd.ActiveDocument.Styles("TOC 2")
                .AutomaticallyUpdate = True
            End With

        Catch ex As Exception

        End Try

        Try
            With wd.ActiveDocument.Styles("TOC 3").Font
                If boolBLUEHYPERLINK Then
                    .ColorIndex = BlueHyperlinkColor ' Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                End If
            End With
            With wd.ActiveDocument.Styles("TOC 3")
                .AutomaticallyUpdate = True
            End With

        Catch ex As Exception

        End Try

        Try
            With wd.ActiveDocument.Styles("TOC 4").Font
                If boolBLUEHYPERLINK Then
                    .ColorIndex = BlueHyperlinkColor ' Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                End If
            End With
            With wd.ActiveDocument.Styles("TOC 4")
                .AutomaticallyUpdate = True
            End With

        Catch ex As Exception

        End Try

        Try
            With wd.ActiveDocument.Styles("TOC 5").Font
                If boolBLUEHYPERLINK Then
                    .ColorIndex = BlueHyperlinkColor ' Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                End If
            End With
            With wd.ActiveDocument.Styles("TOC 5")
                .AutomaticallyUpdate = True
            End With

        Catch ex As Exception

        End Try

        Try
            With wd.ActiveDocument.Styles("Table of Figures").Font
                If boolBLUEHYPERLINK Then
                    .ColorIndex = BlueHyperlinkColor ' Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                End If
            End With
            With wd.ActiveDocument.Styles("Table of Figures")
                .AutomaticallyUpdate = True
            End With

        Catch ex As Exception

        End Try

    End Sub

    Sub HyperlinkTables(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal rng1 As Microsoft.Office.Interop.Word.Range)

        'first save try to head off a hang problem
        wd.ActiveDocument.Save()

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim tbl1 As System.Data.DataTable
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim boolF As Boolean
        Dim str11 As String

        rng1.Select()
        mySel = wd.Selection

        tbl1 = tblTableN
        int1 = tblTableN.Rows.Count

        Dim pv As Short
        Dim pmax As Short
        Dim bool As Boolean
        Dim bool2 As Boolean

        Dim var1

        Dim oMax As Int64
        Dim oVal As Int64
        oMax = frmH.pb2.Maximum
        oVal = frmH.pb2.Value

        pmax = int1

        If int1 = 0 Then
            GoTo end1
        End If

        pmax = int1
        pv = 0
        frmH.pb1.Value = 1
        frmH.pb1.Maximum = pmax
        bool = frmH.pb1.Visible
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        'frmH.pb2.Value = 1
        'frmH.pb2.Maximum = pmax + 1
        'bool2 = frmH.pb2.Visible
        'frmH.pb2.Visible = True
        'frmH.pb2.Refresh()

        'wdd.visible = True

        boolF = True

        Dim intSC As Int64
        Dim intSCMax As Int64

        intSCMax = 100

        Call FillReturnLabelPosition(wd)

        With wd

            'wdd.visible = True

            For Count1 = int1 - 1 To 0 Step -1
                pv = pv + 1
                'frmH.pb2.Value = pv
                'frmH.pb2.Refresh()

                'go back to home
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                int2 = tbl1.Rows.Item(Count1).Item("TableNumber")

                str1 = "Table_" & int2

                str2 = CStr(int2)

                '20170712 LEE: New logic
                str2 = tbl1.Rows.Item(Count1).Item("TableNameNew")

                mySel.Find.ClearFormatting()

                intSC = 0
                intSCMax = 10
                frmH.pb1.Maximum = intSCMax
                frmH.pb1.Value = 0
                frmH.pb1.Refresh()

                str11 = "Creating Table Hyperlinks"
                str11 = str11 & ChrW(10) & "Finding '" & str1 & "'..."
                frmH.lblProgress.Text = str11
                frmH.lblProgress.Refresh()

                If boolDoHyperlinks And boolUseHyperlinks Then

                    Try

                        With mySel.Find
                            .Forward = True
                            .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                            .Format = True
                            .MatchCase = True
                            .MatchWholeWord = True
                            Do While .Execute(str1, , True)

                                intSC = intSC + 1
                                If intSC > intSCMax Then
                                    intSCMax = intSCMax + 10
                                End If
                                frmH.pb1.Value = intSC
                                frmH.pb1.Refresh()

                                Try

                                    '20181015 LEE:
                                    'If table caption is bold, then table reference is bold, and font-size replaces expected
                                    'unbold and put back font size
                                    Dim wsel As Microsoft.Office.Interop.Word.Selection
                                    wsel = wd.Selection
                                    Dim fs As Single
                                    Try
                                        fs = wsel.Font.Size
                                    Catch ex As Exception
                                        var1 = ex.Message
                                        var1 = var1
                                    End Try

                                    Call InsertTableCrossRef(wd, str2, int2)
                                    'wd.Selection.InsertCrossReference(ReferenceType:="Table", ReferenceKind:=Microsoft.Office.Interop.Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=str2, InsertAsHyperlink:=True, IncludePosition:=False) ', SeparateNumbers:=False, SeparatorString:=" ")
                                    'make bookmark for position
                                    'Dim wsel As Microsoft.Office.Interop.Word.Selection
                                    wsel = wd.Selection
                                    With wd.ActiveDocument.Bookmarks
                                        .Add(Range:=wsel.Range, Name:="TN")
                                        .ShowHidden = False
                                    End With
                                    'wdd.visible = True

                                    'go back to TN
                                    wd.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToBookmark, Name:="TN")
                                    'delete tn
                                    Try
                                        wd.ActiveDocument.Bookmarks("TN").Delete()
                                    Catch ex As Exception

                                    End Try

                                    'format font blue
                                    wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                                    'wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                                    'wd.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                                    If boolBLUEHYPERLINK Then
                                        Call BlueHyperlink(wd)
                                    End If

                                    '20181015 LEE:
                                    'If table caption is bold, then table reference is bold, and font-size replaces expected
                                    'unbold and put back font size
                                    wsel = wd.Selection
                                    wsel.Font.Size = fs
                                    wsel.Font.Bold = False

                                    'wdd.visible = True

                                    wd.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                                    'modify field code to conserve formatting
                                    'This is done because update will eliminated BlueHyperlink formatting
                                    wd.ActiveWindow.View.ShowFieldCodes = True ' Not ActiveWindow.View.ShowFieldCodes
                                    wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
                                    wd.Selection.TypeText(Text:=" \* MERGEFORMAT ")
                                    wd.ActiveWindow.View.ShowFieldCodes = False 'Not ActiveWindow.View.ShowFieldCodes
                                Catch ex As Exception

                                End Try

                            Loop

                            frmH.pb1.Value = frmH.pb1.Maximum
                            frmH.pb1.Refresh()

                        End With

                    Catch ex As Exception

                    End Try

                End If

            Next

        End With

        ''wdd.visible = False
        mySel.Find.ClearFormatting()

        'frmH.pb2.Maximum = oMax
        'frmH.pb2.Value = oVal
        'frmH.pb2.Refresh()

        'frmH.pb1.Visible = bool
        'frmH.pb2.Visible = bool2

end1:

    End Sub

    Sub PrepareWatson(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim dbPath As String
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        'Dim wStudyID As Long
        'Dim wSpeciesID As Long
        'Dim wProjectID As Long
        Dim wAnalyteID As Long
        'Dim arrAnalytes(7, 50) '1=AnalyteDescription, 2=AnalyteID, 3=AnalyteIndex
        '4=BQL, 5=AQL, 6=Conc Units, 7=AcceptedRuns
        Dim arrAnalyticalRuns()
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim Count4 As Short
        Dim Count5 As Short
        Dim Count6 As Short
        Dim Count7 As Short
        Dim var1, var2, var3, var4, var5, var6, var7, var8, var9
        Dim int1 As Short
        Dim int2 As Short
        Dim arrRegCon()
        Dim arrTemp(2, 50)
        Dim num1 As Object
        Dim num2 As Object
        Dim num3 As Object
        Dim arrBCStdActual()
        Dim arrLegend(3, 20)
        Dim ctLegend As Short
        Dim lng1 As Long
        Dim lng2 As Long
        Dim boolPortrait As Boolean
        Dim intLastAnal As Short
        Dim arrOrder(4, 100)
        '1=ColumnHeader, 2=Include(X), 3=Order, 4=ReportColumnHeader
        Dim ctCols 'number of columns in a table
        Dim strSub1 As String
        Dim strSub2 As String
        Dim pos1 As Short
        Dim pos2 As Short
        Dim inttemprows As Short
        Dim arrLastAnal(2, 50)
        Dim boolTable As Boolean
        Dim numSum As Object
        Dim numMean As Object
        Dim numSD As Object
        Dim drows() As DataRow
        Dim dtbl As System.Data.DataTable
        Dim tbl As System.Data.DataTable
        Dim dv As System.Data.DataView

        'make sure Word is 100%!!!!
        wd.ActiveWindow.ActivePane.View.Zoom.Percentage = 100

        '***Start 1
        'record wStudyID and wProjectID
        boolPortrait = True

        dtbl = tblStudies
        str1 = "id_tblStudies = " & id_tblStudies
        drows = dtbl.Select(str1)
        wStudyID = drows(0).Item("int_WatsonStudyID")
        wProjectID = drows(0).Item("int_WatsonProjectID")

        If wSpeciesID = 0 Then
            MsgBox("Hmmm. The study species is not configured in the Watson database. Please investigate and correct. This Workbook Preparation action is terminated.", vbInformation + vbOKOnly, "Species must be configured...")
            boolReportCont = False
            GoTo end1
        End If

        'retrieve active analytes in table globalanalytes using projectid
        'arrAnalytes already configured

        '***End 1
        'generate watson tables
        Call Watson(wd)
        If boolCont Then
        Else
            GoTo end2
        End If


end2:

        'wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

end1:

    End Sub

    Function DoIndReportSections(ByVal wd, ByVal strN, ByVal intRow, ByVal dvRS) As Boolean

        Dim str2 As String
        Dim int1 As Short
        Dim bool As Boolean
        Dim boolGo As Boolean
        Dim var1
        Dim strFilter As String
        Dim HLevel As Short
        Dim intFilter As Short
        Dim dv As System.Data.DataView
        Dim idWS As Int64

        dv = dvRS

        DoIndReportSections = True
        strFilter = ""
        intFilter = 0
        If InStr(1, strN, "Review Analytical Runs", CompareMethod.Text) > 0 Then
        ElseIf InStr(1, strN, "Summary Table", CompareMethod.Text) > 0 Then
        ElseIf InStr(1, strN, "Analytical Reference Std", CompareMethod.Text) > 0 Or InStr(1, strN, "QA Event Table", CompareMethod.Text) > 0 Or InStr(1, strN, "Add/Edit Contributors", CompareMethod.Text) > 0 Or StrComp(strN, "Cover Page", CompareMethod.Text) = 0 Or InStr(1, strN, "Sample Receipt Records", CompareMethod.Text) > 0 Or StrComp(strN, "AppTableFig", CompareMethod.Text) = 0 Then
            If InStr(1, strN, "Analytical Reference Std", CompareMethod.Text) > 0 Then
                strFilter = "Analytical Reference Standard Characterization"
                intFilter = 134
            ElseIf InStr(1, strN, "QA Event Table", CompareMethod.Text) > 0 Then
                strFilter = "QA Table"
                intFilter = 3
            ElseIf InStr(1, strN, "Add/Edit Contributors", CompareMethod.Text) > 0 Then
                strFilter = "Contributing Personnel"
                intFilter = 5
                'ElseIf StrComp(strN, "Cover Page", CompareMethod.Text) = 0 Then
                '    strFilter = "Cover Page"
                '    intFilter = 61
            ElseIf InStr(1, strN, "Sample Receipt Records", CompareMethod.Text) > 0 Then
                strFilter = "Sample Receipt"
                intFilter = 9
            End If

            Dim strPath As String
            Dim strPathGuWu As String
            Dim dtbl As System.Data.DataTable
            Dim charS As String
            Dim boolInclude As Boolean
            Dim boolGuWu As Boolean
            Dim boolStatement As Boolean
            Dim dg As DataGrid
            Dim ct As Short
            Dim intRType As Short
            Dim var2, var3
            Dim rm, lm, rt
            Dim intr As Short
            Dim intc As Short
            Dim boolCont As Boolean
            Dim tbl1 As System.Data.DataTable
            Dim tbl2 As System.Data.DataTable
            Dim strF1 As String
            Dim strF2 As String
            Dim dr1() As DataRow
            Dim dr2() As DataRow
            Dim Count1 As Short
            Dim intCBS As Int64
            Dim charSectionName As String
            Dim intTNum As Short
            Dim strS As String
            'Dim dgv As DataGridView
            Dim boolFirstReturn As Boolean
            'Dim dv1 as system.data.dataview
            Dim boolA As Short
            Dim int2 As Short
            Dim strText As String
            Dim boolPB As Boolean

            tbl1 = tblConfigBodySections
            tbl2 = tblReportStatements
            'dgv = frmH.dgvReportStatements
            'dv = dgv.DataSource
            dv = dvRS
            'dv1 = dvRS
            'strS = "intOrder ASC"
            'dv.Sort = strS
            ct = dv.Count
            'first determine report type
            If Len(strFilter) = 0 Then
                int1 = intRow
            Else
                'dv1 = frmH.dgvReportStatements.DataSource
                int1 = FindRowDVByCol(intFilter, dv, "ID_TBLCONFIGBODYSECTIONS")
            End If

            If int1 = -1 Then
                'intRType = dv(int1).Item("id_tblConfigReportType") 'any row will do
                intCBS = intFilter 'dv(int1).Item("id_tblConfigBodySections")
                strText = strN
                boolA = True
                boolPB = True
                HLevel = 1
                If boolA = -1 Then
                    boolInclude = True
                Else
                    boolInclude = False
                End If
                'boolA = dv(int1).Item("boolGuWu")
                If boolA = -1 Then
                    boolGuWu = True
                Else
                    boolGuWu = False
                End If
                'boolA = dv(int1).Item("boolUseStatements")
                If boolA = -1 Then
                    boolStatement = True
                Else
                    boolStatement = False
                End If
                boolStatement = False

                charS = "" 'NZ(dv(int1).Item("charStatement"), "")
                idWS = -1 'dv(int1).Item("ID_TBLWORDSTATEMENTS")
                charSectionName = strN

                int2 = 0 ' dv(int1).Item("ID_TBLCONFIGBODYSECTIONS")
            Else
                'intRType = dv(int1).Item("id_tblConfigReportType") 'any row will do
                intCBS = dv(int1).Item("id_tblConfigBodySections")
                strText = dv(int1).Item("CHARHEADINGTEXT")
                boolA = NZ(dv(int1).Item("boolInclude"), True)
                boolPB = NZ(dv(int1).Item("boolPB"), True)
                HLevel = NZ(dv(int1).Item("NUMHEADINGLEVEL"), 1)
                If boolA = -1 Then
                    boolInclude = True
                Else
                    boolInclude = False
                End If
                boolA = dv(int1).Item("boolGuWu")
                If boolA = -1 Then
                    boolGuWu = True
                Else
                    boolGuWu = False
                End If
                boolA = dv(int1).Item("boolUseStatements")
                If boolA = -1 Then
                    boolStatement = True
                Else
                    boolStatement = False
                End If

                charS = NZ(dv(int1).Item("charStatement"), "")
                var1 = dv(int1).Item("CHARSECTIONNAME")
                idWS = dv(int1).Item("ID_TBLWORDSTATEMENTS")
                charSectionName = NZ(var1, strN)

                int2 = dv(int1).Item("ID_TBLCONFIGBODYSECTIONS")
            End If




            strF2 = "charSectionName = '" & charSectionName & "'"
            dr2 = tbl1.Select(strF2)
            If dr2.Length = 0 Then
                intTNum = 0
            Else
                intTNum = dr2(0).Item("intWordTableNumber")
            End If
            frmH.lblProgress.Text = "Generating " & charSectionName & " Section..."
            frmH.lblProgress.Refresh()
            ctPB = ctPB + 1
            If ctPB > frmH.pb1.Maximum Then
                ctPB = 1
            End If
            frmH.pb1.Value = ctPB
            frmH.pb1.Refresh()

            'wdd.visible = True

            Call DoReportBodySection(intCBS, boolInclude, wd, boolFirstReturn, charS, charSectionName, boolGuWu, boolStatement, intTNum, HLevel, strText, boolPB, idWS)

            'find id
            If int2 = 10 Then
            Else
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                wd.Selection.WholeStory()
                Call SearchReplace(wd, "Report Body", wd.selection.range, False, "", 1, 0, 0, False, False, False, 0)
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                wd.Selection.WholeStory()

                'wdd.visible = False
                Call SignatureSearch(wd)
                'wdd.visible = True
                wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                'frmH.Activate()
            End If

        ElseIf StrComp(strN, "Report Body Sections", CompareMethod.Text) = 0 Then

            Dim strPath As String
            Dim strPathGuWu As String
            Dim dtbl As System.Data.DataTable
            Dim charS As String
            Dim boolInclude As Boolean
            Dim boolGuWu As Boolean
            Dim boolStatement As Boolean
            'Dim dv as system.data.dataview
            Dim dg As DataGrid
            Dim ct As Short
            Dim intRType As Short
            Dim var2, var3
            Dim rm, lm, rt
            Dim intr As Short
            Dim intc As Short
            Dim boolCont As Boolean
            Dim tbl1 As System.Data.DataTable
            Dim tbl2 As System.Data.DataTable
            Dim strF1 As String
            Dim strF2 As String
            Dim dr1() As DataRow
            Dim dr2() As DataRow
            Dim Count1 As Short
            Dim intCBS As Int64
            Dim charSectionName As String
            Dim intTNum As Short
            Dim strS As String
            Dim dgv As DataGridView
            Dim boolFirstReturn As Boolean
            Dim boolA As Short
            Dim strText As String
            Dim boolPB As Boolean

            boolCont = True 'for legacy purposes
            strErrMsg = ""
            intErrCount = 0

            tbl1 = tblConfigBodySections
            tbl2 = tblReportStatements
            'dgv = frmH.dgvReportStatements
            'dv = dgv.DataSource
            'strS = "intOrder ASC"
            'dv.Sort = strS
            ct = dv.Count
            'first determine report type
            'intRType = dv(0).Item("id_tblConfigReportType")
            int1 = intRow
            intRType = dv(int1).Item("id_tblConfigReportType")
            idWS = dv(int1).Item("ID_TBLWORDSTATEMENTS")

            boolFirstReturn = True
            strErrMsg = ""
            If intRType >= -1 Then 'Sample Report

                'dv = frmH.dgvReportStatements.DataSource
                'int1 = frmH.dgvReportStatements.CurrentRow.Index

                intCBS = dv(int1).Item("id_tblConfigBodySections")
                strText = dv(int1).Item("CHARHEADINGTEXT")
                boolA = dv(int1).Item("boolInclude")
                If boolA = -1 Then
                    boolInclude = True
                Else
                    boolInclude = False
                End If
                boolA = dv(int1).Item("boolGuWu")
                If boolA = -1 Then
                    boolGuWu = True
                Else
                    boolGuWu = False
                End If
                boolA = dv(int1).Item("boolUseStatements")
                If boolA = -1 Then
                    boolStatement = True
                Else
                    boolStatement = False
                End If

                'HereHere

                'boolInclude = dv(int1).Item("boolInclude")
                'boolGuWu = dv(int1).Item("boolGuWu")
                'boolStatement = dv(int1).Item("boolUseStatements")
                charS = NZ(dv(int1).Item("charStatement"), "")
                HLevel = NZ(dv(int1).Item("NUMHEADINGLEVEL"), 1)

                'strF1 = "id_tblConfigBodySections = " & intCBS & " AND id_tblStudies = " & id_tblStudies
                'dr1 = tbl2.Select(strF1)
                'var1 = dr1(0).Item("charSectionName")
                'boolPB = dr1(0).Item("boolPB")

                var1 = dv(int1).Item("CHARSECTIONNAME")
                'boolPB = dv(int1).Item("boolPB")
                'var2 = dv(int1).Item("boolPB")
                'boolPB = dv(int1).Item("boolPB")
                'var2 = NZ(var2, False)
                var2 = dv(int1).Item("boolPB")
                var2 = NZ(var2, False)
                'boolPB = dv(int1).Item("boolPB")
                boolPB = var2
                idWS = dv(int1).Item("ID_TBLWORDSTATEMENTS")

                'int1 = dr1.Length
                'charSectionName = dr1(0).Item("charSectionName") 'remember, a nonbound field named 'charSectionName' has been added to tblReportStatements
                charSectionName = NZ(var1, strN)
                strF2 = "charSectionName = '" & charSectionName & "'"
                dr2 = tbl1.Select(strF2)
                If dr2.Length = 0 Then
                    intTNum = 0
                Else
                    intTNum = dr2(0).Item("intWordTableNumber")
                End If
                frmH.lblProgress.Text = "Generating " & charSectionName & " Section..."
                frmH.lblProgress.Refresh()
                'ctPB = ctPB + 1
                'If ctPB > frmH.pb1.Maximum Then
                '    ctPB = 1
                'End If
                'frmH.pb1.Value = ctPB

                ctPB = ctPB + 1
                If ctPB > frmH.pb1.Maximum Then
                    ctPB = 1
                    If frmH.pb1.Maximum < ctPB Then
                        frmH.pb1.Maximum = 10
                    End If
                End If
                frmH.pb1.Value = ctPB

                frmH.pb1.Refresh()
                Call DoReportBodySection(intCBS, boolInclude, wd, boolFirstReturn, charS, charSectionName, boolGuWu, boolStatement, intTNum, HLevel, strText, boolPB, idWS)

            End If

        Else
            MsgBox("This does not apply to the chosen Table of Contents selection.", MsgBoxStyle.Information, "Not applicable...")
            boolGo = False
            DoIndReportSections = False
        End If

end1:

    End Function

    Sub DoReportBodySection(ByVal intCBS, ByVal boolInclude, ByVal wd, ByVal boolFirstReturn, ByVal charS, ByVal charSectionName, ByVal boolGuWu, ByVal boolStatement, ByVal intTNum, ByVal HLevel, ByVal strText, ByVal boolPB, ByVal idWS)

        Dim wrdSelection As Microsoft.Office.Interop.Word.Selection
        Dim int1 As Short
        Dim intSecs As Short
        Dim intCurSec As Short
        Dim str1 As String
        Dim var1
        Dim boolSkip As Boolean
        Dim boolDoText As Boolean
        'Dim boolInsertPB As Boolean


        boolSkip = False
        boolDoText = True
        'boolInsertPB = False
        var1 = charSectionName
        If StrComp(var1, "Summary", CompareMethod.Text) = 0 Then
            Dim var2
            var2 = var1
        End If


        'wdd.visible = True

        If boolInclude Then

            With wd
                wrdSelection = .selection()

                If boolPB Then
                    If intCBS = 141 Then 'Appendix then
                        wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

                        'linktoprevious=false
                        'delete any contents in the footer
                        intSecs = wd.ActiveDocument.Sections.Count

                        With wd.ActiveDocument.Sections(intSecs)
                            .Footers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False
                            .Footers(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Text = ""
                        End With

                    Else
                        wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
                    End If
                End If


                If HLevel = 0 Then
                Else
                    'frmH.lblProgress.Text = "Changed Heading Style to Heading " & HLevel
                    'frmH.lblProgress.Refresh()
                    wrdSelection.Style = .ActiveDocument.Styles.item("Heading " & HLevel)
                    'frmH.lblProgress.Text = "Finished Heading Style to Heading " & HLevel
                    'frmH.lblProgress.Refresh()
                End If

                Select Case intCBS
                    Case Is = 61 'Cover page
                        boolDoText = False
                    Case Is = 135 'Table of Contents
                        boolDoText = False
                    Case Is = 136 'Table of Tables
                        boolDoText = False
                    Case Is = 137 'Table of Figures
                        boolDoText = False
                    Case Is = 138 'Table of Appendices
                        boolDoText = False
                    Case Is = 139 'Figure Section
                        boolDoText = False
                    Case Is = 140 'Table Section
                        boolDoText = False
                    Case Is = 141 'Appendix Section
                        boolDoText = False
                    Case Is = 341 'Entire report
                        boolDoText = False
                End Select
                'If boolInsertPB And HLevel <> 1 Then
                '    wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak)
                'End If
                If boolDoText Then
                    'replace hyphens
                    str1 = Replace(strText, "-", NBH, 1, -1, CompareMethod.Text)
                    strText = str1
                    wrdSelection.TypeText(Text:=strText)
                    .Selection.TypeParagraph()
                    wrdSelection = .selection()
                    wrdSelection.Style = .ActiveDocument.Styles.item("Normal")
                End If
                'wrdSelection = .selection()
                'wrdSelection.Style = .ActiveDocument.Styles.item("Normal")
            End With

            'wdd.visible = True


            If Len(charS) = 0 Then
                strErrMsg = strErrMsg & Chr(10) & charSectionName
                intErrCount = intErrCount + 1
                boolGuWu = True
            End If
            Select Case intCBS
                Case 3 'QATable
                    'boolSkip = True
            End Select

            'wdd.visible = True
            'If StrComp(strN, "Analytical Run Summary", CompareMethod.Text) = 0 Then
            'ElseIf StrComp(strN, "Summary Table", CompareMethod.Text) = 0 Then
            'ElseIf StrComp(strN, "Analytical Reference Standard", CompareMethod.Text) = 0 Or StrComp(strN, "QA Event Table", CompareMethod.Text) = 0
            ' Or StrComp(strN, "Contributing Personnel", CompareMethod.Text) = 0 Or StrComp(strN, "Cover Page", CompareMethod.Text) = 0 Or 
            'StrComp(strN, "Sample Receipt", CompareMethod.Text) = 0 Or StrComp(strN, "AppTableFig", CompareMethod.Text) = 0 Then

            If boolSkip Then
            Else
                If intTNum = 0 Then
                    Select Case charSectionName
                        Case "Analytical Run Summary" 'ignore
                        Case "Summary Table" 'ignore
                        Case "Analytical Reference Std"
                            Call AnalRefStandards(wd, "None")
                        Case "QA Event Table"
                            Call QATable(wd)
                        Case "Contributing Personnel"
                            Call ContributingPersonnel(wd, True)
                        Case "Sample Receipt Records"
                            Call GuWuSampleReceiptStatement01(wd)
                        Case "AppTableFig" 'ignore

                    End Select
                Else
                    'Call PasteStatement(intCBS, wd, boolGuWu, boolStatement, charS, charSectionName, intTNum, idWS)
                End If
            End If

            'wdd.visible = True


            Select Case intCBS
                Case 5
                    'Call ContributingPersonnel(wd)
                Case 61 'cover page
                    With wd.ActiveDocument.Bookmarks
                        .Add(Range:=wrdSelection.Range, Name:="EOCOV")
                        .ShowHidden = False
                    End With
                    'wd.Selection.InsertBreak(Type:=Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage)

                    'wdd.visible = True


            End Select

        End If
        'End If

    End Sub

    Sub Watson(ByRef wd As Microsoft.Office.Interop.Word.Application)

        Dim wrdselection As Microsoft.Office.Interop.Word.Selection
        Dim tm, bm, tot, initFS
        Dim dgv As DataGridView
        Dim intRows As Short
        Dim Count1 As Short
        Dim Count2 As Short
        Dim var1
        Dim dv As System.Data.DataView
        Dim dv1 As System.Data.DataView
        Dim int1 As Short
        Dim int2 As Short
        Dim arr1(3)
        Dim str1 As String
        Dim intCt As Short
        Dim boolGo As Boolean
        Dim wdobj As Object
        Dim tbl As System.Data.DataTable
        Dim row() As DataRow
        Dim idTR As Int64
        Dim strM As String

        tbl = tblConfigReportTables

        'wdd.visible = True

        dgv = frmH.dgvReportTableConfiguration
        dv = dgv.DataSource
        str1 = "BOOLINCLUDE = TRUE"
        'dv.RowFilter = str1
        intRows = dv.Count

        ctAnalyticalRuns = 0
        ctCalibrStds = 0

        ctPB = 0
        ctPBMax = intRows

        frmH.pb1.Value = 0
        frmH.pb1.Maximum = intRows
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        'wdd.visible = True

        Dim intT As Int64
        Dim intSave
        'intRows = 4 'for testing
        Try
            wd.Options.Pagination = False
            intSave = wd.Options.SaveInterval
            wd.Options.SaveInterval = 0
            wd.Options.CheckGrammarAsYouType = Not wd.Options.CheckGrammarAsYouType
            wd.Options.CheckSpellingAsYouType = Not wd.Options.CheckSpellingAsYouType
        Catch ex As Exception

        End Try

        wrdselection = wd.Selection()
        boolTableSection = True
        'record public section
        intTableSection = wd.Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndSectionNumber)
        With wd.ActiveDocument.Bookmarks
            .Add(Range:=wrdselection.Range, Name:="StartTables")
            .ShowHidden = False
        End With

        Try

            'If gboolReadOnlyTables And intRows > 0 Then
            '    xlROT = New Microsoft.Office.Interop.Excel.Application
            '    xlROT.Workbooks.Add()

            'End If

            Dim strF As String
            Dim strF1 As String
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINCLUDE <> 0"
            Dim rows2() As DataRow = tblReportTable.Select(strF)
            Dim rows1() As DataRow

            intTTot = getIntTTot(False)

            'now get IntStd Table
            Dim tbl2 As System.Data.DataTable
            tbl2 = tblAssignedSamples
            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND BOOLINTSTD <> 0"

            Dim dv10 As System.Data.DataView = New DataView(tbl2, strF, "", DataViewRowState.CurrentRows)
            Dim tbl10 As System.Data.DataTable = dv10.ToTable("a", True, "CHARANALYTE", "BOOLINTSTD")

            intTCur = 0

            boolTableSectionStart = True

            '20170107 LEE:
            'must clear tblQCTables before calling
            Try
                tblQCTables.Clear()
                tblQCTables.AcceptChanges()
            Catch ex As Exception

            End Try

            Dim boolV As Boolean = wd.Visible

            For Count1 = 0 To intRows - 1

                wd.Visible = boolV

                frmH.pb1.Value = Count1
                frmH.pb1.Refresh()
                var1 = dv(Count1).Item("BOOLINCLUDE")
                idTR = dv(Count1).Item("ID_TBLREPORTTABLE")
                boolPlaceHolder = dv(Count1).Item("BOOLPLACEHOLDER") 'global
                If var1 Then

                    intT = dv(Count1).Item("ID_TBLCONFIGREPORTTABLES")
                    row = tbl.Select("ID_TBLCONFIGREPORTTABLES = " & intT)
                    str1 = row(0).Item("CHARTABLENAME")
                    str1 = dv(Count1).Item("CHARHEADINGTEXT")

                    Cursor.Current = Cursors.WaitCursor

                    Try

                        Call PrepareTable(intT, wd, idTR, intT, Count1, str1)

                    Catch ex As Exception

                        If intT = 4 Then
                            strM = "There was a problem preparing table:"
                            strM = strM & ChrW(10) & ChrW(10) & str1 & ChrW(10) & ChrW(10)
                            strM = strM & "It is possible there is an inconsistency in QC configuration for this study." & ChrW(10) & ChrW(10)
                            strM = strM & "Please activate the 'Sample/QC/Calibr Std Details' tab and inspect the QC Levels section for "
                            strM = strM & "QC Level assignment inconsistencies. The user may have to manually assign QC Samples."
                            strM = strM & ChrW(10) & ChrW(10) & "The report generation process will continue."
                            strM = strM & ChrW(10) & ChrW(10)
                            strM = strM & ex.Message
                        Else
                            'strM = "Problem preparing table " & str1 & "."
                            'strM = strM & ChrW(10) & ChrW(10) & "The report generation process will continue."
                            strM = "There was a problem preparing table:"
                            strM = strM & ChrW(10) & ChrW(10) & str1
                            strM = strM & ChrW(10) & ChrW(10) & "The report generation process will continue."
                            strM = strM & ChrW(10) & ChrW(10)
                            strM = strM & ex.Message

                        End If
                        If boolDisableWarnings Then
                        Else
                            MsgBox(strM, MsgBoxStyle.Information, "Problem preparing table...")
                        End If
                        frmH.Refresh()
                        'wdd.visible = True


                    End Try
                    If boolEntireReport Then
                    Else
                        'goto end of document
                        wd.Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    End If
                End If

            Next

            wd.Visible = boolV

            strM = "Finished creating tables ...."
            frmH.lblProgress.Text = strM
            frmH.Refresh()

            'Try
            '    xlROT.ActiveWorkbook.Close(False)
            '    xlROT.Quit()
            'Catch ex As Exception

            'End Try

        Catch ex As Exception

        End Try

        Try
            wd.Options.Pagination = True
            wd.Options.SaveInterval = intSave
            wd.Options.CheckGrammarAsYouType = Not wd.Options.CheckGrammarAsYouType
            wd.Options.CheckSpellingAsYouType = Not wd.Options.CheckSpellingAsYouType
        Catch ex As Exception
            var1 = var1
        End Try

end1:

    End Sub

    Sub HyperlinkAppendices(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal rng1 As Microsoft.Office.Interop.Word.Range)

        System.Windows.Forms.Application.DoEvents()

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim tbl1 As System.Data.DataTable
        Dim mySel As Microsoft.Office.Interop.Word.Selection
        Dim str4 As String
        Dim pos1, pos2
        rng1.Select()
        mySel = wd.Selection
        pos1 = mySel.Start
        pos2 = mySel.End

        tbl1 = tblAppendix
        int1 = tblAppendix.Rows.Count

        Dim pv As Short
        Dim pmax As Short
        Dim bool As Boolean
        pmax = int1
        pv = 0
        frmH.pb1.Value = 0
        frmH.pb1.Maximum = pmax
        bool = frmH.pb1.Visible
        frmH.pb1.Visible = True
        frmH.pb1.Refresh()

        'wdd.visible = True

        Dim strDel As String
        If boolUseHyperlinks Then
            strDel = "_"
        Else
            strDel = ChrW(160)
        End If
        Dim strPut As String

        If boolDoHyperlinks Then

            With wd
                For Count1 = int1 - 1 To 0 Step -1

                    pv = pv + 1
                    frmH.pb1.Value = pv
                    frmH.pb1.Refresh()

                    mySel.Select()
                    Dim var1
                    var1 = tbl1.Rows.Item(Count1).Item("AppendixNumber") 'returns number
                    int2 = var1
                    'str1 = "Appendix " & AppendixLetter(int2)
                    'str1 = "Appendix_" & int2
                    If gboolDisplayAttachments Then
                        str1 = "Attachment_" & int2
                        strPut = "Attachment" & strDel & int2
                    Else
                        str1 = "Appendix_" & int2

                        'for appendix, must find if number or character

                        strPut = "Appendix" & strDel & int2
                    End If


                    str2 = CStr(Count1 + 1) 'InsertCrossReference takes number formatted as text, not letter
                    mySel.Find.ClearFormatting()

                    'wdd.visible = True

                    With mySel.Find
                        .Forward = True
                        .Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue
                        .Format = True
                        .MatchCase = True
                        Do While .Execute(str1)
                            'wd.Selection.InsertCrossReference(ReferenceType:="Appendix", ReferenceKind:=Microsoft.Office.Interop.Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=str2, InsertAsHyperlink:=True, IncludePosition:=False) ', SeparateNumbers:=False, SeparatorString:=" ")

                            If boolUseHyperlinks Then

                                Try
                                    If gboolDisplayAttachments Then
                                        wd.Selection.InsertCrossReference(ReferenceType:="Attachment", ReferenceKind:=Microsoft.Office.Interop.Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=str2, InsertAsHyperlink:=True, IncludePosition:=False) ', SeparateNumbers:=False, SeparatorString:=" ")
                                    Else
                                        wd.Selection.InsertCrossReference(ReferenceType:="Appendix", ReferenceKind:=Microsoft.Office.Interop.Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=str2, InsertAsHyperlink:=True, IncludePosition:=False) ', SeparateNumbers:=False, SeparatorString:=" ")
                                    End If
                                Catch ex As Exception
                                    Exit Do
                                End Try

                                wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdWord, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                                '.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                                If boolBLUEHYPERLINK Then
                                    Call BlueHyperlink(wd)
                                End If
                                wd.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                                'modify field code to conserve formatting
                                'This is done because update will eliminated BlueHyperlink formatting
                                wd.ActiveWindow.View.ShowFieldCodes = True ' Not ActiveWindow.View.ShowFieldCodes
                                wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
                                wd.Selection.TypeText(Text:=" \* MERGEFORMAT ")
                                wd.ActiveWindow.View.ShowFieldCodes = False 'Not ActiveWindow.View.ShowFieldCodes

                            Else

                                wd.Selection.TypeText(Text:=strPut)

                            End If



                        Loop
                    End With
                Next
            End With

            mySel.Find.ClearFormatting()

        End If


        frmH.pb1.Visible = bool

    End Sub

    Sub HyperlinkFigures(ByVal wd As Microsoft.Office.Interop.Word.Application, ByVal intFig As Short, ByVal strFLabel As String)

        Dim str1 As String
        Dim str2 As String
        Dim int1 As Short
        Dim int2 As Short
        Dim Count1 As Short
        Dim strLabel As String
        Dim pos1 As Int64
        Dim strM1 As String
        Dim strM2 As String
        Dim strM3 As String
        Dim boolE As Boolean = False
        Dim strE As String = ""
        Dim rng1 As Microsoft.Office.Interop.Word.Range
        Dim strCap As String

        If InStr(1, strFLabel, "Figure", CompareMethod.Text) > 0 Then
            strLabel = "Figure"
            strCap = strLabel
        ElseIf InStr(1, strFLabel, "Appendix", CompareMethod.Text) > 0 Then

            strLabel = "Appendix"
            If gboolDisplayAttachments Then
                strCap = "Attachment"
            Else
                strCap = strLabel
            End If
        End If

        'wdd.visible = True

        With wd

            Try

                If boolDoHyperlinks And boolUseHyperlinks Then

                    'str1 = "Figure_" & int2
                    str1 = strFLabel & "_" & intFig
                    str2 = CStr(intFig) 'InsertCrossReference takes number formatted as text, not letter
                    Try
                        'wd.Selection.InsertCrossReference(ReferenceType:=strLabel, ReferenceKind:=Microsoft.Office.Interop.Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=str2, InsertAsHyperlink:=True, IncludePosition:=False) ', SeparateNumbers:=False, SeparatorString:=" ")
                        wd.Selection.InsertCrossReference(ReferenceType:=strCap, ReferenceKind:=Microsoft.Office.Interop.Word.WdReferenceKind.wdOnlyLabelAndNumber, ReferenceItem:=str2, InsertAsHyperlink:=True, IncludePosition:=False) ', SeparateNumbers:=False, SeparatorString:=" ")
                    Catch ex As Exception
                        strE = ex.Message
                        boolE = True
                    End Try
                    pos1 = wd.Selection.End
                    rng1 = wd.Selection.Range

                    If boolE Then
                        wd.Selection.TypeText(Text:=strE)
                        pos1 = wd.Selection.End
                        rng1 = wd.Selection.Range
                    Else

                        .Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdWord, Count:=1, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)

                        'format font blue
                        'wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
                        '.Selection.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                        Try
                            If boolBLUEHYPERLINK Then
                                Call BlueHyperlink(wd)
                            End If
                        Catch ex As Exception
                            strM3 = ex.Message
                        End Try
                        wd.Selection.MoveRight(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=1)

                        'modify field code to conserve formatting
                        'This is done because update will eliminated BlueHyperlink formatting
                        wd.ActiveWindow.View.ShowFieldCodes = True ' Not ActiveWindow.View.ShowFieldCodes
                        wd.Selection.MoveLeft(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdCharacter, Count:=2)
                        wd.Selection.TypeText(Text:=" \* MERGEFORMAT ")
                        wd.ActiveWindow.View.ShowFieldCodes = False 'Not ActiveWindow.View.ShowFieldCodes
                    End If

                    'go back to pos1
                    rng1.Select()

                End If

            Catch ex As Exception
                Dim var1
                strM1 = "Problem in HyperlinkFigures" & ChrW(10) & ChrW(10)
                strM1 = strM1 & ex.Message
            End Try

        End With

        'frmH.pb1.Visible = bool


    End Sub

    Sub EnterFooters(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim ctSec As Integer
        Dim var1
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim str5 As String
        Dim int1 As Long
        Dim int2 As Long
        Dim int3 As Long
        Dim int4 As Short
        Dim int5 As Short
        Dim int6 As Short
        Dim wds As Microsoft.Office.Interop.Word.Selection
        Dim pos1, pos2
        Dim pb1VO As Short
        Dim pb1MO As Short
        Dim boolF As Boolean = False

        'Header rules:
        '1st page has logo, no study info
        '2nd page has no logo, has study info

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)



        frmH.pb1.Maximum = ctPBMax

        Dim rm, lm, rt

        pb1VO = frmH.pb1.Value
        pb1MO = frmH.pb1.Maximum

        With wd

            GoTo end2

            'get project number
            Dim dgv As DataGridView
            Dim dv As System.Data.DataView
            Dim dr1() As DataRow
            Dim tbl As System.Data.DataTable
            Dim strProj As String
            Dim strAddressTitle As String
            Dim strStudy As String
            Dim strFor As String
            Dim strF As String

            'dgv = frmH.dgvDataCompany
            'dv = dgv.DataSource
            'int1 = FindRowDV("Corporate Study/Project Number", dv)
            'str1 = dv(int1).Item("Value")
            'strProj = "Project Number: " & str1

            ''get submittedto address
            'str2 = frmH.cbxSubmittedTo.Text
            'str1 = GetAddressTitle(str2, tblCorporateAddresses)
            ''Replace carriage returns with space
            'strAddressTitle = Replace(str1, Chr(10), " ", 1, -1, CompareMethod.Text)
            'str2 = "Sponsor Study Number"
            'int1 = FindRowDV(str2, dv)
            'strStudy = dv(int1).Item("Value")
            'str1 = "For " & strAddressTitle
            'strFor = str1 & " Study " & strStudy

            Dim CHARFLT As String
            Dim CHARFRT As String
            Dim CHARFLB As String
            Dim CHARFRB As String

            Dim boolF2 As Boolean = True
            Dim boolGo As Boolean = False
            Dim boolPN As Boolean = False 'page number
            Dim boolTP As Boolean = False 'total pages
            Dim rng As Microsoft.Office.Interop.Word.Range
            Dim intSPN As Short
            Dim intSTP As Short
            Dim intLPN As Short
            Dim intLTP As Short
            Dim strCH As String
            Dim strA As String
            Dim strB As String

            intLPN = Len("[PAGENUMBER]")
            intLTP = Len("[TOTALPAGES]")

            rng = wd.Selection.Range

            If id_tblReports = 0 Then 'go get it
                id_tblReports = frmH.dgvReports("ID_TBLREPORTS", frmH.dgvReports.CurrentRow.Index).Value
            End If

            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTS = " & id_tblReports
            tbl = tblReportHeaders
            dr1 = tbl.Select(strF)


            CHARFLT = NZ(dr1(0).Item("CHARFLT"), "")
            If Len(CHARFLT) = 0 Then
            Else
                CHARFLT = SearchReplace(wd, "Report Body", rng, True, CHARFLT, 1, 0, 0, False, True, True, 0)
            End If

            CHARFRT = NZ(dr1(0).Item("CHARFRT"), "")
            If Len(CHARFRT) = 0 Then
            Else
                CHARFRT = SearchReplace(wd, "Report Body", rng, True, CHARFRT, 1, 0, 0, False, True, True, 0)
            End If

            CHARFLB = NZ(dr1(0).Item("CHARFLB"), "")
            If Len(CHARFLB) = 0 Then
            Else
                CHARFLB = SearchReplace(wd, "Report Body", rng, True, CHARFLB, 1, 0, 0, False, True, True, 0)
            End If

            CHARFRB = NZ(dr1(0).Item("CHARFRB"), "")
            If Len(CHARFRB) = 0 Then
            Else
                CHARFRB = SearchReplace(wd, "Report Body", rng, True, CHARFRB, 1, 0, 0, False, True, True, 0)
            End If

            If Len(CHARFLT) = 0 And Len(CHARFRT) = 0 And Len(CHARFLB) = 0 And Len(CHARFRB) = 0 Then
                boolF = False
            Else
                boolF = True
                If Len(CHARFLB) = 0 And Len(CHARFRB) = 0 Then
                    boolF2 = False
                End If
            End If

            Dim intSec As Short = 2
            ctSec = .ActiveDocument.Sections.Count
            If ctSec = 1 Then
                intSec = 1
            End If

            If boolF Then 'continue
            Else
                GoTo end2
            End If

            'lm = .Selection.PageSetup.LeftMargin

            ''in order to make first page different and keep the header content from the template, do the following
            'If .ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
            '    .ActiveWindow.Panes(2).Close()
            'End If
            'If .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
            '    .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            'End If
            '.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader

            'Try
            '    .ActiveWindow.ActivePane.View.PreviousHeaderFooter()
            'Catch ex As Exception

            'End Try


            'wdd.visible = True

            'goto 2nd section
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToSection, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst, Count:=intSec, Name:="")
            .ActiveDocument.Sections(intSec).Footers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False

            'move to header of next section
            '.ActiveWindow.ActivePane.View.NextHeaderFooter()
            Try
                If .ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
                    .ActiveWindow.Panes(2).Close()
                End If
                If .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                    .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                End If
                .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter

                'select all
                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
                'delete any extra rows and graphics
                pos1 = .Selection.Start
                .Selection.WholeStory()
                .Selection.Delete()

                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

                .Selection.ParagraphFormat.TabStops.ClearAll()
                'set righttab at right margin
                'rm = .Selection.PageSetup.RightMargin
                rm = .ActiveDocument.Sections(intSec).PageSetup.RightMargin
                lm = .ActiveDocument.Sections(intSec).PageSetup.LeftMargin

                If .ActiveDocument.Sections(intSec).PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                    rt = (11 * 72) - lm - rm
                    '.Selection.ParagraphFormat.TabStops.Add(Position:=684, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                Else
                    rt = (8.5 * 72) - lm - rm
                    '.Selection.ParagraphFormat.TabStops.Add(Position:=432, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                End If

                ''debug.writeline("lm: " & lm & ", rm: " & rm & ", rt: " & rt)
                .Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                boolGo = False
                boolPN = False
                boolTP = False
                strCH = CHARFLT
                intSTP = 0
                intSPN = 0
                If InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text) > 0 Then
                    intSPN = InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text)
                    boolPN = True
                    boolGo = True
                End If
                If InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text) > 0 Then
                    intSTP = InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text)
                    boolTP = True
                    boolGo = True
                End If

                If boolGo Then
                    If boolPN And boolTP Then
                        If intSPN > intSTP Then
                            'str1 & [PN] & str2 & [TP] & str3
                            int1 = intSTP + intLTP
                            int2 = intSPN + intLPN
                            int3 = Len(strCH)
                            int4 = intSTP
                            int5 = intSPN
                            str1 = Mid(strCH, 1, int4 - 1)
                            str2 = Mid(strCH, int1, int5 - int1)
                            strB = "[PAGENUMBER]"
                            strA = "[TOTALPAGES]"
                            If int2 >= int3 Then
                                str3 = ""
                            Else
                                str3 = Mid(strCH, int2, Len(strCH) - int2)
                            End If
                            .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                            .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                            .Selection.TypeText(Text:=str3)

                        Else
                            'str1 & [PN] & str2 & [TP] & str3
                            int1 = intSPN + intLPN
                            int2 = intSTP + intLTP
                            int3 = Len(strCH)
                            int4 = intSPN
                            int5 = intSTP
                            str1 = Mid(strCH, 1, int4 - 1)
                            str2 = Mid(strCH, int1, int5 - int1)
                            strA = "[PAGENUMBER]"
                            strB = "[TOTALPAGES]"
                            If int2 >= int3 Then
                                str3 = ""
                            Else
                                str3 = Mid(strCH, int2, Len(strCH) - int2)
                            End If
                            .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                            .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                            .Selection.TypeText(Text:=str3)

                        End If
                    ElseIf boolPN And boolTP = False Then
                        strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=strCH)
                        wds = .Selection
                        .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                    ElseIf boolPN = False And boolTP Then
                        strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=strCH)
                        wds = .Selection
                        .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                    End If
                Else
                    .Selection.TypeText(Text:=strCH)
                End If

                .Selection.TypeText(Text:=vbTab)

                boolGo = False
                boolPN = False
                boolTP = False
                strCH = CHARFRT
                intSTP = 0
                intSPN = 0
                If InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text) > 0 Then
                    intSPN = InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text)
                    boolPN = True
                    boolGo = True
                    'strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                    '.Selection.TypeText(Text:=strCH)
                    'wds = .Selection
                    '.Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                    ''.Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldEmpty, Text:="PAGE  ", PreserveFormatting:=True)
                End If
                If InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text) > 0 Then
                    intSTP = InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text)
                    boolTP = True
                    boolGo = True
                    'strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                    '.Selection.TypeText(Text:=strCH)
                    'wds = .Selection
                    '.Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                End If

                If boolGo Then
                    If boolPN And boolTP Then
                        If intSPN > intSTP Then
                            'str1 & [PN] & str2 & [TP] & str3
                            int1 = intSTP + intLTP
                            int2 = intSPN + intLPN
                            int3 = Len(strCH)
                            int4 = intSTP
                            int5 = intSPN
                            str1 = Mid(strCH, 1, int4 - 1)
                            str2 = Mid(strCH, int1, int5 - int1)
                            strB = "[PAGENUMBER]"
                            strA = "[TOTALPAGES]"
                            If int2 >= int3 Then
                                str3 = ""
                            Else
                                str3 = Mid(strCH, int2, Len(strCH) - int2)
                            End If
                            .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                            .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                            .Selection.TypeText(Text:=str3)

                        Else
                            'str1 & [PN] & str2 & [TP] & str3
                            int1 = intSPN + intLPN
                            int2 = intSTP + intLTP
                            int3 = Len(strCH)
                            int4 = intSPN
                            int5 = intSTP
                            str1 = Mid(strCH, 1, int4 - 1)
                            str2 = Mid(strCH, int1, int5 - int1)
                            strA = "[PAGENUMBER]"
                            strB = "[TOTALPAGES]"
                            If int2 >= int3 Then
                                str3 = ""
                            Else
                                str3 = Mid(strCH, int2, Len(strCH) - int2)
                            End If
                            .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                            .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                            .Selection.TypeText(Text:=str3)

                        End If
                    ElseIf boolPN And boolTP = False Then
                        strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=strCH)
                        wds = .Selection
                        .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                    ElseIf boolPN = False And boolTP Then
                        strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=strCH)
                        wds = .Selection
                        .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                    End If
                Else
                    .Selection.TypeText(Text:=strCH)
                End If

                If boolF2 Then
                    .Selection.TypeParagraph()

                    boolGo = False
                    boolPN = False
                    boolTP = False
                    strCH = CHARFLB
                    intSTP = 0
                    intSPN = 0
                    If InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text) > 0 Then
                        intSPN = InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text)
                        boolPN = True
                        boolGo = True
                    End If
                    If InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text) > 0 Then
                        intSTP = InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text)
                        boolTP = True
                        boolGo = True
                    End If

                    If boolGo Then
                        If boolPN And boolTP Then
                            If intSPN > intSTP Then
                                'str1 & [PN] & str2 & [TP] & str3
                                int1 = intSTP + intLTP
                                int2 = intSPN + intLPN
                                int3 = Len(strCH)
                                int4 = intSTP
                                int5 = intSPN
                                str1 = Mid(strCH, 1, int4 - 1)
                                str2 = Mid(strCH, int1, int5 - int1)
                                strB = "[PAGENUMBER]"
                                strA = "[TOTALPAGES]"
                                If int2 >= int3 Then
                                    str3 = ""
                                Else
                                    str3 = Mid(strCH, int2, Len(strCH) - int2)
                                End If
                                .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                                .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                                .Selection.TypeText(Text:=str3)

                            Else
                                'str1 & [PN] & str2 & [TP] & str3
                                int1 = intSPN + intLPN
                                int2 = intSTP + intLTP
                                int3 = Len(strCH)
                                int4 = intSPN
                                int5 = intSTP
                                str1 = Mid(strCH, 1, int4 - 1)
                                str2 = Mid(strCH, int1, int5 - int1)
                                strA = "[PAGENUMBER]"
                                strB = "[TOTALPAGES]"
                                If int2 >= int3 Then
                                    str3 = ""
                                Else
                                    str3 = Mid(strCH, int2, Len(strCH) - int2)
                                End If
                                .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                                .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                                .Selection.TypeText(Text:=str3)

                            End If
                        ElseIf boolPN And boolTP = False Then
                            strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=strCH)
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                        ElseIf boolPN = False And boolTP Then
                            strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=strCH)
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                        End If
                    Else
                        .Selection.TypeText(Text:=strCH)
                    End If

                    .Selection.TypeText(Text:=vbTab)

                    boolGo = False
                    boolPN = False
                    boolTP = False
                    strCH = CHARFRB
                    intSTP = 0
                    intSPN = 0
                    If InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text) > 0 Then
                        intSPN = InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text)
                        boolPN = True
                        boolGo = True

                    End If
                    If InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text) > 0 Then
                        intSTP = InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text)
                        boolTP = True
                        boolGo = True
                    End If

                    If boolGo Then
                        If boolPN And boolTP Then
                            If intSPN > intSTP Then
                                'str1 & [PN] & str2 & [TP] & str3
                                int1 = intSTP + intLTP
                                int2 = intSPN + intLPN
                                int3 = Len(strCH)
                                int4 = intSTP
                                int5 = intSPN
                                str1 = Mid(strCH, 1, int4 - 1)
                                str2 = Mid(strCH, int1, int5 - int1)
                                strB = "[PAGENUMBER]"
                                strA = "[TOTALPAGES]"
                                If int2 >= int3 Then
                                    str3 = ""
                                Else
                                    str3 = Mid(strCH, int2, Len(strCH) - int2)
                                End If
                                .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                                .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                                .Selection.TypeText(Text:=str3)

                            Else
                                'str1 & [PN] & str2 & [TP] & str3
                                int1 = intSPN + intLPN
                                int2 = intSTP + intLTP
                                int3 = Len(strCH)
                                int4 = intSPN
                                int5 = intSTP
                                str1 = Mid(strCH, 1, int4 - 1)
                                str2 = Mid(strCH, int1, int5 - int1)
                                strA = "[PAGENUMBER]"
                                strB = "[TOTALPAGES]"
                                If int2 >= int3 Then
                                    str3 = ""
                                Else
                                    str3 = Mid(strCH, int2, Len(strCH) - int2)
                                End If
                                .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                                .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                                .Selection.TypeText(Text:=str3)

                            End If
                        ElseIf boolPN And boolTP = False Then
                            strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=strCH)
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                        ElseIf boolPN = False And boolTP Then
                            strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=strCH)
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                        End If
                    Else
                        .Selection.TypeText(Text:=strCH)
                    End If

                End If
                .Selection.TypeParagraph()
                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

                'wdd.visible = True

                int3 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndSectionNumber)
                'loop through all section headers

                frmH.pb1.Value = 0
                frmH.pb1.Maximum = ctSec
                frmH.pb1.Value = int3

                'Try
                '    .ActiveWindow.ActivePane.View.NextHeaderFooter()
                '    'ensure you're still in header
                '    .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader

                'Catch ex As Exception
                '    GoTo end1
                'End Try

                For Count1 = 3 To ctSec
                    int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndSectionNumber)
                    frmH.lblProgress.Text = "Creating Report Footers (" & Count1 & " of " & ctSec & ")..."
                    frmH.lblProgress.Refresh()
                    frmH.pb1.Value = Count1
                    frmH.pb1.Refresh()

                    'wdd.visible = True

                    .ActiveDocument.Sections(Count1).Footers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False

                    '.Selection.HeaderFooter.LinkToPrevious = False 'this action returns you to the previous header
                    '.ActiveWindow.ActivePane.View.NextHeaderFooter()
                    'goto correct page
                    ''ensure headerfooter index hasn't change
                    '.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

                    '.Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToSection, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst, Count:=Count1, Name:="")
                    '.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader

                    'set tabs according to pagesetup
                    'rm = .Selection.PageSetup.LeftMargin
                    'lm = .Selection.PageSetup.RightMargin

                    rm = .ActiveDocument.Sections(Count1).PageSetup.RightMargin
                    lm = .ActiveDocument.Sections(Count1).PageSetup.LeftMargin

                    If .ActiveDocument.Sections(Count1).PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                        rt = (11 * 72) - lm - rm
                        '.Selection.ParagraphFormat.TabStops.Add(Position:=684, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    Else
                        rt = (8.5 * 72) - lm - rm
                        '.Selection.ParagraphFormat.TabStops.Add(Position:=432, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    End If
                    .ActiveDocument.Sections(Count1).Footers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.TabStops.ClearAll()
                    .ActiveDocument.Sections(Count1).Footers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                    'wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)
                    '.Selection.WholeStory()
                    '.Selection.ParagraphFormat.TabStops.ClearAll()
                    '.Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                    'If Count1 = ctSec Then
                    'Else
                    '    .ActiveWindow.ActivePane.View.NextHeaderFooter()
                    'End If
                Next

            Catch ex As Exception

            End Try

end2:

            ctSec = .ActiveDocument.Sections.Count

            Try
                .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
            Catch ex As Exception

            End Try

            Dim pw, ph

            'wdd.visible = True
            If boolF Then
            Else 'do search replace in header

                'wdd.visible = True

                'goto 1st section
                intSec = 1
                Try
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToSection, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst, Count:=intSec, Name:="")
                    '.ActiveDocument.Sections(intSec).Footers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False
                Catch ex As Exception

                End Try


                ''move to header of next section
                '.ActiveWindow.ActivePane.View.NextHeaderFooter()

                'wdd.visible = True

                Try

                    If .ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
                        .ActiveWindow.Panes(2).Close()
                    End If
                    If .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                        .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                    End If
                    .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageFooter

                    Dim boolLinked As Boolean = False
                    frmH.pb1.Value = 0
                    frmH.pb1.Maximum = intSec
                    For Count1 = 1 To ctSec

                        'frmH.lblProgress.Text = "Search/Replacing Report Footers (" & Count1 & " of " & ctSec & ")..."
                        'frmH.lblProgress.Refresh()
                        'frmH.pb1.Value = Count1
                        'frmH.pb1.Refresh()

                        Try

                            If Count1 = 1 Then

                                If .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable) Then
                                    'format table
                                    .Selection.Tables.Item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                                End If

                                .Selection.WholeStory()

                                'do search replace
                                rng = wd.Selection.Range
                                Dim boolAA As Boolean

                                'wdd.visible = True

                                'don't need this
                                ''boolA = SearchReplace(wd, "Report Body", rng, True, CHARHLB, 1, 0, 0, FALSE)
                                'Try
                                '    boolAA = SearchReplace(wd, "Report Body", rng, False, "aa", 1, 0, 0, True)
                                'Catch ex As Exception
                                '    Dim varaaa
                                '    varaaa = "a"
                                'End Try

                            Else
                                .ActiveWindow.ActivePane.View.NextHeaderFooter()
                            End If

                            boolLinked = .Selection.HeaderFooter.LinkToPrevious
                            If boolLinked And Count1 <> ctSec Then
                            Else


                                frmH.lblProgress.Text = "Search/Replacing Report Footers (" & Count1 & " of " & ctSec & ")..."
                                frmH.lblProgress.Refresh()
                                frmH.pb1.Value = Count1
                                frmH.pb1.Refresh()

                                If .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable) Then
                                    'format table
                                    .Selection.Tables.Item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                                End If

                                .Selection.WholeStory()

                                'do search replace
                                rng = wd.Selection.Range
                                Dim boolAB As Boolean

                                'wdd.visible = True


                                'boolA = SearchReplace(wd, "Report Body", rng, True, CHARHLB, 1, 0, 0, FALSE)
                                Try
                                    boolAB = SearchReplace(wd, "Report Body", rng, False, "aa", 1, 0, 0, True, True, True, 0)
                                Catch ex As Exception
                                    Dim varaaa
                                    varaaa = "a"
                                End Try

                                '20180329 LEE:
                                'need to start doing custom field codes
                                Call SearchReplaceCustomFieldCode(wd)

                                'don't do this anymore. It's done in LinkPrevious
                                'int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndSectionNumber)
                                'frmH.lblProgress.Text = "Configuring Report Footers (" & Count1 & " of " & ctSec & ")..."
                                'frmH.lblProgress.Refresh()
                                'frmH.pb1.Value = Count1
                                'frmH.pb1.Refresh()

                                'rm = .ActiveDocument.Sections(Count1).PageSetup.RightMargin
                                'lm = .ActiveDocument.Sections(Count1).PageSetup.LeftMargin
                                'pw = .ActiveDocument.Sections(Count1).PageSetup.PageWidth
                                'ph = .ActiveDocument.Sections(Count1).PageSetup.PageHeight


                                ''wdd.visible = True

                                'If Count1 = 1 Then
                                'Else
                                '    .ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False
                                'End If



                                'If .ActiveDocument.Sections(Count1).PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                                '    rt = pw - lm - rm
                                '    '.Selection.ParagraphFormat.TabStops.Add(Position:=684, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                                'Else
                                '    rt = pw - lm - rm
                                '    '.Selection.ParagraphFormat.TabStops.Add(Position:=432, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                                'End If
                                'wd.Selection.WholeStory()
                                '.ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.TabStops.ClearAll()
                                '.ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                            End If

                        Catch ex As Exception

                        End Try

                    Next


                    frmH.lblProgress.Text = "Search/Replacing Report Footers (" & ctSec & " of " & ctSec & ")..."
                    frmH.lblProgress.Refresh()
                    frmH.pb1.Value = 0
                    frmH.pb1.Maximum = ctSec
                    frmH.pb1.Value = ctSec
                    frmH.pb1.Refresh()

                    Try
                        .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
                    Catch ex As Exception

                    End Try

                Catch ex As Exception

                End Try
            End If

end1:

            Try
                .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
            Catch ex As Exception

            End Try

        End With

        frmH.pb1.Value = 0
        frmH.pb1.Maximum = pb1MO
        frmH.pb1.Value = pb1VO



    End Sub

    Sub EnterHeaders(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim ctSec As Integer
        Dim var1
        Dim Count1 As Short
        Dim Count2 As Short
        Dim Count3 As Short
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim str5 As String
        Dim int1 As Long
        Dim int2 As Long
        Dim int3 As Long
        Dim int4 As Short
        Dim int5 As Short
        Dim int6 As Short
        Dim wds As Microsoft.Office.Interop.Word.Selection
        Dim pos1, pos2
        Dim pb1VO As Short
        Dim pb1MO As Short
        Dim rm, lm, rt

        'Header rules:
        '1st page has logo, no study info
        '2nd page has no logo, has study info

        wd.Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

        frmH.pb1.Maximum = ctPBMax

        'wdd.visible = True


        pb1VO = frmH.pb1.Value
        pb1MO = frmH.pb1.Maximum

        Dim boolH As Boolean = False

        boolH = True

        With wd

            GoTo end2

            'get project number
            Dim dgv As DataGridView
            Dim dv As System.Data.DataView
            Dim dr1() As DataRow
            Dim tbl As System.Data.DataTable
            Dim strProj As String
            Dim strAddressTitle As String
            Dim strStudy As String
            Dim strFor As String
            Dim strF As String

            'dgv = frmH.dgvDataCompany
            'dv = dgv.DataSource
            'int1 = FindRowDV("Corporate Study/Project Number", dv)
            'str1 = dv(int1).Item("Value")
            'strProj = "Project Number: " & str1

            ''get submittedto address
            'str2 = frmH.cbxSubmittedTo.Text
            'str1 = GetAddressTitle(str2, tblCorporateAddresses)
            ''Replace carriage returns with space
            'strAddressTitle = Replace(str1, Chr(10), " ", 1, -1, CompareMethod.Text)
            'str2 = "Sponsor Study Number"
            'int1 = FindRowDV(str2, dv)
            'strStudy = dv(int1).Item("Value")
            'str1 = "For " & strAddressTitle
            'strFor = str1 & " Study " & strStudy

            Dim CHARHLT As String
            Dim CHARHRT As String
            Dim CHARHLB As String
            Dim CHARHRB As String


            Dim boolH2 As Boolean = True
            Dim boolGo As Boolean = False
            Dim boolPN As Boolean = False 'page number
            Dim boolTP As Boolean = False 'total pages
            Dim rng As Microsoft.Office.Interop.Word.Range
            Dim intSPN As Short
            Dim intSTP As Short
            Dim intLPN As Short
            Dim intLTP As Short
            Dim strCH As String
            Dim strA As String
            Dim strB As String

            intLPN = Len("[PAGENUMBER]")
            intLTP = Len("[TOTALPAGES]")

            rng = wd.Selection.Range

            If id_tblReports = 0 Then 'go get it
                id_tblReports = frmH.dgvReports("ID_TBLREPORTS", frmH.dgvReports.CurrentRow.Index).Value
            End If

            strF = "ID_TBLSTUDIES = " & id_tblStudies & " AND ID_TBLREPORTS = " & id_tblReports
            tbl = tblReportHeaders
            dr1 = tbl.Select(strF)
            Dim intRows As Short
            intRows = dr1.Length

            CHARHLT = NZ(dr1(0).Item("CHARHLT"), "")
            If Len(CHARHLT) = 0 Then
            Else
                CHARHLT = SearchReplace(wd, "Report Body", rng, True, CHARHLT, 1, 0, 0, False, True, True, 0)
            End If

            CHARHRT = NZ(dr1(0).Item("CHARHRT"), "")
            If Len(CHARHRT) = 0 Then
            Else
                CHARHRT = SearchReplace(wd, "Report Body", rng, True, CHARHRT, 1, 0, 0, False, True, True, 0)
            End If

            CHARHLB = NZ(dr1(0).Item("CHARHLB"), "")
            If Len(CHARHLB) = 0 Then
            Else
                CHARHLB = SearchReplace(wd, "Report Body", rng, True, CHARHLB, 1, 0, 0, False, True, True, 0)
            End If

            CHARHRB = NZ(dr1(0).Item("CHARHRB"), "")
            If Len(CHARHRB) = 0 Then
            Else
                CHARHRB = SearchReplace(wd, "Report Body", rng, True, CHARHRB, 1, 0, 0, False, True, True, 0)
            End If

            If Len(CHARHLT) = 0 And Len(CHARHRT) = 0 And Len(CHARHLB) = 0 And Len(CHARHRB) = 0 Then
                boolH = False
            Else
                boolH = True
                If Len(CHARHLB) = 0 And Len(CHARHRB) = 0 Then
                    boolH2 = False
                End If
            End If

            Dim intSec As Short = 2
            ctSec = .ActiveDocument.Sections.Count
            If ctSec = 1 Then
                intSec = 1
            End If

            If boolH Then 'continue
            Else
                GoTo end2
            End If

            'wdd.visible = True

            'goto 2nd section
            .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToSection, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst, Count:=intSec, Name:="")
            .ActiveDocument.Sections(intSec).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False

            'move to header of next section
            '.ActiveWindow.ActivePane.View.NextHeaderFooter()
            Try
                If .ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
                    .ActiveWindow.Panes(2).Close()
                End If
                If .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                    .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                End If
                .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader

                'select all
                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
                'delete any extra rows and graphics
                pos1 = .Selection.Start
                .Selection.WholeStory()
                .Selection.Delete()

                .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft

                .Selection.ParagraphFormat.TabStops.ClearAll()
                'set righttab at right margin
                'rm = .Selection.PageSetup.RightMargin
                rm = .ActiveDocument.Sections(intSec).PageSetup.RightMargin
                lm = .ActiveDocument.Sections(intSec).PageSetup.LeftMargin

                If .ActiveDocument.Sections(intSec).PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                    rt = (11 * 72) - lm - rm
                    '.Selection.ParagraphFormat.TabStops.Add(Position:=684, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                Else
                    rt = (8.5 * 72) - lm - rm
                    '.Selection.ParagraphFormat.TabStops.Add(Position:=432, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                End If

                ''debug.writeline("lm: " & lm & ", rm: " & rm & ", rt: " & rt)
                .Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                boolGo = False
                boolPN = False
                boolTP = False
                strCH = CHARHLT
                intSTP = 0
                intSPN = 0
                If InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text) > 0 Then
                    intSPN = InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text)
                    boolPN = True
                    boolGo = True
                End If
                If InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text) > 0 Then
                    intSTP = InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text)
                    boolTP = True
                    boolGo = True
                End If

                If boolGo Then
                    If boolPN And boolTP Then
                        If intSPN > intSTP Then
                            'str1 & [PN] & str2 & [TP] & str3
                            int1 = intSTP + intLTP
                            int2 = intSPN + intLPN
                            int3 = Len(strCH)
                            int4 = intSTP
                            int5 = intSPN
                            str1 = Mid(strCH, 1, int4 - 1)
                            str2 = Mid(strCH, int1, int5 - int1)
                            strB = "[PAGENUMBER]"
                            strA = "[TOTALPAGES]"
                            If int2 >= int3 Then
                                str3 = ""
                            Else
                                str3 = Mid(strCH, int2, Len(strCH) - int2)
                            End If
                            .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                            .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                            .Selection.TypeText(Text:=str3)

                        Else
                            'str1 & [PN] & str2 & [TP] & str3
                            int1 = intSPN + intLPN
                            int2 = intSTP + intLTP
                            int3 = Len(strCH)
                            int4 = intSPN
                            int5 = intSTP
                            str1 = Mid(strCH, 1, int4 - 1)
                            str2 = Mid(strCH, int1, int5 - int1)
                            strA = "[PAGENUMBER]"
                            strB = "[TOTALPAGES]"
                            If int2 >= int3 Then
                                str3 = ""
                            Else
                                str3 = Mid(strCH, int2, Len(strCH) - int2)
                            End If
                            .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                            .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                            .Selection.TypeText(Text:=str3)

                        End If
                    ElseIf boolPN And boolTP = False Then
                        strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=strCH)
                        wds = .Selection
                        .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                    ElseIf boolPN = False And boolTP Then
                        strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=strCH)
                        wds = .Selection
                        .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                    End If
                Else
                    .Selection.TypeText(Text:=strCH)
                End If

                .Selection.TypeText(Text:=vbTab)

                boolGo = False
                boolPN = False
                boolTP = False
                strCH = CHARHRT
                intSTP = 0
                intSPN = 0
                If InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text) > 0 Then
                    intSPN = InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text)
                    boolPN = True
                    boolGo = True
                    'strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                    '.Selection.TypeText(Text:=strCH)
                    'wds = .Selection
                    '.Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                    ''.Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldEmpty, Text:="PAGE  ", PreserveFormatting:=True)
                End If
                If InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text) > 0 Then
                    intSTP = InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text)
                    boolTP = True
                    boolGo = True
                    'strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                    '.Selection.TypeText(Text:=strCH)
                    'wds = .Selection
                    '.Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                End If

                If boolGo Then
                    If boolPN And boolTP Then
                        If intSPN > intSTP Then
                            'str1 & [PN] & str2 & [TP] & str3
                            int1 = intSTP + intLTP
                            int2 = intSPN + intLPN
                            int3 = Len(strCH)
                            int4 = intSTP
                            int5 = intSPN
                            str1 = Mid(strCH, 1, int4 - 1)
                            str2 = Mid(strCH, int1, int5 - int1)
                            strB = "[PAGENUMBER]"
                            strA = "[TOTALPAGES]"
                            If int2 >= int3 Then
                                str3 = ""
                            Else
                                str3 = Mid(strCH, int2, Len(strCH) - int2)
                            End If
                            .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                            .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                            .Selection.TypeText(Text:=str3)

                        Else
                            'str1 & [PN] & str2 & [TP] & str3
                            int1 = intSPN + intLPN
                            int2 = intSTP + intLTP
                            int3 = Len(strCH)
                            int4 = intSPN
                            int5 = intSTP
                            str1 = Mid(strCH, 1, int4 - 1)
                            str2 = Mid(strCH, int1, int5 - int1)
                            strA = "[PAGENUMBER]"
                            strB = "[TOTALPAGES]"
                            If int2 >= int3 Then
                                str3 = ""
                            Else
                                str3 = Mid(strCH, int2, Len(strCH) - int2)
                            End If
                            .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                            .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                            .Selection.TypeText(Text:=str3)

                        End If
                    ElseIf boolPN And boolTP = False Then
                        strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=strCH)
                        wds = .Selection
                        .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                    ElseIf boolPN = False And boolTP Then
                        strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                        .Selection.TypeText(Text:=strCH)
                        wds = .Selection
                        .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                    End If
                Else
                    .Selection.TypeText(Text:=strCH)
                End If

                If boolH2 Then
                    .Selection.TypeParagraph()

                    boolGo = False
                    boolPN = False
                    boolTP = False
                    strCH = CHARHLB
                    intSTP = 0
                    intSPN = 0
                    If InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text) > 0 Then
                        intSPN = InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text)
                        boolPN = True
                        boolGo = True
                    End If
                    If InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text) > 0 Then
                        intSTP = InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text)
                        boolTP = True
                        boolGo = True
                    End If

                    If boolGo Then
                        If boolPN And boolTP Then
                            If intSPN > intSTP Then
                                'str1 & [PN] & str2 & [TP] & str3
                                int1 = intSTP + intLTP
                                int2 = intSPN + intLPN
                                int3 = Len(strCH)
                                int4 = intSTP
                                int5 = intSPN
                                str1 = Mid(strCH, 1, int4 - 1)
                                str2 = Mid(strCH, int1, int5 - int1)
                                strB = "[PAGENUMBER]"
                                strA = "[TOTALPAGES]"
                                If int2 >= int3 Then
                                    str3 = ""
                                Else
                                    str3 = Mid(strCH, int2, Len(strCH) - int2)
                                End If
                                .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                                .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                                .Selection.TypeText(Text:=str3)

                            Else
                                'str1 & [PN] & str2 & [TP] & str3
                                int1 = intSPN + intLPN
                                int2 = intSTP + intLTP
                                int3 = Len(strCH)
                                int4 = intSPN
                                int5 = intSTP
                                str1 = Mid(strCH, 1, int4 - 1)
                                str2 = Mid(strCH, int1, int5 - int1)
                                strA = "[PAGENUMBER]"
                                strB = "[TOTALPAGES]"
                                If int2 >= int3 Then
                                    str3 = ""
                                Else
                                    str3 = Mid(strCH, int2, Len(strCH) - int2)
                                End If
                                .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                                .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                                .Selection.TypeText(Text:=str3)

                            End If
                        ElseIf boolPN And boolTP = False Then
                            strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=strCH)
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                        ElseIf boolPN = False And boolTP Then
                            strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=strCH)
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                        End If
                    Else
                        .Selection.TypeText(Text:=strCH)
                    End If

                    .Selection.TypeText(Text:=vbTab)

                    boolGo = False
                    boolPN = False
                    boolTP = False
                    strCH = CHARHRB
                    intSTP = 0
                    intSPN = 0
                    If InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text) > 0 Then
                        intSPN = InStr(1, strCH, "[PAGENUMBER]", CompareMethod.Text)
                        boolPN = True
                        boolGo = True

                    End If
                    If InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text) > 0 Then
                        intSTP = InStr(1, strCH, "[TOTALPAGES]", CompareMethod.Text)
                        boolTP = True
                        boolGo = True
                    End If

                    If boolGo Then
                        If boolPN And boolTP Then
                            If intSPN > intSTP Then
                                'str1 & [PN] & str2 & [TP] & str3
                                int1 = intSTP + intLTP
                                int2 = intSPN + intLPN
                                int3 = Len(strCH)
                                int4 = intSTP
                                int5 = intSPN
                                str1 = Mid(strCH, 1, int4 - 1)
                                str2 = Mid(strCH, int1, int5 - int1)
                                strB = "[PAGENUMBER]"
                                strA = "[TOTALPAGES]"
                                If int2 >= int3 Then
                                    str3 = ""
                                Else
                                    str3 = Mid(strCH, int2, Len(strCH) - int2)
                                End If
                                .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                                .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                                .Selection.TypeText(Text:=str3)

                            Else
                                'str1 & [PN] & str2 & [TP] & str3
                                int1 = intSPN + intLPN
                                int2 = intSTP + intLTP
                                int3 = Len(strCH)
                                int4 = intSPN
                                int5 = intSTP
                                str1 = Mid(strCH, 1, int4 - 1)
                                str2 = Mid(strCH, int1, int5 - int1)
                                strA = "[PAGENUMBER]"
                                strB = "[TOTALPAGES]"
                                If int2 >= int3 Then
                                    str3 = ""
                                Else
                                    str3 = Mid(strCH, int2, Len(strCH) - int2)
                                End If
                                .Selection.TypeText(Text:=Replace(str1, strA, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                                .Selection.TypeText(Text:=Replace(str2, strB, "", 1, -1, CompareMethod.Text))
                                wds = .Selection
                                .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                                .Selection.TypeText(Text:=str3)

                            End If
                        ElseIf boolPN And boolTP = False Then
                            strCH = Replace(strCH, "[PAGENUMBER]", "", 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=strCH)
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
                        ElseIf boolPN = False And boolTP Then
                            strCH = Replace(strCH, "[TOTALPAGES]", "", 1, -1, CompareMethod.Text)
                            .Selection.TypeText(Text:=strCH)
                            wds = .Selection
                            .Selection.Fields.Add(Range:=wds.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
                        End If
                    Else
                        .Selection.TypeText(Text:=strCH)
                    End If

                End If
                .Selection.TypeParagraph()
                .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdStory)

                .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

                'wdd.visible = True

                int3 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndSectionNumber)
                'loop through all section headers

                frmH.pb1.Value = 0
                frmH.pb1.Maximum = ctSec
                frmH.pb1.Value = int3

                For Count1 = 3 To ctSec
                    int1 = .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdActiveEndSectionNumber)
                    frmH.lblProgress.Text = "Creating Report Headers (" & Count1 & " of " & ctSec & ")..."
                    frmH.lblProgress.Refresh()
                    frmH.pb1.Value = Count1
                    frmH.pb1.Refresh()

                    .ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False

                    rm = .ActiveDocument.Sections(Count1).PageSetup.RightMargin
                    lm = .ActiveDocument.Sections(Count1).PageSetup.LeftMargin

                    If .ActiveDocument.Sections(Count1).PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                        rt = (11 * 72) - lm - rm
                        '.Selection.ParagraphFormat.TabStops.Add(Position:=684, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    Else
                        rt = (8.5 * 72) - lm - rm
                        '.Selection.ParagraphFormat.TabStops.Add(Position:=432, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                    End If
                    .ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.TabStops.ClearAll()
                    .ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                Next
            Catch ex As Exception

            End Try

end1:

            .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
end2:

            ctSec = .ActiveDocument.Sections.Count

            Dim pw
            Dim ph

            'wdd.visible = True
            boolH = False
            If boolH Then
            Else 'do search replace in header

                'wdd.visible = True

                'goto 1st section
                intSec = 1
                Try
                    .Selection.GoTo(What:=Microsoft.Office.Interop.Word.WdGoToItem.wdGoToSection, Which:=Microsoft.Office.Interop.Word.WdGoToDirection.wdGoToFirst, Count:=intSec, Name:="")
                Catch ex As Exception

                End Try
                '.ActiveDocument.Sections(intSec).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = False

                'move to header of next section
                Try

                    'wdd.visible = True

                    Try
                        If .ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
                            .ActiveWindow.Panes(2).Close()
                        End If
                        If .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                            .ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
                        End If
                        .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader
                    Catch ex As Exception

                    End Try



                    Dim boolLinked As Boolean = False
                    frmH.pb1.Value = 0
                    frmH.pb1.Maximum = ctSec

                    For Count1 = 1 To ctSec

                        Try
                            If Count1 = 1 Then


                                frmH.lblProgress.Text = "Search/Replacing Report Headers (" & Count1 & " of " & ctSec & ")..."
                                frmH.lblProgress.Refresh()
                                frmH.pb1.Value = Count1
                                frmH.pb1.Refresh()

                                If .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable) Then
                                    'format table
                                    .Selection.Tables.Item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                                End If

                                .Selection.WholeStory()

                                'do search replace
                                rng = wd.Selection.Range
                                Dim boolAA As Boolean

                            Else
                                .ActiveWindow.ActivePane.View.NextHeaderFooter()
                            End If

                            'boolLinked = .Selection.HeaderFooter.LinkToPrevious
                            boolLinked = .ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious

                            frmH.lblProgress.Text = "Search/Replacing Report Headers (" & Count1 & " of " & ctSec & ")..."
                            frmH.lblProgress.Refresh()
                            frmH.pb1.Value = Count1
                            frmH.pb1.Refresh()

                            If boolLinked And Count1 <> ctSec Then

                            Else



                                If .Selection.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable) Then
                                    'format table
                                    .Selection.Tables.Item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
                                End If

                                'wdd.visible = True

                                .Selection.WholeStory()

                                'do search replace
                                rng = wd.Selection.Range
                                Dim boolAB As Boolean

                                'wdd.visible = True


                                'boolA = SearchReplace(wd, "Report Body", rng, True, CHARHLB, 1, 0, 0, FALSE)
                                Try
                                    boolAB = SearchReplace(wd, "Report Body", rng, False, "aa", 1, 0, 0, True, True, True, 0)
                                Catch ex As Exception
                                    Dim varaaa
                                    varaaa = "a"
                                End Try

                                '20180329 LEE:
                                'need to start doing custom field codes
                                Call SearchReplaceCustomFieldCode(wd)

                                'don't do this anymore. Done in LinkPrevious

                                'rm = .ActiveDocument.Sections(Count1).PageSetup.RightMargin
                                'lm = .ActiveDocument.Sections(Count1).PageSetup.LeftMargin
                                'ph = .ActiveDocument.Sections(Count1).PageSetup.PageHeight
                                'pw = .ActiveDocument.Sections(Count1).PageSetup.PageWidth

                                ''.Selection.WholeStory()
                                'If .ActiveDocument.Sections(Count1).PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape Then
                                '    'rt = (11 * 72) - lm - rm
                                '    rt = pw - lm - rm
                                '    '.Selection.ParagraphFormat.TabStops.Add(Position:=684, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                                'Else
                                '    'rt = (8.5 * 72) - lm - rm
                                '    rt = pw - lm - rm
                                '    '.Selection.ParagraphFormat.TabStops.Add(Position:=432, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
                                'End If
                                'wd.Selection.WholeStory()
                                '.ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.TabStops.ClearAll()
                                '.ActiveDocument.Sections(Count1).Headers.Item(Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                                ''Don't go to footer! Shit gets screwed up. Use EnterFooter routine
                                ''go to current footer
                                'If wd.Selection.HeaderFooter.IsHeader = True Then

                                '    wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryFooter ' wdSeekCurrentPageFooter
                                'Else
                                '    wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryHeader 'wdSeekCurrentPageHeader
                                'End If
                                ''for some reason, that previous event brings me back one section
                                ''so forward one

                                ''wdd.visible = True

                                'wd.Selection.WholeStory()
                                'wd.Selection.ParagraphFormat.TabStops.ClearAll()
                                'wd.Selection.ParagraphFormat.TabStops.Add(Position:=rt, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)

                                ''go back to current header
                                'If wd.Selection.HeaderFooter.IsHeader = True Then
                                '    wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryFooter '
                                'Else
                                '    wd.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryHeader '
                                'End If

                            End If

                        Catch ex As Exception

                        End Try

                    Next

                    frmH.lblProgress.Text = "Search/Replacing Report Headers (" & ctSec & " of " & ctSec & ")..."
                    frmH.lblProgress.Refresh()
                    frmH.pb1.Value = 0
                    frmH.pb1.Maximum = ctSec
                    frmH.pb1.Value = ctSec
                    frmH.pb1.Refresh()

                    pw = .ActiveDocument.PageSetup.PageWidth
                    ph = .ActiveDocument.PageSetup.PageHeight

                Catch ex As Exception

                End Try

                Try
                    .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument
                Catch ex As Exception

                End Try

            End If

        End With

        frmH.pb1.Value = 0
        frmH.pb1.Maximum = pb1MO
        frmH.pb1.Value = pb1VO





    End Sub

    Sub FormatTOFigures(ByVal wd As Microsoft.Office.Interop.Word.Application)

        Dim strTitle As String

        strTitle = "Table of Figures"

        Try
            With wd
                '.ActiveDocument.Styles.Add(Name:=strTitle, Type:=Microsoft.Office.Interop.Word.WdStyleType.wdStyleTypeParagraph)
                .ActiveDocument.Styles(strTitle).AutomaticallyUpdate = True
                With .ActiveDocument.Styles(strTitle).Font
                    '.Name = "Times New Roman"
                    '.Size = 10
                    '.Bold = False
                    '.Italic = False
                    '.StrikeThrough = False
                    '.DoubleStrikeThrough = False
                    '.Outline = False
                    '.Emboss = False
                    '.Shadow = False
                    '.Hidden = False
                    '.SmallCaps = True
                    '.AllCaps = False
                    .ColorIndex = BlueHyperlinkColor '  Microsoft.Office.Interop.Word.WdColor.wdColorBlue
                    '.Engrave = False
                    '.Superscript = False
                    '.Subscript = False
                    '.Scaling = 100
                    '.Kerning = 0
                    '.Animation = wdAnimationNone
                End With
                With .ActiveDocument.Styles(strTitle).ParagraphFormat
                    .LeftIndent = 84 'InchesToPoints(1.17)
                    .RightIndent = 18 'InchesToPoints(0.26)
                    .SpaceBefore = 0
                    '.SpaceBeforeAuto = False
                    .SpaceAfter = 0
                    '.SpaceAfterAuto = False
                    .LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceSingle
                    .Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
                    .WidowControl = True
                    .KeepWithNext = False
                    .KeepTogether = False
                    .PageBreakBefore = False
                    .NoLineNumber = False
                    .Hyphenation = True
                    .FirstLineIndent = -84 'InchesToPoints(-1.17)
                    '.OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText
                    Try
                        .OutlineLevel = Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText ' Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevelBodyText
                    Catch ex As Exception

                    End Try

                    .CharacterUnitLeftIndent = 0
                    .CharacterUnitRightIndent = 0
                    .CharacterUnitFirstLineIndent = 0
                    Try
                        .LineUnitBefore = 0
                        .LineUnitAfter = 0
                    Catch ex As Exception

                    End Try
                End With
                .ActiveDocument.Styles(strTitle).NoSpaceBetweenParagraphsOfSameStyle = False
                .ActiveDocument.Styles(strTitle).ParagraphFormat.TabStops.ClearAll()
                .ActiveDocument.Styles(strTitle).ParagraphFormat.TabStops.Add(Position:=432, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderDots)

                With .ActiveDocument.Styles(strTitle).ParagraphFormat
                    'With .Shading
                    '    .Texture = wdTextureNone
                    '    .ForegroundPatternColor = wdColorAutomatic
                    '    .BackgroundPatternColor = wdColorAutomatic
                    'End With
                    .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    .Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
                    With .Borders
                        .DistanceFromTop = 1
                        .DistanceFromLeft = 4
                        .DistanceFromBottom = 1
                        .DistanceFromRight = 4
                        .Shadow = False
                    End With
                End With
                '.ActiveDocument.Styles(strTitle).LanguageID = wdEnglishUS
                '.ActiveDocument.Styles(strTitle).NoProofing = False
                '.ActiveDocument.Styles(strTitle).Frame.Delete()
            End With
        Catch ex As Exception

        End Try

    End Sub

    Sub GenerateFCReport()

        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim boolSTB As Boolean = True
        Try
            boolSTB = wd.Application.ShowWindowsInTaskbar
        Catch ex As Exception

        End Try

        Cursor.Current = Cursors.WaitCursor

        Dim dtbl As System.Data.DataTable
        Dim rows() As DataRow
        Dim intL As Short
        Dim Count1 As Short
        Dim wdSel As Microsoft.Office.Interop.Word.Selection
        Dim strPath As String
        Dim strM As String
        Dim intRows As Short
        Dim intCols As Short
        Dim str1 As String
        Dim str2 As String
        Dim strF As String
        Dim strS As String

        Dim pw, lm, rm

        dtbl = tblFieldCodes
        intL = dtbl.Rows.Count

        intRows = intL
        intCols = 4

        strF = "ID_TBLFIELDCODES > 0"
        strS = "CHARFIELDCODE ASC"

        '                        Case 1
        'strCN = "CHARFIELDCODE"
        '                Case 2
        'strCN = "CHARDESCRIPTION"
        '                Case 3
        'strCN = "CHAREXAMPLE"

        Dim dv As System.Data.DataView = New DataView(dtbl, strF, strS, DataViewRowState.CurrentRows)
        Dim tbl As System.Data.DataTable = dv.ToTable("a", True, "ID_TBLFIELDCODES", "CHARFIELDCODE", "CHARDESCRIPTION", "CHAREXAMPLE")
        'change first column
        intRows = tbl.Rows.Count
        For Count1 = 0 To intRows - 1
            tbl.Rows(Count1).BeginEdit()
            tbl.Rows(Count1).Item("ID_TBLFIELDCODES") = Count1 + 1
            tbl.Rows(Count1).EndEdit()
        Next

        Dim dv1 As System.Data.DataView = New DataView(tbl)
        Dim dgv As DataGridView

        dgv = frmH.dgvFieldCodes

        dgv.DataSource = dv1
        dgv.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText

        'select all rows in dgv
        dgv.SelectAll()

        intRows = dv1.Count

        'select all rows in dgv

        rows = dtbl.Select(strF, strS)

        wd.Documents.Add()
        strPath = GetNewTempFile(True)

        strPath = Replace(strPath, ".xml", ".docx", 1, -1, CompareMethod.Text)

        Try
            wd.ActiveDocument.SaveAs(FileName:=strPath, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles:=True, ReadOnlyRecommended:=False)

        Catch ex As Exception

            Try
                strPath = Replace(strPath, ".docx", ".docx", 1, -1, CompareMethod.Text)
                wd.ActiveDocument.SaveAs(FileName:=strPath, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles:=True, ReadOnlyRecommended:=False)

            Catch ex1 As Exception
                strM = "Hmmm. There was a problem preparing this report:" & ChrW(10) & ChrW(10) & strPath & ChrW(10) & ChrW(10) & ex1.Message & ChrW(10) & ChrW(10)
                strM = strM & "Please contact your StudyDoc Administrator."
                MsgBox(strM, MsgBoxStyle.Information, "Problem...")
                GoTo end1
            End Try
        End Try

        'wd.ActiveDocument.SaveAs(strPath)

        With wd

            With .ActiveDocument.PageSetup
                .Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape
            End With

            pw = wd.ActiveDocument.PageSetup.PageWidth
            lm = wd.ActiveDocument.PageSetup.LeftMargin
            rm = wd.ActiveDocument.PageSetup.RightMargin

            If wd.ActiveWindow.View.SplitSpecial <> Microsoft.Office.Interop.Word.WdSpecialPane.wdPaneNone Then
                wd.ActiveWindow.Panes(2).Close()
            End If
            If wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdNormalView Or wd.ActiveWindow. _
                ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView Then
                wd.ActiveWindow.ActivePane.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView
            End If

            wdSel = wd.Selection

            .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekCurrentPageHeader

            .Selection.ParagraphFormat.TabStops.ClearAll()
            .Selection.ParagraphFormat.TabStops.Add(Position:=pw - rm - lm, Alignment:=Microsoft.Office.Interop.Word.WdTabAlignment.wdAlignTabRight, Leader:=Microsoft.Office.Interop.Word.WdTabLeader.wdTabLeaderSpaces)
            .Selection.Font.Bold = True
            .Selection.TypeText(Text:="LABIntegrity StudyDoc Field Code Report")
            .Selection.Font.Bold = False
            .Selection.TypeText(Text:=vbTab & "Page ")
            .Selection.Fields.Add(Range:=wdSel.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage)
            .Selection.TypeText(Text:=" of ")
            .Selection.Fields.Add(Range:=wdSel.Range, Type:=Microsoft.Office.Interop.Word.WdFieldType.wdFieldNumPages)
            .Selection.TypeParagraph()
            .Selection.TypeText(Text:=Format(Now, "MMMM dd, yyyy hh:mm:ss tt"))
            .Selection.TypeParagraph()
            .Selection.MoveUp(Microsoft.Office.Interop.Word.WdUnits.wdLine, 1)
            .Selection.HomeKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdLine, Extend:=Microsoft.Office.Interop.Word.WdMovementType.wdExtend)
            With .Selection.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom)
                .LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
            End With
            .ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument

            wdSel = wd.Selection()
            'intRows = 152
            .ActiveDocument.Tables.Add(Range:=wdSel.Range, NumRows:=intRows + 1, NumColumns:=intCols, DefaultTableBehavior:=Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord9TableBehavior, AutoFitBehavior:=Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)
            .Selection.Font.Size = 10
            .Selection.Tables.Item(1).Rows.AllowBreakAcrossPages = False

            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderTop).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderLeft).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderRight).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderHorizontal).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderVertical).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalDown).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            '.Selection.Borders.Item(Microsoft.Office.Interop.Word.WdBorderType.wdBorderDiagonalUp).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            '.Selection.Font.Size = 11

            'enter headings
            .Selection.Tables.Item(1).Cell(1, 1).Select()
            str1 = "#"
            .Selection.Text = str1

            .Selection.Tables.Item(1).Cell(1, 2).Select()
            str1 = "Field Code"
            .Selection.Text = str1

            .Selection.Tables.Item(1).Cell(1, 3).Select()
            str1 = "Description"
            .Selection.Text = str1

            .Selection.Tables.Item(1).Cell(1, 4).Select()
            str1 = "Example"
            .Selection.Text = str1

            .Selection.Tables.Item(1).Cell(1, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.Font.Bold = True

            Dim strPaste As String
            Dim strPasteT As String
            Dim Count2 As Short
            Dim Count3 As Short
            Dim int1 As Int16
            Dim strCN As String

            'int1 = 0
            'For Count2 = 0 To intRows - 1

            '    int1 = int1 + 1

            '    strPasteT = ""
            '    For Count3 = 1 To intCols - 1

            '        Select Case Count3
            '            Case 1
            '                strCN = "CHARFIELDCODE"
            '            Case 2
            '                strCN = "CHARDESCRIPTION"
            '            Case 3
            '                strCN = "CHAREXAMPLE"
            '        End Select

            '        str1 = NZ(rows(Count2).Item(strCN), "NA")

            '        If InStr(1, str1, ChrW(10), CompareMethod.Text) > 0 Then
            '            str1 = Replace(str1, ChrW(10), ChrW(11), 1, -1, CompareMethod.Text)
            '        End If

            '        If Count3 = 1 Then
            '            strPasteT = CStr(int1) & ChrW(9) & str1
            '        Else
            '            strPasteT = strPasteT & ChrW(9) & """" & CStr(str1) & """"
            '            'strPasteT = strPasteT & ChrW(9) & str1
            '        End If
            '    Next

            'If Count2 = 1 Then
            '    strPaste = strPasteT
            'Else
            '    strPaste = strPaste & ChrW(10) & strPasteT
            'End If



            'If Count2 = 0 Then
            '    strPaste = strPasteT
            'Else
            '    strPaste = strPaste & ChrW(10) & strPasteT
            'End If

            'Next
            .Selection.Tables.Item(1).Cell(2, 1).Select()

            'wdd.visible = True

            'send strpaste to clipboard
            Try
                Clipboard.Clear()
            Catch ex As Exception

            End Try
            Try
                Clipboard.SetDataObject(dgv.GetClipboardContent())
            Catch ex As Exception

            End Try
            'select appropriate rows

            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
            'paste from clipboard
            Try
                .Selection.PasteAndFormat(Microsoft.Office.Interop.Word.WdRecoveryType.wdPasteDefault)
            Catch ex As Exception

            End Try

            'the paste action removes the range object and any table formatting, must reset it
            Call GlobalTableParaFormat(wd)

            'autofit table
            Call AutoFitTable(wd, False)

            .Selection.Tables.Item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)
            .Selection.Tables.Item(1).AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitFixed)

            'autosizing isn't optimal. do it manually

            With .Selection.Tables(1)

                Dim totW As Double
                Dim num1 As Double
                For Count1 = 1 To intCols
                    num1 = .Columns(Count1).Width
                    totW = totW + num1
                Next

                '.Columns(1).Select()
                'wd.Selection.Font.Size = 10 '.Selection.Font.Size - 1
                .Columns(1).Width = totW * 0.05
                .Columns(2).Width = totW * 0.25
                .Columns(3).Width = totW * 0.45
                .Columns(4).Width = totW * 0.25


            End With

            'make column1 font smaller because it's not autfitting
            '.Selection.Tables(1).Columns(1).Select()
            '.Selection.Font.Size = 10 '.Selection.Font.Size - 1

            .Selection.Tables.Item(1).Cell(2, 1).Select()
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdRow, Extend:=True)
            .Selection.EndKey(Unit:=Microsoft.Office.Interop.Word.WdUnits.wdColumn, Extend:=True)
            .Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft
            .Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom
            .Selection.Font.Size = 10


        End With

        wd.ActiveDocument.Save()


        Try
            wd.ActiveDocument.Close(False)
        Catch ex As Exception

        End Try

        Try
            wd.Application.ShowWindowsInTaskbar = boolSTB
        Catch ex As Exception

        End Try

        wd.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges)
        wd = Nothing

        Threading.Thread.Sleep(250)

        Call OpenAFR(strPath, "", False, boolSTB, True, False)

        'frmH.lblProgress.Visible = False
        'frmH.pb1.Visible = False
        'frmH.pb2.Visible = False

        'frmH.panProgress.Visible = False
        'frmH.panProgress.Refresh()

        frmH.Refresh()

        frmWordStatement.Activate()

end1:

        Cursor.Current = Cursors.Default

    End Sub

End Module
