Option Compare Text

Imports System
Imports System.IO
Imports System.Text


Public Class frmWordStatement

    Public boolRO As Boolean = True 'Open document as readonly
    Public strPath As String
    Public strPathT As String
    Public id As Int64 ' this is id_tblWordStatements
    Public boolCancel As Boolean = True
    Public OldRow As Short
    Public boolFormLoad As Boolean = True
    Public intDGVCol As Short
    Public varBefore = ""
    Public varAfter = ""
    Public varTitleBefore = ""
    Public varTitleAfter = ""
    Public varBeforeTemp = ""
    Public varAfterTemp = ""
    Public strSection As String
    Public idNew As Int64
    Public boolReport As Boolean = False 'boolReport is a Report Template
    Public strReport As String
    Public boolFocus As Boolean = False
    Public strPathPDF As String
    Public boolHold As Boolean
    Public boolSTB As Boolean = True
    Public boolTested As Boolean = False
    Public boolEdit As Boolean = False
    Public boolReadOnly As Boolean = False
    Public boolOutlier As Boolean = False
    Public boolcmdAddStatementJustClicked = False

    Public boolGoBack As Boolean = False ''20181023 LEE:


    Sub EnableButtons(btn As String)

        'Dim str1 As String
        'str1 = "The Edit button must be clicked in order to access the Report Template."
        'Me.lblDGV.Text = str1

        'position lblBlink

        Dim h, w
        Dim h1, w1
        Dim l

        h1 = Me.lblBlink.Height
        w1 = Me.lblBlink.Width

        h = Me.Height
        w = Me.Width
        l = Me.panEdraw.Left


        Me.lblBlink.Left = ((w - l) - w1) / 2
        Me.lblBlink.Left = ((l + Me.panEdraw.Width) - w1) / 2
        Me.lblBlink.Top = (h - h1) / 2
        Me.ov1.Refresh()  'Refresh overlay as well; otherwise we get multiple lblBlinks shown.

        'don't use these buttons anymore
        'make them invisible
        Me.panEditReports.Visible = False

        Dim dgv As DataGridView = Me.dgvReportStatements
        Dim dgv1 As DataGridView = Me.dgvVersions

        Select Case btn

            Case "Edit"

                Me.cmdEdit.Enabled = False
                Me.cmdCancelEdit.Enabled = True
                Me.cmdSave.Enabled = True
                Me.cmdExit.Enabled = False
                Me.cmdCompareDocs.Enabled = False
                Me.cmdInsertNew.Enabled = True
                Me.cmdOpenExisting.Enabled = True
                Me.cmdFieldCode.Enabled = True

                Me.cmdEditTitle.Enabled = False

                Me.cmdPrint1.Enabled = True
                Me.cmdDeactivateTemplates.Enabled = False

                Me.cmdWord1.Enabled = True
                Me.cmdPDF.Enabled = True
                Me.cmdPrint.Enabled = True

                'Me.panEditReports.Visible = True

                dgv.Enabled = False
                dgv.BackgroundColor = Color.Gray
                dgv.DefaultCellStyle.BackColor = Color.Gray

                dgv1.Enabled = False
                dgv1.BackgroundColor = Color.Gray
                dgv1.DefaultCellStyle.BackColor = Color.Gray

                Me.lblBlink.Visible = False

                Call DoWrite()

                'Me.panEdraw.Enabled = True

                '888
                'tSave.Enabled = False


            Case "Cancel", "Save"

                Me.cmdEdit.Enabled = True
                Me.cmdCancelEdit.Enabled = False
                Me.cmdSave.Enabled = False
                Me.cmdExit.Enabled = True
                Me.cmdCompareDocs.Enabled = True
                Me.cmdInsertNew.Enabled = False
                Me.cmdOpenExisting.Enabled = False
                Me.cmdFieldCode.Enabled = False
          
                Me.cmdPrint1.Enabled = False
                Me.cmdDeactivateTemplates.Enabled = True

                Me.cmdWord1.Enabled = False
                Me.cmdPDF.Enabled = False
                Me.cmdPrint.Enabled = False

                Me.cmdEditTitle.Enabled = True

                'Me.panEditReports.Visible = False

                dgv.Enabled = True
                dgv.BackgroundColor = Color.White
                dgv.DefaultCellStyle.BackColor = Color.White

                dgv1.Enabled = True
                dgv1.BackgroundColor = Color.White
                dgv1.DefaultCellStyle.BackColor = Color.White

                Me.lblBlink.Visible = True

                Call DoReadOnly()

                'Me.panEdraw.Enabled = False

                '8888
                'tSave.Enabled = True

        End Select

        'debug
        Dim var1, var2
        var1 = Me.panEdraw.Top
        var2 = Me.lblBlink.Top

        'MsgBox("pan: " & var1 & ", lbl: " & var2)

    End Sub


    Sub ovInit()

        Try
            Me.ov1.CloseDoc()
            Me.ov1.Dispose()
        Catch ex As Exception

        End Try

        'Me.ov1.LicenseName = "Gubbs Inc"

        'Me.ov1.LicenseKey = "ED99-5500-1211-ABBD" 'v6
        'Me.ov1.LicenseCode = "EDO8-5556-1211-ABEB" 'v8

        Me.ov1.LicenseName = "LabIntegrity7631358702" 'v8.812
        Me.ov1.LicenseCode = "EDO8-5573-1234-ABEB" 'v8.812


    End Sub

    Private Sub frmWordStatement_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        Call ClearTemp()

    End Sub

    Private Sub frmWordStatement_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        '20181023 LEE:
        If boolGoBack Then
        Else
            Dim strM As String

            strM = "This window cannot be closed with the Close button."
            strM = strM & ChrW(10) & ChrW(10) & "Please use the Go Back button in the Saved mode."

            MsgBox(strM, vbInformation, "Invalid action...")

            e.Cancel = True
        End If

    End Sub

    Private Sub frmWord_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.WindowState = FormWindowState.Maximized

        Call CorrectActive()

        Call DoubleBufferControl(Me, "dgv")

        'Call ControlDefaults(Me)

        boolFormLoad = True

        Me.ov1.Dock = DockStyle.Fill
        'Me.wb1.Dock = DockStyle.Fill

        boolHold = False

        Me.tSave.Enabled = False

        Dim strT As String

        If gDoPDF Then
            Me.ov1.Visible = False
            'Me.wb1.Visible = True
            strT = "LABIntegrity StudyDoc" & ChrW(8482) & " PDF Document"
        Else
            Me.ov1.Visible = True
            'Me.wb1.Visible = False
            strT = "LABIntegrity StudyDoc" & ChrW(8482) & " Microsoft" & ChrW(174) & " Word Document"
        End If

        Dim str1 As String
        If gboolET Then
            str1 = "The document must be Checked Out in order to access the Report Template."
        Else
            str1 = "The Edit button must be clicked in order to access the Report Template."
        End If
        Me.lblBlink.Text = str1

        Me.Text = strT

        'Me.ov1.Caption = "Microsoft" & ChrW(174) & "Word Viewer"


        'Me.ov1.LicenseKey = "ED99-5500-1211-ABBD"

        'Me.ov1.LicenseName = "Gubbs Inc"

        'Me.ov1.LicenseKey = "ED99-5500-1211-ABBD" 'v6
        'Me.ov1.LicenseCode = "EDO8-5556-1211-ABEB" 'v8.371

        Me.ov1.LicenseName = "LabIntegrity7631358702" 'v8.812
        Me.ov1.LicenseCode = "EDO8-5573-1234-ABEB" 'v8.812

        boolFormLoad = True
        Dim t, h, l, w

        w = My.Computer.Screen.WorkingArea.Width
        h = My.Computer.Screen.WorkingArea.Height

        Me.Top = 0
        Me.Left = 0
        Me.Width = w
        Me.Height = h

        t = Me.Top
        h = Me.Height
        l = Me.Left
        w = Me.Width

        'Me.panWdWB.Top = 0
        'Me.panWdWB.Height = h
        'Me.panWdWB.Width = w - Me.panWord.Left - 30

        'Me.panWdWB.ContextMenuStrip = Me.cmsfrmWord
        'Me.panWdWB.Visible = False

        'Me.panWord.Left = Me.pan1.Left + Me.pan1.Width
        'Me.panWord.Top = Me.lblSection.Top + Me.lblSection.Height + 2
        'Me.panWord.Height = h - (40 + Me.panWord.Top) '(Me.CHARTITLE.Top + Me.CHARTITLE.Height + 2)
        'Me.panWord.Width = w - Me.panWord.Left - 30

        '20181112 LEE:
        'Don't mess with panEdraw width, top or left; only do width
        ''it is anchored
        'Me.panEdraw.Left = Me.pan1.Left + Me.pan1.Width
        ''Me.panEdraw.Top = Me.cmdWord1.Top + Me.pan1.Top ' Me.lblSection.Top + Me.lblSection.Height + 2

        'Me.panEdraw.Left = Me.pan1.Left + Me.pan1.Width
        'Me.panEdraw.Top = Me.pan1.Top + Me.cmdWord1.Top

        Me.panEdraw.Height = h - (40 + Me.panEdraw.Top) '(Me.CHARTITLE.Top + Me.CHARTITLE.Height + 2)
        'Me.panEdraw.Width = w - Me.panEdraw.Left - 30

        'Me.panEdraw.Left = Me.panWord.Left
        'Me.panEdraw.Top = Me.panWord.Top
        'Me.panEdraw.Width = Me.panWord.Width
        'Me.panEdraw.Height = Me.panWord.Height
        Me.panEdraw.ContextMenuStrip = Me.cmsfrmWord

        'Me.panWord.ContextMenuStrip = Me.cmsfrmWord
        'Me.panWord.Visible = False
        Me.panEdraw.Visible = True

        Me.pan2.Left = 0
        Me.pan2.Top = 0

        'Me.lblEdit.Visible = False

        If boolReport Then 'boolReport is a Report Template, not a generated report

        Else
            Call LoadAFRReport(True)
            'Call LoadAFR(True)
        End If


        'Me.lblSection.Visible = True


        Me.cmdSave.Enabled = False


        boolFormLoad = False

        'Me.afrWord.Refresh()

        'Me.ov1.Refresh()

        If gDoPDF Then
            'rearrange buttons
            Me.cmdWord.Visible = False ' True
            Me.cmdOpenPDF.Visible = True
            'Me.wb1.Visible = True
            Me.ov1.Visible = False
        Else
            Me.cmdOpenPDF.Visible = True ' False
            'Me.wb1.Visible = False
            Me.ov1.Visible = True
        End If

        Me.dgvReportStatements.AutoResizeRows()

        If boolReport And Me.panEditReports.Visible Then
            Call EnableButtons("Cancel")
        End If

        Try
            Call DoReadOnly()
        Catch ex As Exception

        End Try

        'move cmdShow
        Me.cmdShow.Top = Me.dgvVersions.Top + Me.cmdShow.Height + 10
        Me.cmdShow.SendToBack()

        If gboolET Then
            Me.cmdEdit.Text = "Check &Out"
            Me.cmdSave.Text = "Check &In"

        Else
            Me.cmdEdit.Text = "&Edit"
            Me.cmdSave.Text = "&Save"

            Me.panList.Top = Me.pan1.Top + Me.pan2.Top + Me.pan2.Height

            'Me.panList.Height = Me.dgvReportStatements.Top + Me.dgvReportStatements.Height + 5
            Me.cmdShow.Visible = True
            Me.panList.Height = Me.lblVersions.Top + Me.lblVersions.Height
            Me.lblVersions.Visible = False

            Me.cmdCompareDocs.Visible = False
            Me.cmdExit.Top = Me.cmdCompareDocs.Top
            Me.panSave.Height = Me.cmdExit.Top + Me.cmdExit.Height + 10
            Me.panSave.Top = Me.panList.Top + Me.panList.Height + 5

            If Me.panSave.Visible Then
            Else
                Me.cmdShow.Visible = False
                '20160713 LEE: Do not want to show anything at startup
                'This can take a long time on some customer's machines
                'makes performance seem slow
                'Call DoVersionsSelChange()

            End If

        End If

        Me.panEditReports.Top = Me.pan1.Top + Me.panList.Top + Me.dgvReportStatements.Top
        Me.panEditReports.Left = Me.pan1.Left + Me.pan1.Width + 5


        'don't use lblEdit anymore
        'Me.lblEdit.Visible = False

        frmWordStatement_ToolTipsSet()

    End Sub

    Sub FormLoad()

        'Dim strpathT As String 'this is public
        Dim Count1 As Int16
        Dim strW As String
        Dim fs As FileStream
        Dim strM As String
        Dim var1
        Dim str1 As String

        'Me.ov1.ProtectDoc(-1)'v8

        'str1 = "Report Statement(s) for section: " & strSection
        'Me.lblSection.Text = str1

        Call EnableOVstuffReportTemplate(Me.ov1)

        boolFormLoad = True

        Me.cmdAddStatement.Text = "Create &New Template..."
        Me.lblEditTitles.Text = "Edit Template Names"

        'Call LoadAFR() 'don't need this. dgv_selectionchange will fire and trigger this

        Call dgvReportStatementsConfigure()

        Call ReportStatementsChange()

        'select the current row
        Dim dgv As DataGridView
        Dim intRow As Short
        Dim intCol As Short
        dgv = Me.dgvReportStatements

        'find intcol
        intDGVCol = 0
        For Count1 = 0 To dgv.Columns.Count - 1
            var1 = dgv.Columns(Count1).Name
            If StrComp(var1, "CHARTITLE", CompareMethod.Text) = 0 Then
                intCol = Count1
                intDGVCol = intCol
                Exit For
            End If
        Next
        If dgv.Rows.Count = 0 Then
        Else
            intRow = 0
            For Count1 = 0 To dgv.Rows.Count - 1
                var1 = dgv("ID_TBLWORDSTATEMENTS", Count1).Value
                If var1 = id Then
                    intRow = Count1
                    Exit For
                End If
            Next

            'this will signal selchange event
            dgv.CurrentCell = dgv.Rows(intRow).Cells(intCol)
            dgv.Rows(intRow).Selected = True
        End If

        boolFormLoad = False

        OldRow = intRow
        Me.CHARTITLE.Text = dgv("CHARTITLE", intRow).Value
        varTitleBefore = dgv("CHARTITLE", intRow).Value
        varTitleAfter = varTitleBefore

        Dim boolPDFTrue As Boolean = False
        If gDoPDF Then

            Call GoToWord(True)

        Else
            Me.lblSection.Text = strReport
        End If

        'If Me.pan2.Visible Then
        If gDoPDF Then
            'rearrange buttons
            Dim n1, n2, n3, n4, n5, n6, n7
            n1 = Me.pan2.Width
            n2 = Me.cmdWord.Width
            n3 = n2 * 2
            n4 = (n1 - n3) / 3
            Me.cmdWord.Left = n4
            Me.cmdOpenPDF.Left = n4 + n2 + n4

            Me.cmdWord.Visible = False ' True
            Me.cmdOpenPDF.Visible = True
            'Me.wb1.Visible = True
            Me.ov1.Visible = False
            Me.panRefresh.Visible = False
        Else
            Me.cmdOpenPDF.Visible = True 'False
            'Me.wb1.Visible = False
            Me.ov1.Visible = True
            Me.panRefresh.Visible = True

        End If

        'place some controls
        Call PlaceControls()

        'End If


        'Me.cmdSave.Enabled = False

        'boolFormLoad = False 'wait until actual form_load

    End Sub

    Sub PlaceControls()

        Me.panRefresh.Top = Me.pan1.Top + Me.pan1.Height
        Me.panRefresh.Left = Me.pan1.Left

    End Sub

    Sub LoadAFRReport(ByVal boolOV As Boolean)

        If gDoPDF Then
            'rearrange buttons
            Me.cmdWord.Visible = False ' True
            Me.cmdOpenPDF.Visible = True
            'Me.wb1.Visible = True
            Me.ov1.Visible = False
            'Exit Sub
        Else
            Me.cmdOpenPDF.Visible = True ' False
            'Me.wb1.Visible = False
            Me.ov1.Visible = True
        End If

        Cursor.Current = Cursors.WaitCursor

        Dim strS As String


        strPathT = strReport
        Dim var1

        If boolOV Then
            'shouldn't have to
            '20140228 yes i do
            Try
                Me.ov1.CloseDoc() 'v8
                'Me.ov1.Close() 'v6
            Catch ex As Exception
                MsgBox("LoadAFRReport (Me.ov1.CloseDoc()): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

            Threading.Thread.Sleep(100) 'give the doc some time to establish, or get a read-only error

            'before using ov1 to open doc
            'must open plain wd to find ShowWindowsInTaskbar setting
            Dim wd As New Microsoft.Office.Interop.Word.Application
            Dim wdDoc As Microsoft.Office.Interop.Word.Document
            Try
                boolSTB = wd.ShowWindowsInTaskbar
            Catch ex As Exception
                boolSTB = True
            End Try
            Try
                wd.Quit(False)
            Catch ex As Exception
                var1 = var1
            End Try
            'If boolTested Then
            'Else
            '    Dim wd As New Microsoft.Office.Interop.Word.Application
            '    boolSTB = wdDoc.Application.ShowWindowsInTaskbar
            '    wd.Quit()
            '    wd = Nothing
            '    boolTested = True
            'End If

            Try
                'Me.ov1.Open(strReport, "Word.Application") 'v8
                'Me.ov1.Open(strReport) 'v6
                Me.ov1.OpenWord(strReport)

            Catch ex As Exception
                MsgBox("LoadAFRReport (Me.ov1.Open(...)): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

            'Dim wdDoc As Microsoft.Office.Interop.Word.Document

            Try
                wdDoc = Me.ov1.ActiveDocument

            Catch ex As Exception
                Exit Sub
            End Try

            Try
                wdDoc.Application.ShowWindowsInTaskbar = boolSTB
            Catch ex As Exception
                var1 = ex.Message
            End Try


            Try
                Call SetWordStuff(wdDoc)
            Catch ex As Exception

            End Try


        Else

            Dim wdDoc As Microsoft.Office.Interop.Word.Document
            Try
                wdDoc = Me.ov1.ActiveDocument

            Catch ex As Exception
                Exit Sub
            End Try


            Try
                Call SetWordStuff(wdDoc)
            Catch ex As Exception

            End Try

        End If


        If gDoPDF Then
            'rearrange buttons
            Me.cmdWord.Visible = False ' True
            Me.cmdOpenPDF.Visible = True
            'Me.wb1.Visible = True
            Me.ov1.Visible = False
        Else
            Me.cmdOpenPDF.Visible = True ' False
            'Me.wb1.Visible = False
            Me.ov1.Visible = True
        End If

        Cursor.Current = Cursors.Default


    End Sub

    Sub dgvReportStatementsConfigure()

        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim Count1 As Short
        Dim strF As String
        Dim strS As String

        dgv1 = Me.dgvReportStatements
        dgv2 = Me.dgvVersions

        strF = "CHARWORDSTATEMENT = 'Active'"
        strS = "CHARTITLE ASC"
        Dim dv1 As New DataView(tblWordStatements, strF, strS, DataViewRowState.CurrentRows)
        dv1.AllowDelete = False
        dv1.AllowEdit = False
        dv1.AllowNew = False

        strF = "ID_TBLWORDSTATEMENTSVERSIONS < 1"
        strS = "INTWORDVERSION DESC"
        Dim dv2 As New DataView(tblWordStatementsVERSIONS, strF, strS, DataViewRowState.CurrentRows)
        dv2.AllowDelete = False
        dv2.AllowEdit = False
        dv2.AllowNew = False


        dgv1.DataSource = dv1

        dgv1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        'dv = dgv2.DataSource
        'dv.AllowEdit = False
        'dgv1.DataSource = dv
        For Count1 = 0 To dgv1.ColumnCount - 1
            dgv1.Columns(Count1).Visible = False
            'dgv1.Columns(Count1).SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        dgv1.Columns("CHARTITLE").Visible = True
        dgv1.Columns("CHARTITLE").HeaderText = "Template"
        dgv1.RowHeadersWidth = 20

        dgv1.ReadOnly = True
        dgv1.AutoResizeColumns()
        dgv1.AutoResizeRows()
        dgv1.AutoResizeRows()



        dgv2.DataSource = dv2

        For Count1 = 0 To dgv2.Columns.Count - 1
            dgv2.Columns(Count1).Visible = False
        Next

        'dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        dgv2.Columns("INTWORDVERSION").Visible = True
        dgv2.Columns("INTWORDVERSION").HeaderText = "Ver."
        dgv2.Columns("INTWORDVERSION").DisplayIndex = 0

        dgv2.Columns("CHARDESCRIPTION").Visible = True
        dgv2.Columns("CHARDESCRIPTION").HeaderText = "Description"
        dgv2.Columns("CHARDESCRIPTION").DisplayIndex = 1

        dgv2.Columns("CHARUSERID").Visible = True
        dgv2.Columns("CHARUSERID").HeaderText = "Owner"
        ' dgv2.Columns("CHARUSERID").DisplayIndex = dgv2.Columns.Count - 1

        'for some reason, the order is all goofy

        'dgv2.Columns("CHARUSERID").DisplayIndex = dgv2.Columns("CHARDESCRIPTION").DisplayIndex + 1


        dgv2.RowHeadersWidth = 20

        dgv2.ReadOnly = True
        dgv2.AutoResizeColumns()
        dgv2.AutoResizeRows()
        dgv2.AutoResizeRows()


    End Sub


    Sub InsertFC(ByVal boolShow As Boolean)

        Dim boolM As Boolean
        Dim strM As String

        strM = ""
        boolM = False

        Dim pos As Int64
        Dim strT As String
        Dim str1 As String
        Dim strL As String
        Dim strR As String
        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        Dim wdApp As Microsoft.Office.Interop.Word.Application

        'record position of cursor in text box
        'wd = wb.Document

        wdDoc = Me.ov1.ActiveDocument
        'wdDoc = Me.afrWord.ActiveDocument
        'wdDoc = Me.wbFrmWd.Document
        wdApp = wdDoc.Application


        'wd_doc_RBS = wd 'frmH.wbRBS.Document
        'wd_app_RBS = wd_doc_RBS.Application

        'MsgBox(wd_app_RBS.Selection.Start)

        Dim frm As New frmFieldCodes

        'Me.Cursor = New Cursor(Cursor.Current.Handle)

        'frm.Location = new system.drawing.point(Cursor.Position.X, Cursor.Position.Y + 10)

        frm.ShowDialog()

        'Me.afrWord.Refresh()

        If frm.boolCancel Then


            'wdapp..Collapse(Direction:=Microsoft.Office.Interop.Word.WdCollapseDirection.wdCollapseStart)

        Else

            str1 = frm.strFC
            wdApp.Selection.TypeText(Text:=str1)

        End If

        GoTo end1

        Try
            wdApp.DisplayAlerts = False
        Catch ex As Exception

        End Try


        frm.Dispose()

        Try

        Catch ex As Exception

        End Try

        Try
            wdDoc.Close()
        Catch ex As Exception

        End Try

        Try
            wdApp.DisplayAlerts = True
        Catch ex As Exception

        End Try
        Try
            wdApp.Quit()
        Catch ex As Exception

        End Try

        Try
            wdDoc = Nothing
        Catch ex As Exception

        End Try

        Try
            wdApp = Nothing
        Catch ex As Exception

        End Try




end1:
    End Sub

    Private Sub cmdFieldCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFieldCode.Click

        Dim strM As String

        If Me.cmdEdit.Enabled Then
            strM = "Document is not in Edit mode."
            MsgBox(strM, vbInformation, "Invalid action...")
        Else
            Call InsertFC(False)
        End If


    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        '20181023 LEE:
        boolGoBack = True

        Call DoExit()

    End Sub

    Sub DoExit()

        Dim var1
        Try
            Me.ov1.CloseDoc()
        Catch ex As Exception
            var1 = var1
        End Try

        Try
            Me.ov1.Dispose()
        Catch ex As Exception
            var1 = var1
        End Try

        Try
            Call ClearTemp()
        Catch ex As Exception
            var1 = var1
        End Try

        boolCancel = True
        'Me.Visible = False
        'Me.Close()

        'Me.Dispose()

        ''try getting rid of this

        If boolReport Then

            Try
                frmH.BringToFront()
            Catch ex As Exception

            End Try

            Try
                frmH.Visible = True
            Catch ex As Exception

            End Try

            Try
                frmH.Activate()
            Catch ex As Exception

            End Try

            Try
                frmH.Refresh()
            Catch ex As Exception

            End Try

            Try
                frmH.Focus()
            Catch ex As Exception

            End Try

            'Me.Visible = False

            'Try
            '    Me.Close()
            'Catch ex As Exception
            '    var1 = var1
            'End Try

            'Try
            '    Me.Dispose()
            'Catch ex As Exception
            '    var1 = var1
            'End Try

        Else

            If BOOLFORCEFINALREPORTPDF Then
            Else
                Try
                    frmH.Visible = True
                Catch ex As Exception

                End Try
                frmH.BringToFront()
                frmH.Refresh()
            End If

        End If

        Me.Visible = False

        Try
            Me.Close()
        Catch ex As Exception
            var1 = var1
        End Try

        Try
            Me.Dispose()
        Catch ex As Exception
            var1 = var1
        End Try

        'frmH.BringToFront()
        'frmH.Refresh()


    End Sub

    Sub DoCancel()

        Cursor.Current = Cursors.WaitCursor

        Dim strPath As String
        strPath = Me.lblSection.Text
        Dim var1

        'Me.ov1.Open(strPath, "Word.Application")
        If Me.ov1.IsOpened Then
            Me.ov1.CloseDoc()
        End If
        Try
            Me.ov1.Open(strPath, "Word.Application")
        Catch ex As Exception
            var1 = "a" 'debug
        End Try

        Cursor.Current = Cursors.Default

    End Sub


    Sub DoSave(strDescr As String)


        Dim strReportO As String

        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        'Dim wdApp As Microsoft.Office.Interop.Word.Application
        Dim var1 As Object
        Dim strt

        'If varBefore = varAfter And varTitleBefore = varTitleAfter Then
        '    MsgBox("Save canceled. No changes have been made.", MsgBoxStyle.Information, "Invalid action...")
        '    Exit Sub
        'End If

        Cursor.Current = Cursors.WaitCursor

        Me.cmdCancel.Enabled = False
        Me.cmdEditStatements.Enabled = True

        Dim oPath As String
        Dim oName As String
        Dim strUPath As String = "C:\LabIntegrity\StudyDoc\Temp" ' "C:\Labintegrity\StudyDoc\Temp"
        Dim strNPath As String

        Dim strExt As String
        Dim strR As String
        Dim bool2007 As Boolean
        Dim strPath As String
        Dim boolM As Boolean
        Dim ver As Short

        Try

            wdDoc = Me.ov1.ActiveDocument

            'strt = wdDoc.Application.Selection.Start

            'find path of wddoc
            oPath = wdDoc.Path
            oName = wdDoc.FullName

            ''The following can't happen any more
            'If InStr(1, oPath, strUPath, CompareMethod.Text) < 1 Then
            '    'save locally
            '    'user probably opened a new document
            '    strNPath = Replace(oName, oPath, strUPath, 1, -1, CompareMethod.Text)
            '    strReportO = GetNewTempFile()
            '    'must save this as 2003 xml which saves vba module
            '    wdDoc.SaveAs(strReportO, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXML) 'this is WD 2003 xml format

            '    '****
            'Else
            '    strReportO = strReport
            'End If

            var1 = strPathT
            'get new file
            strPathT = GetNewTempFile(True)
            strReport = strPathT
            Me.lblSection.Text = strPathT

            ver = CInt(wdDoc.Application.Version)
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If

            'first save original doc
            wdDoc.Application.DisplayAlerts = False
            Try
                wdDoc.Save()
            Catch ex As Exception
                MsgBox("DoSave first save record" & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

            Try
                'now save as new file for fs to open
 
                If bool2007 Then
                    wdDoc.SaveAs(strPathT, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXML) 'this is WD 2003 xml format
                Else
                    'https://msdn.microsoft.com/en-us/library/office/ff839952.aspx
                    'wdFormatFlatXML 19 Open XML file format saved as a single XML file.
                    'wdFormatFlatXML 20 Open XML file format with macros enabled saved as a single XML file.
                    'Word 2010 only
                    wdDoc.SaveAs2(strPathT, FileFormat:=20)
                End If

            Catch ex As Exception
                MsgBox("now save as new file for fs to open" & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try


            wdDoc.Application.DisplayAlerts = True

            'keep setting this

            Call SetWordStuff(wdDoc)

            'try killing docs here
            If Me.ov1.IsOpened Then
                Try
                    Me.ov1.CloseDoc()
                Catch ex As Exception

                End Try
            End If


            Pause(0.25)

            Cursor.Current = Cursors.WaitCursor

            If gboolAuditTrail Then
                'clear audittrailtemp
                tblAuditTrailTemp.Clear()
                idSE = 0
                Call FillAuditTrailTemp(tblWordStatements)
            End If

            If boolGuWuOracle Then
                Try
                    ta_tblWordStatements.Update(tblWordStatements)
                Catch ex As DBConcurrencyException
                    'ds2005.TBLWORDSTATEMENTS.Merge('ds2005.TBLWORDSTATEMENTS, True)
                End Try
            ElseIf boolGuWuAccess Then
                Try
                    ta_tblWordStatementsAcc.Update(tblWordStatements)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLWORDSTATEMENTS.Merge('ds2005Acc.TBLWORDSTATEMENTS, True)
                End Try
            ElseIf boolGuWuSQLServer Then
                Try
                    ta_tblWordStatementsSQLServer.Update(tblWordStatements)
                Catch ex As DBConcurrencyException
                    'ds2005Acc.TBLWORDSTATEMENTS.Merge('ds2005Acc.TBLWORDSTATEMENTS, True)
                End Try
            End If

            '****

            If gboolAuditTrail Then
                Dim dt1 As DateTime
                dt1 = Now

                Call RecordAuditTrail(False, dt1)
            End If

            're-open original
            'do not re-open original
            'state should be blank document
            'Try
            '    Me.ov1.OpenWord(oName)
            'Catch ex As Exception
            '    var1 = var1
            'End Try

            Cursor.Current = Cursors.WaitCursor

            Call UpdateDatabaseSave(strPathT, strDescr)

            Cursor.Current = Cursors.WaitCursor

        Catch ex As Exception

            Dim strM As String
            strM = "A problem occurred when attempting to open the Word file. Please try again." & ChrW(10) & ex.Message
            MsgBox(strM, MsgBoxStyle.Information, "Problem...")

        End Try


        boolCancel = False

        Me.CHARTITLE.ReadOnly = True
        Me.CHARTITLE.Enabled = False

        Me.dgvReportStatements.AutoResizeRows()

        Cursor.Current = Cursors.Default

        wdDoc = Nothing
        ' wdApp = Nothing

        'restore old strreport
        strReport = strReportO
        Me.lblSection.Text = strReport

        Me.chkCreateNew.Checked = False


end1:

        boolHold = False


        'Me.Close()

    End Sub


    Sub UpdateDatabaseSave(ByVal strPathT As String, strD As String)

        'this routine will populate the table tblWordDocs

        Dim dtbl1 As System.Data.DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim strPath As String
        Dim intL As Int64
        Dim Count1 As Int16
        Dim Count10 As Short
        Dim intMax As Int64

        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim str4 As String
        Dim strSQL As String
        Dim dt As Date
        Dim idHere As Int64
        Dim var1, var2
        Dim intVer As Int64


        Dim intRow As Int32
        intRow = Me.dgvReportStatements.CurrentRow.Index
        idHere = Me.dgvReportStatements("ID_TBLWORDSTATEMENTS", intRow).Value


        Dim strT As String

        Dim strPathTT As String

        varTitleAfter = Me.CHARTITLE.Text


        dtbl1 = tblWordStatements
        strT = "tblWorddocs"
        dt = Now

        Dim constr As String
        If boolGuWuAccess Then
            constr = constrIni
        ElseIf boolGuWuSQLServer Then
            constr = constrIni '"Provider=SQLOLEDB;" & constrIni
        ElseIf boolGuWuOracle Then
            constr = constrIniGuWuODBC
        End If

        Dim myConnection As OleDb.OleDbConnection
        Dim myCommand As New OleDb.OleDbCommand

        myConnection = New OleDb.OleDbConnection(constr)
        myConnection.Open()
        myCommand.Connection = myConnection
        myCommand.CommandType = CommandType.Text


        'get version
        If Me.chkCreateNew.Checked Then
            intVer = 1
        Else
            intVer = GetWordVersion(idHere, True)
        End If

        If gboolET Then
        Else
            'first delete existing records
            'get the most recent version


            strSQL = "DELETE FROM " & strT & " WHERE ID_TBLWORDSTATEMENTS = " & idHere & " AND INTWORDVERSION = " & intVer

            myCommand.CommandText = strSQL
            'cmd.ExecuteNonQuery()
            Try
                myCommand.ExecuteNonQuery()
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try

        End If


        'now get new maxid
        intMax = GetMaxID("tblWordDocs", 1, True)

        Dim boolE As Boolean

        Count1 = 0

        Dim fs As FileStream

        strReport = strPathT

        If Me.ov1.IsOpened Then

            Dim wdDoc As Microsoft.Office.Interop.Word.Document
            Try
                wdDoc = Me.ov1.ActiveDocument
            Catch ex As Exception

            End Try

            Try
                Call SetWordStuff(wdDoc)
            Catch ex As Exception

            End Try

            Try
                wdDoc.Close()
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                wdDoc = Nothing
            Catch ex As Exception
                var1 = var1
            End Try


            Try
                Me.ov1.CloseDoc()
            Catch ex As Exception
                var1 = var1
            End Try



        End If

        fs = File.OpenRead(strPathT)

        Pause(0.25) 'give open a chance to complete

        intL = fs.Length


        Dim strW As String
        Dim int2 As Short
        Dim b(intL) As Byte
        Dim chars() As Char
        Dim c As Char
        Dim Count2 As Int64


        fs.Read(b, 0, b.Length)

        'Dim temp As UTF8Encoding = New UTF8Encoding(True)
        Dim temp As System.Text.Encoding = System.Text.Encoding.UTF8
        Dim utf8Decoder As Decoder = Encoding.UTF8.GetDecoder()
        Dim charCount As Integer = utf8Decoder.GetCharCount(b, 0, b.Length)
        chars = New Char(charCount - 1) {}
        Dim charsDecodedCount As Integer = utf8Decoder.GetChars(b, 0, b.Length, chars, 0)

        strW = ""
        Count1 = 0
        Count2 = 0

        If gboolET Then
            If Me.chkCreateNew.Checked Then
            Else
                'increment intver
                intVer = intVer + 1
            End If
        End If

        'Write in blocks of 2000 characters.
        'NDL Note: used "charCount - 1" as that is what was used previously.  
        '          "chars.Length" yields a different (slightly smaller) number.
        Dim intLastSegment As Int32 = Math.Truncate((charCount - 1) / 2000)
        For Count1 = 0 To intLastSegment
            intMax = intMax + 1
            If (Count1 <> intLastSegment) Then  'Write in blocks of 2000
                strW = New String(chars, Count1 * 2000, 2000)
            Else
                'Write the remainding characters
                strW = New String(chars, Count1 * 2000, (charCount - 1) Mod 2000)
            End If
            strW = Replace(strW, ChrW(39), ChrW(8217), 1, -1, CompareMethod.Text)
            'strSQL = "INSERT INTO " & strT & "  VALUES (" & intMax & "," & idHere & ",'" & strW & "',#" & dt & "#," & intVer & ");"
            strSQL = "INSERT INTO " & strT & "  VALUES (" & intMax & "," & idHere & ",'" & strW & "'," & ReturnDate(dt) & "," & intVer & ");"

            myCommand.CommandText = strSQL

            Try
                myCommand.ExecuteNonQuery()
            Catch ex As Exception
                var1 = ex.Message
                var1 = var1
            End Try
            strW = ""
        Next

        fs.Close()
        fs.Dispose()

        Call PutMaxID("tblWordDocs", intMax)

        If gboolET And Me.chkCreateNew.Checked = False Then

            id = GetMaxID("TBLWORDSTATEMENTSVERSIONS", 1, True) 'max comes incremented by 1
            '20190219 LEE: Don't need anymore. Used GetMaxID
            'Call PutMaxID("TBLWORDSTATEMENTSVERSIONS", id)

            Dim nr As DataRow = tblWordStatementsVERSIONS.NewRow
            nr.BeginEdit()
            nr("ID_TBLWORDSTATEMENTSVERSIONS") = id
            nr("ID_TBLWORDSTATEMENTS") = idHere
            nr("INTWORDVERSION") = intVer
            nr("CHARDESCRIPTION") = strD

            nr("ID_TBLPERSONNEL") = id_tblPersonnel
            nr("ID_TBLUSERACCOUNTS") = id_tblUserAccounts
            nr("CHARUSERID") = gUserID

            nr("UPSIZE_TS") = dt
            nr.EndEdit()
            tblWordStatementsVERSIONS.Rows.Add(nr)

        End If

        'clear audittrailtemp
        tblAuditTrailTemp.Clear()
        idSE = 0

        Call FillAuditTrailTemp(tblWordStatementsVERSIONS)

        If boolGuWuOracle Then
            'Try
            '    ta_TBLWORDSTATEMENTSVERSIONS.Update(TBLWORDSTATEMENTSVERSIONS)
            'Catch ex As DBConcurrencyException
            '    'ds2005.TBLWORDSTATEMENTSVERSIONS.Merge('ds2005.TBLCONFIGURATION, True)
            'End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_TBLWORDSTATEMENTSVERSIONSAcc.Update(TBLWORDSTATEMENTSVERSIONS)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLWORDSTATEMENTSVERSIONS.Merge('ds2005Acc.TBLCONFIGURATION, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_TBLWORDSTATEMENTSVERSIONSSQLServer.Update(TBLWORDSTATEMENTSVERSIONS)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLWORDSTATEMENTSVERSIONS.Merge('ds2005Acc.TBLCONFIGURATION, True)
            End Try
        End If

        'now the data is saved to the database
        'I shouldn't need to re-open it

        Pause(0.25)

        If gboolAuditTrail Then

            Call FillAuditTrailTemp(tblWorddocs)
        End If

        If gboolAuditTrail Then
            Dim dt1 As DateTime
            dt1 = Now

            'record tblaudittrailtemp
            Call RecordAuditTrail(False, dt1)
        End If


        myCommand.Connection = Nothing
        If myConnection.State = ADODB.ObjectStateEnum.adStateOpen Then
            myConnection.Close()
        End If
        myConnection = Nothing

end1:


    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click


        Dim strM As String
        Dim intR As Short
        Dim strT As String = Me.cmdSave.Text
        Dim strT1 As String

        If InStr(1, strT, "Check", CompareMethod.Text) > 0 Then
            strM = "Do you wish to Check In?"
            strT1 = "Check In?"
        Else
            strM = "Do you wish to Save?"
            strT1 = "Save"

            intR = MsgBox(strM, vbYesNo, strT1)
            If intR = 6 Then
            Else
                GoTo end1
            End If

        End If

   

        Dim strDescr As String = "Saved changes"
        If Me.chkCreateNew.Checked Then
            strDescr = "Initial version"
        End If

        Dim var1
        Dim Count1 As Short


        If gboolET Then
            Dim frm1 As New frmVersionDescr
            If Me.chkCreateNew.Checked Then
                frm1.rtbD.Text = strDescr
            End If
            frm1.ShowDialog()
            If frm1.boolCancel Then
                frm1.Close()
                frm1.Dispose()
                GoTo end1
            Else
                strDescr = frm1.rtbD.Text
                frm1.Close()
                frm1.Dispose()
            End If
        Else
            strDescr = "Saved changes"
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

            'clear audittrailtemp
            tblAuditTrailTemp.Clear()
            idSE = 0

        End If

        Call EnableButtons("Save")


        Me.cmdAddStatement.Enabled = True

        Call DoWrite()

        Call DoSave(strDescr)

        'after this action, there seems to be two Word processes open
        'they must be killed
        If Me.ov1.IsOpened Then
            Dim wdDoc As Microsoft.Office.Interop.Word.Document
            Try
                wdDoc = Me.ov1.ActiveDocument
            Catch ex As Exception
                var1 = var1
            End Try


            Try
                Call SetWordStuff(wdDoc)
            Catch ex As Exception

            End Try

            Try
                wdDoc.Application.Quit()
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                Me.ov1.CloseDoc()
            Catch ex As Exception
                var1 = var1
            End Try
        Else
            var1 = var1
        End If

        Call DoReadOnly()

        Call ClearTemp()

        'make sure first row in dgvVersions
        Try
            Me.dgvVersions.Rows(0).Selected = True

        Catch ex As Exception

        End Try

        Try
            'Me.dgvVersions.CurrentCell = Me.dgvVersions.CurrentRow.Cells(1)
            'find 1st visible column
            For Count1 = 0 To Me.dgvVersions.ColumnCount - 1
                If Me.dgvVersions.Columns(Count1).Visible Then
                    Me.dgvVersions.CurrentCell = Me.dgvVersions.Item(Count1, 0)
                    Exit For
                End If
            Next


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



end2:

        Me.dgvReportStatements.Visible = True
        Me.dgvVersions.Visible = True

        Me.chkCreateNew.Checked = False


end1:

    End Sub


    Sub LoadAFR(ByVal boolOV As Boolean, boolFromNew As Boolean, idWS As Int64, idVer As Int16)

        'look here for compatability mode stuff
        'http://office.microsoft.com/en-us/word-help/use-word-2013-to-open-documents-created-in-earlier-versions-of-word-HA102749315.aspx

        Dim var1, var2
        Dim dtbl2 As DataTable
        Dim rows1() As DataRow
        Dim rows2() As DataRow
        Dim strF As String
        Dim strS As String
        Dim intL As Int64
        'Dim strpathT As String 'this is public
        Dim Count1 As Int16
        Dim strW As String
        Dim fs As FileStream
        Dim strM As String
        Dim intRow As Int16
        Dim id As Int64
        Dim intV As Int64

        Dim dgv1 As DataGridView = Me.dgvReportStatements
        Dim dgv2 As DataGridView = Me.dgvVersions

        If boolFromNew Then
            id = idWS
            intV = idVer
        Else
            If dgv1.CurrentRow Is Nothing Then
                If dgv1.RowCount = 0 Then
                    GoTo end1
                Else
                    dgv1.Rows(0).Selected = True
                End If
            End If

            intRow = dgv1.CurrentRow.Index
            id = dgv1("ID_TBLWORDSTATEMENTS", intRow).Value

            If dgv2.CurrentRow Is Nothing Then
                If dgv2.RowCount = 0 Then
                    GoTo end1
                Else
                    dgv2.Rows(0).Selected = True
                End If
            End If

            intRow = dgv2.CurrentRow.Index
            intV = dgv2("INTWORDVERSION", intRow).Value
        End If

        strF = "ID_TBLWORDSTATEMENTS = " & id & " AND INTWORDVERSION = " & intV
        strS = "ID_TBLWORDDOCS ASC"

        Call OpenWordDocs(id, intV)

        dtbl2 = tblWorddocs
        rows2 = dtbl2.Select(strF, strS)
        intL = rows2.Length

        Cursor.Current = Cursors.WaitCursor

        Dim boolE As Boolean
        boolE = True
        Count1 = 0
        strPathT = ""

        '20140228
        'try using createxml

        strPathT = Createxml(id, intV)
        strReport = strPathT
        Me.lblSection.Text = strReport

        Pause(0.25)

        If gDoPDF Then
            GoTo end1
        End If

        If boolOV Then

            Try
                Me.ov1.Open(strPathT, "Word.Application") 'v8
            Catch ex As Exception
                var1 = ex.Message
                MsgBox("Problem opening Word .xml document." & ChrW(10) & ChrW(10) & var1)
            End Try
        End If

        var1 = var1

        'If boolOV Then

        '    '20142028
        '    Try
        '        Me.ov1.CloseDoc() 'v8
        '        'Me.ov1.Close() 'v6
        '    Catch ex As Exception
        '        MsgBox("LoadAFR (Me.ov1.CloseDoc()): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
        '    End Try

        '    'Me.ov1.CreateNew("Word.Document")
        '    Threading.Thread.Sleep(250) 'give the doc some time to establish, or get a read-only error

        '    Dim wdDoc As Microsoft.Office.Interop.Word.Document
        '    'before using ov1 to open doc
        '    'must open plain wd to find ShowWindowsInTaskbar setting
        '    If boolTested Then
        '    Else
        '        Dim wd As New Microsoft.Office.Interop.Word.Application
        '        boolSTB = wd.Application.ShowWindowsInTaskbar
        '        wd.Quit()
        '        wd = Nothing
        '        boolTested = True
        '    End If

        '    Try
        '        'Me.ov1.Open(strPathT) 'v6
        '        Try
        '            wdDoc = Me.ov1.ActiveDocument
        '            wdDoc.Application.Quit()
        '            Me.ov1.CloseDoc()
        '        Catch ex As Exception
        '            var1 = var1
        '        End Try

        '        Try
        '            Me.ov1.CloseDoc()
        '        Catch ex As Exception
        '            var1 = var1
        '        End Try

        '        Try
        '            Me.ov1.Open(strPathT, "Word.Application") 'v8
        '        Catch ex As Exception
        '            Try
        '                'Call ovInit()
        '                'Call DoWrite()
        '                Me.ov1.Open(strPathT, "Word.Application") 'v8
        '            Catch ex1 As Exception
        '                MsgBox("DoWrite (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ChrW(10) & strPathT & ChrW(10) & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
        '            End Try
        '        End Try


        '    Catch ex As Exception
        '        MsgBox("LoadAFR (Me.ov1.Open(strPathT, 'Word.Application') 'v8): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
        '    End Try

        'Me.ov1.Refresh()


        'Try
        '    wdDoc = Me.ov1.ActiveDocument

        '    'wdDoc.Application.ShowWindowsInTaskbar

        'Catch ex As Exception
        '    Exit Sub
        'End Try

        'Try
        '    wdDoc.Application.ShowWindowsInTaskbar = boolSTB
        'Catch ex As Exception

        'End Try


        'Else


        'End If


end1:

        Cursor.Current = Cursors.Default

    End Sub


    Private Sub cmiFieldCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmiFieldCode.Click
        Call InsertFC(False)
    End Sub

    Sub RefreshAFR()
        Try
            'Me.afrWord.Refresh()
            Me.ov1.Refresh()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub dgvReportStatements_CellContextMenuStripNeeded(sender As Object, e As DataGridViewCellContextMenuStripNeededEventArgs) Handles dgvReportStatements.CellContextMenuStripNeeded

    End Sub

    Private Sub dgvReportStatements_DoubleClick(sender As Object, e As EventArgs) Handles dgvReportStatements.DoubleClick



    End Sub


    Private Sub dgvReportStatements_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvReportStatements.SelectionChanged

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        'NDL: Do nothing if programmatically changed during creation of a new template
        If boolcmdAddStatementJustClicked = True Then
            Exit Sub
        End If

        Try
            Dim var1
            var1 = Me.dgvReportStatements.CurrentRow.Index
            var1 = var1 'debug

            If Me.dgvVersions.Visible Then

                If Me.ov1.IsOpened Then

                    Dim wdDoc As Microsoft.Office.Interop.Word.Document

                    Try
                        wdDoc = Me.ov1.ActiveDocument
                    Catch ex As Exception
                        var1 = var1
                    End Try


                    Try
                        wdDoc.Application.DisplayAlerts = False
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        Call SetWordStuff(wdDoc)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        Me.ov1.CloseDoc()
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        wdDoc.Application.DisplayAlerts = True
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    Try
                        wdDoc.Application.Quit(False)
                    Catch ex As Exception
                        var1 = var1
                    End Try

                    'Try
                    '    Me.ov1.CloseDoc()
                    'Catch ex As Exception
                    '    var1 = var1
                    'End Try
                Else
                    var1 = var1
                End If

                Call ReportStatementsChange()

                var1 = var1

            End If
        Catch ex As Exception

        End Try

    End Sub

    Sub ReportStatementsChange()

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        Dim id As Int64
        Dim dgv1 As DataGridView = Me.dgvReportStatements
        Dim dgv2 As DataGridView = Me.dgvVersions
        Dim intRow As Int32

        If dgv1.CurrentRow Is Nothing Then
            If dgv1.RowCount = 0 Then
                GoTo end1
            Else
                dgv1.Rows(0).Selected = True
            End If
        End If

        intRow = dgv1.CurrentRow.Index

        id = dgv1("ID_TBLWORDSTATEMENTS", intRow).Value

        Dim strF As String
        Dim strS As String

        strF = "ID_TBLWORDSTATEMENTS = " & id
        strS = "INTWORDVERSION DESC"

        Me.CHARTITLE.Text = dgv1("CHARTITLE", intRow).Value

        Dim dv As DataView



        Try
            'dv = dgv2.DataSource
            Dim boolS As Boolean
            boolS = boolHold
            boolHold = True
            dv = dgv2.DataSource
            dv.RowFilter = strF
            dv.Sort = strS
            boolHold = boolS
            If Me.cmdShow.Visible Or boolEdit Then
                'If Me.cmdShow.Visible Then
            Else
                Call VersionChange()
            End If
            dgv1.AutoResizeRows()
            dgv2.AutoResizeRows()
        Catch ex As Exception

        End Try


end1:

    End Sub

    Sub VersionChange()

        If boolFormLoad Then
            Exit Sub
        End If

        If Me.cmdShow.Visible Then
            Exit Sub
        End If

        Dim strM As String
        Dim intM As Short
        Dim dgv1 As DataGridView
        Dim dgv2 As DataGridView
        Dim intRow As Short
        Dim bool As Boolean

        dgv2 = Me.dgvVersions

        If dgv2.CurrentRow Is Nothing Then
            If dgv2.RowCount = 0 Then
                GoTo end1
            End If
        End If

        Call LoadAFR(True, False, 0, 0)

        Call DoReadOnly()

end1:

    End Sub


    Private Sub frmWordStatement_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseEnter


    End Sub


    Sub DoEditTitle(btn As String)

        Select Case btn
            Case "Edit"
                Me.cmdCancel.Enabled = True
                Me.cmdEditStatements.Enabled = False
                Me.cmdSaveStatements.Enabled = True

                Me.CHARTITLE.Enabled = True
                Me.CHARTITLE.ReadOnly = False
                Me.CHARTITLE.Focus()
            Case "Cancel", "Save"
                Me.cmdCancel.Enabled = False
                Me.cmdEditStatements.Enabled = True
                Me.cmdSaveStatements.Enabled = False

                Me.CHARTITLE.Enabled = False
                Me.CHARTITLE.ReadOnly = True
                'Me.CHARTITLE.Focus()

                Me.panSave.Enabled = True
                Me.panList.Enabled = True
                Me.panButtons.Enabled = True

                Me.panEditReports.Visible = False

        End Select

    End Sub

    Private Sub cmdEditStatements_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEditStatements.Click

        Call DoEditTitle("Edit")

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.CHARTITLE.Text = Me.dgvReportStatements("CHARTITLE", Me.dgvReportStatements.CurrentRow.Index).Value

        Me.panEditReports.Visible = False

        Call DoEditTitle("Cancel")

        'Dim dgv As DataGridView
        'Dim intRow As Short

        'dgv = Me.dgvReportStatements
        'intRow = dgv.CurrentRow.Index

        'Dim str1 As String

        'str1 = dgv("CHARTITLE", intRow).Value
        'Me.CHARTITLE.Text = str1

        'Call DoEditTitle("Cancel")

        'Call ClickMainCancel()

    End Sub

    

    Function RNUnique(ByVal tbl As System.Data.DataTable, ByVal strField As String, ByVal strInput As String) As Boolean

        Dim rows() As DataRow
        Dim strF As String
        Dim strM As String
        Dim strM1 As String
        Dim strStatus As String

        RNUnique = False

        strF = strField & " = '" & strInput & "'"
        tbl = tblWordStatements
        rows = tbl.Select(strF)
        If rows.Length = 0 Then
            RNUnique = True
        Else
            RNUnique = False
        End If

        If RNUnique = False Then

            strStatus = rows(0).Item("CHARWORDSTATEMENT")
            strM1 = "Thep proposed report template name conflicts with an existing " & strStatus & " template." & ChrW(10) & ChrW(10) & "Name must be unique."
            MsgBox(strM1, vbInformation, "Invalid action...")

        End If

    End Function

    Private Sub cmdAddStatement_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddStatement.Click

        'NDL There are some _selectionchanged routines we want to avoid during this subroutine.  We will
        'use boolcmdAddStatementJustClicked as a signal to skip them.
        boolcmdAddStatementJustClicked = True


        Dim frm As New frmReportTemplateAdd

        Dim Count1 As Int16
        Dim Count2 As Int16
        Dim dgv1 As DataGridView = Me.dgvReportStatements
        Dim dgv2 As DataGridView = frm.dgvReportStatements

        Dim intRow1 As Int16
        Dim intRow2 As Int16

        Dim dv1 As DataView = New DataView(tblWordStatements, "ID_TBLWORDSTATEMENTS > 0", "CHARTITLE ASC", DataViewRowState.CurrentRows)
        dv1.AllowDelete = False
        dv1.AllowEdit = False
        dv1.AllowNew = False

        dgv2.DataSource = dv1 'dgv1.DataSource
        dgv2.RowHeadersWidth = dgv1.RowHeadersWidth
        dgv2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Count1 = 0 To dgv1.ColumnCount - 1
            dgv2.Columns(Count1).HeaderText = dgv1.Columns(Count1).HeaderText
            dgv2.Columns(Count1).Visible = dgv1.Columns(Count1).Visible
            'dgv2.Columns(Count1).Width = dgv1.Columns(Count1).Width
        Next
        'make CHARWORDSTATEMENT visible
        dgv2.Columns("CHARWORDSTATEMENT").Visible = True
        dgv2.Columns("CHARWORDSTATEMENT").HeaderText = "Status"
        dgv2.AutoResizeColumns()

        dgv1 = Me.dgvVersions
        dgv2 = frm.dgvVersions
        dgv2.DataSource = dgv1.DataSource
        dgv2.RowHeadersWidth = dgv1.RowHeadersWidth
        For Count1 = 0 To dgv1.ColumnCount - 1
            dgv2.Columns(Count1).HeaderText = dgv1.Columns(Count1).HeaderText
            dgv2.Columns(Count1).Visible = dgv1.Columns(Count1).Visible
            dgv2.Columns(Count1).Width = dgv1.Columns(Count1).Width
        Next

        If Me.lblVersions.Visible Then
            frm.boolShowVersions = True
        Else
            frm.boolShowVersions = False
        End If

        frm.ShowDialog()

        Dim strInput As String
        Dim strPath As String
        Dim strM1 As String
        Dim var1, var2
        Dim boolCancel As Boolean
        Dim idWS As Int64 = 0
        Dim idVer As Int16 = 1

        strInput = frm.txtName.Text 'this has been validated in frm
        strPath = frm.txtFilePath.Text 'this has been validated in frm

        dgv1 = frm.dgvReportStatements
        dgv2 = frm.dgvVersions

        Try
            intRow1 = dgv1.CurrentRow.Index
            idWS = dgv1("ID_TBLWORDSTATEMENTS", intRow1).Value
        Catch ex As Exception

        End Try

        Try
            intRow2 = dgv2.CurrentRow.Index
            idVer = dgv2("INTWORDVERSION", intRow2).Value
        Catch ex As Exception

        End Try

        boolCancel = frm.boolCancel

        frm.Dispose()

        If boolCancel Then
        Else


            varTitleAfter = strInput
            Me.chkCreateNew.Checked = True

            If frm.rbBlank.Checked Then
                Call MakeBlankReportTemplate(strInput, strPath, False, False, idWS, idVer)
            ElseIf frm.rbDocument.Checked Then
                Call MakeBlankReportTemplate(strInput, strPath, True, False, idWS, idVer)
            ElseIf frm.rbTemplate.Checked Then
                Call MakeBlankReportTemplate(strInput, strPath, False, True, idWS, idVer)
            Else

            End If

        End If

        'NDL There are some _selectionchanged routines we want to avoid during this subroutine.  We will
        'use boolcmdAddStatementJustClicked as a signal to skip them.
        boolcmdAddStatementJustClicked = False

end4:

 

    End Sub

    Sub MakeBlankReportTemplate(strInput As String, strPath As String, boolExistingWord As Boolean, boolExistingTemplate As Boolean, idWS As Int64, idVer As Int16)

        Dim strM1 As String
        Dim var1, var2
        Dim Count1 As Int16

        'make a new document
        Dim boolE As Boolean
        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim boolSTB As Boolean = True
        Try
            boolSTB = wd.Application.ShowWindowsInTaskbar
        Catch ex As Exception

        End Try

        Dim strS As String
        Dim strPathX As String

        boolE = True
        Count1 = 0
        strPathT = ""

        '20140228
        If boolExistingWord Then
            strPathT = strPath
        ElseIf boolExistingTemplate Then

            'get existing template
            Call LoadAFR(True, True, idWS, idVer)

        Else
            strPathT = GetNewTempFile(True)
        End If

        strReport = strPathT
        Me.lblSection.Text = strPathT

        'get an id
        'now get new maxid and assign it to public idNew ID_TBLWORDSTATEMENTS
        Dim tblM As System.Data.DataTable
        Dim rowM() As DataRow
        Dim strM As String
        Dim intMax As Int64
        Dim dt As Date = Now

        intMax = GetMaxID("TBLWORDSTATEMENTS", 1, True) 'is incremented one already
        idNew = intMax
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("TBLWORDSTATEMENTS", intMax)

        '****
        Dim newRow As DataRow = tblWordStatements.NewRow
        newRow.BeginEdit()
        newRow("ID_TBLWORDSTATEMENTS") = idNew
        newRow("ID_TBLCONFIGBODYSECTIONS") = 341
        newRow("INTWORDTABLENUMBER") = 331
        newRow("CHARTITLE") = strInput
        newRow("CHARWORDSTATEMENT") = "Active"
        newRow("UPSIZE_TS") = dt
        newRow.EndEdit()

        tblWordStatements.Rows.Add(newRow)

        Me.dgvReportStatements.AutoResizeRows()

        Dim idSS As Int64 = idNew

        'now add an increment
        intMax = GetMaxID("TBLWORDSTATEMENTSVERSIONS", 1, True) 'is incremented one already
        idNew = intMax
        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("TBLWORDSTATEMENTSVERSIONS", intMax)

        Dim newRow1 As DataRow = TBLWORDSTATEMENTSVERSIONS.NewRow
        newRow1.BeginEdit()
        newRow1("ID_TBLWORDSTATEMENTSVERSIONS") = idNew
        newRow1("ID_TBLWORDSTATEMENTS") = idSS
        newRow1("INTWORDVERSION") = 1
        newRow1("CHARDESCRIPTION") = "Initial Version"

        newRow1("ID_TBLPERSONNEL") = id_tblPersonnel
        newRow1("ID_TBLUSERACCOUNTS") = id_tblUserAccounts
        newRow1("CHARUSERID") = gUserID

        newRow1("UPSIZE_TS") = dt
        newRow1.EndEdit()

        TBLWORDSTATEMENTSVERSIONS.Rows.Add(newRow1)

        'now choose the appropriate row in dgv
        Dim dgv As DataGridView = Me.dgvReportStatements
        Dim dgv1 As DataGridView = Me.dgvVersions

        For Count1 = 0 To dgv.RowCount - 1
            var1 = dgv("ID_TBLWORDSTATEMENTS", Count1).Value
            If var1 = idSS Then
                'dgv.Rows(Count1).Selected = True
                dgv.CurrentCell = dgv.Item("CHARTITLE", Count1)
                If boolExistingWord Then
                Else
                    'call this again 
                    Call ReportStatementsChange()
                End If
                Exit For
            End If
        Next

        '****

        'now disable/enable buttons
        'Me.CHARTITLE.Enabled = True
        'Me.CHARTITLE.ReadOnly = False
        'Me.CHARTITLE.Visible = False

        Me.dgvReportStatements.Enabled = False
        Me.dgvReportStatements.BackgroundColor = Color.Gray

        Me.dgvVersions.Enabled = False
        Me.dgvVersions.BackgroundColor = Color.Gray

        Me.lblDGV.Visible = False ' True

        Me.cmdExit.Text = "E&xit and Cancel"
        Me.cmdExit.Text = "E&xit"

        'Me.CHARTITLE.Enabled = True

        'Me.panEditReports.Visible = True

        Me.cmdAddStatement.Enabled = False

end1:

        'Call RefreshAFR()

        Me.ov1.Refresh()

        Me.dgvReportStatements.AutoResizeRows()

        Me.lblBlink.Visible = False

        Call DoWrite()

        Call EnableButtons("Edit")

        'Call InsertNewDocument()


        If boolExistingWord Then

            Try
                'Me.afrWord.Open(strPathT, boolRO, "Word.Document", "a", "a")
                'Me.ov1.Open(strPathT, "Word.Application") 'v8
                Me.ov1.OpenWord(strPathT)

            Catch ex As Exception
                MsgBox("cmdAddStatement (me.ov1.open(...)): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

        ElseIf boolExistingTemplate Then

            'do nothing

        Else

            '******

            wd.Documents.Add()

            'wd.ActiveDocument.SaveAs(FileName:=strPathT, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXML)
            wd.ActiveDocument.SaveAs(FileName:=strPathT, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)

            Try
                wd.Application.ShowWindowsInTaskbar = boolSTB
            Catch ex As Exception

            End Try

            wd.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges)

            Threading.Thread.Sleep(100) 'give the doc some time to save, or get a read-only error

            Try
                Me.ov1.CloseDoc()
            Catch ex As Exception
                MsgBox("cmdAddStatement (me.ov1.CloseDoc()): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

            Pause(0.25)

            Try
                'Me.afrWord.Open(strPathT, boolRO, "Word.Document", "a", "a")
                Me.ov1.Open(strPathT, "Word.Application") 'v8
                'Me.ov1.Open(strPathT) 'v6
            Catch ex As Exception
                MsgBox("cmdAddStatement (me.ov1.open(...)): " & ex.Message, MsgBoxStyle.Exclamation, "Problem...")
            End Try

            '******
        End If



        'Me.dgvReportStatements.Visible = False
        'Me.dgvVersions.Visible = False


end4:
        boolcmdAddStatementJustClicked = False

    End Sub

    Private Sub cmdWord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWord.Click

        Dim strM As String

        'Public boolReport As Boolean = False 'boolReport is a Report Template
        strM = "User does not have permission to open document in Word."
        If boolReport Then
            If BOOLALLOWREPORTTEMPLATEWORD Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            If BOOLALLOWFINALREPORTWORD Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If


        If Me.ov1.IsOpened Then
        Else
            strM = "A document must be loaded."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        Call DoWrite()

        Call GoToWord(False)

        Call ClearTemp()

        Exit Sub

        '    ActiveDocument.SaveAs2(FileName:="assjj.docx", FileFormat:= _
        'wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        ':=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        ':=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        'SaveAsAOCELetter:=False, CompatibilityMode:=14)
        '    Selection.TypeText(Text:=";aljf")
        '    Selection.TypeParagraph()
        '    ActiveDocument.SaveAs2(FileName:="assjj.docm", FileFormat:= _
        '        wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
        '        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        '        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        '        :=False, SaveAsAOCELetter:=False, CompatibilityMode:=14)


        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        Dim strPath As String
        Dim strR As String
        Dim boolM As Boolean
        Dim bool2007 As Boolean
        Dim ver

        'legend
        'ver = wdDoc.Application.Version
        '14=2010
        '12=2007
        '10=2003

        'Select Case Val(strAppVersion)
        '    Case Is < 9
        '        GetEdition = "Pre Office 2000: " & strERR_MSG
        '    Case Is < 10
        '        'Office 2000            strSku = Mid$(strGuid, 4, 2)            
        '        GetEdition = GetEdition2000(strSku)
        '    Case Is < 11
        '        'Office 2002            strSku = Mid$(strGuid, 4, 2)            
        '        GetEdition = GetEdition2002(strSku)
        '    Case Is < 12
        '        'Office 2003            strSku = Mid$(strGuid, 4, 2)            
        '        GetEdition = GetEdition2003(strSku)
        '    Case Is < 13
        '        'Office 2007            strSku = Mid$(strGuid, 11, 4)            
        '        GetEdition = GetEdition2007(strSku)
        '    Case Is < 15
        '        'Office 2010            strSku = Mid$(strGuid, 11, 4)           
        '        GetEdition = GetEdition2010(strSku)
        '    Case Is < 16
        '        'Office 2013            strSku = Mid$(strGuid, 11, 4)            
        '        GetEdition = GetEdition2013(strSku)
        '    Case Else
        '        GetEdition = "Post Office 2013: " & strERR_MSG
        'End Select


        boolM = False
        Try
            wdDoc = Me.ov1.ActiveDocument
            ver = CInt(wdDoc.Application.Version)
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If
            strPath = GetNewTempFile(True)
            boolHasMacro = False
            If bool2007 Then
                'Note: Word2007 has no compatability mode parameter
                If wdDoc.HasVBProject Then
                    strR = Replace(strPath, ".xml", ".docm", 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
                    boolM = True
                    boolHasMacro = True
                Else
                    strR = Replace(strPath, ".xml", ".docx", 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                    boolM = False
                End If
            Else
                If wdDoc.HasVBProject Then
                    strR = Replace(strPath, ".xml", ".docm", 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs2(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolM = True
                    boolHasMacro = True
                Else
                    strR = Replace(strPath, ".xml", ".docx", 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs2(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolM = False
                End If
            End If

            Try
                Call SetWordStuff(wdDoc)
            Catch ex As Exception

            End Try

        Catch ex As Exception

        End Try

        Pause(0.25)

        Me.Visible = False


        Try
            'Me.afrWord.Close()
            Me.ov1.CloseDoc() 'v8
        Catch ex As Exception

        End Try


        Try
            '20140228
            'added this because lots of winword processes open after closing frmWordStatement
            Me.ov1.Dispose()
        Catch ex As Exception

        End Try

        'Me.ov1.Close() 'v6

        Dim wd As New Microsoft.Office.Interop.Word.Application

        Dim boolSTB As Boolean = True
        Try
            boolSTB = wd.Application.ShowWindowsInTaskbar
        Catch ex As Exception

        End Try

        Try
            wd.Application.NormalTemplate.Saved = True
        Catch ex As Exception

        End Try

        If boolM Then
            wd.Documents.Open(FileName:=strR, Format:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
        Else
            wd.Documents.Open(FileName:=strR, Format:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
        End If

        'legend:
        '      Documents.Open(FileName:="page_4.html", ConfirmConversions:=False, _
        'ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        'PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        'WritePasswordTemplate:="", Format:=wdOpenFormatAuto, XMLTransform:="")

        'Try
        '    wd.Application.NormalTemplate.Saved = True
        'Catch ex As Exception

        'End Try

        wd.Visible = True


        Try
            wd.Application.ShowWindowsInTaskbar = boolSTB
        Catch ex As Exception

        End Try

        Me.Close()
        Me.Dispose()

end1:


    End Sub

    Private Sub cmdExit2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit2.Click

        '20181023 LEE:
        boolGoBack = True

        'Call DoExitDoc2()

        Call DoExit()

        Dim var1
        var1 = var1 'degub


    End Sub

    Sub DoExitDoc2()

        Try
            'Me.afrWord.Close()
            Me.ov1.CloseDoc() 'v8
            'Me.ov1.Close() 'v6
        Catch ex As Exception

        End Try

        '
        Try
            '20140228
            'added this because lots of winword processes open after closing frmWordStatement
            Me.ov1.Dispose()
        Catch ex As Exception

        End Try

        boolCancel = True

        Call ClearTemp()

        'Me.Visible = False

        Me.Close()
        Me.Dispose()
        'Me.Close()

        If BOOLFORCEFINALREPORTPDF And boolReport = False Then
        Else
            frmH.BringToFront()
        End If


    End Sub

    Private Sub pan1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles panEditReports.MouseEnter


    End Sub

    Private Sub pan1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles panEditReports.MouseLeave

    End Sub

    Private Sub panEdraw_GotFocus(sender As Object, e As System.EventArgs) Handles panEdraw.GotFocus

        '20170824 LEE: Not needed anymore

        'Dim strM As String

        'If Me.cmdEdit.Enabled And Me.pan2.Visible = False Then
        '    strM = "A: Please note that the 'Edit' button has not been clicked." & ChrW(10) & ChrW(10)
        '    strM = strM & "Any changes made to the document cannot be saved unless the 'Edit' button is clicked."
        '    MsgBox(strM, MsgBoxStyle.Information, "Non-edit mode...")
        'End If

    End Sub

    Private Sub panEdraw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles panEdraw.MouseClick



    End Sub


    Private Sub panEdraw_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles panEdraw.MouseEnter


    End Sub

    Sub GoToWord(boolPDF As Boolean)

        Dim strM As String
        Dim var1
        Dim strExt As String


        'If Me.cmdSave.Visible And gGoToWord = False Then
        '20160713 LEE: gGoToWord is deprecated. Default = False
        If gGoToWord = False Then

            strM = "Are you sure you wish to open this document in Microsoft" & ChrW(174) & " Word?" & ChrW(10) & ChrW(10)
            strM = strM & "If any changes have not been saved, they will be lost."

            strM = "The Loaded Document will be opened in Microsoft" & ChrW(174) & " Word." & ChrW(10) & ChrW(10)
            strM = strM & "This StudyDoc window will remain open."

            If boolReport Then
                var1 = MsgBox(strM, MsgBoxStyle.OkCancel, "Are you sure...")
                If var1 = 1 Then
                Else
                    Exit Sub
                End If
            Else
                If BOOLFORCEFINALREPORTPDF Then
                Else
                    var1 = MsgBox(strM, MsgBoxStyle.OkCancel, "Are you sure...")
                    If var1 = 1 Then
                    Else
                        Exit Sub
                    End If
                End If
            End If

        End If



        '    ActiveDocument.SaveAs2(FileName:="assjj.docx", FileFormat:= _
        'wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        ':=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        ':=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        'SaveAsAOCELetter:=False, CompatibilityMode:=14)
        '    Selection.TypeText(Text:=";aljf")
        '    Selection.TypeParagraph()
        '    ActiveDocument.SaveAs2(FileName:="assjj.docm", FileFormat:= _
        '        wdFormatXMLDocumentMacroEnabled, LockComments:=False, Password:="", _
        '        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        '        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        '        :=False, SaveAsAOCELetter:=False, CompatibilityMode:=14)


        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        Dim strPath As String
        Dim strR As String
        Dim boolM As Boolean
        Dim bool2007 As Boolean
        Dim ver As Short

        'legend
        'ver = wdDoc.Application.Version
        '14=2010
        '12=2007
        '10=2003

        'Select Case Val(strAppVersion)
        '    Case Is < 9
        '        GetEdition = "Pre Office 2000: " & strERR_MSG
        '    Case Is < 10
        '        'Office 2000            strSku = Mid$(strGuid, 4, 2)            
        '        GetEdition = GetEdition2000(strSku)
        '    Case Is < 11
        '        'Office 2002            strSku = Mid$(strGuid, 4, 2)            
        '        GetEdition = GetEdition2002(strSku)
        '    Case Is < 12
        '        'Office 2003            strSku = Mid$(strGuid, 4, 2)            
        '        GetEdition = GetEdition2003(strSku)
        '    Case Is < 13
        '        'Office 2007            strSku = Mid$(strGuid, 11, 4)            
        '        GetEdition = GetEdition2007(strSku)
        '    Case Is < 15
        '        'Office 2010            strSku = Mid$(strGuid, 11, 4)           
        '        GetEdition = GetEdition2010(strSku)
        '    Case Is < 16
        '        'Office 2013            strSku = Mid$(strGuid, 11, 4)            
        '        GetEdition = GetEdition2013(strSku)
        '    Case Else
        '        GetEdition = "Post Office 2013: " & strERR_MSG
        'End Select

        'before using ov1 to open doc
        'must open plain wd to find ShowWindowsInTaskbar setting
        If boolTested Then
        Else
            Dim wd1 As New Microsoft.Office.Interop.Word.Application
            boolSTB = wd1.Application.ShowWindowsInTaskbar
            wd1.Quit()
            wd1 = Nothing
            boolTested = True
        End If

        Dim strOrigPath As String

        boolM = False
        Try
            wdDoc = Me.ov1.ActiveDocument
            strOrigPath = wdDoc.FullName

            'save the doc
            'DONOT save the doc if in view mode
            'Edits are not allowed in this view
            If boolRO Then
            Else
                wdDoc.Save()
            End If

            'If boolRO Then
            '    'reopen original
            '    Me.ov1.CloseDoc()
            '    Me.ov1.OpenWord(strOrigPath)
            '    wdDoc = Me.ov1.ActiveDocument
            'End If

            Try
                wdDoc.Application.ShowWindowsInTaskbar = boolSTB
            Catch ex As Exception
                var1 = var1 'debug
            End Try
            ver = CInt(wdDoc.Application.Version)
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If
            strPath = GetNewTempFileReport(False)
            boolHasMacro = False
            If bool2007 Then
                'Note: Word2007 has no compatability mode parameter
                If wdDoc.HasVBProject Then
                    strExt = ".docm"
                    strR = Replace(strPath, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
                    boolM = True
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strR = Replace(strPath, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                    boolM = False
                End If
            Else
                If wdDoc.HasVBProject Then
                    strExt = ".docm"
                    strR = Replace(strPath, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs2(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolM = True
                    boolHasMacro = True
                Else
                    strExt = ".docx"
                    strR = Replace(strPath, ".xml", strExt, 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs2(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolM = False
                End If
            End If


            Try
                Call SetWordStuff(wdDoc)
            Catch ex As Exception

            End Try

        Catch ex As Exception

        End Try

        Pause(0.25)

        'close new file
        'wdDoc.Close()

        Me.ov1.CloseDoc(False)
        Pause(0.25)
        Me.ov1.OpenWord(strOrigPath)

        If boolRO Then
            Call DoReadOnly()
        End If

        '20160712 Don't close anymore
        'Me.Visible = False

        ''20160712 Don't close anymore
        'Try
        '    'Me.afrWord.Close()
        '    Me.ov1.CloseDoc() 'v8
        'Catch ex As Exception

        'End Try

        ''20160712 Don't close anymore
        'Try
        '    '20140228
        '    'added this because lots of winword processes open after closing frmWordStatement
        '    Me.ov1.Dispose()
        'Catch ex As Exception

        'End Try

        'Me.ov1.Close() 'v6

        Dim wd As New Microsoft.Office.Interop.Word.Application

        Try
            wd.Application.NormalTemplate.Saved = True
        Catch ex As Exception

        End Try

        'If boolM Then
        '    wd.Documents.Open(FileName:=strR, Format:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
        'Else
        '    wd.Documents.Open(FileName:=strR, Format:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
        'End If


        Try
            If boolM Then
                wd.Documents.Open(FileName:=strR)
            Else
                wd.Documents.Open(FileName:=strR)
            End If
        Catch ex As Exception

        End Try

        'legend:
        '      Documents.Open(FileName:="page_4.html", ConfirmConversions:=False, _
        'ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="", _
        'PasswordTemplate:="", Revert:=False, WritePasswordDocument:="", _
        'WritePasswordTemplate:="", Format:=wdOpenFormatAuto, XMLTransform:="")

        Try
            wd.Application.NormalTemplate.Saved = True
        Catch ex As Exception

        End Try

        If boolPDF Then
            Dim strP As String

            'strP = Replace(strReport, strExt, ".pdf", 1, -1, CompareMethod.Text)
            'strP = CreatePDF(wd, strReport)

            strP = Replace(strR, strExt, ".pdf", 1, -1, CompareMethod.Text)
            strP = CreatePDF(wd, strReport)

            Try
                wd.Application.ShowWindowsInTaskbar = boolSTB
            Catch ex As Exception

            End Try

            wd.Quit(False)

            'don't need to start in Reader
            'already done in CreatePDF
            Try
                System.Diagnostics.Process.Start(strP)
            Catch ex As Exception
                strM = "There was problem opening this file as a PDF."
                strM = strM & ChrW(10) & ChrW(10) & "It may be that there is not a configured default PDF viewer on this workstation."
                MsgBox(strM, MsgBoxStyle.Information, "Problem...")

            End Try
        Else
            wd.Visible = True
        End If

        Pause(2)

        Try
            wd.Application.ShowWindowsInTaskbar = boolSTB
        Catch ex As Exception
            var1 = var1 'debug
        End Try

        If boolReport Then
        Else
            If BOOLFORCEFINALREPORTPDF Then
                Call DoExit() ' ExitDoc2()
                Me.Visible = False
            End If
        End If

        ''20160712 Don't close anymore
        'Me.Close()
        'Me.Dispose()


    End Sub


    Private Sub cmdOpenPDF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOpenPDF.Click

        Dim strM As String

        'Public boolReport As Boolean = False 'boolReport is a Report Template
        strM = "User does not have permission to open document as PDF."
        If boolReport Then
            If BOOLALLOWREPORTTEMPLATEPDF Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            If BOOLALLOWPDFREPORT Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If


        If Me.ov1.IsOpened Then
        Else
            strM = "A document must be loaded."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        If boolReport Then 'Report Template
            If BOOLALLOWREPORTTEMPLATEPDF Then
            Else
                strM = "User does not have permissions to create a PDF file."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            'BOOLALLOWPDFREPORT = False'testing
            If BOOLALLOWPDFREPORT Then
            Else
                strM = "User does not have permissions to create a PDF file."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If

        Call DoWrite()

        Call GoToWord(True)

        Call ClearTemp()

end1:

    End Sub

    Private Sub cmdWord1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdWord1.Click

        Dim strM As String
        Dim var1

        'Public boolReport As Boolean = False 'boolReport is a Report Template
        strM = "User does not have permission to open document in Word."
        If boolReport Then
            If BOOLALLOWREPORTTEMPLATEWORD Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            If BOOLALLOWFINALREPORTWORD Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If

        If Me.cmdEdit.Enabled Then
            strM = "Document is not in Edit mode."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        'strM = "The Loaded Document will be opened in Microsoft" & ChrW(174) & " Word." & ChrW(10) & ChrW(10)
        'strM = strM & "This StudyDoc window will remain open."

        'var1 = MsgBox(strM, MsgBoxStyle.OkCancel, "Are you sure...")
        'If var1 = 1 Then
        'Else
        '    GoTo end1
        'End If



        Call DoWrite()

        Call GoToWord(False)

        'Call ClearTemp()

end1:

    End Sub


    Sub DoReadOnly()

        'wdAllowOnlyRevisions = 0,
        'wdAllowOnlyComments = 1,
        'wdAllowOnlyFormFields = 2,
        'wdAllowOnlyReading = 3,
        'wdNoProtection = -1,

        Dim var1
        Try
            Me.ov1.ProtectDoc(3)
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try


    End Sub

    Sub DoWrite()

        Try
            Me.ov1.UnProtectDoc()
        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        'a document must be loaded in ov1 in order get a handle

        Dim lng As Int64 = Me.ov1.GetOfficeHwnd

        MsgBox(lng)

        Me.ov1.SetAppFocus()

    End Sub


    Private Sub cmdEdit_Click(sender As System.Object, e As System.EventArgs) Handles cmdEdit.Click

        Call EditClick()

    End Sub

    Sub EditClick()

        Dim int1 As Int64
        'int1 = Me.dgvVersions.CurrentRow.Index


        If Me.boolEdit Then
            Me.cmdShow.Visible = False
        End If

        Call EnableButtons("Edit")

        Call VersionChange()

        Call DoWrite()

        Me.cmdAddStatement.Enabled = False

        Me.CHARTITLE.Enabled = False

end1:

        Call ClearTemp()

    End Sub

    Private Sub cmdCancelEdit_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancelEdit.Click

        Dim strM As String
        Dim intR As Short

        strM = "Do you wish to Cancel?"
        intR = MsgBox(strM, vbYesNo, "Cancel?")
        If intR = 6 Then
            Call ClickMainCancel()
        End If

    End Sub

    Sub ClickMainCancel()

        If Me.panEditReports.Visible Then
            Me.dgvReportStatements.Visible = True
            Me.dgvVersions.Visible = True

        Else


        End If

        If Me.chkCreateNew.Checked Then

            tblWordStatements.RejectChanges()
            tblWordStatementsVERSIONS.RejectChanges()
            Me.chkCreateNew.Checked = False

        End If

        Try 'If word document isn't open (for instance, tried to open exising word document, Word gave user "Read Only" option,
            'and user cancelled out of it), then ov1 doesn't have an "IsOpened" option.  That's why it is put into the "Try".
            'NDL 2-Feb-2016
            If Me.ov1.IsOpened Then
                Me.ov1.CloseDoc()
            End If
        Catch ex As Exception
        End Try

        Call ClearTemp()

        Call EnableButtons("Cancel")

        Me.cmdAddStatement.Enabled = True


    End Sub

    Private Sub cmdSaveStatements_Click(sender As System.Object, e As System.EventArgs) Handles cmdSaveStatements.Click

        Dim str1 As String
        Dim intRow As Integer

        Call DoSaveReportTitle()

        Dim dgv As DataGridView = Me.dgvReportStatements
        intRow = dgv.CurrentRow.Index

        str1 = Me.CHARTITLE.Text

        'do audit trail

        dgv("CHARTITLE", intRow).Value = str1

        Call DoEditTitle("Save")

end1:

    End Sub

    Sub DoSaveReportTitle()

        Dim dtbl As System.Data.DataTable
        dtbl = tblWordStatements
        Dim newRow As DataRow = dtbl.NewRow
        Dim var2, var3
        Dim dgv As DataGridView
        Dim dt As Date = Now
        Dim var1
        Dim Count1 As Int32

        Dim id As Int64
        Dim strTitle As String = Me.CHARTITLE.Text

        dgv = Me.dgvReportStatements
        id = dgv("ID_TBLWORDSTATEMENTS", dgv.CurrentRow.Index).Value

        Dim strF As String = "ID_TBLWORDSTATEMENTS = " & id
        Dim rows() As DataRow = dtbl.Select(strf)
        Try
            rows(0).BeginEdit()
            rows(0).Item("CHARTITLE") = strTitle
            rows(0).Item("UPSIZE_TS") = dt
            rows(0).EndEdit()
        Catch ex As Exception
            var1 = ex.Message
            var1 = var1
        End Try

        'must also do tblReportStatements
        rows = tblReportStatements.Select(strF)
        For Count1 = 0 To rows.Length - 1
            rows(Count1).BeginEdit()
            rows(Count1).Item("CHARSTATEMENT") = strTitle
            rows(Count1).EndEdit()
        Next
        tblReportStatements.AcceptChanges()

        'enter audittrail


        '*****

        Cursor.Current = Cursors.WaitCursor

        If gboolAuditTrail Then
            'clear audittrailtemp
            tblAuditTrailTemp.Clear()
            idSE = 0
            Call FillAuditTrailTemp(tblWordStatements)
        End If

        If boolGuWuOracle Then
            Try
                ta_tblWordStatements.Update(tblWordStatements)
            Catch ex As DBConcurrencyException
                'ds2005.TBLWORDSTATEMENTS.Merge('ds2005.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblWordStatementsAcc.Update(tblWordStatements)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLWORDSTATEMENTS.Merge('ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblWordStatementsSQLServer.Update(tblWordStatements)
            Catch ex As DBConcurrencyException
                'ds2005Acc.TBLWORDSTATEMENTS.Merge('ds2005Acc.TBLWORDSTATEMENTS, True)
            End Try
        End If


        If boolGuWuOracle Then
            Try
                ta_tblReportStatements.Update(tblReportStatements)
            Catch ex As DBConcurrencyException
                'ds2005.tblReportStatements.Merge('ds2005.tblReportStatements, True)
            End Try
        ElseIf boolGuWuAccess Then
            Try
                ta_tblReportStatementsAcc.Update(tblReportStatements)
            Catch ex As DBConcurrencyException
                'ds2005Acc.tblReportStatements.Merge('ds2005Acc.tblReportStatements, True)
            End Try
        ElseIf boolGuWuSQLServer Then
            Try
                ta_tblReportStatementsSQLServer.Update(tblReportStatements)
            Catch ex As DBConcurrencyException
                'ds2005Acc.tblReportStatements.Merge('ds2005Acc.tblReportStatements, True)
            End Try
        End If


        '****

        If gboolAuditTrail Then
            Call RecordAuditTrail(False, dt)
        End If


        '*****



        Me.dgvReportStatements.AutoResizeRows()

        boolHold = False

        varTitleBefore = ""
        varTitleAfter = Me.CHARTITLE.Text


    End Sub

    Sub DoSaveReportTitleBU()

        'this is supposed to be an update, not a new row

        Dim dtbl As System.Data.DataTable
        dtbl = tblWordStatements
        Dim newRow As DataRow = dtbl.NewRow
        Dim dt As Date
        Dim var2, var3
        Dim dgv As DataGridView

        boolHold = True

        dgv = Me.dgvReportStatements
        dt = Now
        var2 = dgv("ID_TBLCONFIGBODYSECTIONS", 0).Value
        var3 = dgv("INTWORDTABLENUMBER", 0).Value

        idNew = GetMaxID("TBLWORDSTATEMENTS", 1, True)

        newRow.BeginEdit()
        newRow("ID_TBLWORDSTATEMENTS") = idNew
        newRow("ID_TBLCONFIGBODYSECTIONS") = var2
        newRow("INTWORDTABLENUMBER") = var3
        newRow("CHARTITLE") = Me.CHARTITLE.Text
        newRow("CHARWORDSTATEMENT") = "Active"
        newRow("UPSIZE_TS") = dt
        newRow.EndEdit()

        dtbl.Rows.Add(newRow)

        '20190219 LEE: Don't need anymore. Used GetMaxID
        'Call PutMaxID("TBLWORDSTATEMENTS", idNew)

        Me.dgvReportStatements.AutoResizeRows()

        boolHold = False

        varTitleBefore = ""
        varTitleAfter = Me.CHARTITLE.Text

    End Sub

    Private Sub ov1_RegionChanged(sender As Object, e As System.EventArgs) Handles ov1.RegionChanged

        '20170824 LEE: Not needed anymore

        'Dim strM As String

        'If Me.cmdEdit.Enabled And Me.pan2.Visible = False Then
        '    strM = "B: Please note that the 'Edit' button has not been clicked." & ChrW(10) & ChrW(10)
        '    strM = strM & "Any changes made to the document cannot be saved unless the 'Edit' button is clicked."
        '    MsgBox(strM, MsgBoxStyle.Information, "Non-edit mode...")
        'End If

    End Sub

    Private Sub tSave_Tick(sender As System.Object, e As System.EventArgs) Handles tSave.Tick

        Dim boolF As Boolean = Me.lblBlink.Font.Bold

        If boolF Then
            Me.lblBlink.Font = New Font(Me.lblBlink.Font, FontStyle.Regular)
        Else
            Me.lblBlink.Font = New Font(Me.lblBlink.Font, FontStyle.Bold)
        End If

        Me.tSave.Interval = 1000


    End Sub

    Private Sub cmdCompareDocs_Click(sender As Object, e As EventArgs) Handles cmdCompareDocs.Click

        Dim dgvV As DataGridView = Me.dgvVersions
        Dim strM As String
        Dim intR As Short

        strM = "Ensure that the desired Report Template is selected in the Report Templates table."
        intR = MsgBox(strM, vbOKCancel, "Ensure...")
        If intR = 1 Then
        Else
            GoTo end1
        End If

        If dgvV.RowCount < 2 Then
            strM = "This document has only one version."
            strM = strM & ChrW(10) & ChrW(10) & "There must be at least two versions to compare."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        Dim intRow As Int32

        intRow = Me.dgvReportStatements.CurrentRow.Index
        Dim id As Int64
        id = Me.dgvReportStatements("ID_TBLWORDSTATEMENTS", intRow).Value
        Dim strTitle As String
        strTitle = Me.dgvReportStatements("CHARTITLE", intRow).Value

        Dim frm As New frmDocumentCompare

        frm.boolFormLoad = True
        frm.gDoc = "Report Template"
        frm.txtWSID.Text = id.ToString
        frm.txtWordTemplate.Text = strTitle
        frm.strPrevForm = "WordStatement"
        frm.Text = "Document Compare"

        Dim dv As New DataView
        dv = Me.dgvVersions.DataSource

        'cannot use datasource with combobox
        'I want to show two items: intVersion and charDescription

        Dim boolHT As Boolean
        boolHT = boolHold
        boolHold = True

        Dim Count1 As Int64
        Dim num1
        Dim str1 As String
        Dim str2 As String
        Dim str3 As String
        Dim var1

        For Count1 = 0 To dv.Count - 1
            num1 = NZ(dv(Count1).Item("INTWORDVERSION"), 0)
            str2 = NZ(dv(Count1).Item("CHARDESCRIPTION"), 0).ToString
            str3 = Format(num1, "00000") & "   " & str2
            frm.cbxWRT1.Items.Add(str3)
            frm.cbxWRT2.Items.Add(str3)
        Next
        frm.cbxWRT1.SelectedIndex = -1
        frm.cbxWRT2.SelectedIndex = -1

        boolHold = boolHT

        frm.boolFormLoad = False
        frm.frm = Me
        Cursor.Current = Cursors.Default
        frm.Show(Me)
        Me.Visible = False

        GoTo end1

        Try
            frm.Dispose()
        Catch ex As Exception
            var1 = var1
        End Try

        Call ClearTemp()

        Cursor.Current = Cursors.Default
        Me.BringToFront()
        Me.Activate()
        Me.Visible = True

end1:

    End Sub

    Private Sub dgvVersions_CellContentDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvVersions.CellContentDoubleClick

        If Me.lblReadOnly.Visible Then
        Else
            Call EditClick()
        End If

    End Sub

    Private Sub dgvVersions_MouseEnter(sender As Object, e As EventArgs) Handles dgvVersions.MouseEnter

        Me.cmdShow.Visible = False

    End Sub


    Private Sub dgvVersions_SelectionChanged(sender As Object, e As EventArgs) Handles dgvVersions.SelectionChanged

        Dim var1

        'NDL: Do nothing if programmatically changed during creation of a new template
        If boolcmdAddStatementJustClicked Then
            Exit Sub
        End If

        If Me.ov1.IsOpened Then

            Dim wdDoc As Microsoft.Office.Interop.Word.Document

            Try
                wdDoc = Me.ov1.ActiveDocument
            Catch ex As Exception
                var1 = var1
            End Try


            Try
                wdDoc.Application.DisplayAlerts = False
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                Call SetWordStuff(wdDoc)
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                Me.ov1.CloseDoc()
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                wdDoc.Application.DisplayAlerts = True
            Catch ex As Exception
                var1 = var1
            End Try

            Try
                wdDoc.Application.Quit(False)
            Catch ex As Exception
                var1 = var1
            End Try

            'Try
            '    Me.ov1.CloseDoc()
            'Catch ex As Exception
            '    var1 = var1
            'End Try
        Else
            var1 = var1
        End If

        Call DoVersionsSelChange()

    End Sub

    Sub DoVersionsSelChange()

        'for some reason, this is looping

        If boolFormLoad Then
            Exit Sub
        End If

        If boolHold Then
            Exit Sub
        End If

        If boolReadOnly Then
            Dim boolH As Boolean
            boolH = boolHold
            boolHold = True
            Call VersionChange()
            boolHold = boolH
        End If

        Me.dgvReportStatements.AutoResizeRows()
        Me.dgvVersions.AutoResizeRows()

        Call ClearTemp()

    End Sub

    Private Sub dgvVersions_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvVersions.CellContentClick

    End Sub

    Private Sub dgvReportStatements_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReportStatements.CellContentClick

    End Sub

    Private Sub cmdShow_Click(sender As Object, e As EventArgs) Handles cmdShow.Click

        Me.cmdShow.Visible = False

        If boolEdit Then
        Else
            Call VersionChange()
        End If

    End Sub

    Private Sub dgvReportStatements_MouseEnter(sender As Object, e As EventArgs) Handles dgvReportStatements.MouseEnter

        Me.cmdShow.Visible = False

    End Sub

    Private Sub cmdInsertNew_Click(sender As Object, e As EventArgs) Handles cmdInsertNew.Click

        Dim strM As String
        Dim intR As Short

        strM = "Are you sure you want to clear the document?"
        intR = MsgBox(strM, vbOKCancel, "Continue?")

        If intR = 1 Then
        Else
            GoTo end1
        End If

        Call InsertNewDocument()

end1:


    End Sub

    Sub InsertNewDocument()

        Try
            Me.ov1.CloseDoc()
            Try
                Me.ov1.CreateNew("Word.Application")

                Call ResetToTemp()

            Catch ex As Exception
                MsgBox("CreateNew: " & ex.Message)
                GoTo end1
            End Try

        Catch ex As Exception
            MsgBox("closedoc: " & ex.Message)
            GoTo end1
        End Try

end1:

    End Sub

    Private Sub cmdOpenExisting_Click(sender As Object, e As EventArgs) Handles cmdOpenExisting.Click

        Call MakeTemplateFromExistingWordDoc(False, "")

    End Sub

    Sub MakeTemplateFromExistingWordDoc(boolFromNew As Boolean, strPath As String)

        Dim strM As String
        Dim intR As Short

        If boolFromNew Then
        Else
            strM = "Are you sure you want to open an existing document?"
            intR = MsgBox(strM, vbOKCancel, "Continue?")

            If intR = 1 Then
            Else
                GoTo end1
            End If
        End If

        'Me.ov1.OpenFileDialog("Microsoft Excel Files(*.xl;*.xlsx;*.xlsb;*.xlam;*.xltx;*.xltm;*.xls;*.xlt;*.xla;*.xlm;*.xlw)|*.xl;*.xlsx;*.xlsb;*.xlam;*.xltx;*.xltm;*.xls;*.xlt;*.xla;*.xlm;*.xlw||")

        'Me.ov1.OpenFileDialog()

        Try
            Me.ov1.CloseDoc()
            Try

                If boolFromNew Then
                    Me.ov1.OpenWord(strPath)
                Else
                    Me.ov1.OpenFileDialog("Microsoft Word Files(*.doc*)|*.doc*||")

                    'the previous action leaves the current document as the original document
                    'must save as a temp file, then re-open

                    Call ResetToTemp()

                End If



            Catch ex As Exception
                MsgBox("MakeTemplateFromExistingWordDoc: " & ex.Message, vbInformation, "Problem opening existing Word document...")
                GoTo end1
            End Try

        Catch ex As Exception
            MsgBox("closedoc: " & ex.Message)
            GoTo end1
        End Try

end1:

    End Sub

    Sub ResetToTemp()

        Dim wdDoc As Microsoft.Office.Interop.Word.Document
        Dim ver
        Dim bool2007 As Boolean = False
        Dim strR As String
        Dim boolM As Boolean = False

        Call ClearTemp()

        Try
            wdDoc = Me.ov1.ActiveDocument
            ver = CInt(wdDoc.Application.Version)
            If ver < 13 Then
                bool2007 = True
            Else
                bool2007 = False
            End If
            strPath = GetNewTempFile(True)
            boolHasMacro = False
            If bool2007 Then
                'Note: Word2007 has no compatability mode parameter
                If wdDoc.HasVBProject Then
                    strR = Replace(strPath, ".xml", ".docm", 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled)
                    boolM = True
                    boolHasMacro = True
                Else
                    strR = Replace(strPath, ".xml", ".docx", 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument)
                    boolM = False
                End If
            Else
                If wdDoc.HasVBProject Then
                    strR = Replace(strPath, ".xml", ".docm", 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs2(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolM = True
                    boolHasMacro = True
                Else
                    strR = Replace(strPath, ".xml", ".docx", 1, -1, CompareMethod.Text)
                    wdDoc.SaveAs2(FileName:=strR, FileFormat:=Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode:=Microsoft.Office.Interop.Word.WdCompatibilityMode.wdCurrent)
                    boolM = False
                End If
            End If

            Try
                Call SetWordStuff(wdDoc)
            Catch ex As Exception

            End Try

        Catch ex As Exception

        End Try

        Pause(0.25)

    End Sub

    Private Sub frmWordStatement_ToolTipsSet()
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

        'General Buttons
        toolTip1.SetToolTip(Me.cmdWord1, "Open in Microsoft Word (independent of StudyDoc)")
        toolTip1.SetToolTip(Me.cmdFieldCode, "Add StudyDoc Field Code to document")
        toolTip1.SetToolTip(Me.cmdAddStatement, "Create new Microsoft Word template")
        toolTip1.SetToolTip(Me.cmdOpenExisting, "Import existing document")
        toolTip1.SetToolTip(Me.cmdInsertNew, "Create new blank document")
        toolTip1.SetToolTip(Me.cmdExit, "Exit Word Template editor")
        toolTip1.SetToolTip(Me.cmdInsertNew, "Create new blank document")
        toolTip1.SetToolTip(Me.cmdCompareDocs, "Compare two versions of the document using Word")
        toolTip1.SetToolTip(Me.cmdWord, "Open in Microsoft Word (independent of StudyDoc)")
        toolTip1.SetToolTip(Me.cmdOpenPDF, "Open in PDF Viewer (independent of StudyDoc)")
        toolTip1.SetToolTip(Me.cmdExit2, "Exit Word Template viewer")
    End Sub

    Private Sub cmdDeactivateTemplates_Click(sender As Object, e As EventArgs) Handles cmdDeactivateTemplates.Click

        Call DeactivateTemplates()

    End Sub

    Sub DeactivateTemplates()

        Dim frm As New frmWordStatementsActiveTemplates

        Dim a, b, c, d

        a = Me.pan1.Left + Me.pan1.Width
        b = Me.pan1.Top

        Dim BorderWidth As Single = (Me.Width - Me.ClientSize.Width) / 2
        Dim TitlebarHeight As Single = Me.Height - Me.ClientSize.Height - 2 * BorderWidth

        frm.Top = b + TitlebarHeight + 10
        frm.Left = a + 10


        frm.ShowDialog()

        Call UpdateRSW()

    End Sub

    Private Sub cmdPDF_Click(sender As Object, e As EventArgs) Handles cmdPDF.Click

        Call DoPDF()

    End Sub

    Sub DoPDF()

        Dim strM As String

        If Me.ov1.IsOpened Then
        Else
            strM = "A document must be loaded."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        'Public boolReport As Boolean = False 'boolReport is a Report Template
        If boolReport Then 'Report Template
            If BOOLALLOWREPORTTEMPLATEPDF Then
            Else
                strM = "User does not have permissions to create a PDF file."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            'BOOLALLOWPDFREPORT = False'testing
            If BOOLALLOWPDFREPORT Or BOOLFORCEFINALREPORTPDF Then
            Else
                strM = "User does not have permissions to create a PDF file."
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If

        Call DoWrite()

        Call GoToWord(True)

        Call ClearTemp()

end1:

    End Sub

    Private Sub cmdPrint_Click(sender As Object, e As EventArgs) Handles cmdPrint.Click

        Dim strM As String

        'Public boolReport As Boolean = False 'boolReport is a Report Template
        strM = "User does not have permission to print this document."
        If boolReport Then
            If BOOLALLOWRTEMPLATEPRINT Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            If BOOLALLOWFINALREPORTPRINT Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If

        If Me.ov1.IsOpened Then
        Else
            strM = "A document must be loaded."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        'Dim strPath As String

        'strPath = Me.ov1.GetDocumentFullName
        'Me.ov1.CloseDoc()
        'Me.ov1.OpenWord(strPath)

        Me.ov1.PrintDialog()

end1:

    End Sub

    Private Sub cmdPrint1_Click(sender As Object, e As EventArgs) Handles cmdPrint1.Click

        Dim strM As String

        'Public boolReport As Boolean = False 'boolReport is a Report Template
        strM = "User does not have permission to print this document."
        If boolReport Then
            If BOOLALLOWRTEMPLATEPRINT Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        Else
            If BOOLALLOWFINALREPORTPRINT Then
            Else
                MsgBox(strM, vbInformation, "Invalid action...")
                GoTo end1
            End If
        End If

        If Me.ov1.IsOpened Then
        Else
            strM = "A document must be loaded."
            MsgBox(strM, vbInformation, "Invalid action...")
            GoTo end1
        End If

        Me.ov1.PrintDialog()

end1:

    End Sub

  
    Private Sub cmdEditTitle_Click(sender As Object, e As EventArgs) Handles cmdEditTitle.Click

        Me.panSave.Enabled = False
        Me.panList.Enabled = False
        Me.panButtons.Enabled = False

        Me.panEditReports.Enabled = True
        Me.CHARTITLE.Enabled = False
        Me.cmdEditStatements.Enabled = True
        Me.cmdCancel.Enabled = True
        Me.panEditReports.Visible = True

    End Sub

    Private Sub CHARTITLE_TextChanged(sender As Object, e As EventArgs) Handles CHARTITLE.TextChanged

    End Sub


    Private Sub CHARTITLE_Validating(sender As Object, e As ComponentModel.CancelEventArgs) Handles CHARTITLE.Validating

        Dim boolGo As Boolean

        If Me.panEditReports.Visible Then
        Else
            boolGo = False
            GoTo end2
        End If

        'If Me.cmdEditStatements.Enabled Then
        '    boolGo = True
        '    GoTo end2
        'End If

        Dim strInput As String
        Dim strM1 As String

        strInput = Me.CHARTITLE.Text
        'blank is not acceptable
        If Len(strInput) = 0 Then
            strM1 = "Proposed report template name cannot be blank."
            MsgBox(strM1, MsgBoxStyle.Information, "Invalid entry...")
            e.Cancel = True
            GoTo end1

        End If
        'check to ensure name is unique
        If RNUnique(tblWordStatements, "CHARTITLE", strInput) Then
            boolGo = True
        Else
            boolGo = False
            e.Cancel = True
        End If


end2:

        If boolGo Then
            varTitleAfter = Me.CHARTITLE.Text

        End If

end1:

    End Sub

    Private Sub dgvReportStatements_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReportStatements.CellDoubleClick

        If Me.lblReadOnly.Visible Then
        Else
            Call EditClick()
        End If

    End Sub
End Class